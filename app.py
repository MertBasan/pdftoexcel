"""
Bank Statement → CSV/Excel Extractor (Accountant-Grade)
========================================================

Key design points
-----------------
* Text-based extraction (NOT OCR). PDFs from major Turkish banks are digitally
  generated and fully text-extractable; OCR would only add error.
* Per-bank parsers. Currently: Halkbank. Bank is auto-detected by fingerprint.
* Two mathematical accuracy checks on every file:
    1. Balance chain: prev_Bakiye + current_TUTAR == current_Bakiye for every
       consecutive row pair. If even one row fails, extraction broke.
    2. Final balance: last extracted Bakiye must equal the "Hesap Bakiyesi"
       printed on the statement header.
  Both green ⇒ TUTAR and Bakiye are bit-exact what's in the PDF.

Output columns (per user spec)
------------------------------
TARİH | Saat | TUTAR | Bakiye | AÇIKLAMA | DEKONT

Notes per bank
--------------
Halkbank PDFs do NOT contain a Saat column or a separate DEKONT column. For
Halkbank:
  * Saat is left empty (source has no time-of-day field).
  * DEKONT is best-effort: the trailing 10+ digit reference number after the
    last "/" in POS-style descriptions. HAVALE/transfer rows have no DEKONT.

Tested against a 48-page, 1,644-transaction Halkbank statement: 0 balance-chain
mismatches, final balance reconciles to the header (18.707,63 = 18.707,63).
"""

from __future__ import annotations

import io
import os
import re
import zipfile
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from typing import Callable, Optional

import pandas as pd
import pdfplumber
import streamlit as st


# =============================================================================
# 1. Number parsing
# =============================================================================

def parse_tr_decimal(s: str) -> Decimal:
    """Convert Turkish-formatted number string to Decimal.

    Examples: '1.234,56' -> Decimal('1234.56'); '-476,19' -> Decimal('-476.19').
    """
    s = s.strip()
    if not s:
        raise InvalidOperation("empty number string")
    neg = s.startswith('-')
    if neg:
        s = s[1:]
    s = s.replace('.', '').replace(',', '.')
    d = Decimal(s)
    return -d if neg else d


def format_tr_decimal(d: Optional[Decimal]) -> str:
    if d is None:
        return ''
    s = f"{d:,.2f}"
    return s.replace(',', 'X').replace('.', ',').replace('X', '.')


# =============================================================================
# 2. Result containers
# =============================================================================

@dataclass
class TransactionRow:
    date: str
    time: str = ''
    amount: Decimal = Decimal(0)
    balance: Decimal = Decimal(0)
    description: str = ''
    receipt: str = ''


@dataclass
class StatementMetadata:
    bank: str = ''
    customer_no: str = ''
    customer_name: str = ''
    account_no: str = ''
    iban: str = ''
    branch: str = ''
    currency: str = ''
    period_start: str = ''
    period_end: str = ''
    stated_balance: Optional[Decimal] = None


@dataclass
class BalanceIssue:
    row_index: int
    date: str
    expected: Decimal
    actual: Decimal
    diff: Decimal


@dataclass
class StatementResult:
    source_filename: str
    metadata: StatementMetadata
    rows: list[TransactionRow] = field(default_factory=list)
    balance_issues: list[BalanceIssue] = field(default_factory=list)
    final_balance_match: Optional[bool] = None
    parser_warnings: list[str] = field(default_factory=list)


# =============================================================================
# 3. Bank detection
# =============================================================================

def detect_bank(full_text: str) -> Optional[str]:
    head = full_text[:5000].upper()
    if 'HALKBANK' in head or 'TÜRKIYE HALK BANKASI' in head or 'TURKIYE HALK BANKASI' in head:
        return 'HALKBANK'
    if 'AKBANK' in head or 'AKPOS' in head:
        return 'AKBANK'
    if 'GARANTİ' in head or 'GARANTI BANKASI' in head:
        return 'GARANTI'
    if 'İŞ BANKASI' in head or 'IS BANKASI' in head or 'TÜRKİYE İŞ BANKASI' in head:
        return 'ISBANK'
    if 'YAPI VE KREDİ' in head or 'YAPI KREDI' in head:
        return 'YAPIKREDI'
    if 'ZİRAAT' in head or 'ZIRAAT BANKASI' in head:
        return 'ZIRAAT'
    return None


# =============================================================================
# 4. Halkbank parser
# =============================================================================

_HB_SKIP_CONT_PREFIXES = (
    'Müşteri', 'Hesap', 'TCKN', 'IBAN', 'Şube', 'Döviz', 'Üretim', 'Dönemi',
    'Bakiye Bilgi', 'Bloke', 'Kullanıl', 'Toplam Kredi',
    'Türkiye Halk', 'yerine kullanıl', 'Uyuşmazlık',
)
_HB_SKIP_EXACT = {'HESAP ÖZETİ'}

_HB_NUM = r'-?[\d.]+,\d{2}'
_HB_ROW = re.compile(rf'^(\d{{2}}-\d{{2}}-\d{{4}})\s+({_HB_NUM})\s+({_HB_NUM})\s+(.*)$')

# Best-effort DEKONT for Halkbank: last "/" followed by 10+ digits at end of
# the joined description. Catches POS Satis / POS Aidat / Komisyon receipts.
# HAVALE rows end with a masked IBAN (TR54***...8675) and produce no DEKONT.
_HB_DEKONT_RE = re.compile(r'/(\d{10,})\s*$')


def _hb_extract_dekont(description: str) -> str:
    m = _HB_DEKONT_RE.search(description)
    return m.group(1) if m else ''


def _hb_strip_right_column(line: str) -> str:
    """Halkbank's cover page is a two-column layout. pdfplumber merges
    same-baseline content from both columns onto a single line. Strip away
    the right-column part so we can read left-column values cleanly.

    Example input:
        'Şube Kodu / Adı :GÜLTEPE / ISTANBUL SB. Kullanılabilir Kredi Limiti :0,00'
    Example output:
        'Şube Kodu / Adı :GÜLTEPE / ISTANBUL SB.'
    """
    # Right-column labels that may be merged onto a left-column line
    right_anchors = (
        r'\s+(?:Üretim\s+Zamanı|Dönemi|Hesap\s+Bakiyesi|Bloke\s+Bakiyesi|'
        r'Kullanılabilir(?:\s+\w+)*|Toplam\s+Kredi|Bakiye\s+Bilgileriniz|'
        r'Hesap\s+Özeti)\s*:'
    )
    return re.split(right_anchors, line, maxsplit=1)[0].rstrip()


def _hb_extract_customer_name(text: str) -> str:
    """Extract the (possibly multi-line) customer name.

    The label 'Müşteri Adı / Ünvanı :' wraps to two visual lines in the PDF.
    Because of that, pdfplumber emits the customer name's FIRST line BEFORE
    the label line, and the SECOND line AFTER it. Concretely:

        line N-1: SEM LOKANTA SANAYI VE TICARET LIMITED  Dönemi :01.01.2026...
        line N:   Müşteri Adı / Ünvanı :
        line N+1: SIRKETI
        line N+2: TCKN / VKN :7600928****

    So we collect:
      * the line above the label (stripping right-column garbage)
      * lines below the label until we hit the next field label
    """
    lines = text.split('\n')
    label_idx = next(
        (i for i, l in enumerate(lines) if 'Müşteri Adı / Ünvanı' in l),
        None,
    )
    if label_idx is None:
        return ''

    next_label_re = re.compile(r'^(TCKN|Hesap|Bakiye|Şube|Döviz|IBAN|Bloke|Toplam)\b')
    parts: list[str] = []

    # Line above the label — likely the first line of the name
    if label_idx > 0:
        cand = _hb_strip_right_column(lines[label_idx - 1]).strip()
        # Skip if it's actually a different label (e.g. "Müşteri Numarası : ...")
        if cand and not cand.startswith(('Müşteri', 'Hesap', 'TCKN', 'IBAN',
                                         'Şube', 'Döviz', 'Bakiye', 'Bloke',
                                         'Kullanıl', 'Toplam', 'Üretim', 'Dönemi')):
            parts.append(cand)

    # Lines below the label until next field
    for j in range(label_idx + 1, min(label_idx + 5, len(lines))):
        cand = _hb_strip_right_column(lines[j]).strip()
        if not cand:
            continue
        if next_label_re.match(cand):
            break
        parts.append(cand)

    return ' '.join(parts).strip()


def _hb_extract_metadata(pdf_bytes: bytes) -> StatementMetadata:
    """Extract cover-page metadata from page 1 (no cropping — pdfplumber's
    text ordering on cropped regions is unreliable for this layout).
    """
    md = StatementMetadata(bank='HALKBANK')
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text = pdf.pages[0].extract_text() or ''
    except Exception:
        return md

    # Strip right-column garbage from each line before regex matching
    cleaned_lines = [_hb_strip_right_column(l) for l in text.split('\n')]
    cleaned = '\n'.join(cleaned_lines)

    def grab(pattern: str, source: str = cleaned) -> Optional[str]:
        m = re.search(pattern, source)
        return m.group(1).strip() if m else None

    md.customer_no   = grab(r'Müşteri Numarası\s*:\s*(\S+)') or ''
    md.customer_name = _hb_extract_customer_name(text)
    md.account_no    = grab(r'Hesap No\s*:\s*(\S+)') or ''
    md.iban          = grab(r'IBAN\s*:\s*(\S+)') or ''
    md.branch        = grab(r'Şube Kodu / Adı\s*:\s*([^\n]+)') or ''
    md.currency      = grab(r'Döviz Cinsi\s*:\s*(\S+)') or ''

    # Period and stated balance live in the right column — search RAW text
    period = grab(r'Dönemi\s*:\s*([\d./]+\s*-\s*[\d./]+)', source=text)
    if period and '-' in period:
        a, b = period.split('-', 1)
        md.period_start, md.period_end = a.strip(), b.strip()

    bal = grab(r'Hesap Bakiyesi\s*:\s*([\-\d.,]+)', source=text)
    if bal:
        try:
            md.stated_balance = parse_tr_decimal(bal)
        except InvalidOperation:
            pass

    return md


def _hb_parse_transactions(full_text: str) -> tuple[list[TransactionRow], list[str]]:
    """Halkbank wraps long descriptions across multiple visual lines, breaking
    mid-character. Continuation lines must be joined WITHOUT a separator so
    that '...687' + '2' = '...6872' stays intact.
    """
    rows: list[TransactionRow] = []
    warnings: list[str] = []
    current: Optional[TransactionRow] = None

    for raw in full_text.split('\n'):
        line = raw.rstrip('\r').rstrip()
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith('İşlem Tarihi') or stripped.startswith('Sayfa No'):
            continue
        if stripped in _HB_SKIP_EXACT:
            continue

        m = _HB_ROW.match(line)
        if m:
            if current is not None:
                rows.append(current)
            date_str, amount_str, balance_str, desc = m.groups()
            try:
                amount = parse_tr_decimal(amount_str)
                balance = parse_tr_decimal(balance_str)
            except InvalidOperation as exc:
                warnings.append(f"Skipped malformed row at {date_str}: {exc}")
                current = None
                continue
            current = TransactionRow(
                date=date_str,
                amount=amount,
                balance=balance,
                description=desc.strip(),
            )
        else:
            if current is None:
                continue
            if stripped.startswith(_HB_SKIP_CONT_PREFIXES):
                continue
            current.description += stripped

    if current is not None:
        rows.append(current)

    for r in rows:
        r.receipt = _hb_extract_dekont(r.description)

    return rows, warnings


def parse_halkbank(pdf_bytes: bytes, full_text: str, source_filename: str) -> StatementResult:
    metadata = _hb_extract_metadata(pdf_bytes)
    rows, warnings = _hb_parse_transactions(full_text)
    return StatementResult(
        source_filename=source_filename,
        metadata=metadata,
        rows=rows,
        parser_warnings=warnings,
    )


# =============================================================================
# 5. Parser registry
# =============================================================================

PARSERS: dict[str, Callable[[bytes, str, str], StatementResult]] = {
    'HALKBANK': parse_halkbank,
    # TODO: 'AKBANK': parse_akbank — different layout (has Saat & DEKONT columns)
}


# =============================================================================
# 6. Validation
# =============================================================================

def validate_balance_chain(rows: list[TransactionRow]) -> list[BalanceIssue]:
    issues: list[BalanceIssue] = []
    for i in range(1, len(rows)):
        expected = rows[i - 1].balance + rows[i].amount
        if expected != rows[i].balance:
            issues.append(BalanceIssue(
                row_index=i,
                date=rows[i].date,
                expected=expected,
                actual=rows[i].balance,
                diff=rows[i].balance - expected,
            ))
    return issues


def validate_final_balance(result: StatementResult) -> Optional[bool]:
    if not result.rows or result.metadata.stated_balance is None:
        return None
    return result.rows[-1].balance == result.metadata.stated_balance


# =============================================================================
# 7. PDF text extraction
# =============================================================================

def extract_pdf_text(pdf_bytes: bytes) -> str:
    parts: list[str] = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or '')
    return '\n'.join(parts)


# =============================================================================
# 8. Pipeline
# =============================================================================

def process_pdf(pdf_bytes: bytes, filename: str) -> StatementResult:
    text = extract_pdf_text(pdf_bytes)
    bank = detect_bank(text)

    if bank is None:
        result = StatementResult(source_filename=filename, metadata=StatementMetadata())
        result.parser_warnings.append(
            "Unrecognized bank. Currently supported: " + ", ".join(PARSERS.keys())
        )
        return result

    if bank not in PARSERS:
        result = StatementResult(source_filename=filename,
                                 metadata=StatementMetadata(bank=bank))
        result.parser_warnings.append(
            f"Detected {bank} but no parser is implemented yet. "
            f"Currently implemented: {', '.join(PARSERS.keys())}."
        )
        return result

    result = PARSERS[bank](pdf_bytes, text, filename)
    result.balance_issues = validate_balance_chain(result.rows)
    result.final_balance_match = validate_final_balance(result)
    return result


# =============================================================================
# 9. Output
# =============================================================================

# Per user spec — only these six columns.
OUTPUT_COLUMNS = ['TARİH', 'Saat', 'TUTAR', 'Bakiye', 'AÇIKLAMA', 'DEKONT']


def result_to_dataframe(result: StatementResult) -> pd.DataFrame:
    if not result.rows:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    df = pd.DataFrame({
        'TARİH':    [r.date for r in result.rows],
        'Saat':     [r.time for r in result.rows],
        'TUTAR':    [float(r.amount) for r in result.rows],
        'Bakiye':   [float(r.balance) for r in result.rows],
        'AÇIKLAMA': [r.description for r in result.rows],
        'DEKONT':   [r.receipt for r in result.rows],
    })[OUTPUT_COLUMNS]
    df['TARİH'] = pd.to_datetime(df['TARİH'], format='%d-%m-%Y', errors='coerce').dt.date
    return df


def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    """Use openpyxl — pandas' default xlsx engine, no extra install needed
    beyond pandas itself in most setups.

    Important: DEKONT contains long numeric-looking strings with leading zeros
    (e.g. '000001813395187'). We force that column to text format so Excel
    doesn't reinterpret it as a float and show '1.81E+09'.
    """
    from openpyxl.utils import get_column_letter

    # Make sure DEKONT is stored as string with leading zeros preserved
    df = df.copy()
    if 'DEKONT' in df.columns:
        df['DEKONT'] = df['DEKONT'].astype(str).replace({'nan': '', 'None': ''})

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Hareketler', index=False)
        ws = writer.sheets['Hareketler']

        for i, col in enumerate(df.columns, start=1):
            try:
                sample = df[col].astype(str)
                width = min(60, max(12, int(sample.str.len().quantile(0.95)) + 2))
            except Exception:
                width = 18
            ws.column_dimensions[get_column_letter(i)].width = width

        # Money formatting for TUTAR / Bakiye
        for col_name in ('TUTAR', 'Bakiye'):
            if col_name in df.columns:
                col_letter = get_column_letter(df.columns.get_loc(col_name) + 1)
                for row_idx in range(2, len(df) + 2):
                    ws[f'{col_letter}{row_idx}'].number_format = '#,##0.00'

        # CRITICAL: force DEKONT to text format. The values are numeric-looking
        # strings with leading zeros like '000001813395187' — without this,
        # Excel reads them as floats and shows '1.81E+09', destroying the data.
        if 'DEKONT' in df.columns:
            col_letter = get_column_letter(df.columns.get_loc('DEKONT') + 1)
            for row_idx in range(2, len(df) + 2):
                cell = ws[f'{col_letter}{row_idx}']
                cell.number_format = '@'
                # Also force the value itself to be a string
                if cell.value is not None and cell.value != '':
                    cell.value = str(cell.value)
    return buf.getvalue()


def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode('utf-8-sig')


# =============================================================================
# 10. Filename
# =============================================================================

_SAFE_RE = re.compile(r'[^A-Za-z0-9._\-]+')

def safe_filename(name: str, max_len: int = 80) -> str:
    cleaned = _SAFE_RE.sub('_', name).strip('_')
    return cleaned[:max_len] if cleaned else 'output'


def build_output_basename(result: StatementResult) -> str:
    pdf_base = os.path.splitext(os.path.basename(result.source_filename))[0]
    cust = result.metadata.customer_name
    if cust:
        cust_short = ' '.join(cust.split()[:3])
        return safe_filename(f"{pdf_base}__{cust_short}")
    return safe_filename(pdf_base)


# =============================================================================
# 11. Upload handling (PDF + ZIP)
# =============================================================================

def iter_uploaded(uploaded_files) -> list[tuple[str, bytes]]:
    out: list[tuple[str, bytes]] = []
    for uf in uploaded_files:
        name = uf.name
        data = uf.read()
        lname = name.lower()
        if lname.endswith('.pdf'):
            out.append((name, data))
        elif lname.endswith('.zip'):
            try:
                with zipfile.ZipFile(io.BytesIO(data)) as z:
                    for inner in z.namelist():
                        if inner.lower().endswith('.pdf') and not inner.endswith('/'):
                            out.append((os.path.basename(inner), z.read(inner)))
            except zipfile.BadZipFile:
                st.error(f"Could not read {name} as ZIP.")
        else:
            st.warning(f"Skipping unsupported file: {name}")
    return out


# =============================================================================
# 12. Streamlit UI
# =============================================================================

def render_metadata(md: StatementMetadata) -> None:
    cols = st.columns(3)
    cols[0].markdown(f"**Bank**: {md.bank or '—'}")
    cols[0].markdown(f"**Müşteri**: {md.customer_name or '—'}")
    cols[1].markdown(f"**Hesap No**: {md.account_no or '—'}")
    cols[1].markdown(f"**IBAN**: {md.iban or '—'}")
    cols[2].markdown(f"**Dönem**: {md.period_start or '—'} → {md.period_end or '—'}")
    cols[2].markdown(f"**Şube**: {md.branch or '—'}")


def render_validation(result: StatementResult) -> None:
    n_rows = len(result.rows)
    issues = result.balance_issues
    final_match = result.final_balance_match

    if n_rows == 0:
        st.error("❌ No transactions parsed.")
        return

    if not issues:
        st.success(
            f"✅ Balance chain verified across **{n_rows-1}** consecutive row pairs. "
            f"All {n_rows} transactions reconcile mathematically — TUTAR and Bakiye "
            f"are bit-exact what's in the PDF."
        )
    else:
        st.error(
            f"❌ Balance mismatches on {len(issues)} row(s). "
            f"DO NOT TRUST this output — extraction is broken for at least one row."
        )
        with st.expander(f"Show first {min(10, len(issues))} balance issues"):
            for iss in issues[:10]:
                st.write(
                    f"Row {iss.row_index} ({iss.date}): "
                    f"expected {format_tr_decimal(iss.expected)}, "
                    f"actual {format_tr_decimal(iss.actual)}, "
                    f"diff {format_tr_decimal(iss.diff)}"
                )

    if final_match is True:
        st.success(
            f"✅ Final balance matches statement header: "
            f"**{format_tr_decimal(result.metadata.stated_balance)}** "
            f"(no rows missed or duplicated)."
        )
    elif final_match is False:
        last = result.rows[-1].balance
        stated = result.metadata.stated_balance
        st.error(
            f"❌ Final balance mismatch — last row says "
            f"**{format_tr_decimal(last)}**, statement header says "
            f"**{format_tr_decimal(stated)}**. Some transactions may have been "
            f"missed or duplicated."
        )

    for w in result.parser_warnings:
        st.warning(w)


def main() -> None:
    st.set_page_config(page_title="Bank Statement Extractor", layout="wide")
    st.title("Bank Statement → Excel/CSV")
    st.caption(
        "Accountant-grade extraction for Turkish bank PDFs. "
        "Currently supports: " + ", ".join(PARSERS.keys()) + "."
    )

    if 'results' not in st.session_state:
        st.session_state.results = []

    uploaded = st.file_uploader(
        "Upload PDF statements or a ZIP folder containing them",
        type=['pdf', 'zip'],
        accept_multiple_files=True,
    )

    col_a, col_b = st.columns([1, 1])
    if col_a.button("Process uploads", type="primary", disabled=not uploaded):
        files = iter_uploaded(uploaded)
        if not files:
            st.warning("No PDFs found in uploads.")
        new_results: list[StatementResult] = []
        progress = st.progress(0.0, text="Processing…")
        for i, (name, data) in enumerate(files, start=1):
            try:
                res = process_pdf(data, name)
            except Exception as exc:
                res = StatementResult(source_filename=name, metadata=StatementMetadata())
                res.parser_warnings.append(f"Fatal error: {exc}")
            new_results.append(res)
            progress.progress(i / len(files), text=f"Processed {i}/{len(files)}")
        progress.empty()
        st.session_state.results.extend(new_results)
        st.success(f"Added {len(new_results)} file(s). Total: "
                   f"{len(st.session_state.results)}.")

    if col_b.button("Clear all", disabled=not st.session_state.results):
        st.session_state.results = []
        st.rerun()

    if not st.session_state.results:
        st.info("No statements processed yet. Upload PDFs above to begin.")
        return

    st.divider()
    st.subheader("Per-file results")
    for idx, res in enumerate(st.session_state.results):
        n_rows = len(res.rows)
        ok = (not res.balance_issues) and (res.final_balance_match is not False) and n_rows > 0
        badge = "✅" if ok else ("⚠️" if n_rows > 0 else "❌")
        with st.expander(f"{badge} {res.source_filename} — {n_rows} rows", expanded=False):
            render_metadata(res.metadata)
            st.markdown("---")
            render_validation(res)

            if n_rows > 0:
                df = result_to_dataframe(res)
                st.dataframe(df.head(20), use_container_width=True)

                base = build_output_basename(res)
                dl1, dl2 = st.columns(2)
                dl1.download_button(
                    f"⬇ {base}.xlsx",
                    data=df_to_xlsx_bytes(df),
                    file_name=f"{base}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"xlsx_{idx}",
                )
                dl2.download_button(
                    f"⬇ {base}.csv",
                    data=df_to_csv_bytes(df),
                    file_name=f"{base}.csv",
                    mime="text/csv",
                    key=f"csv_{idx}",
                )

    st.divider()
    st.subheader("Combined output (all uploaded files)")
    all_dfs = [result_to_dataframe(r) for r in st.session_state.results if r.rows]
    if not all_dfs:
        st.info("No rows yet to combine.")
        return
    combined = pd.concat(all_dfs, ignore_index=True)

    before = len(combined)
    combined = combined.drop_duplicates(
        subset=['TARİH', 'TUTAR', 'Bakiye', 'AÇIKLAMA'],
        keep='first',
    ).reset_index(drop=True)
    after = len(combined)
    if before != after:
        st.info(f"Removed {before - after} duplicate row(s) across uploads.")

    st.write(f"**Total rows: {len(combined):,}**")
    st.dataframe(combined.head(50), use_container_width=True)

    dl1, dl2 = st.columns(2)
    dl1.download_button(
        "⬇ combined.xlsx",
        data=df_to_xlsx_bytes(combined),
        file_name="combined_bank_statements.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    dl2.download_button(
        "⬇ combined.csv",
        data=df_to_csv_bytes(combined),
        file_name="combined_bank_statements.csv",
        mime="text/csv",
    )


if __name__ == '__main__':
    main()
