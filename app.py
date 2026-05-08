"""
Bank Statement -> CSV/Excel Extractor (Accountant-Grade)
========================================================

Supported banks: Halkbank (V1 + V2 layouts), Akbank, Ziraat (vector + scanned/OCR).
Unknown banks: generic fallback parser with auto-detected row patterns.

Output columns: TARIH | Saat | TUTAR | Bakiye | ACIKLAMA | DEKONT

DEPLOYMENT
----------
This app is deployment-friendly. It runs on:

  * Streamlit Community Cloud — push the repo with `requirements.txt`
    and `packages.txt` (lists tesseract-ocr, tesseract-ocr-tur, poppler-utils).
    No code changes needed.
  * Linux server / Docker — `apt install tesseract-ocr tesseract-ocr-tur
    poppler-utils` then `pip install -r requirements.txt`.
  * macOS local — `brew install tesseract tesseract-lang poppler`
    then `pip install -r requirements.txt`.
  * Windows local — install tesseract from
    https://github.com/UB-Mannheim/tesseract/wiki (default path
    C:\\Program Files\\Tesseract-OCR) and poppler from
    https://github.com/oschwartz10612/poppler-windows/releases (default
    path C:\\Program Files\\poppler\\Library\\bin). Override either path
    via env vars TESSERACT_CMD and POPPLER_PATH if installed elsewhere.

If the OCR libraries / system binaries are missing, the app still works
for vector PDFs (Halkbank, Akbank, vector-mode Ziraat). Only scanned
PDFs need OCR.
"""

from __future__ import annotations

import io
import os
import re
import sys
import zipfile
from dataclasses import dataclass, field
from datetime import date as _date
from decimal import Decimal, InvalidOperation
from typing import Callable, Optional

import pandas as pd
import pdfplumber
import streamlit as st

# -----------------------------------------------------------------------------
# Cross-platform OCR setup
# -----------------------------------------------------------------------------
# pytesseract + pdf2image are optional (only needed for scanned PDFs).
# On Linux/macOS/Streamlit-Cloud, the tesseract & poppler binaries live on the
# system PATH (installed via apt/brew/packages.txt) and `pdf2image` accepts
# `poppler_path=None` to mean "use PATH".
# On Windows the binaries usually aren't on PATH, so we look in default install
# locations and let env vars override.

OCR_AVAILABLE: bool = False
OCR_IMPORT_ERROR: Optional[str] = None
POPPLER_PATH: Optional[str] = None  # None = use system PATH

try:
    import pytesseract  # type: ignore
    from pdf2image import convert_from_bytes  # type: ignore

    if sys.platform == 'win32':
        # Tesseract: env var TESSERACT_CMD wins; otherwise default install path.
        _tess_cmd = os.environ.get('TESSERACT_CMD') or \
            r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        if os.path.exists(_tess_cmd):
            pytesseract.pytesseract.tesseract_cmd = _tess_cmd
        # Poppler: env var POPPLER_PATH wins; otherwise default install path.
        _poppler = os.environ.get('POPPLER_PATH') or \
            r'C:\Program Files\poppler\Library\bin'
        if os.path.exists(_poppler):
            POPPLER_PATH = _poppler

    # Confirm the tesseract binary is actually callable.
    pytesseract.get_tesseract_version()
    OCR_AVAILABLE = True
except ImportError as _exc:
    OCR_IMPORT_ERROR = (
        f"Python OCR libraries not installed: {_exc}. "
        "Install with: pip install pytesseract pdf2image"
    )
except Exception as _exc:
    OCR_IMPORT_ERROR = (
        f"OCR libraries imported but tesseract binary not found: {_exc}. "
        "Install tesseract via your system package manager (apt/brew/installer) "
        "and ensure the 'tur' (Turkish) language pack is included."
    )


# =============================================================================
# 1. Number parsing
# =============================================================================

def parse_tr_decimal(s: str) -> Decimal:
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
    parser_warnings: list[str] = field(default_factory=list)  # actionable
    parser_info: list[str] = field(default_factory=list)      # informational


def _fix_ocr_date(s: str) -> str:
    """Validate DD.MM.YYYY (or DD-MM-YYYY) and try common OCR digit
    substitutions if invalid.

    OCR commonly misreads digits in similar-looking fonts:
      5 <-> 3 (curve at top of 5 mistaken for 3, or vice versa)
      8 <-> 0 / 6 <-> 0 / 9 <-> 4 / 1 <-> 7

    Returns a normalized 'DD.MM.YYYY' string, or '' if unfixable.
    """
    if not s:
        return ''
    m = re.match(r'^\s*(\d{2})[.\-/\s](\d{2})[.\-/\s](\d{4})\s*$', s)
    if not m:
        return ''
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))

    def _valid(dd: int, mm: int, yy: int) -> bool:
        try:
            _date(yy, mm, dd)
            return True
        except (ValueError, TypeError):
            return False

    if _valid(d, mo, y):
        return f'{d:02d}.{mo:02d}.{y}'

    # OCR digit confusion table — only substitutions known to occur.
    confusions = {
        '0': ['8', '6', '9'], '1': ['7', '4'], '2': ['7'],
        '3': ['5', '8'],      '4': ['9', '1'], '5': ['3', '6', '8'],
        '6': ['0', '5', '8'], '7': ['1', '2'], '8': ['0', '3', '5', '6'],
        '9': ['4', '0'],
    }
    # Try fixing the day digits (most common: '50' read for '30').
    digits = f'{d:02d}'
    for pos, ch in enumerate(digits):
        for repl in confusions.get(ch, []):
            cand_d = int(digits[:pos] + repl + digits[pos + 1:])
            if _valid(cand_d, mo, y):
                return f'{cand_d:02d}.{mo:02d}.{y}'
    # Then try the month digits.
    digits = f'{mo:02d}'
    for pos, ch in enumerate(digits):
        for repl in confusions.get(ch, []):
            cand_mo = int(digits[:pos] + repl + digits[pos + 1:])
            if _valid(d, cand_mo, y):
                return f'{d:02d}.{cand_mo:02d}.{y}'
    return ''


# =============================================================================
# 3. Bank detection
# =============================================================================

def detect_bank(full_text: str) -> Optional[str]:
    """Detect issuing bank by looking at the statement header (first ~500 chars
    where the bank logo / brand sits). This avoids false positives when one
    bank's name appears inside another bank's transaction descriptions
    (e.g. a Ziraat statement that mentions 'Türkiye Garanti Bankası' as a
    FAST-transfer destination)."""
    def _check(window: str) -> Optional[str]:
        # Order doesn't matter much within the header — only one bank's brand
        # should appear there. We still keep it stable for predictability.
        if 'ZİRAAT' in window or 'ZIRAAT BANKASI' in window or 'ZIRAAT' in window:
            return 'ZIRAAT'
        if 'AKBANK' in window or 'AKPOS' in window:
            return 'AKBANK'
        if 'HALKBANK' in window or 'TÜRKIYE HALK BANKASI' in window or 'TURKIYE HALK BANKASI' in window:
            return 'HALKBANK'
        if 'GARANTİ' in window or 'GARANTI BANKASI' in window:
            return 'GARANTI'
        if 'İŞ BANKASI' in window or 'IS BANKASI' in window or 'TÜRKİYE İŞ BANKASI' in window:
            return 'ISBANK'
        if 'YAPI VE KREDİ' in window or 'YAPI KREDI' in window:
            return 'YAPIKREDI'
        return None

    upper = full_text.upper()
    # Header-only detection first — strongly preferred.
    head_hit = _check(upper[:500])
    if head_hit:
        return head_hit
    # Fallback: scan a wider window, but accept that this is less reliable.
    return _check(upper[:5000])


# =============================================================================
# 4. PDF text extraction with OCR fallback
# =============================================================================

def _is_pdf_scanned(pdf_bytes: bytes) -> bool:
    """A PDF is considered scanned if pdfplumber finds zero text characters
    across all pages (only images / no embedded text layer)."""
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            total_chars = sum(len(p.chars) for p in pdf.pages)
        return total_chars == 0
    except Exception:
        return False


def extract_pdf_text(pdf_bytes: bytes) -> str:
    """Vector-text extraction. Returns '' if the PDF is image-only."""
    parts: list[str] = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or '')
    return '\n'.join(parts)


def _ocr_pages(pdf_bytes: bytes, dpi: int = 400, lang: str = 'tur') -> list[dict]:
    """OCR every page and return list of {text, words, image_size}.
    `words` is the per-token positional data (used by Ziraat OCR parser)."""
    if not OCR_AVAILABLE:
        raise RuntimeError(
            "OCR libraries not installed. To process scanned PDFs install:\n"
            "  pip install pytesseract pdf2image\n"
            "  apt-get install tesseract-ocr tesseract-ocr-tur poppler-utils\n"
            f"Original import error: {OCR_IMPORT_ERROR}"
        )
    images = convert_from_bytes(pdf_bytes, dpi=dpi, poppler_path=POPPLER_PATH)
    results = []
    for img in images:
        data = pytesseract.image_to_data(
            img, lang=lang, output_type=pytesseract.Output.DICT, config='--psm 6'
        )
        # Group OCR words into lines by block/par/line, sorted by y then x
        lines: dict = {}
        for i, txt in enumerate(data['text']):
            if not txt.strip():
                continue
            key = (data['block_num'][i], data['par_num'][i], data['line_num'][i])
            lines.setdefault(key, []).append((data['left'][i], data['top'][i], txt))
        sorted_lines = sorted(
            lines.items(), key=lambda kv: min(t for _, t, _ in kv[1])
        )
        line_objs = []
        text_lines = []
        for _key, words in sorted_lines:
            words.sort()  # by left x
            tokens = [t for _, _, t in words]
            line_objs.append({'tokens': tokens, 'top': min(t for _, t, _ in words)})
            text_lines.append(' '.join(tokens))
        results.append({
            'text': '\n'.join(text_lines),
            'lines': line_objs,
            'size': img.size,
        })
    return results


# =============================================================================
# 5. Halkbank parser  (unchanged from original)
# =============================================================================

_HB_SKIP_CONT_PREFIXES = (
    'Müşteri', 'Hesap', 'TCKN', 'IBAN', 'Şube', 'Döviz', 'Üretim', 'Dönemi',
    'Bakiye Bilgi', 'Bloke', 'Kullanıl', 'Toplam Kredi',
    'Türkiye Halk', 'yerine kullanıl', 'Uyuşmazlık',
    'MÜŞTERİ', 'Dönem',
)
_HB_SKIP_EXACT = {'HESAP ÖZETİ'}

_HB_NUM = r'-?[\d.]+,\d{2}'
_HB_ROW_V1 = re.compile(rf'^(\d{{2}}-\d{{2}}-\d{{4}})\s+({_HB_NUM})\s+({_HB_NUM})\s+(.*)$')

_HB_NUM_UNSIGNED = r'[\d.]+,\d{2}'
_HB_ROW_V2 = re.compile(
    rf'^(\d{{2}}\.\d{{2}}\.\d{{4}})\s+'
    rf'\d{{2}}\.\d{{2}}\.\d{{4}}\s+'
    rf'({_HB_NUM_UNSIGNED})\s+([+-])\s+'
    rf'({_HB_NUM_UNSIGNED})\s+([+-])\s+'
    rf'(.*)$'
)

_HB_DEKONT_RE = re.compile(r'/(\d{10,})\s*$')


def _hb_extract_dekont(description: str) -> str:
    m = _HB_DEKONT_RE.search(description)
    return m.group(1) if m else ''


def _hb_strip_right_column(line: str) -> str:
    right_anchors = (
        r'\s+(?:Üretim\s+Zamanı|Dönemi|Hesap\s+Bakiyesi|Bloke\s+Bakiyesi|'
        r'Kullanılabilir(?:\s+\w+)*|Toplam\s+Kredi|Bakiye\s+Bilgileriniz|'
        r'Hesap\s+Özeti)\s*:'
    )
    return re.split(right_anchors, line, maxsplit=1)[0].rstrip()


def _hb_extract_customer_name(text: str) -> str:
    lines = text.split('\n')
    for line in lines:
        m = re.search(r'Müşteri Adı\s*/\s*Ünvanı\s*:\s*(.+)', line)
        if m:
            name = m.group(1).strip()
            if name:
                return name
    label_idx = next(
        (i for i, l in enumerate(lines) if 'Müşteri Adı / Ünvanı' in l), None)
    if label_idx is None:
        return ''
    next_label_re = re.compile(r'^(TCKN|Hesap|Bakiye|Şube|Döviz|IBAN|Bloke|Toplam)\b')
    parts: list[str] = []
    if label_idx > 0:
        cand = _hb_strip_right_column(lines[label_idx - 1]).strip()
        if cand and not cand.startswith(('Müşteri', 'Hesap', 'TCKN', 'IBAN',
                                         'Şube', 'Döviz', 'Bakiye', 'Bloke',
                                         'Kullanıl', 'Toplam', 'Üretim', 'Dönemi',
                                         'MÜŞTERİ', 'Dönem')):
            parts.append(cand)
    for j in range(label_idx + 1, min(label_idx + 5, len(lines))):
        cand = _hb_strip_right_column(lines[j]).strip()
        if not cand:
            continue
        if next_label_re.match(cand):
            break
        parts.append(cand)
    return ' '.join(parts).strip()


def _hb_extract_metadata(pdf_bytes: bytes) -> StatementMetadata:
    md = StatementMetadata(bank='HALKBANK')
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text = pdf.pages[0].extract_text() or ''
    except Exception:
        return md
    cleaned_lines = [_hb_strip_right_column(l) for l in text.split('\n')]
    cleaned = '\n'.join(cleaned_lines)

    def grab(pattern: str, source: str = cleaned) -> Optional[str]:
        m = re.search(pattern, source)
        return m.group(1).strip() if m else None

    md.customer_no = grab(r'Müşteri Numarası\s*:\s*(\S+)') or grab(r'Müşteri No\s*:\s*(\S+)') or ''
    md.customer_name = _hb_extract_customer_name(text)
    md.account_no = grab(r'Hesap No\s*:\s*(\S+)') or ''
    md.iban = grab(r'IBAN\s*:\s*(\S+)') or ''
    md.branch = grab(r'Şube Kodu / Adı\s*:\s*([^\n]+)') or ''
    if not md.branch:
        code = grab(r'Şube Kodu\s*:\s*(\S+)') or ''
        name = grab(r'Şube Adı\s*:\s*([^\n]+)') or ''
        if code or name:
            md.branch = f"{code} / {name}".strip(' /')
    md.currency = grab(r'Döviz Cinsi\s*:\s*(\S+)') or ''
    period = (
        grab(r'Dönemi\s*:\s*([\d./]+\s*-\s*[\d./]+)', source=text) or
        grab(r'Dönem\s*\(Tarih Aralığı\)\s*:\s*([\d./]+\s*-\s*[\d./]+)', source=text) or
        grab(r'Dönem[^:]*:\s*([\d./]+\s*-\s*[\d./]+)', source=text)
    )
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


def _normalize_date(date_str: str) -> str:
    return date_str.replace('.', '-')


def _hb_parse_transactions(full_text: str) -> tuple[list[TransactionRow], list[str]]:
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
        m = _HB_ROW_V1.match(line)
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
            current = TransactionRow(date=date_str, amount=amount, balance=balance,
                                     description=desc.strip())
            continue
        m2 = _HB_ROW_V2.match(line)
        if m2:
            if current is not None:
                rows.append(current)
            date_str, amount_str, amount_sign, balance_str, balance_sign, desc = m2.groups()
            try:
                amount = parse_tr_decimal(amount_str)
                balance = parse_tr_decimal(balance_str)
            except InvalidOperation as exc:
                warnings.append(f"Skipped malformed row at {date_str}: {exc}")
                current = None
                continue
            if amount_sign == '-':
                amount = -amount
            if balance_sign == '-':
                balance = -balance
            current = TransactionRow(date=_normalize_date(date_str), amount=amount,
                                     balance=balance, description=desc.strip())
            continue
        if current is None:
            continue
        if stripped.startswith(_HB_SKIP_CONT_PREFIXES):
            continue
        if stripped.startswith('Ekstrenize') or stripped.startswith('Türkiye Halk Bankası'):
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
    return StatementResult(source_filename=source_filename, metadata=metadata,
                           rows=rows, parser_warnings=warnings)


# =============================================================================
# 6. Akbank parser  (unchanged from original)
# =============================================================================

_AK_NUM = r'-?[\d.]+,\d{2}'
_AK_ROW = re.compile(
    rf'^(\d{{2}}\.\d{{2}}\.\d{{4}})\s+'
    rf'(\d{{2}}:\d{{2}})\s+'
    rf'\d{{2}}\.\d{{2}}\.\d{{4}}\s+'
    rf'(\d+)\s+'
    rf'({_AK_NUM})\s+'
    rf'({_AK_NUM})\s+'
    rf'([BA])\s+'
    rf'(.*)$'
)
_AK_SKIP_LINE_STARTS = (
    'HESAP HAREKET', 'Tarih :', 'Düzenleyen', 'Hesap Şube', 'Hesap No',
    'Döviz', 'TARİH', 'AKBANK',
)


def _ak_extract_metadata(pdf_bytes: bytes) -> StatementMetadata:
    md = StatementMetadata(bank='AKBANK')
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text = pdf.pages[0].extract_text() or ''
    except Exception:
        return md

    def grab(pattern: str) -> Optional[str]:
        m = re.search(pattern, text)
        return m.group(1).strip() if m else None

    md.customer_name = grab(r'Ad Soyad\s*:\s*([^\n]+)') or ''
    md.account_no = grab(r'Hesap No\s*:\s*(\S+)') or ''
    md.iban = grab(r'IBAN\s*:\s*(\S+)') or ''
    branch_m = re.search(r'Hesap Şube\s*:\s*(.+?)(?:\s{2,}|\s+(?:IBAN|Ad Soyad|Tarih))', text)
    md.branch = branch_m.group(1).strip() if branch_m else (grab(r'Hesap Şube\s*:\s*([^\n]+)') or '')
    curr = grab(r'Döviz\s*:\s*([^\n]+)')
    if curr:
        parts = curr.split('-')
        md.currency = parts[-1].strip() if len(parts) > 1 else curr.strip()
    period = grab(r'Tarih Aralığı\s*:\s*([\d./]+\s*-\s*[\d./]+)')
    if period and '-' in period:
        a, b = period.split('-', 1)
        md.period_start, md.period_end = a.strip(), b.strip()
    return md


def _ak_parse_transactions(full_text: str) -> tuple[list[TransactionRow], list[str]]:
    rows: list[TransactionRow] = []
    warnings: list[str] = []
    current: Optional[TransactionRow] = None
    for raw in full_text.split('\n'):
        line = raw.rstrip('\r').rstrip()
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith(_AK_SKIP_LINE_STARTS):
            continue
        if re.match(r'^\d+\s*/\s*\d+$', stripped):
            continue
        m = _AK_ROW.match(line)
        if m:
            if current is not None:
                rows.append(current)
            date_str, time_str, fis_no, amount_str, balance_str, ba, desc = m.groups()
            try:
                amount = parse_tr_decimal(amount_str)
                balance = parse_tr_decimal(balance_str)
            except InvalidOperation as exc:
                warnings.append(f"Skipped malformed row at {date_str}: {exc}")
                current = None
                continue
            amount = -abs(amount) if ba == 'B' else abs(amount)
            current = TransactionRow(
                date=_normalize_date(date_str), time=time_str,
                amount=amount, balance=balance,
                description=desc.strip(), receipt=fis_no,
            )
            continue
        if current is None:
            continue
        current.description += ' ' + stripped
    if current is not None:
        rows.append(current)
    rows.reverse()
    return rows, warnings


def parse_akbank(pdf_bytes: bytes, full_text: str, source_filename: str) -> StatementResult:
    metadata = _ak_extract_metadata(pdf_bytes)
    rows, warnings = _ak_parse_transactions(full_text)
    return StatementResult(source_filename=source_filename, metadata=metadata,
                           rows=rows, parser_warnings=warnings)


# =============================================================================
# 7. Ziraat Bankası parser — supports BOTH vector PDFs and scanned/OCR PDFs
# =============================================================================

_ZR_DATE_RE = re.compile(r'^\d{2}\.\d{2}\.\d{4}$')

# Strict Turkish-decimal money pattern: optional sign, optional thousands groups,
# mandatory comma followed by exactly 2 digits at end.
_TR_MONEY_RE = re.compile(r'^-?\d{1,3}(?:\.\d{3})*,\d{2}$|^-?\d+,\d{2}$')

# Header / footer / metadata line prefixes to skip in OCR text
_ZR_OCR_SKIP_PREFIXES = (
    'Tarih', 'Fiş', 'Sayın', 'Müşteri', 'Müşter', 'Adres', 'HENDESE',
    'Şube', 'IBAN', 'Dönem', 'Day', 'Döviz', 'Borç', 'Alacak',
    'Taraflar', 'Merkez', 'Ticaret Sicil', 'www.', 'ğ', 'EEE',
)


def _normalize_ocr_number(s: str) -> str:
    """Fix common OCR mistakes in Turkish-formatted numbers.

    Turkish money: `1.234.567,89` (period = thousands separator, comma = decimal).
    OCR often reads the decimal comma as a period, producing e.g. `75.001.60`
    or `37.51`. Rule: if the number has no comma but ends in `.XX`, that final
    period IS the decimal — convert it to a comma.
    """
    s = s.strip()
    if not s or ',' in s:
        return s
    if re.search(r'\.\d{2}$', s):
        idx = s.rfind('.')
        s = s[:idx] + ',' + s[idx + 1:]
    return s


def _try_repair_no_separator(tok: str) -> Optional[str]:
    """Repair tokens where OCR dropped the decimal entirely.

    e.g. `3751` for `37,51`. Only attempted on pure digit strings of length ≥ 3.
    The repaired number is later validated by the balance chain — if it makes
    the chain consistent we keep it (and surface a warning); if not, the user
    sees a clear flag instead of a silent miscount.
    """
    s = tok.strip()
    sign = ''
    if s.startswith('-'):
        sign = '-'
        s = s[1:]
    if not s.isdigit() or len(s) < 3 or len(s) > 12:
        return None
    return f"{sign}{s[:-2]},{s[-2:]}"


def _is_money(tok: str) -> bool:
    return bool(_TR_MONEY_RE.match(tok))


def _ziraat_extract_metadata(full_text: str) -> StatementMetadata:
    md = StatementMetadata(bank='ZİRAAT')

    def grab(pattern: str) -> Optional[str]:
        m = re.search(pattern, full_text, re.IGNORECASE)
        return m.group(1).strip() if m else None

    cust_raw = grab(r'Say[ıi]n\s*:?\s*([^\n]+)') or ''
    md.customer_name = re.split(r'\s{3,}|\s+[Şş]ube', cust_raw)[0].strip()
    # Strip any leading "O:" or "0:" left over from OCR'd "Sayın :" prefix
    md.customer_name = re.sub(r'^[O0]\s*:\s*', '', md.customer_name).strip()
    md.account_no = grab(r'M[üu][şs]ter[i/]+Hesap No\s*:\s*(\S+)') or grab(r'Hesap No\s*:\s*(\S+)') or ''
    md.iban = grab(r'IBAN\s*:\s*(\S+)') or ''
    md.branch = grab(r'[Şş]ube Kodu\s*:\s*([^\n\t]+)') or ''
    md.currency = grab(r'D[öo]viz Cinsi\s*:\s*(\S+)') or ''
    # Period: OCR often misreads ':' as '1', mangles digits, and turns inner
    # date periods into spaces. Find the first "two-dates-with-hyphen"
    # pattern within ~30 chars after 'Dönem'.
    pm = re.search(
        r'D[öo]nem.{0,30}?(\d{2}[.\-\s]\d{2}[.\-\s]\d{4})\s*[-–]\s*(\d{2}[.\-\s]\d{2}[.\-\s]\d{4})',
        full_text,
        re.DOTALL,
    )
    if pm:
        # Validate & repair OCR digit confusions (e.g. day '50' -> '30').
        md.period_start = _fix_ocr_date(pm.group(1).strip()) or ''
        md.period_end   = _fix_ocr_date(pm.group(2).strip()) or ''
    return md


def _ziraat_parse_vector(pdf_bytes: bytes) -> tuple[list[TransactionRow], list[str]]:
    """Parse Ziraat statement when PDF has an embedded text layer (uses tables)."""
    rows: list[TransactionRow] = []
    warnings: list[str] = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                for table in (page.extract_tables() or []):
                    for trow in table:
                        if not trow or len(trow) < 5:
                            continue
                        date_val = (trow[0] or '').strip()
                        fis_no = (trow[1] or '').strip()
                        desc = ' '.join((trow[2] or '').split())
                        tutar_s = (trow[3] or '').strip()
                        bakiye_s = (trow[4] or '').strip()
                        if not _ZR_DATE_RE.match(date_val):
                            continue
                        if not tutar_s or not bakiye_s:
                            continue
                        try:
                            amount = parse_tr_decimal(tutar_s)
                            balance = parse_tr_decimal(bakiye_s)
                        except (InvalidOperation, Exception) as exc:
                            warnings.append(f"Skipped Ziraat row at {date_val}: {exc}")
                            continue
                        rows.append(TransactionRow(
                            date=_normalize_date(date_val),
                            amount=amount, balance=balance,
                            description=desc, receipt=fis_no,
                        ))
    except Exception as exc:
        warnings.append(f"Ziraat table extraction error: {exc}")
    rows.reverse()  # newest-first → chronological
    return rows, warnings


def _ziraat_normalize_fis(fis: str) -> str:
    """Clean OCR artifacts from Ziraat Fiş No.

    Legitimate Ziraat fiş numbers are 6 chars: 'F' + 5 digits (e.g. F12750)
    or alphanumeric (e.g. FBWW24). OCR commonly produces:
      - leading junk: '(OF12750', 'oOF12679', '(F10836'
      - I/1 and O/0 confusion: 'FI7539' (real F17539), 'FO9116' (real F09116)
      - spurious extra char: 'FI12679' (real F12679, I is junk insert)

    Strategy: strip leading non-F junk; if result starts F[IlO]:
      - len 6: convert position-1 char (I/l->1, O->0) -- preserves length
      - len 7: drop the spurious char -- restores canonical 6-char length
    """
    if not fis:
        return fis
    fis = re.sub(r'^[(\s]+', '', fis)
    fis = re.sub(r'^[oO0]+(?=F)', '', fis)  # strip leading O/0/o ONLY if F follows
    fis = re.sub(r'[)\s]+$', '', fis)
    m = re.match(r'^(F)([IlO])(\d{4,5})$', fis)
    if m:
        body = m.group(3)
        if len(fis) == 6:
            fixed = {'I': '1', 'l': '1', 'O': '0'}[m.group(2)]
            fis = m.group(1) + fixed + body
        elif len(fis) == 7:
            fis = m.group(1) + body
    return fis


def _ziraat_parse_ocr(pdf_bytes: bytes) -> tuple[list[TransactionRow], list[str], str]:
    """Parse Ziraat statement when PDF is image-only (uses OCR with positional data).

    Returns (rows, warnings, full_ocr_text). The OCR text is also used to extract
    metadata (customer name, IBAN, period, etc.).
    """
    rows: list[TransactionRow] = []
    warnings: list[str] = []
    full_text_parts: list[str] = []

    pages = _ocr_pages(pdf_bytes, dpi=400, lang='tur')
    for page_data in pages:
        full_text_parts.append(page_data['text'])

        pending_desc: list[str] = []
        for line in page_data['lines']:
            tokens = line['tokens']
            if not tokens:
                continue

            if _ZR_DATE_RE.match(tokens[0]):
                date = tokens[0]
                # Find money tokens AFTER normalization & repair
                normalized: list[tuple[int, str, bool]] = []  # (idx, value, was_repaired)
                repair_log: list[tuple[str, str]] = []
                for i, t in enumerate(tokens):
                    if i == 0:  # date itself — never treat as money
                        normalized.append((i, t, False))
                        continue
                    n = _normalize_ocr_number(t)
                    if _is_money(n):
                        normalized.append((i, n, n != t))
                    else:
                        repaired = _try_repair_no_separator(t)
                        if repaired and _is_money(repaired):
                            normalized.append((i, repaired, True))
                            repair_log.append((t, repaired))
                        else:
                            normalized.append((i, t, False))

                money_idxs = [(i, v, r) for i, v, r in normalized
                              if i > 0 and _is_money(v)]
                if len(money_idxs) >= 2:
                    tutar_idx, tutar_s, tutar_repaired = money_idxs[-2]
                    _bal_idx, bakiye_s, bakiye_repaired = money_idxs[-1]
                    # Fiş No: first token after the date that doesn't look like OCR
                    # junk. OCR often inserts noise tokens like "(o" / "(O" / "o"
                    # before the real fiş — skip them.
                    fis = ''
                    junk_re = re.compile(r'^[(\s]*[oO0]?[)\s]*$')
                    for ti in range(1, tutar_idx):
                        cand = re.sub(r'^[(\s]*[oO0]?\s*', '', tokens[ti]).strip()
                        cand = re.sub(r'\s*[)\s]*$', '', cand).strip()
                        if not cand or junk_re.match(cand):
                            continue
                        if len(cand) >= 3:
                            fis = _ziraat_normalize_fis(cand)
                            break
                    desc_inline = ' '.join(
                        tokens[ti + 1:tutar_idx]
                    ).strip() if fis else ' '.join(tokens[1:tutar_idx]).strip()
                    full_desc_parts = pending_desc + (
                        [desc_inline] if desc_inline else []
                    )
                    full_desc = ' '.join(full_desc_parts).strip()
                    pending_desc = []
                    try:
                        amount = parse_tr_decimal(tutar_s)
                        balance = parse_tr_decimal(bakiye_s)
                    except (InvalidOperation, Exception) as exc:
                        warnings.append(f"Skipped Ziraat OCR row at {date}: {exc}")
                        continue
                    rows.append(TransactionRow(
                        date=_normalize_date(date),
                        amount=amount, balance=balance,
                        description=full_desc, receipt=fis,
                    ))
                    if tutar_repaired or bakiye_repaired:
                        for orig, fixed in repair_log:
                            if fixed in (tutar_s, bakiye_s):
                                warnings.append(
                                    f"OCR repair on {date}: '{orig}' → '{fixed}'. "
                                    f"Verify against the source PDF."
                                )
                else:
                    warnings.append(
                        f"Ziraat OCR row at {date}: only {len(money_idxs)} valid "
                        f"number(s), expected 2. Tokens: {tokens}. ROW SKIPPED."
                    )
                    pending_desc = []
            else:
                # Not a transaction line — could be description continuation OR
                # header/footer/boilerplate. Skip known boilerplate; otherwise
                # buffer as a possible description continuation.
                line_text = ' '.join(tokens)
                if any(line_text.startswith(p) for p in _ZR_OCR_SKIP_PREFIXES):
                    continue
                pending_desc.append(line_text)

    rows.reverse()  # PDF lists newest-first → chronological
    return rows, warnings, '\n'.join(full_text_parts)


def parse_ziraat(pdf_bytes: bytes, full_text: str, source_filename: str) -> StatementResult:
    """Ziraat parser. Picks vector or OCR path based on whether the PDF
    has an embedded text layer."""
    is_scanned = _is_pdf_scanned(pdf_bytes)
    if is_scanned:
        if not OCR_AVAILABLE:
            md = StatementMetadata(bank='ZİRAAT')
            return StatementResult(
                source_filename=source_filename, metadata=md, rows=[],
                parser_warnings=[
                    "This Ziraat PDF is image-only (scanned). OCR is required "
                    "but pytesseract / pdf2image / tesseract are not installed. "
                    f"Details: {OCR_IMPORT_ERROR or 'see deployment notes at top of file.'}"
                ],
            )
        rows, raw_msgs, ocr_text = _ziraat_parse_ocr(pdf_bytes)
        metadata = _ziraat_extract_metadata(ocr_text)
        # Split: OCR-detected note + repair notes are informational; everything
        # else (skipped rows, parse errors) is actionable.
        info: list[str] = [
            "Scanned/image-only PDF detected — used OCR (Turkish). "
            "Verify amounts against the source PDF."
        ]
        warnings: list[str] = []
        for m in raw_msgs:
            (info if m.startswith('OCR repair') else warnings).append(m)
        return StatementResult(
            source_filename=source_filename, metadata=metadata,
            rows=rows, parser_warnings=warnings, parser_info=info,
        )
    rows, warnings = _ziraat_parse_vector(pdf_bytes)
    metadata = _ziraat_extract_metadata(full_text)
    return StatementResult(source_filename=source_filename, metadata=metadata,
                           rows=rows, parser_warnings=warnings)


# =============================================================================
# 8. Generic fallback parser  (unchanged from original)
# =============================================================================

_GEN_DATE = r'\d{2}[.\-/]\d{2}[.\-/]\d{4}'
_GEN_TIME = r'\d{2}:\d{2}'
_GEN_NUM = r'-?[\d.]+,\d{2}'

_GEN_PATTERNS = [
    ('akbank_like', re.compile(
        rf'^({_GEN_DATE})\s+({_GEN_TIME})\s+{_GEN_DATE}\s+\d+\s+'
        rf'({_GEN_NUM})\s+({_GEN_NUM})\s+([BA])\s+(.*)$'
    )),
    ('halkbank_v2_like', re.compile(
        rf'^({_GEN_DATE})\s+{_GEN_DATE}\s+'
        rf'([\d.]+,\d{{2}})\s+([+-])\s+'
        rf'([\d.]+,\d{{2}})\s+([+-])\s+(.*)$'
    )),
    ('halkbank_v1_like', re.compile(
        rf'^({_GEN_DATE})\s+({_GEN_NUM})\s+({_GEN_NUM})\s+(.*)$'
    )),
]
_GEN_SKIP_STARTS = (
    'HESAP', 'TARİH', 'TARIH', 'İŞLEM', 'SAYFA', 'AKBANK', 'HALKBANK',
    'GARANTİ', 'ZİRAAT', 'MÜŞTERİ', 'MUSTERI', 'IBAN', 'DÖVIZ', 'DOVIZ',
    'DÜZENLEYEN', 'ŞUBE', 'SUBE', 'DÖNEM', 'DONEM', 'TCKN', 'Müşteri',
    'Hesap', 'Şube', 'Döviz', 'Dönem', 'Bakiye', 'Bloke', 'Toplam',
    'Türkiye', 'Ekstrenize', 'yerine',
)


def _generic_parse_transactions(full_text: str) -> tuple[list[TransactionRow], list[str]]:
    rows: list[TransactionRow] = []
    warnings: list[str] = []
    current: Optional[TransactionRow] = None
    for raw in full_text.split('\n'):
        line = raw.rstrip('\r').rstrip()
        stripped = line.strip()
        if not stripped:
            continue
        matched = False
        for pat_name, pat_re in _GEN_PATTERNS:
            m = pat_re.match(line)
            if not m:
                continue
            if current is not None:
                rows.append(current)
            groups = m.groups()
            if pat_name == 'akbank_like':
                date_str, time_str, amount_str, balance_str, ba, desc = groups
                try:
                    amount = parse_tr_decimal(amount_str)
                    balance = parse_tr_decimal(balance_str)
                except InvalidOperation:
                    current = None; matched = True; break
                amount = -abs(amount) if ba == 'B' else abs(amount)
                current = TransactionRow(date=_normalize_date(date_str), time=time_str,
                                         amount=amount, balance=balance, description=desc.strip())
            elif pat_name == 'halkbank_v2_like':
                date_str, amount_str, amount_sign, balance_str, balance_sign, desc = groups
                try:
                    amount = parse_tr_decimal(amount_str)
                    balance = parse_tr_decimal(balance_str)
                except InvalidOperation:
                    current = None; matched = True; break
                if amount_sign == '-': amount = -amount
                if balance_sign == '-': balance = -balance
                current = TransactionRow(date=_normalize_date(date_str), amount=amount,
                                         balance=balance, description=desc.strip())
            elif pat_name == 'halkbank_v1_like':
                date_str, amount_str, balance_str, desc = groups
                try:
                    amount = parse_tr_decimal(amount_str)
                    balance = parse_tr_decimal(balance_str)
                except InvalidOperation:
                    current = None; matched = True; break
                current = TransactionRow(date=_normalize_date(date_str), amount=amount,
                                         balance=balance, description=desc.strip())
            matched = True
            break
        if not matched and current is not None:
            if any(stripped.upper().startswith(s.upper()) for s in _GEN_SKIP_STARTS):
                continue
            if re.match(r'^\d+\s*/\s*\d+$', stripped):
                continue
            current.description += ' ' + stripped
    if current is not None:
        rows.append(current)
    return rows, warnings


def parse_generic(pdf_bytes: bytes, full_text: str, source_filename: str) -> StatementResult:
    md = StatementMetadata(bank='UNKNOWN')

    def grab(pattern: str) -> Optional[str]:
        m = re.search(pattern, full_text[:3000])
        return m.group(1).strip() if m else None

    md.customer_name = grab(r'(?:Ad Soyad|Müşteri Adı|Ünvanı?)\s*:\s*([^\n]+)') or ''
    md.account_no = grab(r'Hesap No\s*:\s*(\S+)') or ''
    md.iban = grab(r'IBAN\s*:\s*(\S+)') or ''
    period = grab(r'(?:Tarih Aralığı|Dönemi?|Dönem)\s*[:(]\s*([\d./]+\s*-\s*[\d./]+)')
    if period and '-' in period:
        a, b = period.split('-', 1)
        md.period_start, md.period_end = a.strip(), b.strip()
    rows, warnings = _generic_parse_transactions(full_text)
    warnings.insert(0,
        "⚠ Used generic fallback parser — bank not recognized. "
        "Results may be incomplete. Please verify carefully."
    )
    if len(rows) >= 2:
        try:
            first_d = pd.to_datetime(rows[0].date, format='%d-%m-%Y', errors='coerce')
            last_d = pd.to_datetime(rows[-1].date, format='%d-%m-%Y', errors='coerce')
            if first_d is not None and last_d is not None and first_d > last_d:
                rows.reverse()
                warnings.append("Rows appeared reverse-chronological; reversed to oldest-first.")
        except Exception:
            pass
    return StatementResult(source_filename=source_filename, metadata=md,
                           rows=rows, parser_warnings=warnings)


# =============================================================================
# 9. Parser registry
# =============================================================================

PARSERS: dict[str, Callable[[bytes, str, str], StatementResult]] = {
    'HALKBANK': parse_halkbank,
    'AKBANK': parse_akbank,
    'ZIRAAT': parse_ziraat,
}


# =============================================================================
# 10. Validation
# =============================================================================

def validate_balance_chain(rows: list[TransactionRow]) -> list[BalanceIssue]:
    issues: list[BalanceIssue] = []
    for i in range(1, len(rows)):
        expected = rows[i - 1].balance + rows[i].amount
        if expected != rows[i].balance:
            issues.append(BalanceIssue(
                row_index=i, date=rows[i].date,
                expected=expected, actual=rows[i].balance,
                diff=rows[i].balance - expected,
            ))
    return issues


def validate_final_balance(result: StatementResult) -> Optional[bool]:
    if not result.rows or result.metadata.stated_balance is None:
        return None
    return result.rows[-1].balance == result.metadata.stated_balance


# =============================================================================
# 11. Pipeline
# =============================================================================

def process_pdf(pdf_bytes: bytes, filename: str) -> StatementResult:
    text = extract_pdf_text(pdf_bytes)

    # Detection: try the embedded text layer first. If it's empty (scanned PDF),
    # fall back to a small OCR sample of page 1 just for bank detection.
    bank = detect_bank(text) if text.strip() else None
    if bank is None and not text.strip() and OCR_AVAILABLE:
        try:
            sample = _ocr_pages(pdf_bytes, dpi=300, lang='tur')
            if sample:
                bank = detect_bank(sample[0]['text'])
        except Exception:
            pass

    if bank is not None and bank in PARSERS:
        result = PARSERS[bank](pdf_bytes, text, filename)
    elif bank is not None:
        result = parse_generic(pdf_bytes, text, filename)
        result.metadata.bank = bank
        result.parser_warnings.insert(
            0, f"Detected {bank} but no dedicated parser. Using generic fallback."
        )
    else:
        result = parse_generic(pdf_bytes, text, filename)
        if not text.strip():
            result.parser_warnings.insert(
                0,
                "PDF has no embedded text (scanned/image-only) and OCR was "
                "not available or did not detect a known bank. Install OCR "
                "dependencies (see top of file) or upload a vector PDF."
            )
        else:
            result.parser_warnings.insert(
                0, f"[DEBUG] Bank not detected. First 300 chars: {repr(text[:300])}"
            )

    result.balance_issues = validate_balance_chain(result.rows)
    result.final_balance_match = validate_final_balance(result)
    return result


# =============================================================================
# 12. Output — DataFrame & file generation with proper data types
# =============================================================================

OUTPUT_COLUMNS = ['TARİH', 'Saat', 'TUTAR', 'Bakiye', 'AÇIKLAMA', 'DEKONT']


def result_to_dataframe(result: StatementResult) -> pd.DataFrame:
    """Build a DataFrame with proper dtypes:
    - TARİH:    datetime64[ns]   (real dates, not strings)
    - TUTAR:    float64          (monetary number)
    - Bakiye:   float64          (monetary number)
    - DEKONT:   string           (always text — display compatibility)
    - Saat / AÇIKLAMA: string

    DEKONT is always a string here so concatenated DataFrames from different
    banks (where one is all-digit, another alphanumeric) display cleanly
    through pandas/PyArrow. The Excel writer (df_to_xlsx_bytes) decides
    per-export whether to *promote* DEKONT to Int64 in the workbook.
    """
    if not result.rows:
        return pd.DataFrame({
            'TARİH':    pd.Series(dtype='datetime64[ns]'),
            'Saat':     pd.Series(dtype='string'),
            'TUTAR':    pd.Series(dtype='float64'),
            'Bakiye':   pd.Series(dtype='float64'),
            'AÇIKLAMA': pd.Series(dtype='string'),
            'DEKONT':   pd.Series(dtype='string'),
        })[OUTPUT_COLUMNS]

    df = pd.DataFrame({
        'TARİH':    [r.date for r in result.rows],
        'Saat':     [r.time for r in result.rows],
        'TUTAR':    [float(r.amount) for r in result.rows],
        'Bakiye':   [float(r.balance) for r in result.rows],
        'AÇIKLAMA': [r.description for r in result.rows],
        'DEKONT':   [r.receipt for r in result.rows],
    })[OUTPUT_COLUMNS]

    df['TARİH'] = pd.to_datetime(df['TARİH'], format='%d-%m-%Y', errors='coerce')
    df['Saat'] = df['Saat'].astype('string').fillna('')
    df['TUTAR'] = pd.to_numeric(df['TUTAR'], errors='coerce').astype('float64')
    df['Bakiye'] = pd.to_numeric(df['Bakiye'], errors='coerce').astype('float64')
    df['AÇIKLAMA'] = df['AÇIKLAMA'].astype('string').fillna('')
    df['DEKONT'] = df['DEKONT'].astype('string').fillna('')
    return df


def _inject_ignored_errors(xlsx_bytes: bytes, sqref: str) -> bytes:
    """Inject an `<ignoredErrors>` XML block into the saved xlsx so Excel
    suppresses the green-triangle 'Number stored as text' warning over `sqref`.

    Used when DEKONT is a text column containing numeric-looking receipts
    (e.g. 10-digit Halkbank dekonts) — the data must remain text but the
    warning is noise."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes), 'r') as zin:
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                content = zin.read(item)
                if item == 'xl/worksheets/sheet1.xml':
                    txt = content.decode('utf-8')
                    block = (
                        '<ignoredErrors>'
                        f'<ignoredError sqref="{sqref}" numberStoredAsText="1"/>'
                        '</ignoredErrors>'
                    )
                    # ignoredErrors must come before extLst / after mergeCells per schema;
                    # placing it just before </worksheet> works for Excel in practice.
                    if '<ignoredErrors' not in txt:
                        txt = txt.replace('</worksheet>', block + '</worksheet>')
                    content = txt.encode('utf-8')
                zout.writestr(item, content)
    return out.getvalue()


def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    """Write Excel with proper data types so Excel sees numbers as numbers
    and dates as dates — no green triangles."""
    from openpyxl.utils import get_column_letter

    df = df.copy()

    # DEKONT comes in as `string` dtype (uniform across files). Decide here,
    # at export time, whether to promote it to Int64 (when every non-empty
    # value is purely digits — Halkbank-style 10-digit receipts) so Excel
    # treats it as a number with no green-triangle warning. Otherwise keep
    # it as text and inject `<ignoredErrors>` only if any cell looks numeric.
    dekont_is_numeric = False
    if 'DEKONT' in df.columns and len(df):
        non_empty = df['DEKONT'].astype('string').fillna('').str.strip()
        non_empty = non_empty[non_empty != '']
        if len(non_empty) > 0 and non_empty.str.fullmatch(r'\d+').all():
            df['DEKONT'] = pd.to_numeric(df['DEKONT'], errors='coerce').astype('Int64')
            dekont_is_numeric = True

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

        n_data_rows = len(df)
        last_row = n_data_rows + 1

        if 'TARİH' in df.columns and n_data_rows:
            col_letter = get_column_letter(df.columns.get_loc('TARİH') + 1)
            for row_idx in range(2, last_row + 1):
                ws[f'{col_letter}{row_idx}'].number_format = 'DD.MM.YYYY'

        for col_name in ('TUTAR', 'Bakiye'):
            if col_name in df.columns and n_data_rows:
                col_letter = get_column_letter(df.columns.get_loc(col_name) + 1)
                for row_idx in range(2, last_row + 1):
                    ws[f'{col_letter}{row_idx}'].number_format = '#,##0.00;-#,##0.00'

        if dekont_is_numeric and 'DEKONT' in df.columns and n_data_rows:
            col_letter = get_column_letter(df.columns.get_loc('DEKONT') + 1)
            for row_idx in range(2, last_row + 1):
                ws[f'{col_letter}{row_idx}'].number_format = '0'

    xlsx_bytes = buf.getvalue()

    # Text DEKONT containing some numeric-looking values → suppress green triangles.
    if 'DEKONT' in df.columns and not dekont_is_numeric and len(df):
        has_numericish = df['DEKONT'].astype(str).str.fullmatch(r'\d+').any()
        if has_numericish:
            col_letter = get_column_letter(df.columns.get_loc('DEKONT') + 1)
            sqref = f'{col_letter}2:{col_letter}{len(df) + 1}'
            xlsx_bytes = _inject_ignored_errors(xlsx_bytes, sqref)

    return xlsx_bytes


def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    df = df.copy()
    # Format date column as DD-MM-YYYY for CSV (universal & locale-safe)
    if 'TARİH' in df.columns:
        df['TARİH'] = pd.to_datetime(df['TARİH'], errors='coerce').dt.strftime('%d-%m-%Y')
    for col in ('TUTAR', 'Bakiye'):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    return df.to_csv(index=False, float_format='%.2f').encode('utf-8-sig')


# =============================================================================
# 13. Filename
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
# 14. Upload handling (PDF + ZIP)
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
# 15. Streamlit UI
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
        for w in result.parser_warnings:
            st.warning(w)
        for info in result.parser_info:
            st.info(info)
        return

    if not issues:
        st.success(
            f"✅ Balance chain verified across **{n_rows-1}** consecutive row pairs. "
            f"All {n_rows} transactions reconcile mathematically."
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
            f"**{format_tr_decimal(result.metadata.stated_balance)}**."
        )
    elif final_match is False:
        last = result.rows[-1].balance
        stated = result.metadata.stated_balance
        st.error(
            f"❌ Final balance mismatch — last row says "
            f"**{format_tr_decimal(last)}**, statement header says "
            f"**{format_tr_decimal(stated)}**."
        )

    # Actionable warnings always shown.
    for w in result.parser_warnings:
        st.warning(w)

    # Informational notes (e.g. "OCR was used", repairs that the balance
    # chain already validated) → hidden in an expander when everything
    # reconciled, surfaced inline when there are issues.
    if result.parser_info:
        if issues or final_match is False:
            for info in result.parser_info:
                st.info(info)
        else:
            with st.expander(f"Parser notes ({len(result.parser_info)})"):
                for info in result.parser_info:
                    st.write(f"• {info}")


def main() -> None:
    st.set_page_config(page_title="Bank Statement Extractor", layout="wide")
    st.title("Bank Statement → Excel/CSV")
    st.caption(
        "Accountant-grade extraction for Turkish bank PDFs. "
        "Supported: " + ", ".join(PARSERS.keys()) + ". "
        "Other banks: generic fallback parser."
    )

    if not OCR_AVAILABLE:
        with st.expander("ℹ OCR not available (only needed for scanned PDFs)"):
            st.write(
                "Vector PDFs (Halkbank, Akbank, vector-mode Ziraat) work fine. "
                "Scanned/image-only PDFs won't be readable until OCR is enabled."
            )
            st.code(
                "# Linux / Streamlit Cloud (via packages.txt):\n"
                "tesseract-ocr\ntesseract-ocr-tur\npoppler-utils\n\n"
                "# macOS:\nbrew install tesseract tesseract-lang poppler\n\n"
                "# Windows: install from\n"
                "#   https://github.com/UB-Mannheim/tesseract/wiki\n"
                "#   https://github.com/oschwartz10612/poppler-windows/releases",
                language='text',
            )
            if OCR_IMPORT_ERROR:
                st.caption(f"Detail: {OCR_IMPORT_ERROR}")

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
                st.dataframe(df.head(20), width='stretch')

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
    st.dataframe(combined.head(50), width='stretch')

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