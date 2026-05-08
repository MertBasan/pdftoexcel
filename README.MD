# Bank Statement → Excel/CSV Extractor

Streamlit app that converts Turkish bank-statement PDFs into clean Excel /
CSV with proper data types (real dates, real numbers — no green-triangle
"Number stored as text" warnings) and a balance-chain integrity check.

**Supported banks**: Halkbank (V1 + V2 layouts), Akbank, Ziraat (vector +
scanned/OCR). Other banks fall through to a generic parser.

---

## Sharing the app — Streamlit Community Cloud (recommended)

This is the zero-friction path. Anyone with the URL can use it; nothing
to install on their end.

1. Push this folder to a public (or private) GitHub repository. The repo
   must contain at least:
   - `app.py`
   - `requirements.txt`  ← Python deps
   - `packages.txt`      ← apt packages (provides tesseract & poppler binaries)
2. Go to <https://share.streamlit.io>, sign in with GitHub, click
   **New app**, pick the repo and `app.py` as the entry point.
3. Click **Deploy**. Streamlit Cloud reads `packages.txt` and runs
   `apt-get install` on the listed packages, then `pip install -r
   requirements.txt`. First boot takes ~3 minutes; subsequent loads are
   instant.
4. Share the resulting `https://<your-app>.streamlit.app` URL.

Users hit the URL in any browser. No Python, no tesseract, no anything.

---

## Local development

### macOS

```bash
brew install tesseract tesseract-lang poppler
python -m venv env && source env/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

### Linux (Ubuntu / Debian)

```bash
sudo apt install tesseract-ocr tesseract-ocr-tur poppler-utils
python -m venv env && source env/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

### Windows

1. Install Tesseract from
   <https://github.com/UB-Mannheim/tesseract/wiki> — accept the default
   install path `C:\Program Files\Tesseract-OCR`. **Make sure to tick
   the Turkish language pack** during install.
2. Install poppler from
   <https://github.com/oschwartz10612/poppler-windows/releases> — extract
   to `C:\Program Files\poppler` so the binaries land at
   `C:\Program Files\poppler\Library\bin`.
3. Then:
   ```powershell
   python -m venv env
   .\env\Scripts\Activate.ps1
   pip install -r requirements.txt
   streamlit run app.py
   ```

If you installed tesseract or poppler somewhere else, set env vars before
running streamlit:

```powershell
$env:TESSERACT_CMD = "D:\tools\Tesseract-OCR\tesseract.exe"
$env:POPPLER_PATH  = "D:\tools\poppler\Library\bin"
streamlit run app.py
```

---

## What's in the box

| File | Purpose |
| --- | --- |
| `app.py` | The Streamlit app and all parsers. |
| `requirements.txt` | Python deps for pip (also read by Streamlit Cloud). |
| `packages.txt` | apt packages for Streamlit Cloud (tesseract, poppler). |
| `README.md` | This file. |

Vector PDFs work without OCR. OCR is only consulted when a PDF has zero
embedded text (i.e. is scanned), e.g. some Ziraat exports.

---

## Output guarantees

For every uploaded file, the app:

1. Detects the issuing bank from the page header.
2. Picks the right parser (or falls back to a generic one).
3. Validates that **every consecutive (balance + amount = next balance)**
   pair holds. Any mismatch is surfaced loudly.
4. Writes Excel with real data types: `TARİH` is a date with format
   `DD.MM.YYYY`, `TUTAR` and `Bakiye` are numbers with format
   `#,##0.00;-#,##0.00`, `DEKONT` is integer when all values are pure
   digits and text otherwise (with the `numberStoredAsText` warning
   suppressed for numeric-looking text).

If the balance chain doesn't reconcile, **the output is marked
untrustworthy in big red text**. Don't ship it without checking.