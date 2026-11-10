# pdf2docx-healer

[![PyPI version](https://img.shields.io/pypi/v/pdf2docx-healer.svg?logo=pypi&logoColor=white)](https://pypi.org/project/pdf2docx-healer/)
[![Python versions](https://img.shields.io/pypi/pyversions/pdf2docx-healer.svg?logo=python&logoColor=white)](https://pypi.org/project/pdf2docx-healer/)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/kunalgoesbyken/pdf2docx-healer/blob/main/LICENSE)
[![CI/CD](https://img.shields.io/github/actions/workflow/status/kunalgoesbyken/pdf2docx-healer/publish.yml?logo=github&label=publish)](https://github.com/kunalgoesbyken/pdf2docx-healer/actions)

**A drop-in replacement for `pdf2docx` that actually preserves your formatting.**

`pdf2docx` is a great PDF-to-DOCX converter, but it drops bullet lists, loses hyperlinks, mangles CJK fonts, and chokes on scanned PDFs. `pdf2docx-healer` wraps `pdf2docx` and heals all of these issues in a post-processing pass — so your Word documents come out looking the way they should.

---

## Why this exists

| Problem | `pdf2docx` alone | With `pdf2docx-healer` |
|---------|------------------|------------------------|
| Bullet lists (`•`, `-`, `*`) | Flattened to plain text, no Word list style | Proper `List Bullet` style with real Word numbering |
| Numbered lists (`1.`, `a.`, `i.`) | Lost or merged into one paragraph | `List Number` style; lettered/roman via OOXML injection |
| Nested lists (3+ levels) | Indentation lost | Level detected from indent, applied to Word |
| Hyperlinks | URL text is plain, not clickable | Wrapped in real `<w:hyperlink>` elements, blue and underlined |
| CJK fonts (Chinese/Japanese/Korean) | Font names like `SimSun` may not resolve | Fallback chain maps to system-available CJK fonts |
| Scanned PDFs (image-only) | "Words count: 0" warning, empty output | OCR via Tesseract, then normal conversion |
| Section headers styled as lists | Headers like "4. Numbered List" get list style | Detected as headers, kept as Normal paragraphs |

---

## Install

```bash
pip install pdf2docx-healer
```

OCR for scanned PDFs runs through PyMuPDF's built-in Tesseract integration, so it needs no extra Python package. Install [Tesseract](https://github.com/tesseract-ocr/tesseract) with the language data you need, and make sure `tesseract` is on `PATH` or that `TESSDATA_PREFIX` points to its `tessdata` folder. Without Tesseract, scanned PDFs are still converted but a warning is logged and no OCR is applied.

---

## Quick start

### Python API

```python
from docx_healer import heal

# Simplest usage — output goes to "report.docx"
heal("report.pdf", "report.docx")
```

```python
from docx_healer import heal, HealerConfig

# Full control via config
config = HealerConfig(
    ocr_enabled=True,          # OCR for scanned/image PDFs
    ocr_lang="eng",            # Tesseract language code
    ocr_dpi=300,               # OCR resolution
    ocr_threshold=0.3,         # Fraction of textless pages to trigger OCR
    fix_lists=True,            # Detect & style bullet/numbered lists
    fix_hyperlinks=True,       # Wrap URL text in clickable hyperlinks
    fix_fonts=True,            # Map CJK/unavailable fonts to system fonts
    aggressive_lists=False,    # More aggressive paragraph splitting
    verbose=True,              # Print progress
)

heal("scanned_report.pdf", "output.docx", config=config)
```

`heal()` returns the output path. It raises `FileNotFoundError` for a missing input and `InvalidInputError` (a `ValueError`) for a non-PDF input, a missing output directory, an output path equal to the input, or an invalid config value. `IncompatiblePdf2docxError` is raised if the installed `pdf2docx` does not expose the internals used for list detection; use `fix_lists=False` to skip that step. Problems in an individual post-processing pass (lists, hyperlinks, fonts) are logged as warnings on the `docx_healer` logger and the remaining passes still run.

The output file is written only after conversion and post-processing finish, so a failure never leaves a half-written `.docx` or replaces an existing one. Temporary files created for OCR are always removed, and your input PDF is never modified or deleted.

### Command line

```bash
# Basic conversion
pdf2docx-heal input.pdf -o output.docx

# Scanned PDF with OCR
pdf2docx-heal input.pdf --ocr --ocr-lang eng

# Quiet mode (no progress output)
pdf2docx-heal input.pdf -q

# Skip specific fixes
pdf2docx-heal input.pdf --no-lists --no-hyperlinks
```

Run `pdf2docx-heal --help` to see all options. The command exits with status 0 on success, 1 when the conversion fails (missing or invalid input, conversion error), and 2 for invalid options.

---

## What it fixes

- **Bullet lists** — Detects Unicode (`•`, `◦`, `▪`, `–`) and ASCII (`-`, `*`, `+`) bullets, applies Word's `List Bullet` style. Nested bullets (up to 5 levels) detected from indentation.
- **Numbered lists** — Detects decimal (`1.`), parenthesized (`(1)`), lettered (`a.`), roman (`i.`), and outline (`1.1`) numbering. Lettered/roman use OOXML injection with correct `numFmt` since Word's built-in styles only support decimal.
- **Hyperlinks** — Scans runs for `http://`, `https://`, `www.`, `mailto:`, `ftp://` and wraps them in `<w:hyperlink>` elements with external relationship targets. Multiple URLs in one run all get converted, trailing punctuation such as the period ending a sentence is left outside the link, and URLs in table cells, nested tables, headers and footers are covered too.
- **CJK font fallback** — Maps embedded font names (`SimSun`, `MS-Mincho`, `HYGoThic-Medium`) to system-available equivalents across Windows/macOS/Linux. Character-range detection maps unknown fonts by script (CJK, Arabic, Hebrew, Thai, Devanagari). Other fonts, for example Calibri, are left as they are.
- **Scanned PDF OCR** — Detects image-only PDFs and runs Tesseract OCR via PyMuPDF. Falls back gracefully if Tesseract isn't installed.
- **Smart header detection** — Headers like `"4. Numbered List"` are detected via sequential-reset analysis and kept as Normal paragraphs instead of being styled as list items.

---

## Requirements

- Python 3.8+
- `pdf2docx >= 0.5.9, < 0.6`, `PyMuPDF >= 1.23.0`, `python-docx >= 0.8.11`, `lxml`
- [Tesseract](https://github.com/tesseract-ocr/tesseract) (optional, for OCR; used through PyMuPDF)

---

## Development

```bash
pip install -e ".[dev]"
pytest
ruff check src tests
mypy
```
