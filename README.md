# pdf2docx-healer

[![PyPI version](https://img.shields.io/pypi/v/pdf2docx-healer.svg?logo=pypi&logoColor=white)](https://pypi.org/project/pdf2docx-healer/)
[![Python versions](https://img.shields.io/pypi/pyversions/pdf2docx-healer.svg?logo=python&logoColor=white)](https://pypi.org/project/pdf2docx-healer/)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/krockxz/pdf2docx-healer/blob/main/LICENSE)
[![CI/CD](https://img.shields.io/github/actions/workflow/status/krockxz/pdf2docx-healer/publish.yml?logo=github&label=publish)](https://github.com/krockxz/pdf2docx-healer/actions)

**A drop-in replacement for `pdf2docx` that actually preserves your formatting.**

`pdf2docx` is a great PDF-to-DOCX converter, but it has a frustrating habit of dropping bullet lists, losing hyperlinks, mangling CJK fonts, and choking on scanned PDFs. `pdf2docx-healer` wraps `pdf2docx` and heals all of these issues in a post-processing pass — so your Word documents come out looking the way they should.

---

## Why this exists

| Problem | `pdf2docx` alone | With `pdf2docx-healer` |
|---------|------------------|------------------------|
| Bullet lists (`•`, `-`, `*`) | Often flattened to plain text, no Word list style | Proper `List Bullet` style with real Word numbering |
| Numbered lists (`1.`, `a.`, `i.`) | Lost or merged into one paragraph | `List Number` style; lettered/roman via OOXML injection |
| Nested lists (3+ levels) | Indentation lost | Level detected from indent, applied to Word |
| Hyperlinks | URL text is plain, not clickable | Wrapped in real `<w:hyperlink>` elements with blue/underline |
| CJK fonts (Chinese/Japanese/Korean) | Font names like `SimSun` may not resolve | Fallback chain maps to system-available CJK fonts |
| Scanned PDFs (image-only) | "Words count: 0" warning, empty output | OCR via Tesseract, then normal conversion |
| Section headers styled as lists | Headers like "4. Numbered List" get list style | Detected as headers, kept as Normal paragraphs |

---

## Install

```bash
pip install pdf2docx-healer
```

For OCR support on scanned PDFs, also install [Tesseract](https://github.com/tesseract-ocr/tesseract) and the optional extra:

```bash
pip install "pdf2docx-healer[ocr]"
```

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

Run `pdf2docx-heal --help` to see all options.

---

## What it fixes

### Bullet lists

Detects Unicode bullets (`•`, `◦`, `▪`, `–`, etc.) and ASCII bullets (`-`, `*`, `+`) and applies Word's built-in `List Bullet` style. Nested bullets (up to 5 levels) are detected from indentation and mapped to the right list level.

### Numbered lists

Detects decimal (`1.`), parenthesized (`(1)`), lettered (`a.`, `b.`), roman (`i.`, `ii.`), and outline (`1.1`, `1.2`) numbering. Decimal and parenthesized use Word's `List Number` style. Lettered and roman formats use OOXML numbering injection with the correct `numFmt` (`lowerLetter`, `lowerRoman`) since Word's built-in styles only support decimal.

### Hyperlinks

Scans all paragraph runs for URL patterns (`http://`, `https://`, `www.`, `mailto:`, `ftp://`) and wraps them in proper OOXML `<w:hyperlink>` elements with external relationship targets. Multiple URLs in a single run are all converted. Hyperlink text gets blue color and underline styling.

### CJK font fallback

Maps PDF-embedded font names (like `SimSun`, `MS-Mincho`, `HYGoThic-Medium`) to system-available equivalents across Windows, macOS, and Linux. Falls back through a chain: e.g. `SimSun` → `宋体` → `Microsoft YaHei` → `微軟雅黑` → `Arial Unicode MS` → `Noto Sans CJK SC`. Character-range detection also maps unknown fonts based on the script being rendered (CJK, Arabic, Hebrew, Thai, Devanagari, Cyrillic).

### Scanned PDF OCR

Detects image-only PDFs (no text layer) and runs them through PyMuPDF's OCR pipeline (requires Tesseract). If Tesseract isn't installed, it falls back gracefully instead of crashing. The OCR'd PDF is then converted normally.

### Smart header detection

Section headers that look like list items (e.g. `"4. Numbered List"`) are detected via sequential-reset analysis and title-case heuristics, and kept as Normal paragraphs instead of being styled as list items.

---

## How it works

`pdf2docx-healer` runs a 4-phase pipeline:

```
┌─────────────┐    ┌──────────────┐    ┌─────────────┐    ┌──────────────┐
│  1. Pre-parse│ -> │ 2. Intercept │ -> │  3. Convert │ -> │ 4. Post-process│
│  (OCR if     │    │  (monkey-    │    │  (pdf2docx  │    │  (lists, links,│
│  needed)     │    │   patch)     │    │   core)     │    │   fonts)       │
└─────────────┘    └──────────────┘    └─────────────┘    └──────────────┘
```

1. **Pre-parse** — If OCR is enabled, detects whether the PDF is scanned and runs Tesseract OCR to add a text layer.
2. **Intercept** — Monkey-patches `pdf2docx`'s dead `is_list_item()` function with real bullet/number detection, and forces `list_not_table=True` so list blocks aren't parsed as tables.
3. **Convert** — Runs `pdf2docx` with the patched internals.
4. **Post-process** — Opens the output DOCX with `python-docx` and fixes lists (splitting, styling, numbering XML), hyperlinks (OOXML injection), and fonts (fallback mapping).

---

## Configuration reference

### `HealerConfig` fields

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `ocr_enabled` | `bool` | `False` | Enable OCR for scanned PDFs |
| `ocr_lang` | `str` | `"eng"` | Tesseract language code |
| `ocr_dpi` | `int` | `300` | OCR resolution in DPI |
| `ocr_threshold` | `float` | `0.3` | Fraction of textless pages to trigger OCR |
| `fix_lists` | `bool` | `True` | Detect & style bullet/numbered lists |
| `fix_hyperlinks` | `bool` | `True` | Wrap URL text in clickable hyperlinks |
| `fix_fonts` | `bool` | `True` | Map CJK/unavailable fonts to system fonts |
| `aggressive_lists` | `bool` | `False` | More aggressive paragraph splitting |
| `verbose` | `bool` | `True` | Print progress output |

### CLI flags

| Flag | Description |
|------|-------------|
| `pdf` | Input PDF file path (positional) |
| `-o, --output` | Output DOCX path (default: input with `.docx`) |
| `--ocr` | Enable OCR for scanned PDFs |
| `--ocr-lang` | Tesseract language code (default: `eng`) |
| `--ocr-dpi` | OCR resolution in DPI (default: `300`) |
| `--ocr-threshold` | Fraction of textless pages to trigger OCR (default: `0.3`) |
| `--no-lists` | Skip list detection and formatting |
| `--no-hyperlinks` | Skip hyperlink extraction |
| `--no-font-fix` | Skip CJK font fallback mapping |
| `--aggressive` | Use aggressive paragraph splitting |
| `-q, --quiet` | Suppress progress output |

---

## Requirements

- Python 3.8+
- `pdf2docx >= 0.5.0`
- `PyMuPDF >= 1.23.0`
- `python-docx >= 0.8.11`
- `lxml`
- [Tesseract](https://github.com/tesseract-ocr/tesseract) (optional, for OCR)

---

## Limitations

- `pdf2docx` (the underlying engine) is no longer actively maintained. This package works around its bugs but can't fix fundamental parsing limitations.
- OCR requires Tesseract installed separately. Without it, scanned PDFs fall back to image-only output.
- Hyperlink text that `pdf2docx` drops during conversion (due to overlapping link annotations) cannot be recovered — only URL text that survives conversion gets wrapped.
- Some paragraph merge patterns by `pdf2docx` (without `\n` separators) may persist.

---

## License

MIT — see [LICENSE](https://github.com/krockxz/pdf2docx-healer/blob/main/LICENSE).

## Contributing

Issues and pull requests welcome at [github.com/krockxz/pdf2docx-healer](https://github.com/krockxz/pdf2docx-healer).
