<p align="center">
  <img src="https://raw.githubusercontent.com/kunalgoesbyken/pdf2docx-healer/main/docs/assets/logo.svg" alt="pdf2docx-healer" width="440">
</p>

<p align="center">
  <b>Turn a PDF into a Word document you can actually edit.</b><br>
  Real headings, real lists, working links, tables, and right-to-left text in the right order.<br>
  Runs on your computer, so your files never leave it.
</p>

<p align="center">
  <a href="https://pypi.org/project/pdf2docx-healer/"><img src="https://img.shields.io/pypi/v/pdf2docx-healer.svg?logo=pypi&logoColor=white" alt="PyPI version"></a>
  <a href="https://pypi.org/project/pdf2docx-healer/"><img src="https://img.shields.io/pypi/pyversions/pdf2docx-healer.svg?logo=python&logoColor=white" alt="Python versions"></a>
  <a href="https://github.com/kunalgoesbyken/pdf2docx-healer/actions"><img src="https://img.shields.io/github/actions/workflow/status/kunalgoesbyken/pdf2docx-healer/ci.yml?branch=main&logo=github&label=CI" alt="CI"></a>
  <a href="https://github.com/kunalgoesbyken/pdf2docx-healer/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License: MIT"></a>
</p>

```bash
uvx --from pdf2docx-healer pdf2docx-heal report.pdf --mode editable
```

<p align="center">
  <img src="https://raw.githubusercontent.com/kunalgoesbyken/pdf2docx-healer/main/docs/assets/benchmark.svg" alt="Benchmark: pdf2docx-healer keeps more words in reading order and is the only tool that finds headings, against pdf2docx, pdf2docx-plus and LibreOffice" width="880">
</p>

---

## Why

Most PDF to Word converters copy how the page looks. You get headings that are only big bold text, bullets typed as characters, two-column papers read across the columns, and Hebrew or Arabic words in the wrong order. Editing that file is slow and painful.

pdf2docx-healer rebuilds the structure instead. It is a drop-in upgrade for `pdf2docx`, with the same one-line call.

| | pdf2docx | pdf2docx-plus | LibreOffice | **pdf2docx-healer** |
| --- | :---: | :---: | :---: | :---: |
| Word heading styles (navigation pane, table of contents) | no | no | no | **yes** |
| Real Word lists | no | some | no | **yes** |
| Internal links (citations, "see section 2") | no | no | no | **yes** |
| Words kept in reading order | 0.85 | 0.86 | 0.70 | **0.94** editable, 0.88 default |
| Pages lost out of 230 | 0 | 0 | 36 | **0** |

The numbers come from a public benchmark of 12 real PDFs plus one generated test PDF. [Full results and where rivals still win](#benchmark).

Good for:
- **Papers and reports** you need to edit, cite or reuse.
- **Translation work**, where a CAT tool needs a clean DOCX.
- **Hebrew, Arabic, Chinese, Hindi** and other non-Latin documents.
- **Pipelines** that need document structure, not just text.

---

## Install

With [uv](https://docs.astral.sh/uv/):

```bash
uv tool install pdf2docx-healer                  # the pdf2docx-heal command, in its own environment
uvx --from pdf2docx-healer pdf2docx-heal in.pdf  # run once without installing
uv add pdf2docx-healer                           # as a dependency of your project
```

With pip or pipx:

```bash
pip install pdf2docx-healer
pipx install pdf2docx-healer
```

Python 3.10 to 3.13 on Windows, macOS and Linux.

---

## Quick start

```bash
pdf2docx-heal report.pdf                      # writes report.docx, keeping the page layout
pdf2docx-heal report.pdf --mode editable      # a reflowing document that is easy to edit
pdf2docx-heal scan.pdf --ocr --ocr-lang eng   # scanned PDFs (needs Tesseract)
```

```python
from docx_healer import heal, HealerConfig

heal("report.pdf", "report.docx")
heal("report.pdf", "report.docx", config=HealerConfig(mode="editable"))
```

### Two modes

| | `fidelity` (default) | `editable` |
| --- | --- | --- |
| How | `pdf2docx` page by page, then repair passes | Rebuilds the document from the PDF's text, tables and images |
| Looks like | The PDF page | A normal Word document that reflows |
| Best for | Forms, layouts you must keep | Papers, reports, books, right-to-left text, anything you will edit |

Both modes add Word heading styles (from the PDF outline, or font sizes when there is none), real Word lists, clickable web and mail links, internal links for citations and cross references, and font fallbacks for CJK, Arabic, Hebrew, Thai and Devanagari.

Editable mode also reads two-column pages in the right order and keeps right-to-left paragraphs in order. Repeating headers, footers, logos and page numbers move into the Word header and footer. Lists come through with their real format: bullets, `1.`, `a)`, `iii.`, and outlines like `1.2.3`.

---

## Benchmark

Run on 2026-11-26 with healer 0.3.0, pdf2docx 0.5.13, pdf2docx-plus 0.6.10 and LibreOffice (headless, `writer_pdf_import`). Every number is computed from the PDF itself and the produced DOCX; nothing is read from logs. The full table, per-document rows, the corpus and the exact definitions are in [`bench/README.md`](https://github.com/kunalgoesbyken/pdf2docx-healer/blob/main/bench/README.md), and `python -m bench.run` reproduces it.

| Tool | Word recall | Recall any order | Headings | Real lists | Number shown | Internal links | Tables | Images | Pages lost | Median s |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| **healer editable** | **0.941** | 0.991 | **0.79** | **1.00** | **1.00** | 0.85 | **0.88** | 0.84 | **0/230** | 6.3 |
| **healer fidelity** | 0.879 | **0.992** | 0.70 | 0.76 | **1.00** | **0.95** | 0.50 | 0.66 | **0/230** | **5.9** |
| pdf2docx-plus | 0.855 | 0.986 | 0.00 | 0.68 | 0.88 | 0.00 | 0.64 | 0.66 | **0/230** | 7.3 |
| pdf2docx | 0.848 | 0.958 | 0.00 | 0.00 | **1.00** | 0.00 | 0.33 | 0.66 | **0/230** | 6.7 |
| LibreOffice | 0.698 | 0.929 | 0.00 | 0.00 | 0.00 | 0.00 | 0.23 | **0.93** | 36/230 | 11.0 |

Some highlights from the per-document rows:
- **Two-column paper (LLaMA):** editable keeps 0.986 of the words in reading order, against 0.78 for the best rival.
- **Hebrew:** editable keeps 0.964, against 0.60 for every other tool.
- **Arabic:** editable keeps 0.768, against 0.48.
- **Hindi syllabus:** LibreOffice drops all 36 pages; healer drops none.

Where a rival still wins:
- **Images:** LibreOffice scores highest because the score counts only raster images. On the two-column paper, healer turns the vector charts into pictures (counted as extra) and LibreOffice leaves them out (counted as correct).
- **Tables in fidelity mode:** pdf2docx-plus scores 0.64 against healer's 0.50. Editable mode scores 0.88.
- **Single documents:** editable is a little behind on the Chinese text (0.958 against 0.993 for pdf2docx-plus) and on the Mixtral paper (0.928 against 0.951). Fidelity mode scores 1.000 on the Chinese text and 0.950 on Mixtral.

---

## Options

| CLI | `HealerConfig` | Default | What it does |
| --- | --- | --- | --- |
| `--mode` | `mode` | `fidelity` | `fidelity` or `editable` |
| `--ocr`, `--ocr-lang`, `--ocr-dpi`, `--ocr-threshold` | `ocr_enabled`, `ocr_lang`, `ocr_dpi`, `ocr_threshold` | off, `eng`, 300, 0.3 | OCR pages without text |
| `--no-headings` | `fix_headings` | on | Fidelity: heading styles |
| `--no-lists` | `fix_lists` | on | Fidelity: real Word lists |
| `--no-hyperlinks` | `fix_hyperlinks` | on | Clickable web and mail links |
| `--no-internal-links` | `fix_internal_links` | on | Fidelity: links to other places in the document |
| `--no-font-fix` | `fix_fonts` | on | Font fallbacks |
| `--no-table-filter` | `fix_tables` | on | Fidelity: drop tables that are only page layout |
| `--no-page-recovery` | `recover_pages` | on | Fidelity: rebuild a page pdf2docx fails on |
| `--aggressive` | `aggressive_lists` | off | Fidelity: split paragraphs more eagerly to find lists |
| `-q` | `verbose=False` | | No progress output |

Editable mode always builds headings, lists and internal links, so the three switches marked Fidelity only apply to the default mode.

`heal()` returns the output path. It raises `FileNotFoundError` for a missing input. It raises `InvalidInputError` (a `ValueError`) for a non-PDF or unreadable input, a missing output folder, an output path equal to the input, or a bad option. A failure in one repair pass is logged as a warning on the `docx_healer` logger and the other passes still run. The command exits with 0 on success, 1 on failure and 2 for bad options.

The output is written only when the whole conversion has finished, so a failure never leaves a half-written file and never replaces an existing one. Your PDF is never changed.

### Safety

- **Links:** only `http`, `https`, `mailto` and `ftp` links are kept as links; anything else, such as `file:`, `javascript:` or network paths, becomes plain text. This check always runs, and if it fails the conversion fails.
- **Large PDFs:** editable mode caps images, image size and pixels, drawings and table time, so a hostile PDF cannot exhaust memory.
- **OCR:** runs through PyMuPDF and needs [Tesseract](https://github.com/tesseract-ocr/tesseract) on `PATH` or `TESSDATA_PREFIX`. Without it, scanned pages are kept as images and a warning is logged.

---

## Requirements

- Python 3.10+
- `pdf2docx >= 0.5.9, < 0.6`, `PyMuPDF >= 1.27.1`, `python-docx >= 1.1.0`, `lxml >= 4.9`
- [Tesseract](https://github.com/tesseract-ocr/tesseract), optional, for OCR

---

## Found a PDF that converts badly?

[Open an issue](https://github.com/kunalgoesbyken/pdf2docx-healer/issues/new) and attach the PDF, or a link to it, with what went wrong. Bad conversions are the fastest way to make this better.

---

## Development

```bash
uv sync --extra dev
uv run pytest
uv run ruff check src tests bench
uv run mypy
```
