# pdf2docx-healer: Formatting-Preserving Wrapper around pdf2docx

## Problem
pdf2docx loses formatting: bullet lists become invisible borderless tables, plain-text URLs are not clickable, CJK fonts break, scanned PDFs produce empty output. pdf2docx itself is maintained (0.5.13 was released in May 2026 and it sees about 1.8 million downloads a month), so this project wraps it and repairs the output; it does not replace it.

How much better the output is gets measured, not claimed: see `bench/README.md` for the corpus, the metrics and the latest scoreboard.

## Root Causes Found
Checked against pdf2docx 0.5.13.

1. `is_list_item()` in `pdf2docx/common/share.py` starts with `return False`, so the rest of the function is dead code and no text is ever recognised as a list item.
2. `Converter.default_settings` in `pdf2docx/converter.py` already sets `list_not_table` to `True` and passes its keyword arguments through to `Blocks.collect_stream_lines`. That method closes a stream table only when `kwargs.get('list_not_table') and is_list_item(row[0].text)` is true. Because `is_list_item` always returns `False`, the option has no effect and bullet rows can still be grouped into a table.
3. Font names from PyMuPDF (e.g. SimSun, MS-Mincho) don't exist on target systems.
4. Image-only PDF pages produce no text (no OCR integration).

## Architecture
`heal()` in `converter.py` runs these steps in order. The output file is written only after every step up to post-processing has succeeded.

### Step 1: Pre-parse (ocr_pipeline.py, opt-in)
- When `ocr_enabled` is set, detect scanned pages (too little embedded text) and OCR them with PyMuPDF's Tesseract integration into a temporary copy of the PDF. The input file is never modified.
- Without Tesseract installed a warning is logged and no OCR is applied.

### Step 2: Convert (patches.py)
- Convert with `pdf2docx.Converter`.
- While the conversion runs, and only then, replace `pdf2docx.layout.Blocks.is_list_item` (the name bound inside the Blocks module) with a working detector and wrap `Blocks.collect_stream_lines` so `list_not_table` is always `True`. Both are restored afterwards.
- If the installed pdf2docx lacks these internals, `IncompatiblePdf2docxError` is raised before anything is patched.

### Step 3: Post-process (post_processor.py)
Each pass opens and saves the DOCX on its own. A failing pass is logged and skipped, and the report lists it. Passes run in this order:
- tables: flatten layout tables that are not data (hook in place, returns 0 until the logic lands)
- headings: Word heading styles from the PDF outline and font sizes (hook in place, returns 0 until the logic lands)
- lists: detect bullet and number prefixes, split merged paragraphs, apply Word list styles or custom OOXML numbering (`list_detector.py`)
- hyperlinks: wrap plain-text URLs in `w:hyperlink` elements in body, tables, headers and footers (`link_injector.py`)
- internal links: recreate PDF internal links as bookmark hyperlinks (hook in place, returns 0 until the logic lands)
- fonts: map CJK and unavailable fonts to system fonts (`font_mapper.py`)

### Modes
- `fidelity` (default): the steps above.
- `editable`: planned, a reflowing DOCX built from the PDF structure (`flow_writer.py`). It currently raises `NotImplementedError`, and the benchmark reports it as "not available".

## Dependencies
- pdf2docx >= 0.5.9, < 0.6
- PyMuPDF >= 1.23.0
- python-docx >= 0.8.11
- lxml (included with python-docx)
- Tesseract OCR (optional, system-installed)
