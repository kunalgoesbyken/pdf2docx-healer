# pdf2docx-healer: Formatting-Preserving PDF-to-DOCX Wrapper

## Problem
pdf2docx loses formatting: bullet lists become invisible borderless tables, hyperlinks vanish, CJK fonts break, scanned PDFs produce empty output. The library is abandoned by Artifex.

## Root Causes Found
1. `is_list_item()` in `pdf2docx/common/share.py:202` always returns False (dead code)
2. `list_not_table` kwarg is dropped in the `Layout._parse_table()` → `TablesConstructor.stream_tables()` call chain
3. Font names from PyMuPDF (e.g. SimSun, MS-Mincho) don't exist on target systems
4. Image-only PDF pages produce no text (no OCR integration)

## Architecture: 4-Phase Hybrid Pipeline

### Phase 1: Pre-parse (ocr_pipeline.py)
- Detect scanned pages via empty `page.get_text('text')`
- PyMuPDF `get_textpage_ocr(flags=3)` → `get_text('rawdict', textpage=tp)`
- Produces pdf2docx-compatible rawdict, zero pipeline changes needed

### Phase 2: Intercept (patches.py)
- Patch `pdf2docx.layout.Blocks.is_list_item` (the name bound inside the Blocks module) to actually detect bullets
- Patch `pdf2docx.layout.Blocks.collect_stream_lines` to force `list_not_table=True`
- Patches apply only for the duration of the conversion and fail loudly on incompatible pdf2docx versions

### Phase 3: Enrich (link_injector.py)
- pdf2docx already extracts hyperlinks via shape pipeline
- Supplement: scan store() JSON for missed URL patterns in spans

### Phase 4: Post-process (post_processor.py)
- Open .docx with python-docx
- Detect list paragraphs by bullet/number prefix
- Apply Word list styles (built-in or custom OOXML numbering)
- Font fallback mapping for CJK/unavailable fonts
- Normalize indentation from pdf2docx coordinate calculations

## Dependencies
- pdf2docx >= 0.5.9, < 0.6
- PyMuPDF >= 1.23.0
- python-docx >= 0.8.11
- lxml (included with python-docx)
- Tesseract OCR (optional, system-installed)
