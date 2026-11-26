# Changelog

## 0.4.0

### Added
- Editable mode copies the look of the PDF: font family (with metric-compatible stand-ins such as Times New Roman for Nimbus Roman), size, colour, bold and italic of every run; paragraph alignment, indents, space before and exact line spacing; page size, margins and columns, with a new section where the column layout changes; centred and right-aligned lines keep their breaks. Heading and title styles no longer add their own blue Calibri look.
- Text running up or down the margin and blocks that sit beside other text are pinned at their place on the page. Text on a coloured box keeps the colour behind it as paragraph shading, and text that would be white on white keeps the body colour.
- Word hyphenates and squeezes word spaces when the PDF was typeset that way, so justified lines fill the same.
- `python -m bench.visual` renders the DOCX (Word or LibreOffice) and compares it with the PDF word by word, with side-by-side pictures.

### Changed
- Default mode's font fixer uses the same font family resolver.

### Fixed
- Lines in a tight block, such as a list of author names, are no longer marked as headings.
- A paragraph that starts with an indented first line right after a short last line is no longer merged into the paragraph above.
- A bold line with superscript marks, such as the stars after author names, still counts as bold.

## 0.3.2

### Changed
- README: a header with the icon, a card showing a real PDF page next to the Word document healer makes from it, and a plain comparison with pdf2docx, pdf2docx-plus and LibreOffice.

## 0.3.1

### Changed
- New README with a logo, a benchmark chart and a feature comparison; images use absolute links so they also show on PyPI.
- `python -m bench.chart <results.json> <out.svg>` draws the benchmark chart from a results file.

## 0.3.0

### Added
- `--mode editable` (`HealerConfig(mode="editable")`): builds a reflowing Word document straight from the PDF, with real headings, real Word lists (bullets, numbers, letters, roman, dotted outlines such as 1.2.3), tables, pictures, two-column reading order, right-to-left paragraphs, and running headers, footers and page numbers moved into the Word header and footer. Work on large or hostile PDFs is capped by `FlowLimits` (images, image size and pixels, drawings, table time).
- Headings: Word heading styles from the PDF outline (bookmarks), with font-size detection when there is no outline.
- Internal links: citations, footnote marks and cross references become Word bookmarks and links.
- Email addresses become `mailto:` links.
- Page recovery: a page pdf2docx fails on is rebuilt from its text and images instead of being dropped.
- A reproducible benchmark (`bench/`) against pdf2docx, pdf2docx-plus and LibreOffice on 12 public PDFs plus a generated one.
- `--version`, Windows and macOS CI, Python 3.13 support, and a CI job on the oldest allowed dependencies.

### Fixed
- Lists: a page is never dropped, numbering restarts where the PDF restarts, lettered and roman formats are kept, nesting follows the indent, headings and prose that start with a number are left alone, a list no longer continues across a table, and splitting a paragraph no longer duplicates its section break.
- Fidelity mode keeps the spaces between words, drops tables that are only page layout, and no longer turns a rotated margin label into a table around the whole page.
- Link safety: every link target that is not `http`, `https`, `mailto` or `ftp` (for example `file:`, `javascript:` or a UNC path) is turned into plain text. This check always runs, and if it fails the conversion fails instead of keeping an unchecked file. Links pdf2docx nests inside runs, which Word's schema does not allow, are lifted to paragraph level.
- A quadratic scan over images made large documents slow.

### Changed
- Requires Python 3.10 or newer, `PyMuPDF>=1.27.1`, `python-docx>=1.1.0` and `lxml>=4.9`. Older PyMuPDF groups figures differently, python-docx 1.0 cannot be imported by this package, and 0.8 does not read link text.
- Built and published with uv; each release is also attached to its GitHub release.

## 0.2.0

### Fixed
- Deleting the input PDF: with `ocr_threshold` below 0.3 and a scanned fraction between that value and 0.3, the input PDF was removed after conversion. OCR now receives the configured threshold, and only the temporary OCR copy created by the package is ever deleted.
- The list patches never took effect: `Blocks.is_list_item` was left unpatched because the `Blocks` module binds its own reference. Patches are now applied for the duration of the conversion only, restored afterwards, and raise `IncompatiblePdf2docxError` when the installed pdf2docx lacks the expected internals.
- Hyperlink and font passes now cover table cells (including nested tables), headers, footers, and runs inside hyperlinks.
- Fonts: unknown fonts such as Calibri are kept instead of being replaced by Arial, PDF subset prefixes are stripped, `eastAsia` and `cs` font attributes are mapped, and Cyrillic text no longer forces Times New Roman.
- Hyperlinks: trailing punctuation is no longer part of the link, link text is blue and underlined, uppercase `MAILTO:` and `HTTPS://` schemes are handled, and runs holding anything besides plain text are left untouched.
- Injected list numbering definitions now place `abstractNum` before `num`, define all nine levels, and write `numPr` in schema order.
- A failing post-processing pass is logged as a warning (not only in verbose mode) and reported by `post_process_docx`.
- OCR failures (missing Tesseract, per-page errors) are logged as warnings, and the temporary OCR PDF is removed on every path.
- `heal()` no longer changes the root logger level after a quiet conversion.

### Changed
- `heal()` validates its input (missing file, non-PDF, output directory, output equal to input) and writes the output atomically.
- `HealerConfig` validates `ocr_threshold`, `ocr_dpi` and `ocr_lang`.
- The CLI exits with status 1 on failure and 2 on invalid options.
- `ocr_pdf_file` always creates and returns a new temporary file. `post_process_docx` returns a `PostProcessReport`. `apply_patches` and `is_patched` are replaced by the `list_item_patches()` context manager.
- Requires `pdf2docx>=0.5.9,<0.6`.
- The `[ocr]` extra was removed because `pytesseract` was never used. OCR runs through PyMuPDF and needs a Tesseract installation.
- Removed the unused `numbering.py` module.
- Added a test suite, ruff and mypy configuration, a LICENSE file, and CI.
