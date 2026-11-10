# Changelog

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
