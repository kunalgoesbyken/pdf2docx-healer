"""Command-line interface for pdf2docx-healer."""

import argparse
import sys
from pathlib import Path
from typing import List, Optional

from .converter import HealerConfig, heal
from .errors import InvalidInputError
from .log import logger

EXIT_OK = 0
EXIT_FAILURE = 1
EXIT_USAGE = 2


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pdf2docx-heal",
        description="Convert PDF to DOCX preserving bullet lists, hyperlinks, and CJK fonts.",
    )
    parser.add_argument("pdf", type=Path, help="Input PDF file path")
    parser.add_argument("-o", "--output", type=Path, default=None,
                        help="Output DOCX file path (default: input filename with .docx extension)")
    parser.add_argument("--ocr", action="store_true",
                        help="Enable OCR for scanned/image-based PDFs")
    parser.add_argument("--ocr-lang", default="eng",
                        help="Tesseract language code for OCR (default: eng)")
    parser.add_argument("--ocr-dpi", type=int, default=300,
                        help="OCR resolution in DPI (default: 300)")
    parser.add_argument("--ocr-threshold", type=float, default=0.3,
                        help="Fraction of textless pages (0-1) above which OCR runs (default: 0.3)")
    parser.add_argument("--no-lists", action="store_true",
                        help="Skip bullet/numbered list detection and formatting")
    parser.add_argument("--no-hyperlinks", action="store_true",
                        help="Skip hyperlink extraction")
    parser.add_argument("--no-font-fix", action="store_true",
                        help="Skip CJK font fallback mapping")
    parser.add_argument("--no-headings", action="store_true",
                        help="Skip heading styles from the PDF outline and font sizes")
    parser.add_argument("--no-internal-links", action="store_true",
                        help="Skip internal links (citations, cross references)")
    parser.add_argument("--no-table-filter", action="store_true",
                        help="Keep every table pdf2docx proposes, including text layout")
    parser.add_argument("--no-page-recovery", action="store_true",
                        help="Do not rebuild pages that pdf2docx fails to convert")
    parser.add_argument("--mode", choices=("fidelity", "editable"), default="fidelity",
                        help="fidelity keeps the page layout (pdf2docx); editable builds a "
                             "reflowing document with real headings, lists and tables")
    parser.add_argument("--aggressive", action="store_true",
                        help="Use aggressive paragraph splitting for list detection")
    parser.add_argument("-q", "--quiet", action="store_true",
                        help="Suppress progress output")
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        config = HealerConfig(
            ocr_enabled=args.ocr,
            ocr_lang=args.ocr_lang,
            ocr_dpi=args.ocr_dpi,
            ocr_threshold=args.ocr_threshold,
            fix_lists=not args.no_lists,
            fix_hyperlinks=not args.no_hyperlinks,
            fix_fonts=not args.no_font_fix,
            fix_headings=not args.no_headings,
            fix_internal_links=not args.no_internal_links,
            fix_tables=not args.no_table_filter,
            recover_pages=not args.no_page_recovery,
            mode=args.mode,
            aggressive_lists=args.aggressive,
            verbose=not args.quiet,
        )
    except InvalidInputError as exc:
        parser.error(str(exc))

    try:
        heal(args.pdf, args.output, config=config)
    except Exception as exc:
        logger.debug("Conversion failed", exc_info=True)
        print(f"Error: {exc}", file=sys.stderr)
        return EXIT_FAILURE
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
