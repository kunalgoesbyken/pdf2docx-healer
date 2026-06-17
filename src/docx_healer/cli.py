"""Command-line interface for pdf2docx-healer."""

import argparse
import sys
from pathlib import Path
from .converter import heal, HealerConfig


def main():
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
                        help="Fraction of textless pages to trigger OCR (default: 0.3)")
    parser.add_argument("--no-lists", action="store_true",
                        help="Skip bullet/numbered list detection and formatting")
    parser.add_argument("--no-hyperlinks", action="store_true",
                        help="Skip hyperlink extraction")
    parser.add_argument("--no-font-fix", action="store_true",
                        help="Skip CJK font fallback mapping")
    parser.add_argument("--aggressive", action="store_true",
                        help="Use aggressive paragraph splitting for list detection")
    parser.add_argument("-q", "--quiet", action="store_true",
                        help="Suppress progress output")

    args = parser.parse_args()

    if not args.pdf.exists():
        print(f"Error: PDF file not found: {args.pdf}", file=sys.stderr)
        sys.exit(1)

    output = args.output or args.pdf.with_suffix(".docx")

    config = HealerConfig(
        ocr_enabled=args.ocr,
        ocr_lang=args.ocr_lang,
        ocr_dpi=args.ocr_dpi,
        ocr_threshold=args.ocr_threshold,
        fix_lists=not args.no_lists,
        fix_hyperlinks=not args.no_hyperlinks,
        fix_fonts=not args.no_font_fix,
        aggressive_lists=args.aggressive,
        verbose=not args.quiet,
    )

    heal(str(args.pdf), str(output), config=config)


if __name__ == "__main__":
    main()
