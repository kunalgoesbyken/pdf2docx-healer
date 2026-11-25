"""Heading styles from the PDF outline and font sizes."""

from .files import PathLike


def fix_docx_headings(docx_path: PathLike, pdf_path: PathLike, verbose: bool = False) -> int:
    """Apply Word heading styles to paragraphs that are headings in the PDF, in place."""
    return 0
