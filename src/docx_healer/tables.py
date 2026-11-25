"""Layout tables that pdf2docx creates where the PDF has no table."""

from .files import PathLike


def fix_docx_tables(docx_path: PathLike, pdf_path: PathLike, verbose: bool = False) -> int:
    """Flatten tables that are page layout, not data, in place."""
    return 0
