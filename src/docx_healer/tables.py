"""Table pass: pdf2docx tables are kept as they are because layout cannot be told from data."""

from .files import PathLike


def fix_docx_tables(docx_path: PathLike, pdf_path: PathLike, verbose: bool = False) -> int:
    """Leave every table untouched and return 0.

    The DOCX borders say nothing about layout versus data, and PyMuPDF find_tables misses
    borderless data tables, so no rule flattens layout tables without also flattening data.
    """
    return 0
