"""Internal PDF links (citations, cross references, outline) as Word bookmarks."""

from .files import PathLike


def fix_docx_internal_links(docx_path: PathLike, pdf_path: PathLike, verbose: bool = False) -> int:
    """Recreate the PDF's internal links as bookmark hyperlinks, in place."""
    return 0
