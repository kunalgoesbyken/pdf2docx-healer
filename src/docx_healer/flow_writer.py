"""Editable mode: build a reflowing DOCX from the PDF structure."""

from .files import PathLike


def write_editable_docx(pdf_path: PathLike, docx_path: PathLike, verbose: bool = False) -> None:
    """Write a reflowing DOCX with headings, lists, tables, images and links."""
    raise NotImplementedError("editable mode is not implemented yet")
