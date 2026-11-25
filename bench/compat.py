"""Import PyMuPDF under either of its module names."""

try:
    import pymupdf
except ImportError:
    import fitz as pymupdf

__all__ = ["pymupdf"]
