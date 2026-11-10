"""OCR pipeline for scanned/image-based PDF pages."""

import importlib
import os
import tempfile
from typing import Any, Optional

from .errors import OcrUnavailableError
from .files import PathLike, remove_quietly
from .log import logger, progress

DEFAULT_MIN_TEXT_CHARS = 50


def _load_pymupdf() -> Any:
    try:
        return importlib.import_module("pymupdf")
    except ImportError:
        return importlib.import_module("fitz")


pymupdf = _load_pymupdf()


def is_scanned_page(page: Any, min_text_chars: int = DEFAULT_MIN_TEXT_CHARS) -> bool:
    """Return True when the page has too little embedded text to be text-based."""
    if len(page.get_text("text").strip()) < min_text_chars:
        return True

    text_blocks = [b for b in page.get_text("rawdict").get("blocks", []) if b.get("type") == 0]
    if not text_blocks:
        return True

    total_chars = sum(
        len(span.get("chars", []))
        for block in text_blocks
        for line in block.get("lines", [])
        for span in line.get("spans", [])
    )
    return total_chars < min_text_chars


def needs_ocr(
    doc: Any,
    threshold: float = 0.3,
    min_text_chars: int = DEFAULT_MIN_TEXT_CHARS,
) -> bool:
    """Return True when the fraction of scanned pages exceeds ``threshold``."""
    if len(doc) == 0:
        return False
    scanned = sum(1 for page in doc if is_scanned_page(page, min_text_chars))
    return scanned / len(doc) > threshold


def require_tesseract() -> None:
    """Raise OcrUnavailableError unless PyMuPDF can locate a Tesseract install."""
    get_tessdata = getattr(pymupdf, "get_tessdata", None)
    if get_tessdata is None:
        return
    try:
        get_tessdata()
    except Exception as exc:
        raise OcrUnavailableError(
            "Tesseract OCR is not available. Install Tesseract and make sure it is on PATH "
            f"or that TESSDATA_PREFIX points to its tessdata folder ({exc})."
        ) from exc


def _ocr_page_bytes(page: Any, language: str, dpi: int) -> Optional[bytes]:
    try:
        pixmap = page.get_pixmap(dpi=dpi)
        return pixmap.pdfocr_tobytes(language=language, compress=True)
    except Exception as exc:
        logger.warning(
            "OCR failed for page %d, keeping it without a text layer: %s", page.number + 1, exc
        )
        return None


def _append_page(
    target: Any, source: Any, index: int, language: str, dpi: int, verbose: bool
) -> None:
    page = source[index]
    if is_scanned_page(page):
        progress(verbose, f"OCR page {index + 1}/{len(source)}...")
        ocr_bytes = _ocr_page_bytes(page, language, dpi)
        if ocr_bytes is not None:
            with pymupdf.open(stream=ocr_bytes, filetype="pdf") as page_doc:
                target.insert_pdf(page_doc)
            return
    target.insert_pdf(source, from_page=index, to_page=index)


def ocr_pdf(doc: Any, language: str = "eng", dpi: int = 300, verbose: bool = False) -> Any:
    """Return a new document in which every scanned page of ``doc`` has a text layer."""
    result = pymupdf.open()
    try:
        for index in range(len(doc)):
            _append_page(result, doc, index, language, dpi, verbose)
    except BaseException:
        result.close()
        raise
    return result


def ocr_pdf_file(
    pdf_path: PathLike, language: str = "eng", dpi: int = 300, verbose: bool = False
) -> str:
    """OCR ``pdf_path`` into a newly created temporary PDF and return its path.

    The returned file is always created by this call and is owned by the caller,
    who is responsible for deleting it. The input file is never returned.
    """
    progress(verbose, f"Running OCR (lang={language}, dpi={dpi})...")
    handle, tmp_path = tempfile.mkstemp(prefix="docx_healer_ocr_", suffix=".pdf")
    os.close(handle)
    try:
        with pymupdf.open(os.fspath(pdf_path)) as source:
            result = ocr_pdf(source, language=language, dpi=dpi, verbose=verbose)
        try:
            result.save(tmp_path, garbage=4, deflate=True)
        finally:
            result.close()
    except BaseException:
        remove_quietly(tmp_path)
        raise
    progress(verbose, f"OCR complete: {tmp_path}")
    return tmp_path


def prepare_ocr_copy(
    pdf_path: PathLike,
    language: str = "eng",
    dpi: int = 300,
    threshold: float = 0.3,
    verbose: bool = False,
) -> Optional[str]:
    """Create an OCR'd temporary copy when the PDF needs it, otherwise return None.

    A non-None result is always a file created here that the caller must delete.
    """
    with pymupdf.open(os.fspath(pdf_path)) as source:
        if not needs_ocr(source, threshold=threshold):
            return None
    try:
        require_tesseract()
    except OcrUnavailableError as exc:
        logger.warning("Skipping OCR of a scanned PDF: %s", exc)
        return None
    return ocr_pdf_file(pdf_path, language=language, dpi=dpi, verbose=verbose)
