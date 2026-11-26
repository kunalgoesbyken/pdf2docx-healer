"""Convert with pdf2docx page by page and rebuild any page it fails on as plain paragraphs."""

import copy
import io
from typing import Any, Dict, List, Optional, Tuple

from docx import Document
from docx.document import Document as DocumentType
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

from .log import logger, progress
from .ocr_pipeline import pymupdf

MARGIN_LIMIT_PT = 72.0
MIN_TEXT_AREA_PT = 36.0

_FLAG_SUPERSCRIPT = 1
_FLAG_ITALIC = 2
_FLAG_BOLD = 16
_BLOCK_TEXT = 0
_BLOCK_IMAGE = 1
_XML_ILLEGAL = {code: None for code in range(32) if code not in (9, 10, 13)}
_XML_ILLEGAL.update({0xFFFE: None, 0xFFFF: None})

Snapshot = Tuple[int, Any]


def convert_with_recovery(converter: Any, docx_path: str, verbose: bool) -> List[int]:
    """Run pdf2docx page by page, rebuild failed pages in place, and save the DOCX.

    A page counts as failed when pdf2docx did not finalize it or raised while writing it.
    Returns the 1-based numbers of the rebuilt pages.
    """
    settings = converter.default_settings
    converter.load_pages()
    pages = list(converter.pages)
    total = len(pages)
    reasons: Dict[int, BaseException] = {}

    try:
        converter.pages.parse(_AnnouncedDocument(converter.fitz_doc, verbose), **settings)
    except Exception as exc:
        for page in pages:
            reasons[page.id] = exc
    else:
        for position, page in enumerate(pages, start=1):
            progress(verbose, f"Parsing page {position}/{total}")
            try:
                page.parse(**settings)
            except Exception as exc:
                reasons[page.id] = exc

    document = Document()
    recovered: List[int] = []
    for position, page in enumerate(pages, start=1):
        progress(verbose, f"Writing page {position}/{total}")
        rebuild = not page.finalized
        if not rebuild:
            snapshot = _snapshot(document)
            try:
                page.make_docx(document)
            except Exception as exc:
                _restore(document, snapshot)
                reasons[page.id] = exc
                rebuild = True
        if rebuild:
            _rebuild_page(document, converter.fitz_doc, page.id, reasons.get(page.id))
            recovered.append(page.id + 1)
    document.save(docx_path)
    return recovered


def _rebuild_page(
    document: DocumentType, pdf: Any, index: int, reason: Optional[BaseException]
) -> None:
    snapshot = _snapshot(document)
    try:
        skipped = append_plain_page(document, pdf[index])
    except Exception as exc:
        _restore(document, snapshot)
        _open_section(document)
        document.add_paragraph()
        logger.error(
            "Page %d could not be converted (%s) and could not be rebuilt (%s); left it blank",
            index + 1,
            _describe(reason),
            _describe(exc),
        )
        return
    logger.warning(
        "Page %d could not be converted with its layout (%s); rebuilt it as plain text%s",
        index + 1,
        _describe(reason),
        f" ({skipped} image(s) skipped)" if skipped else "",
    )


def append_plain_page(document: DocumentType, pdf_page: Any) -> int:
    """Append a PDF page as paragraphs in reading order; return how many images were skipped."""
    blocks = [b for b in pdf_page.get_text("dict")["blocks"] if _has_content(b)]
    rect = pdf_page.rect
    section = _open_section(document)
    section.page_width = Pt(rect.width)
    section.page_height = Pt(rect.height)
    left, right, top, bottom = _margins(blocks, rect)
    section.left_margin, section.right_margin = Pt(left), Pt(right)
    section.top_margin, section.bottom_margin = Pt(top), Pt(bottom)
    text_width = max(rect.width - left - right, MIN_TEXT_AREA_PT)

    skipped = 0
    previous_bottom: Optional[float] = None
    for block in blocks:
        x0, y0, x1, y1 = block["bbox"]
        gap = 0.0 if previous_bottom is None else max(0.0, y0 - previous_bottom)
        if block["type"] == _BLOCK_TEXT:
            _add_text_block(document, block, gap)
        elif not _add_image_block(document, block, min(x1 - x0, text_width), gap):
            skipped += 1
        previous_bottom = y1
    if not blocks:
        document.add_paragraph()
    return skipped


class _AnnouncedDocument:
    """Proxy a PyMuPDF document and report each page pdf2docx opens for analysis."""

    def __init__(self, document: Any, verbose: bool) -> None:
        self._document = document
        self._verbose = verbose

    def __getitem__(self, index: int) -> Any:
        progress(self._verbose, f"Analyzing page {index + 1}/{len(self._document)}")
        return self._document[index]

    def __iter__(self) -> Any:
        return iter(self._document)

    def __len__(self) -> int:
        return len(self._document)

    def __getattr__(self, name: str) -> Any:
        return getattr(self._document, name)


def _describe(error: Optional[BaseException]) -> str:
    if error is None:
        return "pdf2docx did not produce the page"
    return f"{type(error).__name__}: {error}"


def _snapshot(document: DocumentType) -> Snapshot:
    body = document.element.body
    return len(body), copy.deepcopy(body.sectPr)


def _restore(document: DocumentType, snapshot: Snapshot) -> None:
    count, sentinel = snapshot
    body = document.element.body
    children = list(body)
    for child in children[count - 1:]:
        body.remove(child)
    body.append(sentinel)


def _open_section(document: DocumentType) -> Any:
    has_content = any(child.tag != qn("w:sectPr") for child in document.element.body)
    if has_content:
        return document.add_section(WD_SECTION.NEW_PAGE)
    return document.sections[0]


def _has_content(block: Dict[str, Any]) -> bool:
    if block["type"] == _BLOCK_IMAGE:
        return bool(block.get("image"))
    return any(
        span["text"].strip() for line in block["lines"] for span in line["spans"]
    )


def _margins(blocks: List[Dict[str, Any]], rect: Any) -> Tuple[float, float, float, float]:
    if not blocks:
        return (MARGIN_LIMIT_PT,) * 4
    left = min(b["bbox"][0] for b in blocks) - rect.x0
    top = min(b["bbox"][1] for b in blocks) - rect.y0
    right = rect.x1 - max(b["bbox"][2] for b in blocks)
    bottom = rect.y1 - max(b["bbox"][3] for b in blocks)
    left, right, top, bottom = (
        min(max(value, 0.0), MARGIN_LIMIT_PT) for value in (left, right, top, bottom)
    )
    if rect.width - left - right < MIN_TEXT_AREA_PT:
        left = right = 0.0
    return left, right, top, bottom


def _clean(text: str) -> str:
    return text.translate(_XML_ILLEGAL)


def _add_text_block(document: DocumentType, block: Dict[str, Any], gap: float) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(gap)
    paragraph.paragraph_format.space_after = Pt(0)
    first_line = True
    for line in block["lines"]:
        spans = [span for span in line["spans"] if span["text"]]
        if not spans:
            continue
        if not first_line:
            paragraph.add_run().add_break()
        first_line = False
        for span in spans:
            run = paragraph.add_run(_clean(span["text"]))
            flags = span["flags"]
            run.font.size = Pt(span["size"])
            run.font.bold = bool(flags & _FLAG_BOLD)
            run.font.italic = bool(flags & _FLAG_ITALIC)
            if flags & _FLAG_SUPERSCRIPT:
                run.font.superscript = True
            if span["color"]:
                run.font.color.rgb = RGBColor.from_string(f"{span['color']:06X}")


def _add_image_block(
    document: DocumentType, block: Dict[str, Any], width: float, gap: float
) -> bool:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(gap)
    paragraph.paragraph_format.space_after = Pt(0)
    for data in _image_candidates(block["image"]):
        try:
            paragraph.add_run().add_picture(io.BytesIO(data), width=Pt(max(width, 1.0)))
            return True
        except Exception as exc:
            logger.debug("Image could not be embedded as %s: %s", block.get("ext"), exc)
    document.element.body.remove(paragraph._p)
    return False


def _image_candidates(data: bytes) -> Any:
    yield data
    try:
        pixmap = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.Pixmap(data))
        yield pixmap.tobytes("png")
    except Exception as exc:
        logger.debug("Image could not be re-encoded: %s", exc)
