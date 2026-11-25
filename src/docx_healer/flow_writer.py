"""Editable mode: build a reflowing DOCX from the PDF structure."""

from typing import Any, List, Optional, Set, Tuple

import pymupdf
from docx import Document
from docx.shared import Pt

from . import flow_extract
from .errors import InvalidInputError
from .files import PathLike
from .flow_docx import Emitter
from .flow_headings import OutlineEntry, assign_headings
from .flow_layout import (
    assemble,
    body_size,
    drop_backgrounds,
    drop_running,
    fill_cells,
    join_across_segments,
    merge_tables,
    order_page,
)
from .flow_links import resolve_anchors
from .flow_lists import ListInstance, apply_lists
from .flow_model import Block, Budget, FlowLimits, PageData, Para, TableItem
from .flow_text import clean_text, vocabulary
from .log import logger, progress

_MIN_MARGIN = 36.0
_MAX_MARGIN = 144.0
_CENTER_SHARE = 0.05
_MAX_TITLE = 255
_WIDE_IMAGE = 0.9


def write_editable_docx(
    pdf_path: PathLike, docx_path: PathLike, verbose: bool = False,
    limits: Optional[FlowLimits] = None,
) -> None:
    """Write a reflowing DOCX with headings, lists, tables, images and links."""
    budget = Budget(limits or FlowLimits())
    with pymupdf.open(str(pdf_path)) as pdf:
        if pdf.needs_pass:
            raise InvalidInputError("The PDF is password protected")
        pages = _extract_pages(pdf, budget, verbose)
        toc = _outline(pdf, budget.limits.max_outline_entries)
        title = (pdf.metadata or {}).get("title") or ""
    blocks, instances, size = _build_blocks(pages, toc)
    _write(blocks, instances, pages, size, title, docx_path)
    progress(verbose, f"Editable document written ({len(blocks)} blocks)")


def _extract_pages(pdf: Any, budget: Budget, verbose: bool) -> List[PageData]:
    limit = budget.limits.max_pages
    if len(pdf) > limit:
        logger.warning(
            "The PDF has %d pages; only the first %d are converted (max_pages)", len(pdf), limit
        )
    pages: List[PageData] = []
    for index in range(min(len(pdf), limit)):
        try:
            data = flow_extract.extract_page(pdf[index], budget)
        except Exception as exc:
            logger.warning("Skipped page %d: %s", index + 1, exc)
            logger.debug("Page %d failure details", index + 1, exc_info=True)
            continue
        if not data.has_text:
            logger.warning("Page %d has no text (scanned or blank), skipped", index + 1)
            data.images = []
        pages.append(data)
        progress(verbose and index % 25 == 0, f"Read page {index + 1}")
    return pages


def _outline(pdf: Any, limit: int) -> List[OutlineEntry]:
    try:
        raw = pdf.get_toc(simple=False)
    except Exception as exc:
        logger.debug("Could not read the outline: %s", exc)
        return []
    if len(raw) > limit:
        logger.warning(
            "The outline has %d entries; only the first %d are used (max_outline_entries)",
            len(raw), limit,
        )
    entries = []
    for entry in raw[:limit]:
        top = None
        dest = entry[3] if len(entry) > 3 and isinstance(entry[3], dict) else {}
        try:
            if dest.get("nameddest"):
                top = float(pdf.resolve_link("#" + dest["nameddest"])[2])
            elif dest.get("to") is not None:
                top = float(dest["to"].y)
        except Exception:
            top = None
        entries.append((int(entry[0]), str(entry[1]), int(entry[2]), top))
    return entries


def _line_groups(pages: List[PageData]) -> List[List[str]]:
    groups: List[List[str]] = []
    for page in pages:
        groups.extend([line.text for line in unit.lines] for unit in page.units)
        for table in page.tables:
            groups.extend([line.text for line in cell.lines] for cell in table.cells)
    return groups


def _build_blocks(
    pages: List[PageData], toc: List[OutlineEntry]
) -> Tuple[List[Block], List[ListInstance], float]:
    drop_running(pages)
    size = body_size(pages)
    vocab: Set[str] = vocabulary(_line_groups(pages))
    blocks: List[Block] = []
    for page in pages:
        try:
            blocks.extend(_page_blocks(page, size, vocab))
        except Exception as exc:
            logger.warning("Skipped page %d: %s", page.number + 1, exc)
            logger.debug("Page %d failure details", page.number + 1, exc_info=True)
    blocks = join_across_segments(blocks, vocab)
    assign_headings([b for b in blocks if isinstance(b, Para)], toc, size)
    instances = apply_lists(blocks)
    return merge_tables(blocks), instances, size


def _page_blocks(page: PageData, size: float, vocab: Set[str]) -> List[Block]:
    drop_backgrounds(page)
    blocks = assemble(order_page(page, size), vocab, page.width)
    for block in blocks:
        if isinstance(block, TableItem):
            fill_cells(block, vocab, page.width)
    return blocks


def _margins(blocks: List[Block], page: PageData) -> Tuple[float, float]:
    boxes = [b.bbox for b in blocks if isinstance(b, (Para, TableItem))]
    lefts = [box[0] for box in boxes]
    rights = [box[2] for box in boxes]
    left = min(lefts, default=_MIN_MARGIN)
    right = page.width - max(rights, default=page.width - _MIN_MARGIN)
    return _clamp(left), _clamp(right)


def _clamp(value: float) -> float:
    return max(_MIN_MARGIN, min(_MAX_MARGIN, value))


def _write(
    blocks: List[Block], instances: List[ListInstance], pages: List[PageData], size: float,
    title: str, docx_path: PathLike,
) -> None:
    doc = Document()
    page = pages[0] if pages else PageData(0, 595.0, 842.0)
    left, right = _margins(blocks, page)
    section = doc.sections[0]
    section.page_width, section.page_height = Pt(page.width), Pt(page.height)
    section.left_margin, section.right_margin = Pt(left), Pt(right)
    section.top_margin = section.bottom_margin = Pt(_MIN_MARGIN * 1.5)
    if title:
        doc.core_properties.title = clean_text(title)[:_MAX_TITLE]
    text_width = page.width - left - right
    anchors = resolve_anchors(blocks, size)
    emitter = Emitter(doc, size, text_width, page.width, anchors)
    emitter.register_instances(instances, left)
    for block in blocks:
        try:
            _emit(emitter, doc, block, page.width, text_width)
        except Exception as exc:
            logger.warning("Skipped a block on page %d: %s", block.page + 1, exc)
            logger.debug("Block failure details", exc_info=True)
    doc.save(str(docx_path))


def _emit(emitter: Emitter, doc: Any, block: Block, page_width: float, text_width: float) -> None:
    if isinstance(block, Para):
        emitter.paragraph(doc.add_paragraph(), block)
    elif isinstance(block, TableItem):
        emitter.table(block)
    else:
        middle = (block.bbox[0] + block.bbox[2]) / 2
        centered = abs(middle - page_width / 2) < _CENTER_SHARE * page_width
        emitter.image(block, centered and block.width < _WIDE_IMAGE * text_width)
