"""Editable mode: build a reflowing DOCX from the PDF structure."""

import copy
from dataclasses import replace
from typing import Any, List, Optional, Set, Tuple

import pymupdf
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.text import WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from . import flow_extract
from .errors import InvalidInputError
from .files import PathLike
from .flow_docx import Emitter, FontBook, twips, write_running
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
from .flow_model import (
    Block,
    Budget,
    FlowLimits,
    PageData,
    Para,
    Rect,
    Running,
    RunningImage,
    RunningItem,
    TableItem,
)
from .flow_page import Look, Section, plan_pages
from .flow_text import clean_text, vocabulary
from .log import logger, progress

_MIN_MARGIN = 36.0
_MIN_LINE = 0.5
_MIN_HYPHENS = 2
_HYPHEN_SHARE = 0.005
_MAX_TITLE = 255
_MAX_PAGE_POINTS = 1584.0
_COLS_SUCCESSORS = (
    "w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg", "w:textDirection", "w:bidi",
    "w:rtlGutter", "w:docGrid", "w:printerSettings", "w:sectPrChange",
)


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
    blocks, instances, size, running = _build_blocks(pages, toc, budget.limits)
    _write(blocks, instances, pages, size, title, running, docx_path)
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
            logger.warning(
                "Page %d has no text (scanned or blank); only its images are kept", index + 1
            )
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
    pages: List[PageData], toc: List[OutlineEntry], limits: FlowLimits
) -> Tuple[List[Block], List[ListInstance], float, Running]:
    running = drop_running(pages)
    size = body_size(pages)
    vocab: Set[str] = vocabulary(_line_groups(pages))
    blocks: List[Block] = []
    for page in pages:
        try:
            blocks.extend(_page_blocks(page, size, vocab, limits.max_marker_lookahead))
        except Exception as exc:
            logger.warning("Skipped page %d: %s", page.number + 1, exc)
            logger.debug("Page %d failure details", page.number + 1, exc_info=True)
    blocks = join_across_segments(blocks, vocab)
    assign_headings([b for b in blocks if isinstance(b, Para)], toc, size)
    instances = apply_lists(blocks)
    return merge_tables(blocks), instances, size, running


def _page_blocks(page: PageData, size: float, vocab: Set[str], lookahead: int) -> List[Block]:
    drop_backgrounds(page)
    blocks = assemble(order_page(page, size), vocab, page.width, lookahead)
    for block in blocks:
        if isinstance(block, TableItem):
            fill_cells(block, vocab, page.width)
    return blocks


def _write(
    blocks: List[Block], instances: List[ListInstance], pages: List[PageData], size: float,
    title: str, running: Running, docx_path: PathLike,
) -> None:
    doc = Document()
    plan = plan_pages(blocks, pages, size)
    sections = plan.sections or [_blank_section(pages)]
    largest = max(max(s.width, s.height) for s in sections)
    scale = min(1.0, _MAX_PAGE_POINTS / max(largest, 1.0))
    first = sections[0]
    _apply_section(doc, first, running, scale)
    if title:
        doc.core_properties.title = clean_text(title)[:_MAX_TITLE]
    paras = [b for b in blocks if isinstance(b, Para)]
    fonts = FontBook(paras, size)
    _hyphenation(doc, paras)
    fonts.register(doc)
    text_width = (first.width - first.left - first.right) * scale
    emitter = Emitter(doc, size, text_width, resolve_anchors(blocks, size), fonts)
    emitter.register_instances(instances, _list_margins(blocks, plan.looks, len(instances)))
    write_running(doc.sections[0], running, size, fonts)
    starts = {section.start: section for section in sections[1:]}
    for index, (block, look) in enumerate(zip(blocks, plan.looks)):
        section = starts.get(index)
        if section is not None:
            _close_section(doc)
            _apply_section(doc, section, running, scale)
        try:
            _emit(emitter, doc, block, _scaled(look, scale))
        except Exception as exc:
            logger.warning("Skipped a block on page %d: %s", block.page + 1, exc)
            logger.debug("Block failure details", exc_info=True)
    doc.save(str(docx_path))


def _list_margins(blocks: List[Block], looks: List[Look], count: int) -> List[float]:
    margins: List[Optional[float]] = [None] * count
    for block, look in zip(blocks, looks):
        if isinstance(block, Para) and block.num is not None and block.num[0] < count:
            if margins[block.num[0]] is None:
                margins[block.num[0]] = look.origin
    return [margin or 0.0 for margin in margins]


def _hyphenation(doc: Any, paras: List[Para]) -> None:
    """Hyphenate and squeeze spaces like the typesetter when the PDF breaks words at line ends."""
    lines = sum(para.lines for para in paras)
    if sum(para.hyphens for para in paras) < max(_MIN_HYPHENS, _HYPHEN_SHARE * lines):
        return
    settings = doc.settings.element
    compat = settings.find(qn("w:compat"))
    if compat is not None:
        compat.insert(0, OxmlElement("w:wpJustification"))
    anchor = settings.find(qn("w:defaultTabStop"))
    element = OxmlElement("w:autoHyphenation")
    if anchor is not None:
        anchor.addnext(element)
    else:
        settings.insert(0, element)


def _blank_section(pages: List[PageData]) -> Section:
    page = pages[0] if pages else PageData(0, 595.0, 842.0)
    return Section(0, page.width, page.height, _MIN_MARGIN, _MIN_MARGIN, _MIN_MARGIN, _MIN_MARGIN)


def _scaled(look: Look, scale: float) -> Look:
    if scale == 1.0:
        return look
    return replace(
        look, before=look.before * scale, left=look.left * scale, right=look.right * scale,
        first=look.first * scale, width=look.width * scale,
    )


def _distances(running: Running, section: Section) -> Tuple[float, float]:
    tops = [_item_box(item)[1] for item in running.header]
    bottoms = [_item_box(item)[3] for item in running.footer]
    header = min(tops) if tops else min(_MIN_MARGIN, section.top / 2)
    footer = section.height - max(bottoms) if bottoms else min(_MIN_MARGIN, section.bottom / 2)
    return max(0.0, header), max(0.0, footer)


def _item_box(item: RunningItem) -> Rect:
    return item.image.bbox if isinstance(item, RunningImage) else item.box


def _close_section(doc: Any) -> None:
    """End the current section at the last paragraph, so no empty line is added for the break."""
    body = doc.element.body
    current = body.sectPr
    last = body[-2] if len(body) > 1 else None
    if last is None or last.tag != qn("w:p"):
        holder = doc.add_paragraph()
        holder.paragraph_format.line_spacing = Pt(_MIN_LINE)
        holder.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        last = holder._p
    ppr = last.get_or_add_pPr()
    ppr.insert_element_before(copy.deepcopy(current), "w:pPrChange")
    title_page = current.find(qn("w:titlePg"))
    if title_page is not None:
        current.remove(title_page)


def _apply_section(doc: Any, plan: Section, running: Running, scale: float) -> None:
    section = doc.sections[-1]
    section.start_type = WD_SECTION.NEW_PAGE if plan.new_page else WD_SECTION.CONTINUOUS
    section.orientation = (
        WD_ORIENT.LANDSCAPE if plan.width > plan.height else WD_ORIENT.PORTRAIT
    )
    section.page_width, section.page_height = Pt(plan.width * scale), Pt(plan.height * scale)
    section.left_margin, section.right_margin = Pt(plan.left * scale), Pt(plan.right * scale)
    section.top_margin, section.bottom_margin = Pt(plan.top * scale), Pt(plan.bottom * scale)
    header, footer = _distances(running, plan)
    section.header_distance, section.footer_distance = Pt(header * scale), Pt(footer * scale)
    sectpr = section._sectPr
    old = sectpr.find(qn("w:cols"))
    if old is not None:
        sectpr.remove(old)
    sectpr.insert_element_before(_columns(plan, scale), *_COLS_SUCCESSORS)
    bidi = sectpr.find(qn("w:bidi"))
    if plan.rtl and bidi is None:
        sectpr.insert_element_before(
            OxmlElement("w:bidi"), "w:rtlGutter", "w:docGrid", "w:printerSettings",
            "w:sectPrChange",
        )
    elif not plan.rtl and bidi is not None:
        sectpr.remove(bidi)


def _columns(plan: Section, scale: float) -> Any:
    cols = OxmlElement("w:cols")
    ranges = plan.columns
    if len(ranges) < 2:
        cols.set(qn("w:space"), "720")
        return cols
    ordered = sorted(ranges)
    gaps = [b[0] - a[1] for a, b in zip(ordered, ordered[1:])]
    cols.set(qn("w:num"), str(len(ranges)))
    cols.set(qn("w:space"), str(twips(min(gaps) * scale)))
    cols.set(qn("w:equalWidth"), "0")
    flow = ranges if not plan.rtl else list(reversed(ordered))
    for index, (low, high) in enumerate(flow):
        col = OxmlElement("w:col")
        col.set(qn("w:w"), str(twips((high - low) * scale)))
        if index + 1 < len(flow):
            following = flow[index + 1]
            gap = following[0] - high if following[0] >= high else low - following[1]
            col.set(qn("w:space"), str(twips(max(0.0, gap) * scale)))
        cols.append(col)
    return cols


def _emit(emitter: Emitter, doc: Any, block: Block, look: Look) -> None:
    if isinstance(block, Para):
        emitter.paragraph(doc.add_paragraph(), block, look)
    elif isinstance(block, TableItem):
        emitter.table(block, look)
    else:
        emitter.image(block, look)
