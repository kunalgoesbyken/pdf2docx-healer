"""Internal PDF links (citations, cross references, outline) as Word bookmarks."""

import os
from typing import Any, Dict, Iterator, List, NamedTuple, Optional, Set, Tuple

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .docx_walk import fold_char, paged_body_elements
from .files import PathLike
from .link_injector import LinkSpan, is_plain_text_run, link_spans_in_run, text_run
from .log import logger, progress
from .ocr_pipeline import pymupdf

MAX_INTERNAL_LINKS = 20000
MAX_SEARCH_WORK = 4_000_000_000
MIN_CONTEXT_CHARS = 8
CONTEXT_RADII = (0, 3, 6, 10, 16, 24, 40, 64, 96)
DESTINATION_TOLERANCE = 1.0
BOOKMARK_PREFIX = "pdf_dest_"

_Slot = Optional[Tuple[Any, int, bool]]
_Box = Tuple[float, float, float, float]


class _BudgetExceeded(Exception):
    pass


class _Link(NamedTuple):
    source: int
    rect: _Box
    target: int
    point: Tuple[float, float]


class _PdfLine(NamedTuple):
    bbox: _Box
    block: int
    start: int
    end: int


class _PdfBlock(NamedTuple):
    text: str
    boxes: List[_Box]
    starts: List[bool]


class _PdfPage(NamedTuple):
    blocks: List[_PdfBlock]
    lines: List[_PdfLine]


class _Anchor(NamedTuple):
    block: int
    start: int
    end: int


class _Scope:
    """Folded text of the paragraphs one PDF page maps to, with paragraph start offsets."""

    def __init__(self, paragraphs: List[Any], budget: "_Budget") -> None:
        self.paragraphs = paragraphs
        self.starts: List[int] = []
        parts = []
        offset = 0
        for paragraph in paragraphs:
            text = "".join(piece for piece, _ in _paragraph_slots(paragraph))
            self.starts.append(offset)
            parts.append(text)
            offset += len(text) + 1
        self.text = "\n".join(parts)
        self.budget = budget

    def find_unique(self, window: str) -> Optional[int]:
        """Return the offset of ``window`` when it occurs exactly once, else None."""
        self.budget.spend(len(self.text))
        first = self.text.find(window)
        if first < 0:
            return None
        self.budget.spend(len(self.text) - first)
        return first if self.text.find(window, first + 1) < 0 else -1

    def paragraph_at(self, offset: int) -> Tuple[int, int]:
        low, high = 0, len(self.starts) - 1
        while low < high:
            middle = (low + high + 1) // 2
            if self.starts[middle] <= offset:
                low = middle
            else:
                high = middle - 1
        return low, offset - self.starts[low]


class _Budget:
    def __init__(self) -> None:
        self.spent = 0

    def spend(self, amount: int) -> None:
        self.spent += amount
        if self.spent > MAX_SEARCH_WORK:
            raise _BudgetExceeded()


def fix_docx_internal_links(docx_path: PathLike, pdf_path: PathLike, verbose: bool = False) -> int:
    """Recreate the PDF's internal links as bookmark hyperlinks, in place.

    Each PDF link is recreated only when both its anchor text and its destination
    paragraph are found with a unique, long enough surrounding context; text that
    is already inside a hyperlink is never touched. Returns the number of links added.
    """
    with pymupdf.open(os.fspath(pdf_path)) as pdf:
        links = _collect_links(pdf)
        if not links:
            return 0
        if len(links) > MAX_INTERNAL_LINKS:
            logger.warning(
                "Skipped internal links: %d links exceed the limit of %d",
                len(links),
                MAX_INTERNAL_LINKS,
            )
            return 0
        doc = Document(str(docx_path))
        pages = paged_body_elements(doc)
        trusted = len(pages) == pdf.page_count
        scopes = _ScopeCache(doc, pages, trusted)
        try:
            destinations = _resolve_destinations(pdf, links, scopes)
            placements = _place_anchors(pdf, links, destinations, scopes)
        except _BudgetExceeded:
            logger.warning(
                "Skipped internal links: text search exceeded the work limit of %d", MAX_SEARCH_WORK
            )
            return 0
    if not placements:
        return 0
    _apply(doc, placements)
    doc.save(str(docx_path))
    progress(verbose, f"Added {len(placements)} internal link(s)")
    return len(placements)


def _collect_links(pdf: Any) -> List[_Link]:
    """Return the internal links with their destination as a point in top left page coordinates.

    PyMuPDF reports ``to`` in page coordinates for ``LINK_GOTO`` but in raw PDF coordinates
    (origin bottom left) for ``LINK_NAMED``, which comes from ``Document.resolve_names``.
    """
    links = []
    for number, page in enumerate(pdf):
        for link in page.get_links():
            if link["kind"] not in (pymupdf.LINK_GOTO, pymupdf.LINK_NAMED):
                continue
            target, point = link.get("page", -1), link.get("to")
            if not 0 <= target < pdf.page_count or not isinstance(point, pymupdf.Point):
                continue
            if link["kind"] == pymupdf.LINK_NAMED:
                point = pymupdf.Point(point) * pdf[target].transformation_matrix
            rect = link["from"]
            links.append(
                _Link(number, (rect.x0, rect.y0, rect.x1, rect.y1), target, (point.x, point.y))
            )
    return links


class _ScopeCache:
    def __init__(self, doc: Any, pages: List[List[Any]], trusted: bool) -> None:
        self.doc = doc
        self.pages = pages
        self.trusted = trusted
        self.budget = _Budget()
        self.cache: Dict[int, _Scope] = {}

    def get(self, page: int) -> _Scope:
        key = page if self.trusted else -1
        scope = self.cache.get(key)
        if scope is None:
            if self.trusted:
                paragraphs = [p for element in self.pages[page] for p in element.iter(qn("w:p"))]
            else:
                paragraphs = list(self.doc.element.body.iter(qn("w:p")))
            scope = self.cache[key] = _Scope(paragraphs, self.budget)
        return scope


def _paragraph_slots(paragraph: Any) -> Iterator[Tuple[str, _Slot]]:
    """Yield each folded character of a paragraph with the run character it came from.

    Text outside plain direct runs (existing hyperlinks, fields, drawings) has no slot,
    so it takes part in matching but can never be linked.
    """
    for child in paragraph:
        if child.tag == qn("w:r") and is_plain_text_run(child):
            for index, char in enumerate(child.text):
                for part, piece in enumerate(fold_char(char)):
                    yield piece, (child, index, part == 0)
        elif child.tag != qn("w:pPr"):
            for node in child.iter(qn("w:t")):
                for char in node.text or "":
                    for piece in fold_char(char):
                        yield piece, None


def _read_page(page: Any) -> _PdfPage:
    blocks: List[_PdfBlock] = []
    lines: List[_PdfLine] = []
    for block in page.get_text("rawdict")["blocks"]:
        if block["type"] != 0:
            continue
        text: List[str] = []
        boxes: List[_Box] = []
        starts: List[bool] = []
        for line in block["lines"]:
            begin = len(text)
            for span in line["spans"]:
                for char in span["chars"]:
                    for part, piece in enumerate(fold_char(char["c"])):
                        text.append(piece)
                        boxes.append(tuple(char["bbox"]))
                        starts.append(part == 0)
            if len(text) > begin:
                lines.append(_PdfLine(tuple(line["bbox"]), len(blocks), begin, len(text)))
        if text:
            blocks.append(_PdfBlock("".join(text), boxes, starts))
    return _PdfPage(blocks, lines)


def _windows(start: int, end: int, size: int) -> List[Tuple[int, int]]:
    """Context windows around ``[start, end)``, shortest first and long enough to be distinctive."""
    edges = sorted({*CONTEXT_RADII, size})
    shortest = min(MIN_CONTEXT_CHARS, size)
    windows = {(max(0, start - left), min(size, end + right)) for left in edges for right in edges}
    return sorted((w for w in windows if w[1] - w[0] >= shortest), key=lambda w: (w[1] - w[0], w))


def _locate(scope: _Scope, block: _PdfBlock, start: int, end: int) -> Optional[Tuple[int, int]]:
    """Find ``block[start:end]`` in the scope as (paragraph index, offset in that paragraph).

    The text is widened with its surroundings in the block, on either side, until it occurs
    exactly once; windows that are missing (a PDF block can span several paragraphs) rule out every
    larger window, ambiguous ones are widened, and None is returned when none is unique.
    """
    missing: List[Tuple[int, int]] = []
    for low, high in _windows(start, end, len(block.text)):
        if any(low <= gone_low and gone_high <= high for gone_low, gone_high in missing):
            continue
        found = scope.find_unique(block.text[low:high])
        if found is None:
            missing.append((low, high))
        elif found >= 0:
            paragraph, offset = scope.paragraph_at(found)
            return paragraph, offset + (start - low)
    return None


def _destination_line(page: _PdfPage, point: Tuple[float, float]) -> Optional[_PdfLine]:
    x, y = point
    below = [line for line in page.lines if line.bbox[3] >= y - DESTINATION_TOLERANCE]
    inside = [
        line
        for line in below
        if line.bbox[0] - DESTINATION_TOLERANCE <= x <= line.bbox[2] + DESTINATION_TOLERANCE
    ]
    return min(inside or below, key=lambda line: (line.bbox[1], line.bbox[0]), default=None)


def _resolve_destinations(
    pdf: Any, links: List[_Link], scopes: _ScopeCache
) -> Dict[Tuple[int, float, float], Any]:
    wanted: Dict[int, Set[Tuple[float, float]]] = {}
    for link in links:
        wanted.setdefault(link.target, set()).add(link.point)
    resolved: Dict[Tuple[int, float, float], Any] = {}
    for number in sorted(wanted):
        page = _read_page(pdf[number])
        scope = scopes.get(number)
        for point in wanted[number]:
            line = _destination_line(page, point)
            found = _locate(scope, page.blocks[line.block], line.start, line.end) if line else None
            if found is not None:
                key = (number, *point)
                resolved[key] = scope.paragraphs[found[0]]
    return resolved


def _anchor_chars(page: _PdfPage, rect: _Box) -> Optional[_Anchor]:
    hits: List[Tuple[int, int]] = []
    for line in page.lines:
        bbox = line.bbox
        if bbox[0] > rect[2] or bbox[2] < rect[0] or bbox[1] > rect[3] or bbox[3] < rect[1]:
            continue
        boxes = page.blocks[line.block].boxes
        for index in range(line.start, line.end):
            x0, y0, x1, y1 = boxes[index]
            if rect[0] <= (x0 + x1) / 2 <= rect[2] and rect[1] <= (y0 + y1) / 2 <= rect[3]:
                hits.append((line.block, index))
    if not hits or len({block for block, _ in hits}) != 1:
        return None
    block, first = hits[0]
    last = hits[-1][1]
    starts = page.blocks[block].starts
    contiguous = sorted(index for _, index in hits) == list(range(first, last + 1))
    on_edges = starts[first] and (last + 1 == len(starts) or starts[last + 1])
    return _Anchor(block, first, last + 1) if contiguous and on_edges else None


class _Placement(NamedTuple):
    destination: Any
    ranges: List[Tuple[Any, int, int]]


def _place_anchors(
    pdf: Any,
    links: List[_Link],
    destinations: Dict[Tuple[int, float, float], Any],
    scopes: _ScopeCache,
) -> List[_Placement]:
    by_page: Dict[int, List[_Link]] = {}
    for link in links:
        if (link.target, *link.point) in destinations:
            by_page.setdefault(link.source, []).append(link)
    claimed: Dict[Any, Set[int]] = {}
    placements: List[_Placement] = []
    for number in sorted(by_page):
        page = _read_page(pdf[number])
        scope = scopes.get(number)
        slot_cache: Dict[int, List[_Slot]] = {}
        for link in by_page[number]:
            anchor = _anchor_chars(page, link.rect)
            if anchor is None:
                continue
            located = _locate(scope, page.blocks[anchor.block], anchor.start, anchor.end)
            if located is None:
                continue
            index, offset = located
            if index not in slot_cache:
                slot_cache[index] = [slot for _, slot in _paragraph_slots(scope.paragraphs[index])]
            ranges = _run_ranges(slot_cache[index], offset, anchor.end - anchor.start)
            if ranges is None or _overlaps_claimed(claimed, ranges):
                continue
            for run, start, end in ranges:
                claimed.setdefault(run, set()).update(range(start, end))
            key = (link.target, *link.point)
            placements.append(_Placement(destinations[key], ranges))
    return placements


def _overlaps_claimed(claimed: Dict[Any, Set[int]], ranges: List[Tuple[Any, int, int]]) -> bool:
    return any(
        not claimed.get(run, set()).isdisjoint(range(start, end)) for run, start, end in ranges
    )


def _run_ranges(
    slots: List[_Slot], offset: int, length: int
) -> Optional[List[Tuple[Any, int, int]]]:
    """Turn a character range into per-run ranges, or None when any part is not linkable."""
    chosen = [slot for slot in slots[offset : offset + length] if slot is not None]
    if len(chosen) != length:
        return None
    following = slots[offset + length] if offset + length < len(slots) else None
    if not chosen[0][2] or (following is not None and not following[2]):
        return None
    ranges: List[Tuple[Any, int, int]] = []
    for run, index, _ in chosen:
        if ranges and ranges[-1][0] is run:
            ranges[-1] = (run, ranges[-1][1], index + 1)
        else:
            ranges.append((run, index, index + 1))
    return ranges


def _next_bookmark_ids(doc: Any) -> Tuple[Set[str], int]:
    names: Set[str] = set()
    highest = -1
    for start in doc.element.body.iter(qn("w:bookmarkStart")):
        names.add(start.get(qn("w:name")) or "")
        identifier = start.get(qn("w:id")) or ""
        if identifier.isdigit():
            highest = max(highest, int(identifier))
    return names, highest + 1


def _add_bookmark(paragraph: Any, name: str, identifier: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(identifier))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(identifier))
    properties = paragraph.find(qn("w:pPr"))
    paragraph.insert(0 if properties is None else paragraph.index(properties) + 1, start)
    paragraph.append(end)


def _anchor_link(text: str, rpr: Optional[Any], name: str) -> Any:
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), name)
    hyperlink.set(qn("w:history"), "1")
    hyperlink.append(text_run(text, rpr))
    return hyperlink


def _apply(doc: Any, placements: List[_Placement]) -> None:
    taken, identifier = _next_bookmark_ids(doc)
    names: Dict[Any, str] = {}
    spans: Dict[Any, List[LinkSpan]] = {}
    serial = 0
    for placement in placements:
        name = names.get(placement.destination)
        if name is None:
            serial += 1
            while f"{BOOKMARK_PREFIX}{serial}" in taken:
                serial += 1
            name = names[placement.destination] = f"{BOOKMARK_PREFIX}{serial}"
            _add_bookmark(placement.destination, name, identifier)
            identifier += 1
        for run, start, end in placement.ranges:
            spans.setdefault(run, []).append(LinkSpan(start, end, name))
    for run, run_spans in spans.items():
        link_spans_in_run(run, sorted(run_spans), _anchor_link)
