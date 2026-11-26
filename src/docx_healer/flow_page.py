"""Page geometry for editable mode: sections, columns and how each block sits on the page.

Every measure comes from the PDF line boxes. Word sets a line with an exact line pitch P so
that its baseline lies 0.8 P below the top of the line, so a PDF line with its baseline at y
fills y - 0.8 P to y + 0.2 P of the Word flow. Gaps between blocks become space before.
"""

from dataclasses import dataclass, field
from statistics import median
from typing import Any, Dict, List, Optional, Tuple

from .flow_model import Block, PageData, Para, Rect, TableItem

Range = Tuple[float, float]

_ASCENT = 0.8
_MIN_PITCH = 0.5
_MAX_PITCH = 3.0
_DEFAULT_RATIO = 1.2
_TOLERANCE = 0.3
_MIN_TOLERANCE = 1.5
_NARROW = 2.0
_BOTTOM_ROOM = 18.0
_FRAME_ROOM = 1.15
_BESIDE = 0.5
_FRAME_WIDTH = 1.05
_SAME = 2.0
_MIN_TEXT_WIDTH = 72.0


@dataclass
class Look:
    """How one block is placed: alignment, indents, space before and line pitch, in points."""

    align: str = "left"
    before: float = 0.0
    left: float = 0.0
    right: float = 0.0
    first: float = 0.0
    pitch: float = 0.0
    width: float = 0.0
    page_break: bool = False
    column_break: bool = False
    origin: float = 0.0
    frame: Optional[Rect] = None
    turn: str = ""


@dataclass
class Section:
    """A run of blocks sharing page size, margins and columns, in points."""

    start: int
    width: float
    height: float
    left: float
    right: float
    top: float
    bottom: float
    columns: List[Range] = field(default_factory=list)
    new_page: bool = True
    rtl: bool = False


@dataclass
class Plan:
    looks: List[Look]
    sections: List[Section]


@dataclass
class _Frame:
    left: float
    right: float
    top: float
    bottom: float


def _seg(block: Block) -> Tuple[Any, ...]:
    return block.seg or (block.page, 0, -1)


def _column(seg: Tuple[Any, ...]) -> int:
    return seg[2] if isinstance(seg[1], int) else -1


def _end_page(block: Block) -> int:
    if isinstance(block, Para) and block.boxes:
        return int(block.boxes[-1][0][0])
    return _seg(block)[0]


def _first_boxes(para: Para) -> List[Rect]:
    boxes = [box for seg, box, _ in para.boxes if seg == para.seg]
    return boxes or [para.bbox]


def _measured_pitch(para: Para) -> Optional[float]:
    steps = [
        b[2] - a[2] for a, b in zip(para.boxes, para.boxes[1:])
        if a[0] == b[0] and _MIN_PITCH * para.size < b[2] - a[2] < _MAX_PITCH * para.size
    ]
    return median(steps) if steps else None


def _pitches(blocks: List[Block]) -> Dict[int, float]:
    measured = {id(b): _measured_pitch(b) for b in blocks if isinstance(b, Para)}
    ratios = [
        value / b.size for b in blocks
        if isinstance(b, Para) and b.size > 0 and (value := measured[id(b)])
    ]
    ratio = median(ratios) if ratios else _DEFAULT_RATIO
    return {
        id(b): measured[id(b)] or ratio * b.size for b in blocks if isinstance(b, Para)
    }


def _extent(block: Block, pitch: float) -> List[Tuple[int, Rect]]:
    """Return the page and Word flow rectangle of each line or of the whole block."""
    if isinstance(block, Para):
        return [
            (int(seg[0]), (box[0], base - _ASCENT * pitch, box[2], base + (1 - _ASCENT) * pitch))
            for seg, box, base in block.boxes
        ] or [(block.page, block.bbox)]
    return [(_seg(block)[0], block.bbox)]


def _frames(blocks: List[Block], pitches: Dict[int, float]) -> Dict[int, _Frame]:
    frames: Dict[int, _Frame] = {}
    for block in blocks:
        for page, rect in _extent(block, pitches.get(id(block), 0.0)):
            frame = frames.get(page)
            if frame is None:
                frames[page] = _Frame(rect[0], rect[2], rect[1], rect[3])
                continue
            frame.left, frame.right = min(frame.left, rect[0]), max(frame.right, rect[2])
            frame.top, frame.bottom = min(frame.top, rect[1]), max(frame.bottom, rect[3])
    return frames


def _ranges(page: PageData, frame: _Frame, used: Dict[int, Range]) -> List[Range]:
    """Return the run's columns in flow order: the page's gutters, widened to fit the content."""
    gutters = [g for g in page.gutters if frame.left < g[0] and g[1] < frame.right] if used else []
    bounds = [frame.left] + [x for gap in gutters for x in gap] + [frame.right]
    ranges = [list(pair) for pair in zip(bounds[::2], bounds[1::2])]
    for index, (low, high) in used.items():
        if 0 <= index < len(ranges):
            ranges[index][0] = min(ranges[index][0], low)
            ranges[index][1] = max(ranges[index][1], high)
    for left, right in zip(ranges, ranges[1:]):
        right[0] = max(right[0], left[0] + _MIN_TEXT_WIDTH / 2)
        left[1] = min(left[1], right[0] - _MIN_TOLERANCE)
    result = [(low, high) for low, high in ranges]
    return result[::-1] if page.rtl else result


def _section(
    start: int, page: PageData, frame: _Frame, used: Dict[int, Range], new_page: bool
) -> Section:
    right = max(frame.right, frame.left + _MIN_TEXT_WIDTH)
    return Section(
        start, page.width, page.height, max(0.0, frame.left), max(0.0, page.width - right),
        max(0.0, frame.top), max(0.0, min(page.height - frame.bottom, _BOTTOM_ROOM)),
        _ranges(page, frame, used), new_page, page.rtl,
    )


def _columns_used(
    blocks: List[Block], pitches: Dict[int, float]
) -> Dict[Tuple[Any, ...], Dict[int, Range]]:
    """Return, per page and run, the horizontal extent of the content in each column."""
    used: Dict[Tuple[Any, ...], Dict[int, Range]] = {}
    for block in blocks:
        seg = _seg(block)
        column = _column(seg)
        if column < 0:
            continue
        columns = used.setdefault((seg[0], seg[1]), {})
        for page, rect in _extent(block, pitches.get(id(block), 0.0)):
            if page == seg[0]:
                low, high = columns.get(column, (rect[0], rect[2]))
                columns[column] = (min(low, rect[0]), max(high, rect[2]))
    return used


def _fits(section: Section, other: Section, frame: _Frame) -> bool:
    """Whether a page can continue the section after a plain page break."""
    same = all(
        abs(a - b) <= _SAME for a, b in (
            (section.width, other.width), (section.height, other.height),
            (section.left, other.left), (section.right, other.right),
        )
    )
    columns = len(section.columns) == len(other.columns) and all(
        abs(a[0] - b[0]) <= _SAME and abs(a[1] - b[1]) <= _SAME
        for a, b in zip(section.columns, other.columns)
    )
    room = frame.top >= section.top - _SAME and frame.bottom <= section.height - section.bottom
    return same and columns and room and section.rtl == other.rtl


def _flow_index(section: Section, page: PageData, column: int) -> int:
    if len(section.columns) < 2 or column < 0:
        return 0
    index = len(section.columns) - 1 - column if page.rtl else column
    return max(0, min(len(section.columns) - 1, index))


def _mirror(rects: List[Rect], low: float, high: float) -> Tuple[List[Rect], float, float]:
    return [(-r[2], r[1], -r[0], r[3]) for r in rects], -high, -low


def align_lines(rects: List[Rect], low: float, high: float, size: float, rtl: bool) -> Look:
    """Find the alignment and indents that put these lines where they are in the column.

    For right-to-left text the measures are mirrored, so left means the start side, which is
    how Word reads alignment and indents of a right-to-left paragraph.
    """
    if rtl:
        rects, low, high = _mirror(rects, low, high)
    tolerance = max(_MIN_TOLERANCE, _TOLERANCE * size)
    lefts = [r[0] for r in rects]
    rights = [r[2] for r in rects]
    if len(rects) == 1:
        space_left, space_right = lefts[0] - low, high - rights[0]
        if space_left > tolerance and space_right > tolerance \
                and abs(space_left - space_right) <= 2 * tolerance:
            return Look("center")
        if space_left > tolerance and abs(space_right) <= tolerance:
            return Look("right")
        return Look("left", left=max(0.0, space_left))
    rest = lefts[1:]
    body = rights[:-1]
    indent = max(0.0, min(rest) - low)
    first = lefts[0] - low - indent
    centers = [(a + b) / 2 for a, b in zip(lefts, rights)]
    if max(centers) - min(centers) <= tolerance < max(lefts) - min(lefts):
        widest = max(b - a for a, b in zip(lefts, rights)) * _FRAME_WIDTH
        spare = max(0.0, high - low - widest)
        shift = 2 * median(centers) - low - high
        return Look(
            "center", left=max(0.0, (spare + shift) / 2), right=max(0.0, (spare - shift) / 2)
        )
    ragged_left = max(rest) - min(rest) > tolerance or lefts[1] > lefts[0] + tolerance
    if ragged_left and max(rights) - min(rights) <= tolerance:
        return Look("right", right=max(0.0, high - max(rights)))
    if max(rest) - min(rest) <= tolerance:
        edge = max(body)
        even = max(body) - min(body) <= tolerance
        if even and (len(body) > 1 or high - edge <= tolerance):
            return Look("both", left=indent, first=first, right=max(0.0, high - edge))
        spare = high - max(rights)
        right = spare if spare > _NARROW * size else 0.0
        return Look("left", left=indent, first=first, right=right)
    return Look("left", left=indent, first=first)


def _box_look(rect: Rect, low: float, high: float, size: float) -> Look:
    look = align_lines([rect], low, high, size, False)
    look.width = high - low
    if look.align == "right":
        look.left = 0.0
    return look


def _rotated(block: Block) -> bool:
    return _seg(block)[1:2] == ("rotated",)


def _framed(block: Block, look: Look) -> Look:
    """Place text running up or down the page in a frame where the PDF has it."""
    seg = _seg(block)
    turn = seg[3] if len(seg) > 3 else ""
    if not turn:
        return look
    left, top, right, bottom = block.bbox
    length = (bottom - top) * _FRAME_ROOM
    shift = length - (bottom - top) if turn == "btLr" else 0.0
    return Look(look.align, frame=(left, top - shift, right, top - shift + length), turn=turn)


def _beside(block: Block, look: Look, extent: List[Tuple[int, Rect]]) -> Look:
    """Pin a block that sits beside earlier content in the same row where the PDF has it."""
    left, top = min(r[0] for _, r in extent), extent[0][1][1]
    right, bottom = max(r[2] for _, r in extent), extent[-1][1][3]
    extra = (right - left) * (_FRAME_WIDTH - 1)
    shift = {"center": extra / 2, "right": extra}.get(look.align, 0.0)
    return Look(
        look.align, pitch=look.pitch, width=right - left + extra,
        frame=(left - shift, top, right - shift + extra, bottom),
    )


def plan_pages(blocks: List[Block], pages: List[PageData], body: float) -> Plan:
    """Work out the sections and the look of every block."""
    by_number = {page.number: page for page in pages}
    pitches = _pitches(blocks)
    frames = _frames([b for b in blocks if not _rotated(b)], pitches)
    used = _columns_used(blocks, pitches)
    looks: List[Look] = []
    sections: List[Section] = []
    cursor = start_y = bottom = 0.0
    flow = 0
    layout: Any = None
    last_page = -1
    for index, block in enumerate(blocks):
        if _rotated(block):
            looks.append(_framed(block, Look()))
            continue
        seg = _seg(block)
        page = by_number.get(seg[0]) or PageData(seg[0], 595.0, 842.0)
        frame = frames.get(seg[0]) or _Frame(0.0, page.width, 0.0, page.height)
        column = _column(seg)
        multi = column >= 0 and len(page.gutters) > 0
        key = ("columns", seg[0], seg[1]) if multi else ("single",)
        columns = used.get((seg[0], seg[1]), {}) if multi else {}
        look = Look()
        if not sections or seg[0] > last_page:
            candidate = _section(index, page, frame, columns, True)
            if sections and key == layout and _fits(sections[-1], candidate, frame):
                look.page_break = True
            else:
                sections.append(candidate)
            cursor = start_y = bottom = sections[-1].top
            flow = 0
        elif key != layout:
            sections.append(_section(index, page, frame, columns, False))
            cursor = start_y = bottom
            flow = 0
        layout = key
        section = sections[-1]
        target = _flow_index(section, page, column)
        if target > flow:
            look.column_break = True
            cursor = start_y
        flow = max(flow, target)
        low, high = section.columns[flow]
        pitch = pitches.get(id(block), 0.0)
        extent = _extent(block, pitch)
        if isinstance(block, Para):
            placed = align_lines(_first_boxes(block), low, high, block.size, block.rtl)
            placed.pitch = pitch
            placed.width = high - low
        else:
            placed = _box_look(block.bbox, low, high, body)
        placed.page_break, placed.column_break = look.page_break, look.column_break
        placed.origin = low
        top = extent[0][1][1]
        broke = look.page_break or look.column_break
        overlaps = cursor - top > _BESIDE * body and extent[-1][0] == seg[0]
        if overlaps and not broke and not isinstance(block, TableItem):
            looks.append(_beside(block, placed, extent))
            bottom = max(bottom, extent[-1][1][3])
            continue
        looks.append(placed)
        placed.before = max(0.0, top - cursor)
        end_page, end = extent[-1]
        if isinstance(block, Para) and block.boxes and end_page == seg[0]:
            flow = max(flow, _flow_index(section, page, _column(block.boxes[-1][0])))
        cursor = end[3]
        bottom = max(bottom, cursor) if end_page == seg[0] else cursor
        last_page = max(last_page, end_page, _end_page(block))
    return Plan(looks, sections)
