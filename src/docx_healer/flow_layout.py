"""Reading order, running-content removal and paragraph assembly for editable mode."""

import math
import re
from collections import Counter
from typing import Any, Dict, List, Optional, Set, Tuple

from .flow_lists import parse_marker
from .flow_model import Block, Line, PageData, Para, Rect, Span, TableItem, Unit, union
from .flow_text import join_decision, strong_direction

Seg = Tuple[Any, ...]

_BAND = 0.1
_REPEAT_SHARE = 0.5
_COLUMN_WIDTH = 8.0
_MIN_CROSSING = 1
_CROSSING_SHARE = 0.2
_MIN_COLUMN_LINES = 3
_MIN_TOTAL_LINES = 6
_MAX_GUTTERS = 8
_GRID_TOLERANCE = 6.0
_LOW_SHARE = 0.75
_ROW_TOP = 0.5
_PARA_GAP = 0.35
_SIDE_BY_SIDE = -0.6
_SIZE_JUMP = 0.15
_INDENT = 0.8
_CENTER_SHARE = 0.03
_TERMINAL = re.compile(r"[.!?:;\u2026\u3002\uff01\uff1f][\"'\u201d\u2019)\]]*$")
_START_TOKENS = {
    ("decimal", 1), ("lowerLetter", 1), ("upperLetter", 1), ("lowerRoman", 1), ("upperRoman", 1),
}
_BACKGROUND_SHARE = 0.9


def _normalise(text: str) -> str:
    return re.sub(r"\d+", "#", text.strip())


def drop_running(pages: List[PageData]) -> int:
    """Remove lines and images that repeat at the same position on many pages."""
    text_pages = [p for p in pages if p.units]
    if len(text_pages) < 2:
        return 0
    needed = max(2, math.ceil(_REPEAT_SHARE * len(text_pages)))
    seen: Dict[Tuple[Any, ...], Set[int]] = {}
    for page in pages:
        for key in _running_keys(page):
            seen.setdefault(key, set()).add(page.number)
    repeated = {key for key, where in seen.items() if len(where) >= needed}
    dropped = 0
    for page in pages:
        for unit in page.units:
            kept = [line for line in unit.lines if _line_key(page, line) not in repeated]
            dropped += len(unit.lines) - len(kept)
            unit.lines = kept
        page.units = [unit for unit in page.units if unit.lines]
        images = [img for img in page.images if img.key not in repeated]
        dropped += len(page.images) - len(images)
        page.images = images
    return dropped


def _in_margin(page: PageData, line: Line) -> bool:
    return line.bbox[3] <= _BAND * page.height or line.bbox[1] >= (1 - _BAND) * page.height


def _line_key(page: PageData, line: Line) -> Optional[Tuple[Any, ...]]:
    if not _in_margin(page, line):
        return None
    return ("line", _normalise(line.text), round(line.bbox[1] / 6))


def _running_keys(page: PageData) -> Set[Tuple[Any, ...]]:
    keys: Set[Tuple[Any, ...]] = {img.key for img in page.images}
    for unit in page.units:
        for line in unit.lines:
            key = _line_key(page, line)
            if key is not None:
                keys.add(key)
    return keys


def body_size(pages: List[PageData]) -> float:
    counts: Counter = Counter()
    for page in pages:
        for unit in page.units:
            for line in unit.lines:
                counts[round(line.size * 2) / 2] += len(line.text)
        for table in page.tables:
            for cell in table.cells:
                for line in cell.lines:
                    counts[round(line.size * 2) / 2] += len(line.text)
    return counts.most_common(1)[0][0] if counts else 10.0


def _width(rect: Rect) -> float:
    return rect[2] - rect[0]


def _gutters(units: List[Unit], size: float, extras: List[Rect]) -> List[Tuple[float, float]]:
    columns = [u for u in units if _width(u.bbox) >= _COLUMN_WIDTH * size]
    total = sum(len(u.lines) for u in columns)
    if total < _MIN_TOTAL_LINES:
        return []
    events: List[Tuple[float, int]] = []
    for unit in columns:
        box = unit.bbox
        events.append((box[0], len(unit.lines)))
        events.append((box[2], -len(unit.lines)))
    events.sort()
    allowed = max(_MIN_CROSSING, _CROSSING_SHARE * total)
    minimum = max(6.0, 0.7 * size)
    found: List[Tuple[float, float]] = []
    covering = 0
    start: Optional[float] = None
    previous = events[0][0]
    for x, delta in events:
        if x > previous:
            if covering <= allowed:
                start = previous if start is None else start
            elif start is not None:
                found.append((start, previous))
                start = None
        covering += delta
        previous = x
    first, last = events[0][0], events[-1][0]
    inner = [g for g in found if g[1] - g[0] >= minimum and g[0] > first and g[1] < last]
    return [gap for gap in inner[:_MAX_GUTTERS] if _valid_gutter(gap, columns, extras)]


def _valid_gutter(gap: Tuple[float, float], columns: List[Unit], extras: List[Rect]) -> bool:
    left = [u for u in columns if u.bbox[2] <= gap[0] + 1]
    right = [u for u in columns if u.bbox[0] >= gap[1] - 1]
    enough = all(sum(len(u.lines) for u in side) >= _MIN_COLUMN_LINES for side in (left, right))
    sides = (
        [u.bbox for u in left] + [e for e in extras if e[2] <= gap[0] + 1],
        [u.bbox for u in right] + [e for e in extras if e[0] >= gap[1] - 1],
    )
    return enough and _parallel(*sides)


def _parallel(left: List[Rect], right: List[Rect]) -> bool:
    a, b = union(left), union(right)
    overlap = min(a[3], b[3]) - max(a[1], b[1])
    return overlap >= 0.3 * min(a[3] - a[1], b[3] - b[1])


def _crosses(rect: Rect, gutters: List[Tuple[float, float]]) -> bool:
    return any(rect[0] < low - 1 and rect[2] > high + 1 for low, high in gutters)


def _column(rect: Rect, gutters: List[Tuple[float, float]]) -> int:
    center = (rect[0] + rect[2]) / 2
    return sum(1 for low, high in gutters if (low + high) / 2 < center)


def _row_sort(items: List[Any]) -> List[Any]:
    rows: List[List[Any]] = []
    firsts: List[Rect] = []
    for item in sorted(items, key=lambda i: (i.bbox[1], i.bbox[0])):
        top, bottom = item.bbox[1], item.bbox[3]
        if rows:
            first = firsts[-1]
            if top - first[1] < _ROW_TOP * min(bottom - top, first[3] - first[1]):
                rows[-1].append(item)
                continue
        rows.append([item])
        firsts.append(item.bbox)
    return [item for row in rows for item in sorted(row, key=lambda i: i.bbox[0])]


def order_page(page: PageData, size: float) -> List[Tuple[Any, Seg]]:
    """Return the page items in reading order, each tagged with its column segment."""
    flow = [u for u in page.units if not u.rotated]
    items: List[Any] = flow + list(page.tables) + list(page.images)
    gutters = _gutters(flow, size, [t.bbox for t in page.tables] + [i.bbox for i in page.images])
    ordered: List[Tuple[Any, Seg]] = []
    run: List[Any] = []
    runs = 0

    def flush() -> None:
        nonlocal run, runs
        by_column: Dict[int, List[Any]] = {}
        for item in run:
            by_column.setdefault(_column(item.bbox, gutters), []).append(item)
        for column in sorted(by_column):
            for item in _row_sort(by_column[column]):
                ordered.append((item, (page.number, runs, column)))
        run = []
        runs += 1

    for item in sorted(items, key=lambda i: (i.bbox[1], i.bbox[0])):
        if gutters and _crosses(item.bbox, gutters):
            flush()
            ordered.append((item, (page.number, runs, -1)))
            runs += 1
        else:
            run.append(item)
    flush()
    for index, unit in enumerate(u for u in page.units if u.rotated):
        ordered.append((unit, (page.number, "rotated", index)))
    return ordered


def _text_x(line: Line, length: int) -> float:
    index = len(line.text) - len(line.text.lstrip()) + length
    if index < len(line.head_x):
        return line.head_x[index]
    return line.bbox[0] + 18.0


def _marker_starts(cur: Optional[Para], line: Line) -> bool:
    marker = parse_marker(line.text.lstrip())
    if marker is None or strong_direction(line.text) == "rtl" or cur is None:
        return False
    if marker.bullet:
        if not marker.weak:
            return True
        return cur.marker is not None and cur.marker.bullet
    starts = any(interp in _START_TOKENS for interp in marker.interps)
    if cur.marker is not None and not cur.marker.bullet:
        return starts or cur.marker.delim == marker.delim
    return starts and _TERMINAL.search(cur.last_line_text.rstrip()) is not None


def _continues(cur: Para, prev: Line, line: Line, page_width: float) -> bool:
    size = cur.size
    if abs(line.size - cur.size) > _SIZE_JUMP * size or line.bold != prev.bold:
        return False
    gap = line.bbox[1] - prev.bbox[3]
    if gap > _PARA_GAP * size or gap < _SIDE_BY_SIDE * size:
        return False
    if line.unit != prev.unit and (line.bbox[0] >= prev.bbox[2] or line.bbox[2] <= prev.bbox[0]):
        return False
    delta = line.bbox[0] - prev.bbox[0]
    centered = abs(
        (line.bbox[0] + line.bbox[2]) / 2 - (prev.bbox[0] + prev.bbox[2]) / 2
    ) < _CENTER_SHARE * page_width
    if cur.marker is not None:
        full = prev.bbox[2] >= prev.edge - 0.5 * size
        return line.bbox[0] >= cur.marker_x - 0.5 * size or full
    if cur.lines >= 2 and not centered and abs(delta) > _INDENT * size:
        return False
    return True


def _new_para(line: Line, seg: Seg, page_width: float) -> Para:
    marker = parse_marker(line.text.lstrip())
    if marker is not None and strong_direction(line.text) == "rtl":
        marker = None
    spans = [s.clone() for s in line.spans]
    para = Para(
        spans=spans, page=line.page, seg=seg, bbox=line.bbox, x0=line.bbox[0],
        size=line.size, bold=line.bold, last_line_text=line.text,
    )
    _trim(para.spans, leading=True)
    if marker is not None:
        para.marker = marker
        para.marker_x = line.bbox[0]
        para.text_x = _text_x(line, marker.length)
    return para


def _trim(spans: List[Span], leading: bool = False, trailing: bool = True) -> None:
    if leading and spans:
        spans[0].text = spans[0].text.lstrip()
    if trailing and spans:
        spans[-1].text = spans[-1].text.rstrip()


def _append_line(para: Para, line: Line, vocab: Set[str]) -> None:
    previous = para.last_line_text.rstrip()
    following = line.text.lstrip()
    drop, space = join_decision(previous, following, vocab)
    spans = [s.clone() for s in line.spans]
    _trim(spans, leading=True, trailing=False)
    _trim(para.spans)
    if drop and para.spans:
        para.spans[-1].text = para.spans[-1].text[:-1]
    if space and para.spans:
        last = para.spans[-1]
        if last.link is not None:
            para.spans.append(Span(" ", last.size, last.bold, last.italic, last.mono))
        else:
            last.text += " "
    para.spans.extend(spans)
    para.lines += 1
    para.bbox = union([para.bbox, line.bbox])
    para.last_line_text = line.text


def finish_para(para: Para) -> None:
    _trim(para.spans)
    merged: List[Span] = []
    for span in para.spans:
        if not span.text:
            continue
        blank = not span.text.strip() and span.link is None and merged and merged[-1].link is None
        if merged and (merged[-1].same_style(span) or blank):
            merged[-1].text += span.text
        else:
            merged.append(span)
    para.spans = merged
    para.rtl = strong_direction(para.text) == "rtl"


def _flat(items: List[Tuple[Any, Seg]]) -> List[Tuple[Any, Seg]]:
    stream: List[Tuple[Any, Seg]] = []
    for item, seg in items:
        if isinstance(item, Unit):
            stream.extend((line, seg) for line in item.lines)
        else:
            stream.append((item, seg))
    return stream


def _marker_only(line: Line) -> bool:
    text = line.text.strip()
    return 0 < len(text) <= 6 and parse_marker(text + " x") is not None


def _beside(marker: Line, line: Line) -> bool:
    top = max(marker.bbox[1], line.bbox[1])
    bottom = min(marker.bbox[3], line.bbox[3])
    height = min(marker.bbox[3] - marker.bbox[1], line.bbox[3] - line.bbox[1])
    gap = line.bbox[0] - marker.bbox[2]
    return bottom - top >= 0.5 * height and -1 <= gap <= 4 * marker.size


def _glue_markers(stream: List[Tuple[Any, Seg]]) -> List[Tuple[Any, Seg]]:
    glued: List[Tuple[Any, Seg]] = []
    index = 0
    while index < len(stream):
        item, seg = stream[index]
        following = stream[index + 1][0] if index + 1 < len(stream) else None
        if isinstance(item, Line) and isinstance(following, Line) and _marker_only(item) \
                and _beside(item, following):
            spans = [s.clone() for s in item.spans]
            spans[-1].text = spans[-1].text.rstrip() + " "
            heads = item.head_x[: len(item.text.strip())]
            merged = Line(
                spans + following.spans, union([item.bbox, following.bbox]), following.page,
                following.unit, following.size, following.bold,
                heads + [item.bbox[2]] + following.head_x,
            )
            glued.append((merged, seg))
            index += 2
            continue
        glued.append((item, seg))
        index += 1
    return glued


def assemble(
    items: List[Tuple[Any, Seg]], vocab: Set[str], page_width: float
) -> List[Block]:
    """Turn ordered items into paragraphs, tables and images."""
    blocks: List[Block] = []
    cur: Optional[Para] = None
    prev: Optional[Line] = None
    for item, seg in _glue_markers(_flat(items)):
        if not isinstance(item, Line):
            if cur is not None:
                finish_para(cur)
            cur, prev = None, None
            blocks.append(item)
            continue
        if cur is not None and prev is not None and cur.seg == seg \
                and _continues(cur, prev, item, page_width) \
                and not _marker_starts(cur, item):
            _append_line(cur, item, vocab)
        else:
            if cur is not None:
                finish_para(cur)
            cur = _new_para(item, seg, page_width)
            blocks.append(cur)
        prev = item
    if cur is not None:
        finish_para(cur)
    return blocks


def join_across_segments(blocks: List[Block], vocab: Set[str]) -> List[Block]:
    """Join a paragraph cut by a column or page break back together."""
    result: List[Block] = []
    for block in blocks:
        last = result[-1] if result else None
        if isinstance(block, Para) and isinstance(last, Para) and _joinable(last, block):
            line = Line(block.spans, block.bbox, block.page, 0, block.size, block.bold)
            _append_line(last, line, vocab)
            last.lines += block.lines - 1
            last.last_line_text = block.last_line_text
            finish_para(last)
            continue
        result.append(block)
    return result


def _joinable(first: Para, second: Para) -> bool:
    if first.seg == second.seg or first.seg[1:2] == ("rotated",) or second.seg[1:2] == ("rotated",):
        return False
    if first.marker is not None and not first.text.strip():
        return False
    if second.marker is not None or not second.text:
        return False
    if abs(first.size - second.size) > 0.6 or first.bold != second.bold or first.rtl != second.rtl:
        return False
    text = first.last_line_text.rstrip()
    if not text:
        return False
    head = second.text[0]
    if text[-1] in "-\u2010\u00ad":
        return True
    return _TERMINAL.search(text) is None and not (head.isupper() or head.isdigit())


def table_header_rows(table: TableItem) -> int:
    row = [cell for cell in table.cells if cell.r0 == 0]
    texts = [cell for cell in row if cell.lines]
    if table.rows < 2 or not texts:
        return 0
    return 1 if all(all(line.bold for line in cell.lines) for cell in texts) else 0


def fill_cells(table: TableItem, vocab: Set[str], page_width: float) -> None:
    for cell in table.cells:
        ordered = sorted(cell.lines, key=lambda line: (line.bbox[1], line.bbox[0]))
        unit = Unit(ordered)
        paras = assemble([(unit, ("cell", id(cell)))], vocab, page_width)
        cell.paras = [p for p in paras if isinstance(p, Para)]
        for para in cell.paras:
            para.marker = None
    table.header_rows = table_header_rows(table)


def merge_tables(blocks: List[Block]) -> List[Block]:
    """Merge tables that continue across a column or page break."""
    result: List[Block] = []
    for block in blocks:
        last = result[-1] if result else None
        if (
            isinstance(block, TableItem) and isinstance(last, TableItem)
            and _continues_table(last, block)
        ):
            _absorb(last, block)
            continue
        result.append(block)
    return result


def _continues_table(a: TableItem, b: TableItem) -> bool:
    if a.page == b.page:
        moved = b.bbox[1] < a.bbox[1]
    else:
        moved = a.bbox[3] >= _LOW_SHARE * a.height and b.bbox[1] <= (1 - _LOW_SHARE) * b.height
    return moved and _same_grid(a, b)


def _same_grid(a: TableItem, b: TableItem) -> bool:
    widths_a = [y - x for x, y in zip(a.col_edges, a.col_edges[1:])]
    widths_b = [y - x for x, y in zip(b.col_edges, b.col_edges[1:])]
    return len(widths_a) == len(widths_b) and all(
        abs(x - y) <= _GRID_TOLERANCE for x, y in zip(widths_a, widths_b)
    )


def _row_texts(table: TableItem, row: int) -> List[str]:
    cells = sorted((c for c in table.cells if c.r0 == row), key=lambda c: c.c0)
    return [" ".join(p.text for p in c.paras) for c in cells]


def _absorb(first: TableItem, second: TableItem) -> None:
    skip = 0
    if first.header_rows and second.header_rows and _row_texts(first, 0) == _row_texts(second, 0):
        skip = second.header_rows
    for cell in second.cells:
        if cell.r0 < skip:
            continue
        cell.r0 += first.rows - skip
        cell.r1 += first.rows - skip
        first.cells.append(cell)
    first.rows += second.rows - skip


def drop_backgrounds(page: PageData) -> None:
    area = page.width * page.height
    page.images = [
        img for img in page.images
        if (img.bbox[2] - img.bbox[0]) * (img.bbox[3] - img.bbox[1]) < _BACKGROUND_SHARE * area
    ]
