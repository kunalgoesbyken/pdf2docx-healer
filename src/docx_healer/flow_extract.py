"""Per-page extraction of lines, links, images and tables."""

import re
import time
import zlib
from bisect import bisect_right
from dataclasses import replace
from typing import Any, Dict, List, Optional, Tuple

import pymupdf

from .flow_model import (
    Budget,
    Cell,
    ImageItem,
    Line,
    LinkKey,
    PageData,
    Rect,
    Span,
    TableItem,
    Unit,
)
from .flow_text import clean_text, collapse_spaces
from .log import logger

_BAND = 20.0
_BOLD_NAME = re.compile(r"bold|black|heavy", re.IGNORECASE)
_ITALIC_NAME = re.compile(r"italic|oblique", re.IGNORECASE)
_SUPER_RATIO = 0.85
_RISE = 0.2
_DROP = 0.15
_HEAD_CHARS = 16
_ROW_OVERLAP = 0.6
_ROW_GAP = 1.5
_MIN_IMAGE_SIDE = 12.0
_MIN_IMAGE_THICKNESS = 4.0
_FIGURE_GAP_SHARE = 0.5
_FIGURE_ALIGN_SHARE = 0.5
_EDGE_TOLERANCE = 2.0
_MIN_STROKES = 4
_SPARSE_CELLS = 12
_MIN_FILLED = 0.4
_MAX_FIGURES = 12
_MIN_EVIDENCE = 3
_MIN_FIGURE_WIDTH = 60.0
_MIN_FIGURE_HEIGHT = 40.0
_FIGURE_PAGE_SHARE = 0.85
_FIGURE_TEXT_SHARE = 0.2
_FIGURE_DPI = 144
_FIGURE_MAX_PIXELS = 2000
_LINK_KINDS = (pymupdf.LINK_URI, pymupdf.LINK_GOTO, pymupdf.LINK_NAMED)


def extract_page(page: Any, budget: Budget) -> PageData:
    """Return the page content as units of lines, tables and images."""
    rect = page.rect * page.derotation_matrix
    data = PageData(number=page.number, width=rect.width, height=rect.height)
    links = _page_links(page, budget)
    blocks = _page_blocks(page, budget)
    dominant = _dominant_direction(blocks)
    unit_index = 0
    for block in blocks:
        if block["type"] == 1:
            _add_image(data, block, budget)
            continue
        lines = [
            line for line in
            (_build_line(raw, page.number, unit_index, links) for raw in block["lines"])
            if line is not None
        ]
        if not lines:
            continue
        along = [raw for raw in block["lines"] if _direction(raw) == dominant]
        lines = _glue_rows(lines)
        edge = max(line.bbox[2] for line in lines)
        for line in lines:
            line.edge = edge
        data.units.append(Unit(lines, rotated=len(along) < len(block["lines"]) / 2))
        unit_index += 1
    _add_tables(page, data, budget)
    _add_figures(page, data, budget)
    return data


def _page_blocks(page: Any, budget: Budget) -> List[Dict[str, Any]]:
    """Read the page's blocks without ever decoding an image over max_image_pixels."""
    cap = budget.limits.max_image_pixels
    infos = page.get_image_info(xrefs=True)
    if all(info["width"] * info["height"] <= cap for info in infos):
        return list(page.get_text("rawdict")["blocks"])
    flags = pymupdf.TEXTFLAGS_RAWDICT & ~pymupdf.TEXT_PRESERVE_IMAGES
    blocks = list(page.get_text("rawdict", flags=flags)["blocks"])
    for info in infos:
        if info["width"] * info["height"] > cap:
            _warn_once(
                budget, f"pixels{page.number}",
                f"Skipped an image on page {page.number + 1} larger than max_image_pixels={cap}",
            )
            continue
        if not info["xref"]:
            continue
        try:
            extracted = page.parent.extract_image(info["xref"])
        except Exception as exc:
            logger.debug("Could not read image xref %d: %s", info["xref"], exc)
            continue
        blocks.append({
            "type": 1, "bbox": info["bbox"], "width": info["width"],
            "height": info["height"], "image": extracted["image"], "ext": extracted["ext"],
        })
    return blocks


def _direction(raw: Dict[str, Any]) -> Tuple[int, int]:
    return (round(raw["dir"][0]), round(raw["dir"][1]))


def _dominant_direction(blocks: List[Dict[str, Any]]) -> Tuple[int, int]:
    counts: Dict[Tuple[int, int], int] = {}
    for block in blocks:
        if block["type"] != 0:
            continue
        for raw in block["lines"]:
            key = _direction(raw)
            counts[key] = counts.get(key, 0) + sum(len(s["chars"]) for s in raw["spans"])
    return max(counts, key=lambda k: counts[k]) if counts else (1, 0)


def _same_row(left: Line, right: Line) -> bool:
    top = max(left.bbox[1], right.bbox[1])
    bottom = min(left.bbox[3], right.bbox[3])
    height = min(left.bbox[3] - left.bbox[1], right.bbox[3] - right.bbox[1])
    limit = _ROW_GAP * left.size
    forward = right.bbox[0] - left.bbox[2]
    backward = left.bbox[0] - right.bbox[2]
    near = -1.0 <= forward <= limit or -1.0 <= backward <= limit
    return bottom - top >= _ROW_OVERLAP * height and near


def _glue_rows(lines: List[Line]) -> List[Line]:
    """Join lines of one block that sit on the same baseline row, such as a section number."""
    glued: List[Line] = []
    for line in lines:
        if glued and _same_row(glued[-1], line):
            glued[-1] = _join_row(glued[-1], line)
        else:
            glued.append(line)
    return glued


def _join_row(left: Line, right: Line) -> Line:
    spans = list(left.spans)
    if not left.text.endswith(" ") and not right.text.startswith(" "):
        spans.append(Span(" ", left.spans[-1].size, font=left.spans[-1].font))
    spans.extend(right.spans)
    heads = (left.head_x + [left.bbox[2]] + right.head_x)[:_HEAD_CHARS]
    bbox = (
        min(left.bbox[0], right.bbox[0]), min(left.bbox[1], right.bbox[1]),
        max(left.bbox[2], right.bbox[2]), max(left.bbox[3], right.bbox[3]),
    )
    size = left.size if len(left.text) >= len(right.text) else right.size
    return Line(spans, bbox, left.page, left.unit, size, left.bold and right.bold, heads)


def _page_links(page: Any, budget: Budget) -> Dict[int, List[Tuple[Rect, LinkKey]]]:
    index: Dict[int, List[Tuple[Rect, LinkKey]]] = {}
    count = 0
    cap = budget.limits.max_links_per_page
    for link in page.get_links():
        key = _link_key(page.parent, link)
        if key is None:
            continue
        if count >= cap:
            _warn_once(
                budget, f"links{page.number}",
                f"Page {page.number + 1} has more than max_links_per_page={cap} links; "
                "the rest stay plain text",
            )
            break
        count += 1
        rect = (pymupdf.Rect(link["from"]) & page.rect) * page.derotation_matrix
        if rect.is_empty or rect.is_infinite:
            continue
        box = (rect.x0, rect.y0, rect.x1, rect.y1)
        for band in range(int(box[1] // _BAND), int(box[3] // _BAND) + 1):
            index.setdefault(band, []).append((box, key))
    return index


def _link_key(doc: Any, link: Dict[str, Any]) -> Optional[LinkKey]:
    kind = link.get("kind")
    if kind not in _LINK_KINDS:
        return None
    if kind == pymupdf.LINK_URI:
        uri = link.get("uri") or ""
        return ("uri", uri) if uri else None
    if kind == pymupdf.LINK_GOTO:
        target = link.get("to")
        page_number = link.get("page", -1)
        if page_number < 0 or target is None or not hasattr(target, "x"):
            return None
        return ("goto", page_number, float(target.x), float(target.y))
    name = link.get("nameddest")
    if not name:
        return None
    try:
        page_number, x, y = doc.resolve_link("#" + name)
    except Exception:
        return None
    return ("goto", page_number, float(x), float(y)) if page_number >= 0 else None


def _warn_once(budget: Budget, key: str, message: str) -> None:
    if key not in budget.warned:
        budget.warned.add(key)
        logger.warning(message)


def _link_for(links: Dict[int, List[Tuple[Rect, LinkKey]]], bbox: Rect) -> Optional[LinkKey]:
    center_x = (bbox[0] + bbox[2]) / 2
    center_y = (bbox[1] + bbox[3]) / 2
    for box, key in links.get(int(center_y // _BAND), ()):
        if box[0] <= center_x <= box[2] and box[1] <= center_y <= box[3]:
            return key
    return None


def _rect(values: Any) -> Rect:
    return (float(values[0]), float(values[1]), float(values[2]), float(values[3]))


def _style(raw_span: Dict[str, Any], dominant: Optional[Dict[str, Any]]) -> Span:
    flags = raw_span["flags"]
    font = raw_span.get("font", "")
    sup, sub = _script_position(raw_span, dominant, flags)
    return Span(
        "", float(raw_span["size"]), bold=bool(flags & 16) or bool(_BOLD_NAME.search(font)),
        italic=bool(flags & 2) or bool(_ITALIC_NAME.search(font)), mono=bool(flags & 8),
        sup=sup, sub=sub, font=font,
    )


def _build_line(
    raw: Dict[str, Any], page_number: int, unit: int, links: Dict[int, List[Tuple[Rect, LinkKey]]]
) -> Optional[Line]:
    spans: List[Span] = []
    head_x: List[float] = []
    dominant = _dominant_span(raw["spans"])
    for raw_span in raw["spans"]:
        style = _style(raw_span, dominant)
        pieces: List[Tuple[str, Optional[LinkKey]]] = []
        for char in raw_span["chars"]:
            key = _link_for(links, char["bbox"]) if links else None
            if pieces and pieces[-1][1] == key:
                pieces[-1] = (pieces[-1][0] + char["c"], key)
            else:
                pieces.append((char["c"], key))
            if len(head_x) < _HEAD_CHARS:
                head_x.append(char["bbox"][0])
        spans.extend(replace(style, text=text, link=key) for text, key in pieces)
    for span in spans:
        span.text = collapse_spaces(clean_text(span.text))
    spans = [span for span in spans if span.text]
    visible = [span for span in spans if span.text.strip()]
    if not visible:
        return None
    sizes: Dict[float, int] = {}
    for span in spans:
        sizes[span.size] = sizes.get(span.size, 0) + len(span.text)
    size = max(sizes, key=lambda s: sizes[s])
    bold = all(span.bold for span in visible)
    return Line(spans, _rect(raw["bbox"]), page_number, unit, size, bold, head_x)


def _dominant_span(spans: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    best = None
    best_chars = 0
    for span in spans:
        count = len(span["chars"])
        if count > best_chars:
            best, best_chars = span, count
    return best


def _script_position(
    span: Dict[str, Any], dominant: Optional[Dict[str, Any]], flags: int
) -> Tuple[bool, bool]:
    if flags & 1:
        return True, False
    if dominant is None or span is dominant:
        return False, False
    main_size = float(dominant["size"])
    if float(span["size"]) > _SUPER_RATIO * main_size:
        return False, False
    rise = float(dominant["origin"][1]) - float(span["origin"][1])
    if rise > _RISE * main_size:
        return True, False
    if -rise > _DROP * main_size:
        return False, True
    return False, False


def _add_image(data: PageData, block: Dict[str, Any], budget: Budget) -> None:
    bbox = _rect(block["bbox"])
    width, height = bbox[2] - bbox[0], bbox[3] - bbox[1]
    image = block.get("image")
    if not image:
        return
    if max(width, height) < _MIN_IMAGE_SIDE or min(width, height) < _MIN_IMAGE_THICKNESS:
        logger.debug("Skipped decorative image on page %d", data.number + 1)
        return
    if budget.images_left <= 0 or budget.bytes_left < len(image):
        _warn_once(
            budget, "images",
            "Image budget (max_images / max_image_bytes) exceeded; remaining images are skipped",
        )
        return
    budget.images_left -= 1
    budget.bytes_left -= len(image)
    key = ("img", zlib.crc32(image), len(image), round(bbox[0] / 6), round(bbox[1] / 6))
    data.images.append(ImageItem(bbox, image, block.get("ext", "png"), key, data.number))


def _add_tables(page: Any, data: PageData, budget: Budget) -> None:
    if not data.units:
        return
    limits = budget.limits
    if budget.table_seconds_left <= 0:
        _warn_once(
            budget, "tables",
            f"Table detection stopped: table_seconds={limits.table_seconds} used up; "
            "later tables stay paragraphs",
        )
        return
    items = _drawing_items(page, budget)
    if items is None or items < _MIN_STROKES:
        return
    started = time.perf_counter()
    try:
        found = page.find_tables().tables
    except Exception as exc:
        logger.warning("Table detection failed on page %d: %s", page.number + 1, exc)
        found = []
    budget.table_seconds_left -= time.perf_counter() - started
    taken: List[Rect] = []
    for table in found:
        item = _table_item(table, data.number, data.height, limits.max_table_cells)
        if item is None or any(_overlap(item.bbox, other) for other in taken):
            continue
        inside = _lines_inside(item, data)
        if not _dense(item, inside):
            logger.debug("Skipped a sparse table on page %d", data.number + 1)
            continue
        taken.append(item.bbox)
        _fill_table(item, data, inside)
        data.tables.append(item)
    data.tables.sort(key=lambda t: (t.bbox[1], t.bbox[0]))


def _overlap(a: Rect, b: Rect) -> bool:
    width = min(a[2], b[2]) - max(a[0], b[0])
    height = min(a[3], b[3]) - max(a[1], b[1])
    if width <= 0 or height <= 0:
        return False
    smaller = min((a[2] - a[0]) * (a[3] - a[1]), (b[2] - b[0]) * (b[3] - b[1]))
    return width * height > 0.5 * smaller


def _edges(values: List[float]) -> List[float]:
    edges: List[float] = []
    for value in sorted(values):
        if not edges or value - edges[-1] > _EDGE_TOLERANCE:
            edges.append(value)
    return edges


def _edge_index(edges: List[float], value: float) -> int:
    return min(range(len(edges)), key=lambda i: abs(edges[i] - value))


def _table_item(
    table: Any, page_number: int, height: float, max_cells: int
) -> Optional[TableItem]:
    rects = [_rect(cell) for row in table.rows for cell in row.cells if cell is not None]
    if len(rects) < 4:
        return None
    xs = _edges([v for r in rects for v in (r[0], r[2])])
    ys = _edges([v for r in rects for v in (r[1], r[3])])
    rows, cols = len(ys) - 1, len(xs) - 1
    if rows < 2 or cols < 2 or rows * cols > max_cells:
        return None
    cells = []
    for r in rects:
        c0, c1 = _edge_index(xs, r[0]), _edge_index(xs, r[2])
        r0, r1 = _edge_index(ys, r[1]), _edge_index(ys, r[3])
        if c1 > c0 and r1 > r0:
            cells.append(Cell(r0, r1, c0, c1, r))
    if not cells:
        return None
    bbox = (xs[0], ys[0], xs[-1], ys[-1])
    grid = {(r, c): cell for cell in cells for r in range(cell.r0, cell.r1)
            for c in range(cell.c0, cell.c1)}
    return TableItem(
        bbox, rows, cols, xs, cells, page=page_number, height=height, row_edges=ys, grid=grid
    )


def _lines_inside(item: TableItem, data: PageData) -> List[Line]:
    box = item.bbox
    return [
        line for unit in data.units for line in unit.lines
        if box[0] - 1 <= (line.bbox[0] + line.bbox[2]) / 2 <= box[2] + 1
        and box[1] - 1 <= (line.bbox[1] + line.bbox[3]) / 2 <= box[3] + 1
    ]


def _dense(item: TableItem, inside: List[Line]) -> bool:
    """A grid is a table when enough of the cells in its text-bearing columns hold text."""
    if len(item.cells) < _SPARSE_CELLS:
        return True
    filled = {id(cell): cell for cell in (_cell_for(item, line) for line in inside)}
    columns = {c for cell in filled.values() for c in range(cell.c0, cell.c1)}
    live = [cell for cell in item.cells if any(c in columns for c in range(cell.c0, cell.c1))]
    return len(filled) >= _MIN_FILLED * len(live)


def _fill_table(item: TableItem, data: PageData, inside: List[Line]) -> None:
    for line in inside:
        _cell_for(item, line).lines.append(line)
    gone = {id(line) for line in inside}
    for unit in data.units:
        unit.lines = [line for line in unit.lines if id(line) not in gone]
    data.units = [unit for unit in data.units if unit.lines]


def _cell_for(item: TableItem, line: Line) -> Cell:
    cx = (line.bbox[0] + line.bbox[2]) / 2
    cy = (line.bbox[1] + line.bbox[3]) / 2
    row = min(max(bisect_right(item.row_edges, cy) - 1, 0), item.rows - 1)
    col = min(max(bisect_right(item.col_edges, cx) - 1, 0), item.cols - 1)
    cell = item.grid.get((row, col))
    if cell is not None:
        return cell
    same_row = [c for (r, _), c in item.grid.items() if r == row]
    return same_row[0] if same_row else item.cells[0]


def _drawing_items(page: Any, budget: Budget) -> Optional[int]:
    """Count the page's path segments, or None when the page is over a drawing limit."""
    limits = budget.limits
    paths = page.get_cdrawings()
    items = sum(len(path["items"]) for path in paths)
    if len(paths) > limits.max_drawings or items > limits.max_drawing_items:
        _warn_once(
            budget, f"drawings{page.number}",
            f"Page {page.number + 1} has more than max_drawings={limits.max_drawings} paths "
            f"or max_drawing_items={limits.max_drawing_items} segments; "
            "its tables and figures stay as text",
        )
        return None
    return items


def _add_figures(page: Any, data: PageData, budget: Budget) -> None:
    if not data.units or budget.table_seconds_left <= 0:
        return
    if _drawing_items(page, budget) is None:
        return
    started = time.perf_counter()
    try:
        _rasterise_figures(page, data, budget)
    finally:
        budget.table_seconds_left -= time.perf_counter() - started


def _rasterise_figures(page: Any, data: PageData, budget: Budget) -> None:
    try:
        drawings = page.get_drawings()
        clusters = page.cluster_drawings(drawings=drawings)
    except Exception as exc:
        logger.debug("Vector figures skipped on page %d: %s", data.number + 1, exc)
        return
    taken = [t.bbox for t in data.tables] + [i.bbox for i in data.images]
    clusters = _merge_clusters(clusters[: budget.limits.max_figure_clusters])
    for cluster in clusters[:_MAX_FIGURES]:
        rect = cluster * page.derotation_matrix
        box = (rect.x0, rect.y0, rect.x1, rect.y1)
        if rect.width < _MIN_FIGURE_WIDTH or rect.height < _MIN_FIGURE_HEIGHT:
            continue
        if rect.width * rect.height > _FIGURE_PAGE_SHARE * data.width * data.height:
            continue
        if any(_overlap(box, other) for other in taken):
            continue
        if _evidence(drawings, cluster) < _MIN_EVIDENCE:
            continue
        inside = [(u, line) for u in data.units for line in u.lines if _inside(line, box)]
        area = sum(
            (line.bbox[2] - line.bbox[0]) * (line.bbox[3] - line.bbox[1]) for _, line in inside
        )
        if area > _FIGURE_TEXT_SHARE * rect.width * rect.height:
            continue
        image = _raster(page, cluster, budget)
        if image is None:
            continue
        figure_key = ("fig", zlib.crc32(image), len(image))
        data.images.append(ImageItem(box, image, "png", figure_key, data.number))
        data.figures += 1
        taken.append(box)
        gone = {id(line) for _, line in inside}
        for unit in data.units:
            unit.lines = [line for line in unit.lines if id(line) not in gone]
        data.units = [unit for unit in data.units if unit.lines]


def _overlap_length(low_a: float, high_a: float, low_b: float, high_b: float) -> float:
    return min(high_a, high_b) - max(low_a, low_b)


def _adjacent(a: Rect, b: Rect) -> bool:
    """True when two clusters are stacked or side by side, aligned and close together."""
    across = _overlap_length(a[0], a[2], b[0], b[2])
    down = _overlap_length(a[1], a[3], b[1], b[3])
    narrow = min(a[2] - a[0], b[2] - b[0])
    short = min(a[3] - a[1], b[3] - b[1])
    stacked = across >= _FIGURE_ALIGN_SHARE * narrow and -down <= _FIGURE_GAP_SHARE * short
    beside = down >= _FIGURE_ALIGN_SHARE * short and -across <= _FIGURE_GAP_SHARE * narrow
    return stacked or beside


def _merge_clusters(clusters: List[Any]) -> List[Any]:
    """Join the clusters of one multi-part figure, such as stacked subplots, into one."""
    boxes: List[Rect] = [(c.x0, c.y0, c.x1, c.y1) for c in clusters]
    changed = True
    while changed:
        changed = False
        joined: List[Rect] = []
        for box in boxes:
            for index, other in enumerate(joined):
                if _adjacent(box, other):
                    joined[index] = (
                        min(box[0], other[0]), min(box[1], other[1]),
                        max(box[2], other[2]), max(box[3], other[3]),
                    )
                    changed = True
                    break
            else:
                joined.append(box)
        boxes = joined
    return [pymupdf.Rect(*box) for box in sorted(boxes, key=lambda b: (b[1], b[0]))]


def _curved(item: Any) -> bool:
    if item[0] in ("c", "qu"):
        return True
    return item[0] == "l" and abs(item[1].x - item[2].x) > 0.5 and abs(item[1].y - item[2].y) > 0.5


def _evidence(drawings: List[Dict[str, Any]], cluster: Any) -> int:
    """Count shapes in the cluster that rules, empty boxes and table grids do not produce."""
    total = 0
    for drawing in drawings:
        if not cluster.contains(drawing["rect"]):
            continue
        if drawing.get("fill") is not None:
            total += len(drawing["items"])
        else:
            total += sum(1 for item in drawing["items"] if _curved(item))
    return total


def _inside(line: Line, box: Rect) -> bool:
    cx = (line.bbox[0] + line.bbox[2]) / 2
    cy = (line.bbox[1] + line.bbox[3]) / 2
    return box[0] <= cx <= box[2] and box[1] <= cy <= box[3]


def _raster(page: Any, clip: Any, budget: Budget) -> Optional[bytes]:
    dpi = min(_FIGURE_DPI, _FIGURE_MAX_PIXELS * 72.0 / max(clip.width, clip.height, 1.0))
    try:
        data = bytes(page.get_pixmap(clip=clip, dpi=int(dpi)).tobytes("png"))
    except Exception as exc:
        logger.debug("Could not rasterise a vector figure: %s", exc)
        return None
    if budget.images_left <= 0 or budget.bytes_left < len(data):
        _warn_once(
            budget, "images",
            "Image budget (max_images / max_image_bytes) exceeded; remaining images are skipped",
        )
        return None
    budget.images_left -= 1
    budget.bytes_left -= len(data)
    return data
