"""Data model and limits shared by the editable-mode modules."""

from dataclasses import dataclass, field, replace
from typing import Any, Dict, List, Optional, Set, Tuple, Union

Rect = Tuple[float, float, float, float]
LinkKey = Tuple[Any, ...]


@dataclass(frozen=True)
class FlowLimits:
    """Upper bounds that keep editable-mode work finite on hostile inputs."""

    max_pages: int = 2000
    max_links_per_page: int = 2000
    max_images: int = 2000
    max_outline_entries: int = 5000
    max_image_bytes: int = 128 * 1024 * 1024
    table_seconds: float = 120.0
    max_drawings: int = 5000
    max_drawing_items: int = 200000
    max_image_pixels: int = 50_000_000
    max_table_cells: int = 20000
    max_marker_lookahead: int = 40
    max_figure_clusters: int = 200


@dataclass
class Budget:
    """Mutable counters that enforce the limits across pages."""

    limits: FlowLimits
    images_left: int = 0
    bytes_left: int = 0
    table_seconds_left: float = 0.0
    warned: Set[str] = field(default_factory=set)

    def __post_init__(self) -> None:
        self.images_left = self.limits.max_images
        self.bytes_left = self.limits.max_image_bytes
        self.table_seconds_left = self.limits.table_seconds


@dataclass
class Span:
    """A run of text with uniform formatting."""

    text: str
    size: float
    bold: bool = False
    italic: bool = False
    mono: bool = False
    sup: bool = False
    sub: bool = False
    font: str = ""
    link: Optional[LinkKey] = None

    def clone(self) -> "Span":
        return replace(self)

    def same_style(self, other: "Span") -> bool:
        return (
            self.bold == other.bold and self.italic == other.italic and self.mono == other.mono
            and self.sup == other.sup and self.sub == other.sub and self.link == other.link
            and abs(self.size - other.size) < 0.6
        )


@dataclass
class Line:
    """One text line with its geometry."""

    spans: List[Span]
    bbox: Rect
    page: int
    unit: int
    size: float
    bold: bool
    head_x: List[float] = field(default_factory=list)
    edge: float = 0.0

    @property
    def text(self) -> str:
        return "".join(span.text for span in self.spans)


@dataclass
class Unit:
    """A block of lines as grouped by the PDF text extractor."""

    lines: List[Line]
    rotated: bool = False

    @property
    def bbox(self) -> Rect:
        return union([line.bbox for line in self.lines])


@dataclass
class ImageItem:
    """A raster image placed on a page."""

    bbox: Rect
    data: bytes
    ext: str
    key: Tuple[Any, ...]
    page: int = 0

    @property
    def width(self) -> float:
        return self.bbox[2] - self.bbox[0]


@dataclass
class Cell:
    """A table cell in grid coordinates."""

    r0: int
    r1: int
    c0: int
    c1: int
    rect: Rect
    lines: List[Line] = field(default_factory=list)
    paras: List["Para"] = field(default_factory=list)


@dataclass
class TableItem:
    """A detected table with its grid and the text lines inside it."""

    bbox: Rect
    rows: int
    cols: int
    col_edges: List[float]
    cells: List[Cell]
    page: int = 0
    header_rows: int = 0
    height: float = 0.0
    row_edges: List[float] = field(default_factory=list)
    grid: Dict[Tuple[int, int], Cell] = field(default_factory=dict)


@dataclass
class PageData:
    """Everything extracted from one page."""

    number: int
    width: float
    height: float
    units: List[Unit] = field(default_factory=list)
    tables: List[TableItem] = field(default_factory=list)
    images: List[ImageItem] = field(default_factory=list)
    figures: int = 0

    @property
    def has_text(self) -> bool:
        return bool(self.units) or bool(self.tables) or self.figures > 0


@dataclass
class Marker:
    """A list marker found at the start of a paragraph."""

    bullet: bool
    weak: bool
    interps: List[Tuple[str, int]]
    delim: str
    glyph: str
    length: int
    path: Tuple[int, ...] = ()


@dataclass
class Para:
    """A paragraph assembled from lines."""

    spans: List[Span]
    page: int
    seg: Tuple[Any, ...]
    bbox: Rect
    x0: float
    size: float
    bold: bool
    lines: int = 1
    marker: Optional[Marker] = None
    marker_x: float = 0.0
    text_x: float = 0.0
    kind: str = "para"
    level: int = 0
    rtl: bool = False
    num: Optional[Tuple[int, int]] = None
    bookmark: Optional[str] = None
    last_line_text: str = ""

    @property
    def text(self) -> str:
        return "".join(span.text for span in self.spans)


def union(rects: List[Rect]) -> Rect:
    return (
        min(r[0] for r in rects), min(r[1] for r in rects),
        max(r[2] for r in rects), max(r[3] for r in rects),
    )


Block = Union[Para, TableItem, ImageItem]


@dataclass
class RunningLine:
    """A header or footer line that repeats across pages; a None part is the page number."""

    parts: List[Optional[str]]
    size: float
    bold: bool
    align: str
    rtl: bool
    first_page: bool


@dataclass
class Running:
    """Repeating header and footer lines removed from the page bodies."""

    header: List[RunningLine] = field(default_factory=list)
    footer: List[RunningLine] = field(default_factory=list)
