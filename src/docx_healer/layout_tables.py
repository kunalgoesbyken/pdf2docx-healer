"""Keep pdf2docx from turning text layout into Word tables.

pdf2docx proposes a stream table for every group of text lines that do not flow as one
column, which fabricates tables for multi-column prose, wrapped definition lists, list
markers and text pieces inside cells. The proposals come from one method,
``Blocks.collect_stream_lines``, before any block is moved into a table, so that is the
one place filtered here.
"""

import functools
import importlib
import statistics
from contextlib import contextmanager
from dataclasses import dataclass
from typing import Any, Callable, Iterator, List, Optional

from .log import logger
from .patches import PATCH_LOCK

_BLOCKS_MODULE = "pdf2docx.layout.Blocks"
_SHARE_MODULE = "pdf2docx.common.share"
_COLLECT = "collect_stream_lines"


@dataclass(frozen=True)
class TableRules:
    """When a group of text lines still counts as a table."""

    min_rows: int = 3
    min_columns: int = 3
    min_row_share: float = 0.5
    max_cell_chars: int = 15
    max_group_lines: int = 4000


def _is_shaded(lines: Any, shadings: List[Any]) -> bool:
    return any(shading.bbox.intersects(lines.bbox) for shading in shadings)


def _visible_shadings(potential: Any, white: int) -> List[Any]:
    return [s for s in potential if not s.is_determined and s.color != white]


def _median_cell_chars(lines: Any) -> float:
    return statistics.median(len(line.text.strip()) for line in lines)


def looks_tabular(lines: Any, rules: TableRules) -> bool:
    """Say whether text lines grouped by pdf2docx are laid out like a table of data.

    Needs at least ``min_columns`` columns and ``min_rows`` rows, a row count close to the
    count of physical text rows (wrapped prose has far fewer), and cells that hold a label
    or a number rather than a sentence (at most ``max_cell_chars`` characters at the median).
    """
    columns = len(lines.group_by_columns())
    if columns < rules.min_columns:
        return False
    rows = len(lines.group_by_rows())
    if rows < rules.min_rows:
        return False
    if rows < rules.min_row_share * len(lines.group_by_physical_rows()):
        return False
    return _median_cell_chars(lines) <= rules.max_cell_chars


def _resolve_collect() -> Optional[Callable[..., Any]]:
    try:
        blocks = importlib.import_module(_BLOCKS_MODULE).Blocks
    except (ImportError, AttributeError):
        return None
    collect = vars(blocks).get(_COLLECT)
    return collect if callable(collect) else None


def _white() -> Optional[int]:
    try:
        share = importlib.import_module(_SHARE_MODULE)
        return int(share.rgb_value([1, 1, 1]))
    except (ImportError, AttributeError):
        return None


@contextmanager
def layout_table_filter(rules: TableRules = TableRules()) -> Iterator[None]:
    """Drop stream-table candidates that are not tabular, then restore pdf2docx.

    A candidate is kept as pdf2docx proposed it when it overlaps a visible shading (a
    callout or header band), holds a rotated line (pdf2docx draws that as a table cell) or
    has more than ``max_group_lines`` lines (it is not analysed, with one debug log).
    Does nothing when the installed pdf2docx has a different layout.
    """
    with PATCH_LOCK:
        original = _resolve_collect()
        white = _white()
        if original is None or white is None:
            logger.debug("pdf2docx stream table collector not found; layout tables not filtered")
            yield
            return

        blocks_class = importlib.import_module(_BLOCKS_MODULE).Blocks

        @functools.wraps(original)
        def collect(self: Any, potential_shadings: Any, *args: Any, **kwargs: Any) -> Any:
            groups = original(self, potential_shadings, *args, **kwargs)
            shadings = _visible_shadings(potential_shadings, white)
            return [group for group in groups if _keep(group, shadings, rules)]

        blocks_class.collect_stream_lines = collect
        try:
            yield
        finally:
            blocks_class.collect_stream_lines = original


def _keep(lines: Any, shadings: List[Any], rules: TableRules) -> bool:
    if not lines:
        return False
    if len(lines) > rules.max_group_lines:
        logger.debug("Stream table candidate of %d lines is over max_group_lines", len(lines))
        return True
    if any(line.is_vertical_text for line in lines) or _is_shaded(lines, shadings):
        return True
    return looks_tabular(lines, rules)
