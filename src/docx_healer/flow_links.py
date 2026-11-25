"""Resolving internal PDF links to bookmarks on the paragraphs they point at."""

from typing import Dict, Iterator, List

from .flow_model import Block, LinkKey, Para, TableItem

_REACH = 6.0


def iter_paras(blocks: List[Block]) -> Iterator[Para]:
    for block in blocks:
        if isinstance(block, Para):
            yield block
        elif isinstance(block, TableItem):
            for cell in block.cells:
                yield from cell.paras


def _distance(para: Para, x: float, y: float) -> float:
    left, top, right, bottom = para.bbox
    if bottom < y - 2:
        return float("inf")
    dy = max(0.0, top - y)
    dx = max(0.0, left - 5 - x, x - right - 5)
    return dy + 2 * dx


def resolve_anchors(blocks: List[Block], body: float) -> Dict[LinkKey, str]:
    """Attach a bookmark to the paragraph each internal link points at."""
    keys = {
        span.link for para in iter_paras(blocks) for span in para.spans
        if span.link is not None and span.link[0] == "goto"
    }
    by_page: Dict[int, List[Para]] = {}
    for block in blocks:
        if isinstance(block, Para):
            by_page.setdefault(block.page, []).append(block)
    anchors: Dict[LinkKey, str] = {}
    created = 0
    for key in sorted(keys, key=str):
        _, page, x, y = key
        best, best_distance = None, _REACH * body
        for para in by_page.get(page, ()):
            distance = _distance(para, x, y)
            if distance < best_distance:
                best, best_distance = para, distance
        if best is None:
            continue
        if best.bookmark is None:
            created += 1
            best.bookmark = f"_pdf_ref_{created}"
        anchors[key] = best.bookmark
    return anchors
