"""Heading detection from the PDF outline or from font statistics."""

import re
from typing import Dict, List, Optional, Tuple

from .flow_model import Para
from .log import logger

_MAX_LEVEL = 9
_MAX_LINES = 3
_MAX_CHARS = 200
_MAX_STAT_CHARS = 160
_SIZE_RATIO = 1.12
_BOLD_RATIO = 1.05
_SIZE_MERGE = 0.6
_MAX_SHARE = 0.4
_PREFIX = re.compile(r"^[0-9A-Za-z.]{1,8}\s+")
_OUTLINE_SHARE = 0.5
_MAX_PREFIX = 8
_TITLE_RATIO = 1.05
_TIGHT_GAP = 0.5

OutlineEntry = Tuple[int, str, int, Optional[float]]


def _norm(text: str) -> str:
    return "".join(ch for ch in text.casefold() if ch.isalnum())


def _title_matches(title: str, text: str) -> bool:
    wanted, found = _norm(title), _norm(text)
    if not wanted or not found:
        return False
    if wanted == found:
        return True
    if found.endswith(wanted) and len(found) - len(wanted) <= _MAX_PREFIX:
        return True
    stripped = _norm(_PREFIX.sub("", text, count=1))
    return bool(stripped) and stripped in (wanted, _norm(_PREFIX.sub("", title, count=1)))


def outline_headings(paras_by_page: Dict[int, List[Para]], toc: List[OutlineEntry]) -> int:
    """Mark paragraphs matching outline entries as headings; return how many matched."""
    used: set = set()
    matched = 0
    for level, title, page_number, dest_y in toc:
        best: Optional[Para] = None
        best_distance = float("inf")
        for para in paras_by_page.get(page_number - 1, ()):
            if id(para) in used or para.lines > _MAX_LINES or len(para.text) > _MAX_CHARS:
                continue
            if not _title_matches(title, para.text):
                continue
            distance = abs(para.bbox[1] - dest_y) if dest_y is not None else para.bbox[1]
            if distance < best_distance:
                best, best_distance = para, distance
        if best is not None:
            used.add(id(best))
            best.kind = "heading"
            best.level = max(1, min(_MAX_LEVEL, level))
            matched += 1
    return matched


def statistic_headings(paras: List[Para], body: float) -> int:
    """Mark short paragraphs set larger than the body text as headings by size rank."""
    candidates = _candidates(paras, body)
    if not candidates or len(candidates) > _MAX_SHARE * max(1, len(paras)):
        return 0
    sizes: List[float] = []
    for size in sorted({round(p.size * 2) / 2 for p in candidates}, reverse=True):
        if not sizes or sizes[-1] - size > _SIZE_MERGE:
            sizes.append(size)
    for para in candidates:
        rank = min(range(len(sizes)), key=lambda i: abs(sizes[i] - para.size))
        para.kind = "heading"
        para.level = min(_MAX_LEVEL, rank + 1)
    return len(candidates)


def _is_candidate(para: Para, body: float) -> bool:
    if para.kind != "para" or para.lines > _MAX_LINES or len(para.text) > _MAX_STAT_CHARS:
        return False
    if para.seg[1:2] == ("rotated",):
        return False
    if len(para.text.strip()) < 2:
        return False
    ratio = para.size / body if body else 1.0
    return ratio >= _SIZE_RATIO or (para.bold and ratio >= _BOLD_RATIO)


def _tight(first: Para, second: Para) -> bool:
    if not first.boxes or not second.boxes or abs(first.size - second.size) > _SIZE_MERGE:
        return False
    (seg_a, above, _), (seg_b, below, _) = first.boxes[-1], second.boxes[0]
    if seg_a[0] != seg_b[0] or min(above[2], below[2]) <= max(above[0], below[0]):
        return False
    return abs(below[1] - above[3]) <= _TIGHT_GAP * max(first.size, second.size)


def _candidates(paras: List[Para], body: float) -> List[Para]:
    """Return heading candidates, leaving out lines that are part of a tight block of text.

    A block is a run of paragraphs set in the same size with less than half a line between
    them, like a list of author names. It holds headings only when every paragraph in it
    looks like one and all of it together is no longer than one heading can be.
    """
    chosen: List[Para] = []
    group: List[Para] = []
    for index, para in enumerate(paras):
        group.append(para)
        following = paras[index + 1] if index + 1 < len(paras) else None
        if following is not None and _tight(para, following):
            continue
        looks = [p for p in group if _is_candidate(p, body)]
        lines = sum(p.lines for p in group)
        chars = sum(len(p.text) for p in group)
        if len(looks) == len(group) and lines <= _MAX_LINES and chars <= _MAX_STAT_CHARS:
            chosen.extend(looks)
        group = []
    return chosen


def assign_headings(paras: List[Para], toc: List[OutlineEntry], body: float) -> Tuple[str, int]:
    """Choose outline headings when the PDF has a usable outline, else statistics."""
    by_page: Dict[int, List[Para]] = {}
    for para in paras:
        by_page.setdefault(para.page, []).append(para)
    if toc:
        matched = outline_headings(by_page, toc)
        if matched >= max(1, _OUTLINE_SHARE * len(toc)):
            return "outline", matched + _extend_outline(paras, body)
        logger.debug("Outline matched %d of %d entries; using font statistics", matched, len(toc))
        for para in paras:
            if para.kind == "heading":
                para.kind, para.level = "para", 0
    return "statistics", statistic_headings(paras, body)


def _extend_outline(paras: List[Para], body: float) -> int:
    """Give outline-less paragraphs that look like the outline's headings the same level."""
    levels: Dict[Tuple[float, bool], int] = {}
    for para in paras:
        if para.kind == "heading":
            key = (round(para.size * 2) / 2, para.bold)
            levels[key] = min(levels.get(key, _MAX_LEVEL), para.level)
    if not levels:
        return 0
    largest = max(size for size, _ in levels)
    added = 0
    for para in _candidates(paras, body):
        if para.kind != "para":
            continue
        level = levels.get((round(para.size * 2) / 2, para.bold))
        if level is not None:
            para.kind, para.level = "heading", level
        elif para.size > _TITLE_RATIO * largest:
            para.kind = "title"
        else:
            continue
        added += 1
    return added
