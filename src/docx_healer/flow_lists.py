"""List marker parsing and numbering plans for editable mode."""

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from .constants import BULLET_CODEPOINTS
from .flow_model import Block, Marker, Para

_BALLOT_BOXES = frozenset({0x2610, 0x2611, 0x2612})
_WEAK_GLYPHS = frozenset("-*+\u2013\u2014")
_SHAPES = range(0x25A0, 0x2600)
_STRONG_GLYPHS = frozenset(
    chr(cp) for cp in [*BULLET_CODEPOINTS, *_SHAPES] if cp not in _BALLOT_BOXES
) - _WEAK_GLYPHS
_PRIVATE_USE = (0xE000, 0xF8FF)
_NUMBER = re.compile(r"^(?P<open>\()?(?P<tok>\d{1,3}|[A-Za-z]{1,6})(?P<close>[.)])(?=\s)\s*")
_ROMAN_VALUES = {"i": 1, "v": 5, "x": 10, "l": 50, "c": 100, "d": 500, "m": 1000}
_MAX_LEVEL = 8
_MIN_INDENT_TWIPS = 200
_MAX_INDENT_TWIPS = 720
_LEVEL_TOLERANCE = 0.6
_DEFAULT_HANG = 360
_DEFAULT_STEP = 360


def parse_marker(text: str) -> Optional[Marker]:
    """Return the list marker at the start of ``text`` or None."""
    if not text:
        return None
    first = text[0]
    code = ord(first)
    strong = first in _STRONG_GLYPHS or _PRIVATE_USE[0] <= code <= _PRIVATE_USE[1]
    if strong or first in _WEAK_GLYPHS:
        rest = text[1:]
        if not rest.strip():
            return None
        if not strong and not rest[0].isspace():
            return None
        length = 1 + len(rest) - len(rest.lstrip())
        private = _PRIVATE_USE[0] <= code <= _PRIVATE_USE[1]
        glyph = "\u2022" if private or first == "\u00b7" else first
        return Marker(True, not strong, [], "", glyph, length)
    match = _NUMBER.match(text)
    if match is None or not text[match.end():].strip():
        return None
    opened, close = bool(match.group("open")), match.group("close")
    if opened and close != ")":
        return None
    interps = _interpretations(match.group("tok"))
    if not interps:
        return None
    delim = "(%)" if opened else f"%{close}"
    return Marker(False, False, interps, delim, "", match.end())


def _roman(token: str) -> Optional[int]:
    if token != token.lower() and token != token.upper():
        return None
    values = [_ROMAN_VALUES.get(ch, 0) for ch in token.lower()]
    if 0 in values:
        return None
    total = sum(-v if v < following else v for v, following in zip(values, values[1:] + [0]))
    return total if _to_roman(total) == token.lower() else None


def _to_roman(value: int) -> str:
    pairs = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
             (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
    out = ""
    for number, symbol in pairs:
        while value >= number:
            out += symbol
            value -= number
    return out


def _interpretations(token: str) -> List[Tuple[str, int]]:
    if token.isdigit():
        if len(token) > 1 and token.startswith("0"):
            return []
        return [("decimal", int(token))]
    if not token.isalpha():
        return []
    result: List[Tuple[str, int]] = []
    lower = token.islower()
    if len(token) == 1:
        index = ord(token.lower()) - ord("a") + 1
        result.append(("lowerLetter" if lower else "upperLetter", index))
    roman = _roman(token)
    if roman is not None:
        result.append(("lowerRoman" if lower else "upperRoman", roman))
    return result


@dataclass
class LevelDef:
    """Numbering definition of one level of a list instance."""

    bullet: bool
    fmt: str
    start: int
    delim: str
    glyph: str
    marker_x: float
    text_x: float


@dataclass
class ListInstance:
    """One Word numbering instance and the paragraphs that belong to it."""

    levels: Dict[int, LevelDef] = field(default_factory=dict)
    counters: Dict[int, int] = field(default_factory=dict)
    members: List[Tuple[Para, int]] = field(default_factory=list)


def _levels_for(region: List[Para], size: float) -> List[int]:
    xs = sorted({round(p.marker_x, 1) for p in region})
    clusters: List[float] = []
    for x in xs:
        if not clusters or x - clusters[-1] > _LEVEL_TOLERANCE * size:
            clusters.append(x)
    return [
        max(i for i, c in enumerate(clusters) if c <= p.marker_x + _LEVEL_TOLERANCE * size)
        for p in region
    ]


def _signature(marker: Marker, level: LevelDef) -> bool:
    if marker.bullet:
        return level.bullet and level.glyph == marker.glyph
    return (not level.bullet) and level.delim == marker.delim


def _next_sibling_fits(
    region: List[Para], levels: List[int], index: int, interp: Tuple[str, int]
) -> bool:
    for later in range(index + 1, len(region)):
        if levels[later] < levels[index]:
            return False
        if levels[later] == levels[index]:
            marker = region[later].marker
            return marker is not None and (interp[0], interp[1] + 1) in marker.interps
    return False


def _place(instance: ListInstance, para: Para, level: int, region: List[Para],
           levels: List[int], index: int) -> bool:
    marker = para.marker
    assert marker is not None
    existing = instance.levels.get(level)
    if marker.bullet:
        if existing is None:
            instance.levels[level] = LevelDef(
                True, "bullet", 1, "", marker.glyph, para.marker_x, para.text_x,
            )
        elif not _signature(marker, existing):
            return False
        for deeper in [k for k in instance.counters if k > level]:
            del instance.counters[deeper]
        instance.counters[level] = instance.counters.get(level, 0) + 1
        return True
    if existing is None:
        choice = next(
            (i for i in marker.interps if _next_sibling_fits(region, levels, index, i)),
            marker.interps[0],
        )
        instance.levels[level] = LevelDef(
            False, choice[0], choice[1], marker.delim, "", para.marker_x, para.text_x,
        )
        value = choice[1]
    else:
        if not _signature(marker, existing):
            return False
        expected = instance.counters.get(level, existing.start - 1) + 1
        if (existing.fmt, expected) not in marker.interps:
            return False
        value = expected
    for deeper in [k for k in instance.counters if k > level]:
        del instance.counters[deeper]
    instance.counters[level] = value
    return True


def plan_region(region: List[Para], size: float) -> List[ListInstance]:
    """Split a run of marker paragraphs into validated list instances."""
    levels = _levels_for(region, size)
    instances: List[ListInstance] = []
    current: Optional[ListInstance] = None
    for index, para in enumerate(region):
        level = levels[index]
        if level > _MAX_LEVEL:
            continue
        if current is not None and _place(current, para, level, region, levels, index):
            current.members.append((para, level))
            continue
        if current is not None and level > min(m[1] for m in current.members):
            continue
        current = ListInstance()
        instances.append(current)
        if _place(current, para, level, region, levels, index):
            current.members.append((para, level))
    return [i for i in instances if _accepted(i)]


def _accepted(instance: ListInstance) -> bool:
    base = min((level for _, level in instance.members), default=0)
    tops = [p for p, level in instance.members if level == base]
    if not tops:
        return False
    marker = tops[0].marker
    assert marker is not None
    if marker.bullet and not marker.weak:
        return True
    return len(tops) >= 2


def level_indents(definition: LevelDef, margin: float) -> Tuple[int, int]:
    """Return (left, hanging) in twips for a level."""
    hanging = int((definition.text_x - definition.marker_x) * 20)
    hanging = max(_MIN_INDENT_TWIPS, min(_MAX_INDENT_TWIPS, hanging or _DEFAULT_HANG))
    left = int(max(0.0, definition.marker_x - margin) * 20) + hanging
    return left, hanging


def _strip_marker(para: Para) -> None:
    remaining = para.marker.length if para.marker is not None else 0
    kept = []
    for span in para.spans:
        if remaining >= len(span.text):
            remaining -= len(span.text)
            continue
        if remaining:
            span.text = span.text[remaining:]
            remaining = 0
        kept.append(span)
    para.spans = kept


def apply_lists(blocks: List[Block]) -> List[ListInstance]:
    """Turn validated marker paragraphs into list items and return the numbering instances."""
    found: List[ListInstance] = []
    region: List[Para] = []

    def close() -> None:
        if not region:
            return
        for instance in plan_region(region, region[0].size):
            index = len(found)
            found.append(instance)
            for para, level in instance.members:
                _strip_marker(para)
                para.kind, para.level, para.num = "list", level, (index, level)
        region.clear()

    for block in blocks:
        if isinstance(block, Para) and block.kind == "para" and block.marker is not None:
            region.append(block)
        else:
            close()
    close()
    return found
