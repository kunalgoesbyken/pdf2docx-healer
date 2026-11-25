"""Render the visible number of a numbered paragraph from numbering.xml, following ECMA-376."""

import re
from typing import Any, Dict, List, NamedTuple, Optional, Set, Tuple

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
ROMAN = (
    (1000, "m"),
    (900, "cm"),
    (500, "d"),
    (400, "cd"),
    (100, "c"),
    (90, "xc"),
    (50, "l"),
    (40, "xl"),
    (10, "x"),
    (9, "ix"),
    (5, "v"),
    (4, "iv"),
    (1, "i"),
)
PLACEHOLDER = re.compile(r"%([1-9])")
MAX_ROMAN = 3999
LEVELS = 9


class Label(NamedTuple):
    text: str
    kind: str


class Level(NamedTuple):
    start: int
    fmt: str
    text: str
    restart: Optional[int]
    legal: bool
    style: Optional[str]


def to_roman(value: int) -> str:
    """Return the lowercase roman numeral of value, or an empty string outside 1..3999."""
    if value < 1 or value > MAX_ROMAN:
        return ""
    parts = []
    for number, symbol in ROMAN:
        while value >= number:
            parts.append(symbol)
            value -= number
    return "".join(parts)


def to_letters(value: int) -> str:
    """Return a, b, ... z, aa, bb, ... as Word counts letters, or empty for value below 1."""
    if value < 1:
        return ""
    return chr(ord("a") + (value - 1) % 26) * ((value - 1) // 26 + 1)


def format_number(value: int, fmt: str) -> Tuple[str, bool]:
    """Return (text, approximated) for value in a numFmt; unknown formats render as decimal."""
    if fmt == "decimal":
        return str(value), False
    if fmt == "decimalZero":
        return f"{value:02d}", False
    if fmt == "lowerRoman":
        return to_roman(value), False
    if fmt == "upperRoman":
        return to_roman(value).upper(), False
    if fmt == "lowerLetter":
        return to_letters(value), False
    if fmt == "upperLetter":
        return to_letters(value).upper(), False
    return str(value), True


def attr(element: Any, name: str) -> Optional[str]:
    """Return the w:val of a child element, or None when the child is absent."""
    child = element.find(f"{W}{name}")
    return None if child is None else child.get(f"{W}val")


def parse_level(lvl: Any) -> Level:
    """Read one w:lvl element, applying the spec defaults for omitted children."""
    start = attr(lvl, "start")
    restart = attr(lvl, "lvlRestart")
    text = attr(lvl, "lvlText")
    return Level(
        start=int(start) if start is not None else 0,
        fmt=attr(lvl, "numFmt") or "decimal",
        text=text if text is not None else "",
        restart=int(restart) if restart is not None else None,
        legal=lvl.find(f"{W}isLgl") is not None,
        style=attr(lvl, "pStyle"),
    )


def levels_of(parent: Any) -> Dict[int, Level]:
    """Map ilvl to Level for the w:lvl children of an abstractNum."""
    found: Dict[int, Level] = {}
    for lvl in parent.findall(f"{W}lvl"):
        found[int(lvl.get(f"{W}ilvl", "0"))] = parse_level(lvl)
    return found


class NumberingRenderer:
    """Walks numbered paragraphs in document order and yields the label each one displays."""

    def __init__(self, root: Any) -> None:
        self.approx_formats: Set[str] = set()
        self.abstract_levels: Dict[int, Dict[int, Level]] = {}
        self.num_abstract: Dict[int, int] = {}
        self.num_overrides: Dict[int, Dict[int, Tuple[Optional[int], Optional[Level]]]] = {}
        self.counters: Dict[int, List[Optional[int]]] = {}
        self.started: Set[int] = set()
        if root is not None:
            self.load(root)

    def load(self, root: Any) -> None:
        """Index abstractNum and num elements."""
        for abstract in root.findall(f"{W}abstractNum"):
            self.abstract_levels[int(abstract.get(f"{W}abstractNumId", "-1"))] = levels_of(abstract)
        for num in root.findall(f"{W}num"):
            num_id = int(num.get(f"{W}numId", "-1"))
            ref = attr(num, "abstractNumId")
            if ref is None:
                continue
            self.num_abstract[num_id] = int(ref)
            overrides = {}
            for override in num.findall(f"{W}lvlOverride"):
                start = attr(override, "startOverride")
                lvl = override.find(f"{W}lvl")
                overrides[int(override.get(f"{W}ilvl", "0"))] = (
                    int(start) if start is not None else None,
                    parse_level(lvl) if lvl is not None else None,
                )
            self.num_overrides[num_id] = overrides

    def level(self, num_id: int, ilvl: int) -> Optional[Level]:
        """Return the effective level definition, with lvlOverride replacing the abstract one."""
        override = self.num_overrides.get(num_id, {}).get(ilvl)
        if override is not None and override[1] is not None:
            return override[1]
        return self.abstract_levels.get(self.num_abstract.get(num_id, -1), {}).get(ilvl)

    def ilvl_for_style(self, num_id: int, style_id: str) -> Optional[int]:
        """Return the level whose pStyle is style_id, as a numbered paragraph style requires."""
        levels = self.abstract_levels.get(self.num_abstract.get(num_id, -1), {})
        for ilvl, level in sorted(levels.items()):
            if level.style == style_id:
                return ilvl
        return None

    def apply_start_overrides(self, num_id: int, abstract_id: int) -> None:
        """Seed counters from startOverride the first time a num instance is used."""
        if num_id in self.started:
            return
        self.started.add(num_id)
        counters = self.counters.setdefault(abstract_id, [None] * LEVELS)
        for ilvl, (start, lvl) in self.num_overrides.get(num_id, {}).items():
            value = start if start is not None else (lvl.start if lvl is not None else None)
            if value is not None and ilvl < LEVELS:
                counters[ilvl] = value - 1

    def next_label(self, num_id: int, ilvl: int) -> Optional[Label]:
        """Advance the counters for one numbered paragraph and return its label."""
        abstract_id = self.num_abstract.get(num_id)
        if num_id == 0 or abstract_id is None or abstract_id not in self.abstract_levels:
            return None
        level = self.level(num_id, ilvl)
        if level is None or not 0 <= ilvl < LEVELS:
            return None
        self.apply_start_overrides(num_id, abstract_id)
        counters = self.counters[abstract_id]
        previous = counters[ilvl]
        counters[ilvl] = level.start if previous is None else previous + 1
        self.restart_deeper(num_id, abstract_id, ilvl)
        if level.fmt == "bullet":
            return Label(level.text, "bullet")
        if level.fmt == "none":
            return Label("", "none")
        return Label(self.render(num_id, abstract_id, ilvl, level), "number")

    def restart_deeper(self, num_id: int, abstract_id: int, ilvl: int) -> None:
        """Reset the counters of deeper levels as their lvlRestart rule demands."""
        counters = self.counters[abstract_id]
        for deeper in range(ilvl + 1, LEVELS):
            deeper_level = self.level(num_id, deeper)
            restart_after = deeper - 1
            if deeper_level is not None and deeper_level.restart is not None:
                if deeper_level.restart == 0:
                    continue
                if deeper_level.restart - 1 < deeper:
                    restart_after = deeper_level.restart - 1
            if ilvl <= restart_after:
                counters[deeper] = None

    def render(self, num_id: int, abstract_id: int, ilvl: int, level: Level) -> str:
        """Substitute %1..%9 in lvlText with the formatted counter of each referenced level."""
        counters = self.counters[abstract_id]

        def substitute(match: Any) -> str:
            ref = int(match.group(1)) - 1
            if ref > ilvl:
                return ""
            ref_level = self.level(num_id, ref)
            if ref_level is None:
                return ""
            value = counters[ref]
            if value is None:
                value = ref_level.start
            fmt = "decimal" if level.legal else ref_level.fmt
            text, approx = format_number(value, fmt)
            if approx:
                self.approx_formats.add(fmt)
            return text

        return PLACEHOLDER.sub(substitute, level.text)
