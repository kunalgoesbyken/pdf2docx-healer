"""Parse the list marker that opens a line of text."""

import unicodedata
from dataclasses import dataclass
from typing import Dict, Optional, Tuple

from .constants import BULLET_CODEPOINTS, WEAK_BULLETS

Key = Tuple[str, str]

BULLET = "bullet"
LOWER_LETTER = "lowerLetter"
UPPER_LETTER = "upperLetter"
LOWER_ROMAN = "lowerRoman"
UPPER_ROMAN = "upperRoman"
DECIMAL = "decimal"

_MAX_DIGITS = 3
_MAX_ROMAN_LENGTH = 8
_ROMAN_DIGITS = (
    (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
    (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i"),
)
_ROMAN_VALUES = {"i": 1, "v": 5, "x": 10, "l": 50, "c": 100, "d": 500, "m": 1000}
_CLOSERS = {".": ".", ")": ")"}


@dataclass(frozen=True)
class Marker:
    """A list marker: its (numFmt, shape) readings, strength and where the text starts."""

    readings: Dict[Key, int]
    strong: bool
    end: int

    @property
    def is_bullet(self) -> bool:
        return any(fmt == BULLET for fmt, _ in self.readings)


def _roman(value: int) -> str:
    out = []
    for size, glyph in _ROMAN_DIGITS:
        count, value = divmod(value, size)
        out.append(glyph * count)
    return "".join(out)


def _roman_value(token: str) -> Optional[int]:
    if not token or any(char not in _ROMAN_VALUES for char in token):
        return None
    total = 0
    for index, char in enumerate(token):
        value = _ROMAN_VALUES[char]
        following = _ROMAN_VALUES[token[index + 1]] if index + 1 < len(token) else 0
        total += -value if value < following else value
    return total if total > 0 and _roman(total) == token else None


def _readings(token: str, shape: str) -> Dict[Key, int]:
    readings: Dict[Key, int] = {}
    if token.isdigit():
        value = int(token)
        if len(token) <= _MAX_DIGITS and value > 0 and str(value) == token:
            readings[(DECIMAL, shape)] = value
        return readings
    if not (token.islower() or token.isupper()) or len(token) > _MAX_ROMAN_LENGTH:
        return readings
    lower = token.islower()
    if len(token) == 1:
        value = ord(token.lower()) - ord("a") + 1
        readings[(LOWER_LETTER if lower else UPPER_LETTER, shape)] = value
    roman = _roman_value(token.lower())
    if roman is not None:
        readings[(LOWER_ROMAN if lower else UPPER_ROMAN, shape)] = roman
    return readings


def _content_start(text: str, index: int) -> Optional[int]:
    if index >= len(text) or not text[index].isspace():
        return None
    while index < len(text) and text[index].isspace():
        index += 1
    return index if index < len(text) else None


def _bullet(text: str, start: int, bare: bool) -> Optional[Marker]:
    glyph = text[start]
    strong = ord(glyph) in BULLET_CODEPOINTS or unicodedata.category(glyph) == "Co"
    if not strong and glyph not in WEAK_BULLETS:
        return None
    end = _content_start(text, start + 1)
    if end is None:
        if bare and strong and not text[start + 1:].strip():
            end = len(text)
        else:
            return None
    key = (BULLET, "" if strong else glyph)
    return Marker({key: 0}, strong, end)


def _number(text: str, start: int) -> Optional[Marker]:
    opened = text[start] == "("
    index = start + 1 if opened else start
    token_end = index
    while token_end < len(text) and text[token_end].isascii() and text[token_end].isalnum():
        token_end += 1
    if token_end == index or token_end >= len(text):
        return None
    closer = text[token_end]
    if opened:
        shape = "()" if closer == ")" else None
    else:
        shape = _CLOSERS.get(closer)
    if shape is None:
        return None
    readings = _readings(text[index:token_end], shape)
    end = _content_start(text, token_end + 1)
    if not readings or end is None:
        return None
    return Marker(readings, False, end)


def parse_marker(text: str, bare: bool = False) -> Optional[Marker]:
    """Return the marker that starts ``text`` (``bare`` accepts a lone bullet glyph), else None."""
    start = len(text) - len(text.lstrip())
    if start >= len(text):
        return None
    return _bullet(text, start, bare) or _number(text, start)
