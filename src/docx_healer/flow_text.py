"""Text cleaning and line-joining helpers for editable mode."""

import re
import unicodedata
from typing import List, Optional, Set, Tuple

_FOLD = re.compile("[\ufb00-\ufb06\ufb50-\ufdff\ufe70-\ufefe]")
_INVALID = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff\ud800-\udfff]")
_SPACES = re.compile("[ \t]+")
_WORD = re.compile(r"[^\W\d_]+(?:-[^\W\d_]+)*")
_TAIL_WORD = re.compile(r"([^\W\d_]+(?:-[^\W\d_]+)*)[-\u2010]$")
_HEAD_WORD = re.compile(r"^[^\W\d_]+")
_SOFT_HYPHEN = "\u00ad"
_HYPHENS = "-\u2010"
_NO_SPACE_RANGES = (
    (0x0E00, 0x0EFF), (0x1000, 0x109F), (0x1780, 0x17FF), (0x2E80, 0x2FDF), (0x3000, 0x30FF),
    (0x3130, 0x318F), (0x3400, 0x4DBF), (0x4E00, 0x9FFF), (0xF900, 0xFAFF), (0xFF00, 0xFFEF),
    (0x20000, 0x2FA1F),
)
_MIN_STEM_LEN = 2


def clean_text(text: str) -> str:
    """Fold ligatures and presentation forms and drop characters XML cannot hold."""
    text = _FOLD.sub(lambda m: unicodedata.normalize("NFKC", m.group()), text)
    return _INVALID.sub("", text)


def collapse_spaces(text: str) -> str:
    return _SPACES.sub(" ", text)


def is_no_space_char(char: str) -> bool:
    code = ord(char)
    return any(low <= code <= high for low, high in _NO_SPACE_RANGES)


def is_rtl_char(char: str) -> bool:
    return unicodedata.bidirectional(char) in ("R", "AL")


def is_ltr_char(char: str) -> bool:
    return unicodedata.bidirectional(char) == "L"


def vocabulary(groups: List[List[str]]) -> Set[str]:
    """Lower-cased whole words found in the document, ignoring pieces cut by a line-end hyphen."""
    words: Set[str] = set()
    for lines in groups:
        carried = False
        for text in lines:
            found = list(_WORD.finditer(text.strip()))
            stripped = text.strip()
            cut_tail = (
                bool(found) and stripped[-1:] in _HYPHENS and found[-1].end() == len(stripped) - 1
            )
            start = 1 if carried and found else 0
            end = len(found) - 1 if cut_tail else len(found)
            words.update(m.group().lower() for m in found[start:end])
            carried = cut_tail
    return words


def join_decision(previous: str, following: str, vocab: Set[str]) -> Tuple[bool, bool]:
    """Return (drop_trailing_hyphen, add_space) for joining two consecutive lines."""
    if not previous or not following:
        return False, False
    last, first = previous[-1], following[0]
    if last == _SOFT_HYPHEN:
        return True, False
    if is_no_space_char(last) or is_no_space_char(first):
        return False, False
    if last not in _HYPHENS:
        return False, not (previous.endswith(" ") or following.startswith(" "))
    tail = _TAIL_WORD.search(previous)
    head = _HEAD_WORD.match(following)
    if tail is None:
        return False, not previous.endswith(" ")
    stem = tail.group(1)
    if "-" in stem or head is None or not following[0].islower():
        return False, False
    rest = head.group().lower()
    low = stem.lower()
    if f"{low}-{rest}" in vocab:
        return False, False
    if f"{low}{rest}" in vocab:
        return True, False
    if len(low) < _MIN_STEM_LEN or low in vocab:
        return False, False
    return True, False


def strong_direction(text: str) -> Optional[str]:
    """Return "rtl" or "ltr" for the majority strong direction of the text, if any."""
    rtl = sum(1 for char in text if is_rtl_char(char))
    ltr = sum(1 for char in text if is_ltr_char(char))
    if rtl == 0 and ltr == 0:
        return None
    return "rtl" if rtl > ltr else "ltr"
