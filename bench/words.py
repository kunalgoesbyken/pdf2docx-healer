"""Word tokenizing and bounded, order-aware alignment of two word sequences."""

import unicodedata
from bisect import bisect_left
from collections import Counter
from typing import Dict, List, Sequence, Tuple

WINDOW = 32
CONFIRM = 2
GRAMS = (3, 2, 1)
MIN_PAGE_WORDS = 5
PAGE_LOST_BELOW = 0.5
SHINGLE = 8
JOINERS = "‌‍"
SINGLE_CHAR_RANGES = (
    (0x3040, 0x30FF),
    (0x3400, 0x4DBF),
    (0x4E00, 0x9FFF),
    (0xF900, 0xFAFF),
    (0x20000, 0x2FA1F),
)


def is_single_char_script(char: str) -> bool:
    """Return True for scripts written without spaces, where one character is one word."""
    code = ord(char)
    return any(low <= code <= high for low, high in SINGLE_CHAR_RANGES)


def tokenize(text: str) -> List[str]:
    """Split text into folded words; CJK is one word per character."""
    tokens: List[str] = []
    current: List[str] = []
    for char in unicodedata.normalize("NFKC", text).casefold():
        if char in JOINERS:
            continue
        if is_single_char_script(char):
            if current:
                tokens.append("".join(current))
                current = []
            tokens.append(char)
        elif unicodedata.category(char)[0] in "LNM":
            current.append(char)
        elif current:
            tokens.append("".join(current))
            current = []
    if current:
        tokens.append("".join(current))
    return tokens


def confirmed(
    ref: Sequence[str], hyp: Sequence[str], i: int, j: int, ref_end: int, hyp_end: int
) -> bool:
    """Check that the words after a candidate resync point also agree."""
    for step in range(1, CONFIRM + 1):
        if i + step >= ref_end or j + step >= hyp_end:
            return True
        if ref[i + step] != hyp[j + step]:
            return False
    return True


def resync(
    ref: Sequence[str], hyp: Sequence[str], span: Tuple[int, int, int, int], window: int
) -> Tuple[int, int]:
    """Find the cheapest skip (ref words, hyp words) that lands on matching text, or (1, 1)."""
    i, ref_end, j, hyp_end = span
    offsets: Dict[str, List[int]] = {}
    for offset in range(min(window, hyp_end - j)):
        offsets.setdefault(hyp[j + offset], []).append(offset)
    best_cost = None
    best = (1, 1)
    for di in range(min(window, ref_end - i)):
        if best_cost is not None and di >= best_cost:
            break
        for dj in offsets.get(ref[i + di], ()):
            cost = di + dj
            if best_cost is not None and cost >= best_cost:
                break
            if confirmed(ref, hyp, i + di, j + dj, ref_end, hyp_end):
                best_cost = cost
                best = (di, dj)
                break
    return best


def local_align(
    ref: Sequence[str],
    hyp: Sequence[str],
    span: Tuple[int, int, int, int],
    flags: Tuple[bytearray, bytearray],
    window: int,
) -> None:
    """Two-pointer diff that can only skip `window` words at a time; used between anchors."""
    i, ref_end, j, hyp_end = span
    while i < ref_end and j < hyp_end:
        if ref[i] == hyp[j]:
            flags[0][i] = 1
            flags[1][j] = 1
            i += 1
            j += 1
            continue
        di, dj = resync(ref, hyp, (i, ref_end, j, hyp_end), window)
        i += di
        j += dj


def unique_grams(
    words: Sequence[str], start: int, end: int, size: int
) -> Dict[Tuple[str, ...], int]:
    """Map each n-gram that occurs exactly once in words[start:end] to its position."""
    seen: Dict[Tuple[str, ...], int] = {}
    for k in range(start, end - size + 1):
        gram = tuple(words[k : k + size])
        seen[gram] = -1 if gram in seen else k
    return {gram: pos for gram, pos in seen.items() if pos >= 0}


def longest_increasing(pairs: List[Tuple[int, int]]) -> List[Tuple[int, int]]:
    """Return the longest subsequence of (i, j) pairs, sorted by i, whose j values increase."""
    tails: List[int] = []
    tail_index: List[int] = []
    previous = [-1] * len(pairs)
    for index, (_, j) in enumerate(pairs):
        slot = bisect_left(tails, j)
        if slot:
            previous[index] = tail_index[slot - 1]
        if slot == len(tails):
            tails.append(j)
            tail_index.append(index)
        else:
            tails[slot] = j
            tail_index[slot] = index
    chain: List[Tuple[int, int]] = []
    cursor = tail_index[-1] if tail_index else -1
    while cursor >= 0:
        chain.append(pairs[cursor])
        cursor = previous[cursor]
    chain.reverse()
    return chain


def find_anchors(
    ref: Sequence[str], hyp: Sequence[str], span: Tuple[int, int, int, int]
) -> Tuple[int, List[Tuple[int, int]]]:
    """Return (n, ordered anchor pairs) for the longest n-gram size that has unique shared grams."""
    rs, re_, hs, he = span
    for size in GRAMS:
        ref_grams = unique_grams(ref, rs, re_, size)
        hyp_grams = unique_grams(hyp, hs, he, size)
        pairs = sorted(
            (pos, hyp_grams[gram]) for gram, pos in ref_grams.items() if gram in hyp_grams
        )
        if pairs:
            return size, longest_increasing(pairs)
    return 0, []


def align(
    ref: Sequence[str], hyp: Sequence[str], window: int = WINDOW
) -> Tuple[bytearray, bytearray]:
    """Mark the words of both sequences that match in order.

    Equal prefixes and suffixes are trimmed, unique shared n-grams anchor the rest (longest
    increasing chain, so a moved block is lost), and the gaps use a windowed diff: O(n log n).
    """
    flags = (bytearray(len(ref)), bytearray(len(hyp)))
    work = [(0, len(ref), 0, len(hyp))]
    while work:
        rs, re_, hs, he = work.pop()
        while rs < re_ and hs < he and ref[rs] == hyp[hs]:
            flags[0][rs] = flags[1][hs] = 1
            rs += 1
            hs += 1
        while rs < re_ and hs < he and ref[re_ - 1] == hyp[he - 1]:
            flags[0][re_ - 1] = flags[1][he - 1] = 1
            re_ -= 1
            he -= 1
        if rs >= re_ or hs >= he:
            continue
        size, anchors = find_anchors(ref, hyp, (rs, re_, hs, he))
        if not anchors:
            local_align(ref, hyp, (rs, re_, hs, he), flags, window)
            continue
        prev_r, prev_h = rs, hs
        for i, j in anchors:
            for offset in range(size):
                flags[0][i + offset] = flags[1][j + offset] = 1
            if i > prev_r or j > prev_h:
                work.append((prev_r, max(i, prev_r), prev_h, max(j, prev_h)))
            prev_r = max(prev_r, i + size)
            prev_h = max(prev_h, j + size)
        work.append((prev_r, re_, prev_h, he))
    return flags


def ratio(part: int, whole: int) -> float:
    """Return part over whole, defined as 1.0 when both are empty and 0.0 for an empty whole."""
    if whole == 0:
        return 1.0 if part == 0 else 0.0
    return part / whole


def word_scores(ref: Sequence[str], hyp: Sequence[str]) -> Dict[str, float]:
    """Return ordered recall and precision plus the order-free recall of hyp against ref."""
    ref_flags, hyp_flags = align(ref, hyp)
    available = Counter(hyp)
    unordered = 0
    for word, count in Counter(ref).items():
        unordered += min(count, available.get(word, 0))
    if hyp:
        precision = sum(hyp_flags) / len(hyp)
    else:
        precision = 0.0 if ref else 1.0
    return {
        "recall": ratio(sum(ref_flags), len(ref)),
        "precision": precision,
        "recall_unordered": ratio(unordered, len(ref)),
    }


def shingles(text: str) -> List[int]:
    """Return hashes of every SHINGLE-character window of text."""
    return [hash(text[k : k + SHINGLE]) for k in range(len(text) - SHINGLE + 1)]


def pages_lost(pages: Sequence[Sequence[str]], hyp: Sequence[str]) -> Tuple[int, int]:
    """Return (pages with text, pages whose text is mostly absent from hyp).

    A page is lost when under half of its 8-character windows, taken over the page words
    with spaces removed, occur anywhere in the output text. Word fusing, splitting and
    reordering inside a page do not matter; a dropped page does.
    """
    present = set(shingles("".join(hyp)))
    text_pages = 0
    lost = 0
    for words in pages:
        text = "".join(words)
        if len(words) < MIN_PAGE_WORDS or len(text) < SHINGLE:
            continue
        text_pages += 1
        windows = shingles(text)
        found = sum(1 for window in windows if window in present)
        if found / len(windows) < PAGE_LOST_BELOW:
            lost += 1
    return text_pages, lost
