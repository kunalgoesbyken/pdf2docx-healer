"""Heading styles from the PDF outline and font sizes."""

import os
import unicodedata
from collections import Counter
from typing import Any, Dict, List, NamedTuple, Optional, Set, Tuple

from docx import Document
from docx.document import Document as DocumentType
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from .docx_walk import fold_text, paged_body_elements, paragraph_plain_text
from .files import PathLike
from .log import progress
from .ocr_pipeline import pymupdf

MAX_KEY_CHARS = 300
MAX_NUMBER_PREFIX_CHARS = 8
MAX_TITLE_PARAGRAPHS = 3
MAX_HEADING_LEVEL = 9
FALLBACK_LEVELS = 4
FALLBACK_MAX_LINE_CHARS = 120
FALLBACK_MAX_BLOCK_LINES = 3
FALLBACK_MAX_CANDIDATE_SHARE = 0.25
FALLBACK_MIN_REPEAT_PAGES = 3
FALLBACK_MIN_SIZE_RATIO = 1.1
FALLBACK_MIN_SIZE_GAIN = 1.0
SIZE_STEP = 0.5

_IGNORED_CATEGORIES = ("P", "S")


class _Target(NamedTuple):
    level: int
    key: str
    page: int


class _Item(NamedTuple):
    position: int
    key: str


def fix_docx_headings(docx_path: PathLike, pdf_path: PathLike, verbose: bool = False) -> int:
    """Apply Word heading styles to paragraphs that are headings in the PDF, in place.

    Headings come from the PDF outline, or from font sizes when there is none. Only
    body paragraphs are styled, never table cells, headers or footers. Returns the
    number of headings styled.
    """
    doc = Document(str(docx_path))
    with pymupdf.open(os.fspath(pdf_path)) as pdf:
        targets = _outline_targets(pdf) or _font_targets(pdf)
        page_count = pdf.page_count
    if not targets:
        return 0

    pages = paged_body_elements(doc)
    default_style = doc.styles.default(WD_STYLE_TYPE.PARAGRAPH)
    default_id = default_style.style_id if default_style is not None else None
    runs_by_page, paragraphs = _paragraph_runs(pages, default_id)
    trusted = len(pages) == page_count
    matched = _match(targets, runs_by_page, trusted)

    styles: Dict[int, Any] = {}
    count = 0
    for target, group in matched:
        style = styles.get(target.level)
        if style is None:
            style = styles[target.level] = _heading_style(doc, target.level)
        for index in group:
            Paragraph(paragraphs[index], doc).style = style
        count += 1
    if count:
        doc.save(str(docx_path))
        progress(verbose, f"Styled {count} heading(s)")
    return count


def _title_key(text: str) -> str:
    folded = fold_text(text)
    return "".join(
        char for char in folded if unicodedata.category(char)[0] not in _IGNORED_CATEGORIES
    )


def _without_number(key: str) -> str:
    """Drop a leading section number such as ``1``, ``3.2.1`` or ``A.1`` from a title key."""
    lead = 1 if len(key) > 1 and key[0].isalpha() and unicodedata.category(key[1]) == "Nd" else 0
    limit = min(len(key), MAX_NUMBER_PREFIX_CHARS)
    while lead < limit and unicodedata.category(key[lead]) == "Nd":
        lead += 1
    if lead == 0 or lead >= len(key):
        return key
    return key[lead:]


def _variants(key: str) -> List[str]:
    bare = _without_number(key)
    return [key] if bare == key else [key, bare]


def _outline_targets(pdf: Any) -> List[_Target]:
    targets = []
    for level, title, page in pdf.get_toc(simple=True):
        key = _title_key(title)
        if key and len(key) <= MAX_KEY_CHARS:
            targets.append(_Target(min(max(level, 1), MAX_HEADING_LEVEL), key, page))
    return targets


def _round_size(size: float) -> float:
    return round(size / SIZE_STEP) * SIZE_STEP


def _block_lines(block: Dict[str, Any]) -> List[Tuple[str, float]]:
    lines = []
    for line in block.get("lines", []):
        spans = [span for span in line["spans"] if span["text"].strip()]
        if not spans:
            continue
        main = max(spans, key=lambda span: len(span["text"].strip()))
        lines.append(
            ("".join(span["text"] for span in line["spans"]).strip(), _round_size(main["size"]))
        )
    return lines


def _font_targets(pdf: Any) -> List[_Target]:
    """Find headings as short blocks set in a clearly larger size than the body text."""
    blocks: List[Tuple[int, List[Tuple[str, float]]]] = []
    chars_by_size: Counter = Counter()
    for number, page in enumerate(pdf):
        for block in page.get_text("dict")["blocks"]:
            lines = _block_lines(block) if block["type"] == 0 else []
            if lines:
                blocks.append((number, lines))
                for text, size in lines:
                    chars_by_size[size] += len(text)
    if not chars_by_size:
        return []
    body = chars_by_size.most_common(1)[0][0]
    threshold = max(body * FALLBACK_MIN_SIZE_RATIO, body + FALLBACK_MIN_SIZE_GAIN)

    candidates = []
    pages_by_key: Dict[str, Set[int]] = {}
    for number, lines in blocks:
        sizes = {size for _, size in lines}
        if len(lines) > FALLBACK_MAX_BLOCK_LINES or len(sizes) != 1 or min(sizes) < threshold:
            continue
        text = " ".join(line for line, _ in lines)
        key = _title_key(text)
        if (
            not key
            or len(text) > FALLBACK_MAX_LINE_CHARS
            or not any(char.isalpha() for char in key)
        ):
            continue
        candidates.append((number, key, sizes.pop(), len(lines)))
        pages_by_key.setdefault(key, set()).add(number)

    total_lines = sum(len(lines) for _, lines in blocks)
    if (
        sum(line_count for *_, line_count in candidates)
        > total_lines * FALLBACK_MAX_CANDIDATE_SHARE
    ):
        return []
    candidates = [c for c in candidates if len(pages_by_key[c[1]]) < FALLBACK_MIN_REPEAT_PAGES]
    ranked = sorted({size for _, _, size, _ in candidates}, reverse=True)[:FALLBACK_LEVELS]
    levels = {size: rank + 1 for rank, size in enumerate(ranked)}
    return [
        _Target(levels[size], key, number + 1)
        for number, key, size, _ in candidates
        if size in levels
    ]


def _eligible_key(element: Any, default_id: Optional[str]) -> Optional[str]:
    properties = element.find(qn("w:pPr"))
    if properties is not None:
        if properties.find(qn("w:numPr")) is not None:
            return None
        style = properties.find(qn("w:pStyle"))
        if style is not None and style.get(qn("w:val")) != default_id:
            return None
    key = _title_key(paragraph_plain_text(element))
    return key if key and len(key) <= MAX_KEY_CHARS else None


def _paragraph_runs(
    pages: List[List[Any]], default_id: Optional[str]
) -> Tuple[List[List[List[_Item]]], List[Any]]:
    """Split each page into runs of adjacent, eligible, body level paragraphs."""
    paragraphs: List[Any] = []
    runs_by_page: List[List[List[_Item]]] = []
    for elements in pages:
        runs: List[List[_Item]] = []
        current: List[_Item] = []
        for element in elements:
            key = _eligible_key(element, default_id) if element.tag == qn("w:p") else None
            if key is None:
                if current:
                    runs.append(current)
                current = []
                continue
            paragraphs.append(element)
            current.append(_Item(len(paragraphs) - 1, key))
        if current:
            runs.append(current)
        runs_by_page.append(runs)
    return runs_by_page, paragraphs


def _build_index(runs: List[List[_Item]]) -> Dict[str, List[Tuple[int, ...]]]:
    index: Dict[str, List[Tuple[int, ...]]] = {}
    for run in runs:
        for start in range(len(run)):
            joined = ""
            for end in range(start, min(start + MAX_TITLE_PARAGRAPHS, len(run))):
                joined += run[end].key
                if len(joined) > MAX_KEY_CHARS:
                    break
                group = tuple(item.position for item in run[start : end + 1])
                for variant in _variants(joined):
                    index.setdefault(variant, []).append(group)
    return index


def _lookup(index: Dict[str, List[Tuple[int, ...]]], key: str) -> List[Tuple[int, ...]]:
    found: Dict[Tuple[int, ...], None] = {}
    for variant in _variants(key):
        found.update(dict.fromkeys(index.get(variant, ())))
    return list(found)


def _match(
    targets: List[_Target], runs_by_page: List[List[List[_Item]]], trusted: bool
) -> List[Tuple[_Target, Tuple[int, ...]]]:
    page_indexes = [_build_index(runs) for runs in runs_by_page] if trusted else []
    whole = _build_index([run for runs in runs_by_page for run in runs])
    key_counts = Counter(target.key for target in targets)
    used: Set[int] = set()
    matched = []
    for target in targets:
        if trusted and 1 <= target.page <= len(page_indexes):
            options = _lookup(page_indexes[target.page - 1], target.key)
        else:
            options = _lookup(whole, target.key)
            if key_counts[target.key] != 1 or len(options) != 1:
                continue
        group = next((option for option in options if used.isdisjoint(option)), None)
        if group is not None:
            used.update(group)
            matched.append((target, group))
    return matched


def _heading_style(doc: DocumentType, level: int) -> Any:
    """Return ``Heading <level>`` made visually neutral so styling never changes how text looks."""
    name = f"Heading {level}"
    styles = doc.styles
    normal = styles.default(WD_STYLE_TYPE.PARAGRAPH)
    existing = [
        style for style in styles if style.name == name and style.type == WD_STYLE_TYPE.PARAGRAPH
    ]
    style = (
        existing[0] if existing else styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH, builtin=True)
    )
    element = style.element
    for tag in ("w:pPr", "w:rPr", "w:link"):
        for stale in element.findall(qn(tag)):
            element.remove(stale)
    if normal is not None:
        style.base_style = normal
        style.next_paragraph_style = normal
    style.quick_style = True
    style.priority = 9
    outline = OxmlElement("w:outlineLvl")
    outline.set(qn("w:val"), str(level - 1))
    element.get_or_add_pPr().append(outline)
    return style
