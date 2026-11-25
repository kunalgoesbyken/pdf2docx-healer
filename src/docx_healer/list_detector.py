"""Post-conversion list detection and Word style application."""

import copy
from dataclasses import dataclass, field
from typing import Any, Dict, FrozenSet, List, Optional

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from .list_markers import BULLET, Key, Marker, parse_marker
from .log import logger, progress

_INDENT_TOLERANCE_PT = 6.0
_HEADING_SIZE_RATIO = 1.25
_NUMBERING_LEVELS = 9
_STYLE_LEVELS = 5
_INDENT_TWIPS = 720
_HANGING_TWIPS = 360
_SYMBOL_BULLET = chr(0xF0B7)
_SHAPE_TEXT = {".": "%{n}.", ")": "%{n})", "()": "(%{n})"}
_KEEP_WHEN_CUTTING = frozenset({qn("w:pPr"), qn("w:rPr")})


@dataclass
class _Item:
    paragraph: int
    line: int
    marker: Marker
    indent: float
    leading_text: bool
    size: Optional[float]
    run: Optional["_Run"] = None
    level: int = 0


@dataclass
class _Run:
    """Consecutive items at one indent that continue the same sequence."""

    indent: float
    items: List[_Item] = field(default_factory=list)
    firsts: Dict[Key, int] = field(default_factory=dict)
    lasts: Dict[Key, int] = field(default_factory=dict)
    key: Optional[Key] = None
    num_id: Optional[int] = None

    def accepts(self, marker: Marker) -> bool:
        return any(
            key in self.lasts and value == self.lasts[key] + _step(key)
            for key, value in marker.readings.items()
        )

    def add(self, item: _Item) -> None:
        readings = item.marker.readings
        if not self.items:
            self.firsts = dict(readings)
            self.lasts = dict(readings)
        else:
            self.lasts = {
                key: value for key, value in readings.items()
                if key in self.lasts and value == self.lasts[key] + _step(key)
            }
            self.firsts = {key: self.firsts[key] for key in self.lasts}
        self.items.append(item)

    @property
    def strong(self) -> bool:
        return any(item.marker.strong for item in self.items)


def _step(key: Key) -> int:
    return 0 if key[0] == BULLET else 1


def _heading_style_ids(document: Any) -> FrozenSet[str]:
    ids = set()
    for style in document.styles:
        name = (style.name or "").lower()
        pPr = style.element.pPr
        outlined = pPr is not None and pPr.find(qn("w:outlineLvl")) is not None
        titled = name.startswith("heading") or name == "title"
        if style.type == WD_STYLE_TYPE.PARAGRAPH and (outlined or titled):
            ids.add(style.style_id)
    return frozenset(ids)


def _is_heading(paragraph: Paragraph, heading_ids: FrozenSet[str]) -> bool:
    pPr: Any = paragraph._p.pPr
    if pPr is None:
        return False
    return pPr.find(qn("w:outlineLvl")) is not None or pPr.style in heading_ids


def _has_numbering(paragraph: Paragraph) -> bool:
    pPr = paragraph._p.pPr
    return pPr is not None and pPr.numPr is not None


def _text_atoms(p_elem: Any) -> List[Any]:
    return [
        el for el in p_elem.iter(qn("w:t"), qn("w:tab"), qn("w:br"))
        if el.getparent().tag == qn("w:r")
        and (el.tag != qn("w:br") or el.get(qn("w:type")) in (None, "textWrapping"))
    ]


def _atom_text(atom: Any) -> str:
    if atom.tag == qn("w:t"):
        return atom.text or ""
    return "\n" if atom.tag == qn("w:br") else "\t"


def _line_texts(atoms: List[Any]) -> List[str]:
    return "".join(_atom_text(atom) for atom in atoms).split("\n")


def _font_size(p_elem: Any) -> Optional[float]:
    """Return the font size in points of the last run that holds text, if it is set."""
    for run in reversed(list(p_elem.iter(qn("w:r")))):
        if run.find(qn("w:t")) is None:
            continue
        size = run.find(f"{qn('w:rPr')}/{qn('w:sz')}")
        return int(size.get(qn("w:val"))) / 2 if size is not None else None
    return None


def _left_indent_pt(paragraph: Paragraph) -> float:
    indent = paragraph.paragraph_format.left_indent
    return indent.pt if indent is not None else 0.0


def _scan(paragraphs: List[Paragraph], heading_ids: FrozenSet[str]) -> List[_Item]:
    items: List[_Item] = []
    for index, paragraph in enumerate(paragraphs):
        if _is_heading(paragraph, heading_ids) or _has_numbering(paragraph):
            continue
        lines = _line_texts(_text_atoms(paragraph._p))
        markers = [(number, parse_marker(line)) for number, line in enumerate(lines)]
        found = [(number, marker) for number, marker in markers if marker is not None]
        if not found:
            continue
        indent = _left_indent_pt(paragraph)
        size = _font_size(paragraph._p)
        for position, (number, marker) in enumerate(found):
            leading = position == 0 and number > 0
            items.append(_Item(index, number, marker, indent, leading, size))
    return items


def _adjacent(previous: int, item: _Item) -> bool:
    return item.paragraph in (previous, previous + 1) and not item.leading_text


def _is_larger(parent: _Item, child: _Item) -> bool:
    if parent.size is None or child.size is None:
        return False
    return parent.size > child.size * _HEADING_SIZE_RATIO


def _resolve_runs(items: List[_Item], aggressive: bool) -> None:
    stack: List[_Run] = []

    def finish(run: _Run) -> None:
        if not (aggressive or run.strong or len(run.items) >= 2):
            return
        run.key = next(iter(run.lasts))
        for item in run.items:
            item.run = run

    previous = -2
    for item in items:
        if not _adjacent(previous, item):
            while stack:
                finish(stack.pop())
        previous = item.paragraph
        while stack and stack[-1].indent > item.indent + _INDENT_TOLERANCE_PT:
            finish(stack.pop())
        if stack and abs(stack[-1].indent - item.indent) <= _INDENT_TOLERANCE_PT:
            if stack[-1].accepts(item.marker):
                stack[-1].add(item)
                continue
            finish(stack.pop())
        if stack and _is_larger(stack[-1].items[-1], item):
            finish(stack.pop())
        run = _Run(item.indent)
        run.add(item)
        stack.append(run)
    while stack:
        finish(stack.pop())


def _assign_levels(items: List[_Item]) -> None:
    stretch: List[_Item] = []

    def close() -> None:
        bases: List[float] = []
        for indent in sorted({item.indent for item in stretch}):
            if not bases or indent - bases[-1] > _INDENT_TOLERANCE_PT:
                bases.append(indent)
        for item in stretch:
            item.level = max(i for i, base in enumerate(bases) if base <= item.indent)
        stretch.clear()

    previous = -2
    for item in items:
        if item.run is None or not _adjacent(previous, item):
            close()
        if item.run is not None:
            stretch.append(item)
        previous = item.paragraph
    close()


def _drop_from(node: Any, root: Any, forward: bool) -> None:
    first = True
    while node is not root:
        parent = node.getparent()
        for sibling in list(node.itersiblings(preceding=not forward)):
            if forward or sibling.tag not in _KEEP_WHEN_CUTTING:
                parent.remove(sibling)
        if first:
            parent.remove(node)
            first = False
        node = parent


def _strip_trailing_space(p_elem: Any) -> None:
    for atom in reversed(_text_atoms(p_elem)):
        if atom.tag != qn("w:t"):
            break
        atom.text = (atom.text or "").rstrip()
        if atom.text:
            break


def _split_at(p_elem: Any, line_break: Any) -> Any:
    path = []
    node = line_break
    while node is not p_elem:
        parent = node.getparent()
        path.append(parent.index(node))
        node = parent
    tail = copy.deepcopy(p_elem)
    twin = tail
    for position in reversed(path):
        twin = twin[position]
    _drop_from(line_break, p_elem, forward=True)
    _drop_from(twin, tail, forward=False)
    _strip_trailing_space(p_elem)
    return tail


def _split_lines(paragraph: Paragraph, starts: List[int]) -> List[Any]:
    p_elem = paragraph._p
    breaks = [atom for atom in _text_atoms(p_elem) if atom.tag == qn("w:br")]
    parts = [p_elem]
    for start in reversed(starts[1:]):
        tail = _split_at(p_elem, breaks[start - 1])
        p_elem.addnext(tail)
        parts.insert(1, tail)
    return parts


def _cut_prefix(p_elem: Any, count: int) -> None:
    for atom in _text_atoms(p_elem):
        if count <= 0:
            break
        if atom.tag == qn("w:t"):
            text = atom.text or ""
            atom.text = text[count:]
            atom.set(qn("xml:space"), "preserve")
            count -= len(text)
        else:
            atom.getparent().remove(atom)
            count -= 1


class _Numbering:
    """Creates the numbering definitions of one document on demand."""

    def __init__(self, document: Any) -> None:
        self._document = document
        self._bullet_id: Optional[int] = None

    @property
    def _element(self) -> Any:
        return self._document.part.numbering_part.element

    def bullet(self) -> int:
        if self._bullet_id is None:
            self._bullet_id = self._add(BULLET, "", 1)
        return self._bullet_id

    def numbered(self, key: Key, start: int) -> int:
        return self._add(key[0], key[1], start)

    def _add(self, fmt: str, shape: str, start: int) -> int:
        element = self._element
        abs_num_id = _next_id(element, "w:abstractNum", "w:abstractNumId")
        num_id = _next_id(element, "w:num", "w:numId")
        _insert_abstract_num(element, _build_abstract_num(abs_num_id, fmt, shape, start))
        num = OxmlElement("w:num")
        num.set(qn("w:numId"), str(num_id))
        _mk_child(num, "w:abstractNumId", abs_num_id)
        element.append(num)
        return num_id


def _insert_abstract_num(numbering_elm: Any, abs_num: Any) -> None:
    first_num = numbering_elm.find(qn("w:num"))
    if first_num is None:
        numbering_elm.append(abs_num)
    else:
        first_num.addprevious(abs_num)


def _build_abstract_num(abs_num_id: int, fmt: str, shape: str, start: int) -> Any:
    abs_num = OxmlElement("w:abstractNum")
    abs_num.set(qn("w:abstractNumId"), str(abs_num_id))
    nsid = f"{abs_num_id + 1:08X}"
    _mk_child(abs_num, "w:nsid", nsid)
    _mk_child(abs_num, "w:multiLevelType", "hybridMultilevel")
    _mk_child(abs_num, "w:tmpl", nsid)
    for level in range(_NUMBERING_LEVELS):
        abs_num.append(_build_level(level, fmt, shape, start))
    return abs_num


def _build_level(level: int, fmt: str, shape: str, start: int) -> Any:
    lvl = OxmlElement("w:lvl")
    lvl.set(qn("w:ilvl"), str(level))
    lvl.set(qn("w:tplc"), str(level))
    _mk_child(lvl, "w:start", start)
    _mk_child(lvl, "w:numFmt", fmt)
    text = _SYMBOL_BULLET if fmt == BULLET else _SHAPE_TEXT[shape].format(n=level + 1)
    _mk_child(lvl, "w:lvlText", text)
    _mk_child(lvl, "w:lvlJc", "left")

    pPr = OxmlElement("w:pPr")
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), str(_INDENT_TWIPS * (level + 1)))
    ind.set(qn("w:hanging"), str(_HANGING_TWIPS))
    pPr.append(ind)
    lvl.append(pPr)

    if fmt == BULLET:
        rPr = OxmlElement("w:rPr")
        rFonts = OxmlElement("w:rFonts")
        for attribute in ("w:ascii", "w:hAnsi"):
            rFonts.set(qn(attribute), "Symbol")
        rFonts.set(qn("w:hint"), "default")
        rPr.append(rFonts)
        lvl.append(rPr)
    return lvl


def _mk_child(parent: Any, tag: str, val: Any) -> Any:
    el = OxmlElement(tag)
    el.set(qn("w:val"), str(val))
    parent.append(el)
    return el


def _next_id(numbering_elm: Any, element_xpath: str, attr_name: str) -> int:
    ids = numbering_elm.xpath(f"./{element_xpath}/@{attr_name}")
    if not ids:
        return 0 if "abstractNum" in element_xpath else 1
    return max(int(i) for i in ids) + 1


class _Styles:
    """Looks up the built-in list style ids of one document once each."""

    def __init__(self, document: Any) -> None:
        self._styles = document.styles
        self._ids: Dict[str, Optional[str]] = {}

    def list_style_id(self, base: str, level: int) -> Optional[str]:
        name = base if level == 0 else f"{base} {level + 1}"
        if name not in self._ids:
            try:
                self._ids[name] = self._styles[name].style_id
            except KeyError:
                self._ids[name] = None
        return self._ids[name] if level < _STYLE_LEVELS else None


def _apply(
    paragraph: Paragraph, run: _Run, key: Key, level: int, numbering: _Numbering, styles: _Styles
) -> None:
    is_bullet = key[0] == BULLET
    pPr: Any = paragraph._p.get_or_add_pPr()
    style_id = styles.list_style_id("List Bullet" if is_bullet else "List Number", level)
    if style_id is not None:
        pPr.style = style_id
        if is_bullet:
            return
    if is_bullet:
        num_id = numbering.bullet()
    else:
        if run.num_id is None:
            run.num_id = numbering.numbered(key, run.firsts[key])
        num_id = run.num_id
    pPr._remove_numPr()
    numPr = pPr.get_or_add_numPr()
    numPr.get_or_add_ilvl().val = min(level, _NUMBERING_LEVELS - 1)
    numPr.get_or_add_numId().val = num_id


def _apply_paragraph(
    paragraph: Paragraph, items: List[_Item], numbering: _Numbering, styles: _Styles
) -> int:
    starts = sorted({0, *(item.line for item in items)})
    parts = _split_lines(paragraph, starts) if len(starts) > 1 else [paragraph._p]
    by_line = {item.line: item for item in items}
    fixed = 0
    for start, p_elem in zip(starts, parts):
        item = by_line.get(start)
        if item is None or item.run is None or item.run.key is None:
            continue
        _cut_prefix(p_elem, item.marker.end)
        target = Paragraph(p_elem, paragraph._parent)
        _apply(target, item.run, item.run.key, item.level, numbering, styles)
        fixed += 1
    return fixed


def fix_lists_in_docx(docx_path: Any, aggressive: bool = False, verbose: bool = False) -> int:
    """Turn bullet and numbered paragraphs into real Word lists, returning how many were fixed.

    A numbered marker only becomes a list item inside a run of at least two items that count
    up by one, so the numbers a reader sees stay the ones in the PDF; aggressive mode also
    accepts a single numbered item. Each list gets its own numbering definition, format and
    start value. Heading paragraphs are never touched.
    """
    doc = Document(str(docx_path))
    paragraphs = list(doc.paragraphs)
    items = _scan(paragraphs, _heading_style_ids(doc))
    _resolve_runs(items, aggressive)
    _assign_levels(items)

    grouped: Dict[int, List[_Item]] = {}
    for item in items:
        grouped.setdefault(item.paragraph, []).append(item)

    numbering = _Numbering(doc)
    styles = _Styles(doc)
    fixed = 0
    for index in sorted(grouped, reverse=True):
        group = grouped[index]
        if not any(item.run is not None for item in group):
            continue
        try:
            fixed += _apply_paragraph(paragraphs[index], group, numbering, styles)
        except Exception as exc:
            logger.warning("Skipping a list paragraph that could not be styled: %s", exc)
            logger.debug("List paragraph failure details", exc_info=True)

    if fixed > 0:
        doc.save(docx_path)
        progress(verbose, f"Fixed {fixed} list items")
    return fixed
