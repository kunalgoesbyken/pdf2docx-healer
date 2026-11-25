"""Read the measurable facts of a DOCX: words, headings, list labels, links, tables, images."""

import shlex
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from lxml import etree

from .numbering import NumberingRenderer
from .words import tokenize

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
}
W, R, MC, A, V = (f"{{{NS[key]}}}" for key in ("w", "r", "mc", "a", "v"))
STYLE_HOPS = 20
EXTRA_PART_SUFFIXES = ("header+xml", "footer+xml", "footnotes+xml", "endnotes+xml")
OUTLINE_BODY_TEXT = 9


class InvalidDocx(Exception):
    """Raised when the output cannot count as a valid DOCX."""


@dataclass
class ParaFact:
    text: str = ""
    heading: Optional[int] = None
    numbered: bool = False
    ilvl: Optional[int] = None
    label: Optional[str] = None
    kind: Optional[str] = None
    in_table: bool = False


@dataclass
class DocxFacts:
    words: List[str] = field(default_factory=list)
    paragraphs: List[ParaFact] = field(default_factory=list)
    ext_links: List[str] = field(default_factory=list)
    anchors: List[str] = field(default_factory=list)
    bookmarks: List[str] = field(default_factory=list)
    bookmark_dup_names: int = 0
    resolved_anchors: int = 0
    dangling_anchors: int = 0
    broken_links: int = 0
    tables: int = 0
    images: int = 0
    list_paras: int = 0
    list_levels: List[int] = field(default_factory=list)
    unresolved_numbering: int = 0
    approx_formats: List[str] = field(default_factory=list)


@dataclass
class StyleInfo:
    name: str = ""
    based_on: Optional[str] = None
    outline: Optional[int] = None
    num_id: Optional[int] = None
    ilvl: Optional[int] = None


def parse_hyperlink_instruction(instruction: str) -> Optional[Tuple[str, str]]:
    """Return ("url", target) or ("anchor", name) for a HYPERLINK field instruction."""
    lexer = shlex.shlex(instruction, posix=False)
    lexer.whitespace_split = True
    tokens = [token.strip('"') for token in lexer]
    if not tokens or tokens[0].upper() != "HYPERLINK":
        return None
    index = 1
    url = None
    while index < len(tokens):
        token = tokens[index]
        if token.lower() == "\\l" and index + 1 < len(tokens):
            return "anchor", tokens[index + 1]
        if token.lower() in ("\\o", "\\t", "\\m") and index + 1 < len(tokens):
            index += 2
            continue
        if not token.startswith("\\") and url is None:
            url = token
        index += 1
    return ("url", url) if url else None


def int_or_none(value: Optional[str]) -> Optional[int]:
    """Parse an optional integer attribute."""
    try:
        return int(value) if value is not None else None
    except ValueError:
        return None


def child_val(element: Any, name: str) -> Optional[str]:
    """Return the w:val of a direct child, or None."""
    child = element.find(f"{W}{name}")
    return None if child is None else child.get(f"{W}val")


def read_styles(root: Any) -> Dict[str, StyleInfo]:
    """Index paragraph styles by id with the properties heading and list detection need."""
    styles: Dict[str, StyleInfo] = {}
    if root is None:
        return styles
    for element in root.findall(f"{W}style"):
        info = StyleInfo(
            name=child_val(element, "name") or "", based_on=child_val(element, "basedOn")
        )
        ppr = element.find(f"{W}pPr")
        if ppr is not None:
            info.outline = int_or_none(child_val(ppr, "outlineLvl"))
            num_pr = ppr.find(f"{W}numPr")
            if num_pr is not None:
                info.num_id = int_or_none(child_val(num_pr, "numId"))
                info.ilvl = int_or_none(child_val(num_pr, "ilvl"))
        styles[element.get(f"{W}styleId", "")] = info
    return styles


def style_chain(styles: Dict[str, StyleInfo], style_id: Optional[str]) -> List[StyleInfo]:
    """Return the style and its basedOn ancestors, bounded and cycle-safe."""
    chain: List[StyleInfo] = []
    seen: Set[str] = set()
    while style_id and style_id in styles and style_id not in seen and len(chain) < STYLE_HOPS:
        seen.add(style_id)
        chain.append(styles[style_id])
        style_id = styles[style_id].based_on
    return chain


def heading_level(chain: List[StyleInfo], direct_outline: Optional[int]) -> Optional[int]:
    """Return the heading level from outlineLvl (direct, then style) or a Heading N style."""
    outline = direct_outline
    if outline is None:
        for info in chain:
            if info.outline is not None:
                outline = info.outline
                break
    if outline is not None:
        return outline + 1 if outline < OUTLINE_BODY_TEXT else None
    for info in chain:
        name = info.name.lower()
        if name.startswith("heading ") and name[8:].isdigit():
            return int(name[8:])
    return None


def part_root(package_part: Any) -> Any:
    """Parse the XML of a package part, or None when it is absent or unreadable."""
    try:
        return etree.fromstring(package_part.blob)
    except (etree.XMLSyntaxError, ValueError):
        return None


def related_roots(document: Any) -> Dict[str, Any]:
    """Return the parsed numbering and styles parts reachable from the main document part."""
    roots: Dict[str, Any] = {}
    for rel in document.part.rels.values():
        if rel.is_external:
            continue
        for kind in ("numbering", "styles"):
            if rel.reltype.endswith("/" + kind):
                roots[kind] = part_root(rel.target_part)
    return roots


def extra_words(document: Any) -> List[str]:
    """Tokenize header, footer, footnote and endnote parts."""
    words: List[str] = []
    for part in document.part.package.iter_parts():
        if not str(part.content_type).endswith(EXTRA_PART_SUFFIXES):
            continue
        root = part_root(part)
        if root is None:
            continue
        for paragraph in root.iter(f"{W}p"):
            text = "".join(
                node.text or ""
                for node in paragraph.iter(f"{W}t")
                if not any(True for _ in node.iterancestors(f"{MC}Fallback"))
            )
            words.extend(tokenize(text))
    return words


class Walker:
    """Single document-order pass over the main document part."""

    def __init__(self, document: Any, roots: Dict[str, Any]) -> None:
        self.document = document
        self.facts = DocxFacts()
        self.styles = read_styles(roots.get("styles"))
        self.numbering = NumberingRenderer(roots.get("numbering"))
        self.fallback_depth = 0
        self.table_depth = 0
        self.stack: List[Tuple[ParaFact, List[str]]] = []
        self.fields: List[List[str]] = []
        self.bookmark_starts: List[Tuple[str, str, int]] = []
        self.bookmark_ids: Set[str] = set()
        self.bookmark_ends: Dict[str, List[int]] = {}
        self.sequence = 0
        self.body_words: List[str] = []

    def run(self) -> DocxFacts:
        """Walk the body, then resolve bookmarks and anchors."""
        for event, element in etree.iterwalk(self.document.element.body, events=("start", "end")):
            if element.tag == f"{MC}Fallback":
                self.fallback_depth += 1 if event == "start" else -1
            elif self.fallback_depth:
                continue
            elif event == "start":
                self.on_start(element)
            else:
                self.on_end(element)
        self.resolve()
        self.facts.words = self.body_words
        return self.facts

    def on_start(self, element: Any) -> None:
        """Dispatch the start of an element."""
        tag = element.tag
        if tag == f"{W}p":
            self.start_paragraph(element)
        elif tag == f"{W}t":
            self.append_text(element.text or "")
        elif tag in (f"{W}tab", f"{W}br", f"{W}cr"):
            self.append_text(" ")
        elif tag == f"{W}noBreakHyphen":
            self.append_text("-")
        elif tag == f"{W}tbl":
            self.table_depth += 1
            self.facts.tables += 1
        elif tag == f"{W}hyperlink":
            self.on_hyperlink(element)
        elif tag == f"{W}fldSimple":
            self.on_instruction(element.get(f"{W}instr") or "")
        elif tag == f"{W}fldChar":
            self.on_field_char(element.get(f"{W}fldCharType"))
        elif tag == f"{W}instrText" and self.fields:
            self.fields[-1].append(element.text or "")
        elif tag == f"{W}bookmarkStart":
            self.on_bookmark_start(element)
        elif tag == f"{W}bookmarkEnd":
            self.sequence += 1
            self.bookmark_ends.setdefault(element.get(f"{W}id") or "", []).append(self.sequence)
        elif tag in (f"{A}blip", f"{V}imagedata"):
            self.facts.images += 1

    def on_end(self, element: Any) -> None:
        """Close paragraphs and tables."""
        if element.tag == f"{W}p" and self.stack:
            record, buffer = self.stack.pop()
            record.text = "".join(buffer)
            if record.kind == "number" and record.label:
                self.body_words.extend(tokenize(record.label))
            self.body_words.extend(tokenize(record.text))
        elif element.tag == f"{W}tbl":
            self.table_depth -= 1

    def append_text(self, text: str) -> None:
        """Add text to the innermost open paragraph."""
        if self.stack:
            self.stack[-1][1].append(text)

    def start_paragraph(self, element: Any) -> None:
        """Open a paragraph record and number it in document order."""
        record = ParaFact(in_table=self.table_depth > 0)
        ppr = element.find(f"{W}pPr")
        style_id = child_val(ppr, "pStyle") if ppr is not None else None
        chain = style_chain(self.styles, style_id)
        direct_outline = int_or_none(child_val(ppr, "outlineLvl")) if ppr is not None else None
        record.heading = heading_level(chain, direct_outline)
        self.number(record, ppr, style_id, chain)
        self.facts.paragraphs.append(record)
        self.stack.append((record, []))

    def number(
        self, record: ParaFact, ppr: Any, style_id: Optional[str], chain: List[StyleInfo]
    ) -> None:
        """Resolve numId and ilvl (direct, then style chain) and label the paragraph."""
        num_id: Optional[int] = None
        ilvl: Optional[int] = None
        direct = ppr.find(f"{W}numPr") if ppr is not None else None
        if direct is not None:
            num_id = int_or_none(child_val(direct, "numId"))
            ilvl = int_or_none(child_val(direct, "ilvl"))
        from_style = False
        if num_id is None:
            for info in chain:
                if info.num_id is not None:
                    num_id, from_style = info.num_id, True
                    break
        if num_id is None or num_id == 0:
            return
        if ilvl is None and from_style and style_id:
            ilvl = self.numbering.ilvl_for_style(num_id, style_id)
        ilvl = 0 if ilvl is None else ilvl
        label = self.numbering.next_label(num_id, ilvl)
        if label is None:
            self.facts.unresolved_numbering += 1
            return
        record.numbered, record.ilvl, record.label, record.kind = True, ilvl, label.text, label.kind

    def on_hyperlink(self, element: Any) -> None:
        """Record a w:hyperlink as external (r:id) or internal (anchor)."""
        rel_id = element.get(f"{R}id")
        if rel_id:
            rel = self.document.part.rels.get(rel_id)
            if rel is None:
                self.facts.broken_links += 1
            elif rel.is_external:
                self.facts.ext_links.append(rel.target_ref)
            return
        anchor = element.get(f"{W}anchor")
        if anchor:
            self.facts.anchors.append(anchor)

    def on_field_char(self, kind: Optional[str]) -> None:
        """Track complex fields so HYPERLINK instructions split over runs are joined."""
        if kind == "begin":
            self.fields.append([])
        elif kind == "end" and self.fields:
            self.on_instruction("".join(self.fields.pop()))

    def on_instruction(self, instruction: str) -> None:
        """Record the target of a HYPERLINK field instruction."""
        parsed = parse_hyperlink_instruction(instruction)
        if parsed is None:
            return
        if parsed[0] == "url":
            self.facts.ext_links.append(parsed[1])
        else:
            self.facts.anchors.append(parsed[1])

    def on_bookmark_start(self, element: Any) -> None:
        """Record a bookmarkStart; a repeated id makes the document invalid."""
        book_id = element.get(f"{W}id") or ""
        if book_id in self.bookmark_ids:
            raise InvalidDocx(f"duplicate bookmark id {book_id}")
        self.bookmark_ids.add(book_id)
        self.sequence += 1
        self.bookmark_starts.append((book_id, element.get(f"{W}name") or "", self.sequence))

    def resolve(self) -> None:
        """Keep bookmarks that have a later bookmarkEnd with the same id, first name wins."""
        names: List[str] = []
        known: Set[str] = set()
        duplicates: Set[str] = set()
        for book_id, name, seq in self.bookmark_starts:
            if not any(end > seq for end in self.bookmark_ends.get(book_id, [])):
                continue
            if name in known:
                duplicates.add(name)
            else:
                known.add(name)
                names.append(name)
        self.facts.bookmarks = names
        self.facts.bookmark_dup_names = len(duplicates)
        self.facts.resolved_anchors = sum(1 for anchor in self.facts.anchors if anchor in known)
        self.facts.dangling_anchors = len(self.facts.anchors) - self.facts.resolved_anchors
        numbered = [p for p in self.facts.paragraphs if p.numbered]
        self.facts.list_paras = len(numbered)
        self.facts.list_levels = sorted({p.ilvl for p in numbered if p.ilvl is not None})
        self.facts.approx_formats = sorted(self.numbering.approx_formats)


def docx_facts(path: Any) -> DocxFacts:
    """Open a DOCX and measure it; raise InvalidDocx when it is not a valid document."""
    from docx import Document

    path = Path(path)
    try:
        if not zipfile.is_zipfile(str(path)):
            raise InvalidDocx("not a zip package")
        document = Document(str(path))
        facts = Walker(document, related_roots(document)).run()
        facts.words = facts.words + extra_words(document)
    except InvalidDocx:
        raise
    except Exception as exc:
        raise InvalidDocx(f"{type(exc).__name__}: {exc}") from exc
    return facts
