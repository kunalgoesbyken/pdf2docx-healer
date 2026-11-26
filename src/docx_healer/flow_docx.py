"""Writing the assembled blocks into a python-docx document."""

import io
from collections import Counter
from dataclasses import replace
from typing import Any, Dict, List, Optional, Tuple
from xml.sax.saxutils import escape

import pymupdf
from docx.document import Document as DocumentType
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.image.image import Image as DocxImage
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Emu, Pt, RGBColor
from docx.text.paragraph import Paragraph
from lxml import etree

from .constants import FONT_CLASS_FALLBACKS
from .flow_lists import LevelDef, ListInstance, level_indents
from .flow_model import (
    ImageItem,
    LinkKey,
    Para,
    Running,
    RunningImage,
    RunningItem,
    RunningLine,
    Span,
    TableItem,
)
from .flow_page import Look
from .flow_text import is_rtl_char
from .font_mapper import font_family, normalize_font_name
from .log import logger

_COMPLEX_START = 0x0590
_LEVELS = 9
_FALLBACK_STEP = 360
_BODY_SIZE = 0.6
_MIN_GAP = 0.5
_MAX_FONTS = 200
_DEFAULT_FAMILY = "Times New Roman"
_WHITE = 0xFFFFFF
_MIN_CONTRAST = 1.5
_QUOTE = {'"': "&quot;"}
_BIDI_SUCCESSORS = (
    "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing",
    "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
    "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
    "w:pPrChange",
)
_SHADE_SUCCESSORS = (
    "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
    "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi",
) + _BIDI_SUCCESSORS
_DIRECTION_SUCCESSORS = _BIDI_SUCCESSORS[_BIDI_SUCCESSORS.index("w:textDirection") + 1 :]
_FRAME_SUCCESSORS = (
    "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd",
) + _SHADE_SUCCESSORS
_JUSTIFY = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


class Numbering:
    """Creates abstractNum and num elements for list instances."""

    def __init__(self, doc: DocumentType) -> None:
        self.element = doc.part.numbering_part.element
        self.abstract_id = self._max("w:abstractNum", "w:abstractNumId")
        self.num_id = self._max("w:num", "w:numId")

    def _max(self, tag: str, attribute: str) -> int:
        values = [int(e.get(qn(attribute)) or 0) for e in self.element.findall(qn(tag))]
        return max(values, default=0)

    def add(self, instance: ListInstance, margin: float) -> int:
        self.abstract_id += 1
        self.num_id += 1
        base = instance.levels[min(instance.levels)]
        levels = "".join(
            _level_xml(i, instance.levels.get(i), base, margin) for i in range(_LEVELS)
        )
        abstract = parse_xml(
            f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{self.abstract_id}">'
            f'<w:multiLevelType w:val="hybridMultilevel"/>{levels}</w:abstractNum>'
        )
        first_num = self.element.find(qn("w:num"))
        if first_num is not None:
            first_num.addprevious(abstract)
        else:
            self.element.append(abstract)
        num = parse_xml(
            f'<w:num {nsdecls("w")} w:numId="{self.num_id}">'
            f'<w:abstractNumId w:val="{self.abstract_id}"/></w:num>'
        )
        cleanup = self.element.find(qn("w:numIdMacAtCleanup"))
        if cleanup is not None:
            cleanup.addprevious(num)
        else:
            self.element.append(num)
        return self.num_id


def _level_xml(index: int, level: Optional[LevelDef], base: LevelDef, margin: float) -> str:
    if level is None:
        bullet = base.bullet
        fmt = "bullet" if bullet else "decimal"
        text = "•" if bullet else f"%{index + 1}."
        left, hanging = _FALLBACK_STEP * (index + 2), _FALLBACK_STEP
        start = 1
    else:
        fmt = level.fmt
        text = level.glyph if level.bullet else _level_text(level, index)
        left, hanging = level_indents(level, margin)
        start = level.start
    return (
        f'<w:lvl w:ilvl="{index}"><w:start w:val="{start}"/><w:numFmt w:val="{fmt}"/>'
        f'<w:lvlText w:val="{escape(text)}"/><w:lvlJc w:val="left"/>'
        f'<w:pPr><w:ind w:left="{left}" w:hanging="{hanging}"/></w:pPr></w:lvl>'
    )


def _level_text(level: LevelDef, index: int) -> str:
    ancestors = "".join(f"%{k + 1}." for k in range(index)) if level.dotted else ""
    return ancestors + level.delim.replace("%", f"%{index + 1}")


class FontBook:
    """The document's body font and colour, and the family each PDF font name stands for."""

    def __init__(self, paras: List[Para], body: float) -> None:
        self._names: Dict[str, str] = {}
        chars: Counter = Counter()
        colors: Counter = Counter()
        classes: Dict[str, Counter] = {}
        for para in paras:
            for span in para.spans:
                weight = len(span.text.strip())
                if not weight:
                    continue
                if abs(span.size - body) < _BODY_SIZE:
                    colors[span.color] += weight
                family = self._resolve(span.font)
                if not family:
                    continue
                if abs(span.size - body) < _BODY_SIZE and not span.mono:
                    chars[family] += weight
                kind = "modern" if span.mono else "roman" if span.serif else "swiss"
                classes.setdefault(family, Counter())[kind] += weight
        self.body = chars.most_common(1)[0][0] if chars else _DEFAULT_FAMILY
        self.color = colors.most_common(1)[0][0] if colors else 0
        self.classes = {name: votes.most_common(1)[0][0] for name, votes in classes.items()}

    def _resolve(self, name: str) -> str:
        found = self._names.get(name)
        if found is None:
            found = self._names[name] = font_family(name)
        return found

    def family(self, name: str, mono: bool = False) -> str:
        found = self._resolve(name)
        if found:
            return found
        return FONT_CLASS_FALLBACKS["modern"] if mono else self.body

    def register(self, doc: DocumentType) -> None:
        """List each family in the font table with its class, so a missing font gets a like one."""
        try:
            part = doc.part.part_related_by(RT.FONT_TABLE)
            table = etree.fromstring(part.blob)
        except Exception as exc:
            logger.debug("No font table to extend: %s", exc)
            return
        known = {font.get(qn("w:name")) for font in table.findall(qn("w:font"))}
        for name, kind in sorted(self.classes.items()):
            if name in known or len(known) >= _MAX_FONTS:
                continue
            known.add(name)
            pitch = "fixed" if kind == "modern" else "variable"
            table.append(parse_xml(
                f'<w:font {nsdecls("w")} w:name="{escape(name, _QUOTE)}">'
                f'<w:altName w:val="{FONT_CLASS_FALLBACKS[kind]}"/>'
                f'<w:family w:val="{kind}"/><w:pitch w:val="{pitch}"/></w:font>'
            ))
        part._blob = etree.tostring(
            table, xml_declaration=True, encoding="UTF-8", standalone=True
        )


class Emitter:
    """Writes paragraphs, tables and images into one document."""

    def __init__(
        self, doc: DocumentType, body: float, text_width: float, anchors: Dict[LinkKey, str],
        fonts: FontBook,
    ) -> None:
        self.doc = doc
        self.body = body
        self.text_width = text_width
        self.anchors = anchors
        self.fonts = fonts
        self.numbering = Numbering(doc)
        self.num_ids: Dict[int, int] = {}
        self.bookmark_id = 0
        plain_styles(doc, fonts.body, body, fonts.color)

    def register_instances(self, instances: List[ListInstance], margins: List[float]) -> None:
        """Add each list's numbering, its indents measured from the column it sits in."""
        for index, (instance, margin) in enumerate(zip(instances, margins)):
            self.num_ids[index] = self.numbering.add(instance, margin)

    def paragraph(self, paragraph: Paragraph, para: Para, look: Optional[Look] = None) -> None:
        if para.kind == "heading":
            paragraph.style = self.doc.styles[f"Heading {para.level}"]
        elif para.kind == "title":
            paragraph.style = self.doc.styles["Title"]
        p = paragraph._p
        listed = para.kind == "list" and para.num is not None
        if para.kind == "list" and para.num is not None:
            ppr: Any = p.get_or_add_pPr()
            numpr = ppr.get_or_add_numPr()
            numpr.get_or_add_ilvl().val = para.level
            numpr.get_or_add_numId().val = self.num_ids[para.num[0]]
        if para.rtl:
            _set_bidi(p.get_or_add_pPr())
        if _shaded(para.shade):
            _shade(p.get_or_add_pPr(), para.shade or 0)
        if look is not None:
            if look.frame is not None:
                _frame(p.get_or_add_pPr(), look.frame, look.turn)
            if not look.turn:
                _apply_look(paragraph, look, indents=not listed and look.frame is None)
        self._runs(paragraph, para)
        if look is not None and look.column_break:
            _column_break(paragraph)
        if para.bookmark:
            self._bookmark(p, para.bookmark)

    def _bookmark(self, p: Any, name: str) -> None:
        self.bookmark_id += 1
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), str(self.bookmark_id))
        start.set(qn("w:name"), name)
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), str(self.bookmark_id))
        ppr = p.find(qn("w:pPr"))
        p.insert(0 if ppr is None else 1, start)
        p.append(end)

    def _target(self, link: Optional[LinkKey]) -> Optional[Tuple[str, str]]:
        if link is None:
            return None
        if link[0] == "uri":
            return ("uri", link[1])
        name = self.anchors.get(link)
        return ("anchor", name) if name else None

    def _runs(self, paragraph: Paragraph, para: Para) -> None:
        current: Optional[Tuple[str, str]] = None
        container: Any = None
        for span in para.spans:
            target = self._target(span.link)
            if target != current:
                current = target
                container = self._open_link(paragraph, target) if target else None
            run = paragraph.add_run(span.text)
            self._format(run, span, para)
            if container is not None:
                container.append(run._r)

    def _open_link(self, paragraph: Paragraph, target: Tuple[str, str]) -> Any:
        link = OxmlElement("w:hyperlink")
        if target[0] == "uri":
            rid = paragraph.part.relate_to(target[1], RT.HYPERLINK, is_external=True)
            link.set(qn("r:id"), rid)
        else:
            link.set(qn("w:anchor"), target[1])
        link.set(qn("w:history"), "1")
        paragraph._p.append(link)
        return link

    def _format(self, run: Any, span: Span, para: Para) -> None:
        if span.bold:
            run.bold = run.font.cs_bold = True
        if span.italic:
            run.italic = run.font.cs_italic = True
        size = span.size
        if span.sup:
            run.font.superscript = True
            size = para.size
        elif span.sub:
            run.font.subscript = True
            size = para.size
        if abs(size - self.body) >= _BODY_SIZE / 2:
            _set_size(run._r.get_or_add_rPr(), size)
        colour = span.color if _shaded(para.shade) or _readable(span.color) else self.fonts.color
        if colour != self.fonts.color:
            run.font.color.rgb = _rgb(colour)
        if any(ord(ch) >= _COMPLEX_START for ch in span.text):
            _complex_fonts(run, span)
        else:
            family = self.fonts.family(span.font, span.mono)
            if family != self.fonts.body:
                run.font.name = family
        if any(is_rtl_char(ch) for ch in span.text):
            run.font.rtl = True

    def table(self, table: TableItem, look: Optional[Look] = None) -> None:
        if look is not None and (look.before > _MIN_GAP or look.page_break or look.column_break):
            spacer(self.doc.add_paragraph(), look)
        columns = len(table.col_edges) - 1
        grid = self.doc.add_table(rows=table.rows, cols=columns)
        grid.style = self.doc.styles["Table Grid"]
        widths = [b - a for a, b in zip(table.col_edges, table.col_edges[1:])]
        room = look.width if look is not None and look.width > 0 else self.text_width
        scale = min(1.0, room / max(sum(widths), 1.0))
        for index, width in enumerate(widths):
            grid.columns[index].width = Pt(width * scale)
            for row in grid.rows:
                row.cells[index].width = Pt(width * scale)
        if look is not None and look.left > _MIN_GAP:
            indent = OxmlElement("w:tblInd")
            indent.set(qn("w:w"), str(twips(look.left)))
            indent.set(qn("w:type"), "dxa")
            grid._tbl.tblPr.append(indent)
        self._merge(grid, table)
        for cell in table.cells:
            self._fill_cell(grid.cell(cell.r0, cell.c0), cell.paras)
        if table.header_rows:
            for row in grid.rows[: table.header_rows]:
                row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))

    def _merge(self, grid: Any, table: TableItem) -> None:
        for cell in table.cells:
            if cell.r1 - cell.r0 > 1 or cell.c1 - cell.c0 > 1:
                try:
                    grid.cell(cell.r0, cell.c0).merge(grid.cell(cell.r1 - 1, cell.c1 - 1))
                except Exception as exc:
                    logger.debug("Table cell merge skipped: %s", exc)

    def _fill_cell(self, target: Any, paras: List[Para]) -> None:
        for index, para in enumerate(paras):
            paragraph = target.paragraphs[0] if index == 0 else target.add_paragraph()
            self.paragraph(paragraph, para)

    def image(self, image: ImageItem, look: Look) -> None:
        paragraph = self.doc.add_paragraph()
        room = look.width if look.width > 0 else self.text_width
        if not _add_picture(paragraph, image, room):
            body = paragraph._p.getparent()
            if body is not None:
                body.remove(paragraph._p)
            return
        if look.frame is not None:
            _frame(paragraph._p.get_or_add_pPr(), look.frame, look.turn)
        _apply_look(paragraph, replace(look, pitch=0.0), indents=look.frame is None)
        if look.column_break:
            _column_break(paragraph)


def _half_points(size: float) -> float:
    return max(1.0, round(size * 2) / 2)


def _shaded(shade: Optional[int]) -> bool:
    return shade is not None and shade != _WHITE


def _readable(color: int) -> bool:
    """Whether text in this colour can be read on a white page."""
    channels = [((color >> shift) & 0xFF) / 255 for shift in (16, 8, 0)]
    linear = [c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4 for c in channels]
    luminance = 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2]
    return 1.05 / (luminance + 0.05) >= _MIN_CONTRAST


def _shade(ppr: Any, fill: int) -> None:
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), f"{fill:06X}")
    ppr.insert_element_before(shading, *_SHADE_SUCCESSORS)


def _set_size(rpr: Any, size: float) -> None:
    """Set the size for every script, as Word sizes right-to-left and complex text by szCs."""
    points = _half_points(size)
    sz = rpr.get_or_add_sz()
    sz.val = Pt(points)
    complex_size = rpr.find(qn("w:szCs"))
    if complex_size is None:
        complex_size = OxmlElement("w:szCs")
        sz.addnext(complex_size)
    complex_size.set(qn("w:val"), str(int(points * 2)))


def twips(points: float) -> int:
    return int(round(points * 20))


def _rgb(color: int) -> RGBColor:
    return RGBColor((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF)


def _apply_look(paragraph: Paragraph, look: Look, indents: bool) -> None:
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(look.before)
    if look.pitch > 0:
        fmt.line_spacing = Pt(look.pitch)
        fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    fmt.alignment = _JUSTIFY[look.align]
    if indents:
        if look.left > _MIN_GAP:
            fmt.left_indent = Pt(look.left)
        if look.right > _MIN_GAP:
            fmt.right_indent = Pt(look.right)
        if abs(look.first) > _MIN_GAP:
            fmt.first_line_indent = Pt(look.first)
    if look.page_break:
        fmt.page_break_before = True


def _frame(ppr: Any, rect: Tuple[float, float, float, float], turn: str) -> None:
    """Pin the paragraph to its place on the page, its text running up or down the page."""
    frame = OxmlElement("w:framePr")
    for name, value in (
        ("w:w", twips(rect[2] - rect[0])), ("w:h", twips(rect[3] - rect[1])),
        ("w:x", twips(rect[0])), ("w:y", twips(rect[1])),
    ):
        frame.set(qn(name), str(value))
    frame.set(qn("w:hRule"), "exact" if turn else "atLeast")
    frame.set(qn("w:wrap"), "none")
    frame.set(qn("w:hAnchor"), "page")
    frame.set(qn("w:vAnchor"), "page")
    ppr.insert_element_before(frame, *_FRAME_SUCCESSORS)
    if not turn:
        return
    direction = OxmlElement("w:textDirection")
    direction.set(qn("w:val"), turn)
    ppr.insert_element_before(direction, *_DIRECTION_SUCCESSORS)


def spacer(paragraph: Paragraph, look: Look) -> None:
    """Make an empty paragraph exactly as tall as the gap before a block that is not text."""
    fmt = paragraph.paragraph_format
    fmt.line_spacing = Pt(max(look.before, _MIN_GAP))
    fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    if look.page_break:
        fmt.page_break_before = True
    if look.column_break:
        _column_break(paragraph)


def _column_break(paragraph: Paragraph) -> None:
    run = OxmlElement("w:r")
    brk = OxmlElement("w:br")
    brk.set(qn("w:type"), "column")
    run.append(brk)
    ppr = paragraph._p.find(qn("w:pPr"))
    paragraph._p.insert(0 if ppr is None else 1, run)


def plain_styles(doc: DocumentType, family: str, size: float, color: int) -> None:
    """Make the default, Normal, Title and heading styles add nothing the PDF does not show."""
    defaults = doc.styles.element.find(qn("w:docDefaults"))
    if defaults is not None:
        for fonts in defaults.iter(qn("w:rFonts")):
            for attribute in [a for a in fonts.attrib if a.lower().endswith("theme")]:
                del fonts.attrib[attribute]
            fonts.set(qn("w:ascii"), family)
            fonts.set(qn("w:hAnsi"), family)
        for spacing in list(defaults.iter(qn("w:spacing"))):
            spacing.getparent().remove(spacing)
    normal = doc.styles["Normal"]
    normal.font.name = family
    _set_size(normal.element.get_or_add_rPr(), size)
    if color:
        normal.font.color.rgb = _rgb(color)
    normal.paragraph_format.widow_control = False
    for style in doc.styles:
        if style.type == WD_STYLE_TYPE.PARAGRAPH and (
            style.name == "Title" or style.name.startswith("Heading ")
        ):
            _strip_style(style.element)


def _strip_style(element: Any) -> None:
    for tag in ("w:rPr", "w:pPr"):
        found = element.find(qn(tag))
        if found is None:
            continue
        for child in list(found):
            if child.tag != qn("w:outlineLvl"):
                found.remove(child)


def _add_picture(paragraph: Paragraph, image: ImageItem, text_width: float) -> bool:
    """Put the image into the paragraph at most text_width wide; False if it cannot be embedded."""
    data = _usable_image(image)
    if data is None:
        return False
    try:
        paragraph.add_run().add_picture(io.BytesIO(data), width=Pt(min(image.width, text_width)))
    except Exception as exc:
        logger.warning("Could not embed an image: %s", exc)
        return False
    return True


def _usable_image(image: ImageItem) -> Optional[bytes]:
    """Return image bytes python-docx can embed, re-encoding as PNG when it cannot read them."""
    try:
        DocxImage.from_blob(image.data)
        return image.data
    except Exception as exc:
        logger.debug("Re-encoding an image python-docx cannot read (%s): %s", image.ext, exc)
    try:
        pixmap = pymupdf.Pixmap(image.data)
        if pixmap.colorspace is None or pixmap.colorspace.n != 3 or pixmap.alpha:
            pixmap = pymupdf.Pixmap(pymupdf.csRGB, pixmap)
        return bytes(pixmap.tobytes("png"))
    except Exception as exc:
        logger.warning("Skipped an image in an unsupported format (%s): %s", image.ext, exc)
        return None


def _complex_fonts(run: Any, span: Span) -> None:
    name = normalize_font_name(span.font, span.text)
    run.font.name = name
    fonts = run._r.get_or_add_rPr().rFonts
    for attribute in ("w:eastAsia", "w:cs"):
        fonts.set(qn(attribute), name)


def _set_bidi(ppr: Any) -> None:
    ppr.insert_element_before(OxmlElement("w:bidi"), *_BIDI_SUCCESSORS)


_ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
}


def write_running(section: Any, running: Running, body: float, fonts: FontBook) -> None:
    """Write the repeating lines and pictures into the section's Word header and footer parts."""
    width = Emu(section.page_width - section.left_margin - section.right_margin).pt
    for part, items in ((section.header, running.header), (section.footer, running.footer)):
        _fill_part(part, items, body, width, fonts)
        first = [item for item in items if item.first_page]
        if len(first) != len(items):
            section.different_first_page_header_footer = True
    first_header = section.first_page_header
    first_footer = section.first_page_footer
    if section.different_first_page_header_footer:
        _fill_part(first_header, [x for x in running.header if x.first_page], body, width, fonts)
        _fill_part(first_footer, [x for x in running.footer if x.first_page], body, width, fonts)


def _fill_part(
    part: Any, items: List[RunningItem], body: float, width: float, fonts: FontBook
) -> None:
    if not items:
        return
    part.is_linked_to_previous = False
    for index, item in enumerate(items):
        paragraph = part.paragraphs[0] if index == 0 else part.add_paragraph()
        paragraph.alignment = _ALIGNMENTS[item.align]
        if isinstance(item, RunningImage):
            _add_picture(paragraph, item.image, width)
            continue
        line = item
        if line.rtl:
            _set_bidi(paragraph._p.get_or_add_pPr())
        for piece in line.parts:
            if piece is None:
                _page_field(paragraph, line, body, fonts)
                continue
            run = paragraph.add_run(piece)
            _style_run(run, line, body, fonts)
            if line.rtl:
                run.font.rtl = True


def _style_run(run: Any, line: RunningLine, body: float, fonts: FontBook) -> None:
    if line.bold:
        run.bold = run.font.cs_bold = True
    if line.italic:
        run.italic = run.font.cs_italic = True
    if abs(line.size - body) >= _BODY_SIZE / 2:
        _set_size(run._r.get_or_add_rPr(), line.size)
    if line.color != fonts.color:
        run.font.color.rgb = _rgb(line.color)
    family = fonts.family(line.font)
    if family != fonts.body:
        run.font.name = family


def _page_field(paragraph: Paragraph, line: RunningLine, body: float, fonts: FontBook) -> None:
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), " PAGE ")
    run = paragraph.add_run("1")
    _style_run(run, line, body, fonts)
    field.append(run._r)
    paragraph._p.append(field)
