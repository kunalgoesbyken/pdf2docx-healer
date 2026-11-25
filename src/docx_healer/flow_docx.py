"""Writing the assembled blocks into a python-docx document."""

import io
from typing import Any, Dict, List, Optional, Tuple
from xml.sax.saxutils import escape

import pymupdf
from docx.document import Document as DocumentType
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, RGBColor
from docx.text.paragraph import Paragraph

from .flow_lists import LevelDef, ListInstance, level_indents
from .flow_model import ImageItem, LinkKey, Para, Span, TableItem
from .flow_text import is_rtl_char
from .font_mapper import normalize_font_name
from .log import logger

_RELATIVE_SIZE = 0.12
_COMPLEX_START = 0x0590
_MONO_FONT = "Courier New"
_LEVELS = 9
_LINK_BLUE = RGBColor(0x05, 0x63, 0xC1)
_CENTER_SHARE = 0.05
_FALLBACK_STEP = 360
_BIDI_SUCCESSORS = (
    "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing",
    "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
    "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
    "w:pPrChange",
)


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
        text = "\u2022" if bullet else f"%{index + 1}."
        left, hanging = _FALLBACK_STEP * (index + 2), _FALLBACK_STEP
        start = 1
    else:
        fmt = level.fmt
        text = level.glyph if level.bullet else level.delim.replace("%", f"%{index + 1}")
        left, hanging = level_indents(level, margin)
        start = level.start
    return (
        f'<w:lvl w:ilvl="{index}"><w:start w:val="{start}"/><w:numFmt w:val="{fmt}"/>'
        f'<w:lvlText w:val="{escape(text)}"/><w:lvlJc w:val="left"/>'
        f'<w:pPr><w:ind w:left="{left}" w:hanging="{hanging}"/></w:pPr></w:lvl>'
    )


class Emitter:
    """Writes paragraphs, tables and images into one document."""

    def __init__(
        self, doc: DocumentType, body: float, text_width: float, page_width: float,
        anchors: Dict[LinkKey, str],
    ) -> None:
        self.doc = doc
        self.body = body
        self.text_width = text_width
        self.page_width = page_width
        self.anchors = anchors
        self.numbering = Numbering(doc)
        self.num_ids: Dict[int, int] = {}
        self.bookmark_id = 0
        self.styles: Dict[float, str] = {}
        self._link_style()
        doc.styles["Normal"].font.size = Pt(max(1.0, round(body * 2) / 2))

    def _link_style(self) -> None:
        styles = self.doc.styles
        if "Hyperlink" not in [s.name for s in styles]:
            style = styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER)
            style.font.color.rgb = _LINK_BLUE
            style.font.underline = True

    def register_instances(self, instances: List[ListInstance], margin: float) -> None:
        for index, instance in enumerate(instances):
            self.num_ids[index] = self.numbering.add(instance, margin)

    def _size_style(self, size: float) -> Optional[str]:
        if abs(size / self.body - 1) <= _RELATIVE_SIZE:
            return None
        key = round(size * 2) / 2
        name = self.styles.get(key)
        if name is None:
            name = f"Text {key:g}pt"
            style = self.doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            style.base_style = self.doc.styles["Normal"]
            style.font.size = Pt(key)
            self.styles[key] = name
        return name

    def paragraph(self, paragraph: Paragraph, para: Para) -> None:
        heading = para.kind == "heading"
        if heading:
            paragraph.style = self.doc.styles[f"Heading {para.level}"]
        elif para.kind == "title":
            paragraph.style = self.doc.styles["Title"]
            heading = True
        else:
            name = self._size_style(para.size)
            if name:
                paragraph.style = self.doc.styles[name]
        p = paragraph._p
        if para.kind == "list" and para.num is not None:
            ppr: Any = p.get_or_add_pPr()
            numpr = ppr.get_or_add_numPr()
            numpr.get_or_add_ilvl().val = para.level
            numpr.get_or_add_numId().val = self.num_ids[para.num[0]]
        if para.rtl:
            _set_bidi(p.get_or_add_pPr())
        self._runs(paragraph, para, heading)
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

    def _runs(self, paragraph: Paragraph, para: Para, heading: bool) -> None:
        current: Optional[Tuple[str, str]] = None
        container: Any = None
        for span in para.spans:
            target = self._target(span.link)
            if target != current:
                current = target
                container = self._open_link(paragraph, target) if target else None
            run = paragraph.add_run(span.text)
            self._format(run, span, para, heading)
            if container is not None:
                run.style = self.doc.styles["Hyperlink"]
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

    def _format(self, run: Any, span: Span, para: Para, heading: bool) -> None:
        if span.bold and not heading:
            run.bold = True
        if span.italic:
            run.italic = True
        if span.sup:
            run.font.superscript = True
        elif span.sub:
            run.font.subscript = True
        elif not heading and abs(span.size / para.size - 1) > _RELATIVE_SIZE:
            run.font.size = Pt(round(span.size * 2) / 2)
        if span.mono:
            run.font.name = _MONO_FONT
        elif any(ord(ch) >= _COMPLEX_START for ch in span.text):
            _complex_fonts(run, span)
        if any(is_rtl_char(ch) for ch in span.text):
            run.font.rtl = True

    def table(self, table: TableItem) -> None:
        columns = len(table.col_edges) - 1
        grid = self.doc.add_table(rows=table.rows, cols=columns)
        grid.style = self.doc.styles["Table Grid"]
        widths = [b - a for a, b in zip(table.col_edges, table.col_edges[1:])]
        scale = min(1.0, self.text_width / max(sum(widths), 1.0))
        for index, width in enumerate(widths):
            grid.columns[index].width = Pt(width * scale)
            for row in grid.rows:
                row.cells[index].width = Pt(width * scale)
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

    def image(self, image: ImageItem, center: bool) -> None:
        data = _usable_image(image)
        if data is None:
            return
        width = min(image.width, self.text_width)
        paragraph = self.doc.add_paragraph()
        try:
            paragraph.add_run().add_picture(io.BytesIO(data), width=Pt(width))
        except Exception as exc:
            logger.warning("Could not embed an image: %s", exc)
            body = paragraph._p.getparent()
            if body is not None:
                body.remove(paragraph._p)
            return
        if center:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _usable_image(image: ImageItem) -> Optional[bytes]:
    if image.ext.lower() in ("png", "jpeg", "jpg", "gif", "bmp", "tiff", "tif"):
        return image.data
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
