from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple, Union

from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docx_healer.ocr_pipeline import pymupdf

Item = Union[str, Tuple[str, str]]

_VISUAL_RUN = ("rFonts", "b", "bCs", "i", "iCs", "color", "sz", "szCs", "caps", "u")
_VISUAL_PARA = ("keepNext", "keepLines", "spacing", "ind", "jc", "pageBreakBefore", "pBdr", "shd")


def add_direct_run(paragraph, text: str, size: int = 24, bold: bool = False) -> None:
    run = paragraph.add_run(text)
    props = run._r.get_or_add_rPr()
    fonts = OxmlElement("w:rFonts")
    for name in ("ascii", "hAnsi", "eastAsia"):
        fonts.set(qn("w:" + name), "DirectFont")
    props.append(fonts)
    if bold:
        props.append(OxmlElement("w:b"))
    italic = OxmlElement("w:i")
    italic.set(qn("w:val"), "0")
    props.append(italic)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "000000")
    props.append(color)
    point = OxmlElement("w:sz")
    point.set(qn("w:val"), str(size))
    props.append(point)


def make_paged_docx(path: Path, pages: Sequence[Sequence[Item]]) -> Path:
    doc = Document()
    for index, items in enumerate(pages):
        if index:
            doc.add_section(WD_SECTION.NEW_PAGE)
        for item in items:
            if isinstance(item, tuple):
                table = doc.add_table(rows=1, cols=1)
                add_direct_run(table.cell(0, 0).paragraphs[0], item[1])
            else:
                add_direct_run(doc.add_paragraph(), item)
    doc.save(str(path))
    return path


def make_toc_pdf(path: Path, toc: Sequence[Sequence[Any]], page_count: int) -> Path:
    doc = pymupdf.open()
    for _ in range(page_count):
        doc.new_page()
    if toc:
        doc.set_toc([list(entry) for entry in toc])
    doc.save(str(path))
    doc.close()
    return path


def body_styles(path: Path) -> List[Tuple[str, Optional[str]]]:
    doc = Document(str(path))
    result = []
    for child in doc.element.body.iterchildren(qn("w:p")):
        text = "".join(t.text or "" for t in child.iter(qn("w:t")))
        style = child.find(qn("w:pPr") + "/" + qn("w:pStyle"))
        result.append((text, style.get(qn("w:val")) if style is not None else None))
    return result


def _style_chain(doc, style_id: Optional[str]) -> List[Any]:
    styles = doc.styles.element
    chain: List[Any] = []
    seen = set()
    while style_id and style_id not in seen:
        seen.add(style_id)
        found = [
            s for s in styles.iterchildren(qn("w:style")) if s.get(qn("w:styleId")) == style_id
        ]
        if not found:
            break
        chain.append(found[0])
        based = found[0].find(qn("w:basedOn"))
        style_id = based.get(qn("w:val")) if based is not None else None
    return list(reversed(chain))


def _merge(target: Dict[str, Dict[str, str]], props: Any, wanted: Sequence[str]) -> None:
    if props is None:
        return
    for child in props:
        name = child.tag.split("}")[1]
        if name in wanted:
            attributes = {key.split("}")[1]: value for key, value in child.attrib.items()}
            if not attributes:
                attributes = {"val": "1"}
            if name == "rFonts":
                target.setdefault(name, {}).update(attributes)
            else:
                target[name] = attributes


def effective_formatting(
    doc, paragraph_element: Any, run_element: Any
) -> Dict[str, Dict[str, str]]:
    style = paragraph_element.find(qn("w:pPr") + "/" + qn("w:pStyle"))
    style_id = style.get(qn("w:val")) if style is not None else "Normal"
    defaults = doc.styles.element.find(qn("w:docDefaults"))
    result: Dict[str, Dict[str, str]] = {}
    if defaults is not None:
        _merge(result, defaults.find(qn("w:rPrDefault") + "/" + qn("w:rPr")), _VISUAL_RUN)
        _merge(result, defaults.find(qn("w:pPrDefault") + "/" + qn("w:pPr")), _VISUAL_PARA)
    for node in _style_chain(doc, style_id):
        _merge(result, node.find(qn("w:rPr")), _VISUAL_RUN)
        _merge(result, node.find(qn("w:pPr")), _VISUAL_PARA)
    _merge(result, paragraph_element.find(qn("w:pPr")), _VISUAL_PARA)
    _merge(result, run_element.find(qn("w:rPr")), _VISUAL_RUN)
    return result
