from collections import defaultdict
from typing import Dict, List, Optional, Sequence, Tuple, Union

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

from factories import make_docx

Item = Union[str, Tuple[str, ...]]

ROMAN = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
         (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]


def roman(value: int) -> str:
    out = ""
    for size, glyph in ROMAN:
        while value >= size:
            out += glyph
            value -= size
    return out


def format_value(fmt: str, value: int) -> str:
    if fmt == "decimal":
        return str(value)
    if fmt in ("lowerLetter", "upperLetter"):
        glyph = chr(ord("a") + value - 1)
        return glyph.upper() if fmt == "upperLetter" else glyph
    if fmt in ("lowerRoman", "upperRoman"):
        glyph = roman(value)
        return glyph.upper() if fmt == "upperRoman" else glyph
    raise AssertionError(fmt)


def build(doc, items: Sequence[Item]) -> None:
    for item in items:
        text, indent_pt, style = (item, None, None) if isinstance(item, str) else (tuple(item) + (None, None))[:3]
        paragraph = doc.add_paragraph()
        for index, piece in enumerate(text.split("\n")):
            if index:
                paragraph.add_run().add_break()
            paragraph.add_run(piece)
        if indent_pt is not None:
            paragraph.paragraph_format.left_indent = Pt(indent_pt)
        if style:
            paragraph.style = doc.styles[style]


def make(tmp_path, items: Sequence[Item], name: str = "in.docx"):
    return make_docx(tmp_path / name, lambda doc: build(doc, items))


def _level(numbering, num_id: str, ilvl: int):
    for num in numbering.findall(qn("w:num")):
        if num.get(qn("w:numId")) != num_id:
            continue
        abstract_id = num.find(qn("w:abstractNumId")).get(qn("w:val"))
        override = None
        for node in num.findall(qn("w:lvlOverride")):
            if int(node.get(qn("w:ilvl"))) == ilvl:
                override = node
        for abstract in numbering.findall(qn("w:abstractNum")):
            if abstract.get(qn("w:abstractNumId")) == abstract_id:
                for lvl in abstract.findall(qn("w:lvl")):
                    if int(lvl.get(qn("w:ilvl"))) == ilvl:
                        start = lvl.find(qn("w:start")).get(qn("w:val"))
                        if override is not None and override.find(qn("w:startOverride")) is not None:
                            start = override.find(qn("w:startOverride")).get(qn("w:val"))
                        return (
                            lvl.find(qn("w:numFmt")).get(qn("w:val")),
                            lvl.find(qn("w:lvlText")).get(qn("w:val")),
                            int(start),
                        )
    raise AssertionError(f"numId {num_id} level {ilvl} is not defined")


def info(paragraph) -> Dict[str, object]:
    """Describe how a paragraph is numbered: style, numId, ilvl, numFmt, start, lvlText."""
    pPr = paragraph._p.pPr
    numPr = pPr.numPr if pPr is not None else None
    result: Dict[str, object] = {"style": paragraph.style.name, "numId": None, "ilvl": None}
    if numPr is not None and numPr.numId is not None:
        num_id = str(numPr.numId.val)
        ilvl = numPr.ilvl.val if numPr.ilvl is not None else 0
        numbering = paragraph.part.numbering_part.element
        fmt, text, start = _level(numbering, num_id, ilvl)
        result.update(numId=num_id, ilvl=ilvl, numFmt=fmt, lvlText=text, start=start)
    return result


def rendered(path) -> List[str]:
    """Return what a reader sees: the auto-number label plus text for each non-empty paragraph."""
    doc = Document(str(path))
    counters: Dict[str, Dict[int, int]] = defaultdict(dict)
    lines = []
    for paragraph in doc.paragraphs:
        if not paragraph.text.strip():
            continue
        details = info(paragraph)
        label = ""
        if details["numId"] is not None:
            num_id, ilvl = str(details["numId"]), int(details["ilvl"])  # type: ignore[arg-type]
            fmt = details["numFmt"]
            if fmt == "bullet":
                label = "• "
            else:
                state = counters[num_id]
                state[ilvl] = state.get(ilvl, int(details["start"]) - 1) + 1  # type: ignore[arg-type]
                for deeper in [level for level in state if level > ilvl]:
                    del state[deeper]
                label = str(details["lvlText"]).replace(f"%{ilvl + 1}", format_value(str(fmt), state[ilvl])) + " "
        elif str(details["style"]).startswith("List Bullet"):
            label = "• "
        elif str(details["style"]).startswith("List Number"):
            label = "<style-numbered> "
        lines.append(label + paragraph.text)
    return lines


def paragraphs_info(path) -> List[Tuple[str, Dict[str, object]]]:
    doc = Document(str(path))
    return [(p.text, info(p)) for p in doc.paragraphs if p.text.strip()]


def num_ids(path, texts: Optional[Sequence[str]] = None) -> List[Optional[object]]:
    return [details["numId"] for text, details in paragraphs_info(path) if texts is None or text in texts]


def test_renderer_counts_from_start_and_formats_letters():
    assert format_value("lowerRoman", 4) == "iv"
    assert format_value("upperLetter", 2) == "B"
    assert format_value("decimal", 7) == "7"
