from pathlib import Path
from typing import Callable, List, Optional

import pymupdf
from docx import Document
from docx.oxml.ns import qn

from docx_healer.flow_writer import write_editable_docx

LEFT = 72
BODY = 10
PPR_ORDER = [
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr", "w:widowControl",
    "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens",
    "w:kinsoku", "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE",
    "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection",
    "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr",
    "w:sectPr", "w:pPrChange",
]
RPR_ORDER = [
    "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps", "w:strike",
    "w:dstrike", "w:outline", "w:shadow", "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid",
    "w:vanish", "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position", "w:sz",
    "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd", "w:fitText", "w:vertAlign",
    "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath",
]


def new_page(doc):
    doc.new_page()
    return doc[len(doc) - 1]


def put(page, x, y, text, size=BODY, font="helv"):
    page.insert_text((x, y), text, fontsize=size, fontname=font)


def lines(page, x, y, texts, size=BODY, font="helv", lead=None):
    lead = lead or size * 1.2
    for text in texts:
        put(page, x, y, text, size, font)
        y += lead
    return y


def para_gap(size=BODY):
    return size * 0.9


def build_pdf(path: Path, build: Callable) -> Path:
    doc = pymupdf.open()
    build(doc)
    doc.save(str(path))
    doc.close()
    return path


def convert(tmp_path: Path, build: Callable, name: str = "t", **kwargs) -> Path:
    pdf = build_pdf(tmp_path / f"{name}.pdf", build)
    out = tmp_path / f"{name}.docx"
    write_editable_docx(pdf, out, **kwargs)
    return out


def text_of(element) -> str:
    return "".join(t.text or "" for t in element.iter(qn("w:t")))


def items(path: Path) -> List[dict]:
    body = Document(str(path)).element.body
    found = []
    for child in body:
        if child.tag == qn("w:p"):
            found.append(paragraph_info(child))
        elif child.tag == qn("w:tbl"):
            found.append(table_info(child))
    return found


def paragraph_info(p) -> dict:
    ppr = p.find(qn("w:pPr"))
    style = None
    num = None
    bidi = False
    if ppr is not None:
        pstyle = ppr.find(qn("w:pStyle"))
        style = pstyle.get(qn("w:val")) if pstyle is not None else None
        numpr = ppr.find(qn("w:numPr"))
        if numpr is not None:
            num = (
                int(numpr.find(qn("w:numId")).get(qn("w:val"))),
                int(numpr.find(qn("w:ilvl")).get(qn("w:val"))),
            )
        bidi = ppr.find(qn("w:bidi")) is not None
    return {"kind": "p", "style": style, "text": text_of(p), "num": num, "bidi": bidi, "el": p}


def table_info(tbl) -> dict:
    rows = []
    header_rows = 0
    for tr in tbl.findall(qn("w:tr")):
        trpr = tr.find(qn("w:trPr"))
        if trpr is not None and trpr.find(qn("w:tblHeader")) is not None:
            header_rows += 1
        rows.append([text_of(tc) for tc in tr.findall(qn("w:tc"))])
    return {"kind": "tbl", "rows": rows, "header_rows": header_rows, "el": tbl}


def paragraphs(path: Path) -> List[dict]:
    return [item for item in items(path) if item["kind"] == "p"]


def texts(path: Path) -> List[str]:
    return [p["text"] for p in paragraphs(path) if p["text"].strip()]


def tables(path: Path) -> List[dict]:
    return [item for item in items(path) if item["kind"] == "tbl"]


def _roman(value: int) -> str:
    pairs = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
             (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
    out = ""
    for number, symbol in pairs:
        while value >= number:
            out += symbol
            value -= number
    return out


def _format(fmt: str, value: int) -> str:
    if fmt == "decimal":
        return str(value)
    if fmt == "lowerLetter":
        return chr(ord("a") + value - 1)
    if fmt == "upperLetter":
        return chr(ord("A") + value - 1)
    if fmt == "lowerRoman":
        return _roman(value)
    if fmt == "upperRoman":
        return _roman(value).upper()
    return "?"


def numbering_levels(path: Path) -> dict:
    numbering = Document(str(path)).part.numbering_part.element
    abstract = {}
    for an in numbering.findall(qn("w:abstractNum")):
        levels = {}
        for lvl in an.findall(qn("w:lvl")):
            start = lvl.find(qn("w:start"))
            levels[int(lvl.get(qn("w:ilvl")))] = (
                int(start.get(qn("w:val"))) if start is not None else 0,
                lvl.find(qn("w:numFmt")).get(qn("w:val")),
                lvl.find(qn("w:lvlText")).get(qn("w:val")),
            )
        abstract[an.get(qn("w:abstractNumId"))] = levels
    nums = {}
    for num in numbering.findall(qn("w:num")):
        nums[int(num.get(qn("w:numId")))] = abstract[
            num.find(qn("w:abstractNumId")).get(qn("w:val"))
        ]
    return nums


def labels(path: Path) -> List[Optional[str]]:
    nums = numbering_levels(path)
    counters: dict = {}
    result: List[Optional[str]] = []
    for p in paragraphs(path):
        if p["num"] is None:
            result.append(None)
            continue
        num_id, level = p["num"]
        defs = nums[num_id]
        state = counters.setdefault(num_id, {})
        for deeper in [k for k in state if k > level]:
            del state[deeper]
        state[level] = state.get(level, defs[level][0] - 1) + 1
        start, fmt, text = defs[level]
        if fmt == "bullet":
            result.append(text)
            continue
        label = text
        for index in range(level + 1):
            index_start, index_fmt, _ = defs[index]
            value = state.get(index, index_start)
            label = label.replace(f"%{index + 1}", _format(index_fmt, value))
        result.append(label)
    return result


def grid(page, x0, y0, widths, row_h, rows, skip_h=(), size=BODY, header_font="hebo"):
    xs = [x0]
    for width in widths:
        xs.append(xs[-1] + width)
    ys = [y0 + i * row_h for i in range(len(rows) + 1)]
    shape = page.new_shape()
    for ri, y in enumerate(ys):
        if ri in skip_h:
            shape.draw_line((xs[1], y), (xs[-1], y))
        else:
            shape.draw_line((xs[0], y), (xs[-1], y))
    for x in xs:
        shape.draw_line((x, ys[0]), (x, ys[-1]))
    shape.finish(width=0.8, color=(0, 0, 0))
    shape.commit()
    for ri, row in enumerate(rows):
        for ci, cell in enumerate(row):
            if cell:
                put(page, xs[ci] + 4, ys[ri] + row_h - 8, cell, size, header_font if ri == 0 else "helv")
    return ys[-1]


def png_bytes(width: int, height: int) -> bytes:
    pixmap = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, width, height))
    pixmap.set_rect(pixmap.irect, (200, 30, 30))
    return pixmap.tobytes("png")


def runs(page, x, y, parts, size=BODY):
    for part in parts:
        text, font = part[0], part[1]
        part_size = part[2] if len(part) > 2 else size
        rise = part[3] if len(part) > 3 else 0
        put(page, x, y - rise, text, part_size, font)
        x += pymupdf.get_text_length(text, fontname=font, fontsize=part_size)
    return x


def html_block(page, rect, html):
    page.insert_htmlbox(pymupdf.Rect(*rect), html)


def problems(path: Path) -> List[str]:
    doc = Document(str(path))
    body = doc.element.body
    found: List[str] = []
    num_ids = set(numbering_levels(path))
    for numid in body.iter(qn("w:numId")):
        if int(numid.get(qn("w:val"))) not in num_ids:
            found.append(f"unknown numId {numid.get(qn('w:val'))}")
    names = [b.get(qn("w:name")) for b in body.iter(qn("w:bookmarkStart"))]
    ids = [b.get(qn("w:id")) for b in body.iter(qn("w:bookmarkStart"))]
    ends = [b.get(qn("w:id")) for b in body.iter(qn("w:bookmarkEnd"))]
    if len(set(names)) != len(names) or len(set(ids)) != len(ids) or sorted(ids) != sorted(ends):
        found.append("bookmark names or ids are not unique and paired")
    for link in body.iter(qn("w:hyperlink")):
        anchor = link.get(qn("w:anchor"))
        rid = link.get(qn("r:id"))
        if anchor is not None and anchor not in names:
            found.append(f"dangling anchor {anchor}")
        if rid is not None and rid not in doc.part.rels:
            found.append(f"dangling relationship {rid}")
    for tc in body.iter(qn("w:tc")):
        if tc[-1].tag != qn("w:p") and tc[-1].tag != qn("w:tbl"):
            found.append("cell does not end with a paragraph")
    for tbl in body.iter(qn("w:tbl")):
        cols = len(tbl.find(qn("w:tblGrid")))
        for tr in tbl.findall(qn("w:tr")):
            span = 0
            for tc in tr.findall(qn("w:tc")):
                grid_span = tc.find(f"{qn('w:tcPr')}/{qn('w:gridSpan')}")
                span += int(grid_span.get(qn("w:val"))) if grid_span is not None else 1
            if span != cols:
                found.append(f"row spans {span} of {cols} columns")
    for tag, order in (("w:pPr", PPR_ORDER), ("w:rPr", RPR_ORDER)):
        for element in body.iter(qn(tag)):
            names = ["w:" + child.tag.split("}")[1] for child in element]
            positions = [order.index(name) for name in names if name in order]
            if positions != sorted(positions):
                found.append(f"{tag} children out of schema order: {names}")
    return found
