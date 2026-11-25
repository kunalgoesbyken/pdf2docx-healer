from typing import Any, Dict, List

import pytest
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

from convert_support import red_png
from docx_healer.ocr_pipeline import pymupdf
from docx_healer.page_recovery import append_plain_page


class StubPage:
    def __init__(self, blocks: List[Dict[str, Any]], width: float = 400, height: float = 600):
        self._blocks = blocks
        self.rect = pymupdf.Rect(0, 0, width, height)

    def get_text(self, option: str) -> Dict[str, Any]:
        assert option == "dict"
        return {"blocks": self._blocks}


def text_block(*lines: str, bbox=(50, 50, 350, 80), flags: int = 0, color: int = 0) -> Dict:
    return {
        "type": 0,
        "bbox": bbox,
        "lines": [
            {"spans": [{"text": line, "size": 11.0, "flags": flags, "color": color}]}
            for line in lines
        ],
    }


def image_block(data: bytes, bbox=(50, 100, 150, 200)) -> Dict:
    return {"type": 1, "bbox": bbox, "image": data, "ext": "png"}


def paragraph_texts(document) -> List[str]:
    return [
        "".join(t.text or "" for t in p.iter(qn("w:t")))
        for p in document.element.body.iter(qn("w:p"))
    ]


def real_page(draw) -> Any:
    doc = pymupdf.open()
    page = doc.new_page()
    draw(page)
    return page


def test_bold_italic_size_and_colour_are_kept():
    def draw(page):
        page.insert_text((72, 100), "Bold words", fontname="hebo", fontsize=14, color=(1, 0, 0))
        page.insert_text((72, 130), "Slanted words", fontname="heit", fontsize=9)

    document = Document()
    append_plain_page(document, real_page(draw))

    runs = {run.text: run for p in document.paragraphs for run in p.runs if run.text}
    bold, slanted = runs["Bold words"], runs["Slanted words"]
    assert bold.font.bold and not bold.font.italic
    assert slanted.font.italic and not slanted.font.bold
    assert bold.font.size == Pt(14)
    assert str(bold.font.color.rgb) == "FF0000"
    assert slanted.font.color.rgb is None


def test_lines_of_a_block_stay_separate_without_adding_spaces():
    document = Document()
    page = StubPage([text_block("日本語の", "テキスト")])

    append_plain_page(document, page)

    paragraph = document.paragraphs[0]
    assert paragraph.text == "日本語の\nテキスト"
    assert len(paragraph._p.findall(f".//{qn('w:br')}")) == 1


def test_xml_illegal_control_characters_are_stripped_but_text_survives():
    document = Document()
    page = StubPage([text_block("clean\x00te\x08xt\x0b￾ end tab\there")])

    append_plain_page(document, page)

    assert document.paragraphs[0].text == "cleantext end tab\there"


def test_blocks_without_visible_text_are_skipped():
    document = Document()
    page = StubPage([text_block("   "), text_block("kept")])

    append_plain_page(document, page)

    assert [t for t in paragraph_texts(document) if t] == ["kept"]


def test_page_size_and_margins_follow_the_pdf_page():
    document = Document()
    page = StubPage([text_block("body", bbox=(30, 40, 370, 560))], width=400, height=600)

    append_plain_page(document, page)

    section = document.sections[0]
    assert (section.page_width.pt, section.page_height.pt) == (400, 600)
    assert (section.left_margin.pt, section.right_margin.pt) == (30, 30)
    assert (section.top_margin.pt, section.bottom_margin.pt) == (40, 40)


@pytest.mark.parametrize("bbox", [(0, 0, 400, 600), (390, 10, 400, 20), (-50, -50, 900, 900)])
def test_margins_never_leave_a_degenerate_text_area(bbox):
    document = Document()
    page = StubPage([text_block("body", bbox=bbox)], width=400, height=600)

    append_plain_page(document, page)

    section = document.sections[0]
    area = section.page_width.pt - section.left_margin.pt - section.right_margin.pt
    assert area >= 36
    assert min(section.left_margin.pt, section.top_margin.pt) >= 0


def test_page_without_content_still_opens_a_page():
    document = Document()
    document.add_paragraph("before")

    append_plain_page(document, StubPage([]))

    assert len(document.sections) == 2
    assert len(document.paragraphs) == 3


def test_new_page_starts_after_table_only_content():
    document = Document()
    document.add_table(rows=1, cols=1)

    append_plain_page(document, StubPage([text_block("after the table")]))

    assert len(document.sections) == 2


def test_gap_between_blocks_becomes_spacing_before():
    document = Document()
    page = StubPage(
        [text_block("one", bbox=(50, 50, 350, 70)), text_block("two", bbox=(50, 90, 350, 110))]
    )

    append_plain_page(document, page)

    first, second = [p for p in document.paragraphs if p.text]
    assert first.paragraph_format.space_before.pt == 0
    assert second.paragraph_format.space_before.pt == 20


def test_overlapping_blocks_never_get_negative_spacing():
    document = Document()
    page = StubPage(
        [text_block("one", bbox=(50, 50, 350, 90)), text_block("two", bbox=(50, 60, 350, 80))]
    )

    append_plain_page(document, page)

    assert document.paragraphs[1].paragraph_format.space_before.pt == 0


def test_picture_is_embedded_and_counted_as_kept():
    document = Document()

    skipped = append_plain_page(document, StubPage([image_block(red_png())]))

    assert skipped == 0
    assert len(list(document.element.body.iter(qn("w:drawing")))) == 1


def test_image_in_a_format_python_docx_rejects_is_reencoded():
    pixmap = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 8, 8), False)
    pixmap.set_rect(pixmap.irect, (10, 20, 30))
    document = Document()

    skipped = append_plain_page(document, StubPage([image_block(pixmap.tobytes("pnm"))]))

    assert skipped == 0
    assert len(list(document.element.body.iter(qn("w:drawing")))) == 1


def test_undecodable_image_is_skipped_and_leaves_no_empty_paragraph():
    document = Document()
    before = len(document.paragraphs)

    skipped = append_plain_page(document, StubPage([image_block(b"not an image at all")]))

    assert skipped == 1
    assert len(document.paragraphs) == before
