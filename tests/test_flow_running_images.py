import pymupdf
import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from flow_fixtures import LEFT, convert, new_page, png_bytes, problems, put


def _pictures(part):
    return len(part._element.findall(".//" + "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"))


def _pages(doc, count, build):
    for number in range(count):
        page = new_page(doc)
        put(page, LEFT, 200, f"Page {number + 1} body sentence with its own words.")
        build(page, number)


def test_logo_repeated_on_every_page_moves_into_the_header(tmp_path):
    def build(doc):
        _pages(doc, 3, lambda page, _: page.insert_image(
            pymupdf.Rect(480, 20, 560, 60), stream=png_bytes(80, 40)))

    out = convert(tmp_path, build)

    section = Document(str(out)).sections[0]
    assert len(Document(str(out)).inline_shapes) == 0
    assert _pictures(section.header) == 1
    assert section.header.paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.RIGHT
    assert not section.different_first_page_header_footer
    assert problems(out) == []


def test_logo_missing_from_the_first_page_goes_to_the_footer_after_page_one(tmp_path):
    def build(doc):
        _pages(doc, 4, lambda page, number: number and page.insert_image(
            pymupdf.Rect(270, 780, 330, 820), stream=png_bytes(60, 40)))

    out = convert(tmp_path, build)

    section = Document(str(out)).sections[0]
    assert section.different_first_page_header_footer
    assert _pictures(section.footer) == 1
    assert _pictures(section.first_page_footer) == 0
    assert section.footer.paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert problems(out) == []


def test_logo_and_running_title_share_the_header_in_reading_order(tmp_path):
    def build(doc):
        def add(page, _):
            put(page, LEFT, 70, "Annual Report 2024")
            page.insert_image(pymupdf.Rect(LEFT, 10, LEFT + 60, 50), stream=png_bytes(60, 40))
        _pages(doc, 3, add)

    out = convert(tmp_path, build)

    header = Document(str(out)).sections[0].header
    assert [_pictures(p) for p in header.paragraphs] == [1, 0]
    assert header.paragraphs[1].text == "Annual Report 2024"


def test_wide_repeated_banner_is_capped_to_the_text_width(tmp_path):
    def build(doc):
        _pages(doc, 3, lambda page, _: page.insert_image(
            pymupdf.Rect(0, 5, 595, 60), stream=png_bytes(300, 30)))

    out = convert(tmp_path, build)

    doc = Document(str(out))
    section = doc.sections[0]
    text_width = section.page_width - section.left_margin - section.right_margin
    extents = section.header._element.findall(
        ".//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent")
    assert len(extents) == 1
    assert int(extents[0].get("cx")) <= text_width


def test_same_picture_repeated_at_two_places_gives_two_running_pictures(tmp_path):
    def build(doc):
        def add(page, _):
            page.insert_image(pymupdf.Rect(LEFT, 10, LEFT + 60, 50), stream=png_bytes(60, 40))
            page.insert_image(pymupdf.Rect(LEFT, 790, LEFT + 60, 830), stream=png_bytes(60, 40))
        _pages(doc, 3, add)

    out = convert(tmp_path, build)

    section = Document(str(out)).sections[0]
    assert (_pictures(section.header), _pictures(section.footer)) == (1, 1)


@pytest.mark.parametrize("keep_proportion", [False, True])
def test_repeated_picture_under_the_body_text_is_dropped_not_moved_to_the_header(
    tmp_path, keep_proportion
):
    def build(doc):
        _pages(doc, 3, lambda page, _: page.insert_image(
            pymupdf.Rect(0, 0, 595, 842), stream=png_bytes(40, 40), overlay=False,
            keep_proportion=keep_proportion))

    out = convert(tmp_path, build)

    doc = Document(str(out))
    assert len(doc.inline_shapes) == 0
    assert _pictures(doc.sections[0].header) == 0
    assert _pictures(doc.sections[0].footer) == 0


def test_image_on_one_page_only_stays_in_the_body(tmp_path):
    def build(doc):
        _pages(doc, 4, lambda page, number: number == 2 and page.insert_image(
            pymupdf.Rect(100, 300, 300, 400), stream=png_bytes(100, 50)))

    out = convert(tmp_path, build)

    doc = Document(str(out))
    assert len(doc.inline_shapes) == 1
    assert _pictures(doc.sections[0].header) == 0
