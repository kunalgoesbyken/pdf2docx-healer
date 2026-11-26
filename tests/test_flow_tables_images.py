import pymupdf
from docx import Document
from docx.oxml.ns import qn

from flow_fixtures import LEFT, convert, grid, new_page, png_bytes, put, tables, texts

HEADER = ["Name", "Value", "Notes"]


def _pdf_table(page, y):
    return grid(
        page, LEFT, y, [100, 100, 100], 26,
        [HEADER, ["tall", "1", "x"], ["", "2", "y"], ["c", "3", "z"]],
        skip_h=(2,),
    )


def test_ruled_table_becomes_a_word_table_with_header_and_merges_and_no_duplicate_text(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 80, "Intro sentence before the table.")
        _pdf_table(page, 100)
        put(page, LEFT, 300, "Closing sentence after the table.")

    out = convert(tmp_path, build)

    found = tables(out)
    assert len(found) == 1
    table = found[0]
    assert table["header_rows"] == 1
    assert table["el"].findall(qn("w:tr"))[0].find(qn("w:trPr")).find(qn("w:tblHeader")) is not None
    assert table["rows"] == [HEADER, ["tall", "1", "x"], ["", "2", "y"], ["c", "3", "z"]]
    assert table["el"].findall(f".//{qn('w:vMerge')}")
    assert texts(out) == ["Intro sentence before the table.", "Closing sentence after the table."]


def test_table_cut_by_a_page_break_is_one_table_with_a_single_header(tmp_path):
    def build(doc):
        page = new_page(doc)
        grid(page, LEFT, 640, [150, 150], 26, [["Item", "Price"], ["apple", "1"], ["pear", "2"]])
        page = new_page(doc)
        grid(page, LEFT, 90, [150, 150], 26, [["Item", "Price"], ["plum", "3"], ["fig", "4"]])

    out = convert(tmp_path, build)

    found = tables(out)
    assert len(found) == 1
    assert [row[0] for row in found[0]["rows"]] == ["Item", "apple", "pear", "plum", "fig"]
    assert found[0]["header_rows"] == 1


def test_images_are_embedded_scaled_and_tiny_ones_skipped(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 80, "A caption sentence sits above the figure.")
        page.insert_image(pymupdf.Rect(72, 100, 472, 300), stream=png_bytes(300, 150))
        page.insert_image(pymupdf.Rect(500, 100, 510, 110), stream=png_bytes(10, 10))
        put(page, LEFT, 340, "Text continues below the figure here.")

    out = convert(tmp_path, build)

    shapes = Document(str(out)).inline_shapes
    assert len(shapes) == 1
    assert shapes[0].width.pt <= 595 - 2 * 36
    assert abs(shapes[0].width / shapes[0].height - 2.0) < 0.05
    assert texts(out) == ["A caption sentence sits above the figure.", "Text continues below the figure here."]
