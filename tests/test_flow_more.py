import pymupdf
from docx import Document

from flow_fixtures import (
    LEFT,
    convert,
    grid,
    html_block,
    lines,
    new_page,
    paragraphs,
    png_bytes,
    problems,
    put,
    tables,
    texts,
)


def test_two_known_words_split_by_a_hyphen_stay_compound(tmp_path):
    def build(doc):
        page = new_page(doc)
        y = lines(page, LEFT, 100, ["It replaces sequence-", "aligned recurrence everywhere."])
        put(page, LEFT, y + 20, "Each sequence is aligned with the next one.")

    out = convert(tmp_path, build)

    assert texts(out)[0] == "It replaces sequence-aligned recurrence everywhere."


def test_section_number_on_its_own_text_run_is_glued_to_the_heading(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 100, "3.1", 14, "hebo")
        put(page, LEFT + 36, 100, "Encoder and Decoder", 14, "hebo")
        lines(page, LEFT, 140, ["Body text sits here and goes on for a while", "before it stops."])
        doc.set_toc([[1, "Encoder and Decoder", 1]])

    out = convert(tmp_path, build)

    assert [(p["style"], p["text"]) for p in paragraphs(out) if p["text"]][0] == (
        "Heading1", "3.1 Encoder and Decoder",
    )


def test_vector_figure_becomes_one_picture_and_its_labels_are_not_duplicated(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 80, "Text before the figure sits above it.")
        shape = page.new_shape()
        for index in range(6):
            shape.draw_rect(pymupdf.Rect(100 + index * 30, 120 + index * 15, 160 + index * 30, 170 + index * 15))
        shape.finish(width=1, color=(0, 0, 1), fill=(0.8, 0.9, 1))
        shape.commit()
        put(page, 110, 140, "node", 6)
        put(page, LEFT, 330, "Text after the figure sits below it.")

    out = convert(tmp_path, build)

    assert len(Document(str(out)).inline_shapes) == 1
    assert texts(out) == ["Text before the figure sits above it.", "Text after the figure sits below it."]


def test_page_with_a_rotation_flag_keeps_its_text_and_links(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 100, "Rotated page keeps this sentence readable.")
        page.insert_link({
            "kind": pymupdf.LINK_URI, "from": page.search_for("sentence")[0], "uri": "https://example.org/r",
        })
        page.set_rotation(90)

    out = convert(tmp_path, build)

    assert texts(out) == ["Rotated page keeps this sentence readable."]


def test_plain_header_row_is_not_marked_and_a_lone_box_is_not_a_table(tmp_path):
    def build(doc):
        page = new_page(doc)
        grid(page, LEFT, 100, [100, 100], 26, [["Name", "Value"], ["a", "1"]], header_font="helv")
        page.draw_rect(pymupdf.Rect(72, 300, 300, 340), color=(0, 0, 0))
        put(page, 80, 325, "Boxed note text inside a frame.")

    out = convert(tmp_path, build)

    found = tables(out)
    assert len(found) == 1
    assert found[0]["header_rows"] == 0
    assert "Boxed note text inside a frame." in texts(out)


def test_full_page_background_and_repeated_logo_images_are_dropped(tmp_path):
    def build(doc):
        for number in range(3):
            page = new_page(doc)
            page.insert_image(pymupdf.Rect(0, 0, 595, 842), stream=png_bytes(40, 40), overlay=False)
            page.insert_image(pymupdf.Rect(480, 20, 560, 60), stream=png_bytes(80, 40))
            put(page, LEFT, 200, f"Page {number + 1} body sentence with its own words.")
        page = doc[2]
        page.insert_image(pymupdf.Rect(100, 300, 300, 400), stream=png_bytes(100, 50))

    out = convert(tmp_path, build)

    assert len(Document(str(out)).inline_shapes) == 1


def test_rich_document_is_structurally_valid(tmp_path):
    def build(doc):
        new_page(doc)
        new_page(doc)
        page = doc[0]
        put(page, LEFT, 90, "Report Title Here", 20, "hebo")
        y = lines(page, LEFT, 130, ["A paragraph with a link inside it that goes", "on for a second line."])
        page.insert_link({
            "kind": pymupdf.LINK_GOTO, "from": page.search_for("link inside")[0], "page": 1,
            "to": pymupdf.Point(72, 100), "zoom": 0,
        })
        y = lines(page, LEFT + 14, y + 10, ["1. first step", "2. second step"])
        grid(page, LEFT, y + 20, [100, 100], 26, [["Name", "Value"], ["a", "1"]])
        html_block(doc[1], (72, 90, 500, 160), '<p dir="rtl">שלום עולם זה טקסט בעברית</p>')
        put(doc[1], LEFT, 200, "Target paragraph on page two.")

    out = convert(tmp_path, build)

    assert problems(out) == []


def test_table_cut_by_a_column_break_is_one_table(tmp_path):
    left = ["Left column line one runs along here", "left column line two keeps going on",
            "left column line three is here too"]
    right = ["Right column line one runs along here", "right column line two keeps going on",
             "right column line three is here too"]

    def build(doc):
        page = new_page(doc)
        y = lines(page, LEFT, 100, left)
        grid(page, LEFT, y + 20, [100, 100], 26, [["Item", "Price"], ["apple", "1"], ["pear", "2"]])
        y = grid(page, 330, 100, [100, 100], 26, [["Item", "Price"], ["plum", "3"], ["fig", "4"]])
        lines(page, 330, y + 20, right)

    out = convert(tmp_path, build)

    found = tables(out)
    assert len(found) == 1
    assert [row[0] for row in found[0]["rows"]] == ["Item", "apple", "pear", "plum", "fig"]
    assert texts(out) == [" ".join(left), " ".join(right)]


def test_single_weak_bullet_and_lone_number_stay_literal(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 100, "- 5 degrees colder than yesterday morning.")
        put(page, LEFT, 140, "Then a plain sentence follows after it here.")

    out = convert(tmp_path, build)

    assert texts(out) == ["- 5 degrees colder than yesterday morning.", "Then a plain sentence follows after it here."]
    assert all(p["num"] is None for p in paragraphs(out))
