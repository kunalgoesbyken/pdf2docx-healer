import logging

import pymupdf
from docx import Document
from docx.oxml.ns import qn

from docx_healer.flow_model import FlowLimits
from flow_fixtures import (
    LEFT,
    convert,
    grid,
    html_block,
    labels,
    lines,
    new_page,
    numbering_levels,
    paragraphs,
    png_bytes,
    put,
    tables,
    text_of,
    texts,
)

_LATE = ["3.", "4.", "5."]
_OUTLINE = [
    ("1.", 0, "Overview section"),
    ("1.1", 1, "Scope of work"),
    ("1.2", 1, "Intended audience"),
    ("2.", 0, "Method section"),
    ("2.1", 1, "Data sources used"),
]


def _shapes(path):
    return len(Document(str(path)).inline_shapes)


def _without_jfif(data):
    assert data[:2] == b"\xff\xd8" and data[2:4] == b"\xff\xe0"
    length = int.from_bytes(data[4:6], "big")
    return data[:2] + data[4 + length:]


def _jpeg(width, height):
    pixmap = pymupdf.Pixmap(pymupdf.csGRAY, pymupdf.IRect(0, 0, width, height))
    pixmap.set_rect(pixmap.irect, (90,))
    return pixmap.tobytes("jpeg")


def _bars(page, x, y, width, height, count=6):
    shape = page.new_shape()
    step = width / count
    for index in range(count):
        top = y + height * (0.2 + 0.1 * (index % 4))
        shape.draw_rect(pymupdf.Rect(x + index * step, top, x + (index + 0.8) * step, y + height))
    shape.finish(width=0.5, color=(0, 0, 0), fill=(0.3, 0.5, 0.9))
    shape.commit()


def _item(page, x, y, marker, text):
    put(page, x, y, marker)
    put(page, x + 28, y, text)
    return y + 16


def _list_paragraphs(path):
    return [p for p in paragraphs(path) if p["num"] is not None]


def test_scanned_page_keeps_its_page_sized_image(tmp_path, caplog):
    def build(doc):
        page = new_page(doc)
        page.insert_image(page.rect, stream=png_bytes(60, 80))

    with caplog.at_level(logging.WARNING, logger="docx_healer"):
        out = convert(tmp_path, build)

    assert _shapes(out) == 1
    assert texts(out) == []
    assert "no text" in caplog.text.lower()


def test_scanned_pages_with_the_same_image_each_keep_it(tmp_path):
    def build(doc):
        for _ in range(3):
            page = new_page(doc)
            page.insert_image(page.rect, stream=png_bytes(60, 80))

    assert _shapes(convert(tmp_path, build)) == 3


def test_text_page_still_drops_a_page_sized_background(tmp_path):
    def build(doc):
        for number in range(2):
            page = new_page(doc)
            page.insert_image(page.rect, stream=png_bytes(30, 40), overlay=False)
            put(page, LEFT, 200, f"Page {number} body words that are different here.")

    assert _shapes(convert(tmp_path, build)) == 0


def test_jpeg_that_word_cannot_recognise_is_transcoded_and_embedded(tmp_path):
    plain = _without_jfif(_jpeg(80, 60))

    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 80, "A caption above the picture sits here.")
        page.insert_image(pymupdf.Rect(72, 100, 272, 250), stream=plain)

    assert _shapes(convert(tmp_path, build)) == 1


def test_icon_is_kept_while_hairlines_and_specks_are_skipped(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 60, "Text with small pictures around it sits here.")
        page.insert_image(pymupdf.Rect(72, 100, 89, 119), stream=png_bytes(34, 38))
        page.insert_image(pymupdf.Rect(72, 150, 372, 152), stream=png_bytes(300, 2))
        page.insert_image(pymupdf.Rect(400, 100, 410, 110), stream=png_bytes(10, 10))

    assert _shapes(convert(tmp_path, build)) == 1


def test_one_raster_placed_twice_and_a_vector_figure_give_three_pictures(tmp_path):
    data = png_bytes(60, 40)

    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 60, "Text above the pictures sits here.")
        page.insert_image(pymupdf.Rect(72, 100, 172, 160), stream=data)
        page.insert_image(pymupdf.Rect(300, 100, 400, 160), stream=data)
        _bars(page, 72, 300, 240, 90)

    assert _shapes(convert(tmp_path, build)) == 3


def test_stacked_subcharts_make_one_figure_and_far_figures_stay_apart(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 60, "Text above the charts sits here alone.")
        for index in range(3):
            _bars(page, 100, 100 + index * 50, 300, 44)
            put(page, 220, 140 + index * 50, f"sub {index}", 7)
        _bars(page, 100, 500, 300, 80)
        put(page, LEFT, 600, "Caption below the charts sits here alone.")

    out = convert(tmp_path, build)

    assert _shapes(out) == 2
    assert texts(out) == [
        "Text above the charts sits here alone.", "Caption below the charts sits here alone.",
    ]


def test_list_that_starts_at_three_after_a_sentence_keeps_its_numbers(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 100, "The next list picks up at three and goes on.")
        y = 116
        for marker, word in zip(_LATE, ["backup", "writers", "switch"]):
            y = _item(page, LEFT, y, marker, f"Take the {word}")

    out = convert(tmp_path, build)

    assert [p["text"] for p in _list_paragraphs(out)] == [
        "Take the backup", "Take the writers", "Take the switch",
    ]
    assert [label for label in labels(out) if label] == _LATE
    assert numbering_levels(out)[_list_paragraphs(out)[0]["num"][0]][0][0] == 3


def test_late_number_without_a_sibling_stays_text(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 100, "The next paragraph ends here for good.")
        _item(page, LEFT, 116, "3.", "Take the backup")

    out = convert(tmp_path, build)

    assert _list_paragraphs(out) == []
    assert "3. Take the backup" in texts(out)[-1]


def test_late_number_after_an_open_sentence_stays_in_the_paragraph(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 100, "A sentence that does not end yet and carries on")
        y = _item(page, LEFT, 112, "3.", "alpha")
        _item(page, LEFT, y, "4.", "beta")

    assert _list_paragraphs(convert(tmp_path, build)) == []


def test_sibling_search_is_bounded_and_exactly_at_the_limit_works(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 100, "The next list picks up at three and goes on.")
        _item(page, LEFT, 116, "3.", "alpha")
        put(page, LEFT + 28, 130, "continuation one")
        put(page, LEFT + 28, 142, "continuation two")
        _item(page, LEFT, 154, "4.", "beta")

    inside = convert(tmp_path, build, name="a", limits=FlowLimits(max_marker_lookahead=3))
    outside = convert(tmp_path, build, name="b", limits=FlowLimits(max_marker_lookahead=2))

    assert len(_list_paragraphs(inside)) == 2
    assert _list_paragraphs(outside) == []


def test_dotted_outline_numbers_become_a_multilevel_list_with_the_pdf_labels(tmp_path):
    def build(doc):
        page = new_page(doc)
        y = 100
        for marker, level, text in _OUTLINE:
            y = _item(page, LEFT + level * 28, y, marker, text)

    out = convert(tmp_path, build)

    assert [p["text"] for p in _list_paragraphs(out)] == [t for _, _, t in _OUTLINE]
    assert [p["num"][1] for p in _list_paragraphs(out)] == [0, 1, 1, 0, 1]
    assert [label for label in labels(out) if label] == [m for m, _, _ in _OUTLINE]


def test_dotted_outline_flush_left_still_nests_by_depth(tmp_path):
    def build(doc):
        page = new_page(doc)
        y = 100
        for marker, _, text in _OUTLINE:
            y = _item(page, LEFT, y, marker, text)

    out = convert(tmp_path, build)

    assert [p["num"][1] for p in _list_paragraphs(out)] == [0, 1, 1, 0, 1]
    assert [label for label in labels(out) if label] == [m for m, _, _ in _OUTLINE]


def test_dotted_numbers_without_a_numbered_parent_stay_text(tmp_path):
    def build(doc):
        page = new_page(doc)
        y = _item(page, LEFT, 100, "1.1", "Scope of work")
        y = _item(page, LEFT, y + 10, "1.2", "Intended audience")
        put(page, LEFT, y + 30, "The figure 3.5 million people live here and 2.5 more.")

    out = convert(tmp_path, build)

    assert _list_paragraphs(out) == []
    assert any("3.5 million" in t for t in texts(out))


def test_dotted_number_that_skips_a_step_is_not_a_list_item(tmp_path):
    def build(doc):
        page = new_page(doc)
        y = _item(page, LEFT, 100, "1.", "Overview section")
        y = _item(page, LEFT + 28, y, "1.1", "Scope of work")
        y = _item(page, LEFT + 28, y, "1.3", "Skipped one")
        _item(page, LEFT, y, "2.", "Method section")

    out = convert(tmp_path, build)

    shown = [label for label in labels(out) if label]
    assert "1.3" not in shown
    assert shown[:2] == ["1.", "1.1"]


def _running(doc, pages=4):
    for number in range(1, pages + 1):
        page = new_page(doc)
        put(page, LEFT, 36, "Annual Report 2024")
        put(page, 300, 790, str(number))
        put(page, LEFT, 120, f"Unique body sentence number {number} sits here alone.")


def test_running_header_and_page_number_move_into_word_header_and_footer(tmp_path):
    out = convert(tmp_path, _running)

    section = Document(str(out)).sections[0]
    assert [p.text for p in section.header.paragraphs if p.text] == ["Annual Report 2024"]
    footer = section.footer._element.xml
    assert 'PAGE' in footer
    assert len(texts(out)) == 4
    assert all("Annual Report" not in t for t in texts(out))


def test_text_that_changes_in_the_margin_stays_in_the_body(tmp_path):
    def build(doc):
        for number, title in enumerate(["Alpha", "Beta", "Gamma"], start=1):
            page = new_page(doc)
            put(page, LEFT, 36, f"Chapter {title}")
            put(page, LEFT, 120, f"Unique body sentence number {number} sits here alone.")

    out = convert(tmp_path, build)

    section = Document(str(out)).sections[0]
    assert section.header.is_linked_to_previous or not [p for p in section.header.paragraphs if p.text]
    assert texts(out)[0] == "Chapter Alpha"


def test_header_text_with_markup_characters_is_kept_exactly(tmp_path):
    def build(doc):
        for number in range(1, 4):
            page = new_page(doc)
            put(page, LEFT, 36, "R&D <Q3> report")
            put(page, LEFT, 120, f"Unique body sentence number {number} sits here alone.")

    out = convert(tmp_path, build)

    header = Document(str(out)).sections[0].header
    assert [p.text for p in header.paragraphs if p.text] == ["R&D <Q3> report"]


def test_two_separate_running_lines_both_reach_the_header(tmp_path):
    def build(doc):
        for number in range(1, 4):
            page = new_page(doc)
            put(page, LEFT, 36, "Left running title")
            put(page, 400, 36, "Right running tag")
            put(page, LEFT, 120, f"Unique body sentence number {number} sits here alone.")

    out = convert(tmp_path, build)

    header = Document(str(out)).sections[0].header
    assert [p.text for p in header.paragraphs if p.text] == ["Left running title", "Right running tag"]


def _rtl_columns(page, right_word, left_word):
    for x0, word in ((330, right_word), (72, left_word)):
        html = "".join(f'<p dir="rtl">{word} {i} אחד שניים שלושה ארבעה.</p>' for i in range(4))
        html_block(page, (x0, 100, x0 + 190, 300), html)


def test_right_to_left_page_reads_the_right_column_first(tmp_path):
    def build(doc):
        _rtl_columns(new_page(doc), "ימין", "שמאל")

    out = convert(tmp_path, build)

    found = texts(out)
    assert found[0].startswith("ימין")
    assert found[-1].startswith("שמאל")


def test_left_to_right_page_with_one_hebrew_unit_still_reads_left_first(tmp_path):
    def build(doc):
        page = new_page(doc)
        lines(page, 72, 100, [f"Left column text line {i} goes on here" for i in range(4)])
        lines(page, 330, 100, [f"Right column text line {i} goes on here" for i in range(3)])
        html_block(page, (330, 200, 520, 228), '<p dir="rtl">אחד שניים שלושה</p>')

    out = convert(tmp_path, build)

    assert texts(out)[0].startswith("Left column")


def test_table_with_empty_spacer_columns_is_still_a_table(tmp_path):
    rows = [["Name", "", "", "", "Role", "", "", "", "Team"]] + [
        [f"n{i}", "", "", "", f"r{i}", "", "", "", f"t{i}"] for i in range(5)
    ]

    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 80, "Intro sentence before the grid sits here.")
        grid(page, 72, 100, [50, 5, 5, 5, 120, 5, 5, 5, 90], 22, rows)

    out = convert(tmp_path, build)

    assert len(tables(out)) == 1


def test_lattice_that_is_mostly_empty_stays_text(tmp_path):
    rows = [["only", "", "", "", ""]] + [["", "", "", "", ""] for _ in range(5)]

    def build(doc):
        page = new_page(doc)
        grid(page, 72, 100, [60, 60, 60, 60, 60], 22, rows)

    assert tables(convert(tmp_path, build)) == []


def test_huge_scanned_page_is_clamped_to_the_word_page_limit(tmp_path):
    def build(doc):
        page = doc.new_page(width=2320, height=3408)
        page.insert_image(page.rect, stream=png_bytes(60, 80))

    out = convert(tmp_path, build)

    size = Document(str(out)).sections[0]
    assert size.page_width.twips <= 31680 and size.page_height.twips <= 31680
    assert abs(size.page_width / size.page_height - 2320 / 3408) < 0.01
    assert _shapes(out) == 1


def test_scans_among_text_pages_are_not_treated_as_a_running_image(tmp_path):
    def build(doc):
        for number in range(4):
            page = new_page(doc)
            put(page, LEFT, 200, f"Page {number} body sentence with its own words.")
        for _ in range(4):
            page = new_page(doc)
            page.insert_image(page.rect, stream=png_bytes(60, 80))

    assert _shapes(convert(tmp_path, build)) == 4


def test_running_text_missing_from_the_first_page_uses_a_different_first_page_header(tmp_path):
    def build(doc):
        for number in range(1, 5):
            page = new_page(doc)
            if number > 1:
                put(page, LEFT, 36, "Annual Report 2024")
            put(page, LEFT, 120, f"Unique body sentence number {number} sits here alone.")

    out = convert(tmp_path, build)

    section = Document(str(out)).sections[0]
    assert section.different_first_page_header_footer
    assert [p.text for p in section.header.paragraphs if p.text] == ["Annual Report 2024"]
    assert [p.text for p in section.first_page_header.paragraphs if p.text] == []


def test_page_number_with_words_around_it_becomes_a_field_between_the_words(tmp_path):
    def build(doc):
        for number in range(1, 5):
            page = new_page(doc)
            put(page, LEFT, 790, f"Page {number} of 4")
            put(page, LEFT, 120, f"Unique body sentence number {number} sits here alone.")

    out = convert(tmp_path, build)

    footer = Document(str(out)).sections[0].footer._element.xml
    assert "PAGE" in footer
    assert text_of(Document(str(out)).sections[0].footer._element) == "Page 1 of 4"


def test_a_number_that_varies_without_following_the_page_is_not_put_in_the_header(tmp_path):
    def build(doc):
        for number in range(1, 5):
            page = new_page(doc)
            put(page, LEFT, 36, f"Section {number * 7} notes")
            put(page, LEFT, 120, f"Unique body sentence number {number} sits here alone.")

    out = convert(tmp_path, build)

    section = Document(str(out)).sections[0]
    assert [p.text for p in section.header.paragraphs if p.text] == []


def test_figure_cluster_budget_leaves_the_rest_as_text(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 60, "Text above the charts sits here alone.")
        for index in range(3):
            _bars(page, 100, 100 + index * 200, 300, 80)

    inside = convert(tmp_path, build, name="a", limits=FlowLimits(max_figure_clusters=3))
    outside = convert(tmp_path, build, name="b", limits=FlowLimits(max_figure_clusters=2))

    assert _shapes(inside) == 3
    assert _shapes(outside) == 2


def test_internal_link_into_a_table_cell_gets_a_bookmark_in_the_cell(tmp_path):
    rows = [["Name", "Role", "Team"]] + [[f"n{i}", f"r{i}", f"t{i}"] for i in range(3)]

    def build(doc):
        new_page(doc)
        page = new_page(doc)
        put(page, LEFT, 80, "Intro sentence before the grid sits here.")
        grid(page, 72, 100, [100, 100, 100], 26, rows)
        first = doc[0]
        put(first, LEFT, 100, "Jump to the roles table now.")
        link = first.search_for("roles table")[0]
        first.insert_link({
            "kind": pymupdf.LINK_GOTO, "from": link, "page": 1,
            "to": pymupdf.Point(176, 154), "zoom": 0,
        })

    out = convert(tmp_path, build)

    xml = Document(str(out)).element
    anchors = [h.get(qn("w:anchor")) for h in xml.iter(qn("w:hyperlink"))]
    cell_marks = [
        b.get(qn("w:name")) for tc in xml.iter(qn("w:tc")) for b in tc.iter(qn("w:bookmarkStart"))
    ]
    assert len(anchors) == 1 and anchors == cell_marks


def test_a_wrapped_line_that_starts_with_a_decimal_number_stays_in_its_paragraph(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 100, "The survey counted the people living in the north region.")
        put(page, LEFT, 112, "1.1 million people were found there across the last year.")

    out = convert(tmp_path, build)

    assert len(texts(out)) == 1
