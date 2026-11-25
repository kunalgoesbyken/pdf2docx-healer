import logging

import pymupdf
import pytest
from docx import Document

from docx_healer import flow_extract
from docx_healer.flow_model import FlowLimits
from docx_healer.flow_text import clean_text
from flow_fixtures import (
    LEFT,
    convert,
    html_block,
    lines,
    new_page,
    para_gap,
    paragraphs,
    png_bytes,
    put,
    texts,
)


def _blank(doc):
    new_page(doc)


def test_blank_pdf_gives_a_valid_empty_document(tmp_path):
    out = convert(tmp_path, _blank)

    assert Document(str(out)).element.body is not None
    assert texts(out) == []


def test_page_with_only_an_image_keeps_the_image_and_warns(tmp_path, caplog):
    def build(doc):
        page = new_page(doc)
        page.insert_image(pymupdf.Rect(72, 100, 372, 400), stream=png_bytes(300, 300))

    with caplog.at_level(logging.WARNING, logger="docx_healer"):
        out = convert(tmp_path, build)

    assert len(Document(str(out)).inline_shapes) == 1
    assert texts(out) == []
    assert "no text" in caplog.text.lower()


def test_two_columns_with_full_width_title_and_footnote_read_in_order(tmp_path):
    left = ["Left column first paragraph begins here and", "runs on to a second line before it ends."]
    left2 = ["Left column second paragraph starts here and", "also has a second line before it closes."]
    right = ["Right column first paragraph begins here and", "runs on to a second line before it ends."]
    right2 = ["Right column second paragraph starts here and", "also has a second line before it closes."]
    title = "A full width title line that spans both of the columns on this page"
    foot = "A full width footnote line that spans both of the columns on this page"

    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 100, title)
        y = lines(page, LEFT, 140, left)
        lines(page, LEFT, y + para_gap(), left2)
        y = lines(page, 330, 140, right)
        lines(page, 330, y + para_gap(), right2)
        put(page, LEFT, 400, foot)

    out = convert(tmp_path, build)

    assert texts(out) == [
        title,
        " ".join(left),
        " ".join(left2),
        " ".join(right),
        " ".join(right2),
        foot,
    ]


def test_three_columns_are_detected_without_a_hardcoded_count(tmp_path):
    cols = [
        [
            f"Column {n} line one runs along the page",
            f"column {n} line two keeps going on",
            f"column {n} line three is here too",
            f"column {n} line four finishes it",
        ]
        for n in ("A", "B", "C")
    ]

    def build(doc):
        page = new_page(doc)
        for index, x in enumerate((50, 230, 410)):
            lines(page, x, 120, cols[index])

    out = convert(tmp_path, build)

    assert texts(out) == [" ".join(col) for col in cols]


def test_line_end_hyphens_join_only_when_the_word_is_safe(tmp_path):
    body = [
        "The trans-",
        "duction model uses well-",
        "known layers for end-to-",
        "end learning with Smith-",
        "Jones as the author.",
    ]

    def build(doc):
        page = new_page(doc)
        y = lines(page, LEFT, 100, body)
        put(page, LEFT, y + 20, "A well-known fact is stated here.")

    out = convert(tmp_path, build)

    assert texts(out)[0] == (
        "The transduction model uses well-known layers for end-to-end learning "
        "with Smith-Jones as the author."
    )


def test_running_header_and_page_numbers_are_dropped_but_body_numbers_stay(tmp_path):
    def build(doc):
        for number in range(1, 5):
            page = new_page(doc)
            put(page, LEFT, 36, "Annual Report 2024")
            put(page, 300, 790, str(number))
            put(page, LEFT, 120, f"Unique body sentence number {number} sits here alone.")
            if number == 2:
                put(page, LEFT, 300, "42")

    out = convert(tmp_path, build)

    assert texts(out) == [
        "Unique body sentence number 1 sits here alone.",
        "Unique body sentence number 2 sits here alone.",
        "42",
        "Unique body sentence number 3 sits here alone.",
        "Unique body sentence number 4 sits here alone.",
    ]


def test_paragraph_continues_across_a_page_break_only_when_the_sentence_is_open(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 735, "This sentence is cut by the page and continues onto the")
        put(page, LEFT, 747, "next page where it")
        page = new_page(doc)
        put(page, LEFT, 110, "finally finishes after the break.")
        put(page, LEFT, 700, "This paragraph ends cleanly at the page end.")
        page = new_page(doc)
        put(page, LEFT, 110, "New paragraph starts on the last page.")

    out = convert(tmp_path, build)

    assert texts(out) == [
        "This sentence is cut by the page and continues onto the next page where it "
        "finally finishes after the break.",
        "This paragraph ends cleanly at the page end.",
        "New paragraph starts on the last page.",
    ]


def test_cjk_lines_join_without_a_space_and_latin_lines_with_one(tmp_path):
    first = ["这是一个很长的中文句子用来测试", "换行时不能插入空格字符"]

    def build(doc):
        page = new_page(doc)
        lines(page, LEFT, 100, first, font="china-s")
        lines(page, LEFT, 200, ["latin line one", "continues here"])

    out = convert(tmp_path, build)

    assert texts(out) == ["".join(first), "latin line one continues here"]


def test_hebrew_paragraph_is_bidi_and_in_logical_order(tmp_path):
    def build(doc):
        page = new_page(doc)
        html_block(page, (72, 100, 500, 200), '<p dir="rtl">שלום עולם זה טקסט בעברית</p>')

    out = convert(tmp_path, build)

    paras = [p for p in paragraphs(out) if p["text"].strip()]
    assert [p["text"] for p in paras] == ["שלום עולם זה טקסט בעברית"]
    assert paras[0]["bidi"]
    assert "<w:rtl" in Document(str(out)).element.xml


def test_ligatures_fold_and_control_characters_are_removed(tmp_path):
    assert clean_text("a\x01b\x00cﬁx") == "abcfix"

    def build(doc):
        page = new_page(doc)
        html_block(page, (72, 100, 500, 200), "<p>The ﬁne print and ﬂow</p>")

    out = convert(tmp_path, build)

    assert texts(out) == ["The fine print and flow"]


def test_a_failing_page_is_logged_and_the_rest_still_written(tmp_path, monkeypatch, caplog):
    original = flow_extract.extract_page

    def flaky(page, *args, **kwargs):
        if page.number == 1:
            raise RuntimeError("boom")
        return original(page, *args, **kwargs)

    monkeypatch.setattr(flow_extract, "extract_page", flaky)

    def build(doc):
        for word in ("first", "second", "third"):
            put(new_page(doc), LEFT, 120, f"Only the {word} page here.")

    with caplog.at_level(logging.WARNING, logger="docx_healer"):
        out = convert(tmp_path, build)

    assert texts(out) == ["Only the first page here.", "Only the third page here."]
    assert "page 2" in caplog.text


@pytest.mark.parametrize("pages,limit,kept", [(5, 3, 3), (3, 3, 3)])
def test_page_limit_keeps_the_first_pages_and_warns_only_when_exceeded(
    tmp_path, caplog, pages, limit, kept
):
    def build(doc):
        for number in range(pages):
            put(new_page(doc), LEFT, 120, f"Page {number + 1} carries one marker sentence.")

    with caplog.at_level(logging.WARNING, logger="docx_healer"):
        out = convert(tmp_path, build, limits=FlowLimits(max_pages=limit))

    assert len(texts(out)) == kept
    assert ("max_pages" in caplog.text) == (pages > limit)
