import time

import pytest
from docx import Document
from docx.oxml.ns import qn

from docx_healer.headings import fix_docx_headings
from docx_healer.ocr_pipeline import pymupdf
from factories import make_docx
from paged import body_styles, effective_formatting, make_paged_docx, make_toc_pdf

H1, H2, H3 = "Heading1", "Heading2", "Heading3"


def styled(path):
    return {text: style for text, style in body_styles(path) if style}


@pytest.mark.parametrize(
    "docx_text, title",
    [
        ("INTRODUCTION", "Introduction"),
        ("Intro   duction", "Introduction"),
        ("MultiHead Attention", "Multi-Head Attention"),
        ("1 Introduction", "Introduction"),
        ("3.2.1 Scaled Dot-Product Attention", "Scaled Dot-Product Attention"),
        ("A.1 Proof of the lemma", "Proof of the lemma"),
        ("2. 概要と目的", "概要と目的"),
        ("1. مقدمة", "مقدمة"),
        ("ﬁnal remarks", "final remarks"),
    ],
)
def test_title_matches_despite_case_spacing_hyphens_and_numbering(tmp_path, docx_text, title):
    pdf = make_toc_pdf(tmp_path / "a.pdf", [[1, title, 1]], 1)
    path = make_paged_docx(tmp_path / "a.docx", [["lead in text", docx_text, "tail text"]])

    assert fix_docx_headings(path, pdf) == 1

    assert styled(path) == {docx_text: H1}


@pytest.mark.parametrize(
    "docx_text",
    ["Figure 3 Introduction", "Introduction to the model", "see the Introduction", "Intro"],
)
def test_text_that_only_contains_the_title_is_not_a_heading(tmp_path, docx_text):
    pdf = make_toc_pdf(tmp_path / "a.pdf", [[1, "Introduction", 1]], 1)
    path = make_paged_docx(tmp_path / "a.docx", [[docx_text]])
    before = path.read_bytes()

    assert fix_docx_headings(path, pdf) == 0
    assert path.read_bytes() == before


def test_levels_follow_the_outline(tmp_path):
    toc = [[1, "Model", 1], [2, "Encoder", 1], [3, "Layers", 2], [1, "Results", 2]]
    pdf = make_toc_pdf(tmp_path / "a.pdf", toc, 2)
    path = make_paged_docx(
        tmp_path / "a.docx", [["Model", "Encoder"], ["Layers", "Results", "body"]]
    )

    assert fix_docx_headings(path, pdf) == 4

    assert styled(path) == {"Model": H1, "Encoder": H2, "Layers": H3, "Results": H1}


def test_duplicate_titles_are_matched_by_page_not_by_first_hit(tmp_path):
    pdf = make_toc_pdf(tmp_path / "a.pdf", [[1, "Introduction", 1], [2, "Introduction", 3]], 3)
    path = make_paged_docx(
        tmp_path / "a.docx",
        [["Introduction", "text"], ["Introduction"], ["lead", "Introduction", "text"]],
    )

    assert fix_docx_headings(path, pdf) == 2

    paragraphs = body_styles(path)
    assert [style for text, style in paragraphs if text == "Introduction"] == [H1, None, H2]


def test_title_on_another_page_than_the_outline_says_is_left_alone(tmp_path):
    pdf = make_toc_pdf(tmp_path / "a.pdf", [[1, "Results", 2]], 2)
    path = make_paged_docx(tmp_path / "a.docx", [["Results", "text"], ["other text"]])

    assert fix_docx_headings(path, pdf) == 0
    assert styled(path) == {}


def test_title_split_over_paragraphs_styles_every_piece(tmp_path):
    toc = [[1, "Model", 1], [2, "Scaled Dot-Product Attention", 1]]
    pdf = make_toc_pdf(tmp_path / "a.pdf", toc, 1)
    path = make_paged_docx(
        tmp_path / "a.docx", [["Model", "Scaled Dot-Product", "Attention", "body"]]
    )

    assert fix_docx_headings(path, pdf) == 2

    assert styled(path) == {"Model": H1, "Scaled Dot-Product": H2, "Attention": H2}


def test_heading_inside_a_table_is_never_styled(tmp_path):
    pdf = make_toc_pdf(tmp_path / "a.pdf", [[1, "Results", 1]], 1)
    path = make_paged_docx(tmp_path / "a.docx", [[("tbl", "Results"), "body"]])
    before = path.read_bytes()

    assert fix_docx_headings(path, pdf) == 0

    assert path.read_bytes() == before


def test_headers_and_footers_are_never_styled(tmp_path):
    pdf = make_toc_pdf(tmp_path / "a.pdf", [[1, "Results", 1]], 1)
    path = make_paged_docx(tmp_path / "a.docx", [["body"]])
    doc = Document(str(path))
    doc.sections[0].header.paragraphs[0].text = "Results"
    doc.save(str(path))

    assert fix_docx_headings(path, pdf) == 0

    header = Document(str(path)).sections[0].header.paragraphs[0]
    assert header.style.name != "Heading 1"


def test_paragraphs_with_their_own_style_are_skipped(tmp_path):
    pdf = make_toc_pdf(tmp_path / "a.pdf", [[1, "Results", 1]], 1)
    path = make_docx(
        tmp_path / "a.docx", lambda doc: doc.add_paragraph("Results", style="List Bullet")
    )

    assert fix_docx_headings(path, pdf) == 0
    assert styled(path) == {"Results": "ListBullet"}


def test_pass_is_idempotent(tmp_path):
    pdf = make_toc_pdf(tmp_path / "a.pdf", [[1, "Results", 1]], 1)
    path = make_paged_docx(tmp_path / "a.docx", [["Results", "text"]])

    assert fix_docx_headings(path, pdf) == 1
    assert fix_docx_headings(path, pdf) == 0
    assert styled(path) == {"Results": H1}


def test_without_page_sections_only_unambiguous_titles_are_styled(tmp_path):
    toc = [[1, "Introduction", 1], [1, "Methods", 1], [1, "Introduction", 2]]
    pdf = make_toc_pdf(tmp_path / "a.pdf", toc, 2)

    def build(doc):
        for text in ("Introduction", "text", "Methods", "Introduction", "text"):
            doc.add_paragraph(text)

    path = make_docx(tmp_path / "a.docx", build)

    assert fix_docx_headings(path, pdf) == 1

    assert styled(path) == {"Methods": H1}


def test_styling_does_not_change_how_any_paragraph_looks(tmp_path):
    toc = [[1, "Introduction", 1], [2, "Background", 1]]
    pdf = make_toc_pdf(tmp_path / "a.pdf", toc, 1)
    path = make_paged_docx(tmp_path / "a.docx", [["Introduction", "Background", "plain body"]])

    def snapshot():
        doc = Document(str(path))
        return [
            effective_formatting(doc, p, r)
            for p in doc.element.body.iterchildren(qn("w:p"))
            for r in p.iter(qn("w:r"))
        ]

    before = snapshot()

    assert fix_docx_headings(path, pdf) == 2

    assert styled(path) == {"Introduction": H1, "Background": H2}
    assert snapshot() == before


def test_heading_styles_carry_outline_levels_and_are_created_when_missing(tmp_path):
    pdf = make_toc_pdf(tmp_path / "a.pdf", [[1, "Results", 1]], 1)
    path = make_paged_docx(tmp_path / "a.docx", [["Results"]])
    doc = Document(str(path))
    styles = doc.styles.element
    for node in list(styles.iterchildren(qn("w:style"))):
        if node.get(qn("w:styleId")) == "Heading1":
            styles.remove(node)
    doc.save(str(path))

    assert fix_docx_headings(path, pdf) == 1

    styles = Document(str(path)).styles.element
    node = [s for s in styles.iterchildren(qn("w:style")) if s.get(qn("w:styleId")) == "Heading1"][
        0
    ]
    assert node.find(qn("w:name")).get(qn("w:val")).lower() == "heading 1"
    assert node.find(qn("w:pPr") + "/" + qn("w:outlineLvl")).get(qn("w:val")) == "0"


def test_many_outline_entries_are_handled_in_linear_time(tmp_path):
    count = 3000
    toc = [[1, f"Topic {index}", 1 + index % 3] for index in range(count)]
    pdf = make_toc_pdf(tmp_path / "a.pdf", toc, 3)
    pages = [[f"Topic {index}" for index in range(page, count, 3)] for page in range(3)]
    path = make_paged_docx(tmp_path / "a.docx", pages)

    start = time.perf_counter()
    assert fix_docx_headings(path, pdf) == count
    assert time.perf_counter() - start < 60


def build_font_pdf(path):
    doc = pymupdf.open()
    for number in range(3):
        doc.new_page()
        page = doc[number]
        page.insert_text((72, 40), "Running header text", fontsize=9)
        y = 90
        if number == 0:
            page.insert_text((72, y), "Big Document Title", fontsize=22)
            y += 40
        page.insert_text((72, y), f"Chapter {number + 1} Overview", fontsize=16)
        y += 30
        for line in range(8):
            page.insert_text(
                (72, y), f"This is ordinary body text line {line} of page {number}.", fontsize=10
            )
            y += 14
        page.insert_text((72, y + 10), f"Part {number + 1} Details", fontsize=13)
        y += 40
        for line in range(8):
            page.insert_text(
                (72, y), f"More ordinary body text line {line} for the page {number}.", fontsize=10
            )
            y += 14
    doc.save(str(path))
    doc.close()
    return path


def font_docx(path):
    pages = []
    for number in range(3):
        items = ["Running header text"]
        if number == 0:
            items.append("Big Document Title")
        items.append(f"Chapter {number + 1} Overview")
        items += [f"This is ordinary body text line {line} of page {number}." for line in range(8)]
        items.append(f"Part {number + 1} Details")
        items += [
            f"More ordinary body text line {line} for the page {number}." for line in range(8)
        ]
        pages.append(items)
    return make_paged_docx(path, pages)


def test_without_an_outline_larger_short_lines_become_headings_by_size_rank(tmp_path):
    pdf = build_font_pdf(tmp_path / "a.pdf")
    path = font_docx(tmp_path / "a.docx")

    assert fix_docx_headings(path, pdf) == 7

    assert styled(path) == {
        "Big Document Title": H1,
        "Chapter 1 Overview": H2,
        "Chapter 2 Overview": H2,
        "Chapter 3 Overview": H2,
        "Part 1 Details": H3,
        "Part 2 Details": H3,
        "Part 3 Details": H3,
    }


def test_without_an_outline_and_without_larger_text_nothing_is_styled(tmp_path):
    pdf = tmp_path / "a.pdf"
    doc = pymupdf.open()
    doc.new_page()
    for line in range(10):
        doc[0].insert_text((72, 100 + 14 * line), f"Uniform body text on line {line}.", fontsize=11)
    doc.save(str(pdf))
    doc.close()
    path = make_paged_docx(
        tmp_path / "a.docx", [[f"Uniform body text on line {line}." for line in range(10)]]
    )
    before = path.read_bytes()

    assert fix_docx_headings(path, pdf) == 0
    assert path.read_bytes() == before


def test_without_an_outline_text_that_is_mostly_large_has_no_headings(tmp_path):
    pdf = tmp_path / "a.pdf"
    doc = pymupdf.open()
    doc.new_page()
    for line in range(2):
        doc[0].insert_text(
            (72, 100 + 14 * line), f"Small line {line} of the page text.", fontsize=8
        )
    for line in range(6):
        doc[0].insert_text(
            (72, 200 + 30 * line), f"Large line {line} of the page text.", fontsize=20
        )
    doc.save(str(pdf))
    doc.close()
    items = [f"Small line {n} of the page text." for n in range(2)] + [
        f"Large line {n} of the page text." for n in range(6)
    ]
    path = make_paged_docx(tmp_path / "a.docx", [items])

    assert fix_docx_headings(path, pdf) == 0
