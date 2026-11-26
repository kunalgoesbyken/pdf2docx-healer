import pymupdf
import pytest
from docx import Document
from docx.oxml.ns import qn

from docx_healer.flow_docx import FontBook
from docx_healer.flow_headings import statistic_headings
from docx_healer.flow_model import Para, Span
from docx_healer.flow_page import align_lines
from docx_healer.font_mapper import font_family, normalize_font_name
from flow_fixtures import LEFT, convert, lines, new_page, paragraphs, problems, put

WIDE = [
    "The quick brown fox jumps over the lazy dog and keeps on running far",
    "away from the farm until the evening sun goes down behind the hills",
]


def _ppr(paragraph, tag):
    ppr = paragraph["el"].find(qn("w:pPr"))
    return None if ppr is None else ppr.find(qn(tag))


def _find(path, start):
    return next(p for p in paragraphs(path) if p["text"].startswith(start))


def _jc(paragraph):
    found = _ppr(paragraph, "w:jc")
    return None if found is None else found.get(qn("w:val"))


@pytest.mark.parametrize(
    "rects, rtl, align",
    [
        ([(72, 0, 300, 10), (72, 12, 300, 22), (72, 24, 200, 34)], False, "both"),
        ([(84, 0, 300, 10), (72, 12, 300, 22), (72, 24, 150, 34)], False, "both"),
        ([(72, 0, 280, 10), (72, 12, 240, 22), (72, 24, 290, 34)], False, "left"),
        ([(100, 0, 272, 10), (130, 12, 242, 22)], False, "center"),
        ([(150, 0, 222, 10)], False, "center"),
        ([(200, 0, 300, 10), (150, 12, 300, 22), (250, 24, 300, 34)], False, "right"),
        ([(90, 0, 200, 10)], False, "left"),
        ([(150, 0, 300, 10)], True, "left"),
        ([(72, 0, 300, 10), (72, 12, 300, 22), (200, 24, 300, 34)], True, "both"),
    ],
)
def test_alignment_is_read_from_the_line_edges(rects, rtl, align):
    assert align_lines(rects, 72, 300, 10, rtl).align == align


def test_indents_come_from_the_column_edges():
    look = align_lines([(90, 0, 280, 10), (78, 12, 280, 22), (78, 24, 200, 34)], 72, 300, 10, False)

    assert (look.align, look.left, look.first, look.right) == ("both", 6, 12, 20)


@pytest.mark.parametrize(
    "name, family",
    [
        ("ABCDEF+NimbusRomNo9L-Medi", "Times New Roman"),
        ("HelveticaWorld-Bold", "Arial"),
        ("CMR10", "Cambria"),
        ("SFTT1000", "Courier New"),
        ("ABCDEF+SourceSerifPro-Regular", "Source Serif Pro"),
        ("ITCFranklinGothicStd-Dem", "ITC Franklin Gothic Std"),
        ("Arial,Bold", "Arial"),
        ("CIDFont+F3", ""),
        ("T1_0", ""),
        ("", ""),
    ],
)
def test_font_names_resolve_to_their_family(name, family):
    assert font_family(name) == family


def test_default_mode_font_fixer_uses_the_same_family():
    assert normalize_font_name("ABCDEF+NimbusSanL-Bold") == "Arial"
    assert normalize_font_name("SourceSerifPro-It") == "Source Serif Pro"


def test_font_table_stays_bounded_for_many_fonts():
    spans = [Span(f"word{i} ", 10, font=f"Family{i:04d}Name-Regular") for i in range(400)]
    book = FontBook([Para(spans, 0, (0, 0, 0), (0, 0, 1, 1), 0, 10, False)], 10)
    doc = Document()

    book.register(doc)

    part = doc.part.part_related_by(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
    )
    assert part.blob.count(b"<w:font ") <= 200


def _two_columns(page, y, count=8):
    for x, side in ((LEFT, "Left"), (330, "Right")):
        lines(page, x, y, [f"{side} column line {i} has a few words" for i in range(count)])


def test_full_width_title_then_two_columns_become_sections(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, 200, 80, "A Centered Title Line", 16, "hebo")
        _two_columns(page, 130)

    out = convert(tmp_path, build)

    body = Document(str(out)).element.body
    cols = [c for c in body.iter(qn("w:cols")) if c.get(qn("w:num")) == "2"]
    types = [t.get(qn("w:val")) for t in body.iter(qn("w:type"))]
    assert cols and "continuous" in types
    assert _jc(_find(out, "A Centered Title")) == "center"
    assert not problems(out)


def test_colour_size_and_font_follow_the_pdf(tmp_path):
    def build(doc):
        page = new_page(doc)
        lines(page, LEFT, 100, WIDE)
        page.insert_text((LEFT, 160), "Blue words in times", fontsize=13, fontname="tiro",
                         color=(0, 0, 1))

    out = convert(tmp_path, build)

    run = _find(out, "Blue words")["el"].find(f"{qn('w:r')}/{qn('w:rPr')}")
    assert run.find(qn("w:color")).get(qn("w:val")) == "0000FF"
    assert run.find(qn("w:sz")).get(qn("w:val")) == "26"
    assert run.find(qn("w:rFonts")).get(qn("w:ascii")) == "Times New Roman"


def test_heading_styles_carry_no_theme_look(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 100, "Section Heading", 14, "hebo")
        for y in (130, 180, 230):
            lines(page, LEFT, y, WIDE)

    out = convert(tmp_path, build)

    styles = Document(str(out)).styles
    for name in ("Heading 1", "Title"):
        rpr = styles[name].element.find(qn("w:rPr"))
        assert rpr is None or len(rpr) == 0


def _para(text, top, size, bold, page=0):
    box = (150.0, top, 450.0, top + size)
    return Para(
        [Span(text, size, bold=bold)], page, (page, 0, -1), box, 150.0, size, bold,
        boxes=[((page, 0, -1), box, top + size * 0.8)],
    )


def test_tight_block_of_large_bold_lines_is_not_a_heading():
    authors = [
        _para("Ann Author*, Ben Writer*, Cat Person", 114, 12, False),
        _para("Eve Jones, Fay Brown, Gus Green", 130, 12, True),
        _para("Jo Grey, Kim Stone, Lee Wood", 144, 12, True),
        _para("Max Field*, Ned Hill*", 157, 12, False),
    ]
    heading = _para("Introduction", 300, 12, True)
    body = [_para(f"Body text paragraph {i}", 330 + 40 * i, 10, False) for i in range(6)]

    statistic_headings(authors + [heading] + body, 10)

    assert [p.kind for p in authors] == ["para"] * 4
    assert heading.kind == "heading"


def test_each_pdf_page_starts_a_word_page(tmp_path):
    def build(doc):
        for word in ("First", "Second"):
            page = new_page(doc)
            lines(page, LEFT, 100, [f"{word} page text ends with a full stop."])

    out = convert(tmp_path, build)

    second = _find(out, "Second page")
    body = Document(str(out)).element.body
    breaks = _ppr(second, "w:pageBreakBefore") is not None
    sections = len(list(body.iter(qn("w:sectPr")))) > 1
    assert breaks or sections


def test_white_text_on_a_dark_box_keeps_its_background(tmp_path):
    def build(doc):
        page = new_page(doc)
        lines(page, LEFT, 100, WIDE)
        page.draw_rect(pymupdf.Rect(LEFT - 10, 180, 400, 220), color=None, fill=(0, 0, 128 / 255))
        page.insert_text((LEFT, 205), "White words on navy", fontsize=12, color=(1, 1, 1))
        page.insert_text((LEFT, 300), "White words on paper", fontsize=12, color=(1, 1, 1))

    out = convert(tmp_path, build)

    shading = _ppr(_find(out, "White words on navy"), "w:shd")
    assert shading is not None and shading.get(qn("w:fill")) == "000080"
    bare = _find(out, "White words on paper")["el"].find(f".//{qn('w:color')}")
    assert bare is None or bare.get(qn("w:val")) != "FFFFFF"


def test_text_up_the_margin_is_pinned_in_a_rotated_frame(tmp_path):
    def build(doc):
        page = new_page(doc)
        for y in (100, 150, 200):
            lines(page, LEFT, y, WIDE)
        page.insert_text((30, 500), "Sideways margin note", fontsize=14, rotate=90)

    out = convert(tmp_path, build)

    note = _find(out, "Sideways margin")
    frame = _ppr(note, "w:framePr")
    assert frame is not None and frame.get(qn("w:wrap")) == "none"
    assert _ppr(note, "w:textDirection").get(qn("w:val")) == "btLr"
    assert not problems(out)


def test_block_beside_earlier_text_is_pinned_where_it_is(tmp_path):
    def build(doc):
        page = new_page(doc)
        lines(page, LEFT, 100, ["Name of the person filling the form", "Street and city"])
        put(page, 420, 100, "Office use only")
        for y in (200, 250):
            lines(page, LEFT, y, WIDE)

    out = convert(tmp_path, build)

    pinned = [p for p in paragraphs(out) if _ppr(p, "w:framePr") is not None]
    assert [p["text"].strip() for p in pinned] == ["Office use only"]
    assert not problems(out)


def test_space_before_and_line_pitch_are_exact(tmp_path):
    def build(doc):
        page = new_page(doc)
        lines(page, LEFT, 100, WIDE, lead=14)
        lines(page, LEFT, 160, WIDE, lead=14)

    out = convert(tmp_path, build)

    spacing = _ppr(_find(out, WIDE[0]), "w:spacing")
    assert spacing.get(qn("w:lineRule")) == "exact"
    assert spacing.get(qn("w:line")) == "280"
    later = [p for p in paragraphs(out) if p["text"].startswith(WIDE[0])][1]
    assert int(_ppr(later, "w:spacing").get(qn("w:before"))) > 0


def test_centred_lines_keep_their_breaks(tmp_path):
    first = "A centred block of several words that runs wide"
    second = "and a shorter second line"

    def build(doc):
        page = new_page(doc)
        for y, text in ((100, first), (114, second)):
            width = pymupdf.get_text_length(text, "helv", 12)
            put(page, (page.rect.width - width) / 2, y, text, 12)
        for y in (170, 220):
            lines(page, LEFT, y, WIDE)

    out = convert(tmp_path, build)

    block = _find(out, first)
    assert _jc(block) == "center"
    assert len(block["el"].findall(f".//{qn('w:br')}")) == 1
