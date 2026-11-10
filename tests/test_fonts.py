import pytest
from docx import Document
from docx.oxml.ns import qn

from docx_healer.font_mapper import fix_docx_fonts, normalize_font_name
from docx_healer.link_injector import fix_docx_hyperlinks
from factories import make_docx

ARIAL = ("Arial", "Arial", "Arial")


@pytest.mark.parametrize(
    ("name", "text", "expected"),
    [
        ("SimSun", "", "SimSun"),
        ("ArialMT", "", "Arial"),
        ("arialmt", "", "Arial"),
        ("Helvetica", "", "Arial"),
        ("ABCDEF+TimesNewRomanPSMT", "", "Times New Roman"),
        ("Calibri", "plain latin text", "Calibri"),
        ("ABCDEF+Calibri", "plain latin text", "Calibri"),
        ("C++ Mono", "", "C++ Mono"),
        ("Calibri", "中文", "SimSun"),
        ("Calibri", "документ", "Calibri"),
        ("Verdana", "שלום", "David"),
        ("", "anything", "Arial"),
        (None, "anything", "Arial"),
    ],
)
def test_normalize_font_name(name, text, expected):
    assert normalize_font_name(name, text) == expected


def set_fonts(run, name):
    run.font.name = name
    run.font._element.rPr.rFonts.set(qn("w:eastAsia"), name)


def font_triple(run_element):
    rfonts = run_element.find(qn("w:rPr")).find(qn("w:rFonts"))
    return tuple(rfonts.get(qn(key)) for key in ("w:ascii", "w:hAnsi", "w:eastAsia"))


def test_unknown_fonts_are_preserved_and_pdf_names_are_mapped(tmp_path):
    def build(doc):
        paragraph = doc.add_paragraph()
        set_fonts(paragraph.add_run("kept"), "Calibri")
        set_fonts(paragraph.add_run("mapped"), "MS-Mincho")

    path = make_docx(tmp_path / "a.docx", build)

    assert fix_docx_fonts(path) == 1

    runs = list(Document(str(path)).paragraphs[0]._p.iter(qn("w:r")))
    assert font_triple(runs[0]) == ("Calibri", "Calibri", "Calibri")
    assert font_triple(runs[1]) == ("MS Mincho", "MS Mincho", "MS Mincho")


def test_fonts_are_fixed_in_tables_headers_and_footers(tmp_path):
    def build(doc):
        table = doc.add_table(rows=1, cols=1)
        set_fonts(table.cell(0, 0).paragraphs[0].add_run("cell"), "ArialMT")
        set_fonts(doc.sections[0].header.paragraphs[0].add_run("head"), "ArialMT")
        set_fonts(doc.sections[0].footer.paragraphs[0].add_run("foot"), "ArialMT")

    path = make_docx(tmp_path / "a.docx", build)

    assert fix_docx_fonts(path) == 3

    doc = Document(str(path))
    section = doc.sections[0]
    paragraphs = [
        doc.tables[0].cell(0, 0).paragraphs[0],
        section.header.paragraphs[0],
        section.footer.paragraphs[0],
    ]
    assert [font_triple(p._p.find(qn("w:r"))) for p in paragraphs] == [ARIAL] * 3


def test_hyperlink_runs_get_the_same_font_treatment(tmp_path):
    def build(doc):
        set_fonts(doc.add_paragraph().add_run("see https://example.com"), "Helvetica")

    path = make_docx(tmp_path / "a.docx", build)
    fix_docx_hyperlinks(path)

    fix_docx_fonts(path)

    paragraph = Document(str(path)).paragraphs[0]
    assert {font_triple(r) for r in paragraph._p.iter(qn("w:r"))} == {ARIAL}
