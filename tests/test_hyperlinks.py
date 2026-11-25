import zipfile

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor

from docx_healer.link_injector import fix_docx_hyperlinks
from docx_inspect import body_hyperlinks, hyperlinks, paragraph_text, run_text
from factories import make_docx


def test_url_in_body_is_linked_without_trailing_punctuation(tmp_path):
    text = "See https://example.com/docs, or www.example.org."
    path = make_docx(tmp_path / "a.docx", lambda doc: doc.add_paragraph(text))

    assert fix_docx_hyperlinks(path) == 2

    assert body_hyperlinks(path) == [
        ("https://example.com/docs", "https://example.com/docs"),
        ("www.example.org", "https://www.example.org"),
    ]
    assert paragraph_text(Document(str(path)).paragraphs[0]) == text


def test_mailto_and_uppercase_schemes_are_not_double_prefixed(tmp_path):
    path = make_docx(
        tmp_path / "a.docx",
        lambda doc: doc.add_paragraph("Write MAILTO:team@example.com or HTTPS://Example.com/x"),
    )

    fix_docx_hyperlinks(path)

    assert [target for _, target in body_hyperlinks(path)] == [
        "MAILTO:team@example.com",
        "HTTPS://Example.com/x",
    ]


def test_urls_in_table_cells_and_nested_tables_are_linked(tmp_path):
    def build(doc):
        outer = doc.add_table(rows=1, cols=2)
        outer.cell(0, 0).paragraphs[0].text = "Docs: https://example.com/outer"
        nested = outer.cell(0, 1).add_table(rows=1, cols=1)
        nested.cell(0, 0).paragraphs[0].text = "Nested https://example.com/inner"

    path = make_docx(tmp_path / "a.docx", build)

    assert fix_docx_hyperlinks(path) == 2

    assert sorted(target for _, target in body_hyperlinks(path)) == [
        "https://example.com/inner",
        "https://example.com/outer",
    ]


def test_merged_cells_are_not_processed_twice(tmp_path):
    def build(doc):
        table = doc.add_table(rows=1, cols=2)
        merged = table.cell(0, 0).merge(table.cell(0, 1))
        merged.paragraphs[0].text = "https://example.com/merged"

    path = make_docx(tmp_path / "a.docx", build)

    assert fix_docx_hyperlinks(path) == 1
    assert len(body_hyperlinks(path)) == 1


def test_urls_in_header_and_footer_are_linked_on_their_own_part(tmp_path):
    def build(doc):
        doc.add_paragraph("body")
        section = doc.sections[0]
        section.header.paragraphs[0].text = "Header https://example.com/header"
        section.footer.paragraphs[0].text = "Footer https://example.com/footer"

    path = make_docx(tmp_path / "a.docx", build)

    assert fix_docx_hyperlinks(path) == 2

    section = Document(str(path)).sections[0]
    assert hyperlinks(section.header, section.header._element) == [
        ("https://example.com/header", "https://example.com/header")
    ]
    assert hyperlinks(section.footer, section.footer._element) == [
        ("https://example.com/footer", "https://example.com/footer")
    ]


def test_documents_without_headers_do_not_gain_header_parts(tmp_path):
    path = make_docx(tmp_path / "a.docx", lambda doc: doc.add_paragraph("see https://example.com"))

    fix_docx_hyperlinks(path)

    with zipfile.ZipFile(path) as archive:
        assert not [n for n in archive.namelist() if "header" in n or "footer" in n]


def test_run_formatting_is_kept_and_link_is_blue_and_underlined(tmp_path):
    def build(doc):
        run = doc.add_paragraph().add_run("go to https://example.com now")
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)

    path = make_docx(tmp_path / "a.docx", build)

    fix_docx_hyperlinks(path)

    paragraph = Document(str(path)).paragraphs[0]
    runs = list(paragraph._p.iter(qn("w:r")))
    assert [run_text(r) for r in runs] == ["go to ", "https://example.com", " now"]
    assert all(r.find(qn("w:rPr")).find(qn("w:b")) is not None for r in runs)
    link_properties = runs[1].find(qn("w:rPr"))
    assert link_properties.find(qn("w:u")).get(qn("w:val")) == "single"
    assert link_properties.find(qn("w:color")).get(qn("w:val")) == "0563C1"


def test_runs_with_non_text_content_are_left_alone(tmp_path):
    def build(doc):
        run = doc.add_paragraph().add_run("https://example.com")
        page_break = OxmlElement("w:br")
        page_break.set(qn("w:type"), "page")
        run._r.append(page_break)

    path = make_docx(tmp_path / "a.docx", build)

    assert fix_docx_hyperlinks(path) == 0

    paragraph = Document(str(path)).paragraphs[0]
    assert paragraph._p.find(".//" + qn("w:br")) is not None


def test_pass_is_idempotent_and_skips_existing_hyperlinks(tmp_path):
    path = make_docx(tmp_path / "a.docx", lambda doc: doc.add_paragraph("https://example.com/a"))

    assert fix_docx_hyperlinks(path) == 1
    assert fix_docx_hyperlinks(path) == 0
    assert len(body_hyperlinks(path)) == 1


def test_file_is_untouched_when_nothing_matches(tmp_path):
    path = make_docx(tmp_path / "a.docx", lambda doc: doc.add_paragraph("no links here"))
    before = path.read_bytes()

    assert fix_docx_hyperlinks(path) == 0
    assert path.read_bytes() == before


def _linked(tmp_path, text):
    path = make_docx(tmp_path / "a.docx", lambda doc: doc.add_paragraph(text))
    count = fix_docx_hyperlinks(path)
    return count, body_hyperlinks(path)


@pytest.mark.parametrize(
    "text, expected",
    [
        ("contact ops@example.com now", [("ops@example.com", "mailto:ops@example.com")]),
        ("mail ops@example.com.", [("ops@example.com", "mailto:ops@example.com")]),
        ("(ops@example.com), or", [("ops@example.com", "mailto:ops@example.com")]),
        ("<ops@example.com>", [("ops@example.com", "mailto:ops@example.com")]),
        ("Ops@Example.COM", [("Ops@Example.COM", "mailto:Ops@Example.COM")]),
        ("a.b+tag@mail.example.co.uk!", [("a.b+tag@mail.example.co.uk", "mailto:a.b+tag@mail.example.co.uk")]),
        (
            "one@a.org and two@b.org",
            [("one@a.org", "mailto:one@a.org"), ("two@b.org", "mailto:two@b.org")],
        ),
    ],
)
def test_plain_email_addresses_become_mailto_links(tmp_path, text, expected):
    count, found = _linked(tmp_path, text)

    assert found == expected
    assert count == len(expected)


@pytest.mark.parametrize(
    "text",
    [
        "ping @handle about it",
        "user@localhost is a dev host",
        "install pkg@1.2.3 first",
        "git clone git@github.com:org/repo.git",
        "price @ 5.00 each",
        "name@ example.com",
        "x@y.z",
        "@scope/pkg@latest",
        "a@b@c.com",
        "trailing dots ops..@example.com",
    ],
)
def test_text_that_only_looks_like_an_email_is_not_linked(tmp_path, text):
    count, found = _linked(tmp_path, text)

    assert (count, found) == (0, [])


def test_email_inside_a_url_is_not_linked_twice(tmp_path):
    count, found = _linked(tmp_path, "see https://example.com/u/ops@example.com and https://user@host.example.com/x")

    assert [target for _, target in found] == [
        "https://example.com/u/ops@example.com",
        "https://user@host.example.com/x",
    ]
    assert count == 2


def test_explicit_mailto_is_linked_once(tmp_path):
    count, found = _linked(tmp_path, "write mailto:ops@example.com today")

    assert found == [("mailto:ops@example.com", "mailto:ops@example.com")]
    assert count == 1


def test_email_in_a_table_cell_is_linked(tmp_path):
    def build(doc):
        table = doc.add_table(rows=1, cols=1)
        table.cell(0, 0).paragraphs[0].text = "ops@example.com"

    path = make_docx(tmp_path / "a.docx", build)

    assert fix_docx_hyperlinks(path) == 1
    assert body_hyperlinks(path) == [("ops@example.com", "mailto:ops@example.com")]
