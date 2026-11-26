import time

import pytest

from bench.facts import InvalidDocx, docx_facts
from bench_docx_builder import (
    abstract,
    build_docx,
    lvl,
    num,
    numbering_xml,
    para,
    start_override,
    style,
    styles_xml,
)


def bookmark(book_id, name, text="x"):
    return (
        f'<w:bookmarkStart w:id="{book_id}" w:name="{name}"/>'
        f'<w:r><w:t>{text}</w:t></w:r><w:bookmarkEnd w:id="{book_id}"/>'
    )


def test_empty_docx_is_valid_with_zero_everything(tmp_path):
    facts = docx_facts(build_docx(tmp_path / "a.docx"))
    assert facts.words == []
    assert facts.paragraphs == []
    assert (facts.tables, facts.images, facts.list_paras) == (0, 0, 0)
    assert facts.ext_links == [] and facts.anchors == []


def test_text_in_tables_and_header_is_counted(tmp_path):
    body = (
        "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>alpha beta</w:t></w:r></w:p></w:tc>"
        "<w:tc><w:p><w:r><w:t>gamma</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
    )
    header = "<w:p><w:r><w:t>running header</w:t></w:r></w:p>"
    facts = docx_facts(build_docx(tmp_path / "a.docx", body, header=header))
    assert facts.words == ["alpha", "beta", "gamma", "running", "header"]
    assert facts.tables == 1
    assert facts.paragraphs[0].in_table is True


def test_text_only_in_header_counts(tmp_path):
    header = "<w:p><w:r><w:t>only here</w:t></w:r></w:p>"
    assert docx_facts(build_docx(tmp_path / "a.docx", "", header=header)).words == ["only", "here"]


def test_alternate_content_fallback_text_is_counted_once(tmp_path):
    body = (
        '<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:t>boxed text</w:t></mc:Choice>'
        "<mc:Fallback><w:t>boxed text</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>"
    )
    facts = docx_facts(build_docx(tmp_path / "a.docx", body))
    assert facts.words == ["boxed", "text"]


def test_runs_in_one_paragraph_join_without_a_space_and_tabs_split(tmp_path):
    body = (
        "<w:p><w:r><w:t>foo</w:t></w:r><w:r><w:t>bar</w:t></w:r><w:r><w:tab/></w:r>"
        "<w:r><w:t>baz</w:t></w:r></w:p>"
    )
    assert docx_facts(build_docx(tmp_path / "a.docx", body)).words == ["foobar", "baz"]


def test_list_numbering_start_three_and_override_restart(tmp_path):
    numbering = numbering_xml(
        abstract(0, lvl(0, start=3)), num(1, 0) + num(2, 0, start_override(0, 1))
    )
    body = para("a", 1) + para("b", 1) + para("c", 2) + para("plain")
    facts = docx_facts(build_docx(tmp_path / "a.docx", body, numbering=numbering))
    assert [p.label for p in facts.paragraphs] == ["3.", "4.", "1.", None]
    assert facts.list_paras == 3 and facts.list_levels == [0]


def test_numid_zero_and_dangling_numid_are_not_lists(tmp_path):
    numbering = numbering_xml(abstract(0, lvl(0)), num(1, 0))
    body = para("a", 0) + para("b", 99) + para("c", 1)
    facts = docx_facts(build_docx(tmp_path / "a.docx", body, numbering=numbering))
    assert [p.numbered for p in facts.paragraphs] == [False, False, True]
    assert facts.unresolved_numbering == 1


def test_list_through_paragraph_style_numbering(tmp_path):
    numbering = numbering_xml(abstract(0, lvl(0, fmt="bullet", text="o")), num(4, 0))
    styles = styles_xml(
        style("ListBullet", "List Bullet", ppr='<w:numPr><w:numId w:val="4"/></w:numPr>')
    )
    facts = docx_facts(
        build_docx(
            tmp_path / "a.docx", para("x", style="ListBullet"), numbering=numbering, styles=styles
        )
    )
    assert facts.paragraphs[0].numbered is True
    assert facts.paragraphs[0].kind == "bullet"


def test_headings_by_style_name_outline_level_and_cycle(tmp_path):
    styles = styles_xml(
        style("Heading2", "heading 2")
        + style("Custom", "My Title", ppr='<w:outlineLvl w:val="2"/>')
        + style("Loop", "Loop", based_on="Loop")
        + style("Child", "Child", based_on="Heading2")
    )
    body = (
        para("one", style="Heading2")
        + para("two", style="Custom")
        + para("three", style="Loop")
        + para("four", style="Child")
        + para("five", extra_ppr='<w:outlineLvl w:val="0"/>')
        + para("body", extra_ppr='<w:outlineLvl w:val="9"/>')
    )
    facts = docx_facts(build_docx(tmp_path / "a.docx", body, styles=styles))
    assert [p.heading for p in facts.paragraphs] == [2, 3, None, 2, 1, None]


def test_hyperlinks_external_internal_and_fields(tmp_path):
    rels = [("rId5", "hyperlink", "https://example.com/a", True)]
    body = (
        '<w:p><w:hyperlink r:id="rId5"><w:r><w:t>ext</w:t></w:r></w:hyperlink>'
        '<w:hyperlink r:id="rId99"><w:r><w:t>broken</w:t></w:r></w:hyperlink>'
        '<w:hyperlink w:anchor="target"><w:r><w:t>int</w:t></w:r></w:hyperlink>'
        '<w:hyperlink w:anchor="missing"><w:r><w:t>gone</w:t></w:r></w:hyperlink>'
        "<w:fldSimple w:instr=' HYPERLINK \"https://example.org/f\" '><w:r><w:t>fld</w:t></w:r></w:fldSimple>"
        "<w:fldSimple w:instr=' HYPERLINK \\l \"target\" '><w:r><w:t>fld2</w:t></w:r></w:fldSimple>"
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.net/c" </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>cplx</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
        + "<w:p>"
        + bookmark(1, "target")
        + "</w:p>"
    )
    facts = docx_facts(build_docx(tmp_path / "a.docx", body, rels=rels))
    assert facts.ext_links == [
        "https://example.com/a",
        "https://example.org/f",
        "https://example.net/c",
    ]
    assert facts.anchors == ["target", "missing", "target"]
    assert facts.bookmarks == ["target"]
    assert facts.resolved_anchors == 2 and facts.dangling_anchors == 1
    assert facts.broken_links == 1


def test_bookmark_without_end_is_not_a_bookmark(tmp_path):
    body = (
        '<w:p><w:bookmarkStart w:id="1" w:name="open"/><w:r><w:t>x</w:t></w:r></w:p>'
        '<w:p><w:hyperlink w:anchor="open"><w:r><w:t>l</w:t></w:r></w:hyperlink></w:p>'
    )
    facts = docx_facts(build_docx(tmp_path / "a.docx", body))
    assert facts.bookmarks == [] and facts.dangling_anchors == 1


def test_duplicate_bookmark_names_are_valid_first_wins(tmp_path):
    body = "<w:p>" + bookmark(1, "same") + bookmark(2, "same") + "</w:p>"
    facts = docx_facts(build_docx(tmp_path / "a.docx", body))
    assert facts.bookmarks == ["same"] and facts.bookmark_dup_names == 1


def test_duplicate_bookmark_ids_are_invalid(tmp_path):
    body = "<w:p>" + bookmark(1, "a") + bookmark(1, "b") + "</w:p>"
    with pytest.raises(InvalidDocx):
        docx_facts(build_docx(tmp_path / "a.docx", body))


def test_truncated_zip_is_invalid(tmp_path):
    path = build_docx(tmp_path / "a.docx", para("x"))
    path.write_bytes(path.read_bytes()[:60])
    with pytest.raises(InvalidDocx):
        docx_facts(path)


def test_missing_file_is_invalid(tmp_path):
    with pytest.raises(InvalidDocx):
        docx_facts(tmp_path / "nope.docx")


def test_images_counts_drawings_but_not_fallback_copies(tmp_path):
    body = (
        '<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><a:blip r:embed="x"/></w:drawing>'
        '</mc:Choice><mc:Fallback><w:pict><v:imagedata r:id="y"/></w:pict></mc:Fallback>'
        '</mc:AlternateContent></w:r><w:r><w:drawing><a:blip r:embed="z"/></w:drawing></w:r></w:p>'
    )
    assert docx_facts(build_docx(tmp_path / "a.docx", body)).images == 2


def test_header_pictures_count_once_without_fallback_copies(tmp_path):
    header = (
        '<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><a:blip r:embed="x"/>'
        '</w:drawing></mc:Choice><mc:Fallback><w:pict><v:imagedata r:id="y"/></w:pict></mc:Fallback>'
        "</mc:AlternateContent></w:r></w:p>"
    )
    body = '<w:p><w:r><w:drawing><a:blip r:embed="z"/></w:drawing></w:r></w:p>'
    assert docx_facts(build_docx(tmp_path / "a.docx", body, header=header)).images == 2


def test_30k_paragraphs_and_deep_style_chain_stay_linear(tmp_path):
    chain = "".join(style(f"S{i}", f"S{i}", based_on=f"S{i + 1}") for i in range(1000))
    body = "".join(para("w", style="S0") for _ in range(30_000))
    path = build_docx(tmp_path / "a.docx", body, styles=styles_xml(chain))
    started = time.perf_counter()
    facts = docx_facts(path)
    assert len(facts.paragraphs) == 30_000
    assert time.perf_counter() - started < 30
