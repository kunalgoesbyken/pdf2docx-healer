import re
import time
from typing import Any, Dict, List, Optional, Sequence

from docx import Document
from docx.oxml.ns import qn

from docx_healer import internal_links
from docx_healer.internal_links import fix_docx_internal_links
from docx_healer.link_injector import fix_docx_hyperlinks
from docx_healer.ocr_pipeline import pymupdf
from factories import make_docx
from paged import make_paged_docx

BOOKMARK_NAME = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,39}$")


def build_pdf(
    path,
    pages: Sequence[Sequence[str]],
    links: Sequence[Dict[str, Any]],
    named: Optional[Dict[str, Any]] = None,
    fontname: str = "helv",
):
    doc = pymupdf.open()
    for _ in pages:
        doc.new_page()
    for number, lines in enumerate(pages):
        for index, line in enumerate(lines):
            doc[number].insert_text((72, 100 + 20 * index), line, fontsize=11, fontname=fontname)
    for spec in links:
        if "uri" in spec:
            rect = doc[spec["page"]].search_for(spec["anchor"])[spec.get("hit", 0)]
            doc[spec["page"]].insert_link(
                {"kind": pymupdf.LINK_URI, "from": rect, "uri": spec["uri"]}
            )
            continue
        source = doc[spec["page"]]
        rect = spec.get("rect") or source.search_for(spec["anchor"])[spec.get("hit", 0)]
        if "name" in spec:
            source.insert_link({"kind": pymupdf.LINK_NAMED, "from": rect, "name": spec["name"]})
            continue
        target = doc[spec["to_page"]]
        if "to_y" in spec:
            top = spec["to_y"]
        else:
            top = target.search_for(spec["to_text"])[0].y0 - 3
        source.insert_link(
            {
                "kind": pymupdf.LINK_GOTO,
                "from": rect,
                "page": spec["to_page"],
                "to": pymupdf.Point(72, top),
            }
        )
    for name, (page, text) in (named or {}).items():
        target = doc[page]
        top = target.search_for(text)[0].y0 - 3
        catalog = doc.pdf_catalog()
        existing = doc.xref_get_key(catalog, "Dests")
        body = f"/{name} [{target.xref} 0 R /XYZ 72 {target.rect.height - top} 0]"
        prior = existing[1][2:-2] if existing[0] == "dict" else ""
        doc.xref_set_key(catalog, "Dests", f"<<{prior} {body}>>")
    doc.save(str(path))
    doc.close()
    return path


def anchors(path) -> List[Any]:
    doc = Document(str(path))
    found = []
    for link in doc.element.body.iter(qn("w:hyperlink")):
        name = link.get(qn("w:anchor"))
        if name:
            found.append(("".join(t.text or "" for t in link.iter(qn("w:t"))), name))
    return found


def bookmarks(path) -> Dict[str, str]:
    doc = Document(str(path))
    found = {}
    for start in doc.element.body.iter(qn("w:bookmarkStart")):
        paragraph = start.getparent()
        found[start.get(qn("w:name"))] = "".join(t.text or "" for t in paragraph.iter(qn("w:t")))
    return found


def all_text(path) -> List[str]:
    doc = Document(str(path))
    return [
        "".join(t.text or "" for t in p.iter(qn("w:t"))) for p in doc.element.body.iter(qn("w:p"))
    ]


SOURCE = "Recurrent models such as long short-term memory [13] and gated networks"
TARGET = "[13] Hochreiter and Schmidhuber long short-term memory"


def citation_fixture(tmp_path, extra_links=(), docx_pages=None):
    pdf = build_pdf(
        tmp_path / "a.pdf",
        [[SOURCE], ["Bibliography", "[12] Other entry here", TARGET]],
        [{"page": 0, "anchor": "13", "to_page": 1, "to_text": TARGET}, *extra_links],
    )
    pages = docx_pages or [[SOURCE], ["Bibliography", "[12] Other entry here", TARGET]]
    return pdf, make_paged_docx(tmp_path / "a.docx", pages)


def test_link_over_part_of_a_word_wraps_only_the_covered_characters(tmp_path):
    pdf, path = citation_fixture(tmp_path)
    before = all_text(path)

    assert fix_docx_internal_links(path, pdf) == 1

    [(text, name)] = anchors(path)
    assert text == "13"
    assert bookmarks(path) == {name: TARGET}
    assert all_text(path) == before


def test_link_run_keeps_the_direct_formatting_of_the_original_run(tmp_path):
    pdf, path = citation_fixture(tmp_path)

    fix_docx_internal_links(path, pdf)

    doc = Document(str(path))
    link = next(doc.element.body.iter(qn("w:hyperlink")))
    props = link.find(qn("w:r") + "/" + qn("w:rPr"))
    assert props.find(qn("w:sz")).get(qn("w:val")) == "24"
    assert props.find(qn("w:rFonts")).get(qn("w:ascii")) == "DirectFont"


def test_links_to_the_same_destination_share_one_bookmark(tmp_path):
    source = "first cite [13] then again [13] and done"
    pdf = build_pdf(
        tmp_path / "a.pdf",
        [[source], ["Bibliography", TARGET]],
        [
            {"page": 0, "anchor": "13", "hit": 0, "to_page": 1, "to_text": TARGET},
            {"page": 0, "anchor": "13", "hit": 1, "to_page": 1, "to_text": TARGET},
        ],
    )
    path = make_paged_docx(tmp_path / "a.docx", [[source], ["Bibliography", TARGET]])

    assert fix_docx_internal_links(path, pdf) == 2

    found = anchors(path)
    assert [text for text, _ in found] == ["13", "13"]
    assert len({name for _, name in found}) == 1
    assert len(bookmarks(path)) == 1


def test_destination_at_the_top_of_a_page_bookmarks_the_first_paragraph(tmp_path):
    pdf = build_pdf(
        tmp_path / "a.pdf",
        [["see the appendix section below please"], ["Appendix opening line", "second line"]],
        [{"page": 0, "anchor": "appendix", "to_page": 1, "to_y": 0}],
    )
    path = make_paged_docx(
        tmp_path / "a.docx",
        [["see the appendix section below please"], ["Appendix opening line", "second line"]],
    )

    assert fix_docx_internal_links(path, pdf) == 1

    assert list(bookmarks(path).values()) == ["Appendix opening line"]


def test_named_destinations_resolve_and_unknown_names_are_skipped(tmp_path):
    pdf = build_pdf(
        tmp_path / "a.pdf",
        [[SOURCE, "second line that mentions memory"], ["Bibliography", TARGET]],
        [
            {"page": 0, "anchor": "13", "name": "cite.hochreiter"},
            {"page": 0, "anchor": "second", "name": "no.such.dest"},
        ],
        named={"cite.hochreiter": (1, TARGET)},
    )
    path = make_paged_docx(
        tmp_path / "a.docx",
        [[SOURCE, "second line that mentions memory"], ["Bibliography", TARGET]],
    )

    assert fix_docx_internal_links(path, pdf) == 1

    assert [text for text, _ in anchors(path)] == ["13"]
    assert list(bookmarks(path).values()) == [TARGET]


def test_anchor_text_missing_from_the_docx_is_skipped_without_a_bookmark(tmp_path):
    pdf, path = citation_fixture(
        tmp_path, docx_pages=[["Totally different words here"], ["Bibliography", TARGET]]
    )
    before = path.read_bytes()

    assert fix_docx_internal_links(path, pdf) == 0
    assert path.read_bytes() == before


def test_destination_missing_from_the_docx_is_skipped_without_a_dangling_link(tmp_path):
    pdf, path = citation_fixture(
        tmp_path, docx_pages=[[SOURCE], ["Bibliography", "[12] Other entry here"]]
    )
    before = path.read_bytes()

    assert fix_docx_internal_links(path, pdf) == 0
    assert path.read_bytes() == before


def test_ambiguous_anchor_context_is_skipped(tmp_path):
    line = "Repeated sentence with a marker [3] in the same words"
    pdf = build_pdf(
        tmp_path / "a.pdf",
        [[line, line], ["Bibliography", "[3] Third entry about things"]],
        [
            {
                "page": 0,
                "anchor": "3",
                "hit": 0,
                "to_page": 1,
                "to_text": "[3] Third entry about things",
            }
        ],
    )
    path = make_paged_docx(
        tmp_path / "a.docx", [[line, line], ["Bibliography", "[3] Third entry about things"]]
    )
    before = path.read_bytes()

    assert fix_docx_internal_links(path, pdf) == 0
    assert path.read_bytes() == before


def test_rect_over_two_lines_links_the_text_across_both(tmp_path):
    pdf_path = tmp_path / "a.pdf"
    doc = pymupdf.open()
    doc.new_page()
    doc.new_page()
    doc[0].insert_text((72, 100), "as shown in Figure", fontsize=11)
    doc[0].insert_text((72, 114), "3 for details", fontsize=11)
    doc[1].insert_text((72, 100), "Figure 3: The model overview", fontsize=11)
    rect = pymupdf.Rect(70, 88, 300, 118)
    doc[0].insert_link(
        {"kind": pymupdf.LINK_GOTO, "from": rect, "page": 1, "to": pymupdf.Point(72, 87)}
    )
    doc.save(str(pdf_path))
    doc.close()
    path = make_paged_docx(
        tmp_path / "a.docx",
        [["as shown in Figure 3 for details"], ["Figure 3: The model overview"]],
    )

    assert fix_docx_internal_links(path, pdf_path) == 1

    assert [text for text, _ in anchors(path)] == ["as shown in Figure 3 for details"]


def test_existing_hyperlinks_are_never_nested(tmp_path):
    url = "https://example.com/ref"
    source = f"read the paper at {url} for background"
    pdf = build_pdf(
        tmp_path / "a.pdf",
        [[source], ["Bibliography", TARGET]],
        [{"page": 0, "anchor": "example.com/ref", "to_page": 1, "to_text": TARGET}],
    )
    path = make_paged_docx(tmp_path / "a.docx", [[source], ["Bibliography", TARGET]])
    assert fix_docx_hyperlinks(path) == 1

    assert fix_docx_internal_links(path, pdf) == 0

    assert anchors(path) == []
    assert not bookmarks(path)


def test_external_pdf_links_are_left_to_the_url_pass(tmp_path):
    pdf = build_pdf(
        tmp_path / "a.pdf", [[SOURCE]], [{"page": 0, "anchor": "13", "uri": "https://example.com"}]
    )
    path = make_paged_docx(tmp_path / "a.docx", [[SOURCE]])
    before = path.read_bytes()

    assert fix_docx_internal_links(path, pdf) == 0
    assert path.read_bytes() == before


def test_cjk_anchor_text_is_linked(tmp_path):
    source = "详见图1以及附录中的说明"
    target = "图1 模型结构总览"
    pdf = build_pdf(
        tmp_path / "a.pdf",
        [[source], ["附录", target]],
        [{"page": 0, "anchor": "图1", "to_page": 1, "to_text": target}],
        fontname="china-s",
    )
    path = make_paged_docx(tmp_path / "a.docx", [[source], ["附录", target]])

    assert fix_docx_internal_links(path, pdf) == 1

    assert [text for text, _ in anchors(path)] == ["图1"]
    assert list(bookmarks(path).values()) == [target]


def test_anchor_inside_a_table_cell_is_linked(tmp_path):
    pdf = build_pdf(
        tmp_path / "a.pdf",
        [[SOURCE], ["Bibliography", TARGET]],
        [{"page": 0, "anchor": "13", "to_page": 1, "to_text": TARGET}],
    )
    path = make_paged_docx(tmp_path / "a.docx", [[("tbl", SOURCE)], ["Bibliography", TARGET]])

    assert fix_docx_internal_links(path, pdf) == 1

    assert [text for text, _ in anchors(path)] == ["13"]


def test_destination_on_a_later_line_bookmarks_the_whole_paragraph(tmp_path):
    first = "A long reference entry whose first line"
    second = "continues here on the second line"
    pdf = build_pdf(
        tmp_path / "a.pdf",
        [[SOURCE], [first, second]],
        [{"page": 0, "anchor": "13", "to_page": 1, "to_text": second}],
    )
    path = make_paged_docx(tmp_path / "a.docx", [[SOURCE], [f"{first} {second}"]])

    assert fix_docx_internal_links(path, pdf) == 1

    assert list(bookmarks(path).values()) == [f"{first} {second}"]


def test_bookmark_names_are_valid_unique_and_ids_do_not_collide(tmp_path):
    pdf, path = citation_fixture(tmp_path)
    doc = Document(str(path))
    body = doc.element.body
    for number, name in ((0, "_GoBack"), (1, "pdf_dest_1")):
        start = body.makeelement(
            qn("w:bookmarkStart"), {qn("w:id"): str(number), qn("w:name"): name}
        )
        end = body.makeelement(qn("w:bookmarkEnd"), {qn("w:id"): str(number)})
        body.insert(0, end)
        body.insert(0, start)
    doc.save(str(path))

    assert fix_docx_internal_links(path, pdf) == 1

    doc = Document(str(path))
    starts = list(doc.element.body.iter(qn("w:bookmarkStart")))
    ends = list(doc.element.body.iter(qn("w:bookmarkEnd")))
    ids = [s.get(qn("w:id")) for s in starts]
    assert len(set(ids)) == len(ids) == 3
    assert sorted(ids) == sorted(e.get(qn("w:id")) for e in ends)
    names = [s.get(qn("w:name")) for s in starts]
    assert len(set(names)) == 3
    assert all(BOOKMARK_NAME.match(n) for n in names if n != "_GoBack")
    assert anchors(path)[0][1] in names


def test_without_page_sections_unique_anchors_are_still_linked(tmp_path):
    pdf = build_pdf(
        tmp_path / "a.pdf",
        [[SOURCE], ["Bibliography", TARGET]],
        [{"page": 0, "anchor": "13", "to_page": 1, "to_text": TARGET}],
    )

    def build(doc):
        for text in (SOURCE, "Bibliography", TARGET):
            doc.add_paragraph(text)

    path = make_docx(tmp_path / "a.docx", build)

    assert fix_docx_internal_links(path, pdf) == 1

    assert [text for text, _ in anchors(path)] == ["13"]


def test_pass_is_idempotent(tmp_path):
    pdf, path = citation_fixture(tmp_path)

    assert fix_docx_internal_links(path, pdf) == 1
    assert fix_docx_internal_links(path, pdf) == 0
    assert len(anchors(path)) == 1


def build_many_links_pdf(path, tokens_per_line, lines):
    doc = pymupdf.open()
    doc.new_page()
    doc.new_page()
    doc[1].insert_text((72, 100), "Shared destination paragraph text", fontsize=11)
    texts = []
    for line in range(lines):
        row = " ".join(f"[{line * tokens_per_line + n}]" for n in range(tokens_per_line))
        doc[0].insert_text((20, 20 + 6 * line), row, fontsize=4)
        texts.append(row)
    top = doc[1].search_for("Shared destination")[0].y0 - 3
    for word in doc[0].get_text("words"):
        doc[0].insert_link(
            {
                "kind": pymupdf.LINK_GOTO,
                "from": pymupdf.Rect(word[:4]),
                "page": 1,
                "to": pymupdf.Point(72, top),
            }
        )
    doc.save(str(path))
    doc.close()
    return texts


def test_thousands_of_links_are_linked_in_bounded_time(tmp_path):
    texts = build_many_links_pdf(tmp_path / "a.pdf", 30, 50)
    path = make_paged_docx(tmp_path / "a.docx", [texts, ["Shared destination paragraph text"]])

    start = time.perf_counter()
    assert fix_docx_internal_links(path, tmp_path / "a.pdf") == 1500
    assert time.perf_counter() - start < 120

    assert len(anchors(path)) == 1500
    assert len(bookmarks(path)) == 1


def test_link_count_over_the_limit_leaves_the_document_untouched(tmp_path, monkeypatch):
    texts = build_many_links_pdf(tmp_path / "a.pdf", 4, 2)
    path = make_paged_docx(tmp_path / "a.docx", [texts, ["Shared destination paragraph text"]])
    before = path.read_bytes()
    monkeypatch.setattr(internal_links, "MAX_INTERNAL_LINKS", 7)

    assert fix_docx_internal_links(path, tmp_path / "a.pdf") == 0
    assert path.read_bytes() == before


def test_link_count_at_the_limit_still_works(tmp_path, monkeypatch):
    texts = build_many_links_pdf(tmp_path / "a.pdf", 4, 2)
    path = make_paged_docx(tmp_path / "a.docx", [texts, ["Shared destination paragraph text"]])
    monkeypatch.setattr(internal_links, "MAX_INTERNAL_LINKS", 8)

    assert fix_docx_internal_links(path, tmp_path / "a.pdf") == 8
