import logging

import pymupdf
from docx import Document
from docx.oxml.ns import qn

from docx_healer.flow_model import FlowLimits
from docx_inspect import body_hyperlinks
from flow_fixtures import LEFT, convert, items, new_page, paragraphs, put, runs, texts


def _link(page, needle, uri):
    rect = page.search_for(needle)[0]
    page.insert_link({"kind": pymupdf.LINK_URI, "from": rect, "uri": uri})


def test_external_link_wraps_exactly_the_covered_characters(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 100, "See the project page for details and more.")
        _link(page, "project page", "https://example.org/project")

    out = convert(tmp_path, build)

    assert body_hyperlinks(out) == [("project page", "https://example.org/project")]
    assert texts(out) == ["See the project page for details and more."]


def test_internal_link_points_at_a_bookmark_in_the_target_paragraph(tmp_path):
    def build(doc):
        new_page(doc)
        new_page(doc)
        page = doc[0]
        put(page, LEFT, 100, "Jump to the glossary entry or to nowhere at all.")
        first = page.search_for("glossary entry")[0]
        nowhere = page.search_for("nowhere")[0]
        page.insert_link({"kind": pymupdf.LINK_GOTO, "from": first, "page": 1, "to": pymupdf.Point(72, 288), "zoom": 0})
        page.insert_link({"kind": pymupdf.LINK_GOTO, "from": nowhere, "page": 1, "to": pymupdf.Point(500, 700), "zoom": 0})
        page = doc[1]
        put(page, LEFT, 100, "First paragraph on the second page.")
        put(page, LEFT, 300, "Glossary paragraph that is the link target.")
        put(page, LEFT, 400, "Third paragraph on the second page.")

    out = convert(tmp_path, build)

    xml = Document(str(out)).element
    links = xml.findall(f".//{qn('w:hyperlink')}")
    assert len(links) == 1
    anchor = links[0].get(qn("w:anchor"))
    starts = {b.get(qn("w:name")): b for b in xml.iter(qn("w:bookmarkStart"))}
    assert set(starts) == {anchor}
    start = starts[anchor]
    ends = [e.get(qn("w:id")) for e in xml.iter(qn("w:bookmarkEnd"))]
    assert start.get(qn("w:id")) in ends
    owner = start.getparent()
    assert "Glossary paragraph" in "".join(t.text for t in owner.iter(qn("w:t")))
    assert "".join(t.text for t in links[0].iter(qn("w:t"))) == "glossary entry"


def test_inline_formatting_is_kept_without_sizing_every_run(tmp_path):
    def build(doc):
        page = new_page(doc)
        runs(page, LEFT, 100, [
            ("Plain ", "helv"), ("bold", "hebo"), (" ", "helv"), ("italic", "heit"), (" ", "helv"),
            ("mono", "cour"), (" E=mc", "helv"), ("2", "helv", 6, 4), (" end.", "helv"),
        ])

    out = convert(tmp_path, build)

    paras = [p for p in paragraphs(out) if p["text"].strip()]
    assert [p["text"] for p in paras] == ["Plain bold italic mono E=mc2 end."]
    by_text = {}
    for run in paras[0]["el"].iter(qn("w:r")):
        by_text["".join(t.text for t in run.iter(qn("w:t"))).strip()] = run
    assert by_text["bold"].find(f"{qn('w:rPr')}/{qn('w:b')}") is not None
    assert by_text["italic"].find(f"{qn('w:rPr')}/{qn('w:i')}") is not None
    assert "Courier New" in by_text["mono"].find(f"{qn('w:rPr')}/{qn('w:rFonts')}").get(qn("w:ascii"))
    assert by_text["2"].find(f"{qn('w:rPr')}/{qn('w:vertAlign')}").get(qn("w:val")) == "superscript"
    assert not list(paras[0]["el"].iter(qn("w:sz")))


def test_link_cap_leaves_the_rest_as_plain_text_and_warns_only_over_the_limit(tmp_path, caplog):
    words = [f"word{n}" for n in range(8)]

    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 100, " ".join(words))
        for word in words:
            _link(page, word, f"https://example.org/{word}")

    with caplog.at_level(logging.WARNING, logger="docx_healer"):
        over = convert(tmp_path, build, name="over", limits=FlowLimits(max_links_per_page=5))
    assert len(body_hyperlinks(over)) == 5
    assert "max_links_per_page" in caplog.text
    assert texts(over) == [" ".join(words)]

    caplog.clear()
    with caplog.at_level(logging.WARNING, logger="docx_healer"):
        exact = convert(tmp_path, build, name="exact", limits=FlowLimits(max_links_per_page=8))
    assert len(body_hyperlinks(exact)) == 8
    assert "max_links_per_page" not in caplog.text
    assert len(items(exact)) == 1
