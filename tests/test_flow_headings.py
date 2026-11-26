from docx_healer.converter import HealerConfig, heal
from docx_inspect import body_hyperlinks
from flow_fixtures import LEFT, build_pdf, convert, lines, new_page, para_gap, paragraphs, put


def _styles(path):
    return [(p["style"], p["text"]) for p in paragraphs(path) if p["text"].strip()]


def _body(page, y, count=3):
    for index in range(count):
        y = lines(
            page, LEFT, y,
            [f"Body paragraph {index} runs for a couple of lines of ordinary text", "and then finishes right here."],
        ) + para_gap()
    return y


def test_outline_levels_become_heading_styles_even_with_number_prefixes(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 100, "1 Introduction", 14, "hebo")
        _body(page, 130)
        page = new_page(doc)
        put(page, LEFT, 100, "2 Background", 14, "hebo")
        put(page, LEFT, 200, "2.1 Details", 12, "hebo")
        _body(page, 230)
        doc.set_toc([[1, "Introduction", 1], [1, "Background", 2], [2, "Details", 2]])

    out = convert(tmp_path, build)

    heads = [(s, t) for s, t in _styles(out) if s and s.startswith("Heading")]
    assert heads == [("Heading1", "1 Introduction"), ("Heading1", "2 Background"), ("Heading2", "2.1 Details")]


def test_font_statistics_pick_heading_levels_and_skip_bold_lead_ins(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 100, "Document Title Here", 20, "hebo")
        put(page, LEFT, 140, "A Section Heading", 14, "hebo")
        y = _body(page, 170)
        put(page, LEFT, y + 6, "Note.", 10, "hebo")
        put(page, LEFT + 30, y + 6, "Then regular words continue in the same paragraph.", 10)

    out = convert(tmp_path, build)

    styled = {t: s for s, t in _styles(out) if s}
    assert styled == {"Document Title Here": "Heading1", "A Section Heading": "Heading2"}
    assert any(t.startswith("Note.") for _, t in _styles(out))


def test_heal_in_editable_mode_links_plain_urls_and_keeps_headings(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 100, "Quarterly Report", 20, "hebo")
        _body(page, 140)
        put(page, LEFT, 400, "Read more at https://example.com/notes for details.")

    pdf = build_pdf(tmp_path / "in.pdf", build)
    out = tmp_path / "out.docx"

    heal(str(pdf), str(out), HealerConfig(mode="editable", verbose=False))

    assert [t for s, t in _styles(out) if s == "Heading1"] == ["Quarterly Report"]
    assert ("https://example.com/notes", "https://example.com/notes") in body_hyperlinks(out)
