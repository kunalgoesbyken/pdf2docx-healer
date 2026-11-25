from flow_fixtures import LEFT, convert, items, labels, new_page, numbering_levels, paragraphs, put


def _bullet(page, x, y, text, glyph="·"):
    put(page, x, y, glyph)
    put(page, x + 14, y, text)
    return y + 14


def _numbered(page, x, y, marker, text):
    put(page, x, y, f"{marker} {text}")
    return y + 14


def _list_paragraphs(path):
    return [p for p in paragraphs(path) if p["num"] is not None]


def test_bullets_nest_by_indent_and_lose_their_glyph(tmp_path):
    def build(doc):
        page = new_page(doc)
        y = _bullet(page, LEFT, 100, "top one")
        y = _bullet(page, 100, y, "second level a", "-")
        y = _bullet(page, 100, y, "second level b", "-")
        y = _bullet(page, 128, y, "third level")
        y = _bullet(page, LEFT, y, "top two")

    out = convert(tmp_path, build)

    listed = _list_paragraphs(out)
    assert [p["text"] for p in listed] == [
        "top one", "second level a", "second level b", "third level", "top two",
    ]
    assert [p["num"][1] for p in listed] == [0, 1, 1, 2, 0]
    assert len({p["num"][0] for p in listed}) == 1
    assert labels(out) == ["•", "-", "-", "•", "•"]


def test_three_formats_and_a_late_start_each_restart_and_show_the_pdf_number(tmp_path):
    markers = ["1.", "2.", "3.", "a)", "b)", "c)", "i.", "ii.", "iii.", "4.", "5."]

    def build(doc):
        page = new_page(doc)
        y = 100
        for index, marker in enumerate(markers):
            y = _numbered(page, LEFT, y, marker, f"item number {index} text")

    out = convert(tmp_path, build)

    listed = _list_paragraphs(out)
    assert [p["text"] for p in listed] == [f"item number {i} text" for i in range(len(markers))]
    assert [label for label in labels(out) if label] == markers
    ids = [p["num"][0] for p in listed]
    assert len(set(ids)) == 4
    levels = numbering_levels(out)
    assert [levels[i][0][1] for i in dict.fromkeys(ids)] == [
        "decimal", "lowerLetter", "lowerRoman", "decimal",
    ]
    assert [levels[i][0][0] for i in dict.fromkeys(ids)] == [1, 1, 1, 4]


def test_numbered_list_keeps_counting_across_a_page_break(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 100, "Steps follow below in order.")
        _numbered(page, LEFT, 715, "1.", "alpha step")
        _numbered(page, LEFT, 729, "2.", "beta step")
        page = new_page(doc)
        _numbered(page, LEFT, 110, "3.", "gamma step")
        _numbered(page, LEFT, 124, "4.", "delta step")

    out = convert(tmp_path, build)

    listed = _list_paragraphs(out)
    assert [p["text"] for p in listed] == ["alpha step", "beta step", "gamma step", "delta step"]
    assert len({p["num"][0] for p in listed}) == 1
    assert [label for label in labels(out) if label] == ["1.", "2.", "3.", "4."]


def test_unsure_markers_stay_literal_text(tmp_path):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 100, "2024. was a good year for the whole team here.")
        put(page, LEFT, 140, "A. Smith wrote the first report on this topic.")
        put(page, LEFT, 180, "1. Only item in this lonely list.")
        put(page, LEFT, 220, "A closing sentence sits between the examples.")
        y = _numbered(page, LEFT, 260, "1.", "first")
        y = _numbered(page, LEFT, y, "2.", "second")
        _numbered(page, LEFT, y, "4.", "fourth skips three")

    out = convert(tmp_path, build)

    listed = _list_paragraphs(out)
    assert [p["text"] for p in listed] == ["first", "second"]
    literal = [p["text"] for p in paragraphs(out) if p["num"] is None and p["text"]]
    assert literal == [
        "2024. was a good year for the whole team here.",
        "A. Smith wrote the first report on this topic.",
        "1. Only item in this lonely list.",
        "A closing sentence sits between the examples.",
        "4. fourth skips three",
    ]


def test_nested_numbering_restarts_under_every_new_parent(tmp_path):
    def build(doc):
        page = new_page(doc)
        y = _numbered(page, LEFT, 100, "1.", "parent one")
        y = _numbered(page, 100, y, "a.", "child one a")
        y = _numbered(page, 100, y, "b.", "child one b")
        y = _numbered(page, LEFT, y, "2.", "parent two")
        y = _numbered(page, 100, y, "a.", "child two a")
        _numbered(page, 100, y, "b.", "child two b")

    out = convert(tmp_path, build)

    assert [p["num"][1] for p in _list_paragraphs(out)] == [0, 1, 1, 0, 1, 1]
    assert [label for label in labels(out) if label] == ["1.", "a.", "b.", "2.", "a.", "b."]
    assert len(items(out)) == 6
