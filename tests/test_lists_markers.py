import pytest

from docx_healer.list_markers import (
    BULLET,
    DECIMAL,
    LOWER_LETTER,
    LOWER_ROMAN,
    UPPER_LETTER,
    UPPER_ROMAN,
    parse_marker,
)


@pytest.mark.parametrize(
    "text, readings",
    [
        ("1. item", {(DECIMAL, "."): 1}),
        ("  12) item", {(DECIMAL, ")"): 12}),
        ("(3) item", {(DECIMAL, "()"): 3}),
        ("a. item", {(LOWER_LETTER, "."): 1}),
        ("B) item", {(UPPER_LETTER, ")"): 2}),
        ("i. item", {(LOWER_LETTER, "."): 9, (LOWER_ROMAN, "."): 1}),
        ("IV. item", {(UPPER_ROMAN, "."): 4}),
        ("(xii) item", {(LOWER_ROMAN, "()"): 12}),
        ("C. item", {(UPPER_LETTER, "."): 3, (UPPER_ROMAN, "."): 100}),
        ("•  item", {(BULLET, ""): 0}),
        ("- item", {(BULLET, "-"): 0}),
        ("\titem", {(BULLET, ""): 0}),
    ],
)
def test_markers_and_their_readings(text, readings):
    marker = parse_marker(text)
    assert marker is not None
    assert marker.readings == readings


@pytest.mark.parametrize(
    "text, end",
    [("1.  item", 4), ("  •\t item", 5), ("(a) x", 4)],
)
def test_end_is_where_the_text_starts(text, end):
    marker = parse_marker(text)
    assert marker is not None
    assert text[marker.end:] == text[end:]


@pytest.mark.parametrize(
    "text",
    ["", "   ", "1.", "1. ", "item", "1.5 x", "01. x", "0. x", "1000. x", "ii x", "iiii. x", "LLC. x", "Ab. x",
     "1a. x", "e.g. x", "(1 x", "1) ", "٣. x", "１. x", "-5 x", "--- x", "•", "•x", "a.b x"],
)
def test_not_markers(text):
    assert parse_marker(text) is None


def test_lone_bullet_glyph_counts_only_when_bare_is_allowed():
    assert parse_marker("•") is None
    assert parse_marker("•", bare=True) is not None
    assert parse_marker("-", bare=True) is None
    assert parse_marker("1.", bare=True) is None


def test_strength_separates_glyphs_from_ascii_markers():
    assert parse_marker("■ x").strong
    assert parse_marker(" x").strong
    assert not parse_marker("- x").strong
    assert not parse_marker("1. x").strong
