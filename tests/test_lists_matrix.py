import time

import pytest
from docx import Document

from docx_healer.list_detector import fix_lists_in_docx
from test_lists_support import make, num_ids, paragraphs_info, rendered

BULLET = "•"


def run(tmp_path, items):
    path = make(tmp_path, items)
    fix_lists_in_docx(path)
    return path


def untouched(tmp_path, lines):
    path = make(tmp_path, lines)
    before = [p.text for p in Document(str(path)).paragraphs]
    assert fix_lists_in_docx(path) == 0
    after = Document(str(path))
    assert [p.text for p in after.paragraphs] == before
    assert all(p.style.name == "Normal" for p in after.paragraphs)


def test_back_to_back_decimal_lists_restart(tmp_path):
    lines = ["1. one", "2. two", "3. three", "1. alpha", "2. beta"]
    path = run(tmp_path, lines)
    assert rendered(path) == lines
    ids = num_ids(path)
    assert len(set(ids[:3])) == 1 and len(set(ids[3:])) == 1 and ids[0] != ids[3]


def test_lists_of_different_formats_keep_their_format(tmp_path):
    lines = ["a. apple", "b. berry", "i. first", "ii. second", "iii. third", "A. up", "B. up too",
             "I. top", "II. next", "(1) paren", "(2) paren two", "1) close", "2) close two"]
    path = run(tmp_path, lines)
    assert rendered(path) == lines
    fmts = [d["numFmt"] for _, d in paragraphs_info(path)]
    assert fmts == ["lowerLetter"] * 2 + ["lowerRoman"] * 3 + ["upperLetter"] * 2 + ["upperRoman"] * 2 + ["decimal"] * 4


def test_list_starting_at_three_keeps_its_start(tmp_path):
    path = run(tmp_path, ["3. third", "4. fourth", "5. fifth"])
    assert rendered(path) == ["3. third", "4. fourth", "5. fifth"]
    assert {d["start"] for _, d in paragraphs_info(path)} == {3}


def test_list_interrupted_by_a_paragraph_then_continued(tmp_path):
    lines = ["1. one", "2. two", "A normal paragraph in between.", "3. three", "4. four"]
    path = run(tmp_path, lines)
    assert rendered(path) == lines


def test_gap_in_sequence_keeps_the_literal_text(tmp_path):
    lines = ["1. one", "2. two", "4. four"]
    path = run(tmp_path, lines)
    assert rendered(path) == lines


def test_outline_numbers_are_left_alone(tmp_path):
    untouched(tmp_path, ["3.1 Encoder and Decoder Stacks", "3.2 Attention", "1.1. First clause here"])


def test_nested_levels_follow_relative_indent(tmp_path):
    path = run(tmp_path, [
        (f"{BULLET}  top", 19), (f"{BULLET}  top again", 19),
        ("○  middle", 42), ("■  deepest", 73), (f"{BULLET}  back to top", 19),
    ])
    details = paragraphs_info(path)
    assert [text for text, _ in details] == ["top", "top again", "middle", "deepest", "back to top"]
    assert [d["style"] for _, d in details] == ["List Bullet", "List Bullet", "List Bullet 2", "List Bullet 3", "List Bullet"]


def test_indent_offset_does_not_shift_the_level(tmp_path):
    path = run(tmp_path, [(f"{BULLET} a", 100), (f"{BULLET} b", 100)])
    assert {d["style"] for _, d in paragraphs_info(path)} == {"List Bullet"}


@pytest.mark.parametrize("glyph", ["■", "►", "➢", "○", "", "", "", "", "•", "▪", "❖", "✔"])
def test_common_bullet_glyphs_are_detected(tmp_path, glyph):
    path = run(tmp_path, [f"{glyph}  item text"])
    assert rendered(path) == [f"{BULLET} item text"]


def test_lone_numbered_line_is_not_a_list(tmp_path):
    untouched(tmp_path, ["4. Results"])


def test_numbered_heading_keeps_its_number(tmp_path):
    path = make(tmp_path, [("1. Goals", None, "Heading 2"), ("2. Steps", None, "Heading 2"), ("Title line", None, "Title")])
    fix_lists_in_docx(path)
    assert [p.text for p in Document(str(path)).paragraphs] == ["1. Goals", "2. Steps", "Title line"]
    assert num_ids(path) == [None, None, None]


def test_numbered_run_of_headings_next_to_body_text_is_untouched_when_headings_are_styled(tmp_path):
    path = make(tmp_path, [("3.1 Encoder and Decoder Stacks", None, "Heading 3"), "Encoder: the encoder is composed of a stack."])
    assert fix_lists_in_docx(path) == 0


@pytest.mark.parametrize("fragment", [
    ["n"], ["d + n"], ["n", "d + n"], ["+ n"], ["k=Pdk i=1qiki, has mean 0 and variance dk."],
    ["- 5 degrees"], ["* note"], ["> quote"], ["~ roughly"], ["— a quote"],
])
def test_math_and_sentence_fragments_are_not_list_items(tmp_path, fragment):
    untouched(tmp_path, fragment)


@pytest.mark.parametrize("line", [
    "2024. was a good year", "e.g. a thing", "i.e. a thing", "Fig. 3. Results", "U.S. army news",
    "U. S. Army news", "No. 5 on the list", "1.5 million people", "3.Results", "10.5% growth", "LLC. Enter the name",
])
def test_abbreviations_and_numbers_in_prose_are_not_markers(tmp_path, line):
    untouched(tmp_path, [line, "Plain follow-up sentence."])


@pytest.mark.parametrize("line", [f"{BULLET}item", "-item", "1.item", "a.b. c", "(1)x", "■item"])
def test_marker_without_following_space_is_not_a_marker(tmp_path, line):
    untouched(tmp_path, [line])


def test_non_english_text_after_markers(tmp_path):
    lines = ["1. 你好世界", "2. مرحبا بالعالم",
             "3. नमस्ते दुनिया",
             f"{BULLET} 日本語のテキスト", f"{BULLET} ภาษาไทย"]
    path = run(tmp_path, lines)
    assert rendered(path) == lines


def test_non_ascii_digits_are_not_markers(tmp_path):
    untouched(tmp_path, ["٣. text", "٤. text", "１. wide", "２. wide"])


def test_weak_ascii_bullets_need_two_in_a_row(tmp_path):
    path = run(tmp_path, ["- first", "- second", "plain text", "- lonely stays"])
    assert rendered(path) == [f"{BULLET} first", f"{BULLET} second", "plain text", "- lonely stays"]


def test_alphabetic_and_roman_ambiguity_is_resolved_by_the_sequence(tmp_path):
    cases = [
        (["i. one", "ii. two"], "lowerRoman"), (["i. one", "j. two"], "lowerLetter"),
        (["c. one", "d. two"], "lowerLetter"), (["x. one", "xi. two"], "lowerRoman"),
        (["v. one", "vi. two"], "lowerRoman"), (["I. one", "II. two"], "upperRoman"),
        (["H. one", "I. two"], "upperLetter"),
    ]
    for index, (lines, fmt) in enumerate(cases):
        path = make(tmp_path, lines, f"case{index}.docx")
        fix_lists_in_docx(path)
        assert rendered(path) == lines
        assert {d["numFmt"] for _, d in paragraphs_info(path)} == {fmt}


def test_non_canonical_roman_is_not_a_marker(tmp_path):
    untouched(tmp_path, ["iiii. text", "iiiii. text", "LLC. Enter", "mmmm. text"])


def test_single_line_breaks_in_one_paragraph_become_separate_items(tmp_path):
    path = run(tmp_path, ["1. one \n2. two \n3. three", "1. alpha \n2. beta"])
    assert rendered(path) == ["1. one", "2. two", "3. three", "1. alpha", "2. beta"]
    ids = num_ids(path)
    assert len(set(ids[:3])) == 1 and ids[3] != ids[0]


def test_intro_line_before_bullets_in_the_same_paragraph(tmp_path):
    path = run(tmp_path, [f"Intro line\n{BULLET} a\n{BULLET} b\nclosing words"])
    assert rendered(path) == ["Intro line", f"{BULLET} a", f"{BULLET} b\nclosing words"]


def test_numbered_parent_with_lettered_children(tmp_path):
    lines = ["1. parent", ("a. child", 36), ("b. child two", 36), "2. second parent", ("a. again", 36), ("b. and more", 36)]
    path = run(tmp_path, lines)
    assert rendered(path) == ["1. parent", "a. child", "b. child two", "2. second parent", "a. again", "b. and more"]
    details = [d for _, d in paragraphs_info(path)]
    assert [d["ilvl"] for d in details] == [0, 1, 1, 0, 1, 1]


def test_bullets_with_numbered_children(tmp_path):
    path = run(tmp_path, [f"{BULLET} parent", ("1. child", 36), ("2. child two", 36), f"{BULLET} parent two"])
    assert rendered(path) == [f"{BULLET} parent", "1. child", "2. child two", f"{BULLET} parent two"]


def test_formatting_and_marker_in_separate_runs_survive(tmp_path):
    def build_doc(doc):
        paragraph = doc.add_paragraph()
        paragraph.add_run(BULLET)
        paragraph.add_run("  ")
        paragraph.add_run("bold part ").bold = True
        paragraph.add_run("plain part")

    from factories import make_docx

    path = make_docx(tmp_path / "f.docx", build_doc)
    fix_lists_in_docx(path)
    paragraph = Document(str(path)).paragraphs[0]
    assert paragraph.text == "bold part plain part"
    assert [run.bold for run in paragraph.runs if run.text.strip()] == [True, None]


def test_tab_after_marker_is_consumed(tmp_path):
    path = run(tmp_path, [f"{BULLET}\titem", f"{BULLET}\tsecond"])
    assert rendered(path) == [f"{BULLET} item", f"{BULLET} second"]


def test_second_run_changes_nothing(tmp_path):
    path = run(tmp_path, ["1. one", "2. two", f"{BULLET} x", "3.1 outline"])
    before = rendered(path)
    assert fix_lists_in_docx(path) == 0
    assert rendered(path) == before


def test_aggressive_mode_also_styles_single_numbered_lines(tmp_path):
    path = make(tmp_path, ["4. Results"])
    assert fix_lists_in_docx(path, aggressive=True) == 1
    assert rendered(path) == ["4. Results"]


def test_very_long_document_is_linear(tmp_path):
    lines = [f"{index % 40 + 1}. item number {index}" for index in range(6000)]
    path = make(tmp_path, lines)
    started = time.perf_counter()
    fix_lists_in_docx(path)
    assert time.perf_counter() - started < 60
    assert rendered(path) == lines


def test_unstyled_numbered_headings_around_a_list_stay_headings(tmp_path):
    from docx.shared import Pt

    from factories import make_docx

    def build_doc(doc):
        for text, size, indent in (("1. Goals", 16, 0), ("\u2022  Ship it", 11, 19), ("\u2022  Test it", 11, 19),
                                   ("2. Steps", 16, 0), ("1. Write it", 11, 19), ("2. Run it", 11, 19)):
            paragraph = doc.add_paragraph()
            paragraph.add_run(text).font.size = Pt(size)
            paragraph.paragraph_format.left_indent = Pt(indent)

    path = make_docx(tmp_path / "h.docx", build_doc)
    fix_lists_in_docx(path)

    assert rendered(path) == ["1. Goals", f"{BULLET} Ship it", f"{BULLET} Test it", "2. Steps", "1. Write it", "2. Run it"]
    styles = [d["style"] for _, d in paragraphs_info(path)]
    assert styles == ["Normal", "List Bullet", "List Bullet", "Normal", "List Number", "List Number"]


def test_nested_children_in_a_slightly_smaller_font_keep_the_parent_list(tmp_path):
    from docx.shared import Pt

    from factories import make_docx

    def build_doc(doc):
        for text, size, indent in (("1. parent", 12, 0), ("a. child", 10, 36), ("b. child", 10, 36), ("2. parent", 12, 0)):
            paragraph = doc.add_paragraph()
            paragraph.add_run(text).font.size = Pt(size)
            paragraph.paragraph_format.left_indent = Pt(indent)

    path = make_docx(tmp_path / "n.docx", build_doc)
    fix_lists_in_docx(path)

    assert rendered(path) == ["1. parent", "a. child", "b. child", "2. parent"]
    assert [d["ilvl"] for _, d in paragraphs_info(path)] == [0, 1, 1, 0]
