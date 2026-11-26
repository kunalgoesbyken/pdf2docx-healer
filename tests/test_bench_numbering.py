from lxml import etree

from bench.numbering import NumberingRenderer, format_number
from bench_docx_builder import abstract, lvl, num, numbering_xml, start_override


def renderer(abstracts, nums):
    return NumberingRenderer(etree.fromstring(numbering_xml(abstracts, nums).encode()))


def run(rend, items):
    out = []
    for num_id, ilvl in items:
        label = rend.next_label(num_id, ilvl)
        out.append(None if label is None else label.text)
    return out


def test_start_three_decimal():
    rend = renderer(abstract(0, lvl(0, start=3)), num(1, 0))
    assert run(rend, [(1, 0)] * 3) == ["3.", "4.", "5."]


def test_second_num_with_start_override_restarts():
    rend = renderer(abstract(0, lvl(0)), num(1, 0) + num(2, 0, start_override(0, 1)))
    assert run(rend, [(1, 0)] * 3 + [(2, 0)] * 2) == ["1.", "2.", "3.", "1.", "2."]


def test_two_nums_on_one_abstract_without_override_continue():
    rend = renderer(abstract(0, lvl(0)), num(1, 0) + num(2, 0))
    assert run(rend, [(1, 0), (1, 0), (2, 0), (2, 0)]) == ["1.", "2.", "3.", "4."]


def test_start_override_value_wins_over_lvl_start():
    rend = renderer(abstract(0, lvl(0, start=1)), num(1, 0, start_override(0, 7)))
    assert run(rend, [(1, 0), (1, 0)]) == ["7.", "8."]


def test_lower_roman_second_level_restarts_after_first_level():
    levels = lvl(0) + lvl(1, fmt="lowerRoman", text="%2.")
    rend = renderer(abstract(0, levels), num(1, 0))
    items = [(1, 0), (1, 1), (1, 1), (1, 0), (1, 1)]
    assert run(rend, items) == ["1.", "i.", "ii.", "2.", "i."]


def test_multilevel_text_references_parent_counter():
    levels = lvl(0) + lvl(1, text="%1.%2")
    rend = renderer(abstract(0, levels), num(1, 0))
    items = [(1, 0), (1, 1), (1, 1), (1, 0), (1, 1)]
    assert run(rend, items) == ["1.", "1.1", "1.2", "2.", "2.1"]


def test_level_without_parent_uses_parent_start_value():
    levels = lvl(0, start=4) + lvl(1, text="%1.%2")
    rend = renderer(abstract(0, levels), num(1, 0))
    assert run(rend, [(1, 1)]) == ["4.1"]


def test_lvl_restart_zero_never_restarts():
    levels = lvl(0) + lvl(1, fmt="lowerRoman", text="%2.", extra='<w:lvlRestart w:val="0"/>')
    rend = renderer(abstract(0, levels), num(1, 0))
    items = [(1, 0), (1, 1), (1, 0), (1, 1)]
    assert run(rend, items) == ["1.", "i.", "2.", "ii."]


def test_missing_numfmt_is_decimal_and_missing_start_is_zero():
    rend = renderer(abstract(0, lvl(0, fmt=None, start=None)), num(1, 0))
    assert run(rend, [(1, 0), (1, 0)]) == ["0.", "1."]


def test_bullet_level_has_bullet_kind():
    rend = renderer(abstract(0, lvl(0, fmt="bullet", text="•")), num(1, 0))
    label = rend.next_label(1, 0)
    assert label.kind == "bullet" and label.text == "•"


def test_num_id_zero_missing_num_and_missing_abstract_have_no_label():
    rend = renderer(abstract(0, lvl(0)), num(1, 0) + num(2, 9))
    assert rend.next_label(0, 0) is None
    assert rend.next_label(5, 0) is None
    assert rend.next_label(2, 0) is None


def test_missing_numbering_part_has_no_labels():
    assert NumberingRenderer(None).next_label(1, 0) is None


def test_missing_level_definition_has_no_label():
    rend = renderer(abstract(0, lvl(0)), num(1, 0))
    assert rend.next_label(1, 3) is None


def test_format_number_letters_and_roman_edges():
    assert format_number(27, "upperLetter") == ("AA", False)
    assert format_number(28, "lowerLetter") == ("bb", False)
    assert format_number(4, "lowerRoman") == ("iv", False)
    assert format_number(0, "lowerRoman") == ("", False)
    assert format_number(1994, "upperRoman") == ("MCMXCIV", False)
    assert format_number(7, "decimalZero") == ("07", False)
    assert format_number(-3, "upperLetter") == ("", False)


def test_unknown_format_renders_decimal_and_is_reported():
    assert format_number(5, "cardinalText") == ("5", True)
    rend = renderer(abstract(0, lvl(0, fmt="ideographTraditional")), num(1, 0))
    assert run(rend, [(1, 0)]) == ["1."]
    assert rend.approx_formats == {"ideographTraditional"}


def test_is_legal_forces_decimal_for_parent_references():
    levels = lvl(0, fmt="upperRoman") + lvl(1, text="%1.%2", extra="<w:isLgl/>")
    rend = renderer(abstract(0, levels), num(1, 0))
    assert run(rend, [(1, 0), (1, 1)]) == ["I.", "1.1"]


def test_ilvl_for_style_reads_abstract_level_pstyle():
    levels = lvl(0) + lvl(1, extra='<w:pStyle w:val="ListLevel2"/>')
    rend = renderer(abstract(0, levels), num(1, 0))
    assert rend.ilvl_for_style(1, "ListLevel2") == 1
    assert rend.ilvl_for_style(1, "Other") is None
