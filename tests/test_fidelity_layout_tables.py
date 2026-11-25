from pathlib import Path
from typing import Callable

import pytest
from pdf2docx import Converter

from convert_support import body_text, make_borderless_table_pdf, make_ruled_table_pdf, table_count
from docx_healer import HealerConfig, heal
from docx_healer.layout_tables import TableRules, layout_table_filter
from layout_support import (
    CJK_ROWS,
    NUMERIC_ROWS,
    SENTENCE_ROWS,
    make_boxed_text_pdf,
    make_callout_pdf,
    make_data_table_pdf,
    make_definition_list_pdf,
    make_marker_list_pdf,
    make_nested_pieces_pdf,
    make_rotated_label_pdf,
    make_three_column_prose_pdf,
    make_two_column_pdf,
    nested_table_count,
    table_texts,
    text_directions,
)

Maker = Callable[[Path], Path]
CONFIG = HealerConfig(verbose=False)


def convert(maker: Maker, tmp_path: Path, config: HealerConfig = CONFIG) -> Path:
    pdf = maker(tmp_path / "in.pdf")
    out = tmp_path / "out.docx"
    heal(pdf, out, config)
    return out


def order(text: str, *words: str) -> bool:
    positions = [text.index(word) for word in words]
    return positions == sorted(positions)


@pytest.mark.parametrize(
    "maker",
    [make_three_column_prose_pdf, make_definition_list_pdf, make_marker_list_pdf, make_two_column_pdf],
    ids=["three-column-prose", "definition-list", "marker-list", "two-column-prose"],
)
def test_text_layout_is_not_turned_into_a_table(maker, tmp_path):
    out = convert(maker, tmp_path)

    assert table_count(out) == 0


def test_grid_of_sentences_is_not_a_table(tmp_path):
    out = convert(lambda p: make_data_table_pdf(p, SENTENCE_ROWS, gap=150), tmp_path)

    assert table_count(out) == 0
    assert order(body_text(out), "Explain the unit", "Describe each", "Plan the next")


def test_three_column_prose_keeps_each_column_in_one_piece(tmp_path):
    out = convert(make_three_column_prose_pdf, tmp_path)

    text = body_text(out)
    assert order(text, "col0 column line 00", "col0 column line 13")


def test_text_pieces_inside_a_ruled_cell_do_not_become_a_nested_table(tmp_path):
    out = convert(make_nested_pieces_pdf, tmp_path)

    assert nested_table_count(out) == 0
    assert table_count(out) == 1
    assert "Second cell" in table_texts(out)[0]


def test_two_column_page_keeps_the_left_column_before_the_right(tmp_path):
    out = convert(make_two_column_pdf, tmp_path)

    text = body_text(out)
    assert order(text, "left column line 00", "left column line 13", "right column line 00")


def test_borderless_data_table_survives(tmp_path):
    out = convert(lambda p: make_data_table_pdf(p, NUMERIC_ROWS), tmp_path)

    assert table_count(out) == 1
    assert order(table_texts(out)[0], "Model", "alpha", "delta")
    assert "71.2" in table_texts(out)[0]


def test_cjk_borderless_data_table_survives(tmp_path):
    out = convert(lambda p: make_data_table_pdf(p, CJK_ROWS, fontname="china-s"), tmp_path)

    assert table_count(out) == 1
    assert "北方" in table_texts(out)[0]


def test_ruled_table_survives(tmp_path):
    out = convert(make_ruled_table_pdf, tmp_path)

    assert table_count(out) == 1


def test_single_cell_box_survives(tmp_path):
    out = convert(make_boxed_text_pdf, tmp_path)

    assert table_texts(out) == ["Boxed notice inside a single cell"]


def test_shaded_callout_survives_as_a_table(tmp_path):
    out = convert(make_callout_pdf, tmp_path)

    assert table_count(out) >= 1
    assert "Shaded callout first line" in "".join(table_texts(out))
    assert "Plain paragraph after the callout" not in "".join(table_texts(out))


def test_rotated_margin_label_keeps_its_text_direction(tmp_path):
    out = convert(make_rotated_label_pdf, tmp_path)

    assert text_directions(out) == 1
    assert "Margin stamp rotated" in body_text(out)


def test_switching_the_filter_off_restores_pdf2docx_tables(tmp_path):
    out = convert(make_three_column_prose_pdf, tmp_path, HealerConfig(verbose=False, fix_tables=False))

    assert table_count(out) == 1


def pdf2docx_tables(pdf: Path, out: Path, rules: TableRules) -> int:
    converter = Converter(str(pdf))
    try:
        with layout_table_filter(rules):
            converter.convert(str(out))
    finally:
        converter.close()
    return table_count(out)


def test_groups_over_the_line_budget_keep_pdf2docx_behaviour(tmp_path):
    pdf = make_three_column_prose_pdf(tmp_path / "in.pdf")

    over = pdf2docx_tables(pdf, tmp_path / "over.docx", TableRules(max_group_lines=41))
    at_limit = pdf2docx_tables(pdf, tmp_path / "at.docx", TableRules(max_group_lines=42))

    assert (over, at_limit) == (1, 0)


def test_row_and_column_minimums_are_injectable(tmp_path):
    pdf = make_data_table_pdf(tmp_path / "in.pdf", NUMERIC_ROWS[:2])

    default = pdf2docx_tables(pdf, tmp_path / "default.docx", TableRules())
    relaxed = pdf2docx_tables(pdf, tmp_path / "relaxed.docx", TableRules(min_rows=2))

    assert (default, relaxed) == (0, 1)


def test_cell_length_limit_is_injectable(tmp_path):
    pdf = make_data_table_pdf(tmp_path / "in.pdf", SENTENCE_ROWS, gap=150)

    default = pdf2docx_tables(pdf, tmp_path / "default.docx", TableRules())
    relaxed = pdf2docx_tables(pdf, tmp_path / "relaxed.docx", TableRules(max_cell_chars=40))

    assert (default, relaxed) == (0, 1)


def test_the_filter_is_removed_when_the_context_exits():
    from pdf2docx.layout.Blocks import Blocks

    before = Blocks.collect_stream_lines
    with layout_table_filter(TableRules()):
        assert Blocks.collect_stream_lines is not before
    assert Blocks.collect_stream_lines is before


def test_the_filter_does_nothing_when_pdf2docx_has_another_layout(monkeypatch):
    from pdf2docx.layout.Blocks import Blocks

    monkeypatch.delattr(Blocks, "collect_stream_lines")
    with layout_table_filter(TableRules()):
        assert not hasattr(Blocks, "collect_stream_lines")


def test_borderless_three_row_table_from_the_shared_fixture_keeps_its_text(tmp_path):
    out = convert(make_borderless_table_pdf, tmp_path)

    assert order(body_text(out), "Region", "North", "South")
