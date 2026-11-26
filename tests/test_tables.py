from docx import Document
from docx.oxml.ns import qn

from convert_support import (
    body_text,
    make_borderless_table_pdf,
    make_ruled_table_pdf,
    table_count,
)
from docx_healer import HealerConfig, heal


def test_ruled_table_survives_the_full_pipeline(tmp_path):
    pdf = make_ruled_table_pdf(tmp_path / "in.pdf")

    heal(pdf, tmp_path / "out.docx", HealerConfig(verbose=False))

    assert table_count(tmp_path / "out.docx") >= 1
    text = body_text(tmp_path / "out.docx")
    order = [text.index(word) for word in ("Region", "North")]
    assert order == sorted(order)
    for word in ("Revenue", "Growth", "120", "4%"):
        assert word in text


def test_borderless_data_table_keeps_every_cell_in_order(tmp_path):
    pdf = make_borderless_table_pdf(tmp_path / "in.pdf")

    heal(pdf, tmp_path / "out.docx", HealerConfig(verbose=False))

    text = body_text(tmp_path / "out.docx")
    order = [text.index(word) for word in ("Region", "North", "South")]
    assert order == sorted(order)
    for word in ("Revenue", "Growth", "120", "95", "-2%"):
        assert word in text


def test_cell_borders_are_not_dropped_for_ruled_tables(tmp_path):
    pdf = make_ruled_table_pdf(tmp_path / "in.pdf")

    heal(pdf, tmp_path / "out.docx", HealerConfig(verbose=False))

    body = Document(str(tmp_path / "out.docx")).element.body
    visible = [
        side
        for borders in body.iter(qn("w:tcBorders"))
        for side in borders
        if side.get(qn("w:val")) not in (None, "nil", "none")
    ]
    assert visible
