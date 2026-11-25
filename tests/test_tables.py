import pytest
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

from convert_support import (
    body_text,
    make_borderless_table_pdf,
    make_ruled_table_pdf,
    table_count,
)
from docx_healer import HealerConfig, heal
from docx_healer.tables import fix_docx_tables
from factories import make_docx, make_pdf


def borderless_single_cell(doc):
    table = doc.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "Text held by a single cell table"


def nested_tables(doc):
    outer = doc.add_table(rows=1, cols=1)
    cell = outer.cell(0, 0)
    inner = cell.add_table(rows=2, cols=2)
    for index, value in enumerate(["A", "B", "C", "D"]):
        inner.cell(index // 2, index % 2).text = value
    cell.add_paragraph("after")


def exact_height_row(doc):
    table = doc.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "Fixed height row"
    row = table.rows[0]
    row.height = Pt(36)


def two_by_three(doc):
    table = doc.add_table(rows=2, cols=3)
    for index, value in enumerate(["Region", "Revenue", "Growth", "North", "120", "4%"]):
        table.cell(index // 3, index % 3).text = value


@pytest.mark.parametrize(
    "build",
    [borderless_single_cell, nested_tables, exact_height_row, two_by_three],
    ids=lambda build: build.__name__,
)
def test_table_pass_leaves_every_table_shape_untouched(tmp_path, build):
    path = make_docx(tmp_path / "in.docx", build)
    pdf = make_pdf(tmp_path / "in.pdf", ["page"])
    before = path.read_bytes()

    changed = fix_docx_tables(path, pdf)

    assert changed == 0
    assert path.read_bytes() == before


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
