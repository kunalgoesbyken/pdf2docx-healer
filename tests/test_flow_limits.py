import logging

import pymupdf
import pytest
from docx import Document

from docx_healer.flow_model import FlowLimits
from flow_fixtures import LEFT, convert, grid, new_page, png_bytes, put, tables, texts


def _warnings(caplog, name):
    return [r for r in caplog.records if name in r.getMessage()]


@pytest.mark.parametrize("entries,limit,warned", [(6, 3, True), (3, 3, False)])
def test_outline_cap_warns_only_when_exceeded(tmp_path, caplog, entries, limit, warned):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 100, "Heading One", 14, "hebo")
        put(page, LEFT, 140, "Some body text that goes on and on for a while here.")
        doc.set_toc([[1, f"Entry {n}", 1] for n in range(entries)])

    with caplog.at_level(logging.WARNING, logger="docx_healer"):
        convert(tmp_path, build, limits=FlowLimits(max_outline_entries=limit))

    assert bool(_warnings(caplog, "max_outline_entries")) is warned


@pytest.mark.parametrize("budget,kept", [(1, 1), (2, 2)])
def test_image_cap_keeps_the_first_images_and_warns_once(tmp_path, caplog, budget, kept):
    def build(doc):
        page = new_page(doc)
        put(page, LEFT, 60, "Two pictures follow this sentence on the page.")
        page.insert_image(pymupdf.Rect(72, 100, 272, 200), stream=png_bytes(100, 50))
        page.insert_image(pymupdf.Rect(72, 300, 272, 400), stream=png_bytes(110, 55))

    with caplog.at_level(logging.WARNING, logger="docx_healer"):
        out = convert(tmp_path, build, limits=FlowLimits(max_images=budget))

    assert len(Document(str(out)).inline_shapes) == kept
    assert len(_warnings(caplog, "max_images")) == (1 if budget < 2 else 0)


def test_table_time_budget_falls_back_to_plain_paragraphs(tmp_path, caplog):
    def build(doc):
        grid(new_page(doc), LEFT, 100, [100, 100], 26, [["Name", "Value"], ["a", "1"]])

    with caplog.at_level(logging.WARNING, logger="docx_healer"):
        out = convert(tmp_path, build, limits=FlowLimits(table_seconds=0.0))

    assert tables(out) == []
    assert "Name" in " ".join(texts(out))
    assert _warnings(caplog, "table_seconds")


def test_table_cell_cap_leaves_the_text_as_paragraphs(tmp_path):
    def build(doc):
        grid(new_page(doc), LEFT, 100, [100, 100, 100], 26, [["A", "B", "C"], ["d", "e", "f"]])

    capped = convert(tmp_path, build, name="capped", limits=FlowLimits(max_table_cells=5))
    exact = convert(tmp_path, build, name="exact", limits=FlowLimits(max_table_cells=6))

    assert tables(capped) == []
    assert len(tables(exact)) == 1
