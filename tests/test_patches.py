import importlib

import pytest
from docx import Document

from docx_healer import HealerConfig, IncompatiblePdf2docxError, heal, patches
from docx_healer.patches import _is_list_item, list_item_patches
from factories import TEXT_PAGE, make_list_with_prices_pdf, make_pdf

BLOCKS = "pdf2docx.layout.Blocks"


@pytest.mark.parametrize(
    "text",
    ["• item", "•", "· ", "- item", "* item", "1. item", "12) item", "a. item", "(iv) item",
     "■ item", " item", " item", "– item"],
)
def test_list_markers_are_detected(text):
    assert _is_list_item(text)


@pytest.mark.parametrize(
    "text",
    ["", "   ", "Hello world", "-5 degrees", "*note", "1.5 million", "2024", "-", "1.", "(a)",
     "LLC. Enter the name", "1.2. item", "iiii. text", "Fig. 3. Results", "e.g. a thing", "U.S. army"],
)
def test_non_list_text_is_not_detected(text):
    assert not _is_list_item(text)


@pytest.mark.parametrize(
    "value",
    [[["1. a", "b"], ["2. c", "d"]], [], None, 5, b"item", object(), 1.5, ("•",)],
)
def test_non_string_input_never_raises(value):
    assert not _is_list_item(value)


def test_private_use_bullet_is_returned_as_written():
    assert _is_list_item(" item") == ""


def test_a_failing_marker_parser_never_costs_a_page(monkeypatch):
    blocks = importlib.import_module(BLOCKS)

    def broken(*args, **kwargs):
        raise RuntimeError("parser blew up")

    monkeypatch.setattr(patches, "parse_marker", broken, raising=False)

    with list_item_patches():
        assert blocks.is_list_item("• item") is None

def test_patches_apply_inside_the_context_and_are_restored():
    blocks = importlib.import_module(BLOCKS)
    original_is_list_item = blocks.is_list_item
    original_collect = blocks.Blocks.collect_stream_lines

    with list_item_patches():
        assert blocks.is_list_item is _is_list_item
        assert blocks.Blocks.collect_stream_lines is not original_collect

    assert blocks.is_list_item is original_is_list_item
    assert blocks.Blocks.collect_stream_lines is original_collect


def test_patches_are_restored_after_an_error():
    blocks = importlib.import_module(BLOCKS)
    original_is_list_item = blocks.is_list_item

    with pytest.raises(RuntimeError):
        with list_item_patches():
            raise RuntimeError("fail inside")

    assert blocks.is_list_item is original_is_list_item


def test_collect_stream_lines_is_forced_to_keep_lists_out_of_tables(monkeypatch):
    blocks = importlib.import_module(BLOCKS)
    seen = []

    def recording(self, potential_shadings, line_separate_threshold, **kwargs):
        seen.append(kwargs)

    monkeypatch.setattr(blocks.Blocks, "collect_stream_lines", recording)

    with list_item_patches():
        blocks.Blocks.collect_stream_lines(object(), [], 5.0)

    assert seen == [{"list_not_table": True}]


def test_incompatible_collect_signature_fails_loudly_without_patching(monkeypatch):
    blocks = importlib.import_module(BLOCKS)
    original_is_list_item = blocks.is_list_item

    def old_signature(self, potential_shadings, line_separate_threshold):
        return []

    monkeypatch.setattr(blocks.Blocks, "collect_stream_lines", old_signature)

    with pytest.raises(IncompatiblePdf2docxError, match="pdf2docx"):
        with list_item_patches():
            pass

    assert blocks.is_list_item is original_is_list_item


def test_missing_is_list_item_fails_loudly(monkeypatch):
    blocks = importlib.import_module(BLOCKS)
    monkeypatch.delattr(blocks, "is_list_item")

    with pytest.raises(IncompatiblePdf2docxError, match="is_list_item"):
        with list_item_patches():
            pass


def test_heal_surfaces_incompatibility_and_writes_nothing(tmp_path, monkeypatch):
    blocks = importlib.import_module(BLOCKS)
    monkeypatch.delattr(blocks, "is_list_item")
    pdf = make_pdf(tmp_path / "in.pdf", [TEXT_PAGE])

    with pytest.raises(IncompatiblePdf2docxError):
        heal(pdf, tmp_path / "out.docx", HealerConfig(verbose=False))

    assert not (tmp_path / "out.docx").exists()


def test_bullet_rows_become_a_list_instead_of_a_table(tmp_path):
    pdf = make_list_with_prices_pdf(tmp_path / "in.pdf")
    output = tmp_path / "out.docx"

    heal(pdf, output, HealerConfig(verbose=False))

    doc = Document(str(output))
    table_rows = sum(len(table.rows) for table in doc.tables)
    bullets = [p for p in doc.paragraphs if p.style.name == "List Bullet"]
    assert len(bullets) >= 3
    assert table_rows <= 1


def test_patches_do_not_leak_into_later_conversions(tmp_path):
    pdf = make_list_with_prices_pdf(tmp_path / "in.pdf")

    heal(pdf, tmp_path / "healed.docx", HealerConfig(verbose=False))
    heal(pdf, tmp_path / "plain.docx", HealerConfig(verbose=False, fix_lists=False))

    plain = Document(str(tmp_path / "plain.docx"))
    assert sum(len(table.rows) for table in plain.tables) == 4
