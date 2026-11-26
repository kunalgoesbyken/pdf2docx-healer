import pytest
from docx import Document

from convert_support import (
    PAGE_TEXTS,
    body_paragraphs,
    body_text,
    fail_document_analysis,
    fail_make,
    fail_parse,
    make_encrypted_pdf,
    make_image_page_pdf,
    make_zero_page_pdf,
    picture_count,
)
from docx_healer import HealerConfig, InvalidInputError, heal
from docx_healer.ocr_pipeline import pymupdf
from factories import make_pdf

LEFTOVER = "Partially written content"


def quiet(**overrides) -> HealerConfig:
    return HealerConfig(verbose=False, **overrides)


def positions(path, words):
    text = body_text(path)
    return [text.index(word) for word in words]


def section_count(path) -> int:
    return len(Document(str(path)).sections)


def test_parse_error_on_the_first_page_keeps_its_text_in_place(tmp_path, monkeypatch, caplog):
    pdf = make_pdf(tmp_path / "in.pdf", PAGE_TEXTS)
    fail_parse(monkeypatch, {0})

    with caplog.at_level("WARNING", logger="docx_healer"):
        heal(pdf, tmp_path / "out.docx", quiet())

    found = positions(tmp_path / "out.docx", PAGE_TEXTS)
    assert found == sorted(found)
    assert body_text(tmp_path / "out.docx").count(PAGE_TEXTS[0]) == 1
    warned = [r for r in caplog.records if r.name == "docx_healer" and "Page 1" in r.getMessage()]
    assert len(warned) == 1
    assert section_count(tmp_path / "out.docx") == 3


def test_make_error_on_the_last_page_replaces_the_partial_page(tmp_path, monkeypatch, caplog):
    pdf = make_pdf(tmp_path / "in.pdf", PAGE_TEXTS)
    fail_make(monkeypatch, {2})

    with caplog.at_level("WARNING", logger="docx_healer"):
        heal(pdf, tmp_path / "out.docx", quiet())

    text = body_text(tmp_path / "out.docx")
    assert [text.count(word) for word in PAGE_TEXTS] == [1, 1, 1]
    found = positions(tmp_path / "out.docx", PAGE_TEXTS)
    assert found == sorted(found)
    assert sum("Page 3" in r.getMessage() for r in caplog.records) == 1
    assert section_count(tmp_path / "out.docx") == 3


def test_make_error_in_the_middle_keeps_neighbouring_pages_intact(tmp_path, monkeypatch):
    pdf = make_pdf(tmp_path / "in.pdf", PAGE_TEXTS)
    fail_make(monkeypatch, {1}, leftover=LEFTOVER)

    heal(pdf, tmp_path / "out.docx", quiet())

    text = body_text(tmp_path / "out.docx")
    assert LEFTOVER not in text
    assert [text.count(word) for word in PAGE_TEXTS] == [1, 1, 1]
    found = positions(tmp_path / "out.docx", PAGE_TEXTS)
    assert found == sorted(found)
    assert section_count(tmp_path / "out.docx") == 3


def test_every_page_failing_still_produces_every_page(tmp_path, monkeypatch, caplog):
    pdf = make_pdf(tmp_path / "in.pdf", PAGE_TEXTS)
    fail_parse(monkeypatch, {0, 1, 2})

    with caplog.at_level("WARNING", logger="docx_healer"):
        heal(pdf, tmp_path / "out.docx", quiet())

    text = body_text(tmp_path / "out.docx")
    assert [text.count(word) for word in PAGE_TEXTS] == [1, 1, 1]
    assert section_count(tmp_path / "out.docx") == 3
    assert sum("rebuilt" in r.getMessage() for r in caplog.records) == 3


def test_parse_and_make_failures_on_different_pages_are_both_recovered(tmp_path, monkeypatch):
    pdf = make_pdf(tmp_path / "in.pdf", PAGE_TEXTS)
    fail_parse(monkeypatch, {0})
    fail_make(monkeypatch, {2})

    heal(pdf, tmp_path / "out.docx", quiet())

    text = body_text(tmp_path / "out.docx")
    assert [text.count(word) for word in PAGE_TEXTS] == [1, 1, 1]


def test_image_only_page_that_fails_keeps_its_picture(tmp_path, monkeypatch):
    pdf = make_image_page_pdf(tmp_path / "in.pdf", text_pages_before=1)
    fail_parse(monkeypatch, {1})

    heal(pdf, tmp_path / "out.docx", quiet())

    assert picture_count(tmp_path / "out.docx") == 1
    assert body_text(tmp_path / "out.docx").count(PAGE_TEXTS[0]) == 1


def test_document_analysis_failure_rebuilds_every_page(tmp_path, monkeypatch, caplog):
    pdf = make_pdf(tmp_path / "in.pdf", PAGE_TEXTS)
    fail_document_analysis(monkeypatch)

    with caplog.at_level("WARNING", logger="docx_healer"):
        heal(pdf, tmp_path / "out.docx", quiet())

    text = body_text(tmp_path / "out.docx")
    assert [text.count(word) for word in PAGE_TEXTS] == [1, 1, 1]
    assert sum("rebuilt" in r.getMessage() for r in caplog.records) == 3


def test_page_that_cannot_even_be_read_is_left_blank_without_aborting(
    tmp_path, monkeypatch, caplog
):
    pdf = make_pdf(tmp_path / "in.pdf", PAGE_TEXTS)
    fail_parse(monkeypatch, {1})
    original = pymupdf.Page.get_text

    def get_text(self, *args, **kwargs):
        if self.number == 1 and args and args[0] == "dict":
            raise RuntimeError("unreadable page")
        return original(self, *args, **kwargs)

    monkeypatch.setattr(pymupdf.Page, "get_text", get_text)

    with caplog.at_level("ERROR", logger="docx_healer"):
        heal(pdf, tmp_path / "out.docx", quiet())

    text = body_text(tmp_path / "out.docx")
    assert PAGE_TEXTS[0] in text and PAGE_TEXTS[2] in text
    assert PAGE_TEXTS[1] not in text
    assert section_count(tmp_path / "out.docx") == 3
    assert sum("left it blank" in r.getMessage() for r in caplog.records) == 1


def test_recovery_disabled_keeps_the_upstream_behaviour(tmp_path, monkeypatch, caplog):
    pdf = make_pdf(tmp_path / "in.pdf", PAGE_TEXTS)
    fail_parse(monkeypatch, {1})

    with caplog.at_level("WARNING"):
        heal(pdf, tmp_path / "out.docx", quiet(recover_pages=False))

    text = body_text(tmp_path / "out.docx")
    assert PAGE_TEXTS[1] not in text
    assert PAGE_TEXTS[0] in text and PAGE_TEXTS[2] in text
    assert "rebuilt" not in caplog.text


def test_recovery_disabled_does_not_swallow_document_analysis_failure(tmp_path, monkeypatch):
    pdf = make_pdf(tmp_path / "in.pdf", PAGE_TEXTS)
    fail_document_analysis(monkeypatch)

    with pytest.raises(RuntimeError, match="document analysis"):
        heal(pdf, tmp_path / "out.docx", quiet(recover_pages=False))

    assert not (tmp_path / "out.docx").exists()


@pytest.mark.parametrize("recover", [True, False])
def test_password_protected_pdf_is_rejected_clearly(tmp_path, recover):
    pdf = make_encrypted_pdf(tmp_path / "in.pdf")

    with pytest.raises(InvalidInputError, match="password"):
        heal(pdf, tmp_path / "out.docx", quiet(recover_pages=recover))

    assert not (tmp_path / "out.docx").exists()


@pytest.mark.parametrize("recover", [True, False])
def test_pdf_without_pages_is_rejected_clearly(tmp_path, recover):
    pdf = make_zero_page_pdf(tmp_path / "in.pdf")

    with pytest.raises(InvalidInputError, match="no pages"):
        heal(pdf, tmp_path / "out.docx", quiet(recover_pages=recover))

    assert not (tmp_path / "out.docx").exists()


def test_healthy_pdf_is_untouched_by_recovery(tmp_path, caplog):
    pdf = make_pdf(tmp_path / "in.pdf", PAGE_TEXTS)

    with caplog.at_level("WARNING", logger="docx_healer"):
        heal(pdf, tmp_path / "out.docx", quiet())

    assert "rebuilt" not in caplog.text
    paragraphs = [p for p in body_paragraphs(tmp_path / "out.docx") if p]
    assert paragraphs == PAGE_TEXTS


def test_five_hundred_failing_pages_are_all_recovered(tmp_path, monkeypatch):
    pdf = make_pdf(tmp_path / "in.pdf", [f"Mango{index} orchard" for index in range(500)])
    fail_parse(monkeypatch, set(range(500)))

    heal(pdf, tmp_path / "out.docx", quiet())

    paragraphs = [p for p in body_paragraphs(tmp_path / "out.docx") if p]
    assert paragraphs == [f"Mango{index} orchard" for index in range(500)]
