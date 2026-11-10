from pathlib import Path

import pytest

from docx_healer import HealerConfig, heal, ocr_pipeline
from docx_healer.errors import OcrUnavailableError
from docx_healer.ocr_pipeline import ocr_pdf_file, prepare_ocr_copy
from factories import TEXT_PAGE, make_pdf


def quiet_config(**overrides) -> HealerConfig:
    return HealerConfig(verbose=False, **overrides)


def test_input_pdf_survives_when_threshold_is_below_the_default(tmp_path, fake_ocr):
    pdf = make_pdf(tmp_path / "in.pdf", [TEXT_PAGE, TEXT_PAGE, TEXT_PAGE, None])
    output = tmp_path / "out.docx"

    heal(pdf, output, quiet_config(ocr_enabled=True, ocr_threshold=0.2))

    assert pdf.exists()
    assert output.exists()
    assert fake_ocr == ["ocr"]


def test_input_pdf_survives_when_ocr_is_not_needed(tmp_path, fake_ocr):
    pdf = make_pdf(tmp_path / "in.pdf", [TEXT_PAGE, TEXT_PAGE])

    heal(pdf, tmp_path / "out.docx", quiet_config(ocr_enabled=True))

    assert pdf.exists()
    assert fake_ocr == []


def test_ocr_copy_is_created_then_removed_and_never_the_input(
    tmp_path, scratch_tempdir, fake_ocr, monkeypatch
):
    pdf = make_pdf(tmp_path / "in.pdf", [None, None])
    created = []
    original = ocr_pipeline.ocr_pdf_file

    def recording_ocr_pdf_file(*args, **kwargs):
        path = original(*args, **kwargs)
        created.append(Path(path))
        assert Path(path).exists()
        return path

    monkeypatch.setattr(ocr_pipeline, "ocr_pdf_file", recording_ocr_pdf_file)

    heal(pdf, tmp_path / "out.docx", quiet_config(ocr_enabled=True))

    assert len(created) == 1
    assert created[0] != pdf
    assert not created[0].exists()
    assert pdf.exists()
    assert list(scratch_tempdir.iterdir()) == []


def test_ocr_copy_is_removed_when_conversion_fails(
    tmp_path, scratch_tempdir, fake_ocr, monkeypatch
):
    pdf = make_pdf(tmp_path / "in.pdf", [None])

    def failing_convert(*args, **kwargs):
        raise RuntimeError("conversion exploded")

    monkeypatch.setattr("docx_healer.converter._convert", failing_convert)

    with pytest.raises(RuntimeError, match="conversion exploded"):
        heal(pdf, tmp_path / "out.docx", quiet_config(ocr_enabled=True))

    assert pdf.exists()
    assert list(scratch_tempdir.iterdir()) == []


def test_prepare_ocr_copy_returns_none_below_threshold(tmp_path, fake_ocr):
    pdf = make_pdf(tmp_path / "in.pdf", [TEXT_PAGE, TEXT_PAGE, TEXT_PAGE, None])

    assert prepare_ocr_copy(pdf, threshold=0.3) is None
    copy = prepare_ocr_copy(pdf, threshold=0.2)

    assert copy is not None
    assert Path(copy) != pdf
    Path(copy).unlink()


def test_ocr_pdf_file_always_returns_a_new_file(tmp_path, fake_ocr):
    pdf = make_pdf(tmp_path / "in.pdf", [TEXT_PAGE])

    result = Path(ocr_pdf_file(pdf))

    assert result != pdf
    assert result.exists()
    result.unlink()
    assert pdf.exists()


def test_missing_tesseract_warns_and_still_converts(tmp_path, monkeypatch, caplog):
    pdf = make_pdf(tmp_path / "in.pdf", [None, None])

    def unavailable():
        raise OcrUnavailableError("Tesseract OCR is not available")

    monkeypatch.setattr(ocr_pipeline, "require_tesseract", unavailable)

    with caplog.at_level("WARNING", logger="docx_healer"):
        heal(pdf, tmp_path / "out.docx", quiet_config(ocr_enabled=True))

    assert (tmp_path / "out.docx").exists()
    assert pdf.exists()
    assert "Skipping OCR" in caplog.text


def test_failed_page_ocr_warns_keeps_the_page_and_cleans_up(
    tmp_path, scratch_tempdir, fake_ocr, monkeypatch, caplog
):
    pdf = make_pdf(tmp_path / "in.pdf", [None])

    def broken(self, *args, **kwargs):
        raise RuntimeError("tessdata for 'xyz' not found")

    monkeypatch.setattr(ocr_pipeline.pymupdf.Pixmap, "pdfocr_tobytes", broken)

    with caplog.at_level("WARNING", logger="docx_healer"):
        heal(pdf, tmp_path / "out.docx", quiet_config(ocr_enabled=True, ocr_lang="xyz"))

    assert (tmp_path / "out.docx").exists()
    assert "OCR failed for page 1" in caplog.text
    assert list(scratch_tempdir.iterdir()) == []
