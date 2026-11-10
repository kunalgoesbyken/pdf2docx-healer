import pytest

from docx_healer import HealerConfig, InvalidInputError, heal
from factories import TEXT_PAGE, make_pdf


def quiet() -> HealerConfig:
    return HealerConfig(verbose=False)


def test_missing_pdf_raises_file_not_found(tmp_path):
    with pytest.raises(FileNotFoundError):
        heal(tmp_path / "missing.pdf", tmp_path / "out.docx", quiet())


def test_non_pdf_file_is_rejected(tmp_path):
    text = tmp_path / "notes.pdf"
    text.write_text("just some text")

    with pytest.raises(InvalidInputError, match="Not a PDF"):
        heal(text, tmp_path / "out.docx", quiet())


def test_directory_input_is_rejected(tmp_path):
    with pytest.raises(InvalidInputError, match="Not a file"):
        heal(tmp_path, tmp_path / "out.docx", quiet())


def test_missing_output_directory_is_rejected(tmp_path):
    pdf = make_pdf(tmp_path / "in.pdf", [TEXT_PAGE])

    with pytest.raises(InvalidInputError, match="Output directory"):
        heal(pdf, tmp_path / "nope" / "out.docx", quiet())


def test_output_directory_path_is_rejected(tmp_path):
    pdf = make_pdf(tmp_path / "in.pdf", [TEXT_PAGE])

    with pytest.raises(InvalidInputError, match="is a directory"):
        heal(pdf, tmp_path, quiet())


def test_output_may_not_overwrite_the_input(tmp_path):
    pdf = make_pdf(tmp_path / "in.pdf", [TEXT_PAGE])
    original = pdf.read_bytes()

    with pytest.raises(InvalidInputError, match="overwrite the input"):
        heal(pdf, pdf, quiet())

    assert pdf.read_bytes() == original


def test_default_output_sits_next_to_the_input(tmp_path):
    pdf = make_pdf(tmp_path / "report.pdf", [TEXT_PAGE])

    result = heal(pdf, config=quiet())

    assert result == str(tmp_path / "report.docx")
    assert (tmp_path / "report.docx").exists()


@pytest.mark.parametrize(
    "overrides",
    [{"ocr_threshold": -0.1}, {"ocr_threshold": 1.5}, {"ocr_dpi": 0}, {"ocr_lang": " "}],
)
def test_invalid_config_is_rejected(overrides):
    with pytest.raises(InvalidInputError):
        HealerConfig(**overrides)


def test_failed_conversion_leaves_no_output_and_keeps_existing_file(tmp_path, monkeypatch):
    pdf = make_pdf(tmp_path / "in.pdf", [TEXT_PAGE])
    output = tmp_path / "out.docx"
    output.write_bytes(b"previous result")

    def failing_convert(*args, **kwargs):
        raise RuntimeError("boom")

    monkeypatch.setattr("docx_healer.converter._convert", failing_convert)

    with pytest.raises(RuntimeError):
        heal(pdf, output, quiet())

    assert output.read_bytes() == b"previous result"
    assert sorted(p.name for p in tmp_path.iterdir()) == ["in.pdf", "out.docx"]


def test_successful_conversion_leaves_no_staging_files(tmp_path):
    pdf = make_pdf(tmp_path / "in.pdf", [TEXT_PAGE])

    heal(pdf, tmp_path / "out.docx", quiet())

    assert sorted(p.name for p in tmp_path.iterdir()) == ["in.pdf", "out.docx"]
