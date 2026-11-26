from importlib import metadata

import pytest

import docx_healer
from docx_healer.cli import EXIT_FAILURE, EXIT_OK, main
from factories import TEXT_PAGE, make_pdf


def test_successful_run_exits_zero_and_writes_the_output(tmp_path):
    pdf = make_pdf(tmp_path / "in.pdf", [TEXT_PAGE])
    output = tmp_path / "out.docx"

    assert main([str(pdf), "-o", str(output), "-q"]) == EXIT_OK
    assert output.exists()


def test_missing_input_exits_non_zero_with_a_message(tmp_path, capsys):
    assert main([str(tmp_path / "missing.pdf"), "-q"]) == EXIT_FAILURE

    assert "PDF file not found" in capsys.readouterr().err


def test_non_pdf_input_exits_non_zero(tmp_path, capsys):
    not_pdf = tmp_path / "notes.pdf"
    not_pdf.write_text("hello")

    assert main([str(not_pdf), "-q"]) == EXIT_FAILURE
    assert "Not a PDF" in capsys.readouterr().err


def test_conversion_failure_exits_non_zero(tmp_path, monkeypatch, capsys):
    pdf = make_pdf(tmp_path / "in.pdf", [TEXT_PAGE])

    def failing_convert(*args, **kwargs):
        raise RuntimeError("converter crashed")

    monkeypatch.setattr("docx_healer.converter._convert", failing_convert)

    assert main([str(pdf), "-q"]) == EXIT_FAILURE
    assert "converter crashed" in capsys.readouterr().err


@pytest.mark.parametrize("flag", [["--ocr-threshold", "2"], ["--ocr-dpi", "0"]])
def test_invalid_numeric_options_exit_with_usage_error(tmp_path, flag):
    pdf = make_pdf(tmp_path / "in.pdf", [TEXT_PAGE])

    with pytest.raises(SystemExit) as raised:
        main([str(pdf), *flag])

    assert raised.value.code == 2


def test_version_matches_installed_metadata():
    assert docx_healer.__version__ == metadata.version("pdf2docx-healer")


def test_version_flag_prints_the_version_and_exits_zero(capsys):
    with pytest.raises(SystemExit) as raised:
        main(["--version"])

    assert raised.value.code == 0
    assert capsys.readouterr().out.strip() == f"pdf2docx-heal {docx_healer.__version__}"
