from docx import Document

from docx_healer import post_processor
from docx_healer.post_processor import post_process_docx
from factories import make_docx


def build(doc):
    doc.add_paragraph("see https://example.com now")


def test_failed_pass_is_warned_reported_and_does_not_stop_the_others(tmp_path, monkeypatch, caplog):
    path = make_docx(tmp_path / "a.docx", build)

    def broken(*args, **kwargs):
        raise ValueError("list pass blew up")

    monkeypatch.setattr(post_processor, "fix_lists_in_docx", broken)

    with caplog.at_level("WARNING", logger="docx_healer"):
        report = post_process_docx(path, verbose=False)

    assert report.failed_passes == ["lists"]
    assert report.hyperlinks_added == 1
    assert "lists pass failed" in caplog.text
    assert "list pass blew up" in caplog.text


def test_failed_pass_is_logged_even_when_not_verbose(tmp_path, monkeypatch, caplog):
    path = make_docx(tmp_path / "a.docx", build)

    def broken(*args, **kwargs):
        raise RuntimeError("fonts blew up")

    monkeypatch.setattr(post_processor, "fix_docx_fonts", broken)

    with caplog.at_level("WARNING", logger="docx_healer"):
        report = post_process_docx(path, verbose=False)

    assert report.failed_passes == ["fonts"]
    assert any(record.levelname == "WARNING" for record in caplog.records)


def test_disabled_passes_do_not_run(tmp_path, monkeypatch):
    path = make_docx(tmp_path / "a.docx", build)

    def forbidden(*args, **kwargs):
        raise AssertionError("pass should not run")

    for name in ("fix_lists_in_docx", "fix_docx_hyperlinks", "fix_docx_fonts"):
        monkeypatch.setattr(post_processor, name, forbidden)

    report = post_process_docx(path, fix_lists=False, fix_hyperlinks=False, fix_fonts=False)

    assert report.failed_passes == []
    assert Document(str(path)).paragraphs[0].text == "see https://example.com now"


def test_successful_run_reports_counts_and_no_failures(tmp_path):
    path = make_docx(tmp_path / "a.docx", build)

    report = post_process_docx(path)

    assert report.hyperlinks_added == 1
    assert report.failed_passes == []
