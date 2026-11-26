import logging

from convert_support import PAGE_TEXTS, fail_parse
from docx_healer import HealerConfig, heal
from factories import make_pdf


def test_verbose_run_reports_each_page(tmp_path, capsys):
    pdf = make_pdf(tmp_path / "in.pdf", PAGE_TEXTS)

    heal(pdf, tmp_path / "out.docx", HealerConfig(verbose=True))

    lines = capsys.readouterr().out.splitlines()
    for number in (1, 2, 3):
        assert any(f"page {number}/3" in line.lower() for line in lines)


def test_quiet_run_prints_nothing(tmp_path, capsys):
    pdf = make_pdf(tmp_path / "in.pdf", PAGE_TEXTS)

    heal(pdf, tmp_path / "out.docx", HealerConfig(verbose=False))

    assert capsys.readouterr().out == ""


def test_conversion_leaves_the_root_logger_as_it_found_it(tmp_path, monkeypatch):
    pdf = make_pdf(tmp_path / "in.pdf", PAGE_TEXTS)
    fail_parse(monkeypatch, {1})
    root = logging.getLogger()
    level, handlers = root.level, list(root.handlers)

    heal(pdf, tmp_path / "out.docx", HealerConfig(verbose=True))

    assert root.level == level
    assert root.handlers == handlers
