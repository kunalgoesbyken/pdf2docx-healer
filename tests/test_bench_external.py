import sys
import time
from types import SimpleNamespace

from bench import external
from bench.report import aggregate, merge_rows
from bench.runner import run_isolated
from bench.worker import TOOLS, run_tool

DONE = "done: in.pdf -> out.docx  pages=5 ok=3 failed=2  in 1.20s"


def fake_run(tmp_path, returncode=0, stderr="", write=True, name="out.docx"):
    def run(argv, **kwargs):
        if write:
            target = tmp_path / name
            target.write_bytes(b"docx")
        return SimpleNamespace(returncode=returncode, stdout="", stderr=stderr)

    return run


def test_logged_failures_reads_the_done_line():
    assert external.logged_failures("noise\n" + DONE + "\n") == 2


def test_logged_failures_is_none_without_a_done_line():
    assert external.logged_failures("INFO: starting\nERROR: boom") is None


def test_plus_tool_is_registered():
    assert "pdf2docx-plus" in TOOLS
    assert "libreoffice" in TOOLS


def test_plus_missing_executable_is_unavailable(monkeypatch):
    monkeypatch.setattr(external, "find_plus", lambda: None)
    result = run_tool(external.convert_plus, "in.pdf", "out.docx")
    assert result["status"] == "unavailable"
    assert "BENCH_PDF2DOCX_PLUS" in result["error"]


def test_plus_page_failures_with_exit_code_one_still_count_as_output(monkeypatch, tmp_path):
    monkeypatch.setattr(external, "find_plus", lambda: "plus")
    monkeypatch.setattr(external.subprocess, "run", fake_run(tmp_path, 1, DONE))
    result = run_tool(external.convert_plus, "in.pdf", str(tmp_path / "out.docx"))
    assert result["status"] == "ok"
    assert result["note"] == "logged page failures: 2"


def test_plus_exit_zero_without_output_is_an_error(monkeypatch, tmp_path):
    monkeypatch.setattr(external, "find_plus", lambda: "plus")
    monkeypatch.setattr(external.subprocess, "run", fake_run(tmp_path, 0, "", write=False))
    result = run_tool(external.convert_plus, "in.pdf", str(tmp_path / "out.docx"))
    assert result["status"] == "error"


def test_plus_crash_without_output_reports_exit_code_and_stderr_tail(monkeypatch, tmp_path):
    monkeypatch.setattr(external, "find_plus", lambda: "plus")
    monkeypatch.setattr(
        external.subprocess, "run", fake_run(tmp_path, 3, "Traceback\nboom", write=False)
    )
    result = run_tool(external.convert_plus, "in.pdf", str(tmp_path / "out.docx"))
    assert result["status"] == "error"
    assert "exit code 3" in result["error"] and "boom" in result["error"]


def test_plus_command_uses_default_profile_and_output_flag(monkeypatch, tmp_path):
    seen = {}

    def run(argv, **kwargs):
        seen["argv"] = argv
        (tmp_path / "out.docx").write_bytes(b"x")
        return SimpleNamespace(returncode=0, stdout="", stderr="")

    monkeypatch.setattr(external, "find_plus", lambda: "plus-exe")
    monkeypatch.setattr(external.subprocess, "run", run)
    external.convert_plus("in.pdf", str(tmp_path / "out.docx"))
    assert seen["argv"] == ["plus-exe", "convert", "in.pdf", "--output", str(tmp_path / "out.docx")]


def test_libreoffice_missing_is_unavailable(monkeypatch):
    monkeypatch.setattr(external, "find_soffice", lambda: None)
    result = run_tool(external.convert_libreoffice, "in.pdf", "out.docx")
    assert result["status"] == "unavailable"


def test_libreoffice_moves_the_converted_file_to_the_requested_path(monkeypatch, tmp_path):
    seen = {}

    def run(argv, **kwargs):
        seen["argv"] = argv
        outdir = argv[argv.index("--outdir") + 1]
        (tmp_path / "work").mkdir(exist_ok=True)
        with open(f"{outdir}/in.docx", "wb") as handle:
            handle.write(b"docx")
        return SimpleNamespace(returncode=0, stdout="", stderr="")

    monkeypatch.setattr(external, "find_soffice", lambda: "soffice")
    monkeypatch.setattr(external.subprocess, "run", run)
    target = tmp_path / "result.docx"
    external.convert_libreoffice(str(tmp_path / "in.pdf"), str(target))
    assert target.read_bytes() == b"docx"
    assert "--infilter=writer_pdf_import" in seen["argv"]
    assert seen["argv"][seen["argv"].index("--convert-to") + 1] == "docx"
    assert "--headless" in seen["argv"]


def test_libreoffice_without_output_is_an_error(monkeypatch, tmp_path):
    monkeypatch.setattr(external, "find_soffice", lambda: "soffice")
    monkeypatch.setattr(
        external.subprocess,
        "run",
        lambda argv, **kw: SimpleNamespace(returncode=0, stdout="", stderr=""),
    )
    result = run_tool(
        external.convert_libreoffice, str(tmp_path / "in.pdf"), str(tmp_path / "o.docx")
    )
    assert result["status"] == "error"


def test_timeout_kills_grandchildren_too(tmp_path):
    ticks = tmp_path / "ticks.txt"
    grandchild = (
        f"import time\nwhile True:\n    open({str(ticks)!r}, 'a').write('x')\n    time.sleep(0.2)\n"
    )
    parent = (
        "import subprocess, sys, time\n"
        f"subprocess.Popen([sys.executable, '-c', {grandchild!r}])\n"
        "time.sleep(60)\n"
    )
    result = run_isolated([sys.executable, "-c", parent], timeout=3)
    assert result["status"] == "timeout"
    time.sleep(1.5)
    size = ticks.stat().st_size
    time.sleep(1.5)
    assert ticks.stat().st_size == size


def test_merge_rows_replaces_same_pair_and_keeps_the_rest():
    old = [
        {"doc": "a", "tool": "x", "status": "ok", "seconds": 1, "scores": {"word_recall": 1.0}},
        {"doc": "a", "tool": "y", "status": "ok", "seconds": 1, "scores": {"word_recall": 0.5}},
    ]
    new = [
        {"doc": "a", "tool": "y", "status": "ok", "seconds": 2, "scores": {"word_recall": 0.7}},
        {"doc": "a", "tool": "z", "status": "ok", "seconds": 3, "scores": {"word_recall": 0.9}},
    ]
    merged = merge_rows(old, new)
    assert [(r["doc"], r["tool"], r["scores"]["word_recall"]) for r in merged] == [
        ("a", "x", 1.0),
        ("a", "y", 0.7),
        ("a", "z", 0.9),
    ]
    assert aggregate(merged)["z"]["n_ok"] == 1


def test_libreoffice_uses_the_configured_profile_folder(monkeypatch, tmp_path):
    seen = {}

    def run(argv, **kwargs):
        seen["argv"] = argv
        outdir = argv[argv.index("--outdir") + 1]
        with open(f"{outdir}/in.docx", "wb") as handle:
            handle.write(b"docx")
        return SimpleNamespace(returncode=0, stdout="", stderr="")

    profile = tmp_path / "lo-profile"
    monkeypatch.setenv(external.PROFILE_ENV, str(profile))
    monkeypatch.setattr(external, "find_soffice", lambda: "soffice")
    monkeypatch.setattr(external.subprocess, "run", run)
    external.convert_libreoffice(str(tmp_path / "in.pdf"), str(tmp_path / "o.docx"))
    assert f"-env:UserInstallation={profile.resolve().as_uri()}" in seen["argv"]
