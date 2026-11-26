import sys
import time

from bench.runner import run_isolated
from bench.worker import run_tool


def test_isolated_ok_returns_stdout_and_seconds():
    result = run_isolated([sys.executable, "-c", "print('hello')"], timeout=30)
    assert result["status"] == "ok"
    assert result["stdout"].strip() == "hello"
    assert result["seconds"] >= 0


def test_isolated_crash_is_reported_and_does_not_raise():
    result = run_isolated([sys.executable, "-c", "import os; os._exit(3)"], timeout=30)
    assert result["status"] == "crash"
    assert result["returncode"] == 3


def test_isolated_timeout_kills_the_child():
    started = time.perf_counter()
    result = run_isolated([sys.executable, "-c", "import time; time.sleep(60)"], timeout=1)
    assert result["status"] == "timeout"
    assert time.perf_counter() - started < 20


def test_run_tool_not_implemented_is_unavailable():
    def editable(pdf, out):
        raise NotImplementedError("editable mode is not implemented yet")

    result = run_tool(editable, "in.pdf", "out.docx")
    assert result["status"] == "unavailable"
    assert "not implemented" in result["error"]


def test_run_tool_exception_is_error_with_type_name():
    def broken(pdf, out):
        raise ValueError("bad pdf")

    result = run_tool(broken, "in.pdf", "out.docx")
    assert result["status"] == "error"
    assert result["error"].startswith("ValueError")


def test_run_tool_without_output_file_is_error(tmp_path):
    result = run_tool(lambda pdf, out: None, "in.pdf", str(tmp_path / "missing.docx"))
    assert result["status"] == "error"
    assert "output" in result["error"]


def test_run_tool_success_records_seconds(tmp_path):
    out = tmp_path / "o.docx"

    def good(pdf, target):
        open(target, "wb").write(b"x")

    result = run_tool(good, "in.pdf", str(out))
    assert result["status"] == "ok" and result["seconds"] >= 0
