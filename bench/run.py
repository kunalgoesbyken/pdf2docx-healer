"""Run the benchmark: fetch the corpus, convert with every tool, score, write results."""

import argparse
import datetime
import json
import os
import platform
import subprocess
import sys
from importlib import metadata
from pathlib import Path
from typing import Any, Dict, List, Optional

from . import external, fetch
from .compat import pymupdf
from .facts import InvalidDocx, docx_facts
from .report import merge_rows, scoreboard
from .runner import run_isolated
from .score import failure_scores, score
from .synthetic import build
from .truth import pdf_truth
from .worker import PREFIX, TOOLS

REPO = Path(__file__).resolve().parents[1]
CACHE = REPO / "bench" / ".cache"
RESULTS = REPO / "bench" / "results"
DEFAULT_TIMEOUT = 900
SLOW_FACTOR = 4
SYNTHETIC_ID = "synthetic"


def parse_result(run: Dict[str, Any]) -> Dict[str, Any]:
    """Turn a runner result into the worker's JSON result, or a crash/timeout record."""
    if run["status"] == "timeout":
        return {"status": "timeout", "seconds": run["seconds"], "error": "killed after the timeout"}
    for line in reversed(run["stdout"].splitlines()):
        if line.startswith(PREFIX):
            return json.loads(line[len(PREFIX) :])
    detail = run["stderr"].strip().splitlines()[-1:] or [""]
    return {
        "status": "crash",
        "seconds": run["seconds"],
        "error": f"exit code {run['returncode']}: {detail[0][:200]}",
    }


def convert(tool: str, pdf: Path, out: Path, timeout: float, healer_src: Path) -> Dict[str, Any]:
    """Run one conversion in a child process whose import path starts at healer_src."""
    env = dict(os.environ)
    env["PYTHONPATH"] = os.pathsep.join([str(healer_src), str(REPO)])
    argv = [sys.executable, "-m", "bench.worker", tool, str(pdf), str(out)]
    return parse_result(run_isolated(argv, timeout, cwd=str(REPO), env=env))


def run_pair(
    doc: Dict[str, Any],
    tool: str,
    pdf: Path,
    truth: Dict[str, Any],
    timeout: float,
    keep: bool,
    healer_src: Path,
) -> Dict[str, Any]:
    """Convert and score one (document, tool) pair."""
    out_dir = CACHE / "out"
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"{doc['id']}.{tool}.docx"
    out.unlink(missing_ok=True)
    result = convert(tool, pdf, out, timeout, healer_src)
    expected = doc.get("expected")
    scores = failure_scores(truth, expected)
    if result["status"] == "ok":
        try:
            scores = score(truth, docx_facts(out), expected)
        except InvalidDocx as exc:
            result["status"], result["error"] = "invalid", str(exc)[:300]
    elif result["status"] == "unavailable":
        scores = {}
    if not keep:
        out.unlink(missing_ok=True)
    return {
        "doc": doc["id"],
        "category": doc.get("category", ""),
        "tool": tool,
        "status": result["status"],
        "seconds": result["seconds"],
        "error": result.get("error"),
        "scores": scores,
        "note": result.get("note"),
    }


def git_state(repo: Path = REPO) -> str:
    """Return the short commit hash, with +dirty when the checkout has uncommitted changes."""

    def git(*args: str) -> str:
        return subprocess.run(
            ["git", *args], cwd=str(repo), capture_output=True, text=True
        ).stdout.strip()

    commit = git("rev-parse", "--short", "HEAD") or "unknown"
    return commit + (
        "+dirty" if git("status", "--porcelain", "--", "src", "bench", "pyproject.toml") else ""
    )


def versions() -> Dict[str, str]:
    """Collect the versions of everything that influences the numbers."""
    found = {"python": platform.python_version(), "pymupdf": pymupdf.__version__}
    for name in ("pdf2docx", "python-docx", "lxml"):
        try:
            found[name] = metadata.version(name)
        except metadata.PackageNotFoundError:
            found[name] = "missing"
    try:
        import docx_healer

        found["pdf2docx-healer"] = docx_healer.__version__
    except ImportError:
        found["pdf2docx-healer"] = "missing"
    found.update(external_versions())
    found["platform"] = platform.platform()
    return found


def external_versions() -> Dict[str, str]:
    """Versions of the converters that live outside this environment, when they are installed."""
    found: Dict[str, str] = {}
    plus = external.find_plus()
    if plus:
        python = next(
            (
                str(p)
                for p in (Path(plus).parent / "python.exe", Path(plus).parent / "python")
                if p.exists()
            ),
            None,
        )
        query = "import importlib.metadata as m; print(m.version('pdf2docx-plus'))"
        if python:
            answer = external.run_command([python, "-c", query]).stdout.strip().splitlines()
            found["pdf2docx-plus"] = answer[-1] if answer else "unknown"
    soffice = external.find_soffice()
    if soffice:
        found["libreoffice"] = libreoffice_version(soffice)
    return found


def libreoffice_version(soffice: str) -> str:
    """Return `soffice --version`, or the build id from version.ini when it prints nothing."""
    answer = external.run_command([soffice, "--version"]).stdout.strip().splitlines()
    if answer:
        return answer[0]
    ini = Path(soffice).resolve().parent / "version.ini"
    if ini.exists():
        for line in ini.read_text(encoding="utf-8", errors="replace").splitlines():
            if line.startswith("buildid="):
                return "build " + line.split("=", 1)[1][:12]
    return "unknown"


def collect_documents(args: argparse.Namespace, corpus: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Fetch the chosen corpus entries and add the synthetic document."""
    ids = [name for name in (args.only or []) if name != SYNTHETIC_ID]
    documents: List[Dict[str, Any]] = []
    if ids or not args.only:
        for entry in fetch.select(corpus, ids or None, args.include_slow):
            try:
                path = fetch.fetch_entry(entry, fetch.CACHE)
            except fetch.FetchError as exc:
                print(f"skip  {entry['id']}: {exc}", flush=True)
                continue
            documents.append({**entry, "path": path})
    if not args.only or SYNTHETIC_ID in args.only:
        synthetic = CACHE / "synthetic" / "synthetic.pdf"
        synthetic.parent.mkdir(parents=True, exist_ok=True)
        expected = build(synthetic)
        documents.append(
            {"id": SYNTHETIC_ID, "category": "synthetic", "path": synthetic, "expected": expected}
        )
    return documents


def main(argv: Optional[List[str]] = None) -> int:
    """Entry point: python -m bench.run."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tools", nargs="*", default=list(TOOLS), choices=list(TOOLS))
    parser.add_argument(
        "--only", nargs="*", help="document ids (add 'synthetic' for the generated PDF)"
    )
    parser.add_argument("--include-slow", action="store_true")
    parser.add_argument("--timeout", type=float, default=DEFAULT_TIMEOUT)
    parser.add_argument("--keep-output", action="store_true", help="keep the produced DOCX files")
    parser.add_argument(
        "--healer-src",
        type=Path,
        default=REPO / "src",
        help="src folder of the pdf2docx-healer checkout to measure (default: this checkout)",
    )
    parser.add_argument(
        "--merge",
        type=Path,
        help="results JSON from an earlier run; run the same documents and add the new rows to it",
    )
    args = parser.parse_args(argv)
    previous = json.loads(args.merge.read_text(encoding="utf-8")) if args.merge else None
    if previous and not args.only:
        args.only = previous["meta"]["documents"]
    documents = collect_documents(args, fetch.load_corpus())
    rows: List[Dict[str, Any]] = []
    for doc in documents:
        truth = pdf_truth(doc["path"], CACHE / "truth")
        timeout = args.timeout * (SLOW_FACTOR if doc.get("slow") else 1)
        for tool in args.tools:
            row = run_pair(
                doc, tool, doc["path"], truth, timeout, args.keep_output, args.healer_src
            )
            rows.append(row)
            recall = row["scores"].get("word_recall")
            shown = "-" if recall is None else f"{recall:.3f}"
            print(
                f"{row['status']:<11} {doc['id']:<24} {tool:<16} "
                f"recall {shown} {row['seconds']:.1f}s",
                flush=True,
            )
    meta = {
        "date": datetime.date.today().isoformat(),
        "bench_commit": git_state(REPO),
        "healer_commit": git_state(args.healer_src.resolve().parent),
        "versions": versions(),
        "timeout": args.timeout,
        "documents": [doc["id"] for doc in documents],
    }
    if previous:
        rows = merge_rows(previous["rows"], rows)
        if not any(tool.startswith("healer") for tool in args.tools):
            meta["healer_commit"] = previous["meta"].get(
                "healer_commit", previous["meta"].get("commit", "unknown")
            )
        meta["versions"] = {**previous["meta"]["versions"], **versions()}
        meta["merged_tools"] = sorted(set(args.tools))
    RESULTS.mkdir(parents=True, exist_ok=True)
    stem = RESULTS / meta["date"]
    stem.with_suffix(".json").write_text(
        json.dumps({"meta": meta, "rows": rows}, indent=1, ensure_ascii=False), encoding="utf-8"
    )
    board = scoreboard(rows, meta)
    stem.with_suffix(".md").write_text(board, encoding="utf-8")
    print(board)
    return 0


if __name__ == "__main__":
    sys.exit(main())
