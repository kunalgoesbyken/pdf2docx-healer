"""Converters that run outside this Python environment: pdf2docx-plus and LibreOffice."""

import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import List, Optional

PLUS_ENV = "BENCH_PDF2DOCX_PLUS"
SOFFICE_ENV = "BENCH_SOFFICE"
PROFILE_ENV = "BENCH_SOFFICE_PROFILE"
SOFFICE_NAMES = ("soffice", "libreoffice")
SOFFICE_PATHS = (
    "C:/Program Files/LibreOffice/program/soffice.exe",
    "C:/Program Files (x86)/LibreOffice/program/soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
)
ERROR_TAIL = 300


def find_plus() -> Optional[str]:
    """Locate the pdf2docx-plus command: BENCH_PDF2DOCX_PLUS, else PATH."""
    return os.environ.get(PLUS_ENV) or shutil.which("pdf2docx-plus")


def find_soffice() -> Optional[str]:
    """Locate LibreOffice: BENCH_SOFFICE, else PATH, else the usual install folders."""
    configured = os.environ.get(SOFFICE_ENV)
    if configured:
        return configured
    for name in SOFFICE_NAMES:
        found = shutil.which(name)
        if found:
            return found
    return next((path for path in SOFFICE_PATHS if os.path.exists(path)), None)


def logged_failures(text: str) -> Optional[int]:
    """Read the failed page count from the last 'done:' line pdf2docx-plus logs."""
    for line in reversed(text.splitlines()):
        if "done:" in line and "failed=" in line:
            token = line.split("failed=")[1].split()[0]
            return int(token) if token.isdigit() else None
    return None


def error_tail(text: str) -> str:
    """Return the last part of a process output as one line."""
    return " ".join(text.split())[-ERROR_TAIL:]


def run_command(argv: List[str]) -> "subprocess.CompletedProcess[str]":
    """Run a converter command, capturing its output."""
    return subprocess.run(
        argv,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdin=subprocess.DEVNULL,
    )


def has_output(path: str) -> bool:
    """Return True when the converter wrote a non-empty file."""
    return os.path.exists(path) and os.path.getsize(path) > 0


def convert_plus(pdf: str, out: str) -> Optional[str]:
    """pdf2docx-plus with its default profile (fidelity); page failures are returned as a note."""
    exe = find_plus()
    if exe is None:
        raise NotImplementedError(
            f"pdf2docx-plus not found; install it in its own environment and set {PLUS_ENV}"
        )
    process = run_command([exe, "convert", pdf, "--output", out])
    if not has_output(out):
        raise RuntimeError(f"exit code {process.returncode}: {error_tail(process.stderr)}")
    failures = logged_failures(process.stderr + process.stdout)
    return f"logged page failures: {failures}" if failures else None


def convert_libreoffice(pdf: str, out: str) -> None:
    """LibreOffice headless with the Writer PDF import filter and its own user profile.

    The profile is BENCH_SOFFICE_PROFILE when set, otherwise a throwaway folder, so a stale
    lock from an earlier run can never block a conversion.
    """
    soffice = find_soffice()
    if soffice is None:
        raise NotImplementedError(f"LibreOffice not found; install it or set {SOFFICE_ENV}")
    with tempfile.TemporaryDirectory() as work:
        outdir = Path(work) / "out"
        outdir.mkdir()
        profile = Path(os.environ.get(PROFILE_ENV) or Path(work) / "profile")
        argv = [
            soffice,
            f"-env:UserInstallation={profile.resolve().as_uri()}",
            "--headless",
            "--infilter=writer_pdf_import",
            "--convert-to",
            "docx",
            "--outdir",
            str(outdir),
            pdf,
        ]
        process = run_command(argv)
        produced = outdir / (Path(pdf).stem + ".docx")
        if not produced.exists():
            raise RuntimeError(f"exit code {process.returncode}: {error_tail(process.stderr)}")
        shutil.move(str(produced), out)
