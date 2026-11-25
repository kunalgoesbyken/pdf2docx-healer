"""Child process entry point: convert one PDF with one tool and print one JSON result line."""

import json
import logging
import os
import sys
import time
from typing import Any, Callable, Dict, List, Optional

from . import external

PREFIX = "BENCH_RESULT "
ERROR_LIMIT = 300


def convert_pdf2docx(pdf: str, out: str) -> None:
    """Plain pdf2docx with its defaults."""
    from pdf2docx import Converter

    converter = Converter(pdf)
    try:
        converter.convert(out)
    finally:
        converter.close()


def convert_healer_fidelity(pdf: str, out: str) -> None:
    """pdf2docx-healer with the default HealerConfig."""
    from docx_healer import HealerConfig, heal

    heal(pdf, out, HealerConfig(verbose=False))


def convert_healer_editable(pdf: str, out: str) -> None:
    """pdf2docx-healer in editable mode."""
    from docx_healer import HealerConfig, heal

    heal(pdf, out, HealerConfig(mode="editable", verbose=False))


TOOLS: Dict[str, Callable[[str, str], Optional[str]]] = {
    "pdf2docx": convert_pdf2docx,
    "healer-fidelity": convert_healer_fidelity,
    "healer-editable": convert_healer_editable,
    "pdf2docx-plus": external.convert_plus,
    "libreoffice": external.convert_libreoffice,
}


def describe(exc: BaseException) -> str:
    """Return a short single-line description of an exception."""
    return f"{type(exc).__name__}: {exc}".replace("\n", " ")[:ERROR_LIMIT]


def run_tool(convert: Callable[[str, str], Optional[str]], pdf: str, out: str) -> Dict[str, Any]:
    """Call convert and classify the outcome as ok, unavailable (NotImplementedError) or error.

    A string returned by convert is kept as the note of the result.
    """
    started = time.perf_counter()
    status, error, note = "ok", None, None
    try:
        note = convert(pdf, out)
        if not os.path.exists(out) or os.path.getsize(out) == 0:
            status, error = "error", "no output file was written"
    except NotImplementedError as exc:
        status, error = "unavailable", describe(exc)
    except Exception as exc:
        status, error = "error", describe(exc)
    return {
        "status": status,
        "seconds": time.perf_counter() - started,
        "error": error,
        "note": note,
    }


def main(argv: List[str]) -> int:
    """Usage: python -m bench.worker TOOL PDF OUT."""
    tool, pdf, out = argv
    logging.getLogger().setLevel(logging.WARNING)
    result = run_tool(TOOLS[tool], pdf, out)
    sys.stdout.write("\n" + PREFIX + json.dumps(result) + "\n")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
