"""Logging helpers shared across docx_healer."""

import logging

logger = logging.getLogger("docx_healer")


def progress(verbose: bool, message: str) -> None:
    if verbose:
        print(f"[healer] {message}")
