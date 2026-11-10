"""Orchestrates OCR, pdf2docx conversion, and DOCX post-processing."""

import logging
import os
import tempfile
from contextlib import contextmanager, nullcontext
from dataclasses import dataclass
from pathlib import Path
from typing import ContextManager, Iterator, Optional

from .errors import InvalidInputError
from .files import PathLike, remove_quietly, resolve_input_pdf, resolve_output_path
from .log import progress
from .ocr_pipeline import prepare_ocr_copy
from .patches import list_item_patches
from .post_processor import post_process_docx


@dataclass
class HealerConfig:
    """Configuration for the pdf2docx-healer pipeline."""

    ocr_enabled: bool = False
    ocr_lang: str = "eng"
    ocr_dpi: int = 300
    ocr_threshold: float = 0.3

    fix_lists: bool = True
    fix_hyperlinks: bool = True
    fix_fonts: bool = True

    aggressive_lists: bool = False

    verbose: bool = True

    def __post_init__(self) -> None:
        if not 0.0 <= self.ocr_threshold <= 1.0:
            raise InvalidInputError(
                f"ocr_threshold must be between 0 and 1, got {self.ocr_threshold}"
            )
        if self.ocr_dpi <= 0:
            raise InvalidInputError(f"ocr_dpi must be positive, got {self.ocr_dpi}")
        if not self.ocr_lang.strip():
            raise InvalidInputError("ocr_lang must not be empty")


def heal(
    pdf_path: PathLike,
    docx_path: Optional[PathLike] = None,
    config: Optional[HealerConfig] = None,
) -> str:
    """Convert a PDF to DOCX preserving formatting (bullets, hyperlinks, fonts).

    Steps: optionally OCR scanned pages, convert with pdf2docx, then repair lists,
    hyperlinks and fonts in the result. The output file is only written when every
    step up to post-processing succeeded. Returns the output path.

    Raises FileNotFoundError for a missing input and InvalidInputError for an
    unusable input, output location, or configuration.
    """
    config = config or HealerConfig()
    pdf = resolve_input_pdf(pdf_path)
    output = resolve_output_path(docx_path, pdf)

    with _source_pdf(pdf, config) as source, _staged_output(output) as staged:
        _convert(source, staged, config)
        post_process_docx(
            staged,
            fix_lists=config.fix_lists,
            fix_hyperlinks=config.fix_hyperlinks,
            fix_fonts=config.fix_fonts,
            aggressive_lists=config.aggressive_lists,
            verbose=config.verbose,
        )

    progress(config.verbose, f"Done: {output}")
    return str(output)


@contextmanager
def _source_pdf(pdf: Path, config: HealerConfig) -> Iterator[str]:
    """Yield the PDF to convert; delete only the OCR copy this call created."""
    ocr_copy = None
    if config.ocr_enabled:
        ocr_copy = prepare_ocr_copy(
            pdf,
            language=config.ocr_lang,
            dpi=config.ocr_dpi,
            threshold=config.ocr_threshold,
            verbose=config.verbose,
        )
    try:
        yield ocr_copy or str(pdf)
    finally:
        if ocr_copy is not None:
            remove_quietly(ocr_copy)


@contextmanager
def _staged_output(output: Path) -> Iterator[str]:
    """Yield a scratch DOCX beside ``output`` and move it into place on success."""
    handle, name = tempfile.mkstemp(prefix=".docx_healer_", suffix=".docx", dir=output.parent)
    os.close(handle)
    try:
        yield name
        os.replace(name, output)
    finally:
        remove_quietly(name)


@contextmanager
def _root_log_level(level: Optional[int]) -> Iterator[None]:
    if level is None:
        yield
        return
    root = logging.getLogger()
    previous = root.level
    root.setLevel(level)
    try:
        yield
    finally:
        root.setLevel(previous)


def _convert(pdf_path: str, docx_path: str, config: HealerConfig) -> None:
    from pdf2docx import Converter

    progress(config.verbose, "Converting...")
    patches: ContextManager[None] = list_item_patches() if config.fix_lists else nullcontext()
    quiet_level = None if config.verbose else logging.WARNING
    with _root_log_level(quiet_level), patches:
        converter = Converter(pdf_path)
        try:
            converter.convert(docx_path)
        finally:
            converter.close()
