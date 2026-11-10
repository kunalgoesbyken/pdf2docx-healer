"""Path validation and safe file removal."""

import os
from pathlib import Path
from typing import Optional, Union

from .errors import InvalidInputError
from .log import logger

PathLike = Union[str, "os.PathLike[str]"]

_PDF_MAGIC = b"%PDF-"
_PDF_HEADER_WINDOW = 1024


def resolve_input_pdf(pdf_path: PathLike) -> Path:
    path = Path(pdf_path)
    if not path.exists():
        raise FileNotFoundError(f"PDF file not found: {path}")
    if not path.is_file():
        raise InvalidInputError(f"Not a file: {path}")
    with path.open("rb") as handle:
        header = handle.read(_PDF_HEADER_WINDOW)
    if _PDF_MAGIC not in header:
        raise InvalidInputError(f"Not a PDF file: {path}")
    return path


def resolve_output_path(docx_path: Optional[PathLike], pdf: Path) -> Path:
    output = Path(docx_path) if docx_path is not None else pdf.with_suffix(".docx")
    if output.is_dir():
        raise InvalidInputError(f"Output path is a directory: {output}")
    if not output.parent.is_dir():
        raise InvalidInputError(f"Output directory does not exist: {output.parent}")
    if output.exists() and output.samefile(pdf):
        raise InvalidInputError(f"Output would overwrite the input PDF: {output}")
    return output


def remove_quietly(path: PathLike) -> None:
    try:
        os.unlink(path)
    except FileNotFoundError:
        return
    except OSError as exc:
        logger.warning("Could not remove temporary file %s: %s", path, exc)
