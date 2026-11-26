"""pdf2docx-healer: Formatting-preserving PDF-to-DOCX converter."""

from .converter import HealerConfig, heal
from .errors import (
    HealerError,
    IncompatiblePdf2docxError,
    InvalidInputError,
    OcrUnavailableError,
)

__version__ = "0.3.2"
__all__ = [
    "heal",
    "HealerConfig",
    "HealerError",
    "IncompatiblePdf2docxError",
    "InvalidInputError",
    "OcrUnavailableError",
]
