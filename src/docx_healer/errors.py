"""Exceptions raised by docx_healer."""


class HealerError(Exception):
    """Base class for all errors raised deliberately by docx_healer."""


class InvalidInputError(HealerError, ValueError):
    """The input PDF, output path, or configuration is unusable."""


class IncompatiblePdf2docxError(HealerError, RuntimeError):
    """The installed pdf2docx does not expose the internals the healer patches."""


class OcrUnavailableError(HealerError, RuntimeError):
    """OCR was requested but Tesseract is not usable."""
