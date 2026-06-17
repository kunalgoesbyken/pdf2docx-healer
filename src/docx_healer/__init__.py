"""pdf2docx-healer: Formatting-preserving PDF-to-DOCX converter."""

from .converter import heal, HealerConfig

__version__ = "0.1.3"
__all__ = ["heal", "HealerConfig"]
