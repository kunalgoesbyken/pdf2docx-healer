from pathlib import Path
from typing import Callable, Optional, Sequence

from docx import Document
from docx.document import Document as DocumentType

from docx_healer.ocr_pipeline import pymupdf

TEXT_PAGE = "This page carries plenty of embedded text, so it is never treated as scanned."


def make_pdf(path: Path, pages: Sequence[Optional[str]]) -> Path:
    doc = pymupdf.open()
    for text in pages:
        page = doc.new_page()
        if text:
            page.insert_text((72, 100), text, fontsize=11)
    doc.save(str(path))
    doc.close()
    return path


def make_list_with_prices_pdf(path: Path) -> Path:
    doc = pymupdf.open()
    page = doc.new_page()
    y = 100
    for index, item in enumerate(["Apples and oranges", "Bread from the bakery", "Cheese", "Milk and eggs"]):
        page.insert_text((72, y), "•", fontsize=12)
        page.insert_text((90, y), item, fontsize=12)
        page.insert_text((300, y), f"${index + 1}.00", fontsize=12)
        y += 22
    doc.save(str(path))
    doc.close()
    return path


def make_docx(path: Path, build: Callable[[DocumentType], None]) -> Path:
    doc = Document()
    build(doc)
    doc.save(str(path))
    return path
