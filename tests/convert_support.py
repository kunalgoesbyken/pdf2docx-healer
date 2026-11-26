from pathlib import Path
from typing import Any, List, Optional, Set

import pytest
from docx import Document
from docx.oxml.ns import qn
from pdf2docx.page.Page import Page
from pdf2docx.page.Pages import Pages

from docx_healer.ocr_pipeline import pymupdf

PAGE_TEXTS = [
    "Zeppelin harbour lantern",
    "Quokka meadow compass",
    "Marimba glacier violin",
]


def body_paragraphs(path: Path) -> List[str]:
    body = Document(str(path)).element.body
    return ["".join(t.text or "" for t in p.iter(qn("w:t"))) for p in body.iter(qn("w:p"))]


def body_text(path: Path) -> str:
    return " ".join(body_paragraphs(path))


def picture_count(path: Path) -> int:
    return len(list(Document(str(path)).element.body.iter(qn("w:drawing"))))


def table_count(path: Path) -> int:
    return len(list(Document(str(path)).element.body.iter(qn("w:tbl"))))


def red_png() -> bytes:
    pixmap = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 40, 40), False)
    pixmap.set_rect(pixmap.irect, (200, 30, 30))
    return pixmap.tobytes("png")


def make_image_page_pdf(path: Path, text_pages_before: int = 0) -> Path:
    doc = pymupdf.open()
    for index in range(text_pages_before):
        doc.new_page().insert_text((72, 100), PAGE_TEXTS[index], fontsize=11)
    doc.new_page().insert_image(pymupdf.Rect(72, 100, 172, 200), stream=red_png())
    doc.save(str(path))
    doc.close()
    return path


def make_zero_page_pdf(path: Path) -> Path:
    path.write_bytes(
        b"%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n"
        b"2 0 obj<</Type/Pages/Kids[]/Count 0>>endobj\n"
        b"trailer<</Root 1 0 R/Size 3>>\n%%EOF\n"
    )
    return path


def make_encrypted_pdf(path: Path) -> Path:
    doc = pymupdf.open()
    doc.new_page().insert_text((72, 100), PAGE_TEXTS[0], fontsize=11)
    doc.save(str(path), encryption=pymupdf.PDF_ENCRYPT_AES_256, owner_pw="owner", user_pw="user")
    doc.close()
    return path


def make_ruled_table_pdf(path: Path) -> Path:
    doc = pymupdf.open()
    page = doc.new_page()
    xs, ys = [72, 172, 272, 372], [100, 130, 160]
    for x in xs:
        page.draw_line((x, ys[0]), (x, ys[-1]), width=1)
    for y in ys:
        page.draw_line((xs[0], y), (xs[-1], y), width=1)
    for row, values in enumerate([("Region", "Revenue", "Growth"), ("North", "120", "4%")]):
        for col, value in enumerate(values):
            page.insert_text((xs[col] + 6, ys[row] + 20), value, fontsize=11)
    doc.save(str(path))
    doc.close()
    return path


def make_borderless_table_pdf(path: Path) -> Path:
    doc = pymupdf.open()
    page = doc.new_page()
    rows = [("Region", "Revenue", "Growth"), ("North", "120", "4%"), ("South", "95", "-2%")]
    for row, values in enumerate(rows):
        for col, value in enumerate(values):
            page.insert_text((72 + col * 120, 100 + row * 24), value, fontsize=11)
    doc.save(str(path))
    doc.close()
    return path


def fail_parse(monkeypatch: pytest.MonkeyPatch, failing: Set[int]) -> None:
    original = Page.parse

    def parse(self: Page, **settings: object) -> object:
        if self.id in failing:
            raise RuntimeError(f"forced parse failure on {self.id + 1}")
        return original(self, **settings)

    monkeypatch.setattr(Page, "parse", parse)


def fail_make(
    monkeypatch: pytest.MonkeyPatch, failing: Set[int], leftover: Optional[str] = None
) -> None:
    original = Page.make_docx

    def make_docx(self: Page, doc: Any) -> None:
        if self.id in failing:
            if leftover is not None:
                original(self, doc)
                doc.add_paragraph(leftover)
            raise RuntimeError(f"forced make failure on {self.id + 1}")
        original(self, doc)

    monkeypatch.setattr(Page, "make_docx", make_docx)


def fail_document_analysis(monkeypatch: pytest.MonkeyPatch) -> None:
    def parse(self: Pages, fitz_doc: object, **settings: object) -> None:
        raise RuntimeError("forced document analysis failure")

    monkeypatch.setattr(Pages, "parse", parse)
