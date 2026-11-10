import tempfile
from pathlib import Path
from typing import List

import pytest

from docx_healer import ocr_pipeline
from docx_healer.ocr_pipeline import pymupdf

OCR_TEXT = "Recognised text produced by the stubbed OCR engine for this scanned page."


@pytest.fixture
def scratch_tempdir(tmp_path, monkeypatch) -> Path:
    scratch = tmp_path / "scratch"
    scratch.mkdir()
    monkeypatch.setattr(tempfile, "tempdir", str(scratch))
    return scratch


@pytest.fixture
def fake_ocr(monkeypatch) -> List[str]:
    calls: List[str] = []

    def pdfocr_tobytes(self, *args, **kwargs):
        calls.append("ocr")
        doc = pymupdf.open()
        doc.new_page().insert_text((72, 100), OCR_TEXT, fontsize=11)
        data = doc.tobytes()
        doc.close()
        return data

    monkeypatch.setattr(pymupdf.Pixmap, "pdfocr_tobytes", pdfocr_tobytes)
    monkeypatch.setattr(ocr_pipeline, "require_tesseract", lambda: None)
    return calls
