import hashlib
import zipfile
from contextlib import nullcontext
from typing import Any, Dict, List

from pdf2docx.image.ImagesExtractor import ImagesExtractor

from convert_support import red_png
from docx_healer import HealerConfig, converter, heal, speedups
from docx_healer.ocr_pipeline import pymupdf

HIDE = "_hide_page_text_and_images"


class FakeDocument:
    def __init__(self) -> None:
        self.updated: Dict[int, bytes] = {}

    def xref_stream(self, xref: int) -> bytes:
        return b"q /Im1 Do Q"

    def update_stream(self, xref: int, stream: bytes) -> None:
        self.updated[xref] = stream


class FakePage:
    def __init__(self, xobjects: int) -> None:
        self.parent = FakeDocument()
        self.listings = 0
        self._xobjects = xobjects

    def get_xobjects(self) -> List[Any]:
        return [(xref, "", 0, None) for xref in range(1, self._xobjects + 1)]

    def get_contents(self) -> List[int]:
        return []

    def get_images(self, full: bool = False) -> List[Any]:
        self.listings += 1
        return [(1, 0, 0, 0, 8, "DeviceRGB", "", "Im1", "DCTDecode")]


def drawing_heavy_pdf(path):
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_text((72, 60), "Figure with vector shapes", fontsize=11)
    for index in range(12):
        page.draw_circle((100 + index * 25, 200 + (index % 3) * 20), 8, fill=(0.2, 0.4, 0.8))
        page.draw_line((72, 300 + index * 4), (300, 300 + index * 4), width=0.5)
    page.insert_image(pymupdf.Rect(72, 400, 172, 500), stream=red_png())
    doc.save(str(path))
    doc.close()
    return path


def part_hashes(path) -> Dict[str, str]:
    with zipfile.ZipFile(path) as archive:
        return {name: hashlib.md5(archive.read(name)).hexdigest() for name in archive.namelist()}


def test_images_are_listed_once_per_pass_instead_of_once_per_stream():
    plain, cached = FakePage(40), FakePage(40)

    ImagesExtractor._hide_page_text_and_images(plain, rm_text=False, rm_image=True)
    with speedups.cached_image_listing():
        ImagesExtractor._hide_page_text_and_images(cached, rm_text=False, rm_image=True)

    assert plain.listings == 40
    assert cached.listings == 1
    assert cached.parent.updated == plain.parent.updated
    assert len(cached.parent.updated) == 40


def test_the_original_method_is_restored_afterwards():
    original = vars(ImagesExtractor)[HIDE]

    with speedups.cached_image_listing():
        assert vars(ImagesExtractor)[HIDE] is not original

    assert vars(ImagesExtractor)[HIDE] is original


def test_the_original_method_is_restored_after_an_error():
    original = vars(ImagesExtractor)[HIDE]

    try:
        with speedups.cached_image_listing():
            raise RuntimeError("boom")
    except RuntimeError:
        pass

    assert vars(ImagesExtractor)[HIDE] is original


def test_unknown_pdf2docx_layout_is_left_alone(monkeypatch):
    monkeypatch.setattr(speedups, "_resolve_extractor", lambda: None)
    original = vars(ImagesExtractor)[HIDE]

    with speedups.cached_image_listing():
        assert vars(ImagesExtractor)[HIDE] is original


def test_cached_listing_leaves_the_converted_document_byte_identical(tmp_path, monkeypatch):
    pdf = drawing_heavy_pdf(tmp_path / "in.pdf")
    config = HealerConfig(verbose=False)
    calls: List[int] = []
    original = vars(ImagesExtractor)[HIDE].__func__

    def spy(page: Any, *args: Any, **kwargs: Any) -> Any:
        calls.append(1)
        return original(page, *args, **kwargs)

    monkeypatch.setattr(ImagesExtractor, HIDE, staticmethod(spy))
    heal(pdf, tmp_path / "fast.docx", config)
    monkeypatch.setattr(converter, "cached_image_listing", nullcontext)
    heal(pdf, tmp_path / "plain.docx", config)

    assert len(calls) >= 2
    assert part_hashes(tmp_path / "fast.docx") == part_hashes(tmp_path / "plain.docx")
