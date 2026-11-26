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


def merged_table_pdf(path):
    doc = pymupdf.open()
    page = doc.new_page()
    xs, ys = [72, 172, 272, 372], [100, 130, 160, 190]
    for x in xs:
        page.draw_line((x, ys[0] if x in (72, 372) else ys[1]), (x, ys[-1]), width=1)
    for y in ys:
        page.draw_line((xs[0], y), (xs[-1], y), width=1)
    page.insert_text((80, 120), "Merged header", fontsize=11)
    for row in (1, 2):
        for col in range(3):
            page.insert_text((xs[col] + 6, ys[row] + 20), f"r{row}c{col}", fontsize=11)
    doc.save(str(path))
    doc.close()
    return path


def make_table(rows: int = 3, columns: int = 3) -> Any:
    from docx import Document

    return Document().add_table(rows=rows, cols=columns)


def grid_pattern(table: Any) -> List[int]:
    first: Dict[int, int] = {}
    return [first.setdefault(id(cell._tc), len(first)) for cell in table._cells]


def test_the_cell_grid_is_built_once_per_table(monkeypatch):
    from docx.table import Table

    builds: List[int] = []
    original = vars(Table)["_cells"]

    def counting(table: Any) -> Any:
        builds.append(1)
        return original.fget(table)

    monkeypatch.setattr(Table, "_cells", property(counting))
    table = make_table()
    with speedups.cached_table_cells():
        cells = [table.cell(row, col) for row in range(3) for col in range(3) for _ in range(4)]

    assert len(cells) == 36
    assert len(builds) == 1


def test_merging_cells_drops_the_cached_grid():
    table = make_table()
    plain = make_table()
    plain.cell(0, 0).merge(plain.cell(1, 1))
    with speedups.cached_table_cells():
        assert len(table._cells) == 9
        table.cell(0, 0).merge(table.cell(1, 1))
        pattern = grid_pattern(table)

    assert pattern == grid_pattern(plain)
    assert len(set(pattern)) == 6


def test_a_failed_merge_still_drops_the_cached_grid():
    table = make_table()
    with speedups.cached_table_cells():
        table.cell(0, 0).merge(table.cell(0, 1))
        stale = table._cells
        try:
            table.cell(0, 0).merge(table.cell(2, 2))
        except ValueError:
            pass
        assert table._cells is not stale


def test_table_internals_are_restored_afterwards():
    from docx.table import Table, _Cell

    cells, merge = vars(Table)["_cells"], vars(_Cell)["merge"]

    with speedups.cached_table_cells():
        assert vars(Table)["_cells"] is not cells

    assert vars(Table)["_cells"] is cells
    assert vars(_Cell)["merge"] is merge
    assert not hasattr(make_table(), "_docx_healer_cells")


def test_unknown_python_docx_layout_is_left_alone(monkeypatch):
    from docx.table import Table

    monkeypatch.setattr(speedups, "_resolve_table_types", lambda: None)
    original = vars(Table)["_cells"]

    with speedups.cached_table_cells():
        assert vars(Table)["_cells"] is original


def test_cached_cells_leave_a_merged_table_document_byte_identical(tmp_path, monkeypatch):
    pdf = merged_table_pdf(tmp_path / "in.pdf")
    config = HealerConfig(verbose=False)

    heal(pdf, tmp_path / "fast.docx", config)
    monkeypatch.setattr(converter, "cached_table_cells", nullcontext)
    heal(pdf, tmp_path / "plain.docx", config)

    assert 'w:gridSpan' in zipfile.ZipFile(tmp_path / "fast.docx").read("word/document.xml").decode()
    assert part_hashes(tmp_path / "fast.docx") == part_hashes(tmp_path / "plain.docx")
