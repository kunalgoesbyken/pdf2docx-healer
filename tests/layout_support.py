from pathlib import Path
from typing import List, Sequence

from docx import Document
from docx.oxml.ns import qn

from docx_healer.ocr_pipeline import pymupdf

FONT = 10
LEFT_COLUMN = 72
RIGHT_COLUMN = 320
COLUMN_WIDTH = 210
BODY_LINES = 14


def prose_line(column: str, index: int) -> str:
    return f"{column} column line {index:02d} carries ordinary running words"


def _save(doc: "pymupdf.Document", path: Path) -> Path:
    doc.save(str(path))
    doc.close()
    return path


def make_two_column_pdf(path: Path) -> Path:
    doc = pymupdf.open()
    page = doc.new_page()
    for index in range(BODY_LINES):
        y = 100 + index * 13
        page.insert_text((LEFT_COLUMN, y), prose_line("left", index), fontsize=FONT)
        page.insert_text((RIGHT_COLUMN, y), prose_line("right", index), fontsize=FONT)
    return _save(doc, path)


def make_data_table_pdf(
    path: Path, rows: Sequence[Sequence[str]], gap: int = 90, fontname: str = "helv"
) -> Path:
    doc = pymupdf.open()
    page = doc.new_page()
    for row, values in enumerate(rows):
        for col, value in enumerate(values):
            page.insert_text((72 + col * gap, 100 + row * 20), value, fontsize=11, fontname=fontname)
    return _save(doc, path)


NUMERIC_ROWS = [
    ("Model", "Params", "Score", "Speed"),
    ("alpha", "7B", "61.1", "12"),
    ("bravo", "13B", "64.3", "9"),
    ("charlie", "33B", "68.9", "5"),
    ("delta", "65B", "71.2", "3"),
]

CJK_ROWS = [
    ("地区", "收入", "增长"),
    ("北方", "120", "四"),
    ("南方", "95", "二"),
    ("东方", "88", "三"),
]


def make_boxed_text_pdf(path: Path) -> Path:
    doc = pymupdf.open()
    page = doc.new_page()
    page.draw_rect(pymupdf.Rect(72, 90, 400, 130), width=1)
    page.insert_text((80, 114), "Boxed notice inside a single cell", fontsize=11)
    return _save(doc, path)


def make_callout_pdf(path: Path) -> Path:
    doc = pymupdf.open()
    page = doc.new_page()
    page.draw_rect(pymupdf.Rect(72, 90, 450, 140), color=None, fill=(0.85, 0.9, 1.0))
    page.insert_text((80, 110), "Shaded callout first line", fontsize=11)
    page.insert_text((80, 126), "Shaded callout second line", fontsize=11)
    page.insert_text((72, 180), "Plain paragraph after the callout", fontsize=11)
    return _save(doc, path)


def make_nested_pieces_pdf(path: Path) -> Path:
    doc = pymupdf.open()
    page = doc.new_page()
    xs, ys = [72, 272, 472], [100, 160, 220]
    for x in xs:
        page.draw_line((x, ys[0]), (x, ys[-1]), width=1)
    for y in ys:
        page.draw_line((xs[0], y), (xs[-1], y), width=1)
    page.insert_text((80, 125), "Prac", fontsize=11)
    page.insert_text((115, 138), "tical", fontsize=11)
    page.insert_text((280, 125), "Second cell", fontsize=11)
    page.insert_text((80, 185), "Third cell", fontsize=11)
    page.insert_text((280, 185), "Fourth cell", fontsize=11)
    return _save(doc, path)


def make_marker_list_pdf(path: Path) -> Path:
    doc = pymupdf.open()
    page = doc.new_page()
    items = [
        "Explain the purpose of the unit clearly",
        "Describe the main stages of the process",
        "Compare two approaches with evidence",
        "Summarise the findings in plain words",
    ]
    for index, item in enumerate(items):
        y = 100 + index * 22
        page.insert_text((72, y), f"{index + 1}.", fontsize=11)
        page.insert_text((100, y), item, fontsize=11)
    return _save(doc, path)


def make_spaced_words_pdf(path: Path, words: Sequence[str]) -> Path:
    doc = pymupdf.open()
    page = doc.new_page()
    x = 72.0
    for word in words:
        page.insert_text((x, 100), word, fontsize=11)
        x += pymupdf.get_text_length(word, fontsize=11) + 3.5
    return _save(doc, path)


def table_texts(path: Path) -> List[str]:
    body = Document(str(path)).element.body
    return [
        "".join(t.text or "" for t in table.iter(qn("w:t")))
        for table in body.iter(qn("w:tbl"))
    ]


def nested_table_count(path: Path) -> int:
    body = Document(str(path)).element.body
    return sum(1 for table in body.iter(qn("w:tbl")) if table.getparent().tag == qn("w:tc"))


def make_three_column_prose_pdf(path: Path) -> Path:
    doc = pymupdf.open()
    page = doc.new_page()
    for index in range(BODY_LINES):
        y = 100 + index * 13
        for column, x in enumerate((60, 225, 390)):
            page.insert_text((x, y), prose_line(f"col{column}", index)[:34], fontsize=FONT)
    return _save(doc, path)


def make_definition_list_pdf(path: Path) -> Path:
    doc = pymupdf.open()
    page = doc.new_page()
    y = 100
    terms = ["Alpha", "Bravo", "Charlie"]
    for term in terms:
        page.insert_text((72, y), term, fontsize=11)
        for line in range(3):
            page.insert_text((180, y + line * 13), f"{term} explains the matter in wrapped line {line}", fontsize=10)
        y += 60
    return _save(doc, path)


SENTENCE_ROWS = [
    ("Explain the unit well", "Use the given example", "Check against notes"),
    ("Describe each stage", "Compare both methods", "Write a short summary"),
    ("List all your sources", "Mark claims needing proof", "Share the draft early"),
    ("Plan the next steps", "Review progress weekly", "Report to the team"),
]


def make_rotated_label_pdf(path: Path) -> Path:
    doc = pymupdf.open()
    page = doc.new_page()
    for index in range(8):
        page.insert_text((120, 100 + index * 14), f"Body paragraph line {index} with ordinary words", fontsize=11)
    page.insert_text((60, 300), "Margin stamp rotated", fontsize=11, rotate=90)
    return _save(doc, path)

