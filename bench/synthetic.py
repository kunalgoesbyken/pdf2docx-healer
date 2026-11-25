"""Generate a synthetic PDF whose headings, lists, links, tables and images are known exactly."""

from pathlib import Path
from typing import Any, Dict, List, Tuple

from .compat import pymupdf

LEFT = 72
INDENT = 28
LABEL_GAP = 34
LINE = 18
BODY = 11
SIZES = {1: 22, 2: 16, 3: 13}
BULLETS = ("•", "–", "-")
HANDBOOK = "https://example.com/handbook"
ANNEX = "https://example.org/annex"
SOURCE = "https://example.net/source"

NUMBERED_1: List[Tuple[str, str, int]] = [
    ("1.", "Write the migration", 0),
    ("2.", "Run it on staging", 0),
    ("a.", "Check row counts", 1),
    ("b.", "Compare checksums", 1),
    ("3.", "Roll out to production", 0),
]
BULLET_ITEMS: List[Tuple[str, int]] = [
    ("Ship the billing service", 0),
    ("Cut tail latency", 0),
    ("Cache hot keys", 1),
    ("Drop the batch query", 1),
    ("Add a batch loader", 2),
    ("Hire two engineers", 0),
]
START_THREE = [("3.", "Take the backup"), ("4.", "Stop the writers"), ("5.", "Flip the switch")]
ROMAN = [("i.", "Draft the notice"), ("ii.", "Review the notice"), ("iii.", "Send the notice")]
LETTERS = [("a)", "Collect the logs"), ("b)", "Sort the logs"), ("c)", "Archive the logs")]
OUTLINE_ITEMS: List[Tuple[str, str, int]] = [
    ("1.", "Overview section", 0),
    ("1.1", "Scope of work", 1),
    ("1.2", "Intended audience", 1),
    ("2.", "Method section", 0),
    ("2.1", "Data sources used", 1),
]
TABLE = [
    ["Region", "Revenue", "Growth"],
    ["North", "120", "4%"],
    ["South", "95", "-2%"],
    ["East", "143", "7%"],
]


class Canvas:
    """Tracks the cursor on the current page and records what was drawn."""

    def __init__(self, doc: Any) -> None:
        self.doc = doc
        self.page = doc.new_page()
        self.y = 90.0
        self.expected: Dict[str, Any] = {
            "headings": [],
            "lists": [],
            "ext_links": [],
            "internal_links": 0,
            "tables": 0,
            "images": 0,
            "table_cells": TABLE,
        }
        self.outline: List[List[Any]] = []
        self.pending: List[Tuple[int, Any, int]] = []

    def new_page(self) -> None:
        """Start the next page."""
        self.page = self.doc.new_page()
        self.y = 90.0

    def heading(self, level: int, text: str) -> None:
        """Draw a heading and register it for the outline."""
        self.y += 10
        self.page.insert_text((LEFT, self.y), text, fontsize=SIZES[level], fontname="hebo")
        self.outline.append([level, text, self.page.number + 1])
        self.expected["headings"].append([level, text])
        self.y += LINE + 6

    def line(self, text: str, x: float = LEFT) -> None:
        """Draw one body line."""
        self.page.insert_text((x, self.y), text, fontsize=BODY, fontname="helv")
        self.y += LINE

    def link_line(self, parts: List[Tuple[str, Any]]) -> None:
        """Draw a line of plain and linked segments; a link is a URI or a target page."""
        x = float(LEFT)
        for text, target in parts:
            self.page.insert_text((x, self.y), text, fontsize=BODY, fontname="helv")
            width = pymupdf.get_text_length(text, fontname="helv", fontsize=BODY)
            rect = pymupdf.Rect(x, self.y - BODY, x + width, self.y + 3)
            if isinstance(target, str):
                self.page.insert_link({"kind": pymupdf.LINK_URI, "from": rect, "uri": target})
                self.expected["ext_links"].append(target)
            elif target is not None:
                self.pending.append((self.page.number, rect, target))
                self.expected["internal_links"] += 1
            x += width
        self.y += LINE

    def item(self, label: str, text: str, level: int, kind: str) -> None:
        """Draw a list item with a visible label and register the expectation."""
        x = LEFT + INDENT * level
        self.page.insert_text((x, self.y), label, fontsize=BODY, fontname="helv")
        self.page.insert_text((x + LABEL_GAP, self.y), text, fontsize=BODY, fontname="helv")
        self.expected["lists"].append({"text": text, "label": label, "level": level, "kind": kind})
        self.y += LINE

    def place_internal_links(self) -> None:
        """Insert the internal links once every target page exists."""
        for page_number, rect, target in self.pending:
            self.doc[page_number].insert_link(
                {
                    "kind": pymupdf.LINK_GOTO,
                    "from": rect,
                    "page": target,
                    "to": pymupdf.Point(LEFT, 90),
                }
            )

    def table(self) -> None:
        """Draw a ruled table."""
        widths = (120, 100, 100)
        height = 24
        top = self.y
        for row_index, row in enumerate(TABLE):
            x = LEFT
            for width, cell in zip(widths, row):
                rect = pymupdf.Rect(
                    x, top + row_index * height, x + width, top + (row_index + 1) * height
                )
                self.page.draw_rect(rect, width=0.8)
                self.page.insert_text((x + 6, rect.y1 - 8), cell, fontsize=BODY, fontname="helv")
                x += width
        self.y = top + len(TABLE) * height + 16
        self.expected["tables"] += 1

    def image(self) -> None:
        """Place a small solid-colour image."""
        pixmap = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 90, 60), False)
        pixmap.set_rect(pixmap.irect, (200, 60, 60))
        self.page.insert_image(pymupdf.Rect(LEFT, self.y, LEFT + 90, self.y + 60), pixmap=pixmap)
        self.expected["images"] += 1
        self.y += 80


def build(path: Any) -> Dict[str, Any]:
    """Write the synthetic PDF to path and return the exact expectations."""
    doc = pymupdf.open()
    canvas = Canvas(doc)
    canvas.heading(1, "Quarterly Operations Report")
    canvas.link_line(
        [
            ("Read the ", None),
            ("operations handbook", HANDBOOK),
            (" and the ", None),
            ("annex notes", ANNEX),
            (" before the review.", None),
        ]
    )
    canvas.link_line([("Jump to ", None), ("the results table", 2), (" when you are ready.", None)])
    canvas.heading(2, "1. Goals")
    for text, level in BULLET_ITEMS:
        canvas.item(BULLETS[level], text, level, "bullet")
    canvas.y += 8
    for label, text, level in NUMBERED_1:
        canvas.item(label, text, level, "number")
    canvas.new_page()
    canvas.heading(2, "2. Steps")
    canvas.line("The next list picks up at three, then come roman and letter lists.")
    for label, text in START_THREE:
        canvas.item(label, text, 0, "number")
    canvas.y += 8
    for label, text in ROMAN:
        canvas.item(label, text, 0, "number")
    canvas.y += 8
    for label, text in LETTERS:
        canvas.item(label, text, 0, "number")
    canvas.y += 8
    for label, text, level in OUTLINE_ITEMS:
        canvas.item(label, text, level, "number")
    canvas.y += 8
    canvas.link_line(
        [
            ("Back to ", None),
            ("the top", 0),
            (" or read the ", None),
            ("source code", SOURCE),
            (".", None),
        ]
    )
    canvas.new_page()
    canvas.heading(2, "3. Numbers")
    canvas.table()
    canvas.image()
    canvas.heading(3, "3.1 Regional detail")
    canvas.line("The north region grew fastest this quarter.")
    canvas.place_internal_links()
    doc.set_toc(canvas.outline)
    doc.save(str(Path(path)), deflate=True)
    doc.close()
    return canvas.expected
