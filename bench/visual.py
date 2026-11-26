"""Compare how a converted DOCX looks against its PDF.

The DOCX is laid out by LibreOffice Writer and saved as a PDF. Every word of the original is
then looked up on the same page of that rendering: is it there, is it in the same place, and
does it keep its size, colour, weight, slant and serif or sans font. Side-by-side pictures of
each page are written so a person can check the numbers by eye.
"""

import argparse
import re
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

import pymupdf

from . import fetch
from .external import render_pdf, render_pdf_word
from .run import CACHE, REPO, convert

BOLD = re.compile(r"bold|black|heavy|semibold|demi|medi", re.IGNORECASE)
ITALIC = re.compile(r"italic|oblique|ital", re.IGNORECASE)
SERIF_FLAG = 4
MONO_FLAG = 8
BOLD_FLAG = 16
ITALIC_FLAG = 2
SIZE_TOLERANCE = 0.6
PLACE_X = 0.04
PLACE_Y = 0.025
PICTURE_DPI = 60
GAP = 24
RENDERERS = {"libreoffice": render_pdf, "word": render_pdf_word}
CHECKS = ("found", "placed", "size", "colour", "bold", "italic", "font")


@dataclass
class Word:
    text: str
    x: float
    y: float
    size: float
    colour: int
    bold: bool
    italic: bool
    font: str


def _font_class(span: Dict[str, Any]) -> str:
    if span["flags"] & MONO_FLAG:
        return "mono"
    return "serif" if span["flags"] & SERIF_FLAG else "sans"


def _span_at(spans: List[Dict[str, Any]], x: float, y: float) -> Optional[Dict[str, Any]]:
    for span in spans:
        left, top, right, bottom = span["bbox"]
        if left <= x <= right and top <= y <= bottom:
            return span
    return None


def page_words(page: Any) -> List[Word]:
    """Return the page's words with the style of the span each one sits in."""
    spans = [
        span
        for block in page.get_text("dict")["blocks"]
        for line in block.get("lines", [])
        for span in line["spans"]
        if span["text"].strip()
    ]
    words = []
    for left, top, right, bottom, text, *_ in page.get_text("words"):
        x, y = (left + right) / 2, (top + bottom) / 2
        span = _span_at(spans, x, y)
        if span is None:
            continue
        name = span["font"]
        words.append(
            Word(
                text,
                x / page.rect.width,
                y / page.rect.height,
                span["size"],
                span["color"],
                bool(span["flags"] & BOLD_FLAG or BOLD.search(name)),
                bool(span["flags"] & ITALIC_FLAG or ITALIC.search(name)),
                _font_class(span),
            )
        )
    return words


def compare_page(original: List[Word], rendered: List[Word]) -> Dict[str, int]:
    """Count, for one page, how many original words keep each property in the rendering."""
    by_text: Dict[str, List[Word]] = {}
    for word in rendered:
        by_text.setdefault(word.text, []).append(word)
    counts = dict.fromkeys(CHECKS, 0)
    counts["words"] = len(original)
    for word in original:
        candidates = by_text.get(word.text)
        if not candidates:
            continue
        match = min(candidates, key=lambda other: abs(other.x - word.x) + abs(other.y - word.y))
        counts["found"] += 1
        counts["placed"] += abs(match.x - word.x) <= PLACE_X and abs(match.y - word.y) <= PLACE_Y
        counts["size"] += abs(match.size - word.size) <= SIZE_TOLERANCE
        counts["colour"] += match.colour == word.colour
        counts["bold"] += match.bold == word.bold
        counts["italic"] += match.italic == word.italic
        counts["font"] += match.font == word.font
    return counts


def side_by_side(original: Any, rendered: Any, index: int, out: Path) -> None:
    """Write one picture with the PDF page on the left and the rendered DOCX page on the right."""
    left = original[index].rect
    right = rendered[index].rect if index < len(rendered) else left
    sheet = pymupdf.open()
    page = sheet.new_page(
        width=left.width + GAP + right.width, height=max(left.height, right.height)
    )
    page.show_pdf_page(pymupdf.Rect(0, 0, left.width, left.height), original, index)
    if index < len(rendered):
        page.show_pdf_page(
            pymupdf.Rect(left.width + GAP, 0, left.width + GAP + right.width, right.height),
            rendered,
            index,
        )
    page.get_pixmap(dpi=PICTURE_DPI).save(str(out))
    sheet.close()


def compare(pdf: Path, rendered_pdf: Path, pictures: Optional[Path], label: str) -> Dict[str, Any]:
    """Compare every page of a PDF with its rendered DOCX."""
    totals = dict.fromkeys(("words", *CHECKS), 0)
    with pymupdf.open(str(pdf)) as original, pymupdf.open(str(rendered_pdf)) as rendered:
        for index, page in enumerate(original):
            words = page_words(page)
            other = page_words(rendered[index]) if index < len(rendered) else []
            for key, value in compare_page(words, other).items():
                totals[key] += value
            if pictures is not None:
                side_by_side(original, rendered, index, pictures / f"{label}-p{index + 1}.png")
        pages = (len(original), len(rendered))
    share = {key: totals[key] / totals["words"] if totals["words"] else 0.0 for key in CHECKS}
    return {"pages": pages, **share}


def first_pages(pdf: Path, count: int, out: Path) -> Path:
    """Save the first ``count`` pages of a PDF as a new file."""
    with pymupdf.open(str(pdf)) as source:
        part = pymupdf.open()
        part.insert_pdf(source, to_page=min(count, len(source)) - 1)
        part.save(str(out))
        part.close()
    return out


def run_document(
    entry: Dict[str, Any], tool: str, pages: int, pictures: Optional[Path], renderer: str
) -> Dict[str, Any]:
    """Convert the first pages of one corpus document with one tool and compare the look."""
    source = fetch.fetch_entry(entry, fetch.CACHE)
    with tempfile.TemporaryDirectory() as work:
        pdf = first_pages(source, pages, Path(work) / f"{entry['id']}.pdf")
        docx = Path(work) / f"{entry['id']}.docx"
        result = convert(tool, pdf, docx, 900, REPO / "src")
        if result["status"] != "ok":
            return {"doc": entry["id"], "error": result.get("error") or result["status"]}
        rendered = Path(work) / "rendered.pdf"
        RENDERERS[renderer](str(docx), str(rendered))
        return {"doc": entry["id"], **compare(pdf, rendered, pictures, f"{entry['id']}.{tool}")}


def table(rows: Sequence[Dict[str, Any]]) -> str:
    """Render the results as a markdown table with a mean row."""
    header = "| Document | Pages | " + " | ".join(CHECKS) + " |"
    lines = [header, "|" + " --- |" * (len(CHECKS) + 2)]
    good = [row for row in rows if "error" not in row]
    for row in rows:
        if "error" in row:
            lines.append(f"| {row['doc']} | failed: {row['error'][:60]} |")
            continue
        pages = "{}/{}".format(*row["pages"])
        lines.append(
            f"| {row['doc']} | {pages} | " + " | ".join(f"{row[c]:.2f}" for c in CHECKS) + " |"
        )
    if good:
        means = [sum(row[c] for row in good) / len(good) for c in CHECKS]
        lines.append("| mean | | " + " | ".join(f"{m:.2f}" for m in means) + " |")
    return "\n".join(lines)


def main(argv: Optional[List[str]] = None) -> int:
    """Compare the look of converted documents with their PDFs."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tool", default="healer-editable")
    parser.add_argument("--only", nargs="*")
    parser.add_argument("--pages", type=int, default=3)
    parser.add_argument("--renderer", choices=list(RENDERERS), default="libreoffice")
    parser.add_argument("--pictures", type=Path, default=CACHE / "visual")
    args = parser.parse_args(argv)
    args.pictures.mkdir(parents=True, exist_ok=True)
    entries = [
        entry
        for entry in fetch.load_corpus()["documents"]
        if not entry.get("slow") and (not args.only or entry["id"] in args.only)
    ]
    rows: List[Dict[str, Any]] = [
        run_document(e, args.tool, args.pages, args.pictures, args.renderer) for e in entries
    ]
    print(table(rows))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
