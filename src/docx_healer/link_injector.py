"""Hyperlink injection via python-docx post-processing."""

import copy
import re
from typing import Any, List, Optional, Tuple

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from .docx_walk import iter_all_paragraphs
from .files import PathLike
from .log import progress

_URL_PATTERN = re.compile(
    r'https?://[^\s<>"\')\]]+|'
    r'ftp://[^\s<>"\')\]]+|'
    r'www\.[^\s<>"\')\]]+\.[a-zA-Z]{2,}[^\s<>"\')\]]*|'
    r'mailto:[^\s<>"\')\]]+',
    re.IGNORECASE,
)
_TRAILING_PUNCTUATION = ".,;:!?"
_HYPERLINK_BLUE = RGBColor(0x05, 0x63, 0xC1)
_PLAIN_RUN_CHILDREN = frozenset(
    qn(tag) for tag in ("w:rPr", "w:t", "w:tab", "w:cr", "w:noBreakHyphen")
)


def fix_docx_hyperlinks(docx_path: PathLike, verbose: bool = False) -> int:
    """Wrap plain-text URLs in clickable hyperlinks, in place.

    Covers body text, (nested) table cells, and headers/footers.
    Returns the number of hyperlinks added.
    """
    doc = Document(str(docx_path))
    count = sum(_hyperlink_urls_in_paragraph(p) for p in iter_all_paragraphs(doc))
    if count > 0:
        doc.save(str(docx_path))
        progress(verbose, f"Added {count} hyperlink(s)")
    return count


def _url_spans(text: str) -> List[Tuple[int, int]]:
    spans = []
    for match in _URL_PATTERN.finditer(text):
        end = match.start() + len(match.group().rstrip(_TRAILING_PUNCTUATION))
        if end > match.start():
            spans.append((match.start(), end))
    return spans


def _target_url(url: str) -> str:
    if "://" in url or url.lower().startswith("mailto:"):
        return url
    return "https://" + url


def _is_plain_text_child(child: Any) -> bool:
    if child.tag == qn("w:br"):
        return child.get(qn("w:type")) in (None, "textWrapping")
    return child.tag in _PLAIN_RUN_CHILDREN


def _is_plain_text_run(run_element: Any) -> bool:
    return all(_is_plain_text_child(child) for child in run_element)


def _text_run(text: str, rpr: Optional[Any]) -> Any:
    run = OxmlElement("w:r")
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    run.text = text
    return run


def _hyperlink_urls_in_paragraph(paragraph: Paragraph) -> int:
    """Replace URL text in the paragraph's direct runs with hyperlink elements.

    Splits each run into text-before / hyperlink / text-after pieces, for every URL
    in the run. Runs holding anything besides plain text are left untouched.
    """
    return sum(_hyperlink_urls_in_run(paragraph, run) for run in list(paragraph.runs))


def _hyperlink_urls_in_run(paragraph: Paragraph, run: Run) -> int:
    text = run.text
    spans = _url_spans(text)
    if not spans or not _is_plain_text_run(run._element):
        return 0

    run_element = run._element
    parent = run_element.getparent()
    if parent is None:
        return 0
    position = parent.index(run_element)
    rpr = run_element.find(qn("w:rPr"))

    pieces = []
    cursor = 0
    for start, end in spans:
        if start > cursor:
            pieces.append(_text_run(text[cursor:start], rpr))
        pieces.append(_build_hyperlink(paragraph, text[start:end], rpr))
        cursor = end
    if cursor < len(text):
        pieces.append(_text_run(text[cursor:], rpr))

    parent.remove(run_element)
    for offset, piece in enumerate(pieces):
        parent.insert(position + offset, piece)
    return len(spans)


def _build_hyperlink(paragraph: Paragraph, url_text: str, rpr: Optional[Any]) -> Any:
    relationship_id = paragraph.part.relate_to(
        _target_url(url_text), RT.HYPERLINK, is_external=True
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), relationship_id)

    link_run = _text_run(url_text, rpr)
    hyperlink.append(link_run)
    styled = Run(link_run, paragraph)
    styled.font.underline = True
    styled.font.color.rgb = _HYPERLINK_BLUE
    return hyperlink
