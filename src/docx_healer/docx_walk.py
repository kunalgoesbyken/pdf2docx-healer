"""Iteration over every paragraph and run of a document, including tables and headers/footers."""

import unicodedata
from functools import lru_cache
from typing import Any, Iterator, List, Tuple

from docx.document import Document as DocumentType
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


def _own_header_footers(doc: DocumentType) -> List[Any]:
    containers = []
    for section in doc.sections:
        for container in (
            section.header,
            section.footer,
            section.first_page_header,
            section.first_page_footer,
            section.even_page_header,
            section.even_page_footer,
        ):
            if not container.is_linked_to_previous:
                containers.append(container)
    return containers


def _roots(doc: DocumentType) -> List[Tuple[Any, Any]]:
    """Return (xml root, python-docx parent) pairs for the body and each own header/footer.

    Linked headers and footers are skipped so shared content is visited once and no
    empty header part is created as a side effect.
    """
    roots = [(doc.element.body, doc)]
    roots.extend((container._element, container) for container in _own_header_footers(doc))
    return roots


def iter_all_paragraphs(doc: DocumentType) -> Iterator[Paragraph]:
    """Yield body paragraphs, paragraphs in (nested) tables, and header/footer paragraphs."""
    for root, parent in _roots(doc):
        for element in list(root.iter(qn("w:p"))):
            yield Paragraph(element, parent)


def iter_all_run_elements(doc: DocumentType) -> Iterator[Any]:
    """Yield every ``w:r`` element, including runs wrapped in hyperlinks."""
    for root, _ in _roots(doc):
        yield from list(root.iter(qn("w:r")))


_CONTINUING_SECTION_TYPES = frozenset(("continuous", "nextColumn"))
_DROPPED_CATEGORIES = frozenset(("Cc", "Cf", "Zs", "Zl", "Zp"))


@lru_cache(maxsize=65536)
def fold_char(char: str) -> str:
    """Fold one character for comparison: NFKC, casefolded, whitespace and controls dropped."""
    if char.isspace() or unicodedata.category(char) in _DROPPED_CATEGORIES:
        return ""
    return unicodedata.normalize("NFKC", char).casefold()


def fold_text(text: str) -> str:
    return "".join(fold_char(char) for char in text)


def paragraph_plain_text(paragraph_element: Any) -> str:
    """Return the text of every ``w:t`` under a paragraph, in document order."""
    return "".join(node.text or "" for node in paragraph_element.iter(qn("w:t")))


def _section_properties(element: Any) -> Any:
    if element.tag == qn("w:sectPr"):
        return element
    if element.tag == qn("w:p"):
        return element.find(qn("w:pPr") + "/" + qn("w:sectPr"))
    return None


def paged_body_elements(doc: DocumentType) -> List[List[Any]]:
    """Group the body's direct children by page.

    A page starts at every section whose break type is not continuous or next column,
    which is how pdf2docx separates PDF pages. A section's properties sit at its end
    and describe how that section starts.
    """
    pages: List[List[Any]] = []
    pending: List[Any] = []
    for child in doc.element.body.iterchildren():
        pending.append(child)
        properties = _section_properties(child)
        if properties is None:
            continue
        kind = properties.find(qn("w:type"))
        continues = kind is not None and kind.get(qn("w:val")) in _CONTINUING_SECTION_TYPES
        if continues and pages:
            pages[-1].extend(pending)
        else:
            pages.append(pending)
        pending = []
    if pending and pages:
        pages[-1].extend(pending)
    elif pending:
        pages.append(pending)
    return pages
