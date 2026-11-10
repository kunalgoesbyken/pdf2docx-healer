"""Iteration over every paragraph and run of a document, including tables and headers/footers."""

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
