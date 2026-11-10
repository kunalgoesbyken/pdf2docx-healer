from typing import List, Tuple

from docx import Document
from docx.oxml.ns import qn


def paragraph_text(paragraph) -> str:
    return "".join(t.text or "" for t in paragraph._p.iter(qn("w:t")))


def run_text(run_element) -> str:
    return "".join(t.text or "" for t in run_element.iter(qn("w:t")))


def hyperlinks(part_owner, root) -> List[Tuple[str, str]]:
    found = []
    for link in root.iter(qn("w:hyperlink")):
        target = part_owner.part.rels[link.get(qn("r:id"))].target_ref
        found.append((run_text(link), target))
    return found


def body_hyperlinks(path) -> List[Tuple[str, str]]:
    doc = Document(str(path))
    return hyperlinks(doc, doc.element.body)
