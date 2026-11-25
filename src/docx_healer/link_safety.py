"""Turn external links that open anything other than a web page or mail client into plain text."""

import copy
import unicodedata
from typing import Any
from urllib.parse import urlsplit

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .files import PathLike
from .log import progress

SAFE_SCHEMES = frozenset({"http", "https", "mailto", "ftp"})


def is_safe_uri(uri: str) -> bool:
    """True when the target is an absolute http, https, mailto or ftp address."""
    cleaned = "".join(ch for ch in uri if unicodedata.category(ch)[0] != "C").strip()
    if not cleaned or cleaned.startswith(("\\", "//")):
        return False
    try:
        scheme = urlsplit(cleaned).scheme
    except ValueError:
        return False
    return scheme.lower() in SAFE_SCHEMES


def unlink_unsafe_targets(docx_path: PathLike, verbose: bool = False) -> int:
    """Lift hyperlinks nested in runs to paragraph level, then unwrap unsafe external ones.

    pdf2docx writes ``w:hyperlink`` inside ``w:r``, which the schema does not allow and
    python-docx cannot read. Every unsafe external link keeps its text as plain runs.
    """
    doc = Document(str(docx_path))
    part = doc.part
    removed = 0
    lifted = 0
    for link in list(doc.element.body.iter(qn("w:hyperlink"))):
        if _lift(link):
            lifted += 1
        rid = link.get(qn("r:id"))
        rel = part.rels.get(rid) if rid else None
        if rel is None or rel.reltype != RT.HYPERLINK or is_safe_uri(rel.target_ref):
            continue
        _unwrap(link)
        part.drop_rel(rid)
        removed += 1
    if removed or lifted:
        doc.save(str(docx_path))
    if removed:
        progress(verbose, f"Removed {removed} link(s) to non-web targets")
    return removed


def _lift(link: Any) -> bool:
    run = link.getparent()
    if run is None or run.tag != qn("w:r"):
        return False
    after = list(link.itersiblings())
    run.addnext(link)
    if after:
        tail = OxmlElement("w:r")
        rpr = run.find(qn("w:rPr"))
        if rpr is not None:
            tail.append(copy.deepcopy(rpr))
        for child in after:
            tail.append(child)
        link.addnext(tail)
    if all(child.tag == qn("w:rPr") for child in run):
        run.getparent().remove(run)
    return True


def _unwrap(link: Any) -> None:
    parent = link.getparent()
    index = parent.index(link)
    for child in list(link):
        style = child.find(f"{qn('w:rPr')}/{qn('w:rStyle')}")
        if style is not None and style.get(qn("w:val")) == "Hyperlink":
            style.getparent().remove(style)
        parent.insert(index, child)
        index += 1
    parent.remove(link)
