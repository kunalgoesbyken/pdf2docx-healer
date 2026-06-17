"""Hyperlink injection via python-docx post-processing (not store/restore)."""

import re
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree

# URL pattern
_URL_PATTERN = re.compile(
    r'https?://[^\s<>"\')\]]+|'
    r'ftp://[^\s<>"\')\]]+|'
    r'www\.[^\s<>"\')\]]+\.[a-zA-Z]{2,}[^\s<>"\')\]]*|'
    r'mailto:[^\s<>"\')\]]+',
    re.IGNORECASE
)


def fix_docx_hyperlinks(docx_path, verbose=False):
    """Post-process: find URL text in runs and wrap them in OOXML hyperlinks.

    Opens the DOCX, scans all runs for URLs, and replaces plain-text URLs
    with clickable hyperlinks in-place.

    Returns number of hyperlinks added.
    """
    from docx import Document
    doc = Document(docx_path)
    count = 0

    for paragraph in doc.paragraphs:
        count += _hyperlink_urls_in_paragraph(paragraph)

    if count > 0:
        doc.save(docx_path)
        if verbose:
            print(f"[healer] Added {count} hyperlink(s) in {docx_path}")

    return count


def _hyperlink_urls_in_paragraph(paragraph):
    """Find URLs in paragraph runs and convert them to hyperlinks.

    Strategy: for each run containing URL(s), split the run into
    text-before / hyperlink / text-after segments, processing ALL URLs
    in the run (not just the first).

    NOTE: This function modifies the paragraph XML tree by inserting and
    removing elements.  python-docx's in-memory ``paragraph.runs`` view
    is not refreshed after these changes, so the caller must not rely on
    it being up-to-date.  Because we iterate over a snapshot (``list()``)
    and save the document immediately afterwards, this is safe.
    """
    count = 0
    runs = list(paragraph.runs)

    for run in runs:
        if not run.text:
            continue
        urls = _URL_PATTERN.findall(run.text)
        if not urls:
            continue

        # Check if run is already inside a hyperlink
        if run._element.getparent() is not None:
            parent_tag = run._element.getparent().tag
            if parent_tag == qn('w:hyperlink'):
                continue

        run_elem = run._element
        parent = run_elem.getparent()
        orig_rPr = run_elem.find(qn('w:rPr'))
        insert_pos = parent.index(run_elem)
        parent.remove(run_elem)

        remaining_text = run.text

        for url_match in urls:
            original_url = url_match
            target_url = url_match
            if '://' not in target_url and not target_url.startswith('mailto:'):
                target_url = 'https://' + target_url

            idx = remaining_text.find(original_url)
            if idx < 0:
                continue

            before = remaining_text[:idx]
            after = remaining_text[idx + len(original_url):]

            if before:
                before_run = OxmlElement('w:r')
                if orig_rPr is not None:
                    before_run.append(copy_rPr(orig_rPr))
                before_run.text = before
                parent.insert(insert_pos, before_run)
                insert_pos += 1

            hyperlink = _build_hyperlink(paragraph, target_url, run)
            parent.insert(insert_pos, hyperlink)
            hyperlink[0].text = original_url
            insert_pos += 1
            count += 1

            remaining_text = after

        if remaining_text:
            after_run = OxmlElement('w:r')
            if orig_rPr is not None:
                after_run.append(copy_rPr(orig_rPr))
            after_run.text = remaining_text
            parent.insert(insert_pos, after_run)

    return count


def _build_hyperlink(paragraph, url, source_run):
    """Build an OOXML w:hyperlink element referencing an external URL."""
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)

    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    # Build the run inside the hyperlink, copying formatting
    hl_run = OxmlElement('w:r')
    rPr = source_run._element.find(qn('w:rPr'))
    if rPr is not None:
        hl_run.append(copy_rPr(rPr))

    # Add hyperlink style (blue + underline)
    _ensure_hyperlink_style(hl_run)

    hyperlink.append(hl_run)
    return hyperlink


def _ensure_hyperlink_style(run_elem):
    """Ensure the run inside a hyperlink has blue color and underline."""
    rPr = run_elem.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        run_elem.insert(0, rPr)

    # Add underline if not present
    if rPr.find(qn('w:u')) is None:
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)

    # Add blue color if not present
    if rPr.find(qn('w:color')) is None:
        color = OxmlElement('w:color')
        color.set(qn('w:val'), '0563C1')
        rPr.append(color)


def copy_rPr(source_rPr):
    """Deep copy a w:rPr element, with fallback if lxml serialisation fails."""
    try:
        return etree.fromstring(etree.tostring(source_rPr))
    except Exception:
        import copy
        return copy.deepcopy(source_rPr)
