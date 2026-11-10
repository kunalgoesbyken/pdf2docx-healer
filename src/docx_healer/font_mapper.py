"""Font fallback mapping for CJK and unavailable fonts."""

from typing import Any, Dict, List, Optional

from docx import Document
from docx.oxml.ns import qn

from .constants import CHAR_RANGE_FONT_MAP, FONT_FALLBACK_MAP
from .docx_walk import iter_all_run_elements
from .files import PathLike
from .log import progress

_RFONTS_ATTRIBUTES = ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs")
_SUBSET_PREFIX_LENGTH = 6
_FOLDED_FALLBACKS: Dict[str, List[str]] = {
    name.casefold(): chain for name, chain in FONT_FALLBACK_MAP.items()
}


def _strip_subset_prefix(name: str) -> str:
    prefix, plus, base = name.partition("+")
    if plus and len(prefix) == _SUBSET_PREFIX_LENGTH and prefix.isalpha() and prefix.isupper():
        return base
    return name


def _preferred(font: str) -> str:
    chain = FONT_FALLBACK_MAP.get(font)
    return chain[0] if chain else font


def _font_for_script(text: str) -> Optional[str]:
    for (low, high), font in CHAR_RANGE_FONT_MAP:
        if any(low <= ord(char) <= high for char in text):
            return _preferred(font)
    return None


def normalize_font_name(original_name: Optional[str], text_sample: str = "") -> str:
    """Resolve a PDF font name to one commonly available on the target system.

    Known PDF font names map to their preferred system font. Unknown names are kept
    (minus any PDF subset prefix) unless ``text_sample`` needs a script-specific font.
    Falls back to Arial only for an empty name.
    """
    if not original_name:
        return "Arial"

    name = _strip_subset_prefix(original_name.strip())
    chain = _FOLDED_FALLBACKS.get(name.casefold())
    if chain:
        return chain[0]
    return _font_for_script(text_sample) or name


def _fix_run_fonts(run_element: Any) -> bool:
    rpr = run_element.rPr
    if rpr is None or rpr.rFonts is None:
        return False

    text = run_element.text
    changed = False
    for attribute in _RFONTS_ATTRIBUTES:
        original = rpr.rFonts.get(qn(attribute))
        if not original:
            continue
        normalized = normalize_font_name(original, text)
        if normalized != original:
            rpr.rFonts.set(qn(attribute), normalized)
            changed = True
    return changed


def fix_docx_fonts(docx_path: PathLike, verbose: bool = False) -> int:
    """Map PDF font names in a DOCX to available fonts, in place.

    Covers body text, (nested) table cells, headers/footers, and hyperlink runs.
    Returns the number of runs whose fonts changed.
    """
    doc = Document(str(docx_path))
    fixed = sum(_fix_run_fonts(run) for run in iter_all_run_elements(doc))
    if fixed > 0:
        doc.save(str(docx_path))
        progress(verbose, f"Fixed fonts in {fixed} run(s)")
    return fixed
