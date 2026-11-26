"""Font fallback mapping for CJK and unavailable fonts."""

import re
from typing import Any, Dict, List, Optional

from docx import Document
from docx.oxml.ns import qn

from .constants import CHAR_RANGE_FONT_MAP, FONT_FALLBACK_MAP, FONT_FAMILY_EQUIVALENTS
from .docx_walk import iter_all_run_elements
from .files import PathLike
from .log import progress

_RFONTS_ATTRIBUTES = ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs")
_SUBSET_PREFIX_LENGTH = 6
_STYLE_SEPARATORS = re.compile(r"[-,]")
_TRAILING_TAGS = re.compile(r"(?:PSMT|PS|MT)$")
_WORD = re.compile(r"[^\W\d_]{3,}")
_CAMEL = re.compile(r"(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])")
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


def _key(name: str) -> str:
    return "".join(ch for ch in name.casefold() if ch.isalnum())


def font_family(original_name: Optional[str]) -> str:
    """Return the font family a PDF font name stands for, as a name Word and LibreOffice know.

    The subset prefix and the style part of a PostScript name are dropped
    (``ABCDEF+NimbusRomNo9L-Medi`` is the Nimbus Roman family). Families with a widely
    installed equivalent of the same widths map to it (Nimbus Roman to Times New Roman);
    any other family keeps its own name, split into words when it is written in camel case.
    """
    if not original_name or not original_name.strip():
        return ""
    name = _strip_subset_prefix(original_name.strip())
    known = _FOLDED_FALLBACKS.get(name.casefold())
    if known:
        return known[0]
    base = _STYLE_SEPARATORS.split(name, maxsplit=1)[0].strip() or name
    base = _TRAILING_TAGS.sub("", base) or base
    if "+" in base or not _WORD.search(base):
        return ""
    key = _key(base)
    for prefixes, family in FONT_FAMILY_EQUIVALENTS:
        if any(key.startswith(prefix) for prefix in prefixes):
            return family
    return base if " " in base else _CAMEL.sub(" ", base)


def normalize_font_name(original_name: Optional[str], text_sample: str = "") -> str:
    """Resolve a PDF font name to one commonly available on the target system.

    Known PDF font names map to their preferred system font. A ``text_sample`` that needs a
    script-specific font gets that font; any other name resolves to its font family.
    Falls back to Arial only for an empty name.
    """
    if not original_name or not original_name.strip():
        return "Arial"
    name = _strip_subset_prefix(original_name.strip())
    chain = _FOLDED_FALLBACKS.get(name.casefold())
    if chain:
        return chain[0]
    return _font_for_script(text_sample) or font_family(name) or name


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
