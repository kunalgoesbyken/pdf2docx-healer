"""Font fallback mapping for CJK and unavailable fonts."""

from .constants import FONT_FALLBACK_MAP, CHAR_RANGE_FONT_MAP


def normalize_font_name(original_name, text_sample=''):
    """Resolve a PyMuPDF/PDF font name to one commonly available on the target system.
    
    Args:
        original_name: Font name from pdf2docx span (e.g., 'SimSun', 'TimesNewRomanPSMT').
        text_sample: Sample characters to use for character-based fallback.
    
    Returns:
        A resolved font name string.
    """
    if not original_name:
        return 'Arial'
    
    name = original_name.strip()
    
    # Exact match in fallback map
    if name in FONT_FALLBACK_MAP:
        return FONT_FALLBACK_MAP[name][0]
    
    # Check normalized name (strip prefix like 'BAAAAA+' from subset fonts)
    if '+' in name:
        base_name = name.split('+', 1)[-1]
    else:
        base_name = name
    
    if base_name in FONT_FALLBACK_MAP:
        return FONT_FALLBACK_MAP[base_name][0]
    
    # Case-insensitive lookup
    name_upper = base_name.upper()
    for key, fallback in FONT_FALLBACK_MAP.items():
        if key.upper() == name_upper:
            return fallback[0]
    
    # Character-based fallback (check text content)
    if text_sample:
        for (lo, hi), fallback_font in CHAR_RANGE_FONT_MAP:
            if any(lo <= ord(c) <= hi for c in text_sample):
                # Look up the range-matched font in our fallback map
                if fallback_font in FONT_FALLBACK_MAP:
                    return FONT_FALLBACK_MAP[fallback_font][0]
                return fallback_font
    
    # Ultimate fallback
    return 'Arial'


def fix_docx_fonts(docx_path, verbose=False):
    """Open a DOCX and fix font names using the fallback map.
    
    Args:
        docx_path: Path to the DOCX file to fix in-place.
        verbose: Print progress information.
    
    Returns:
        Number of font names fixed.
    """
    from docx import Document
    
    doc = Document(docx_path)
    fixed = 0
    
    for paragraph in doc.paragraphs:
        try:
            for run in paragraph.runs:
                original = run.font.name
                if original:
                    normalized = normalize_font_name(original, run.text)
                    if normalized != original:
                        run.font.name = normalized
                        fixed += 1
        except Exception:
            if verbose:
                import traceback
                traceback.print_exc()
            continue
    
    if fixed > 0:
        doc.save(docx_path)
        if verbose:
            print(f"[healer] Fixed {fixed} font name(s) in {docx_path}")
    
    return fixed
