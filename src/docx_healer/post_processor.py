"""Post-processing: list detection, hyperlinks, and font mapping on final DOCX."""

from .list_detector import fix_lists_in_docx
from .font_mapper import fix_docx_fonts
from .link_injector import fix_docx_hyperlinks


def post_process_docx(docx_path, fix_lists=True, fix_hyperlinks=True,
                      fix_fonts=True, aggressive_lists=False, verbose=False):
    """Run all post-processing steps on a generated DOCX.

    Args:
        docx_path: Path to the DOCX file (modified in-place).
        fix_lists: Apply list detection and Word list styles.
        fix_hyperlinks: Wrap URL text in clickable hyperlinks.
        fix_fonts: Apply CJK font fallback mapping.
        aggressive_lists: Style even solo numbered paragraphs.
        verbose: Print progress information.

    Returns:
        Dict with counts of fixes applied.
    """
    results = {'lists_fixed': 0, 'hyperlinks_added': 0, 'fonts_fixed': 0}

    if fix_lists:
        try:
            results['lists_fixed'] = fix_lists_in_docx(
                docx_path, aggressive=aggressive_lists, verbose=verbose
            )
        except Exception as e:
            if verbose:
                print(f"[healer] List fix error: {e}")

    if fix_hyperlinks:
        try:
            results['hyperlinks_added'] = fix_docx_hyperlinks(
                docx_path, verbose=verbose
            )
        except Exception as e:
            if verbose:
                print(f"[healer] Hyperlink fix error: {e}")

    if fix_fonts:
        try:
            results['fonts_fixed'] = fix_docx_fonts(docx_path, verbose=verbose)
        except Exception as e:
            if verbose:
                print(f"[healer] Font fix error: {e}")

    return results
