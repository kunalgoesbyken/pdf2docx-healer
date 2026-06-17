"""Monkey-patches to pdf2docx internals for bullet list preservation."""

import re

# Unicode bullet codepoints
_BULLET_CODEPOINTS = frozenset({
    0x2022, 0x2023, 0x2043, 0x204C, 0x204D, 0x2219,
    0x25AA, 0x25CB, 0x25CF, 0x25D8, 0x25E6, 0x2619,
    0x2765, 0x2767, 0x29BE, 0x29BF, 0xF0B7,
    0x2013, 0x2014, 0x00B7, 0x2713, 0x2714,
    0x2610, 0x2611, 0x2612,
})

_BULLET_CHARS = frozenset({
    '-', '*', '+', '~', '>', '\u00bb',  # », guillemet
})

# Map private-use bullet 0xF0B7 to standard bullet
_BULLET_MAP = {0xF0B7: '\u2022'}

# Numbered list patterns: "1.", "(1)", "1)", "a.", "(a)", "i.", "(i)"
_NUMBERED_PATTERNS = [
    re.compile(r'^\s*\d{1,3}[\.\)]\s'),
    re.compile(r'^\s*[a-zA-Z][\.\)]\s'),
    re.compile(r'^\s*[ivxlcdm]+[\.\)]\s', re.IGNORECASE),
    re.compile(r'^\s*\(\d{1,3}\)\s'),
    re.compile(r'^\s*\([a-zA-Z]\)\s'),
    re.compile(r'^\s*\([ivxlcdm]+\)\s', re.IGNORECASE),
    re.compile(r'^\s*\d{1,3}\.\d{1,3}[\.\)]\s'),
]


def _is_list_item(text, bullets=True, numbers=True):
    """Detect if text starts with a bullet or number prefix.
    
    Returns the bullet character if detected (truthy), None otherwise.
    """
    if not text or not text.strip():
        return None

    stripped = text.strip()
    if not stripped:
        return None

    first_char = stripped[0]
    cp = ord(first_char)

    if bullets and cp in _BULLET_CODEPOINTS:
        return _BULLET_MAP.get(cp, first_char)

    if bullets and first_char in _BULLET_CHARS:
        return first_char

    if numbers:
        for pattern in _NUMBERED_PATTERNS:
            if pattern.match(stripped):
                return stripped

    return None


_patches_applied = False


def apply_patches(verbose=False):
    """Apply all monkey-patches to pdf2docx. Safe to call multiple times."""
    global _patches_applied
    if _patches_applied:
        return

    try:
        # Patch 1: Fix is_list_item function
        import pdf2docx.common.share as share_mod
        share_mod.is_list_item = _is_list_item
        if verbose:
            print("[healer] Patched pdf2docx.common.share.is_list_item")

        # Patch 2: Force list_not_table=True in collect_stream_lines
        import pdf2docx.layout.Blocks as blocks_mod
        _orig_collect = blocks_mod.Blocks.collect_stream_lines

        def _patched_collect(self, potential_shadings, line_separate_threshold, **kwargs):
            kwargs['list_not_table'] = True
            return _orig_collect(self, potential_shadings, line_separate_threshold, **kwargs)

        blocks_mod.Blocks.collect_stream_lines = _patched_collect
        if verbose:
            print("[healer] Patched Blocks.collect_stream_lines to force list_not_table=True")

        _patches_applied = True

    except ImportError as e:
        raise ImportError(
            "pdf2docx is required. Install with: pip install pdf2docx"
        ) from e


def is_patched():
    """Check if patches have been applied."""
    return _patches_applied
