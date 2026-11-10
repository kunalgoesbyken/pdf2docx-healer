"""Post-conversion list detection and Word style application."""

import copy
import re
from typing import Any, List, Tuple

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from .constants import BULLET_CODEPOINTS
from .log import logger, progress


def _normalize_text(text):
    """Fix double-encoded UTF-8 characters common in pdf2docx output."""
    if not text:
        return text
    has_suspect = any(0xC0 <= ord(c) <= 0xFD for c in text)
    if not has_suspect:
        return text
    try:
        return text.encode('latin-1').decode('utf-8')
    except (UnicodeEncodeError, UnicodeDecodeError):
        return text


_UNICODE_BULLETS = ''.join(chr(cp) for cp in sorted(BULLET_CODEPOINTS) if cp > 0x7F)
_ASCII_BULLETS   = r'\-\*\>\+\~' + ''.join(
    chr(cp) for cp in sorted(BULLET_CODEPOINTS) if cp <= 0x7F
)
_ALL_BULLETS     = _UNICODE_BULLETS + _ASCII_BULLETS

_BULLET_LINE_START = re.compile(
    r'(?:^|\n\s*)([' + re.escape(_ALL_BULLETS) + r'])\s+'
)
# Outline patterns (e.g. "1.1.", "1.1)", "1.1 ") placed FIRST to match before
# the simpler \d{1,3}[\.\)]\s+ which would grab just "1." from "1.1. foo"
_NUMBERED_START = re.compile(
    r'(?:^|\n\s*)(\d{1,3}\.\d{1,3}[\.\)]?\s+|'
    r'\d{1,3}[\.\)]\s+|[a-zA-Z][\.\)]\s+|[ivxlcdm]+[\.\)]\s+|'
    r'\(\d{1,3}\)\s+|\([a-zA-Z]\)\s+|\([ivxlcdm]+\)\s+)',
    re.IGNORECASE
)
_BULLET_ANYWHERE = re.compile(
    r'(?:'
    r'([' + re.escape(_UNICODE_BULLETS) + r'])(?=\s+\S)'
    r'|'
    r'(?:^|\n\s*)([' + re.escape(_ASCII_BULLETS) + r'])(?=\s+\S)'
    r')'
)
# Same reordering: outline first so "1.1." is not truncated to "1."
_NUMBERED_ANYWHERE = re.compile(
    r'(?:^|\n\s*)(\d{1,3}\.\d{1,3}[\.\)]?\s+|'
    r'\d{1,3}[\.\)]\s+|[a-zA-Z][\.\)]\s+|[ivxlcdm]+[\.\)]\s+|'
    r'\(\d{1,3}\)\s+|\([a-zA-Z]\)\s+|\([ivxlcdm]+\)\s+)',
    re.IGNORECASE
)

# Extract just the numeric value from a prefix for sequential checks
_NUMERIC_EXTRACT = re.compile(r'^\[?\(?(\d+|[a-zA-Z]|[ivxlcdm]+)')

_BULLET_STYLES = {0: 'List Bullet', **{i: f'List Bullet {i + 1}' for i in range(1, 5)}}
_NUMBER_STYLES = {0: 'List Number', **{i: f'List Number {i + 1}' for i in range(1, 5)}}
_NUMBERING_LEVELS = 9
_INDENT_TWIPS = 720
_HANGING_TWIPS = 360


def detect_number_format(prefix):
    """Detect OOXML numFmt from a list prefix string.

    Examines the prefix and returns the appropriate Word numbering format:
    'decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman'.
    """
    m = _NUMERIC_EXTRACT.search(prefix)
    if not m:
        return 'decimal'
    raw = m.group(1)

    # Outline patterns (e.g. "1.1", "1.1.") stay decimal
    if re.search(r'\d\.\d', prefix):
        return 'decimal'

    # Roman numerals (check before letters since 'i'/'v'/'x' are also letters)
    roman_set = {
        'i','ii','iii','iv','v','vi','vii','viii','ix','x',
        'xi','xii','xiii','xiv','xv','xvi','xvii','xviii','xix','xx',
    }
    if raw.lower() in roman_set:
        return 'lowerRoman' if raw.islower() else 'upperRoman'

    # Single letters
    if raw.isalpha() and len(raw) == 1:
        return 'lowerLetter' if raw.islower() else 'upperLetter'

    # Default: decimal
    return 'decimal'


def detect_list_prefix(text):
    """Detect if paragraph text starts with a list prefix."""
    text = _normalize_text(text)
    if not text or not text.strip():
        return None, None, text, 0

    match = _BULLET_LINE_START.match(text)
    if match:
        return 'bullet', match.group(1), text[match.end():], 0

    match = _NUMBERED_START.match(text)
    if match:
        return 'number', match.group(1).strip(), text[match.end():], 0

    return None, None, text, 0


def split_merged_list_paragraph(text):
    """Split text containing multiple list items into individual items."""
    text = _normalize_text(text)
    items = []
    pos = 0

    while pos < len(text):
        rest = text[pos:]
        bullet_match = _BULLET_ANYWHERE.search(rest)
        numbered_match = _NUMBERED_ANYWHERE.search(rest)

        matches = []
        if bullet_match:
            matches.append(('bullet', bullet_match, bullet_match.group(1) or bullet_match.group(2)))
        if numbered_match:
            matches.append(('number', numbered_match, numbered_match.group(1)))

        if not matches:
            break

        matches.sort(key=lambda x: x[1].start())
        match_type, match, prefix = matches[0]
        end_in_rest = match.end()
        absolute_end = pos + end_in_rest
        next_pos = _find_next_boundary(text, absolute_end)

        item_text = text[absolute_end:next_pos].strip()
        items.append((match_type, prefix, item_text, pos + match.start(), next_pos))
        pos = next_pos

    return items


def _find_next_boundary(text, start):
    rest = text[start:]
    candidates = []
    bm = _BULLET_ANYWHERE.search(rest)
    nm = _NUMBERED_ANYWHERE.search(rest)
    if bm:
        candidates.append(start + bm.start())
    if nm:
        candidates.append(start + nm.start())
    return min(candidates) if candidates else len(text)


def detect_indent_level(paragraph, base_left_margin=None):
    """Detect list indentation level from paragraph left indent in points."""
    pf = paragraph.paragraph_format
    left_indent = pf.left_indent
    if left_indent is None:
        return 0
    if hasattr(left_indent, 'pt'):
        left_pt = left_indent.pt
    elif hasattr(left_indent, 'emu'):
        left_pt = left_indent.emu / 12700
    else:
        return 0

    delta = left_pt - (base_left_margin or 0)
    if delta < 18:
        return 0
    if delta < 54:
        return 1
    if delta < 90:
        return 2
    if delta < 126:
        return 3
    return min(4, int(delta / 36))


def _is_likely_header(prefix_str, item_text):
    """Heuristic: is this numbered line likely a section header, not a list item?

    Signals:
      - Text contains another list-like pattern inside (describing format)
      - Text is short (<= 4 words) and title-case with no function words
      - Text is all title-case words (common header style)
    """
    stripped = item_text.strip()
    if not stripped:
        return True

    # Contains patterns like "(1. 2. 3.)" or "(a. b. c.)" -- describing list format
    if re.search(r'\([\d\.,\s]+\)', stripped) or re.search(r'\([a-z\.,\s]+\)', stripped):
        return True

    words = stripped.split()
    function_words = {'a','an','the','is','are','was','were','be','of','in','to','for',
                      'and','or','but','with','from','this','that','it','its','on','at',
                      'by','has','have','had','not','no','so','as','if','we','he','she'}

    if len(words) <= 3:
        # Short text with no function words = likely header
        if not any(w.lower() in function_words for w in words):
            return True

    # All words title-case (every word capitalized) = likely header
    if len(words) >= 2 and all(w[0].isupper() for w in words if w[0].isalpha()):
        if sum(1 for w in words if w.lower() in function_words) <= 1:
            return True

    return False


def _sequential_reset(group):
    """Check if the first numbered item's prefix suggests a reset.

    If group[0] has prefix number N > 1 and group[1] has prefix number M < N,
    the first item is likely a section header (the real list starts at group[1]).
    """
    if len(group) < 2:
        return False
    if group[0][1] != 'number':
        return False

    def _num_val(prefix):
        m = _NUMERIC_EXTRACT.search(prefix)
        if not m:
            return None
        raw = m.group(1).lower()
        roman_map = {'i':1,'ii':2,'iii':3,'iv':4,'v':5,'vi':6,'vii':7,'viii':8,'ix':9,'x':10}
        if raw in roman_map:
            return roman_map[raw]
        if raw.isalpha() and len(raw) == 1:
            return ord(raw) - ord('a') + 1
        if raw.isdigit():
            return int(raw)
        return None

    first = _num_val(group[0][2])
    second = _num_val(group[1][2])
    if first is not None and second is not None and first > second:
        return True
    # Mixed-type group: numbered header followed by bullet items
    # (e.g., "1. Section Title\n• item\n• item" merged by pdf2docx)
    if first is not None and group[1][1] == 'bullet':
        return True
    return False


def apply_builtin_list_style(paragraph, list_type, level=0, num_fmt=None, _num_id=None):
    """Apply a built-in Word list style to a paragraph.

    For non-decimal numbered lists (letters, Roman), falls through to XML
    injection since built-in styles only support decimal.

    Returns a num_id for cached reuse across a group, or None.
    """
    # For non-decimal numbered lists, the built-in styles use decimal only.
    # Fall through to XML injection which can specify the correct numFmt.
    if list_type == 'number' and num_fmt and num_fmt != 'decimal':
        return _apply_list_via_xml(paragraph, list_type, level, num_fmt, _num_id)

    style_map = _BULLET_STYLES if list_type == 'bullet' else _NUMBER_STYLES
    style_name = style_map.get(min(level, 4), style_map[0])
    try:
        paragraph.style = paragraph.part.document.styles[style_name]
    except KeyError:
        return _apply_list_via_xml(paragraph, list_type, level, num_fmt, _num_id)
    return None


def _apply_list_via_xml(paragraph, list_type, level=0, num_fmt=None, _num_id=None):
    """Apply list formatting directly via OOXML numbering injection.

    If _num_id is provided, reuses the existing numbering definition.
    Returns the num_id for reuse by subsequent items.
    """
    numbering_elm = paragraph.part.numbering_part.element
    num_id = _num_id if _num_id is not None else _add_numbering(numbering_elm, list_type, num_fmt)

    pPr = paragraph._element.get_or_add_pPr()
    pPr._remove_numPr()
    numPr = pPr.get_or_add_numPr()
    numPr.get_or_add_ilvl().val = level
    numPr.get_or_add_numId().val = num_id
    return num_id


def _add_numbering(numbering_elm, list_type, num_fmt):
    abs_num_id = _get_next_id(numbering_elm, 'w:abstractNum', 'w:abstractNumId')
    num_id = _get_next_id(numbering_elm, 'w:num', 'w:numId')
    _insert_abstract_num(numbering_elm, _build_abstract_num(abs_num_id, list_type, num_fmt))

    num = OxmlElement('w:num')
    num.set(qn('w:numId'), str(num_id))
    _mk_child(num, 'w:abstractNumId', abs_num_id)
    numbering_elm.append(num)
    return num_id


def _insert_abstract_num(numbering_elm, abs_num):
    first_num = numbering_elm.find(qn('w:num'))
    if first_num is None:
        numbering_elm.append(abs_num)
    else:
        first_num.addprevious(abs_num)


def _build_abstract_num(abs_num_id, list_type, num_fmt):
    abs_num = OxmlElement('w:abstractNum')
    abs_num.set(qn('w:abstractNumId'), str(abs_num_id))
    nsid = f'{abs_num_id + 1:08X}'
    _mk_child(abs_num, 'w:nsid', nsid)
    _mk_child(abs_num, 'w:multiLevelType', 'hybridMultilevel')
    _mk_child(abs_num, 'w:tmpl', nsid)
    for level in range(_NUMBERING_LEVELS):
        abs_num.append(_build_level(level, list_type, num_fmt or 'decimal'))
    return abs_num


def _build_level(level, list_type, num_fmt):
    lvl = OxmlElement('w:lvl')
    lvl.set(qn('w:ilvl'), str(level))
    lvl.set(qn('w:tplc'), str(level))
    _mk_child(lvl, 'w:start', 1)
    if list_type == 'bullet':
        _mk_child(lvl, 'w:numFmt', 'bullet')
        _mk_child(lvl, 'w:lvlText', '')
    else:
        _mk_child(lvl, 'w:numFmt', num_fmt)
        _mk_child(lvl, 'w:lvlText', f'%{level + 1}.')
    _mk_child(lvl, 'w:lvlJc', 'left')

    pPr = OxmlElement('w:pPr')
    ind = OxmlElement('w:ind')
    ind.set(qn('w:left'), str(_INDENT_TWIPS * (level + 1)))
    ind.set(qn('w:hanging'), str(_HANGING_TWIPS))
    pPr.append(ind)
    lvl.append(pPr)

    if list_type == 'bullet':
        rPr = OxmlElement('w:rPr')
        rFonts = OxmlElement('w:rFonts')
        for attribute in ('w:ascii', 'w:hAnsi'):
            rFonts.set(qn(attribute), 'Symbol')
        rFonts.set(qn('w:hint'), 'default')
        rPr.append(rFonts)
        lvl.append(rPr)
    return lvl


def _mk_child(parent, tag, val):
    el = OxmlElement(tag)
    el.set(qn('w:val'), str(val))
    parent.append(el)
    return el


def _get_next_id(numbering_elm, element_xpath, attr_name):
    ids = numbering_elm.xpath(f'./{element_xpath}/@{attr_name}')
    if not ids:
        return 0 if 'abstractNum' in element_xpath else 1
    return max(int(i) for i in ids) + 1


def _split_paragraph_at_list_items(paragraph):
    """Split a merged paragraph into individual list-item and non-list paragraphs.

    Handles text like "Normal text\\n• item 1\\n• item 2\\nMore normal text"
    by creating Normal paragraphs for the non-list gaps between list items.
    Also handles a single list item surrounded by non-list text.
    """
    text = paragraph.text
    items = split_merged_list_paragraph(text)

    if not items:
        return []

    # Build a complete segment list: (type, text) where type is 'bullet',
    # 'number', or 'normal'
    segments: List[Tuple[Any, ...]] = []
    prev_end = 0
    for ptype, prefix, item_text, start, end in items:
        if start > prev_end:
            gap_text = text[prev_end:start].strip()
            if gap_text:
                segments.append(('normal', gap_text))
        segments.append((ptype, prefix, item_text))
        prev_end = end
    if prev_end < len(text):
        tail_text = text[prev_end:].strip()
        if tail_text:
            segments.append(('normal', tail_text))

    # If only one segment and it's a list item with no surrounding text,
    # just replace in place
    if len(segments) == 1 and segments[0][0] != 'normal':
        ptype, prefix, item_text = segments[0]
        _replace_paragraph_text(paragraph, item_text)
        return [(paragraph, ptype, prefix, item_text)]

    if len(segments) == 1:
        # Only normal text, no list items — shouldn't happen but handle gracefully
        return [(paragraph, 'normal', '', segments[0][1])]

    results = []
    parent = paragraph._element.getparent()
    base_elem = paragraph._element

    insert_idx = 0
    for seg_idx, seg in enumerate(segments):
        if seg_idx == 0:
            if seg[0] == 'normal':
                _replace_paragraph_text(paragraph, seg[1])
                results.append((paragraph, 'normal', '', seg[1]))
            else:
                ptype, prefix, item_text = seg
                _replace_paragraph_text(paragraph, item_text)
                results.append((paragraph, ptype, prefix, item_text))
            insert_idx = 1
        else:
            new_p_elem = copy.deepcopy(base_elem)
            _remove_numPr(new_p_elem)
            parent.insert(parent.index(base_elem) + insert_idx, new_p_elem)
            insert_idx += 1
            new_para = Paragraph(new_p_elem, paragraph.part)
            if seg[0] == 'normal':
                _replace_paragraph_text(new_para, seg[1])
                results.append((new_para, 'normal', '', seg[1]))
            else:
                ptype, prefix, item_text = seg
                _replace_paragraph_text(new_para, item_text)
                results.append((new_para, ptype, prefix, item_text))

    return results


def _remove_numPr(p_elem):
    """Remove numbering properties from a paragraph XML element."""
    pPr = p_elem.find(qn('w:pPr'))
    if pPr is not None:
        for numPr in pPr.findall(qn('w:numPr')):
            pPr.remove(numPr)


def fix_lists_in_docx(docx_path, aggressive=False, verbose=False):
    """Open a DOCX, detect list paragraphs, and apply proper Word list styles.

    Uses sequential-number checking and header heuristics to avoid styling
    section headers (like "4. Numbered List") as list items.
    """
    doc = Document(str(docx_path))
    paragraphs = list(doc.paragraphs)
    fixed_count = 0

    i = 0
    while i < len(paragraphs):
        p = paragraphs[i]
        text = _normalize_text(p.text)

        if not text.strip():
            i += 1
            continue

        items = split_merged_list_paragraph(text)
        if not items:
            i += 1
            continue

        list_type = items[0][0]

        # ── Collect all consecutive items into a group ──
        group = []  # [(paragraph, list_type, prefix, item_text, orig_index)]

        # Always split (even for 1 item) to preserve surrounding non-list text
        split_results = _split_paragraph_at_list_items(p)
        paragraphs = list(doc.paragraphs)
        if split_results:
            for sp in split_results:
                group.append((sp[0], sp[1], sp[2], sp[3], i))
            # Determine list_type from the first non-normal segment
            for sp in split_results:
                if sp[1] != 'normal':
                    list_type = sp[1]
                    break
        else:
            group.append((p, items[0][0], items[0][1], items[0][2], i))

        i += 1

        while i < len(paragraphs):
            next_p = paragraphs[i]
            next_text = _normalize_text(next_p.text)
            if not next_text.strip():
                i += 1
                continue

            next_items = split_merged_list_paragraph(next_text)
            if not next_items:
                break
            if next_items[0][0] != list_type:
                break

            if len(next_items) == 1:
                ptype, prefix, item_text, _, _ = next_items[0]
                group.append((next_p, ptype, prefix, item_text, i))
            else:
                split_results = _split_paragraph_at_list_items(next_p)
                paragraphs = list(doc.paragraphs)
                for sp in split_results:
                    group.append((sp[0], sp[1], sp[2], sp[3], i))
            i += 1

        # ── Decision: style or restore ──
        try:
            should_style = _should_style_group(group, list_type, aggressive)
            # Cache numId per group so lettered/roman items share one definition
            _num_id_cache = None

            if should_style:
                for para, ptype, prefix, item_text, _ in group:
                    if ptype == 'normal':
                        continue
                    level = detect_indent_level(para)
                    if prefix and item_text.strip():
                        _replace_paragraph_text(para, item_text)
                    num_fmt = detect_number_format(prefix) if ptype == 'number' else None
                    _num_id_cache = apply_builtin_list_style(
                        para, ptype, level, num_fmt, _num_id_cache
                    )
                    fixed_count += 1
            else:
                for para, ptype, prefix, item_text, _ in group:
                    if ptype == 'normal':
                        continue
                    current = para.text.strip()
                    expected = (prefix + ' ' + item_text).strip()
                    if current != expected:
                        _replace_paragraph_text(para, expected)
        except Exception as exc:
            logger.warning("Skipping a list group that could not be styled: %s", exc)
            logger.debug("List group failure details", exc_info=True)
            continue

    # ── Cleanup pass: remove list styling from falsely-detected headers ──
    cleanup_count = _cleanup_falsely_styled_headers(doc)
    # ── Cleanup pass: split list-styled paragraphs containing non-list lines ──
    split_count = _split_list_paragraphs_with_normal_lines(doc)
    fixed_count += cleanup_count + split_count

    if fixed_count > 0:
        doc.save(docx_path)
        progress(verbose, f"Fixed {fixed_count} list items")

    return fixed_count


def _cleanup_falsely_styled_headers(doc):
    """Remove List Bullet/List Number style from paragraphs that pdf2docx
    incorrectly styled as lists but which are actually section headers.

    Only targets the 'trailing header' pattern: a short, title-case paragraph
    at the end of a list group (preceded by a long list item, followed by
    non-list content). This is very conservative to avoid removing real
    list items.
    """
    count = 0

    paragraphs = list(doc.paragraphs)
    for idx, p in enumerate(paragraphs):
        style_name = p.style.name
        if style_name not in ('List Bullet', 'List Number'):
            continue

        # Must NOT have numPr (real list items from our patch have numPr)
        pPr = p._element.find(qn('w:pPr'))
        if pPr is not None and pPr.find(qn('w:numPr')) is not None:
            continue

        text = _normalize_text(p.text).strip()
        if not text:
            continue

        # Must NOT have a list prefix in the text
        kind, prefix, _, _ = detect_list_prefix(text)
        if kind is not None:
            continue

        # ── Header shape check ──
        words = text.split()
        if len(words) > 4:
            continue
        if text.endswith(('.', ',', ';', ':', '!', '?')):
            continue
        # EVERY word must start with uppercase (title-case) — this excludes
        # "First bullet", "Dash bullet" etc. where only the first word is capitalized
        if not all(w[0].isupper() for w in words if w):
            continue

        # ── Context check: must be at the END of a list group ──
        # Preceding paragraph must be same list style with LONG text (> 30 chars)
        if idx == 0:
            continue
        prev_p = paragraphs[idx - 1]
        if prev_p.style.name != style_name:
            continue
        prev_text = _normalize_text(prev_p.text).strip()
        if len(prev_text) <= 30:
            continue

        # Following paragraph must NOT be same list style (list ended)
        if idx + 1 < len(paragraphs):
            next_p = paragraphs[idx + 1]
            if next_p.style.name == style_name:
                continue

        # All checks passed — this is a trailing header, reset to Normal
        p.style = doc.styles['Normal']
        if pPr is not None:
            for numPr in pPr.findall(qn('w:numPr')):
                pPr.remove(numPr)
        count += 1

    return count


def _split_list_paragraphs_with_normal_lines(doc):
    """Split list-styled paragraphs that contain \\n-separated non-list lines.

    pdf2docx sometimes merges a list item and subsequent normal text into one
    paragraph. This detects \\n within a List Bullet/List Number paragraph
    where the text after \\n doesn't look like a list item, and splits it into
    a separate Normal paragraph.
    """
    count = 0

    for p in list(doc.paragraphs):
        style_name = p.style.name
        if style_name not in ('List Bullet', 'List Number'):
            continue

        text = _normalize_text(p.text)
        if '\n' not in text:
            continue

        lines = text.split('\n')
        first_line = lines[0].strip()
        remaining = '\n'.join(lines[1:]).strip()
        if not remaining:
            continue

        # Check if remaining text starts with a list prefix
        kind, _, _, _ = detect_list_prefix(remaining)
        if kind is not None:
            continue  # multi-item list — handled by the main splitting logic

        # The remaining text is normal content — split it into a new paragraph
        p_elem = p._element

        # Set first line text on original paragraph
        _replace_paragraph_text(p, first_line)

        # Create a new Normal paragraph for the remaining text
        new_p = copy.deepcopy(p_elem)
        # Remove list style and numPr from the new paragraph
        new_pPr = new_p.find(qn('w:pPr'))
        if new_pPr is not None:
            for pStyle in new_pPr.findall(qn('w:pStyle')):
                new_pPr.remove(pStyle)
            for numPr in new_pPr.findall(qn('w:numPr')):
                new_pPr.remove(numPr)
        # Insert after original
        p_elem.addnext(new_p)
        # Set text on new paragraph
        new_para = Paragraph(new_p, p._parent)
        _replace_paragraph_text(new_para, remaining)
        count += 1

    return count


def _should_style_group(group, list_type, aggressive):
    """Decide whether a collected group should receive list styles.

    Rules:
      - Bullet groups: always style (bullets are unambiguous list markers)
      - Aggressive mode: always style
      - Numbered groups of 1 item: check if it's a header
      - Numbered groups of 2+ items: check for sequential reset (header first)
    """
    if list_type == 'bullet' or aggressive:
        return True

    # ── Numbered group checks ──
    if len(group) == 1:
        # Solo numbered line: style unless it looks like a header
        ptype, prefix, item_text = group[0][1], group[0][2], group[0][3]
        return not _is_likely_header(prefix, item_text)

    # 2+ numbered items: check for header as first item
    if _sequential_reset(group):
        # First item is a header — remove it, style the rest
        header = group.pop(0)
        # Restore header text
        para, ptype, prefix, item_text, _ = header
        _replace_paragraph_text(para, prefix + ' ' + item_text)
        return True

    return True


def _replace_paragraph_text(paragraph, new_text):
    """Replace paragraph text preserving formatting and hyperlinks from the first run.

    Instead of clearing all runs and creating a bare run (which loses hyperlink
    wrappers), this preserves the first run's XML ancestor chain so that
    hyperlinks, field codes, and other wrapper elements survive the text change.
    """
    if not paragraph.runs:
        paragraph.clear()
        paragraph.add_run(new_text)
        return

    first_run = paragraph.runs[0]
    first_r = first_run._element
    p_elem = paragraph._element

    # Walk up from the first run to find the direct child of the paragraph.
    # This preserves wrapper elements like w:hyperlink.
    container = first_r
    while container.getparent() is not None and container.getparent() is not p_elem:
        container = container.getparent()

    # Remove all other direct children of the paragraph, EXCEPT pPr
    # (pPr holds the paragraph style and must be preserved)
    for child in list(p_elem):
        if child is not container and child.tag != qn('w:pPr'):
            p_elem.remove(child)

    # If the container is a wrapper (e.g. hyperlink), keep only the first run
    if container is not first_r:
        for child in list(container):
            if child is not first_r:
                container.remove(child)

    # Set the text in the first run's w:t element(s)
    t_elems = first_r.findall(qn('w:t'))
    if t_elems:
        t_elems[0].text = new_text
        for t in t_elems[1:]:
            first_r.remove(t)
    else:
        t_elem = OxmlElement('w:t')
        t_elem.text = new_text
        t_elem.set(qn('xml:space'), 'preserve')
        first_r.append(t_elem)
