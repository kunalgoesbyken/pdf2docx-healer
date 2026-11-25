"""Metrics for one (PDF, converter) pair: PDF truth against the facts of the produced DOCX."""

from typing import Any, Dict, List, Optional, Tuple

from .facts import DocxFacts, ParaFact
from .words import pages_lost, tokenize, word_scores

RATIO_KEYS = (
    "word_recall",
    "word_precision",
    "word_recall_unordered",
    "headings_recall",
    "heading_level_acc",
    "list_real",
    "list_level_ok",
    "number_ok",
    "ext_link_recall",
    "int_link_recall",
    "tables_score",
    "images_score",
)
LIST_KEYS = ("list_real", "list_level_ok", "number_ok")
WORD_KEYS = ("word_recall", "word_precision", "word_recall_unordered")


def similarity(expected: int, found: int) -> float:
    """Return min over max of two counts, 1.0 when both are zero."""
    if expected == 0 and found == 0:
        return 1.0
    return min(expected, found) / max(expected, found)


def normalize_uri(uri: str) -> str:
    """Compare URIs ignoring surrounding whitespace and a trailing slash."""
    return uri.strip().rstrip("/")


def heading_scores(
    outline: List[List[Any]], paragraphs: List[ParaFact]
) -> Tuple[Optional[float], Optional[float]]:
    """Return (recall, level accuracy) of PDF outline entries among DOCX heading paragraphs."""
    available: Dict[str, List[int]] = {}
    for paragraph in paragraphs:
        if paragraph.heading is not None:
            key = " ".join(tokenize(paragraph.text))
            available.setdefault(key, []).append(paragraph.heading)
    entries = [(int(level), " ".join(tokenize(str(title)))) for level, title in outline]
    entries = [(level, key) for level, key in entries if key]
    if not entries:
        return None, None
    found = right = 0
    for level, key in entries:
        levels = available.get(key)
        if not levels:
            continue
        found += 1
        if level in levels:
            levels.remove(level)
            right += 1
        else:
            levels.pop(0)
    return found / len(entries), right / len(entries)


def symbol_prefix(text: str) -> bool:
    """Return True when a paragraph starts with a visible marker such as a bullet."""
    first = text.lstrip()[:1]
    return bool(first) and not first.isalnum()


def match_items(
    expected: List[Dict[str, Any]], paragraphs: List[ParaFact]
) -> List[Optional[Tuple[ParaFact, str]]]:
    """Pair expected list items with paragraphs: plain (text only) or literal (label in text)."""
    index: Dict[Tuple[str, ...], List[int]] = {}
    for position, paragraph in enumerate(paragraphs):
        index.setdefault(tuple(tokenize(paragraph.text)), []).append(position)
    used = set()
    matches: List[Optional[Tuple[ParaFact, str]]] = []
    for item in expected:
        text_tokens = tokenize(item["text"])
        label_tokens = tokenize(item["label"]) if item["kind"] == "number" else []
        found = None
        for form, key in (
            ("plain", tuple(text_tokens)),
            ("literal", tuple(label_tokens + text_tokens)),
        ):
            if form == "literal" and not label_tokens:
                continue
            for position in index.get(key, []):
                if position not in used:
                    used.add(position)
                    found = (paragraphs[position], form)
                    break
            if found:
                break
        matches.append(found)
    return matches


def item_scores(
    item: Dict[str, Any], match: Optional[Tuple[ParaFact, str]]
) -> Tuple[bool, bool, bool]:
    """Return (real list, level ok, visible number ok) for one expected list item."""
    if match is None:
        return False, False, False
    paragraph, form = match
    real = paragraph.numbered and paragraph.kind == item["kind"]
    level_ok = paragraph.numbered and paragraph.ilvl == item["level"]
    if item["kind"] == "bullet":
        visible = (paragraph.numbered and paragraph.kind == "bullet") or (
            not paragraph.numbered and symbol_prefix(paragraph.text)
        )
    elif form == "literal":
        visible = not paragraph.numbered and paragraph.text.lstrip().startswith(item["label"])
    else:
        visible = (
            paragraph.numbered and paragraph.kind == "number" and paragraph.label == item["label"]
        )
    return bool(real), bool(level_ok), bool(visible)


def list_scores(
    expected: List[Dict[str, Any]], paragraphs: List[ParaFact]
) -> Dict[str, Optional[float]]:
    """Score the expected list items of a synthetic document against the produced paragraphs."""
    if not expected:
        return {key: None for key in LIST_KEYS}
    totals = [0, 0, 0]
    for item, match in zip(expected, match_items(expected, paragraphs)):
        for slot, value in enumerate(item_scores(item, match)):
            totals[slot] += int(value)
    return {key: totals[slot] / len(expected) for slot, key in enumerate(LIST_KEYS)}


def reference_words(truth: Dict[str, Any]) -> List[str]:
    """Flatten the per-page reference words."""
    return [word for page in truth["pages"] for word in page]


def defined_metrics(truth: Dict[str, Any], expected: Optional[Dict[str, Any]]) -> Dict[str, bool]:
    """Say which ratio metrics the PDF (and expectations) define."""
    defined = {key: True for key in RATIO_KEYS}
    for key in WORD_KEYS:
        defined[key] = bool(reference_words(truth))
    defined["headings_recall"] = defined["heading_level_acc"] = any(
        tokenize(str(title)) for _, title in truth["outline"]
    )
    defined["ext_link_recall"] = bool(truth["ext_uris"])
    defined["int_link_recall"] = truth["internal_links"] > 0
    for key in LIST_KEYS:
        defined[key] = bool(expected and expected.get("lists"))
    return defined


def failure_scores(
    truth: Dict[str, Any], expected: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """Scores of a run that produced no valid DOCX: every defined ratio metric is zero."""
    defined = defined_metrics(truth, expected)
    text_pages = sum(1 for page in truth["pages"] if len(page) >= 5)
    scores: Dict[str, Any] = {key: (0.0 if defined[key] else None) for key in RATIO_KEYS}
    scores.update(
        {
            "valid": False,
            "text_pages": text_pages,
            "pages_lost": text_pages,
            "list_paras": 0,
            "list_levels": [],
            "int_links_dangling": 0,
            "tables_docx": 0,
            "images_docx": 0,
            "headings_docx": 0,
            "tables_pdf": truth["tables"],
            "images_pdf": truth["images"],
            "int_links_pdf": truth["internal_links"],
        }
    )
    return scores


def score(
    truth: Dict[str, Any], facts: Optional[DocxFacts], expected: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """Compute every metric; facts=None means the conversion failed."""
    if facts is None:
        return failure_scores(truth, expected)
    defined = defined_metrics(truth, expected)
    words = word_scores(reference_words(truth), facts.words)
    text_pages, lost = pages_lost(truth["pages"], facts.words)
    heading_recall, heading_level = heading_scores(truth["outline"], facts.paragraphs)
    wanted = {normalize_uri(uri) for uri in truth["ext_uris"]}
    got = {normalize_uri(uri) for uri in facts.ext_links}
    lists = (
        list_scores(expected["lists"], facts.paragraphs)
        if expected and expected.get("lists")
        else {}
    )
    result: Dict[str, Any] = {
        "word_recall": words["recall"],
        "word_precision": words["precision"],
        "word_recall_unordered": words["recall_unordered"],
        "headings_recall": heading_recall,
        "heading_level_acc": heading_level,
        "ext_link_recall": len(wanted & got) / len(wanted) if wanted else None,
        "int_link_recall": (
            min(1.0, facts.resolved_anchors / truth["internal_links"])
            if truth["internal_links"]
            else None
        ),
        "tables_score": similarity(truth["tables"], facts.tables),
        "images_score": similarity(truth["images"], facts.images),
        "list_real": lists.get("list_real"),
        "list_level_ok": lists.get("list_level_ok"),
        "number_ok": lists.get("number_ok"),
        "valid": True,
        "text_pages": text_pages,
        "pages_lost": lost,
        "list_paras": facts.list_paras,
        "list_levels": facts.list_levels,
        "int_links_dangling": facts.dangling_anchors,
        "tables_docx": facts.tables,
        "images_docx": facts.images,
        "headings_docx": sum(1 for p in facts.paragraphs if p.heading is not None),
        "tables_pdf": truth["tables"],
        "images_pdf": truth["images"],
        "int_links_pdf": truth["internal_links"],
        "approx_formats": facts.approx_formats,
        "unresolved_numbering": facts.unresolved_numbering,
    }
    for key in RATIO_KEYS:
        if not defined[key]:
            result[key] = None
    return result
