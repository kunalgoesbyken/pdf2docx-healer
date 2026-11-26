"""Aggregate per-document rows into per-tool numbers and a pasteable markdown scoreboard."""

from statistics import median
from typing import Any, Dict, List, Optional

from .score import RATIO_KEYS

FAILED = ("error", "crash", "timeout", "invalid")
COLUMNS = (
    ("Word recall", "word_recall", "{:.3f}"),
    ("Word precision", "word_precision", "{:.3f}"),
    ("Recall any order", "word_recall_unordered", "{:.3f}"),
    ("Heading recall", "headings_recall", "{:.2f}"),
    ("Heading level", "heading_level_acc", "{:.2f}"),
    ("List real", "list_real", "{:.2f}"),
    ("List level", "list_level_ok", "{:.2f}"),
    ("Number shown", "number_ok", "{:.2f}"),
    ("Ext links", "ext_link_recall", "{:.2f}"),
    ("Int links", "int_link_recall", "{:.2f}"),
    ("Tables", "tables_score", "{:.2f}"),
    ("Images", "images_score", "{:.2f}"),
)


def mean(values: List[float]) -> Optional[float]:
    """Return the arithmetic mean, or None for an empty list."""
    return sum(values) / len(values) if values else None


def aggregate(rows: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    """Return one aggregate per tool; failed runs score zero, unavailable runs are left out."""
    tools: Dict[str, List[Dict[str, Any]]] = {}
    for row in rows:
        tools.setdefault(row["tool"], []).append(row)
    table: Dict[str, Dict[str, Any]] = {}
    for tool, tool_rows in tools.items():
        counted = [row for row in tool_rows if row["status"] != "unavailable"]
        entry: Dict[str, Any] = {
            "n_ok": sum(1 for row in counted if row["status"] == "ok"),
            "n_failed": sum(1 for row in counted if row["status"] in FAILED),
            "n_unavailable": len(tool_rows) - len(counted),
            "pages_lost": sum(row["scores"].get("pages_lost", 0) for row in counted),
            "text_pages": sum(row["scores"].get("text_pages", 0) for row in counted),
            "seconds_median": median([row["seconds"] for row in counted if row["status"] == "ok"])
            if any(row["status"] == "ok" for row in counted)
            else None,
            "seconds_total": sum(row["seconds"] or 0 for row in counted),
        }
        for key in RATIO_KEYS:
            entry[key] = mean(
                [row["scores"][key] for row in counted if row["scores"].get(key) is not None]
            )
        table[tool] = entry
    return table


def merge_rows(existing: List[Dict[str, Any]], new: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Combine two result sets; a new row replaces the old row of the same document and tool.

    Rows stay grouped by document, documents and tools keep their order of first appearance.
    """
    combined = {(row["doc"], row["tool"]): row for row in existing + new}
    docs: Dict[str, int] = {}
    tools: Dict[str, int] = {}
    for doc, tool in combined:
        docs.setdefault(doc, len(docs))
        tools.setdefault(tool, len(tools))
    return sorted(combined.values(), key=lambda row: (docs[row["doc"]], tools[row["tool"]]))


def cell(value: Optional[float], pattern: str) -> str:
    """Format a metric cell, a dash for undefined."""
    return "-" if value is None else pattern.format(value)


def aggregate_table(table: Dict[str, Dict[str, Any]]) -> List[str]:
    """Render the per-tool table."""
    header = ["Tool", "Ok", "Failed", *[title for title, _, _ in COLUMNS], "Pages lost", "Median s"]
    lines = ["| " + " | ".join(header) + " |", "|" + " --- |" * len(header)]
    for tool, entry in table.items():
        if entry["n_ok"] + entry["n_failed"] == 0:
            lines.append(
                "| " + " | ".join([tool, "not available"] + [""] * (len(header) - 2)) + " |"
            )
            continue
        cells = [tool, str(entry["n_ok"]), str(entry["n_failed"])]
        cells += [cell(entry[key], pattern) for _, key, pattern in COLUMNS]
        cells.append(f"{entry['pages_lost']}/{entry['text_pages']}")
        cells.append(cell(entry["seconds_median"], "{:.1f}"))
        lines.append("| " + " | ".join(cells) + " |")
    return lines


def document_table(rows: List[Dict[str, Any]]) -> List[str]:
    """Render one line per (document, converter)."""
    header = [
        "Document",
        "Converter",
        "Status",
        "Recall",
        "Precision",
        "Pages lost",
        "List paras",
        "Ext links",
        "Int links",
        "Tables d/p",
        "Images d/p",
        "Seconds",
    ]
    lines = ["| " + " | ".join(header) + " |", "|" + " --- |" * len(header)]
    for row in rows:
        scores = row["scores"]
        if row["status"] == "unavailable":
            lines.append(
                f"| {row['doc']} | {row['tool']} | not available |" + " |" * (len(header) - 3)
            )
            continue
        status = row["status"] if row["status"] == "ok" else f"{row['status']}"
        lines.append(
            "| "
            + " | ".join(
                [
                    row["doc"],
                    row["tool"],
                    status,
                    cell(scores.get("word_recall"), "{:.3f}"),
                    cell(scores.get("word_precision"), "{:.3f}"),
                    f"{scores.get('pages_lost', 0)}/{scores.get('text_pages', 0)}",
                    str(scores.get("list_paras", 0)),
                    cell(scores.get("ext_link_recall"), "{:.2f}"),
                    cell(scores.get("int_link_recall"), "{:.2f}"),
                    f"{scores.get('tables_docx', 0)}/{scores.get('tables_pdf', 0)}",
                    f"{scores.get('images_docx', 0)}/{scores.get('images_pdf', 0)}",
                    cell(row["seconds"], "{:.1f}"),
                ]
            )
            + " |"
        )
    return lines


def scoreboard(rows: List[Dict[str, Any]], meta: Dict[str, Any]) -> str:
    """Return the markdown scoreboard: run metadata, the aggregate table and the document table."""
    versions = ", ".join(f"{name} {value}" for name, value in meta.get("versions", {}).items())
    lines = [
        f"Run {meta.get('date', '')}: pdf2docx-healer at commit "
        f"`{meta.get('healer_commit', meta.get('commit', ''))}`, benchmark at commit "
        f"`{meta.get('bench_commit', '')}` ({versions}).",
        "",
        *aggregate_table(aggregate(rows)),
        "",
        *document_table(rows),
        "",
    ]
    notes = sorted(
        {
            f"{row['tool']}: {row['error']}"
            for row in rows
            if row["status"] == "unavailable" and row.get("error")
        }
    )
    lines += [f"- not available, {note}" for note in notes]
    return "\n".join(lines).rstrip() + "\n"
