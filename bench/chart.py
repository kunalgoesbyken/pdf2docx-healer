"""Draw the benchmark summary as an SVG of small bar charts, one per score."""

import argparse
import json
from html import escape
from pathlib import Path
from typing import Any, Dict, List, Sequence, Tuple

from .report import aggregate

PANELS = (
    ("Words kept in reading order", "word_recall"),
    ("Headings found", "headings_recall"),
    ("Lists that are real Word lists", "list_real"),
    ("Tables", "tables_score"),
)
FOCUS = "healer"
WIDTH = 880
COLUMNS = 2
PANEL_WIDTH = 410
LABEL_WIDTH = 140
BAR_WIDTH = 200
ROW = 26
BAR = 16
TOP = 104
PANEL_GAP = 40
STYLE = """
.surface { fill: #ffffff; stroke: #e3e6ea; }
.title { fill: #1f2933; font-weight: 700; font-size: 20px; }
.sub, .value, .label { fill: #4f5b67; font-size: 13px; }
.label.focus, .value.focus { fill: #1f2933; font-weight: 600; }
.head { fill: #1f2933; font-weight: 600; font-size: 15px; }
.base { stroke: #c3cad3; stroke-width: 1; }
.bar { fill: #8c96a1; }
.bar.focus { fill: #257de0; }
@media (prefers-color-scheme: dark) {
  .surface { fill: #191919; stroke: #33383d; }
  .title, .head, .label.focus, .value.focus { fill: #e8ecf0; }
  .sub, .value, .label { fill: #a9b3bd; }
  .base { stroke: #4a5561; }
  .bar { fill: #6b7580; }
  .bar.focus { fill: #338ced; }
}
"""


def bar_path(x: float, y: float, width: float, height: float, radius: float = 4) -> str:
    """Return a bar anchored square at x and rounded at its data end."""
    radius = min(radius, width / 2, height / 2)
    return (
        f"M{x:.1f} {y:.1f}h{width - radius:.1f}a{radius} {radius} 0 0 1 {radius} {radius}"
        f"v{height - 2 * radius:.1f}a{radius} {radius} 0 0 1 -{radius} {radius}"
        f"h-{width - radius:.1f}z"
    )


def panel(x: float, y: float, title: str, values: Sequence[Tuple[str, float]]) -> List[str]:
    """Return the SVG elements of one titled bar panel."""
    parts = [f'<text class="head" x="{x}" y="{y}">{escape(title)}</text>']
    base = x + LABEL_WIDTH
    for index, (tool, value) in enumerate(values):
        focus = " focus" if tool.startswith(FOCUS) else ""
        top = y + 14 + index * ROW
        middle = top + BAR / 2 + 4.5
        parts.append(
            f'<text class="label{focus}" x="{base - 10}" y="{middle}" '
            f'text-anchor="end">{escape(tool)}</text>'
        )
        if value > 0:
            shape = bar_path(base, top, value * BAR_WIDTH, BAR)
            parts.append(
                f'<path class="bar{focus}" d="{shape}">'
                f"<title>{escape(tool)}: {value:.2f}</title></path>"
            )
        parts.append(
            f'<text class="value{focus}" x="{base + value * BAR_WIDTH + 8:.1f}" '
            f'y="{middle}">{value:.2f}</text>'
        )
    parts.append(
        f'<line class="base" x1="{base}" x2="{base}" y1="{y + 10}" '
        f'y2="{y + 18 + len(values) * ROW - ROW + BAR}"/>'
    )
    return parts


def render(rows: List[Dict[str, Any]], meta: Dict[str, Any]) -> str:
    """Return the full SVG for one benchmark run."""
    table = aggregate(rows)
    tools = sorted(
        (tool for tool, entry in table.items() if entry["n_ok"] + entry["n_failed"]),
        key=lambda tool: -(table[tool]["word_recall"] or 0),
    )
    panel_height = 14 + len(tools) * ROW + PANEL_GAP
    rows_of_panels = -(-len(PANELS) // COLUMNS)
    height = TOP + rows_of_panels * panel_height
    pages = table[tools[0]]["text_pages"] if tools else 0
    subtitle = (
        f"{len(meta.get('documents', []))} PDFs, {pages} pages, run {meta.get('date', '')}. "
        "Higher is better. Blue is pdf2docx-healer, gray is other tools."
    )
    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{WIDTH}" height="{height}" '
        f'viewBox="0 0 {WIDTH} {height}" role="img" aria-label="Benchmark results" '
        'font-family="Segoe UI, Inter, Helvetica, Arial, sans-serif">',
        f"<style>{STYLE}</style>",
        f'<rect class="surface" x="0.5" y="0.5" width="{WIDTH - 1}" '
        f'height="{height - 1}" rx="12"/>',
        '<text class="title" x="32" y="44">PDF to Word: how much structure survives</text>',
        f'<text class="sub" x="32" y="68">{escape(subtitle)}</text>',
    ]
    for index, (title, key) in enumerate(PANELS):
        x = 32 + (index % COLUMNS) * PANEL_WIDTH
        y = TOP + (index // COLUMNS) * panel_height
        values = [(tool, table[tool][key] or 0.0) for tool in tools]
        parts += panel(x, y, title, values)
    parts.append("</svg>")
    return "\n".join(parts) + "\n"


def main() -> None:
    """Write the chart for a results file."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("results", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    data = json.loads(args.results.read_text(encoding="utf-8"))
    args.output.write_text(render(data["rows"], data.get("meta", {})), encoding="utf-8")


if __name__ == "__main__":
    main()
