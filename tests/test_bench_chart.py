from bench.chart import PANELS, render


def row(tool, recall, headings):
    scores = {"word_recall": recall, "headings_recall": headings, "text_pages": 4}
    return {"doc": "d", "tool": tool, "status": "ok", "seconds": 1.0, "scores": scores}


def test_chart_orders_by_recall_and_marks_healer_rows():
    rows = [row("pdf2docx", 0.5, 0.0), row("healer-editable", 0.9, 0.75)]
    svg = render(rows, {"documents": ["d"], "date": "2026-01-02"})
    assert svg.index(">healer-editable<") < svg.index(">pdf2docx<")
    assert 'class="label focus"' in svg and 'class="bar focus"' in svg
    assert "0.90" in svg and "0.75" in svg
    assert "1 PDFs, 4 pages, run 2026-01-02" in svg
    assert all(title in svg for title, _ in PANELS)


def test_zero_scores_draw_no_bar_but_keep_the_value():
    svg = render([row("pdf2docx", 0.0, 0.0)], {})
    assert 'class="bar' not in svg
    assert ">0.00<" in svg
