from bench.facts import docx_facts
from bench.report import aggregate, scoreboard
from bench.score import RATIO_KEYS, failure_scores, score
from bench.words import tokenize
from bench_docx_builder import abstract, build_docx, lvl, num, numbering_xml, para


def truth_for(pages, **extra):
    base = {
        "pages": [tokenize(p) for p in pages],
        "outline": [],
        "ext_uris": [],
        "internal_links": 0,
        "tables": 0,
        "images": 0,
    }
    base.update(extra)
    return base


def facts_of(tmp_path, body, **kwargs):
    return docx_facts(build_docx(tmp_path / "x.docx", body, **kwargs))


def test_empty_output_scores_zero_recall_and_zero_precision(tmp_path):
    truth = truth_for(["one two three four five six seven eight"])
    result = score(truth, facts_of(tmp_path, ""))
    assert result["word_recall"] == 0.0 and result["word_precision"] == 0.0
    assert result["list_paras"] == 0


def test_pages_lost_comes_from_content(tmp_path):
    pages = [" ".join(f"p{p}w{i}" for i in range(40)) for p in range(3)]
    body = para(pages[0]) + para(pages[2])
    result = score(truth_for(pages), facts_of(tmp_path, body))
    assert result["pages_lost"] == 1 and result["text_pages"] == 3


def test_headings_recall_and_level_accuracy(tmp_path):
    from bench_docx_builder import style, styles_xml

    styles = styles_xml(style("H1", "heading 1") + style("H2", "heading 2"))
    body = para("Intro", style="H1") + para("Scope", style="H1") + para("Skip")
    truth = truth_for(
        ["Intro Scope Skip Missing"], outline=[[1, "Intro"], [2, "Scope"], [1, "Missing"]]
    )
    result = score(truth, facts_of(tmp_path, body, styles=styles))
    assert result["headings_recall"] == 2 / 3
    assert result["heading_level_acc"] == 1 / 3


def test_headings_not_defined_without_outline(tmp_path):
    result = score(truth_for(["a b c"]), facts_of(tmp_path, para("a b c")))
    assert result["headings_recall"] is None and result["heading_level_acc"] is None


def test_external_links_recall_by_distinct_uri(tmp_path):
    rels = [("rId5", "hyperlink", "https://example.com/a", True)]
    body = '<w:p><w:hyperlink r:id="rId5"><w:r><w:t>a</w:t></w:r></w:hyperlink></w:p>'
    truth = truth_for(
        ["a"], ext_uris=["https://example.com/a", "https://example.com/a", "https://example.org/b"]
    )
    result = score(truth, facts_of(tmp_path, body, rels=rels))
    assert result["ext_link_recall"] == 0.5


def test_internal_links_resolve_dangling_and_cap(tmp_path):
    body = (
        '<w:p><w:bookmarkStart w:id="1" w:name="t"/><w:bookmarkEnd w:id="1"/>'
        '<w:hyperlink w:anchor="t"><w:r><w:t>a</w:t></w:r></w:hyperlink>'
        '<w:hyperlink w:anchor="zz"><w:r><w:t>b</w:t></w:r></w:hyperlink></w:p>'
    )
    result = score(truth_for(["a b"], internal_links=4), facts_of(tmp_path, body))
    assert result["int_link_recall"] == 0.25
    assert result["int_links_dangling"] == 1


def test_tables_and_images_similarity(tmp_path):
    body = "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
    result = score(truth_for(["x"], tables=2), facts_of(tmp_path, body))
    assert result["tables_score"] == 0.5
    assert result["images_score"] == 1.0


def test_failure_scores_are_zero_for_defined_metrics_and_none_otherwise():
    truth = truth_for(["a b c d e f g h"], outline=[[1, "A"]], ext_uris=["https://x.org"], tables=1)
    failed = failure_scores(truth)
    assert failed["word_recall"] == 0.0
    assert failed["headings_recall"] == 0.0
    assert failed["ext_link_recall"] == 0.0
    assert failed["tables_score"] == 0.0
    assert failed["int_link_recall"] is None
    assert failed["valid"] is False


def synthetic_expected():
    return {
        "lists": [
            {"text": "alpha item", "label": "3.", "level": 0, "kind": "number"},
            {"text": "beta item", "label": "4.", "level": 0, "kind": "number"},
            {"text": "gamma item", "label": "•", "level": 1, "kind": "bullet"},
        ]
    }


def test_list_metrics_for_real_numbering_start_three(tmp_path):
    numbering = numbering_xml(
        abstract(0, lvl(0, start=3) + lvl(1, fmt="bullet", text="•")), num(1, 0)
    )
    body = para("alpha item", 1) + para("beta item", 1) + para("gamma item", 1, ilvl=1)
    result = score(
        truth_for(["alpha beta gamma item"]),
        facts_of(tmp_path, body, numbering=numbering),
        synthetic_expected(),
    )
    assert (result["list_real"], result["list_level_ok"], result["number_ok"]) == (1.0, 1.0, 1.0)


def test_list_metrics_wrong_start_is_caught(tmp_path):
    numbering = numbering_xml(
        abstract(0, lvl(0, start=1) + lvl(1, fmt="bullet", text="•")), num(1, 0)
    )
    body = para("alpha item", 1) + para("beta item", 1) + para("gamma item", 1, ilvl=1)
    result = score(
        truth_for(["x"]), facts_of(tmp_path, body, numbering=numbering), synthetic_expected()
    )
    assert result["list_real"] == 1.0
    assert result["number_ok"] == 1 / 3


def test_list_metrics_literal_numbers_are_visible_but_not_real(tmp_path):
    body = para("3. alpha item") + para("4. beta item") + para("• gamma item")
    result = score(truth_for(["x"]), facts_of(tmp_path, body), synthetic_expected())
    assert result["list_real"] == 0.0
    assert result["number_ok"] == 1.0
    assert result["list_level_ok"] == 0.0


def test_list_metrics_doubled_label_is_not_ok(tmp_path):
    numbering = numbering_xml(
        abstract(0, lvl(0, start=3) + lvl(1, fmt="bullet", text="•")), num(1, 0)
    )
    body = para("3. alpha item", 1) + para("beta item", 1) + para("gamma item", 1, ilvl=1)
    result = score(
        truth_for(["x"]), facts_of(tmp_path, body, numbering=numbering), synthetic_expected()
    )
    assert result["number_ok"] == 2 / 3


def test_missing_list_item_counts_against_all_list_metrics(tmp_path):
    result = score(truth_for(["x"]), facts_of(tmp_path, para("alpha item")), synthetic_expected())
    assert result["list_real"] == 0.0 and result["number_ok"] == 0.0


def row(tool, doc, status="ok", **metrics):
    base = {"doc": doc, "tool": tool, "status": status, "seconds": 1.0, "scores": metrics}
    return base


def test_aggregate_scores_failures_as_zero_and_skips_unavailable():
    rows = [
        row("a", "d1", word_recall=1.0, word_precision=1.0, valid=True, pages_lost=0, text_pages=4),
        row(
            "a",
            "d2",
            status="timeout",
            word_recall=0.0,
            word_precision=0.0,
            valid=False,
            pages_lost=5,
            text_pages=5,
        ),
        row("b", "d1", status="unavailable"),
        row("b", "d2", status="unavailable"),
    ]
    table = aggregate(rows)
    assert table["a"]["word_recall"] == 0.5
    assert table["a"]["n_failed"] == 1 and table["a"]["n_ok"] == 1
    assert table["a"]["pages_lost"] == 5
    assert table["b"]["n_unavailable"] == 2 and table["b"]["word_recall"] is None


def test_aggregate_means_ignore_undefined_metrics():
    rows = [
        row("a", "d1", ext_link_recall=None, word_recall=0.8),
        row("a", "d2", ext_link_recall=1.0, word_recall=0.6),
    ]
    table = aggregate(rows)
    assert table["a"]["ext_link_recall"] == 1.0
    assert abs(table["a"]["word_recall"] - 0.7) < 1e-9


def test_scoreboard_markdown_has_one_aggregate_table_and_one_document_table():
    rows = [
        row("a", "d1", word_recall=1.0, word_precision=1.0),
        row("b", "d1", status="unavailable"),
    ]
    text = scoreboard(rows, {"commit": "abc1234"})
    assert text.count("| Tool |") == 1
    assert text.count("| Document |") == 1
    assert "not available" in text
    assert "abc1234" in text


def test_ratio_keys_are_scored_in_failure_rows():
    truth = truth_for(
        ["a b c d e f g h"], outline=[[1, "A"]], ext_uris=["u"], tables=1, internal_links=2
    )
    failed = failure_scores(truth)
    assert all(key in failed for key in RATIO_KEYS)


def test_image_only_reference_leaves_word_metrics_undefined(tmp_path):
    truth = truth_for([""], images=1)
    result = score(truth, facts_of(tmp_path, ""))
    assert result["word_recall"] is None and result["word_precision"] is None
    assert result["images_score"] == 0.0
    assert failure_scores(truth)["word_recall"] is None
