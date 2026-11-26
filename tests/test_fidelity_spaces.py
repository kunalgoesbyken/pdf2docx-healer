from typing import Any, Dict, List

import pytest
from pdf2docx.text.Spans import Spans

from convert_support import body_text
from docx_healer import HealerConfig, heal
from docx_healer.patches import keep_word_spaces
from layout_support import make_spaced_words_pdf


def texts(spans: Spans) -> List[str]:
    return [span.text for span in spans if span is not None]


def raw(text: str, styled: bool = False) -> Dict[str, Any]:
    span: Dict[str, Any] = {"text": text, "bbox": (0, 0, 10, 10), "size": 10}
    if styled:
        span["style"] = [{"type": 1, "color": 0}]
    return span


def test_words_placed_apart_stay_separate_words(tmp_path):
    pdf = make_spaced_words_pdf(tmp_path / "in.pdf", ["alpha", "bravo", "charlie", "delta"])

    heal(pdf, tmp_path / "out.docx", HealerConfig(verbose=False))

    assert "alpha bravo charlie delta" in body_text(tmp_path / "out.docx")


@pytest.mark.parametrize(
    ("spans", "expected"),
    [
        ([raw("a"), raw(" "), raw("b")], ["a", " ", "b"]),
        ([raw(" "), raw("a"), raw(" "), raw("b"), raw(" ")], ["a", " ", "b"]),
        ([raw(" "), raw("  ")], []),
        ([raw("a"), raw(" "), raw(" "), raw("b")], ["a", " ", " ", "b"]),
        ([raw(" ", styled=True), raw("a")], [" ", "a"]),
        ([raw("a"), raw(" ", styled=True)], ["a", " "]),
        ([], []),
        ([raw("全"), raw(" "), raw("文")], ["全", " ", "文"]),
    ],
    ids=[
        "interior",
        "edges-dropped",
        "only-blanks",
        "two-blanks",
        "styled-leading",
        "styled-trailing",
        "empty",
        "cjk",
    ],
)
def test_only_interior_blank_spans_are_kept(spans, expected):
    with keep_word_spaces():
        restored = Spans().restore(spans)

    assert texts(restored) == expected


def test_blank_span_handling_is_restored_on_exit():
    before = Spans.restore
    with keep_word_spaces():
        assert Spans.restore is not before
    assert Spans.restore is before
    assert texts(Spans().restore([raw("a"), raw(" "), raw("b")])) == ["a", "b"]


def test_no_change_when_pdf2docx_has_another_layout(monkeypatch):
    monkeypatch.delattr(Spans, "restore")
    with keep_word_spaces():
        assert "restore" not in vars(Spans)
