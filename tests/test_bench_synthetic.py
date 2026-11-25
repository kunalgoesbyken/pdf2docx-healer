import json

import pytest

from bench.compat import pymupdf
from bench.synthetic import build
from bench.truth import pdf_truth
from bench.words import tokenize


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    path = tmp_path_factory.mktemp("syn") / "synthetic.pdf"
    return path, build(path)


def test_expected_is_json_serialisable(built):
    json.dumps(built[1])


def test_truth_matches_the_generator(built):
    path, expected = built
    truth = pdf_truth(path)
    assert len(truth["pages"]) == 3
    assert truth["outline"] == [list(item) for item in expected["headings"]]
    assert truth["ext_uris"] == sorted(expected["ext_links"])
    assert truth["internal_links"] == expected["internal_links"]
    assert truth["tables"] == expected["tables"] == 1
    assert truth["images"] == expected["images"] == 1


def test_every_list_item_follows_its_label_in_the_pdf_text(built):
    path, expected = built
    with pymupdf.open(str(path)) as doc:
        words = [word[4] for page in doc for word in page.get_text("words")]
    assert len(expected["lists"]) >= 15
    for item in expected["lists"]:
        needle = item["text"].split()
        starts = [i for i in range(len(words)) if words[i : i + len(needle)] == needle]
        assert len(starts) == 1, item
        if item["kind"] == "number":
            assert words[starts[0] - 1] == item["label"], item


def test_list_item_texts_are_unique_and_levels_vary(built):
    items = built[1]["lists"]
    texts = [tuple(tokenize(item["text"])) for item in items]
    assert len(set(texts)) == len(texts)
    assert {item["level"] for item in items} >= {0, 1, 2}
    assert {item["kind"] for item in items} == {"number", "bullet"}


def test_numbered_labels_cover_start_three_roman_letter_and_multilevel(built):
    labels = {item["label"] for item in built[1]["lists"]}
    assert {"3.", "4.", "5.", "i.", "ii.", "a)", "b)", "1.1", "1.2", "2.1"} <= labels


def test_pdf_truth_cache_roundtrip(built, tmp_path):
    path, _ = built
    first = pdf_truth(path, cache_dir=tmp_path)
    assert list(tmp_path.glob("*.json"))
    assert pdf_truth(path, cache_dir=tmp_path) == first
