import hashlib
import json
from pathlib import Path

import pytest

from bench.compat import pymupdf
from bench.fetch import FetchError, check_pdf, entry_sha256, fetch_entry, sha256_of

CORPUS = Path(__file__).resolve().parents[1] / "bench" / "corpus.json"
REQUIRED = ("id", "url", "sha256", "license", "source_page", "tests", "category")
CATEGORIES = {
    "two-column",
    "report-lists",
    "form-tables",
    "slides",
    "outline",
    "internal-links",
    "cjk",
    "rtl-arabic",
    "rtl-hebrew",
    "devanagari",
    "scanned",
    "table-heavy",
    "long",
}


@pytest.fixture(scope="module")
def entries():
    return json.loads(CORPUS.read_text(encoding="utf-8"))["documents"]


def test_corpus_size_and_schema(entries):
    assert 12 <= len(entries) <= 20
    for entry in entries:
        for key in REQUIRED:
            assert entry.get(key), (entry.get("id"), key)
        assert entry["url"].startswith("https://")
        assert len(entry["sha256"]) == 64 and int(entry["sha256"], 16) >= 0
        assert entry["source_page"].startswith("https://")


def test_corpus_ids_unique_and_every_required_category_covered(entries):
    ids = [entry["id"] for entry in entries]
    assert len(set(ids)) == len(ids)
    assert CATEGORIES <= {entry["category"] for entry in entries}


def test_slow_flag_is_set_only_on_the_long_document(entries):
    slow = [entry for entry in entries if entry.get("slow")]
    assert slow
    assert {entry["category"] for entry in slow} == {"long"}
    assert all(entry.get("slow") for entry in entries if entry["category"] == "long")


def make_pdf(path, pages=2):
    doc = pymupdf.open()
    for _ in range(pages):
        doc.new_page().insert_text((72, 100), "hello", fontsize=11)
    doc.save(str(path))
    doc.close()
    return path


def test_check_pdf_accepts_a_real_pdf_and_counts_pages(tmp_path):
    assert check_pdf(make_pdf(tmp_path / "a.pdf", 3)) == 3


def test_check_pdf_rejects_html_masquerading_as_pdf(tmp_path):
    path = tmp_path / "a.pdf"
    path.write_bytes(b"<html>blocked</html>")
    with pytest.raises(FetchError):
        check_pdf(path)


def test_check_pdf_rejects_truncated_pdf(tmp_path):
    good = make_pdf(tmp_path / "a.pdf").read_bytes()
    bad = tmp_path / "b.pdf"
    bad.write_bytes(good[: len(good) // 3])
    with pytest.raises(FetchError):
        check_pdf(bad)


def test_fetch_entry_uses_cache_and_verifies_sha(tmp_path):
    source = make_pdf(tmp_path / "src.pdf")
    entry = {"id": "doc", "url": source.as_uri(), "sha256": sha256_of(source)}
    cache = tmp_path / "cache"
    first = fetch_entry(entry, cache)
    assert first.read_bytes() == source.read_bytes()
    source.unlink()
    assert fetch_entry(entry, cache) == first


def test_fetch_entry_sha_mismatch_raises_and_leaves_no_file(tmp_path):
    source = make_pdf(tmp_path / "src.pdf")
    entry = {"id": "doc", "url": source.as_uri(), "sha256": hashlib.sha256(b"x").hexdigest()}
    cache = tmp_path / "cache"
    with pytest.raises(FetchError):
        fetch_entry(entry, cache)
    assert not list(cache.glob("*.pdf"))


def make_image_pdf(path, title):
    doc = pymupdf.open()
    page = doc.new_page()
    pixmap = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 40, 30), False)
    pixmap.set_rect(pixmap.irect, (10, 120, 200))
    page.insert_image(pymupdf.Rect(72, 72, 112, 102), pixmap=pixmap)
    doc.set_metadata({"title": title})
    doc.save(str(path))
    doc.close()
    return path


def test_image_digest_ignores_a_rebuilt_pdf_wrapper(tmp_path):
    first = make_image_pdf(tmp_path / "a.pdf", "one")
    second = make_image_pdf(tmp_path / "b.pdf", "two, rebuilt by the server")
    assert sha256_of(first) != sha256_of(second)
    entry = {"id": "scan", "url": first.as_uri(), "sha256_of": "images"}
    entry["sha256"] = entry_sha256(first, entry)
    entry["url"] = second.as_uri()
    assert fetch_entry(entry, tmp_path / "cache").exists()


def test_image_digest_still_catches_different_images(tmp_path):
    first = make_image_pdf(tmp_path / "a.pdf", "one")
    entry = {"id": "scan", "url": first.as_uri(), "sha256_of": "images"}
    entry["sha256"] = hashlib.sha256(b"other image").hexdigest()
    with pytest.raises(FetchError):
        fetch_entry(entry, tmp_path / "cache")
