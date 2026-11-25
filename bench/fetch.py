"""Download the corpus into bench/.cache/pdfs and verify every file; nothing is executed."""

import argparse
import hashlib
import json
import shutil
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any, Dict, List, Optional

from .compat import pymupdf

ROOT = Path(__file__).resolve().parent
CORPUS = ROOT / "corpus.json"
CACHE = ROOT / ".cache" / "pdfs"
USER_AGENT = "bench-fetch/1.0 (open benchmark corpus downloader)"
RETRY_STATUS = (429, 500, 502, 503, 504)
RETRY_WAIT = (20, 60, 120, 240)
TAIL = 4096


class FetchError(Exception):
    """Raised when a corpus file cannot be downloaded or fails verification."""


def sha256_of(path: Path) -> str:
    """Return the hex SHA-256 of a file."""
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def images_sha256(path: Path) -> str:
    """Hash only the decoded page images, for servers that rebuild the PDF wrapper per request."""
    digest = hashlib.sha256()
    with pymupdf.open(str(path)) as doc:
        for page in doc:
            for image in page.get_images(full=True):
                digest.update(doc.extract_image(image[0])["image"])
    return digest.hexdigest()


def entry_sha256(path: Path, entry: Dict[str, Any]) -> str:
    """Return the digest the corpus pins for this entry: the file hash, or the image hash."""
    return images_sha256(path) if entry.get("sha256_of") == "images" else sha256_of(path)


def check_pdf(path: Path) -> int:
    """Check the header, the end marker and that PyMuPDF opens it; return the page count."""
    data = path.read_bytes()
    if not data.startswith(b"%PDF-"):
        raise FetchError(f"{path.name}: not a PDF (missing %PDF- header)")
    if b"%%EOF" not in data[-TAIL:]:
        raise FetchError(f"{path.name}: truncated PDF (no %%EOF near the end)")
    try:
        with pymupdf.open(str(path)) as doc:
            pages = len(doc)
    except Exception as exc:
        raise FetchError(f"{path.name}: PyMuPDF cannot open it: {exc}") from exc
    if pages < 1:
        raise FetchError(f"{path.name}: no pages")
    return pages


def download(url: str, target: Path) -> None:
    """Download url to target, waiting and retrying on rate limits and server errors."""
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    for attempt in range(len(RETRY_WAIT) + 1):
        try:
            with urllib.request.urlopen(request, timeout=180) as response, open(
                target, "wb"
            ) as out:
                shutil.copyfileobj(response, out)
            return
        except urllib.error.HTTPError as exc:
            if exc.code not in RETRY_STATUS or attempt == len(RETRY_WAIT):
                raise FetchError(f"{url}: HTTP {exc.code}") from exc
            time.sleep(RETRY_WAIT[attempt])
        except (urllib.error.URLError, OSError) as exc:
            raise FetchError(f"{url}: {exc}") from exc


def fetch_entry(entry: Dict[str, Any], cache_dir: Path, pin: bool = False) -> Path:
    """Return the verified local path of a corpus entry, downloading it when missing."""
    cache_dir.mkdir(parents=True, exist_ok=True)
    final = cache_dir / f"{entry['id']}.pdf"
    expected = entry.get("sha256", "")
    if final.exists():
        if pin:
            entry["pages"], entry["sha256"] = check_pdf(final), entry_sha256(final, entry)
            return final
        if entry_sha256(final, entry) == expected:
            return final
        final.unlink()
    partial = cache_dir / f"{entry['id']}.part"
    try:
        download(entry["url"], partial)
        pages = check_pdf(partial)
        actual = entry_sha256(partial, entry)
        if not pin and actual != expected:
            raise FetchError(
                f"{entry['id']}: sha256 {actual} does not match corpus.json {expected}"
            )
        entry["pages"] = pages
        entry["sha256"] = actual
        partial.replace(final)
    finally:
        partial.unlink(missing_ok=True)
    return final


def load_corpus() -> Dict[str, Any]:
    """Read bench/corpus.json."""
    return json.loads(CORPUS.read_text(encoding="utf-8"))


def select(
    corpus: Dict[str, Any], only: Optional[List[str]], include_slow: bool
) -> List[Dict[str, Any]]:
    """Choose entries by id, or every non-slow entry (plus slow ones on request)."""
    entries = corpus["documents"]
    if only:
        unknown = set(only) - {entry["id"] for entry in entries}
        if unknown:
            raise FetchError(f"unknown document ids: {sorted(unknown)}")
        return [entry for entry in entries if entry["id"] in only]
    return [entry for entry in entries if include_slow or not entry.get("slow")]


def main(argv: Optional[List[str]] = None) -> int:
    """Fetch and verify the corpus; with --pin record sha256 and page counts into corpus.json."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--only", nargs="*", help="document ids to fetch")
    parser.add_argument("--include-slow", action="store_true")
    parser.add_argument(
        "--pin", action="store_true", help="write sha256 and pages into corpus.json"
    )
    args = parser.parse_args(argv)
    corpus = load_corpus()
    failures = 0
    for entry in select(corpus, args.only, args.include_slow):
        try:
            path = fetch_entry(entry, CACHE, pin=args.pin)
            print(
                f"ok    {entry['id']:<24} {entry.get('pages', '?'):>4} pages "
                f"{path.stat().st_size:>9} bytes"
            )
        except FetchError as exc:
            failures += 1
            print(f"FAIL  {entry['id']:<24} {exc}")
    if args.pin:
        CORPUS.write_text(json.dumps(corpus, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
