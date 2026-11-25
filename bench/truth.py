"""Ground truth read from the PDF itself: words, pages, outline, links, tables, images."""

import hashlib
import json
from pathlib import Path
from typing import Any, Dict, List, Optional

from .compat import pymupdf
from .words import tokenize

SCHEMA = 3


def sha256_file(path: Path) -> str:
    """Return the hex SHA-256 of a file."""
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def count_tables(page: Any) -> int:
    """Count the tables PyMuPDF finds on a page; a page it cannot analyse counts as none."""
    try:
        return len(page.find_tables().tables)
    except Exception:
        return 0


def read_truth(path: Path) -> Dict[str, Any]:
    """Measure one PDF."""
    pages: List[List[str]] = []
    ext_uris = set()
    internal = 0
    tables = 0
    images = 0
    image_only = 0
    with pymupdf.open(str(path)) as doc:
        outline = [[int(item[0]), str(item[1])] for item in doc.get_toc(simple=True)]
        for page in doc:
            words = tokenize(page.get_text(), keep_split=True)
            pages.append(words)
            placed = len(page.get_image_info())
            images += placed
            image_only += 1 if not words and placed else 0
            tables += count_tables(page)
            for link in page.get_links():
                if link.get("kind") == pymupdf.LINK_URI and link.get("uri"):
                    ext_uris.add(link["uri"].strip())
                elif link.get("kind") == pymupdf.LINK_GOTO:
                    internal += 1
    return {
        "schema": SCHEMA,
        "pymupdf": pymupdf.__version__,
        "pages": pages,
        "outline": outline,
        "ext_uris": sorted(ext_uris),
        "internal_links": internal,
        "tables": tables,
        "images": images,
        "image_only_pages": image_only,
    }


def pdf_truth(path: Any, cache_dir: Optional[Path] = None) -> Dict[str, Any]:
    """Return the truth of a PDF, cached on disk by content hash when cache_dir is given."""
    path = Path(path)
    if cache_dir is None:
        return read_truth(path)
    cache_dir = Path(cache_dir)
    cache_dir.mkdir(parents=True, exist_ok=True)
    cached = cache_dir / f"{sha256_file(path)}.json"
    if cached.exists():
        data = json.loads(cached.read_text(encoding="utf-8"))
        if data.get("schema") == SCHEMA and data.get("pymupdf") == pymupdf.__version__:
            return data
    data = read_truth(path)
    cached.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return data
