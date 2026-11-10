"""Scoped monkey-patches to pdf2docx internals for bullet list preservation."""

import functools
import importlib
import inspect
import re
import threading
from contextlib import contextmanager
from importlib import metadata
from typing import Any, Callable, Iterator, Optional

from .constants import BULLET_CODEPOINTS
from .errors import IncompatiblePdf2docxError

_MARKER_CHARS = frozenset("-*+~>»")
_BULLET_MAP = {0xF0B7: "•"}
_NUMBERED_PREFIX = re.compile(
    r"^(?:\d{1,3}\.\d{1,3}[.)]|\d{1,3}[.)]|[a-z][.)]|[ivxlcdm]+[.)]"
    r"|\((?:\d{1,3}|[a-z]|[ivxlcdm]+)\))(?:\s|$)",
    re.IGNORECASE,
)
_BLOCKS_MODULE = "pdf2docx.layout.Blocks"
_PATCH_LOCK = threading.RLock()


def _is_list_item(text: str, bullets: bool = True, numbers: bool = True) -> Optional[str]:
    """Return the marker when ``text`` starts with a bullet or number prefix, else None."""
    stripped = text.strip() if text else ""
    if not stripped:
        return None

    first = stripped[0]
    if bullets and ord(first) in BULLET_CODEPOINTS:
        return _BULLET_MAP.get(ord(first), first)
    if bullets and first in _MARKER_CHARS and (len(stripped) == 1 or stripped[1].isspace()):
        return first
    if numbers and _NUMBERED_PREFIX.match(stripped):
        return stripped
    return None


def _pdf2docx_version() -> str:
    try:
        return metadata.version("pdf2docx")
    except metadata.PackageNotFoundError:
        return "unknown"


def _incompatible(reason: str) -> IncompatiblePdf2docxError:
    return IncompatiblePdf2docxError(
        f"pdf2docx {_pdf2docx_version()} is not compatible with the list patches: {reason}. "
        "Install a supported version with: pip install 'pdf2docx>=0.5.9,<0.6' "
        "or disable list fixing (fix_lists=False / --no-lists)."
    )


def _resolve_targets() -> "tuple[Any, Any, Callable[..., Any]]":
    try:
        blocks_module = importlib.import_module(_BLOCKS_MODULE)
    except ImportError as exc:
        raise _incompatible(f"cannot import {_BLOCKS_MODULE}") from exc

    blocks_class = getattr(blocks_module, "Blocks", None)
    collect = getattr(blocks_class, "collect_stream_lines", None)
    if not callable(collect):
        raise _incompatible("Blocks.collect_stream_lines is missing")
    if not callable(getattr(blocks_module, "is_list_item", None)):
        raise _incompatible(f"{_BLOCKS_MODULE}.is_list_item is missing")

    parameters = inspect.signature(collect).parameters.values()
    accepts_kwargs = any(p.kind is inspect.Parameter.VAR_KEYWORD for p in parameters)
    if not accepts_kwargs:
        raise _incompatible("Blocks.collect_stream_lines does not accept the list_not_table option")
    return blocks_module, blocks_class, collect


@contextmanager
def list_item_patches() -> Iterator[None]:
    """Patch pdf2docx so bullet rows are not turned into tables, then restore it.

    Raises IncompatiblePdf2docxError before touching anything when the installed
    pdf2docx lacks the expected internals.
    """
    with _PATCH_LOCK:
        blocks_module, blocks_class, original_collect = _resolve_targets()
        original_is_list_item = blocks_module.is_list_item

        @functools.wraps(original_collect)
        def collect_without_list_tables(self: Any, *args: Any, **kwargs: Any) -> Any:
            return original_collect(self, *args, **{**kwargs, "list_not_table": True})

        blocks_module.is_list_item = _is_list_item
        blocks_class.collect_stream_lines = collect_without_list_tables
        try:
            yield
        finally:
            blocks_module.is_list_item = original_is_list_item
            blocks_class.collect_stream_lines = original_collect
