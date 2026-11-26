"""Scoped monkey-patches to pdf2docx internals for bullet list preservation."""

import functools
import importlib
import inspect
import threading
from contextlib import contextmanager
from importlib import metadata
from typing import Any, Callable, Iterator, Optional

from .errors import IncompatiblePdf2docxError
from .list_markers import parse_marker
from .log import logger

_BLOCKS_MODULE = "pdf2docx.layout.Blocks"
PATCH_LOCK = threading.RLock()


def _is_list_item(text: Any, bullets: bool = True, numbers: bool = True) -> Optional[str]:
    """Return the marker when a row starts with a list marker, else None; never raises."""
    if not isinstance(text, str):
        return None
    try:
        marker = parse_marker(text, bare=True)
    except Exception:
        logger.debug("List marker parsing failed and was ignored", exc_info=True)
        return None
    if marker is None or not (bullets if marker.is_bullet else numbers):
        return None
    return text[: marker.end].strip()


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
    with PATCH_LOCK:
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


_SPANS_MODULE = "pdf2docx.text.Spans"
_TEXT_SPAN_MODULE = "pdf2docx.text.TextSpan"
_IMAGE_SPAN_MODULE = "pdf2docx.image.ImageSpan"


def _resolve_span_types() -> "Optional[tuple[Any, Any, Any]]":
    try:
        spans = importlib.import_module(_SPANS_MODULE).Spans
        text_span = importlib.import_module(_TEXT_SPAN_MODULE).TextSpan
        image_span = importlib.import_module(_IMAGE_SPAN_MODULE).ImageSpan
    except (ImportError, AttributeError):
        return None
    if not inspect.isfunction(vars(spans).get("restore")):
        return None
    return spans, text_span, image_span


@contextmanager
def keep_word_spaces() -> Iterator[None]:
    """Keep the blank spans between two visible spans of a line, then restore pdf2docx.

    PyMuPDF reports the gap between two words as its own whitespace-only span. pdf2docx
    drops every unstyled blank span, so the words around it run together. Blank spans at
    the start or end of a line are still dropped. Does nothing when the installed pdf2docx
    has a different layout.
    """
    with PATCH_LOCK:
        types = _resolve_span_types()
        if types is None:
            logger.debug("pdf2docx Spans.restore not found; blank spans are not kept")
            yield
            return
        spans_class, text_span, image_span = types
        original = vars(spans_class)["restore"]

        def restore(self: Any, raws: Any) -> Any:
            built = []
            for raw in raws:
                if "image" in raw:
                    built.append((image_span(raw), True))
                else:
                    span = text_span(raw)
                    built.append((span, bool(span.text.strip()) or bool(span.style)))
            visible = [index for index, (_, shown) in enumerate(built) if shown]
            first, last = (visible[0], visible[-1]) if visible else (0, -1)
            for index, (span, shown) in enumerate(built):
                self.append(span if shown or first < index < last else None)
            return self

        spans_class.restore = restore
        try:
            yield
        finally:
            spans_class.restore = original
