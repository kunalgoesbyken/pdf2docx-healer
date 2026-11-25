"""Scoped speedups for pdf2docx and python-docx internals that leave the output unchanged."""

import functools
import importlib
import threading
from contextlib import contextmanager
from typing import Any, Callable, Dict, Iterator, Optional, Tuple

from .log import logger

_EXTRACTOR_MODULE = "pdf2docx.image.ImagesExtractor"
_HIDE_METHOD = "_hide_page_text_and_images"
_TABLE_MODULE = "docx.table"
_CELLS = "_cells"
_CELLS_CACHE = "_docx_healer_cells"
_LOCK = threading.RLock()


class _CachedImageListing:
    """Proxy a PyMuPDF page and answer repeated ``get_images`` calls from a cache."""

    def __init__(self, page: Any) -> None:
        self._page = page
        self._listings: Dict[Tuple[Any, ...], Any] = {}

    def get_images(self, *args: Any, **kwargs: Any) -> Any:
        key = (args, tuple(sorted(kwargs.items())))
        if key not in self._listings:
            self._listings[key] = self._page.get_images(*args, **kwargs)
        return self._listings[key]

    def __getattr__(self, name: str) -> Any:
        return getattr(self._page, name)


def _resolve_extractor() -> Any:
    try:
        module = importlib.import_module(_EXTRACTOR_MODULE)
    except ImportError:
        return None
    extractor = getattr(module, "ImagesExtractor", None)
    if extractor is None or not isinstance(vars(extractor).get(_HIDE_METHOD), staticmethod):
        return None
    return extractor


@contextmanager
def cached_image_listing() -> Iterator[None]:
    """List a page's images once per pdf2docx hide-images pass instead of once per content stream.

    pdf2docx calls ``page.get_images`` inside a loop over every xobject of the page, which
    makes drawing-heavy pages quadratic. The listing does not depend on the streams being
    edited, so answering it from a per-call cache returns the same result. Does nothing when
    the installed pdf2docx has a different layout.
    """
    with _LOCK:
        extractor = _resolve_extractor()
        if extractor is None:
            logger.debug("pdf2docx image extractor not found; skipping the image listing cache")
            yield
            return

        original = vars(extractor)[_HIDE_METHOD]
        wrapped: Callable[..., Any] = original.__func__

        @functools.wraps(wrapped)
        def hide(page: Any, *args: Any, **kwargs: Any) -> Any:
            return wrapped(_CachedImageListing(page), *args, **kwargs)

        setattr(extractor, _HIDE_METHOD, staticmethod(hide))
        try:
            yield
        finally:
            setattr(extractor, _HIDE_METHOD, original)


def _resolve_table_types() -> Optional[Tuple[Any, Any]]:
    try:
        module = importlib.import_module(_TABLE_MODULE)
        table_class, cell_class = module.Table, module._Cell
    except (ImportError, AttributeError):
        return None
    cells = vars(table_class).get(_CELLS)
    if not isinstance(cells, property) or cells.fget is None:
        return None
    if not callable(vars(cell_class).get("merge")):
        return None
    return table_class, cell_class


@contextmanager
def cached_table_cells() -> Iterator[None]:
    """Build the cell grid of a python-docx table once instead of once per ``table.cell`` call.

    ``Table.cell`` rebuilds the whole grid on every call and pdf2docx calls it several times
    per cell, which makes big tables quadratic. The grid only changes when cells are merged,
    so it is kept on the table and dropped after every ``_Cell.merge``. Does nothing when
    the installed python-docx has a different layout.
    """
    with _LOCK:
        types = _resolve_table_types()
        if types is None:
            logger.debug("python-docx table layout not recognised; skipping the cell grid cache")
            yield
            return
        table_class, cell_class = types
        original_cells = vars(table_class)[_CELLS]
        original_merge = vars(cell_class)["merge"]

        def cells(table: Any) -> Any:
            grid = table.__dict__.get(_CELLS_CACHE)
            if grid is None:
                grid = table.__dict__[_CELLS_CACHE] = original_cells.fget(table)
            return grid

        @functools.wraps(original_merge)
        def merge(cell: Any, *args: Any, **kwargs: Any) -> Any:
            try:
                return original_merge(cell, *args, **kwargs)
            finally:
                getattr(cell._parent, "__dict__", {}).pop(_CELLS_CACHE, None)

        setattr(table_class, _CELLS, property(cells))
        cell_class.merge = merge
        try:
            yield
        finally:
            setattr(table_class, _CELLS, original_cells)
            cell_class.merge = original_merge
