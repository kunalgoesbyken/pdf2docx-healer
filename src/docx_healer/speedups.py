"""Scoped speedups for pdf2docx internals that leave its output unchanged."""

import functools
import importlib
import threading
from contextlib import contextmanager
from typing import Any, Callable, Dict, Iterator, Tuple

from .log import logger

_EXTRACTOR_MODULE = "pdf2docx.image.ImagesExtractor"
_HIDE_METHOD = "_hide_page_text_and_images"
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
