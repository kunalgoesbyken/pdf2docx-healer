"""Post-processing passes that repair the generated DOCX in place."""

from dataclasses import dataclass, field
from typing import Callable, List, Optional

from .files import PathLike
from .font_mapper import fix_docx_fonts
from .headings import fix_docx_headings
from .internal_links import fix_docx_internal_links
from .link_injector import fix_docx_hyperlinks
from .list_detector import fix_lists_in_docx
from .log import logger
from .tables import fix_docx_tables


@dataclass
class PostProcessReport:
    """Counts of fixes applied and the names of passes that failed."""

    tables_flattened: int = 0
    headings_fixed: int = 0
    lists_fixed: int = 0
    hyperlinks_added: int = 0
    internal_links_added: int = 0
    fonts_fixed: int = 0
    failed_passes: List[str] = field(default_factory=list)


def _run_pass(report: PostProcessReport, name: str, docx_path: PathLike,
              action: Callable[[], int]) -> int:
    try:
        return action()
    except Exception as exc:
        report.failed_passes.append(name)
        logger.warning("The %s pass failed and was skipped for %s: %s", name, docx_path, exc)
        logger.debug("%s pass failure details", name, exc_info=True)
        return 0


def post_process_docx(
    docx_path: PathLike,
    pdf_path: Optional[PathLike] = None,
    fix_tables: bool = True,
    fix_headings: bool = True,
    fix_lists: bool = True,
    fix_hyperlinks: bool = True,
    fix_internal_links: bool = True,
    fix_fonts: bool = True,
    aggressive_lists: bool = False,
    verbose: bool = False,
) -> PostProcessReport:
    """Run all post-processing passes on a generated DOCX, modifying it in place.

    Passes that read the source PDF (tables, headings, internal links) run only when
    ``pdf_path`` is given. Headings run before lists so numbered headings are never
    restyled as list items. Each pass opens and saves the file on its own, so a failing
    pass leaves the file as the previous pass left it. Failures are logged as warnings
    and listed in the returned report instead of aborting the remaining passes.
    """
    report = PostProcessReport()
    has_pdf = pdf_path is not None

    if fix_tables and has_pdf:
        report.tables_flattened = _run_pass(
            report, "tables", docx_path,
            lambda: fix_docx_tables(docx_path, pdf_path, verbose=verbose),
        )
    if fix_headings and has_pdf:
        report.headings_fixed = _run_pass(
            report, "headings", docx_path,
            lambda: fix_docx_headings(docx_path, pdf_path, verbose=verbose),
        )
    if fix_lists:
        report.lists_fixed = _run_pass(
            report, "lists", docx_path,
            lambda: fix_lists_in_docx(docx_path, aggressive=aggressive_lists, verbose=verbose),
        )
    if fix_hyperlinks:
        report.hyperlinks_added = _run_pass(
            report, "hyperlinks", docx_path,
            lambda: fix_docx_hyperlinks(docx_path, verbose=verbose),
        )
    if fix_internal_links and has_pdf:
        report.internal_links_added = _run_pass(
            report, "internal links", docx_path,
            lambda: fix_docx_internal_links(docx_path, pdf_path, verbose=verbose),
        )
    if fix_fonts:
        report.fonts_fixed = _run_pass(
            report, "fonts", docx_path,
            lambda: fix_docx_fonts(docx_path, verbose=verbose),
        )
    return report
