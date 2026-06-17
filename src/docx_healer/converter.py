"""Main ConverterHealer — orchestrates the 4-phase pipeline."""

import os
import json
import tempfile
from dataclasses import dataclass


@dataclass
class HealerConfig:
    """Configuration for the pdf2docx-healer pipeline."""
    # OCR settings
    ocr_enabled: bool = False
    ocr_lang: str = 'eng'
    ocr_dpi: int = 300
    ocr_threshold: float = 0.3
    
    # Feature toggles
    fix_lists: bool = True
    fix_hyperlinks: bool = True
    fix_fonts: bool = True
    
    # List detection
    aggressive_lists: bool = False
    
    # Output
    verbose: bool = True


def heal(pdf_path, docx_path=None, config=None):
    """Convert PDF to DOCX preserving formatting (bullets, hyperlinks, fonts).

    This is the main entry point running the 4-phase pipeline:
    1. Pre-parse: OCR scanned PDFs if needed
    2. Intercept: Monkey-patch pdf2docx for list detection
    3. Convert: Run pdf2docx with patched internals
    4. Post-process: List styles, hyperlinks, font mapping on final DOCX
    """
    if config is None:
        config = HealerConfig()
    
    if docx_path is None:
        docx_path = os.path.splitext(pdf_path)[0] + '.docx'
    
    verbose = config.verbose
    
    # === Phase 1: Pre-parse — OCR detection ===
    actual_pdf = pdf_path
    tmp_pdf = None
    
    if config.ocr_enabled:
        from .ocr_pipeline import needs_ocr, ocr_pdf_file
        import fitz
        
        src_doc = fitz.open(pdf_path)
        try:
            if needs_ocr(src_doc, threshold=config.ocr_threshold):
                src_doc.close()
                tmp_pdf = ocr_pdf_file(
                    pdf_path,
                    language=config.ocr_lang,
                    dpi=config.ocr_dpi,
                    verbose=verbose
                )
                actual_pdf = tmp_pdf
            else:
                src_doc.close()
        except Exception as e:
            if verbose:
                print(f"[healer] OCR skipped: {e}")
            try:
                src_doc.close()
            except Exception:
                pass
    
    try:
        # === Phase 2: Intercept — Apply monkey-patches ===
        if config.fix_lists:
            from .patches import apply_patches
            apply_patches(verbose=verbose)
        
        # === Phase 3: Convert ===
        _convert(actual_pdf, docx_path, verbose=verbose)
        
        # === Phase 4: Post-process ===
        from .post_processor import post_process_docx
        post_process_docx(
            docx_path,
            fix_lists=config.fix_lists,
            fix_hyperlinks=config.fix_hyperlinks,
            fix_fonts=config.fix_fonts,
            aggressive_lists=config.aggressive_lists,
            verbose=verbose,
        )
        
        if verbose:
            print(f"[healer] Done: {docx_path}")
        
        return docx_path
    
    finally:
        if tmp_pdf and os.path.exists(tmp_pdf):
            try:
                os.unlink(tmp_pdf)
            except OSError:
                pass


def _convert(pdf_path, docx_path, verbose=False):
    """Run pdf2docx conversion with default settings."""
    if verbose:
        print("[healer] Converting...")

    from pdf2docx import Converter
    if not verbose:
        import logging
        logging.getLogger().setLevel(logging.WARNING)

    cv = Converter(pdf_path)
    try:
        cv.convert(docx_path)
    finally:
        cv.close()
