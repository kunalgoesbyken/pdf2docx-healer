"""OCR pipeline for scanned/image-based PDF pages."""

import tempfile
import os
import fitz


def is_scanned_page(page, min_text_chars=50):
    """Detect if a PDF page is image-based (minimal or no embedded text).
    
    Args:
        page: A fitz.Page object.
        min_text_chars: Minimum characters required to consider a page text-based.
    
    Returns:
        True if the page appears scanned/image-based.
    """
    text = page.get_text('text').strip()
    if len(text) < min_text_chars:
        return True

    # Secondary check: count text blocks
    raw = page.get_text('rawdict')
    blocks = raw.get('blocks', [])
    text_blocks = [b for b in blocks if b.get('type') == 0]
    
    if not text_blocks:
        return True

    # Count total characters across all spans
    total_chars = sum(
        len(span.get('chars', []))
        for block in text_blocks
        for line in block.get('lines', [])
        for span in line.get('spans', [])
    )
    
    return total_chars < min_text_chars


def needs_ocr(fitz_doc, threshold=0.3, min_text_chars=50):
    """Determine if a PDF document needs OCR.
    
    Args:
        fitz_doc: An opened fitz.Document.
        threshold: Fraction of textless pages that triggers full-OCR mode.
        min_text_chars: Minimum characters per page.
    
    Returns:
        True if the document should be OCR'd.
    """
    if len(fitz_doc) == 0:
        return False
    
    textless = sum(1 for page in fitz_doc if is_scanned_page(page, min_text_chars))
    return (textless / len(fitz_doc)) > threshold


def ocr_pdf(fitz_doc, language='eng', dpi=300, verbose=False):
    """OCR a PDF document using PyMuPDF's built-in Tesseract integration.
    
    Returns a new fitz.Document with OCR'd text layer.
    
    Args:
        fitz_doc: An opened fitz.Document (image-based PDF).
        language: Tesseract language code (e.g., 'eng', 'eng+chi_sim').
        dpi: OCR resolution.
        verbose: Print progress.
    
    Returns:
        A new fitz.Document with searchable text layer.
    """
    ocr_doc = fitz.open()
    
    for i in range(len(fitz_doc)):
        page = fitz_doc[i]
        
        if is_scanned_page(page):
            if verbose:
                print(f"[healer] OCR page {i + 1}/{len(fitz_doc)}...")
            
            # Render page as pixmap
            pix = page.get_pixmap(dpi=dpi)
            
            # Convert to 1-page PDF with OCR text layer
            try:
                ocr_bytes = pix.pdfocr_tobytes(
                    language=language,
                    compress=True
                )
                ocr_doc.insert_pdf(fitz.open('pdf', ocr_bytes))
            except Exception as e:
                if verbose:
                    print(f"[healer] OCR failed for page {i + 1}: {e}")
                    print("[healer] Is Tesseract installed? Falling back to image-only page.")
                # Insert original page as-is
                ocr_doc.insert_pdf(fitz_doc, from_page=i, to_page=i)
        else:
            # Page has text, copy as-is
            ocr_doc.insert_pdf(fitz_doc, from_page=i, to_page=i)
    
    return ocr_doc


def ocr_pdf_file(pdf_path, language='eng', dpi=300, verbose=False):
    """OCR a PDF file and save to a temporary path.
    
    Args:
        pdf_path: Path to input PDF.
        language: Tesseract language code.
        dpi: OCR resolution.
        verbose: Print progress.
    
    Returns:
        Path to the OCR'd temporary PDF file.
    """
    src_doc = fitz.open(pdf_path)
    
    if not needs_ocr(src_doc):
        src_doc.close()
        return pdf_path  # No OCR needed
    
    if verbose:
        print(f"[healer] Detected scanned PDF. Running OCR (lang={language}, dpi={dpi})...")
    
    try:
        ocr_doc = ocr_pdf(src_doc, language=language, dpi=dpi, verbose=verbose)
    finally:
        src_doc.close()
    
    # Write to temp file
    tmp = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
    tmp_path = tmp.name
    tmp.close()
    
    ocr_doc.save(tmp_path, garbage=4, deflate=True)
    ocr_doc.close()
    
    if verbose:
        print(f"[healer] OCR complete: {tmp_path}")
    
    return tmp_path
