# pdf2docx-healer

Formatting-preserving PDF-to-DOCX converter that fixes:

- Bullet and numbered lists (proper Word list styles)
- Hyperlinks (extracted from PDF annotations)
- CJK and unavailable font fallback
- Scanned/image-based PDFs (via OCR)

## Usage

```python
from docx_healer import heal

heal("input.pdf", "output.docx")
```

```bash
pdf2docx-heal input.pdf -o output.docx
```
