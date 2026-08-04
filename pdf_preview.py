"""Prévia em imagem para PDFs já gerados pelo Excel."""
from __future__ import annotations


def pdf_page_count(pdf_bytes: bytes) -> int:
    import fitz
    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        return document.page_count
    finally:
        document.close()


def render_pdf_page(pdf_bytes: bytes, page_number: int = 0, crop_to_content: bool = False,
                    zoom: float = 1.0) -> bytes:
    import fitz
    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        page = document.load_page(page_number)
        clip = None
        if crop_to_content:
            rectangles = [fitz.Rect(block[:4]) for block in page.get_text("blocks")]
            # Linhas e bordas não fazem parte dos blocos de texto. Incluí-las
            # evita cortar o contorno real da planilha no PNG.
            rectangles.extend(drawing["rect"] for drawing in page.get_drawings() if not drawing["rect"].is_empty)
            if rectangles:
                content = rectangles[0]
                for rectangle in rectangles[1:]:
                    content.include_rect(rectangle)
                # Margem branca visível ao redor do print, sem a área vazia da página.
                margin = 14
                clip = fitz.Rect(
                    max(page.rect.x0, content.x0 - margin), max(page.rect.y0, content.y0 - margin),
                    min(page.rect.x1, content.x1 + margin), min(page.rect.y1, content.y1 + margin),
                )
        scale = 1.5 * max(0.5, min(2.5, zoom))
        return page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clip, alpha=False).tobytes("png")
    finally:
        document.close()
