"""
PDF reading utilities for section-page analysis.

（EZI-0003：从仓库外归档目录的 enforce/pdf_reader.py 迁入；
归档版本本身已改用仓库内 earlyz.validation.pdf_backend（PDFium），此处保持一致。）
"""
import re
from pathlib import Path

# PDFium（earlyz.validation.pdf_backend，2026-09-30 替换 PyMuPDF/AGPL）。
# 坐标口径与 fitz 一致：左上原点、y 向下、pt；文本单位由「块」改为「行」（标题为单行，判定不变）。
from earlyz.validation.pdf_backend import read_pages


def _normalize(text: str) -> str:
    """Remove all whitespace for fuzzy matching."""
    return re.sub(r'\s+', '', text)


def get_section_start_pages(pdf_path: Path, sections: list[str]) -> list[int | None]:
    """Return the starting page number (1-based) for each section heading.

    Uses LAST match (reverse search) with y=50-200 range only.
    - y=50-200: content headings on chapter start pages
    - y<50: headers (excluded)
    - y>200: TOC entries (excluded)
    No full-page fallback — it picks up header text.
    """
    norm_headings = [_normalize(h) for h in sections]
    if not norm_headings:
        return []
    # 一次读全册文本行（句柄在 read_pages 内关闭）
    pages = read_pages(pdf_path, text=False, lines=True)
    start_pages: list[int | None] = []

    for heading_norm in norm_headings:
        found: int | None = None

        # LAST match (reverse search) avoids TOC entries before chapter content
        for info in reversed(pages):
            for x0, y0, x1, y1, text in info.lines:
                text_norm = _normalize(text)
                if heading_norm in text_norm and (50 < y0 <= 200):
                    found = info.index + 1
                    break
            if found:
                break

        start_pages.append(found)

    return start_pages
