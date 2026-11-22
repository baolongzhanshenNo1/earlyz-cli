# -*- coding: utf-8 -*-
"""自制最小示例对（合成内容——与任何真实学校/期刊/合同模板无关）。

运行：python examples/make_examples.py
产物：examples/示例论文.docx（裸内容） + examples/示例模板.docx（规范模板）
试排：earlyz format examples/示例论文.docx --spec examples/示例模板.docx
"""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, Twips

HERE = Path(__file__).resolve().parent


def _set_font(run, east: str, size: float, *, bold: bool = False) -> None:
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    run.font.bold = bold
    rpr = run._r.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), east)


def _style(p, *, outline: int | None = None) -> None:
    ppr = p._p.get_or_add_pPr()
    if outline is not None:
        ol = ppr.makeelement(qn("w:outlineLvl"), {qn("w:val"): str(outline)})
        ppr.append(ol)


def build_paper(path: Path) -> None:
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin, sec.bottom_margin = Twips(1440), Twips(1440)
    sec.left_margin, sec.right_margin = Twips(1800), Twips(1800)
    p = doc.add_paragraph("示例论文题名")
    _style(p, outline=0)
    _set_font(p.runs[0], "黑体", 16.0, bold=True)
    for text, lvl, east, size in (
        ("1 引言", 0, "黑体", 14.0),
        ("本文用于演示 EarlyZ CLI 的排版链路。", None, "宋体", 12.0),
        ("1.1 研究背景", 1, "黑体", 12.0),
        ("这一小节演示二级标题与正文的取值迁移。", None, "宋体", 12.0),
        ("2 方法", 0, "黑体", 14.0),
        ("正文段落若干，用于演示行距与缩进。", None, "宋体", 12.0),
    ):
        p = doc.add_paragraph(text)
        _style(p, outline=lvl)
        _set_font(p.runs[0], east, size)
    doc.save(str(path))


def build_template(path: Path) -> None:
    """规范模板：与示例论文同构，但取值不同（题名居中、正文 10.5pt、页边距收紧）。"""
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin, sec.bottom_margin = Twips(1580), Twips(1020)
    sec.left_margin, sec.right_margin = Twips(1300), Twips(1300)
    rows = (
        ("示例论文题名（模板样式行）", 0, "黑体", 15.0, WD_ALIGN_PARAGRAPH.CENTER),
        ("1 引言", 0, "黑体", 13.0, WD_ALIGN_PARAGRAPH.LEFT),
        ("模板正文样例段落，宋体五号。", None, "宋体", 10.5, WD_ALIGN_PARAGRAPH.JUSTIFY),
        ("1.1 研究背景", 1, "黑体", 10.5, WD_ALIGN_PARAGRAPH.LEFT),
        ("模板二级标题下的正文样例。", None, "宋体", 10.5, WD_ALIGN_PARAGRAPH.JUSTIFY),
        ("2 方法", 0, "黑体", 13.0, WD_ALIGN_PARAGRAPH.LEFT),
        ("模板第二个正文段落。", None, "宋体", 10.5, WD_ALIGN_PARAGRAPH.JUSTIFY),
    )
    for text, lvl, east, size, align in rows:
        p = doc.add_paragraph(text)
        p.alignment = align
        _style(p, outline=lvl)
        _set_font(p.runs[0], east, size)
    doc.save(str(path))


if __name__ == "__main__":
    build_paper(HERE / "示例论文.docx")
    build_template(HERE / "示例模板.docx")
    print("示例对已生成：examples/示例论文.docx + examples/示例模板.docx")
