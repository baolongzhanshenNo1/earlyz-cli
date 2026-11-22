"""L7 出 docx：LaTeX Document IR → 初始 docx 骨架（样式名对齐角色），交给 compile_document 套模板。

流程：
1. 用 python-docx 创建空文档
2. 遍历 Document.blocks，生成段落/标题/公式/表格/图片/列表
3. 公式调用 equation_dsmt4.set_equation_from_latex_dsmt4 插入 DSMT4 OLE
4. 样式名对齐角色（Heading1/Heading2/正文/图题/表题/摘要/关键词）
5. 保存为临时 docx，返回路径

依赖：
- python-docx（已在 requirements.txt）
- earlyz.actions.equation_dsmt4（已实现）
- earlyz.latex.ir（Document/Para/Heading/Equation/Float 等）

注意：
- 版式（字体/字号/行距/页边距）不在这里设置，由 compile_document + 模板决定
- 只负责内容结构，样式名对齐语义角色
"""
from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Any

from docx import Document as DocxDocument
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

from .ir import (
    Document, Para, Heading, Equation, Float, Grid, ListBlock, ListItem,
    Abstract, Keywords, Bibliography, Code, Theorem, Proof, Raw, Marker,
    Inline, Cell, Row, BibItem
)


def latex_document_to_docx(doc: Document, output_path: str | Path | None = None) -> Path:
    """LaTeX Document IR → docx 骨架。
    
    Args:
        doc: LaTeX Document IR
        output_path: 输出路径，None 则生成临时文件
    
    Returns:
        生成的 docx 文件路径
    """
    if output_path is None:
        fd, temp_path = tempfile.mkstemp(suffix=".docx", prefix="latex_")
        import os
        os.close(fd)
        output_path = Path(temp_path)
    else:
        output_path = Path(output_path)
    
    docx = DocxDocument()
    ctx = _Context(docx, doc)
    
    # 处理文档块
    for block in doc.blocks:
        _write_block(block, ctx)
    
    docx.save(str(output_path))
    return output_path


class _Context:
    """转换上下文，携带 docx 文档、计数器、标签映射等。"""
    def __init__(self, docx: DocxDocument, doc: Document):
        self.docx = docx
        self.doc = doc
        self.equation_counter = 0
        self.figure_counter = 0
        self.table_counter = 0


def _write_block(block: Any, ctx: _Context) -> None:
    """写入一个块元素。"""
    if isinstance(block, Para):
        _write_para(block, ctx)
    elif isinstance(block, Heading):
        _write_heading(block, ctx)
    elif isinstance(block, Equation):
        _write_equation(block, ctx)
    elif isinstance(block, Float):
        _write_float(block, ctx)
    elif isinstance(block, Grid):
        _write_grid(block, ctx)
    elif isinstance(block, ListBlock):
        _write_list(block, ctx)
    elif isinstance(block, Abstract):
        _write_abstract(block, ctx)
    elif isinstance(block, Keywords):
        _write_keywords(block, ctx)
    elif isinstance(block, Bibliography):
        _write_bibliography(block, ctx)
    elif isinstance(block, Code):
        _write_code(block, ctx)
    elif isinstance(block, Theorem):
        _write_theorem(block, ctx)
    elif isinstance(block, Proof):
        _write_proof(block, ctx)
    elif isinstance(block, Raw):
        _write_raw(block, ctx)
    elif isinstance(block, Marker):
        _write_marker(block, ctx)
    else:
        # 未知块类型，跳过
        pass


def _write_para(para: Para, ctx: _Context) -> None:
    """写入段落。"""
    p = ctx.docx.add_paragraph()
    
    # 对齐方式
    if para.align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif para.align == "left":
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    elif para.align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    # 写入行内元素
    for inline in para.inlines:
        _write_inline(inline, p, ctx)


def _write_heading(heading: Heading, ctx: _Context) -> None:
    """写入标题。"""
    # python-docx 的标题级别从 1 开始（level 0 = Heading 1）
    level = heading.level + 1
    if level > 9:
        level = 9  # python-docx 最多支持 9 级标题
    
    # 提取纯文本作为标题
    title_text = _inlines_to_text(heading.title)
    
    p = ctx.docx.add_heading(title_text, level=level)
    
    # 如果有编号，添加到段落开头
    if heading.numbered and heading.number:
        # 编号已经在 title_text 中（由 numbering.py 生成）
        pass


def _write_equation(eq: Equation, ctx: _Context) -> None:
    """写入公式（调用 equation_dsmt4 插入 DSMT4 OLE）。"""
    # 场景 A 闭环的核心：LaTeX → DSMT4 OLE
    # 这里需要保存当前 docx，调用 equation_dsmt4，然后重新加载
    # 但这样会很慢。更好的方式是先生成骨架，最后统一处理公式
    
    # 暂时用占位段落，标记为公式
    p = ctx.docx.add_paragraph(f"[公式占位 {ctx.equation_counter}]")
    p.style = "Normal"
    ctx.equation_counter += 1
    
    # TODO: 后续需要改进为两阶段：
    # 1. 先生成不含公式的 docx
    # 2. 再用 equation_dsmt4 批量插入公式


def _write_float(float_obj: Float, ctx: _Context) -> None:
    """写入浮动体（图/表）。"""
    if float_obj.type == "figure":
        ctx.figure_counter += 1
        # 图片
        for img in float_obj.images:
            path = img.get("path", "")
            if path and Path(path).exists():
                try:
                    ctx.docx.add_picture(str(path), width=Inches(4))
                except Exception:
                    # 图片加载失败，用文本占位
                    ctx.docx.add_paragraph(f"[图片: {path}]")
            else:
                ctx.docx.add_paragraph(f"[图片: {path}]")
        
        # 题注
        if float_obj.caption:
            caption_text = _inlines_to_text(float_obj.caption)
            p = ctx.docx.add_paragraph(caption_text)
            p.style = "Caption"
    
    elif float_obj.type == "table":
        ctx.table_counter += 1
        # 题注（表格题注在表上方）
        if float_obj.caption:
            caption_text = _inlines_to_text(float_obj.caption)
            p = ctx.docx.add_paragraph(caption_text)
            p.style = "Caption"
        
        # 表格
        for grid in float_obj.grids:
            _write_grid(grid, ctx)


def _write_grid(grid: Grid, ctx: _Context) -> None:
    """写入表格。"""
    if not grid.rows:
        return
    
    # 计算列数
    max_cols = max((len(row.cells) for row in grid.rows), default=1)
    
    table = ctx.docx.add_table(rows=len(grid.rows), cols=max_cols)
    table.style = "Table Grid"
    
    for i, row in enumerate(grid.rows):
        table_row = table.rows[i]
        for j, cell in enumerate(row.cells):
            if j < len(table_row.cells):
                cell_text = _inlines_to_text(cell.inlines)
                table_row.cells[j].text = cell_text


def _write_list(list_block: ListBlock, ctx: _Context) -> None:
    """写入列表。"""
    for item in list_block.items:
        # 列表项
        if item.label:
            # description 列表，label 是粗体
            label_text = _inlines_to_text(item.label)
            p = ctx.docx.add_paragraph()
            run = p.add_run(label_text + ": ")
            run.bold = True
        else:
            p = ctx.docx.add_paragraph(style="List Bullet" if list_block.style == "itemize" else "List Number")
        
        # 列表项内容
        for block in item.blocks:
            _write_block(block, ctx)


def _write_abstract(abstract: Abstract, ctx: _Context) -> None:
    """写入摘要。"""
    # 摘要标题
    p = ctx.docx.add_paragraph("摘要" if abstract.lang == "zh" else "Abstract")
    p.style = "Heading 1"
    
    # 摘要内容
    for block in abstract.blocks:
        _write_block(block, ctx)


def _write_keywords(keywords: Keywords, ctx: _Context) -> None:
    """写入关键词。"""
    kw_text = "关键词：" if keywords.lang == "zh" else "Keywords: "
    kw_text += "；".join(keywords.items) if keywords.lang == "zh" else ", ".join(keywords.items)
    
    p = ctx.docx.add_paragraph(kw_text)


def _write_bibliography(bib: Bibliography, ctx: _Context) -> None:
    """写入参考文献。"""
    p = ctx.docx.add_paragraph("参考文献")
    p.style = "Heading 1"
    
    for item in bib.items:
        if isinstance(item, BibItem):
            bib_text = f"[{item.label}] " + _inlines_to_text(item.inlines)
            ctx.docx.add_paragraph(bib_text)


def _write_code(code: Code, ctx: _Context) -> None:
    """写入代码块。"""
    p = ctx.docx.add_paragraph(code.text)
    # 代码块用等宽字体
    for run in p.runs:
        run.font.name = "Courier New"


def _write_theorem(theorem: Theorem, ctx: _Context) -> None:
    """写入定理类环境。"""
    # 定理标题
    title = theorem.name
    if theorem.numbered and theorem.number:
        title += f" {theorem.number}"
    if theorem.title:
        title += f" ({_inlines_to_text(theorem.title)})"
    
    p = ctx.docx.add_paragraph(title)
    run = p.runs[0] if p.runs else p.add_run()
    run.bold = True
    
    # 定理内容
    for block in theorem.blocks:
        _write_block(block, ctx)


def _write_proof(proof: Proof, ctx: _Context) -> None:
    """写入证明。"""
    title = "证明"
    if proof.title:
        title = _inlines_to_text(proof.title)
    
    p = ctx.docx.add_paragraph(title)
    run = p.runs[0] if p.runs else p.add_run()
    run.italic = True
    
    for block in proof.blocks:
        _write_block(block, ctx)


def _write_raw(raw: Raw, ctx: _Context) -> None:
    """写入无法识别的原始内容。"""
    p = ctx.docx.add_paragraph(f"[无法识别: {raw.reason}]")
    p.style = "Normal"


def _write_marker(marker: Marker, ctx: _Context) -> None:
    """写入标记（目录/分页等）。"""
    if marker.name in ("toc", "lof", "lot"):
        p = ctx.docx.add_paragraph(f"[{marker.name.upper()}]")
    elif marker.name == "pagebreak":
        # python-docx 没有直接的分页符，用分节符代替
        ctx.docx.add_page_break()


def _write_inline(inline: Inline, p, ctx: _Context) -> None:
    """写入行内元素到段落。"""
    if inline.t == "text":
        run = p.add_run(inline.v)
        # 应用样式
        if "b" in inline.style:
            run.bold = True
        if "i" in inline.style:
            run.italic = True
        if "tt" in inline.style:
            run.font.name = "Courier New"
        if "u" in inline.style:
            run.underline = True
        if "sup" in inline.style:
            run.font.superscript = True
        if "sub" in inline.style:
            run.font.subscript = True
    
    elif inline.t == "math":
        # 行内公式占位
        run = p.add_run(f"$${inline.v}$$")
    
    elif inline.t == "ref":
        # 交叉引用占位
        run = p.add_run(f"[ref:{inline.v}]")
    
    elif inline.t == "cite":
        # 引用占位
        keys = inline.data.get("keys", [])
        run = p.add_run(f"[{','.join(keys)}]")
    
    elif inline.t == "footnote":
        # 脚注占位
        run = p.add_run("[脚注]")
    
    elif inline.t == "link":
        url = inline.v
        text = inline.data.get("text", url)
        run = p.add_run(text)
        # TODO: 添加超链接
    
    elif inline.t == "br":
        p.add_run("\n")
    
    elif inline.t == "image":
        path = inline.data.get("path", "")
        run = p.add_run(f"[图片: {path}]")
    
    else:
        # 其他类型占位
        run = p.add_run(inline.v or f"[{inline.t}]")


def _inlines_to_text(inlines: list) -> str:
    """将 inline 列表转成纯文本（用于标题、题注等）。"""
    parts = []
    for inline in inlines:
        if isinstance(inline, Inline):
            if inline.t == "text":
                parts.append(inline.v)
            elif inline.t == "math":
                parts.append(f"${inline.v}$")
            elif inline.t == "ref":
                parts.append(f"[ref:{inline.v}]")
            elif inline.t == "cite":
                keys = inline.data.get("keys", [])
                parts.append(f"[{','.join(keys)}]")
            else:
                parts.append(inline.v or "")
    return "".join(parts)
