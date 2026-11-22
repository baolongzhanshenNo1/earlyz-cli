# -*- coding: utf-8 -*-
"""EZI-0073 星号内联格式注记说明行剥除 + 手写目录单制表符形态。

两案真机实锤：
- 计算机学报：源段「1.1  二级标题 *字体为5号黑体*标题2」**带 outlineLvl=1**
  （模板作者把说明样例行设成真标题级别）——0071 标题豁免只让位 `_spec`，
  星号内联形态漏网 → 混进 L2 标题样本 → 字体众数被带偏。
- 重庆科技：手写目录行「1 正文格式说明\\t1」是**单制表符两段**——旧正则要求
  ≥2 个 tab → 漏判混进 L1 样本 → jc 众数 2/4 打平带偏。
"""

from __future__ import annotations

import zipfile
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def _mk(path: Path, paras: list[tuple[str, int | None]]) -> Path:
    """造 docx：每段 (text, outline_lvl)；tab 以**字面量**写进 ``w:t``
    （真重庆案形态：flow_text 只拼 ``w:t`` 不转 ``<w:tab/>`` 元素——
    python-docx 的 ``	`` 会被转成 ``w:tab`` 元素导致夹具失真）。"""
    doc = Document()
    for text, lvl in paras:
        p = doc.add_paragraph()
        r = p.add_run()
        wt = r._element.makeelement(qn("w:t"), {qn("xml:space"): "preserve"})
        wt.text = text
        r._element.append(wt)
        if lvl is not None:
            ppr = p._element.get_or_add_pPr()
            ol = OxmlElement("w:outlineLvl")
            ol.set(qn("w:val"), str(lvl))
            ppr.append(ol)
    doc.save(str(path))
    return path


def test_star_inline_predicate_shapes():
    from earlyz.actions.new_actions import _star_inline_format_note
    assert _star_inline_format_note("1.1  二级标题 *字体为5号黑体*标题2")
    assert _star_inline_format_note("1.1.1  三级标题 *字体为5号宋体*标题3")
    assert _star_inline_format_note("附录A. 附录内容名称 *字体为黑体三号*")
    for neg in ("1.1 图的格式说明", "2.1 实验结果与分析",
                "参考文献", "表 5-1 实验参数设置"):
        assert not _star_inline_format_note(neg), neg


def test_strip_removes_star_inline_with_outline(tmp_path):
    """带 outlineLvl 的说明样例行也要剥（豁免让位谓词）。"""
    from earlyz.actions.new_actions import strip_instruction_paragraphs
    fx = _mk(tmp_path / "a.docx", [
        ("1.1  二级标题 *字体为5号黑体*标题2", 1),
        ("1.1.1  三级标题 *字体为5号宋体*标题3", 2),
        ("1.1 图的格式说明", 1),
        ("2.1 实验结果与分析", 2),
    ])
    r = strip_instruction_paragraphs(fx, {})
    assert r["changed"] and r["removed"] == 2, r
    x = zipfile.ZipFile(fx).read("word/document.xml").decode("utf-8")
    assert "字体为5号黑体" not in x
    assert "图的格式说明" in x and "实验结果与分析" in x


def test_strip_idempotent(tmp_path):
    from earlyz.actions.new_actions import strip_instruction_paragraphs
    fx = _mk(tmp_path / "a.docx", [("1.1  二级标题 *字体为5号黑体*标题2", 1)])
    r1 = strip_instruction_paragraphs(fx, {})
    r2 = strip_instruction_paragraphs(fx, {})
    assert r1["changed"] and r2["removed"] == 0


def test_manual_toc_single_tab_form():
    """单制表符「标题\\t页码」也算手写目录条目（重庆实锤）。"""
    from earlyz.validation.layout_fidelity import _MANUAL_TOC_ENTRY_RE
    for pos in ("1 正文格式说明\t1", "1.1 论文格式基本要求\t1",
                "1.1.\t编写目的\t4", "2.4.2 中华人民共和国法定计量单位\t6"):
        assert _MANUAL_TOC_ENTRY_RE.match(pos), pos
    for neg in ("1 正文格式说明", "图 3-1 系统架构图",
                "本文研究基于深度学习方法\t2023 年的数据集"):
        assert not _MANUAL_TOC_ENTRY_RE.match(neg), neg


def test_heading_samples_exclude_single_tab_toc(tmp_path):
    """标题样本池：单 tab 目录行排除、真标题保留。"""
    from earlyz.validation.layout_fidelity import _load_docx, heading_samples_by_level
    fx = _mk(tmp_path / "h.docx", [
        ("1 正文格式说明\t1", 0),
        ("2 图表及公式的格式说明\t3", 0),
        ("1 正文格式说明", 0),
        ("2 图表及公式的格式说明", 0),
    ])
    doc = _load_docx(fx)
    samples, excluded = heading_samples_by_level(doc)
    texts = [s["text"] for s in samples[1]]
    assert texts == ["1 正文格式说明", "2 图表及公式的格式说明"], texts
    assert excluded["手写目录条目"] == 2, excluded
