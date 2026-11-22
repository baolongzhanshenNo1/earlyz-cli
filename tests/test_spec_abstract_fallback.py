# -*- coding: utf-8 -*-
"""EZI-0072 大摘要模板 body 规则缺失——全程无「摘要」region 的兜底。

真机病灶（3-2-2）：大摘要模板整篇即摘要、无「摘要」标题 → region 全程
None → 角色收集链把 role=body 段全 ``continue``，body 池 1 条 < min_cluster
→ ``rule:tpl.body`` 不生成，正文行距 360/auto 无人执行（产物保持源 480）。

修法：body 池不足 min_cluster 时重收 role=body 段（第二兜底），并同步放开
``_body_line/_body_size`` 的 cover_end 取样窗（大摘要无前后部结构，cover_end
实测吞到 66/67 段）。
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

from earlyz.spec.compilers import compile_format_spec


def _mk_doc(path: Path, body_line_twips: int = 360) -> Path:
    """大摘要形态：无「摘要」标题、无编号章、封面字段短行 + 真实摘要段。

    段落全部 role=body（无 outlineLvl）；行距走直接格式 w:spacing/@w:line。
    """
    doc = Document()
    for short in ("学        院：               ", "学        号：               "):
        p = doc.add_paragraph(short)
        p.paragraph_format.line_spacing = 1.5
    for i in range(6):
        p = doc.add_paragraph(f"本文研究能源类文本中长句英译的句式重构问题，第 {i + 1} 段"
                              "摘要正文，用于验证大摘要模板正文行距的取值与应用链路。")
        ppr = p._element.get_or_add_pPr()
        sp = ppr.find(qn("w:spacing"))
        if sp is None:
            from docx.oxml import OxmlElement
            sp = OxmlElement("w:spacing")
            ppr.append(sp)
        sp.set(qn("w:line"), str(body_line_twips))
        sp.set(qn("w:lineRule"), "auto")
        for r in p.runs:
            r.font.size = Pt(14.0)
    doc.save(str(path))
    return path


def test_abstract_fallback_creates_body_rule(tmp_path):
    """兜底触发：全程无摘要 region → rule:tpl.body 生成且 line=模板众数。"""
    tpl = _mk_doc(tmp_path / "tpl_abs.docx", body_line_twips=360)
    res = compile_format_spec("uploaded_template_docx", str(tpl), "thesis")
    body = next((r for r in res.spec.rules if r.rule_id == "rule:tpl.body"), None)
    assert body is not None, [r.rule_id for r in res.spec.rules]
    assert body.expected.get("line") == 360, body.expected
    assert body.expected.get("line_rule") == "auto", body.expected


def test_fallback_pool_format_dicts(tmp_path):
    """兜底池元素与 _collect_samples 同构（格式字典，非 IR 段对象——
    修复过程中的实锤坑：塞 IR 对象会让下游 s.get 崩/静默无规则）。"""
    from earlyz.ir.analyzer import analyze_document
    from earlyz.spec.decompiler import _abstract_body_fallback_paras
    tpl = _mk_doc(tmp_path / "tpl_abs2.docx")
    ir = analyze_document(str(tpl))
    pool = _abstract_body_fallback_paras(ir)
    assert len(pool) >= 3
    for d in pool:
        assert isinstance(d, dict), type(d)


def test_cover_window_opens_when_fallback(tmp_path):
    """兜底触发时 cover_end 失效：真摘要段（索引 0~2 的封面字段之后）
    的 360 行距须进 body 规则（修复前 cover_end 吞整篇 → line=276 离群）。"""
    tpl = _mk_doc(tmp_path / "tpl_win.docx", body_line_twips=360)
    res = compile_format_spec("uploaded_template_docx", str(tpl), "thesis")
    body = next(r for r in res.spec.rules if r.rule_id == "rule:tpl.body")
    assert body.expected.get("line") == 360, body.expected


def test_normal_template_unaffected(tmp_path):
    """零风险回归：带编号章 + 摘要标题的正常模板不走第二兜底（正文规则
    仍由原链路生成，取值不变）。"""
    doc = Document()
    h = doc.add_heading("摘要", level=1)
    for i in range(4):
        p = doc.add_paragraph(f"正常摘要正文第 {i + 1} 段，验证正常链路。")
        ppr = p._element.get_or_add_pPr()
        from docx.oxml import OxmlElement
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:line"), "300")
        sp.set(qn("w:lineRule"), "auto")
        ppr.append(sp)
    ch = doc.add_heading("1 绪论", level=1)
    for i in range(5):
        p = doc.add_paragraph(f"正常正文第 {i + 1} 段，正常模板的正文样本。")
        ppr = p._element.get_or_add_pPr()
        from docx.oxml import OxmlElement
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:line"), "300")
        sp.set(qn("w:lineRule"), "auto")
        ppr.append(sp)
    tpl = tmp_path / "tpl_normal.docx"
    doc.save(str(tpl))
    res = compile_format_spec("uploaded_template_docx", str(tpl), "thesis")
    ids = [r.rule_id for r in res.spec.rules]
    assert "rule:tpl.body" in ids, ids
    assert "rule:tpl.heading1" in ids, ids
