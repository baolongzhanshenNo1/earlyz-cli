# -*- coding: utf-8 -*-
"""EZI-0071 格式说明段识别谓词——源剥除第四闸 / 内容 hash 豁免 / 判定层
正文池排除**三消费同源**。

真机形态库：3-2-1 源 18 条实锤（「字体：黑体；字号：三号；间距：1.5倍
行距，段前12磅」「存档论文打印装订…」），跨语料真文负检（矿大/人社部/
汕尾/期刊 287 长段）零误判。
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="xml" ContentType="application/xml"/></Types>')
RELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')

# 校准正样本（3-2-1 源实测形态）
POS = [
    "字体：黑体；字号：三号；间距：1.5倍行距，段前12磅，段后6磅。",
    '标题“摘要”：黑体，三号，居中，段前12磅，段后6磅，1.5倍行距',
    "英文题目：Times New Roman，四号1.5倍行距，所有下划线所填信息中文为楷体GB2312",
    "存档论文打印装订的学位论文左侧面需有书脊，书脊文字从上向下为：硕士"
    "专业学位论文  英文论文题目  姓名  中国矿业大学（北京）  20XX年6月。字号：",
    "下划线所填信息：（1）字体：楷体GB2312，字母与数字为Times New Roman；字号：四号",
    "奇数页页眉：Times New Roman，五号，居中内容为本章序号、本章名称段前、段后均为0行",
]
# 校准负样本（真文：提及个别格式词不命中）
NEG = [
    "本文提出一种基于深度强化学习的移动边缘计算卸载方法，实验表明该方法"
    "在200回合内收敛，平均系统奖励比基线提高12.6%。",
    "本文的排版遵循学校要求：正文采用宋体，但主要内容是翻译策略的比较分析"
    "与语料库研究方法的探讨，包括comparative translation quality assessment"
    "的多个维度与评价标准。",
    "参考文献是学术写作的重要组成部分，作者应当按照规范格式著录文献。",
]


def test_predicate_positive_shapes():
    from earlyz.ir.analyzer import is_format_instruction_text
    for x in POS:
        assert is_format_instruction_text(x), x[:40]


def test_predicate_negative_real_body():
    from earlyz.ir.analyzer import is_format_instruction_text
    for x in NEG:
        assert not is_format_instruction_text(x), x[:40]


def _mk(path: Path, paras: list[str]) -> Path:
    body = "".join(f"<w:p><w:r><w:t>{x}</w:t></w:r></w:p>" for x in paras)
    doc = (f'<?xml version="1.0"?><w:document {W}><w:body>{body}'
           f"<w:sectPr/></w:body></w:document>")
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", doc)
    return path


def test_strip_fourth_gate(tmp_path):
    from earlyz.actions.new_actions import strip_instruction_paragraphs
    fx = _mk(tmp_path / "a.docx", POS[:2] + NEG[:1])
    r = strip_instruction_paragraphs(fx, {})
    assert r["changed"] and r["removed"] == 2, r
    x = zipfile.ZipFile(fx).read("word/document.xml").decode("utf-8")
    assert "黑体；字号" not in x
    assert "深度强化学习" in x          # 真文保留


def test_strip_idempotent_and_fuse(tmp_path):
    from earlyz.actions.new_actions import strip_instruction_paragraphs
    fx = _mk(tmp_path / "a.docx", POS)
    r1 = strip_instruction_paragraphs(fx, {})
    r2 = strip_instruction_paragraphs(fx, {})
    assert r1["changed"] and r2["removed"] == 0
    # 熔断分层（EZI-0071）：包裹形态 >20 → 疑似模板本体全弃
    wrapped = _mk(tmp_path / "b.docx", [f"【模板说明第{i}条占位内容】" for i in range(25)])
    rw = strip_instruction_paragraphs(wrapped, {})
    assert not rw["changed"] and rw.get("aborted"), rw
    # 格式说明段独立上限 60：25 条真源粘贴不该熔断（3-2-1 实锤 22 条）
    spec25 = _mk(tmp_path / "c.docx", [POS[0] + f"补充说明第{i}项" for i in range(25)])
    rs = strip_instruction_paragraphs(spec25, {})
    assert rs["changed"] and rs["removed"] == 25, rs
    # >60 条说明 → 仍疑似模板本体全弃
    spec61 = _mk(tmp_path / "d.docx", [POS[0] + f"补充说明第{i}项" for i in range(61)])
    rb = strip_instruction_paragraphs(spec61, {})
    assert not rb["changed"] and rb.get("aborted"), rb


def test_hash_exemption_same_predicate():
    from earlyz.ir.analyzer import is_format_instruction_text
    import earlyz.ir.analyzer as an
    import inspect
    src = inspect.getsource(an._normalized_content_hash)
    assert "is_format_instruction_text" in src, "内容 hash 豁免须与谓词同源"


def test_judge_body_pool_excludes_instructions(tmp_path):
    """正文池：2 说明段被排除，真文保留（池计数断言——props 按取样协议不含
    text 键，保留性由「真文段入池、说明段出池」的计数差证明）。"""
    from earlyz.validation.layout_fidelity import (
        _load_docx, document_paragraphs, body_paragraphs)
    tpl = _mk(tmp_path / "t.docx", POS[:2] + NEG[:1])
    doc = _load_docx(tpl)
    pool = body_paragraphs(document_paragraphs(doc), exclude_form_labels=False)
    assert len(pool) == 1, f"说明段必须出池、真文必须留池：{len(pool)}"
    # 对照：没有说明段时 3 段全入池
    ctl = _mk(tmp_path / "c.docx", NEG[:3])
    pool2 = body_paragraphs(document_paragraphs(_load_docx(ctl)),
                            exclude_form_labels=False)
    assert len(pool2) == 3
