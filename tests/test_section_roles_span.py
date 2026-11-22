# -*- coding: utf-8 -*-
"""EZI-0075 toc 短节限定/长节透明化 + 附件冒号形词表（引擎+判定双端同源）。

真机实锤（建设施工 2017 版，模板 118 节）：
- 产物节1 = 1628 段合同正文大节，内含一行「目	录」→ 被 toc 独占规则吞成
  纯 toc 节 → align 配对模板唯一目录节（1680）→ 沿用再灌给附件节；
- 「附件 3 ：」冒号形独立标题行，引擎/判定词表都不认 → 模板表单区真值
  （1680）无角色可配对。
修法：toc 独占限短节（≤80 段），长节目录透明；附件词表补冒号形（短行、
无制表符、可选尾冒号——目录条目「附件2⇥45」仍被页码/tab 双闸排除）。
"""

from __future__ import annotations

from earlyz.actions.section_geometry import (
    _para_role, _detect_roles, _region_roles, _TOC_PAGE_MAX_PARAS)
from earlyz.validation.layout_fidelity import _section_roles


# ---------- 引擎侧：附件词表 ----------

def test_engine_para_role_colon_attachment():
    assert _para_role("附件 3 ：") == "attachment"
    assert _para_role("附件 10 ：") == "attachment"
    assert _para_role("附件2") == "attachment"          # 原裸形不回归
    assert _para_role("附件") == "attachment"            # 无编号裸形


def test_engine_para_role_attachment_false_positives():
    """目录条目（tab+页码）、带标题的合并行、正文句都不算附件。"""
    assert _para_role("附件 2" + chr(9) + "45") is None           # tab 条目
    assert _para_role("附件 2 ：发包人供应材料设备一览表附件 3 ：工程质量保修书") is None
    assert _para_role("附件：详见合同专用条款约定执行") is None    # 冒号后有正文
    assert _para_role("如附件所述，双方应按照法律法规履行各自义务") is None


# ---------- 引擎侧：toc 长节透明 ----------

def test_engine_detect_roles_long_span_toc_transparent():
    """1628 段大节含一行「目录」→ 不是 toc 节（目录透明、正文身份保留）。"""
    texts = ["目" + chr(9) + "录"] + ["第%d章 示例条款内容" % i for i in range(1, 200)]
    roles = _detect_roles(texts, first=False)
    assert "toc" not in roles
    assert "body" in roles


def test_engine_detect_roles_short_toc_page_exclusive():
    """真目录页（短节）独占语义不变——防参考文献/致谢条目污染。"""
    texts = ["目" + chr(9) + "录", "参考文献45", "致谢59", "附录62"]
    roles = _detect_roles(texts, first=False)
    assert roles == {"toc"}


def test_engine_region_roles_same_semantics():
    """模板侧同口径：长节目录透明；短目录节独占。"""
    long_texts = ["目" + chr(9) + "录"] + ["第%d章 正文内容示例" % i for i in range(1, 120)]
    assert "toc" not in _region_roles(long_texts)
    assert "toc" in _region_roles(["目" + chr(9) + "录", "参考文献45", "致谢59"])


def test_engine_region_roles_attachment():
    assert "attachment" in _region_roles(["附件 3 ：", "工程质量保修书", "发包人（全称）："])
    assert "attachment" in _region_roles(["附件2"])


# ---------- 判定侧：同源 ----------

def test_judge_section_roles_long_span_toc_transparent():
    texts = ["目" + chr(9) + "录"] + ["第%d章 示例条款内容" % i for i in range(1, 200)]
    roles = _section_roles(texts, first=False)
    assert "toc" not in roles and "body" in roles


def test_judge_section_roles_short_toc_exclusive():
    texts = ["目" + chr(9) + "录", "参考文献45", "致谢59", "附录62"]
    assert _section_roles(texts, first=False) == ("toc",)


def test_judge_section_roles_attachment_colon():
    roles = _section_roles(["附件 3 ：", "工程质量保修书", "发包人（全称）：甲公司"], first=False)
    assert "attachment" in roles
    # tab 条目不算
    assert "attachment" not in _section_roles(["附件 2" + chr(9) + "45"], first=False)
    # 合并长行不算
    assert "attachment" not in _section_roles(
        ["附件 2 ：发包人供应材料设备一览表附件 3 ：工程质量保修书"], first=False)


def test_threshold_constant_shared():
    """两端节长上限同值（80）——词表/口径同源的机械钉。"""
    assert _TOC_PAGE_MAX_PARAS == 80
    from earlyz.validation.layout_fidelity import _TOC_SECTION_MAX_PARAS
    assert _TOC_SECTION_MAX_PARAS == 80


# ---------- 端到端：配对语义 ----------

def test_pair_margin_attachment_role():
    """模板表单节（attachment, 1680）与产物附件节按角色配对成功。"""
    from earlyz.validation.layout_fidelity import _pair_margin_sections
    tpl_roles = [("cover",), ("toc",), ("attachment",), ("body",)]
    tpl_margins = [
        {"top": 1580, "right": 1300, "bottom": 1020, "left": 1300},
        {"top": 1580, "right": 1680, "bottom": 1020, "left": 1300},
        {"top": 1580, "right": 1680, "bottom": 1020, "left": 1300},   # 表单节
        {"top": 1580, "right": 1200, "bottom": 1020, "left": 1300},
    ]
    prod_roles = [("cover",), ("body",), ("attachment",)]
    pairs, mode, _ = _pair_margin_sections(tpl_roles, prod_roles, tpl_margins=tpl_margins)
    by_prod = dict(pairs)
    assert mode == "role"
    assert by_prod[2] == 2, by_prod          # 产物附件节 → 模板表单节
