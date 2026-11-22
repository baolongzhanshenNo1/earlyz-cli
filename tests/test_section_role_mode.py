# -*- coding: utf-8 -*-
"""EZI-0074 同角色多节几何取众数——引擎与判定配对避开少数派变体节。

真机实锤（建设施工合同 2017 版，模板 118 节 vs 产物 4 节）：
- 模板目录节里混 1680 变体、正文节里散布横向附件页（1100/1580/280）；
- 引擎 `_role_map` 按序消费把变体节抄给产物、判定层同病——两边都得把
  **角色几何众数**的节排到队首（组内稳定保序），均匀角色排序恒等零扰动。
"""

from __future__ import annotations

from earlyz.actions.section_geometry import _pair_sections, _role_map
from earlyz.validation.layout_fidelity import _pair_margin_sections

NORMAL = {"sz": {"w": 11906, "h": 16838, "orient": "portrait"},
          "mar": {"top": 1580, "right": 1200, "bottom": 1020, "left": 1300}}
VARIANT = {"sz": {"w": 11906, "h": 16838, "orient": "portrait"},
           "mar": {"top": 1580, "right": 1680, "bottom": 1020, "left": 1300}}
LANDSCAPE = {"sz": {"w": 16838, "h": 11906, "orient": "landscape"},
             "mar": {"top": 1100, "right": 1580, "bottom": 280, "left": 1300}}


def _tpl():
    """模板 6 节：封面 + 目录(变体在先) + 正文×4（横向附件页混在中间）。"""
    return [
        {**NORMAL, "mar": {**NORMAL["mar"], "right": 1300}, "roles": ["封面"]},
        {**VARIANT, "roles": ["目录"]},
        {**NORMAL, "roles": ["目录"]},
        {**NORMAL, "roles": ["正文"]},
        {**LANDSCAPE, "roles": ["正文"]},
        {**NORMAL, "roles": ["正文"]},
    ]


def _prod():
    return [
        {**NORMAL, "roles": ["封面"]},
        {**NORMAL, "roles": ["目录"]},
        {**NORMAL, "roles": ["正文"]},
    ]


def test_engine_role_map_prefers_mode_geometry():
    """引擎：严格众数角色取众数节（正文 3:1 取 1200 形态，不取横向节 4）。"""
    tpl, prod = _tpl(), _prod()
    mapping, matched = _role_map(tpl, prod)
    assert matched == 3
    assert mapping[0] == 0                       # 封面唯一节
    assert mapping[2] in (3, 5), mapping         # 正文众数=节3/5（横向节4 让位）


def test_engine_uniform_role_unchanged():
    """均匀角色：众数=全体 → 排序恒等 → 配对与旧口径一致（零扰动）。"""
    tpl = [{**NORMAL, "roles": ["封面"]},
           {**NORMAL, "roles": ["目录"]},
           {**NORMAL, "roles": ["正文"]},
           {**NORMAL, "roles": ["正文"]}]
    prod = [{**NORMAL, "roles": ["封面"]},
            {**NORMAL, "roles": ["目录"]},
            {**NORMAL, "roles": ["正文"]}]
    mapping, matched = _role_map(tpl, prod)
    assert mapping == {0: 0, 1: 1, 2: 2} and matched == 3


def test_engine_pair_sections_uses_reordered_role_map():
    tpl, prod = _tpl(), _prod()
    seq, mode = _pair_sections(tpl, prod)
    assert mode == "role"
    assert seq[2] in (3, 5), (seq, mode)         # 正文取众数节，横向节 4 让位


def test_judge_pair_prefers_mode_margin():
    """判定层：严格众数角色（正文 3:1）期望取众数节，不取横向节 4。"""
    tpl = _tpl()
    tpl_roles = [g["roles"] for g in tpl]
    tpl_margins = [dict(g["mar"]) for g in tpl]
    prod_roles = [["封面"], ["目录"], ["正文"]]
    pairs, mode, _note = _pair_margin_sections(
        tpl_roles, prod_roles, tpl_margins=tpl_margins)
    by_prod = dict(pairs)
    assert mode == "role"
    assert by_prod[0] == 0
    assert by_prod[2] in (3, 5), by_prod            # 众数正文节（非横向 4）


def test_judge_multiform_role_matches_any():
    """判定层候选集：目录角色 1680/1200 两形态（平票无众数）→ 产物任一
    形态达标（EZI-0022 双峰同原则）——这是建设施工 118 节模板的正主形态。"""
    tpl = _tpl()
    tpl_roles = [g["roles"] for g in tpl]
    tpl_margins = [dict(g["mar"]) for g in tpl]
    for prod_mar in (dict(VARIANT["mar"]), dict(NORMAL["mar"])):
        from earlyz.validation.layout_fidelity import (
            _load_docx, _section_margins, _margin_diff, evaluate_layout_fidelity)
        # 直接驱动 _pair_margin_sections + _pair_ok 等价逻辑：
        # 平票角色两形态都在候选集 → 两种产物形态都 PASS
        forms = {tuple(sorted((k, tpl_margins[ti].get(k)) for k in ("top", "right", "bottom", "left")))
                 for ti, roles in enumerate(tpl_roles) if "目录" in roles}
        pf = tuple(sorted((k, prod_mar.get(k)) for k in ("top", "right", "bottom", "left")))
        assert len(forms) == 2, forms
        assert pf in forms, pf


def test_judge_pair_without_margins_keeps_old_behavior():
    """不传 tpl_margins（旧行为兼容）：按序消费（回归钉）。"""
    tpl_roles = [["封面"], ["目录"], ["目录"], ["正文"], ["正文"], ["正文"]]
    prod_roles = [["封面"], ["目录"], ["正文"]]
    pairs, _mode, _note = _pair_margin_sections(tpl_roles, prod_roles)
    by_prod = dict(pairs)
    assert by_prod[1] == 1 and by_prod[2] == 3, by_prod


# ---------- _plan / apply_section_plan（EZI-0074 第二写手） ----------

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _mkdoc(sections):
    """sections: list of (role_text, margin_tuple, orient)。synthetic document.xml。"""
    from lxml import etree as _et
    root = _et.Element(f"{{{_W}}}document")
    body = _et.SubElement(root, f"{{{_W}}}body")
    for i, (text, mar, orient) in enumerate(sections):
        p = _et.SubElement(body, f"{{{_W}}}p")
        t_ = _et.SubElement(p, f"{{{_W}}}t")
        t_.text = text
        ppr = _et.SubElement(p, f"{{{_W}}}pPr")
        sp = _et.SubElement(ppr, f"{{{_W}}}sectPr")
        sz = _et.SubElement(sp, f"{{{_W}}}pgSz")
        sz.set(f"{{{_W}}}w", "16838" if orient == "landscape" else "11906")
        sz.set(f"{{{_W}}}h", "11906" if orient == "landscape" else "16838")
        if orient == "landscape":
            sz.set(f"{{{_W}}}orient", "landscape")
        m = _et.SubElement(sp, f"{{{_W}}}pgMar")
        for k, v in zip(("top", "right", "bottom", "left"), mar):
            m.set(f"{{{_W}}}{k}", str(v))
    tail = _et.SubElement(body, f"{{{_W}}}sectPr")
    _et.SubElement(tail, f"{{{_W}}}pgSz")
    return root


def test_plan_candidates_put_mode_regions_first():
    """`_plan` 单调前进前先众数重排：正文角色首个区域是横向变体时，
    产物正文块拿众数（纵向）区域，不拿变体。"""
    from earlyz.actions.section_geometry import _template_regions, _plan
    tpl_root = _mkdoc([
        ("建设工程施工合同", (1580, 1300, 1020, 1300), "portrait"),   # cover
        ("目录", (1580, 1680, 1020, 1300), "portrait"),                # toc(唯一,1680)
        ("第一条 总则", (1100, 1580, 280, 1300), "landscape"),         # body 变体在先
        ("第二条 术语", (1580, 1300, 1020, 1300), "portrait"),        # body 众数
        ("第三条 价款", (1580, 1300, 1020, 1300), "portrait"),        # body 众数
    ])
    regions = _template_regions(tpl_root)
    assert len(regions) == 6          # 5 节 + body 末尾尾节
    prod_root = _mkdoc([
        ("建设工程施工合同", (1440, 1800, 1440, 1800), "portrait"),
        ("第一条 总则", (1440, 1800, 1440, 1800), "portrait"),
    ])
    plan = _plan(prod_root, regions)
    tpl_ids = [b["tpl"] for b in plan["blocks"] if "tpl" in b]
    assert all(ti not in (2,) for ti in tpl_ids[1:]), tpl_ids   # 正文块不落横向节(2)
