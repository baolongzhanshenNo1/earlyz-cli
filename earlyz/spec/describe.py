"""§19.5 P1：规则清单人话描述 + provenance 溯源（模板详情页数据源）。

把 FormatSpec.rules 渲染为面向用户的中文清单（UI 文案铁律：禁开发术语，
"配置/渲染/素材"等禁用，用「这段文字……」句式）——详情页/修正入口展示。

规则来源溯源：providence.source（user_confirmed/written_standard/
template_ooxml/sample_paragraph/inferred）→ 中文标签。
"""

from __future__ import annotations

from typing import Any

_ROLE_CN = {
    "title": "题目",
    "chapter_title": "章标题",
    "section_title": "节标题",
    "body": "正文",
    "keywords": "关键词",
    "abstract_cn": "中文摘要",
    "abstract_en": "英文摘要",
    "caption": "图表题注",
    "reference": "参考文献",
    "appendix": "附录",
    "table_cell": "表格内容",
    "header": "页眉",
    "footer": "页脚",
}

_JC_CN = {"center": "居中", "left": "左对齐", "right": "右对齐", "both": "两端对齐"}

_PROP_RENDER = {
    "font_size_pt": lambda v: f"字号 {float(v):g} 磅",
    "font_east_asia": lambda v: f"中文字体 {v}",
    "font_ascii": lambda v: f"西文字体 {v}",
    "jc": lambda v: _JC_CN.get(str(v), f"对齐 {v}"),
    "first_line_chars": lambda v: f"首行缩进 {float(v):g} 字符",
    "first_line_twips": lambda v: f"首行缩进 {int(v)//20} 字符",
    "left_twips": lambda v: f"左缩进 {int(v)//20} 字符",
    "line_twips": lambda v: f"行距 {v} twips",
    "before_twips": lambda v: f"段前 {int(v)//20} 磅",
    "after_twips": lambda v: f"段后 {int(v)//20} 磅",
    "bold": lambda v: "加粗" if v else "不加粗",
    "italic": lambda v: "斜体" if v else "不斜体",
    "outline_lvl": lambda v: f"大纲层级 {v}",
    "p_style": lambda v: f"段落样式 {v}",
    "vert_align": lambda v: f"上下标 {v}",
    "line_rule": lambda v: f"行距规则 {v}",
}

_SOURCE_CN = {
    "user_confirmed": "用户确认",
    "written_standard": "书面标准",
    "template_ooxml": "模板导出",
    "sample_paragraph": "示例段落",
    "inferred": "推断得出",
}


def describe_expected(expected: dict[str, Any]) -> str:
    """expected（FormatProperties）→ 人话短语列表（按模板顺序）。"""
    parts: list[str] = []
    order = ["font_east_asia", "font_ascii", "font_size_pt", "bold", "italic",
             "first_line_chars", "first_line_twips", "left_twips", "jc",
             "line_twips", "before_twips", "after_twips", "outline_lvl",
             "p_style", "vert_align", "line_rule"]
    for key in order:
        v = expected.get(key)
        if v is None:
            continue
        render = _PROP_RENDER.get(key)
        if render is None:
            parts.append(f"{key} {v}")
            continue
        try:
            parts.append(render(v))
        except (TypeError, ValueError):
            parts.append(f"{key} {v}")
    # 未知键兜底（顺序之外但存在 → 保留显示，不静默丢属性）
    for key, v in expected.items():
        if key not in order and v is not None:
            parts.append(f"{key} {v}")
    return "、".join(parts)


def _role_human(rule) -> str:
    role = (rule.selector or {}).get("semantic_role")
    if role:
        title_key = (rule.selector or {}).get("title_key")
        base = _ROLE_CN.get(role, role)
        return f"{base}（{title_key}）" if title_key else base
    return "整篇"


def describe_spec_rules(spec) -> list[dict]:
    """规则清单（人话 + expected 原文 + provenance 溯源）。

    Each item: {rule_id, human_role, human, expected, provenance_source}
    """
    out = []
    for rule in (spec.rules or []):
        human = describe_expected(rule.expected or {})
        src = ""
        if rule.provenance is not None:
            src = _SOURCE_CN.get(rule.provenance.source, rule.provenance.source)
        out.append({
            "rule_id": rule.rule_id,
            "human_role": _role_human(rule),
            "human": human or "（无格式属性）",
            "expected": rule.expected,
            "provenance_source": src,
        })
    return out


def describe_pack_rules(pack_id: str) -> dict:
    """模板详情页数据（§19.5 P1）：规则清单（人话+溯源）+ 用户层覆盖数。

    pack_id：system pack（thesis/gov/…）或 user:<sha>（用户上传模板）。
    编译失败/未知包 → {"error": ...}（不抛）。
    """
    from earlyz.spec.compilers import compile_format_spec
    from earlyz.spec.packs import find_user_pack_file

    try:
        if pack_id.startswith("user:"):
            user_file = find_user_pack_file(pack_id)
            if user_file is None:
                return {"error": f"用户模板不存在: {pack_id}"}
            comp = compile_format_spec("uploaded_template_docx", str(user_file), "thesis")
        else:
            # 内置规范包已下架（2026-09-21 用户拍板）：详情页一律返回下架提示
            return {"error": "内置模板已下架，请上传自己的模板文件（或选择「我的模板」）查看规则。"}
    except Exception as e:  # noqa: BLE001
        return {"error": f"{type(e).__name__}: {e}"}
    spec = comp.spec
    from earlyz.suggest.proposals import load_overrides
    overrides = load_overrides(spec.spec_id)
    return {
        "pack_id": pack_id,
        "spec_id": spec.spec_id,
        "rules": describe_spec_rules(spec),
        "user_overrides_count": len(overrides),
    }
