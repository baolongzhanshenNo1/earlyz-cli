"""早睡 (EarlyZ) 3.2 — 格式级联解析器（§4.7，§8.6 explain_format 核心）。

解释"某个段落的某个格式属性最终为什么是现在的值"：

    直接格式(direct pPr/rPr)
      → 段落样式(pStyle) → basedOn 样式链（循环保护）
        → docDefaults
          → theme（字体回退）/ docGrid（行距网格上下文）

纯只读分析（R0），不改文档。复用 ir/analyzer 的属性解析器
（_parse_ppr_format / _parse_styles / _parse_doc_defaults），保证
与 IR 分析同源、口径一致。
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from earlyz.ir.models import FormatProperties

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# 用户可读属性名 → FormatProperties 字段
_PROPERTY_ALIASES: dict[str, list[str]] = {
    "font": ["font_east_asia", "font_ascii"],
    "font_east_asia": ["font_east_asia"],
    "font_ascii": ["font_ascii"],
    "size": ["font_size_pt"],
    "font_size": ["font_size_pt"],
    "line_spacing": ["line", "line_rule", "line_spacing"],
    "bold": ["bold"],
    "italic": ["italic"],
    "underline": ["underline"],
    "color": ["color"],
    "jc": ["jc"],
    "alignment": ["jc"],
    "indent": ["ind_first_line_chars", "ind_first_line", "ind_left", "ind_right"],
    "first_line_indent": ["ind_first_line_chars", "ind_first_line"],
    "spacing": ["spacing_before", "spacing_after"],
    "keep_next": ["keep_next"],
    "keep_lines": ["keep_lines"],
    "page_break_before": ["page_break_before"],
    "outline_lvl": ["outline_lvl"],
}

_FIELD_LABELS = {
    "font_east_asia": "中文字体",
    "font_ascii": "西文字体",
    "font_size_pt": "字号(pt)",
    "line": "行距(240=单倍)",
    "line_rule": "行距规则",
    "line_spacing": "行距倍数",
    "bold": "加粗",
    "italic": "斜体",
    "underline": "下划线",
    "color": "颜色",
    "jc": "对齐",
    "ind_first_line_chars": "首行缩进(字符×100)",
    "ind_first_line": "首行缩进(twips)",
    "ind_left": "左缩进(twips)",
    "ind_right": "右缩进(twips)",
    "spacing_before": "段前(twips)",
    "spacing_after": "段后(twips)",
    "keep_next": "与下段同页",
    "keep_lines": "段中不分页",
    "page_break_before": "段前分页",
    "outline_lvl": "大纲级别",
}


@dataclass
class CascadeLevel:
    level: str                     # direct/style/style_chain/doc_defaults/doc_grid
    source: str                    # 部件/样式来源描述
    style_id: Optional[str] = None
    style_name: Optional[str] = None
    values: dict = field(default_factory=dict)   # 字段 → 值
    wins: bool = False
    note: str = ""


def _resolve_fields(properties: list[str] | None) -> list[str]:
    if not properties:
        return list(_FIELD_LABELS.keys())
    out: list[str] = []
    for p in properties:
        alias = _PROPERTY_ALIASES.get(p)
        if alias is None:
            raise ValueError(f"未知格式属性: {p}（可选: {sorted(_PROPERTY_ALIASES)}）")
        for f in alias:
            if f not in out:
                out.append(f)
    return out


def _parse_pstyle_id(document_xml: str, p_index: int) -> Optional[str]:
    """第 p_index（1 起）个正文段落的 pStyle 样式 ID。"""
    from lxml import etree

    root = etree.fromstring(document_xml.encode("utf-8"))
    paragraphs = root.findall(f".//{W}p")
    if p_index < 1 or p_index > len(paragraphs):
        return None
    p = paragraphs[p_index - 1]
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        return None
    pstyle = ppr.find(f"{W}pStyle")
    if pstyle is None:
        return None
    return pstyle.get(f"{W}val")


def _direct_props(document_xml: str, p_index: int, fields: list[str]) -> FormatProperties:
    """段落直接格式（pPr + 首 run 的 rPr）。"""
    from earlyz.ir.analyzer import _parse_ppr_format, _parse_rpr_format
    from lxml import etree

    root = etree.fromstring(document_xml.encode("utf-8"))
    paragraphs = root.findall(f".//{W}p")
    if p_index < 1 or p_index > len(paragraphs):
        raise ValueError(f"段落不存在: 第 {p_index} 段（全文共 {len(paragraphs)} 段）")
    p = paragraphs[p_index - 1]
    ppr = p.find(f"{W}pPr")
    props = _parse_ppr_format(ppr) if ppr is not None else FormatProperties()
    r = p.find(f"{W}r")
    if r is not None:
        rpr = r.find(f"{W}rPr")
        if rpr is not None:
            rp = _parse_rpr_format(rpr)
            for f in fields:
                v = getattr(rp, f)
                if v is not None:
                    setattr(props, f, v)
    return props


def _style_chain_map(styles_xml: str) -> dict[str, dict]:
    """style_id → {name, based_on, props}。"""
    from earlyz.ir.analyzer import _parse_ppr_format, _parse_rpr_format
    from lxml import etree

    root = etree.fromstring(styles_xml.encode("utf-8"))
    out: dict[str, dict] = {}
    for st in root.findall(f"{W}style"):
        sid = st.get(f"{W}styleId")
        if not sid:
            continue
        name_el = st.find(f"{W}name")
        based = None
        b = st.find(f"{W}basedOn")
        if b is not None:
            based = b.get(f"{W}val")
        props = FormatProperties()
        ppr = st.find(f"{W}pPr")
        if ppr is not None:
            props = _parse_ppr_format(ppr)
        rpr = st.find(f"{W}rPr")
        if rpr is not None:
            rp = _parse_rpr_format(rpr)
            for f in rp.__dict__:
                v = getattr(rp, f)
                if v is not None:
                    setattr(props, f, v)
        out[sid] = {
            "name": name_el.get(f"{W}val") if name_el is not None else None,
            "based_on": based,
            "props": props,
        }
    return out


def explain_paragraph_format(
    docx_path: str | Path,
    paragraph: int | str = 1,
    properties: Optional[list[str]] = None,
) -> dict:
    """解释第 N 段（或匹配文本的段落）格式属性的完整生效链（§8.6）。

    Returns:
        {
          "node_id": "paragraph:1@word/document.xml",
          "locator": 1,
          "effective": {field: value},
          "chains": {field: [{level, source, style_id, style_name, value, wins, note}]},
          "winner": {field: level},
        }
    """
    src = Path(docx_path)
    if not src.exists():
        raise FileNotFoundError(f"文档不存在: {src}")
    fields = _resolve_fields(properties)

    from earlyz.ir.analyzer import _parse_doc_defaults
    from lxml import etree

    with zipfile.ZipFile(str(src)) as z:
        document_xml = z.read("word/document.xml").decode("utf-8", errors="replace")
        styles_xml = z.read("word/styles.xml").decode("utf-8", errors="replace")

    # ── 段落定位 ──
    root = etree.fromstring(document_xml.encode("utf-8"))
    paragraphs = root.findall(f".//{W}p")
    p_index: Optional[int] = None
    if isinstance(paragraph, int):
        p_index = paragraph
        if p_index < 1 or p_index > len(paragraphs):
            raise ValueError(f"段落不存在: 第 {p_index} 段（全文共 {len(paragraphs)} 段）")
    else:
        needle = str(paragraph)
        for i, p in enumerate(paragraphs, start=1):
            text = "".join(t.text or "" for t in p.findall(f".//{W}t"))
            if needle in text:
                p_index = i
                break
        if p_index is None:
            raise ValueError(f"找不到包含文本「{needle}」的段落")

    # ── 级联各层 ──
    levels: list[CascadeLevel] = []
    style_map = _style_chain_map(styles_xml)
    defaults = _parse_doc_defaults(styles_xml)

    # 层 1：直接格式
    direct = _direct_props(document_xml, p_index, fields)
    levels.append(CascadeLevel(
        level="direct", source="word/document.xml pPr/rPr",
        values={f: getattr(direct, f) for f in fields},
    ))

    # 层 2+：样式链（pStyle → basedOn → …，循环保护 ≤8 层）
    style_id = _parse_pstyle_id(document_xml, p_index)
    seen: set[str] = set()
    cur = style_id
    first = True
    while cur and cur in style_map and cur not in seen and len(seen) < 8:
        seen.add(cur)
        info = style_map[cur]
        levels.append(CascadeLevel(
            level="style" if first else "style_chain",
            source="word/styles.xml",
            style_id=cur,
            style_name=info["name"],
            values={f: getattr(info["props"], f) for f in fields},
            note="段落样式" if first else f"basedOn 继承自 {info.get('based_on') or '—'}",
        ))
        first = False
        cur = info.get("based_on")
    if cur and cur not in style_map and cur not in seen:
        levels.append(CascadeLevel(
            level="style_chain",
            source="word/styles.xml",
            style_id=cur,
            style_name=None,
            values={},
            note=f"basedOn 指向不存在的样式 {cur}（样式定义缺失）",
        ))

    # 层 3：docDefaults（rPrDefault 来自 analyzer 助手；pPrDefault 行距/对齐补齐）
    _dd = dict(defaults)
    try:
        styles_root = etree.fromstring(styles_xml.encode("utf-8"))
        dd = styles_root.find(f"{W}docDefaults")
        if dd is not None:
            ppr_def = dd.find(f"{W}pPrDefault")
            if ppr_def is not None:
                ppr = ppr_def.find(f"{W}pPr")
                if ppr is not None:
                    from earlyz.ir.analyzer import _parse_ppr_format

                    for k, v in _parse_ppr_format(ppr).__dict__.items():
                        if v is not None:
                            _dd.setdefault(k, v)
    except Exception:
        pass
    levels.append(CascadeLevel(
        level="doc_defaults", source="word/styles.xml docDefaults",
        values={f: _dd.get(f) for f in fields},
    ))

    # 层 4：docGrid 行距网格上下文（影响 spacing 生效，仅说明）
    dg = root.find(f".//{W}docGrid")
    if dg is not None:
        line_pitch = dg.get(f"{W}linePitch")
        levels.append(CascadeLevel(
            level="doc_grid", source="word/document.xml sectPr docGrid",
            values={},
            note=f"行距网格 linePitch={line_pitch}（若段落未关 snapToGrid，小行距会被网格取整）",
        ))

    # ── 合成：每字段的生效链 + 胜者 ──
    chains: dict[str, list[dict]] = {}
    winners: dict[str, str] = {}
    for f in fields:
        chain: list[dict] = []
        winner = None
        for lv in levels:
            v = lv.values.get(f)
            if v is None:
                continue
            is_win = winner is None
            chain.append({
                "level": lv.level,
                "source": lv.source,
                "style_id": lv.style_id,
                "style_name": lv.style_name,
                "value": v,
                "wins": is_win,
                "note": (lv.note if lv.note else None) if is_win else
                        (f"被上层「{winner}」覆盖" if winner else lv.note),
            })
            if is_win:
                winner = lv.level
        chains[f] = chain
        winners[f] = winner or "none"

    effective = {
        f: next(
            (c["value"] for c in chains.get(f, []) if c["wins"]),
            None,
        )
        for f in fields
    }

    # 网格提示附加到行距相关字段的链尾
    grid_note = next((lv.note for lv in levels if lv.level == "doc_grid" and lv.note), None)
    if grid_note:
        for f in ("line", "line_rule", "line_spacing", "spacing_before", "spacing_after"):
            if f in chains and chains[f]:
                chains[f][-1]["grid_note"] = grid_note

    return {
        "node_id": f"paragraph:{p_index}@word/document.xml",
        "locator": paragraph,
        "p_index": p_index,
        "p_style_id": style_id,
        "properties": fields,
        "property_labels": {f: _FIELD_LABELS.get(f, f) for f in fields},
        "effective": effective,
        "winner": winners,
        "chains": chains,
    }
