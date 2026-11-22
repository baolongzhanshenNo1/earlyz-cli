"""早睡 (EarlyZ) 3.2 — Phase 2 样式系统 action（§7.1 P0 清单）。

create_or_update_style      — 创建或更新样式（幂等：值相同无变化）
clone_style_from_spec        — 从规范模板 docx 克隆样式（含深拷贝，失败可回滚）
clear_direct_formatting      — 清除段落/run 直接格式（保留 pStyle/numPr/sectPr）
map_style                    — 全文样式映射（from_style_id → to_style_id）

全部：确定性、幂等、带 verifier（无 verifier 的 mutation 不注册，ADR-008）。
复用 new_actions 的 zip 读写基础设施。
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional
from lxml import etree
import zipfile

from earlyz.actions.new_actions import (
    _read_part,
    _write_part,
    _text_of,
    _locate_paragraphs,
    _read_styles_map,
    _role_of,
)

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# pPr 中保留的结构性元素（直接格式清除时不动）
_PPR_KEEP = {f"{W}pStyle", f"{W}numPr", f"{W}sectPr", f"{W}keepNext",
             f"{W}keepLines", f"{W}widowControl"}


def _find_style(root: etree._Element, style_id: str) -> Optional[etree._Element]:
    for style in root.findall(f"{W}style"):
        if style.get(f"{W}styleId") == style_id:
            return style
    return None


# ────────────────────────────────────────────────────────────────
# create_or_update_style
# ────────────────────────────────────────────────────────────────

# FormatProperties 字段名 → (容器元素 tag, 属性名) 映射（样式上下文）
_STYLE_RPR_MAP = {
    "font_east_asia": (f"{W}rFonts", "eastAsia"),
    "font_ascii": (f"{W}rFonts", "ascii"),
    "font_h_ansi": (f"{W}rFonts", "hAnsi"),
    "font_size_pt": (f"{W}sz", "val"),  # 半磅
    "bold": (f"{W}b", None),
    "italic": (f"{W}i", None),
    "color": (f"{W}color", "val"),
}
_STYLE_PPR_MAP = {
    "jc": (f"{W}jc", "val"),
    "spacing_before": (f"{W}spacing", "before"),
    "spacing_after": (f"{W}spacing", "after"),
    "line": (f"{W}spacing", "line"),
    "line_rule": (f"{W}spacing", "lineRule"),
    "ind_first_line_chars": (f"{W}ind", "firstLineChars"),
    "ind_first_line": (f"{W}ind", "firstLine"),
    "ind_left": (f"{W}ind", "left"),
    "ind_right": (f"{W}ind", "right"),
}


def _apply_style_properties(rpr: etree._Element, ppr: etree._Element,
                            properties: dict, rpr_created: bool) -> None:
    for prop, value in properties.items():
        if value is None:
            continue
        if prop in _STYLE_RPR_MAP:
            tag, attr = _STYLE_RPR_MAP[prop]
            if attr is None:
                el = rpr.find(tag)
                if value is True and el is None:
                    etree.SubElement(rpr, tag)
                elif value is False and el is not None:
                    rpr.remove(el)
            else:
                el = rpr.find(tag)
                if el is None:
                    el = etree.SubElement(rpr, tag)
                if prop == "font_size_pt":
                    el.set(f"{W}{attr}", str(int(round(value * 2))))
                else:
                    el.set(f"{W}{attr}", str(value))
        elif prop in _STYLE_PPR_MAP:
            tag, attr = _STYLE_PPR_MAP[prop]
            el = ppr.find(tag)
            if el is None:
                el = etree.SubElement(ppr, tag)
            if isinstance(value, bool):
                el.set(f"{W}{attr}", "1" if value else "0")
            else:
                el.set(f"{W}{attr}", str(int(value)) if isinstance(value, (int, float)) else str(value))


def create_or_update_style(target_path: Path, params: dict) -> dict:
    """创建或更新 styles.xml 中的样式。

    params: {style_id, name?, type?, based_on?, properties: {...}}
    """
    target_path = Path(target_path)
    style_id = params["style_id"]
    properties = params.get("properties") or {}
    root = _read_part(target_path, "word/styles.xml")

    style = _find_style(root, style_id)
    created = style is None
    if style is None:
        style = etree.Element(f"{W}style", {f"{W}styleId": style_id})
        style.set(f"{W}type", params.get("type", "paragraph"))
        if params.get("name"):
            name_el = etree.SubElement(style, f"{W}name")
            name_el.set(f"{W}val", params["name"])
        root.append(style)
    if params.get("based_on"):
        based = style.find(f"{W}basedOn")
        if based is None:
            based = etree.SubElement(style, f"{W}basedOn")
        based.set(f"{W}val", params["based_on"])

    rpr = style.find(f"{W}rPr")
    ppr = style.find(f"{W}pPr")
    if properties:
        if rpr is None:
            rpr = etree.SubElement(style, f"{W}rPr")
        if ppr is None:
            ppr = etree.SubElement(style, f"{W}pPr")
        _apply_style_properties(rpr, ppr, properties, created)

    _write_part(target_path, "word/styles.xml", root)
    return {"changed": created, "created": created, "style_id": style_id}


def verify_style_created(target_path: Path, params: dict) -> bool:
    """验证样式存在且关键属性匹配。"""
    try:
        target_path = Path(target_path)
        root = _read_part(target_path, "word/styles.xml")
        style = _find_style(root, params["style_id"])
        if style is None:
            return False
        properties = params.get("properties") or {}
        if properties.get("font_east_asia"):
            rf = style.find(f"{W}rPr/{W}rFonts")
            if rf is None or rf.get(f"{W}eastAsia") != properties["font_east_asia"]:
                return False
        if properties.get("font_ascii"):
            rf = style.find(f"{W}rPr/{W}rFonts")
            if rf is None or rf.get(f"{W}ascii") != properties["font_ascii"]:
                return False
        if properties.get("font_size_pt") is not None:
            sz = style.find(f"{W}rPr/{W}sz")
            if sz is None or sz.get(f"{W}val") != str(int(round(properties["font_size_pt"] * 2))):
                return False
        return True
    except Exception:
        return False


# ────────────────────────────────────────────────────────────────
# clone_style_from_spec
# ────────────────────────────────────────────────────────────────

def clone_style_from_spec(target_path: Path, params: dict) -> dict:
    """从规范模板 docx 克隆样式到目标文档。

    params: {style_id, source_style_id, source_docx_path}
    深拷贝整个 <w:style> 元素（含 rPr/pPr/basedOn）。
    """
    target_path = Path(target_path)
    style_id = params["style_id"]
    source_style_id = params.get("source_style_id", style_id)
    source_path = Path(params["source_docx_path"])

    src_root = _read_part(source_path, "word/styles.xml")
    src_style = _find_style(src_root, source_style_id)
    if src_style is None:
        raise ValueError(f"源文档无样式 {source_style_id}")

    dst_root = _read_part(target_path, "word/styles.xml")
    existing = _find_style(dst_root, style_id)
    if existing is not None:
        dst_root.remove(existing)  # 替换（幂等更新语义）
    clone = etree.fromstring(etree.tostring(src_style))  # 深拷贝
    clone.set(f"{W}styleId", style_id)
    dst_root.append(clone)
    _write_part(target_path, "word/styles.xml", dst_root)
    return {"changed": True, "style_id": style_id, "source_style_id": source_style_id}


def verify_style_cloned(target_path: Path, params: dict) -> bool:
    """验证目标文档样式存在且与源样式 rPr 关键属性一致。"""
    try:
        target_path = Path(target_path)
        source_path = Path(params["source_docx_path"])
        style_id = params["style_id"]
        source_style_id = params.get("source_style_id", style_id)
        dst_root = _read_part(target_path, "word/styles.xml")
        dst_style = _find_style(dst_root, style_id)
        if dst_style is None:
            return False
        src_root = _read_part(source_path, "word/styles.xml")
        src_style = _find_style(src_root, source_style_id)
        if src_style is None:
            return False
        # 比较 rPr 序列化（关键属性一致性）
        src_rpr = src_style.find(f"{W}rPr")
        dst_rpr = dst_style.find(f"{W}rPr")
        if src_rpr is None and dst_rpr is None:
            return True
        if src_rpr is None or dst_rpr is None:
            return False
        return etree.tostring(src_rpr) == etree.tostring(dst_rpr)
    except Exception:
        return False


# ────────────────────────────────────────────────────────────────
# clear_direct_formatting
# ────────────────────────────────────────────────────────────────

def clear_direct_formatting(target_path: Path, params: dict) -> dict:
    """清除段落的直接格式（pPr 非结构性元素 + run rPr），保留样式引用。

    params: {selector: {semantic_role, occurrence, text_hash}} 或 {semantic_role}
    """
    target_path = Path(target_path)
    selector = params.get("selector") or {"semantic_role": params.get("semantic_role")}
    root = _read_part(target_path, "word/document.xml")
    paras = _locate_paragraphs(root, selector, _read_styles_map(target_path))
    if not paras:
        raise ValueError(f"no paragraph matched selector: {selector}")

    changed = 0
    for p in paras:
        ppr = p.find(f"{W}pPr")
        if ppr is not None:
            for child in list(ppr):
                if child.tag not in _PPR_KEEP:
                    ppr.remove(child)
            if len(ppr) == 0:
                p.remove(ppr)
            changed += 1
        # run 直接格式
        for r in p.findall(f"{W}r"):
            rpr = r.find(f"{W}rPr")
            if rpr is not None:
                r.remove(rpr)
                changed += 1
    _write_part(target_path, "word/document.xml", root)
    return {"changed": changed > 0, "paragraphs": len(paras)}


def verify_no_direct_formatting(target_path: Path, params: dict) -> bool:
    """验证目标段落无直接格式（pPr 仅结构性元素、run 无 rPr）。"""
    try:
        target_path = Path(target_path)
        selector = params.get("selector") or {"semantic_role": params.get("semantic_role")}
        root = _read_part(target_path, "word/document.xml")
        paras = _locate_paragraphs(root, selector, _read_styles_map(target_path))
        if not paras:
            return False
        for p in paras:
            ppr = p.find(f"{W}pPr")
            if ppr is not None:
                for child in ppr:
                    if child.tag not in _PPR_KEEP:
                        return False
            for r in p.findall(f"{W}r"):
                if r.find(f"{W}rPr") is not None:
                    return False
        return True
    except Exception:
        return False


# ────────────────────────────────────────────────────────────────
# map_style
# ────────────────────────────────────────────────────────────────

def map_style(target_path: Path, params: dict) -> dict:
    """全文样式映射：所有引用 from_style_id 的段落改为 to_style_id。

    params: {from_style_id, to_style_id}
    """
    target_path = Path(target_path)
    from_id = params["from_style_id"]
    to_id = params["to_style_id"]
    root = _read_part(target_path, "word/document.xml")
    changed = 0
    for p in root.iter(f"{W}p"):
        pstyle = p.find(f"{W}pPr/{W}pStyle")
        if pstyle is not None and pstyle.get(f"{W}val") == from_id:
            pstyle.set(f"{W}val", to_id)
            changed += 1
    _write_part(target_path, "word/document.xml", root)
    return {"changed": changed > 0, "mapped": changed}


def verify_style_mapped(target_path: Path, params: dict) -> bool:
    """验证无段落引用 from_style_id。"""
    try:
        target_path = Path(target_path)
        root = _read_part(target_path, "word/document.xml")
        for p in root.iter(f"{W}p"):
            pstyle = p.find(f"{W}pPr/{W}pStyle")
            if pstyle is not None and pstyle.get(f"{W}val") == params["from_style_id"]:
                return False
        return True
    except Exception:
        return False

def apply_paragraph_style(target_path: Path, params: dict) -> dict:
    """按语义选择器给段落挂模板命名样式（w:pStyle）。

    params: {selector, style_id}
    与 set_paragraph_format 的区别：后者写直接格式（pPr/rPr 属性），
    本动作挂样式引用——模板的"正文/标题"命名样式得以真正生效，
    用户后续在 Word 里改样式也能一处改全篇。

    幂等：已挂同一 style_id 的段落跳过。
    pStyle 必须位于 pPr 首个子元素（OOXML schema 顺序），缺失时补建 pPr。
    """
    from earlyz.actions.new_actions import _locate_paragraphs, _read_part, _read_styles_map, _write_part

    target_path = Path(target_path)
    selector = params.get("selector") or {}
    style_id = params.get("style_id")
    if not style_id:
        raise ValueError("apply_paragraph_style 需要 style_id")
    root = _read_part(target_path, "word/document.xml")
    paras = _locate_paragraphs(root, selector, _read_styles_map(target_path))
    changed = 0
    for p in paras:
        ppr = p.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            p.insert(0, ppr)
        ps = ppr.find(f"{W}pStyle")
        if ps is None:
            ps = etree.Element(f"{W}pStyle")
            ppr.insert(0, ps)
        if ps.get(f"{W}val") == style_id:
            continue
        ps.set(f"{W}val", style_id)
        changed += 1
    if changed:
        _write_part(target_path, "word/document.xml", root)
    return {"changed": changed > 0, "styled": changed, "style_id": style_id}


def verify_paragraph_style(target_path: Path, params: dict) -> bool:
    """验证选择器命中的段落全部引用指定样式。"""
    try:
        from earlyz.actions.new_actions import _locate_paragraphs, _read_part, _read_styles_map

        target_path = Path(target_path)
        selector = params.get("selector") or {}
        style_id = params.get("style_id")
        root = _read_part(target_path, "word/document.xml")
        paras = _locate_paragraphs(root, selector, _read_styles_map(target_path))
        if not paras:
            return True  # 无匹配 = 无可验证（与软跳过口径一致）
        for p in paras:
            ps = p.find(f"{W}pPr/{W}pStyle")
            if ps is None or ps.get(f"{W}val") != style_id:
                return False
        return True
    except Exception:
        return False

# 语义角色 → 模板样式名/ID 的匹配模式（按 Word 内置名与中文名）
_ROLE_STYLE_PATTERNS: dict[str, list[str]] = {
    "body": [r"^正文$", r"^正文文本$", r"^Body", r"^Normal$", r"^NormalIndent$"],
    "heading1": [r"^标题\s*1$", r"^Heading\s*1$", r"^标题一$"],
    "heading2": [r"^标题\s*2$", r"^Heading\s*2$", r"^标题二$"],
    "heading3": [r"^标题\s*3$", r"^Heading\s*3$", r"^标题三$"],
    "caption": [r"^题注$", r"^Caption$"],
}


def infer_role_styles(template_path: Path) -> dict[str, str]:
    """从模板 styles.xml 推断「语义角色 → 命名样式 ID」。

    只做名称匹配，不做版面猜测：模板定义样式，产品负责套用。
    """
    import re as _re

    root = _read_part(Path(template_path), "word/styles.xml")
    out: dict[str, str] = {}
    for style in root.iter(f"{W}style"):
        if style.get(f"{W}type") != "paragraph":
            continue
        sid = style.get(f"{W}styleId") or ""
        name_el = style.find(f"{W}name")
        name = (name_el.get(f"{W}val") if name_el is not None else "") or ""
        for role, pats in _ROLE_STYLE_PATTERNS.items():
            if role in out:
                continue
            if any(_re.match(p, name, _re.I) or _re.match(p, sid, _re.I) for p in pats):
                out[role] = sid
    return out


def _strip_style_jc(target_path: Path, style_id: str) -> bool:
    """删掉产物 styles.xml 里某个段落样式的 `w:jc`（只动对齐，不动字号/加粗）。"""
    from lxml import etree

    styles_root = _read_part(Path(target_path), "word/styles.xml")
    for st in styles_root.iter(f"{W}style"):
        if st.get(f"{W}styleId") != style_id:
            continue
        ppr = st.find(f"{W}pPr")
        if ppr is None:
            return False
        jc = ppr.find(f"{W}jc")
        if jc is None:
            return False
        ppr.remove(jc)
        _write_part(Path(target_path), "word/styles.xml", styles_root)
        return True
    return False


def wire_template_styles(target_path: Path, template_path: Path,
                         spec=None) -> dict:
    """把模板的命名样式克隆进产物并挂到对应角色段落上。

    非阻断：任何角色失败只记录，不影响排版交付。

    `spec` 给了就做一道对齐护栏（2026-11-01）：模板内置的 `Heading 1` 样式常带
    `jc=center` 而模板自己这一级段落根本没用它（用 `BodyText` + 直接格式左对齐）。
    按名字克隆挂上去 → 产物标题整片被拉成居中（建设项目工程总承包合同 32 段
    center≠both）。字号/加粗仍由该样式提供（产物一直靠它），只把**对齐**还给规则：
    该级规则没有 `jc` 时，删掉克隆样式里的 `w:jc`。
    """
    target_path = Path(target_path)
    template_path = Path(template_path)
    applied: dict = {}
    errors: list[str] = []
    rule_jc: dict[str, str | None] = {}
    if spec is not None:
        for r in getattr(spec, "rules", []) or []:
            sel = getattr(r, "selector", None) or {}
            role = sel.get("semantic_role")
            if role and not sel.get("title_key"):
                # 只看**整级规则**（不带 title_key 的那条）：同名标题规则（heading1#目录
                # jc=center）不代表整级标题怎么对齐——把它算进来就会漏掉总承包合同这种
                # 「整级无 jc、只有目录标题居中」的模板。
                # `expected` 可能是 bool（部分规则用它标「有期望」），不是 dict 就不读。
                _exp = getattr(r, "expected", None)
                _jc = _exp.get("jc") if isinstance(_exp, dict) else None
                if _jc is not None or role not in rule_jc:
                    rule_jc[role] = _jc
    for role, sid in infer_role_styles(template_path).items():
        try:
            clone_style_from_spec(target_path, {
                "style_id": sid, "source_style_id": sid,
                "source_docx_path": str(template_path),
            })
            info = apply_paragraph_style(target_path, {
                "selector": {"semantic_role": role}, "style_id": sid,
            })
            if info.get("styled"):
                applied[role] = {"style_id": sid, "styled": info["styled"]}
                if role.startswith("heading") and not rule_jc.get(role):
                    if _strip_style_jc(target_path, sid):
                        applied[role]["jc_stripped"] = True
        except Exception as e:  # noqa: BLE001
            errors.append(f"{role}: {e}")
    if errors:
        applied["_errors"] = errors[:5]
    return applied
