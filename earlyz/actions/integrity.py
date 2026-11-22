"""动作执行完整性层（Layer 0，2026-10-18）。

背景：审计发现三类"看不见的失败"——
1. `changed` 取的是"没抛异常的动作条数"，空操作也报成功；
2. 破坏性动作（删 headerReference、删 sectPr）无备份、无后置断言、失败不回滚；
3. 参数无 schema，键名写错被静默忽略（`border:"none"` 就是这么没生效的）。

本模块只做三件事，不改任何动作的排版语义：
- `snapshot()` / `diff_parts()`：按 part 记录文档前后签名，`changed` 以真实字节差为准；
- `validate_params()`：关键动作的必填/未知键校验（未知键直接拒，不再静默忽略）；
- `check_guard()`：破坏性动作的后置不变量断言，违反即回滚。
"""
from __future__ import annotations

import hashlib
import io
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

# 参与签名的 part：正文 + 页眉页脚 + 设置 + rels + content types。
# 只签这些，避免把图片/嵌入对象的大二进制反复哈希。
_SIGN_EXACT = ("[Content_Types].xml",)


def _signed(name: str) -> bool:
    """签名所有 XML part：正文/页眉页脚/样式/编号/脚注/设置/rels。

    只签 XML（图片等二进制不参与），避免每次动作都整体哈希几十 MB。
    """
    if name in _SIGN_EXACT:
        return True
    if name.startswith("word/_rels/"):
        return True
    return name.startswith("word/") and name.endswith(".xml")


@dataclass
class Snapshot:
    raw: bytes
    parts: dict[str, str] = field(default_factory=dict)
    sects: list[dict] = field(default_factory=list)
    header_texts: dict[str, str] = field(default_factory=dict)
    footer_texts: dict[str, str] = field(default_factory=dict)
    header_parts: int = 0
    footer_parts: int = 0
    xml_ok: bool = True


def _sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _sects_of(z: zipfile.ZipFile) -> list[dict]:
    """每节的页眉/页脚引用、页边距、页码类型、分节类型。"""
    out: list[dict] = []
    try:
        root = etree.fromstring(z.read("word/document.xml"))
    except Exception:
        return out
    for sp in root.iter(f"{W}sectPr"):
        typ = sp.find(f"{W}type")
        out.append({
            "type": (typ.get(f"{W}val") if typ is not None else "nextPage"),
            "headers": [(h.get(f"{W}type"), h.get(f"{R}id")) for h in sp.findall(f"{W}headerReference")],
            "footers": [(f.get(f"{W}type"), f.get(f"{R}id")) for f in sp.findall(f"{W}footerReference")],
            "has_pgMar": sp.find(f"{W}pgMar") is not None,
            "has_pgSz": sp.find(f"{W}pgSz") is not None,
            "has_pgNumType": sp.find(f"{W}pgNumType") is not None,
            "titlePg": sp.find(f"{W}titlePg") is not None,
        })
    return out


def _texts_of(z: zipfile.ZipFile, prefix: str) -> dict[str, str]:
    out: dict[str, str] = {}
    for name in z.namelist():
        if not (name.startswith(f"word/{prefix}") and name.endswith(".xml")):
            continue
        try:
            root = etree.fromstring(z.read(name))
        except Exception:
            continue
        out[name] = "\u0001".join(t.text or "" for t in root.iter(f"{W}t"))
    return out


def snapshot(path: Path) -> Snapshot:
    raw = Path(path).read_bytes()
    snap = Snapshot(raw=raw)
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as z:
            for name in z.namelist():
                if _signed(name) and name.endswith((".xml", ".rels")):
                    snap.parts[name] = _sha(z.read(name))
            snap.sects = _sects_of(z)
            snap.header_texts = _texts_of(z, "header")
            snap.footer_texts = _texts_of(z, "footer")
            snap.header_parts = len(snap.header_texts)
            snap.footer_parts = len(snap.footer_texts)
    except Exception:
        snap.xml_ok = False
    return snap


def diff_parts(before: Snapshot, after: Snapshot) -> dict:
    changed = sorted(k for k in before.parts if after.parts.get(k) != before.parts[k])
    added = sorted(set(after.parts) - set(before.parts))
    removed = sorted(set(before.parts) - set(after.parts))
    return {"changed": changed, "added": added, "removed": removed}


def is_changed(diff: dict) -> bool:
    return bool(diff["changed"] or diff["added"] or diff["removed"])


# ─────────────────────────────────────────────────────────────
# 参数校验：必填 + 未知键
# ─────────────────────────────────────────────────────────────

# 段落/字符属性动作允许的键（format_headers / format_footers 的 properties 内层）
_FORMAT_PROP_KEYS = {
    "p_bdr_bottom", "ppr_extra_xml", "rpr_extra_xml",
    "font_eastAsia", "font_ascii", "font_size_pt", "size_pt", "bold", "italic",
    "underline", "color", "jc", "alignment", "indent_first_line_chars",
    "line_spacing", "line_rule", "space_before_pt", "space_after_pt",
}

PARAM_RULES: dict[str, dict] = {
    "clear_section_headers": {"required": ["section_index"],
                              "allowed": {"section_index", "allow_document_wide"}},
    "clear_section_footers": {"required": ["section_index"],
                              "allowed": {"section_index", "allow_document_wide"}},
    "create_section_headers": {
        "required": ["section_headers"],
        "allowed": {"section_headers", "font_eastAsia", "font_ascii", "font_size", "border"},
    },
    "format_headers": {"required": ["properties"], "allowed": {"properties", "sections", "parts"},
                       "property_keys": _FORMAT_PROP_KEYS},
    "format_footers": {"required": ["properties"], "allowed": {"properties", "sections", "parts"},
                       "property_keys": _FORMAT_PROP_KEYS},
    "set_header_text": {"required": ["header_path", "text"], "allowed": {"header_path", "text"}},
    "set_header_font": {"required": ["header_path"], "allowed": {"header_path", "font_eastAsia",
                                                                 "font_ascii", "size_pt", "font_size_pt", "bold"}},
    "insert_section_break": {"required": ["heading_text"], "allowed": {"heading_text"}},
    "set_sectpr_type": {"required": ["section_index", "val"], "allowed": {"section_index", "val"}},
    "add_footer_reference": {"required": ["section_index"], "allowed": {"section_index", "footer_path"}},
    "add_page_number": {"required": ["footer_path"], "allowed": {"footer_path", "format", "alignment",
                                                                 "font_eastAsia", "font_ascii", "font_size"}},
    "set_title_page_flag": {"required": ["enabled"], "allowed": {"section_index", "enabled"}},
    "set_even_odd_headers_flag": {"required": ["enabled"], "allowed": {"enabled"}},
    "apply_template_header_footer": {
        "required": ["template_parts", "template_slots"],
        "allowed": {"template_parts", "template_slots", "create_missing"},
    },
    "build_section_skeleton": {"required": ["boundaries"], "allowed": {"boundaries"}},
    "ensure_footer_parts": {"required": [], "allowed": {"format", "alignment", "size_pt", "skip_first"}},
    "split_section_parts": {"required": [], "allowed": set()},
}


def validate_params(action: str, params: dict) -> list[str]:
    """返回错误列表；空列表表示通过。未登记的动作不做校验（保持原行为）。"""
    rule = PARAM_RULES.get(action)
    if rule is None:
        return []
    params = params or {}
    errors: list[str] = []
    for key in rule["required"]:
        if key not in params:
            errors.append(f"缺少必填参数 {key}")
    unknown = sorted(set(params) - set(rule["allowed"]))
    if unknown:
        errors.append(f"未知参数 {unknown}，该动作只接受 {sorted(rule['allowed'])}")
    prop_keys = rule.get("property_keys")
    if prop_keys is not None and isinstance(params.get("properties"), dict):
        bad = sorted(set(params["properties"]) - set(prop_keys))
        if bad:
            errors.append(f"properties 里有不支持的键 {bad}，可用键：{sorted(prop_keys)}")
    return errors


# ─────────────────────────────────────────────────────────────
# 破坏性动作后置守卫
# ─────────────────────────────────────────────────────────────

# action → 语义类型
DESTRUCTIVE_ACTIONS = {
    "clear_section_headers", "clear_section_footers", "remove_extra_sectpr",
    "create_section_headers", "insert_section_break", "set_header_text",
    "apply_template_header_footer", "format_headers", "format_footers",
    "clone_header_footer_from_spec", "normalize_footer_parts",
    "repair_section_references", "add_footer_reference", "add_page_number",
    "build_section_skeleton",
}


def _scoped_parts(path: Path, params: dict, prefix: str) -> set[str] | None:
    """按 sections / parts 参数解析出本次动作真正作用的部件；None 表示全部。"""
    parts = params.get("parts")
    if isinstance(parts, list) and parts:
        return {p for p in parts if isinstance(p, str)}
    sections = params.get("sections")
    if not (isinstance(sections, list) and sections):
        return None
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    rid = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    try:
        with zipfile.ZipFile(path) as z:
            doc = etree.fromstring(z.read("word/document.xml"))
            rels = etree.fromstring(z.read("word/_rels/document.xml.rels"))
    except Exception:  # noqa: BLE001
        return None
    rid2target = {r.get("Id"): str(r.get("Target") or "") for r in rels.iter(f"{{{rel_ns}}}Relationship")}
    sects = list(doc.iter(f"{W}sectPr"))
    out: set[str] = set()
    for si in sections:
        if not isinstance(si, int) or si < 0 or si >= len(sects):
            continue
        for ref in sects[si].findall(f"{W}{prefix}Reference"):
            target = rid2target.get(ref.get(rid))
            if target:
                out.add("word/" + target.lstrip("/"))
    return out or None


def check_post(action: str, path: Path, params: dict) -> list[str]:
    """动作执行后的语义断言（Layer 0 自有，不依赖旧 verifier 接线）。

    覆盖今天事故里的三个语义：
    - `format_headers/footers` 声称改了边框 → 边框必须真在/真没了；
    - `create_section_headers` → 每个被点名的节必须挂上引用；
    - `insert_section_break` → 每个 sectPr 必须带 pgSz/pgMar。
    """
    problems: list[str] = []
    props = (params or {}).get("properties") or {}

    if action in ("format_headers", "format_footers") and "p_bdr_bottom" in props:
        want = bool(props.get("p_bdr_bottom"))
        prefix = "header" if action == "format_headers" else "footer"
        scoped = _scoped_parts(path, params or {}, prefix)
        try:
            with zipfile.ZipFile(path) as z:
                for name in z.namelist():
                    if not (name.startswith(f"word/{prefix}") and name.endswith(".xml")):
                        continue
                    if scoped is not None and name not in scoped:
                        continue
                    root = etree.fromstring(z.read(name))
                    for pel in root.iter(f"{W}p"):
                        has_text = bool("".join(t.text or "" for t in pel.iter(f"{W}t")).strip())
                        ppr = pel.find(f"{W}pPr")
                        pbdr = ppr.find(f"{W}pBdr") if ppr is not None else None
                        bottom = pbdr.find(f"{W}bottom") if pbdr is not None else None
                        exists = bottom is not None and (bottom.get(f"{W}val") or "single") != "none"
                        if want and has_text and not exists:
                            problems.append(f"{name} 的段落没有下边框")
                        if not want and exists:
                            problems.append(f"{name} 仍有下边框，删除未生效")
        except Exception as e:  # noqa: BLE001
            problems.append(f"读取 {prefix} 部件失败：{e}")

    if action == "create_section_headers":
        entries = (params or {}).get("section_headers")
        if not isinstance(entries, list) or not entries:
            problems.append("section_headers 为空，本动作不应记为成功")
        else:
            try:
                with zipfile.ZipFile(path) as z:
                    root = etree.fromstring(z.read("word/document.xml"))
                sects = list(root.iter(f"{W}sectPr"))
                for entry in entries:
                    si = entry.get("sectpr_index") if isinstance(entry, dict) else None
                    if not isinstance(si, int) or si < 0 or si >= len(sects):
                        problems.append(f"节序号 {si!r} 不存在（共 {len(sects)} 节）")
                    elif not sects[si].findall(f"{W}headerReference"):
                        problems.append(f"第 {si} 节创建后仍没有 headerReference")
            except Exception as e:  # noqa: BLE001
                problems.append(f"读取 document.xml 失败：{e}")

    if action == "insert_section_break":
        try:
            with zipfile.ZipFile(path) as z:
                root = etree.fromstring(z.read("word/document.xml"))
            for i, sp in enumerate(root.iter(f"{W}sectPr")):
                if sp.find(f"{W}pgSz") is None or sp.find(f"{W}pgMar") is None:
                    problems.append(f"第 {i} 节缺 pgSz/pgMar（分节后版式会跳变）")
        except Exception as e:  # noqa: BLE001
            problems.append(f"读取 document.xml 失败：{e}")

    return problems


def check_guard(action: str, before: Snapshot, after: Snapshot, params: dict) -> list[str]:
    """破坏性动作的不变量检查。返回问题列表；非空即应回滚。"""
    if action not in DESTRUCTIVE_ACTIONS:
        return []
    if not after.xml_ok:
        return ["产物 document.xml 已无法解析"]

    problems: list[str] = []
    n_before, n_after = len(before.sects), len(after.sects)

    # 通用：页眉/页脚部件不许凭空消失
    if after.header_parts < before.header_parts:
        problems.append(f"页眉部件被删除（{before.header_parts} → {after.header_parts}）")
    if after.footer_parts < before.footer_parts:
        problems.append(f"页脚部件被删除（{before.footer_parts} → {after.footer_parts}）")

    # 通用：未被点名的节，其引用不许变
    target = params.get("section_index")
    if isinstance(target, int) and target < min(n_before, n_after):
        for i in range(min(n_before, n_after)):
            if i == target:
                continue
            if before.sects[i]["headers"] != after.sects[i]["headers"]:
                problems.append(f"第 {i} 节页眉引用被一并改动")
            if before.sects[i]["footers"] != after.sects[i]["footers"]:
                problems.append(f"第 {i} 节页脚引用被一并改动")

    if action in ("clear_section_headers", "clear_section_footers"):
        if n_after != n_before:
            problems.append(f"节数被改动（{n_before} → {n_after}）")
        # 2026-10-18 真机事故：单节文档清「封面页眉」= 把全文唯一一套页眉摘掉。
        # 语义上这是「全文去页眉」，必须显式声明；否则回滚并要求改用边框能力。
        key = "headers" if action == "clear_section_headers" else "footers"
        before_total = sum(len(s[key]) for s in before.sects)
        after_total = sum(len(s[key]) for s in after.sects)
        if before_total and after_total == 0 and not params.get("allow_document_wide"):
            problems.append(
                f"该动作会移除全文所有{'页眉' if key == 'headers' else '页脚'}引用"
                f"（文档共 {n_after} 节）；只想改样式请用 format_headers(p_bdr_bottom=false)，"
                f"确需全局移除请显式传 allow_document_wide=true"
            )

    if action in ("format_headers", "format_footers"):
        # 只改格式：页眉/页脚文字与引用一律不许变
        if before.header_texts != after.header_texts:
            problems.append("页眉文字被改动（本动作只允许改格式）")
        if before.footer_texts != after.footer_texts:
            problems.append("页脚文字被改动（本动作只允许改格式）")
        if [s["headers"] for s in before.sects] != [s["headers"] for s in after.sects]:
            problems.append("页眉引用被改动（本动作只允许改格式）")

    if action == "insert_section_break":
        if n_after != n_before + 1:
            problems.append(f"节数应 +1，实际 {n_before} → {n_after}")
        else:
            new_sects = [s for s in after.sects if not s["has_pgMar"] or not s["has_pgSz"]]
            if new_sects:
                problems.append("新分节缺少 pgSz/pgMar（版式会跳变）")
            total_refs_before = sum(len(s["headers"]) for s in before.sects)
            total_refs_after = sum(len(s["headers"]) for s in after.sects)
            if total_refs_after == 0 and total_refs_before > 0:
                problems.append("分节后页眉引用全丢")

    if action == "create_section_headers":
        if n_after != n_before:
            problems.append(f"节数被改动（{n_before} → {n_after}）")
        for i, s in enumerate(before.sects):
            if i < n_after and s["headers"] and not after.sects[i]["headers"]:
                problems.append(f"第 {i} 节原有页眉引用被挤掉")

    if action == "build_section_skeleton":
        # 幂等：第二次跑允许 0 变化；只有"节数反而减少"才拦
        if n_after < n_before:
            problems.append(f"节数减少（{n_before} → {n_after}）")
        for i, s in enumerate(after.sects):
            if not (s["has_pgSz"] and s["has_pgMar"]):
                problems.append(f"第 {i} 节缺 pgSz/pgMar：新节版式不完整")
        if before.header_parts and not any(s["headers"] for s in after.sects):
            problems.append("分节后页眉引用全丢")

    if action == "remove_extra_sectpr":
        if n_after > n_before:
            problems.append(f"节数不应增加（{n_before} → {n_after}）")
        for i, s in enumerate(after.sects):
            if not s["has_pgSz"]:
                problems.append(f"第 {i} 节缺 pgSz")

    return problems
