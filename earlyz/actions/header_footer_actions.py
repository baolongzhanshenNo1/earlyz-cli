"""早睡 (EarlyZ) 3.2 — Phase 2 复杂页眉页脚（first/even/default 三态，论文前置部分核心）。

set_header_reference      — 创建 header part（文本 + 可选 PAGE 页码域）并绑定到
                            section 的 headerReference（default/first/even 三态）
set_footer_reference      — 同上（footer part）
set_title_page_flag       — 启用/关闭 section 首页不同（titlePg）
set_even_odd_headers_flag — 启用/关闭奇偶页不同（settings.xml evenAndOddHeaders）

实现要点：
- header/footer part 命名唯一（header1.xml 起递增），防重复；
- first 态自动启用 titlePg；even 态自动启用 evenAndOddHeaders；
- 幂等：同参数重复执行不产生重复 part/引用；
- 全部带 verifier。
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Optional
from lxml import etree

from earlyz.actions.new_actions import _read_part, _write_part

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"

_HEADER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
_FOOTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"


def _next_part_number(z: zipfile.ZipFile, prefix: str) -> int:
    nums = []
    for n in z.namelist():
        m = re.match(rf"^word/{prefix}(\d+)\.xml$", n)
        if m:
            nums.append(int(m.group(1)))
    return (max(nums) + 1) if nums else 1


def _ensure_part_registered(path: Path, part_name: str, content_type: str) -> None:
    """注册 ContentTypes Override + document.xml.rels（幂等）。"""
    # ContentTypes
    ct_root = _read_part(path, "[Content_Types].xml")
    exists = any(
        o.get("PartName") == f"/{part_name}"
        for o in ct_root.findall(f"{CT}Override")
    )
    if not exists:
        ov = etree.SubElement(ct_root, f"{CT}Override")
        ov.set("PartName", f"/{part_name}")
        ov.set("ContentType", content_type)
        _write_part(path, "[Content_Types].xml", ct_root)

    # document.xml.rels
    rels_root = _read_part(path, "word/_rels/document.xml.rels")
    targets = [r.get("Target") for r in rels_root.findall(f"{PR}Relationship")]
    rel_id = None
    if part_name.split("/")[-1] not in [t.split("/")[-1] for t in targets if t]:
        # 取最大 rId 递增
        ids = [int(r.get("Id", "rId0")[3:]) for r in rels_root.findall(f"{PR}Relationship")]
        rel_id = f"rId{max(ids) + 1 if ids else 2}"
        rel = etree.SubElement(rels_root, f"{PR}Relationship")
        rel.set("Id", rel_id)
        rel.set("Type", _HEADER_REL_TYPE if "header" in part_name else _FOOTER_REL_TYPE)
        rel.set("Target", part_name.split("/")[-1])
        _write_part(path, "word/_rels/document.xml.rels", rels_root)


def _bind_reference(path: Path, section_index: int, ref_tag: str, ref_type: str, rel_id: str) -> None:
    """在 section 的 sectPr 上设置 headerReference/footerReference（幂等替换）。

    OOXML CT_SectPr 顺序约束：headerReference/footerReference 必须位于
    pgSz 等布局元素之前（否则 Word 打开报错）。插入点取第一个非引用元素，
    找不到则 append（空/异常 sectPr 兜底）。
    """
    doc_root = _read_part(path, "word/document.xml")
    sects = list(doc_root.iter(f"{W}sectPr"))
    if section_index >= len(sects):
        raise ValueError(f"section index {section_index} 超出范围（共 {len(sects)} 节）")
    sect = sects[section_index]
    # 删除同类型旧引用（幂等）。default 引用无 w:type 属性（None），
    # ref_type="default" 时必须匹配无 type 的旧引用，否则旧引用残留。
    for old in sect.findall(f"{W}{ref_tag}"):
        old_type = old.get(f"{W}type")
        if old_type == ref_type or (ref_type == "default" and old_type is None):
            sect.remove(old)
    # 计算插入点：第一个既非 headerReference 也非 footerReference 的子元素
    insert_at = len(sect)
    for i, child in enumerate(sect):
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag not in ("headerReference", "footerReference"):
            insert_at = i
            break
    ref = etree.Element(f"{W}{ref_tag}")
    ref.set(f"{W}type", ref_type)
    ref.set(f"{R}id", rel_id)
    sect.insert(insert_at, ref)
    _write_part(path, "word/document.xml", doc_root)


def _set_settings_flag(path: Path, tag: str, enabled: bool) -> None:
    # 极简 OOXML 可以不带 word/settings.xml（合成夹具实测），
    # 缺失时按需新建最小 settings 部件并登记 Content_Types，
    # 而不是抛 KeyError 让整份排版失败。
    try:
        settings_root = _read_part(path, "word/settings.xml")
    except KeyError:
        settings_root = etree.Element(f"{W}settings", nsmap={"w": W[1:-1]})
        settings_root.append(etree.Element(f"{W}zoom", {f"{W}percent": "100"}))
        _write_part(path, "word/settings.xml", settings_root)
        _ensure_part_registered(
            path, "word/settings.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
        )
    el = settings_root.find(f"{W}{tag}")
    if enabled and el is None:
        etree.SubElement(settings_root, f"{W}{tag}")
        _write_part(path, "word/settings.xml", settings_root)
    elif not enabled and el is not None:
        settings_root.remove(el)
        _write_part(path, "word/settings.xml", settings_root)


def _part_content(text: str, page_number: bool, suffix: str = "") -> bytes:
    """构造 hdr/ftr 内容：文本段落 + 可选 PAGE 域。"""
    hdr = etree.Element(f"{W}hdr")
    p = etree.SubElement(hdr, f"{W}p")
    ppr = etree.SubElement(p, f"{W}pPr")
    jc = etree.SubElement(ppr, f"{W}jc")
    jc.set(f"{W}val", "center")
    if text:
        r = etree.SubElement(p, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        t.text = text
    if page_number:
        r = etree.SubElement(p, f"{W}r")
        etree.SubElement(r, f"{W}fldChar", {f"{W}fldCharType": "begin"})
        r2 = etree.SubElement(p, f"{W}r")
        instr = etree.SubElement(r2, f"{W}instrText")
        instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        instr.text = " PAGE "
        r3 = etree.SubElement(p, f"{W}r")
        etree.SubElement(r3, f"{W}fldChar", {f"{W}fldCharType": "end"})
    if suffix:
        r = etree.SubElement(p, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        t.text = suffix
    return etree.tostring(hdr, xml_declaration=True, encoding="UTF-8", standalone=True)


def _create_part(path: Path, prefix: str, content_type: str, content: bytes) -> str:
    """创建 word/headerN.xml 或 footerN.xml，返回 part 名。"""
    with zipfile.ZipFile(path, "r") as z:
        n = _next_part_number(z, prefix)
    part_name = f"word/{prefix}{n}.xml"
    with zipfile.ZipFile(path, "a") as z:
        zi = zipfile.ZipInfo(part_name)
        zi.compress_type = zipfile.ZIP_DEFLATED  # Word 要求 XML part DEFLATED（STORED 报文件损坏）
        z.writestr(zi, content)
    _ensure_part_registered(path, part_name, content_type)
    return part_name


def _header_content_type() -> str:
    return "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"


def _footer_content_type() -> str:
    return "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"


# ────────────────────────────────────────────────────────────────
# set_header_reference
# ────────────────────────────────────────────────────────────────

def set_header_reference(target_path: Path, params: dict) -> dict:
    """设置 section 页眉（三态）。

    params: {section_index?, header_type: default/first/even, text?, page_number?, suffix?}
    幂等：同参数重复执行不产生重复 part（已有绑定且内容一致 → no-op）。
    """
    target_path = Path(target_path)
    ref_type = params.get("header_type", "default")
    section_index = int(params.get("section_index", -1))  # -1 = 最后节（正文节）
    text = params.get("text") or ""
    page_number = bool(params.get("page_number", False))
    suffix = params.get("suffix") or ""

    # 幂等检查：现有绑定已满足期望 → no-op
    if verify_header_reference(target_path, params):
        return {"changed": False, "idempotent": True, "type": ref_type}

    content = _part_content(text, page_number, suffix)
    part_name = _create_part(target_path, "header", _header_content_type(), content)
    # rel id：重新读 rels 找该 part 对应 rId
    rels_root = _read_part(target_path, "word/_rels/document.xml.rels")
    rel_id = None
    for r in rels_root.findall(f"{PR}Relationship"):
        if (r.get("Target") or "").endswith(part_name.split("/")[-1]):
            rel_id = r.get("Id")
            break
    if rel_id is None:
        raise RuntimeError(f"header part 未注册 rel: {part_name}")

    with zipfile.ZipFile(target_path) as z:
        n_sects = len(list(etree.fromstring(z.read("word/document.xml")).iter(f"{W}sectPr")))
    if section_index < 0:
        section_index = n_sects - 1
    _bind_reference(target_path, section_index, "headerReference", ref_type, rel_id)
    if ref_type == "first":
        set_title_page_flag(target_path, {"section_index": section_index, "enabled": True})
    elif ref_type == "even":
        set_even_odd_headers_flag(target_path, {"enabled": True})
    return {"changed": True, "part": part_name, "type": ref_type}


def _verify_part_bound(path: Path, part_prefix: str, ref_tag: str, ref_type: str, text: str, page_number: bool) -> bool:
    try:
        path = Path(path)
        doc_root = _read_part(path, "word/document.xml")
        sect = list(doc_root.iter(f"{W}sectPr"))[-1]
        ref = sect.find(f"{W}{ref_tag}")
        if ref is None or ref.get(f"{W}type") != ref_type:
            return False
        rel_id = ref.get(f"{R}id")
        rels_root = _read_part(path, "word/_rels/document.xml.rels")
        target = None
        for r in rels_root.findall(f"{PR}Relationship"):
            if r.get("Id") == rel_id:
                target = r.get("Target")
                break
        if not target:
            return False
        part_root = _read_part(path, f"word/{target.split('/')[-1]}")
        texts = "".join(t.text or "" for t in part_root.iter(f"{W}t"))
        if text and text not in texts:
            return False
        if page_number:
            instr = "".join(i.text or "" for i in part_root.iter(f"{W}instrText"))
            if "PAGE" not in instr:
                return False
        return True
    except Exception:
        return False


def verify_header_reference(target_path: Path, params: dict) -> bool:
    return _verify_part_bound(
        target_path, "header", "headerReference",
        params.get("header_type", "default"), params.get("text") or "",
        bool(params.get("page_number", False)),
    )


# ────────────────────────────────────────────────────────────────
# set_footer_reference
# ────────────────────────────────────────────────────────────────

def set_footer_reference(target_path: Path, params: dict) -> dict:
    """设置 section 页脚（三态）。

    params: {section_index?, footer_type: default/first/even, text?, page_number?, suffix?}
    幂等：同参数重复执行不产生重复 part。
    """
    target_path = Path(target_path)
    ref_type = params.get("footer_type", "default")
    section_index = int(params.get("section_index", -1))
    text = params.get("text") or ""
    page_number = bool(params.get("page_number", False))
    suffix = params.get("suffix") or ""

    # 幂等检查
    if verify_footer_reference(target_path, params):
        return {"changed": False, "idempotent": True, "type": ref_type}

    content = _part_content(text, page_number, suffix)
    part_name = _create_part(target_path, "footer", _footer_content_type(), content)
    rels_root = _read_part(target_path, "word/_rels/document.xml.rels")
    rel_id = None
    for r in rels_root.findall(f"{PR}Relationship"):
        if (r.get("Target") or "").endswith(part_name.split("/")[-1]):
            rel_id = r.get("Id")
            break
    if rel_id is None:
        raise RuntimeError(f"footer part 未注册 rel: {part_name}")

    with zipfile.ZipFile(target_path) as z:
        n_sects = len(list(etree.fromstring(z.read("word/document.xml")).iter(f"{W}sectPr")))
    if section_index < 0:
        section_index = n_sects - 1
    _bind_reference(target_path, section_index, "footerReference", ref_type, rel_id)
    if ref_type == "first":
        set_title_page_flag(target_path, {"section_index": section_index, "enabled": True})
    elif ref_type == "even":
        set_even_odd_headers_flag(target_path, {"enabled": True})
    return {"changed": True, "part": part_name, "type": ref_type}


def verify_footer_reference(target_path: Path, params: dict) -> bool:
    return _verify_part_bound(
        target_path, "footer", "footerReference",
        params.get("footer_type", "default"), params.get("text") or "",
        bool(params.get("page_number", False)),
    )


# ────────────────────────────────────────────────────────────────
# set_title_page_flag / set_even_odd_headers_flag
# ────────────────────────────────────────────────────────────────

def set_title_page_flag(target_path: Path, params: dict) -> dict:
    """启用/关闭 section 首页不同（titlePg）。"""
    target_path = Path(target_path)
    section_index = int(params.get("section_index", -1))
    enabled = bool(params.get("enabled", True))
    doc_root = _read_part(target_path, "word/document.xml")
    sects = list(doc_root.iter(f"{W}sectPr"))
    if section_index < 0:
        section_index = len(sects) - 1
    if section_index >= len(sects):
        raise ValueError(f"section index {section_index} 超出范围")
    sect = sects[section_index]
    el = sect.find(f"{W}titlePg")
    if enabled and el is None:
        # CT_SectPr 顺序约束：titlePg 必须位于 cols 之后、docGrid 之前；
        # insert(0,)/addnext(type) 会压到 headerReference/pgSz 前 → Word 报损坏
        from earlyz.pipeline.print_layout import _insert_sectpr_child
        _insert_sectpr_child(sect, etree.Element(f"{W}titlePg"))
        _write_part(target_path, "word/document.xml", doc_root)
    elif not enabled and el is not None:
        sect.remove(el)
        _write_part(target_path, "word/document.xml", doc_root)
    return {"changed": True, "title_page": enabled}


def verify_title_page_flag(target_path: Path, params: dict) -> bool:
    try:
        target_path = Path(target_path)
        section_index = int(params.get("section_index", -1))
        enabled = bool(params.get("enabled", True))
        doc_root = _read_part(target_path, "word/document.xml")
        sects = list(doc_root.iter(f"{W}sectPr"))
        if section_index < 0:
            section_index = len(sects) - 1
        has = sects[section_index].find(f"{W}titlePg") is not None
        return has == enabled
    except Exception:
        return False


def set_even_odd_headers_flag(target_path: Path, params: dict) -> dict:
    """启用/关闭奇偶页不同（settings.xml evenAndOddHeaders）。"""
    target_path = Path(target_path)
    enabled = bool(params.get("enabled", True))
    _set_settings_flag(target_path, "evenAndOddHeaders", enabled)
    return {"changed": True, "even_odd": enabled}


def verify_even_odd_headers_flag(target_path: Path, params: dict) -> bool:
    try:
        target_path = Path(target_path)
        enabled = bool(params.get("enabled", True))
        settings = _read_part(target_path, "word/settings.xml")
        has = settings.find(f"{W}evenAndOddHeaders") is not None
        return has == enabled
    except Exception:
        return False

def ensure_title_pg_flag(target_path: Path, params: dict) -> dict:
    """按「是否存在 first 页眉/页脚引用」对齐 sectPr 的 titlePg 标志。

    规则（模板定义结构，产品负责一致化）：
      · 某节有 type="first" 的 header/footer 引用 → 必须有 titlePg（否则 Word 不显示）
      · 某节没有 first 引用 → 不应有 titlePg（否则首页留白却没有内容）
    幂等；返回每节处理结果。
    """
    target_path = Path(target_path)
    root = _read_part(target_path, "word/document.xml")
    changed = 0
    details = []
    for sect in root.iter(f"{W}sectPr"):
        has_first = any(
            ref.get(f"{W}type") == "first"
            for ref in list(sect.findall(f"{W}headerReference"))
                       + list(sect.findall(f"{W}footerReference"))
        )
        flag = sect.find(f"{W}titlePg")
        if has_first and flag is None:
            # titlePg 在 schema 中位于 pgSz 之前
            el = etree.Element(f"{W}titlePg")
            anchor = None
            for tag in ("pgSz", "pgMar", "type"):
                anchor = sect.find(f"{W}{tag}")
                if anchor is not None:
                    break
            if anchor is not None:
                anchor.addprevious(el)
            else:
                sect.insert(0, el)
            changed += 1
            details.append("added")
        elif not has_first and flag is not None:
            sect.remove(flag)
            changed += 1
            details.append("removed")
        else:
            details.append("ok")
    if changed:
        _write_part(target_path, "word/document.xml", root)
    return {"changed": changed > 0, "sections": len(details),
            "actions": details[:8]}


def verify_title_pg_flag(target_path: Path, params: dict) -> bool:
    try:
        target_path = Path(target_path)
        root = _read_part(target_path, "word/document.xml")
        for sect in root.iter(f"{W}sectPr"):
            has_first = any(
                ref.get(f"{W}type") == "first"
                for ref in list(sect.findall(f"{W}headerReference"))
                       + list(sect.findall(f"{W}footerReference"))
            )
            flag = sect.find(f"{W}titlePg")
            if has_first != (flag is not None):
                return False
        return True
    except Exception:
        return False
