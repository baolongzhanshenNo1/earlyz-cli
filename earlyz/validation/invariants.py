"""文档结构不变量门禁（Layer 2，2026-10-18）。

目的：把「节/页眉/页脚/页码」这类结构性事实变成可断言的检查，防止
"结构塌了但每一步都报成功"再次发生。

检查项（errors 代表包结构已经不可信；warnings 代表能力缺口/可疑形态）：
- document.xml 可解析；至少 1 个 sectPr；每个 sectPr 带 pgSz + pgMar；
- 每个 header/footer 引用：rId 能在 rels 解析、目标部件存在、部件可解析、
  在 [Content_Types].xml 有 Override；
- 同一节同一类型不得出现重复引用；
- 多节文档：正文节应有 footerReference（否则没有页码）；
- titlePg 打开时必须存在 first 槽位引用；
- 存在但无人引用的 header/footer 部件 → warning（孤儿部件）；
- settings 里 evenAndOddHeaders 未开却引用了 even 槽位 → warning（奇偶页眉不生效）。
"""
from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"


def check_document_invariants(path: str | Path) -> dict:
    path = Path(path)
    errors: list[str] = []
    warnings: list[str] = []
    stats: dict = {}

    if not path.is_file():
        return {"ok": False, "errors": ["文件不存在"], "warnings": [], "stats": {}}

    try:
        z = zipfile.ZipFile(path)
    except Exception as e:  # noqa: BLE001
        return {"ok": False, "errors": [f"无法打开 docx：{e}"], "warnings": [], "stats": {}}

    names = set(z.namelist())
    if "word/document.xml" not in names:
        return {"ok": False, "errors": ["缺少 word/document.xml"], "warnings": [], "stats": {}}
    try:
        doc = etree.fromstring(z.read("word/document.xml"))
    except Exception as e:  # noqa: BLE001
        return {"ok": False, "errors": [f"document.xml 解析失败：{e}"], "warnings": [], "stats": {}}

    rels: dict[str, str] = {}
    rels_part = "word/_rels/document.xml.rels"
    if rels_part in names:
        try:
            rroot = etree.fromstring(z.read(rels_part))
            rels = {r.get("Id"): str(r.get("Target") or "") for r in rroot.iter(f"{PR}Relationship")}
        except Exception as e:  # noqa: BLE001
            errors.append(f"document.xml.rels 解析失败：{e}")
    else:
        errors.append("缺少 word/_rels/document.xml.rels")

    overrides: set[str] = set()
    if "[Content_Types].xml" in names:
        try:
            croot = etree.fromstring(z.read("[Content_Types].xml"))
            overrides = {o.get("PartName") or "" for o in croot.iter(f"{CT}Override")}
        except Exception as e:  # noqa: BLE001
            errors.append(f"[Content_Types].xml 解析失败：{e}")

    sects = list(doc.iter(f"{W}sectPr"))
    stats["sections"] = len(sects)
    if not sects:
        errors.append("文档没有任何 sectPr")

    referenced: set[str] = set()
    header_parts = sorted(n for n in names if n.startswith("word/header") and n.endswith(".xml"))
    footer_parts = sorted(n for n in names if n.startswith("word/footer") and n.endswith(".xml"))
    stats["header_parts"] = len(header_parts)
    stats["footer_parts"] = len(footer_parts)

    titlepg_without_first = 0
    even_refs = 0
    for si, sp in enumerate(sects):
        if sp.find(f"{W}pgSz") is None:
            errors.append(f"第 {si} 节缺 pgSz")
        if sp.find(f"{W}pgMar") is None:
            errors.append(f"第 {si} 节缺 pgMar")

        for kind in ("header", "footer"):
            refs = sp.findall(f"{W}{kind}Reference")
            seen_types: list[str] = []
            for ref in refs:
                t = ref.get(f"{W}type") or "default"
                if t in seen_types:
                    errors.append(f"第 {si} 节 {kind} 槽位 {t} 重复引用")
                seen_types.append(t)
                if t == "even":
                    even_refs += 1
                rid = ref.get(RID)
                target = rels.get(rid or "")
                if not target:
                    errors.append(f"第 {si} 节 {kind} 引用 {rid!r} 在 rels 中不存在")
                    continue
                part = "word/" + target.lstrip("/").replace("word/", "")
                referenced.add(part)
                if part not in names:
                    errors.append(f"第 {si} 节引用的部件 {part} 不在包内")
                    continue
                if ("/" + part) not in overrides:
                    errors.append(f"部件 {part} 缺少 [Content_Types] Override（Word 会报损坏）")
                try:
                    etree.fromstring(z.read(part))
                except Exception as e:  # noqa: BLE001
                    errors.append(f"部件 {part} 解析失败：{e}")

        if sp.find(f"{W}titlePg") is not None:
            if not any((r.get(f"{W}type") == "first") for r in sp.findall(f"{W}headerReference")):
                titlepg_without_first += 1

    if titlepg_without_first:
        warnings.append(f"{titlepg_without_first} 个节启用了 titlePg 但没有 first 页眉引用（首页页眉不会生效）")

    if len(sects) > 1:
        body_like = [si for si, sp in enumerate(sects) if si > 0]
        if body_like and all(len(sects[si].findall(f"{W}footerReference")) == 0 for si in body_like):
            warnings.append("多节文档但第 1 节之后都没有 footerReference（正文没有页码）")

    orphans = [p for p in (header_parts + footer_parts) if p not in referenced]
    if orphans:
        warnings.append(f"{len(orphans)} 个页眉/页脚部件无人引用（孤儿部件）：{', '.join(orphans[:4])}")

    if even_refs:
        try:
            settings = etree.fromstring(z.read("word/settings.xml")) if "word/settings.xml" in names else None
        except Exception:  # noqa: BLE001
            settings = None
        if settings is None or settings.find(f"{W}evenAndOddHeaders") is None:
            warnings.append("引用了 even 页眉/页脚，但 settings.xml 未开启 evenAndOddHeaders（奇偶页不生效）")

    stats["referenced_parts"] = len(referenced)
    z.close()
    return {"ok": not errors, "errors": errors, "warnings": warnings, "stats": stats}


def assert_invariants(path: str | Path) -> dict:
    """测试/脚本用：errors 非空直接断言失败，warnings 原样返回。"""
    result = check_document_invariants(path)
    assert result["ok"], "文档不变量失败：\n" + "\n".join(result["errors"][:8])
    return result
