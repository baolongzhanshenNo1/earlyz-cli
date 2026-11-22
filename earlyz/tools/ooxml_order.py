"""OOXML 子元素顺序：唯一真源 + 收尾规范化 + 违规检查（2026-09-28 用户判决落地）。

## 为什么要有这个模块

OOXML 的 `CT_PPr` / `CT_RPr` / `CT_TblPr` / `CT_TrPr` 都是 **sequence**（有序），
子元素顺序错了 Word 可能**静默忽略整段属性**（不报错、不提示，只是没生效）。
真机实证（2026-09-28 夜班）：真论文产物里 **12 个段落**的 `w:pPr` 把 `keepNext`
排在 `w:rPr` **之后**（schema 规定 `keepNext` 必须第 2 位）——Word 能容忍这种
顺序，但它是**不合规 XML**，且是「改了没生效」类故障的温床。

## 口径（用户拍板：「要修，一切问题都要从源头解决，千万不敢留屎山代码」）

① 本模块是顺序表的**唯一真源**。历史上同一张表在本仓有三份拷贝
   （`new_actions._PPR_SEQ` / `engine.writer._PPR_ORDER` 字典 / 各处 `_CT_*_ORDER`），
   三份之间已经出现**不一致**（writer 的字典缺 `suppressLineNumbers`/`rPr`/
   `sectPr`/`pPrChange` 四项）。其它模块一律 `from earlyz.tools.ooxml_order import ...`，
   **禁止再抄一份**。
② 新增代码往 pPr/rPr/tblPr/trPr 里塞子元素，一律用 `insert_ordered()`（有序插入）。
③ 收尾仍统一规范化一次（`canonicalize_root` / `canonicalize_docx`）：
   - 动作侧单一收口点 = `new_actions._write_part`（全仓唯一写盘函数）；
   - 编译链末尾 = `format_compiler.compile_document` 交付前。
   两层叠加保证**产物零违规**，且将来新增直写点也不会漏。

## 未知子元素怎么办

放末尾、保持相对顺序（Word 对未知元素宽容，放尾最安全）。`w:pPrChange`、
`w:rPrChange` 这类修订元素本就该在最后，落尾正好合规。
"""
from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# ── 唯一真源：四张顺序表（ECMA-376 CT_* sequence 顺序）────────────────────
PPR_SEQ = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd",
    "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
    "sectPr", "pPrChange",
]
RPR_SEQ = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
    "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
    "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
    "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
    "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
    "eastAsianLayout", "specVanish", "oMath",
]
TBLPR_SEQ = [
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
    "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
    "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
    "tblCaption", "tblDescription", "tblPrChange",
]
TRPR_SEQ = [
    "cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter",
    "cantSplit", "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden",
]

#: tag → 顺序表（收尾规范化与检查器共用）
SEQ_BY_TAG = {
    f"{W}pPr": PPR_SEQ,
    f"{W}rPr": RPR_SEQ,
    f"{W}tblPr": TBLPR_SEQ,
    f"{W}trPr": TRPR_SEQ,
}


def _localname(tag) -> str:
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


def insert_ordered(parent, el, seq: list[str]) -> None:
    """按 schema 序列把 el 插入 parent；el 不在表内则 append（落尾）。"""
    name = _localname(el.tag)
    try:
        idx = seq.index(name)
    except ValueError:
        parent.append(el)
        return
    for child in parent:
        cname = _localname(child.tag)
        if cname in seq and seq.index(cname) > idx:
            child.addprevious(el)
            return
    parent.append(el)


def canonicalize_element(elem, seq: list[str]) -> int:
    """按 seq 重排 elem 的直接子元素；返回 1 表示动过。未知子元素落尾。"""
    children = list(elem)
    if len(children) < 2:
        return 0
    idx_map = {name: i for i, name in enumerate(seq)}
    keyed = []
    for i, ch in enumerate(children):
        k = idx_map.get(_localname(ch.tag))
        keyed.append((k if k is not None else len(idx_map) + i, i, ch))
    keyed.sort(key=lambda t: (t[0], t[1]))
    ordered = [ch for _, _, ch in keyed]
    if ordered == children:
        return 0
    for ch in children:
        elem.remove(ch)
    for ch in ordered:
        elem.append(ch)
    return 1


def canonicalize_ppr_position(ppr) -> int:
    """`w:pPr` 必须是 `w:p` 的第一个子元素（否则 Word 静默忽略整段属性）。"""
    p = ppr.getparent()
    if p is None or _localname(p.tag) != "p":
        return 0
    if p.index(ppr) == 0:
        return 0
    p.remove(ppr)
    p.insert(0, ppr)
    return 1


def canonicalize_root(root) -> dict:
    """规范化一个 XML part 内所有 pPr/rPr/tblPr/trPr 的顺序。返回计数。"""
    stats = {"reordered": 0, "moved_ppr": 0}
    for elem in root.iter():
        tag = elem.tag
        seq = SEQ_BY_TAG.get(tag if isinstance(tag, str) else "")
        if seq is None:
            continue
        stats["reordered"] += canonicalize_element(elem, seq)
        if tag == f"{W}pPr":
            stats["moved_ppr"] += canonicalize_ppr_position(elem)
    return stats


def violations(root) -> list[dict]:
    """检查器：返回顺序违规清单（空 = 合规）。供测试与交付前门禁用。"""
    out: list[dict] = []
    for elem in root.iter():
        tag = elem.tag
        seq = SEQ_BY_TAG.get(tag if isinstance(tag, str) else "")
        if seq is None:
            continue
        names = [_localname(c.tag) for c in elem]
        known = [n for n in names if n in seq]
        want = sorted(known, key=seq.index)
        if known != want:
            out.append({"tag": _localname(tag), "got": names, "want_prefix": want})
        if tag == f"{W}pPr":
            p = elem.getparent()
            if p is not None and _localname(p.tag) == "p" and p.index(elem) != 0:
                out.append({"tag": "pPr_position", "got": f"index={p.index(elem)}"})
    return out


def canonicalize_docx(path) -> dict:
    """整包收尾规范化（所有 .xml part）。返回 {part: stats}。"""
    path = Path(path)
    tmp = path.with_suffix(path.suffix + ".canon")
    changed: dict = {}
    with zipfile.ZipFile(path) as zin:
        names = zin.namelist()
        payload = {n: zin.read(n) for n in names}
    for name in names:
        if not name.endswith(".xml"):
            continue
        try:
            root = etree.fromstring(payload[name])
        except etree.XMLSyntaxError:
            continue
        stats = canonicalize_root(root)
        if stats["reordered"] or stats["moved_ppr"]:
            payload[name] = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True)
            changed[name] = stats
    if not changed:
        return {}
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, payload[n])
    tmp.replace(path)
    return changed
