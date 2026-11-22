"""标题自动编号：引用闭合 + 双编号压制（2026-11-08 用户实锤）。

症状：期刊产物一级标题显示「1. 1. 引言」，二三级标题还各带 1.27cm 左缩进。

机制不是「取值算错」，是跨部件引用断了：
  ① 模板标题样式带 <w:numPr><w:numId w:val="7"/>（模板靠自动编号出号）；
  ② 样式被搬进产物，但模板 numbering.xml 里 numId=7 的定义**没跟过来**；
  ③ Word 遇到悬空 numId 不报错 → 按默认列表兜底：再画一个「1.」，并套上
     默认列表缩进 left=720 twips（=1.27cm，逐级累加）。

两步收口：
  · reconcile_numbering_refs          —— 样式引用的 numId 必须在 numbering.xml
    有定义（缺的从模板补齐；补齐后编号与缩进才是「模板那套」）；
  · suppress_duplicate_heading_numbers —— 文本里已经手打序号的标题（内容是内容，
    不许删），显式 numId=0 关掉自动编号，避免出现两个号，缩进随之消失。
"""
from __future__ import annotations

import copy
import re
from pathlib import Path

from lxml import etree

from .new_actions import (W, _cover_paragraphs, _para_full_text, _ppr_insert_ordered,
                          _read_part, _read_styles_map, _write_part)

# 「1.」「2.1」「3．」这类段首手打序号（后面必须跟空白，避免把「2026年」当编号）
# 序号前缀：三种形态都算（2026-11-08 实测漏判）
#   "2.1 基于优化" / "1.1编写目的"（多级，可无空格） / "1.  范围"（数字+空格） / "3．方法"
# 反例必须不匹配："2026年工作总结"（数字后直接跟汉字且无分隔）
_NUM_PREFIX_RE = re.compile(r"^\s*(?:\d+(?:\s*[.．]\s*\d+)+|\d+\s*[.．、]|\d+\s)\s*")


def _style_numbering_id(styles_root, style_id: str) -> str | None:
    """沿 basedOn 链找某样式生效的 numId（None=没有编号）。"""
    seen: set[str] = set()
    sid = style_id
    while sid and sid not in seen:
        seen.add(sid)
        st = None
        for s in styles_root.findall(f"{W}style"):
            if s.get(f"{W}styleId") == sid:
                st = s
                break
        if st is None:
            return None
        np = st.find(f"{W}pPr/{W}numPr")
        if np is not None:
            nid = np.find(f"{W}numId")
            if nid is not None:
                return nid.get(f"{W}val")
        based = st.find(f"{W}basedOn")
        sid = based.get(f"{W}val") if based is not None else None
    return None


def paragraph_numbering_id(p, styles_root) -> str | None:
    """段落生效的 numId：直接格式优先，其次样式（含 basedOn 链）。"""
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        return None
    np = ppr.find(f"{W}numPr")
    if np is not None:
        nid = np.find(f"{W}numId")
        if nid is not None:
            return nid.get(f"{W}val")
    st = ppr.find(f"{W}pStyle")
    if st is not None:
        return _style_numbering_id(styles_root, st.get(f"{W}val") or "")
    return None


def reconcile_numbering_refs(target_path: Path, params: dict) -> dict:
    """样式引用的 numId 必须在 numbering.xml 里解析得到（缺的从模板补）。

    params: {template: 模板 docx 路径}（可选；没有就只报告不补齐）
    """
    target_path = Path(target_path)
    styles_root = _read_part(target_path, "word/styles.xml")
    refs: set[str] = set()
    if styles_root is not None:
        for np in styles_root.iter(f"{W}numPr"):
            nid = np.find(f"{W}numId")
            if nid is not None:
                val = nid.get(f"{W}val")
                if val and val != "0":
                    refs.add(val)
    if not refs:
        return {"ok": True, "changed": False, "missing_before": [], "imported": 0}

    try:
        numbering = _read_part(target_path, "word/numbering.xml")
    except Exception:  # noqa: BLE001
        numbering = None
    if numbering is None:
        numbering = etree.Element(f"{W}numbering")
    have = {n.get(f"{W}numId") for n in numbering.findall(f"{W}num")}
    missing = sorted(refs - have, key=lambda x: (len(x), x))
    if not missing:
        return {"ok": True, "changed": False, "missing_before": [], "imported": 0}

    template = (params or {}).get("template") or ""
    imported = 0
    unresolved = list(missing)
    if template and Path(template).exists():
        try:
            tpl = _read_part(Path(template), "word/numbering.xml")
        except Exception:  # noqa: BLE001
            tpl = None
        if tpl is not None:
            tpl_nums = {n.get(f"{W}numId"): n for n in tpl.findall(f"{W}num")}
            abs_by_id = {a.get(f"{W}abstractNumId"): a for a in tpl.findall(f"{W}abstractNum")}
            have_abs = {a.get(f"{W}abstractNumId") for a in numbering.findall(f"{W}abstractNum")}
            for mid in list(unresolved):
                num = tpl_nums.get(mid)
                if num is None:
                    continue
                clone = copy.deepcopy(num)
                aid_el = clone.find(f"{W}abstractNumId")
                aid = aid_el.get(f"{W}val") if aid_el is not None else None
                if aid is None or (aid not in have_abs and aid not in abs_by_id):
                    continue                      # 模板缺 abstractNum → 无法解析
                if aid not in have_abs:
                    numbering.append(copy.deepcopy(abs_by_id[aid]))
                    have_abs.add(aid)
                numbering.append(clone)
                have.add(mid)
                imported += 1
                unresolved.remove(mid)
            if imported:
                _write_part(target_path, "word/numbering.xml", numbering)
    return {"ok": True, "changed": imported > 0, "missing_before": missing,
            "imported": imported, "unresolved": unresolved}


def verify_numbering_refs(target_path: Path, params: dict) -> bool:
    styles_root = _read_part(Path(target_path), "word/styles.xml")
    refs: set[str] = set()
    if styles_root is not None:
        for np in styles_root.iter(f"{W}numPr"):
            nid = np.find(f"{W}numId")
            if nid is not None and (nid.get(f"{W}val") or "0") != "0":
                refs.add(nid.get(f"{W}val"))
    if not refs:
        return True
    try:
        numbering = _read_part(Path(target_path), "word/numbering.xml")
    except Exception:  # noqa: BLE001
        return False
    if numbering is None:
        return False
    have = {n.get(f"{W}numId") for n in numbering.findall(f"{W}num")}
    return not (refs - have)


def suppress_duplicate_heading_numbers(target_path: Path, params: dict) -> dict:
    """标题段文本里已有手打序号 → 显式关掉样式继承的自动编号（numId=0）。

    文本属于内容，不许删；压掉的是「样式带来的那一层编号」，默认列表缩进随之消失。
    """
    target_path = Path(target_path)
    root = _read_part(target_path, "word/document.xml")
    styles_root = _read_part(target_path, "word/styles.xml")
    if root is None or styles_root is None:
        return {"ok": True, "changed": False, "suppressed": 0}
    body = root.find(f"{W}body")
    styles_map = _read_styles_map(target_path)
    cover = _cover_paragraphs(root, styles_map)
    suppressed = 0
    for p in root.iter(f"{W}p"):
        if p.getparent() is not body:
            continue
        if cover and any(p is c for c in cover):
            continue
        ppr = p.find(f"{W}pPr")
        if ppr is None:
            continue
        st = ppr.find(f"{W}pStyle")
        if st is not None and (st.get(f"{W}val") or "").upper().startswith("TOC"):
            continue
        txt = _para_full_text(p)
        if not txt or not _NUM_PREFIX_RE.match(txt):
            continue
        if paragraph_numbering_id(p, styles_root) in (None, "0"):
            continue
        np = ppr.find(f"{W}numPr")
        if np is None:
            np = etree.Element(f"{W}numPr")
            _ppr_insert_ordered(ppr, np)
        nid = np.find(f"{W}numId")
        if nid is None:
            nid = etree.SubElement(np, f"{W}numId")
        nid.set(f"{W}val", "0")
        suppressed += 1
    if suppressed:
        _write_part(target_path, "word/document.xml", root)
    return {"ok": True, "changed": suppressed > 0, "suppressed": suppressed}


def verify_no_duplicate_heading_numbers(target_path: Path, params: dict) -> bool:
    root = _read_part(Path(target_path), "word/document.xml")
    styles_root = _read_part(Path(target_path), "word/styles.xml")
    if root is None or styles_root is None:
        return True
    body = root.find(f"{W}body")
    for p in root.iter(f"{W}p"):
        if p.getparent() is not body:
            continue
        txt = _para_full_text(p)
        if not txt or not _NUM_PREFIX_RE.match(txt):
            continue
        if paragraph_numbering_id(p, styles_root) not in (None, "0"):
            return False
    return True
