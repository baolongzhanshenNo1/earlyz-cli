# -*- coding: utf-8 -*-
"""纯文本表格转换（2026-10-08 用户拍板）——粘贴/下载回来的表数据块 → Word 表格。

背景：用户文档里「表 5-1 仿真参数」这类表格常以纯文本形态贴进来（每行一段
文字），引擎此前只给题注套格式，数据区无处理（用户报「表格无法识别」）。

流程（提案→确认；2026-10-08 用户拍板原话：「先给行列方案过目，用户确认后
再落盘。需要用户确认的时候一定是先和用户确认好方案再进行排版，不能先斩
后奏」）：
- apply=false：扫描候选数据块（题注段 + 连续短行），出候选清单与行列方案
  预览（提供了 cells/columns 时给矩阵预览；未提供时列出候选行并提示给方案）；
- apply=true：**必须**带 match（题注）+ cells 或 columns（已确认的行列方案）
  才落盘——把选中候选的数据行段落替换为表格（题注保留原位）。表边框/单元格
  边距优先取模板首表（template_path），否则三线表默认。

边界：不发明内容——单元格文本逐字来自确认后的 cells（或数据行按确认列数
切分）；转换后原数据行删除（内容进表格）。幂等：转换完成的块不再成为候选。
"""

from __future__ import annotations

import copy
import re
from pathlib import Path

from lxml import etree

from earlyz.actions.new_actions import W, _para_full_text, _read_part, _write_part

_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_WS_SPLIT_RE = re.compile(r"[\s\u3000]+")
_CAPTION_RE = re.compile(
    r"^表\s*[\d一二三四五六七八九十]+(?:\s*[-–—.．]\s*\d+)?(?![\d])\s*\S")
_FIG_LINE_RE = re.compile(r"^[图]\s*\d")
# 标题形态（数字小节号/第X章/中文序号）——数据块收集到此为止（2026-10-08：
# 「5.2 收敛性分析」曾被吃进「表 5-1」的数据块，导致 25 行无法整齐切分）
_HEADING_LIKE_RE = re.compile(
    r"^(?:\d+[.．]\d+(?:[.．]\d+)*|第\s*[一二三四五六七八九十百\d]+\s*[章节])"
    r"\s*[\u4e00-\u9fff]"
    r"|^[一二三四五六七八九十]{1,3}\s*[、．]")
_STOP_CHARS = "。；！？!?;"
MIN_LINES = 6
MAX_LINE_LEN = 60


def _ptext(p) -> str:
    return _para_full_text(p).strip()


def _is_data_line(t: str) -> bool:
    if not t or len(t) > MAX_LINE_LEN:
        return False
    if t[-1] in _STOP_CHARS:
        return False
    return True


def _candidates(root) -> list[dict]:
    """候选数据块：题注段（「表 N-M …」）+ 紧随其后 ≥6 行短数据行。

    返回按文档序的候选列表；坐标 = body 直接子段（w:p）列表下标。
    """
    body = root.find(f"{W}body")
    if body is None:
        return []
    paras = [el for el in body if etree.QName(el).localname == "p"]
    out: list[dict] = []
    i = 0
    while i < len(paras):
        t = _ptext(paras[i])
        if not _CAPTION_RE.match(t):
            i += 1
            continue
        j = i + 1
        skipped = 0
        while j < len(paras) and not _ptext(paras[j]) and skipped < 2:
            j += 1
            skipped += 1
        lines: list[tuple[int, str]] = []
        k = j
        while k < len(paras):
            tt = _ptext(paras[k])
            if not _is_data_line(tt):
                break
            if _CAPTION_RE.match(tt) or _FIG_LINE_RE.match(tt):
                break
            if _HEADING_LIKE_RE.match(tt):
                break
            lines.append((k, tt))
            k += 1
        digits = sum(1 for _, tt in lines if re.search(r"\d", tt))
        if len(lines) >= MIN_LINES and digits >= 2:
            out.append({"caption": t, "caption_index": i,
                        "lines": [tt for _, tt in lines],
                        "line_indexes": [k2 for k2, _ in lines],
                        "start": j, "end": k - 1})
            i = k
        else:
            i += 1
    return out


def _resolve_matrix(lines: list[str], params: dict):
    """按确认后的行列方案出矩阵；不可整齐切分返回 None（不猜）。"""
    cells = params.get("cells")
    if cells:
        try:
            if isinstance(cells[0], (list, tuple)):
                m = [[str(x) for x in row] for row in cells]
            else:
                cols = int(params.get("columns") or 0)
                if cols <= 0 or len(cells) % cols:
                    return None
                flat = [str(x) for x in cells]
                m = [flat[i:i + cols] for i in range(0, len(flat), cols)]
        except Exception:  # noqa: BLE001
            return None
        return m if m and all(m) else None
    cols = int(params.get("columns") or 0)
    if cols >= 2 and len(lines) % cols == 0:
        return [lines[i:i + cols] for i in range(0, len(lines), cols)]
    return None


def _template_table_props(template_path: str | None):
    """从模板首表取 tblBorders / tblCellMar 的 XML 副本（无则 None）。"""
    if not template_path:
        return None, None
    try:
        tpath = Path(template_path)
        if not tpath.exists():
            return None, None
        troot = _read_part(tpath, "word/document.xml")
        for tbl in troot.iter(f"{W}tbl"):
            tbl_pr = tbl.find(f"{W}tblPr")
            if tbl_pr is None:
                continue
            tb = tbl_pr.find(f"{W}tblBorders")
            cm = tbl_pr.find(f"{W}tblCellMar")
            return (copy.deepcopy(tb) if tb is not None else None,
                    copy.deepcopy(cm) if cm is not None else None)
    except Exception:  # noqa: BLE001 — 模板读取失败回退默认三线表
        pass
    return None, None


def _build_table(matrix: list[list[str]], *, borders_xml=None, cellmar_xml=None,
                 header: bool = True):
    """构建 w:tbl（默认三线表：顶/底 1.5pt + 表头下 0.75pt，无竖线）。"""
    ncols = max(len(r) for r in matrix)
    tbl = etree.Element(f"{W}tbl")
    tbl_pr = etree.SubElement(tbl, f"{W}tblPr")
    tbl_w = etree.SubElement(tbl_pr, f"{W}tblW")
    tbl_w.set(f"{W}w", "0")
    tbl_w.set(f"{W}type", "auto")
    jc = etree.SubElement(tbl_pr, f"{W}jc")
    jc.set(f"{W}val", "center")
    if borders_xml is not None:
        tbl_pr.append(borders_xml)
    else:
        tb = etree.SubElement(tbl_pr, f"{W}tblBorders")
        for side, val, sz in (("top", "single", "12"), ("bottom", "single", "12"),
                              ("left", "none", "0"), ("right", "none", "0"),
                              ("insideH", "none", "0"), ("insideV", "none", "0")):
            el = etree.SubElement(tb, f"{W}{side}")
            el.set(f"{W}val", val)
            el.set(f"{W}sz", sz)
            el.set(f"{W}space", "0")
            el.set(f"{W}color", "auto")
    if cellmar_xml is not None:
        tbl_pr.append(cellmar_xml)
    grid = etree.SubElement(tbl, f"{W}tblGrid")
    for _ in range(ncols):
        etree.SubElement(grid, f"{W}gridCol")   # 缺 w → 自动布局
    for ri, row in enumerate(matrix):
        tr = etree.SubElement(tbl, f"{W}tr")
        tr_pr = etree.SubElement(tr, f"{W}trPr")
        if header and ri == 0:
            th = etree.SubElement(tr_pr, f"{W}tblHeader")
            th.set(f"{W}val", "true")
        cant = etree.SubElement(tr_pr, f"{W}cantSplit")
        cant.set(f"{W}val", "true")
        for ci in range(ncols):
            txt = row[ci] if ci < len(row) else ""
            tc = etree.SubElement(tr, f"{W}tc")
            if header and ri == 0:
                tc_pr = etree.SubElement(tc, f"{W}tcPr")
                tcb = etree.SubElement(tc_pr, f"{W}tcBorders")
                btm = etree.SubElement(tcb, f"{W}bottom")
                btm.set(f"{W}val", "single")
                btm.set(f"{W}sz", "6")
                btm.set(f"{W}space", "0")
                btm.set(f"{W}color", "auto")
            p = etree.SubElement(tc, f"{W}p")
            ppr = etree.SubElement(p, f"{W}pPr")
            ind = etree.SubElement(ppr, f"{W}ind")
            ind.set(f"{W}firstLine", "0")
            ind.set(f"{W}firstLineChars", "0")
            pjc = etree.SubElement(ppr, f"{W}jc")
            pjc.set(f"{W}val", "center")
            r = etree.SubElement(p, f"{W}r")
            if header and ri == 0:
                rpr = etree.SubElement(r, f"{W}rPr")
                etree.SubElement(rpr, f"{W}b")
            t = etree.SubElement(r, f"{W}t")
            t.set(_XML_SPACE, "preserve")
            t.text = txt
    return tbl


def scan_text_table_candidates(path) -> dict:
    """只读扫描：疑似表格数据块数量与首个候选摘要（交付成功后的提示用）。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        cands = _candidates(root)
        return {"count": len(cands),
                "first_caption": (cands[0]["caption"] if cands else None),
                "first_lines": (len(cands[0]["lines"]) if cands else 0)}
    except Exception:  # noqa: BLE001
        return {"count": 0, "first_caption": None, "first_lines": 0}


def convert_text_table(path, params: dict | None = None) -> dict:
    """纯文本数据块 → 表格（提案→确认→落盘；绝不先斩后奏）。

    params:
      apply: false（默认，提案）/ true（落盘）
      match: 题注文本（包含匹配；apply=true 必填——指定要转换的数据块）
      cells: 确认后的单元格（行列表 [[…], …] 或配合 columns 的平铺列表）
      columns: 列数（与 cells 平铺配合，或对数据行按位置切分）
      header: 首行是否表头（默认 true）
      template_path: 可选——模板 docx（取首表边框/单元格边距）
    """
    path = Path(path)
    params = params or {}
    apply_ = bool(params.get("apply"))
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False, "applied": False, "reason": "no_body"}
    cands = _candidates(root)

    if not apply_:
        if not cands:
            return {"changed": False, "applied": False, "reason": "no_candidates",
                    "candidates": [],
                    "message": "未发现「表 N-M 标题 + 连续短数据行」形态的文本数据块。"}
        preview = []
        for c in cands[:10]:
            m = _resolve_matrix(c["lines"], params)
            item = {"caption": c["caption"], "line_count": len(c["lines"]),
                    "lines": c["lines"][:30]}
            if m is not None:
                item["matrix"] = m
                item["rows"] = len(m)
                item["cols"] = max(len(r) for r in m)
            else:
                item["matrix"] = None
                item["hint"] = (
                    "未提供行列方案——请核对数据行后给出列数/单元格划分再转换"
                    if not params.get("cells") and not params.get("columns")
                    else "按当前 cells/columns 无法整齐切分，请核对")
            preview.append(item)
        return {"changed": False, "applied": False, "candidates": preview,
                "count": len(cands),
                "message": "识别到疑似表格数据块（未落盘）。请核对行列方案（matrix 预览），"
                           "确认后带 match + cells（或 columns）再 apply=true 转换。"}

    # ── apply=true：必须先有确认后的方案，不猜 ──
    match = str(params.get("match") or "")
    if not match:
        return {"changed": False, "applied": False, "reason": "need_match",
                "candidates": [c["caption"] for c in cands][:10],
                "message": "apply=true 需要 match（题注文本）指定数据块；候选见 candidates。"}
    if not params.get("cells") and not params.get("columns"):
        return {"changed": False, "applied": False, "reason": "need_plan",
                "lines": (next((c["lines"] for c in cands if match in c["caption"]),
                               [])[:30]),
                "message": "apply=true 需要 cells（确认后的单元格）或 columns（确认后的"
                           "列数）——行列方案须先与用户确认，禁止先斩后奏。"}
    cand = next((c for c in cands if match in c["caption"]), None)
    if cand is None:
        return {"changed": False, "applied": False, "reason": "match_not_found",
                "candidates": [c["caption"] for c in cands][:10],
                "message": "未找到匹配该题注的数据块；候选见 candidates。"}
    matrix = _resolve_matrix(cand["lines"], params)
    if matrix is None:
        return {"changed": False, "applied": False, "reason": "bad_plan",
                "lines": cand["lines"][:30],
                "message": "按当前 cells/columns 无法整齐切分数据行，请核对行列方案"
                           "（列数需整除行数；cells 平铺长度需为列数整数倍）。"}

    paras = [el for el in body if etree.QName(el).localname == "p"]
    cap_el = paras[cand["caption_index"]]
    line_els = [paras[k] for k in cand["line_indexes"]]
    parent = cap_el.getparent()
    borders_xml, cellmar_xml = _template_table_props(params.get("template_path"))
    tbl = _build_table(matrix, borders_xml=borders_xml, cellmar_xml=cellmar_xml,
                       header=bool(params.get("header", True)))
    for el in line_els:
        if el.getparent() is not None:
            el.getparent().remove(el)
    parent.insert(list(parent).index(cap_el) + 1, tbl)
    _write_part(path, "word/document.xml", root)
    result = {"changed": True, "applied": True, "caption": cand["caption"],
              "rows": len(matrix), "cols": max(len(r) for r in matrix),
              "cells_total": sum(len(r) for r in matrix),
              "message": f"已把「{cand['caption']}」下的 {len(cand['lines'])} 行数据"
                         f"转换为 {len(matrix)} 行 × {max(len(r) for r in matrix)} 列表格。"}
    try:
        result["verified"] = verify_text_table(path, params)
    except Exception:  # noqa: BLE001
        result["verified"] = None
    return result


def verify_text_table(path, params: dict | None = None) -> bool:
    """复查：题注段之后存在表格（新表落位）；apply=false 提案阶段恒真。"""
    try:
        params = params or {}
        if not params.get("apply"):
            return True
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return False
        match = str(params.get("match") or "")
        kids = [el for el in body]
        if not match:
            return any(etree.QName(el).localname == "tbl" for el in kids)
        for i, el in enumerate(kids):
            if etree.QName(el).localname != "p":
                continue
            if match not in _ptext(el):
                continue
            for nxt in kids[i + 1:i + 3]:
                if etree.QName(nxt).localname == "tbl":
                    return True
            return False
        return False
    except Exception:  # noqa: BLE001
        return False
