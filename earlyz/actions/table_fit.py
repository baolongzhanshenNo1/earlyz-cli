"""表格宽度收进版心（2026-11-07）。

**为什么**：模板/源稿里的表格常按 A4 全宽（9060–9746 twips）排版，而目标文档的
版心只有 8306–9070 twips（学位论文页边距 2.5cm 起）。产物照样搬过去，表格就压到
右边距外面——评分器 `table` 项直接判失败（学位论文(word) 4 张、重庆科技 6 张、
北航 1 张）。

**口径**：只收「显式宽度（tblW type=dxa）且超出版心」的表；按比例缩放 `gridCol`
与每行 `tcW`，保证各行列宽之和 = 版心宽。`type=auto`（自动适应窗口）本来就不会
溢出，不动。不增删列、不改内容。
"""
from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _read(path: Path) -> etree._Element | None:
    try:
        with zipfile.ZipFile(path) as z:
            return etree.fromstring(z.read("word/document.xml"))
    except (KeyError, zipfile.BadZipFile, OSError):
        return None


def _write(path: Path, root: etree._Element) -> None:
    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = path.with_suffix(".tbltmp")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            payload = data if item.filename == "word/document.xml" else zin.read(item.filename)
            zout.writestr(item, payload)
    tmp.replace(path)


def _usable_width(sect) -> int | None:
    sz = sect.find(f"{W}pgSz") if sect is not None else None
    mar = sect.find(f"{W}pgMar") if sect is not None else None
    if sz is None or mar is None:
        return None
    try:
        w = int(sz.get(f"{W}w"))
        left = int(mar.get(f"{W}left") or 0)
        right = int(mar.get(f"{W}right") or 0)
    except (TypeError, ValueError):
        return None
    return w - left - right or None


def _section_for(body, tbl):
    """表所属节：表之后第一个 sectPr（含段落 pPr 内的），否则 body 末节。"""
    kids = list(body)
    try:
        idx = kids.index(tbl)
    except ValueError:
        idx = -1
    for el in kids[idx + 1:]:
        tag = el.tag
        if tag == f"{W}sectPr":
            return el
        if tag == f"{W}p":
            ppr = el.find(f"{W}pPr")
            if ppr is not None and ppr.find(f"{W}sectPr") is not None:
                return ppr.find(f"{W}sectPr")
    return body.find(f"{W}sectPr")


def _scale(vals: list[int], total: int) -> list[int]:
    """按比例缩放到 total，误差补到最后一列（保证和恰好等于 total）。"""
    s = sum(vals)
    if s <= 0 or s <= total:
        return vals
    out = [max(1, round(v * total / s)) for v in vals]
    diff = total - sum(out)
    out[-1] = max(1, out[-1] + diff)
    return out


def fit_tables_to_page(target_path) -> dict:
    info: dict = {"changed": False, "tables": 0, "shrunk": 0, "errors": []}
    try:
        path = Path(target_path)
        root = _read(path)
        if root is None:
            info["errors"].append("document.xml unreadable")
            return info
        body = root.find(f"{W}body")
        if body is None:
            info["errors"].append("no body")
            return info
        tables = list(root.iter(f"{W}tbl"))
        info["tables"] = len(tables)
        for tbl in tables:
            cap = _usable_width(_section_for(body, tbl))
            if not cap:
                continue
            tp = tbl.find(f"{W}tblPr")
            tw = tp.find(f"{W}tblW") if tp is not None else None
            if tw is None:
                continue
            typ = tw.get(f"{W}type")
            cur = tw.get(f"{W}w")
            if typ == "auto" or not (cur or "").isdigit():
                continue
            cur_i = int(cur)
            if cur_i <= cap:
                continue
            tw.set(f"{W}w", str(cap))
            grid = tbl.find(f"{W}tblGrid")
            if grid is not None:
                cols = [int(gc.get(f"{W}w") or 0) for gc in grid.findall(f"{W}gridCol")]
                if cols and sum(cols) > cap:
                    for gc, nv in zip(grid.findall(f"{W}gridCol"), _scale(cols, cap)):
                        gc.set(f"{W}w", str(nv))
            # 每行单元格宽度同比缩放（只处理 dxa 且列数与网格一致的简单表）
            for tr in tbl.findall(f"{W}tr"):
                cells = tr.findall(f"{W}tc")
                widths = []
                for tc in cells:
                    tcpr = tc.find(f"{W}tcPr")
                    tcw = tcpr.find(f"{W}tcW") if tcpr is not None else None
                    if tcw is None or (tcw.get(f"{W}type") or "dxa") != "dxa":
                        widths = []
                        break
                    v = tcw.get(f"{W}w")
                    if not (v or "").isdigit():
                        widths = []
                        break
                    widths.append(int(v))
                if not widths or sum(widths) <= cap:
                    continue
                for tc, nv in zip(cells, _scale(widths, cap)):
                    tc.find(f"{W}tcPr").find(f"{W}tcW").set(f"{W}w", str(nv))
            info["shrunk"] += 1
        if info["shrunk"]:
            _write(path, root)
            info["changed"] = True
    except Exception as e:  # noqa: BLE001 — 非阻断
        info["errors"].append(f"{type(e).__name__}: {e}")
    return info
