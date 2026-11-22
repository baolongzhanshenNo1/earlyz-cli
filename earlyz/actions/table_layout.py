# -*- coding: utf-8 -*-
"""表格跨页布局的**纯函数层**（2026-09-26 定案，见 docs/表格跨页处理方案.md）。

设计约束：
  · 全部为纯函数（输入 XML 元素 → 判定/数值），便于单测与「反例回灌」；
  · 不做任何写入 —— 写盘只发生在 action 层（new_actions.keep_table_together）；
  · 估高按 Word 排版模型折算，并能被 scripts/calibrate_table_height.py 校准。

Word 事实（必须记住，否则会写错）：
  · 没有「表格级-不跨页」属性；整表同页只能靠段级 keepNext 链（相邻行锁同页）；
  · 行级 cantSplit 只防**行内**切断，防不住**行间**断开；
  · sectPr 归属其**之前**的内容 → 表的所属节 = 表之后第一个 sectPr。

题注方向（EZI-0011 定案，落地在 ``earlyz.actions.captions.bind_caption_objects``）：
  · 中文规范：**图题在图下方、表题在表上方**；
  · keepNext 是**单向**的（绑「本段 + 后面内容」），所以两种题注方向相反：
    表题在上 → keepNext 挂**表题段**（题绑表）；图题在下 → keepNext 挂**图片段落**
    （图绑题）。把图题也按表题处理＝图题段绑到后面的正文，图与题之间没有绑定；
  · 题与对象之间夹空段时 keepNext 链断在空段上 → 绑之前先拆纯空段（只拆「仅
    含 pPr/空 run」的段，有字/图/书签/域/分页符/sectPr 的一律不动）；
  · cantSplit / tblHeader 与本条互不相干，绑题注时**一个都不许动**（A4）。
"""
from __future__ import annotations

import re

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# 折算常量
EMU_PER_TWIP = 635          # 1 twip = 635 EMU（图片 extent 换算）
CJK_CHAR_TWIPS_12PT = 432   # 12pt CJK 单字宽 ≈ 432 twips（字号 half-pt × 36）
CELL_PADDING_TWIPS = 100    # 单元格上下内边距估算
DEFAULT_LINE_TWIPS = 360    # 无任何行距信息时的兜底行高（≈12pt 1.5 倍）

FORM_SIGNAL_THRESHOLD = 2   # 命中几条信号判为表单/块状表
FORM_ROW_HEIGHT_MIN = 1000  # 「显式行高」信号门槛（twips ≈1.8cm）：数据表常见
                            # 200~500 的窄行高，表单表则是被撑高的空白行（实测 1395/5895/2640）
HEADER_MIN_ROWS = 4         # 表头识别：表至少这么多行
HEADER_MAX_CHARS = 12       # 表头识别：每格文字上限


# ── 节 / 可用页高 ────────────────────────────────────────────────

def _localname(el) -> str:
    """元素本地名 —— 同时兼容 lxml 与 stdlib ElementTree（校准脚本用 stdlib 解析）。"""
    tag = getattr(el, "tag", None)
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


def section_of_table(body, tbl):
    """表所属节的 sectPr：表之后第一个 sectPr（含段落 pPr 内的）。"""
    kids = list(body)
    try:
        idx = kids.index(tbl)
    except ValueError:
        idx = -1
    for el in kids[idx + 1:]:
        if _localname(el) == "sectPr":
            return el
        if _localname(el) == "p":
            ppr = el.find(f"{W}pPr")
            if ppr is not None and ppr.find(f"{W}sectPr") is not None:
                return ppr.find(f"{W}sectPr")
    return body.find(f"{W}sectPr")


def usable_page_height(body, tbl) -> int:
    """可用正文高（twips）= pgSz.h − pgMar.top − pgMar.bottom。"""
    sect = section_of_table(body, tbl)
    if sect is None:
        return 13958                                    # A4 + 上下 2.54cm 兜底
    pgsz = sect.find(f"{W}pgSz")
    pgmar = sect.find(f"{W}pgMar")
    h, top, bottom = 16838, 1440, 1440
    try:
        if pgsz is not None and pgsz.get(f"{W}h"):
            h = int(pgsz.get(f"{W}h"))
        if pgmar is not None:
            if pgmar.get(f"{W}top"):
                top = int(pgmar.get(f"{W}top"))
            if pgmar.get(f"{W}bottom"):
                bottom = int(pgmar.get(f"{W}bottom"))
    except (TypeError, ValueError):
        pass
    return max(1440, h - top - bottom)


# ── 结构读取 ─────────────────────────────────────────────────────

def table_grid(tbl) -> list[int]:
    """tblGrid 列宽（twips）；缺失时按首行 tcW 兜底。"""
    grid = tbl.find(f"{W}tblGrid")
    cols = []
    if grid is not None:
        for gc in grid.findall(f"{W}gridCol"):
            try:
                cols.append(int(gc.get(f"{W}w") or 0))
            except (TypeError, ValueError):
                cols.append(0)
    if not cols:
        first = tbl.find(f"{W}tr")
        if first is not None:
            for tc in first.findall(f"{W}tc"):
                tcpr = tc.find(f"{W}tcPr")
                tcw = tcpr.find(f"{W}tcW") if tcpr is not None else None
                try:
                    cols.append(int(tcw.get(f"{W}w") or 0) if tcw is not None else 0)
                except (TypeError, ValueError):
                    cols.append(0)
    return cols


def row_cells(tr):
    """行的格与「该格覆盖的列数」（gridSpan）与「是否 vMerge 续行」。"""
    out = []
    for tc in tr.findall(f"{W}tc"):
        tcpr = tc.find(f"{W}tcPr")
        span = 1
        vmerge_continue = False
        if tcpr is not None:
            gs = tcpr.find(f"{W}gridSpan")
            if gs is not None and gs.get(f"{W}val"):
                try:
                    span = max(1, int(gs.get(f"{W}val")))
                except ValueError:
                    span = 1
            vm = tcpr.find(f"{W}vMerge")
            if vm is not None and (vm.get(f"{W}val") or "continue") == "continue":
                vmerge_continue = True
        out.append((tc, span, vmerge_continue))
    return out


def cell_width_twips(cols: list[int], cells, index: int) -> int:
    """第 index 个格的有效宽（含 gridSpan 累加列宽）。"""
    start = sum(span for _, span, _ in cells[:index])
    span = cells[index][1]
    if cols:
        seg = cols[start:start + span]
        if seg and sum(seg) > 0:
            return sum(seg)
        return cols[0] if cols else 0
    return 0


def cell_text(tc) -> str:
    return "".join(t.text or "" for t in tc.iter(f"{W}t"))


def cell_image_twips(tc) -> int:
    """格内图片高度（twips）：取格内最大的 extent（形状的 a:ext 与 wp:extent 不叠加）。"""
    tallest = 0
    for el in tc.iter():
        if not isinstance(el.tag, str) or not el.tag.endswith("}extent"):
            continue
        try:
            cy = int(el.get("cy") or 0)
        except (TypeError, ValueError):
            continue
        tallest = max(tallest, cy // EMU_PER_TWIP)
    return tallest


def _spacing_line_twips(pPr) -> int | None:
    """从 pPr 里读 exact/atLeast 行距（twips）。"""
    if pPr is None:
        return None
    sp = pPr.find(f"{W}spacing")
    if sp is None or sp.get(f"{W}lineRule") not in ("exact", "atLeast"):
        return None
    try:
        return int(sp.get(f"{W}line"))
    except (TypeError, ValueError):
        return None


def default_doc_line_twips(styles_root) -> int:
    """文档默认行距（twips）：docDefaults → 默认段落样式 → Normal 样式 → 兜底。

    表格单元格通常不写自己的行距，继承文档默认 —— 用错默认值会把估高算偏
    （实测该模板 460 twips，兜底 360 会低估 22%），直接影响灰区判定。
    """
    if styles_root is None:
        return DEFAULT_LINE_TWIPS
    # 1) docDefaults/pPrDefault/pPr
    for dd in styles_root.findall(f"{W}docDefaults"):
        for pd in dd.findall(f"{W}pPrDefault"):
            v = _spacing_line_twips(pd.find(f"{W}pPr"))
            if v:
                return v
    # 2) type=paragraph 且 default=1 的样式
    normal = None
    for st in styles_root.findall(f"{W}style"):
        if st.get(f"{W}type") != "paragraph":
            continue
        sid = st.get(f"{W}styleId") or ""
        if st.get(f"{W}default") in ("1", "true"):
            v = _spacing_line_twips(st.find(f"{W}pPr"))
            if v:
                return v
        if sid.lower() in ("normal", "a", "正文"):
            # 不能写 `normal or st`：lxml 元素现在会对空元素做真值判断并告警，
            # 将来恒为 True（空 <w:style/> 会被当成有效 Normal，行距取到空值）。
            if normal is None:
                normal = st
    # 3) Normal 样式
    if normal is not None:
        v = _spacing_line_twips(normal.find(f"{W}pPr"))
        if v:
            return v
    return DEFAULT_LINE_TWIPS


# ── 估高（Word 排版模型） ────────────────────────────────────────

def build_style_map(styles_root) -> dict:
    """styles.xml → {styleId: {line, lineRule, sz, basedOn}}，供解析继承链。

    单元格段落通常不写自己的行距/字号，全靠样式继承；不解析样式的估高会
    严重走偏（真机实测：不解析把 9930 估成 19075，直接误判成「放不下」）。
    """
    out: dict = {}
    if styles_root is None:
        return out
    for st in styles_root.findall(f"{W}style"):
        sid = st.get(f"{W}styleId")
        if not sid:
            continue
        rec: dict = {"line": None, "lineRule": None, "sz": None,
                     "basedOn": st.get(f"{W}basedOn")}
        v = _spacing_line_twips(st.find(f"{W}pPr"))
        if v:
            rec["line"] = v
            sp = st.find(f"{W}pPr").find(f"{W}spacing")
            rec["lineRule"] = sp.get(f"{W}lineRule")
        if rec["line"] is None:
            sp = st.find(f"{W}pPr")
            sp = sp.find(f"{W}spacing") if sp is not None else None
            if sp is not None and sp.get(f"{W}line"):
                try:                                # auto 行距（倍数）也记下来
                    rec["line"] = int(sp.get(f"{W}line"))
                    rec["lineRule"] = sp.get(f"{W}lineRule")
                except (TypeError, ValueError):
                    pass
        rpr = st.find(f"{W}rPr")
        if rpr is not None:
            for holder in (rpr, st.find(f"{W}pPr").find(f"{W}rPr")
                           if st.find(f"{W}pPr") is not None else None):
                if holder is None:
                    continue
                s = holder.find(f"{W}sz")
                if s is not None and s.get(f"{W}val"):
                    try:
                        rec["sz"] = int(s.get(f"{W}val"))
                        break
                    except ValueError:
                        pass
        out[sid] = rec
    return out


def _style_chain(style_map: dict, style_id) -> list:
    """样式继承链（自身 → basedOn…），最多 3 层，防环。"""
    chain = []
    seen = set()
    sid = style_id
    while sid and sid not in seen and len(chain) < 4:
        seen.add(sid)
        rec = style_map.get(sid)
        if rec is None:
            break
        chain.append(rec)
        sid = rec.get("basedOn")
    return chain


def resolve_line_twips(p, default_line: int, style_map: dict | None = None) -> int:
    """段落有效行高（twips）：直系 spacing → 样式链 → 文档默认 → 兜底。"""
    ppr = p.find(f"{W}pPr")
    own = _spacing_line_twips(ppr)
    if own:
        return own
    style_id = None
    if ppr is not None:
        ps = ppr.find(f"{W}pStyle")
        if ps is not None:
            style_id = ps.get(f"{W}val")
    for rec in _style_chain(style_map or {}, style_id):
        if rec.get("lineRule") in ("exact", "atLeast") and rec.get("line"):
            return int(rec["line"])
    sz = resolve_font_halfpt(p, style_map)
    if style_map:
        for rec in _style_chain(style_map, style_id):
            if rec.get("line") and rec.get("lineRule") in ("auto", None):
                try:                                    # auto 行距：值/240 = 倍数
                    mult = max(1.0, int(rec["line"]) / 240.0)
                    return max(default_line, int(sz * 15 * mult))
                except (TypeError, ValueError):
                    pass
    return max(default_line, int(sz * 15))


def resolve_font_halfpt(p, style_map: dict | None = None) -> int:
    """段落有效字号（半磅）：run → pPr.rPr → 样式链 → 24（12pt）。"""
    sz = None
    ppr = p.find(f"{W}pPr")
    if ppr is not None:
        rpr = ppr.find(f"{W}rPr")
        if rpr is not None:
            s = rpr.find(f"{W}sz")
            if s is not None and s.get(f"{W}val"):
                try:
                    sz = int(s.get(f"{W}val"))
                except ValueError:
                    pass
    for r in p.iter(f"{W}r"):
        rpr = r.find(f"{W}rPr")
        if rpr is not None:
            s = rpr.find(f"{W}sz")
            if s is not None and s.get(f"{W}val"):
                try:
                    sz = max(sz or 0, int(s.get(f"{W}val")))
                except ValueError:
                    pass
    if sz:
        return sz
    style_id = None
    if ppr is not None:
        ps = ppr.find(f"{W}pStyle")
        if ps is not None:
            style_id = ps.get(f"{W}val")
    for rec in _style_chain(style_map or {}, style_id):
        if rec.get("sz"):
            return int(rec["sz"])
    return 24


def paragraph_line_twips(p, default_line: int, style_map: dict | None = None) -> int:
    """兼容旧调用：段落行高（twips）。"""
    return resolve_line_twips(p, default_line, style_map)


def chars_per_line(width_twips: int, line_twips: int) -> int:
    """按格宽与行高折算每行可容纳的 CJK 字数（保守：按 12pt 字宽基准）。"""
    if width_twips <= 0:
        return 1
    return max(1, width_twips // CJK_CHAR_TWIPS_12PT)


def estimate_row_height(tr, cols: list[int], default_line: int,
                        style_map: dict | None = None) -> int:
    """单行高（twips）：trHeight 与内容估算取大者（Word 的 atLeast 语义）。

    逐段计行数（**不能**用整格文字长度 × 段数，否则长格被重复计高）。
    """
    trpr = tr.find(f"{W}trPr")
    declared = 0
    if trpr is not None:
        th = trpr.find(f"{W}trHeight")
        if th is not None and th.get(f"{W}val"):
            try:
                declared = int(th.get(f"{W}val"))
            except ValueError:
                declared = 0
    cells = row_cells(tr)
    tallest = 0
    for i, (tc, _span, vmerge_continue) in enumerate(cells):
        if vmerge_continue:                             # 纵向合并的续行不重复计高
            continue
        width = cell_width_twips(cols, cells, i)
        cell_h = CELL_PADDING_TWIPS
        for p in tc.findall(f"{W}p"):
            ph = resolve_line_twips(p, default_line, style_map)
            text = "".join(t.text or "" for t in p.iter(f"{W}t"))
            if text.strip():
                cpl = chars_per_line(width, ph)
                ph = ph * max(1, -(-len(text) // cpl))   # ceil
            cell_h += ph
        cell_h = max(cell_h, cell_image_twips(tc) + CELL_PADDING_TWIPS)
        tallest = max(tallest, cell_h)
    return max(declared, tallest)


def estimate_table_height(tbl, default_line: int = DEFAULT_LINE_TWIPS,
                          style_map: dict | None = None) -> int:
    cols = table_grid(tbl)
    return sum(
        estimate_row_height(tr, cols, default_line, style_map)
        for tr in tbl.findall(f"{W}tr")
    )


# ── 分类器 / 表头识别 ────────────────────────────────────────────

def header_like_signals(tbl) -> dict:
    """首行「像表头」的**原始**判据（不含表单表排除，避免与分类器循环）。"""
    trs = tbl.findall(f"{W}tr")
    if len(trs) < HEADER_MIN_ROWS:
        return {"ok": False, "why": f"rows<{HEADER_MIN_ROWS}"}
    first = trs[0]
    first_cells = row_cells(first)
    texts = [cell_text(tc).strip() for tc, _s, _v in first_cells]
    if not texts or any(not t for t in texts):
        return {"ok": False, "why": "first row has empty cell"}
    if any(len(t) > HEADER_MAX_CHARS for t in texts):
        return {"ok": False, "why": f"cell text >{HEADER_MAX_CHARS} chars"}
    second_cells = row_cells(trs[1])
    n_first = sum(s for _t, s, _v in first_cells)
    n_second = sum(s for _t, s, _v in second_cells)
    if n_first != n_second:
        return {"ok": False, "why": "first/second row column count differs"}
    return {"ok": True, "why": ""}


def table_signals(tbl, default_line: int = DEFAULT_LINE_TWIPS) -> dict:
    """表单/块状表信号集合（每条独立可解释，便于反例回灌）。"""
    trs = tbl.findall(f"{W}tr")
    if not trs:
        return {"signals": [], "detail": {}}
    rows_with_height = 0
    empty_cells = 0
    total_cells = 0
    merged = 0
    for tr in trs:
        trpr = tr.find(f"{W}trPr")
        if trpr is not None:
            th = trpr.find(f"{W}trHeight")
            if th is not None and th.get(f"{W}val"):
                try:
                    if int(th.get(f"{W}val")) >= FORM_ROW_HEIGHT_MIN:
                        rows_with_height += 1
                except ValueError:
                    pass
        for tc, span, vmerge in row_cells(tr):
            total_cells += 1
            if span > 1 or vmerge:
                merged += 1
            if not cell_text(tc).strip():
                empty_cells += 1
    detail = {
        "rows": len(trs),
        "height_ratio": rows_with_height / len(trs),
        "empty_ratio": (empty_cells / total_cells) if total_cells else 0.0,
        "merged_cells": merged,
        "header_like_raw": header_like_signals(tbl)["ok"],
    }
    signals = []
    if detail["height_ratio"] >= 0.5:
        signals.append("explicit_row_heights")
    if detail["empty_ratio"] >= 0.5:
        signals.append("mostly_empty_cells")
    if detail["rows"] <= 6:
        signals.append("few_rows")
    if not detail["header_like_raw"]:
        signals.append("first_row_not_header_like")
    if merged > 0:
        signals.append("has_merged_cells")
    return {"signals": signals, "detail": detail}


def classify_table(tbl, default_line: int = DEFAULT_LINE_TWIPS) -> dict:
    """判定表格类别：form（表单/块状表，需整表同页）或 data（数据表，可跨页）。"""
    sig = table_signals(tbl, default_line)
    kind = "form" if len(sig["signals"]) >= FORM_SIGNAL_THRESHOLD else "data"
    return {"kind": kind, "signals": sig["signals"], "detail": sig["detail"]}


def should_repeat_header(tbl, kind: str | None = None,
                         default_line: int = DEFAULT_LINE_TWIPS) -> bool:
    """是否标 tblHeader：像表头 **且** 不是表单表（表单表硬排除）。"""
    if kind is None:
        kind = classify_table(tbl, default_line)["kind"]
    if kind == "form":
        return False
    return bool(header_like_signals(tbl)["ok"])

_NUM_CELL_RE = re.compile(r"^[−\-+]?[\d,]+(\.\d+)?%?$")
_NOTE_RE = re.compile(r"^(注|说明|备注)\s*[:：]")
_UNIT_RE = re.compile(r"^[（(]?\s*(单位|万元|元|%|mm|cm|m|kg|s|小时|人|个|次)\s*[)）]?$")


def _row_cells(tr):
    """返回 [(col_index, tc, text, span)]，按 gridSpan 展开列号。"""
    out, col = [], 0
    for tc in tr.findall(f"{W}tc"):
        tcpr = tc.find(f"{W}tcPr")
        span = 1
        if tcpr is not None:
            gs = tcpr.find(f"{W}gridSpan")
            if gs is not None:
                try:
                    span = max(1, int(gs.get(f"{W}val") or 1))
                except (TypeError, ValueError):
                    span = 1
        text = "".join(t.text or "" for t in tc.iter(f"{W}t")).strip()
        out.append((col, tc, text, span))
        col += span
    return out


def table_profile(tbl, default_line: int = DEFAULT_LINE_TWIPS) -> dict:
    """表格画像：类别 + 表头行 + 数值列 + 单位行 + 注释行。

    确定性规则，不引入模型；用于驱动数值列对齐、数字不换行、续表等决策。
    """
    rows = tbl.findall(f"{W}tr")
    if not rows:
        return {"kind": "form", "header_rows": 0, "numeric_cols": [], "note_rows": [],
                "unit_row": None, "ncols": 0}
    matrix = [_row_cells(tr) for tr in rows]
    ncols = max((sum(sp for _, _, _, sp in r) for r in matrix), default=0)

    # 表头行：首行单元格文本都很短且不含数值（或首行整行加粗）
    header_rows = 0
    first = matrix[0]
    if first:
        texts = [t for _, _, t, _ in first if t]
        bold_cnt = sum(1 for _, tc, _, _ in first
                       if any(True for b in tc.iter(f"{W}b")))
        numeric = sum(1 for t in texts if _NUM_CELL_RE.match(t))
        if texts and (bold_cnt >= max(1, len(first) // 2)
                      or (numeric == 0 and max(len(t) for t in texts) <= 12)):
            header_rows = 1

    # 注释行 / 单位行
    note_rows = [i for i, r in enumerate(matrix)
                 if r and _NOTE_RE.match(r[0][2] or "")]
    unit_row = None
    for i, r in enumerate(matrix):
        if i == 0 or i in note_rows or not r:
            continue
        texts = [t for _, _, t, _ in r if t]
        if texts and all(_UNIT_RE.match(t) for t in texts):
            unit_row = i
            break

    # 数值列：跳过表头/注释/单位行后，该列 ≥60% 是数值且样本 ≥3
    skip = set(note_rows) | ({unit_row} if unit_row is not None else set())
    body_rows = [r for i, r in enumerate(matrix) if i >= header_rows and i not in skip]
    numeric_cols: list[int] = []
    for col in range(ncols):
        vals = []
        for r in body_rows:
            for c, _, t, sp in r:
                if c == col:
                    vals.append(t)
                    break
        vals = [v for v in vals if v]
        if len(vals) >= 3:
            hit = sum(1 for v in vals if _NUM_CELL_RE.match(v))
            if hit / len(vals) >= 0.6:
                numeric_cols.append(col)

    cls = classify_table(tbl, default_line)
    return {
        "kind": cls["kind"],
        "signals": cls["signals"],
        "header_rows": header_rows,
        "numeric_cols": numeric_cols,
        "note_rows": note_rows,
        "unit_row": unit_row,
        "ncols": ncols,
        "rows": len(rows),
    }
