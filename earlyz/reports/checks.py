"""格式检测报告 · 引擎自检（M1，2026-10-07）。

- check_reference_citations：文后引用序号核查（「未找到该引用序号」类问题的
  证据产出）。事实口径：正文标记总数（含上标状态）/ 文后条目数 / 逐一对应。

铁律：只报事实，不猜平台判定逻辑；不足以下结论时如实说明。
"""

from __future__ import annotations

import re
from pathlib import Path

from earlyz.actions.new_actions import W, _read_part
from earlyz.actions.report_fixes import _para_stream, _para_text, _zone_map

_MARK_RE = re.compile(r"\[(\d+(?:\s*[-–—,，]\s*\d+)*)\]")


def _expand_ids(group: str) -> list[int]:
    ids: list[int] = []
    for piece in re.split(r"[,，]", group):
        piece = piece.strip()
        m = re.fullmatch(r"(\d+)\s*[-–—]\s*(\d+)", piece)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if 0 < a <= b <= a + 200:
                ids.extend(range(a, b + 1))
            continue
        if piece.isdigit():
            ids.append(int(piece))
    return ids


def _styles_superscript_map(path: Path) -> dict[str, bool]:
    """字符样式 → 是否上标（引用标记常经样式上标）。"""
    try:
        root = _read_part(path, "word/styles.xml")
    except Exception:  # noqa: BLE001
        return {}
    out: dict[str, bool] = {}
    for st in root.iter(f"{W}style"):
        sid = st.get(f"{W}styleId")
        if not sid:
            continue
        va = st.find(f"{W}rPr/{W}vertAlign")
        out[sid] = bool(va is not None and va.get(f"{W}val") == "superscript")
    return out


def _is_superscript(run, style_sup: dict[str, bool]) -> bool:
    rpr = run.find(f"{W}rPr")
    if rpr is None:
        return False
    va = rpr.find(f"{W}vertAlign")
    if va is not None:
        return va.get(f"{W}val") == "superscript"
    rs = rpr.find(f"{W}rStyle")
    if rs is not None:
        return style_sup.get(rs.get(f"{W}val") or "", False)
    return False


def check_reference_citations(path: str | Path) -> dict:
    """正文 [n] 引用标记 vs 文后参考文献条目 的核查结果（只读）。"""
    p = Path(path)
    root = _read_part(p, "word/document.xml")
    style_sup = _styles_superscript_map(p)
    paras = [(pel, _para_text(pel)) for pel in root.iter(f"{W}p")]
    zones = _zone_map(paras)

    marker_ids: set[int] = set()
    markers_total = 0
    sup_total = 0
    details: list[dict] = []
    for pi, (pel, text) in enumerate(paras):
        if not text.strip() or zones.get(pi) == "refs":
            continue
        for m in _MARK_RE.finditer(text):
            off = m.start()
            run = None
            for el in pel.iter(f"{W}t"):
                if off < len(el.text or ""):
                    run = next((a for a in el.iterancestors() if a.tag == f"{W}r"), None)
                    break
                off -= len(el.text or "")
            sup = _is_superscript(run, style_sup) if run is not None else False
            ids = _expand_ids(m.group(1))
            markers_total += 1
            sup_total += 1 if sup else 0
            marker_ids.update(ids)
            if not sup:
                details.append({"para": pi, "mark": m.group(0),
                                "snippet": text[max(0, m.start() - 18):m.end() + 18]})

    ref_entries = 0
    for i, (_pel, text) in enumerate(paras):
        if zones.get(i) == "refs" and text.strip():
            ref_entries += 1

    cited = set(i for i in range(1, ref_entries + 1) if i in marker_ids)
    uncited = sorted(set(range(1, ref_entries + 1)) - marker_ids)
    over = sorted(i for i in marker_ids if i > ref_entries)
    return {
        "markers_total": markers_total,
        "superscript_total": sup_total,
        "not_superscript": details[:20],
        "not_superscript_total": len(details),
        "ref_entries": ref_entries,
        "cited_ids": sorted(cited),
        "cited_count": len(cited),
        "uncited_entries": uncited[:50],
        "uncited_total": len(uncited),
        "over_entries": over[:50],
        "over_total": len(over),
    }
