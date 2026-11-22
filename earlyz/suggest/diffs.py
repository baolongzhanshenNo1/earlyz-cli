"""P2 建议层种子：文档级确定性 diff 引擎（§22.1 建议层「差异驱动」）。

比较「引擎产物」与「用户手动修正版」，按稳定指纹（§21.5 同源）逐段对齐，
输出段落级格式属性差异——供归因（attribution）生成根因假设与规则提案
（ADR-015：用户修正只生成规则提案，确认后才发布新 FormatSpec 版本）。

铁律：
- 只读提取直接属性（pPr/rPr），不做 cascade 解析（深度求值属 lint 职责）；
- 对齐只认稳定指纹（文本 + pStyle + drawing/table 标志），禁止裸索引；
- 文本内容变化的段落不参与格式建议（内容增删是用户决策，不进规则提案）。
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from lxml import etree

from earlyz.actions.new_actions import (
    _read_part,
    _stable_paragraph_fingerprint,
    _text_of,
)

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


@dataclass
class FormatDiff:
    """一个段落的一组格式属性差异（稳定定位：text_hash + role + occurrence）。"""

    selector: dict
    role: str | None
    text: str
    property_diffs: dict[str, dict] = field(default_factory=dict)
    target_index: Optional[int] = None  # 文档序（仅诊断呈现，不作定位依据）

    def to_dict(self) -> dict:
        return {
            "selector": self.selector,
            "role": self.role,
            "text": self.text[:60],
            "property_diffs": self.property_diffs,
            "target_index": self.target_index,
        }


# ── 段落级代表性格式提取（直接属性） ────────────────────────────

def _read_para_format(p) -> dict:
    """段落代表性直接格式（pPr + 首个有 rPr 的 run 的 rPr）。

    返回展示名称 → 值（font_size_pt/bold/jc/first_line_chars/line_pt 等）。
    多 run 混排段取其代表性 run；无 run 的段落仅 pPr。
    """
    from earlyz.ir.selectors import normalize_text

    fmt: dict = {}
    ppr = p.find(f"{W}pPr")
    if ppr is not None:
        jc = ppr.find(f"{W}jc")
        if jc is not None:
            fmt["jc"] = jc.get(f"{W}val")
        ind = ppr.find(f"{W}ind")
        if ind is not None:
            flc = ind.get(f"{W}firstLineChars")
            if flc:
                fmt["first_line_chars"] = float(flc)
            else:
                fl = ind.get(f"{W}firstLine")
                if fl:
                    fmt["first_line_twips"] = int(fl)
            left = ind.get(f"{W}left")
            if left:
                fmt["left_twips"] = int(left)
        spacing = ppr.find(f"{W}spacing")
        if spacing is not None:
            line = spacing.get(f"{W}line")
            if line:
                fmt["line_twips"] = int(line)
            rule = spacing.get(f"{W}lineRule")
            if rule:
                fmt["line_rule"] = rule
            before = spacing.get(f"{W}before")
            if before:
                fmt["before_twips"] = int(before)
            after = spacing.get(f"{W}after")
            if after:
                fmt["after_twips"] = int(after)
        ol = ppr.find(f"{W}outlineLvl")
        if ol is not None:
            fmt["outline_lvl"] = int(ol.get(f"{W}val", "0") or "0")
        pstyle = ppr.find(f"{W}pStyle")
        if pstyle is not None:
            fmt["p_style"] = pstyle.get(f"{W}val")

    run = None
    for r in p.iter(f"{W}r"):
        if r.find(f"{W}rPr") is not None:
            run = r
            break
    rpr = run.find(f"{W}rPr") if run is not None else None
    if rpr is not None:
        rfonts = rpr.find(f"{W}rFonts")
        if rfonts is not None:
            ea = rfonts.get(f"{W}eastAsia")
            if ea:
                fmt["font_east_asia"] = ea
        sz = rpr.find(f"{W}sz")
        if sz is not None:
            try:
                fmt["font_size_pt"] = float(sz.get(f"{W}val")) / 2.0
            except (TypeError, ValueError):
                pass
        b = rpr.find(f"{W}b")
        if b is not None:
            fmt["bold"] = b.get(f"{W}val") not in ("0", "false", "off")
        i = rpr.find(f"{W}i")
        if i is not None:
            fmt["italic"] = i.get(f"{W}val") not in ("0", "false", "off")
        baseline = rpr.find(f"{W}vertAlign")
        if baseline is not None:
            fmt["vert_align"] = baseline.get(f"{W}val")

    return fmt


def _para_selector(p, role: str | None, text: str, index: int) -> dict:
    """稳定定位选择器（§4.4 形态：text_hash + role + occurrence 诊断辅助）。"""
    from earlyz.ir.selectors import normalize_text

    norm = normalize_text(text)
    h = hashlib.sha1(norm.encode("utf-8")).hexdigest()[:16]
    sel: dict = {"node_type": "paragraph", "text_hash": h}
    if role:
        sel["semantic_role"] = role
    sel["target_index"] = index  # 诊断辅助（不作定位依据，§4.4）
    return sel


def _paired(base_root, edited_root):
    """稳定指纹对序列：按 (fingerprint, occurrence) 对齐两文档段落。

    返回 list[(base_para, edited_para, role, text, index)]（只含配对成功段）。
    指纹不匹配（文本/结构变化）跳过——内容增删不进格式建议。
    """
    def scan(root):
        paras = list(root.iter(f"{W}p"))
        seen: dict[str, int] = {}
        out = []
        for idx, p in enumerate(paras):
            fp = _stable_paragraph_fingerprint(p)
            occ = seen.get(fp, 0)
            seen[fp] = occ + 1
            out.append((p, fp, occ, idx))
        return out

    base_list = scan(base_root)
    edited_list = scan(edited_root)
    by_key: dict[tuple, list] = {}
    for p, fp, occ, idx in edited_list:
        by_key.setdefault((fp, occ), []).append(p)

    pairs = []
    for p, fp, occ, idx in base_list:
        cands = by_key.get((fp, occ))
        if cands:
            pairs.append((p, cands[0], fp, occ, idx))
    return pairs


def diff_documents(base_path: str | Path, edited_path: str | Path) -> list[FormatDiff]:
    """引擎产物 vs 用户修正版 → 段落级格式差异列表（仅属性变更段）。"""
    base_root = _read_part(str(base_path), "word/document.xml")
    edited_root = _read_part(str(edited_path), "word/document.xml")
    diffs: list[FormatDiff] = []
    for bp, ep, fp, occ, idx in _paired(base_root, edited_root):
        base_fmt = _read_para_format(bp)
        edited_fmt = _read_para_format(ep)
        prop_diffs: dict[str, dict] = {}
        for key in sorted(set(base_fmt) | set(edited_fmt)):
            bv = base_fmt.get(key)
            ev = edited_fmt.get(key)
            if bv != ev:
                prop_diffs[key] = {"base": bv, "edited": ev}
        if not prop_diffs:
            continue
        text = _text_of(bp) or ""
        role = None
        try:
            from earlyz.ir.analyzer import classify_paragraph_role
            from earlyz.spec.decompiler import _field_name  # noqa: F401
            ppr = bp.find(f"{W}pPr")
            numpr = ppr is not None and ppr.find(f"{W}numPr") is not None
            has_drawing = bp.find(f".//{W}drawing") is not None
            r = classify_paragraph_role(
                ppr, "", has_drawing, False, numpr, text,
            )
            role = str(r)
        except Exception:  # noqa: BLE001 — role 失败仅影响归因精确度
            role = None
        diffs.append(FormatDiff(
            selector=_para_selector(bp, role, text, idx),
            role=role,
            text=text,
            property_diffs=prop_diffs,
            target_index=idx,
        ))
    return diffs
