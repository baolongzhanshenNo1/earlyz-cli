"""早睡 (EarlyZ) 3.2 — Visual Diff（§14）：render_diff_preview。

对比 原文档 vs 编译产物 的 IR 级差异：
- 段落/样式/编号/表格属性差异（before → after）
- 内容保持检测（文本变化 = 违规标记）
纯 OOXML 分析，不依赖渲染器（视觉渲染为可选扩展）。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from earlyz.ir.analyzer import analyze_document
from earlyz.ir.models import FormatProperties


def _props(fp: FormatProperties) -> dict:
    return {k: v for k, v in fp.__dict__.items() if v is not None}


@dataclass
class ParagraphDiff:
    node_id: str
    role: str
    text: str
    text_changed: bool = False
    property_changes: dict = field(default_factory=dict)  # prop -> (before, after)


@dataclass
class DiffReport:
    source: str
    target: str
    paragraphs: list[ParagraphDiff] = field(default_factory=list)
    style_changes: list[dict] = field(default_factory=list)
    list_changes: list[dict] = field(default_factory=list)
    text_change_count: int = 0
    property_change_count: int = 0

    def to_dict(self) -> dict:
        return {
            "source": self.source,
            "target": self.target,
            "paragraphs": [
                {
                    "node_id": p.node_id, "role": p.role, "text": p.text[:40],
                    "text_changed": p.text_changed, "property_changes": p.property_changes,
                }
                for p in self.paragraphs
            ],
            "style_changes": self.style_changes,
            "list_changes": self.list_changes,
            "text_change_count": self.text_change_count,
            "property_change_count": self.property_change_count,
        }


def _diff_paragraphs(before, after) -> list[ParagraphDiff]:
    bmap = {p.node_id: p for p in before.paragraphs}
    amap = {p.node_id: p for p in after.paragraphs}
    # 同 role 顺序索引（node_id 含文本 hash，文本变化后需 role+序号回退对齐）
    bseq: dict[str, list] = {}
    aseq: dict[str, list] = {}
    for p in before.paragraphs:
        bseq.setdefault(p.semantic_role, []).append(p)
    for p in after.paragraphs:
        aseq.setdefault(p.semantic_role, []).append(p)

    out: list[ParagraphDiff] = []
    for nid, b in bmap.items():
        a = amap.get(nid)
        if a is None:
            # 回退：同 role 同序号对齐（内容可能变化）
            idx = bseq[b.semantic_role].index(b)
            seq = aseq.get(b.semantic_role, [])
            a = seq[idx] if idx < len(seq) else None
        if a is None:
            continue
        d = ParagraphDiff(node_id=nid, role=b.semantic_role, text=b.text)
        if b.text != a.text:
            d.text_changed = True
        bp, ap = _props(b.direct_format), _props(a.direct_format)
        for prop in set(bp) | set(ap):
            if bp.get(prop) != ap.get(prop):
                d.property_changes[prop] = (bp.get(prop), ap.get(prop))
        if d.text_changed or d.property_changes:
            out.append(d)
    return out


def render_diff_preview(source_docx: str | Path, target_docx: str | Path) -> DiffReport:
    """对比源文档与编译产物，输出 IR 级差异报告。"""
    ir_before = analyze_document(source_docx)
    ir_after = analyze_document(target_docx)

    report = DiffReport(source=str(source_docx), target=str(target_docx))
    report.paragraphs = _diff_paragraphs(ir_before, ir_after)

    # 样式差异
    bstyles = {s.style_id: s for s in ir_before.styles}
    astyles = {s.style_id: s for s in ir_after.styles}
    for sid in bstyles:
        if sid in astyles:
            bp, ap = _props(bstyles[sid].style_format), _props(astyles[sid].style_format)
            changes = {k: (bp.get(k), ap.get(k)) for k in set(bp) | set(ap)
                       if bp.get(k) != ap.get(k)}
            if changes:
                report.style_changes.append({"style_id": sid, "changes": changes})
        else:
            report.style_changes.append({"style_id": sid, "changes": {"__deleted__": (True, None)}})
    for sid in set(astyles) - set(bstyles):
        report.style_changes.append({"style_id": sid, "changes": {"__created__": (None, True)}})

    # 列表差异（abstractNum 级格式）
    bl = {f"{ld.abstract_num_id}:{lv.level}": lv for ld in ir_before.numbering for lv in ld.levels}
    al = {f"{ld.abstract_num_id}:{lv.level}": lv for ld in ir_after.numbering for lv in ld.levels}
    for key in bl:
        if key in al and (bl[key].num_fmt != al[key].num_fmt or bl[key].lvl_text != al[key].lvl_text):
            report.list_changes.append({
                "level": key,
                "num_fmt": (bl[key].num_fmt, al[key].num_fmt),
                "lvl_text": (bl[key].lvl_text, al[key].lvl_text),
            })

    report.text_change_count = sum(1 for p in report.paragraphs if p.text_changed)
    report.property_change_count = sum(len(p.property_changes) for p in report.paragraphs)
    return report
