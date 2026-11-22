"""
Format Linter 与 Violation 模型（§8.5，ADR-009）。

lint_format 比较 DocumentIR 与 FormatSpec，产出结构化 Violation；
绝不修改文档（R0）。比较对象：
- scope=document_default → ir.font_table["doc_defaults"]
- selector.style_id → ir.styles
- selector.semantic_role → ir.paragraphs（P0 用 direct_format，resolved 级联后续扩展）
"""

from __future__ import annotations
import re
from dataclasses import dataclass, field
from typing import Any, Optional
from earlyz.ir.models import DocumentIR, FormatProperties
from earlyz.ir.ooxml_format import structure_bundle_contains
from earlyz.spec.models import FormatSpec






@dataclass
class Violation:
    violation_id: str
    rule_id: str
    node_id: str
    expected: dict
    actual: dict
    source: str
    severity: str = "error"      # error/warning/info
    auto_fixable: bool = True
    risk: str = "R2"
    page: Optional[int] = None
    detail: str = ""

    SEVERITY_ORDER = {"info": 1, "warning": 2, "error": 3}

    def to_dict(self) -> dict:
        return {
            "violation_id": self.violation_id,
            "rule_id": self.rule_id,
            "node_id": self.node_id,
            "page": self.page,
            "expected": self.expected,
            "actual": self.actual,
            "source": self.source,
            "severity": self.severity,
            "auto_fixable": self.auto_fixable,
            "risk": self.risk,
        }


@dataclass
class LintResult:
    violations: list[Violation] = field(default_factory=list)

    @property
    def counts(self) -> dict[str, int]:
        c = {"error": 0, "warning": 0, "info": 0}
        for v in self.violations:
            c[v.severity] = c.get(v.severity, 0) + 1
        return c

    def to_dict(self) -> dict:
        return {
            "total": len(self.violations),
            "counts": self.counts,
            "violations": [v.to_dict() for v in self.violations],
        }


def _props_to_dict(fp: FormatProperties) -> dict:
    return {k: v for k, v in fp.__dict__.items() if v is not None}


# spec 惯例字段名 → IR 字段名（lint 比较时归一化）
_PROP_ALIASES = {
    "first_line_indent_chars": "ind_first_line_chars",
}

# A paragraph without textual runs has no run object to repair. Paragraph-level
# properties still apply; run-only requirements are intentionally inapplicable
# rather than violations (creating a dummy run would mutate content structure).
_RUN_ONLY_PROPERTIES = {
    "font_east_asia", "font_ascii", "font_size_pt", "bold", "italic",
    "underline", "color", "small_caps", "rpr_extra_xml",
}


def _actual_prop(actual: dict, prop: str):
    """按 spec 字段名取实际值（含别名归一化）。"""
    return actual.get(_PROP_ALIASES.get(prop, prop))


def _expected_with_prop(prop: str, value: Any) -> dict:
    return {"property": prop, "value": value}


def _paragraph_actual(p, expected: dict) -> dict:
    """Return paragraph + label-run properties in the spec vocabulary."""
    actual = _props_to_dict(p.direct_format)
    label_text = expected.get("label_text")
    label_props = {
        "label_font_east_asia": "font_east_asia",
        "label_font_ascii": "font_ascii",
        "label_font_size_pt": "font_size_pt",
        "label_bold": "bold",
        "label_italic": "italic",
    }
    if label_text is None and not any(k in expected for k in label_props):
        return actual

    label_run = None
    label_idx = None
    if label_text:
        core_lb = re.sub(r"[\s:：]+", "", str(label_text))
        for i, run in enumerate(p.runs):
            core_run = re.sub(r"[\s:：]+", "", (run.text or "").lstrip())
            if core_run.startswith(core_lb):
                label_run = run
                label_idx = i
                break
        # 标签文本须与模板精确一致（含空格/冒号样式）才算对齐；
        # 「摘要：」vs 模板「摘  要」等差异报 violation，由
        # set_paragraph_format 按模板原文修正（2026-08-21 用户裁决）。
        if label_run is not None and (label_run.text or "").strip() == str(label_text):
            actual["label_text"] = str(label_text)
    if label_run is None and p.runs:
        label_run = p.runs[0]
        label_idx = 0
    if label_run is not None:
        fp = _props_to_dict(label_run.direct_format)
        for spec_key, ir_key in label_props.items():
            if spec_key in expected:
                actual[spec_key] = fp.get(ir_key)
        # 模板标签 run 后可有 <w:tab/>（Word 显示箭头，2026-08-22 用户裁决
        # 「关键词与正文间有一个 tab」）——目标实际值 = 标签 run 后下一 run 是否
        # 含 tab；无 tab → 报 violation 由 set_paragraph_format 按模板补。
        if "label_after_tab" in expected and label_idx is not None:
            runs = getattr(p, "runs", None) or []
            nxt = runs[label_idx + 1] if label_idx + 1 < len(runs) else None
            actual["label_after_tab"] = bool(nxt is not None and getattr(nxt, "has_tab", False))
    return actual


def lint_format(
    ir: DocumentIR,
    spec: FormatSpec,
    severity_floor: str = "info",
) -> LintResult:
    """比较 IR 与 Spec，产出 Violation 列表（§8.5 工具契约）。"""
    result = LintResult()
    floor = Violation.SEVERITY_ORDER.get(severity_floor, 1)
    vid = 0

    # 模板示例残留（2026-08-16）：理解层识别模板示例教学内容，
    # 目标文档出现完全匹配段 = 从模板照搬的残留 → 解释为 content_residual
    _residuals = set(spec.example_residuals or [])
    for p_i, pnode in enumerate(ir.paragraphs):
        if not (pnode.text or "").strip():
            continue  # 含封面区——示例残留大多在封面
        t = re.sub(r"[\s\u3000□]+", "", pnode.text or "")
        t = t.replace("：", ":").replace("；", ";").replace("（", "(").replace("）", ")")
        t = t.replace("，", ",").replace("。", ".").replace("、", ",").strip()
        if t in _residuals:
            result.violations.append(Violation(
                violation_id=f"content_residual:{p_i}",
                rule_id="rule:tpl.content_residual",
                node_id=f"paragraph[{p_i}]",
                expected={"content": "模板示例内容不应出现在目标文档"},
                actual={"content": pnode.text.strip()[:60]},
                source="模板示例教学内容（理解层提取）",
                severity="warning",
                auto_fixable=False,
                risk="R3",
                detail="模板仅用于提取格式；与模板示例文本完全匹配的段落为照搬残留。"
                       "由独立内容清理工具（clean_example_residuals 阶段）处理，"
                       "不进入格式修复链（§10.1）",
            ))
            vid += 1

    # 包级不变量：wp:docPr id 必须包级唯一（§7.1 repair_package_invariants）
    seen_docpr: dict[str, int] = {}
    for fig in ir.figures:
        did = fig.docpr_id
        if did:
            if did in seen_docpr:
                result.violations.append(Violation(
                    violation_id=f"package.docpr_duplicate:{did}",
                    rule_id="rule:tpl.package.docpr_duplicate",
                    node_id=fig.node_id,
                    expected={"docpr_id": "包级唯一"},
                    actual={"docpr_id": did, "duplicate_of": f"figure[{seen_docpr[did]}]"},
                    source="包级不变量（§7.1）",
                    severity="warning",
                    auto_fixable=True,
                    risk="R2",
                    detail="wp:docPr id 重复会导致 Word 对象选择异常；由 repair_package_invariants 修复",
                ))
                vid += 1
            else:
                seen_docpr[did] = fig.index

    for rule in spec.rules:
        if Violation.SEVERITY_ORDER.get(rule.severity, 1) < floor:
            continue
        sel = rule.selector or {}
        scope = sel.get("scope")
        style_id = sel.get("style_id")
        role = sel.get("semantic_role")

        if scope == "document_default":
            actual = ir.font_table.get("doc_defaults", {}) if ir.font_table else {}
            for prop, expected_val in rule.expected.items():
                actual_val = actual.get(prop)
                if actual_val != expected_val:
                    result.violations.append(Violation(
                        violation_id=f"v:{vid}", rule_id=rule.rule_id,
                        node_id=ir.document_id or "document",
                        expected=_expected_with_prop(prop, expected_val),
                        actual=_expected_with_prop(prop, actual_val),
                        source=rule.provenance.detail if rule.provenance else "spec",
                        severity=rule.severity,
                        auto_fixable=rule.auto_fix and rule.risk_ceiling in ("R1", "R2"),
                        risk=rule.risk_ceiling,
                    ))
                    vid += 1

        elif style_id is not None:
            matched = False
            for st in ir.styles:
                if st.style_id != style_id:
                    continue
                matched = True
                actual = _props_to_dict(st.style_format)
                for prop, expected_val in rule.expected.items():
                    if _actual_prop(actual, prop) != expected_val:
                        result.violations.append(Violation(
                            violation_id=f"v:{vid}", rule_id=rule.rule_id,
                            node_id=st.node_id,
                            expected=_expected_with_prop(prop, expected_val),
                            actual=_expected_with_prop(prop, _actual_prop(actual, prop)),
                            source=rule.provenance.detail if rule.provenance else "spec",
                            severity=rule.severity,
                            auto_fixable=rule.auto_fix and rule.risk_ceiling in ("R1", "R2"),
                            risk=rule.risk_ceiling,
                        ))
                        vid += 1
            # 差异驱动：spec 要求但文档缺失的样式 → 产生 violation（创建语义）
            if not matched:
                for prop, expected_val in rule.expected.items():
                    result.violations.append(Violation(
                        violation_id=f"v:{vid}", rule_id=rule.rule_id,
                        node_id=f"style:{style_id}",
                        expected=_expected_with_prop(prop, expected_val),
                        actual=_expected_with_prop(prop, None),
                        source=rule.provenance.detail if rule.provenance else "spec",
                        severity=rule.severity,
                        auto_fixable=rule.auto_fix and rule.risk_ceiling in ("R1", "R2"),
                        risk=rule.risk_ceiling,
                    ))
                    vid += 1

        elif "abstract_num_id" in sel:
            # 多级编号规则：比较 abstractNum 指定级的格式（§Phase2 编号系统）
            level = int(sel.get("level", 0))
            aid = int(sel["abstract_num_id"])
            for ld in ir.numbering:
                if ld.abstract_num_id != aid:
                    continue
                level_node = next((lvl for lvl in ld.levels if lvl.level == level), None)
                actual = {
                    "num_fmt": level_node.num_fmt if level_node else None,
                    "lvl_text": level_node.lvl_text if level_node else None,
                }
                for prop, expected_val in rule.expected.items():
                    if actual.get(prop) != expected_val:
                        result.violations.append(Violation(
                            violation_id=f"v:{vid}", rule_id=rule.rule_id,
                            node_id=f"list_level:{aid}:{level}",
                            expected=_expected_with_prop(prop, expected_val),
                            actual=_expected_with_prop(prop, _actual_prop(actual, prop)),
                            source=rule.provenance.detail if rule.provenance else "spec",
                            severity=rule.severity,
                            auto_fixable=rule.auto_fix and rule.risk_ceiling in ("R1", "R2"),
                            risk=rule.risk_ceiling,
                        ))
                        vid += 1

        elif role is not None and sel.get("scope") is None:
            # 封面区段落归 cover_anchor 规则管，role 规则跳过（防止 body
            # 主导值把封面大标题刷成缩进/两端对齐）
            from earlyz.spec.decompiler import find_cover_end
            from earlyz.ir.analyzer import (
                FRONT_ROLES, heading_rule_applies, zone_is_article_head,
            )
            cover_end = find_cover_end(ir)
            # 与交付门 / writer 同一谓词（2026-09-29）
            article_head = zone_is_article_head(
                p.semantic_role for p in ir.paragraphs[:cover_end]
                if (p.text or "").strip())
            for idx, p in enumerate(ir.paragraphs):
                if p.semantic_role != role:
                    continue
                # 前置页标题 / 同名标题口径（与 writer 同一谓词，2026-09-29）
                if not heading_rule_applies(sel, role, p.text or ""):
                    continue
                # 学位论文封面区仍跳过（含被判成前部角色的校名/题名）；
                # 期刊题名/作者/摘要就在第一个标题之前，必须走 role 规则。
                if idx < cover_end and not (role in FRONT_ROLES and article_head):
                    continue
                actual = _paragraph_actual(p, rule.expected)
                for prop, expected_val in rule.expected.items():
                    if not p.runs and not (p.text or "").strip() and prop in _RUN_ONLY_PROPERTIES:
                        continue
                    if _actual_prop(actual, prop) != expected_val:
                        result.violations.append(Violation(
                            violation_id=f"v:{vid}", rule_id=rule.rule_id,
                            node_id=p.node_id,
                            expected=_expected_with_prop(prop, expected_val),
                            actual=_expected_with_prop(prop, _actual_prop(actual, prop)),
                            source=rule.provenance.detail if rule.provenance else "spec",
                            severity=rule.severity,
                            auto_fixable=rule.auto_fix and rule.risk_ceiling in ("R1", "R2"),
                            risk=rule.risk_ceiling,
                        ))
                        vid += 1

        elif sel.get("scope") in ("header_all", "footer_all"):
            # 页眉/页脚公共格式：只比较格式指纹，不比较模板文字或 PAGE 域。
            part_map = ir.headers if sel.get("scope") == "header_all" else ir.footers
            prefix = "header" if sel.get("scope") == "header_all" else "footer"
            for part, paras in (part_map or {}).items():
                for pi, para in enumerate(paras):
                    actual = _props_to_dict(para["fp"])
                    for prop, expected_val in rule.expected.items():
                        if _actual_prop(actual, prop) != expected_val:
                            result.violations.append(Violation(
                                violation_id=f"v:{vid}", rule_id=rule.rule_id,
                                node_id=f"{prefix}:{part}:{pi}",
                                expected=_expected_with_prop(prop, expected_val),
                                actual=_expected_with_prop(prop, _actual_prop(actual, prop)),
                                source=rule.provenance.detail if rule.provenance else "spec",
                                severity=rule.severity,
                                auto_fixable=rule.auto_fix and rule.risk_ceiling in ("R1", "R2"),
                                risk=rule.risk_ceiling,
                            ))
                            vid += 1

        elif sel.get("scope") == "table_all":
            # 表格对齐规则（T2b）：模板全一致 tblPr jc vs 目标各表
            for t in ir.tables or []:
                actual = {"jc": t.alignment, "tblpr_extra_xml": t.tblpr_extra_xml}
                for prop, expected_val in rule.expected.items():
                    actual_val = _actual_prop(actual, prop)
                    matches = (
                        structure_bundle_contains(actual_val, expected_val)
                        if prop == "tblpr_extra_xml"
                        else actual_val == expected_val
                    )
                    if not matches:
                        result.violations.append(Violation(
                            violation_id=f"v:{vid}", rule_id=rule.rule_id,
                            node_id=t.node_id,
                            expected=_expected_with_prop(prop, expected_val),
                            actual=_expected_with_prop(prop, actual.get(prop)),
                            source=rule.provenance.detail if rule.provenance else "spec",
                            severity=rule.severity,
                            auto_fixable=rule.auto_fix and rule.risk_ceiling in ("R1", "R2"),
                            risk=rule.risk_ceiling,
                        ))
                        vid += 1

        elif sel.get("scope") == "numbering_role":
            role = sel.get("numbering_role")
            expected_level = int(sel.get("level", 0))
            definitions = {d.num_id: d for d in (ir.numbering or [])}
            seen_levels: set[tuple[int, int]] = set()
            for para in ir.paragraphs or []:
                if para.semantic_role != role:
                    continue
                binding = para.list_binding or {}
                num_id = binding.get("num_id")
                ilvl = binding.get("ilvl")
                if num_id in (None, 0) or ilvl != expected_level:
                    continue
                definition = definitions.get(num_id)
                if definition is None or definition.abstract_num_id is None:
                    continue
                key = (definition.abstract_num_id, ilvl)
                if key in seen_levels:
                    continue
                seen_levels.add(key)
                level = next((x for x in definition.levels if x.level == ilvl), None)
                if level is None:
                    continue
                actual = {
                    "num_fmt": level.num_fmt,
                    "lvl_text": level.lvl_text,
                    "start": level.start,
                    "level_extra_xml": level.level_extra_xml,
                }
                for prop, expected_val in rule.expected.items():
                    actual_val = actual.get(prop)
                    matches = (
                        structure_bundle_contains(actual_val, expected_val)
                        if prop == "level_extra_xml"
                        else actual_val == expected_val
                    )
                    if matches:
                        continue
                    result.violations.append(Violation(
                        violation_id=f"v:{vid}", rule_id=rule.rule_id,
                        node_id=f"list_level:{definition.abstract_num_id}:{ilvl}",
                        expected=_expected_with_prop(prop, expected_val),
                        actual=_expected_with_prop(prop, actual_val),
                        source=rule.provenance.detail if rule.provenance else "spec",
                        severity=rule.severity,
                        auto_fixable=rule.auto_fix and rule.risk_ceiling in ("R1", "R2"),
                        risk=rule.risk_ceiling,
                    ))
                    vid += 1

        elif sel.get("scope") == "section_all":
            # 节级页面几何规则（T2）：模板全一致页边距 vs 目标各节
            from earlyz.spec.decompiler import margins_diff
            for s in ir.sections or []:
                diff = margins_diff(s.margins or {}, rule.expected)
                expected_extra = rule.expected.get("section_extra_xml")
                if expected_extra is not None and not structure_bundle_contains(
                    s.section_extra_xml, expected_extra
                ):
                    diff["section_extra_xml"] = expected_extra
                for prop, expected_val in diff.items():
                    result.violations.append(Violation(
                        violation_id=f"v:{vid}", rule_id=rule.rule_id,
                        node_id=s.node_id,
                        expected=_expected_with_prop(prop, expected_val),
                        actual=_expected_with_prop(
                            prop,
                            s.section_extra_xml if prop == "section_extra_xml"
                            else (s.margins or {}).get(prop),
                        ),
                        source=rule.provenance.detail if rule.provenance else "spec",
                        severity=rule.severity,
                        auto_fixable=rule.auto_fix and rule.risk_ceiling in ("R1", "R2"),
                        risk=rule.risk_ceiling,
                    ))
                    vid += 1

        elif "cover_anchor" in sel:
            # 封面锚点规则（T3）：克隆模板封面布局到匹配段落
            from earlyz.spec.decompiler import match_cover_paragraphs
            for p in match_cover_paragraphs(ir, sel):
                actual = _props_to_dict(p.direct_format)
                for prop, expected_val in rule.expected.items():
                    if _actual_prop(actual, prop) != expected_val:
                        result.violations.append(Violation(
                            violation_id=f"v:{vid}", rule_id=rule.rule_id,
                            node_id=p.node_id,
                            expected=_expected_with_prop(prop, expected_val),
                            actual=_expected_with_prop(prop, _actual_prop(actual, prop)),
                            source=rule.provenance.detail if rule.provenance else "spec",
                            severity=rule.severity,
                            auto_fixable=rule.auto_fix and rule.risk_ceiling in ("R1", "R2"),
                            risk=rule.risk_ceiling,
                        ))
                        vid += 1

    return result
