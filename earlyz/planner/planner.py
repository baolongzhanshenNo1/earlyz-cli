"""早睡 (EarlyZ) 3.2 — RepairPlanner（§8.7，ADR-005）。

确定性规划：LLM 不参与；从 Violation 直接编译 Typed Actions。
- 输出 NodeSelector（node_id/role/occurrence），禁止裸索引；
- 同规则多段落合并为批量 action（role 选择器）；
- 超 max_risk 的 violation 跳过并计数；
- preserve_content 时注入内容保持不变量。
"""

from __future__ import annotations
import json

from dataclasses import dataclass, field

from earlyz.ir.models import DocumentIR
from earlyz.lint.engine import Violation
from earlyz.planner.models import RepairAction, RepairPlan
from earlyz.spec.models import FormatSpec

AUTO_FIX_RISKS = ("R1", "R2")


@dataclass
class PlanningPolicy:
    max_risk: str = "R2"
    preserve_content: bool = True
    require_preview_for: list[str] = field(default_factory=lambda: ["R2", "R3"])


def _violation_key(v: Violation, role_rules: set[str]) -> tuple:
    """分组键：role/cover 规则按 rule_id 合并（批量）；scope 规则同 rule
    全节合并为一个动作；其余按 (规则, 节点)。"""
    if v.rule_id in role_rules:
        return (v.rule_id, "*role*")
    return (v.rule_id, v.node_id)


def plan_format_repairs(
    ir: DocumentIR,
    spec: FormatSpec,
    violations: list[Violation],
    policy: PlanningPolicy,
) -> RepairPlan:
    plan = RepairPlan(
        plan_id=f"plan:{ir.document_id}",
        source_violation_count=len(violations),
    )
    if policy.preserve_content:
        plan.invariants.append("content_hash_preserved")
    plan.preconditions.append("document_openable")

    role_rules = {
        v.rule_id for v in violations
        if (spec.rule_by_id(v.rule_id) and (
            "semantic_role" in (spec.rule_by_id(v.rule_id).selector or {})
            or "cover_anchor" in (spec.rule_by_id(v.rule_id).selector or {})
            or (spec.rule_by_id(v.rule_id).selector or {}).get("scope")
               in ("section_all", "header_all", "footer_all", "table_all")
        ))
    }

    groups: dict[tuple, list[Violation]] = {}
    for v in violations:
        if v.risk > policy.max_risk:
            plan.skipped_risks += 1
            continue
        if not v.auto_fixable:
            continue
        groups.setdefault(_violation_key(v, role_rules), []).append(v)

    for (rule_id, node_id), vs in groups.items():
        action = _build_action(ir, spec, vs[0], vs, policy)
        if action is None:
            continue
        plan.actions.append(action)
        plan.affected_nodes.append(node_id)

    # Different semantic rules may resolve to the same structural target
    # (for example heading2 and refs_heading sharing one abstract level).  Do
    # not rewrite the same target twice when action and parameters are equal.
    deduped = []
    by_key = {}
    for action in plan.actions:
        key = (action.action, json.dumps(action.params, ensure_ascii=False, sort_keys=True))
        prior = by_key.get(key)
        if prior is None:
            by_key[key] = action
            deduped.append(action)
            continue
        if action.reason and action.reason not in prior.reason:
            prior.reason = ",".join(filter(None, (prior.reason, action.reason)))
        if action.source and action.source not in prior.source:
            prior.source = "；".join(filter(None, (prior.source, action.source)))
    plan.actions = deduped

    for a in plan.actions:
        for part in ("word/document.xml", "word/styles.xml"):
            if part not in plan.affected_parts:
                plan.affected_parts.append(part)

    # 模板示例残留（2026-08-16）：**不在此生成删除动作**——§10.1 内容变化
    # 必须由明确的非排版工具承担（A+B 架构收口，用户拍板）。lint 的
    # content_residual violation 保留作解释层报告；删除由 compile_document
    # 的独立阶段 clean_example_residuals 显式执行（ActionSpec 声明
    # content_change + validate_delivery 授权放行）。

    return plan


def _build_action(
    ir: DocumentIR,
    spec: FormatSpec,
    first: Violation,
    group: list[Violation],
    policy: PlanningPolicy,
) -> RepairAction | None:
    rule = spec.rule_by_id(first.rule_id)
    sel = rule.selector if rule else {}
    reason = ",".join(v.violation_id for v in group)
    # 决策依据链（C1）：每个 plan 动作记录"依据哪条规范规则"——排版结果可解释性
    # = 产品护城河（用户问"为什么这么改"可溯源）
    source = (rule.provenance.detail if (rule and rule.provenance and rule.provenance.detail)
              else f"规范规则:{first.rule_id}")
    # 非格式属性类 violation（expected 不是 {property,value} 形状）必须在通用推导
    # 之前分流——否则 KeyError 会掀翻整个编译（2026-10-12 事故：封面克隆撞出
    # docPr 重复，1063 条 violation 里唯一一条非 property 形状的把它全干掉了，
    # 用户侧只看到「编译异常: 'property'」）。
    if first.rule_id == "rule:tpl.package.docpr_duplicate":
        # 包级不变量（§7.1 P0）：docPr id 重复 → repair_package_invariants
        return RepairAction(
            action="repair_package_invariants",
            params={"mode": "docpr"},
            risk=_risk_of(first), reason=reason,
            source="包级不变量（§7.1 P0）：wp:docPr id 必须包级唯一",
        )
    expected = {v.expected["property"]: v.expected["value"]
                for v in group if isinstance(v.expected, dict) and "property" in v.expected}
    if not expected:
        # 组内全部是报告层 violation（无 property 形状）：不生成排版动作，
        # 保留在 lint 报告里——规划层绝不能让整个编译崩掉。
        return None

    if sel.get("scope") == "document_default":
        return RepairAction(
            action="set_doc_defaults", params=expected,
            risk=_risk_of(first), reason=reason, source=source,
        )
    if "style_id" in sel:
        return RepairAction(
            action="create_or_update_style",
            params={"style_id": sel["style_id"], "properties": expected},
            risk=_risk_of(first), reason=reason, source=source,
        )
    if "abstract_num_id" in sel:
        params = {"abstract_num_id": int(sel["abstract_num_id"]), "ilvl": int(sel.get("level", 0))}
        params.update(expected)
        return RepairAction(
            action="set_list_level_format",
            params=params,
            risk=_risk_of(first), reason=reason, source=source,
        )
    if sel.get("scope") == "numbering_role":
        try:
            _, abstract_num_id, ilvl = first.node_id.split(":", 2)
        except ValueError:
            return None
        params = {
            "abstract_num_id": int(abstract_num_id),
            "ilvl": int(ilvl),
            **expected,
        }
        return RepairAction(
            action="set_list_level_format", params=params,
            risk=_risk_of(first), reason=reason, source=source,
        )
    if "semantic_role" in sel:
        role = sel["semantic_role"]
        # role 规则语义 = 该 role 全部段落应用（批量 action；已达标段幂等 no-op）
        # 同名标题口径（title_key / skip_title_keys）必须随 selector 下发给 writer
        selector = {k: sel[k] for k in ("semantic_role", "title_key", "skip_title_keys")
                    if k in sel}
        # 不变量：body 不允许居中/右对齐。2026-10-20 实测某标书模板把
        # 「封面标题（居中 14pt）」抽成正文样式，导致整篇正文居中 14pt。
        # 必须在这里改（参数与 verifier 同源），放到动作里改会让验证器判未生效。
        # 字号不做兜底：公文正文就是三号（16pt），压到 12pt 会破坏公文排版。
        if selector.get("semantic_role") == "body" and expected.get("jc") in ("center", "right", "end"):
            expected = dict(expected)
            expected["jc"] = "both"
        return RepairAction(
            action="set_paragraph_format",
            params={"selector": selector, "properties": expected},
            risk=_risk_of(first), reason=reason, source=source,
        )
    if "cover_anchor" in sel:
        # 封面锚点规则 = 克隆模板封面布局到匹配段落（T3；批量批量幂等）
        selector = {k: sel[k] for k in ("cover_anchor", "match", "norm") if k in sel}
        return RepairAction(
            action="set_paragraph_format",
            params={"selector": selector, "properties": expected},
            risk=_risk_of(first), reason=reason, source=source,
        )
    if sel.get("scope") == "header_all":
        # 页眉格式规则：全部 header part 非空段统一（批量幂等）
        return RepairAction(
            action="format_headers",
            params={"properties": expected},
            risk=_risk_of(first), reason=reason, source=source,
        )
    if sel.get("scope") == "footer_all":
        return RepairAction(
            action="format_footers",
            params={"properties": expected},
            risk=_risk_of(first), reason=reason, source=source,
        )
    if sel.get("scope") == "table_all":
        # 表格对齐规则：全部表格统一 tblPr jc（writer 批量幂等）
        return RepairAction(
            action="set_table_alignment",
            params={"properties": expected},
            risk=_risk_of(first), reason=reason, source=source,
        )
    if sel.get("scope") == "section_all":
        # 节级页面几何规则（T2）：所有 sectPr 统一页边距（writer 复用
        # set_page_margins；section_index=-1 表示全文各节逐一应用）
        return RepairAction(
            action="set_page_margins_all",
            params={"properties": expected},
            risk=_risk_of(first), reason=reason, source=source,
        )
    return None


def _risk_of(v: Violation) -> str:
    return v.risk
