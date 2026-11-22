"""P2 建议层种子：差异归因（§22.1 建议层「归因」职责）。

把 diff_documents 产出的每项格式差异对照当前 FormatSpec 规则，做确定性
根因分类（LLM 叙述属对话层，本模块只产出证据链与分类）：

- engine_bug：规则存在且 edited 值 == 规则要求 → 引擎执行缺陷（产物没做对），
  建议走引擎报告闭环（§22.3 engine_report），不作规则提案；
- user_deliberate：规则缺失或 edited 值 != 规则要求 → 用户有意修正/覆盖，
  生成规则修正候选（ADR-015 提案，确认后才发布）；
- template_gap：规则存在但 edited 值 == 规则要求 且 base 值 == 规则要求 →
  规则本身与用户期望不符（模板取值异常），候选规则修正（罕见）；
- rule_override：规则存在、edited 值 != 规则要求 → 用户对现有规则的覆盖意向，
  生成规则修正候选（同一规则新值）。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from earlyz.spec.models import FormatSpec

from .diffs import FormatDiff


@dataclass
class Attribution:
    """一个差异项的根因假设 + 证据链 + 规则修正候选。"""

    classification: str              # engine_bug / user_deliberate / rule_override / template_gap / doc_issue
    diff: FormatDiff
    matched_rule_id: Optional[str] = None
    evidence: list[str] = field(default_factory=list)
    recommendation: str = ""
    rule_patch: Optional[dict] = None   # 提案的 new_rule（FormatRule dict 形态），仅修正类

    def to_dict(self) -> dict:
        return {
            "classification": self.classification,
            "selector": self.diff.selector,
            "matched_rule_id": self.matched_rule_id,
            "evidence": self.evidence,
            "recommendation": self.recommendation,
            "rule_patch": self.rule_patch,
        }


def _find_rule_for(spec, diff: FormatDiff) -> Optional[str]:
    """按 semantic_role 优先、其次规范化文本 匹配覆盖规则（保守：只认 role）。

    spec=None（无模板/未编译）→ 无规则匹配，全部 user_deliberate 新规则候选。
    规则 selector 的语义 = 该 role 全部段落应用（planner 批量语义），
    text 级规则（cover_anchor 等）不参与建议层匹配（少见且语义特殊）。
    """
    if spec is None or not diff.role:
        return None
    from earlyz.ir.analyzer import heading_rule_applies
    for r in spec.rules:
        sel = r.selector or {}
        # heading 同名标题口径与 lint/writer 同一谓词（2026-09-29）
        if (sel.get("semantic_role") == diff.role
                and heading_rule_applies(sel, diff.role, diff.text or "")):
            return r.rule_id
    return None


def _rule_expected(spec: FormatSpec, rule_id: str, prop: str) -> Optional[object]:
    """从规则 expected 中找某展示属性的期望值（键名对齐 diffs 的展示名）。"""
    rule = spec.rule_by_id(rule_id)
    if rule is None:
        return None
    return (rule.expected or {}).get(prop)


def _evidence_of(diff: FormatDiff, rule_id: Optional[str], rule_hits: dict) -> list[str]:
    ev = []
    for prop, d in diff.property_diffs.items():
        base = d.get("base")
        edited = d.get("edited")
        line = f"{prop}: {base!r} → {edited!r}"
        if rule_id and prop in rule_hits:
            line += f"（规则要求 {rule_hits[prop]!r}）"
        ev.append(line)
    return ev


def attribute_diffs(diffs: list[FormatDiff], spec: FormatSpec) -> list[Attribution]:
    """逐差异归因（确定性；空 spec → 全部 user_deliberate 新规则候选）。"""
    out: list[Attribution] = []
    for diff in diffs:
        if not diff.property_diffs:
            continue
        rule_id = _find_rule_for(spec, diff)
        rule_hits: dict = {}
        if rule_id:
            for prop in diff.property_diffs:
                exp = _rule_expected(spec, rule_id, prop)
                if exp is not None:
                    rule_hits[prop] = exp
        evidence = _evidence_of(diff, rule_id, rule_hits)
        # 与规则完全一致的属性：引擎缺陷（产物没做对）
        engine_misses = [p for p, d in diff.property_diffs.items()
                         if p in rule_hits and d.get("edited") == rule_hits[p]
                         and d.get("base") != d.get("edited")]
        # 用户新值（无规则 或 与规则不一致）
        user_overrides = [p for p, d in diff.property_diffs.items()
                          if p not in rule_hits
                          or d.get("edited") != rule_hits[p]]

        if engine_misses and not user_overrides:
            out.append(Attribution(
                classification="engine_bug",
                diff=diff, matched_rule_id=rule_id, evidence=evidence,
                recommendation=(
                    "引擎执行缺陷：属性值与规范规则一致但产物未落实——"
                    "建议反馈引擎报告（engine_report），不作规则提案"
                ),
            ))
            continue
        if not user_overrides:
            out.append(Attribution(
                classification="doc_issue",
                diff=diff, matched_rule_id=rule_id, evidence=evidence,
                recommendation="文档自身属性异常，无规则线索——建议人工复核",
            ))
            continue
        # 有用户新值 → 规则修正候选（ADR-015）
        if rule_id:
            classification = "rule_override"
            recommendation = (
                f"用户对规则 {rule_id} 的覆盖意向（{', '.join(sorted(user_overrides))}）"
                "——确认后更新规则（原值 → 用户值）"
            )
        else:
            classification = "user_deliberate"
            recommendation = (
                f"模板未覆盖的新格式要求（{', '.join(sorted(user_overrides))}）"
                "——确认后新增规则"
            )
        expected_candidate = dict(rule_hits)
        for prop in user_overrides:
            expected_candidate[prop] = diff.property_diffs[prop]["edited"]
        new_rule = _build_rule_patch(
            spec, diff, rule_id, expected_candidate, recommendation,
        )
        out.append(Attribution(
            classification=classification,
            diff=diff, matched_rule_id=rule_id, evidence=evidence,
            recommendation=recommendation, rule_patch=new_rule,
        ))
    return out


def _build_rule_patch(spec: FormatSpec, diff: FormatDiff, rule_id: Optional[str],
                      expected: dict, recommendation: str) -> dict:
    """构造 FormatRule 的 dict 形态（proposal 载荷；确认后按 rule_id upsert）。"""
    existing = spec.rule_by_id(rule_id) if rule_id else None
    if existing is not None:
        return {
            "rule_id": existing.rule_id,
            "selector": existing.selector,
            "expected": {**(existing.expected or {}), **expected},
            "severity": existing.severity,
            "auto_fix": existing.auto_fix,
            "fix_strategy": existing.fix_strategy,
            "risk_ceiling": existing.risk_ceiling,
        }
    # 新规则：值来自用户修正 + 模板归属（spec_id 在 proposal 层补）
    role = diff.role
    return {
        "rule_id": f"rule:{diff.text[:12] or role or 'para'}",
        "selector": {"semantic_role": role} if role else {
            "node_type": "paragraph",
            "text_hash": diff.selector.get("text_hash"),
        },
        "expected": expected,
        "severity": "error",
        "auto_fix": True,
        "fix_strategy": "direct",
        "risk_ceiling": "R2",
    }
