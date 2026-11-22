"""早睡 (EarlyZ) 3.2 — 学习闭环（§15）：learn_format_correction。

用户对编译结果不满意 → 提交纠正（文档 + 目标属性）→ 生成规则提案
（差异驱动：实际 vs 期望）→ 确认后追加到用户规范 JSON 文件
（provenance=user_confirmed，ADR-015）。
"""

from __future__ import annotations

import json
import re
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from earlyz.ir.analyzer import analyze_document


@dataclass
class CorrectionProposal:
    proposal_id: str
    rule: dict
    context: dict          # {semantic_role / style_id / abstract_num_id ...}
    property: str
    expected_value: object
    actual_value: object
    source_document: str
    status: str = "pending"   # pending / confirmed / rejected

    def to_dict(self) -> dict:
        return {
            "proposal_id": self.proposal_id,
            "rule": self.rule,
            "context": self.context,
            "property": self.property,
            "expected_value": self.expected_value,
            "actual_value": self.actual_value,
            "source_document": self.source_document,
            "status": self.status,
        }


def _rule_id(context: dict, prop: str) -> str:
    base = next(iter(context.values()), "doc")
    return f"rule:{base}.{prop}"


def _build_rule(context: dict, prop: str, expected_value, domain: str, source_doc: str) -> dict:
    return {
        "rule_id": _rule_id(context, prop),
        "selector": dict(context),
        "expected": {prop: expected_value},
        "severity": "error",
        "auto_fix": True,
        "risk_ceiling": "R2",
        "provenance": {
            "source": "user_confirmed",
            "detail": f"用户纠正：{prop} → {expected_value}",
            "artifact_id": source_doc,
        },
        "confidence": 1.0,
    }


def propose_correction(
    document: str | Path,
    context: dict,
    property_name: str,
    expected_value,
    domain: str = "thesis",
) -> CorrectionProposal:
    """从用户纠正生成规则提案：读文档 IR 取实际值，构造差异驱动规则。"""
    ir = analyze_document(document)

    # 取实际值：按 context 类型定位
    actual = None
    if "semantic_role" in context:
        for p in ir.paragraphs:
            if p.semantic_role == context["semantic_role"]:
                actual = getattr(p.direct_format, property_name, None)
                break
    elif "style_id" in context:
        for s in ir.styles:
            if s.style_id == context["style_id"]:
                actual = getattr(s.style_format, property_name, None)
                break
    elif "abstract_num_id" in context:
        aid = int(context["abstract_num_id"])
        lvl = int(context.get("level", 0))
        for ld in ir.numbering:
            if ld.abstract_num_id == aid:
                for lv in ld.levels:
                    if lv.level == lvl:
                        actual = getattr(lv, property_name, None)
                break

    rule = _build_rule(context, property_name, expected_value, domain, str(document))
    return CorrectionProposal(
        proposal_id=f"cp_{uuid.uuid4().hex[:10]}",
        rule=rule,
        context=dict(context),
        property=property_name,
        expected_value=expected_value,
        actual_value=actual,
        source_document=str(document),
    )


def confirm_correction(proposal: CorrectionProposal, spec_json_path: str | Path) -> Path:
    """确认提案：追加规则到用户规范 JSON（幂等：同 rule_id 已存在则跳过）。"""
    p = Path(spec_json_path)
    if not p.exists():
        raise FileNotFoundError(f"spec json not found: {p}")
    data = json.loads(p.read_text(encoding="utf-8"))
    rules = data.setdefault("rules", [])
    if any(r.get("rule_id") == proposal.rule["rule_id"] for r in rules):
        proposal.status = "confirmed"
        return p
    rules.append(proposal.rule)
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    proposal.status = "confirmed"
    return p


def load_user_rules(spec_json_path: str | Path) -> list[dict]:
    """读取用户规范中的 user_confirmed 规则。"""
    data = json.loads(Path(spec_json_path).read_text(encoding="utf-8"))
    return [r for r in data.get("rules", [])
            if r.get("provenance", {}).get("source") == "user_confirmed"]
