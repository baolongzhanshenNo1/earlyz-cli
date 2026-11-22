"""早睡 (EarlyZ) 3.2 — RepairPlan 模型（§8.7）。"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

IR_SCHEMA_VERSION = "3.2"


@dataclass
class RepairAction:
    """一个确定性的 typed action。"""

    action: str
    params: dict[str, Any]
    risk: str = "R1"
    verifier: str = ""
    reason: str = ""          # 产生原因（violation_id 列表）
    source: str = ""          # 依据规则溯源（SpecProvenance.detail，"为什么这么改"）

    def to_dict(self) -> dict:
        return {
            "action": self.action,
            "params": self.params,
            "risk": self.risk,
            "reason": self.reason,
            "source": self.source,
        }


@dataclass
class RepairPlan:
    """确定性修复计划（§8.7 输出契约）。"""

    plan_id: str
    schema_version: str = IR_SCHEMA_VERSION
    actions: list[RepairAction] = field(default_factory=list)
    affected_nodes: list[str] = field(default_factory=list)
    affected_parts: list[str] = field(default_factory=list)
    affected_pages_estimate: list[int] = field(default_factory=list)
    preconditions: list[str] = field(default_factory=list)
    invariants: list[str] = field(default_factory=list)
    skipped_risks: int = 0
    source_violation_count: int = 0

    @property
    def safe_count(self) -> int:
        return sum(1 for a in self.actions if a.risk == "R1")

    @property
    def structural_count(self) -> int:
        return sum(1 for a in self.actions if a.risk == "R2")

    @property
    def destructive_count(self) -> int:
        return sum(1 for a in self.actions if a.risk == "R3")

    def to_dict(self) -> dict:
        return {
            "plan_id": self.plan_id,
            "schema_version": self.schema_version,
            "actions": [a.to_dict() for a in self.actions],
            "safe_count": self.safe_count,
            "structural_count": self.structural_count,
            "destructive_count": self.destructive_count,
            "affected_nodes": self.affected_nodes,
            "affected_parts": self.affected_parts,
            "affected_pages_estimate": self.affected_pages_estimate,
            "preconditions": self.preconditions,
            "invariants": self.invariants,
            "skipped_risks": self.skipped_risks,
        }
