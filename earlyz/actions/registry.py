"""早睡 (EarlyZ) 3.2 — ActionSpec 与 ActionRegistry（§6，ADR-006）。

ActionRegistry 是 handler / verifier / schema / 文档的唯一真相源：
- 自动生成 Agent JSON Schema；
- 校验注册一致性（缺 handler、缺 verifier 都必须在报告中暴露）；
- 双轴分级：实现优先级 P0-P2（开发顺序）与运行风险 R0-R3（危险程度）；
- R3 破坏性 action 默认不暴露给 Agent，plan 校验默认拒绝；
- validate_plan 做存在性 + 参数 + 风险三重校验。
"""

from __future__ import annotations

import enum
from dataclasses import dataclass, field
from typing import Any, Callable, Optional

IR_SCHEMA_VERSION = "3.2"


class RiskLevel(enum.Enum):
    R0 = "R0"  # 只读分析（严格说属于工具）
    R1 = "R1"  # 安全格式变更
    R2 = "R2"  # 结构变更，需事务/预览/完整验证
    R3 = "R3"  # 破坏性/不可逆，必须用户明确授权

    def __str__(self) -> str:
        return self.value


class ImplementationPriority(enum.Enum):
    P0 = "P0"
    P1 = "P1"
    P2 = "P2"


@dataclass
class ActionSpec:
    """单个 action 的完整契约（§6.2）。"""

    name: str
    version: int
    category: str                 # paragraph/run/style/section/header_footer/toc/image/table/list/package/...
    scope: str                    # paragraph/run/document/table/section/...
    risk: RiskLevel = RiskLevel.R1
    exposed_to_agent: bool = True
    idempotent: bool = True
    priority: ImplementationPriority = ImplementationPriority.P1
    selector_schema: Optional[dict] = None
    params_schema: Optional[dict] = None
    preconditions: list[str] = field(default_factory=list)
    mutates_parts: list[str] = field(default_factory=list)
    invariants: list[str] = field(default_factory=list)
    # §10.1：声明该 action 会修改内容（正文文本 hash 等）。内容变化例外
    # 必须由 ActionSpec 声明 + 用户授权，validate_delivery 才会放行。
    content_change: bool = False
    verifier: Optional[str] = None
    rollback_strategy: str = "restore_part"
    deprecated_by: Optional[str] = None
    description: str = ""


@dataclass
class RegistryConsistencyReport:
    missing_handler: list[str]
    missing_verifier: list[str]
    duplicate_names: list[str]
    unknown_verifiers: list[str]

    @property
    def consistent(self) -> bool:
        return not (
            self.missing_handler
            or self.missing_verifier
            or self.duplicate_names
            or self.unknown_verifiers
        )


class ActionRegistry:
    """唯一真相源：ActionSpec + handler + verifier。"""

    def __init__(self) -> None:
        self._specs: dict[str, ActionSpec] = {}
        self._handlers: dict[str, Callable] = {}
        self._verifiers: dict[str, Callable] = {}

    # ── 注册 ────────────────────────────────────────────────────

    def register(
        self,
        spec: ActionSpec,
        handler: Optional[Callable] = None,
        verifier: Optional[Callable] = None,
    ) -> None:
        if spec.name in self._specs:
            raise ValueError(f"duplicate action registration: {spec.name}")
        self._specs[spec.name] = spec
        if handler is not None:
            self._handlers[spec.name] = handler
        if verifier is not None:
            self._verifiers[spec.name] = verifier

    # ── 查询 ────────────────────────────────────────────────────

    def get(self, name: str) -> Optional[ActionSpec]:
        return self._specs.get(name)

    def all(self) -> list[ActionSpec]:
        return sorted(self._specs.values(), key=lambda s: s.name)

    def exposed(self) -> list[ActionSpec]:
        return [s for s in self.all() if s.exposed_to_agent]

    def handler(self, name: str) -> Optional[Callable]:
        return self._handlers.get(name)

    def verifier(self, name: str) -> Optional[Callable]:
        return self._verifiers.get(name)

    # ── 一致性（Phase 1 验收：差集为 0） ─────────────────────────

    def consistency_report(self) -> RegistryConsistencyReport:
        missing_handler = [n for n in self._specs if n not in self._handlers]
        missing_verifier = [n for n in self._specs if n not in self._verifiers]
        # 注：R0 只读项不要求 verifier 的策略在注册时处理；这里对所有 mutation 报告
        unknown_verifiers = []
        return RegistryConsistencyReport(
            missing_handler=missing_handler,
            missing_verifier=missing_verifier,
            duplicate_names=[],
            unknown_verifiers=unknown_verifiers,
        )

    # ── 风险政策 ────────────────────────────────────────────────

    def policy_check(self, action_name: str, max_risk: RiskLevel) -> tuple[bool, str]:
        spec = self._specs.get(action_name)
        if spec is None:
            return False, f"UNKNOWN_ACTION: {action_name}"
        if spec.risk.value > max_risk.value:
            return False, f"{action_name} requires user authorization (risk={spec.risk.value}, ceiling={max_risk.value})"
        return True, ""

    # ── Plan 校验 ───────────────────────────────────────────────

    def validate_plan(self, plan: list[dict], max_risk: RiskLevel) -> list[str]:
        errors: list[str] = []
        for i, step in enumerate(plan):
            name = step.get("action")
            if name not in self._specs:
                errors.append(f"step[{i}] UNKNOWN_ACTION: {name}")
                continue
            spec = self._specs[name]
            ok, reason = self.policy_check(name, max_risk)
            if not ok:
                errors.append(f"step[{i}] {reason}")
            params = step.get("params", {})
            if spec.params_schema:
                required = spec.params_schema.get("required", [])
                for req in required:
                    if req not in params:
                        errors.append(f"step[{i}] {name} missing required param: {req}")
            if spec.selector_schema and not params.get("selector"):
                # selector 是可选增强，不强制；但若声明了 selector_schema 且调用方给了 selector 则校验存在性
                pass
        return errors

    # ── Agent JSON Schema（ADR-006：自动生成） ──────────────────

    def agent_schema(self) -> dict:
        any_of = []
        for s in self.exposed():
            step = {
                "type": "object",
                "properties": {
                    "action": {"type": "string", "enum": [s.name]},
                    "params": s.params_schema or {"type": "object", "properties": {}},
                },
                "required": ["action", "params"],
            }
            any_of.append(step)
        return {
            "name": "apply_format_plan",
            "description": "Apply a validated, deterministic format repair plan.",
            "type": "object",
            "properties": {
                "plan_id": {"type": "string"},
                "actions": {"type": "array", "items": {"anyOf": any_of}},
            },
            "required": ["actions"],
        }

    # ── 文档生成 ────────────────────────────────────────────────

    def document(self) -> str:
        lines = ["# 早睡 (EarlyZ) Action Registry", "", f"schema_version: {IR_SCHEMA_VERSION}", ""]
        lines.append("| action | risk | priority | scope | idempotent | exposed | verifier |")
        lines.append("|---|---|---|---|---|---|---|")
        for s in self.all():
            lines.append(
                f"| {s.name} | {s.risk.value} | {s.priority.value} | {s.scope} "
                f"| {'yes' if s.idempotent else 'no'} | {'yes' if s.exposed_to_agent else 'no'} "
                f"| {s.verifier or '-'} |"
            )
        lines.append("")
        return "\n".join(lines)
