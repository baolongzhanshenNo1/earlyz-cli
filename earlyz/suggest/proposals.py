"""P2 建议层种子：规则提案生命周期（ADR-015）。

用户修正**只生成规则提案**，确认后才发布新 FormatSpec 版本：
- 提案存储 <DATA_DIR>/proposals/<id>.json（append 记录，状态机 pending→confirmed/rejected）；
- 确认 → 发布 = 写 <DATA_DIR>/rules/<spec_id>_overrides.json（按 rule_id upsert）；
- 编译链（compilers.compile_format_spec）末尾 merge user overrides（system < user，ADR-017）；
- 证据链随提案永久保存（「为什么这么改」可溯源——产品护城河）。
"""

from __future__ import annotations

import json
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

try:
    from api.host.artifacts import DATA_DIR  # 测试身份（pytest：sys.path=仓库根）
except ImportError:
    from host.artifacts import DATA_DIR  # 运行时身份（api as root）

_RULES_DIR = Path(DATA_DIR) / "rules"
_PROPOSALS_DIR = Path(DATA_DIR) / "proposals"

VALID_STATUS = ("pending", "confirmed", "rejected")


@dataclass
class RuleProposal:
    """一条规则修正提案（ADR-015 载体）。"""

    proposal_id: str
    spec_id: str
    rule_patch: dict                       # FormatRule dict 形态（确认后 upsert）
    evidence: list[str] = field(default_factory=list)
    classification: str = "user_deliberate"
    status: str = "pending"                # pending/confirmed/rejected
    created_at: float = field(default_factory=time.time)
    resolved_at: Optional[float] = None

    def to_dict(self) -> dict:
        return {
            "proposal_id": self.proposal_id,
            "spec_id": self.spec_id,
            "rule_patch": self.rule_patch,
            "evidence": self.evidence,
            "classification": self.classification,
            "status": self.status,
            "created_at": self.created_at,
            "resolved_at": self.resolved_at,
        }

    @staticmethod
    def from_dict(d: dict) -> "RuleProposal":
        return RuleProposal(
            proposal_id=d["proposal_id"], spec_id=d["spec_id"],
            rule_patch=d.get("rule_patch") or {},
            evidence=d.get("evidence") or [],
            classification=d.get("classification") or "user_deliberate",
            status=d.get("status") or "pending",
            created_at=d.get("created_at") or time.time(),
            resolved_at=d.get("resolved_at"),
        )


def create_proposal(spec_id: str, rule_patch: dict, evidence: list[str],
                    classification: str = "user_deliberate",
                    proposal_id: str | None = None) -> RuleProposal:
    """新建提案（默认 pending；不落盘——调用方按需 save）。"""
    return RuleProposal(
        proposal_id=proposal_id or f"p_{uuid.uuid4().hex[:10]}",
        spec_id=spec_id, rule_patch=rule_patch, evidence=evidence,
        classification=classification,
    )


# ── 提案存储（尽力而为：路径不可写不抛） ───────────────────────

def _proposal_path(proposal_id: str) -> Path:
    if not proposal_id.replace("_", "").isalnum():
        raise ValueError("proposal_id 仅允许字母数字与下划线")
    return _PROPOSALS_DIR / f"{proposal_id}.json"


def save_proposal(p: RuleProposal) -> Path | None:
    try:
        _PROPOSALS_DIR.mkdir(parents=True, exist_ok=True)
        path = _proposal_path(p.proposal_id)
        path.write_text(json.dumps(p.to_dict(), ensure_ascii=False, indent=2),
                        encoding="utf-8")
        return path
    except Exception:  # noqa: BLE001 — 提案存储尽力而为
        return None


def load_proposal(proposal_id: str) -> RuleProposal | None:
    try:
        path = _proposal_path(proposal_id)
        if not path.exists():
            return None
        return RuleProposal.from_dict(json.loads(path.read_text(encoding="utf-8")))
    except Exception:  # noqa: BLE001
        return None


# ── 发布：确认后写 user overrides（system < user, ADR-017） ────

def confirm_proposal(p: RuleProposal) -> Path | None:
    """确认 → 发布（写 overrides，按 rule_id upsert）→ 更新提案状态。"""
    if p.status not in ("pending",):
        return None
    out = publish_overrides(p.spec_id, p.rule_patch)
    if out is None:
        return None
    p.status = "confirmed"
    p.resolved_at = time.time()
    save_proposal(p)
    return out


def reject_proposal(p: RuleProposal) -> Path | None:
    if p.status != "pending":
        return None
    p.status = "rejected"
    p.resolved_at = time.time()
    return save_proposal(p)


def _rules_file(spec_id: str) -> Path:
    safe = spec_id.replace("\\", "").replace("/", "").strip()
    if not safe:
        raise ValueError("spec_id 非法")
    return _RULES_DIR / f"{safe}_overrides.json"


def publish_overrides(spec_id: str, rule_patch: dict) -> Path | None:
    """把规则补丁 upsert 进 <DATA_DIR>/rules/<spec_id>_overrides.json。

    rule_id 相同 → 覆盖（用户新值优先）；不同 → 追加。
    """
    try:
        _RULES_DIR.mkdir(parents=True, exist_ok=True)
        path = _rules_file(spec_id)
        rules: list[dict] = []
        if path.exists():
            rules = json.loads(path.read_text(encoding="utf-8")) or []
        rule_id = (rule_patch or {}).get("rule_id")
        if not rule_id:
            return None
        rules = [r for r in rules if r.get("rule_id") != rule_id]
        rules.append(rule_patch)
        path.write_text(json.dumps(rules, ensure_ascii=False, indent=2),
                        encoding="utf-8")
        return path
    except Exception:  # noqa: BLE001
        return None


def load_overrides(spec_id: str) -> list[dict]:
    """读 user overrides 规则（无文件/损坏 → []；损坏不阻断编译，宁缺毋滥）。"""
    try:
        path = _rules_file(spec_id)
        if not path.exists():
            return []
        data = json.loads(path.read_text(encoding="utf-8"))
        if isinstance(data, list):
            return [r for r in data if isinstance(r, dict) and r.get("rule_id")]
        return []
    except Exception:  # noqa: BLE001
        return []


def merge_spec_overrides(spec) -> Optional[object]:
    """编译链接入点：spec 编译后合并 user overrides（ADR-017 system < user）。

    返回合并后的 spec；无 overrides 或合并失败返回 None（调用方保留原 spec）。
    破坏性最小：upsert 规则（by rule_id），不触碰 selector 之外的字段。
    """
    if spec is None or not getattr(spec, "spec_id", None):
        return None
    try:
        from earlyz.spec.models import FormatRule
        overrides = load_overrides(spec.spec_id)
        if not overrides:
            return None
        merged_any = False
        existing = {r.rule_id: i for i, r in enumerate(spec.rules)}
        for o in overrides:
            try:
                rule = FormatRule(
                    rule_id=o["rule_id"],
                    selector=o.get("selector") or {},
                    expected=o.get("expected") or {},
                    severity=o.get("severity") or "error",
                    auto_fix=bool(o.get("auto_fix", True)),
                    fix_strategy=o.get("fix_strategy") or "direct",
                    risk_ceiling=o.get("risk_ceiling") or "R2",
                )
            except (TypeError, KeyError, ValueError):
                continue
            idx = existing.get(rule.rule_id)
            if idx is not None:
                spec.rules[idx] = rule
            else:
                spec.rules.append(rule)
                existing[rule.rule_id] = len(spec.rules) - 1
            merged_any = True
        return spec if merged_any else None
    except Exception:  # noqa: BLE001
        return None
