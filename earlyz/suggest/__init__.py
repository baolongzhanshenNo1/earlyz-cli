"""早睡 (EarlyZ) 3.2 — P2 建议层种子（§22.1 建议层 + ADR-015）。

- diffs：确定性 OOXML diff 引擎（产物 vs 用户修正版，稳定指纹对齐）；
- attribution：差异归因（engine_bug / user_deliberate / rule_override / template_gap）；
- proposals：RuleProposal 生命周期（pending→confirmed/rejected，确认后发布 user
  overrides；编译链 merge，system < user）。
"""

from .attribution import Attribution, attribute_diffs
from .diffs import FormatDiff, diff_documents
from .proposals import (
    RuleProposal,
    confirm_proposal,
    create_proposal,
    load_overrides,
    load_proposal,
    merge_spec_overrides,
    publish_overrides,
    reject_proposal,
    save_proposal,
)

__all__ = [
    "Attribution", "attribute_diffs",
    "FormatDiff", "diff_documents",
    "RuleProposal", "confirm_proposal", "create_proposal",
    "load_overrides", "load_proposal", "merge_spec_overrides",
    "publish_overrides", "reject_proposal", "save_proposal",
]
