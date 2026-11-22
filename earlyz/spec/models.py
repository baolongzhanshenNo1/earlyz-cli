"""早睡 (EarlyZ) 3.2 — FormatSpec 模型（§5）。"""

from __future__ import annotations

import enum
from dataclasses import dataclass, field
from typing import Any, Optional

IR_SCHEMA_VERSION = "3.2"


class SourcePriority(enum.Enum):
    """规则来源优先级（§5.4，仅模板编制阶段）。"""

    user_confirmed = 5
    written_standard = 4
    template_ooxml = 3
    sample_paragraph = 2
    inferred = 1


@dataclass
class SpecProvenance:
    """一条规则的来源证据。"""

    source: str          # user_confirmed/written_standard/template_ooxml/sample_paragraph/inferred
    detail: str = ""
    artifact_id: Optional[str] = None

    @property
    def priority(self) -> int:
        try:
            return SourcePriority[self.source].value
        except KeyError:
            return SourcePriority.inferred.value


@dataclass
class FormatRule:
    """一条排版规则（§5.2）。"""

    rule_id: str
    selector: dict[str, Any]      # SemanticSelector
    expected: dict[str, Any]      # FormatProperties（dict 形式，便于 JSON 化）
    severity: str = "error"       # error/warning/info
    auto_fix: bool = True
    fix_strategy: str = "direct"
    risk_ceiling: str = "R2"
    provenance: Optional[SpecProvenance] = None
    confidence: float = 1.0


@dataclass
class Invariant:
    """必须保持的不变量（内容 hash / 结构引用）。"""

    name: str
    kind: str = "content"         # content/structure/relationship
    detail: str = ""


@dataclass
class RenderingPolicy:
    renderer: str = "word_preferred"   # word_preferred/fallback/none
    profiles: list[str] = field(default_factory=list)


@dataclass
class DeliveryPolicy:
    odd_page_chapters: bool = False
    blank_page_between_chapters: bool = False
    front_matter_roman: bool = False
    keep_header_footer: bool = True
    max_auto_risk: str = "R2"


@dataclass
class FormatSpec:
    """版本化、可审计的排版规范（§5.1）。"""

    spec_id: str
    name: str
    domain: str
    source_type: str              # uploaded_template_docx/specification_docx/built_in_spec
    schema_version: str = IR_SCHEMA_VERSION
    source_artifact_id: Optional[str] = None
    rules: list[FormatRule] = field(default_factory=list)
    example_residuals: list[str] = field(default_factory=list)  # 模板示例残留（规范化文本）
    invariants: list[Invariant] = field(default_factory=list)
    rendering_policy: RenderingPolicy = field(default_factory=RenderingPolicy)
    delivery_policy: DeliveryPolicy = field(default_factory=DeliveryPolicy)
    provenance: list[SpecProvenance] = field(default_factory=list)
    structure: dict = field(default_factory=dict)  # 目录/图表题/参考文献/节页码（模板抽取）

    def rule_by_id(self, rule_id: str) -> Optional[FormatRule]:
        for r in self.rules:
            if r.rule_id == rule_id:
                return r
        return None

    def to_dict(self) -> dict:
        return {
            "schema_version": self.schema_version,
            "spec_id": self.spec_id,
            "name": self.name,
            "domain": self.domain,
            "source_type": self.source_type,
            "rule_count": len(self.rules),
        }
