"""早睡 (EarlyZ) 3.2 — FormatSpec 三源编译器（§5.3/§5.4）。

运行时规范源互斥：一次任务只允许一个权威源
（uploaded_template_docx / specification_docx / built_in_spec），
禁止静默合并。模板编制阶段可离线对比多个来源，但必须输出冲突报告。

P0 实现范围：
- built_in_spec：JSON 文件 → FormatSpec；
- uploaded_template_docx / specification_docx：docx OOXML → 规则提取
  （docDefaults + 样式级规则）；完整模板编译器（语义角色映射、
  冲突消解）在 T4 后续增量中扩展。
"""

from __future__ import annotations

import json
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

from earlyz.spec.models import FormatRule, FormatSpec, SpecProvenance

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


class SpecSourceConflictError(Exception):
    """运行时规范源互斥被违反。"""


@dataclass
class CompilationResult:
    """编译输出（§8.4）。"""

    spec: Optional[FormatSpec]
    rule_count: int = 0
    conflicts: list[dict] = field(default_factory=list)
    unresolved: list[dict] = field(default_factory=list)
    provenance_summary: dict = field(default_factory=dict)

    def to_dict(self) -> dict:
        return {
            "spec_id": self.spec.spec_id if self.spec else None,
            "spec_version": self.spec.schema_version if self.spec else None,
            "rule_count": self.rule_count,
            "conflicts": self.conflicts,
            "unresolved": self.unresolved,
            "provenance_summary": self.provenance_summary,
        }


_ALLOWED_SOURCES = {"uploaded_template_docx", "specification_docx", "built_in_spec"}
_PACKS_DIR = Path(__file__).resolve().parent / "packs"


def resolve_built_in_spec(path_or_id: str) -> Path:
    """内置规范解析：ID（如 'thesis'/'contract'）→ packs 目录 JSON；否则按路径。"""
    p = Path(path_or_id)
    if p.exists():
        return p
    cand = _PACKS_DIR / f"{path_or_id}.json"
    if cand.exists():
        return cand
    cand = _PACKS_DIR / f"{path_or_id}_pack.json"
    if cand.exists():
        return cand
    raise FileNotFoundError(f"built-in spec not found: {path_or_id}")


def compile_format_spec(
    source_type: str,
    source_path_or_id: str,
    domain: str,
    strict: bool = True,
    secondary_source: Optional[str] = None,
) -> CompilationResult:
    """从唯一权威源编译 FormatSpec（§8.4 工具契约）。"""
    if secondary_source is not None:
        raise SpecSourceConflictError(
            f"spec source mutual exclusion violated: primary={source_type}, "
            f"secondary={secondary_source}"
        )
    if source_type not in _ALLOWED_SOURCES:
        raise ValueError(f"unknown source_type: {source_type} (allowed: {sorted(_ALLOWED_SOURCES)})")

    if source_type == "built_in_spec":
        result = _compile_built_in(source_path_or_id, domain)
    else:
        result = _compile_docx(source_type, source_path_or_id, domain, strict)

    # 注解层优先（2026-10-27）：模板作者把规范写在文本框/批注/正文说明里时，
    # 这些「显式声明」比从样例段落反推可靠，按角色合并进规则（只覆盖注解明确给出的属性）。
    if source_type != "built_in_spec":
        try:
            from earlyz.spec.annotations import (
                annotation_summary,
                extract_annotations,
                merge_annotations,
            )
            anns = extract_annotations(source_path_or_id)
            if anns:
                applied = merge_annotations(result.spec, anns)
                if not isinstance(result.spec.structure, dict):
                    result.spec.structure = {}
                result.spec.structure["annotation_summary"] = annotation_summary(anns)
                result.spec.structure["annotation_applied"] = applied
        except Exception:  # noqa: BLE001 — 注解解析失败不阻断编译
            pass

    # ADR-015/017（P2 建议层）：用户规则 overrides 合并（system < user）。
    # 确认过的规则提案发布后在此生效；无 overrides 原样返回。
    try:
        from earlyz.suggest.proposals import merge_spec_overrides
        merged = merge_spec_overrides(result.spec)
        if merged is not None:
            result.spec = merged
    except Exception:  # noqa: BLE001 — overrides 合并失败不阻断编译
        pass
    return result


# ────────────────────────────────────────────────────────────────
# built_in_spec（JSON）
# ────────────────────────────────────────────────────────────────

def _compile_built_in(path_or_id: str, domain: str) -> CompilationResult:
    p = resolve_built_in_spec(path_or_id)
    raw = json.loads(p.read_text(encoding="utf-8"))
    rules: list[FormatRule] = []
    for r in raw.get("rules", []):
        prov_raw = r.get("provenance", {})
        prov = SpecProvenance(
            source=prov_raw.get("source", "user_confirmed"),
            detail=prov_raw.get("detail", ""),
            artifact_id=prov_raw.get("artifact_id"),
        )
        rules.append(
            FormatRule(
                rule_id=r["rule_id"],
                selector=r.get("selector", {}),
                expected=r.get("expected", {}),
                severity=r.get("severity", "error"),
                auto_fix=r.get("auto_fix", True),
                fix_strategy=r.get("fix_strategy", "direct"),
                risk_ceiling=r.get("risk_ceiling", "R2"),
                provenance=prov,
                confidence=r.get("confidence", 1.0),
            )
        )
    spec = FormatSpec(
        spec_id=f"spec:{domain}-{p.stem}",
        name=raw.get("name", p.stem),
        domain=domain,
        source_type="built_in_spec",
        source_artifact_id=str(p),
        rules=rules,
    )
    return CompilationResult(spec=spec, rule_count=len(rules))


# ────────────────────────────────────────────────────────────────
# docx 模板 / 规范文档
# ────────────────────────────────────────────────────────────────

def _local(tag: str) -> str:
    return tag.split("}")[-1]


def _compile_docx(source_type: str, path: str, domain: str, strict: bool) -> CompilationResult:
    """docx 模板/规范文档 → FormatSpec（TSD 模板语义反编译）。

    2026-08-15 重写：旧实现只读 styles.xml（真实高校模板规范在示例段落
    直接格式里，styles.xml 只有辅助样式）→ 排版能力为零。现走
    decompiler.decompile_template：语义角色规则 + 节级页边距规则 +
    封面锚点规则，styles.xml 的样式定义作为样式型模板兜底样本源。
    """
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"docx not found: {p}")

    from earlyz.spec.decompiler import decompile_template

    spec = decompile_template(p, domain=domain)
    rules = spec.rules
    provenance_summary: dict[str, int] = {}
    for r in rules:
        key = ("cover_anchor" if "cover_anchor" in (r.selector or {})
               else r.selector.get("scope") or r.selector.get("semantic_role") or "other")
        provenance_summary[str(key)] = provenance_summary.get(str(key), 0) + 1

    unresolved = []
    if strict and not rules:
        unresolved.append({"detail": "no rules extracted from template"})
    return CompilationResult(
        spec=spec, rule_count=len(rules),
        unresolved=unresolved, provenance_summary=provenance_summary,
    )

