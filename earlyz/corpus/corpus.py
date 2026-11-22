"""早睡 (EarlyZ) 3.2 — 异常文档基准库 runner（Phase 2 验收：真实异常 DOCX 批量体检）。

每个样本执行：
    1. Package Gate      —— inspect_package（ZIP/ContentTypes/重复 entry）
    2. OOXML 分析        —— analyze_document（可解析性）
    3. 复杂对象完整性    —— 脚注/尾注引用、修订、批注、书签
    4. 交付门禁(静态)    —— validate_delivery(renderer="none")
    5. （可选 --word）   —— 真实 Word 渲染 → LEVEL_A / FAILED

输出逐样本结构化报告 + 汇总；每个失败样本必须给出结构化原因（§12 验收指标）。

manifest.json 形如：
    {
      "schema_version": "3.2",
      "docs": [
        {"file": "ok_plain.docx", "expect": "pass"},
        {"file": "fn_missing.docx", "expect": "fail", "expected_issue": "footnote"}
      ]
    }
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from earlyz.ir.analyzer import analyze_document
from earlyz.ir.complex_objects import (
    analyze_complex_objects,
    complex_object_gate_errors,
)
from earlyz.ir.package import inspect_package
from earlyz.validation.delivery import validate_delivery

MANIFEST = "manifest.json"


@dataclass
class CorpusDocResult:
    file: str
    status: str                      # pass / fail / error
    expected: str = "pass"
    expected_issue: Optional[str] = None
    issues: list[str] = field(default_factory=list)
    stats: dict = field(default_factory=dict)
    word_level: Optional[str] = None

    def to_dict(self) -> dict:
        return {
            "file": self.file,
            "status": self.status,
            "expected": self.expected,
            "expected_issue": self.expected_issue,
            "issues": self.issues,
            "stats": self.stats,
            "word_level": self.word_level,
        }


@dataclass
class CorpusReport:
    total: int = 0
    passed: int = 0
    failed: int = 0
    unexpected: int = 0              # 实际结果与 manifest 预期不符
    docs: list[CorpusDocResult] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "total": self.total,
            "passed": self.passed,
            "failed": self.failed,
            "unexpected": self.unexpected,
            "docs": [d.to_dict() for d in self.docs],
        }


def _load_manifest(corpus_dir: Path) -> dict:
    m = corpus_dir / MANIFEST
    if not m.exists():
        return {"schema_version": "3.2", "docs": [
            {"file": f.name, "expect": "pass"} for f in sorted(corpus_dir.glob("*.docx"))
        ]}
    return json.loads(m.read_text(encoding="utf-8"))


def check_docx(docx: Path, with_word: bool = False) -> CorpusDocResult:
    """单个样本体检。issues 收集所有失败原因（结构化，不静默）。"""
    res = CorpusDocResult(file=docx.name, status="pass")
    stats: dict = {}

    # 1. Package Gate
    try:
        ph = inspect_package(str(docx))
        stats["package_errors"] = len(ph.errors)
        stats["duplicate_entries"] = len(ph.duplicate_entries)
        res.issues.extend(ph.errors)
        if ph.duplicate_entries:
            res.issues.append(f"duplicate zip entries: {ph.duplicate_entries[:3]}")
    except Exception as e:  # noqa: BLE001
        res.issues.append(f"package inspect failed: {type(e).__name__}: {e}")
        res.status = "fail"
        return res

    # 2. OOXML 可解析性
    try:
        ir = analyze_document(str(docx))
        stats["paragraphs"] = len(ir.paragraphs)
        stats["tables"] = len(ir.tables)
        stats["sections"] = len(ir.sections)
    except Exception as e:  # noqa: BLE001
        res.issues.append(f"analyze failed: {type(e).__name__}: {e}")
        res.status = "fail"
        return res

    # 3. 复杂对象完整性（硬错误类）
    try:
        co = analyze_complex_objects(str(docx))
        stats["footnote_refs"] = co["footnotes"]["refs"]
        stats["endnote_refs"] = co["endnotes"]["refs"]
        stats["tracked_changes"] = co["revisions"]["tracked_changes"]
        stats["comment_refs"] = co["comments"]["comment_refs"]
        # 硬错误由下游 validate_delivery 门禁统一收集（避免重复上报）
    except Exception as e:  # noqa: BLE001
        res.issues.append(f"complex object analysis failed: {type(e).__name__}: {e}")

    # 4. 静态交付门禁
    try:
        rep = validate_delivery(str(docx), renderer="none")
        stats["validation_level"] = rep.level
        res.issues.extend(rep.hard_errors)
    except Exception as e:  # noqa: BLE001
        res.issues.append(f"delivery validation failed: {type(e).__name__}: {e}")

    # 5. 可选：真实 Word 渲染
    if with_word:
        try:
            from earlyz.validation.renderers import render_with_word

            info = render_with_word(str(docx))
            stats["word_pages"] = info.get("pages")
            res.word_level = "LEVEL_A_WORD_VERIFIED"
            if info.get("bbox_overflow"):
                res.issues.append(f"bbox overflow: {len(info['bbox_overflow'])} items")
        except Exception as e:  # noqa: BLE001
            res.word_level = "FAILED"
            res.issues.append(f"word render failed: {type(e).__name__}: {e}")

    if res.issues:
        res.status = "fail"
    res.issues = list(dict.fromkeys(res.issues))  # 去重（gate 与 delivery 可能同源）
    res.stats = stats
    return res


def run_corpus_check(corpus_dir: str | Path, with_word: bool = False) -> CorpusReport:
    """批量体检 + manifest 预期比对。"""
    base = Path(corpus_dir)
    manifest = _load_manifest(base)
    report = CorpusReport()
    for entry in manifest.get("docs", []):
        fname = entry.get("file", "")
        docx = base / fname
        if not docx.exists():
            r = CorpusDocResult(file=fname, status="error",
                                expected=entry.get("expect", "pass"),
                                issues=["file missing in corpus dir"])
        else:
            r = check_docx(docx, with_word=with_word)
        r.expected = entry.get("expect", "pass")
        r.expected_issue = entry.get("expected_issue")
        report.docs.append(r)
        report.total += 1
        if r.status == "pass":
            report.passed += 1
        else:
            report.failed += 1
        # 预期比对：状态不符 OR（预期 fail 且期望的缺陷类型未命中）→ unexpected
        if r.status != r.expected:
            report.unexpected += 1
        elif r.expected == "fail" and r.expected_issue and r.expected_issue not in " | ".join(r.issues):
            report.unexpected += 1
    return report
