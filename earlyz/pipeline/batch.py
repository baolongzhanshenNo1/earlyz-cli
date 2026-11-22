"""早睡 (EarlyZ) 3.2 — batch_format：多文档批量排版（§6.1 工具面，企业场景）。

- 每文档独立事务（compile_document），单文档失败不阻断批次；
- 同一规范来源应用于全部文档（三源互斥由 compile_format_spec 保证）；
- 产物独立命名（批次号 + 序号），同一文档多次出现不覆盖；
- 输出 batch certificate（JSON 汇总报告）。
"""

from __future__ import annotations

import json
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from earlyz.pipeline.format_compiler import compile_document


@dataclass
class BatchItemResult:
    index: int
    input_path: str
    status: str                      # applied / failed
    output_path: Optional[str] = None
    lint_total: int = 0
    plan_actions: int = 0
    transaction_status: Optional[str] = None
    validation_level: Optional[str] = None
    content_preserved: Optional[bool] = None
    hard_errors: list = field(default_factory=list)
    error: Optional[str] = None

    def to_dict(self) -> dict:
        return {
            "index": self.index,
            "input_path": self.input_path,
            "status": self.status,
            "output_path": self.output_path,
            "lint_total": self.lint_total,
            "plan_actions": self.plan_actions,
            "transaction_status": self.transaction_status,
            "validation_level": self.validation_level,
            "content_preserved": self.content_preserved,
            "hard_errors": self.hard_errors,
            "error": self.error,
        }


@dataclass
class BatchResult:
    total: int
    applied: int
    failed: int
    items: list[BatchItemResult]
    report_path: Optional[str] = None

    def to_dict(self) -> dict:
        return {
            "total": self.total,
            "applied": self.applied,
            "failed": self.failed,
            "report_path": self.report_path,
            "items": [i.to_dict() for i in self.items],
        }


def run_batch_format(
    input_paths: list[str | Path],
    spec_source_type: str,
    spec_source: str,
    domain: str = "general",
    workdir: str | Path = ".",
    print_mode: bool = False,
) -> BatchResult:
    """批量编译：逐文档 analyze→lint→plan→apply→validate。"""
    if not input_paths:
        raise ValueError("批量排版需要至少一个输入文档")
    work = Path(workdir)
    batch_dir = work / f"batch_{uuid.uuid4().hex[:8]}"
    batch_dir.mkdir(parents=True, exist_ok=True)

    items: list[BatchItemResult] = []
    for i, doc in enumerate(input_paths):
        item = BatchItemResult(index=i, input_path=str(doc), status="pending")
        try:
            res = compile_document(
                doc,
                spec_source_type,
                spec_source,
                domain=domain,
                workdir=batch_dir / f"doc{i}",
                print_mode=print_mode,
            )
            d = res.to_dict()
            item.transaction_status = d["transaction_status"]
            item.lint_total = d["lint_total"]
            item.plan_actions = d["plan_actions"]
            if res.validation:
                item.validation_level = d["validation_level"]
                item.content_preserved = bool(res.validation.content_preserved)
                item.hard_errors = res.validation.hard_errors
            if d["transaction_status"] == "applied" and d["passed"]:
                item.status = "applied"
                item.output_path = d["artifacts"].get("print_output") or d["artifacts"].get("output_path")
            else:
                item.status = "failed"
                item.error = "transaction rolled back" if d["transaction_status"] != "applied" else "validation failed"
        except Exception as e:
            item.status = "failed"
            item.error = str(e)[:500]
        items.append(item)

    applied = sum(1 for i in items if i.status == "applied")
    failed = sum(1 for i in items if i.status == "failed")

    result = BatchResult(total=len(items), applied=applied, failed=failed, items=items)
    report_path = batch_dir / "batch_report.json"
    report_path.write_text(
        json.dumps(result.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8"
    )
    result.report_path = str(report_path)
    return result
