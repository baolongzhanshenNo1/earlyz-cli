"""早睡 (EarlyZ) 3.2 — TransactionExecutor（§8.9，ADR-008）。

原子执行 RepairPlan：
- 不修改原文件：先复制工作副本，全部修改在 candidate 上进行；
- 逐 action 执行（handler 来自 ActionRegistry），执行后跑 verifier；
- 任一 hard failure（handler 异常 / verifier 失败 / 缺 handler）→ 回滚：
  用备份覆盖 candidate，禁止交付半成品；
- journal 记录每个 action 的前后 part hash、status、错误与回滚结果。

P0 语义：逐 action 读改写；批量一次写盘优化在 Writer 适配层（T10）处理。
"""

from __future__ import annotations

import json
import os
import shutil
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from earlyz.actions.registry import ActionRegistry
from earlyz.planner.models import RepairPlan


@dataclass
class TransactionResult:
    transaction_id: str
    status: str                    # applied / rolled_back / failed
    output_path: Optional[str] = None
    journal_path: Optional[str] = None
    backup_path: Optional[str] = None
    action_results: list[dict] = field(default_factory=list)
    hard_error_count: int = 0
    failure_detail: str = ""        # 回滚原因：动作名（#序号）：错误（2026-10-20）

    def to_dict(self) -> dict:
        return {
            "transaction_id": self.transaction_id,
            "status": self.status,
            "output_path": self.output_path,
            "journal_path": self.journal_path,
            "action_results": self.action_results,
            "hard_error_count": self.hard_error_count,
        }


def part_hash(docx_path: str, part: str) -> str:
    """part 内容 sha1（journal 用）。"""
    import hashlib
    with zipfile.ZipFile(docx_path) as z:
        return hashlib.sha1(z.read(part)).hexdigest()


def _freeze_paragraph_targets(plan: RepairPlan, candidate: Path) -> None:
    """Resolve every selector-carrying action against the same pre-mutation document.

    Paragraph formatting can change role classification (for example writing
    outlineLvl turns an article title into heading1). Without this pre-pass,
    later actions in the same plan reclassify a mutated document and target a
    different paragraph or find nothing. Stable text/style/object fingerprints
    keep the whole plan bound to its Analyze-time targets without naked indexes.

    §21.5：覆盖所有带 selector 的段落级 action（set_paragraph_format /
    set_paragraph_tabs / set_paragraph_border_shading / set_run_format ——
    按参数形态判定，天然覆盖未来新增，禁止按 action 名枚举）。
    """
    targets = [a for a in plan.actions
               if (a.params or {}).get("selector")
               and not (a.params or {}).get("_resolved_paragraphs")]
    if not targets:
        return
    from earlyz.actions.new_actions import (
        _capture_resolved_paragraphs, _locate_paragraphs, _read_part,
        _read_styles_map,
    )
    root = _read_part(candidate, "word/document.xml")
    styles = _read_styles_map(candidate)
    for action in targets:
        selector = (action.params or {}).get("selector") or {}
        paras = _locate_paragraphs(root, selector, styles)
        if paras:
            action.params["_resolved_paragraphs"] = _capture_resolved_paragraphs(root, paras)


def _journal_safe_info(info: dict, *, list_cap: int = 50) -> dict:
    """journal 化的 action 返回值：保留标量 + **有上限的标量/扁平字典列表**。

    2026-09-28：原先只留标量 → `shrink_details`（列表）被静默丢掉 → 编译器读不到
    → 用户永远收不到「公式太长已自动缩小字号」的说明（违反用户明确要求：
    「然后给用户说明情况」）。列表放进 journal 有体积风险，故设上限。
    """
    out: dict = {}
    for k, v in info.items():
        if isinstance(v, (int, float, str, bool)) or v is None:
            out[k] = v
        elif isinstance(v, list):
            rows = []
            for item in v[:list_cap]:
                if isinstance(item, (int, float, str, bool)) or item is None:
                    rows.append(item)
                elif isinstance(item, dict) and all(
                        isinstance(x, (int, float, str, bool)) or x is None
                        for x in item.values()):
                    rows.append({str(kk): vv for kk, vv in item.items()})
            if rows:
                out[k] = rows
    return out


class TransactionExecutor:
    def __init__(self, registry: ActionRegistry, workdir: str | Path):
        self.registry = registry
        self.workdir = Path(workdir)
        self.workdir.mkdir(parents=True, exist_ok=True)

    def execute(self, plan: RepairPlan, input_path: str | Path,
                progress_cb=None) -> TransactionResult:
        import hashlib
        import uuid

        src = Path(input_path)
        # Windows 文件名不允许冒号（tx:xxx 会被解析为 NTFS 数据流）→ 用下划线
        tx_id = f"tx_{uuid.uuid4().hex[:12]}"
        candidate = self.workdir / f"{tx_id}-candidate.docx"
        backup = self.workdir / f"{tx_id}-backup.docx"
        journal_path = self.workdir / f"{tx_id}-journal.json"

        # 事务从源文件复制工作副本开始；源文件永不直接修改
        shutil.copy2(src, candidate)
        shutil.copy2(src, backup)
        _freeze_paragraph_targets(plan, candidate)

        journal: dict = {
            "transaction_id": tx_id,
            "plan_id": plan.plan_id,
            "source": str(src),
            "status": "applied",
            "actions": [],
            "rollback": {"needed": False, "restored": False},
        }

        result = TransactionResult(
            transaction_id=tx_id,
            status="applied",
            output_path=str(candidate),
            journal_path=str(journal_path),
            backup_path=str(backup),
        )

        for i, action in enumerate(plan.actions):
            if progress_cb is not None:
                try:
                    progress_cb({
                        "type": "action",
                        "action": action.action,
                        "index": i,
                        "total": len(plan.actions),
                        "params_keys": list((action.params or {}).keys())[:8],
                    })
                except Exception:  # noqa: BLE001 — 进度回调异常不阻断执行
                    pass
            entry: dict = {
                "index": i,
                "action": action.action,
                "params": action.params,
                "status": "ok",
                "before_hash": None,
                "after_hash": None,
                "error": None,
            }
            handler = self.registry.handler(action.action)
            if handler is None:
                entry["status"] = "failed"
                entry["error"] = f"UNKNOWN_ACTION: {action.action} (no handler)"
                result.hard_error_count += 1
                journal["actions"].append(entry)
                # 规划阶段失败：必须进 action_results，否则上层只看到
                # 「transaction rolled back」不知道哪一步（2026-10-20）
                result.action_results.append(entry)
                result.failure_detail = (
                    f"{entry.get('action')}（#{entry.get('index')}）：{entry.get('error')}")
                self._rollback(journal, candidate, backup, result)
                self._write_journal(journal_path, journal)
                return result

            try:
                entry["before_hash"] = part_hash(str(candidate), "word/document.xml")
                info = handler(str(candidate), action.params)
                entry["after_hash"] = part_hash(str(candidate), "word/document.xml")
                # 动作主动跳过（源文档缺该角色，如 affiliation）→ 不跑验证器，
                # 否则验证器必然 False 导致整份回滚（2026-10-20 实测）
                if isinstance(info, dict) and info.get("skipped"):
                    entry["skipped"] = info.get("skipped")
            except Exception as e:  # noqa: BLE001 — handler 异常一律视为 hard failure
                entry["status"] = "failed"
                entry["error"] = f"{type(e).__name__}: {e}"
                result.hard_error_count += 1
                journal["actions"].append(entry)
                # 规划阶段失败：必须进 action_results，否则上层只看到
                # 「transaction rolled back」不知道哪一步（2026-10-20）
                result.action_results.append(entry)
                result.failure_detail = (
                    f"{entry.get('action')}（#{entry.get('index')}）：{entry.get('error')}")
                self._rollback(journal, candidate, backup, result)
                self._write_journal(journal_path, journal)
                return result

            if os.environ.get("EARLYZ_TRACE_MATH"):
                try:
                    import zipfile as _zf
                    from lxml import etree as _et
                    _W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
                    _M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
                    _x = _et.fromstring(_zf.ZipFile(str(candidate)).read("word/document.xml"))
                    _n = len(list(_x.iter(f"{_M}oMath")))
                    print(f"[MATH] {action.action}: oMath={_n}", flush=True)
                except Exception as _e:  # noqa: BLE001
                    print(f"[MATH] {action.action}: probe failed {_e}", flush=True)

            if os.environ.get("EARLYZ_TRACE_MATH"):
                try:
                    import zipfile as _zf
                    from lxml import etree as _et
                    _M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
                    _x = _et.fromstring(_zf.ZipFile(str(candidate)).read("word/document.xml"))
                    print(f"[MATH-POST] {action.action}: oMath={len(list(_x.iter(f'{_M}oMath')))}", flush=True)
                except Exception as _e:  # noqa: BLE001
                    print(f"[MATH-POST] {action.action}: probe failed {_e}", flush=True)

            verifier = self.registry.verifier(action.action)
            if verifier is not None and not entry.get("skipped"):
                try:
                    ok = verifier(str(candidate), action.params)
                except Exception as e:  # noqa: BLE001
                    ok = False
                    entry["error"] = f"verifier exception: {e}"
                if not ok:
                    entry["status"] = "failed"
                    entry["error"] = entry.get("error") or "verifier returned False"
                    result.hard_error_count += 1
                    journal["actions"].append(entry)
                    # post-action 失败也要进 action_results（2026-10-20）
                    result.action_results.append(entry)
                    result.failure_detail = (
                        f"{entry.get('action')}（#{entry.get('index')}）：{entry.get('error')}")
                    self._rollback(journal, candidate, backup, result)
                    self._write_journal(journal_path, journal)
                    return result

            journal["actions"].append(entry)
            result.action_results.append(entry)

        self._write_journal(journal_path, journal)
        return result

    def execute_post_actions(
        self,
        plan: RepairPlan,
        input_path: str | Path,
        tx_id: str,
        candidate: Path,
        backup: Path,
        journal_path: Path,
    ) -> TransactionResult:
        """在既有事务上追加结构修复 action（2026-08-17 路径统一）。

        compile 主链原先把 structural 修复（footer/section_refs/toc/表格/
        题注）硬编码在 format_compiler 里直接调 handler——与 planner 路径
        形成双通道：无 journal 记录、无 verifier 门禁、失败不回滚（用户
        从 journal 看"只有旧 action"即此观察断层的真相）。

        统一后：所有 action 走同一 execute 循环（journal + verifier +
        回滚），只是分两段输入——planner 段（lint 驱动）与 structural 段
        （compile 显式）。产物可追溯性=每次改动都有账可查。
        """
        import hashlib  # noqa: F401 — 与 execute 保持导入一致

        journal: dict | None = None
        if journal_path.exists():
            try:
                journal = json.loads(journal_path.read_text(encoding="utf-8"))
            except Exception:  # noqa: BLE001 — journal 损坏则重建
                journal = None
        if not journal:
            journal = {
                "transaction_id": tx_id,
                "plan_id": plan.plan_id,
                "source": str(input_path),
                "status": "applied",
                "actions": [],
                "rollback": {"needed": False, "restored": False},
            }

        result = TransactionResult(
            transaction_id=tx_id,
            status="applied",
            output_path=str(candidate),
            journal_path=str(journal_path),
            backup_path=str(backup),
        )

        _freeze_paragraph_targets(plan, candidate)
        start_index = len(journal["actions"])
        for i, action in enumerate(plan.actions):
            entry: dict = {
                "index": start_index + i,
                "action": action.action,
                "params": action.params,
                "phase": "structural",
                "status": "ok",
                "before_hash": None,
                "after_hash": None,
                "error": None,
            }
            handler = self.registry.handler(action.action)
            if handler is None:
                entry["status"] = "failed"
                entry["error"] = f"UNKNOWN_ACTION: {action.action} (no handler)"
                result.hard_error_count += 1
                journal["actions"].append(entry)
                # 规划阶段失败：必须进 action_results，否则上层只看到
                # 「transaction rolled back」不知道哪一步（2026-10-20）
                result.action_results.append(entry)
                result.failure_detail = (
                    f"{entry.get('action')}（#{entry.get('index')}）：{entry.get('error')}")
                self._rollback(journal, candidate, backup, result)
                self._write_journal(journal_path, journal)
                return result

            try:
                entry["before_hash"] = part_hash(str(candidate), "word/document.xml")
                info = handler(str(candidate), action.params)
                entry["after_hash"] = part_hash(str(candidate), "word/document.xml")
                if os.environ.get("EARLYZ_TRACE_MATH"):
                    try:
                        import zipfile as _zf
                        from lxml import etree as _et
                        _M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
                        _x = _et.fromstring(_zf.ZipFile(str(candidate)).read("word/document.xml"))
                        print(f"[MATH2] #{entry.get('index')} {action.action}: "
                              f"oMath={len(list(_x.iter(f'{_M}oMath')))}", flush=True)
                    except Exception as _e:  # noqa: BLE001
                        print(f"[MATH2] #{entry.get('index')} {action.action}: probe fail {_e}", flush=True)
                if isinstance(info, dict):
                    entry["info"] = _journal_safe_info(info)
            except Exception as e:  # noqa: BLE001 — structural 失败降级为警告
                # 与 planner 段不同：结构修复失败不回滚整个事务（格式修复
                # 已应用，回滚会丢掉全部成果）；记录失败继续下一 action。
                entry["status"] = "failed"
                entry["error"] = f"{type(e).__name__}: {e}"
                journal["actions"].append(entry)
                continue

            verifier = self.registry.verifier(action.action)
            if verifier is not None:
                try:
                    ok = verifier(str(candidate), action.params)
                except Exception as e:  # noqa: BLE001
                    ok = False
                    entry["error"] = f"verifier exception: {e}"
                if not ok:
                    entry["status"] = "failed"
                    entry["error"] = entry.get("error") or "verifier returned False"
                    journal["actions"].append(entry)
                    # 失败动作也要进 action_results：否则上层只能看到
                    # 「transaction rolled back」，无法定位是哪一步（2026-10-20）
                    result.action_results.append(entry)
                    result.failure_detail = (
                        f"{entry.get('action')}（#{entry.get('index')}）：{entry['error']}")
                    continue

            journal["actions"].append(entry)
            result.action_results.append(entry)

        self._write_journal(journal_path, journal)
        return result

    def _rollback(self, journal: dict, candidate: Path, backup: Path, result: TransactionResult) -> None:
        shutil.copy2(backup, candidate)
        journal["status"] = "rolled_back"
        journal["rollback"] = {"needed": True, "restored": True}
        result.status = "rolled_back"

    def _write_journal(self, path: Path, journal: dict) -> None:
        path.write_text(json.dumps(journal, ensure_ascii=False, indent=2), encoding="utf-8")


def execute_plan(
    plan: RepairPlan,
    registry: ActionRegistry,
    input_path: str | Path,
    workdir: str | Path,
) -> TransactionResult:
    """函数式入口（§8.9 工具契约）。"""
    return TransactionExecutor(registry, workdir).execute(plan, input_path)
