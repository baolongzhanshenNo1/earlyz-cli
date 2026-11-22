"""P3 全自动循环（§22.6 循环调度设计，2026-08-25 拍板边界）。

一次触发 = 1 个 Turn；Turn 内 Step 链：DIAGNOSE → FIX → VERIFY →
（FAIL 再诊断 ≤3）→ REPORT → 用户最终确认（上层接收返回值处理）。

边界（§22.5.1）：
- R1/R2 自动直通；R3（破坏性/内容变化）不入自动批次（上层展示待确认）；
- LLM 决策 ≤4（DIAGNOSE 首诊 1 + REPORT 1 + 再诊断 ≤2）；FIX-VERIFY 迭代 ≤3；
- 超限 → degraded（上层转 P2 对话确认，保留诊断包）；
- actions 必须来自 ActionRegistry（未知 action 跳过+记录）；params 禁裸
  para_index（§4.4）；risk 由 registry 权威判定（非 LLM 自报）；
- LLM 永不直接改 OOXML（铁律 5）——本模块只调度，执行全部走 ActionRegistry。
"""

from __future__ import annotations

import json
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

from earlyz.actions.registry import RiskLevel

DIAGNOSE_SYSTEM_PROMPT = """你是「早睡」(EarlyZ) 自动修复诊断分析师（P3 循环）。输入是文档的结构化诊断指纹（如实事实，无文件名）与模板规则摘要。
你的任务：找出可自动修复的格式缺陷，输出严格 JSON，仅一个键：
{"findings": [{"classification": "engine_bug|template_mismatch|user_doc", "evidence": ["..."], "confidence": 0.9, "actions": [{"action": "<registry action 名>", "params": {...}}]}]}

铁律约束：
1. classification 只能是三种：engine_bug（引擎缺陷）/ template_mismatch（模板差异）/ user_doc（用户文档问题）；
2. action 名必须来自给定 action 清单，禁止自创；定位参数：para_index 必须配 text_snippet（防漂移；裸 para_index 会被拒绝），也可用 node_id/text_hash/semantic_role + occurrence；
3. 只引用指纹中存在的证据，禁止臆测；
4. 只建议修复性 action（R1/R2）；破坏性动作（删除/内容变化/R3）不要建议，留给人工；
5. 证据不足 → confidence < 0.8 且 classification=user_doc/engine_bug，actions 可为空。"""


@dataclass
class P3Config:
    max_llm_calls: int = 4      # DIAGNOSE 首诊 + REPORT + 再诊断 ≤4
    max_iter: int = 3           # FIX-VERIFY 迭代 ≤3
    risk_ceiling: RiskLevel = RiskLevel.R2


@dataclass
class P3Finding:
    """DIAGNOSE 输出的校验后形态（无效发现进 rejected，不执行）。"""

    classification: str
    evidence: list[str] = field(default_factory=list)
    confidence: float = 0.0
    actions: list[dict] = field(default_factory=list)      # 通过校验（R1/R2）
    rejected_actions: list[dict] = field(default_factory=list)  # 被拒（R3/未知/裸索引）


@dataclass
class P3Result:
    status: str                     # ok / degraded / no_findings / empty
    fixed_count: int = 0
    iterations: int = 0
    llm_calls: int = 0
    findings: list[P3Finding] = field(default_factory=list)
    rejected: list[dict] = field(default_factory=list)
    degraded_reason: str = ""
    report_text: str = ""           # REPORT（LLM 产出或降级文案）
    verified_ok: bool = False
    elapsed_s: float = 0.0
    # 2026-08-26 task_events 粒度：循环 Step 轨迹（diagnose/fix/verify/report × iter）
    steps: list[dict] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "status": self.status,
            "fixed_count": self.fixed_count,
            "iterations": self.iterations,
            "llm_calls": self.llm_calls,
            "findings": [
                {"classification": f.classification, "confidence": f.confidence,
                 "evidence": f.evidence, "actions": f.actions,
                 "rejected_actions": f.rejected_actions}
                for f in self.findings
            ],
            "rejected": self.rejected,
            "degraded_reason": self.degraded_reason,
            "report_text": self.report_text,
            "verified_ok": self.verified_ok,
            "steps": self.steps,
            "elapsed_s": round(self.elapsed_s, 1),
        }


# ── DIAGNOSE 契约：LLM 调用（注入点） ─────────────────────────

def default_llm_call(sys_prompt: str, user_prompt: str, *, timeout_s: int = 120) -> str:
    """生产 LLM 调用（LLMClient.chat + JSON 围栏容忍）；测试注入 fake。"""
    import sys
    from pathlib import Path as _P

    api_dir = _P(__file__).resolve().parents[2] / "api"
    if str(api_dir) not in sys.path:
        sys.path.insert(0, str(api_dir))
    from llm_client import LLMClient

    msg = LLMClient(timeout=timeout_s).chat(
        [{"role": "system", "content": sys_prompt},
         {"role": "user", "content": user_prompt}],
        temperature=0, model=None)
    return (msg or {}).get("content", "") or ""


def _extract_json(text: str) -> Optional[dict]:
    import re
    m = re.search(r"\{.*\}", text, re.S)
    if not m:
        return None
    try:
        return json.loads(m.group(0))
    except json.JSONDecodeError:
        return None


# ── 校验：action 白名单 / 禁裸索引 / R3 拦截 ─────────────────

def _risk_of(registry, action_name: str) -> Optional[RiskLevel]:
    try:
        spec = registry.get(action_name)
        return spec.risk if spec else None
    except Exception:  # noqa: BLE001
        return None


def validate_findings(findings: list[dict], registry, spec=None,
                      ceiling: RiskLevel = RiskLevel.R2) -> tuple[list[P3Finding], list[dict]]:
    """校验 DIAGNOSE 输出（§22.6 #6 铁律约束）。返回 (通过, 全量 rejected 记录)。"""
    out: list[P3Finding] = []
    rejected_all: list[dict] = []
    for f in findings or []:
        cls = str(f.get("classification") or "")
        if cls not in ("engine_bug", "template_mismatch", "user_doc"):
            rejected_all.append({"reason": "invalid_classification", "raw": f})
            continue
        f_obj = P3Finding(
            classification=cls,
            evidence=[str(e) for e in (f.get("evidence") or [])],
            confidence=float(f.get("confidence") or 0.0),
        )
        for a in f.get("actions") or []:
            name = (a or {}).get("action") or ""
            params = (a or {}).get("params") or {}
            if not name:
                rejected_all.append({"reason": "missing_action", "raw": a})
                f_obj.rejected_actions.append(a)
                continue
            risk = _risk_of(registry, name)
            if risk is None:
                rejected_all.append({"reason": "unknown_action", "action": name})
                f_obj.rejected_actions.append(a)
                continue
            if risk.value > ceiling.value:  # R3 拦截（§22.5.1 授权边界）
                rejected_all.append({"reason": "risk_exceeds_ceiling", "action": name,
                                     "risk": risk.value})
                f_obj.rejected_actions.append(a)
                continue
            if "para_index" in params and not params.get("text_snippet"):
                rejected_all.append({"reason": "naked_para_index", "action": name})
                f_obj.rejected_actions.append(a)
                continue
            f_obj.actions.append({"action": name, "params": params})
        out.append(f_obj)
    return out, rejected_all


# ── Step 执行 ────────────────────────────────────────────────

def _execute_fix(actions: list[dict], docx_path: str | Path, registry,
                 workdir: str | Path) -> tuple[list[dict], int]:
    """FIX：ActionRegistry handler 批次执行（R1/R2），逐 action verifier 门禁。

    返回 (执行记录, 成功数)。失败 action 记录不中断批次（与 executor 结构段
    语义一致：格式修复推进优于全回滚——事务回滚由上层 execute 语义负责）。
    """
    from earlyz.transaction.executor import part_hash  # noqa: F401

    records: list[dict] = []
    ok = 0
    for a in actions:
        name = a["action"]
        params = a["params"]
        entry = {"action": name, "status": "ok"}
        try:
            handler = registry.handler(name)
            if handler is None:
                entry.update({"status": "failed", "error": "no_handler"})
            else:
                handler(str(docx_path), params)
                verifier = registry.verifier(name)
                if verifier is not None:
                    if not verifier(str(docx_path), params):
                        entry.update({"status": "failed", "error": "verifier_false"})
                    else:
                        ok += 1
                else:
                    ok += 1
        except Exception as e:  # noqa: BLE001
            entry.update({"status": "failed", "error": f"{type(e).__name__}: {e}"})
        records.append(entry)
    return records, ok


def _verify_static(docx_path: str | Path) -> list[str]:
    """VERIFY：静态硬门禁（package/复杂对象/域完整性/字段）——无 Word 渲染。"""
    errors: list[str] = []
    try:
        from earlyz.ir.package import inspect_package
        ph = inspect_package(str(docx_path))
        errors.extend(ph.errors)
    except Exception as e:  # noqa: BLE001
        errors.append(f"package: {type(e).__name__}: {e}")
    try:
        from earlyz.ir.complex_objects import complex_object_gate_errors
        errors.extend(complex_object_gate_errors(str(docx_path)))
    except Exception:  # noqa: BLE001
        pass
    try:
        from earlyz.ir.analyzer import analyze_document
        ir = analyze_document(str(docx_path))
        for f in ir.fields:
            if not f.complete:
                errors.append(f"broken field: {f.node_id}")
    except Exception as e:  # noqa: BLE001
        errors.append(f"analyze: {type(e).__name__}: {e}")
    return errors


def _report(llm_call, findings: list[P3Finding], fixed_count: int,
            rejected: list[dict], verified_ok: bool) -> str:
    """REPORT：LLM 证据链可读化（1 次调用）；失败回退模板文案。"""
    budget = {
        "fixed_count": fixed_count,
        "verified_ok": verified_ok,
        "findings": [{"classification": f.classification, "evidence": f.evidence,
                      "actions": f.actions} for f in findings],
        "rejected": rejected,
    }
    sys_prompt = (
        "你是「早睡」(EarlyZ) 自动修复报告员。输入是 P3 循环的执行结果（严格事实）。"
        "输出一段面向用户的中文总结：已修复什么（证据）、验证结果、被拒绝/未处理的"
        "动作（R3 或无效）及原因。只引用输入中存在的证据，禁止臆测。150 字以内。"
    )
    try:
        content = llm_call(sys_prompt, json.dumps(budget, ensure_ascii=False, default=str))
        text = (content or "").strip()
        if text:
            return text
    except Exception:  # noqa: BLE001
        pass
    # 回退文案（"不静默"铁律）
    parts = [f"本次自动修复完成 {fixed_count} 项" if fixed_count else "本次未产生自动修复"]
    if not verified_ok:
        parts.append("验证未通过（详见诊断包）")
    if rejected:
        parts.append(f"{len(rejected)} 项动作未自动执行（R3/无效/未知）")
    return "；".join(parts) + "。"


def run_p3_loop(docx_path: str | Path, diagnose_payload: dict,
                registry, *, spec=None, config: P3Config | None = None,
                llm_call: Callable | None = None,
                workdir: str | Path | None = None) -> P3Result:
    """P3 循环入口（§22.6 状态机；LLM/registry 注入，纯逻辑可测）。"""
    t0 = time.time()
    cfg = config or P3Config()
    llm_call = llm_call or default_llm_call
    result = P3Result(status="empty")
    llm_remaining = cfg.max_llm_calls
    docx = Path(docx_path)
    work_dir = Path(workdir) if workdir else docx.parent

    def _diagnose(payload: dict) -> tuple[list[P3Finding], list[dict]]:
        nonlocal llm_remaining
        if llm_remaining <= 0:
            return [], [{"reason": "llm_budget_exhausted"}]
        llm_remaining -= 1
        sys_prompt = DIAGNOSE_SYSTEM_PROMPT + "\n\n可选动作清单：" + _action_catalog_hint(registry)
        user_prompt = "诊断指纹：\n" + json.dumps(payload, ensure_ascii=False, default=str)
        if spec is not None:
            try:
                from earlyz.spec.describe import describe_spec_rules
                summary = describe_spec_rules(spec)
                user_prompt += ("\n\n模板规则摘要（供 template_mismatch 对照）：\n"
                                + json.dumps(summary[:30], ensure_ascii=False, default=str))
            except Exception:  # noqa: BLE001 — 摘要失败不影响主诊断
                pass
        content = llm_call(sys_prompt, user_prompt)
        parsed = _extract_json(content or "")
        if parsed is None:
            return [], [{"reason": "unparseable_llm_output"}]
        return validate_findings(parsed.get("findings") or [], registry, spec, cfg.risk_ceiling)

    iter_no = 0
    use_payload = diagnose_payload
    while True:
        # 2026-08-27 残留 Word 孤儿清理（P3 写文件 PermissionError 根因）
        try:
            from earlyz.validation.renderers import _cleanup_orphan_word
            _cleanup_orphan_word()
        except Exception:
            pass
        # 2026-08-27 上轮失败注入（LLM 不再重复失败动作——真机 3 轮浪费根因之一）
        if iter_no > 0:
            prior_fail = [f for s in result.steps if s.get("phase") == "fix"
                          for f in s.get("failed_actions", [])]
            if prior_fail:
                use_payload = {**diagnose_payload, "prior_failures": prior_fail}
        # 1) DIAGNOSE
        findings, rejected = _diagnose(use_payload)
        result.findings.extend(findings)
        result.rejected.extend(rejected)
        result.llm_calls += 1
        result.steps.append({"phase": "diagnose", "iter": iter_no,
                             "status": "ok" if findings else "no_findings",
                             "n_findings": len(findings),
                             "n_rejected": len(rejected)})
        if not findings or not any(f.actions for f in findings):
            result.status = "no_findings"
            break

        # 2) FIX（取本轮全部通过校验的 R1/R2 action）
        iter_no += 1
        result.iterations = iter_no
        actions = [a for f in findings for a in f.actions]
        records, ok = _execute_fix(actions, docx, registry, work_dir)
        result.fixed_count += ok
        failed = [r for r in records if r["status"] == "failed"]
        result.steps.append({"phase": "fix", "iter": iter_no,
                             "status": "ok" if ok else "noop",
                             "executed": ok, "failed": len(failed),
                             "total": len(actions),
                             # 2026-08-27 失败详情进轨迹（下轮诊断/task_events 可见）
                             "failed_actions": [{"action": r["action"], "error": r.get("error")}
                                                for r in failed]})

        # 3) VERIFY（静态硬门禁）
        errors = _verify_static(docx)
        result.verified_ok = not errors
        result.steps.append({"phase": "verify", "iter": iter_no,
                             "status": "pass" if not errors else "fail",
                             "n_errors": len(errors)})
        if result.verified_ok and not failed:
            result.status = "ok"
            break
        if iter_no >= cfg.max_iter:
            result.status = "degraded"
            result.degraded_reason = (
                f"FIX-VERIFY 迭代 {iter_no} 次未收敛（{len(failed)} 项失败, "
                f"{len(errors)} 项验证错误）——转 P2 对话确认"
            )
            break
        # 4) 再诊断：把上次执行失败/验证错误作为证据注入
        use_payload = {
            **diagnose_payload,
            "p3_retry": {
                "iteration": iter_no,
                "failed": failed,
                "verify_errors": errors,
            },
        }

    # 5) REPORT（熔断后仍产出报告——不静默）
    if result.status == "no_findings":
        result.report_text = "未发现可自动修复的格式缺陷（诊断结论见 evidence 分类）。"
    else:
        result.report_text = _report(
            llm_call, result.findings, result.fixed_count, result.rejected,
            result.verified_ok,
        )
    result.steps.append({"phase": "report", "iter": iter_no,
                         "status": "ok", "llm_calls": result.llm_calls,
                         "iterations": result.iterations})
    result.llm_calls += 1
    result.elapsed_s = time.time() - t0
    return result


def _action_catalog_hint(registry, limit: int = 60) -> str:
    """DIAGNOSE 提示用：白名单 action 清单（registry 唯一真相源）。"""
    try:
        specs = [s for s in registry.all() if s.risk.value <= "R2"]
        names = [s.name for s in specs]
    except Exception:  # noqa: BLE001
        names = []
    return ", ".join(names[:limit])
