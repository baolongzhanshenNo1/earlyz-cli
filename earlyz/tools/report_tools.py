"""格式检测报告工具（M1）：read_format_report / fix_format_report_issues。

- read_format_report：解析批注版 docx → 问题清单 + 能力映射 + 预览（只读）。
- fix_format_report_issues：对可自动修复项执行修复（工作副本语义：原文件
  永不修改；每项 handler+verifier 成对，verifier 不过即如实报 verify_failed）。
  返回「修复对照表」行：已修复 / 仅报告 / 不支持 / 失败——三保险之一「不静默」。
"""

from __future__ import annotations

from pathlib import Path

from earlyz.reports import parse_format_report
from earlyz.reports.remediate import MODE_AUTO, plan


def _lean_plan(planned: dict) -> dict:
    """给 LLM 的计划刊要版：去掉冗长字段，保留事实与计数；同类项预览合并。"""
    out = {
        "platform": planned["platform_label"],
        "template_name": planned["template_name"],
        "summary": planned["summary"],
        "auto_fixable_ids": planned["auto_fixable_ids"],
        "issues": [],
        "parse_warnings": planned["parse_warnings"],
    }
    seen_preview: dict[str, int] = {}
    for it in planned["issues"]:
        row = {
            "id": it["id"],
            "group": it["group"],
            "sub": it.get("sub"),
            "name": it["name"],
            "severity": it["severity"],
            "current": it["current"],
            "standard": it["standard"],
            "anchor": (it.get("anchor") or "")[:60],
            "capability": it["capability"].get("mode"),
            "action": it["capability"].get("action"),
            "params": it["capability"].get("params"),
            "note": it["capability"].get("note"),
        }
        prev = it.get("preview")
        if isinstance(prev, dict):
            import json as _json
            key = _json.dumps(prev, sort_keys=True, ensure_ascii=False, default=str)
            if key in seen_preview:
                row["preview"] = {"same_as": seen_preview[key]}
            else:
                seen_preview[key] = it["id"]
                if "convertible_total" in prev:
                    row["preview"] = {
                        "convertible_total": prev["convertible_total"],
                        "kept_total": prev["kept_total"],
                        "zone_skips": prev["zone_skips"],
                        "samples": prev["samples"][:3],
                    }
                elif "markers_total" in prev:
                    row["preview"] = {
                        "markers_total": prev["markers_total"],
                        "superscript_total": prev["superscript_total"],
                        "ref_entries": prev["ref_entries"],
                        "cited_count": prev["cited_count"],
                        "uncited_total": prev["uncited_total"],
                        "over_total": prev["over_total"],
                        "not_superscript_total": prev["not_superscript_total"],
                    }
                elif "footer_parts" in prev:
                    row["preview"] = {"footer_parts": prev["footer_parts"][:8],
                                      "target": prev.get("target")}
                else:
                    row["preview"] = prev
        out["issues"].append(row)
    return out


def read_format_report(report_path: str | Path, doc_path: str | Path | None = None) -> dict:
    """解析检测报告并生成处置计划（只读）。

    report_path = 批注版报告文件（分析问题清单用）；
    doc_path    = 当前工作副本（预览证据针对它算；缺省=report_path）。
    """
    report = parse_format_report(report_path)
    target = doc_path if doc_path is not None else report_path
    planned = plan(report, target, with_preview=True)
    return {
        "ok": True,
        "report_file": str(report_path),
        "doc": str(target),
        "report": _lean_plan(planned),
        "guidance": (
            "处置顺序：① 向用户简述问题清单与处置分类；② 调 fix_format_report_issues "
            "自动修复 auto_fix 项；③ 输出修复对照表 + 建议用户回平台复检。"
            "仅报告/不支持项不得擅自改。"
        ),
    }


def fix_format_report_issues(doc_path: str | Path, report_path: str | Path,
                             issue_ids: list | None = None) -> dict:
    """按报告修复可自动修复项（工作副本语义；逐项 verifier 复查）。"""
    from earlyz.tools.suggest_fixes import _work_copy

    work = Path(doc_path)
    report = parse_format_report(report_path)
    planned = plan(report, work, with_preview=False)
    wanted = set(int(x) for x in issue_ids) if issue_ids else None

    copy_path, backup = _work_copy(work)
    rows: list[dict] = []
    fixed = failed = skipped = covered = 0
    seen: dict[tuple, int] = {}  # (action, params) -> 首个执行该动作的问题 id（同类项合并）
    for it in planned["issues"]:
        iid = it["id"]
        cap = it["capability"]
        if wanted is not None and iid not in wanted:
            continue
        base = {"id": iid, "issue": f"{it['group']}-{it.get('sub') or ''}-{it['name']}",
                "anchor": (it.get("anchor") or "")[:40]}
        if cap["mode"] != MODE_AUTO:
            rows.append({**base, "status": "reported_only",
                         "note": cap.get("note")})
            skipped += 1
            continue
        import json as _json
        key = (cap["action"],
               _json.dumps(cap.get("params") or {}, sort_keys=True, ensure_ascii=False))
        if key in seen:
            rows.append({**base, "status": "covered", "action": cap["action"],
                         "note": f"与问题 #{seen[key]} 同一修复动作（同类项合并执行）"})
            covered += 1
            continue
        seen[key] = iid
        try:
            from earlyz.actions import report_fixes as rf
            handler = getattr(rf, cap["action"])
            vname = f"verify_{cap['action']}"
            if not hasattr(rf, vname) and cap["action"].startswith("set_"):
                vname = f"verify_{cap['action'][4:]}"
            verifier = getattr(rf, vname)
            out = handler(copy_path, cap.get("params") or {})
            ok = verifier(copy_path, cap.get("params") or {})
            detail = None
            if isinstance(out, dict):
                if cap["action"] == "normalize_cjk_punctuation":
                    detail = {"converted": out.get("converted"),
                              "kept_total": out.get("kept_total"),
                              "zone_skips": out.get("zone_skips"),
                              "samples": [f"{c['before'][:40]} → {c['after'][:40]}"
                                          for c in (out.get("changes") or [])[:6]]}
                elif cap["action"] == "set_footer_page_number_font":
                    detail = {"parts": [p["part"] for p in out.get("parts") or []],
                              "before": sorted({x for p in (out.get("parts") or [])
                                                for x in p.get("before") or []}),
                              "size": out.get("size")}
            rows.append({**base, "status": "fixed" if ok else "verify_failed",
                         "action": cap["action"], "detail": detail,
                         "note": cap.get("note")})
            fixed += 1 if ok else 0
            failed += 0 if ok else 1
        except Exception as e:  # noqa: BLE001
            rows.append({**base, "status": "failed", "action": cap.get("action"),
                         "error": f"{type(e).__name__}: {e}"})
            failed += 1

    return {
        "ok": failed == 0,
        "result_path": str(copy_path),
        "backup": str(backup),
        "fixed": fixed, "failed": failed,
        "covered": covered, "reported_only": skipped,
        "rows": rows,
        "guidance": ("修复对照表已生成；请在回复中逐条呈现（已修复/仅报告/不支持），"
                     "并建议用户回平台复检。若全部为 reported_only 则说明无需落盘改动。"),
    }
