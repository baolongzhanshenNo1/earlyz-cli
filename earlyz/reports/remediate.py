"""检测报告问题 → 修复能力映射与预览（M1，2026-10-07 定稿）。

原则：只把「已被真实样本实测验证」的问题类型接自动修复（auto_fix）；
其余一律 check（可核查/出证据）或 unsupported（如实说明）——宁可少做不瞎做。
"""

from __future__ import annotations

import re
from pathlib import Path

from earlyz.actions import report_fixes as rf
from earlyz.reports.checks import check_reference_citations
from earlyz.spec.decompiler import (_ARABIC_SIZE_PT, _ARABIC_Xiao_PT, _CN_SIZE_PT,
                                    _SZ_ARABIC_RE, _SZ_RE)

MODE_AUTO = "auto_fix"
MODE_CHECK = "check"
MODE_UNSUPPORTED = "unsupported"

_CN_DIGITS = {1: "一", 2: "二", 3: "三", 4: "四", 5: "五", 6: "六", 7: "七"}


def _cn_size_from_text(text: str) -> str | None:
    """从规范/现状文本提取字号 → 规范中文名（五号/小四…）；取不到返回 None。"""
    if not text:
        return None
    m = _SZ_RE.search(text)
    if m:
        return m.group(1)
    m = re.search(r"小([1-7])(?!\d)", text)
    if m:
        return f"小{_CN_DIGITS[int(m.group(1))]}"
    m = re.search(r"([1-7])\s*号", text)
    if m:
        return f"{_CN_DIGITS[int(m.group(1))]}号"
    return None


def capability_for(issue: dict, doc_path: str | Path) -> dict:
    """单条问题 → 能力（不触盘；doc_path 仅为签名统一保留）。"""
    name = issue.get("name") or ""
    sub = issue.get("sub") or ""
    group = issue.get("group") or ""
    std = issue.get("standard") or ""
    joined = f"{group}{sub}{name}"
    if "页码" in joined and "字号" in name:
        cn = _cn_size_from_text(std)
        if cn and cn in _CN_SIZE_PT:
            return {"mode": MODE_AUTO, "action": "set_footer_page_number_font",
                    "params": {"cn_size": cn},
                    "note": f"页脚页码段字号 → {cn}（报告规范值）"}
        return {"mode": MODE_CHECK, "checker": "footer_size",
                "note": "规范值未能映射为可执行字号，仅核查通报"}
    if "标点" in name:
        return {"mode": MODE_AUTO, "action": "normalize_cjk_punctuation",
                "params": {},
                "note": "半角标点低风险子集自动归一（参考文献/英文摘要/类代码段不动）"}
    if "引用" in joined:
        return {"mode": MODE_CHECK, "checker": "citations",
                "note": "引擎核查正文标记与文后条目对应关系（只报事实，不猜平台判定）"}
    if "环绕" in joined:
        return {"mode": MODE_CHECK, "checker": "images",
                "note": "引擎有「图片转嵌入型」动作（set_images_inline）；是否适用需人工确认"}
    if "字体" in name or "行距" in name or "缩进" in name or "对齐" in name:
        return {"mode": MODE_CHECK, "checker": "inspect",
                "note": "字体/行距/缩进类：引擎可核查现状；修复取值需模板口径（未接自动修复）"}
    return {"mode": MODE_UNSUPPORTED,
            "note": "引擎暂无该问题类型的修复能力（如实报告，不擅自改）"}


def preview_for(issue: dict, doc_path: str | Path) -> dict | None:
    """当前文档上该问题的预览/证据（只读；异常安全）。"""
    cap = issue.get("capability") or {}
    try:
        if cap.get("action") == "normalize_cjk_punctuation":
            scan = rf.scan_cjk_punctuation(doc_path)
            return {
                "convertible_total": len(scan["convertible"]),
                "samples": [f"…{c['before'][:44]}… → …{c['after'][:44]}…"
                            for c in scan["convertible"][:6]],
                "kept_total": scan["kept_total"],
                "zone_skips": scan["zone_skips"],
            }
        if cap.get("action") == "set_footer_page_number_font":
            sizes = rf.footer_page_number_sizes(doc_path)
            return {
                "footer_parts": [{"part": s["part"],
                                  "sizes": sorted({str(x) for x in s["sizes"]})}
                                 for s in sizes],
                "target": cap.get("params"),
            }
        if cap.get("checker") == "citations":
            return check_reference_citations(doc_path)
        if cap.get("checker") == "footer_size":
            sizes = rf.footer_page_number_sizes(doc_path)
            return {"footer_parts": [s["part"] for s in sizes]}
    except Exception as e:  # noqa: BLE001
        return {"error": f"{type(e).__name__}: {e}"}
    return None


def plan(report: dict, doc_path: str | Path, *, with_preview: bool = True) -> dict:
    """解析结果 → 附能力与预览的处置计划。"""
    issues = []
    for it in report.get("issues") or []:
        cap = capability_for(it, doc_path)
        row = {**it, "capability": cap}
        if with_preview:
            prev = preview_for(row, doc_path)
            if prev is not None:
                row["preview"] = prev
        issues.append(row)
    auto_ids = [i["id"] for i in issues if i["capability"]["mode"] == MODE_AUTO]
    return {
        "platform": report.get("platform"),
        "platform_label": report.get("platform_label"),
        "template_name": report.get("template_name"),
        "summary": report.get("summary"),
        "issues": issues,
        "auto_fixable_ids": auto_ids,
        "parse_warnings": report.get("parse_warnings") or [],
    }
