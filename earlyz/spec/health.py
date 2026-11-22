"""模板规范自检（spec health）。

采样器是静默失败的重灾区：正文样本池为空时，规则仍会生成（取值来自规范注解
或 docDefaults），provenance 依然写 template_ooxml —— 产物看起来正常，实际
没有按模板排版。这里把「样本数 / 取值缺失」变成交付前可见的告警。
"""
from __future__ import annotations

import re
from typing import TYPE_CHECKING

if TYPE_CHECKING:  # pragma: no cover
    from earlyz.spec.models import FormatSpec

_BODY_KEYS = ("font_size_pt", "font_east_asia", "line")


def spec_health_warnings(spec) -> list[str]:
    """返回模板规范的健康告警（空列表 = 健康）。"""
    if spec is None:
        return ["模板规范缺失（spec=None），全部按源文档格式交付"]
    warnings: list[str] = []
    rules = list(getattr(spec, "rules", []) or [])
    body = None
    for r in rules:
        if getattr(r, "rule_id", "") == "rule:tpl.body":
            body = r
            break
    if body is None:
        warnings.append("模板未解析出正文规则：正文将沿用源文档格式（样板可能只含封面/表格）")
        return warnings
    detail = getattr(getattr(body, "provenance", None), "detail", "") or ""
    m = re.search(r"n=(\d+)", detail)
    n = int(m.group(1)) if m else None
    if "from=abstract_body" in detail:
        warnings.append("正文取值来自摘要样本（模板无独立正文），版式按摘要正文推导")
    if n == 0:
        warnings.append("正文规则无模板样本（取值来自注解或默认值），版式可信度低")
    elif n is not None and n < 3:
        warnings.append(f"正文规则样本仅 {n} 条，取值稳定性低")
    expected = getattr(body, "expected", {}) or {}
    missing = [k for k in _BODY_KEYS if expected.get(k) is None]
    if missing:
        warnings.append("正文规则缺字段：" + "、".join(missing) + "（该属性沿用源文档）")
    return warnings
