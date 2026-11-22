"""格式检测报告接入（M1，2026-10-07 边界定稿）。

平台注册表架构：首批 = 维普（泛语）。新增平台 = 新增解析器模块 +
在 _PARSERS 注册（HTML/PDF/其他平台按真实需求+样本再加，不预造）。
"""

from __future__ import annotations

from pathlib import Path

from earlyz.reports.vpcs import PLATFORM_ID as _VPCS_ID
from earlyz.reports import vpcs as _vpcs

__all__ = ["detect_format_report", "parse_format_report", "ReportParseError",
           "PLATFORM_LABELS"]

ReportParseError = _vpcs.ReportParseError

# 平台注册表：platform_id -> (detect_fn, parse_fn)
_PARSERS = {
    _VPCS_ID: (_vpcs.detect, _vpcs.parse),
}

PLATFORM_LABELS = {_VPCS_ID: _vpcs.PLATFORM_LABEL}


def detect_format_report(path: str | Path) -> dict | None:
    """检测 docx 是否为已知平台的格式检测批注版。异常安全（返回 None）。"""
    for _pid, (detect_fn, _parse_fn) in _PARSERS.items():
        try:
            hit = detect_fn(path)
        except Exception:  # noqa: BLE001 —— 检测链路绝不抛出
            hit = None
        if hit:
            return hit
    return None


def parse_format_report(path: str | Path) -> dict:
    """解析格式检测报告 → 统一 IR。

    Raises ReportParseError：不是可识别的检测报告（不猜）。
    """
    last_err: Exception | None = None
    for _pid, (detect_fn, parse_fn) in _PARSERS.items():
        try:
            if detect_fn(path):
                return parse_fn(path)
        except ReportParseError as e:  # noqa: PERF203
            last_err = e
        except Exception as e:  # noqa: BLE001
            last_err = ReportParseError(str(e))
    if last_err is not None:
        raise last_err
    raise ReportParseError("不是已知平台的格式检测报告（当前支持：维普（泛语）批注版 docx）")
