# -*- coding: utf-8 -*-
"""产物命名统一 v3（2026-10-08，用户拍板）——一个产物一个名字，贯穿落盘/下载/列表。

规则：
- 交付产物名 = <文档名>-MMDD-HHMMSS.<ext>；Word 与 PDF 永远同茎
  （下载菜单/打包下载/预览配对全部据此，跨格式永远配得上对）。
- 文档名 = 源文档主名清洗：剥 uid 前缀链（每次上传叠一层）、历史
  -suggest-/旧紧凑戳链、事务/内部包装名（tx_/p3fix_/uid_formatted 等），
  限 40 字符。
- tx_*/p3fix_*/uid_ 等内部名永不作为交付名（tx candidate 只是事务暂存——
  真机事故：`tx_170962199a08-candidate.docx` 直接进了产物列表）。
- 诊断包名跟随交付产物名：<产物主名>-diagnose.json。
- 列表显示名 = 交付名去戳去扩展（旧记录显示层同规则派生，不用改库）。

集中所有「写 OUTPUT + 设 result_filename」的交付点使用，禁止各处自造名字。
"""

from __future__ import annotations

import re
import time
from pathlib import Path

MAX_BASE = 40

_UID_RE = re.compile(r"^(?:[0-9a-f]{8}_)+")
_SUGGEST_CHAIN_RE = re.compile(r"(-suggest-\d{8}-\d{6})+$")
_COMPACT_CHAIN_RE = re.compile(r"(-\d{4}-\d{6})+$")
_TX_RE = re.compile(r"^tx_[0-9a-f]{12}-(?:candidate|backup)$")
_P3FIX_RE = re.compile(r"^p3fix_\d+_[0-9a-f]{8}_")
_INTERNAL_RE = re.compile(r"^[0-9a-f]{8}_(?:formatted|scan_form)$")
# 同名冲突序号只剥「紧跟在我们时间戳/建议戳之后」的 -2/-3（用户自己的 "-2" 结尾名不动）
_CONFLICT_RE = re.compile(r"^(.+(?:-\d{4}-\d{6}|-suggest-\d{8}-\d{6}))-\d{1,2}$")


def clean_stem(stem: str, limit: int | None = MAX_BASE) -> str:
    """文档主名清洗。

    返回空串表示「剥完只剩内部包装名」（tx_/p3fix_/uid_formatted 等）——
    调用方须换真实文档名提示（如 ctx.working / 原始输入名）再清洗。
    limit=None 表示不截断（显示层用——40 字上限只约束文件名）。
    """
    base = Path(str(stem).strip()).stem or str(stem).strip()
    if _TX_RE.match(base) or _INTERNAL_RE.match(base):
        return ""  # 内部名（uid_formatted 等）在剥 uid 前判定，防「剥完剩 formatted」
    m = _CONFLICT_RE.match(base)
    if m:
        base = m.group(1)
    for _ in range(8):  # 多层包装/链式追加——循环剥到稳定
        prev = base
        base = _UID_RE.sub("", base).strip(" _-")
        base = _P3FIX_RE.sub("", base).strip(" _-")
        base = _SUGGEST_CHAIN_RE.sub("", base)
        base = _COMPACT_CHAIN_RE.sub("", base)
        if base == prev:
            break
    if not base.strip(" _-") or _TX_RE.match(base) or _INTERNAL_RE.match(base):
        return ""
    if limit and len(base) > limit:
        base = base[:limit].rstrip(" _-")
    return base


def deliver_name(doc_hint: str, ext: str = ".docx", *, ts: float | None = None) -> str:
    """交付产物名：<文档名>-<MMDD-HHMMSS>.<ext>。

    清洗失败（hint 是内部名）回退「文档排版结果」；ext 带不带点都行。
    """
    base = clean_stem(doc_hint) or "文档排版结果"
    stamp = time.strftime("%m%d-%H%M%S", time.localtime(ts) if ts else time.localtime())
    if not ext.startswith("."):
        ext = "." + ext
    return f"{base}-{stamp}{ext}"


def unique_in(directory: Path, filename: str) -> Path:
    """目录内同名冲突自动 -2/-3 步进（与建议流工作副本同口径）。"""
    directory = Path(directory)
    p = directory / filename
    if not p.exists():
        return p
    stem, dot, ext = filename.rpartition(".") if "." in filename else (filename, "", "")
    k = 2
    while True:
        name = f"{stem}-{k}.{ext}" if dot else f"{stem}-{k}"
        cand = directory / name
        if not cand.exists():
            return cand
        k += 1


def display_name(product: str, fallback: str = "文档排版结果") -> str:
    """列表/标题显示名：交付名（或旧记录 file_path）→ 文档名（去戳去扩展）。

    显示层不截断（40 字上限只约束文件名——旧记录长链名照常全文显示）。
    """
    return clean_stem(product, limit=None) or fallback
