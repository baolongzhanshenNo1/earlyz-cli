"""验证器 provider 化（架构文档 §20 P2——Capability seam 铺开）。

配置化而非 if-else：渲染/验证能力注册表（word_com / officecli / libreoffice / ooxml_static），
SaaS 化前置（CI 用 ooxml_static，桌面用 word_com；环境不适合时自动降级）。

2026-09-23 渲染链切换（台架试验实锤）：LibreOffice 对双栏期刊文档单栏化失真，
默认链移除 libreoffice（保留注册，EARLYZ_RENDER_PROVIDER=libreoffice 可显式选），
中间层改为 officecli 截图渲染（与 Word 分页一致）。

- register_provider：provider 注册（名称 → available/render 工厂）
- resolve_provider：按 preference + 配置选择（EARLYZ_RENDER_PROVIDER 或 config）
- render_with_provider：执行渲染（provider 不可用时按链降级）

行为兼容：现有 validate_delivery 的 renderer="word_preferred" 语义不变——
resolve 优先 word_com，不可用走 officecli，仍不可用走 ooxml_static。
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Callable, Optional

from . import proc_track as _proc


@dataclass
class RenderResult:
    renderer: str
    pages: int = 0
    sections: int = 0
    bbox_overflow: Optional[list] = None
    pdf_path: str = ""
    reason: str = ""
    image_dir: str = ""


# name → {"available": () -> (bool, reason), "render": (docx, timeout) -> dict}
_RENDER_PROVIDERS: dict[str, dict] = {}


def register_provider(name: str,
                      available_fn: Callable[[], tuple[bool, str]],
                      render_fn: Callable) -> None:
    """注册渲染 provider（渲染能力 seam——新增验证器只需注册）。"""
    _RENDER_PROVIDERS[name] = {"available": available_fn, "render": render_fn}

def list_providers() -> list[str]:
    return list(_RENDER_PROVIDERS)

def _config_value(config: dict | None) -> str | None:
    """配置真相：显式 config > 环境变量 > None（默认 word_preferred 语义）。"""
    if config:
        v = config.get("render_provider")
        if v:
            return str(v)
    return os.environ.get("EARLYZ_RENDER_PROVIDER") or None

def resolve_provider(preference: str = "word_preferred",
                     config: dict | None = None) -> str:
    """按 preference + 配置解析 provider（返回实际可用 provider 名或 "none"）。

    - 显式配置（config.render_provider / EARLYZ_RENDER_PROVIDER）优先：
      指定的 provider 不可用 → "none"（并保留 reason 供上层 warning）；
    - 默认链：word_com → officecli → ooxml_static（2026-09-23 起 libreoffice
      移出默认链——双栏渲染失真，见模块 docstring）。
    """
    pref = _config_value(config) or (preference or "word_preferred")
    if pref in ("none", "off"):
        return "none"
    chain = []
    if pref == "word_preferred":
        chain = ["word_com", "officecli", "ooxml_static"]
    elif pref in ("auto", ""):
        chain = ["word_com", "officecli", "ooxml_static"]
    else:
        chain = [pref]
    for name in chain:
        if name in _RENDER_PROVIDERS:
            ok, _reason = _RENDER_PROVIDERS[name]["available"]()
            if ok:
                return name
    return "none"

def render_with_provider(provider: str, docx_path: str,
                         timeout: int = 300) -> RenderResult:
    """执行渲染；provider 名非法/不可用时抛 RuntimeError（上层按 fallback 处理）。"""
    if provider == "none":
        return RenderResult(renderer="none", reason="static-only")
    entry = _RENDER_PROVIDERS.get(provider)
    if entry is None:
        raise RuntimeError(f"render provider not registered: {provider}")
    raw = entry["render"](docx_path, timeout=timeout)
    return RenderResult(
        renderer=str(raw.get("renderer") or provider),
        pages=int(raw.get("pages") or 0),
        sections=int(raw.get("sections") or 0),
        bbox_overflow=raw.get("bbox_overflow"),
        pdf_path=str(raw.get("pdf_path") or ""),
        reason=str(raw.get("reason") or ""),
        image_dir=str(raw.get("image_dir") or ""),
    )


# ── 默认 provider 注册（延迟导入避免 import 环）──────────────────────

def _word_available() -> tuple[bool, str]:
    from earlyz.validation.renderers import word_render_available
    if word_render_available():
        return True, ""
    return False, "Word COM 不可用（未安装/非 Windows）"

def _word_render(docx_path: str, timeout: int = 300) -> dict:
    from earlyz.validation.renderers import render_with_word
    return render_with_word(docx_path, timeout=timeout)

def _officecli_available() -> tuple[bool, str]:
    from earlyz.validation.preview import officecli_available
    if officecli_available():
        return True, ""
    return False, "officecli 不可用（未安装/未配置 EARLYZ_OFFICECLI）"

def _officecli_render(docx_path: str, timeout: int = 300) -> dict:
    """officecli 截图渲染（docx → 页面 PNG → 分页统计）。

    与预览共用缓存（render_pages_cached）——同一产物验证/预览只渲染一次。
    bbox 越界检测不适用截图链（留空；该项由 word_com 门禁兜底）。
    """
    from earlyz.validation.preview import (
        render_pages_cached, sections_count,
    )
    r = render_pages_cached(docx_path, timeout=timeout)
    return {
        "renderer": "officecli",
        "pages": int(r["page_count"]),
        "sections": sections_count(docx_path),
        "bbox_overflow": [],
        "reason": "screenshot-render（bbox 检测由 word_com 门禁承担）",
        "image_dir": str(r["out_dir"]),
    }

def _find_soffice() -> str | None:
    """找 soffice：PATH 优先，其次应用捆绑 LibreOffice（2026-10-06 兜底渲染需要——
    用户机器上 soffice 通常不在 PATH，只在捆绑目录）。"""
    import shutil
    for c in ("soffice", "libreoffice"):
        p = shutil.which(c)
        if p:
            return p
    try:  # 双导入兼容：python 源码布局 api 是顶层包；Nuitka exe 下 host 才是顶层
        try:
            from host.components import find_libreoffice
        except ImportError:
            from api.host.components import find_libreoffice
        return str(find_libreoffice() or "") or None
    except Exception:  # noqa: BLE001 — 解析失败按无捆绑处理
        return None


def _lo_available() -> tuple[bool, str]:
    if _find_soffice():
        return True, ""
    return False, "LibreOffice 不可用（soffice 未安装）"

def _lo_render(docx_path: str, timeout: int = 300) -> dict:
    """LibreOffice CLI 渲染（docx → pdf → 页面分析）。

    ⚠️ 2026-09-23 起移出默认链：双栏文档单栏化失真（台架实锤），
    仅供 EARLYZ_RENDER_PROVIDER=libreoffice 显式选择/回归对比，新代码勿依赖。
    """
    import subprocess
    import tempfile
    from pathlib import Path
    from earlyz.validation.renderers import _analyze_pdf

    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice 不可用")
    with tempfile.TemporaryDirectory(prefix="earlyz-lo-") as td:
        _proc.run(
            [soffice, "--headless", "--convert-to", "pdf",
             "--outdir", td, str(docx_path)],
            check=True, timeout=timeout, capture_output=True,
        )
        pdf = next(Path(td).glob("*.pdf"))
        info = _analyze_pdf(pdf)
    info["renderer"] = "libreoffice"
    info["pdf_path"] = str(pdf) if pdf.exists() else ""
    return info

def _ooxml_static_available() -> tuple[bool, str]:
    return True, "纯 OOXML 静态检查（无渲染引擎依赖）"

def _ooxml_static_render(docx_path: str, timeout: int = 300) -> dict:
    """纯 OOXML 静态检查（无渲染引擎依赖——CI 场景）。

    直接读 package：section 个数 + 近似页数（sectPr 计数——Word 渲染前的静态下界）。
    """
    import zipfile
    import re
    try:
        with zipfile.ZipFile(docx_path) as z:
            doc_xml = z.read("word/document.xml").decode("utf-8", errors="replace")
        sections = len(re.findall(r"<w:sectPr", doc_xml)) + 1
        paragraphs = doc_xml.count("<w:p")
        pages = max(1, (paragraphs + 39) // 40)  # 每页约 40 段的粗下界
    except Exception:
        sections, pages, paragraphs = 0, 0, 0
    return {"renderer": "ooxml_static", "pages": pages, "sections": sections,
            "bbox_overflow": [], "reason": "STATIC"}

register_provider("word_com", _word_available, _word_render)
register_provider("officecli", _officecli_available, _officecli_render)
register_provider("libreoffice", _lo_available, _lo_render)
register_provider("ooxml_static", _ooxml_static_available, _ooxml_static_render)
