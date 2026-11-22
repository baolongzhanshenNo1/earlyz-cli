"""页面预览渲染（officecli 截图链）+ 分页计数（2026-09-23 台架试验结论落地）。

背景：LibreOffice 在双栏期刊文档上单栏化渲染（cmp_p3 实锤），不能用于
「docx → 页面图」的预览与视觉验证——渲染器失真会把好文档显示成坏的。
officecli 内置渲染器与 Word 分页/分栏一致（p2/p3 零差异），故：

- 预览主链 = Word bulk（一次导出 PDF → PDFium 全册出图）；新鲜
  <名>.render.pdf 直接复用免 Word 启动；无 Word（WPS 用户）→ officecli 截图兜底；
- 任何情况下不用 LibreOffice 渲染页面（只保留 .doc→.docx 转换用途）；
- Word 真机仍是交付门禁唯一 LEVEL_A 裁判（providers.word_com）。

性能铁律（2026-09-23 晚实锤）：officecli 逐页截图在 4MB/50 页级文档上
~45-60s/页（全程 40 分钟）——预览「不显示」事故根因。大文档只许 Word bulk；
officecli 逐页仅限无 Word 机器兜底。

分页计数：officecli `view stats` 不报 docx 页数，`--page N` 越界静默回绕到
第 1 页（2026-09-23 实测 p4 字节 == p1）。故按序截图并以「字节哈希回到第 1 页」
判定末页：第 k>1 页哈希 == 第 1 页哈希 → N = k-1。封面页与后续页字节相同的
概率可忽略（页眉/页码差异），并在 MAX_PAGES 封顶防呆。

缓存：按源文件 sha1 命名目录（EARLYZ_PREVIEW_CACHE_DIR 或系统临时目录下
earlyz-preview-cache），manifest.json 复用——同一产物预览/验证只渲染一次。
"""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import tempfile
import threading
from pathlib import Path
from typing import Optional

from . import proc_track as _proc

MAX_PAGES = 500  # 防呆封顶（100 页级文档正常；超过视为探测异常）


def officecli_bin() -> Optional[str]:
    """officecli 可执行文件：EARLYZ_OFFICECLI 显式指定 > PATH。"""
    env = os.environ.get("EARLYZ_OFFICECLI")
    if env and Path(env).exists():
        return env
    return shutil.which("officecli")


def officecli_available() -> bool:
    return officecli_bin() is not None


def _default_cache_root() -> Path:
    root = os.environ.get("EARLYZ_PREVIEW_CACHE_DIR")
    if root:
        return Path(root)
    return Path(tempfile.gettempdir()) / "earlyz-preview-cache"


def _run_officecli(args: list[str], timeout: int = 180) -> None:
    binary = officecli_bin()
    if binary is None:
        raise RuntimeError("officecli 不可用（未安装/未配置 EARLYZ_OFFICECLI）")
    _proc.run(
        [binary, *args],
        check=True, timeout=timeout, capture_output=True,
    )

def _sha1_file(path: Path) -> str:
    return hashlib.sha1(path.read_bytes()).hexdigest()


def _sha1_png(path: Path) -> str:
    return hashlib.sha1(path.read_bytes()).hexdigest()


def render_pages_officecli(docx_path: str | Path, out_dir: Path,
                           timeout: int = 300) -> int:
    """officecli 逐页截图到 out_dir（page_1.png … page_N.png），返回 N。

    回绕探测：--page k 越界静默回绕到第 1 页 → 字节哈希 == 第 1 页即末页+1。
    渲染完成后 close 释放 resident 会话（避免锁住 docx 影响后续 Word 打开）。
    """
    src = Path(docx_path)
    out_dir.mkdir(parents=True, exist_ok=True)
    for old in out_dir.glob("page_*.png"):
        old.unlink()
    rendered: list[Path] = []
    try:
        k = 1
        first_hash: str | None = None
        while k <= MAX_PAGES + 1:
            out = out_dir / f"page_{k}.png"
            _run_officecli(
                ["view", str(src), "screenshot", "--page", str(k), "-o", str(out)],
                timeout=timeout,
            )
            if not out.exists() or out.stat().st_size == 0:
                break  # 版本差异：越界直接不出图
            h = _sha1_png(out)
            if k == 1:
                first_hash = h
                rendered.append(out)
                k += 1
                continue
            if h == first_hash:
                out.unlink()  # 回绕产生的第 1 页副本
                break
            rendered.append(out)
            k += 1
        n = len(rendered)
        if n == 0:
            raise RuntimeError("officecli 未渲染出任何页面")
        # 归一化文件名（page_1..page_N 连续；探测中断也保持连续）
        for idx, p in enumerate(rendered, start=1):
            target = out_dir / f"page_{idx}.png"
            if p != target:
                p.rename(target)
        return n
    except BaseException:
        # 2026-11-06（P1 取消清理）：取消/超时/异常时删掉本次已截的半套图。
        # 不清理会让缓存目录留着 page_1..page_k，前端把"半套"当完整预览。
        # 进入本函数时已清空过旧图，此处删的必然是本次产物。
        for _p in out_dir.glob("page_*.png"):
            try:
                _p.unlink()
            except OSError:
                pass
        raise
    finally:
        # 取消后 close 仍要试（resident 会话会锁住 docx），但限时，别让它再拖 60s
        try:
            _run_officecli(["close", str(src)], timeout=15)
        except Exception:  # noqa: BLE001 — 释放失败不阻断结果
            pass


def render_pages_word(docx_path: str | Path, out_dir: Path,
                      timeout: int = 300,
                      pdf_path: Optional[str | Path] = None) -> int:
    """Word bulk：docx → PDF → PDFium 全册 PNG（预览主链，一次导出全册）。"""
    from earlyz.validation import pdf_backend  # noqa: PLC0415 — 延迟导入

    out_dir.mkdir(parents=True, exist_ok=True)
    for old in out_dir.glob("page_*.png"):
        old.unlink()
    if pdf_path is not None:
        pdf = Path(pdf_path)  # 复用门禁新鲜 render.pdf —— 免 Word 启动
    else:
        from earlyz.validation.renderers import render_with_word
        info = render_with_word(docx_path, timeout=timeout)
        pdf = Path(info["pdf_path"])
    return len(pdf_backend.render_png_pages(pdf, out_dir, dpi=110))


# 同文档并发渲染单飞（2026-10-08）：preview 请求按 cache_key 排队
_PREVIEW_LOCKS: dict = {}
_PREVIEW_LOCKS_GUARD = threading.Lock()


def _preview_key_lock(key: str) -> "threading.Lock":
    with _PREVIEW_LOCKS_GUARD:
        lk = _PREVIEW_LOCKS.get(key)
        if lk is None:
            lk = threading.Lock()
            _PREVIEW_LOCKS[key] = lk
        return lk


def render_pages_cached(docx_path: str | Path,
                        out_dir: Optional[str | Path] = None,
                        timeout: int = 300) -> dict:
    """渲染（或复用缓存）文档全部页面。

    返回 {cache_key, out_dir, page_count, pages: [str], renderer}。
    renderer ∈ {"word", "officecli"}；Word bulk 优先（全册一次导出），无 Word
    回退 officecli 逐页；禁 LibreOffice。新鲜 <名>.render.pdf 直接复用。
    2026-10-08：同文档并发请求单飞（per-key 锁）——反复点预览不再起多路渲染
    互踩（此前并发清页/抢 Word 造成假失败与无谓等待）；后到者等锁后复用。
    """
    src = Path(docx_path)
    if not src.exists():
        raise RuntimeError(f"文档不存在: {src}")
    key = _sha1_file(src)[:16]
    root = Path(out_dir) if out_dir else _default_cache_root()
    target = root / key
    manifest_path = target / "manifest.json"

    def _try_reuse() -> Optional[dict]:
        if not manifest_path.exists():
            return None
        try:
            mf = json.loads(manifest_path.read_text(encoding="utf-8"))
            pages = [str(target / f"page_{i}.png")
                     for i in range(1, int(mf["page_count"]) + 1)]
            if all(Path(p).exists() for p in pages):
                return {"cache_key": key, "out_dir": str(target),
                        "page_count": len(pages), "pages": pages,
                        "renderer": mf.get("renderer", "officecli")}
        except Exception:  # noqa: BLE001 — manifest 损坏 → 重渲染
            pass
        return None

    hit = _try_reuse()
    if hit:
        return hit
    with _preview_key_lock(key):
        hit = _try_reuse()  # 等锁期间可能已被同键请求渲染完成
        if hit:
            return hit
        # 渲染链（2026-09-23 晚修正）：Word bulk 优先——officecli 逐页在 50 页级
        # 文档 ~45-60s/页（预览不显示事故根因）；officecli 仅无 Word 时兜底；禁 LO。
        # 2026-10-08：.pdf 产物直接出图（导出 PDF 后的「改后」预览；旧链路只收
        # docx → PDF 产物预览 400「格式不支持」）。
        renderer = None
        n = 0
        if src.suffix.lower() == ".pdf":
            from earlyz.validation import pdf_backend  # noqa: PLC0415
            target.mkdir(parents=True, exist_ok=True)
            for old in target.glob("page_*.png"):
                old.unlink()
            n = len(pdf_backend.render_png_pages(src, target, dpi=110))
            renderer = "pdf"
        from earlyz.validation.renderers import render_pdf_path
        reuse_pdf = render_pdf_path(src) if renderer is None else None
        if renderer is None and reuse_pdf is not None \
                and reuse_pdf.exists() and reuse_pdf.stat().st_mtime >= src.stat().st_mtime:
            try:
                n = render_pages_word(src, target, pdf_path=reuse_pdf)
                renderer = "word"
            except Exception:  # noqa: BLE001 — 复用失败 → 走正常导出
                renderer = None
        if renderer is None:
            try:
                n = render_pages_word(src, target, timeout=timeout)
                renderer = "word"
            except Exception:  # noqa: BLE001 — 无 Word / COM 故障 → officecli 兜底
                renderer = None
        if renderer is None:
            if not officecli_available():
                raise RuntimeError("无可用页面渲染器（LibreOffice 已禁用）")
            try:
                n = render_pages_officecli(src, target, timeout=timeout)
                renderer = "officecli"
            finally:
                # officecli 每页经 Word 截图 → 结束统一清残留（2026-10-08：
                # 残留实例会拖垮后续 Word 探测/渲染，累积成预览连环假失败）
                try:
                    from earlyz.validation.renderers import _cleanup_orphan_word
                    _cleanup_orphan_word()
                except Exception:  # noqa: BLE001
                    pass
        pages = [str(target / f"page_{i}.png") for i in range(1, n + 1)]
        manifest_path.write_text(
            json.dumps({"page_count": n, "renderer": renderer,
                        "source_sha1": key}, ensure_ascii=False),
            encoding="utf-8",
        )
        return {"cache_key": key, "out_dir": str(target),
                "page_count": n, "pages": pages, "renderer": renderer}


def sections_count(docx_path: str | Path) -> int:
    """节数（sectPr 计数）——与 providers._ooxml_static_render 同口径。"""
    import re
    import zipfile

    try:
        with zipfile.ZipFile(str(docx_path)) as z:
            doc_xml = z.read("word/document.xml").decode("utf-8", errors="replace")
        return len(re.findall(r"<w:sectPr", doc_xml)) + 1
    except Exception:  # noqa: BLE001
        return 0
