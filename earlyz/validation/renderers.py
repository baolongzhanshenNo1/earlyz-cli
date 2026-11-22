"""早睡 (EarlyZ) 3.2 — Render Gate renderers（§11 验证体系，铁律 8/9）。

Windows 上优先使用 Word COM 渲染（docx → PDF → 页面分析），
提供 validate_delivery 的 word_render_callback 契约实现：

    render(docx_path) -> {
        "renderer": "word",
        "pages": int,
        "sections": int,
        "bbox_overflow": [dict],   # 图片/文本块超出页面边界
        "pdf_path": str,
    }

- Word 不可用（未安装/非 Windows）→ raise RendererUnavailableError，
  上层据此降级 LEVEL_C_STATIC_ONLY 并加 warning（不算 hard error）；
- Word 可用但文档打开/导出失败 → raise RenderFailedError，
  上层视为 hard error（文档真实损坏，阻断交付，铁律 8）。

渲染命令复用生产验收轮的 Word COM 方案（PowerShell COM SaveAs2=17），
页面分析用 PDFium（earlyz.validation.pdf_backend，2026-09-30 替换 PyMuPDF）。
"""

from __future__ import annotations

import contextlib
import hashlib
import os
import re
import shutil
import subprocess
import tempfile
import threading
import time
from pathlib import Path
from typing import Optional

from . import proc_track as _proc


class RendererUnavailableError(Exception):
    """Word 不可用（未安装 / 非 Windows / 环境开关关闭）。不是文档问题。"""


class RenderFailedError(Exception):
    """Word 可用但渲染失败（文档损坏 / 打开失败 / 导出失败）。硬错误。"""


# ── Word COM 全局串行（2026-09-25，假失败根因）────────────────────────
# 症状：`render failed: Word COM 执行失败: Exception calling "Quit" ...
#       The RPC server is unavailable (0x800706BA)`，但导出其实已成功
#（产物 .render.pdf 完整、页数正确）→ 交付被判 FAILED。
# 根因链：① Word 是单实例 COM 服务器；② `_cleanup_orphan_word()` 会清掉
# 一切「无主窗口」的 WINWORD（隐藏实例 MainWindowHandle=0）；③ 两处渲染
# 重叠时（预览渲染 vs 交付门禁渲染，来自桌面端侧车/API/CLI 不同线程或进程），
# 先结束的一方把另一方的 Word 进程杀掉 → 后者的 `$w.Quit()` 报 RPC 不可用。
# 因此 Word COM 必须跨线程 + 跨进程排队（铁律：Word COM 必须串行）。
_WORD_COM_PY_LOCK = threading.RLock()


@contextlib.contextmanager
def _word_com_serialized(timeout: float = 900.0):
    """获取 Word COM 独占权（进程内 RLock + 跨进程文件锁）。"""
    if not _WORD_COM_PY_LOCK.acquire(timeout=timeout):
        raise RenderFailedError("Word 渲染排队超时（进程内有未结束的渲染）")
    fh = None
    try:
        lock_path = Path(tempfile.gettempdir()) / "earlyz-word-com.lock"
        fh = open(lock_path, "a+b")
        fh.seek(0)
        deadline = time.time() + timeout
        while True:
            try:
                if os.name == "nt":
                    import msvcrt  # noqa: PLC0415
                    msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl  # noqa: PLC0415
                    fcntl.flock(fh.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                break
            except OSError:
                if time.time() >= deadline:
                    raise RenderFailedError("Word 渲染等待跨进程锁超时")
                time.sleep(0.25)
        yield
    finally:
        if fh is not None:
            try:
                fh.seek(0)
                if os.name == "nt":
                    import msvcrt  # noqa: PLC0415
                    msvcrt.locking(fh.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    import fcntl  # noqa: PLC0415
                    fcntl.flock(fh.fileno(), fcntl.LOCK_UN)
            except Exception:  # noqa: BLE001 — 解锁失败不影响结果
                pass
            try:
                fh.close()
            except Exception:  # noqa: BLE001
                pass
        _WORD_COM_PY_LOCK.release()


_availability_cache: dict = {"ts": 0.0, "value": None}


def word_render_available() -> bool:
    """Word COM 可用性探测（缓存 120s，避免每次编译都探测）。"""
    if os.environ.get("EARLYZ_SKIP_WORD_RENDER") == "1":
        return False
    if os.name != "nt":
        return False
    now = time.time()
    if now - _availability_cache["ts"] < 120:
        return bool(_availability_cache["value"])
    try:
        ps = _run_powershell(
            "$w = New-Object -ComObject Word.Application; $w.Quit(); Write-Output 'OK'",
            timeout=45,
        )
        ok = "OK" in (ps or "")
    except Exception:
        ok = False
    _availability_cache.update(ts=time.time(), value=ok)
    return ok


def build_word_render_callback():
    """返回 validate_delivery 的 word_render_callback；Word 不可用时返回 None。

    调用方（agent_loop.compile_format_tool）用此构造回调，
    不可用时保持 renderer="none" 行为（LEVEL_C + warning）。
    """
    if not word_render_available():
        return None
    return render_with_word


def _cleanup_orphan_word() -> None:
    """清除后台遗留的 WINWORD（无主窗口=COM 残留）。

    根源（2026-08-27 用户任务管理器实证 14 个残留）：subprocess.run 超时 →
    PowerShell 被 kill → Word COM 子进程成为孤儿（Quit 的 finally 没机会执行）。
    孤儿 Word 会锁住产物文件（P3 的 PermissionError 铁证）并累积内存。
    用户手动打开的 Word 有主窗口（MainWindowHandle≠0），不会被误杀。
    """
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             "Get-Process WINWORD -ErrorAction SilentlyContinue | "
             "Where-Object { $_.MainWindowHandle -eq 0 } | Stop-Process -Force"],
            capture_output=True, timeout=30,
        )
    except Exception:
        pass


_PROCESS_HOOK = None


def set_process_hook(fn) -> None:
    """登记子进程钩子（2026-11-06）：api.host.jobs 用它按 PID 精确取消/杀树。

    实现已下沉到 `earlyz/validation/proc_track.py`（soffice/officecli 共用同一钩子）；
    本函数保留为兼容入口。
    """
    global _PROCESS_HOOK
    _PROCESS_HOOK = fn
    _proc.set_process_hook(fn)


def _run_powershell(script: str, timeout: int) -> str:
    # 所有 Word COM 出口（探测/渲染/分页域）都经此函数 → 排队点放这里即全覆盖
    with _word_com_serialized(timeout + 120):
        try:
            # 2026-10-08 预览事故修复：残留「无主窗口」Word 实例会让 New-Object
            # 挂到坏 COM 上 → 探测/渲染连环假失败 → 回退 officecli 逐页渲染
            # （63 页≈1 小时+，用户视角=「预览一直加载不出来」）。持锁先清残留
            # （锁内无并发 Word 渲染，安全），保证每次 COM 从干净状态起步。
            _cleanup_orphan_word()
            # 2026-11-06：用 proc_track.popen 拿 PID——取消时按进程树精确杀，
            # 绝不按镜像名杀 WINWORD（那会连带杀掉用户自己打开的 Word 文档）。
            _proc_handle = _proc.popen(
                ["powershell", "-NoProfile", "-Command", script],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            )
            try:
                _o, _e = _proc_handle.communicate(timeout=timeout)
            except subprocess.TimeoutExpired:
                _proc.kill_tree(_proc_handle.pid)
                try:
                    _proc_handle.communicate(timeout=10)
                except Exception:  # noqa: BLE001
                    pass
                raise
            ps = subprocess.CompletedProcess(_proc_handle.args, _proc_handle.returncode,
                                             _o, _e)
        except subprocess.TimeoutExpired:
            # 超时：PowerShell 被杀 → Word 孤儿残留（2026-08-27 14 进程隐患铁证）
            _cleanup_orphan_word()
            raise
        out = (ps.stdout or b"").decode("utf-8", errors="replace")
        if ps.returncode != 0:
            err = (ps.stderr or b"").decode("utf-8", errors="replace")
            # 退出码非零但脚本已跑到收尾（输出里已有 OK）→ 只是 Quit/RPC 阶段噪声，
            # 不否掉已完成的导出（2026-09-25 假失败修复）
            if "OK" in out:
                _cleanup_orphan_word()
                return out
            raise RuntimeError((err or out or "unknown")[:300])
        _cleanup_orphan_word()  # 正常路径兜底（Quit 静默失败时）；现在有全局锁，安全
        return out


def _extract_word_info(ps_out: str, fallback_sections: Optional[int] = None) -> dict:
    """从 PowerShell 输出解析 Word 信息（页数/节数）。"""
    info: dict = {"sections": fallback_sections}
    for line in ps_out.splitlines():
        line = line.strip()
        if line.startswith("SECTIONS:"):
            try:
                info["sections"] = int(line.split(":", 1)[1].strip())
            except ValueError:
                pass
    return info


def _analyze_pdf(pdf_path: Path) -> dict:
    """PDFium 分析（pdf_backend）：页数 + 图片/文本行越界检测。"""
    from earlyz.validation import pdf_backend  # noqa: PLC0415 — 延迟导入

    overflow: list[dict] = []
    infos = pdf_backend.read_pages(pdf_path, text=False, lines=True, images=True)
    for info in infos:
        pw, ph = info.width, info.height
        for x0, y0, x1, y1 in info.images:
            if x1 > pw + 1.0 or y1 > ph + 1.0 or x0 < -1.0 or y0 < -1.0:
                overflow.append(
                    {
                        "page": info.index + 1,
                        "kind": "image",
                        "bbox": [round(v, 1) for v in (x0, y0, x1, y1)],
                        "page_size": [round(pw, 1), round(ph, 1)],
                    }
                )
        for x0, y0, x1, y1, _text in info.lines:
            # 文本容差 6pt：PDF 字形墨迹测量可略超页边，>6pt 才视为真实越界
            if x1 > pw + 6.0 or y1 > ph + 6.0:
                overflow.append(
                    {
                        "page": info.index + 1,
                        "kind": "text_block",
                        "bbox": [round(v, 1) for v in (x0, y0, x1, y1)],
                        "page_size": [round(pw, 1), round(ph, 1)],
                    }
                )
    return {"pages": len(infos), "bbox_overflow": overflow}


_TABLE_PAGES_PS = r'''
$ErrorActionPreference = 'Stop'
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$doc = $w.Documents.Open('__PATH__', $false, $true)
try {
  $doc.Repaginate()
  $n = $doc.Tables.Count
  "TABLES=" + $n
  for ($i = 1; $i -le $n; $i++) {
    $cells = $doc.Tables.Item($i).Range.Cells
    $c = $cells.Count
    $p1 = $cells.Item(1).Range.Information(3)
    $pl = $cells.Item($c).Range.Information(3)
    $v1 = $cells.Item(1).Range.Information(5)
    $vl = $cells.Item($c).Range.Information(5)
    "TABLE {0} cells={1} first_page={2} last_page={3} v_first={4} v_last={5}" -f $i, $c, $p1, $pl, $v1, $vl
  }
  "PAGES=" + $doc.ComputeStatistics(2)
} finally {
  $doc.Close($false)
}
$w.Quit()
"OK"
'''


def probe_table_pages(docx_path, timeout: int = 300) -> dict:
    """Word 真机读表格分页位置（B 复核 / 校准都用它）。

    返回 {"tables": [{"index","cells","first_page","last_page","v_first","v_last"}],
          "pages": int}

    只用 Cell.Range.Information(3)=wdActiveEndPageNumber（合并格安全）；
    行级 Tables.Rows.Item 在存在 vMerge 时会报错，故不采用。
    Information(5) 是**磅**（pt），不是 twips，仅供同页内相对位置比较。
    """
    path = str(Path(docx_path).resolve()).replace("'", "''")
    if not shutil.which("powershell"):
        raise RendererUnavailableError("powershell 不可用")
    out = _run_powershell(_TABLE_PAGES_PS.replace("__PATH__", path), timeout=timeout)
    tables = []
    pages = 0
    for line in out.splitlines():
        line = line.strip()
        if line.startswith("TABLE "):
            parts = dict(
                (kv.split("=", 1) for kv in line.split()[2:] if "=" in kv)
            )
            try:
                tables.append({
                    "index": int(line.split()[1]),
                    "cells": int(parts.get("cells", 0)),
                    "first_page": int(parts.get("first_page", 0)),
                    "last_page": int(parts.get("last_page", 0)),
                    "v_first": float(parts.get("v_first", 0) or 0),
                    "v_last": float(parts.get("v_last", 0) or 0),
                })
            except ValueError:
                continue
        elif line.startswith("PAGES="):
            try:
                pages = int(line.split("=", 1)[1])
            except ValueError:
                pages = 0
    if "OK" not in out and not tables:
        raise RendererUnavailableError("Word 表格分页探测失败")
    return {"tables": tables, "pages": pages}


def _table_form_indices(product_path) -> list[bool]:
    """按 XML 顺序给出每张表是否「表单/块状表」（层 4.6 判定用，与 A 同源判据）。"""
    try:
        import zipfile as _zf
        from xml.etree import ElementTree as _ET
        from earlyz.actions import table_layout as _TL
        with _zf.ZipFile(product_path) as z:
            body = _ET.fromstring(z.read("word/document.xml")).find(f"{W}body")
        if body is None:
            return []
        return [_TL.classify_table(tbl)["kind"] == "form"
                for tbl in body.iter(f"{W}tbl")]
    except Exception:
        return []


def _table_pagination_pass(product_path, warnings) -> dict:
    """层 4.6（B 主）：Word 真机读「表首格页 / 末格页」，对**表单表**纠偏。

    A（估算 + 分类）负责判定与施工，B 用 Word 自己的排版事实复核：
    表单表被切开、且真机跨度 ≤1 页（说明它装得下一页，只是判决偏保守）→
    force_keep 强制锁同页并复测。无 Word 环境时整层降级为 warning，不影响交付等级。
    """
    info: dict = {"checked": False, "tables": [], "corrected": [], "after": []}
    try:
        from earlyz.validation.renderers import (
            probe_table_pages, RendererUnavailableError)
    except Exception as e:  # noqa: BLE001
        warnings.append(f"table pagination pass skipped: {type(e).__name__}: {e}")
        return info
    try:
        before = probe_table_pages(product_path)
    except RendererUnavailableError as e:
        warnings.append(f"table pagination probe unavailable: {e}")
        return info
    except Exception as e:  # noqa: BLE001 — 探测异常不该否掉交付
        warnings.append(f"table pagination probe failed: {type(e).__name__}: {e}")
        return info
    info["checked"] = True
    info["tables"] = before.get("tables", [])
    info["pages"] = before.get("pages", 0)
    forms = _table_form_indices(product_path)
    try:
        from earlyz.actions.new_actions import keep_table_together
        for t in info["tables"]:
            idx = int(t["index"]) - 1
            is_form = forms[idx] if 0 <= idx < len(forms) else False
            if not is_form or t["last_page"] == t["first_page"]:
                continue
            if t["last_page"] - t["first_page"] > 1:
                continue                    # 真机就超过一页：按定案允许跨页
            res = keep_table_together(product_path,
                                      {"table_index": idx, "force_keep": True})
            if res.get("changed"):
                info["corrected"].append(t["index"])
    except Exception as e:  # noqa: BLE001
        warnings.append(f"table pagination correction failed: {type(e).__name__}: {e}")
        return info
    if info["corrected"]:
        try:
            after = probe_table_pages(product_path)
            info["after"] = after.get("tables", [])
            still = [a for a in info["after"]
                     if a["last_page"] != a["first_page"]]
            if still:
                warnings.append(
                    "table pagination: 纠偏后仍有表跨页（表 "
                    + ", ".join(str(s["index"]) for s in still) + "）")
        except Exception as e:  # noqa: BLE001
            warnings.append(f"table pagination recheck failed: {type(e).__name__}: {e}")
    return info


_BOOKMARK_PAGES_PS = r'''
$ErrorActionPreference = 'Stop'
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$doc = $w.Documents.Open('__PATH__', $false, $true)
try {
  $doc.Repaginate()
  $out = @()
  $doc.Bookmarks.ShowHidden = $true      # _Toc* 是隐藏书签，不开这个取不到
  $cnt = $doc.Bookmarks.Count
  for ($i = 1; $i -le $cnt; $i++) {
    $bm = $doc.Bookmarks.Item($i)
    $n = $bm.Name
    $p = $bm.Range.Information(3)
    # 1 = wdActiveEndAdjustedPageNumber：按文档自己的页码方案（pgNumType 起始值/
    # 罗马数字节）算出的**显示页码**——目录里要写的正是这个值。
    # 3 = wdActiveEndPageNumber 是物理页，二者在「从封面连续编号」的模板里差好几页
    # （2026-11-08 实锤：目录写 1/9/10/11，页脚实际 5/13/14/15，系统性偏小 4）。
    $adj = $bm.Range.Information(1)
    "BM {0} page={1} adj={2}" -f $n, $p, $adj
  }
  "PAGES=" + $doc.ComputeStatistics(2)
} finally {
  $doc.Close($false)
}
$w.Quit()
"OK"
'''


def probe_bookmark_pages(docx_path, timeout: int = 300) -> dict:
    """Word 真机取书签实际页码（目录页码的唯一真源）。

    目录条目指向 `_Toc*` 书签（PAGEREF 的落点），所以「书签名 → 页码」就是
    目录页码的真值：不再由我们自己推算版面（推算必然在分页变化后过期，
    2026-09-26 连续两轮真机翻车即此病因）。

    返回 {"bookmarks": {name: page}, "pages": int, "toc_bookmarks": {name: page}}
    """
    path = str(Path(docx_path).resolve()).replace("'", "''")
    if not shutil.which("powershell"):
        raise RendererUnavailableError("powershell 不可用")
    out = _run_powershell(_BOOKMARK_PAGES_PS.replace("__PATH__", path), timeout=timeout)
    marks: dict = {}
    adjusted: dict = {}
    pages = 0
    for line in out.splitlines():
        line = line.strip()
        if line.startswith("BM "):
            parts = line.split()
            if len(parts) >= 3 and parts[2].startswith("page="):
                try:
                    marks[parts[1]] = int(parts[2].split("=", 1)[1])
                except ValueError:
                    continue
            adj = [x for x in parts if x.startswith("adj=")]
            if adj:
                try:
                    adjusted[parts[1]] = int(adj[0].split("=", 1)[1])
                except ValueError:
                    pass
        elif line.startswith("PAGES="):
            try:
                pages = int(line.split("=", 1)[1])
            except ValueError:
                pages = 0
    if "OK" not in out and not marks:
        raise RendererUnavailableError("Word 书签页码探测失败")
    return {
        "bookmarks": marks,
        "toc_bookmarks": {k: v for k, v in marks.items() if k.startswith("_Toc")},
        # 显示页码（目录条目直接可用的值，无需再减 offset）
        "toc_bookmarks_adj": {k: v for k, v in adjusted.items() if k.startswith("_Toc")},
        "adjusted_bookmarks": adjusted,
        "pages": pages,
    }


def render_pdf_path(docx_path: str | Path) -> Path:
    """render.pdf 的规范落点（2026-10-08 长路径事故修复）。

    真机根因：产物名链式追加「-suggest-时间戳」（每轮 +24 字符），最终
    render.pdf 全路径 264 字符越过 MAX_PATH(260)——Word SaveAs2 在超长目标
    路径上**静默卡死**（不报错、不退出）→ 300s 超时 → officecli 逐页兜底
    （63 页≈小时级）→ 用户视角「预览一直加载不出来」。全路径 ≥250 时改写短
    临时目录——Word 只在可控长度路径上写；PNG/缓存键（按文件内容 sha1）不受
    影响。
    """
    src = Path(docx_path)
    regular = src.with_name(src.stem + ".render.pdf")
    if len(str(regular)) < 250:
        return regular
    tag = hashlib.sha1(str(src.resolve()).encode("utf-8")).hexdigest()[:10]
    d = Path(tempfile.gettempdir()) / "earlyz-render"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"{tag}.render.pdf"


def render_with_word(docx_path: str | Path, timeout: int = 300) -> dict:
    """Word COM 渲染 + 页面分析（word_render_callback 契约实现）。"""
    src = Path(docx_path)
    if not src.exists():
        raise RenderFailedError(f"文档不存在: {src}")
    if not word_render_available():
        raise RendererUnavailableError("Word 不可用（未安装或已跳过）")

    pdf_path = render_pdf_path(src)

    # 长路径 staged copy（2026-10-08 预览卡死事故）：>240 字符的文档路径在
    # Word 打开链路上不可靠（真机：258 字符路径 Open 挂死三次以上；相同字节
    # 的短路径副本（42 字符/259 字符新名）全部秒开）——渲染前把文档复制到短
    # 临时路径，Word 只打开短路径副本，避免长路径在 Word 内部链路（临时文件/
    # 恢复状态/锁文件）里踩奇点。副本在渲染结束后删除。
    staged: Optional[Path] = None
    _open_src = src
    if len(str(src.resolve())) >= 240:
        _tag = hashlib.sha1(str(src.resolve()).encode("utf-8")).hexdigest()[:10]
        staged = Path(tempfile.gettempdir()) / "earlyz-render" / f"{_tag}.docx"
        staged.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, staged)
        _open_src = staged

    # COM 层对含空格的路径会做 URL 编码（%20）导致 Word 找不到文件——
    # 必须用 Windows 原生反斜杠路径。PS 单引号字符串里反斜杠是字面量，唯一需要
    # 转义的是单引号本身（'→''）；漏转义 = 文件名带 ' 即可注入任意 PowerShell
    # （安全审查 H1）。与 table/bookmark 两处探针口径一致。
    _w = str(_open_src.resolve()).replace("'", "''")
    _p = str(pdf_path.resolve()).replace("'", "''")

    script = (
        "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8\n"
        "$OutputEncoding = [System.Text.Encoding]::UTF8\n"
        "$ErrorActionPreference = 'Stop'\n"
        "$w = New-Object -ComObject Word.Application\n"
        "$w.Visible = $false\n"
        "$w.DisplayAlerts = 0\n"
        "try {\n"
        f"  $d = $w.Documents.Open('{_w}', $false, $true)\n"
        f"  $d.SaveAs2('{_p}', 17)\n"
        "  Write-Output ('SECTIONS:' + $d.Sections.Count)\n"
        "  $d.Close($false)\n"
        "  Write-Output 'OK'\n"
        "} catch {\n"
        "  Write-Output ('ERR: ' + $_.Exception.Message)\n"
        "  exit 1\n"
        "} finally {\n"
        "  try { $w.Quit() } catch { }\n"
        "}\n"
    )
    try:
        out = _run_powershell(script, timeout=timeout)
    except subprocess.TimeoutExpired as e:
        raise RenderFailedError(f"Word 渲染超时（>{timeout}s）: {e}") from e
    except RuntimeError as e:
        # Word COM 创建失败（已装但异常）→ 视为渲染失败（文档或环境问题）
        raise RenderFailedError(f"Word COM 执行失败: {e}") from e
    finally:
        if staged is not None:
            try:
                staged.unlink(missing_ok=True)
            except Exception:  # noqa: BLE001 — 清理失败不阻断
                pass

    if "OK" not in out:
        raise RenderFailedError(f"Word 导出失败: {(out or '')[:300]}")

    if not pdf_path.exists():
        raise RenderFailedError("Word 导出未生成 PDF")

    info = _analyze_pdf(pdf_path)
    info.update(_extract_word_info(out))
    info["renderer"] = "word"
    info["pdf_path"] = str(pdf_path)
    return info


_PAGEREF_PS = r"""param([string]$Docx)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
  $d = $w.Documents.Open($Docx, $false, $false)
  $d.Repaginate()
  # 2026-08-23 单栏页码铁证：Fields.Update() 在无 UI 分页下对 PAGEREF
  # 返回错误值（全 '1'——覆盖了源稿正确缓存）。导出 PDF 强制 Word 打印
  # 分页链路（分页+域全部重算）→ 域结果正确 → Save 落盘（PAGE 域缓存
  # 同步为真实页码——域缓存与真实分页一致，Word 打开不再因域重算
  # 触发「是否将更改保存」弹窗）。
  $pdf = [System.IO.Path]::GetTempFileName() + ".pdf"
  $d.ExportAsFixedFormat($pdf, 17)
  $n = 0
  foreach ($f in @($d.Fields)) {
    if ($f.Type -eq 37) {
      $code = '' + $f.Code.Text
      $res = '' + $f.Result.Text
      Write-Output ('PAGEREF' + [char]9 + $code.Trim() + [char]9 + $res.Trim())
      $n++
    }
  }
  $d.Save()
  $d.Close($false)
  Write-Output ('COUNT:' + $n)
  Write-Output 'OK'
} catch {
  Write-Output ('ERR: ' + $_.Exception.Message)
  exit 1
} finally {
  try { $w.Quit() } catch {}
}
"""


_PAGES_PS = r"""param([string]$Docx)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
  $d = $w.Documents.Open($Docx, $false, $true)
  $d.Repaginate()
  $pages = $d.ComputeStatistics(2)
  Write-Output ('PAGES:' + $pages)
  Write-Output 'OK'
} catch {
  Write-Output ('ERR: ' + $_.Exception.Message)
  exit 1
} finally {
  try { $w.Quit() } catch {}
}
"""


def count_doc_pages(docx_path: str | Path, timeout: int = 120) -> int:
    """Word COM 只读打开 + Repaginate，返回物理页数（诊断用，不渲染）。"""
    src = Path(docx_path)
    if not src.exists():
        raise RenderFailedError(f"文档不存在: {src}")
    if not word_render_available():
        raise RendererUnavailableError("Word 不可用（未安装或已跳过）")
    script = Path(tempfile.gettempdir()) / "earlyz_pages_collect.ps1"
    script.write_text(_PAGES_PS, encoding="utf-8")
    try:
        ps = _proc.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-File", str(script), "-Docx", str(src.resolve())],
            capture_output=True, timeout=timeout,
        )
    except subprocess.TimeoutExpired as e:
        raise RenderFailedError(f"页数统计超时（>{timeout}s）") from e
    out = (ps.stdout or b"").decode("utf-8", errors="replace")
    err = (ps.stderr or b"").decode("utf-8", errors="replace")
    if ps.returncode != 0 or "OK" not in out:
        raise RenderFailedError(f"页数统计失败: {(out or err or 'unknown')[:200]}")
    m = re.search(r"PAGES:(\d+)", out)
    if not m:
        raise RenderFailedError(f"页数统计输出异常: {out[:200]}")
    return int(m.group(1))


def collect_pageref_results(docx_path: str | Path, timeout: int = 180) -> dict[str, str]:
    """Word COM 只读打开，刷新 PAGEREF，返回 {书签名: 页码文本}。

    不 Save 文档——避免 Word 回写 lastRenderedPageBreak / updateFields。
    Word 不可用时 raise RendererUnavailableError；打开失败 raise RenderFailedError。
    """
    src = Path(docx_path)
    if not src.exists():
        raise RenderFailedError(f"文档不存在: {src}")
    if not word_render_available():
        raise RendererUnavailableError("Word 不可用（未安装或已跳过）")

    script = Path(tempfile.gettempdir()) / "earlyz_pageref_collect.ps1"
    script.write_text(_PAGEREF_PS, encoding="utf-8")
    try:
        with _word_com_serialized(timeout + 120):  # 同 Word COM 全局锁（跨进程也排队）
            ps = _proc.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy", "Bypass",
                    "-File", str(script),
                    "-Docx", str(src.resolve()),
                ],
                capture_output=True,
                timeout=timeout,
            )
    except subprocess.TimeoutExpired as e:
        raise RenderFailedError(f"PAGEREF 刷新超时（>{timeout}s）") from e

    out = (ps.stdout or b"").decode("utf-8", errors="replace")
    err = (ps.stderr or b"").decode("utf-8", errors="replace")
    if ps.returncode != 0 or "OK" not in out:
        raise RenderFailedError(
            f"PAGEREF 刷新失败: {(out or err or 'unknown')[:300]}"
        )

    results: dict[str, str] = {}
    for line in out.splitlines():
        if not line.startswith("PAGEREF\t"):
            continue
        parts = line.split("\t", 2)
        if len(parts) < 3:
            continue
        m = re.search(r"PAGEREF\s+(\S+)", parts[1], re.I)
        if not m:
            continue
        results[m.group(1)] = parts[2].strip()
    return results

_ROW_PAGES_PS = r'''
$ErrorActionPreference = 'Stop'
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$doc = $w.Documents.Open('__PATH__', $false, $true)
try {
  $doc.Repaginate()
  "TABLES=" + $doc.Tables.Count
  for ($i = 1; $i -le $doc.Tables.Count; $i++) {
    $t = $doc.Tables.Item($i)
    $rows = @()
    for ($r = 1; $r -le $t.Rows.Count; $r++) {
      $rows += $t.Rows.Item($r).Range.Information(3)
    }
    "TABLE {0} ROWS {1}" -f $i, ($rows -join ',')
  }
  "PAGES=" + $doc.ComputeStatistics(2)
} finally {
  $doc.Close($false)
}
'''


def probe_table_row_pages(docx_path, timeout: int = 300) -> dict:
    """每张表「每一行落在第几页」——续表拆分需要精确到行。

    返回 {"tables": [{"index": 1, "row_pages": [1,1,2,...]}], "pages": N}
    """
    path = str(Path(docx_path).resolve()).replace("'", "''")
    out = _run_powershell(_ROW_PAGES_PS.replace("__PATH__", path), timeout=timeout)
    tables: list[dict] = []
    pages = 0
    for line in out.splitlines():
        line = line.strip()
        if line.startswith("TABLE "):
            parts = line.split()
            if len(parts) >= 4 and parts[2] == "ROWS":
                idx = int(parts[1])
                raw = line.split("ROWS", 1)[1].strip()
                row_pages = [int(x) for x in raw.split(",") if x.strip().isdigit()]
                tables.append({"index": idx, "row_pages": row_pages})
        elif line.startswith("PAGES="):
            try:
                pages = int(line.split("=", 1)[1])
            except ValueError:
                pages = 0
    return {"tables": tables, "pages": pages}
