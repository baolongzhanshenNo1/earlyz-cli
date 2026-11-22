"""渲染超时看门狗。

背景：Word COM 渲染偶尔会卡死（真实事故：单个文档渲染 465 秒无产出；
pytest 全量跑到 Word COM 用例时被 Hermes 后端中断杀掉，EXIT=127）。
线程无法强杀，所以渲染必须跑在**独立子进程**里，超时后 terminate
整个进程树并清理残留 Word 进程。

用法：
    from earlyz.validation.render_timeout import render_with_timeout
    info = render_with_timeout(docx_path, timeout_s=180)

默认不启用——由 EARLYZ_RENDER_TIMEOUT 环境变量控制，避免改变既有行为。
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

_CHILD = r"""
import json, sys
sys.path.insert(0, sys.argv[1])
from earlyz.validation.renderers import build_word_render_callback
try:
    info = build_word_render_callback()(sys.argv[2])
    print("__RESULT__" + json.dumps(info, ensure_ascii=False, default=str))
except Exception as e:
    print("__ERROR__" + f"{type(e).__name__}: {e}")
"""


def _kill_word_children() -> None:
    """清理可能残留的 Word / LibreOffice 进程（仅本机开发环境）。"""
    for image in ("WINWORD.EXE", "soffice.bin", "soffice.exe"):
        try:
            subprocess.run(["taskkill", "/F", "/IM", image],
                           capture_output=True, timeout=15)
        except Exception:  # noqa: BLE001
            pass


def render_with_timeout(docx_path: str | Path, timeout_s: int = 180,
                        repo_root: str | Path | None = None) -> dict:
    """在子进程中渲染，超时则终止并抛出 TimeoutError。"""
    docx_path = str(docx_path)
    root = str(repo_root or Path(__file__).resolve().parents[2])
    proc = subprocess.Popen(
        [sys.executable, "-c", _CHILD, root, docx_path],
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
        encoding="utf-8", errors="replace",
        **({"creationflags": 0x08000000} if os.name == "nt" else {}),
    )
    try:
        out, err = proc.communicate(timeout=timeout_s)
    except subprocess.TimeoutExpired:
        proc.kill()
        try:
            proc.communicate(timeout=10)
        except Exception:  # noqa: BLE001
            pass
        _kill_word_children()
        raise TimeoutError(
            f"渲染超时 {timeout_s}s：{Path(docx_path).name}（已终止进程并清理 Word）"
        )
    for line in (out or "").splitlines():
        if line.startswith("__RESULT__"):
            return json.loads(line[len("__RESULT__"):])
        if line.startswith("__ERROR__"):
            raise RuntimeError(line[len("__ERROR__"):])
    raise RuntimeError(f"渲染子进程无结果：rc={proc.returncode} err={(err or '')[-300:]}")


def render_if_guarded(docx_path: str | Path, callback=None) -> dict:
    """按 EARLYZ_RENDER_TIMEOUT 决定走子进程看门狗还是原回调。"""
    raw = os.environ.get("EARLYZ_RENDER_TIMEOUT", "").strip()
    if raw.isdigit() and int(raw) > 0:
        return render_with_timeout(docx_path, int(raw))
    if callback is None:
        from earlyz.validation.renderers import build_word_render_callback
        callback = build_word_render_callback()
    return callback(str(docx_path))
