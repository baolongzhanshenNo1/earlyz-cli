"""子进程登记与统一执行（2026-11-06，后台任务契约 P1）。

为什么需要这个模块：取消要"真取消"，就必须能按 PID 树杀掉**我们自己起的**子进程
（PowerShell 里的 Word、soffice、officecli）。按镜像名杀会连带杀掉用户自己打开的
Word 文档，绝不允许。

边界：
1. 只登记、只执行；不感知任务状态，不写库。
2. 未挂 hook 时行为与 subprocess.run 一致（默认 hook=None）。
3. 超时终止走 `taskkill /T`（整棵树），不是按镜像名。
"""
from __future__ import annotations

import os
import subprocess
from typing import Callable

_PROCESS_HOOK: Callable[[int], None] | None = None


def set_process_hook(fn: Callable[[int], None] | None) -> None:
    global _PROCESS_HOOK
    _PROCESS_HOOK = fn


def track(pid: int) -> None:
    if _PROCESS_HOOK is None or not pid:
        return
    try:
        _PROCESS_HOOK(int(pid))
    except Exception:  # noqa: BLE001 — 登记失败不影响执行
        pass


def kill_tree(pid: int) -> None:
    """按 PID 杀进程树（Windows）。失败静默——进程可能已退出。"""
    try:
        if os.name == "nt":
            subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"],
                           capture_output=True, timeout=15)
        else:
            os.kill(pid, 15)
    except Exception:  # noqa: BLE001
        pass


def popen(cmd, **kwargs) -> subprocess.Popen:
    """启动子进程并登记 PID（供取消时按树精确终止）。"""
    proc = subprocess.Popen(cmd, **kwargs)
    track(proc.pid)
    return proc


def run(cmd, *, timeout: float | None = None, check: bool = False,
        capture_output: bool = True) -> subprocess.CompletedProcess:
    """subprocess.run 的登记版：语义一致，额外登记 PID。"""
    stdout = subprocess.PIPE if capture_output else None
    stderr = subprocess.PIPE if capture_output else None
    proc = popen(cmd, stdout=stdout, stderr=stderr)
    try:
        out, err = proc.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        kill_tree(proc.pid)
        try:
            proc.communicate(timeout=10)
        except Exception:  # noqa: BLE001
            pass
        raise subprocess.TimeoutExpired(cmd, timeout or 0)
    res = subprocess.CompletedProcess(cmd, proc.returncode, out, err)
    if check and proc.returncode != 0:
        raise subprocess.CalledProcessError(proc.returncode, cmd, out, err)
    return res
