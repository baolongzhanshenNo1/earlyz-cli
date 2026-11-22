"""系统字体可用性探测（2026-11-21）。

只读枚举 + 注册表查询，**不复制、不打包、不嵌入任何字体文件**——本产品对字体的
用法始终是「写字体名 + 由 Word/PDF 后端按系统授权渲染」，这里只回答「这台机器装了没」。

用途：目标字体未安装时（典型：公文的「方正小标宋简体」是商业字库，用户机子多半
没有），排版会静默失真。调用方据此改用近似字体并把事实通报给用户，而不是让 Word
自己乱替换。

两条数据源都用，缺一不可：
- GDI `EnumFontFamiliesExW`：返回**本地化**字体家族名（宋体/黑体/微软雅黑…），
  这是模板与 spec 里实际书写的名字；
- 注册表 `…\\Fonts`：补英文名与别名（Times New Roman、SimSun & NSimSun 这类
  复合项按 & 、、（ 拆开）。

返回 None 的语义是「不知道」（非 Windows / 两条路都失败）——调用方必须保持原行为，
不许拿 None 当 False 去做替换。
"""

from __future__ import annotations

import ctypes
import re
from ctypes import wintypes

_names: set[str] | None = None
_probed = False

_SPLIT = re.compile(r"[&,，、()（）]")


def _clean(name: str) -> str:
    """'@宋体' → '宋体'（@ 前缀是竖排变体家族）。"""
    return name.lstrip("@").strip()


def _enum_families_gdi() -> set[str]:
    class LOGFONTW(ctypes.Structure):
        _fields_ = [
            ("lfHeight", wintypes.LONG), ("lfWidth", wintypes.LONG),
            ("lfEscapement", wintypes.LONG), ("lfOrientation", wintypes.LONG),
            ("lfWeight", wintypes.LONG), ("lfItalic", ctypes.c_ubyte),
            ("lfUnderline", ctypes.c_ubyte), ("lfStrikeOut", ctypes.c_ubyte),
            ("lfCharSet", ctypes.c_ubyte), ("lfOutPrecision", ctypes.c_ubyte),
            ("lfClipPrecision", ctypes.c_ubyte), ("lfQuality", ctypes.c_ubyte),
            ("lfPitchAndFamily", ctypes.c_ubyte), ("lfFaceName", wintypes.WCHAR * 32),
        ]

    class ENUMLOGFONTEXW(ctypes.Structure):
        _fields_ = [
            ("elfLogFont", LOGFONTW), ("elfFullName", wintypes.WCHAR * 64),
            ("elfStyle", wintypes.WCHAR * 32), ("elfScript", wintypes.WCHAR * 32),
        ]

    gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    hdc = user32.GetDC(None)
    if not hdc:
        return set()
    found: set[str] = set()
    proc_type = ctypes.WINFUNCTYPE(
        ctypes.c_int, ctypes.POINTER(ENUMLOGFONTEXW),
        ctypes.c_void_p, wintypes.DWORD, ctypes.c_void_p)

    def _cb(lp_elfe, _lpntme, _fonttype, _lparam):
        found.add(_clean(lp_elfe.contents.elfLogFont.lfFaceName))
        return 1

    try:
        lf = LOGFONTW()
        lf.lfCharSet = 1  # DEFAULT_CHARSET：枚举全部字符集
        gdi32.EnumFontFamiliesExW(hdc, ctypes.byref(lf), proc_type(_cb), 0, 0)
    finally:
        user32.ReleaseDC(None, hdc)
    found.discard("")
    return found


def _enum_names_registry() -> set[str]:
    try:
        import winreg
    except Exception:  # noqa: BLE001 — 非 Windows
        return set()
    names: set[str] = set()
    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            key = winreg.OpenKey(
                root, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")
        except OSError:
            continue
        i = 0
        while True:
            try:
                value_name, _value, _type = winreg.EnumValue(key, i)
            except OSError:
                break
            for part in _SPLIT.split(value_name):
                cleaned = _clean(part)
                if cleaned and not re.fullmatch(
                        r"(?i)true\s*type|open\s*type|type\s*1", cleaned):
                    names.add(cleaned)
            i += 1
    return names


def installed_font_names() -> set[str] | None:
    """本机已安装字体家族名集合（本地化名 + 英文别名）；未知返回 None。"""
    global _names, _probed
    if _probed:
        return _names
    _probed = True
    merged: set[str] = set()
    try:
        merged |= _enum_families_gdi()
    except Exception:  # noqa: BLE001
        pass
    try:
        merged |= _enum_names_registry()
    except Exception:  # noqa: BLE001
        pass
    _names = merged or None
    return _names


def font_available(name: str) -> bool | None:
    """True=本机已装；False=确认没装；None=判断不了（调用方保持原行为）。"""
    if not name:
        return None
    names = installed_font_names()
    if names is None:
        return None
    return _clean(name) in names


def reset_cache() -> None:
    """测试/换机钩子：清掉进程内缓存。"""
    global _names, _probed
    _names = None
    _probed = False
