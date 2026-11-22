"""字体名归一：把 LibreOffice/开源替代字体映射回 Windows 常见字体。

背景（2026-10-27 交大本科）：模板文件里 `docDefaults` 写着
`w:eastAsia="Noto Serif SC"`、`w:ascii="Liberation Serif"` —— 这是模板在
Linux/LibreOffice 上编辑留下的替代字体。Windows 版 Word 没有这两款字体，
打开时会自行替换（字形、行宽都变），提取器如果照抄，产物就带着一堆
「本机不存在的字体名」。

口径：只映射**已知的开源克隆字体**，未知字体原样保留。映射表要保守——
把用户真用的 Noto 判成宋体同样是错。
"""
from __future__ import annotations

import re

# 已知替代字体 → Windows 常见等价字体（键一律小写）
FONT_ALIASES: dict[str, str] = {
    "liberation serif": "Times New Roman",
    "liberation sans": "Arial",
    "liberation mono": "Courier New",
    "noto serif sc": "宋体",
    "noto serif cjk sc": "宋体",
    "noto serif": "宋体",
    "noto sans sc": "黑体",
    "noto sans cjk sc": "黑体",
    "noto sans": "Arial",
    "source han serif sc": "宋体",
    "source han serif cjk sc": "宋体",
    "source han sans sc": "黑体",
    "source han sans cjk sc": "黑体",
}

# 宋体/Times 的常见别名，用于去重（"宋体;SimSun" 归一后只留一个）
_EQUIV: dict[str, str] = {
    "simsun": "宋体",
    "宋体": "宋体",
    "simhei": "黑体",
    "黑体": "黑体",
    "times new roman": "Times New Roman",
    "fangsong_gb2312": "仿宋_GB2312",
    "仿宋": "仿宋",
    "楷体_gb2312": "楷体_GB2312",
    "楷体": "楷体",
}


def normalize_font(name: str) -> str:
    """单个字体名归一：替代字体→常见字体；其余原样返回。"""
    raw = (name or "").strip()
    if not raw:
        return raw
    key = raw.lower()
    if key in FONT_ALIASES:
        return FONT_ALIASES[key]
    # "Noto Serif CJK SC" 这类后缀变体
    for alias, target in FONT_ALIASES.items():
        if key.startswith(alias):
            return target
    return raw


def _equivalence(name: str) -> str:
    return _EQUIV.get(name.lower(), name.lower())


def normalize_font_field(value: str) -> str:
    """字体字段归一：支持 "宋体;SimSun" / "Times New Roman" 这类写法。

    逐项映射 → 按等价类去重 → 保留首项原名（"宋体;SimSun" → "宋体"）。
    """
    raw = (value or "").strip()
    if not raw:
        return raw
    parts = [p.strip() for p in re.split(r"[;,，；]", raw) if p.strip()]
    if not parts:
        return raw
    out: list[str] = []
    seen: set[str] = set()
    for p in parts:
        mapped = normalize_font(p)
        key = _equivalence(mapped)
        if key in seen:
            continue
        seen.add(key)
        out.append(mapped)
    return ";".join(out)


FONT_FIELDS = (
    "font_east_asia", "font_ascii", "label_font_east_asia", "label_font_ascii",
)


def normalize_font_fields(expected: dict) -> dict:
    """就地把规则 expected 里的字体字段归一，返回是否改动。"""
    changed = False
    for key in FONT_FIELDS:
        v = expected.get(key)
        if isinstance(v, str) and v:
            nv = normalize_font_field(v)
            if nv != v:
                expected[key] = nv
                changed = True
    return changed
