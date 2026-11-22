"""siunitx 子集 → 纯文本：\\SI \\qty \\si \\unit \\num \\SIrange \\qtyrange \\numrange \\ang \\numlist。

输出按 siunitx 默认（per-mode=power：\\per 记负指数；单位间细空格）；范围连接词默认 " to "，
cls/导言区 \\sisetup{range-phrase = …} 有设置时由 document.py 读出来传入（ustc 是「～」）。
"""
from __future__ import annotations

import re
from collections import deque

from .macros import plain, read_group, read_opt

SI_CMDS = frozenset({"SI", "qty", "si", "unit", "num", "SIrange", "qtyrange", "numrange", "ang",
                     "numlist", "SIlist", "qtylist", "complexnum", "complexqty", "numproduct",
                     "qtyproduct"})
UNITS = {
    "meter": "m", "metre": "m", "gram": "g", "kilogram": "kg", "second": "s", "ampere": "A",
    "kelvin": "K", "mole": "mol", "candela": "cd", "hertz": "Hz", "newton": "N", "pascal": "Pa",
    "joule": "J", "watt": "W", "coulomb": "C", "volt": "V", "farad": "F", "ohm": "Ω",
    "siemens": "S", "weber": "Wb", "tesla": "T", "henry": "H", "degreeCelsius": "°C",
    "lumen": "lm", "lux": "lx", "becquerel": "Bq", "gray": "Gy", "sievert": "Sv", "katal": "kat",
    "liter": "L", "litre": "L", "minute": "min", "hour": "h", "day": "d", "degree": "°",
    "arcminute": "′", "arcsecond": "″", "hectare": "ha", "tonne": "t", "electronvolt": "eV",
    "dalton": "Da", "bel": "B", "decibel": "dB", "neper": "Np", "bar": "bar", "angstrom": "Å",
    "percent": "%", "byte": "B", "bit": "bit", "astronomicalunit": "au", "atomicmassunit": "u",
}
PREFIXES = {
    "yocto": "y", "zepto": "z", "atto": "a", "femto": "f", "pico": "p", "nano": "n", "micro": "µ",
    "milli": "m", "centi": "c", "deci": "d", "deca": "da", "deka": "da", "hecto": "h", "kilo": "k",
    "mega": "M", "giga": "G", "tera": "T", "peta": "P", "exa": "E", "zetta": "Z", "yotta": "Y",
    "kibi": "Ki", "mebi": "Mi", "gibi": "Gi", "tebi": "Ti",
}
_SUP = str.maketrans("0123456789-+", "⁰¹²³⁴⁵⁶⁷⁸⁹⁻⁺")
# siunitx 缩写单位（\\kg \\MHz \\um …）：由 TeX Live 2026 siunitx.sty 的 \\siunitx_declare_unit 定义链
# 解析得出（前缀 + 单位），离线固化，运行时不依赖本机 TeX。
ABBREV = {
    "radian": "rad", "steradian": "sr", "A": "A", "pA": "pA", "nA": "nA", "uA": "µA", "mA": "mA", "kA": "kA",
    "Hz": "Hz", "mHz": "mHz", "kHz": "kHz", "MHz": "MHz", "GHz": "GHz", "THz": "THz", "mol": "mol",
    "fmol": "fmol", "pmol": "pmol", "nmol": "nmol", "umol": "µmol", "mmol": "mmol", "kmol": "kmol", "V": "V",
    "pV": "pV", "nV": "nV", "uV": "µV", "mV": "mV", "kV": "kV", "hl": "hL", "l": "L", "ml": "mL", "ul": "µL",
    "hL": "hL", "L": "L", "mL": "mL", "uL": "µL", "fg": "fg", "pg": "pg", "ng": "ng", "ug": "µg", "mg": "mg",
    "g": "g", "kg": "kg", "W": "W", "nW": "nW", "uW": "µW", "mW": "mW", "kW": "kW", "MW": "MW", "GW": "GW",
    "J": "J", "uJ": "µJ", "mJ": "mJ", "kJ": "kJ", "eV": "eV", "meV": "meV", "keV": "keV", "MeV": "MeV",
    "GeV": "GeV", "TeV": "TeV", "kWh": "kWh", "m": "m", "pm": "pm", "nm": "nm", "um": "µm", "mm": "mm",
    "cm": "cm", "dm": "dm", "km": "km", "K": "K", "dB": "dB", "F": "F", "fF": "fF", "pF": "pF", "nF": "nF",
    "uF": "µF", "mF": "mF", "H": "H", "fH": "fH", "pH": "pH", "nH": "nH", "uH": "µH", "mH": "mH", "C": "C",
    "nC": "nC", "uC": "µC", "mC": "mC", "T": "T", "uT": "µT", "mT": "mT", "N": "N", "mN": "mN", "kN": "kN",
    "MN": "MN", "Pa": "Pa", "kPa": "kPa", "MPa": "MPa", "GPa": "GPa", "mohm": "mΩ", "kohm": "kΩ",
    "Mohm": "MΩ", "s": "s", "as": "as", "fs": "fs", "ps": "ps", "ns": "ns", "us": "µs", "ms": "ms",
    "celsius": "°C",
}
UNITS.update(ABBREV)

_SEP = "\u2009"


def range_phrase_from(src: str, cjk: bool) -> str | None:
    """从 cls/导言区源码里取 \\sisetup 的 range-phrase。cls 常按语言分支各设一次（ustc：中文「～」、
    英文「 to 」），按文档语言挑：中文文档取含汉字/全角的那个，否则取不含的那个。"""
    cands = []
    for m in re.finditer(r"range-phrase\s*=\s*", src):
        i, depth, j = m.end(), 0, m.end()
        if i < len(src) and src[i] == "{":
            for j in range(i, len(src)):
                depth += {"{": 1, "}": -1}.get(src[j], 0)
                if depth == 0:
                    break
            val = src[i + 1:j]
        else:
            val = re.match(r"[^,}\]]*", src[i:]).group(0)
        texts = re.findall(r"\\text\s*\{([^{}]*)\}", val)
        txt = "".join(texts) if texts else re.sub(r"\\[A-Za-z]+", "", val).replace("{", "").replace("}", "")
        txt = txt.replace("~", "\u00a0")
        if txt.strip():
            core = txt.strip()
            cands.append(core if any(ord(c) > 0x2000 for c in core) else " " + core + " ")
    if not cands:
        return None
    wide = [c for c in cands if any(ord(ch) > 0x2000 for ch in c)]
    narrow = [c for c in cands if c not in wide]
    pick = (wide or narrow) if cjk else (narrow or wide)
    return pick[0]


def _power(n: str) -> str:
    return n.translate(_SUP)


def unit_text(toks) -> str:
    """单位记号 → 文本。\\square \\cubic 作用于下一个单位，\\squared \\cubed \\tothe 作用于上一个；
    \\per 让下一个单位取负指数。字面写法（kg.m/s^2）按原样，「.」变「·」。"""
    out: list[list] = []                   # [单位文本, 指数]
    q, prefix, pre_pow, per = deque(toks or ()), "", None, False
    literal = []
    while q:
        t = q.popleft()
        if t.k == "cs":
            v = t.v
            if v in PREFIXES:
                prefix += PREFIXES[v]
            elif v in UNITS:
                p = pre_pow or 1
                out.append([prefix + UNITS[v], -p if per else p])
                prefix, pre_pow, per = "", None, False
            elif v == "per":
                per = True
            elif v in ("square", "cubic"):
                pre_pow = 2 if v == "square" else 3
            elif v in ("squared", "cubed") and out:
                out[-1][1] *= 2 if v == "squared" else 3
            elif v in ("tothe", "raiseto"):
                n = plain(read_group(q) or [])
                if v == "tothe" and out:
                    out[-1][1] = int(n) if n.lstrip("-").isdigit() else out[-1][1]
                elif n.lstrip("-").isdigit():
                    pre_pow = int(n)
            else:
                literal.append(v)
        elif t.k == "sup":
            g = read_group(q) or []
            literal.append(_power(plain(g)))
        elif t.k in ("char", "space"):
            literal.append("·" if t.v == "." else t.v)
        elif t.k == "active":
            literal.append(" ")
    parts = [u + (_power(str(p)) if p != 1 else "") for u, p in out]
    return (_SEP.join(parts) + "".join(literal)).strip()


def num_text(s: str) -> str:
    s = s.strip().replace(" ", "")
    m = re.fullmatch(r"([-+]?[\d.,]+)[eE]([-+]?\d+)", s)
    if m:
        return m.group(1) + "×10" + _power(m.group(2).lstrip("+"))
    return s.replace("+-", "±")


def complex_text(s: str) -> str:
    """\\complexnum{1+-2i} → 1 ± 2i（siunitx 把 +- 记 ±，复数单位 i/j 原样）。"""
    s = s.strip().replace(" ", "")
    m = re.fullmatch(r"([-+]?[\d.,eE]*?)((?:\+-|[-+])[\d.,eE]*[ij])?", s)
    if not m or not m.group(2):
        return num_text(s)
    re_part, im = m.group(1), m.group(2)
    sign = "±" if im.startswith("+-") else im[0]
    im = im[2:] if im.startswith("+-") else im[1:]
    return (num_text(re_part) + " " + sign + " " if re_part else ("" if sign == "+" else sign)) + im


def si(v: str, q: deque, range_phrase: str = " to ") -> str:
    read_opt(q)
    if v in ("complexnum", "complexqty"):
        z = complex_text(plain(read_group(q) or []))
        return z + (_SEP + unit_text(read_group(q)) if v == "complexqty" else "")
    if v in ("numproduct", "qtyproduct"):
        parts = [num_text(x) for x in re.split(r"\s*x\s*", plain(read_group(q) or [])) if x.strip()]
        u = unit_text(read_group(q)) if v == "qtyproduct" else ""
        return " × ".join(x + (_SEP + u if u else "") for x in parts)
    if v in ("SI", "qty"):
        n = num_text(plain(read_group(q) or []))
        return n + _SEP + unit_text(read_group(q))
    if v in ("si", "unit"):
        return unit_text(read_group(q))
    if v == "num":
        return num_text(plain(read_group(q) or []))
    if v in ("SIrange", "qtyrange", "numrange"):
        a = num_text(plain(read_group(q) or []))
        b = num_text(plain(read_group(q) or []))
        u = unit_text(read_group(q)) if v != "numrange" else ""
        return a + (_SEP + u if u else "") + range_phrase + b + (_SEP + u if u else "")
    if v == "ang":
        parts = (plain(read_group(q) or []).split(";") + ["", ""])[:3]
        marks = ("°", "′", "″")
        return "".join(p.strip() + m for p, m in zip(parts, marks) if p.strip())
    items = [num_text(x) for x in plain(read_group(q) or []).split(";") if x.strip()]
    u = unit_text(read_group(q)) if v in ("SIlist", "qtylist") else ""
    sep = [", "] * max(0, len(items) - 2) + ([" and "] if len(items) > 1 else [])   # list-separator / final
    out = ""
    for k, x in enumerate(items):
        out += x + (_SEP + u if u else "") + (sep[k] if k < len(sep) else "")
    return out
