# -*- coding: utf-8 -*-
"""LaTeX → OMML（2026-09-28 夜班）。

对称于 `earlyz/tools/omml_latex.py`（OMML → LaTeX，2026-09-27）。定位：**LaTeX 是
公式的母语**（LLM 写公式默认就是 LaTeX），交付物是 Word 原生公式（OMML）。

设计口径：
· 只编译**明确支持**的子集；遇到不支持的构造**显式抛 `LatexUnsupported`**，
  绝不静默降级成纯文本 —— 静默降级 = 用户看到「改了和没改一样」。
· 公式尺寸取值出自**当前模板**（CJC 模板第 14 条：标准 10.5 磅 / 下标·上标
  5.8 磅 / 次下标·上标 4.5 磅 / 符号 16 磅 / 次符号 10.5 磅）→ 半磅整数
  = 21 / 12 / 9 / 32 / 21。传 sizes 覆盖，缺省即模板值。
· 变量斜体走 `m:rPr/m:sty`（数字、运算符、函数名一律 `p` 直立）。

四条写入铁律（docs/公式与LaTeX残留处理方案.md 第九节）：
① 不进 `m:oMath` 以外的怪容器；② 不用固定行距；③ 编号右对齐走制表位
（动作层负责）；④ 变量斜体走 `m:rPr/m:sty`。
"""
from __future__ import annotations

import re
from typing import Optional

from lxml import etree

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_M = "{%s}" % M
_W = "{%s}" % W

# 模板取值（CJC 模板第 14 条），单位 = 半磅整数
DEFAULT_SIZES = {
    "base": "21",     # 标准 10.5 磅
    "sup": "12",      # 下标/上标 5.8 磅
    "sub": "12",
    "subsup": "9",    # 次下标/上标 4.5 磅
    "sym": "32",      # 符号 16 磅
    "subsym": "21",   # 次符号 10.5 磅
}

MATH_FONT = "Cambria Math"


class LatexUnsupported(Exception):
    """遇到本模块不支持、且不能安全降级的 LaTeX 构造。"""


# ── 符号表 ──────────────────────────────────────────────────────────────────

_GREEK = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
    "zeta": "ζ", "eta": "η", "theta": "θ", "iota": "ι", "kappa": "κ",
    "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ", "pi": "π", "rho": "ρ",
    "sigma": "σ", "tau": "τ", "upsilon": "υ", "phi": "φ", "chi": "χ",
    "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Xi": "Ξ",
    "Pi": "Π", "Sigma": "Σ", "Upsilon": "Υ", "Phi": "Φ", "Psi": "Ψ",
    "Omega": "Ω",
}
_SYMBOLS = {
    # 关系
    "le": "≤", "leq": "≤", "ge": "≥", "geq": "≥", "ne": "≠", "neq": "≠",
    "approx": "≈", "equiv": "≡", "sim": "∼", "simeq": "≃", "propto": "∝",
    "ll": "≪", "gg": "≫", "prec": "≺", "succ": "≻", "subset": "⊂",
    "subseteq": "⊆", "supset": "⊃", "supseteq": "⊇", "in": "∈",
    "notin": "∉", "ni": "∋", "perp": "⊥", "parallel": "∥", "mid": "|",
    # 二元运算
    "times": "×", "cdot": "·", "div": "÷", "pm": "±", "mp": "∓",
    "ast": "∗", "star": "⋆", "circ": "∘", "bullet": "∙", "oplus": "⊕",
    "ominus": "⊖", "otimes": "⊗", "oslash": "⊘", "odot": "⊙", "wedge": "∧",
    "vee": "∨", "cap": "∩", "cup": "∪", "setminus": "∖",
    # 大运算符（走 m:nary 的用 _NARY，这里留通用的）
    # 箭头
    "to": "→", "rightarrow": "→", "leftarrow": "←", "leftrightarrow": "↔",
    "Rightarrow": "⇒", "Leftarrow": "⇐", "Leftrightarrow": "⇔",
    "mapsto": "↦", "uparrow": "↑", "downarrow": "↓", "updownarrow": "↕",
    "longrightarrow": "⟶", "longleftarrow": "⟵",
    # 杂项
    "infty": "∞", "partial": "∂", "nabla": "∇", "forall": "∀", "exists": "∃",
    "nexists": "∄", "emptyset": "∅", "varnothing": "∅", "angle": "∠",
    "therefore": "∴", "because": "∵", "ldots": "…", "dots": "…", "cdots": "⋯",
    "vdots": "⋮", "ddots": "⋱", "prime": "′", "degree": "°", "deg": "°",
    "square": "□", "triangle": "△", "checkmark": "✓", "aleph": "ℵ",
    "hbar": "ℏ", "ell": "ℓ", "Re": "ℜ", "Im": "ℑ", "wp": "℘",
    "neg": "¬", "lnot": "¬", "land": "∧", "lor": "∨", "top": "⊤", "bot": "⊥",
    "models": "⊨", "vdash": "⊢", "dashv": "⊣",
}
_FUNCS = {
    "log", "ln", "lg", "exp", "sin", "cos", "tan", "cot", "sec", "csc",
    "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh", "coth", "max",
    "min", "sup", "inf", "lim", "det", "dim", "gcd", "arg", "ker", "deg",
    "mod", "bmod", "Pr", "tr", "rank", "diag", "sgn",
}
_NARY = {
    "sum": "∑", "prod": "∏", "coprod": "∐", "int": "∫", "iint": "∬",
    "iiint": "∭", "oint": "∮", "bigcup": "⋃", "bigcap": "⋂", "bigoplus": "⨁",
    "bigotimes": "⨂", "bigvee": "⋁", "bigwedge": "⋀", "bigsqcup": "⨆",
}
_ACCENTS = {
    "hat": "^", "widehat": "^", "bar": "¯", "overline": "¯", "tilde": "~",
    "widetilde": "~", "vec": "⃗", "dot": "˙", "ddot": "¨", "check": "ˇ",
    "breve": "˘", "acute": "´", "grave": "`", "mathring": "˚",
}
_SPACES = {"\\,": " ", "\\;": " ", "\\:": " ", "\\!": "", "\\ ": " ",
           "quad": "  ", "qquad": "    ", "thinspace": " ", "enspace": " "}
# 直立文本（不进斜体）
_TEXTY = {"text", "mathrm", "mbox", "operatorname", "textrm", "textup"}
# 加粗
_BOLDY = {"mathbf", "boldsymbol", "bm", "textbf", "pmb"}
# 强制斜体
_ITALICY = {"mathit", "textit", "mathnormal"}


# ── IR ─────────────────────────────────────────────────────────────────────

class _Node:
    __slots__ = ()


class R(_Node):
    """一段文本。sty: 'i' 变量斜体 / 'p' 直立 / 'b' 粗体。"""

    __slots__ = ("text", "sty")

    def __init__(self, text: str, sty: str = "i"):
        self.text = text
        self.sty = sty


class Group(_Node):
    __slots__ = ("kids",)

    def __init__(self, kids: list):
        self.kids = kids


class Sup(_Node):
    __slots__ = ("base", "sup")

    def __init__(self, base, sup):
        self.base, self.sup = base, sup


class Sub(_Node):
    __slots__ = ("base", "sub")

    def __init__(self, base, sub):
        self.base, self.sub = base, sub


class SubSup(_Node):
    __slots__ = ("base", "sub", "sup")

    def __init__(self, base, sub, sup):
        self.base, self.sub, self.sup = base, sub, sup


class Frac(_Node):
    __slots__ = ("num", "den")

    def __init__(self, num, den):
        self.num, self.den = num, den


class Rad(_Node):
    __slots__ = ("deg", "rad")

    def __init__(self, deg, rad):
        self.deg, self.rad = deg, rad


class Nary(_Node):
    __slots__ = ("op", "sub", "sup", "body")

    def __init__(self, op, sub, sup, body):
        self.op, self.sub, self.sup, self.body = op, sub, sup, body


class Delim(_Node):
    __slots__ = ("left", "right", "body")

    def __init__(self, left, right, body):
        self.left, self.right, self.body = left, right, body


class Acc(_Node):
    __slots__ = ("chr", "base")

    def __init__(self, chr, base):
        self.chr, self.base = chr, base


# ── 词法 ───────────────────────────────────────────────────────────────────

_TOKEN_RE = re.compile(r"""
    \\(?:begin|end)\s*\{[^}]*\}      # \begin{...} / \end{...} → 显式不支持
  | \\[A-Za-z]+                      # 命令
  | \\[ ,;!:{}|\\%&#$_^~"-]          # 转义 / 单字符命令
  | \d+(?:\.\d+)?                    # 数字
  | [A-Za-z]                         # 字母
  | \s+                              # 空白
  | .                                # 其它单字符
""", re.X)

_TOK_NUM = "num"
_TOK_LET = "let"
_TOK_CMD = "cmd"
_TOK_CHR = "chr"
_TOK_OPEN = "open"
_TOK_CLOSE = "close"
_TOK_SUP = "sup"
_TOK_SUB = "sub"
_TOK_AMP = "amp"


def _tokenize(src: str) -> list:
    out: list = []
    for m in _TOKEN_RE.finditer(src):
        s = m.group(0)
        if s[0] == "\\":
            if s.startswith("\\begin") or s.startswith("\\end"):
                raise LatexUnsupported(f"不支持环境：{s.strip()}")
            if len(s) == 2 and not s[1].isalpha():
                if s == "\\," or s == "\\;" or s == "\\:" or s == "\\!" or s == "\\ ":
                    out.append((_TOK_CMD, s[1:] if s != "\\ " else "space"))
                elif s[1] in "{}\\|%&#$_":
                    out.append((_TOK_LET, s[1:]))
                else:
                    out.append((_TOK_CHR, s[1:]))
            else:
                out.append((_TOK_CMD, s[1:]))
        elif s.isspace():
            continue
        elif s.isdigit() or (s[0].isdigit() and "." in s):
            out.append((_TOK_NUM, s))
        elif s.isalpha():
            out.append((_TOK_LET, s))
        elif s == "{":
            out.append((_TOK_OPEN, s))
        elif s == "}":
            out.append((_TOK_CLOSE, s))
        elif s == "^":
            out.append((_TOK_SUP, s))
        elif s == "_":
            out.append((_TOK_SUB, s))
        elif s == "&":
            out.append((_TOK_AMP, s))
        else:
            out.append((_TOK_CHR, s))
    return out


# ── 语法 ───────────────────────────────────────────────────────────────────

class _Parser:
    def __init__(self, toks: list):
        self.toks = toks
        self.i = 0

    # — 基础 —
    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def next(self):
        t = self.peek()
        self.i += 1
        return t

    def eof(self) -> bool:
        return self.i >= len(self.toks)

    def parse_seq(self) -> list:
        """解析到末尾 / `}` / `\\right` 为止。"""
        kids: list = []
        while not self.eof():
            kind, val = self.peek()
            if kind == _TOK_CLOSE:
                break
            if kind == _TOK_CMD and val in ("right",):
                break
            if kind == _TOK_SUP or kind == _TOK_SUB:
                raise LatexUnsupported("上下标缺少底式（`^`/`_` 前面没有内容）")
            kids.append(self.parse_atom_with_scripts())
        return kids

    def parse_group(self):
        """`{...}` → Group；裸 token → 单节点。"""
        kind, val = self.peek()
        if kind == _TOK_OPEN:
            self.next()
            kids = self.parse_seq()
            k2, _v2 = self.peek()
            if k2 != _TOK_CLOSE:
                raise LatexUnsupported("花括号不配对")
            self.next()
            return Group(kids)
        return self.parse_atom_with_scripts()

    def parse_atom_with_scripts(self):
        base = self.parse_atom()
        kind, _v = self.peek()
        sub = sup = None
        while kind == _TOK_SUP or kind == _TOK_SUB:
            self.next()
            arg = self.parse_script_arg()
            if kind == _TOK_SUP:
                if sup is not None:
                    raise LatexUnsupported("重复上标")
                sup = arg
            else:
                if sub is not None:
                    raise LatexUnsupported("重复下标")
                sub = arg
            kind, _v = self.peek()
        if sub is None and sup is None:
            return base
        if sub is not None and sup is not None:
            return SubSup(base, sub, sup)
        return Sup(base, sup) if sup is not None else Sub(base, sub)

    def parse_script_arg(self):
        kind, _v = self.peek()
        if kind == _TOK_OPEN:
            return self.parse_group()
        if kind is None:
            raise LatexUnsupported("上下标后缺内容")
        return self.parse_atom()

    def parse_atom(self):
        kind, val = self.next()
        if kind == _TOK_CLOSE:
            raise LatexUnsupported("多余的 `}`")
        if kind == _TOK_OPEN:
            self.i -= 1
            return self.parse_group()
        if kind == _TOK_AMP:
            raise LatexUnsupported("不支持 `&`（表格/对齐环境）")
        if kind == _TOK_CMD:
            return self.handle_cmd(val)
        if kind == _TOK_NUM:
            return R(val, "p")
        if kind == _TOK_LET:
            return R(val, "i")
        if kind == _TOK_CHR:
            return self.handle_char(val)
        raise LatexUnsupported(f"无法解析的记号：{val!r}")

    def handle_char(self, ch: str):
        if ch in "()[]":
            # 隐式定界符：吃到**本层**配对的右界。
            # ⚠️ 不要自己维护括号深度：嵌套括号由 parse_atom_with_scripts 的递归
            # 自己消费掉（含内层右界），再计深度只会「只增不减」，导致外层右界
            # 认不出来而落到 parse_atom → KeyError: ')'（2026-09-28 冒烟实测）。
            pairs = {"(": ")", "[": "]", "|": "|"}
            close = pairs[ch]
            body = []
            while not self.eof():
                k, v = self.peek()
                if k == _TOK_CLOSE:
                    break
                if k == _TOK_CMD and v == "right":
                    break
                if k == _TOK_CHR and v == close:
                    self.next()
                    return Delim(ch, close, Group(body))
                body.append(self.parse_atom_with_scripts())
            raise LatexUnsupported(f"括号 {ch} 不配对")
        if ch in "=+-<>*/±":
            return R(ch, "p")
        return R(ch, "p")

    def handle_cmd(self, name: str):
        if name in _SPACES:
            return R(_SPACES[name], "p")
        if name in _GREEK:
            return R(_GREEK[name], "i")
        if name in _SYMBOLS:
            return R(_SYMBOLS[name], "p")
        if name in _FUNCS:
            return R(name, "p")
        if name in _NARY:
            sub = sup = None
            body = Group([])
            k, v = self.peek()
            if k == _TOK_SUB:
                self.next()
                sub = self.parse_script_arg()
                k, v = self.peek()
            if k == _TOK_SUP:
                self.next()
                sup = self.parse_script_arg()
                k, v = self.peek()
            if k == _TOK_OPEN:
                body = self.parse_group()
            return Nary(_NARY[name], sub, sup, body)
        if name == "frac" or name == "dfrac" or name == "tfrac":
            return Frac(self.parse_group(), self.parse_group())
        if name == "sqrt":
            deg = None
            k, _v = self.peek()
            if k == _TOK_CHR and _v == "[":
                self.next()
                deg_kids = []
                while not self.eof():
                    kk, vv = self.peek()
                    if kk == _TOK_CHR and vv == "]":
                        self.next()
                        break
                    deg_kids.append(self.parse_atom_with_scripts())
                deg = Group(deg_kids)
            return Rad(deg, self.parse_group())
        if name in _ACCENTS:
            return Acc(_ACCENTS[name], self.parse_group())
        if name in _TEXTY or name in _BOLDY or name in _ITALICY:
            body = self.parse_group()
            sty = "p" if name in _TEXTY else ("b" if name in _BOLDY else "i")
            return _restyle(body, sty)
        if name == "left":
            k, v = self.next()
            left = v if k in (_TOK_CHR, _TOK_LET) else "("
            body = self.parse_seq()
            k2, v2 = self.next()
            if k2 != _TOK_CMD or v2 != "right":
                raise LatexUnsupported("`\\left` 没有配对 `\\right`")
            k3, v3 = self.next()
            right = v3 if k3 in (_TOK_CHR, _TOK_LET) else ")"
            if left == "." or k == _TOK_LET and left == ".":
                left = ""
            if right == ".":
                right = ""
            return Delim(left, right, Group(body))
        if name == "right":
            raise LatexUnsupported("孤立的 `\\right`")
        if name in ("displaystyle", "textstyle", "limits", "nolimits", "quad"):
            return R("", "p")
        raise LatexUnsupported(f"不支持的命令：\\{name}")


def _restyle(node, sty: str):
    """把子树里的文本 run 改成指定字体姿态（`\\text{}` / `\\mathbf{}` 等）。"""
    if isinstance(node, R):
        return R(node.text, sty)
    if isinstance(node, Group):
        return Group([_restyle(k, sty) for k in node.kids])
    if isinstance(node, Sup):
        return Sup(_restyle(node.base, sty), node.sup)
    if isinstance(node, Sub):
        return Sub(_restyle(node.base, sty), node.sub)
    if isinstance(node, SubSup):
        return SubSup(_restyle(node.base, sty), node.sub, node.sup)
    if isinstance(node, Frac):
        return Frac(node.num, node.den)
    if isinstance(node, Rad):
        return Rad(node.deg, node.rad)
    if isinstance(node, Nary):
        return Nary(node.op, node.sub, node.sup, node.body)
    if isinstance(node, Delim):
        return Delim(node.left, node.right, node.body)
    if isinstance(node, Acc):
        return Acc(node.chr, _restyle(node.base, sty))
    return node


# ── 渲染（IR → OMML） ──────────────────────────────────────────────────────

def _el(tag: str, **attrs):
    e = etree.Element(tag, nsmap={"m": M, "w": W} if tag == _M + "oMath" else None)
    for k, v in attrs.items():
        e.set(k, v)
    return e


def _run(text: str, sty: str, size: str):
    r = etree.Element(_M + "r")
    rpr_m = etree.SubElement(r, _M + "rPr")
    if sty in ("i", "b"):
        s = etree.SubElement(rpr_m, _M + "sty")
        s.set(_M + "val", sty)
    else:
        s = etree.SubElement(rpr_m, _M + "sty")
        s.set(_M + "val", "p")
    rpr_w = etree.SubElement(r, _W + "rPr")
    fonts = etree.SubElement(rpr_w, _W + "rFonts")
    fonts.set(_W + "ascii", MATH_FONT)
    fonts.set(_W + "hAnsi", MATH_FONT)
    sz = etree.SubElement(rpr_w, _W + "sz")
    sz.set(_W + "val", size)
    szcs = etree.SubElement(rpr_w, _W + "szCs")
    szcs.set(_W + "val", size)
    t = etree.SubElement(r, _M + "t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return r


# 模板「符号：16 磅」= 公式里的**大运算符字符**（∑∫∏ 等）比正文符号大一档。
# Word 对 `m:nary` 的 chr 由数学区默认字号决定，单独给不了字号 —— 所以：
#   · `\sum` 这类走 m:nary 的：chr 字号给不了（如实在报告里说明），上下限走「次符号」；
#   · 直接写字符的 `∑` / `∫`：按「符号 16 磅」渲染。
_SYM_CHARS = set("∑∏∐∫∬∭∮⋃⋂⨁⨂⨆⋁⋀∏")


class _Renderer:
    def __init__(self, sizes: dict):
        self.sizes = sizes

    def sz(self, key: str) -> str:
        return self.sizes.get(key, DEFAULT_SIZES[key])

    def emit(self, node, host, level: int = 0, force_key: Optional[str] = None):
        """把 IR 节点渲染成 OMML 子元素，追加到 host。

        level 决定字号档：0 标准 / 1 上标下标 / 2 次上标下标；
        force_key 覆盖档位（nary 上下限 = 「次符号」10.5 磅）。
        """
        if isinstance(node, R):
            if node.text == "":
                return
            key = force_key or ("base" if level == 0 else
                                ("sup" if level == 1 else "subsup"))
            if force_key is None and level == 0 and node.text.strip() \
                    and all(c in _SYM_CHARS for c in node.text.strip()):
                key = "sym"      # 大运算符字符 → 「符号」16 磅
            host.append(_run(node.text, node.sty, self.sz(key)))
            return
        if isinstance(node, Group):
            for k in node.kids:
                self.emit(k, host, level, force_key)
            return
        if isinstance(node, Sup):
            e = etree.SubElement(host, _M + "sSup")
            base = etree.SubElement(e, _M + "e")
            self.emit(node.base, base, level)
            sup = etree.SubElement(e, _M + "sup")
            self.emit(node.sup, sup, level + 1)
            return
        if isinstance(node, Sub):
            e = etree.SubElement(host, _M + "sSub")
            base = etree.SubElement(e, _M + "e")
            self.emit(node.base, base, level)
            sub = etree.SubElement(e, _M + "sub")
            self.emit(node.sub, sub, level + 1)
            return
        if isinstance(node, SubSup):
            e = etree.SubElement(host, _M + "sSubSup")
            base = etree.SubElement(e, _M + "e")
            self.emit(node.base, base, level)
            sub = etree.SubElement(e, _M + "sub")
            self.emit(node.sub, sub, level + 1)
            sup = etree.SubElement(e, _M + "sup")
            self.emit(node.sup, sup, level + 1)
            return
        if isinstance(node, Frac):
            e = etree.SubElement(host, _M + "f")
            pr = etree.SubElement(e, _M + "fPr")
            ty = etree.SubElement(pr, _M + "type")
            ty.set(_M + "val", "bar")
            num = etree.SubElement(e, _M + "num")
            self.emit(node.num, num, level)
            den = etree.SubElement(e, _M + "den")
            self.emit(node.den, den, level)
            return
        if isinstance(node, Rad):
            e = etree.SubElement(host, _M + "rad")
            pr = etree.SubElement(e, _M + "radPr")
            if node.deg is None:
                hide = etree.SubElement(pr, _M + "degHide")
                hide.set(_M + "val", "1")
            deg = etree.SubElement(e, _M + "deg")
            if node.deg is not None:
                self.emit(node.deg, deg, level + 1)
            el = etree.SubElement(e, _M + "e")
            self.emit(node.rad, el, level)
            return
        if isinstance(node, Nary):
            e = etree.SubElement(host, _M + "nary")
            pr = etree.SubElement(e, _M + "naryPr")
            ch = etree.SubElement(pr, _M + "chr")
            ch.set(_M + "val", node.op)
            lim = etree.SubElement(pr, _M + "limLoc")
            lim.set(_M + "val", "undOvr" if node.op in "∑∏∐⋃⋂" else "subSup")
            sub = etree.SubElement(e, _M + "sub")
            if node.sub is not None:
                self.emit(node.sub, sub, level + 2, force_key="subsym")
            sup = etree.SubElement(e, _M + "sup")
            if node.sup is not None:
                self.emit(node.sup, sup, level + 2, force_key="subsym")
            body = etree.SubElement(e, _M + "e")
            self.emit(node.body, body, level)
            return
        if isinstance(node, Delim):
            e = etree.SubElement(host, _M + "d")
            pr = etree.SubElement(e, _M + "dPr")
            if node.left != "(":
                b = etree.SubElement(pr, _M + "begChr")
                b.set(_M + "val", node.left)
            if node.right != ")":
                c = etree.SubElement(pr, _M + "endChr")
                c.set(_M + "val", node.right)
            el = etree.SubElement(e, _M + "e")
            self.emit(node.body, el, level)
            return
        if isinstance(node, Acc):
            e = etree.SubElement(host, _M + "acc")
            pr = etree.SubElement(e, _M + "accPr")
            ch = etree.SubElement(pr, _M + "chr")
            ch.set(_M + "val", node.chr)
            el = etree.SubElement(e, _M + "e")
            self.emit(node.base, el, level)
            return
        raise LatexUnsupported(f"未知 IR 节点：{type(node).__name__}")


# ── 对外 API ───────────────────────────────────────────────────────────────

def latex_to_omml(latex: str, sizes: Optional[dict] = None):
    """把 LaTeX 数学表达式编译成 `<m:oMath>` 元素（不支持的构造抛 LatexUnsupported）。"""
    if latex is None:
        raise LatexUnsupported("空表达式")
    src = latex.strip()
    # 去掉常见的 `$...$` / `\\(...\\)` / `\\[...\\]` 包裹
    src = re.sub(r"^\$+|\$+$", "", src).strip()
    src = re.sub(r"^\\\(|\\\)$", "", src).strip()
    src = re.sub(r"^\\\[|\\\]$", "", src).strip()
    if not src:
        raise LatexUnsupported("空表达式")
    toks = _tokenize(src)
    kids = _Parser(toks).parse_seq()
    if not kids:
        raise LatexUnsupported("空表达式")
    omml = etree.Element(_M + "oMath", nsmap={"m": M, "w": W})
    _Renderer(dict(DEFAULT_SIZES, **(sizes or {}))).emit(Group(kids), omml, 0)
    return omml


def latex_to_omml_xml(latex: str, sizes: Optional[dict] = None) -> str:
    """同上，返回 XML 字符串（调试 / 快照用）。"""
    return etree.tostring(latex_to_omml(latex, sizes), encoding="unicode")


def is_supported(latex: str) -> bool:
    """能不能编译（给 Agent / 建议层先探路用，不抛异常）。"""
    try:
        latex_to_omml(latex)
        return True
    except LatexUnsupported:
        return False
    except Exception:  # noqa: BLE001 — 任何解析异常一律视为不支持
        return False


def omml_to_text(node) -> str:
    """粗略还原（测试断言用）：把 OMML 里的文本按结构拼回可读串。"""
    out: list = []
    for ch in node.iter():
        if ch.tag == _M + "t" and ch.text:
            out.append(ch.text)
        elif ch.tag == _M + "sSup" or ch.tag == _M + "sSubSup":
            pass
    return "".join(out)
