"""L1 宏展开：定义收集 + 展开 + 条件（常用子集）。

- 定义：\\newcommand \\renewcommand \\providecommand \\DeclareRobustCommand（含 * 与 [n][default]）、
  \\def \\gdef \\edef \\xdef（只支持 #1…#9 无分隔参数；带分隔符的模式参数记 unsupported 不展开）、
  \\let、\\DeclareMathOperator(*)、\\newenvironment \\renewenvironment、\\newif、\\newtheorem 类。
- 展开：用户宏在使用处替换并重扫。protected（解析器原生认识的结构/前置命令）永不展开：
  cls 往往重定义 \\section、\\maketitle，展开就丢结构。
- cls/sty（工程内或本机 TeX 发行版里找到的）只收「公共」宏（lib=True）：名字或宏体带 @ 内部命令/
  排版原语的不收——那是 cls 的排版实现，不是内容；expl3 区段整段跳过。不收的留底在 Expander.raw，
  给 L2 判断「章标题类命令」「双语题注命令」和参数个数。
- 条件：\\newif 标志、\\iftrue/\\iffalse、\\ifdefined、\\ifcsname、\\ifx 求值；\\ifnum/\\ifdim/\\ifodd/\\ifcase
  整段丢弃；其余 \\if@xxx 取真分支。都记 notes。
"""
from __future__ import annotations

from collections import Counter, deque
from dataclasses import dataclass, field

from .tokens import Tok, lex

MAX_EXPANSIONS = 200_000
_DEFINERS = frozenset({"newcommand", "renewcommand", "providecommand", "DeclareRobustCommand",
                       "DeclareMathOperator", "def", "gdef", "edef", "xdef", "let", "newenvironment",
                       "renewenvironment", "newif", "newtheorem", "newdefinition", "newproof",
                       "theoremstyle", "NewDocumentCommand", "RenewDocumentCommand",
                       "ProvideDocumentCommand", "DeclareDocumentCommand", "csdef", "csgdef",
                       "csedef", "csxdef"})
_NUMERIC_IF = frozenset({"ifnum", "ifdim", "ifodd", "ifcase", "ifvmode", "ifhmode", "ifmmode",
                         "ifinner", "ifvoid", "ifhbox", "ifvbox", "ifeof", "if", "ifcat"})
_UNSAFE = frozenset({"expandafter", "csname", "endcsname", "global", "setbox", "hbox", "vbox", "vtop",
                     "kern", "penalty", "futurelet", "afterassignment", "aftergroup", "def", "gdef",
                     "edef", "xdef", "let", "noexpand", "the", "number", "romannumeral", "uppercase",
                     "lowercase", "string", "ExplSyntaxOn", "unexpanded", "detokenize", "box", "copy",
                     "lastbox", "unskip", "unpenalty", "hrule", "vrule", "special", "write", "immediate",
                     "openout", "closeout", "input", "endinput", "catcode", "makeatletter"})
_DETOK = {"bg": "{", "eg": "}", "math": "$", "align": "&", "param": "#", "sup": "^", "sub": "_",
          "active": "~", "space": " ", "par": "\n\n"}


@dataclass
class Macro:
    nargs: int = 0
    default: list | None = None
    body: list = field(default_factory=list)
    lib: bool = False


@dataclass
class EnvDef:
    nargs: int
    default: list | None
    begin: list
    end: list


@dataclass
class MacroTable:
    macros: dict = field(default_factory=dict)
    envs: dict = field(default_factory=dict)
    flags: dict = field(default_factory=dict)          # \newif\iffoo → flags["foo"]
    theorems: dict = field(default_factory=dict)       # 环境名 → {title, numbered, shared, within}
    notes: list = field(default_factory=list)
    unsupported: Counter = field(default_factory=Counter)
    expansions: int = 0


def detok(toks) -> str:
    """token → LaTeX 源码（公式交给 L4 用；控制词后紧跟字母时补空格）。"""
    out, word = [], False
    for t in toks:
        if t.k == "cs":
            out.append("\\" + t.v)
            word = t.v[:1].isalpha() or t.v[:1] == "@"
            continue
        if t.k == "char":
            s = (" " + t.v) if word and (t.v.isalpha() or t.v == "@") else t.v
        elif t.k == "raw":
            v = t.v
            s = ("\\verb|" + v[1] + "|" if v[0] == "verb"
                 else "\\begin{" + v[1] + "}" + v[2] + "\\end{" + v[1] + "}" if v[0] == "env"
                 else "{" + v[1] + "}")
        else:
            s = _DETOK.get(t.k, str(t.v))
        out.append(s)
        word = False
    return "".join(out)


_PLAIN_KINDS = frozenset({"char", "space", "sub", "sup", "align", "active"})


def plain(toks) -> str:
    """只取字符（定义名、环境名、键值、标签、文件名）。_ ^ & ~ 是标签/文件名里的普通字符，保留。"""
    return "".join(t.v if t.k in _PLAIN_KINDS else "" for t in toks).strip()


def _skip_sp(q: deque) -> None:
    while q and q[0].k == "space":
        q.popleft()


def read_group(q: deque) -> list | None:
    """无分隔参数：{…} 去外层括号；否则单个 token（TeX 规则，前导空格跳过）。"""
    _skip_sp(q)
    if not q:
        return None
    t = q.popleft()
    if t.k != "bg":
        return [t]
    depth, out = 1, []
    while q:
        t = q.popleft()
        if t.k == "bg":
            depth += 1
        elif t.k == "eg":
            depth -= 1
            if depth == 0:
                return out
        out.append(t)
    return out


def read_opt(q: deque, open_: str = "[", close: str = "]") -> list | None:
    j = 0
    while j < len(q) and q[j].k == "space":
        j += 1
    if j >= len(q) or not q[j].ch(open_):
        return None
    for _ in range(j + 1):
        q.popleft()
    depth, out = 0, []
    while q:
        t = q.popleft()
        if t.k == "bg":
            depth += 1
        elif t.k == "eg":
            depth -= 1
        elif depth == 0 and t.ch(close):
            return out
        out.append(t)
    return out


def read_star(q: deque) -> bool:
    if q and q[0].ch("*"):
        q.popleft()
        return True
    return False


def _cs_name(toks: list | None) -> str:
    for t in toks or ():
        if t.k == "cs":
            return t.v
        if t.k == "active":
            return "~"
    return ""


def _subst(body: list, args: list) -> list:
    out, i, n = [], 0, len(body)
    while i < n:
        t = body[i]
        if t.k == "param" and i + 1 < n:
            nx = body[i + 1]
            if nx.k == "param":
                out.append(nx)
                i += 2
                continue
            if nx.k == "char" and nx.v.isdigit() and 0 < int(nx.v) <= len(args):
                out.extend(args[int(nx.v) - 1])
                i += 2
                continue
        out.append(t)
        i += 1
    return out


def _safe(name: str, body: list) -> bool:
    if "@" in name:
        return False
    return not any(t.k == "cs" and ("@" in t.v or t.v in _UNSAFE) for t in body)
