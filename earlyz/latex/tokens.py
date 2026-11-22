"""L1 词法：LaTeX 源码 → token 流（TeX catcode 机制的常用子集，The TeXbook 第 8 章）。

- `\\` 后跟 ASCII 字母串 = 控制词（其后空白和一个行尾吞掉）；后跟单个非字母 = 控制符。
  只认 ASCII 字母：XeTeX 把汉字也当字母，`\\LaTeX中文` 会粘成一个命令名，作者实际都会
  加空格或 {}，这里按更宽容的 pdfTeX 口径切开。
- 行尾 = 空格；空行 = \\par；行首空白跳过；`%` 到行尾是注释（连同行尾）。
- verbatim 类内容在词法层整段取原文（\\verb、verbatim/lstlisting/minted/comment 环境、
  \\url \\href 的地址参数），否则里面的 `%` `#` `\\` 会被当成命令。
"""
from __future__ import annotations

import re

VERB_ENVS = frozenset({"verbatim", "verbatim*", "Verbatim", "BVerbatim", "lstlisting",
                       "minted", "comment", "spverbatim"})
_ENV_RE = re.compile(r"[ \t]*\{([A-Za-z*]+)\}")
_VERB_DEFINERS = {"lstnewenvironment": "", "DefineVerbatimEnvironment": "", "newminted": "code",
                  "newtcblisting": "", "NewListingEnvironment": ""}
_NAME_RE = re.compile(r"[ \t]*(?:\[[^\]]*\][ \t]*)?\{([A-Za-z*]+)\}")
_CATCODE_RE = re.compile(r"[ \t]*`\\?(.)[ \t]*=?[ \t]*(\d+)[ \t]*")
_SPECIAL = {"{": "bg", "}": "eg", "$": "math", "&": "align", "#": "param",
            "^": "sup", "_": "sub", "~": "active"}


class Tok:
    """k: cs | char | space | par | bg | eg | math | align | param | sup | sub | active | raw。
    raw 的 v：('verb', 文本) / ('env', 环境名, 文本) / ('arg', 文本)。"""
    __slots__ = ("k", "v", "line", "file")

    def __init__(self, k: str, v="", line: int = 0, file: str = ""):
        self.k, self.v, self.line, self.file = k, v, line, file

    def cs(self, *names) -> bool:
        return self.k == "cs" and (not names or self.v in names)

    def ch(self, c: str) -> bool:
        return self.k == "char" and self.v == c

    def __repr__(self) -> str:
        return f"{self.k}:{self.v!r}"


def _letter(c: str, at: bool) -> bool:
    return ("a" <= c <= "z") or ("A" <= c <= "Z") or (at and c == "@")


def _match_brace(text: str, k: int) -> int:
    """text[k] == '{' → 配平的 '}' 位置（反斜杠转义跳过）；不配平返回 len(text)。"""
    depth, i, n = 0, k, len(text)
    while i < n:
        c = text[i]
        if c == "\\":
            i += 2
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return i
        i += 1
    return n


def lex(text: str, file: str = "", at_letter: bool = False, verb_envs: set | None = None) -> list[Tok]:
    # verb_envs 由调用方共享：cls / setup.tex 里 \\lstnewenvironment 定义的代码环境，正文文件里才用到
    """verb_envs：可变集合，跨文件共享（导言区/setup.tex 里定义的代码环境，正文文件里才用）。"""
    venvs = verb_envs if verb_envs is not None else set(VERB_ENVS)
    text = text.replace("\r\n", "\n").replace("\r", "\n")      # CRLF / CR 行尾统一（verbatim 原文里不留 \r）
    out: list[Tok] = []
    n, i, line = len(text), 0, 1
    at = at_letter
    bol = True                      # 行首（TeX state N）：空白跳过，再遇行尾 = \par

    def emit(k: str, v="") -> None:
        out.append(Tok(k, v, line, file))

    while i < n:
        c = text[i]
        if c == "\r":
            i += 1
            continue
        if c == "\n":
            line += 1
            i += 1
            if bol:
                if out and out[-1].k != "par":
                    emit("par")
            else:
                if out and out[-1].k != "space":
                    emit("space", " ")
                bol = True
            continue
        if c in " \t":
            i += 1
            if not bol and out and out[-1].k not in ("space", "par"):
                emit("space", " ")
            continue
        if c == "%":
            j = text.find("\n", i)
            if j < 0:
                break
            i, line, bol = j + 1, line + 1, True
            continue
        bol = False
        if c == "\\":
            j = i + 1
            if j >= n:
                break
            if _letter(text[j], at):
                k = j
                while k < n and _letter(text[k], at):
                    k += 1
                name = text[j:k]
                while k < n and text[k] in " \t\r":
                    k += 1
                if k < n and text[k] == "\n":
                    k, line, bol = k + 1, line + 1, True
                i = k
            else:
                name, i = text[j], j + 1
                if name == "\n":
                    line += 1
                    name = " "
            if name == "makeatletter":
                at = True
            elif name == "makeatother":
                at = False
            elif name == "catcode":
                m = _CATCODE_RE.match(text, i)
                if m:
                    if m.group(1) == "@":
                        at = m.group(2) == "11"
                    i = m.end()
                    continue
            elif name in ("verb", "lstinline"):
                k = i
                if k < n and text[k] == "*":
                    k += 1
                if name == "lstinline" and k < n and text[k] == "[":
                    k = text.find("]", k) + 1 or n
                if k < n:
                    d = text[k]
                    e = _match_brace(text, k) if d == "{" else text.find(d, k + 1)
                    e = n if e < 0 else e
                    emit("raw", ("verb", text[k + 1:e]))
                    i = e + 1
                    continue
            elif name in ("url", "path", "href", "nolinkurl"):
                k = i
                while k < n and text[k] in " \t":
                    k += 1
                if k < n and text[k] == "{":
                    e = _match_brace(text, k)
                    emit("cs", name)
                    emit("raw", ("arg", text[k + 1:e]))
                    line += text.count("\n", k, e)
                    i = e + 1
                    continue
            elif name in _VERB_DEFINERS:
                m = _NAME_RE.match(text, i)
                if m:
                    venvs.add(m.group(1) + _VERB_DEFINERS[name])
            elif name == "begin":
                m = _ENV_RE.match(text, i)
                if m and m.group(1) in venvs:
                    env = m.group(1)
                    tag = "\\end{" + env + "}"
                    e = text.find(tag, m.end())
                    e = n if e < 0 else e
                    emit("raw", ("env", env, text[m.end():e]))
                    line += text.count("\n", m.end(), e)
                    i = e + len(tag)
                    continue
            emit("cs", name)
            continue
        k = _SPECIAL.get(c)
        emit(k or "char", c)
        i += 1
    return out
