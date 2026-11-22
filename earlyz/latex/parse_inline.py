"""L2 解析（行内）：token → Inline。块级在 parse_blocks.py（继承本类）。

分组语义按 TeX：`{` 压一帧，`}` 弹回；\\bfseries 这类声明只在当前组内生效。
对齐（\\centering）也记在同一个栈里（"@center"），随组结束失效。
"""
from __future__ import annotations

import re
from collections import Counter, deque

from .ir import Inline
from .macros import _skip_sp, detok, plain, read_group, read_opt, read_star
from .siunitx import SI_CMDS, si
from .parse_text import (ACCENTS, CITE_CMDS, COUNTER_FMT, DIMEN_CMDS, DUMMY_TEXT, IGNORE_ARGS, NOOP,
                         REF_CMDS, STYLE_CMDS, STYLE_DECLS, STYLE_RESETS, TEXT_SYMBOLS, TRANSPARENT_ENVS,
                         InlineBuf, accent, inline_text)
from .tokens import Tok


def peek_env(q: deque) -> tuple[str, int]:
    """不消耗地读 `{name}`：返回 (name, 占用 token 数)；不是环境名返回 ('', 0)。"""
    i = 0
    while i < len(q) and q[i].k == "space":
        i += 1
    if i >= len(q) or q[i].k != "bg":
        return "", 0
    name, i = [], i + 1
    while i < len(q) and q[i].k != "eg":
        if q[i].k != "char":
            return "", 0
        name.append(q[i].v)
        i += 1
    return "".join(name), i + 1

# 先跳过若干参数（O=可选 M=必选），再把下一组当正文透明解析
CONTENT_AFTER = {"textcolor": "OM", "colorbox": "OM", "fcolorbox": "OMM", "raisebox": "MOO",
                 "makebox": "OO", "framebox": "OO", "parbox": "OOOM", "resizebox": "MM",
                 "scalebox": "MO", "rotatebox": "OM", "adjustbox": "M", "hyperlink": "M",
                 "hypertarget": "M", "smash": "O", "uppercase": "", "MakeUppercase": "",
                 "MakeLowercase": "", "lowercase": "", "textls": "O", "CJKunderline": "O",
                 "sout": "", "xout": "", "uwave": "", "boxed": "", "fcolorbox*": "OMM",
                 # plain TeX 盒子 / 旧式排版命令：内容照留（对齐、盒子尺寸属版式）
                 "vbox": "", "hbox": "", "vtop": "", "centerline": "", "leftline": "", "rightline": "",
                 "parpic": "OO"}
_UNITS = ("pt", "cm", "mm", "em", "ex", "bp", "in", "pc", "sp", "dd", "cc", "mu", "fil", "fill")


def char_code(q: deque, v: str) -> str:
    """\\char`\\\\ · \\char"5C · \\char'134 · \\char92 · \\symbol{"4E00} → 对应字符；读不出返回空串。"""
    toks = deque(read_group(q) or []) if v == "symbol" else q
    _skip_sp(toks)
    if not toks:
        return ""
    if toks[0].ch("`"):
        toks.popleft()
        if not toks:
            return ""
        n = toks.popleft()
        return n.v if n.k != "cs" or len(n.v) == 1 else ""
    base, digits = 10, "0123456789"
    if toks[0].ch('"'):
        base, digits = 16, "0123456789ABCDEFabcdef"
        toks.popleft()
    elif toks[0].ch("'"):
        base, digits = 8, "01234567"
        toks.popleft()
    s = ""
    while toks and toks[0].k == "char" and toks[0].v in digits:
        s += toks.popleft().v
    try:
        return chr(int(s, base)) if s else ""
    except (ValueError, OverflowError):
        return ""


def _math_end(tok: Tok, closer: str) -> bool:
    return (tok.k == "math") if closer == "$" else tok.cs(closer)


class InlineParser:
    def __init__(self) -> None:
        self.unsupported: Counter = Counter()
        self.notes: list[str] = []
        self.style: list[str] = []
        self.marks: list[int] = []          # 每个未闭合的 { 记下进入时 style 的长度
        self.floor = 0                      # 当前帧可弹的 marks 下限（多余的 } 不越过帧边界）
        self.last_target = None             # LaTeX \refstepcounter 语义：\label 指向最近的编号对象
        self.orphan_labels: list[str] = []
        self.float_ctx = None               # 当前所在浮动体（\caption \label \includegraphics 归它）
        self.si_range = " to "              # siunitx range-phrase（document.py 从 \sisetup 读）

    # ── 状态 ──────────────────────────────────────────────────────
    def styles(self) -> tuple:
        out: list[str] = []
        for s in self.style:
            if not s or s[0] == "@":
                continue
            if s == "!":
                out.clear()
            elif s[0] == "-":
                if s[1:] in out:
                    out.remove(s[1:])
            elif s not in out:
                out.append(s)
        return tuple(out)

    def align(self) -> str:
        for s in reversed(self.style):
            if s and s[0] == "@":
                return s[1:]
        return ""

    def add_label(self, key: str) -> None:
        if not key:
            return
        tgt = self.float_ctx if self.float_ctx is not None and self.last_target is None else self.last_target
        if tgt is not None and hasattr(tgt, "labels"):
            tgt.labels.append(key)
        else:
            self.orphan_labels.append(key)

    def sub_blocks(self, toks) -> list:          # 块级解析器覆盖
        return []

    def take_env(self, name: str, q: deque) -> list:
        """读到与之配平的 \\end{name}（同名嵌套计深度），返回环境体。"""
        out, depth = [], 0
        while q:
            t = q.popleft()
            if t.k == "cs" and t.v in ("begin", "end"):
                nm, n = peek_env(q)
                if nm == name:
                    if t.v == "begin":
                        depth += 1
                    elif depth == 0:
                        for _ in range(n):
                            q.popleft()
                        return out
                    else:
                        depth -= 1
            out.append(t)
        self.notes.append("环境 " + name + " 没有对应的 \\end，已读到文末")
        return out

    # ── 工具 ──────────────────────────────────────────────────────
    def read_math(self, q: deque, closer: str) -> list[Tok]:
        out, depth = [], 0
        while q:
            t = q.popleft()
            if depth == 0 and _math_end(t, closer):
                if closer == "$$" and q and q[0].k == "math":
                    q.popleft()
                return out
            if t.k == "bg":
                depth += 1
            elif t.k == "eg":
                depth -= 1
            out.append(t)
        self.notes.append("有公式没有闭合（缺少 " + closer + "），已读到文末")
        return out

    def skip_dimen(self, q: deque) -> None:
        _skip_sp(q)
        if q and q[0].k == "char" and q[0].v == "=":
            q.popleft()
            _skip_sp(q)
        if q and q[0].k == "cs":
            q.popleft()
            return
        while q and q[0].k == "char" and (q[0].v.isdigit() or q[0].v in ".,-+"):
            q.popleft()
        _skip_sp(q)
        if q and q[0].k == "cs":            # 2\baselineskip
            q.popleft()
            return
        word = ""
        while q and q[0].k == "char" and q[0].v.isalpha() and len(word) < 4:
            word += q.popleft().v
            if word in _UNITS and not (q and q[0].k == "char" and (word + q[0].v) in _UNITS):
                break

    def enter(self) -> tuple:
        frame = (len(self.style), len(self.marks), self.floor)
        self.floor = len(self.marks)
        return frame

    def leave(self, frame: tuple) -> None:
        del self.style[frame[0]:]
        del self.marks[frame[1]:]
        self.floor = frame[2]

    def inline_tokens(self, toks, buf: InlineBuf) -> None:
        q = deque(toks)
        frame = self.enter()
        while q:
            t = q.popleft()
            if not self.inline(t, q, buf):
                if t.k == "par":
                    buf.text(" ", self.styles())
                elif t.k == "raw":
                    buf.add(Inline("code", t.v[-1]))
                else:
                    self.unsupported["行内出现块级命令 \\" + str(t.v)] += 1
        self.leave(frame)

    def inlines(self, toks) -> list[Inline]:
        buf = InlineBuf()
        self.inline_tokens(toks or [], buf)
        return buf.take()

    # ── 行内 ──────────────────────────────────────────────────────
    def inline(self, tok: Tok, q: deque, buf: InlineBuf) -> bool:
        k = tok.k
        if k in ("char", "align", "param", "sup", "sub"):
            buf.text(tok.v, self.styles())
        elif k == "space":
            buf.text(" ", self.styles())
        elif k == "active":
            buf.text("\u00a0", self.styles())
        elif k == "bg":
            self.marks.append(len(self.style))
        elif k == "eg":
            if len(self.marks) > self.floor:
                del self.style[self.marks.pop():]
        elif k == "math":
            if q and q[0].k == "math":
                return False                 # $$ 行间公式归块级
            buf.add(Inline("math", detok(self.read_math(q, "$")).strip()))
        elif k == "raw" and tok.v[0] == "verb":
            buf.add(Inline("code", tok.v[1]))
        elif k == "cs":
            return self.inline_cs(tok, q, buf)
        else:
            return False
        return True

    def inline_cs(self, tok: Tok, q: deque, buf: InlineBuf) -> bool:
        v, st = tok.v, self.styles()
        if v in ACCENTS and (len(v) == 1 or v in "uvHckrbd"):
            arg = read_group(q) or []
            s = "".join(TEXT_SYMBOLS.get(t.v, "") if t.k == "cs" else t.v for t in arg)
            buf.text(accent(s, ACCENTS[v]), st)
        elif v in TEXT_SYMBOLS:
            buf.text(TEXT_SYMBOLS[v], st)
        elif v in ("\\", "newline", "linebreak", "tabularnewline"):
            read_star(q)
            read_opt(q)
            buf.add(Inline("br"))
        elif v == "(":
            buf.add(Inline("math", detok(self.read_math(q, ")")).strip()))
        elif v == "ensuremath":
            buf.add(Inline("math", detok(read_group(q) or []).strip()))
        elif v in STYLE_CMDS:
            arg = read_group(q) or []
            mark = len(self.style)
            self.style.append(STYLE_CMDS[v] or "")
            self.inline_tokens(arg, buf)
            del self.style[mark:]
        elif v in CONTENT_AFTER:
            read_star(q)
            for s in CONTENT_AFTER[v]:
                (read_opt if s == "O" else read_group)(q)
            self.inline_tokens(read_group(q) or [], buf)
        elif v == "texorpdfstring":
            self.inline_tokens(read_group(q) or [], buf)
            read_group(q)
        elif v in STYLE_DECLS:
            self.style.append(STYLE_DECLS[v])
        elif v in STYLE_RESETS:
            self.style.append("!" if STYLE_RESETS[v] is None else "-" + STYLE_RESETS[v])
        elif v in ("centering", "raggedright", "raggedleft"):
            self.style.append({"centering": "@center", "raggedright": "@left", "raggedleft": "@right"}[v])
        elif v in CITE_CMDS:
            read_star(q)
            a, b = read_opt(q), read_opt(q)
            pre, post = (a, b) if b is not None else (None, a)
            keys = [x.strip() for x in plain(read_group(q) or []).split(",") if x.strip()]
            buf.add(Inline("cite", data={"keys": keys, "cmd": v, "pre": inline_text(self.inlines(pre)),
                                         "post": inline_text(self.inlines(post))}))
        elif v in REF_CMDS:
            read_star(q)
            for key in plain(read_group(q) or []).split(","):
                if key.strip():
                    buf.add(Inline("ref", key.strip(), data={"cmd": v}))
        elif v == "label":
            self.add_label(plain(read_group(q) or []))
        elif v in ("footnote", "footnotetext"):
            read_opt(q)
            buf.add(Inline("footnote", data={"blocks": self.sub_blocks(read_group(q) or [])}))
        elif v == "footnotemark":
            read_opt(q)
        elif v in ("url", "path", "nolinkurl"):
            arg = q.popleft().v[1] if q and q[0].k == "raw" else plain(read_group(q) or [])
            buf.add(Inline("link" if v == "url" else "code", arg, data={"text": arg}))
        elif v == "href":
            url = q.popleft().v[1] if q and q[0].k == "raw" else plain(read_group(q) or [])
            buf.add(Inline("link", url, data={"text": inline_text(self.inlines(read_group(q)))}))
        elif v in ("includegraphics", "psfig", "epsfig"):
            read_star(q)
            if v == "includegraphics":
                opts = detok(read_opt(q) or [])
                img = {"path": plain(read_group(q) or []), "opts": opts}
            else:                                   # 旧式 \\psfig{figure=a.eps,width=7cm}
                opts = plain(read_group(q) or [])
                m = re.search(r"(?:figure|file)\s*=\s*([^,]+)", opts)
                img = {"path": m.group(1).strip() if m else "", "opts": opts}
            if self.float_ctx is not None:
                self.float_ctx.images.append(img)
            else:
                buf.add(Inline("image", data=img))
        elif v in SI_CMDS:
            buf.text(si(v, q, self.si_range), st)
        elif v in COUNTER_FMT:
            buf.add(Inline("counter", plain(read_group(q) or []), data={"fmt": v}))
        elif v.startswith("the") and len(v) > 3:
            buf.add(Inline("counter", v[3:], data={"fmt": "the"}))
        elif v == "today":
            buf.add(Inline("raw", "\\today"))
        elif v in IGNORE_ARGS:
            read_star(q)
            o, m = IGNORE_ARGS[v]
            for _ in range(o):
                read_opt(q)
            for _ in range(m):
                read_group(q)
        elif v in DIMEN_CMDS:
            self.skip_dimen(q)
        elif v in ("begingroup", "endgroup"):      # 与 { } 同为分组
            self.inline(Tok("bg" if v == "begingroup" else "eg", "{" if v == "begingroup" else "}"), q, buf)
        elif v in ("char", "symbol"):
            buf.text(char_code(q, v), st)
        elif v == "string":                         # \\string\\cite → 文字「\\cite」
            if q:
                n = q.popleft()
                buf.text("\\" + n.v if n.k == "cs" else str(n.v), st)
        elif v == "par":                            # 行内（单元格 / 样式参数）里的分段 = 换行
            buf.add(Inline("br"))
        elif v in ("begin", "end"):
            self.inline_env(v, q, buf)
        elif v in DUMMY_TEXT:                       # \\lipsum \\zhlipsum \\kant：示例占位文字，留命令不展开
            read_star(q)
            read_opt(q)
            buf.add(Inline("raw", "\\" + v))
        elif v == "tikzstyle":                      # \\tikzstyle{名}=[样式]
            read_group(q)
            _skip_sp(q)
            if q and q[0].ch("="):
                q.popleft()
            read_opt(q)
        elif v in ("subfigure", "subtable", "subfloat"):   # 浮动体外的旧 subfigure 写法：内容 + 分图题注照留
            a, b = read_opt(q), read_opt(q)
            cap = b if b is not None else a
            self.inline_tokens(read_group(q) or [], buf)
            if cap:
                buf.text(" ", st)
                self.inline_tokens(cap, buf)
        elif v in NOOP or v == "noindent":
            pass
        else:
            self.unsupported["\\" + v] += 1
        return True

    def inline_env(self, v: str, q: deque, buf: InlineBuf) -> None:
        """行内（表格单元格、样式命令参数里）遇到环境：透明环境（CJK* 等）当分组；其余取环境体按行内内容解析。"""
        name, n = peek_env(q)
        if not name:
            self.unsupported["\\" + v] += 1
            return
        for _ in range(n):
            q.popleft()
        if v == "end":
            if name in TRANSPARENT_ENVS and len(self.marks) > self.floor:
                del self.style[self.marks.pop():]
            return
        if name in TRANSPARENT_ENVS:
            for _ in range(TRANSPARENT_ENVS[name]):
                read_group(q)
            self.marks.append(len(self.style))
            return
        self.unsupported["行内环境 " + name + "（按行内内容处理）"] += 1
        self.inline_tokens(self.take_env(name, q), buf)
