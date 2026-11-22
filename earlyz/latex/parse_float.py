"""L2 解析（浮动体 / 行间公式）。

浮动体：\\caption \\label \\includegraphics tabular 归最近的浮动体；子图（subfigure 环境、
\\subfloat、\\subcaptionbox）挂 subfloats；有题注的 minipage 是独立编号的图（LaTeX 每个 \\caption
各占一个编号），没题注的 minipage 只是排版容器，内容并回外层。
双语题注：\\bicaption{中}{英} 直接取；cls 自定义的「题注 + 第二行 Fig.\\thefigure 英文」形态
按段里出现本浮动体计数器识别（不认命令名）。
公式：多行环境按顶层 \\\\ 切行，逐行记 \\label \\tag \\nonumber；编号在 numbering.py 推算。
"""
from __future__ import annotations

from collections import deque

from .ir import Equation, Float, Para, Raw
from .macros import detok, plain, read_group, read_opt, read_star
from .parse_table import split_top
from .parse_text import finalize, inline_text
from .tokens import Tok

MULTIROW_MATH = frozenset({"align", "align*", "gather", "gather*", "eqnarray", "eqnarray*", "flalign",
                           "flalign*", "alignat", "alignat*", "multline", "multline*"})
ONE_NUMBER = frozenset({"equation", "multline"})
UNNUMBERED_MATH = frozenset({"displaymath", "$$", "\\["})


def numbered_math(env: str) -> bool:
    return not env.endswith("*") and env not in UNNUMBERED_MATH


class FloatMixin:
    # ── 浮动体 ────────────────────────────────────────────────────
    def parse_float(self, env: str, ftype: str, q: deque, out: list, line: int) -> None:
        if env in ("wrapfigure", "wraptable"):
            read_opt(q)
            read_group(q)
            read_opt(q)
            read_group(q)
        else:
            read_opt(q)
        fl = Float(ftype, env, line=line)
        body = self.take_env(env, q)
        if ftype == "algorithm":
            body = self.split_algorithm(fl, env, body)
        self.fill_float(fl, body)
        out.append(fl)

    def split_algorithm(self, fl: Float, env: str, body: list) -> list:
        """algorithm 浮动体：题注 / 标签照常解析，其余是伪代码（algorithm2e 的 \\KwData \\While 直接写在
        浮动体里，algorithmic 写在内层环境里），整段留原文（Raw），由下游编译成图或原样保留。"""
        keep, rest, q, depth = [], [], deque(body), 0
        while q:
            t = q.popleft()
            if t.k == "bg":
                depth += 1
            elif t.k == "eg":
                depth -= 1
            if depth == 0 and t.cs("caption", "bicaption", "label"):
                keep.append(t)
                if q and q[0].ch("*"):
                    keep.append(q.popleft())
                opt = read_opt(q)
                if opt is not None:
                    keep += [Tok("char", "[")] + opt + [Tok("char", "]")]
                for _ in range(2 if t.v == "bicaption" else 1):
                    g = read_group(q)
                    if g is not None:
                        keep += [Tok("bg", "{")] + g + [Tok("eg", "}")]
                continue
            rest.append(t)
        src = detok(rest).strip()
        if src:
            fl.extra.append(Raw(src, "算法伪代码", env))
        return keep

    def fill_float(self, fl: Float, body: list) -> None:
        saved = (self.float_ctx, self.last_target)
        self.float_ctx, self.last_target = fl, None
        try:
            inner = self.blocks(body)
        finally:
            self.float_ctx, self.last_target = saved
        for b in inner:
            if isinstance(b, Para) and not fl.caption_alt and self._alt_caption(fl, b):
                continue
            fl.extra.append(b)
        fl.numbered = fl.numbered and bool(fl.caption)

    def _alt_caption(self, fl: Float, p: Para) -> bool:
        kind = fl.type
        idx = [i for i, x in enumerate(p.inlines) if x.t == "counter" and x.v in (kind, "chapter")]
        if not any(p.inlines[i].v == kind for i in idx):
            return False
        fl.caption_alt_prefix = inline_text(p.inlines[:idx[0]])
        fl.caption_alt = finalize(p.inlines[idx[-1] + 1:])
        return True

    def parse_subenv(self, env: str, q: deque) -> None:
        """subfigure / subtable / minipage（浮动体内）。"""
        parent = self.float_ctx
        read_opt(q)
        if env == "minipage":
            read_opt(q)
            read_opt(q)
        read_group(q)                                   # 宽度
        sub = Float("sub" if env != "minipage" else parent.type, env)
        self.fill_float(sub, self.take_env(env, q))
        if env == "minipage" and not sub.caption:
            for name in ("images", "grids", "extra", "subfloats", "notes", "labels"):
                getattr(parent, name).extend(getattr(sub, name))
            return
        parent.subfloats.append(sub)

    def subfloat_cmd(self, v: str, q: deque) -> None:
        """\\subfloat[目录项][题注]{内容} · \\subcaptionbox{题注}[宽]{内容} · \\bisubcaptionbox{中}{英}[宽]{内容}。"""
        sub, alt = Float("sub", v), None
        if v in ("subfloat", "subfigure", "subtable"):
            a, b = read_opt(q), read_opt(q)
            cap = b if b is not None else a
        else:
            read_star(q)
            if v == "subcaptionbox":
                read_opt(q)
            cap = read_group(q)
            if v == "bisubcaptionbox":
                alt = read_group(q)
            read_opt(q)
            read_opt(q)
        body = read_group(q) or []
        parent = self.float_ctx
        saved = self.float_ctx, self.last_target
        self.float_ctx, self.last_target = sub, sub        # 题注里的 \\label 归分图（subcaption 语义）
        sub.caption = self.inlines(cap) if cap else []
        sub.caption_alt = self.inlines(alt) if alt else []
        self.float_ctx, self.last_target = saved
        self.fill_float(sub, body)
        parent.subfloats.append(sub)

    def caption_cmd(self, v: str, q: deque, out: list) -> None:
        fl = self.float_ctx
        if v == "captionof":
            star = read_star(q)
            ftype = plain(read_group(q) or [])
            if fl is None:                              # minipage 里的 \captionof：自成一个图/表
                fl = Float(ftype if ftype in ("figure", "table") else "figure", "captionof")
                out.append(fl)
        else:
            star = read_star(q)
        short = read_opt(q)
        text = read_group(q) or []
        alt = None
        if v == "bicaption":
            read_opt(q)
            alt = read_group(q) or []
        if fl is None:
            self.unsupported["浮动体外的 \\" + v] += 1
            out.append(Para(self.inlines(text)))
            return
        if fl.caption:
            if fl.caption_starred and not star:     # \\caption* 说明在前、\\caption 编号题注在后：前者转为说明段
                fl.extra.append(Para(fl.caption))
            elif star:                              # 已有题注，再来的 \\caption* 是说明文字
                fl.extra.append(Para(self.inlines(text)))
                return
            else:
                self.unsupported["一个浮动体里有多个题注"] += 1
                return
        fl.caption = self.inlines(text)
        fl.short = self.inlines(short) if short is not None else None
        if alt is not None:
            fl.caption_alt = self.inlines(alt)
        fl.numbered, fl.caption_starred = not star, star
        if not star:                                # \\caption* 不步进计数器，\\label 不指向它
            self.last_target = fl

    # ── 行间公式 ──────────────────────────────────────────────────
    def parse_equation(self, env: str, q: deque, out: list, line: int) -> None:
        if env.startswith("alignat"):
            read_group(q)
        self.equation(env, self.take_env(env, q), out, line)

    def equation(self, env: str, body: list, out: list, line: int) -> None:
        rows = split_top(body, "row") if env in MULTIROW_MATH else [body]
        while len(rows) > 1 and not any(t.k not in ("space", "par") for t in rows[-1]):
            rows.pop()
        eq = Equation("", env, line=line, group=self.subeq)
        srcs = []
        for i, r in enumerate(rows):
            keep, rq = [], deque(r)
            while rq:
                t = rq.popleft()
                if t.cs("label"):
                    eq.labels.append(plain(read_group(rq) or []))
                    eq.label_rows.append(i)
                elif t.cs("tag"):
                    read_star(rq)
                    eq.tags.append([i, detok(read_group(rq) or []).strip()])
                elif t.cs("nonumber", "notag"):
                    eq.nonumber_rows.append(i)
                else:
                    keep.append(t)
            srcs.append(detok(keep).strip())
        eq.rows = len(rows)
        eq.latex = " \\\\\n".join(srcs)
        if numbered_math(env):
            tagged = {r for r, _ in eq.tags}
            if env in ONE_NUMBER:
                eq.numbers = 0 if (eq.nonumber_rows or eq.tags) else 1
            else:
                eq.numbers = sum(1 for i in range(eq.rows) if i not in eq.nonumber_rows and i not in tagged)
        out.append(eq)
        self.last_target = eq
