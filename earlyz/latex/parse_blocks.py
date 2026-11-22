"""L2 解析（块级）：token → 块列表。行内/表格/浮动体分别在 parse_inline / parse_table / parse_float。

不认识的环境按普通内容透明解析（内容不丢），并记入 unsupported；TikZ、算法伪代码这类
没有文字语义的环境整段留原文（Raw，带原因），由下游决定编译成图还是原样保留。
"""
from __future__ import annotations

import re
from collections import deque

from .front_fields import BUILTIN
from .ir import (Abstract, Bibliography, BibItem, Code, Float, Heading, Keywords, ListBlock,
                 ListItem, Marker, Para, Proof, Raw, Theorem)
from .macros import _skip_sp, detok, plain, read_group, read_opt, read_star
from .parse_float import FloatMixin
from .parse_inline import CONTENT_AFTER, InlineParser
from .parse_text import STYLE_CMDS
from .parse_table import TABULAR_ENVS, TableMixin
from .parse_text import NOOP, SECTION_LEVEL, TRANSPARENT_ENVS, InlineBuf, inline_text, is_cjk_text
from .tokens import Tok, lex

FLOAT_ENVS = {"figure": "figure", "figure*": "figure", "table": "table", "table*": "table",
              "algorithm": "algorithm", "algorithm*": "algorithm", "wrapfigure": "figure",
              "wraptable": "table", "SCfigure": "figure", "sidewaysfigure": "figure",
              "sidewaystable": "table"}
SUB_ENVS = frozenset({"subfigure", "subtable", "minipage"})
LIST_ENVS = {"itemize": "itemize", "enumerate": "enumerate", "description": "description",
             "compactitem": "itemize", "compactenum": "enumerate", "inparaenum": "enumerate",
             "asparaenum": "enumerate", "tablenotes": "itemize", "TableNotes": "itemize",
             "highlights": "itemize"}
ALIGN_ENVS = {"center": "center", "flushleft": "left", "flushright": "right"}
QUOTE_ENVS = frozenset({"quote", "quotation", "verse"})
RAW_ENVS = {"tikzpicture": "TikZ 图（需编译成图片）", "pgfpicture": "PGF 图（需编译成图片）",
            "picture": "picture 绘图（需编译成图片）", "axis": "pgfplots 图（需编译成图片）",
            "algorithmic": "算法伪代码", "algorithmic*": "算法伪代码", "algorithm2e": "算法伪代码",
            "pseudocode": "算法伪代码", "forest": "树图（需编译成图片）", "circuitikz": "电路图（需编译成图片）"}
MARKERS = {"tableofcontents": "toc", "listoffigures": "lof", "listoftables": "lot",
           "frontmatter": "frontmatter", "mainmatter": "mainmatter", "backmatter": "backmatter",
           "appendix": "appendix", "newpage": "pagebreak", "clearpage": "pagebreak",
           "cleardoublepage": "pagebreak", "maketitle": "make:maketitle", "printindex": "index",
           "listofalgorithms": "loa"}

# 封面字段（已由 extract_front 读进 meta）在正文里出现时整组跳过；关键词 / 摘要命令单独成块
_FIELD_CMDS: dict[str, str] = {}
_ABSTRACT_ENVS = {"abstract", "abstract*", "IEEEabstract"}
_SETUP_CMDS: set[str] = set()
for _m in BUILTIN.values():
    if _m.kind == "command":
        _FIELD_CMDS.update(_m.keys)
    if _m.setup:
        _SETUP_CMDS.add(_m.setup)
    _ABSTRACT_ENVS |= {e for e, f in _m.envs.items() if f.startswith("abstract")}
KEYWORD_CMDS = {k for k, f in _FIELD_CMDS.items() if f.startswith("keywords")} | {"keywords", "keyword"}
ABSTRACT_CMDS = {k for k, f in _FIELD_CMDS.items() if f.startswith("abstract")}
META_CMDS = ({k for k in _FIELD_CMDS if k not in KEYWORD_CMDS | ABSTRACT_CMDS} | _SETUP_CMDS |
             {"title", "author", "date", "affiliation", "address", "email", "ead", "cortext", "fntext",
              "tnotetext", "credit", "shorttitle", "shortauthors", "nonumnote", "thesislogo", "thanks",
              "cit"})                  # 自动化学报 \\cit{引用本文格式}：前置元数据
KEYWORD_ENVS = frozenset({"keywords", "keyword", "IEEEkeywords"})
_KW_SPLIT = re.compile(r"\s*[;；,，、]\s*")
_BLOCK_CMDS = frozenset({"par", "noindent", "caption", "captionof", "bicaption", "subfloat",
                         "subcaptionbox", "bisubcaptionbox", "bibliography", "printbibliography",
                         "eqwhere", "item", "[", "label", "figurenote", "tablenote"}) | set(SECTION_LEVEL) \
    | set(MARKERS)
_NOTE_ENVS = frozenset({"tablenotes", "TableNotes"})
_PLAIN_VERB = frozenset({"verbatim", "verbatim*", "comment", "spverbatim", "Verbatim", "BVerbatim"})


def env_name(q: deque) -> str:
    return plain(read_group(q) or [])


class Parser(InlineParser, TableMixin, FloatMixin):
    def __init__(self, theorems: dict | None = None, heading_cmds: dict | None = None,
                 caption_cmds: dict | None = None, cls_envs: dict | None = None,
                 page_cmds: dict | None = None):
        super().__init__()
        self.cls_envs = cls_envs or {}              # cls_envs.env_infos：cls 定义的章级 / 列表 / 表格环境
        self.page_cmds = page_cmds or {}            # cls 生成整页的命令（\\copyrightpage…）→ 有无可选参数
        self.pending_notes: list = []               # 浮动体之前的 TableNotes（threeparttablex），挂到下一个表
        self.theorems = theorems or {}
        self.heading_cmds = heading_cmds or {}      # cls 自定义的章标题类命令 → (级别, 参数个数, 标题)
        self.caption_cmds = caption_cmds or {}      # cls 自定义的题注命令（宏体含 \caption）→ 参数个数
        self.subeq = 0
        self.subeq_seq = 0
        self.bst = ""

    def sub_blocks(self, toks) -> list:
        saved = self.last_target
        try:
            return self.blocks(toks)
        finally:
            self.last_target = saved

    # ── 块级主循环 ────────────────────────────────────────────────
    def blocks(self, toks) -> list:
        q, out, buf = deque(toks), [], InlineBuf()
        st = {"noindent": False, "cont": False}
        frame = self.enter()

        def flush() -> None:
            if buf.has_content():
                out.append(Para(buf.take(), align=self.align(), noindent=st["noindent"], cont=st["cont"]))
                st["noindent"] = False
            else:
                buf.take()
            st["cont"] = False

        while q:
            t = q.popleft()
            if t.k == "par" or t.cs("par"):
                flush()
            elif t.k == "math" and q and q[0].k == "math":
                q.popleft()
                flush()
                self.equation("$$", self.read_math(q, "$$"), out, t.line)
                st["cont"] = True
            elif t.k == "raw" and t.v[0] == "env":
                flush()
                out.append(self.code_block(t.v[1], t.v[2]))
            elif t.k == "cs" and t.v == "begin":
                name = env_name(q)
                if self.block_env(name, q, out, flush, t.line):
                    st["cont"] = name in FLOAT_ENVS or name.rstrip("*") in ("equation", "align", "gather",
                                                                            "multline", "eqnarray", "flalign")
                else:
                    self.inline_tokens(self.take_env(name, q), buf)
            elif t.k == "cs" and t.v == "end":
                env_name(q)                         # 多出来的 \end（宏展开残留），忽略
            elif t.k == "cs" and (t.v in _BLOCK_CMDS or t.v in self.heading_cmds or t.v in META_CMDS
                                  or t.v in KEYWORD_CMDS or t.v in ABSTRACT_CMDS
                                  or t.v in self.caption_cmds or t.v in self.page_cmds
                                  or (t.v in ("subfigure", "subtable") and self.float_ctx is not None)
                                  or self.is_make(t.v, q)):
                self.block_cmd(t, q, out, buf, st, flush)
            elif t.k == "cs" and (t.v in CONTENT_AFTER or t.v in STYLE_CMDS) and self.box_has_block(t.v, q):
                flush()                             # \\centerline{\\begin{tabular}…}：盒子属版式，内容按块解析
            elif not self.inline(t, q, buf):
                self.unsupported["\\" + str(t.v)] += 1
        flush()
        self.leave(frame)
        return out

    def box_has_block(self, v: str, q: deque) -> bool:
        """盒子 / 样式命令的内容里有块级环境（表格、列表、浮动体、行间公式）：吃掉参数，内容放回队首
        （外包一层分组），返回 True；否则不动 q，返回 False（照常走行内）。"""
        from .expand import MATH_ENVS
        probe = deque(q)
        read_star(probe)
        for s in (CONTENT_AFTER.get(v, "") if v in CONTENT_AFTER else ""):
            (read_opt if s == "O" else read_group)(probe)
        content = read_group(probe)
        if not content:
            return False
        names = set(re.findall(r"\\begin\s*\{([^{}]+)\}", detok(content)))
        block = set(TABULAR_ENVS) | set(FLOAT_ENVS) | set(LIST_ENVS) | set(ALIGN_ENVS) | set(MATH_ENVS)
        if not names & block:
            return False
        q.clear()
        q.extend([Tok("bg", "{")] + content + [Tok("eg", "}")] + list(probe))
        return True

    @staticmethod
    def is_make(v: str, q: deque) -> bool:
        """\\makecover \\makedeclaration 这类无参的「生成某页」命令 → Marker；\\makebox{…} 等带参的不算。"""
        if not v.startswith("make") or v in NOOP or v in CONTENT_AFTER or v in MARKERS:
            return False
        i = 0
        while i < len(q) and q[i].k == "space":
            i += 1
        return not (i < len(q) and (q[i].k == "bg" or q[i].ch("[")))

    def block_cmd(self, t: Tok, q: deque, out: list, buf: InlineBuf, st: dict, flush) -> None:
        v = t.v
        if v == "noindent":
            st["noindent"] = True
        elif v == "[":
            flush()
            self.equation("\\[", self.read_math(q, "]"), out, t.line)
            st["cont"] = True
        elif v == "label":
            self.add_label(plain(read_group(q) or []))
        elif v in SECTION_LEVEL or v in self.heading_cmds:
            flush()
            out.append(self.heading(v, q, t.line))
        elif v in self.page_cmds and v not in MARKERS:
            flush()
            read_star(q)
            if self.page_cmds[v]:
                read_opt(q)
            out.append(Marker("make:" + v))
        elif v in MARKERS or v.startswith("make"):
            flush()
            out.append(Marker(MARKERS.get(v, "make:" + v)))
        elif v in ("figurenote", "tablenote"):      # 图注 / 表注（apa7 约定的命令名）
            note = self.sub_blocks(read_group(q) or [])
            if self.float_ctx is not None:
                self.float_ctx.notes.extend(note)
            else:
                flush()
                out.extend(note)
        elif v in ("caption", "captionof", "bicaption"):
            self.caption_cmd(v, q, out)
        elif v in self.caption_cmds:
            args = [read_group(q) or [] for _ in range(self.caption_cmds[v])]
            fake = [Tok("bg", "{")] + (args[0] if args else []) + [Tok("eg", "}")]
            if len(args) > 1:
                fake += [Tok("bg", "{")] + args[1] + [Tok("eg", "}")]
            self.caption_cmd("bicaption" if len(args) > 1 else "caption", deque(fake), out)
        elif v in ("subfloat", "subcaptionbox", "bisubcaptionbox", "subfigure", "subtable") \
                and self.float_ctx is not None:
            self.subfloat_cmd(v, q)
        elif v == "bibliography":
            flush()
            files = [f.strip() for f in plain(read_group(q) or []).split(",") if f.strip()]
            out.append(Bibliography("bibtex", files=files))
        elif v == "printbibliography":
            flush()
            read_opt(q)
            out.append(Bibliography("biblatex"))
        elif v in KEYWORD_CMDS:
            flush()
            out.append(self.keywords(v, read_group(q) or []))
        elif v in ABSTRACT_CMDS:
            flush()
            out.append(self.abstract(v, read_group(q) or []))
        elif v in META_CMDS:                        # 封面字段已由 extract_front 读进 meta
            read_star(q)
            read_opt(q)
            read_group(q)
            if v == "author":
                read_opt(q)                         # Elsevier：\author[1]{姓名}[orcid=…]
        elif v == "eqwhere":
            flush()
            out.extend(self.eqwhere(q))
        elif v == "item":
            read_opt(q)
            self.unsupported["列表环境外的 \\item"] += 1
        else:
            self.inline(t, q, buf)

    def heading(self, v: str, q: deque, line: int) -> Heading:
        if v in self.heading_cmds:
            level, nargs, title = self.heading_cmds[v]      # title：参数序号（1 起）或宏体里的字面标题
            args = [read_group(q) or [] for _ in range(nargs)]
            if isinstance(title, int):
                title = args[title - 1] if 0 < title <= len(args) else []
            h = Heading(level, v, self.inlines(title), numbered=False, line=line, role=v)
            self.last_target = h                    # cls 不编号的章：其后的 \\label 指向本章（编号为空）
            return h
        else:
            star = read_star(q)
            short = read_opt(q)
            title = read_group(q) or []
            h = Heading(SECTION_LEVEL[v], v, [], numbered=not star, line=line)
            h.short = self.inlines(short) if short is not None else None
            saved = self.last_target
            self.last_target = h
            h.title = self.inlines(title)              # 标题里的 \label 归本标题
            self.last_target = saved
        self.last_target = h if h.numbered else None
        return h

    # ── 环境分派 ──────────────────────────────────────────────────
    def block_env(self, name: str, q: deque, out: list, flush, line: int) -> bool:
        from .expand import MATH_ENVS
        if name in MATH_ENVS:
            flush()
            self.parse_equation(name, q, out, line)
        elif name in FLOAT_ENVS:
            flush()
            self.parse_float(name, FLOAT_ENVS[name], q, out, line)
        elif name in SUB_ENVS and self.float_ctx is not None:
            flush()
            self.parse_subenv(name, q)
        elif name in TABULAR_ENVS:
            flush()
            if name.startswith("longtable") and self.float_ctx is None:
                fl = Float("table", name, line=line)
                saved = self.float_ctx, self.last_target
                self.float_ctx, self.last_target = fl, None
                fl.notes.extend(self.pending_notes)
                self.pending_notes = []
                fl.grids.append(self.parse_grid(name, q))
                self.float_ctx, self.last_target = saved
                fl.numbered = fl.numbered and bool(fl.caption)
                out.append(fl)
                self.last_target = fl if fl.numbered else None
            elif self.float_ctx is not None:
                self.float_ctx.grids.append(self.parse_grid(name, q))
            else:
                out.append(self.parse_grid(name, q))
        elif name in LIST_ENVS:
            flush()
            lst = self.parse_list(LIST_ENVS[name], name, q)
            if name in _NOTE_ENVS and self.float_ctx is not None:
                self.float_ctx.notes.extend(lst.items)
            elif name in _NOTE_ENVS:
                self.pending_notes.extend(lst.items)
            elif name == "highlights":
                out.append(Abstract(self.lang_of([lst]), name, [lst]))
            else:
                out.append(lst)
        elif name in self.theorems or name == "proof":
            flush()
            title = read_opt(q)
            body = self.take_env(name, q)
            if name == "proof":
                out.append(Proof(self.sub_blocks(body), self.inlines(title) if title else []))
            else:
                info = self.theorems[name]
                th = Theorem(name, info["title"], self.inlines(title) if title else [],
                             numbered=info["numbered"])
                saved = self.last_target
                self.last_target = th
                th.blocks = self.blocks(body)
                self.last_target = saved
                out.append(th)
        elif name in _ABSTRACT_ENVS:
            flush()
            read_opt(q)
            out.append(self.abstract(name, self.take_env(name, q)))
        elif name in KEYWORD_ENVS:
            flush()
            body = [Tok("char", ";") if x.cs("sep", "and") else x for x in self.take_env(name, q)]
            out.append(self.keywords(name, body))
        elif name == "thebibliography":
            flush()
            out.append(self.thebibliography(q))
        elif name in ALIGN_ENVS or name in QUOTE_ENVS:
            flush()
            self.style.append("@" + ALIGN_ENVS[name] if name in ALIGN_ENVS else "")
            inner = self.blocks(self.take_env(name, q))
            self.style.pop()
            for b in inner:
                if name in QUOTE_ENVS and isinstance(b, Para):
                    b.quote = True
            out.extend(inner)
        elif name == "subequations":
            flush()
            self.subeq_seq += 1
            saved, self.subeq = self.subeq, self.subeq_seq
            out.extend(self.blocks(self.take_env(name, q)))
            self.subeq = saved
        elif name in RAW_ENVS:
            flush()
            body = self.take_env(name, q)
            raw = Raw("\\begin{" + name + "}" + detok(body) + "\\end{" + name + "}", RAW_ENVS[name], name)
            (self.float_ctx.extra if self.float_ctx is not None else out).append(raw)
        elif name in ("lstlisting", "minted"):
            return False
        elif name in self.cls_envs:
            flush()
            self.cls_env(name, q, out, line)
        else:
            flush()
            for _ in range(TRANSPARENT_ENVS.get(name, 0)):
                read_group(q)
            if name not in TRANSPARENT_ENVS:
                self.unsupported["环境 " + name + "（按普通内容处理）"] += 1
            out.extend(self.blocks(self.take_env(name, q)))
        return True

    def code_block(self, env: str, text: str) -> Code:
        """lstlisting[language=C] / minted{python} / cls 定义的代码环境：开头的选项拆出语言名。"""
        lang = ""
        if env not in _PLAIN_VERB:
            m = re.match(r"[ \t]*(?:\[([^\]\n]*)\])?[ \t]*(?:\{([^}\n]*)\})?", text)
            if m and (m.group(1) is not None or m.group(2) is not None):
                opt = m.group(1) or ""
                lm = re.search(r"language\s*=\s*\{?([\w+#-]+)", opt)
                lang = (m.group(2) or "").strip() or (lm.group(1) if lm else "")
                text = text[m.end():]
        return Code(text.strip("\n"), env, lang)

    def cls_env(self, name: str, q: deque, out: list, line: int) -> None:
        """cls 定义的环境（判据见 cls_envs.env_infos）：章级 → 不编号的章标题（role=环境名，标题取 cls 常量，
        取不到留空交给目标模板）；内层列表 / 表格按对应结构解析；其余内容按普通块解析。"""
        e = self.cls_envs[name]
        if e.has_opt:
            read_opt(q)
        for _ in range(e.nreq):
            read_group(q)
        if e.chapter:
            h = Heading(SECTION_LEVEL["chapter"], name, self.inlines(lex(e.title)) if e.title else [],
                        numbered=False, line=line, role=name)
            out.append(h)
            self.last_target = h
        body = self.take_env(name, q)
        if e.inner in ("itemize", "enumerate", "description"):
            out.append(self.list_from(e.inner, body))
        elif e.inner == "table":
            out.append(self.grid_from(name, e.spec, body))
        else:
            out.extend(self.blocks(body))

    def parse_list(self, style: str, env: str, q: deque) -> ListBlock:
        read_opt(q)
        return self.list_from(style, self.take_env(env, q))

    def list_from(self, style: str, body: list) -> ListBlock:
        items, cur, label, depth = [], None, None, 0
        lst = ListBlock(style, [])

        def close() -> None:
            if cur is not None:
                it = ListItem([], self.inlines(label) if label is not None else None)
                saved = self.last_target
                self.last_target = it if style == "enumerate" else None
                it.blocks = self.blocks(cur)
                self.last_target = saved
                lst.items.append(it)

        bq = deque(body)
        while bq:
            t = bq.popleft()
            if t.cs("begin"):
                depth += 1
            elif t.cs("end"):
                depth -= 1
            if depth == 0 and t.cs("item"):
                close()
                label, cur = read_opt(bq), []
                continue
            if cur is not None:
                cur.append(t)
        close()
        return lst

    # ── 前置 / 后置 ───────────────────────────────────────────────
    def lang_of(self, blocks) -> str:
        from .ir import walk
        text = "".join(inline_text(getattr(b, "inlines", None)) for b in walk(blocks))
        return "zh" if is_cjk_text(text) else "en"

    def abstract(self, env: str, body: list) -> Abstract:
        blocks = self.sub_blocks(body)
        return Abstract(self.lang_of(blocks), env, blocks)

    def keywords(self, cmd: str, body: list) -> Keywords:
        text = inline_text(self.inlines(body))
        items = [x for x in _KW_SPLIT.split(text) if x]
        return Keywords("zh" if is_cjk_text(text) else "en", cmd, items)

    def eqwhere(self, q: deque) -> list:
        """xautthesis.cls:343 的「式中：」释义块：\\eqwhere[列宽]{{符号}{释义},…}。
        版式（顶格、破折号对齐、中间行「；」末行「。」）照 cls 注释取，不在这里排。"""
        read_opt(q)
        groups, cur, depth = [], [], 0
        for t in read_group(q) or []:
            if t.k == "bg":
                depth += 1
            elif t.k == "eg":
                depth -= 1
            if depth == 0 and t.ch(","):
                groups.append(cur)
                cur = []
                continue
            cur.append(t)
        groups.append(cur)
        out = [Para(self.inlines([Tok("char", c) for c in "式中："]), noindent=True)]
        for k, g in enumerate(groups):
            gq = deque(g)
            sym, exp = read_group(gq) or [], read_group(gq) or []
            end = "。" if k == len(groups) - 1 else "；"
            out.append(Para(self.inlines(sym + [Tok("char", "——")] + exp + [Tok("char", end)]), noindent=True))
        return out

    def thebibliography(self, q: deque) -> Bibliography:
        read_group(q)
        bib, cur, head = Bibliography("thebibliography"), None, None
        body = deque(self.take_env("thebibliography", q))

        def close() -> None:
            if head is not None:
                bib.items.append(BibItem(head[0], head[1], self.inlines(cur)))

        while body:
            t = body.popleft()
            if t.cs("bibitem"):
                close()
                lab = read_opt(body)
                head, cur = (plain(read_group(body) or []), inline_text(self.inlines(lab)) if lab else ""), []
                continue
            if t.cs("newblock"):
                t = Tok("space", " ")
            if cur is not None:
                cur.append(t)
        close()
        return bib
