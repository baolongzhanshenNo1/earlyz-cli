"""L2 编号：按 TeX 计数器规则给章节/图/表/公式/定理/列表项推算编号，并建标签表。

规则（LaTeX 内核 + amsmath，不按学校猜）：
- book/report 类：章 = 1 级，图/表/公式按章编号；article 类：节 = 顶级，图/表/公式全文连续。
  cls 里 \\counterwithin{figure}{chapter}（或 \\numberwithin）会改这个关系；分隔符取 cls 的
  \\renewcommand{\\thefigure}{\\thechapter-\\arabic{figure}} 实际写法（xaut 是 "-"，标准是 "."）。
- \\appendix 之后章号变 A B C；\\frontmatter / \\backmatter 里章不编号；* 形式不编号。
- 公式：equation/multline 一个号；align/gather/eqnarray 每行一个号（\\nonumber/\\notag 去掉，
  \\tag 自定编号且不占计数器）；subequations 组内 = 父号 + a b c。
产出可与 TeX 编译的 .aux（\\newlabel）逐条对账：tests 用 TeX Live 裁判。
"""
from __future__ import annotations

import re

from .ir import Equation, Float, Heading, ListBlock, Marker, Theorem, walk

_ALPHA = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def _alph(n: int) -> str:
    return _ALPHA[n - 1] if 0 < n <= 26 else str(n)


def _roman(n: int) -> str:
    vals = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"), (50, "l"),
            (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
    out = ""
    for v, s in vals:
        while n >= v:
            out += s
            n -= v
    return out


class Scheme:
    """计数器从属关系与分隔符。from_cls 读 cls 源码里的 counterwithin / \\the<ctr> 定义。"""

    def __init__(self, top: str):
        self.top = top                                   # chapter | section
        self.within = {"figure": top, "table": top, "equation": top} if top == "chapter" else \
            {"figure": None, "table": None, "equation": None}
        self.within["algorithm"] = None                  # algorithm.sty 默认全文连续；[chapter] 选项才按章
        self.sep = {"figure": ".", "table": ".", "equation": ".", "algorithm": "."}
        # 分图引用 = \\thefigure + \\thesubfigure；subcaption / subfig 默认 \\alph（1.1a），cls 改写成 (\\alph) 时为 1.1(a)
        self.sub = {"figure": ("", ""), "table": ("", "")}

    def from_cls(self, src: str) -> "Scheme":
        ctr = r"\{\s*(figure|table|equation|algorithm|algocf)\s*\}"
        par = r"\s*\{\s*(\w+)\s*\}"
        # \counterwithin / \numberwithin / cls 包装（\SJTU@counterwithin { figure } { chapter }）/ \@addtoreset
        for m in re.finditer(r"(?:\\|@)(?:counterwithin|numberwithin)\*?\s*(?:\[[^\]]*\]\s*)?" + ctr + par, src):
            self.within[_ctr(m.group(1))] = m.group(2)
        for m in re.finditer(r"\\@addtoreset\s*" + ctr + par, src):
            self.within[_ctr(m.group(1))] = m.group(2)
        for m in re.finditer(r"\\counterwithout\*?\s*" + ctr + par, src):
            self.within[_ctr(m.group(1))] = None
        # algorithm.sty 的 [chapter]/[section] 包选项；algorithm2e 的 algochapter/algosection
        for m in re.finditer(r"\\PassOptionsToPackage\s*\{\s*(chapter|section)\s*\}\s*\{\s*algorithm\s*\}", src):
            self.within["algorithm"] = m.group(1)
        for m in re.finditer(r"\\usepackage\s*\[([^\]]*)\]\s*\{\s*(algorithm2e|algorithm)\s*\}", src):
            opts = {o.strip() for o in m.group(1).split(",")}
            for lvl in ("chapter", "section"):
                if lvl in opts or "algo" + lvl in opts:
                    self.within["algorithm"] = lvl
        # \the<ctr> 的实际写法定分隔符：LaTeX2e \renewcommand{\thefigure}{\thechapter-\arabic{figure}}，
        # expl3 \cs_set:Npn \thefigure { \thechapter - \__fdu_arabic:n { figure } }（空格在 expl3 里不算）
        for c in ("figure", "table", "equation", "algorithm"):
            m = re.search(r"\\(?:renewcommand\*?|def|cs_set:Npn|cs_gset:Npn|cs_new:Npn|cs_set_nopar:Npn)\s*\{?\\the"
                          + c + r"\}?\s*\{\s*\\the(\w+)([^\\]{0,16}?)\\[\w@:]*arabic(?::n)?\s*\{\s*" + c
                          + r"\s*\}\s*\}", src)
            if m:
                self.within[c] = m.group(1)
                self.sep[c] = re.sub(r"\s+|[{}]", "", m.group(2))
        for m in re.finditer(r"\\(?:renewcommand\*?|def|cs_set:Npn|cs_gset:Npn)\s*\{?\\thesub(figure|table)\}?\s*"
                             r"\{([^{}\\]*)\\alph\s*\{\s*sub(?:figure|table)\s*\}([^{}\\]*)\}", src):
            self.sub[m.group(1)] = (m.group(2).strip(), m.group(3).strip())
        return self


def _ctr(name: str) -> str:
    return "algorithm" if name == "algocf" else name


def assign_numbers(doc, scheme: Scheme) -> None:
    c = {"part": 0, "chapter": 0, "section": 0, "subsection": 0, "subsubsection": 0,
         "figure": 0, "table": 0, "equation": 0, "algorithm": 0}
    thm: dict = {}
    appendix = False
    numbered_matter = True
    sub_parent = {}                                       # subequations 组号 → (父号, 已用个数)
    order = ["part", "chapter", "section", "subsection", "subsubsection", "paragraph", "subparagraph"]
    labels = doc.labels

    def chap_no() -> str:
        return _alph(c["chapter"]) if appendix else str(c["chapter"])

    def head_no(name: str) -> str:
        if scheme.top == "section":
            parts = [(_alph(c["section"]) if appendix else str(c["section"]))]
            for n in order[order.index("section") + 1: order.index(name) + 1]:
                parts.append(str(c[n]))
            return ".".join(parts[: order.index(name) - order.index("section") + 1])
        parts = [chap_no()]
        for n in order[order.index("chapter") + 1: order.index(name) + 1]:
            parts.append(str(c[n]))
        return ".".join(parts[: order.index(name) - order.index("chapter") + 1])

    def obj_no(ctr: str) -> str:
        par = scheme.within.get(ctr)
        n = str(c[ctr])
        if par == "chapter" and scheme.top == "chapter":
            return chap_no() + scheme.sep[ctr] + n
        if par == "section":
            return head_no("section") + scheme.sep[ctr] + n
        return n

    def reset(name: str) -> None:
        idx = order.index(name)
        for n in order[idx + 1:]:
            if n in c:
                c[n] = 0
        for ctr, par in scheme.within.items():
            if par == name:
                c[ctr] = 0
        for k in thm:
            if thm[k]["within"] == name:
                thm[k]["n"] = 0

    def put(keys, kind: str, number: str) -> None:
        for k in keys:
            labels.setdefault(k, {"kind": kind, "number": number})

    for b in walk(doc.blocks):
        if isinstance(b, Marker):
            if b.name == "appendix":
                appendix = True
                c[scheme.top] = 0
            elif b.name in ("frontmatter", "backmatter"):
                numbered_matter = False
            elif b.name == "mainmatter":
                numbered_matter = True
        elif isinstance(b, Heading):
            name = b.name if b.name in order else None
            if name is None:                        # cls 的章级环境 / 命令（致谢、符号说明…）：不编号
                b.number = ""
                put(b.labels, "heading", "")
                continue
            if not b.numbered or (name == "chapter" and not numbered_matter) or \
                    order.index(name) < order.index(scheme.top) or name in ("paragraph", "subparagraph"):
                b.number = ""
                put(b.labels, "heading", "")
                continue
            c[name] = c.get(name, 0) + 1
            reset(name)
            b.number = head_no(name)
            put(b.labels, "heading", b.number)
        elif isinstance(b, Float) and b.type in ("figure", "table", "algorithm") and b.env != "sub":
            if b.numbered:
                c[b.type] += 1
                b.number = obj_no(b.type)
            put(b.labels, b.type, b.number)
            pre, suf = scheme.sub.get(b.type, ("", ""))
            j = 0
            for sf in b.subfloats:
                if sf.type == "sub":
                    letter = chr(ord("a") + j)
                    j += 1
                    sf.number = "(" + letter + ")"             # 分图题注前的标号（subcaption 默认 parens）
                    put(sf.labels, "sub" + b.type, b.number + pre + letter + suf)
        elif isinstance(b, Equation):
            b.number = []
            tags = {r: t for r, t in b.tags}
            no_num = set(b.nonumber_rows)
            multi = b.env.rstrip("*") not in ("equation", "multline")
            for r in range(b.rows):
                if r in tags:
                    b.number.append(tags[r])
                elif b.numbers and (r not in no_num if multi else r == b.rows - 1):
                    if b.group:
                        if b.group not in sub_parent:
                            c["equation"] += 1
                            sub_parent[b.group] = [obj_no("equation"), 0]
                        sub_parent[b.group][1] += 1
                        b.number.append(sub_parent[b.group][0] + chr(ord("a") + sub_parent[b.group][1] - 1))
                    else:
                        c["equation"] += 1
                        b.number.append(obj_no("equation"))
                else:
                    b.number.append("")
            if not multi and b.number:              # equation：编号挂最后一行（单行即第 0 行）
                b.number = [""] * (b.rows - 1) + [next((x for x in b.number if x), "")]
            for k, row in zip(b.labels, b.label_rows or [b.rows - 1] * len(b.labels)):
                num = b.number[row] if multi and row < len(b.number) else next((x for x in b.number if x), "")
                put([k], "equation", num)
        elif isinstance(b, Theorem):
            info = doc_theorem(doc, b.env)
            if b.numbered:
                key = info.get("shared") or b.env
                within = info.get("within", "")
                st = thm.setdefault(key, {"n": 0, "within": scheme.top if within == "@top" else within})
                st["n"] += 1
                par = st["within"]
                b.number = (head_no(par) + "." if par in order and c.get(par) else "") + str(st["n"])
            put(b.labels, "theorem", b.number)
        elif isinstance(b, ListBlock) and b.style == "enumerate":
            for k, it in enumerate(b.items, 1):
                it.number = str(k)
                put(it.labels, "item", it.number)


def doc_theorem(doc, env: str) -> dict:
    return doc.stats.get("theorems", {}).get(env, {})


def roman(n: int) -> str:
    return _roman(n)
