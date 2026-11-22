"""cls/sty 源码的字符串级扫描：「出章标题的环境」「列表 / 表格环境」「生成整页的命令」「expl3 批量定义的定理」。

为什么在字符串层做：这些定义大多在 expl3 区段（\\ExplSyntaxOn…Off，_ : 是字母），load_lib 不收；
L2 只需要知道「这个环境出一个不编号的章标题 / 这是列表 / 这是整页」，不必真展开 cls。
判据只看 cls 自己的代码（调用了哪个章标题原语、开了哪种列表），不按学校名猜；章标题文字取 cls 里的
常量定义，取不到就留空（Heading.role 记环境名，标题交给目标模板）。
"""
from __future__ import annotations

import re
from dataclasses import dataclass

_STD_LISTS = {"itemize": "itemize", "enumerate": "enumerate", "description": "description"}
# 章标题原语：\chapter、cls 包装（\thu@chapter、\__ustc_chapter:w、\__fdu_chapter:V）、SJTU 的 make_heading、
# \addcontentsline{toc}{chapter}。\c@chapter（计数器）这类不算。
_CHAPTER_RE = re.compile(r"\\((?:[A-Za-z]+@|__[A-Za-z]+_)?chapter)(?::[A-Za-z]+)?(?![A-Za-z@_])"
                         r"|\\[A-Za-z@_]*make_heading:[A-Za-z]*"
                         r"|\\addcontentsline\s*\{\s*toc\s*\}\s*\{\s*chapter\s*\}")
_NOT_CHAPTER = ("c@chapter", "p@chapter", "if@chapter", "cl@chapter")
_PAGE_RE = re.compile(r"\\begin\s*\{\s*titlepage\s*\}|\\includepdf\b|\\cleardoublepage\b|\\clearpage\b")
_STRONG_PAGE_RE = re.compile(r"\\begin\s*\{\s*titlepage\s*\}|\\includepdf\b")
_FUNC_RES = (
    re.compile(r"\\cs_(?:new|set|gset)(?:_protected)?(?:_nopar)?:[Nc]p?[nNVvxeo]*\s*(\\[A-Za-z@_:]+)[^{]*"),
    re.compile(r"\\[gex]?def\s*(\\[A-Za-z@]+)[^{]*"),
    re.compile(r"\\(?:new|renew|provide)command\*?\s*\{?\s*(\\[A-Za-z@]+)\s*\}?\s*(?:\[\d\])?\s*(?:\[[^\]]*\])?\s*"),
)
_DOCCMD_RE = re.compile(r"\\(?:New|Renew|Provide|Declare)DocumentCommand\s*\{?\s*(\\[A-Za-z@]+)\s*\}?\s*")
_ENV_RE = re.compile(r"\\(NewDocumentEnvironment|RenewDocumentEnvironment|DeclareDocumentEnvironment"
                     r"|newenvironment|renewenvironment|NewEnviron|RenewEnviron)\s*\*?\s*\{\s*([\w@*]+)\s*\}")
_VERB_RE = re.compile(r"\\(?:lstnewenvironment|DefineVerbatimEnvironment|newtcblisting|NewListingEnvironment)"
                      r"\s*\{\s*([A-Za-z*]+)\s*\}")


@dataclass
class EnvInfo:
    chapter: bool = False
    title: str = ""               # cls 常量里取到的标题原文（LaTeX 源，交给词法器）；取不到为空
    inner: str = ""               # "" | itemize | enumerate | description | table
    spec: str = ""                # inner == table 时的列格式（xparse 可选参数默认值）
    has_opt: bool = False
    nreq: int = 0


def group_at(s: str, i: int) -> tuple[str, int] | None:
    """s[i:] 跳过空白后应是 `{`：返回 (组内文本, 组后位置)。"""
    n = len(s)
    while i < n and s[i] in " \t\r\n":
        i += 1
    if i >= n or s[i] != "{":
        return None
    depth, j = 0, i
    while j < n:
        c = s[j]
        if c == "\\":
            j += 2
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return s[i + 1:j], j + 1
        j += 1
    return None


def _strip_groups(s: str) -> str:
    prev = None
    while prev != s:
        prev, s = s, re.sub(r"\{[^{}]*\}", "", s)
    return s


def _funcs(src: str) -> dict:
    out: dict = {}
    for rx in _FUNC_RES + (_DOCCMD_RE,):
        for m in rx.finditer(src):
            i = m.end()
            if rx is _DOCCMD_RE:                       # 先跳过参数规格组
                g = group_at(src, i)
                if g is None:
                    continue
                i = g[1]
            g = group_at(src, i)
            if g is not None:
                out.setdefault(m.group(1), g[0])
    return out


def _expand(text: str, funcs: dict, depth: int = 0, seen: set | None = None) -> str:
    """把正文里调到的 cls 内部函数体接在后面（最多 3 层），供判据搜索。"""
    seen = seen if seen is not None else set()
    if depth > 3:
        return text
    extra = []
    for name in dict.fromkeys(re.findall(r"\\[A-Za-z@_:]+", text)):
        if name in funcs and name not in seen:
            seen.add(name)
            extra.append(_expand(funcs[name], funcs, depth + 1, seen))
    return text + ("\n" + "\n".join(extra) if extra else "")


def resolve_text(expr: str, src: str, depth: int = 0) -> str:
    """章标题实参 → 文字：字面量原样返回；\\tl_use:N \\x / \\x 查 cls 里第一处常量定义。取不到返回空。"""
    expr = expr.strip()
    if not expr or "#" in expr or depth > 3:
        return ""
    m = re.fullmatch(r"(?:\\tl_use:N\s*)?(\\[A-Za-z@_:]+)", expr)
    if not m:
        return "" if re.search(r"\\[A-Za-z@_]*[@_:]", expr) else expr
    name = re.escape(m.group(1))
    for rx in (r"\\[gex]?def\s*" + name + r"\s*", r"\\(?:re)?newcommand\*?\s*\{?\s*" + name + r"\s*\}?\s*",
               r"\\tl_(?:g?set|const):[Nc]n\s*\{?\s*" + name + r"\s*\}?\s*"):
        d = re.search(rx + r"(?=\{)", src)
        if d:
            g = group_at(src, d.end())
            if g:
                return resolve_text(g[0], src, depth + 1)
    return ""


def _chapter_title(body: str, m: re.Match, src: str) -> str:
    if "make_heading" in m.group(0):
        t = re.search(r"title\s*=\s*([^,}]+)", body[m.end():m.end() + 200])
        return resolve_text(t.group(1), src) if t else ""
    i = m.end()
    tail = body[i:]
    tail = re.sub(r"^\s*\*?\s*(?:\[[^\]]*\])?", "", tail)
    g = group_at(tail, 0)
    if g is not None:
        return resolve_text(g[0], src)
    cs = re.match(r"\s*(\\[A-Za-z@_:]+)", tail)                  # :V 变体直接跟变量
    return resolve_text(cs.group(1), src) if cs else ""


def newlists(src: str) -> dict:
    return {m.group(1): m.group(2) for m in re.finditer(
        r"\\newlist\s*\{\s*([\w@]+)\s*\}\s*\{\s*(itemize|enumerate|description)\s*\}", src)}


def verb_envs(src: str) -> set:
    return set(_VERB_RE.findall(src))


def _env_defs(src: str) -> dict:
    defs: dict = {}
    for m in _ENV_RE.finditer(src):
        kind, name, i = m.group(1), m.group(2), m.end()
        if "Document" in kind:
            g = group_at(src, i)
            if g is None:
                continue
            spec, i = g
            plain = _strip_groups(spec)
            has_opt = bool(re.search(r"(?<![A-Za-z])[Oo](?![A-Za-z])", plain))
            nreq = len(re.findall(r"(?<![A-Za-z+])m(?![A-Za-z])", plain))
            d = re.search(r"O\s*\{", spec)
            default = group_at(spec, d.end() - 1)[0] if d and group_at(spec, d.end() - 1) else ""
        else:
            a = re.match(r"\s*(?:\[(\d)\])?\s*(?:\[([^\]]*)\])?", src[i:])
            n = int(a.group(1)) if a.group(1) else 0
            has_opt, default = a.group(2) is not None, a.group(2) or ""
            nreq, i = max(0, n - (1 if has_opt else 0)), i + a.end()
        g = group_at(src, i)
        if g is not None:
            defs.setdefault(name, (g[0], has_opt, nreq, default))
    return defs


def env_infos(src: str, list_envs: dict) -> dict:
    """cls 里定义、且能认出语义（章标题 / 列表 / 表格）的环境 → EnvInfo。认不出的不收（照旧透明解析）。"""
    funcs, defs, out = _funcs(src), _env_defs(src), {}
    lists = dict(_STD_LISTS) | dict(list_envs) | newlists(src)

    def info(name: str, depth: int = 0) -> EnvInfo | None:
        if name not in defs or depth > 3:
            return None
        body, has_opt, nreq, default = defs[name]
        text = _expand(body, funcs)
        e = EnvInfo(has_opt=has_opt, nreq=nreq)
        for m in _CHAPTER_RE.finditer(text):
            if m.group(1) and m.group(1) in _NOT_CHAPTER:
                continue
            e.chapter, e.title = True, _chapter_title(text, m, src)
            break
        for b in re.finditer(r"\\begin\s*\{\s*([\w@*]+)\s*\}", text):
            inner = b.group(1)
            if inner in lists:
                e.inner = lists[inner]
                break
            if inner in ("longtable", "tabular", "longtable*") and b.start() < len(body):
                e.inner = "table"
                break
            sub = info(inner, depth + 1) if inner != name else None
            if sub is not None:
                e.chapter, e.title = e.chapter or sub.chapter, e.title or sub.title
                e.inner, e.spec = e.inner or sub.inner, e.spec or sub.spec
                break
        if not e.inner:
            if re.search(r"\\(?:longtable|tabular)(?![A-Za-z@_])", body):     # 只看环境体本身
                e.inner, e.spec = "table", default.strip()
            elif re.search(r"\\list(?![A-Za-z@_])", text):                   # trivlist 是 proof 一类，不是列表
                e.inner = ("enumerate" if re.search(r"\\(?:arabic|usecounter)\b", text)
                           else "description" if "makelabel" in text else "itemize")
        return e if (e.chapter or e.inner) else None

    for name in defs:
        e = info(name)
        if e is not None:
            out[name] = e
    for name, style in newlists(src).items():          # enumitem \\newlist：列表环境，可带 [选项]
        out.setdefault(name, EnvInfo(inner=style, has_opt=True))
    return out


def page_cmds(src: str, exclude: set, strong_only: bool = False) -> dict:
    """无必选参数、宏体（含调用链）生成整页的 cls 命令（\\copyrightpage、\\declarationoforiginality…）→ 有无可选参数。
    strong_only：只收 titlepage / \\includepdf 这类确定整页的（\\clearpage 换页的章标题命令不算）。"""
    funcs, out = _funcs(src), {}
    page_re = _STRONG_PAGE_RE if strong_only else _PAGE_RE
    for m in _DOCCMD_RE.finditer(src):
        name = m.group(1)[1:]
        g = group_at(src, m.end())
        if not g or "@" in name or name in exclude:
            continue
        plain = _strip_groups(g[0])
        if re.search(r"(?<![A-Za-z+])m(?![A-Za-z])", plain):
            continue
        body = group_at(src, g[1])
        if body and page_re.search(_expand(body[0], funcs)):
            out[name] = bool(re.search(r"(?<![A-Za-z])[Oo](?![A-Za-z])", plain))
    for name, body in funcs.items():
        n = name[1:]
        if n in out or "@" in n or "_" in n or ":" in n or n in exclude:
            continue
        hdr = re.search(r"\\(?:new|renew|provide)command\*?\s*\{?\s*" + re.escape(name) + r"\s*\}?\s*(\[\d\])?\s*(\[[^\]]*\])?", src)
        if not hdr:
            continue
        nargs = int(hdr.group(1)[1]) if hdr.group(1) else 0
        has_opt = hdr.group(2) is not None
        if nargs - (1 if has_opt else 0) == 0 and page_re.search(_expand(body, funcs)):
            out[n] = has_opt
    return out


def expl_theorems(src: str) -> dict:
    """\\clist_map_inline:nn {theorem, lemma, …} { … \\newtheorem {##1} {…} [chapter] }（SJTU）→ 定理表。"""
    out: dict = {}
    for m in re.finditer(r"\\clist_map_inline:nn\s*", src):
        g1 = group_at(src, m.end())
        g2 = group_at(src, g1[1]) if g1 else None
        if not g2 or "\\newtheorem" not in g2[0]:
            continue
        star = bool(re.search(r"\\newtheorem\s*\*", g2[0]))
        w = re.search(r"\]\s*$|\[\s*(chapter|section)\s*\]", g2[0])
        within = w.group(1) if w and w.group(1) else ""
        for name in (x.strip() for x in g1[0].split(",")):
            if re.fullmatch(r"[A-Za-z]+", name or ""):
                out.setdefault(name, {"title": "", "numbered": not star, "shared": "", "within": within})
    return out
