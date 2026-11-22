"""L1+L2 入口：LaTeX 工程主文件 → Document IR。

流程：词法（逐文件，\\input/\\include/\\subfile/\\import 在 token 层拼接，注释由词法器按 TeX 规则去掉）
→ 切导言区/正文 → 加载 cls 与工程内 .sty 的公共宏 → 展开导言区（收定义）→ 展开正文 →
从 cls 内部宏里认出「章标题类命令」「题注类命令」→ 块级解析 → 编号推算 → 统计。
"""
from __future__ import annotations

import re
from collections import Counter, deque
from pathlib import Path

from .cls_envs import env_infos, expl_theorems, page_cmds, verb_envs
from .expand import Expander
from .front_fields import extract_front
from .ir import Bibliography, Document, Equation, Float, Heading, Inline, Raw, walk
from .macros import Macro, _skip_sp, detok, plain, read_group, read_opt
from .numbering import Scheme, assign_numbers
from .parse_blocks import (_ABSTRACT_ENVS, ABSTRACT_CMDS, FLOAT_ENVS, KEYWORD_CMDS, KEYWORD_ENVS,
                           MARKERS, META_CMDS, Parser)
from .parse_inline import peek_env
from .parse_text import CITE_CMDS, REF_CMDS, SECTION_LEVEL, STYLE_CMDS
from .project import read_text, strip_comments
from .siunitx import range_phrase_from
from .tokens import VERB_ENVS, Tok, lex

_INCLUDE = frozenset({"input", "include", "subfile", "InputIfFileExists", "import", "subimport",
                      "subinputfrom", "inputfrom"})
_COUNTER_FMT = frozenset({"arabic", "roman", "Roman", "alph", "Alph", "chinese", "zhnum", "fnsymbol"})
PROTECTED = (set(SECTION_LEVEL) | set(MARKERS) | set(CITE_CMDS) | set(REF_CMDS) | set(STYLE_CMDS)
             | META_CMDS | KEYWORD_CMDS | ABSTRACT_CMDS | _ABSTRACT_ENVS | KEYWORD_ENVS | set(FLOAT_ENVS)
             | {"caption", "captionof", "bicaption", "label", "footnote", "footnotetext", "includegraphics",
                "item", "begin", "end", "bibitem", "bibliography", "printbibliography", "par", "\\", "url",
                "href", "eqwhere", "tag", "nonumber", "notag", "subfloat", "subcaptionbox",
                "bisubcaptionbox", "thebibliography", "itemize", "enumerate", "description", "tabular",
                "equation", "align", "gather", "multline", "proof", "verb", "newpage", "clearpage"})
MAX_DEPTH = 12


class _Ctx:
    def __init__(self, root: Path, main_dir: Path):
        self.root, self.main_dir = root, main_dir
        self.seen: set[Path] = set()
        self.missing: list[str] = []
        self.files: list[str] = []
        self.encodings: Counter = Counter()
        self.verb: set = set(VERB_ENVS)          # 代码环境名（cls / setup.tex 里定义的也算），跨文件共享


def _resolve(ctx: _Ctx, cur_dir: Path, name: str) -> Path | None:
    name = name.strip().strip('"')
    if not name:
        return None
    for d in (ctx.main_dir, cur_dir, ctx.root):
        for cand in (name, name + ".tex"):
            p = (d / cand)
            if p.is_file():
                try:
                    p.resolve().relative_to(ctx.root.resolve())       # 不读工程外的文件
                except ValueError:
                    continue
                return p
    return None


def _file_name(q: deque) -> str:
    _skip_sp(q)
    if q and q[0].k == "bg":
        return plain(read_group(q) or [])
    out = []
    while q and q[0].k == "char":
        out.append(q.popleft().v)
    return "".join(out)


def lex_project(path: Path, ctx: _Ctx, depth: int = 0) -> list[Tok]:
    text, enc = read_text(path)
    ctx.encodings[enc] += 1
    rel = path.resolve().relative_to(ctx.root.resolve()).as_posix()
    ctx.files.append(rel)
    q, out = deque(lex(text, rel, verb_envs=ctx.verb)), []
    while q:
        t = q.popleft()
        if t.k != "cs" or t.v not in _INCLUDE:
            out.append(t)
            continue
        if t.v in ("import", "subimport", "inputfrom", "subinputfrom"):
            sub = _file_name(q)
            name = sub.rstrip("/") + "/" + _file_name(q)
            base = path.parent if t.v.startswith("sub") else ctx.main_dir
        else:
            name, base = _file_name(q), path.parent
        if t.v == "InputIfFileExists":
            read_group(q)
            read_group(q)
        target = _resolve(ctx, base, name)
        if target is None:
            ctx.missing.append(name)
            continue
        rt = target.resolve()
        if rt in ctx.seen or depth >= MAX_DEPTH:
            continue
        ctx.seen.add(rt)
        if t.v == "include":
            out.append(Tok("cs", "clearpage", t.line, t.file))
        out.extend(lex_project(target, ctx, depth + 1))
    return out


def _split_document(toks: list[Tok]) -> tuple[list, list]:
    q = deque(toks)
    pre: list = []
    while q:
        t = q.popleft()
        if t.cs("begin"):
            name, n = peek_env(q)
            if name == "document":
                for _ in range(n):
                    q.popleft()
                break
        pre.append(t)
    body: list = []
    while q:
        t = q.popleft()
        if t.cs("end"):
            name, n = peek_env(q)
            if name == "document":
                break
        body.append(t)
    return pre, body


def _docclass(pre: list) -> tuple[str, list]:
    q = deque(pre)
    while q:
        t = q.popleft()
        if t.cs("documentclass", "documentstyle"):
            opts = read_opt(q)
            name = plain(read_group(q) or [])
            return name, [o.strip() for o in plain(opts or []).split(",") if o.strip()]
    return "", []


def _heading_info(name: str, raw: dict, depth: int = 0):
    """cls 内部宏是否「生成一个章/节标题」：宏体里出现 \\chapter*{T} / \\section*{T} 或
    \\addcontentsline{toc}{chapter}{T}，或调用了另一个这样的宏。返回 (级别, 标题 token)。"""
    m = raw.get(name)
    if m is None or depth > 3:
        return None
    q = deque(m.body)
    while q:
        t = q.popleft()
        if t.k != "cs":
            continue
        if t.v in ("chapter", "section"):
            if q and q[0].ch("*"):
                q.popleft()
            read_opt(q)
            title = read_group(q)
            if title is not None:
                return SECTION_LEVEL[t.v], title
        elif t.v == "addcontentsline":
            read_group(q)
            lvl = plain(read_group(q) or [])
            title = read_group(q)
            if lvl in ("chapter", "section") and title is not None:
                return SECTION_LEVEL[lvl], title
        elif t.v in raw and t.v != name:
            info = _heading_info(t.v, raw, depth + 1)
            if info:
                callee = raw[t.v]
                args = [read_group(q) or [] for _ in range(callee.nargs)]
                from .macros import _subst
                return info[0], _subst(info[1], args)
    return None


def _cls_commands(exp: Expander) -> tuple[dict, dict]:
    heading_cmds, caption_cmds = {}, {}
    for name, m in exp.raw.items():
        if "@" in name or name in exp.protected:
            continue
        info = _heading_info(name, exp.raw)
        if info:
            level, title = info
            ps = [x for x in title if x.k not in ("space",)]
            if len(ps) == 2 and ps[0].k == "param" and ps[1].k == "char" and ps[1].v.isdigit():
                heading_cmds[name] = (level, m.nargs, int(ps[1].v))
            else:
                heading_cmds[name] = (level, m.nargs, title)
        elif any(t.cs("caption") for t in m.body):
            caption_cmds[name] = m.nargs
    return heading_cmds, caption_cmds


def _keep_page_makers(exp: Expander) -> None:
    """cls 的 \\makecover \\makedeclaration 这类无参「生成某页」宏不展开：封面/声明页由目标模板出，
    展开会把 cls 写死的声明文字当正文。留成 Marker（parse_blocks.is_make）。"""
    for name in [n for n, m in exp.t.macros.items() if n.startswith("make") and m.nargs == 0 and m.lib]:
        exp.raw[name] = exp.t.macros.pop(name)
        exp.protected.add(name)


def _keep_counter_macros(exp: Expander) -> None:
    """\\the<计数器> 的重定义不展开（保留计数器语义，分隔符由 numbering.Scheme 读 cls 源码取）。"""
    for name in [n for n in exp.t.macros if n.startswith("the") and len(n) > 3]:
        m = exp.t.macros[name]
        if any(t.k == "cs" and (t.v in _COUNTER_FMT or t.v.startswith("the")) for t in m.body):
            exp.raw[name] = exp.t.macros.pop(name)


def parse_project(root: str | Path, main_rel: str, cls_path: str = "") -> Document:
    """root = 工程根；main_rel = 主文件相对路径；cls_path = 工程内相对路径或本机 TeX 绝对路径（可空）。"""
    root = Path(root)
    main = root / main_rel
    ctx = _Ctx(root, main.parent)
    ctx.seen.add(main.resolve())
    libs: list[tuple[str, str]] = []                 # (源码, 文件名)：cls + 工程内 .sty
    cls_src = ""
    if cls_path:
        p = Path(cls_path) if Path(cls_path).is_absolute() else root / cls_path
        if p.is_file():
            cls_src, _ = read_text(p)
            libs.append((cls_src, p.name))
    for m in re.finditer(r"\\(?:usepackage|RequirePackage)\s*(?:\[[^\]]*\])?\s*\{([^}]+)\}",
                         strip_comments(read_text(main)[0])):
        for pkg in m.group(1).split(","):
            sty = _resolve(ctx, main.parent, pkg.strip() + ".sty")
            if sty is not None:
                libs.append((read_text(sty)[0], sty.name))
    lib_src = "\n".join(s for s, _ in libs)
    ctx.verb |= verb_envs(lib_src)
    toks = lex_project(main, ctx)
    pre, body = _split_document(toks)
    cls, opts = _docclass(pre)

    exp = Expander(protected=PROTECTED)
    for src, name in libs:
        exp.load_lib(src, name)
    envs = env_infos(lib_src, {})
    exp.protected |= set(envs)                       # cls 环境留给 L2 按语义解析，不展开
    exp.expand(pre)
    _keep_page_makers(exp)
    _keep_counter_macros(exp)
    for name, info in expl_theorems(lib_src).items():
        exp.t.theorems.setdefault(name, info)
    heading_cmds, caption_cmds = _cls_commands(exp)
    excl = set(SECTION_LEVEL) | set(MARKERS) | set(envs)
    for name in page_cmds(lib_src, excl, strong_only=True):   # \\includepdf / titlepage：整页，不当章标题
        heading_cmds.pop(name, None)
    pages = {n: o for n, o in page_cmds(lib_src, excl).items() if n not in heading_cmds}
    exp.protected |= set(pages)
    body_x = exp.expand(body)

    parser = Parser(exp.t.theorems, heading_cmds, caption_cmds, envs, pages)
    cjk_doc = len(re.findall(r"[\u4e00-\u9fff]", detok(body))) > 50
    rp = range_phrase_from(cls_src + "\n" + strip_comments(read_text(main)[0]), cjk_doc)
    if rp:
        parser.si_range = rp
    blocks = parser.blocks(body_x)

    full_src = "\n".join(read_text(root / f)[0] for f in ctx.files)
    front = extract_front(strip_comments(full_src), cls, limit=100_000)
    doc = Document(cls=cls, cls_options=opts, meta=dict(front.fields), blocks=blocks)
    doc.top_level = "chapter" if any(isinstance(b, Heading) and b.name == "chapter" for b in walk(blocks)) \
        else "section"
    doc.stats["theorems"] = exp.t.theorems
    bst = re.findall(r"\\bibliographystyle\s*\{([^}]*)\}", strip_comments(full_src))
    for b in walk(blocks):
        if isinstance(b, Bibliography) and bst:
            b.bst = bst[-1].strip()
    doc.nocite = [k.strip() for g in re.findall(r"\\nocite\s*\{([^}]*)\}", strip_comments(full_src))
                  for k in g.split(",") if k.strip()]
    assign_numbers(doc, Scheme(doc.top_level).from_cls(cls_src + "\n" + strip_comments(full_src)))

    unsup = Counter(exp.t.unsupported) + parser.unsupported
    doc.unsupported = dict(unsup.most_common())
    doc.notes = list(dict.fromkeys(exp.t.notes + parser.notes +
                                   (["工程里找不到这些子文件：" + "、".join(ctx.missing[:8])] if ctx.missing else [])))
    doc.stats.update(_stats(doc, parser, ctx, heading_cmds))
    return doc


def _stats(doc: Document, parser: Parser, ctx: _Ctx, heading_cmds: dict) -> dict:
    c: Counter = Counter()
    refs, cites = [], []

    def scan(inls):
        for i in inls or ():
            if i.t == "ref":
                refs.append(i.v)
            elif i.t == "cite":
                cites.extend(i.data.get("keys", []))
            elif i.t == "footnote":
                c["footnote"] += 1

    for b in walk(doc.blocks):
        name = type(b).__name__
        if isinstance(b, Heading):
            c["heading_" + str(b.level)] += 1
            scan(b.title)
        elif isinstance(b, Float):
            c["float_" + b.type] += 1
            scan(b.caption)
        elif isinstance(b, Equation):
            c["equation"] += 1
            c["equation_numbers"] += sum(1 for x in b.number if x)
        else:
            c[name] += 1
        for attr in ("inlines",):
            scan(getattr(b, attr, None))
        for g in getattr(b, "grids", ()) or ():
            for row in g.rows:
                for cell in row.cells:
                    scan(cell.inlines)
        if isinstance(b, Raw):
            c["raw_" + b.env] += 1
    unresolved = sorted({r for r in refs if r not in doc.labels})
    return {"counts": dict(c), "refs": len(refs), "cites": len(cites), "cite_keys": sorted(set(cites)),
            "unresolved_refs": unresolved, "orphan_labels": parser.orphan_labels,
            "files": ctx.files, "encodings": dict(ctx.encodings), "cls_heading_cmds": sorted(heading_cmds)}
