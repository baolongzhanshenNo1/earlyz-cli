"""L1 宏展开器：定义收集 / 使用处展开 / 条件求值。数据结构与读参工具在 macros.py。"""
from __future__ import annotations

import re
from collections import deque

from .macros import (_DEFINERS, _NUMERIC_IF, MAX_EXPANSIONS, EnvDef, Macro, MacroTable, _cs_name,
                     _safe, _skip_sp, _subst, detok, plain, read_group, read_opt, read_star)
from .tokens import Tok, lex

# 引擎开关：中文文档绝大多数走 XeLaTeX，按 XeTeX 取分支
_ENGINE_IFS = {"ifpdf": False, "ifxetex": True, "ifluatex": False, "ifPDFTeX": False,
               "ifXeTeX": True, "ifLuaTeX": False}
MATH_ENVS = frozenset({"equation", "equation*", "align", "align*", "gather", "gather*", "multline",
                       "multline*", "eqnarray", "eqnarray*", "displaymath", "math", "flalign",
                       "flalign*", "alignat", "alignat*"})


def _is_cond(name: str) -> bool:
    return name.startswith("if") and name not in ("iff", "ifthenelse")


def _peek_env(q: deque) -> tuple[str, int]:
    from .parse_inline import peek_env
    return peek_env(q)


class Expander:
    def __init__(self, table: MacroTable | None = None, protected=frozenset()):
        self.t = table or MacroTable()
        self.protected = set(protected)       # 可变：document.py 加载 cls 后补题注类命令
        self._mm = False                     # 当前是否在数学模式（\ifmmode 用）
        self._over = False
        # cls/sty 里不收的内部定义（带 @ / 排版原语）：不展开，只留底给 L2 判断「像章标题的命令」与参数个数
        self.raw: dict = {}
        self.raw_envs: dict = {}

    # ── 定义 ──────────────────────────────────────────────────────
    def load_lib(self, text: str, file: str) -> int:
        """工程内 cls/sty：只收公共宏（_safe），不输出内容。返回新收的条数。"""
        before = len(self.t.macros) + len(self.t.envs)
        q = deque(lex(text, file, at_letter=True))
        expl = False                      # expl3 区段（_ : 是字母）词法不对，整段不收
        while q:
            tok = q.popleft()
            if tok.k != "cs":
                continue
            if tok.v in ("ExplSyntaxOn", "ExplSyntaxOff"):
                expl = tok.v == "ExplSyntaxOn"
            elif tok.v in _DEFINERS and not expl:
                self._define(tok.v, q, lib=True)
        return len(self.t.macros) + len(self.t.envs) - before

    def _store(self, name: str, m: Macro, lib: bool, provide: bool = False) -> None:
        if not name or name in self.protected:
            return
        if lib and not _safe(name, m.body):
            self.raw.setdefault(name, m)
            return
        if provide and name in self.t.macros:
            return
        m.lib = lib
        self.t.macros[name] = m

    @staticmethod
    def _nargs(n) -> int:
        try:
            return max(0, min(9, int(plain(n)))) if n is not None else 0
        except ValueError:
            return 0

    def _define(self, kind: str, q: deque, lib: bool) -> None:
        if kind in ("newcommand", "renewcommand", "providecommand", "DeclareRobustCommand"):
            read_star(q)
            name = _cs_name(read_group(q))
            n, d = read_opt(q), read_opt(q)
            body = read_group(q) or []
            self._store(name, Macro(self._nargs(n), d, body), lib, kind == "providecommand")
        elif kind == "DeclareMathOperator":
            star = read_star(q)
            name = _cs_name(read_group(q))
            body = read_group(q) or []
            op = ([Tok("cs", "operatorname")] + ([Tok("char", "*")] if star else [])
                  + [Tok("bg", "{")] + body + [Tok("eg", "}")])
            self._store(name, Macro(0, None, op), lib)
        elif kind in ("def", "gdef", "edef", "xdef"):
            _skip_sp(q)
            if not q:
                return
            nt = q.popleft()
            name = nt.v if nt.k == "cs" else ("~" if nt.k == "active" else "")
            params = []
            while q and q[0].k != "bg":
                params.append(q.popleft())
            body = read_group(q) or []
            ok = len(params) % 2 == 0 and all(
                params[i].k == "param" and params[i + 1].k == "char" and params[i + 1].v == str(i // 2 + 1)
                for i in range(0, len(params), 2))
            if not ok:
                if not lib:
                    self.t.unsupported["\\def 带分隔参数"] += 1
                return
            self._store(name, Macro(len(params) // 2, None, body), lib)
        elif kind == "let":
            _skip_sp(q)
            if not q:
                return
            a = q.popleft()
            _skip_sp(q)
            if q and q[0].ch("="):
                q.popleft()
                _skip_sp(q)
            if not q or a.k != "cs":
                return
            b = q.popleft()
            src = self.t.macros.get(b.v) if b.k == "cs" else None
            m = Macro(src.nargs, src.default, list(src.body)) if src else Macro(0, None, [b])
            self._store(a.v, m, lib)
        elif kind in ("newenvironment", "renewenvironment"):
            read_star(q)
            name = plain(read_group(q) or [])
            n, d = read_opt(q), read_opt(q)
            b, e = read_group(q) or [], read_group(q) or []
            if not name or name in self.protected:
                return
            env = EnvDef(self._nargs(n), d, b, e)
            if lib and not (_safe(name, b) and _safe(name, e)):
                self.raw_envs.setdefault(name, env)
                return
            self.t.envs[name] = env
        elif kind == "newif":
            _skip_sp(q)
            if q and q[0].k == "cs" and q[0].v.startswith("if"):
                self.t.flags[q.popleft().v[2:]] = False
        elif kind in ("newtheorem", "newdefinition", "newproof"):
            star = read_star(q)
            pre = read_opt(q)                         # fduthesis：\newtheorem[style=…,counter=thm]{名}{标题}
            name = plain(read_group(q) or [])
            shared = read_opt(q)
            if shared is None and pre is not None:
                m = re.search(r"counter\s*=\s*([A-Za-z]+)", plain(pre))
                shared = [Tok("char", ch) for ch in m.group(1)] if m else None
            title = read_group(q) or []
            within = read_opt(q)
            if name:
                self.t.theorems.setdefault(name, {
                    "title": detok(title).strip(), "numbered": not star and kind != "newproof",
                    "shared": plain(shared) if shared else "", "within": plain(within) if within else ""})
        elif kind == "theoremstyle":
            read_group(q)
        elif kind in ("NewDocumentCommand", "RenewDocumentCommand", "ProvideDocumentCommand",
                      "DeclareDocumentCommand"):
            self._xparse(kind, q, lib)
        elif kind in ("csdef", "csgdef", "csedef", "csxdef"):           # etoolbox：\csdef{名}{体}
            name = plain(read_group(q) or [])
            self._store(name, Macro(0, None, read_group(q) or []), lib)

    def _xparse(self, kind: str, q: deque, lib: bool) -> None:
        """xparse 只认 m / O{默认}（可按 #n 直接替换的两种）；s/o/v/D… 要条件判断，记 unsupported 不展开。"""
        _skip_sp(q)
        if q and q[0].k == "bg":
            name = _cs_name(read_group(q))
        else:
            name = q.popleft().v if q and q[0].k == "cs" else ""
        spec = read_group(q) or []
        body = read_group(q) or []
        ok, n, default, i = True, 0, None, 0
        while i < len(spec):
            s = spec[i]
            i += 1
            if s.k == "space":
                continue
            if s.ch("m"):
                n += 1
            elif s.ch("O") and n == 0 and default is None and i < len(spec) and spec[i].k == "bg":
                depth, d, i = 1, [], i + 1
                while i < len(spec) and depth:
                    if spec[i].k == "bg":
                        depth += 1
                    elif spec[i].k == "eg":
                        depth -= 1
                        if depth == 0:
                            break
                    d.append(spec[i])
                    i += 1
                i += 1
                default, n = d, 1
            else:
                ok = False
                break
        if not ok:
            if not lib:
                self.t.unsupported["\\" + name + "（xparse 参数规格）"] += 1
            return
        self._store(name, Macro(n, default, body), lib, kind == "ProvideDocumentCommand")

    # ── 展开 ──────────────────────────────────────────────────────
    def _args(self, nargs: int, default, q: deque) -> list:
        args: list = []
        if nargs <= 0:
            return args
        if default is not None:
            o = read_opt(q)
            args.append(default if o is None else o)
        while len(args) < nargs:
            args.append(read_group(q) or [])
        return args

    def _push(self, q: deque, toks: list) -> bool:
        if self.t.expansions >= MAX_EXPANSIONS:
            if not self._over:
                self.t.notes.append("宏展开次数超过上限（疑似递归定义），后文的宏不再展开")
                self._over = True
            return False
        self.t.expansions += 1
        q.extendleft(reversed(toks))
        return True

    def expand(self, toks) -> list[Tok]:
        q, out, t = deque(toks), [], self.t
        while q:
            tok = q.popleft()
            if tok.k == "math":
                self._mm = not self._mm
            if tok.k != "cs":
                out.append(tok)
                continue
            v = tok.v
            if v in ("(", "["):
                self._mm = True
            elif v in (")", "]"):
                self._mm = False
            if v in _DEFINERS:
                self._define(v, q, lib=False)
                continue
            if v in ("fi", "else", "or"):
                continue                    # 配对残留；未闭合的条件在 _cond 里已记 notes
            if _is_cond(v) and v not in t.macros:
                self._cond(v, q)
                continue
            if v == "ifthenelse":
                read_group(q)
                a = read_group(q) or []
                read_group(q)
                t.unsupported["\\ifthenelse（取第一个分支）"] += 1
                self._push(q, a)
                continue
            if v.endswith("true") and v[:-4] in t.flags:
                t.flags[v[:-4]] = True
                continue
            if v.endswith("false") and v[:-5] in t.flags:
                t.flags[v[:-5]] = False
                continue
            if v in ("begin", "end"):
                name, n = _peek_env(q)
                if name in MATH_ENVS:
                    self._mm = v == "begin"
                e = t.envs.get(name)
                if e is not None and name not in self.protected:
                    for _ in range(n):
                        q.popleft()
                    body = _subst(e.begin, self._args(e.nargs, e.default, q)) if v == "begin" else e.end
                    if self._push(q, list(body)):
                        continue
                out.append(tok)
                continue
            m = t.macros.get(v)
            if m is not None and v not in self.protected:
                if self._push(q, _subst(m.body, self._args(m.nargs, m.default, q))):
                    continue
            out.append(tok)
        return out

    # ── 条件 ──────────────────────────────────────────────────────
    def _defined(self, name: str) -> bool:
        return name in self.t.macros or name in self.protected

    def _ifx(self, a: Tok | None, b: Tok | None) -> bool:
        if a is None or b is None:
            return False
        if b.k == "cs" and b.v in ("undefined", "@undefined"):
            return not (a.k == "cs" and self._defined(a.v))
        if a.k == "cs" and b.k == "cs":
            if a.v == b.v:
                return True
            ma, mb = self.t.macros.get(a.v), self.t.macros.get(b.v)
            return bool(ma and mb and ma.nargs == mb.nargs and detok(ma.body) == detok(mb.body))
        return a.k == b.k and a.v == b.v

    def _branches(self, q: deque) -> tuple[list, list, bool]:
        yes: list = []
        no: list = []
        cur, depth = yes, 0
        while q:
            tok = q.popleft()
            if tok.k == "cs":
                if _is_cond(tok.v) and tok.v not in self.t.macros:
                    depth += 1
                elif tok.v == "fi":
                    if depth == 0:
                        return yes, no, True
                    depth -= 1
                elif tok.v in ("else", "or") and depth == 0:
                    cur = no
                    continue
            cur.append(tok)
        return yes, no, False

    def _cond(self, v: str, q: deque) -> None:
        t = self.t
        if v == "iftrue":
            val = True
        elif v == "iffalse":
            val = False
        elif v[2:] in t.flags:
            val = t.flags[v[2:]]
        elif v == "ifdefined":
            _skip_sp(q)
            val = self._defined(q.popleft().v) if q else False
        elif v == "ifcsname":
            name = []
            while q and not q[0].cs("endcsname"):
                name.append(q.popleft())
            if q:
                q.popleft()
            val = self._defined(plain(name))
        elif v == "ifx":
            _skip_sp(q)
            a = q.popleft() if q else None
            _skip_sp(q)
            val = self._ifx(a, q.popleft() if q else None)
        elif v == "ifmmode":
            val = self._mm
        elif v in _NUMERIC_IF:
            val = None
        elif v in _ENGINE_IFS:
            val = _ENGINE_IFS[v]
        else:
            val = True
            t.unsupported["未知条件 \\" + v + "（取真分支）"] += 1
        yes, no, closed = self._branches(q)
        if not closed:
            t.notes.append("条件 \\" + v + " 没有对应的 \\fi，后文按真分支处理")
            val = True if val is None else val
        if val is None:
            t.unsupported["数值条件 \\" + v + "（整段略过）"] += 1
            return
        self._push(q, yes if val else no)
