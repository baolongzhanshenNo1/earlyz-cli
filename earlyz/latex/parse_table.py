"""L2 解析（表格）：tabular 类 token → Grid。

行用 `\\\\` 切，单元格用 `&` 切（只认顶层，组内的不算）；booktabs 线、\\hline、\\cline、\\cmidrule
记到行上；\\multicolumn / \\multirow 记跨度；longtable 的 \\endfirsthead/\\endhead 之前的行记为表头。
单元格内容复用行内解析（数学、引用、样式都保留）。
"""
from __future__ import annotations

import re
from collections import deque

from .ir import Cell, Grid, Row
from .macros import _skip_sp, detok, plain, read_group, read_opt, read_star
from .parse_text import NOOP

RULES = {"hline": "hline", "toprule": "toprule", "midrule": "midrule", "bottomrule": "bottomrule",
         "specialrule": "midrule", "Xhline": "hline", "hhline": "hline"}
TABULAR_ENVS = {"tabular": 1, "tabular*": 2, "tabularx": 2, "tabulary": 2, "longtable": 1,
                "longtable*": 1, "array": 1, "tabu": 1, "supertabular": 1, "xltabular": 2,
                "longtabu": 1, "tblr": 1, "longtblr": 1}
_HEAD_MARKS = ("endfirsthead", "endhead")
_FOOT_MARKS = ("endfoot", "endlastfoot")


def col_aligns(spec: str) -> list[str]:
    """列格式 → 每列对齐（l c r p{} X m{} b{}）；@{} !{} >{} <{} 的参数整组跳过。"""
    out, i, n = [], 0, len(spec)
    while i < n:
        c = spec[i]
        if c in "@!><":
            j = spec.find("{", i)
            i = _skip_brace(spec, j) if j >= 0 else i + 1
            continue
        if c == "*":                        # *{3}{c} → 展开
            m = re.match(r"\*\s*\{(\d+)\}\s*\{", spec[i:])
            if m:
                body_start = i + m.end() - 1
                end = _skip_brace(spec, body_start)
                out += col_aligns(spec[body_start + 1:end - 1]) * int(m.group(1))
                i = end
                continue
        if c in "lcrLCRSXJ":
            out.append({"l": "left", "c": "center", "r": "right", "S": "center"}.get(c.lower() if c in "LCR" else c, "left"))
        elif c in "pmbw":
            out.append("left")
            j = spec.find("{", i)
            i = _skip_brace(spec, j) if j >= 0 else i + 1
            if c == "w":
                j = spec.find("{", i)
                i = _skip_brace(spec, j) if j >= 0 else i
            continue
        i += 1
    return out


def _skip_brace(s: str, j: int) -> int:
    depth = 0
    for k in range(j, len(s)):
        if s[k] == "{":
            depth += 1
        elif s[k] == "}":
            depth -= 1
            if depth == 0:
                return k + 1
    return len(s)


def split_top(toks: list, kind: str) -> list[list]:
    """按顶层（不在 {} 里）的 kind token 切开。kind='align' 或 'row'（\\\\ 与 \\tabularnewline）。"""
    parts, cur, depth = [], [], 0
    q = deque(toks)
    while q:
        t = q.popleft()
        if t.k == "bg":
            depth += 1
        elif t.k == "eg":
            depth -= 1
        if depth == 0 and ((kind == "align" and t.k == "align") or
                           (kind == "row" and t.cs("\\", "tabularnewline", "crcr", "cr"))):
            if kind == "row":
                read_star(q)
                read_opt(q)
            parts.append(cur)
            cur = []
            continue
        cur.append(t)
    parts.append(cur)
    return parts


def _strip_rules(toks: list, partial: list) -> tuple[list, list[str]]:
    """去掉行首的线命令（\\hline \\toprule \\cmidrule(lr){2-3} \\cline{2-3}），返回 (剩余, 线类型列表)。"""
    q, rules = deque(toks), []
    while True:
        _skip_sp(q)
        while q and q[0].k == "par":
            q.popleft()
            _skip_sp(q)
        if not q or q[0].k != "cs":
            break
        v = q[0].v
        if v in RULES:
            q.popleft()
            rules.append(RULES[v])
            if v == "specialrule":
                read_group(q), read_group(q), read_group(q)
            elif v in ("Xhline", "hhline"):
                read_group(q)
            else:
                read_opt(q)
        elif v in ("cmidrule", "cline"):
            q.popleft()
            read_opt(q, "(", ")")
            rng = plain(read_group(q) or [])
            m = re.match(r"(\d+)\s*-\s*(\d+)", rng)
            if m:
                partial.append((int(m.group(1)), int(m.group(2))))
        elif v in ("addlinespace", "noalign", "rowcolor", "arrayrulecolor"):
            q.popleft()
            read_opt(q)
            if v in ("noalign", "rowcolor", "arrayrulecolor"):
                read_group(q)
        else:
            break
    return list(q), rules


class TableMixin:
    def parse_grid(self, env: str, q: deque) -> Grid:
        """q 指向环境参数处（\\begin{tabular} 之后）；读到配平的 \\end{env} 为止。"""
        nreq = TABULAR_ENVS.get(env, 1)
        read_opt(q)                          # [t] / longtable 的 [c]
        args = [read_group(q) or [] for _ in range(nreq)]
        spec_src = detok(args[-1]) if args else ""
        return self.grid_from(env, spec_src, self.take_env(env, q))

    def grid_from(self, env: str, spec_src: str, body: list) -> Grid:
        """longtable 按 \\endfirsthead \\endhead \\endfoot \\endlastfoot 分段（每个标记结束它前面那一段）：
        表头取首页表头（没有就取续页表头），表尾取末页表尾（没有就取每页表尾）；续页表头/每页表尾只是
        跨页时重复的版式（「续表」「续下页」），不进 IR。普通 tabular 没有标记，整体就是表体。"""
        aligns = col_aligns(spec_src)
        grid = Grid(env, spec_src, [])
        sections, cur, depth = {}, [], 0
        for t in body:
            if t.k == "bg":
                depth += 1
            elif t.k == "eg":
                depth -= 1
            if depth == 0 and t.k == "cs" and t.v in _HEAD_MARKS + _FOOT_MARKS:
                sections.setdefault(t.v, cur)
                cur = []
                continue
            cur.append(t)
        head = sections.get("endfirsthead", sections.get("endhead"))
        foot = sections.get("endlastfoot", sections.get("endfoot"))
        if head:
            self.add_rows(grid, head, aligns)
            grid.header_rows = len(grid.rows)
        self.add_rows(grid, cur, aligns)
        if foot:
            self.add_rows(grid, foot, aligns)
        if spec_src.strip() and not aligns:
            self.notes.append("表格列格式「" + spec_src[:30] + "」没能识别，列对齐交给模板")
        return grid

    def add_rows(self, grid: Grid, toks: list, aligns: list) -> None:
        for rs in split_top(toks, "row"):
            rs = self.take_captions(rs)
            partial: list = []
            rs, rules = _strip_rules(rs, partial)
            content = any(t.k not in ("space", "par") and not (t.k == "cs" and t.v in NOOP) for t in rs)
            if not content:
                if grid.rows and rules:
                    grid.rows[-1].rule_below = rules[-1]     # 末尾 \bottomrule 挂到上一行
                    grid.rows[-1].partial += partial
                continue
            cells, col = [], 0
            for cs in split_top(rs, "align"):
                cell = self.parse_cell(cs, aligns, col)
                cells.append(cell)
                col += cell.colspan
            grid.rows.append(Row(cells, rule_above=rules[0] if rules else "", partial=partial))

    def take_captions(self, toks: list) -> list:
        """longtable 的 \\caption{…}\\label{…} 写在表行里：归本表（与 LaTeX 一样占表计数器）。"""
        if self.float_ctx is None:
            return toks
        q, keep, depth = deque(toks), [], 0
        while q:
            t = q.popleft()
            if t.k == "bg":
                depth += 1
            elif t.k == "eg":
                depth -= 1
            if depth == 0 and t.cs("caption", "bicaption"):
                self.caption_cmd(t.v, q, [])
                continue
            if depth == 0 and t.cs("label"):
                self.add_label(plain(read_group(q) or []))
                continue
            keep.append(t)
        return keep

    def parse_cell(self, toks: list, aligns: list, col: int) -> Cell:
        q = deque(toks)
        _skip_sp(q)
        colspan, rowspan, align = 1, 1, aligns[col] if col < len(aligns) else ""
        if q and q[0].cs("multicolumn"):
            q.popleft()
            try:
                colspan = max(1, int(plain(read_group(q) or [])))
            except ValueError:
                colspan = 1
            al = col_aligns(detok(read_group(q) or []))
            align = al[0] if al else align
            q = deque(read_group(q) or [])
            _skip_sp(q)
        if q and q[0].cs("multirow", "multirowcell"):
            q.popleft()
            read_opt(q)
            try:
                rowspan = max(1, int(plain(read_group(q) or [])))
            except ValueError:
                rowspan = 1
            read_opt(q)
            read_group(q)                    # 宽度 * / 3cm
            read_opt(q)
            q = deque(read_group(q) or [])
        if q and q[0].cs("makecell", "thead", "shortstack"):
            q.popleft()
            read_opt(q)
            q = deque(read_group(q) or [])
        return Cell(self.inlines(list(q)), colspan, rowspan, align)
