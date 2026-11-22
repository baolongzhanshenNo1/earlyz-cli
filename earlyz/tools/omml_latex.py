"""OMML → LaTeX 文本化（2026-09-27）。

用途：**让公式对 LLM 可见**。Word 原生公式是 OMML（`m:oMath`），里面没有 `w:t`
文本——python-docx / 纯文本抽取都看不到，模型只能看到一段空文本，于是「公式段为什么
没居中」「编号为什么没右对齐」这类问题无从回答，agent 也没法按公式下指令。

本模块把常见 OMML 构造转成 LaTeX 串注入 LLM 上下文；MathType 插入的是 OLE 对象
（二进制流，内容取不到）→ 调用方用段号占位（"公式段 #N（MathType）"），支持按段号操作。

设计：未知构造**递归进子元素**（宁可不完整，绝不丢内容）；不追求 LaTeX 完美，
只求模型能读懂结构（上下标/分式/根式/求和/矩阵/括号）。
"""
from __future__ import annotations

M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# nary 字符 → LaTeX 命令
_NARY = {
    "∑": r"\sum", "∏": r"\prod", "∐": r"\coprod", "∫": r"\int",
    "∬": r"\iint", "∭": r"\iiint", "∮": r"\oint", "⋃": r"\bigcup",
    "⋂": r"\bigcap", "⋁": r"\bigvee", "⋀": r"\bigwedge",
}
# acc 字符 → LaTeX 命令
_ACC = {
    "̂": r"\hat", "^": r"\hat", "̄": r"\bar", "¯": r"\bar", "̇": r"\dot",
    "̈": r"\ddot", "̆": r"\breve", "̌": r"\check", "̃": r"\tilde",
    "→": r"\vec", "⃗": r"\vec",
}


def _txt(el) -> str:
    return "".join((t.text or "") for t in el.iter(f"{M}t"))


def _children_latex(el) -> str:
    """递归拼接子元素的 LaTeX（未知构造的兜底）。"""
    return "".join(_node(child) for child in el)


def _node(el) -> str:
    tag = el.tag
    if tag == f"{M}t":
        return el.text or ""
    if tag == f"{M}r":
        return _children_latex(el)
    if tag in (f"{M}oMath", f"{M}oMathPara", f"{M}e", f"{M}num", f"{M}den",
               f"{M}sub", f"{M}sup", f"{M}deg", f"{M}lim", f"{M}box",
               f"{M}borderBox", f"{M}phant", f"{M}groupChr", f"{M}limLow",
               f"{M}limUpp"):
        return _children_latex(el)
    if tag in (f"{M}ctrlPr", f"{M}rPr", f"{M}argPr", f"{M}argSz", f"{M}fPr",
               f"{M}dPr", f"{M}naryPr", f"{M}radPr", f"{M}sSupPr", f"{M}sSubPr",
               f"{M}sSubSupPr", f"{M}accPr", f"{M}barPr", f"{M}mPr", f"{M}eqArrPr"):
        return ""
    if tag == f"{M}f":                      # 分式
        return r"\frac{%s}{%s}" % (_children_latex(el.find(f"{M}num"))
                                   if el.find(f"{M}num") is not None else "",
                                   _children_latex(el.find(f"{M}den"))
                                   if el.find(f"{M}den") is not None else "")
    if tag == f"{M}sSup":
        base = el.find(f"{M}e")
        sup = el.find(f"{M}sup")
        return "{%s}^{%s}" % (_children_latex(base) if base is not None else "",
                              _children_latex(sup) if sup is not None else "")
    if tag == f"{M}sSub":
        base = el.find(f"{M}e")
        sub = el.find(f"{M}sub")
        return "{%s}_{%s}" % (_children_latex(base) if base is not None else "",
                              _children_latex(sub) if sub is not None else "")
    if tag == f"{M}sSubSup":
        base = el.find(f"{M}e")
        sub = el.find(f"{M}sub")
        sup = el.find(f"{M}sup")
        return "{%s}_{%s}^{%s}" % (
            _children_latex(base) if base is not None else "",
            _children_latex(sub) if sub is not None else "",
            _children_latex(sup) if sup is not None else "")
    if tag == f"{M}sPre":
        base = el.find(f"{M}e")
        sub = el.find(f"{M}sub")
        sup = el.find(f"{M}sup")
        pre = "".join(x for x in (
            "_{%s}" % _children_latex(sub) if sub is not None else "",
            "^{%s}" % _children_latex(sup) if sup is not None else "") if x)
        return "%s{%s}" % (pre, _children_latex(base) if base is not None else "")
    if tag == f"{M}rad":                    # 根式
        deg = el.find(f"{M}deg")
        base = el.find(f"{M}e")
        deg_txt = _children_latex(deg) if deg is not None else ""
        if deg_txt.strip():
            return r"\sqrt[%s]{%s}" % (deg_txt, _children_latex(base) if base is not None else "")
        return r"\sqrt{%s}" % (_children_latex(base) if base is not None else "")
    if tag == f"{M}nary":                   # 求和/积分等
        pr = el.find(f"{M}naryPr")
        chr_ = None
        if pr is not None:
            c = pr.find(f"{M}chr")
            if c is not None:
                chr_ = c.get(f"{M}val")
        op = _NARY.get(chr_ or "∫", r"\int")
        sub = el.find(f"{M}sub")
        sup = el.find(f"{M}sup")
        base = el.find(f"{M}e")
        s = op
        if sub is not None and _children_latex(sub).strip():
            s += "_{%s}" % _children_latex(sub)
        if sup is not None and _children_latex(sup).strip():
            s += "^{%s}" % _children_latex(sup)
        return s + ("{%s}" % _children_latex(base) if base is not None else "")
    if tag == f"{M}d":                      # 括号/定界符
        pr = el.find(f"{M}dPr")
        beg, end = "(", ")"
        if pr is not None:
            b = pr.find(f"{M}begChr")
            e = pr.find(f"{M}endChr")
            if b is not None:
                beg = b.get(f"{M}val") or ""
            if e is not None:
                end = e.get(f"{M}val") or ""
        inner = "".join(_children_latex(x) for x in el.findall(f"{M}e"))
        return f"{beg}{inner}{end}"
    if tag == f"{M}func":                   # 函数
        name = el.find(f"{M}fName")
        base = el.find(f"{M}e")
        return "%s(%s)" % (_children_latex(name) if name is not None else "",
                           _children_latex(base) if base is not None else "")
    if tag == f"{M}acc":                    # 重音符
        pr = el.find(f"{M}accPr")
        ch = None
        if pr is not None:
            c = pr.find(f"{M}chr")
            if c is not None:
                ch = c.get(f"{M}val")
        cmd = _ACC.get(ch or "^", r"\hat")
        base = el.find(f"{M}e")
        return "%s{%s}" % (cmd, _children_latex(base) if base is not None else "")
    if tag == f"{M}bar":
        base = el.find(f"{M}e")
        return r"\bar{%s}" % (_children_latex(base) if base is not None else "")
    if tag == f"{M}m":                      # 矩阵
        rows = []
        for mr in el.findall(f"{M}mr"):
            rows.append(" & ".join(_children_latex(c) for c in mr.findall(f"{M}e")))
        return r"\begin{matrix}%s\end{matrix}" % r" \\ ".join(rows)
    if tag == f"{M}eqArr":                  # 多行公式组
        return r" \\ ".join(_children_latex(x) for x in el.findall(f"{M}e"))
    return _children_latex(el)


def omml_to_latex(el) -> str:
    """把一个 `m:oMath` / `m:oMathPara` 元素转成 LaTeX 串（未知构造递归兜底）。"""
    if el is None:
        return ""
    return _node(el).strip()


def omath_latex_in(paragraph) -> str:
    """段落里全部 OMML（含 oMathPara）的 LaTeX，按出现顺序拼接。"""
    out = []
    for m in paragraph.iter(f"{M}oMath"):
        txt = omml_to_latex(m)
        if txt:
            out.append(txt)
    return " ".join(out)
