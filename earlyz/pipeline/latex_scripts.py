"""LaTeX 脚本残留归一（2026-09-27 用户拍板）。

背景：用户从 AI 复制的内容常带 LaTeX 记法，产物里字面显示、看着就错：
  - `$^{1)}$`（作者单位标注）/ `$^{[1]}$`（文献引用）——模板要求它们是**真上标**；
  - `D_n` / `T_n^off` / `f_n^loc` / `T_n^max`（正文里的伪公式记法）。

本模块只归一**脚本记法**（上标/下标）成真格式（`w:rPr/w:vertAlign`）；裸 `$`、`\\frac{}`、
`\\sum` 等不猜，出 warning 报数（透明，不发明）。

**必须在编译链最前**（analyze 之前、作用于工作副本）：
段指纹 `_stable_paragraph_fingerprint` 含规范化文本，归一后文本变化 → 晚做会打乱计划里的
`_resolved_paragraphs`（命中集合错位 → 动作打空/打错段）。

判据（避免误伤）：
  1. 花括号形：`$^{1)}$` / `^{[1]}` / `$_{2}$`（`$` 成对与否都吃掉相邻的单个 `$`）；
  2. 单 token 形：基字符必须是**单个**字母或数字（数学变量约定）→ `D_n` / `T_n^off` / `2^10` 命中；
     多字母词 `file_name` / `a_b_c` **不动**（`(?<![A-Za-z0-9_])` + `(?!_)` 双闸门）；
  3. `T_n^off`（基_sub^sup）拆成 `T` + 下标 `n` + 上标 `off` 顺序排列（Word 普通 run 无法叠放
     sub-sup，也不生成公式——那是明确不做的事）。
"""
from __future__ import annotations

import copy
import re
import shutil
import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# 花括号形：`$^{…}$` / `$_{…}$`（脚本注记，`$` 成对吃掉）。
# 2026-10-17：`$` 必须成对捕获——旧写法两侧 `\$?` 独立可选，`$E=mc^{2}$`
# 会把行内公式的闭合 `$` 当注记尾巴吃掉，变成 `$E=mc2`，后续公式转对象全失效。
_RE_BRACED = re.compile(
    r"(\$?)\^\{([^{}]{1,40})\}(\$?)|(\$?)_\{([^{}]{1,40})\}(\$?)")

# 单 token 形：基字符 = 单个字母/数字，且前面不是字母数字或下划线；末尾不得紧跟下划线
_RE_TOKEN = re.compile(
    r"(?<![A-Za-z0-9_])"
    r"(?:"
    r"([A-Za-z0-9])_([A-Za-z0-9]{1,12})\^([A-Za-z0-9]{1,12})(?!_)"   # 基_sub^sup
    r"|([A-Za-z0-9])_([A-Za-z0-9]{1,12})(?!_)"                       # 基_sub
    r"|([A-Za-z0-9])\^([A-Za-z0-9]{1,12})(?!_)"                      # 基^sup
    r")"
)

_PARTS = re.compile(r"^word/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$")

# 行内公式区（`$…$`，内容不以 `^`/`_` 开头的才是公式；以脚本记法开头的是注记）
_RE_INLINE_MATH = re.compile(r"\$(?![\^_])([^$\n]{1,200})\$")


def _matches(text: str) -> list[tuple[int, int, str, str]]:
    """返回 [(start, end, kind, content)]，kind ∈ {"sup", "sub"}；按位置排序、去重叠。

    `$` 口径：`$^{1)}$` 这类**脚本注记**的 `$` 成对吃掉；但 `$E=mc^{2}$` 里的 `$`
    是行内公式定界符——只有当本段前面已经吃掉过一个「紧邻脚本的开 `$`」（即
    pending 的奇数个）时，脚本后的 `$` 才当作配对闭合；否则保留给公式转换。
    """
    out: list[tuple[int, int, str, str]] = []
    # 真公式区（`$E=mc^{2}$`）不归脚本归一管——留给 toggle_tex_equations 转对象。
    # 2026-10-17 实测：先归一会把 `^{2}` 拆成上标、`$` 被吃掉，公式再也转不了。
    math_ranges = [(m.start(), m.end()) for m in _RE_INLINE_MATH.finditer(text)]

    def _in_math(a: int, b: int) -> bool:
        return any(a >= s and b <= e for s, e in math_ranges)

    pending_dollars = 0
    for m in _RE_BRACED.finditer(text):
        if _in_math(m.start(), m.end()):
            continue          # 行内公式区整体不动（含 `$`）
        if m.group(2) is not None:
            lead, content, tail = m.group(1), m.group(2), m.group(3)
            kind = "sup"
        else:
            lead, content, tail = m.group(4), m.group(5), m.group(6)
            kind = "sub"
        end = m.end()
        if lead == "$":
            pending_dollars += 1
        elif tail == "$" and pending_dollars % 2 == 1:
            pending_dollars += 1          # 闭合前面注记的开 `$`
        elif tail == "$":
            end = m.end() - 1             # 留给行内公式：保留闭合 `$`
        out.append((m.start(), end, kind, content))
    for m in _RE_TOKEN.finditer(text):
        g = m.groups()
        # 注意：span 必须从标记符（`_` / `^`）算起，把标记符一起吃掉，否则残留 "T_" "^" 字样
        if g[0] is not None:            # 基_sub^sup：基字符留在前段普通文本，只拆两个脚本
            out.append((m.start(2) - 1, m.end(2), "sub", g[1]))
            out.append((m.start(3) - 1, m.end(3), "sup", g[2]))
        elif g[3] is not None:          # 基_sub
            out.append((m.start(5) - 1, m.end(5), "sub", g[4]))
        else:                           # 基^sup
            out.append((m.start(7) - 1, m.end(7), "sup", g[6]))
    # 去重叠：按 start 排序，同 start 取最长（保证 基_sub^sup 整体优先于单独的 基_sub）
    out.sort(key=lambda t: (t[0], -(t[1] - t[0])))
    kept: list[tuple[int, int, str, str]] = []
    last_end = -1
    for item in out:
        if item[0] < last_end:
            continue
        kept.append(item)
        last_end = item[1]
    return kept


def _make_run(rpr_template, text: str, vert: str | None = None):
    r = etree.Element(f"{W}r")
    if rpr_template is not None or vert:
        # 2026-10-08 真机缺陷：裸 run（无 rPr）此前直接跳过整个 rPr 分支 →
        # 上下标 run 连 vertAlign 一起丢（「D_n→Dn」全篇下划线消失，stats 却
        # 声称转了 47 处）。vert 存在时即使无模板也要造一个 rPr 装 vertAlign。
        rpr = (copy.deepcopy(rpr_template) if rpr_template is not None
               else etree.Element(f"{W}rPr"))
        if vert:
            # vertAlign 必须按 CT_RPr 顺序插（否则 Word 静默忽略）
            from earlyz.actions.new_actions import _rpr_insert_ordered
            for old in rpr.findall(f"{W}vertAlign"):
                rpr.remove(old)
            el = etree.Element(f"{W}vertAlign")
            el.set(f"{W}val", "superscript" if vert == "sup" else "subscript")
            _rpr_insert_ordered(rpr, el)
        r.append(rpr)
    t = etree.SubElement(r, f"{W}t")
    if text != text.strip():
        t.set(XML_SPACE, "preserve")
    t.text = text
    return r


def _rewrite_run(run, stats: dict) -> bool:
    """把一个 run 里的脚本记法拆成多个 run。返回是否改动。"""
    t = run.find(f"{W}t")
    if t is None or not t.text:
        return False
    text = t.text
    spans = _matches(text)
    if not spans:
        return False
    parent = run.getparent()
    idx = parent.index(run)
    rpr = run.find(f"{W}rPr")
    new_runs = []
    pos = 0
    for start, end, kind, content in spans:
        if start > pos:
            new_runs.append(_make_run(rpr, text[pos:start]))
        new_runs.append(_make_run(rpr, content, vert=kind))
        stats["superscript" if kind == "sup" else "subscript"] += 1
        pos = end
    if pos < len(text):
        new_runs.append(_make_run(rpr, text[pos:]))
    parent.remove(run)
    for i, r in enumerate(new_runs):
        parent.insert(idx + i, r)
    return True


def _is_plain_text_run(r) -> bool:
    """run 只含 rPr + 一个 w:t（br/tab/drawing/fldChar 一律不算）。"""
    if r.tag != f"{W}r":
        return False
    kids = [k for k in r if not k.tag.endswith("}rPr")]
    return len(kids) == 1 and kids[0].tag == f"{W}t"


def _same_rpr(a, b) -> bool:
    ra, rb = a.find(f"{W}rPr"), b.find(f"{W}rPr")
    if (ra is None) != (rb is None):
        return False
    if ra is None:
        return True
    return etree.tostring(ra) == etree.tostring(rb)


def _merge_adjacent_text_runs(paragraph) -> bool:
    """合并**紧邻**且 rPr 逐字节一致、只含一个 w:t 的 run（跨 run 拆分的脚本记法需要）。

    保守边界：只合并直接相邻的 w:r——中间夹 proofErr/bookmarkStart 等一律跳过
    （书签标记文本范围，跨合并会挪动文本相对位置）；只对**确实含脚本记法**的段落调用。
    """
    changed = False
    while True:
        kids = list(paragraph)
        merged = False
        for i in range(len(kids) - 1):
            a, b = kids[i], kids[i + 1]
            if not _is_plain_text_run(a) or not _is_plain_text_run(b) or not _same_rpr(a, b):
                continue
            ta, tb = a.find(f"{W}t"), b.find(f"{W}t")
            ta.text = (ta.text or "") + (tb.text or "")
            if ta.text != ta.text.strip():
                ta.set(XML_SPACE, "preserve")
            paragraph.remove(b)
            merged = True
            changed = True
            break
        if not merged:
            return changed


def normalize_latex_scripts_in_docx(path: str | Path) -> dict:
    """就地归一一个 docx 的 LaTeX 脚本残留；返回统计（changed/上标数/下标数/残留）。"""
    path = Path(path)
    stats = {"changed": False, "superscript": 0, "subscript": 0,
             "changed_parts": 0, "merged_runs": 0, "dollar_left": 0, "residual": 0}
    with zipfile.ZipFile(path) as z:
        items = [(n, z.read(n)) for n in z.namelist()]

    replaced: dict[str, bytes] = {}
    for name, data in items:
        if not _PARTS.match(name):
            continue
        try:
            root = etree.fromstring(data)
        except etree.XMLSyntaxError:
            continue
        part_changed = False
        for para in root.iter(f"{W}p"):
            # 跨 run 拆分：先看整段文本有没有记法；有才做有界合并（最小改动）
            joined = "".join(x.text or "" for x in para.iter(f"{W}t"))
            if not (_RE_BRACED.search(joined) or _RE_TOKEN.search(joined)):
                continue
            if _merge_adjacent_text_runs(para):
                stats["merged_runs"] += 1
            for run in para.iter(f"{W}r"):
                if _rewrite_run(run, stats):
                    part_changed = True
        if not part_changed:
            # 残留/`$` 统计与是否改动无关：如实报告"有未覆盖形态"，供排查
            for p in root.iter(f"{W}p"):
                joined = "".join(x.text or "" for x in p.iter(f"{W}t"))
                if _RE_BRACED.search(joined) or _RE_TOKEN.search(joined):
                    stats["residual"] += 1
                stats["dollar_left"] += joined.count("$")
            continue
        # 段落级残留统计（跨 run 拆分 / 未覆盖形态）
        for p in root.iter(f"{W}p"):
            joined = "".join(x.text or "" for x in p.iter(f"{W}t"))
            if _RE_BRACED.search(joined) or _RE_TOKEN.search(joined):
                stats["residual"] += 1
            stats["dollar_left"] += joined.count("$")
        stats["changed_parts"] += 1
        replaced[name] = etree.tostring(root, xml_declaration=True,
                                        encoding="UTF-8", standalone=True)

    if not replaced:
        return stats
    stats["changed"] = True
    tmp = path.with_suffix(".latexfix.tmp")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = replaced.get(item.filename, zin.read(item.filename))
            zout.writestr(item, data)
    shutil.move(str(tmp), str(path))
    return stats


def normalize_latex_scripts_in_copy(src: str | Path, workdir: str | Path) -> tuple[Path, dict]:
    """把 src 复制到 workdir（同名）后归一；无改动则原样返回 src（用户原件永不修改）。"""
    src = Path(src)
    workdir = Path(workdir)
    workdir.mkdir(parents=True, exist_ok=True)
    copy_path = workdir / src.name
    shutil.copy2(src, copy_path)
    stats = normalize_latex_scripts_in_docx(copy_path)
    if not stats["changed"]:
        copy_path.unlink(missing_ok=True)
        return src, stats
    return copy_path, stats
