"""交叉引用生成（A2）——M3，2026-10-08。

机制（对齐 docs/方案-格式检测报告接入与A档七项-2026-10-07.md §A2）：
- ① 题注段加书签（`_RefF/_RefT` 系列，run 级拆分保证书签**精确覆盖**「图3-1」
  标签+编号区；公式编号 `_RefE` 覆盖 `(3-1)` 括号区）；
- ② 扫描正文「图 N / 表 N / 式(N)」文本 → 归一化比对**已存在**的编号 →
  待替换清单；确认后替换为 REF 域（` REF _RefF0001 \\h `，缓存=现编号字面）。
- 联动：编号由 A4（renumber_caption_numbers）改字面后，用户 F9 → REF 域从
  书签内容重算 → 引用同步（书签随 run 手术保留）。

边界：
- 只对已存在的编号生成；重号（同 kind 同编号多题注）=歧义 → 不猜（跳过列清单）；
- 正文引用段内已有的 REF 域（含用户自己做的交叉引用）不重复包裹 → 幂等；
- 公式引用须带语境字（如/见/据/由/参/入/按/依/照 + 式），防「格**式(**1**)**」误伤；
- 目录/图表目录条目（含 PAGEREF 域）、封面区、文本框、题注自身一律不参与扫描。
"""

from __future__ import annotations

import copy
import re
from pathlib import Path

from lxml import etree

from earlyz.actions.new_actions import (
    W,
    _EQ_NUM_ANY_RE,
    _cover_paragraphs,
    _equation_para_context,
    _para_full_text,
    _read_part,
    _read_styles_map,
    _set_vert_align,
    _write_part,
)
from earlyz.actions.captions import (
    _bookmark_names,
    _has_pageref,
    _in_textbox,
    _instr_text,
    _max_bookmark_id,
    _number_span,
    _replace_number_span,
    _run_char_len,
    _run_text,
    _walk_captions,
)

_M_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

_KINDS = ("figure", "table", "equation")
_KIND_LABEL = {"figure": "图", "table": "表", "equation": "式"}
_REF_PREFIX = {"figure": "_RefF", "table": "_RefT", "equation": "_RefE"}

# 正文引用 token：图/表 + 编号（数字可带 - . ．· 等分隔）
_REF_TOKEN_RE = re.compile(r"[图表]\s*\d+(?:\s*[-–—\u2010\u2011.．·]\s*\d+)*")
# 公式引用 token：语境字 + 式 + 括号编号（只替换括号编号部分，语境字与「式」字不动）
_EQ_TOKEN_RE = re.compile(
    r"(?<=[如见据由参入按依照])\s*式\s*"
    r"(?P<num>[（(]\s*\d+(?:\s*[-–—.．·]\s*\d+)*\s*[)）])")
# REF 域指令（排除 PAGEREF/STYLEREF）
_REF_INSTR_RE = re.compile(r"(?:(?<=\s)|^)REF\s")
_OUR_REF_RE = re.compile(r"(?:(?<=\s)|^)REF\s+(_Ref[FTE]\d+)")

_PH_OPEN, _PH_CLOSE = "\ue010", "\ue011"
_FW_DIGITS = str.maketrans("０１２３４５６７８９", "0123456789")


def _drtext(p) -> str:
    """段落内**直接子 run** 的可见文本拼接（与 run 手术坐标系一致；不含 hyperlink 内嵌 run）。

    noBreakHyphen 计为 "-"（2026-10-07：题注/引用里的「表5-1」短横线常是该元素，
    漏读会把编号读成「51」，导致引用与题注匹配失败）。
    """
    out: list[str] = []
    for r in p:
        if etree.QName(r).localname != "r":
            continue
        out.append(_run_text(r))
    return "".join(out)


# ── 编号归一 ────────────────────────────────────────────────────

def _norm_num(tok: str) -> str:
    """编号 token 归一：去括号/空白、全角数字转半角、分隔符统一为 `-`。

    "３．１" / "3.1" / "3–1" → "3-1"；"51" → "51"（顺序编号保持原样）。
    """
    t = (tok or "").translate(_FW_DIGITS).strip().strip("()（）").strip()
    t = re.sub(r"\s*[-–—\u2010\u2011.．·]\s*", "-", t)
    return t


def _norm_kinds(value) -> list[str]:
    if not value or value == "all":
        return list(_KINDS)
    if isinstance(value, str):
        sel = {value}
    else:
        sel = {k for k in value if k in _KINDS}
    out = [k for k in _KINDS if k in sel]
    return out or list(_KINDS)


# ── 目标收集（题注 + 公式编号） ─────────────────────────────────

def _caption_targets(root) -> list[dict]:
    """题注目标：{kind, para, num_lit, num_norm, display, bm_start, span_end}。

    坐标 = 段落全 run 文本（与 _replace_number_span 的坐标系一致）。
    display = 标签+编号字面（书签覆盖区，REF 域显示值）。
    """
    out = []
    for c in _walk_captions(root):
        p = c["para"]
        tall = _drtext(p)
        sp = _number_span(tall)
        if sp is None or not tall:
            continue
        lead = len(tall) - len(tall.lstrip())
        out.append({
            "kind": c["kind"], "para": p,
            "num_lit": tall[sp[0]:sp[1]].strip(),
            "num_norm": _norm_num(tall[sp[0]:sp[1]]),
            "display": tall[lead:sp[1]].strip(),
            "bm_start": lead, "span_end": sp[1],
        })
    return out


def _equation_targets(path, root, kinds) -> list[dict]:
    """公式编号目标：行间公式段行尾 `(N)` / `(C-N)` 字面（span=括号区）。"""
    if "equation" not in kinds:
        return []
    try:
        styles = _read_styles_map(path)
    except Exception:  # noqa: BLE001
        styles = {}
    out = []
    for p in root.iter(f"{W}p"):
        try:
            role, _ = _equation_para_context(p, styles, _M_NS)
        except Exception:  # noqa: BLE001
            continue
        if role != "equation_display":
            continue
        text = _para_full_text(p)
        mm = _EQ_NUM_ANY_RE.match(text)
        if not mm:
            continue
        s, e = mm.start("num"), mm.end("num")
        lit = text[s:e]
        out.append({"kind": "equation", "para": p, "num_lit": lit.strip(),
                    "num_norm": _norm_num(lit), "display": lit.strip(),
                    "bm_start": s, "span_end": e, "span": (s, e)})
    return out


# ── 扫描 ────────────────────────────────────────────────────────

def _field_chains(p):
    """段落 run 序列里的域链：yield (i, j, instr_joined, result_range_chars)。

    result_range 按**字符级**计算：separate 之后的文本字符数到 end 之前——
    兼容「单 run 承载整条域」（separate/end 同 run 时必须逐子元素跟踪字符位置，
    2026-10-08 幂等缺陷根因：run 级算法对同 run 域产生反转区间 → 已包裹的引用
    被二次扫描重复替换）。
    """
    runs = [c for c in p if etree.QName(c).localname == "r"]

    def _rlen(r):
        return _run_char_len(r)

    cum = 0
    i = 0
    while i < len(runs):
        if not any((fc.get(f"{W}fldCharType") or "") == "begin"
                   for fc in runs[i].iter(f"{W}fldChar")):
            cum += _rlen(runs[i])
            i += 1
            continue
        depth = 0
        pos = cum  # run i 的字符起点
        sep_pos = None
        instr: list[str] = []
        j = i
        found = False
        while j < len(runs):
            for ch in runs[j]:
                tag = etree.QName(ch).localname
                if tag == "fldChar":
                    k = ch.get(f"{W}fldCharType") or ""
                    if k == "begin":
                        depth += 1
                    elif k == "end":
                        depth -= 1
                    elif k == "separate" and depth == 1 and sep_pos is None:
                        sep_pos = pos
                elif tag == "instrText":
                    instr.append(ch.text or "")
                elif tag == "t":
                    pos += len(ch.text or "")
                elif tag == "noBreakHyphen":
                    pos += 1
            if depth <= 0:
                found = True
                break
            j += 1
        if found:
            rr = (sep_pos, pos) if sep_pos is not None else None
            yield i, j, "".join(instr), rr
        else:
            # 未闭合：防御——继续扫描后续 run
            pass
        cum += sum(_rlen(r) for r in runs[i:j + 1])
        i = j + 1


def _ref_covered_ranges(p) -> list[tuple[int, int]]:
    """段落内 REF 域结果文本的字符区间（幂等：已包裹的引用不再扫）。"""
    ranges = []
    for _i, _j, instr, rr in _field_chains(p):
        if rr and _REF_INSTR_RE.search(instr):
            ranges.append(rr)
    return ranges


def _scan_references(root, targets: list[dict], kinds: list[str]):
    """扫描正文引用 → (matches, unmatched, targets_by)。"""
    by = {k: {} for k in _KINDS}
    dup = set()
    for t in targets:
        kind, norm = t["kind"], t["num_norm"]
        if norm in by[kind]:        # 注意：字典键是编号字符串，不是 (kind, num) 元组
            dup.add((kind, norm))
        else:
            by[kind][norm] = t
    for kind, norm in dup:
        by[kind].pop(norm, None)    # 重号=歧义 → 不猜

    body = root.find(f"{W}body")
    if body is None:
        return [], [], by
    cover = _cover_paragraphs(root, None)
    target_paras = {id(t["para"]) for t in targets}
    matches: list[dict] = []
    unmatched: list[dict] = []

    def _covered(a: int, b: int, ranges) -> bool:
        return any(a < cb and b > ca for ca, cb in ranges)

    for p in body.iter(f"{W}p"):
        if cover and any(p is c for c in cover):
            continue
        if _in_textbox(p) or _has_pageref(p):
            continue
        if id(p) in target_paras:
            continue  # 题注/公式段自身不参与引用扫描
        text = _drtext(p)
        if not text or not any(ch in text for ch in "图表式"):
            continue
        ranges = _ref_covered_ranges(p)
        if "figure" in kinds or "table" in kinds:
            for m in _REF_TOKEN_RE.finditer(text):
                a, b = m.span()
                if a > 0 and text[a - 1] in "图表":
                    continue  # "图表3" 这类连写不拆
                if _covered(a, b, ranges):
                    continue
                head = m.group(0).lstrip()[:1]
                kind = "figure" if head == "图" else "table"
                if kind not in kinds:
                    continue
                lab = re.match(r"[图表]\s*", m.group(0))
                num_lit = m.group(0)[lab.end():]
                tgt = by[kind].get(_norm_num(num_lit))
                if tgt is None:
                    reason = ("重号歧义（同编号多题注）"
                              if (kind, _norm_num(num_lit)) in dup
                              else "未找到对应编号")
                    unmatched.append({"match": m.group(0), "reason": reason,
                                      "snippet": text[max(0, a - 12):b + 12]})
                    continue
                matches.append({"kind": kind, "para": p, "span": (a, b),
                                "match": m.group(0), "target": tgt,
                                "snippet": text[max(0, a - 12):b + 12]})
        if "equation" in kinds:
            for m in _EQ_TOKEN_RE.finditer(text):
                a, b = m.span("num")
                if _covered(a, b, ranges):
                    continue
                num_lit = m.group("num")
                tgt = by["equation"].get(_norm_num(num_lit))
                if tgt is None:
                    reason = ("重号歧义（同编号多公式）"
                              if ("equation", _norm_num(num_lit)) in dup
                              else "未找到对应编号")
                    unmatched.append({"match": m.group(0).strip(),
                                      "reason": reason,
                                      "snippet": text[max(0, a - 12):b + 12]})
                    continue
                matches.append({"kind": "equation", "para": p, "span": (a, b),
                                "match": m.group(0).strip(), "target": tgt,
                                "snippet": text[max(0, a - 12):b + 12]})
    return matches, unmatched, by


# ── run 边界与书签落位 ──────────────────────────────────────────

def _run_len(r) -> int:
    return _run_char_len(r)


def _split_run(r, off: int):
    """把 run 在自身文本偏移 off 处拆成两个（rPr 复制）。

    返回 (left_run, right_run)；off 必须 >0 且 <run 文本总长。
    """
    parent = r.getparent()
    idx = list(parent).index(r)
    rpr = r.find(f"{W}rPr")

    def _mk():
        nr = etree.Element(f"{W}r")
        if rpr is not None:
            nr.append(copy.deepcopy(rpr))
        return nr

    left, right = _mk(), _mk()
    acc = 0
    for ch in r:
        if ch is rpr:
            continue
        tag = etree.QName(ch).localname
        if tag == "t":
            tx = ch.text or ""
            take = max(0, min(len(tx), off - acc))
            if take > 0:
                t = etree.SubElement(left, f"{W}t")
                t.set(_XML_SPACE, "preserve")
                t.text = tx[:take]
            if len(tx) - take > 0:
                t = etree.SubElement(right, f"{W}t")
                t.set(_XML_SPACE, "preserve")
                t.text = tx[take:]
            acc += len(tx)
        elif tag == "noBreakHyphen":
            # NBH 计 1 字符（与 _run_char_len 坐标一致）
            (left if acc < off else right).append(copy.deepcopy(ch))
            acc += 1
        else:
            (left if acc < off else right).append(copy.deepcopy(ch))
    parent.remove(r)
    parent.insert(idx, right)
    parent.insert(idx, left)
    return left, right


def _ensure_char_boundary(p, pos: int) -> None:
    """保证段落全 run 文本坐标系中存在 run 边界于 pos（必要时拆 run）。"""
    if pos <= 0:
        return
    runs = [c for c in p if etree.QName(c).localname == "r"]
    cum = 0
    for r in runs:
        n = _run_len(r)
        if cum < pos < cum + n:
            _split_run(r, pos - cum)
            return
        cum += n


def _insert_point(p, pos: int) -> int:
    """插入下标：位于字符位置 pos 处的元素间隙（在含该位置的 run 之间）。"""
    children = list(p)
    cum = 0
    for i, el in enumerate(children):
        if etree.QName(el).localname != "r":
            continue
        n = _run_len(el)
        if pos <= 0:
            return i
        if cum + n == pos:
            # 跳过后续零字符 run，落点在最后一个零字符 run 之后
            return i + 1
        cum += n
    return len(children)


def _place_ref_bookmark(tgt: dict, name: str, bid: int) -> None:
    """在目标区 [bm_start, span_end) 落书签（start/end 精确覆盖标签+编号）。"""
    p = tgt["para"]
    _ensure_char_boundary(p, tgt["bm_start"])
    _ensure_char_boundary(p, tgt["span_end"])
    be = etree.Element(f"{W}bookmarkEnd")
    be.set(f"{W}id", str(bid))
    p.insert(_insert_point(p, tgt["span_end"]), be)
    bs = etree.Element(f"{W}bookmarkStart")
    bs.set(f"{W}id", str(bid))
    bs.set(f"{W}name", name)
    p.insert(_insert_point(p, tgt["bm_start"]), bs)


def _prune_ref_bookmarks(root, kinds):
    """清理我们命名空间的 `_RefF/T/E` 书签——**在用者保留**。

    返回 (removed, kept_para)：removed=清掉的陈旧书签数；kept_para={name: para}
    （仍被 doc 内 REF 域引用的书签 → 保留原样，供目标复用，避免混态二次运行
    把旧引用指向断链）。
    """
    prefixes = tuple(_REF_PREFIX[k] for k in _KINDS if k in set(kinds))

    def _in_scope(name: str) -> bool:
        return bool(name) and name.startswith(prefixes)

    used: set[str] = set()
    for p in root.iter(f"{W}p"):
        instr = _instr_text(p)
        if "REF" not in instr:
            continue
        for mm in _OUR_REF_RE.finditer(instr):
            used.add(mm.group(1))

    removed = 0
    kept_para: dict[str, object] = {}
    for bs in list(root.iter(f"{W}bookmarkStart")):
        name = bs.get(f"{W}name") or ""
        if not _in_scope(name):
            continue
        if name in used:
            kept_para[name] = bs.getparent()
            continue
        bid = bs.get(f"{W}id")
        if bs.getparent() is not None:
            bs.getparent().remove(bs)
        for be in list(root.iter(f"{W}bookmarkEnd")):
            if be.get(f"{W}id") == bid:
                be.getparent().remove(be)
        removed += 1
    return removed, kept_para


# ── 替换为 REF 域 ───────────────────────────────────────────────

def _placeholder_to_field(p, ph: str, bm: str, cached: str,
                          vert_align: str | None = None) -> bool:
    """把段落里的占位符文本替换为 REF 域（单 run 承载：begin/instr/sep/缓存/end）。

    域指令固定带 `\\* CHARFORMAT`（2026-10-12 实测事故修复）：REF 域结果的默认
    格式采用**被引书签文本的字符格式**——章号引用指向「第 N 章」标题书签，Word
    排版/导出时正文提及被渲染成标题的黑体加粗字号（用户报「第 x 章都被加上了域，
    在 pdf 中被改了字体和字号」）。Word 渲染对照实验：无该开关 → SimHei 18pt +
    TNR-Bold 18pt；加该开关 → SimSun 12pt（保持提及自身格式）。
    vert_align="superscript" → 给承载域的 run 打上标（文献引用 [n] 用）。
    """
    for r in [c for c in p if etree.QName(c).localname == "r"]:
        acc = 0
        for t in r.iter(f"{W}t"):
            tx = t.text or ""
            i = tx.find(ph)
            if i < 0:
                acc += len(tx)
                continue
            off = acc + i
            # 拆出占位符专属 run
            if off > 0:
                _l, rr = _split_run(r, off)
            else:
                rr = r
            if len(ph) < _run_len(rr):
                _mid, _r2 = _split_run(rr, len(ph))
            else:
                _mid = rr
            # _mid 恰为占位符 run：清文本、装域链
            for x in list(_mid.iter(f"{W}t")):
                x.getparent().remove(x)
            for el in ("begin", None, "separate", "text", "end"):
                if el is None:
                    ins = etree.SubElement(_mid, f"{W}instrText")
                    ins.set(_XML_SPACE, "preserve")
                    ins.text = f" REF {bm} \\h \\* CHARFORMAT "
                elif el == "text":
                    tt = etree.SubElement(_mid, f"{W}t")
                    tt.set(_XML_SPACE, "preserve")
                    tt.text = cached
                else:
                    fc = etree.SubElement(_mid, f"{W}fldChar")
                    fc.set(f"{W}fldCharType", el)
            if vert_align:
                _set_vert_align(_mid, vert_align)
            return True
    return False


def _superscript_marker(p, start: int, end: int) -> int:
    """把 [start, end) 及紧邻的引用括号整段打上标（引用标记 [n]/【n】 → 上标）。

    坐标 = `_drtext(p)`（与 _cite_scan 的 span 一致）。只在 run 边界做手术
    （拆 run + 写 vertAlign），不改任何文本 → 后续 _replace_number_span 的
    字符坐标不受影响。返回打上标的 run 数。
    """
    text = _drtext(p)
    a, b = start, end
    if a - 1 >= 0 and text[a - 1] in "[【":
        a -= 1
    if b < len(text) and text[b] in "]】":
        b += 1
    _ensure_char_boundary(p, a)
    _ensure_char_boundary(p, b)
    marked = 0
    cum = 0
    for r in [c for c in p if etree.QName(c).localname == "r"]:
        n = _run_len(r)
        if n == 0:
            continue
        r0, r1 = cum, cum + n
        if r1 > a and r0 < b and _set_vert_align(r, "superscript"):
            marked += 1
        cum = r1
    return marked


# ── 动作入口 ────────────────────────────────────────────────────

def generate_cross_references(path, params: dict | None = None) -> dict:
    """正文交叉引用 → REF 域（指向题注/公式书签；编号改动后 F9 联动）。

    params:
      kinds: "all"|"figure"|"table"|"equation"（默认 all）
      apply: false（默认，只出待替换清单）/ true（落盘）

    只对已存在的编号生成；清单先给用户确认；不发明引用。
    """
    path = Path(path)
    params = params or {}
    kinds = _norm_kinds(params.get("kinds"))
    apply_ = bool(params.get("apply"))
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False, "applied": False, "reason": "no_body"}

    targets = _caption_targets(root) + _equation_targets(path, root, kinds)
    matches, unmatched, _by = _scan_references(root, targets, kinds)
    if not matches:
        return {"changed": False, "applied": False, "reason": "no_references",
                "targets": len(targets), "unmatched": unmatched[:20],
                "message": "正文未发现与现有编号匹配的引用（图/表/式），无待替换项。"}

    totals = {"figure": 0, "table": 0, "equation": 0}
    for m in matches:
        totals[m["kind"]] = totals.get(m["kind"], 0) + 1

    if not apply_:
        proposed = [{"kind": m["kind"], "find": m["match"],
                     "will_display": m["target"]["display"],
                     "snippet": m["snippet"]} for m in matches[:80]]
        return {"changed": False, "applied": False, "proposed": proposed,
                "unmatched": unmatched[:20], "totals": totals,
                "message": "仅出待替换清单（未落盘）；确认后 apply=true 替换为"
                           "交叉引用域（编号改动后 Word 中 F9 联动更新）。"}

    removed_stale, kept_para = _prune_ref_bookmarks(root, kinds)
    max_bid = _max_bookmark_id(root)
    names = _bookmark_names(root)

    # 被引用目标 → 书签（按文档序命名；有在用旧书签落同段的则复用）
    doc_order = {id(el): i for i, el in enumerate(body.iter())}
    referenced: dict[tuple, dict] = {}
    for m in matches:
        referenced[(m["kind"], m["target"]["num_norm"])] = m["target"]
    bm_map: dict[tuple, str] = {}
    added = 0
    consumed: set[str] = set()
    for kind in _KINDS:
        n = 0
        keys = [k for k in referenced if k[0] == kind]
        keys.sort(key=lambda k: doc_order.get(id(referenced[k]["para"]), 0))
        for key in keys:
            tgt = referenced[key]
            reused = None
            for name, para in kept_para.items():
                if name in consumed or not name.startswith(_REF_PREFIX[kind]):
                    continue
                if para is tgt["para"]:
                    reused = name
                    break
            if reused is not None:
                consumed.add(reused)
                bm_map[key] = reused
                continue
            n += 1
            name = f"{_REF_PREFIX[kind]}{n:04d}"
            while name in names:
                n += 10000
                name = f"{_REF_PREFIX[kind]}{n:04d}"
            max_bid += 1
            _place_ref_bookmark(tgt, name, max_bid)
            names.add(name)
            bm_map[key] = name
            added += 1

    # 逐段替换（同段从右往左，保持字符坐标有效）
    by_para: dict[int, list] = {}
    for m in matches:
        by_para.setdefault(id(m["para"]), []).append(m)
    para_of = {id(m["para"]): m["para"] for m in matches}
    changed_items: list[dict] = []
    failed: list[dict] = []
    uid = 0
    for pid, ms in by_para.items():
        p = para_of[pid]
        for m in sorted(ms, key=lambda x: x["span"][0], reverse=True):
            uid += 1
            ph = f"{_PH_OPEN}REF{uid}{_PH_CLOSE}"
            bm = bm_map[(m["kind"], m["target"]["num_norm"])]
            cached = m["target"]["display"]
            err = _replace_number_span(p, m["span"][0], m["span"][1], ph)
            if err:
                failed.append({"match": m["match"], "error": err})
                continue
            if not _placeholder_to_field(p, ph, bm, cached):
                failed.append({"match": m["match"], "error": "field_insert_failed"})
                continue
            changed_items.append({"kind": m["kind"], "find": m["match"],
                                  "display": cached, "bookmark": bm,
                                  "snippet": m["snippet"]})

    if changed_items:
        _write_part(path, "word/document.xml", root)
    result = {"changed": len(changed_items) > 0, "applied": True,
              "changes": changed_items, "failed": failed,
              "unmatched": unmatched[:20], "totals": totals,
              "bookmarks_added": added, "stale_bookmarks_removed": removed_stale}
    try:
        result["verified"] = verify_cross_references(path, params)
    except Exception:  # noqa: BLE001
        result["verified"] = None
    return result


def verify_cross_references(path, params: dict | None = None) -> bool:
    """复查：我们建立的 REF 域（_RefF/T/E）书签全部存在。

    apply=false 是提案阶段（未落盘），恒真（对齐 convert_text_equations 约定）。
    不做「缓存==书签内容」比较——A4 后续改编号后缓存待 F9 刷新属正常状态。
    """
    try:
        params = params or {}
        if not params.get("apply"):
            return True
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        names = _bookmark_names(root)
        for p in root.iter(f"{W}p"):
            instr = _instr_text(p)
            if "REF" not in instr:
                continue
            for mm in _OUR_REF_RE.finditer(instr):
                if mm.group(1) not in names:
                    return False
        return True
    except Exception:  # noqa: BLE001
        return False


# ═══════════════════════════════════════════════════════════════════
# 文献引用联动（2026-10-08 用户拍板：正文 [n] ↔ 文末条目 → 交叉引用域，
# 提案→确认）——与图/表/式交叉引用同机制：条目编号加书签（`_RefCiteNNNN`，
# 精确覆盖数字），正文 [n] 的数字替换为 REF 域（缓存=原编号字面）。条目
# 编号改动后 F9 → 域从书签内容重算 → 引用跟随。本版只处理**单个 [n]**；
# [1,2]/[3-5] 组合与范围引用如实列入 unmatched 不动原文（不发明编号）。
# ═══════════════════════════════════════════════════════════════════

_CITE_PREFIX = "_RefCite"
_CITE_ITEM_RE = re.compile(r"^[\[【]\s*(\d{1,3})\s*[\]】]")
_CITE_TOKEN_RE = re.compile(r"[\[【]\s*(\d{1,3})\s*[\]】]")
_CITE_MULTI_RE = re.compile(
    r"[\[【]\s*\d{1,3}(?:\s*[,，、]\s*\d{1,3})+\s*[\]】]"   # [1,2] / 【1，2】
    r"|[\[【]\s*\d{1,3}\s*[-–—]\s*\d{1,3}\s*[\]】]")        # [3-5]
_OUR_CITE_REF_RE = re.compile(r"(?:(?<=\s)|^)REF\s+(_RefCite\d+)")
_REFS_HEAD_TXT = {"参考文献", "references", "reference"}


def _norm_ws2(s: str) -> str:
    return re.sub(r"[\s\u3000\u25a1]+", "", s or "")


def _is_refs_heading(text: str) -> bool:
    n = _norm_ws2(text)
    return bool(n) and len(n) <= 20 and n.lower() in _REFS_HEAD_TXT


def _cite_targets(root) -> list[dict]:
    """参考文献条目目标：{para, num_lit, num_norm, display, bm_start, span_end}。

    书签精确覆盖条目开头的**数字本身**（`[1]` 的 `1`，不含方括号）——正文侧
    也只替换数字，方括号留在原位；域显示 = 数字，视觉等于原样「[1]」。
    """
    body = root.find(f"{W}body")
    out: list[dict] = []
    if body is None:
        return out
    in_refs = False
    for p in body.iter(f"{W}p"):
        t = _drtext(p)
        if not in_refs:
            if _is_refs_heading(t):
                in_refs = True
            continue
        if _has_pageref(p) or _in_textbox(p):
            continue
        m = _CITE_ITEM_RE.match(t)
        if not m:
            continue
        s, e = m.span(1)
        out.append({"kind": "citation", "para": p, "num_lit": m.group(1),
                    "num_norm": m.group(1), "display": m.group(1),
                    "bm_start": s, "span_end": e})
    return out


def _cite_scan(root, targets: list[dict]):
    """扫描正文（参考文献区之前）的单号 [n] 引用 + 组合/范围引用清单。"""
    by: dict[str, dict] = {}
    dup: set[str] = set()
    for t in targets:
        n = t["num_norm"]
        if n in by:
            dup.add(n)
        else:
            by[n] = t
    for n in dup:
        by.pop(n, None)              # 重号=歧义 → 不猜

    body = root.find(f"{W}body")
    matches: list[dict] = []
    unmatched: list[dict] = []
    if body is None:
        return matches, unmatched, by
    cover = _cover_paragraphs(root, None)
    target_paras = {id(t["para"]) for t in targets}
    for p in body.iter(f"{W}p"):
        t = _drtext(p)
        if not t:
            continue
        if _is_refs_heading(t):
            break                    # 参考文献区起，列表自身不参与扫描
        if cover and any(p is c for c in cover):
            continue
        if _in_textbox(p) or _has_pageref(p):
            continue
        if id(p) in target_paras:
            continue
        if "[" not in t and "【" not in t:
            continue
        ranges = _ref_covered_ranges(p)
        multi_spans: list[tuple[int, int]] = []
        for mm in _CITE_MULTI_RE.finditer(t):
            a, b = mm.span()
            if any(a < cb and b > ca for ca, cb in ranges):
                continue
            multi_spans.append((a, b))
            unmatched.append({"match": mm.group(0).strip(),
                              "reason": "组合/范围引用（本版只改单个 [n]，此项未动）",
                              "snippet": t[max(0, a - 12):b + 12]})
        for mm in _CITE_TOKEN_RE.finditer(t):
            a, b = mm.span()
            if any(a < cb and b > ca for ca, cb in ranges):
                continue             # 已是域（幂等；含用户自己的引用）
            if any(a < cb and b > ca for ca, cb in multi_spans):
                continue             # 已被组合/范围上报，避免重复列
            num = _cn_num_str(mm.group(1))
            s, e = mm.span(1)
            tgt = by.get(num)
            if tgt is None:
                reason = ("重号歧义（同编号多条目）" if num in dup
                          else "未找到对应条目")
                unmatched.append({"match": mm.group(0).strip(), "reason": reason,
                                  "snippet": t[max(0, a - 12):b + 12]})
                continue
            matches.append({"kind": "citation", "para": p, "span": (s, e),
                            "match": mm.group(0).strip(), "target": tgt,
                            "num": num, "snippet": t[max(0, a - 12):b + 12]})
    return matches, unmatched, by


def _prune_cite_bookmarks(root):
    """清理 `_RefCite` 陈旧书签（在用者保留）——与 `_prune_ref_bookmarks` 同口径。"""
    used: set[str] = set()
    for p in root.iter(f"{W}p"):
        instr = _instr_text(p)
        if "REF" not in instr:
            continue
        for mm in _OUR_CITE_REF_RE.finditer(instr):
            used.add(mm.group(1))
    removed = 0
    kept_para: dict[str, object] = {}
    for bs in list(root.iter(f"{W}bookmarkStart")):
        name = bs.get(f"{W}name") or ""
        if not name.startswith(_CITE_PREFIX):
            continue
        if name in used:
            kept_para[name] = bs.getparent()
            continue
        bid = bs.get(f"{W}id")
        if bs.getparent() is not None:
            bs.getparent().remove(bs)
        for be in list(root.iter(f"{W}bookmarkEnd")):
            if be.get(f"{W}id") == bid:
                be.getparent().remove(be)
        removed += 1
    return removed, kept_para


def scan_citation_candidates(path) -> dict:
    """只读扫描：正文 [n] 引用与文末条目的可联动数量（交付成功后的提示用）。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        targets = _cite_targets(root)
        matches, _un, _by = _cite_scan(root, targets)
        return {"items": len(targets), "citations": len(matches)}
    except Exception:  # noqa: BLE001
        return {"items": 0, "citations": 0}


def link_citation_references(path, params: dict | None = None) -> dict:
    """正文 [n] ↔ 参考文献条目 → 交叉引用域（REF；编号改动后 F9 联动）。

    params:
      apply: false（默认，只出待替换清单）/ true（落盘）
      superscript: true/false（可选，默认 **true**）——引用标记 [n] 是否做成上标。
        Chinese 论文的正文引用约定即为上标（2026-10-12 用户口径：「参考文献的
        交叉引用都没有弄成上标」按缺陷报的），故默认上标；要保留正文形态传 false；
        原文标记本来就是上标时无论传什么都会保持。

    只对已存在的条目编号生成；找不到条目/组合/范围引用只报不改；不发明编号。
    域指令带 \\* CHARFORMAT：F9 更新后引用标记保持自身格式，不被条目文本格式
    （宋体小四/五号混排）带跑。
    """
    path = Path(path)
    params = params or {}
    apply_ = bool(params.get("apply"))
    _sup_raw = params.get("superscript")
    sup_ = True if _sup_raw is None else bool(_sup_raw)
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False, "applied": False, "reason": "no_body"}
    targets = _cite_targets(root)
    matches, unmatched, _by = _cite_scan(root, targets)
    if not targets:
        return {"changed": False, "applied": False, "reason": "no_reference_items",
                "unmatched": unmatched[:20],
                "message": "文末未发现「[1] …」式参考文献列表，无待联动项。"}
    if not matches:
        return {"changed": False, "applied": False, "reason": "no_citations",
                "target_items": len(targets), "unmatched": unmatched[:20],
                "message": "正文未发现与文献条目对应的单个 [n] 引用，无待替换项。"
                           "（组合/范围引用等不支持项见列表，均未改动原文。）"}

    totals = {"citations": len(matches)}
    if not apply_:
        proposed = [{"find": m["match"], "num": m["num"],
                     "item": _drtext(m["target"]["para"])[:60],
                     "snippet": m["snippet"]} for m in matches[:80]]
        return {"changed": False, "applied": False, "proposed": proposed,
                "unmatched": unmatched[:20], "totals": totals,
                "message": "仅出待替换清单（未落盘）；确认后 apply=true 把正文 [n] "
                           "替换为交叉引用域（编号改动后 Word 中 F9 联动更新）。"}

    removed_stale, kept_para = _prune_cite_bookmarks(root)
    max_bid = _max_bookmark_id(root)
    names = _bookmark_names(root)
    body_order = {id(el): i for i, el in enumerate(body.iter())}
    used_targets: dict[int, dict] = {}
    for m in matches:
        used_targets.setdefault(id(m["target"]["para"]), m["target"])
    bm_map: dict[int, str] = {}
    added = 0
    consumed: set[str] = set()
    n = 0
    for pid in sorted(used_targets, key=lambda k: body_order.get(k, 0)):
        tgt = used_targets[pid]
        reused = None
        for name, para in kept_para.items():
            if name in consumed or para is not tgt["para"]:
                continue
            reused = name
            break
        if reused is not None:
            consumed.add(reused)
            bm_map[pid] = reused
            continue
        n += 1
        name = f"{_CITE_PREFIX}{n:04d}"
        while name in names:
            n += 10000
            name = f"{_CITE_PREFIX}{n:04d}"
        max_bid += 1
        _place_ref_bookmark(tgt, name, max_bid)
        names.add(name)
        bm_map[pid] = name
        added += 1

    by_para: dict[int, list] = {}
    for m in matches:
        by_para.setdefault(id(m["para"]), []).append(m)
    para_of = {id(m["para"]): m["para"] for m in matches}
    changed_items: list[dict] = []
    failed: list[dict] = []
    uid = 0
    for pid, ms in by_para.items():
        p = para_of[pid]
        for m in sorted(ms, key=lambda x: x["span"][0], reverse=True):
            uid += 1
            ph = f"{_PH_OPEN}CITE{uid}{_PH_CLOSE}"
            bm = bm_map[id(m["target"]["para"])]
            # 上标（默认口径）：先给标记 [n] 打上标再换域——只拆 run 不改文本，
            # 坐标不受影响；承载域的 run 继承该上标（F9 后靠 \* CHARFORMAT 保持）。
            if sup_:
                _superscript_marker(p, m["span"][0], m["span"][1])
            err = _replace_number_span(p, m["span"][0], m["span"][1], ph)
            if err:
                failed.append({"match": m["match"], "error": err})
                continue
            if not _placeholder_to_field(p, ph, bm, m["num"],
                                         vert_align="superscript" if sup_ else None):
                failed.append({"match": m["match"], "error": "field_insert_failed"})
                continue
            changed_items.append({"find": m["match"], "display": m["num"],
                                  "bookmark": bm, "snippet": m["snippet"]})

    if changed_items:
        _write_part(path, "word/document.xml", root)
    result = {"changed": len(changed_items) > 0, "applied": True,
              "changes": changed_items, "failed": failed,
              "unmatched": unmatched[:20], "totals": totals,
              "bookmarks_added": added, "stale_bookmarks_removed": removed_stale}
    try:
        result["verified"] = verify_citation_links(path, params)
    except Exception:  # noqa: BLE001
        result["verified"] = None
    return result


def verify_citation_links(path, params: dict | None = None) -> bool:
    """复查：有对应条目的正文单号 [n] 全部落在 REF 域结果内（无裸编号漏改）。

    apply=false 是提案阶段（未落盘），恒真（对齐既有约定）。
    """
    try:
        params = params or {}
        if not params.get("apply"):
            return True
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        targets = _cite_targets(root)
        by = {t["num_norm"] for t in targets}
        if not by:
            return True
        body = root.find(f"{W}body")
        if body is None:
            return True
        cover = _cover_paragraphs(root, None)
        for p in body.iter(f"{W}p"):
            t = _drtext(p)
            if not t:
                continue
            if _is_refs_heading(t):
                break
            if cover and any(p is c for c in cover):
                continue
            if _in_textbox(p) or _has_pageref(p):
                continue
            ranges = _ref_covered_ranges(p)
            for mm in _CITE_TOKEN_RE.finditer(t):
                if mm.group(1) not in by:
                    continue
                a, b = mm.span(1)
                if not any(a < cb and b > ca for ca, cb in ranges):
                    return False
        return True
    except Exception:  # noqa: BLE001
        return False


# ═══════════════════════════════════════════════════════════════════
# 章号引用联动（2026-10-08 用户线「序号」：正文「第 X 章」提及 → 交叉引用域，
# 提案→确认）——与文献联动同机制：章标题里的「第 N 章」字面加书签
# （`_RefChapNNNN`，精确覆盖「第 N 章」），正文提及整段替换为 REF 域
# （缓存=原字面，`\h` 可跳转）。章号改动后 F9 → 域重算 → 提及跟随。
# 守卫：目录条目（PAGEREF / 点 leader）/文本框/封面区/章标题段自身不参与；
# 只改阿拉伯数字章号；找不到对应章标题的提及如实列出、不动原文；编号不发明。
# ═══════════════════════════════════════════════════════════════════

_CHAP_PREFIX = "_RefChap"
_CN_NUM_PAT = r"[0-9]{1,2}|[零〇一二两三四五六七八九十]{1,3}"
_CHAP_HEAD_RE = re.compile(rf"\s*(第\s*({_CN_NUM_PAT})\s*章)")
_CHAP_TOKEN_RE = re.compile(rf"第\s*({_CN_NUM_PAT})\s*章")
_CN_DIGITS = {"零": 0, "〇": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4,
              "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}


def _cn_num_str(tok: str) -> str:
    """章号归一为阿拉伯数字字符串：「一」→1、「十一」→11、「二十三」→23、'3'→3。

    中文论文大量使用「第一章」式中文数字章号，此前只认阿拉伯数字，
    导致整篇零匹配（2026-10-21 实测）。
    """
    tok = (tok or "").strip()
    if not tok:
        return tok
    if tok.isdigit():
        return str(int(tok))
    if "十" in tok:
        left, _, right = tok.partition("十")
        tens = _CN_DIGITS.get(left, 1) if left else 1
        ones = _CN_DIGITS.get(right, 0) if right else 0
        return str(tens * 10 + ones)
    if all(ch in _CN_DIGITS for ch in tok):
        return str(int("".join(str(_CN_DIGITS[ch]) for ch in tok)))
    return tok
_OUR_CHAP_REF_RE = re.compile(r"(?:(?<=\s)|^)REF\s+(_RefChap\d+)")
_TOC_DOTS_RE2 = re.compile(r"[.．]{4,}|…")


def _is_chap_heading_para(p) -> bool:
    """章标题段：outlineLvl 或 Heading 样式（编译链 ensure_heading_outline 已打）。"""
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        return False
    if ppr.find(f"{W}outlineLvl") is not None:
        return True
    ps = ppr.find(f"{W}pStyle")
    return ps is not None and "heading" in (ps.get(f"{W}val") or "").lower()


def _chap_targets(root) -> list[dict]:
    """章标题目标：{kind, para, num_lit, num_norm, display, bm_start, span_end}。

    书签精确覆盖「第 N 章」字面（含内部空格），REF 域显示值=该字面。
    """
    body = root.find(f"{W}body")
    out: list[dict] = []
    if body is None:
        return out
    for p in body.iter(f"{W}p"):
        if _in_textbox(p) or _has_pageref(p):
            continue
        tall = _drtext(p)
        m = _CHAP_HEAD_RE.match(tall)
        if not m:
            continue
        if _TOC_DOTS_RE2.search(tall):
            continue
        core = tall.strip()
        if not (_is_chap_heading_para(p) or (
                len(core) <= 40 and not re.search(r"[。；，、！？：]", core))):
            continue
        tok = m.group(1)
        # tok 是「第一章」全字面，取其中的数字部分再归一
        num = _cn_num_str(re.sub(r"[第\s章]", "", tok))
        out.append({"kind": "chapter", "para": p, "num_lit": tok,
                    "num_norm": num, "display": tok.strip(),
                    "bm_start": m.start(1), "span_end": m.end(1)})
    return out


def _chap_scan(root, targets: list[dict]):
    """扫描正文「第 X 章」提及 → (matches, unmatched, by)。章标题段自身不扫。"""
    by: dict[str, dict] = {}
    dup: set[str] = set()
    for t in targets:
        n = t["num_norm"]
        if n in by:
            dup.add(n)
        else:
            by[n] = t
    for n in dup:
        by.pop(n, None)
    body = root.find(f"{W}body")
    matches: list[dict] = []
    unmatched: list[dict] = []
    if body is None:
        return matches, unmatched, by
    cover = _cover_paragraphs(root, None)
    target_paras = {id(t["para"]) for t in targets}
    for p in body.iter(f"{W}p"):
        t = _drtext(p)
        if not t or "第" not in t:
            continue
        if cover and any(p is c for c in cover):
            continue
        if _in_textbox(p) or _has_pageref(p):
            continue
        if id(p) in target_paras:
            continue          # 章标题段自身（书签落点）不参与提及扫描
        if _TOC_DOTS_RE2.search(t):
            continue          # 目录条目（点 leader）
        ranges = _ref_covered_ranges(p)
        for mm in _CHAP_TOKEN_RE.finditer(t):
            a, b = mm.span()
            if any(a < cb and b > ca for ca, cb in ranges):
                continue      # 已是域（幂等；含用户自己的引用）
            num = _cn_num_str(mm.group(1))
            tgt = by.get(num)
            if tgt is None:
                reason = ("重号歧义（同章号多标题）" if num in dup
                          else "未找到对应章标题")
                unmatched.append({"match": mm.group(0).strip(), "reason": reason,
                                  "snippet": t[max(0, a - 12):b + 12]})
                continue
            matches.append({"kind": "chapter", "para": p, "span": (a, b),
                            "match": mm.group(0).strip(), "target": tgt,
                            "num": num, "snippet": t[max(0, a - 12):b + 12]})
    return matches, unmatched, by


def _prune_chap_bookmarks(root):
    """清理 `_RefChap` 陈旧书签（在用者保留）——与 `_prune_cite_bookmarks` 同口径。"""
    used: set[str] = set()
    for p in root.iter(f"{W}p"):
        instr = _instr_text(p)
        if "REF" not in instr:
            continue
        for mm in _OUR_CHAP_REF_RE.finditer(instr):
            used.add(mm.group(1))
    removed = 0
    kept_para: dict[str, object] = {}
    for bs in list(root.iter(f"{W}bookmarkStart")):
        name = bs.get(f"{W}name") or ""
        if not name.startswith(_CHAP_PREFIX):
            continue
        if name in used:
            kept_para[name] = bs.getparent()
            continue
        bid = bs.get(f"{W}id")
        if bs.getparent() is not None:
            bs.getparent().remove(bs)
        for be in list(root.iter(f"{W}bookmarkEnd")):
            if be.get(f"{W}id") == bid:
                be.getparent().remove(be)
        removed += 1
    return removed, kept_para


def scan_chapter_candidates(path) -> dict:
    """只读扫描：章标题数与正文「第 X 章」可联动提及数（交付成功后的提示用）。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        targets = _chap_targets(root)
        matches, _un, _by = _chap_scan(root, targets)
        return {"headings": len(targets), "mentions": len(matches)}
    except Exception:  # noqa: BLE001
        return {"headings": 0, "mentions": 0}


def link_chapter_references(path, params: dict | None = None) -> dict:
    """正文「第 X 章」提及 ↔ 章标题 → 交叉引用域（REF；章号改动后 F9 联动）。

    params:
      apply: false（默认，只出待替换清单）/ true（落盘）

    只对已存在的章标题生成；找不到章标题的提及只报不改；不发明编号。
    """
    path = Path(path)
    params = params or {}
    apply_ = bool(params.get("apply"))
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False, "applied": False, "reason": "no_body"}
    targets = _chap_targets(root)
    matches, unmatched, _by = _chap_scan(root, targets)
    if not targets:
        return {"changed": False, "applied": False, "reason": "no_chapter_headings",
                "unmatched": unmatched[:20],
                "message": "未发现「第 N 章」式章标题，无待联动项。"}
    if not matches:
        return {"changed": False, "applied": False, "reason": "no_mentions",
                "target_headings": len(targets), "unmatched": unmatched[:20],
                "message": "正文未发现与章标题对应的「第 X 章」提及，无待替换项。"
                           "（找不到对应章标题的提及见列表，均未改动原文。）"}

    totals = {"mentions": len(matches)}
    if not apply_:
        proposed = [{"find": m["match"], "num": m["num"],
                     "heading": _drtext(m["target"]["para"])[:60],
                     "snippet": m["snippet"]} for m in matches[:80]]
        return {"changed": False, "applied": False, "proposed": proposed,
                "unmatched": unmatched[:20], "totals": totals,
                "message": "仅出待替换清单（未落盘）；确认后 apply=true 把正文"
                           "「第 X 章」替换为交叉引用域（章号改动后 Word 中 F9 联动）。"}

    removed_stale, kept_para = _prune_chap_bookmarks(root)
    max_bid = _max_bookmark_id(root)
    names = _bookmark_names(root)

    body_order = {id(el): i for i, el in enumerate(body.iter())}
    used_targets: dict[int, dict] = {}
    for m in matches:
        used_targets.setdefault(id(m["target"]["para"]), m["target"])
    bm_map: dict[int, str] = {}
    added = 0
    consumed: set[str] = set()
    n = 0
    for pid in sorted(used_targets, key=lambda k: body_order.get(k, 0)):
        tgt = used_targets[pid]
        reused = None
        for name, para in kept_para.items():
            if name in consumed or para is not tgt["para"]:
                continue
            reused = name
            break
        if reused is not None:
            consumed.add(reused)
            bm_map[pid] = reused
            continue
        n += 1
        name = f"{_CHAP_PREFIX}{n:04d}"
        while name in names:
            n += 10000
            name = f"{_CHAP_PREFIX}{n:04d}"
        max_bid += 1
        _place_ref_bookmark(tgt, name, max_bid)
        names.add(name)
        bm_map[pid] = name
        added += 1

    by_para: dict[int, list] = {}
    for m in matches:
        by_para.setdefault(id(m["para"]), []).append(m)
    para_of = {id(m["para"]): m["para"] for m in matches}
    changed_items: list[dict] = []
    failed: list[dict] = []
    uid = 0
    for pid, ms in by_para.items():
        p = para_of[pid]
        for m in sorted(ms, key=lambda x: x["span"][0], reverse=True):
            uid += 1
            ph = f"{_PH_OPEN}CHAP{uid}{_PH_CLOSE}"
            bm = bm_map[id(m["target"]["para"])]
            err = _replace_number_span(p, m["span"][0], m["span"][1], ph)
            if err:
                failed.append({"match": m["match"], "error": err})
                continue
            if not _placeholder_to_field(p, ph, bm, m["match"]):
                failed.append({"match": m["match"], "error": "field_insert_failed"})
                continue
            changed_items.append({"find": m["match"], "display": m["match"],
                                  "bookmark": bm, "snippet": m["snippet"]})

    if changed_items:
        _write_part(path, "word/document.xml", root)
    result = {"changed": len(changed_items) > 0, "applied": True,
              "changes": changed_items, "failed": failed,
              "unmatched": unmatched[:20], "totals": totals,
              "bookmarks_added": added, "stale_bookmarks_removed": removed_stale}
    try:
        result["verified"] = verify_chapter_links(path, params)
    except Exception:  # noqa: BLE001
        result["verified"] = None
    return result


def verify_chapter_links(path, params: dict | None = None) -> bool:
    """复查：有对应章标题的正文「第 X 章」提及全部落在 REF 域结果内（无裸提及漏改）。

    apply=false 是提案阶段（未落盘），恒真（对齐既有约定）。
    """
    try:
        params = params or {}
        if not params.get("apply"):
            return True
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        targets = _chap_targets(root)
        by = {t["num_norm"] for t in targets}
        if not by:
            return True
        body = root.find(f"{W}body")
        if body is None:
            return True
        cover = _cover_paragraphs(root, None)
        target_paras = {id(t["para"]) for t in targets}
        for p in body.iter(f"{W}p"):
            t = _drtext(p)
            if not t or "第" not in t:
                continue
            if cover and any(p is c for c in cover):
                continue
            if _in_textbox(p) or _has_pageref(p):
                continue
            if id(p) in target_paras:
                continue
            if _TOC_DOTS_RE2.search(t):
                continue
            ranges = _ref_covered_ranges(p)
            for mm in _CHAP_TOKEN_RE.finditer(t):
                if mm.group(1) not in by:
                    continue
                a, b = mm.span()
                if not any(a < cb and b > ca for ca, cb in ranges):
                    return False
        return True
    except Exception:  # noqa: BLE001
        return False

_SEC_HEAD_RE = re.compile(rf"\s*(第\s*({_CN_NUM_PAT})\s*([节条]))")
_SEC_TOKEN_RE = re.compile(rf"第\s*({_CN_NUM_PAT})\s*([节条])")


def link_section_references(path, params: dict | None = None) -> dict:
    """正文「第 N 节 / 第 N 条」提及 ↔ 对应标题 → REF 域（P3，默认 dry-run）。

    与章节引用同口径：只处理已存在的标题，找不到对应标题只报不改，不发明编号。
    合同类文档大量使用「第 N 条」，条文编号变动后 F9 可联动。
    """
    path = Path(path)
    params = params or {}
    apply_ = bool(params.get("apply"))
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False, "applied": False, "reason": "no_body"}

    targets: list[dict] = []
    for p in body.iter(f"{W}p"):
        if _in_textbox(p) or _has_pageref(p):
            continue
        txt = _drtext(p).strip()
        m = _SEC_HEAD_RE.match(txt)
        if not m or len(txt) > 60 or _TOC_DOTS_RE2.search(txt):
            continue
        targets.append({"para": p, "key": f"{_cn_num_str(m.group(2))}{m.group(3)}",
                        "display": m.group(0).strip()})
    if not targets:
        return {"changed": False, "applied": False, "reason": "no_section_headings"}

    by: dict[str, dict] = {}
    for t in targets:
        by.setdefault(t["key"], t)
    target_paras = {id(t["para"]) for t in targets}
    matches, unmatched = [], []
    for p in body.iter(f"{W}p"):
        if id(p) in target_paras or _in_textbox(p) or _has_pageref(p):
            continue
        txt = _drtext(p)
        if "第" not in txt or _TOC_DOTS_RE2.search(txt):
            continue
        ranges = _ref_covered_ranges(p)
        for mm in _SEC_TOKEN_RE.finditer(txt):
            a, b = mm.span()
            if any(a < cb and b > ca for ca, cb in ranges):
                continue
            key = f"{_cn_num_str(mm.group(1))}{mm.group(2)}"
            tgt = by.get(key)
            if tgt is None:
                unmatched.append({"match": mm.group(0), "reason": "未找到对应标题",
                                  "snippet": txt[max(0, a - 12):b + 12]})
                continue
            matches.append({"para": p, "span": (a, b), "match": mm.group(0).strip(),
                            "target": tgt, "key": key})
    if not matches:
        return {"changed": False, "applied": False, "reason": "no_mentions",
                "headings": len(targets), "unmatched": unmatched[:20],
                "message": "未发现与「第 N 节/条」标题对应的提及，无待替换项。"}
    if not apply_:
        return {"changed": False, "applied": False,
                "totals": {"mentions": len(matches)},
                "proposed": [{"find": m["match"], "heading": _drtext(m["target"]["para"])[:40]}
                             for m in matches][:30],
                "unmatched": unmatched[:20],
                "message": "仅出清单（未落盘）；apply=true 落盘后条文编号 F9 联动。"}

    used = {}
    for m in matches:
        used.setdefault(id(m["target"]["para"]), m["target"])
    max_bid = _max_bookmark_id(root)
    names = _bookmark_names(root)
    bm_map, n = {}, 0
    for pid, tgt in used.items():
        n += 1
        name = f"_RefSec{n:04d}"
        while name in names:
            n += 10000
            name = f"_RefSec{n:04d}"
        max_bid += 1
        _place_ref_bookmark(tgt, name, max_bid)
        names.add(name)
        bm_map[pid] = name

    by_para: dict[int, list] = {}
    for m in matches:
        by_para.setdefault(id(m["para"]), []).append(m)
    para_of = {id(m["para"]): m["para"] for m in matches}
    uid, failed = 0, []
    for pid, ms in by_para.items():
        p = para_of[pid]
        for m in sorted(ms, key=lambda x: x["span"][0], reverse=True):
            uid += 1
            ph = f"{_PH_OPEN}SEC{uid}{_PH_CLOSE}"
            bm = bm_map[id(m["target"]["para"])]
            if _replace_number_span(p, m["span"][0], m["span"][1], ph):
                failed.append({"match": m["match"], "error": "replace_failed"})
                continue
            if not _placeholder_to_field(p, ph, bm, m["match"]):
                failed.append({"match": m["match"], "error": "field_insert_failed"})
    if uid and not failed:
        _write_part(path, "word/document.xml", root)
    return {"changed": bool(uid and not failed), "applied": True,
            "totals": {"mentions": len(matches)}, "failed": failed[:10]}
