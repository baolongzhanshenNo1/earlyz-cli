"""题注编号重排（A4）+ 图/表目录生成（A1）——M2，2026-10-08。

Word 行为地基（2026-10-08 Word COM 真机探针实测，D:/Hermes/cache/scratch/eqwork/m2probe）：
- `TOC \\h \\z \\c "图"` 域：Word 更新后自动收集「含 SEQ 图 域」的段落为条目
  （条目文本 + 自动页码，按文档序）；隐藏（vanish run）的域名值不进条目文本。
- `SEQ 图 \\* ARABIC \\s 1` 按章重置实测无效（outlineLvl=0 也不重置）→ 章号由
  引擎自算写为字面文本；题注内嵌一个隐藏 SEQ 采集域即可让 F9 由 Word 原生重建
  出正确条目与页码，且 F9 重算隐藏域不会漂移可见编号（编号是字面的）。

设计边界（对齐 docs/方案-格式检测报告接入与A档七项-2026-10-07.md §A1/A4）：
- A4 renumber_caption_numbers：按章遍历 → 图/表计数器按章重排（图 2-1 式）；
  只改编号值本身，不发明间隔/版式；章前题注、无编号题注一律不动（如实报告）；
  apply=false 只出对照清单（重排前给对照清单），true 落盘。
- A1 generate_caption_toc：题注识别（复用 _is_caption_para 收紧判定 + 引用段守卫）
  → 题注加书签（_TofF/_TofT）+ 隐藏采集域 → 插入「图目录/表目录」区（标题段 +
  条目段=引擎缓存 + `TOC \\c` 域壳）；无题注明确报「未发现题注」，不生成空目录；
  插入位置=主目录条目之后（回退：第一章标题之前）。
- 幂等：重排二次运行 changed=0；目录重生成就地替换同区（不堆叠）。
"""

from __future__ import annotations

import copy
import re
from pathlib import Path

from lxml import etree

from earlyz.actions.new_actions import (
    W,
    _CN_DIGITS,
    _cover_paragraphs,
    _heading_chapter_number,
    _is_caption_para,
    _ppr_insert_ordered,
    _read_part,
    _read_styles_map,
    _write_part,
)

# ── 常量 ────────────────────────────────────────────────────────

_KINDS = ("figure", "table")  # 固定顺序：图目录在前、表目录在后
_KIND_CHAR = {"figure": "图", "table": "表"}
_TOF_TITLE_WORDS = {
    "figure": {"图目录", "插图清单", "图索引"},
    "table": {"表目录", "表格清单", "表索引"},
}
_DEFAULT_HEADINGS = {"figure": "图目录", "table": "表目录"}

# 题注编号 token（前缀 图/表 后的第一段数字串，允许 - . ． · 及非断连字符 U+2010/U+2011 分隔）
_NUMBER_RE = re.compile(r"[0-9]+(?:\s*[\-–—\u2010\u2011\.．·]\s*[0-9]+)*")
# 章标题（纯文本兜底）：第X章/第X篇，要求短段无句读（防正文引用）
_CHAPTER_RE = re.compile(r"^\s*第\s*([0-9]{1,2}|[一二三四五六七八九十]{1,3})\s*[章篇]")
# 正文引用段守卫：图 N…为/是/所示…（「图6-1为美妆…」不是题注）
_REF_GUARD_RE = re.compile(
    r"^(?:图|表)\s*\d[\d\.\-–—·．]*\s*(?:为|是|所|中|和|与|及|至|到|均|在)")
_SENT_PUNCT_RE = re.compile(r"[，。；：]")
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


# ── 基础读取 ────────────────────────────────────────────────────

def _ptxt(p) -> str:
    """段落全文本（含隐藏 run 的 w:t——用于域结构判定）。"""
    return "".join(t.text or "" for t in p.iter(f"{W}t"))


def _run_text(r) -> str:
    """run 可见文本（文档序）：w:t 文本 + noBreakHyphen → "-"。

    2026-10-07 王福勋样本实锤：题注「表5-1」的短横线常是 Word「不换行连字符」
    <w:noBreakHyphen/>（无 w:t）——按 w:t 提取会把 5-1 读成 51，误报「格式不一致」
    并触发错误重排；长度/替换手术坐标系必须与本函数一致。
    """
    out: list[str] = []
    for ch in r.iter():
        tag = etree.QName(ch).localname
        if tag == "t":
            out.append(ch.text or "")
        elif tag == "noBreakHyphen":
            out.append("-")
    return "".join(out)


def _run_char_len(r) -> int:
    """run 字符长度（与 _run_text 坐标系一致：noBreakHyphen 计 1 字符）。"""
    n = 0
    for ch in r.iter():
        tag = etree.QName(ch).localname
        if tag == "t":
            n += len(ch.text or "")
        elif tag == "noBreakHyphen":
            n += 1
    return n


def _visible_txt(p) -> str:
    """段落可见文本：跳过 vanish run（隐藏采集域结果不入条目文本）。"""
    out: list[str] = []
    for r in p.iter(f"{W}r"):
        rpr = r.find(f"{W}rPr")
        if rpr is not None:
            v = rpr.find(f"{W}vanish")
            if v is not None and (v.get(f"{W}val") or "1") not in ("0", "false"):
                continue
        out.append(_run_text(r))
    return "".join(out)


def _in_textbox(p) -> bool:
    anc = p.getparent()
    while anc is not None:
        if etree.QName(anc).localname == "txbxContent":
            return True
        anc = anc.getparent()
    return False


def _instr_text(p) -> str:
    """段落内全部域指令拼接（Word 更新后会把指令拆成多个 instrText）。"""
    return "".join(x.text or "" for x in p.iter(f"{W}instrText"))


def _has_pageref(p) -> bool:
    return "PAGEREF" in _instr_text(p)


def _is_caption_para2(text: str) -> bool:
    """题注段判定 = 引擎既有收紧判定 + 正文引用段守卫。"""
    t = (text or "").strip()
    if not _is_caption_para(t):
        return False
    if _REF_GUARD_RE.match(t):
        return False
    return True


def _kind_of(text: str) -> str:
    return "figure" if text.lstrip().startswith("图") else "table"


def _norm_kinds(value) -> list[str]:
    """kinds 参数归一 → 固定顺序列表（figure, table）。"""
    if not value or value == "all":
        return list(_KINDS)
    if isinstance(value, str):
        sel = {value}
    else:
        sel = {k for k in value if k in _KINDS}
    out = [k for k in _KINDS if k in sel]
    return out or list(_KINDS)


def _walk_captions(root) -> list[dict]:
    """单次遍历：题注段 + 当前章号。返回 [{para, kind, text, chapter}]。

    跳过：封面区 / 文本框内 / 目录条目（含 PAGEREF 的段）。
    章号跟踪：outlineLvl/Heading 样式判定优先，回退短段文本「第X章」。
    """
    body = root.find(f"{W}body")
    if body is None:
        return []
    cover_paras = _cover_paragraphs(root, None)
    chapter = None
    out: list[dict] = []
    for p in body.iter(f"{W}p"):
        if cover_paras and any(p is c for c in cover_paras):
            continue
        if _in_textbox(p):
            continue
        if _has_pageref(p):
            continue
        ch = _chapter_num_of(p)
        if ch:
            chapter = ch
            continue
        text = _visible_txt(p).strip()
        if not _is_caption_para2(text):
            continue
        out.append({"para": p, "kind": _kind_of(text), "text": text,
                    "chapter": chapter})
    return out


def _chapter_num_of(p) -> str | None:
    """章号：优先 outlineLvl/Heading 样式判定，回退短段文本模式「第X章」。"""
    n = _heading_chapter_number(p)
    if n:
        return n
    t = _visible_txt(p).strip()
    m = _CHAPTER_RE.match(t)
    if not m:
        return None
    if len(t) > 40 or _SENT_PUNCT_RE.search(t):
        return None
    tok = m.group(1)
    if tok.isdigit():
        return tok
    return _cn_num(tok)


def _cn_num(tok: str) -> str | None:
    """中文数字（≤99）→ 阿拉伯数字串（"十二"→"12"；"二十一"→"21"；"十"→"10"）。"""
    if tok in _CN_DIGITS:
        return str(_CN_DIGITS[tok])
    if "十" not in tok:
        return None
    parts = tok.split("十")
    if len(parts) != 2:
        return None
    tens = _CN_DIGITS.get(parts[0], 1) if parts[0] else 1
    ones = _CN_DIGITS.get(parts[1], 0) if parts[1] else 0
    if not (1 <= tens <= 9 and 0 <= ones <= 9):
        return None
    return str(tens * 10 + ones)


# ── 编号 span 定位与字段安全替换 ────────────────────────────────

def _number_span(text: str) -> tuple[int, int] | None:
    """题注编号 token 的字符区间 [start, end)。找不到返回 None。"""
    m = re.match(r"^\s*(?:图|表)\s*", text)
    if not m:
        return None
    rest = text[m.end():]
    nm = _NUMBER_RE.match(rest)
    if not nm or not nm.group(0):
        return None
    return (m.end() + nm.start(), m.end() + nm.end())


def _fld_kind(run) -> list[str]:
    return [fc.get(f"{W}fldCharType") or "" for fc in run.iter(f"{W}fldChar")]


def _replace_number_span(p, start: int, end: int, new_text: str) -> str | None:
    """把段落文本区间 [start, end) 替换为 new_text（run 级手术）。

    - 区间内嵌的域（如题注旧 SEQ 编号域）整体移除（fldChar/instrText 剥离，
      域结果文本清除）——否则 F9 会重算旧域、在字面编号旁再冒出一个数字。
    - 返回 None（成功）或错误原因（结构异常时不动，如实上报）。
    """
    runs = [c for c in p if etree.QName(c).localname == "r"]
    if not runs:
        return "no_runs"
    spans: list[tuple[int, int]] = []
    cur = 0
    for r in runs:
        n = _run_char_len(r)
        spans.append((cur, cur + n))
        cur += n
    if not (0 <= start < end <= cur):
        return "span_out_of_range"
    ri = rj = None
    for i, (a, b) in enumerate(spans):
        if a <= start < b:
            ri = i
        if a <= end - 1 < b:
            rj = i
    if ri is None or rj is None or ri > rj:
        return "span_ambiguous"
    depth_before: list[int] = []
    d = 0
    for r in runs:
        depth_before.append(d)
        for k in _fld_kind(r):
            d += 1 if k == "begin" else (-1 if k == "end" else 0)
    d0 = depth_before[ri]
    # 域尾范围扩展：扫到域深度回落处（含域 end run）
    scan_end = rj
    d = d0
    for i in range(ri, len(runs)):
        for k in _fld_kind(runs[i]):
            d += 1 if k == "begin" else (-1 if k == "end" else 0)
        if i >= rj and d == d0:
            scan_end = i
            break
    # 目标 run：区间内首个「正常深度」文本 run（承载新字面编号）
    target = None
    for i in range(ri, scan_end + 1):
        r = runs[i]
        if r.find(f".//{W}fldChar") is not None or \
                r.find(f".//{W}instrText") is not None:
            continue
        a, b = spans[i]
        if any(len(t.text or "") for t in r.iter(f"{W}t")) and \
                max(a, start) < min(b, end):
            target = i
            break
    to_delete: list = []
    for i in range(ri, scan_end + 1):
        r = runs[i]
        a, b = spans[i]
        # ① 窗口内的 noBreakHyphen（占 1 字符）：属于被替换的编号区 → 随编号整体移除
        loc = a
        for ch in list(r.iter()):
            tag = etree.QName(ch).localname
            if tag == "t":
                loc += len(ch.text or "")
            elif tag == "noBreakHyphen":
                if start <= loc < end and ch.getparent() is not None:
                    ch.getparent().remove(ch)
                loc += 1
        seg_a, seg_b = max(a, start), min(b, end)
        texts = list(r.iter(f"{W}t"))
        has_text = any(len(t.text or "") for t in texts)
        if seg_a >= seg_b or not has_text:
            # 区间外或空文本：只清域残留（成对删除由整体清理兜底）
            if r.find(f".//{W}fldChar") is not None or \
                    r.find(f".//{W}instrText") is not None:
                for x in list(r.iter()):
                    if etree.QName(x).localname in ("fldChar", "instrText"):
                        x.getparent().remove(x)
                if not any(len(t.text or "") for t in r.iter(f"{W}t")) and \
                        not any(etree.QName(x).localname not in ("rPr", "t")
                                for x in r):
                    to_delete.append(r)
            continue
        # 区间内：先剥域壳
        for x in list(r.iter()):
            if etree.QName(x).localname in ("fldChar", "instrText"):
                x.getparent().remove(x)
        # piece 感知保留：NBH 计 1 字符（已在上一步移除），跨窗文本片段按窗口边界切分
        seg_al, seg_bl = seg_a - a, seg_b - a
        keep_l, keep_r = "", ""
        loc = 0
        for ch in r.iter():
            tag = etree.QName(ch).localname
            if tag == "t":
                tx = ch.text or ""
                if loc + len(tx) <= seg_al:
                    keep_l += tx
                elif loc >= seg_bl:
                    keep_r += tx
                elif tx:
                    if loc < seg_al:
                        keep_l += tx[: seg_al - loc]
                    if loc + len(tx) > seg_bl:
                        keep_r += tx[seg_bl - loc:]
                loc += len(tx)
            elif tag == "noBreakHyphen":
                loc += 1
        if i == target:
            texts[0].text = keep_l + new_text + keep_r
        else:
            texts[0].text = keep_l + keep_r
        for t in texts[1:]:
            t.text = ""
    for r in to_delete:
        if r.getparent() is not None:
            r.getparent().remove(r)
    # 清理区间内被清空的 run（域壳/域结果）
    for i in range(ri, scan_end + 1):
        r = runs[i]
        if r.getparent() is None or (target is not None and r is runs[target]):
            continue
        if any(len(t.text or "") for t in r.iter(f"{W}t")):
            continue
        if any(etree.QName(x).localname not in ("rPr", "t") for x in r):
            continue
        r.getparent().remove(r)
    # 区间内原本没有文本 run（编号全在域内）→ 补一个新 run
    if target is None:
        nr = etree.Element(f"{W}r")
        t = etree.SubElement(nr, f"{W}t")
        t.set(_XML_SPACE, "preserve")
        t.text = new_text
        anchor = None
        for r in runs[ri: scan_end + 1]:
            if r.getparent() is not None:
                anchor = r
                break
        if anchor is not None:
            anchor.addprevious(nr)
        else:
            pp = p.find(f"{W}pPr")
            if pp is not None:
                pp.addnext(nr)
            else:
                p.insert(0, nr)
    return None


# ── A4：题注编号按章重排 ────────────────────────────────────────

def renumber_caption_numbers(path, params: dict | None = None) -> dict:
    """题注编号按章重排（图 2-1 / 表 2-1 式），对照清单随结果返回。

    params:
      kinds: "all"|"figure"|"table"（默认 all）
      style: "dash"（2-1）|"dot"（2.1）| None 时自动（现有编号多数派；无证据 dash）
      apply: false（默认，只出对照清单）/ true（落盘）

    边界：只改编号 token 本身（不动前后空格/其它文字）；章前题注、无编号题注
    不动并如实列入 skipped；题注内嵌的旧编号域（SEQ）整体并入新字面编号。
    """
    path = Path(path)
    params = params or {}
    kinds = _norm_kinds(params.get("kinds"))
    apply_ = bool(params.get("apply"))
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False, "applied": False, "reason": "no_body"}

    walk = _walk_captions(root)
    if not walk:
        return {"changed": False, "applied": False, "reason": "no_captions",
                "message": "未发现题注（图 N/表 N 编号段），无可重排对象。"}
    if not any(c["chapter"] for c in walk):
        return {"changed": False, "applied": False,
                "reason": "no_chapter_headings", "captions": len(walk),
                "message": "未发现章标题（第X章 / outlineLvl），无法按章编号——不动。"}

    changes: list[dict] = []
    skipped: list[dict] = []
    seq = {"figure": 0, "table": 0}
    last = {"figure": None, "table": None}
    chapters_seen: list[str] = []
    for c in walk:
        kind = c["kind"]
        if kind not in kinds:
            continue
        if c["chapter"] is None:
            skipped.append({"kind": kind, "text": c["text"][:50],
                            "reason": "章前题注（无章号上下文）"})
            continue
        span = _number_span(c["text"])
        if span is None:
            skipped.append({"kind": kind, "text": c["text"][:50],
                            "reason": "题注无编号前缀"})
            continue
        if c["chapter"] not in chapters_seen:
            chapters_seen.append(c["chapter"])
        if last[kind] != c["chapter"]:
            last[kind] = c["chapter"]
            seq[kind] = 0
        seq[kind] += 1
        changes.append({"kind": kind, "chapter": c["chapter"], "seq": seq[kind],
                        "old": c["text"][span[0]:span[1]].strip(),
                        "caption": c["text"], "para": c["para"], "span": span})
    if not changes:
        return {"changed": False, "applied": False, "reason": "nothing",
                "message": "选中范围内无可重排的题注编号。", "skipped": skipped}

    style = (params.get("style") or "").strip().lower()
    if style not in ("dash", "dot"):
        dashes = sum(1 for c in changes if "-" in c["old"].replace("—", "-"))
        dots = sum(1 for c in changes if any(ch in c["old"] for ch in ".．·"))
        style = "dot" if dots > dashes else "dash"
    for c in changes:
        sep = "-" if style == "dash" else "."
        c["new"] = f"{c['chapter']}{sep}{c['seq']}"
    pending = [c for c in changes if c["old"] != c["new"]]
    if not pending:
        return {"changed": False, "applied": False, "style_used": style,
                "captions": len(changes), "chapters": chapters_seen,
                "changes": [], "skipped": skipped, "verified": True,
                "message": "题注编号已符合按章序列，无需重排。"}
    if not apply_:
        return {"changed": False, "applied": False, "style_used": style,
                "captions": len(changes), "chapters": chapters_seen,
                "changes": [_pub(c) for c in pending], "skipped": skipped,
                "totals": _totals(pending),
                "message": "仅出对照清单（未落盘）；确认后 apply=true 执行。"}
    done: list[dict] = []
    failed: list[dict] = []
    for c in pending:
        err = _replace_number_span(c["para"], c["span"][0], c["span"][1], c["new"])
        if err:
            failed.append({"old": c["old"], "new": c["new"],
                           "caption": c["caption"][:50], "error": err})
        else:
            done.append(c)
    if done:
        _write_part(path, "word/document.xml", root)
    result = {"changed": len(done) > 0, "applied": True, "style_used": style,
              "captions": len(changes), "chapters": chapters_seen,
              "changes": [_pub(c) for c in done], "skipped": skipped,
              "totals": _totals(done), "failed": failed}
    try:
        result["verified"] = verify_caption_numbers_renumbered(path, params)
    except Exception:  # noqa: BLE001
        result["verified"] = None
    return result


def _pub(c: dict) -> dict:
    return {"kind": c["kind"], "old": c["old"], "new": c["new"],
            "caption": c["caption"][:60]}


def _totals(changes: list[dict]) -> dict:
    t = {"figure": 0, "table": 0}
    for c in changes:
        t[c["kind"]] = t.get(c["kind"], 0) + 1
    return t


def verify_caption_numbers_renumbered(path, params: dict | None = None) -> bool:
    """复查：全部可重排题注的编号 == 重算期望值（豁免项不参与）。

    apply=false 是提案阶段（预览对照清单），恒真（对齐 convert_text_equations 约定）。
    """
    try:
        params = params or {}
        if not params.get("apply"):
            return True
        path = Path(path)
        kinds = _norm_kinds(params.get("kinds"))
        root = _read_part(path, "word/document.xml")
        walk = _walk_captions(root)
        if not walk:
            return True
        style = (params.get("style") or "").strip().lower()
        if style not in ("dash", "dot"):
            dashes = dots = 0
            for c in walk:
                sp = _number_span(c["text"])
                if not sp:
                    continue
                old = c["text"][sp[0]:sp[1]]
                if "-" in old:
                    dashes += 1
                if any(ch in old for ch in ".．·"):
                    dots += 1
            style = "dot" if dots > dashes else "dash"
        sep = "-" if style == "dash" else "."
        seq = {"figure": 0, "table": 0}
        last = {"figure": None, "table": None}
        for c in walk:
            kind = c["kind"]
            if kind not in kinds or c["chapter"] is None:
                continue
            sp = _number_span(c["text"])
            if sp is None:
                continue
            if last[kind] != c["chapter"]:
                last[kind] = c["chapter"]
                seq[kind] = 0
            seq[kind] += 1
            want = f"{c['chapter']}{sep}{seq[kind]}"
            if c["text"][sp[0]:sp[1]].strip() != want:
                return False
        return True
    except Exception:  # noqa: BLE001
        return False


# ── A1：图目录/表目录生成 ───────────────────────────────────────

def _has_seq(p, kind_char: str) -> bool:
    # 注意：真实文档里域指令可能被拆成多个 instrText（"SEQ" + "图"+"\\* ARABIC…"），
    # 必须拼接后判定（2026-10-08 西安理工样本实测）。
    joined = "".join(x.text or "" for x in p.iter(f"{W}instrText"))
    return "SEQ" in joined and kind_char in joined


def _literalize_caption_seqs(p, kind_char: str) -> int:
    """把题注段落里「可见」的旧 SEQ <kind> 域字面化（域 → 其缓存结果文本）。

    为什么必须做：F9 时 Word 会把这类域按全局 SEQ 计数重算（与章无关）——实测
    原文档 4 个可见域在引擎加隐藏采集域后漂移成「图6-35」等怪编号。字面化后
    编号稳定（编号权威 = 引擎写死的字面文本），采集统一交给隐藏域。
    跳过引擎自己的隐藏（vanish）域。返回字面化的域数量。
    """
    runs = [c for c in p if etree.QName(c).localname == "r"]
    if not runs:
        return 0
    done = 0
    i = 0
    while i < len(runs):
        if "begin" not in _fld_kind(runs[i]):
            i += 1
            continue
        depth = 0
        j = i
        found_end = False
        in_result = False
        instr_parts: list[str] = []
        result_parts: list[str] = []
        while j < len(runs):
            rr = runs[j]
            for fc in rr.iter(f"{W}fldChar"):
                k = fc.get(f"{W}fldCharType") or ""
                if k == "begin":
                    depth += 1
                elif k == "end":
                    depth -= 1
                elif k == "separate" and depth == 1:
                    in_result = True
            instr_parts.extend(x.text or "" for x in rr.iter(f"{W}instrText"))
            if in_result:
                result_parts.extend(t.text or "" for t in rr.iter(f"{W}t"))
            if depth <= 0:
                found_end = True
                break
            j += 1
        if not found_end:
            i += 1
            continue
        instr = "".join(instr_parts)
        hidden = all(
            (runs[kk].find(f"{W}rPr") is not None
             and runs[kk].find(f"{W}rPr").find(f"{W}vanish") is not None)
            for kk in range(i, j + 1)
        )
        if not hidden and "SEQ" in instr and kind_char in instr:
            cache = "".join(result_parts).strip()
            parent = runs[i].getparent()
            idx = list(parent).index(runs[i])
            for kk in range(i, j + 1):
                parent.remove(runs[kk])
            if cache:
                nr = etree.Element(f"{W}r")
                t = etree.SubElement(nr, f"{W}t")
                t.set(_XML_SPACE, "preserve")
                t.text = cache
                parent.insert(idx, nr)
            done += 1
        i = j + 1
    return done


def _vanish_rpr():
    rpr = etree.Element(f"{W}rPr")
    v = etree.SubElement(rpr, f"{W}vanish")
    v.set(f"{W}val", "1")
    return rpr


def _seq_runs(kind_char: str, cached: str):
    """隐藏采集域（vanish 全 run；与探针验证过的配置一致）。"""
    runs = []
    for el in ("begin", None, "separate", "text", "end"):
        r = etree.Element(f"{W}r")
        r.append(_vanish_rpr())
        if el == "text":
            t = etree.SubElement(r, f"{W}t")
            t.text = cached
        elif el is None:
            ins = etree.SubElement(r, f"{W}instrText")
            ins.set(_XML_SPACE, "preserve")
            ins.text = f" SEQ {kind_char} "
        else:
            fc = etree.SubElement(r, f"{W}fldChar")
            fc.set(f"{W}fldCharType", el)
        runs.append(r)
    return runs


def _add_bookmark(p, name: str, bid: int) -> None:
    bs = etree.Element(f"{W}bookmarkStart")
    bs.set(f"{W}id", str(bid))
    bs.set(f"{W}name", name)
    be = etree.Element(f"{W}bookmarkEnd")
    be.set(f"{W}id", str(bid))
    ppr = p.find(f"{W}pPr")
    pos = list(p).index(ppr) + 1 if ppr is not None else 0
    p.insert(pos, bs)
    p.append(be)


def _remove_tof_bookmarks(root, kinds) -> int:
    removed = 0
    wanted = {"figure": "_TofF", "table": "_TofT"}
    prefixes = tuple(wanted[k] for k in _KINDS if k in set(kinds))
    for bs in list(root.iter(f"{W}bookmarkStart")):
        name = bs.get(f"{W}name") or ""
        if not name.startswith("_Tof"):
            continue
        if prefixes and not name.startswith(prefixes):
            continue
        bid = bs.get(f"{W}id")
        if bs.getparent() is not None:
            bs.getparent().remove(bs)
        for be in list(root.iter(f"{W}bookmarkEnd")):
            if be.get(f"{W}id") == bid:
                be.getparent().remove(be)
        removed += 1
    return removed


def _max_bookmark_id(root) -> int:
    mx = 0
    for el in list(root.iter(f"{W}bookmarkStart")) + \
            list(root.iter(f"{W}bookmarkEnd")):
        try:
            mx = max(mx, int(el.get(f"{W}id") or 0))
        except ValueError:
            pass
    return mx


def _bookmark_names(root) -> set[str]:
    return {bs.get(f"{W}name") or "" for bs in root.iter(f"{W}bookmarkStart")}


def _field_end_paragraph(body, start_idx: int) -> int:
    """从含域 begin 的段落起，返回域深度回 0 的段落下标（end 所在段）。"""
    children = list(body)
    depth = 0
    seen = False
    for i in range(start_idx, len(children)):
        el = children[i]
        if etree.QName(el).localname != "p":
            continue
        for fc in el.iter(f"{W}fldChar"):
            k = fc.get(f"{W}fldCharType") or ""
            if k == "begin":
                depth += 1
                seen = True
            elif k == "end":
                depth -= 1
        if seen and depth <= 0:
            return i
    return len(children) - 1


def _toc_field_para(body, kind_char: str):
    """找 `TOC ... \\c "<kind>"` 域 begin 所在段落。返回 (idx, para) 或 None。

    指令可能被 Word 拆成多个 instrText（如 `TOC \\h \\z \\c "` + `图` + `"`），
    必须拼接后判定。
    """
    for i, el in enumerate(list(body)):
        if etree.QName(el).localname != "p":
            continue
        joined = _instr_text(el)
        if "TOC" in joined and f'\\c "{kind_char}"' in joined:
            return i, el
    return None


def _zone_bounds(body, kind: str):
    """现有目录区的段落下标范围 (start, end)。无区返回 None。"""
    found = _toc_field_para(body, _KIND_CHAR[kind])
    if found is None:
        return None
    fi, _fp = found
    end_i = _field_end_paragraph(body, fi)
    children = list(body)
    start_i = fi
    if fi > 0 and etree.QName(children[fi - 1]).localname == "p":
        tprev = _ptxt(children[fi - 1]).strip()
        if tprev in (_TOF_TITLE_WORDS[kind] | {_DEFAULT_HEADINGS[kind]}):
            start_i = fi - 1
    return start_i, end_i


def _first_chapter_idx(body) -> int | None:
    """首个章标题段落下标。

    跳过含 PAGEREF 的段落——目录条目文本也长成「第1章 绪论」，不跳过会把目录
    条目当章标题（2026-10-08 真实样本加固）。
    """
    for i, el in enumerate(list(body)):
        if etree.QName(el).localname != "p":
            continue
        if "PAGEREF" in _instr_text(el):
            continue
        if _chapter_num_of(el):
            return i
    return None


def _main_toc_anchor(body) -> int | None:
    """主目录末尾锚点：最后一个含 `_Toc\\d` PAGEREF 的块（段/sdt/tbl）下标。

    须在首章前。真实样本的主目录常包在 <w:sdt> 内（Word 目录内容控件），
    嵌套条目也算数——锚点取容器的直接子元素下标。
    """
    children = list(body)
    first_chapter = _first_chapter_idx(body)

    def _has_toc_ref(p) -> bool:
        return any(re.match(r"^_Toc\d", m.group(1))
                   for m in re.finditer(r"PAGEREF\s+(\S+)", _instr_text(p)))

    main_idxs: list[int] = []
    for i, el in enumerate(children):
        ln = etree.QName(el).localname
        if ln == "p":
            if _has_toc_ref(el):
                main_idxs.append(i)
        elif ln in ("sdt", "tbl"):
            if any(_has_toc_ref(pp) for pp in el.iter(f"{W}p")):
                main_idxs.append(i)
    if first_chapter is not None:
        main_idxs = [i for i in main_idxs if i < first_chapter]
    return main_idxs[-1] if main_idxs else None


def _toc_title_para(root):
    """「目录」标题段（用于抄样式）。"""
    body = root.find(f"{W}body")
    if body is None:
        return None
    for p in body.iter(f"{W}p"):
        t = _visible_txt(p).strip().replace(" ", "")
        if t == "目录":
            return p
    return None


def _title_para(heading_text: str, toc_title_p):
    p = etree.Element(f"{W}p")
    if toc_title_p is not None:
        src_ppr = toc_title_p.find(f"{W}pPr")
        if src_ppr is not None:
            ppr = copy.deepcopy(src_ppr)
            for el in ppr.findall(f"{W}outlineLvl"):
                ppr.remove(el)
            p.append(ppr)
    r = etree.Element(f"{W}r")
    if toc_title_p is not None:
        src_r = next((x for x in toc_title_p.findall(f"{W}r")
                      if "".join(t.text or "" for t in x.iter(f"{W}t")).strip()),
                     None)
        if src_r is not None:
            srpr = src_r.find(f"{W}rPr")
            if srpr is not None:
                r.append(copy.deepcopy(srpr))
    t = etree.SubElement(r, f"{W}t")
    t.set(_XML_SPACE, "preserve")
    t.text = heading_text
    p.append(r)
    return p


def _fld_run(kind: str):
    r = etree.Element(f"{W}r")
    fc = etree.SubElement(r, f"{W}fldChar")
    fc.set(f"{W}fldCharType", kind)
    return r


def _instr_run(instr: str):
    r = etree.Element(f"{W}r")
    ins = etree.SubElement(r, f"{W}instrText")
    ins.set(_XML_SPACE, "preserve")
    ins.text = f" {instr} "
    return r


def _simple_run(text: str):
    r = etree.Element(f"{W}r")
    t = etree.SubElement(r, f"{W}t")
    t.set(_XML_SPACE, "preserve")
    t.text = text
    return r


def _tab_run():
    r = etree.Element(f"{W}r")
    etree.SubElement(r, f"{W}tab")
    return r


def _entry_para(text: str, bm: str, tab_pos: int, style_id: str | None):
    p = etree.Element(f"{W}p")
    ppr = etree.SubElement(p, f"{W}pPr")
    if style_id:
        ps = etree.SubElement(ppr, f"{W}pStyle")
        ps.set(f"{W}val", style_id)
    tabs = etree.SubElement(ppr, f"{W}tabs")
    tab = etree.SubElement(tabs, f"{W}tab")
    tab.set(f"{W}val", "right")
    tab.set(f"{W}leader", "dot")
    tab.set(f"{W}pos", str(tab_pos))
    p.append(_simple_run(text))
    p.append(_tab_run())
    p.append(_fld_run("begin"))
    p.append(_instr_run(f"PAGEREF {bm} \\h"))
    p.append(_fld_run("separate"))
    p.append(_simple_run("1"))
    p.append(_fld_run("end"))
    return p


def _right_tab_pos(body) -> int:
    sect = body.find(f"{W}sectPr")
    tab_pos = 8776
    if sect is not None:
        pg = sect.find(f"{W}pgSz")
        mar = sect.find(f"{W}pgMar")
        if pg is not None and mar is not None:
            try:
                tab_pos = (int(pg.get(f"{W}w", "11906"))
                           - int(mar.get(f"{W}left", "1800"))
                           - int(mar.get(f"{W}right", "1800")))
            except (TypeError, ValueError):
                pass
    return tab_pos


def _wrap_toc_field(entry_paras: list, kind_char: str) -> None:
    """条目段包 `TOC \\h \\z \\c` 域壳（begin/instr/separate 挂首段前，end 挂末段后）。"""
    first, last = entry_paras[0], entry_paras[-1]
    instr = f'TOC \\h \\z \\c "{kind_char}"'
    ppr = first.find(f"{W}pPr")
    pos = list(first).index(ppr) + 1 if ppr is not None else 0
    for off, el in enumerate((_fld_run("begin"), _instr_run(instr),
                              _fld_run("separate"))):
        first.insert(pos + off, el)
    last.append(_fld_run("end"))


def generate_caption_toc(path, params: dict | None = None) -> dict:
    """生成/重建「图目录」「表目录」区（TOC \\c 域 + 引擎缓存条目 + 题注书签/采集域）。

    params:
      kinds: "all"|"figure"|"table"（默认 all）
      heading_texts: {"figure": "图目录", "table": "表目录"} 覆盖标题
      apply: false（默认，只出预览清单）/ true（落盘）

    行为：题注须有可识别编号前缀；无题注时报「未发现题注」不生成空目录；
    已存在同 kind 目录区就地重建（幂等，不堆叠）。
    """
    path = Path(path)
    params = params or {}
    kinds = _norm_kinds(params.get("kinds"))
    apply_ = bool(params.get("apply"))
    headings = dict(_DEFAULT_HEADINGS)
    for k, v in (params.get("heading_texts") or {}).items():
        if k in _KINDS and isinstance(v, str) and v.strip():
            headings[k] = v.strip()
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False, "applied": False, "reason": "no_body"}

    caps = _walk_captions(root)
    per_kind = {k: [c for c in caps if c["kind"] == k] for k in kinds}
    if not any(per_kind.values()):
        return {"changed": False, "applied": False, "reason": "no_captions",
                "found": 0,
                "message": "未发现题注（图 N/表 N 编号段），不生成空目录。"}

    if not apply_:
        preview = {}
        for k in kinds:
            if per_kind[k]:
                preview[k] = {"heading": headings[k], "entries": len(per_kind[k]),
                              "first": [c["text"][:50] for c in per_kind[k][:3]]}
        return {"changed": False, "applied": False, "preview": preview,
                "message": "仅出预览（未落盘）；确认后 apply=true 生成。",
                "totals": {k: len(per_kind[k]) for k in kinds if per_kind[k]}}

    styles = _read_styles_map(path)
    style_id = "TOC2" if "TOC2" in styles else ("TOC1" if "TOC1" in styles else None)
    tab_pos = _right_tab_pos(body)
    toc_title_p = _toc_title_para(root)
    max_bid = _max_bookmark_id(root)
    warnings: list[str] = []
    zones_info: list[dict] = []

    # 已有区检测（重建目标）
    rebuild = {k: _zone_bounds(body, k) for k in kinds}
    # 旧书签清理（全部请求 kind——防此前进度态残留）
    _remove_tof_bookmarks(root, kinds)
    existing_names = _bookmark_names(root)

    anchor = _main_toc_anchor(body)
    any_rebuild = any(rebuild.values())
    if anchor is None and not any_rebuild:
        fc = _first_chapter_idx(body)
        if fc is None or toc_title_p is None:
            return {"changed": False, "applied": False, "reason": "no_anchor",
                    "message": "未发现主目录条目/章节结构，无法定位插入位置——"
                               "请先确保文档有目录区。"}
        anchor = fc - 1

    # ── 构建区（题注书签/采集域 + 条目段 + 域壳）──
    zones: dict[str, list] = {}
    seq_cached = {k: 0 for k in kinds}
    for k in kinds:
        kind_list = per_kind[k]
        if not kind_list:
            warnings.append(f"未发现{_KIND_CHAR[k]}题注，未生成对应目录。")
            continue
        added_bm = added_seq = added_lit = 0
        for n, c in enumerate(kind_list, 1):
            p = c["para"]
            # 旧可见 SEQ 域字面化（防 F9 全局计数漂移；编号权威=字面文本）
            added_lit += _literalize_caption_seqs(p, _KIND_CHAR[k])
            name = f"_Tof{'F' if k == 'figure' else 'T'}%04d" % n
            if name in existing_names:
                n2 = n + 10000
                while name in existing_names:
                    name = f"_Tof{'F' if k == 'figure' else 'T'}%04d" % n2
                    n2 += 10000
            if not _has_seq(p, _KIND_CHAR[k]):
                seq_cached[k] += 1
                for r in _seq_runs(_KIND_CHAR[k], str(seq_cached[k])):
                    p.append(r)
                added_seq += 1
            max_bid += 1
            _add_bookmark(p, name, max_bid)
            existing_names.add(name)
            added_bm += 1
            c["bm"] = name
        entry_paras = [_entry_para(c["text"], c["bm"], tab_pos, style_id)
                       for c in kind_list]
        _wrap_toc_field(entry_paras, _KIND_CHAR[k])
        title_p = _title_para(headings[k], toc_title_p)
        zones[k] = [title_p] + entry_paras
        zones_info.append({"kind": k, "heading": headings[k],
                           "entries": len(entry_paras), "bookmarks": added_bm,
                           "seqs_added": added_seq, "literalized": added_lit,
                           "rebuilt": bool(rebuild[k])})

    # ── 落位 1：重建已有区（就地替换）──
    for k in kinds:
        if k not in zones or not rebuild.get(k):
            continue
        bounds = _zone_bounds(body, k)  # 现场重算（前序 splice 后下标仍准）
        if bounds is None:
            continue
        start_i, end_i = bounds
        children = list(body)
        for el in children[start_i:end_i + 1]:
            body.remove(el)
        for off, el in enumerate(zones[k]):
            body.insert(start_i + off, el)

    # ── 落位 2：新区插入（顺序：紧跟前一 kind 的区尾；否则主目录之后）──
    for k in kinds:
        if k not in zones or rebuild.get(k):
            continue
        cur = None
        prev = _KINDS[_KINDS.index(k) - 1] if k != _KINDS[0] else None
        if prev is not None:
            pb = _zone_bounds(body, prev)
            if pb is not None:
                cur = pb[1] + 1
        if cur is None and anchor is not None:
            cur = anchor + 1
        if cur is None:
            fc = _first_chapter_idx(body)
            cur = fc if fc is not None else 0
        for off, el in enumerate(zones[k]):
            body.insert(cur + off, el)

    _write_part(path, "word/document.xml", root)

    result: dict = {"changed": True, "applied": True, "zones": zones_info,
                    "totals": {k: len(per_kind[k]) for k in kinds if per_kind[k]},
                    "warnings": warnings}
    # 页码真机回写（best effort；无 Word 环境降级说明，整篇编译校验时会再回写）
    try:
        from earlyz.validation.renderers import probe_bookmark_pages
        from earlyz.actions.new_actions import sync_toc_page_numbers
        pm = probe_bookmark_pages(str(path))
        # 显示页码优先（wdActiveEndAdjustedPageNumber）：目录/图目录写的必须是页脚
        # 上显示的页码；物理页 − offset 只对「正文从第 1 页起算」的模板成立。
        adj = pm.get("adjusted_bookmarks") or {}
        bms = adj or (pm.get("bookmarks") or {})
        pmap = {kk: vv for kk, vv in bms.items()
                if kk.startswith("_Toc") or kk.startswith("_Tof")}
        if pmap:
            sync = sync_toc_page_numbers(path, {"page_map": pmap,
                                                "offset": 0 if adj else None})
            result["page_sync"] = {"entries": sync.get("entries"),
                                   "updated": sync.get("updated"),
                                   "offset": sync.get("offset"),
                                   "source": "word-adjusted" if adj else "word"}
        else:
            result["page_sync"] = {"source": "none"}
    except Exception as e:  # noqa: BLE001 —— 不阻断；编译校验层会再回写
        result["page_sync"] = {"source": "deferred",
                               "note": f"页码待编译校验回写：{type(e).__name__}"}
    try:
        result["verified"] = verify_caption_toc(path, params)
    except Exception:  # noqa: BLE001
        result["verified"] = None
    return result


def verify_caption_toc(path, params: dict | None = None) -> bool:
    """复查：每 kind 有 `\\c` 域、条目数==题注数、条目指向成对存在的 _Tof 书签。

    apply=false 是预览阶段（未落盘），恒真（对齐 convert_text_equations 约定）。
    """
    try:
        params = params or {}
        if not params.get("apply"):
            return True
        path = Path(path)
        kinds = _norm_kinds(params.get("kinds"))
        root = _read_part(path, "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return False
        caps = _walk_captions(root)
        names = _bookmark_names(root)
        for k in kinds:
            kind_list = [c for c in caps if c["kind"] == k]
            if not kind_list:
                continue
            bounds = _zone_bounds(body, k)
            if bounds is None:
                return False
            start_i, end_i = bounds
            children = list(body)
            entry_count = 0
            target_names: set[str] = set()
            for el in children[start_i:end_i + 1]:
                if etree.QName(el).localname != "p":
                    continue
                for m in re.finditer(r"PAGEREF\s+(\S+)", _instr_text(el)):
                    target_names.add(m.group(1))
                    entry_count += 1
            if entry_count != len(kind_list):
                return False
            if not target_names or not all(
                    tn.startswith("_Tof") and tn in names for tn in target_names):
                return False
        return True
    except Exception:  # noqa: BLE001
        return False


# ── A5（EZI-0011）：图题绑图 / 表题绑表（keepNext 方向修正）─────────
#
# 中文规范：**图题在图下方、表题在表上方**。`w:keepNext` 是单向的——它把
# 「本段 + 后面内容」绑在同页，不能往前绑。于是两种题注的方向天然相反：
#
#     图   → 题注在下 → 要给**图片段落**加 keepNext（图段绑住后面的图题）
#     表   → 题注在上 → 要给**表题段落**加 keepNext（题段绑住后面的表格）
#
# 此前引擎只有 `normalize_table_caption_block` 一套逻辑（正则把「图/表」写在
# 一起），于是图题也只拿到「自己 keepNext」——绑的是图题**之后**的正文，图与
# 图题之间没有任何绑定，正是用户看到的「图标题和图跨页」。
#
# 第二个隐性断点：题与对象之间夹一个空段时，keepNext 链断在空段上（空段自己
# 没有 keepNext），题注照样能滑到下一页。所以绑之前先把「只含 rPr 的空段」拆掉
# （有字/有图/有书签/有域/有分页符/有 sectPr 的段一律不动）。

#: 图对象容器（DrawingML 图片 / VML 图 / OLE 嵌入对象）
_GRAPHIC_TAGS = ("drawing", "pict", "object")
#: 取「段落自身文字」时要排除的图形内部容器
_GRAPHIC_SCOPE = ("drawing", "pict", "object", "txbxContent")
#: 流内包装层：穿过它们取块级元素（内容控件里的段落仍算流内段落）
_FLOW_WRAPPERS = frozenset({"sdt", "sdtContent", "customXml", "smartTag",
                            "ins", "del", "moveTo", "moveFrom"})
#: 题注文本：图 / 表 + 编号（阿拉伯或中文数字）
_CAPTION_TEXT_RE = re.compile(r"^(图|表)\s*[0-9一二三四五六七八九十]")


def _caption_flow_text(p) -> str:
    """段落自身文字（流内 ``w:t``）；图形/文本框里的字不算段落文字。

    真实样本（``output/毕业设计.docx``）30 张图的题注写在图的文本框里——按
    ``p.iter(w:t)`` 取字会把「图对象」当成「题注段」处理，方向就判错了。
    """
    out: list[str] = []
    for t in p.iter(f"{W}t"):
        anc = t.getparent()
        skip = False
        while anc is not None and anc is not p:
            if etree.QName(anc).localname in _GRAPHIC_SCOPE:
                skip = True
                break
            anc = anc.getparent()
        if not skip:
            out.append(t.text or "")
    return "".join(out).strip()


def _block_has_graphic(el) -> bool:
    return (el is not None and el.tag == f"{W}p"
            and any(el.find(f".//{W}{t}") is not None for t in _GRAPHIC_TAGS))


def _p_has(p, tag: str) -> bool:
    ppr = p.find(f"{W}pPr") if p is not None else None
    return ppr is not None and ppr.find(f"{W}{tag}") is not None


def _ensure_ppr_flag(p, tag: str) -> bool:
    """给段落补 ``w:keepNext`` / ``w:keepLines``（按 CT_PPr schema 序插入）。"""
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        ppr = etree.Element(f"{W}pPr")
        p.insert(0, ppr)
    if ppr.find(f"{W}{tag}") is not None:
        return False
    _ppr_insert_ordered(ppr, etree.Element(f"{W}{tag}"))
    return True


def _drop_ppr_flag(p, tag: str) -> bool:
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        return False
    el = ppr.find(f"{W}{tag}")
    if el is None:
        return False
    ppr.remove(el)
    return True


def _flow_blocks(body) -> list:
    """正文流内块级元素（w:p / w:tbl），穿过 sdt/customXml 等包装层。"""
    def walk(el):
        for ch in el:
            name = etree.QName(ch).localname
            if name in _FLOW_WRAPPERS:
                yield from walk(ch)
            elif name in ("p", "tbl"):
                yield ch

    return list(walk(body)) if body is not None else []


def _droppable_blank(el) -> bool:
    """只含 ``pPr``（且无 sectPr）+ 仅带 ``rPr`` 的空 run 的段落——可安全拆掉。

    有字/有图/书签/域/换行（含分页符）/制表符/超链接/公式的段一律不拆：那些
    都是内容，拆了就是「删内容」。
    """
    if el is None or el.tag != f"{W}p":
        return False
    for child in el:
        name = etree.QName(child).localname
        if name == "pPr":
            if child.find(f"{W}sectPr") is not None:
                return False
            continue
        if name != "r":
            return False
        for rc in child:
            if etree.QName(rc).localname != "rPr":
                return False
    return True


def _table_flag_counts(root) -> dict:
    """``w:cantSplit`` / ``w:tblHeader`` 行数（A4：修题注绑定不许顺手拆掉它们）。"""
    rows = list(root.iter(f"{W}tr"))
    cant = hdr = 0
    for tr in rows:
        trpr = tr.find(f"{W}trPr")
        if trpr is None:
            continue
        if trpr.find(f"{W}cantSplit") is not None:
            cant += 1
        if trpr.find(f"{W}tblHeader") is not None:
            hdr += 1
    return {"rows": len(rows), "cant_split": cant, "tbl_header": hdr}


def _caption_binding_counts(root) -> dict:
    """当前产物的题注绑定实测（与打分器 ``caption_binding_stats`` 同口径）。"""
    body = root.find(f"{W}body")
    blocks = _flow_blocks(body)
    out = {"figure": {"total": 0, "bound": 0, "unbound": 0, "not_adjacent": 0},
           "table": {"total": 0, "bound": 0, "unbound": 0, "not_adjacent": 0},
           "embedded": {"figure": 0, "table": 0}}
    for i, el in enumerate(blocks):
        if el.tag != f"{W}p":
            continue
        m = _CAPTION_TEXT_RE.match(_caption_flow_text(el))
        if m is None:
            continue
        kind = "figure" if m.group(1) == "图" else "table"
        st = out[kind]
        st["total"] += 1
        if kind == "figure":
            prev = blocks[i - 1] if i else None
            if not _block_has_graphic(prev):
                st["not_adjacent"] += 1
            elif _p_has(prev, "keepNext"):
                st["bound"] += 1
            else:
                st["unbound"] += 1
        else:
            nxt = blocks[i + 1] if i + 1 < len(blocks) else None
            if nxt is None or nxt.tag != f"{W}tbl":
                st["not_adjacent"] += 1
            elif _p_has(el, "keepNext"):
                st["bound"] += 1
            else:
                st["unbound"] += 1
    for p in (body.iter(f"{W}p") if body is not None else ()):
        anc = p.getparent()
        inside = False
        while anc is not None:
            if etree.QName(anc).localname in ("txbxContent", "drawing", "pict"):
                inside = True
                break
            anc = anc.getparent()
        if not inside:
            continue
        m = _CAPTION_TEXT_RE.match(_caption_flow_text(p))
        if m is not None:
            out["embedded"]["figure" if m.group(1) == "图" else "table"] += 1
    return out


# ── A1/A2（EZI-0013）：题注抽流 ─────────────────────────────────
#
# 真实产物（`output/毕业设计.docx`）的题注不在版面流里，而是画在图上的**浮动
# 文本框**；Word 还把同一份内容存了两份副本——
#
#     w:p
#       w:r → mc:AlternateContent
#              mc:Choice   → w:drawing → wp:anchor → wps:wsp → wps:txbx → w:txbxContent（题注段）
#              mc:Fallback → w:pict    → v:shape  → v:textbox → w:txbxContent（同一份内容的副本）
#
# 三个后果：① 流内题注 0 条 → 图题绑定永远 N-A；② 题注字号/编号/加粗都按「流内
# 段落」实现，文本框里的字全漏；③ 两个分支都取就成对重复。
#
# 本动作：**只取 Choice**（Fallback 是副本，丢弃并计数）→ 把文本框里的题注段深
# 拷贝成流内段落，插到**图对象段之后**（锚段自身没有真图时，找最近的前一个图段）
# → 再拆掉形状（不拆的话 Word 会把题注画两遍）。深拷贝保留 pPr/rPr 原样 = 字号/
# 中英文字体/对齐/加粗全部继承（A2），只剥分页符与页框。
#
# 幂等：形状拆掉后第二次运行找不到文本框题注 → extracted=0 / changed=False。

#: DrawingML 主命名空间（真图的 `a:blip` 在这里）
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
#: VML 命名空间（`v:imagedata` / `v:textbox`）
_V = "{urn:schemas-microsoft-com:vml}"
#: Markup Compatibility（`mc:AlternateContent` / `mc:Choice` / `mc:Fallback`）
_MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
#: Wordprocessing Drawing（`wp:anchor` / `wp:positionH`——浮动文本框的位置）
_WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
#: 形状容器的非内容子元素（拆形状后只剩这些的 run 一并删掉）
_RUN_SHELL_TAGS = ("rPr", "lastRenderedPageBreak", "proofErr", "bookmarkStart",
                   "bookmarkEnd", "commentRangeStart", "commentRangeEnd")


def _block_has_picture(el) -> bool:
    """块里是否有**真图**（DrawingML 图片 / VML 图片 / OLE 对象）。

    与 `_block_has_graphic` 的区别：后者把「只装题注文本框的 drawing」也算图段，
    用它找图段会把题注挂到自己的锚段上（实测 30 条题注里 15 条正是这种锚段没有
    真图的形态）。
    """
    if el is None or el.tag != f"{W}p":
        return False
    if el.find(f".//{_A}blip") is not None:
        return True
    if el.find(f".//{_V}imagedata") is not None:
        return True
    return el.find(f".//{W}object") is not None


def _is_ancestor(anc, el) -> bool:
    node = el.getparent()
    while node is not None:
        if node is anc:
            return True
        node = node.getparent()
    return False


def _shape_containers(p) -> list:
    """段落里承载文本框的**最外层**形状容器（文档序）。

    `mc:AlternateContent` 优先（Choice/Fallback 是一个对象的两个副本）；没有
    AlternateContent 时退回裸 `w:drawing` / `w:pict`。嵌套在别的文本框里的形状
    不算——那是文本框内嵌文本框，位置语义不同，不抽。
    """
    owned = []
    for el in p.iter(f"{_MC}AlternateContent"):
        if el.find(f".//{W}txbxContent") is None:
            continue
        inside_txbx = False
        node = el.getparent()
        while node is not None and node is not p:
            if etree.QName(node).localname == "txbxContent":
                inside_txbx = True
                break
            node = node.getparent()
        if not inside_txbx:
            owned.append(el)
    outer = [el for el in owned
             if not any(o is not el and _is_ancestor(o, el) for o in owned)]
    if outer:
        return outer
    out = []
    for el in p.iter():
        if etree.QName(el).localname not in ("drawing", "pict"):
            continue
        if el.find(f".//{W}txbxContent") is None:
            continue
        if any(o is not el and _is_ancestor(o, el) for o in out):
            continue
        out.append(el)
    return out


def _shape_branch(shape):
    """形状的「一份」内容：AlternateContent 取 Choice（其次 Fallback），其余取自身。"""
    if shape.tag == f"{_MC}AlternateContent":
        for tag in ("Choice", "Fallback"):
            el = shape.find(f"{_MC}{tag}")
            if el is not None and el.find(f".//{W}txbxContent") is not None:
                return el
        return None
    return shape


def _txbx_paragraphs(branch) -> list:
    """分支里**第一个**文本框的段落（穿过 sdt 等包装层）。"""
    if branch is None:
        return []
    tx = branch.find(f".//{W}txbxContent")
    return _flow_blocks(tx)


def _flow_safe_caption(p, default_jc: str | None = None):
    """深拷贝文本框里的题注段 → 可直接进入版面流（A2：继承原段落格式）。

    浮动文本框的「对齐」不在段落 pPr 里而是在锚点位置上（见 `_shape_h_align`），
    所以原段落没有 `w:jc` 时按形状位置补一条——否则居中题注抽进正文会变左对齐
    （真机实测：图3.1/图4.10/图4.11 三页题注全部贴左，与抽流前不一致）。
    """
    new_p = copy.deepcopy(p)
    ppr = new_p.find(f"{W}pPr")
    if ppr is not None:
        for tag in ("pageBreakBefore", "framePr", "sectPr"):
            el = ppr.find(f"{W}{tag}")
            if el is not None:
                ppr.remove(el)
    added_jc = None
    if default_jc and (ppr is None or ppr.find(f"{W}jc") is None):
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            new_p.insert(0, ppr)
        _ppr_insert_ordered(ppr, etree.Element(f"{W}jc", {f"{W}val": default_jc}))
        added_jc = default_jc
    stripped = 0
    for br in list(new_p.iter(f"{W}br")):
        if (br.get(f"{W}type") or "") == "page":
            br.getparent().remove(br)
            stripped += 1
    return new_p, stripped, added_jc


def _body_column_emu(body) -> int | None:
    """版心宽（EMU）：判「文本框是否居中」的尺子。取不到返回 None（不猜）。"""
    sect = body.find(f"{W}sectPr")
    if sect is None:
        return None
    pg = sect.find(f"{W}pgSz")
    mar = sect.find(f"{W}pgMar")
    if pg is None or mar is None:
        return None
    try:
        w = (int(pg.get(f"{W}w") or 0) - int(mar.get(f"{W}left") or 0)
             - int(mar.get(f"{W}right") or 0))
    except (TypeError, ValueError):
        return None
    return w * 635 if w > 0 else None


def _shape_h_align(shape, col_w_emu: int | None) -> str | None:
    """文本框的水平位置 → 题注段落对齐（原段落没有 `w:jc` 时用它，A2「对齐」继承）。

    * `wp:positionH/wp:align` 直接给 left/center/right → 照搬；
    * 只给 `wp:posOffset` 时按「文本框中心 vs 版心中心」判，容差 10% 版心宽；
    * 量不出来（inside/outside、relativeFrom=page、无锚点）→ None，保持继承。
    """
    anc = shape.find(f".//{_WP}anchor")
    if anc is None:
        return None
    h = anc.find(f"{_WP}positionH")
    if h is None:
        return None
    al = h.find(f"{_WP}align")
    if al is not None and (al.text or "").strip():
        v = (al.text or "").strip().lower()
        return v if v in ("left", "center", "right") else None
    off = h.find(f"{_WP}posOffset")
    ext = anc.find(f"{_WP}extent")
    if off is None or ext is None or not col_w_emu:
        return None
    if (h.get("relativeFrom") or "") not in ("margin", "column"):
        return None
    try:
        x = int((off.text or "").strip())
        cx = int(ext.get(f"cx") or 0)
    except (TypeError, ValueError):
        return None
    if cx <= 0:
        return None
    if abs((x + cx / 2) - col_w_emu / 2) <= col_w_emu * 0.10:
        return "center"
    return None


def _shape_v_offset(shape) -> int | None:
    """形状的纵向偏移（EMU，relativeFrom=paragraph 时才可比）。

    同一个段落里放了多张图 + 多条题注时（真机 图5.3/图5.4），XML 里文本框的先后
    顺序与它们在页面上的上下顺序可以相反——抽流后必须按纵向位置排，否则题注会
    跟着 XML 顺序读成倒序（实测：抽流前「图5.3 在上、图5.4 在下」，抽流后变成
    5.4 前 5.3 后）。
    """
    anc = shape.find(f".//{_WP}anchor")
    if anc is None:
        return None
    v = anc.find(f"{_WP}positionV")
    if v is None or (v.get("relativeFrom") or "") != "paragraph":
        return None
    off = v.find(f"{_WP}posOffset")
    if off is None:
        return None
    try:
        return int((off.text or "").strip())
    except (TypeError, ValueError):
        return None


def _drop_shape(shape, stat: dict) -> None:
    """拆掉整个形状（Choice + Fallback 一起走），顺手清掉只剩外壳的 run。"""
    parent = shape.getparent()
    if parent is None:
        return
    parent.remove(shape)
    stat["shapes_removed"] += 1
    if parent.tag != f"{W}r":
        return
    if all(etree.QName(c).localname in _RUN_SHELL_TAGS for c in parent):
        gp = parent.getparent()
        if gp is not None:
            gp.remove(parent)
            stat["runs_removed"] += 1


def _page_break_after_picture(p) -> bool:
    """段内的分页符是否落在图之后（那样图与题注会被拆到两页上）。"""
    seen_pic = False
    for el in p.iter():
        name = etree.QName(el).localname
        if name in ("blip", "imagedata", "object"):
            seen_pic = True
        elif name == "br" and (el.get(f"{W}type") or "") == "page" and seen_pic:
            return True
    return False


def extract_caption_textboxes(path, params: dict | None = None) -> dict:
    """把锚在图对象文本框里的题注抽成流内题注段（EZI-0013，A1/A2/A5）。

    params:
      apply: false 只出统计不改盘（默认 true，结构段直调）
      kinds: figure / table / all（默认 all）

    行为：

    * 只认**题注文本**（复用 `_is_caption_para2`：图/表 + 编号、短段、无句读、
      非正文引用段）；文本框里的其他文字不动、只计数；
    * `mc:AlternateContent` 的两份副本只取一份（Choice，其次 Fallback），
      Fallback 副本丢弃并计入 `duplicate_branches_ignored`；
    * 题注段落 **deepcopy** 成流内段落后插到图对象段之后：锚段自身有真图 → 紧跟
      该段；锚段只有文本框 → 找最近的前一个真图段（15/30 实测是这种形态）；
      前面没有图时退到后一个对象块（图或表，题在对象上方——表题的常态），
      仍找不到就如实跳过（不猜）；
    * 形状拆掉，题注不会再被画两遍；锚段若只剩空壳则连同删除（只删「只有 pPr
      + 空 run」的段，有字/有图/有书签/有域的段一律不碰）；
    * 题注段里的 `w:br type="page"` / `pageBreakBefore` 一并剥掉——A2 要求图段
      与题注段之间不得有分页符；图段自身分页符只上报不改（不是本动作引入的）。
    """
    path = Path(path)
    params = params or {}
    apply_ = params.get("apply", True) is not False
    kinds = set(_norm_kinds(params.get("kinds")))
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False, "reason": "no_body"}
    cover = _cover_paragraphs(root, _read_styles_map(path))
    before = _caption_binding_counts(root)
    stat: dict = {
        "captions_found": 0, "extracted": 0, "figure": 0, "table": 0,
        "duplicate_branches_ignored": 0, "shapes_removed": 0, "runs_removed": 0,
        "paragraphs_removed": 0, "skipped_non_caption": 0,
        "attached_after_figure": 0, "attached_before_object": 0,
        "skipped_no_figure": 0, "page_breaks_stripped": 0,
        "jc_from_shape": 0, "reordered_by_position": 0,
        "page_break_between": 0, "binding_before": before,
    }

    col_w = _body_column_emu(body)
    blocks = _flow_blocks(body)
    for i, el in enumerate(blocks):
        if el.tag != f"{W}p":
            continue
        if cover and any(el is c for c in cover):
            continue
        shapes = _shape_containers(el)
        if not shapes:
            continue

        found: list[tuple[str, object, str | None, int | None]] = []  # kind, 段, 对齐, 纵向
        for shape in shapes:
            branch = _shape_branch(shape)
            inner = _txbx_paragraphs(branch)
            h_jc = _shape_h_align(shape, col_w)
            v_off = _shape_v_offset(shape)
            got: list[tuple[str, object, str | None, int | None]] = []
            for ip in inner:
                text = _visible_txt(ip).strip()
                if not text:
                    continue
                if not _is_caption_para2(text):
                    stat["skipped_non_caption"] += 1
                    continue
                kind = _kind_of(text)
                if kind not in kinds:
                    continue
                got.append((kind, ip, h_jc, v_off))
            if not got:
                continue
            # Choice / Fallback 是同一份内容：只取一份，副本计数上报（A1 禁重复）
            if (shape.tag == f"{_MC}AlternateContent"
                    and shape.find(f"{_MC}Fallback") is not None):
                stat["duplicate_branches_ignored"] += len(got)
            found.extend(got)

        if not found:
            continue
        stat["captions_found"] += len(found)

        # 同一段落多个形状：按纵向位置排（XML 顺序可能与页面上下顺序相反）
        if len(found) > 1 and all(x[3] is not None for x in found):
            ordered = sorted(found, key=lambda x: x[3])
            if [x[1] for x in ordered] != [x[1] for x in found]:
                stat["reordered_by_position"] += 1
            found = ordered

        # 定位对象段：锚段自身有真图 → 自己；否则最近的前一个真图段（题在图下）；
        # 再否则后一个对象块（图或表，题在对象上方——表题的常态）
        anchor = el if _block_has_picture(el) else None
        before_fig = False
        if anchor is None:
            for j in range(i - 1, -1, -1):
                if _block_has_picture(blocks[j]):
                    anchor = blocks[j]
                    break
        if anchor is None:
            for j in range(i + 1, len(blocks)):
                if _block_has_picture(blocks[j]) or blocks[j].tag == f"{W}tbl":
                    anchor = blocks[j]
                    before_fig = True
                    break
        if anchor is None:
            stat["skipped_no_figure"] += len(found)
            continue

        new_paras = []
        for _kind, ip, h_jc, _v in found:
            new_p, stripped, added_jc = _flow_safe_caption(ip, h_jc)
            new_paras.append(new_p)
            stat["page_breaks_stripped"] += stripped
            if added_jc:
                stat["jc_from_shape"] += 1
        if before_fig:
            for new_p in reversed(new_paras):
                anchor.addprevious(new_p)
            stat["attached_before_object"] += len(new_paras)
        else:
            cur = anchor
            for new_p in new_paras:
                cur.addnext(new_p)
                cur = new_p
            stat["attached_after_figure"] += len(new_paras)

        for kind, _ip, _jc, _v in found:
            stat[kind] += 1
            stat["extracted"] += 1

        for shape in shapes:
            _drop_shape(shape, stat)

        # 锚段不在图段上：拆完形状就只剩空壳 → 删掉（有内容的一律保留）
        if anchor is not el and el.getparent() is not None and _droppable_blank(el):
            el.getparent().remove(el)
            stat["paragraphs_removed"] += 1
        if _page_break_after_picture(anchor):
            stat["page_break_between"] += 1

    after = _caption_binding_counts(root)
    stat["binding_after"] = after if apply_ else before
    stat["embedded_before"] = before["embedded"]
    stat["embedded_after"] = after["embedded"]
    stat["flow_figure_captions_before"] = before["figure"]["total"]
    stat["flow_figure_captions_after"] = after["figure"]["total"]
    changed = stat["extracted"] > 0
    if apply_ and changed:
        _write_part(path, "word/document.xml", root)
    stat["changed"] = changed and apply_
    stat["applied"] = apply_
    if not apply_:
        stat["message"] = "只出统计（未落盘）；确认后 apply=true 执行。"
    return stat


def verify_no_embedded_captions(path, params: dict | None = None) -> bool:
    """复查：正文里不再有「图/表 + 编号」的文本框题注（抽流后应全在流内）。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return False
        for p in body.iter(f"{W}p"):
            if not _in_textbox(p):
                continue
            if _is_caption_para2(_visible_txt(p)):
                return False
        return True
    except Exception:  # noqa: BLE001
        return False


def bind_caption_objects(path, params: dict | None = None) -> dict:
    """图题绑图段、表题绑表段（EZI-0011）：按对象位置定 keepNext 方向。

    params:
      apply: false 只出统计不改盘（默认 true，结构修复段直调）

    行为（只在**流内**题注上做——图形文本框里的题注字不属于版面流，不猜）：

    * 图题（``^图``）：拆掉图与题之间的纯空段 → 给**前一个图片段落**补
      ``keepNext``（+ ``keepLines``）→ 去掉图题段自己身上的 ``keepNext``
      （它绑的是题**后面**的正文，方向反了，会夸大「图题孤悬」）；
    * 表题（``^表``）：拆掉表与题之间的纯空段 → 给**表题段落**补 ``keepNext``
      （题在上，题绑表）。
    * 封面区题注不动；``w:cantSplit`` / ``w:tblHeader`` 一个都不碰，并在结果里
      带上修前/修后的行属性计数当证据（A4）。
    """
    path = Path(path)
    params = params or {}
    apply_ = params.get("apply", True) is not False
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False, "reason": "no_body"}
    cover = _cover_paragraphs(root, _read_styles_map(path))
    before_tbl = _table_flag_counts(root)
    before_cap = _caption_binding_counts(root)

    stat: dict = {
        "figures": 0, "figures_bound": 0, "tables": 0, "tables_bound": 0,
        "blank_paras_removed": 0, "caption_keepnext_removed": 0,
        "keep_next_added": 0, "keep_lines_added": 0,
        "skipped_no_figure": 0, "skipped_no_table": 0, "skipped_cover": 0,
        "embedded_figure_captions": before_cap["embedded"]["figure"],
        "embedded_table_captions": before_cap["embedded"]["table"],
        "table_flags_before": before_tbl, "binding_before": before_cap,
    }

    def _is_cover(el) -> bool:
        return bool(cover) and any(el is c for c in cover)

    # ① 拆空段：keepNext 链过不了空段，空段会让题注照样滑到下一页
    blocks = _flow_blocks(body)
    drop: list = []
    for i, el in enumerate(blocks):
        if el.tag != f"{W}p" or _is_cover(el):
            continue
        m = _CAPTION_TEXT_RE.match(_caption_flow_text(el))
        if m is None:
            continue
        if m.group(1) == "图":
            j = i - 1
            blanks: list = []
            while j >= 0 and _droppable_blank(blocks[j]):
                blanks.append(blocks[j])
                j -= 1
            if blanks and j >= 0 and _block_has_graphic(blocks[j]):
                drop.extend(blanks)
        else:
            j = i + 1
            blanks = []
            while j < len(blocks) and _droppable_blank(blocks[j]):
                blanks.append(blocks[j])
                j += 1
            if blanks and j < len(blocks) and blocks[j].tag == f"{W}tbl":
                drop.extend(blanks)
    for p in drop:
        if p.getparent() is not None:
            p.getparent().remove(p)
            stat["blank_paras_removed"] += 1

    # ② 绑：图题 → 前面的图段；表题 → 自己
    blocks = _flow_blocks(body)
    for i, el in enumerate(blocks):
        if el.tag != f"{W}p" or _is_cover(el):
            continue
        m = _CAPTION_TEXT_RE.match(_caption_flow_text(el))
        if m is None:
            continue
        if m.group(1) == "图":
            stat["figures"] += 1
            prev = blocks[i - 1] if i else None
            if not _block_has_graphic(prev):
                stat["skipped_no_figure"] += 1
                continue
            if _ensure_ppr_flag(prev, "keepNext"):
                stat["keep_next_added"] += 1
            if _ensure_ppr_flag(prev, "keepLines"):
                stat["keep_lines_added"] += 1
            stat["figures_bound"] += 1
            # 方向修正：图题的 keepNext 绑的是题后面那一段，跟图无关
            if _drop_ppr_flag(el, "keepNext"):
                stat["caption_keepnext_removed"] += 1
        else:
            stat["tables"] += 1
            nxt = blocks[i + 1] if i + 1 < len(blocks) else None
            if nxt is None or nxt.tag != f"{W}tbl":
                stat["skipped_no_table"] += 1
                continue
            if _ensure_ppr_flag(el, "keepNext"):
                stat["keep_next_added"] += 1
            stat["tables_bound"] += 1

    after_tbl = _table_flag_counts(root)
    stat["table_flags_after"] = after_tbl
    stat["table_flags_preserved"] = after_tbl == before_tbl
    changed = bool(stat["blank_paras_removed"] or stat["keep_next_added"]
                   or stat["keep_lines_added"] or stat["caption_keepnext_removed"])
    if apply_ and changed:
        _write_part(path, "word/document.xml", root)
    stat["changed"] = changed and apply_
    stat["applied"] = apply_
    stat["binding_after"] = (_caption_binding_counts(root) if apply_ else before_cap)
    if not apply_:
        stat["message"] = "只出统计（未落盘）；确认后 apply=true 执行。"
    return stat


def verify_caption_binding(path, params: dict | None = None) -> bool:
    """复查：图题绑图段、表题绑表，且图题段自己没有 keepNext（方向对）。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return False
        blocks = _flow_blocks(body)
        figs = tabs = 0
        for i, el in enumerate(blocks):
            if el.tag != f"{W}p":
                continue
            m = _CAPTION_TEXT_RE.match(_caption_flow_text(el))
            if m is None:
                continue
            if m.group(1) == "图":
                prev = blocks[i - 1] if i else None
                if not _block_has_graphic(prev):
                    continue          # 不相邻：不判（与打分器同口径）
                figs += 1
                if not _p_has(prev, "keepNext"):
                    return False
                if _p_has(el, "keepNext"):
                    return False      # 方向错：图题绑了后面
            else:
                nxt = blocks[i + 1] if i + 1 < len(blocks) else None
                if nxt is None or nxt.tag != f"{W}tbl":
                    continue
                tabs += 1
                if not _p_has(el, "keepNext"):
                    return False
        return True
    except Exception:  # noqa: BLE001
        return False
