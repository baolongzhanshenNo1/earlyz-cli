"""W8 Toggle TeX + 纯文本公式定型（B/C 两通路，2026-10-06 开工）。

C — toggle_tex_equations：文档里的 `$…$` / `$$…$$` / `\\[…\\]` 文本 → MathType 对象。
    边界文档《公式与LaTeX实现边界-2026-09-29》W8（MathType Toggle TeX 语义）：
    显式定界符即用户意图，自动转换。产物一律 DSMT4 对象，不产 OMML（用户定案
    「不转成 Word 原生公式，原样保留」）。

B — convert_text_equations：人工敲的纯文本公式（`Rn = B log₂(1 + …)` 这类）→ 对象。
    E9 原则「识别结果先给用户确认再替换」：apply=false 只出提案（校验 LaTeX 可
    编译、定位命中情况），用户确认后 apply=true 才落盘。LaTeX 由 Agent 逐符号
    忠实原文书写，禁止臆造记法。

防误伤：货币 `$3.9 … $12.9` 这类非数学 `$…$` 对不做转换（_looks_like_math 守卫），
跳过项在返回里如实列出（skip_reason），由上层通报。
"""
from __future__ import annotations

import re
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# ── 定界符 ─────────────────────────────────────────────────────────────────
_DISPLAY_PATTERNS = (
    re.compile(r"\\\[(.+?)\\\]", re.DOTALL),   # \[ ... \]
    re.compile(r"\$\$(.+?)\$\$", re.DOTALL),   # $$ ... $$
)
_INLINE_PATTERN = re.compile(r"(?<!\$)\$(?!\$)([^$\n]{1,400}?)(?<!\$)\$(?!\$)")


def _looks_like_code(s: str) -> bool:
    """片段像代码而不是数学（2026-11-07）。

    `{foreach $item_list as $item}` 这类模板代码里的 `$…$` 会被当成行内公式，
    转成 MathType 对象后正文文字整段消失（软件编码说明书 content 项实测丢 112 字）。
    这里用代码特征把这类跨度挡在公式转换之外：模板变量、比较/箭头运算、
    代码关键字、`a as b` 形态。宁可漏转公式，也不能吞掉正文。
    """
    s = (s or "").strip()
    if not s:
        return False
    if re.search(r"\$[A-Za-z_]|==|!=|<=|>=|\+\+|->|::|<[A-Za-z/]|/>|;", s):
        return True
    if re.search(r"[{}]\s*[{}]|\"|'", s):
        return True
    if re.search(r"\b(?:foreach|endforeach|if|else|elseif|endif|function|class|var|let|"
                 r"const|return|echo|print|import|from|def|lambda|include|require)\b", s, re.I):
        return True
    if re.match(r"^[A-Za-z_][A-Za-z0-9_]*\s+(?:as|in|of)\b", s):
        return True
    return False


def _looks_like_math(s: str) -> bool:
    """内容像数学才转换——把 `$3.9 … $12.9`（货币）挡在外面。"""
    s = (s or "").strip()
    if not s:
        return False
    if _looks_like_code(s):
        return False
    if any(c in s for c in "=\\^_{}∫∑√±×÷≤≥∈⊂≈≠∝∂∇"
                          "αβγδεζηθικλμνξπρστυφχψω"):
        return True
    if re.search(r"[A-Za-z0-9]\s*[_^]\s*\{?[A-Za-z0-9]", s):   # R_n / x^2 / a_{t}
        return True
    if len(s) <= 24 and re.search(r"[A-Za-z]\d|\d[A-Za-z]", s):  # 3x / x3 / Tn
        return True
    return False


# ── 通用小工具（与 equation_dsmt4 同族）────────────────────────────────────
def _para_full_text(p) -> str:
    return "".join(t.text or "" for t in p.findall(f".//{W}t"))


def _paragraph_text_map(p):
    text_parts, mapping, pos = [], [], 0
    for r in p.findall(f"{W}r"):
        r_text = "".join(t.text or "" for t in r.findall(f"{W}t"))
        if r_text:
            mapping.append((r, pos, pos + len(r_text)))
            text_parts.append(r_text)
            pos += len(r_text)
    return "".join(text_parts), mapping


def _contains_object(p) -> bool:
    for el in p.iter():
        if el.tag.endswith("}object") or el.tag.endswith("}oMath"):
            return True
    return False


# ── C：Toggle TeX ──────────────────────────────────────────────────────────
def _toggle_tex_in(root_dir: Path, params: dict = None) -> dict:
    """`$…$` / `$$…$$` / `\\[…\\]` 文本 → DSMT4 对象（W8，解包目录口径）。

    - 整段只有定界公式（可带段尾编号）→ 整段换公式对象（保留编号、居中）。
    - 行内 `$…$`（同一 run 内）→ 就地拆 run 插入行内对象，不动上下文。
    - 跨 run 的行内公式、非数学 `$…$` 对 → 不动，如实列 skip_reason。
    """
    from earlyz.actions.equation_dsmt4 import (_make_dsmt4_object,
                                              set_equation_dsmt4)

    doc_part = root_dir / "word" / "document.xml"

    display_done, inline_done, skipped = [], [], []

    # ── 第一趟：整段定界公式 ────────────────────────────────────────────
    # set_equation_dsmt4 每次自读盘自写盘，逐条串行调用才安全；本趟结束前
    # 不能在内存树上做别的编辑（会被写回覆盖）。
    tree = etree.parse(str(doc_part))
    whole_targets = []
    for p in tree.getroot().find(f"{W}body").iter(f"{W}p"):
        text = _para_full_text(p).strip()
        if "$" not in text and "\\[" not in text:
            continue
        for pat in _DISPLAY_PATTERNS:
            mm = pat.fullmatch(text)
            if mm and _looks_like_math(mm.group(1)):
                whole_targets.append((text, mm.group(1).strip()))
                break
    for text, latex in whole_targets:
        try:
            set_equation_dsmt4(root_dir, {"match_text": text, "latex": latex,
                                          "font_size_pt": (params or {}).get("font_size_pt")})
            display_done.append({"text": text[:40], "latex": latex})
        except Exception as exc:  # 段文本不唯一等 → 如实留痕
            skipped.append({"text": text[:40], "skip_reason": f"整段转换失败: {exc}"})

    # ── 第二趟：行内片段（重新读盘，一棵树上改完一次性写回）─────────────
    tree = etree.parse(str(doc_part))
    root = tree.getroot()
    plan = []   # [(p, [(st, en, latex, kind), ...] 右→左)]
    for p in root.find(f"{W}body").iter(f"{W}p"):
        text = _para_full_text(p)
        if "$" not in text and "\\[" not in text:
            continue
        spans = []
        for pat in _DISPLAY_PATTERNS:
            for mm in pat.finditer(text):
                spans.append((mm.start(), mm.end(), mm.group(1).strip(), "display"))
        for mm in _INLINE_PATTERN.finditer(text):
            spans.append((mm.start(), mm.end(), mm.group(1).strip(), "inline"))
        if spans:
            plan.append((p, text, sorted(spans, key=lambda s: -s[0])))
    for p, text, spans in plan:
        for st, en, latex, kind in spans:
            if not _looks_like_math(latex):
                skipped.append({"text": text[max(0, st - 8):en + 8][:40],
                                "skip_reason": "非数学内容（疑似货币/普通文本）"})
                continue
            ok = _inline_insert(root_dir, p, st, en, latex, kind,
                                font_size_pt=(params or {}).get("font_size_pt"))
            if ok is True:
                inline_done.append({"text": text[max(0, st - 8):en + 8][:40],
                                    "latex": latex})
            else:
                skipped.append({"text": text[max(0, st - 8):en + 8][:40],
                                "skip_reason": ok or "未转换"})
    if inline_done:
        with open(doc_part, "wb") as f:
            f.write(etree.tostring(root, xml_declaration=True, encoding="UTF-8"))

    changed = bool(display_done or inline_done)
    return {"changed": changed,
            "display_converted": len(display_done),
            "inline_converted": len(inline_done),
            "details": display_done + inline_done,
            "skipped": skipped}


def toggle_tex_equations(path, params: dict = None) -> dict:
    """公开入口：.docx 文件与解包目录双口径（见 equation_dsmt4._work_root）。"""
    from earlyz.actions.equation_dsmt4 import _work_root
    with _work_root(path) as root_dir:
        return _toggle_tex_in(root_dir, params)


def _inline_insert(root_dir: Path, p, st: int, en: int, latex: str, kind: str,
                   font_size_pt: float | None = None):
    """把 [st,en) 文本片段换成行内公式对象（不居中不动段落格式）。

    返回 True=成功；str=失败原因（如实上报，不静默）。
    """
    from earlyz.actions.equation_dsmt4 import _make_dsmt4_object, _center_tab_cap_pt

    text, mapping = _paragraph_text_map(p)
    hit = None
    for r, rs, re_ in mapping:
        if rs <= st and en <= re_:
            hit = (r, rs, re_)
            break
    if hit is None:
        return "公式被 Word 拆成多个格式片段，未自动转换"
    r, rs, _ = hit
    t_elem = r.find(f"{W}t")
    if t_elem is None or not t_elem.text:
        return "格式片段无文本，未自动转换"
    local_st, local_en = st - rs, en - rs
    before, after = t_elem.text[:local_st], t_elem.text[local_en:]

    obj, _info = _make_dsmt4_object(root_dir, p.getroottree().getroot(), latex,
                                    max_width_pt=_center_tab_cap_pt(p),
                                    font_size_pt=font_size_pt)
    obj_run = etree.Element(f"{W}r")
    obj_run.append(obj)

    t_elem.text = before
    t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    anchor = r
    anchor.addnext(obj_run)
    if after:
        after_run = etree.Element(f"{W}r")
        rpr = r.find(f"{W}rPr")
        if rpr is not None:
            after_run.append(etree.fromstring(etree.tostring(rpr)))
        after_t = etree.SubElement(after_run, f"{W}t")
        after_t.text = after
        after_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        obj_run.addnext(after_run)
    return True


# ── B：纯文本公式定型（确认后落盘）────────────────────────────────────────
def _find_single_target(pool, needle: str):
    """定位唯一目标段（与 set_equation_dsmt4 同口径，含 [:24] 前缀兜底）。

    返回 (p, used_needle, error)；error=None 且 p 非 None = 成功。
    """
    cands = [needle]
    if len(needle) > 24:
        cands.append(needle[:24])
    for cand in cands:
        if not cand:
            continue
        hits = [q for q in pool if cand in _para_full_text(q)]
        if len(hits) == 1:
            return hits[0], cand, None
        if len(hits) > 1:
            return None, cand, f"命中 {len(hits)} 段（须恰好 1 段；请提供更长的 match_text）"
    return None, needle, "未在文档中找到该文本（若此前已转换过则属正常）"


def _match_mode(p, needle: str):
    """命中形态判定（A5, 2026-10-08）。

    display = 整段式：needle 前只剩空白/制表位，后只剩空白 + 段尾编号
              （(N)/（N）/(2-1) 等）→ 整段换对象（沿用 set_equation_dsmt4）。
    inline  = 句内式：前后还有真实正文（「…本地执行时延为 T_n^loc = C_n/f_n^loc。 (4)」）
              → 只把公式跨度原位替换为对象，编号留在原位、版式不动。

    返回 (mode, (st, en), note)；mode=None 即无法确定唯一跨度（note=原因）。
    """
    from earlyz.actions.new_actions import _EQ_NUM_ANY_RE

    text, _ = _paragraph_text_map(p)
    n = text.count(needle)
    if n == 0:
        return None, None, "段落中已无该文本（可能已转换）"
    if n > 1:
        return None, None, f"该文本在段落中出现 {n} 次，无法确定唯一跨度"
    idx = text.find(needle)
    before, after = text[:idx], text[idx + len(needle):]
    if before.strip(" \t\u3000") == "":
        res = after.strip(" \t\u3000")
        m = _EQ_NUM_ANY_RE.match(res) if res else None
        if res == "" or (m is not None and not m.group("body").strip()):
            return "display", (idx, idx + len(needle)), ""
    return "inline", (idx, idx + len(needle)), ""


# ── B：纯文本公式定型（确认后落盘）────────────────────────────────────────
def _convert_text_in(root_dir: Path, params: dict = None) -> dict:
    """人工敲的纯文本公式 → DSMT4 对象（E9：先提案后替换；解包目录口径）。

    两种命中形态（A5）：
    - **整段式**（display）：段落=公式文本（可带段尾编号）→ 整段换对象
      （编号保留、居中制表位版式——沿用 set_equation_dsmt4）；
    - **句内式**（inline）：公式粘在正文句子里（「…为 T_n^loc = C_n/f_n^loc。 (4)」）
      → **只把公式跨度原位替换为对象**（编号留在原位、版式不动；复用 _inline_insert）。

    params:
      conversions — [{"match_text": 公式段原文片段, "latex": 目标 LaTeX}]
      apply       — false（默认）= 只校验出提案（含命中形态/跨度预览）；true = 替换
    返回每条 {match_text, latex, ok, applied?, mode, span, error?}，供用户过目。
    """
    from earlyz.mathtype import latex_to_equation
    from earlyz.actions.equation_dsmt4 import set_equation_dsmt4

    convs = list(params.get("conversions") or [])
    apply = bool(params.get("apply"))
    if not convs:
        # 2026-10-07 真机事故兜底：空 conversions 曾静默空转 changed=false，
        # 被 LLM 转述成「无公式被匹配」。显式给出指引，杜绝假成功。
        return {"changed": False, "applied": False, "total": 0, "ok_count": 0,
                "conversions": [],
                "error": ("conversions 为空：需逐条提供 {match_text: 公式段原文, "
                          "latex: 目标LaTeX}（原文取自 inspect_document 的 "
                          "equations(kind=plain) 条目）")}
    doc_part = root_dir / "word" / "document.xml"
    tree = etree.parse(str(doc_part))
    pool = tree.getroot().findall(f".//{W}p")

    rows = []
    display_jobs = []   # (row, used_needle, latex)
    inline_jobs = []    # (row, used_needle, latex)
    for c in convs:
        match_text = (c.get("match_text") or "").strip()
        latex = (c.get("latex") or "").strip()
        row = {"match_text": match_text[:64], "latex": latex, "ok": False, "applied": False}
        if not match_text or not latex:
            row["error"] = "缺 match_text 或 latex"
            rows.append(row)
            continue
        try:
            latex_to_equation(latex)   # 预检：不支持的构造显式报错，不降级
        except Exception as exc:
            row["error"] = f"LaTeX 编译失败: {exc}"
            rows.append(row)
            continue
        target, used, err = _find_single_target(pool, match_text)
        if err:
            row["error"] = err
            rows.append(row)
            continue
        mode, span, note = _match_mode(target, used)
        if mode is None:
            row["error"] = note
            rows.append(row)
            continue
        row["mode"] = mode
        row["span"] = _para_full_text(target)[max(0, span[0] - 10):span[1] + 10][:60]
        if not apply:
            row["ok"] = True
            rows.append(row)
            continue
        rows.append(row)
        if mode == "display":
            display_jobs.append((row, used, latex))
        else:
            inline_jobs.append((row, used, latex))

    # 整段式：逐条串行落盘（set_equation_dsmt4 每次自读自写，逐次调用才安全）
    for row, needle, latex in display_jobs:
        try:
            r = set_equation_dsmt4(root_dir, {"match_text": needle, "latex": latex,
                                              "font_size_pt": (params or {}).get("font_size_pt")})
            row.update(ok=True, applied=True,
                       number_kept=r.get("number_kept"),
                       ole_size=r.get("ole_size"))
        except Exception as exc:
            row["error"] = str(exc)

    # 句内式（A5）：重读盘、一棵树上改完一次性写回（与 toggle 行内趟同法；
    # 同段多处右→左处理防坐标漂移）
    if inline_jobs:
        tree2 = etree.parse(str(doc_part))
        root2 = tree2.getroot()
        pool2 = root2.findall(f".//{W}p")
        jobs = []
        for row, needle, latex in inline_jobs:
            p2, used2, err = _find_single_target(pool2, needle)
            if err:
                row["error"] = err
                continue
            mode2, span2, note2 = _match_mode(p2, used2)
            if mode2 is None:
                row["error"] = note2
                continue
            jobs.append((row, p2, span2, latex))
        jobs.sort(key=lambda j: (id(j[1]), -j[2][0]))
        dirty = False
        for row, p2, (st, en), latex in jobs:
            ok = _inline_insert(root_dir, p2, st, en, latex, "inline",
                                font_size_pt=(params or {}).get("font_size_pt"))
            if ok is True:
                row.update(ok=True, applied=True)
                dirty = True
            else:
                row["error"] = ok or "未转换"
        if dirty:
            with open(doc_part, "wb") as f:
                f.write(etree.tostring(root2, xml_declaration=True, encoding="UTF-8"))

    return {"changed": any(r["applied"] for r in rows),
            "applied": apply,
            "total": len(rows),
            "ok_count": sum(1 for r in rows if r["ok"]),
            "conversions": rows}


def convert_text_equations(path, params: dict = None) -> dict:
    """公开入口：.docx 文件与解包目录双口径（见 equation_dsmt4._work_root）。"""
    from earlyz.actions.equation_dsmt4 import _work_root
    with _work_root(path) as root_dir:
        return _convert_text_in(root_dir, params)


# ── 复查器（catalog 注册用）────────────────────────────────────────────────
def verify_toggle_tex_equations(path, params: dict = None) -> bool:
    """复查：文档里不再残留「像数学」的定界符文本。"""
    from earlyz.actions.equation_dsmt4 import _read_xml_any
    root = _read_xml_any(path, "word/document.xml")
    for p in root.iter(f"{W}p"):
        text = _para_full_text(p)
        for pat in _DISPLAY_PATTERNS:
            for mm in pat.finditer(text):
                if _looks_like_math(mm.group(1)):
                    return False
        for mm in _INLINE_PATTERN.finditer(text):
            if _looks_like_math(mm.group(1)):
                return False
    return True


def verify_convert_text_equations(path, params: dict = None) -> bool:
    """复查：apply=true 后各公式段的原文文本已被对象替换（不再残留）。

    apply=false 是提案阶段，恒真。注意转换后 match_text 不应再命中段落文本
    （set_equation_dsmt4 的二进制幂等兜底除外——文本消失即视为已转）。
    """
    if not (params or {}).get("apply"):
        return True
    from earlyz.actions.equation_dsmt4 import _read_xml_any
    root = _read_xml_any(path, "word/document.xml")
    texts = [_para_full_text(p) for p in root.iter(f"{W}p")]
    for c in (params or {}).get("conversions") or []:
        mt = (c.get("match_text") or "").strip()
        if mt and any(mt in t for t in texts):
            return False
    return True


# ── 公式段「公式—编号」连接样式（点线引导；用户点名才加，默认不加）──────────
# 2026-10-07 用户定案：模板沉默的版式变体——「让用户自行选择，默认不加」。
# 实现为独立动作（agent 可调），交付件默认保持模板口径（制表位无前导符）；
# 用户明确要求「公式和编号之间加点线/虚线/去掉」时由 agent 调本动作。
_LEADER_STYLES = {"dot": "dot", "dash": "dash", "none": None}
_M_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


def _equation_right_tabs(p, ppr):
    """公式段右制表位元素列表（只认编译口径的公式段，见下方调用处）。"""
    if ppr is None:
        return []
    return [t for tabs in ppr.findall(f"{W}tabs")
            for t in tabs if t.get(f"{W}val") == "right"]


def set_equation_leader(path, params: dict) -> dict:
    """公式与其编号之间的连接样式：右制表位前导符（点线/虚线/无）。

    style 必填（让用户自行选择，禁默认悄悄加）：
      "dot"=点线（……）/"dash"=虚线（----）/"none"=去掉（回模板口径）。
    match_text 可选：只处理含该文本的公式段（点名单条公式时用）。
    只动公式段右制表位的 leader 属性——制表位位置/对齐、字体等一律不碰；
    公式段判定与编译定型同一口径（_equation_para_context → equation_display）。
    """
    from earlyz.actions.new_actions import (_equation_para_context, _read_part,
                                            _read_styles_map, _write_part)
    path = Path(path)
    params = params or {}
    style = str(params.get("style") or "").strip().lower()
    if style not in _LEADER_STYLES:
        return {"changed": False,
                "error": 'style 必填："dot"=点线 / "dash"=虚线 / "none"=去掉点线'}
    match_text = str(params.get("match_text") or "").strip()
    want = _LEADER_STYLES[style]
    root = _read_part(path, "word/document.xml")
    style_names = _read_styles_map(path)
    stat = {"equations_seen": 0, "styled": 0, "cleared": 0, "unchanged": 0,
            "skipped_filtered": 0, "skipped_no_right_tab": 0}
    changed = False
    for p in root.iter(f"{W}p"):
        role, ppr = _equation_para_context(p, style_names, _M_NS)
        if role != "equation_display":
            continue
        stat["equations_seen"] += 1
        if match_text and match_text not in _para_full_text(p):
            stat["skipped_filtered"] += 1
            continue
        right_tabs = _equation_right_tabs(p, ppr)
        if not right_tabs:
            stat["skipped_no_right_tab"] += 1
            continue
        touched = False
        for t in right_tabs:
            cur = t.get(f"{W}leader")
            if want is None:
                if cur is not None:
                    del t.attrib[f"{W}leader"]
                    touched = True
            elif cur != want:
                t.set(f"{W}leader", want)
                touched = True
        if touched:
            stat["styled" if want else "cleared"] += 1
            changed = True
        else:
            stat["unchanged"] += 1
    if changed:
        _write_part(path, "word/document.xml", root)
    out = {"changed": changed, "style": style}
    out.update(stat)
    return out


def verify_equation_leader(path, params: dict = None) -> bool:
    """复查：所有（可选过滤的）公式段右制表位 leader 与请求样式一致。

    style="none" 时期望「无 leader 属性」。style 非法/缺失 = 不该落盘的调用 → False。
    """
    from earlyz.actions.new_actions import (_equation_para_context, _read_part,
                                            _read_styles_map)
    params = params or {}
    style = str(params.get("style") or "").strip().lower()
    if style not in _LEADER_STYLES:
        return False
    want = _LEADER_STYLES[style]
    match_text = str(params.get("match_text") or "").strip()
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    style_names = _read_styles_map(path)
    for p in root.iter(f"{W}p"):
        role, ppr = _equation_para_context(p, style_names, _M_NS)
        if role != "equation_display":
            continue
        if match_text and match_text not in _para_full_text(p):
            continue
        for t in _equation_right_tabs(p, ppr):
            cur = t.get(f"{W}leader")
            if want is None:
                if cur is not None:
                    return False
            elif cur != want:
                return False
    return True
