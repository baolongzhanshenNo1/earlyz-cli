"""节几何对齐：把模板的分节纸型/页边距按节套到产物上（2026-11-07）。

**为什么需要**：`set_page_margins_all` 把 spec 里的一套页边距写进**每一个** sectPr。
合同类模板却有 16–118 个节、节几何各不同（封面节 right=1300、正文节 right=1680，
或书脊留白只在最后一节）。旧实现只认「模板首节/末节」两套几何，产物节数不等时
正文节必然拿到封面几何（本科论文：模板 8 节两套边距，产物 9 节目录/正文/参考文献/
致谢全用了封面那套 —— EZI-0016 实测、EZI-0018 回修）。

**口径**：只搬 `w:pgSz` 与 `w:pgMar` 的属性值，不动 `w:type`、不动页眉页脚引用、
不动分节符位置——几何对了、结构仍然由引擎自己决定。映射规则：

- 模板只有一节：全部产物节 ← 该节（等价于旧行为，保证幂等）；
- 节数相等：逐节按索引对位（判定器同口径，见 `layout_fidelity._section_margins`）；
- 节数不等：按**角色**把产物节映射到模板同类节（封面/声明/摘要/ABSTRACT/目录/
  正文/参考文献/致谢/附录）。识别不出角色的节沿用前一节的几何——正文内部子节
  不会掉回封面边距；两侧都识别不出角色时退回旧的「末节←模板末节、其余←模板首节」。

角色只由**标题段**判定（去空白/占位符后精确匹配），不猜正文语义。目录里带页码的
条目（「参考文献……………………59」）会被页码判据排除，不会把目录节认成参考文献节。

    ── 按内容角色建节（EZI-0034，本模块后半部分）────────────────────────────

`align_section_geometry` 只在**已有的**节上搬几何；产物只有一节时它无节可搬
（清洗集 43 对里 8 项页边距 FAIL + 1 项分栏 FAIL 全是这一个成因）。后半部分补上
「建节」这一层：**产物节数 = 属性变化点数**（栏数变 / 边距变 / 页码体系切换），
按 `w:type`、`w:cols`、`w:pgMar`、`w:pgSz` 逐节复制模板对应节，单节模板一律不切。

    plan_sections(doc, template)        # 只读：给出建节计划（角色块 → 模板节 → 是否切）
    apply_section_plan(doc, template)   # 落地：**只补**缺失分节符、复制属性（已有边界绝不删）

判据见 `plan_sections` docstring；OOXML 语义见 `apply_section_plan` docstring。
"""
from __future__ import annotations

import copy
import re
import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

_PGSZ_KEYS = ("w", "h", "orient")
_PGMAR_KEYS = ("top", "right", "bottom", "left", "header", "footer", "gutter")

#: 标题归一：去掉空白、全角空格、模板占位符（□）和省略号/点线。
_PLACEHOLDER_RE = re.compile(r"[\s\u3000\u00a0□■…\u2026]+")

#: 角色标题 → 规范角色名（归一后精确相等；小写用于英文）。
_ROLE_HEADINGS = {
    "声明": "declaration",
    "郑重声明": "declaration",
    "原创性声明": "declaration",
    "学术声明": "declaration",
    "学位论文原创性声明": "declaration",
    "摘要": "abstract",
    "中文摘要": "abstract",
    "abstract": "abstract_en",
    "英文摘要": "abstract_en",
    "目录": "toc",
    "参考文献": "references",
    "references": "references",
    "参考书目": "references",
    "致谢": "acknowledgements",
    "acknowledgements": "acknowledgements",
    "acknowledgments": "acknowledgements",
}

#: 「第 N 章」与「1 / 1.1 / 1．4．1」式章节标题。
_CHAPTER_RE = re.compile(r"^第[0-9一二三四五六七八九十百]+章")
_NUMBERED_RE = re.compile(r"^\d+(?:[.．]\d+){0,3}[、.．]?\s*\S")


def _read(path: Path) -> etree._Element | None:
    try:
        with zipfile.ZipFile(path) as z:
            return etree.fromstring(z.read("word/document.xml"))
    except (KeyError, zipfile.BadZipFile, OSError):
        return None


def _write(path: Path, root: etree._Element) -> None:
    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = path.with_suffix(".geotmp")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            payload = data if item.filename == "word/document.xml" else zin.read(item.filename)
            zout.writestr(item, payload)
    tmp.replace(path)


def _geom(sect) -> tuple[dict, dict]:
    sz = sect.find(f"{W}pgSz")
    mar = sect.find(f"{W}pgMar")
    sz_d = {k: sz.get(f"{W}{k}") for k in _PGSZ_KEYS} if sz is not None else {}
    mar_d = {k: mar.get(f"{W}{k}") for k in _PGMAR_KEYS} if mar is not None else {}
    return ({k: v for k, v in sz_d.items() if v is not None},
            {k: v for k, v in mar_d.items() if v is not None})


def _apply(sect, sz_d: dict, mar_d: dict) -> bool:
    changed = False
    if sz_d:
        sz = sect.find(f"{W}pgSz")
        if sz is None:
            sz = etree.SubElement(sect, f"{W}pgSz")
        for k, v in sz_d.items():
            if sz.get(f"{W}{k}") != str(v):
                sz.set(f"{W}{k}", str(v))
                changed = True
    if mar_d:
        mar = sect.find(f"{W}pgMar")
        if mar is None:
            mar = etree.SubElement(sect, f"{W}pgMar")
        for k, v in mar_d.items():
            if mar.get(f"{W}{k}") != str(v):
                mar.set(f"{W}{k}", str(v))
                changed = True
    return changed


def _norm_heading(text: str) -> str:
    return _PLACEHOLDER_RE.sub("", text or "").strip().lower()


def _section_blocks(root: etree._Element) -> list[tuple[list[str], etree._Element]]:
    """按文档顺序返回 ``[(本节段落文本, 本节 sectPr), ...]``。

    分节符（段内 ``w:pPr/w:sectPr``）结束的段落算作**上一节**的末段；body 级
    ``w:sectPr`` 结束最后一节。只取正文可见段落，页眉页脚不在 document.xml 里。
    """
    body = root.find(f"{W}body")
    if body is None:
        return []
    out: list[tuple[list[str], etree._Element]] = []
    pending: list[str] = []
    for el in body.iter():
        if not isinstance(el.tag, str):
            continue
        if el.tag == f"{W}p":
            txt = "".join(t.text or "" for t in el.iter(f"{W}t")).strip()
            ppr = el.find(f"{W}pPr")
            sp = ppr.find(f"{W}sectPr") if ppr is not None else None
            if sp is not None:
                if txt:
                    pending.append(txt)
                out.append((pending, sp))
                pending = []
            elif txt:
                pending.append(txt)
        elif el.tag == f"{W}sectPr":
            parent = el.getparent()
            if parent is not None and parent.tag == f"{W}body":
                out.append((pending, el))
                pending = []
    return out


def _detect_roles(texts: list[str], *, first: bool) -> set[str]:
    """从一节的标题段判定角色。只认短标题 + 精确匹配，不做语义猜测。"""
    roles: set[str] = set()
    for raw in texts:
        t = _norm_heading(raw)
        if not t or len(t) > 40:
            continue
        role = _ROLE_HEADINGS.get(t)
        if role:
            roles.add(role)
            continue
        if t.startswith("附录") and len(t) <= 8:
            roles.add("appendix")
            continue
        # EZI-0075：附件表单区标题（「附件 3 ：」冒号形也是独立标题行）——与
        # ``_para_role``/判定层词表同口径；判定层缺这一档会让产物附件区配对
        # 不到模板表单节（建设施工 1680 真值无处可对）。
        _raw_clean = (raw or "").strip()
        if len(_raw_clean) <= 10 and "\t" not in _raw_clean and _ATTACHMENT_RE.match(t):
            roles.add("attachment")
            continue
        # 目录条目以页码结尾（「1绪论4」「参考文献59」），不是标题
        if re.search(r"\d$", t):
            continue
        if _CHAPTER_RE.match(t) or _NUMBERED_RE.match(t):
            roles.add("body")
    if "toc" in roles:
        if len(texts) <= _TOC_PAGE_MAX_PARAS:
            # 目录页天然会列出「参考文献/致谢/附录」等条目，只认目录本身，避免污染
            roles = {"toc"} | ({"cover"} & roles)
        else:
            # EZI-0075：**超长节**里的一行目录标题不是这一节的身份——建设施工
            # 实测：1628 段合同正文大节含一行「目录」被独占规则吞成纯 toc 节，
            # 配对到模板唯一目录节（1680）再沿用灌给附件节。长节里目录视为
            # 透明（与 ``_product_blocks`` 的 toc 透明语义同源）。
            roles.discard("toc")
    # 第一节默认是封面；但如果它已经是摘要/正文等明确角色，就不硬加封面
    # （否则「无封面、直接从摘要开始」的文档会被误映射到模板封面节）。
    if first and (not roles or any("封面" in raw or "毕业设计" in raw for raw in texts)):
        roles.add("cover")
    return roles


def _annotate(root: etree._Element) -> list[dict]:
    """逐节读几何 + 角色；与 ``root.iter(sectPr)`` 的顺序保持一致。"""
    sects = list(root.iter(f"{W}sectPr"))
    blocks = _section_blocks(root)
    if len(blocks) != len(sects):  # 结构异常：退回纯 sectPr 顺序，不认角色
        return [{"roles": set(), "sz": _geom(sp)[0], "mar": _geom(sp)[1]} for sp in sects]
    out = []
    for i, (texts, sp) in enumerate(blocks):
        sz_d, mar_d = _geom(sp)
        out.append({"roles": _detect_roles(texts, first=(i == 0)),
                    "sz": sz_d, "mar": mar_d})
    return out


def _role_map(tpl: list[dict], prod: list[dict]) -> tuple[dict[int, int], int]:
    """按角色把产物节映射到模板节（贪心、单调、同角色按出现顺序消费）。

    返回 ``({产物节号: 模板节号}, 命中节数)``。映射只前进不后退；同一角色的
    多个模板节按顺序分配给多个产物节（正文被拆成多节时不会挤在同一节上）。
    """
    role_pos: dict[str, list[int]] = {}
    for ti, g in enumerate(tpl):
        for r in g["roles"]:
            role_pos.setdefault(r, []).append(ti)
    # EZI-0074：同角色多节混**少数派变体**（118 节合同实测：正文节里散布
    # 横向附件页 1100/1580/280、目录 1680 变体）——按序消费会把变体节抄给
    # 产物。消费前把**角色几何众数**的节排到队首（组内稳定保序）；角色
    # 均匀时排序是恒等，配对行为不变（本科 8 节两套边距零扰动）。
    for r, idxs in role_pos.items():
        if len(idxs) < 2:
            continue
        def _geom_key(ti):
            g = tpl[ti]
            return (tuple(sorted((str(k), str(v)) for k, v in (g.get("sz") or {}).items())),
                    tuple(sorted((str(k), str(v)) for k, v in (g.get("mar") or {}).items())))
        from collections import Counter as _C
        mode_key = _C(_geom_key(ti) for ti in idxs).most_common(1)[0][0]
        role_pos[r] = sorted(idxs, key=lambda ti: 0 if _geom_key(ti) == mode_key else 1)
    used = {r: 0 for r in role_pos}
    mapping: dict[int, int] = {}
    last_ti = 0
    for pi, pg in enumerate(prod):
        best: int | None = None
        for r in sorted(pg["roles"]):
            seq = role_pos.get(r) or []
            while used.get(r, 0) < len(seq) and seq[used[r]] < last_ti:
                used[r] += 1
            if used.get(r, 0) < len(seq):
                cand = seq[used[r]]
                if best is None or cand < best:
                    best = cand
        if best is None:
            # EZI-0075：单调前进耗尽但该角色模板节全在身后——合同/标书把表单区
            # 装订在正文前、产物把附件统一放尾部，两侧顺序天然不对称。回退取该
            # 角色队首（mode 重排后=众数节），不让正文几何沿用灌进附件区。
            for r in pg["roles"]:
                seq = role_pos.get(r) or []
                if seq:
                    best = seq[0]
                    break
        if best is not None:
            mapping[pi] = best
            for r in pg["roles"]:
                seq = role_pos.get(r) or []
                if used.get(r, 0) < len(seq) and seq[used[r]] == best:
                    used[r] += 1
            last_ti = best
    return mapping, len(mapping)


def _pair_sections(tpl: list[dict], prod: list[dict]) -> tuple[list[int], str]:
    """给每个产物节选出模板节号，返回 ``(模板索引序列, 映射方式)``。"""
    if len(tpl) == 1:
        return [0] * len(prod), "single-template"
    if len(prod) == len(tpl):
        return list(range(len(prod))), "index"
    mapping, matched = _role_map(tpl, prod)
    if len(set(mapping.values())) >= 2:  # 只认出封面一节时不可信，退回旧规则
        seq: list[int] = []
        cur = 0
        for pi in range(len(prod)):
            if pi in mapping:
                cur = mapping[pi]
            seq.append(cur)
        return seq, "role"
    return ([len(tpl) - 1 if i == len(prod) - 1 else 0 for i in range(len(prod))],
            "first-last")


def align_section_geometry(target_path, template_path) -> dict:
    """把模板对应节的几何写到产物各节。非阻断：失败只记录在返回值里。

    映射方式见模块 docstring：节数相等按索引，节数不等按角色（封面/摘要/目录/
    正文/参考文献/致谢…），角色认不出时退回「首节/末节」旧规则。
    """
    info: dict = {"changed": False, "sections": 0, "errors": []}
    try:
        target_path, template_path = Path(target_path), Path(template_path)
        troot = _read(template_path)
        oroot = _read(target_path)
        if troot is None or oroot is None:
            info["errors"].append("document.xml unreadable")
            return info
        tsecs = list(troot.iter(f"{W}sectPr"))
        osecs = list(oroot.iter(f"{W}sectPr"))
        if not tsecs or not osecs:
            info["errors"].append("no sectPr")
            return info
        tpl = _annotate(troot)
        prod = _annotate(oroot)
        if len(tpl) != len(tsecs) or len(prod) != len(osecs):
            info["errors"].append("sectPr/段落结构不一致")
            return info
        seq, mode = _pair_sections(tpl, prod)
        info["mapping"] = mode
        if mode == "role":
            info["product_roles"] = [sorted(g["roles"]) for g in prod]
            info["template_roles"] = [sorted(g["roles"]) for g in tpl]
            info["role_pairs"] = [
                {"product": pi, "template": ti} for pi, ti in enumerate(seq)
            ]
        touched = 0
        for sect, ti in zip(osecs, seq):
            g = tpl[ti]
            if _apply(sect, g["sz"], g["mar"]):
                touched += 1
        info["sections"] = len(osecs)
        info["template_sections"] = len(tsecs)
        if touched:
            _write(target_path, oroot)
            info["changed"] = True
            info["touched"] = touched
    except Exception as e:  # noqa: BLE001 — 非阻断
        info["errors"].append(f"{type(e).__name__}: {e}")
    return info


# ══════════════════════════════════════════════════════════════════════
# 按内容角色建节（EZI-0034）：产物节数 = 属性变化点数
# ══════════════════════════════════════════════════════════════════════
#
# 与上半部分（几何对齐）的分工：上半部分回答「已有的节该用哪套几何」，本节回答
# 「该有几节、每节从哪开始、每节哪些属性」。两者都只读模板，不猜内容语义。

#: 前置区角色（页码体系里的「罗马数字」那一段）。
_FRONT_ROLES = frozenset({"declaration", "abstract", "abstract_en", "toc", "keywords"})
#: 内容角色 → 页码体系。封面/题名无页码；前置（声明/摘要/Abstract/目录）罗马；其余阿拉伯。
_STAGE_OF_ROLE = {
    "title": "none", "cover": "none",
    "declaration": "front", "abstract": "front", "abstract_en": "front",
    "toc": "front", "keywords": "front",
    "body": "body", "references": "body", "acknowledgements": "body",
    "appendix": "body", "bio": "body", "attachment": "body",
}

#: 页码体系切换时写的 `w:pgNumType`。**不照抄模板**——本科模板连续三节都写 start=1，
#: 照抄会出现「摘要 1、Abstract 1、正文 1」三次重编号（issue 否决项①）。
_PG_NUM_OF_STAGE = {"front": ("upperRoman", 1), "body": ("decimal", 1)}

#: 段落角色词表（前缀匹配，按顺序短路）。与上半部分 `_ROLE_HEADINGS`（精确匹配）分开：
#: 上半部分做几何配对要保守（认错就套错边距），这里做「区域起点」判定，模板首段常是
#: 「摘要*中文摘要内容置于此处…」这种带注解的写法，必须前缀匹配才认得出。
_ROLE_PATTERNS: tuple[tuple[str, tuple[re.Pattern, ...]], ...] = tuple(
    (role, tuple(re.compile(p) for p in pats)) for role, pats in (
        ("declaration", (r"^(?:学术|原创性|独创性|郑重)?声明", r"^本人郑重声明",
                         r"授权使用", r"版权使用授权书", r"^学位论文使用授权")),
        ("abstract_en", (r"^abstract", r"^英文摘要")),
        ("abstract", (r"^摘\s*要", r"^中文摘要")),
        ("toc", (r"^目\s*录", r"^目\s*次", r"^contents")),
        ("references", (r"^参\s*考\s*文\s*献", r"^references?\b", r"^参考书目")),
        ("acknowledgements", (r"^致\s*谢", r"^acknowledg")),
        ("appendix", (r"^附\s*录", r"^appendix\b")),
        ("bio", (r"作者简介", r"^biograph")),
        ("cover", (r"^封面", r"^毕业设计", r"^学位论文$", r"^题目[：:]",
                   r"^密级", r"^udc", r"^论文题目")),
    )
)

#: 「第 N 章」（归一后无空格）、「1 / 1.1 / 1.4.1」式编号、「一、」式条项、罗马数字编号。
_BLOCK_CHAPTER_RE = re.compile(r"^第[0-9一二三四五六七八九十百]+章")
#: 编号标题：数字后必须跟字母/汉字（`1)`、`1）`、`1、` 这类是作者角标/条项标记，不是标题）。
_BLOCK_NUMBERED_RE = re.compile(
    r"^\d+(?:[.．]\d+)*[.．、]?\s*[A-Za-z\u4e00-\u9fff]"
)
_BLOCK_CN_ITEM_RE = re.compile(r"^[一二三四五六七八九十]+[、.．]")
_BLOCK_ROMAN_RE = re.compile(r"^[ivx]+[.．][a-z]")
#: 日期行（「2024年5月」「2024-06」）不是编号标题——否则封面日期会把封面块切成正文块。
_BLOCK_DATE_RE = re.compile(r"^\d{2,4}\s*[年\-/.．]")
#: 「附件 N」= 模板的表单区标题（常是横向页）。**必须在「以页码结尾」判定之前**认出来，
#: 否则会被当成目录条目排除；同时要求原文极短且不含制表符（目录条目「附件2⇥45」不算）。
#: 不给这一档的后果（2026-11-16 实测）：标书模板的「附件2」横向页（1420/1440）被当成
#: 正文节，整篇正文套上了横向纸张。
_ATTACHMENT_RE = re.compile(r"^附\s*件\s*[0-9]{0,2}[：:]?$")
#: 目录页节长上限（EZI-0075）：节段量超过它还含目录行 → 该节不是目录节，
#: 目录行视为透明（真目录节/目录页远短于此；建设施工被吞节实测 1628 段）。
_TOC_PAGE_MAX_PARAS = 80

#: 不限段长的角色前缀：摘要首段常整句开场（IEEE 的 `Abstract—This document …`）。
_LONG_ROLE_PATTERNS: tuple[tuple[str, tuple[re.Pattern, ...]], ...] = tuple(
    (role, tuple(re.compile(p) for p in pats)) for role, pats in (
        ("abstract_en", (r"^abstract",)),
        ("abstract", (r"^摘\s*要",)),
    )
)

#: `sectPr` 子元素顺序（CT_SectPr）。复制完属性后重排，否则 Word 认不出/报损坏。
_CT_SECTPR_ORDER: dict[str, int] = {
    name: i for i, name in enumerate((
        "headerReference", "footerReference", "footnotePr", "endnotePr", "type",
        "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols",
        "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi",
        "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
    ))
}

#: body 一级子节点里不算内容的元素（找「节末段落」时要跳过）。
_NON_PARA_TAGS = frozenset({
    f"{W}bookmarkStart", f"{W}bookmarkEnd", f"{W}commentRangeStart",
    f"{W}commentRangeEnd", f"{W}proofErr", f"{W}permStart", f"{W}permEnd",
})


def _para_role(text: str) -> str | None:
    """单段角色判定（前缀匹配，不猜语义）。

    只认短标题：目录条目以页码结尾（「参考文献……59」）一律排除；日期行不当编号标题。
    摘要按「整句开场」处理（不限段长），其余角色只认 ≤40 字的标题段。
    """
    t = _norm_heading(text)
    if not t:
        return None
    raw = (text or "").strip()
    if len(raw) <= 10 and "\t" not in raw and _ATTACHMENT_RE.match(t):
        return "attachment"          # 表单区标题（「附件2」），先于「以页码结尾」排除
    if len(t) <= 40 and re.search(r"\d$", t):  # 目录条目带页码，不是标题（与上半部分同口径）
        return None
    for role, pats in _LONG_ROLE_PATTERNS:
        for pat in pats:
            if pat.match(t):
                return role
    if len(t) > 40:
        return None
    for role, pats in _ROLE_PATTERNS:
        for pat in pats:
            if pat.match(t):
                return role
    if _BLOCK_DATE_RE.match(t):
        return None
    if (_BLOCK_CHAPTER_RE.match(t) or _BLOCK_CN_ITEM_RE.match(t)
            or _BLOCK_ROMAN_RE.match(t) or _BLOCK_NUMBERED_RE.match(t)):
        return "body"
    return None


def _region_roles(texts: list[str]) -> set[str]:
    """模板一节的角色集合：只看该节开头若干段（节的角色由它怎么开头决定）。"""
    roles: set[str] = set()
    for raw in texts[:20]:
        r = _para_role(raw)
        if r:
            roles.add(r)
    if "toc" in roles:
        if len(texts) <= _TOC_PAGE_MAX_PARAS:
            # 目录页天然列出「参考文献/致谢/附录」条目，只认目录本身（与上半部分同口径）
            roles = {"toc"} | ({"cover"} & roles)
        else:
            roles.discard("toc")   # EZI-0075：长节目录透明（同 _detect_roles）
    return roles


def _template_regions(troot: etree._Element) -> list[dict]:
    """模板逐节：角色集合 + 栏数 + 边距签名 + sectPr 元素（属性复制的来源）。"""
    sects = list(troot.iter(f"{W}sectPr"))
    blocks = _section_blocks(troot)
    if len(blocks) != len(sects):
        blocks = [([], sp) for sp in sects]
    regions: list[dict] = []
    for i, (texts, sp) in enumerate(blocks):
        _, mar = _geom(sp)
        cols = sp.find(f"{W}cols")
        typ = sp.find(f"{W}type")
        roles = _region_roles(texts)
        if not roles:
            # 认不出角色的节按正文（模板里没标「摘要/目录/参考文献」的区域就是正文）；
            # 首节例外：首节没有角色关键词 = 首页块（题名/封面），否则没有「无页码」这一档。
            roles = {"title"} if i == 0 else {"body"}
        regions.append({
            "roles": roles,
            "sect": sp,
            "cols": (cols.get(f"{W}num") or "1") if cols is not None else "1",
            "pgMar": "/".join(mar.get(k) or "?" for k in ("top", "right", "bottom", "left")),
            "type": (typ.get(f"{W}val") or "") if typ is not None else "",
            # 节首段锚（EZI-0057 模板边界镜像用）：该节第一段的归一文本；
            # 空/超长锚 = 不参与镜像（保守：宁可不切，不发明结构）。
            "anchor": _norm_heading(texts[0]) if texts else "",
        })
    return regions


def _product_blocks(body: etree._Element) -> list[dict]:
    """按文档顺序把产物正文切成有序角色块（连续同角色合并）。

    目录（toc）与关键词不独立成块：目录可能是引擎自动生成的（源稿本来没有），
    在目录处切一刀会把「目录+正文」拆成两节，违背「节数 = 属性变化点数」。
    """
    blocks: list[dict] = []
    cur: dict | None = None
    for el in body:
        if el.tag != f"{W}p":
            continue
        text = "".join(t.text or "" for t in el.iter(f"{W}t")).strip()
        role = _para_role(text) if text else None
        if role in ("toc", "keywords"):
            role = None            # 透明：并入当前块
        if role and (cur is None or role != cur["role"]):
            cur = {"role": role, "stage": _STAGE_OF_ROLE.get(role, "body"),
                   "start": el, "start_text": text, "end": el}
            blocks.append(cur)
            continue
        if cur is None:
            # 首个可识别角色之前的内容 = 题名/封面块（无页码）
            cur = {"role": "title", "stage": "none", "start": el,
                   "start_text": text, "end": el}
            blocks.append(cur)
            continue
        cur["end"] = el
    return blocks


def _plan(oroot: etree._Element, regions: list[dict]) -> dict:
    """内部计划：返回带 lxml 元素引用的块/节（对外摘要见 `plan_sections`）。"""
    plan: dict = {"skipped": None, "blocks": [], "sections": []}
    body = oroot.find(f"{W}body")
    if body is None:
        plan["skipped"] = "no-body"
        return plan
    blocks = _product_blocks(body)
    if not blocks:
        plan["skipped"] = "no-content"
        return plan

    role_index: dict[str, list[int]] = {}
    for i, g in enumerate(regions):
        for r in g["roles"]:
            role_index.setdefault(r, []).append(i)
    # EZI-0074：同角色多节混少数派变体（建设施工 118 节实测：正文角色首个
    # 区域是横向附件页 1100/1580/280、目录角色唯一节是 1680）——`_plan` 的
    # 单调前进会把变体区域抄给产物块。与 `_role_map` 同口径：候选序列按
    # **角色几何众数**排到队首（组内稳定保序）；角色均匀时排序恒等。
    def _region_geom_key(i):
        g = regions[i]
        sz = g["sect"].find(f"{W}pgSz")
        orient = (sz.get(f"{W}orient") or "portrait") if sz is not None else "portrait"
        return (g.get("pgMar") or "?", g.get("cols") or "1", orient)
    from collections import Counter as _Counter
    for r, idxs in role_index.items():
        if len(idxs) < 2:
            continue
        mode_key = _Counter(_region_geom_key(i) for i in idxs).most_common(1)[0][0]
        role_index[r] = sorted(idxs, key=lambda i: 0 if _region_geom_key(i) == mode_key else 1)

    def _candidates(role: str) -> list[int]:
        if role in ("title", "cover"):
            return role_index.get("cover") or role_index.get("title") or []
        return role_index.get(role) or []

    # 块 → 模板节：单调前进（只认比上一个块更靠后的同类节）。同角色在模板里出现多次
    # 时（正文被拆成多节），第二个块才会往前走；认不出角色就沿用上一个块的节。
    cur_tpl = 0
    for i, b in enumerate(blocks):
        seq = _candidates(b["role"])
        nxt = next((x for x in seq if x > cur_tpl), None)
        if nxt is None:
            # EZI-0075：**仅附件角色**走乱序回退——合同/标书把表单区装订在正文
            # 前、产物把附件统一放尾部，两侧顺序天然不对称；回退取队首（mode
            # 重排后=众数节），不让正文几何沿用灌进附件块。其余角色保留原沿用
            # 语义：0057 锚镜像的饿死兜底依赖 body 沿用产生的「无变化点」维持
            # 镜像节计划，实测通配回退会把 4 节镜像塌缩成 2 节（T1/T2/T4 红）。
            if b["role"] == "attachment" and seq:
                nxt = seq[0]
            else:
                nxt = seq[0] if (i == 0 and seq) else cur_tpl
        b["tpl"] = nxt
        cur_tpl = nxt

    # 页码体系只在「模板真的区分了前置区」时才作为建节理由：
    #   ① 模板 ≥3 节；② 首节是首页块（题名/封面）；③ 首节之后还有前置区节
    #      （声明/摘要/Abstract/目录）。
    # 只有①②时（合同模板：题名 + 正文，只有一套几何）切了等于把模板没表达的结构
    # 硬塞进去——参考探针在这一类上也是「0 建节」（EZI-0034 v6 审阅表）。
    first_roles = regions[0]["roles"]
    has_front_region = any(g["roles"] & _FRONT_ROLES for g in regions[1:])
    allow_stage = (len(regions) >= 3
                   and bool(first_roles & {"cover", "title"})
                   and has_front_region)
    plan["allow_page_stage"] = allow_stage

    sections: list[dict] = []
    prev: dict | None = None
    prev_stage: str | None = None
    for i, b in enumerate(blocks):
        g = regions[b["tpl"]]
        reasons: list[str] = []
        if i == 0:
            reasons = ["首节"]
        else:
            if g["cols"] != prev["cols"]:
                reasons.append(f"栏数 {prev['cols']}→{g['cols']}")
            if g["pgMar"] != prev["pgMar"]:
                reasons.append("边距变")
            if allow_stage and b["stage"] != prev_stage:
                reasons.append(f"页码体系 {prev_stage}→{b['stage']}")
        b["reasons"] = reasons
        if reasons:
            sections.append({
                "index": len(sections) + 1,
                "role": b["role"],
                "roles": [b["role"]],
                "stage": b["stage"],
                "tpl": b["tpl"],
                "cols": g["cols"],
                "pgMar": g["pgMar"],
                "type": g["type"] or "nextPage(默认)",
                "start": b["start"],
                "start_text": b["start_text"],
                "reasons": reasons,
            })
        else:
            sections[-1]["roles"].append(b["role"])
        prev = g
        prev_stage = b["stage"]

    # ── 模板边界锚镜像（EZI-0057）：角色/属性变化点计划「饿死」（<2 节）时的骨架兜底。
    # 成因：角色词表只认 封面/声明/摘要/ABSTRACT/目录/正文/参考文献/致谢/附录——
    # 篇章式标题（「第一篇 竞争性谈判邀请书」「附页：合同格式」「附件2：」）一个都认不出，
    # 产物整篇并成单个 title 块 → 0 切点 → 1 节（重庆公告 7 节模板 vs 1 节产物实测）。
    # 兜底判据：模板自切多节、节间几何确有 ≥2 组（pgMar/cols 至少两组不同——各节
    # 属性全同时维持 EZI-0034「不切」教义，见 test_no_change_point_keeps_structure），
    # 且除首节外每节的「节首段锚」都能在产物里**按序**找到同文段 → 产物内容就是
    # 模板同族结构 → 直接镜像模板节骨架：产物第 k 节 = 模板第 k region，几何逐节照抄；
    # 切点理由 = 锚本身（不是边距差——边距相同时模板也切了，节数 1:1 判定层才能按
    # 索引逐节对上）。
    # 全有或全无：任一锚空/超长(>60)/乱序/缺失 → 不镜像，维持角色计划（不发明结构）。
    # 激活面实测（源文档代理，2026-11-20）：73 对中 10 对饿死+全锚命中，含重庆 P=7；
    # 角色计划 ≥2 节的案例一律走原路径（本科论文 8 节族不受影响——它们有角色切点）。
    _geoms = {(g["pgMar"], g["cols"]) for g in regions}
    if (len(sections) < 2 and len(regions) >= 2 and len(_geoms) >= 2):
        anchored = _mirror_sections(body, regions)
        if anchored is not None:
            mirror_blocks, mirror_sections = anchored
            blocks, sections = mirror_blocks, mirror_sections
            plan["mirror"] = True

    # 页码体系：同一体系内不重编号（续前节），封面/题名不写 pgNumType。
    # 模板本身不区分前置区（allow_stage=False）时**完全不动** pgNumType——页码体系是
    # 「模板区分了前置区」才有意义的东西，否则等于给期刊论文硬塞罗马数字。
    allow_stage = bool(plan["allow_page_stage"])
    prev_stage = None
    for s in sections:
        if not allow_stage or s["stage"] == "none":
            s["pg_num"] = None
        elif s["stage"] != prev_stage:
            s["pg_num"] = _PG_NUM_OF_STAGE.get(s["stage"])
        else:
            s["pg_num"] = None
        prev_stage = s["stage"]

    plan["blocks"] = blocks
    plan["sections"] = sections
    plan["allow_page_stage"] = allow_stage
    return plan


def _mirror_sections(body: etree._Element, regions: list[dict]):
    """模板边界锚镜像（EZI-0057）：饿死计划（<2 节）时的骨架兜底，判据见 `_plan` 注释。

    返回 ``(blocks, sections)``（带元素引用，形状同角色计划）或 ``None``
    （任一节首段锚 空的/超长(>60)/在产物里按序找不到 → 不镜像，不发明结构）。
    产物第 k 节 = 模板第 k region：几何（cols/pgMar/type）逐节照抄，切点理由 =
    「模板边界锚」本身——边距相同的相邻 region 模板也切了，节数 1:1 判定层才能
    按索引逐节对上（重庆公告 7 节模板 vs 1 节产物，EZI-0057 A1 实测）。
    """
    paras = [
        (el, _norm_heading("".join(t.text or "" for t in el.iter(f"{W}t"))))
        for el in body if el.tag == f"{W}p"
    ]
    if not paras:
        return None
    hits: list[tuple[int, dict]] = []
    search = 0
    for gi in range(1, len(regions)):
        anchor = regions[gi]["anchor"]
        if not anchor or len(anchor) > 60:
            return None
        hit = next((k for k in range(search, len(paras)) if paras[k][1] == anchor), None)
        if hit is None:
            return None
        hits.append((hit, regions[gi]))
        search = hit + 1

    # 落点预检（apply 的 no-anchor 语义前置）：锚段前一个 body 级段落不存在
    # （表格/文档头紧贴锚、且不插空段承载分节符）→ 该 region **并入前一节**，
    # 不让一个落不下的锚废掉整份镜像（重庆实测：节7 锚前是申请书表格链——
    # 表格本属节6 内容，并入后末节几何用节6 的 1797≈1800，判定按索引配对不受影响）。
    blocks: list[dict] = []
    sections: list[dict] = []
    for k, g in enumerate(regions):
        if k > 0:
            pi, _g = hits[k - 1]
            if _prev_body_para(paras[pi][0]) is None:
                if sections:
                    sections[-1]["roles"].append(
                        sorted(g["roles"])[0] if g["roles"] else "body")
                continue
        if k == 0:
            start, stxt = paras[0][0], paras[0][1] or ""
        else:
            pi, _g = hits[k - 1]
            start, stxt = paras[pi][0], paras[pi][1] or ""
        role = sorted(g["roles"])[0] if g["roles"] else "body"
        stage = "none" if k == 0 else "body"
        reason = ("首节" if k == 0
                  else f"模板边界锚:{(g.get('anchor') or '')[:20]}")
        sections.append({
            "index": len(sections) + 1,
            "role": role,
            "roles": [role],
            "stage": stage,
            "tpl": k,
            "cols": g["cols"],
            "pgMar": g["pgMar"],
            "type": g["type"] or "nextPage(默认)",
            "start": start,
            "start_text": stxt,
            "reasons": [reason],
        })
        blocks.append({
            "role": role, "stage": stage, "tpl": k,
            "start": start, "start_text": stxt, "end": start,
            "reasons": [reason],
        })
    if len(sections) < 2:
        return None            # 全部锚都落不下 → 退化成 1 节，维持角色计划
    return blocks, sections


def _public_plan(regions: list[dict], plan: dict) -> dict:
    """把内部计划压成可 JSON 化的摘要（写进 artifacts / issue 证据）。"""
    return {
        "template_sections": len(regions),
        "template_roles": [sorted(r["roles"]) for r in regions],
        "allow_page_stage": bool(plan.get("allow_page_stage")),
        "skipped": plan.get("skipped"),
        "blocks": [{
            "role": b["role"], "stage": b["stage"], "tpl": b["tpl"] + 1,
            "start_text": (b["start_text"] or "")[:40],
            "reasons": list(b["reasons"]),
        } for b in plan.get("blocks") or []],
        "sections": [{
            "index": s["index"], "roles": list(s["roles"]), "stage": s["stage"],
            "tpl": s["tpl"] + 1, "cols": s["cols"], "pgMar": s["pgMar"],
            "type": s["type"],
            "pg_num": ("%s/%s" % s["pg_num"]) if s["pg_num"] else None,
            "start_text": (s["start_text"] or "")[:40],
            "reasons": list(s["reasons"]),
        } for s in plan.get("sections") or []],
        "sections_planned": len(plan.get("sections") or []),
    }


def plan_sections(doc_path, template_path) -> dict:
    """只读：产物该有几节、每节从哪个角色块开始、每节哪些属性（EZI-0034 A1）。

    判据（四选一，别再加回来）：

    1. **栏数变**（`w:cols/@w:num` 与上一块不同）；
    2. **边距变**（`w:pgMar` 与上一块不同）；
    3. **页码体系切换**（封面/题名 无页码 → 前置罗马 → 正文阿拉伯），且**只在模板
       ≥3 节且首节是首页块时生效**（模板只有 2 节说明它没区分前置区，硬切等于发明结构）；
    4. **模板边界锚镜像**（EZI-0057，饿死兜底）：前三条一条都没命中（产物 1 节）、
       且模板多节的每节首段锚都在产物里按序命中 → 1:1 镜像模板节骨架（篇章式标题
       家族——「第一篇/附页/附件2」词表认不出，角色计划必饿死）。全有或全无，
       角色计划 ≥2 节时绝不接管。

    否决项：模板 `pgNumType` 差异（本科模板连续三节都写 start=1，照抄会出现
    「摘要 1、Abstract 1、正文 1」）、页眉页脚 `reference id` 差异（按 11-10 红线不复制
    模板页眉文字，ref 不同不代表用户看得见差别，按这条切会切出 7 节）。

    护栏：单节模板一律不切；没有内容的块不建节；目录/关键词不独立成块。
    """
    info: dict = {"changed": False, "errors": [], "skipped": None,
                  "template_sections": 0, "sections_planned": 0}
    try:
        troot = _read(Path(template_path))
        oroot = _read(Path(doc_path))
        if troot is None or oroot is None:
            info["errors"].append("document.xml unreadable")
            return info
        regions = _template_regions(troot)
        info["template_sections"] = len(regions)
        if len(regions) <= 1:
            info["skipped"] = "single-template"   # 单节模板一律不切（不发明结构）
            info["template_roles"] = [sorted(r["roles"]) for r in regions]
            return info
        plan = _plan(oroot, regions)
        info.update(_public_plan(regions, plan))
        info["errors"] = list(info.get("errors") or [])
    except Exception as e:  # noqa: BLE001 — 非阻断
        info["errors"].append(f"{type(e).__name__}: {e}")
    return info


def _reorder_sectpr(sect: etree._Element) -> None:
    """按 CT_SectPr 顺序重排子元素（`w:type` 必须在 headerReference 之后、pgSz 之前）。"""
    children = list(sect)
    if not children:
        return
    ranked = sorted(
        enumerate(children),
        key=lambda pair: (_CT_SECTPR_ORDER.get(etree.QName(pair[1]).localname, 99), pair[0]),
    )
    for child in children:
        sect.remove(child)
    for _, child in ranked:
        sect.append(child)


def _el_sig(el: etree._Element) -> tuple:
    """元素签名（标签 + 属性 + 子元素递归），**忽略命名空间声明**。

    不能直接比 `etree.tostring`：模板根上声明的 `oel`/`w16du` 等命名空间会跟着子元素
    序列化出来，同一份内容在产物里永远"看起来不同"→ 幂等被打破（EZI-0034 实测：
    第二次 apply 仍报 touched=3，但文档内容一字未变）。
    """
    return (el.tag, tuple(sorted(el.attrib.items())), tuple(_el_sig(c) for c in el))


def _clean_copy(el: etree._Element) -> etree._Element:
    """复制元素但不带模板根上的命名空间声明（产物里只需 `w` 前缀）。"""
    new = etree.Element(el.tag)
    for k, v in el.attrib.items():
        new.set(k, v)
    if el.text:
        new.text = el.text
    for child in el:
        new.append(_clean_copy(child))
    return new


def _copy_from_region(sect: etree._Element, region: dict) -> bool:
    """把模板该节的 `type/pgSz/pgMar/cols/docGrid` **整元素**复制过来（缺就删）。

    整元素复制而不是逐属性赋值：`w:cols` 漏了 `w:num` 双栏不生效、`w:pgMar` 漏一条边距
    白改；`w:type` 必须随模板走（`continuous` 被写成 `nextPage`，CJC/IEEE 会凭空多一页空白）。
    """
    changed = False
    for tag in ("type", "pgSz", "pgMar", "cols", "docGrid"):
        src = region["sect"].find(f"{W}{tag}")
        cur = sect.find(f"{W}{tag}")
        if src is None:
            if cur is not None:
                sect.remove(cur)
                changed = True
            continue
        if cur is not None and _el_sig(cur) == _el_sig(src):
            continue
        new = _clean_copy(src)
        if cur is not None:
            sect.replace(cur, new)
        else:
            sect.append(new)
        changed = True
    return changed


def _set_pg_num_type(sect: etree._Element, fmt_start) -> bool:
    """`w:pgNumType`：按页码规范自定，不照抄模板。`None` = 续前节（删掉元素）。"""
    cur = sect.find(f"{W}pgNumType")
    if fmt_start is None:
        if cur is not None:
            sect.remove(cur)
            return True
        return False
    fmt, start = fmt_start
    if cur is None:
        cur = etree.Element(f"{W}pgNumType")
        sect.append(cur)
    changed = False
    if cur.get(f"{W}fmt") != fmt:
        cur.set(f"{W}fmt", fmt)
        changed = True
    if cur.get(f"{W}start") != str(start):
        cur.set(f"{W}start", str(start))
        changed = True
    return changed


def _page_footer_rids(path: Path) -> set[str]:
    """页脚部件里含 `PAGE` 域的 rel id 集合（用于「封面无页码」）。

    只删带页码域的页脚引用：页脚里若有「内部资料」这类真内容，保留（内容属「排什么」）。
    """
    rids: set[str] = set()
    rel_ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            rels_part = "word/_rels/document.xml.rels"
            if rels_part not in names:
                return rids
            rels = etree.fromstring(z.read(rels_part))
            for rel in rels.iter(f"{rel_ns}Relationship"):
                if not str(rel.get("Type") or "").endswith("/footer"):
                    continue
                target = str(rel.get("Target") or "").lstrip("/")
                part = target if target.startswith("word/") else f"word/{target}"
                if part not in names:
                    continue
                xml = z.read(part).decode("utf-8", "replace")
                if re.search(r">\s*PAGE\s*<", xml):
                    rids.add(str(rel.get("Id")))
    except Exception:  # noqa: BLE001 — 读不到就不删（保守）
        return set()
    return rids


def _prev_body_para(anchor: etree._Element) -> etree._Element | None:
    """切点承载段：新节首段**前一个** body 级段落（分节符写在它的 `w:pPr` 里）。

    前一个兄弟不是段落（表格/`w:sdt`/文档头）时返回 `None`——本模块**不**插空段来
    承载分节符（凭空多一个空段会改版面），这类切点按不可落地处理（见
    `apply_section_plan`：不可落地就整份不动）。
    """
    prev = anchor.getprevious()
    while prev is not None and prev.tag in _NON_PARA_TAGS:
        prev = prev.getprevious()
    if prev is None or prev.tag != f"{W}p":
        return None
    return prev


def _para_sectpr(p: etree._Element) -> etree._Element | None:
    """段落 `w:pPr` 里已有的 `w:sectPr`（有 = 该处已是一条分节边界）。"""
    ppr = p.find(f"{W}pPr")
    return None if ppr is None else ppr.find(f"{W}sectPr")


def _ensure_para_sectpr(prev: etree._Element, template_sect: etree._Element
                        ) -> tuple[etree._Element, bool]:
    """在 `prev`（节末段落）的 `w:pPr` 里复用已有 `sectPr`，没有才新建。

    返回 `(sectPr, created)`。`w:sectPr` 在 CT_PPr 里排最后，所以追加到 `w:pPr` 末尾。
    """
    ppr = prev.find(f"{W}pPr")
    if ppr is None:
        ppr = etree.Element(f"{W}pPr")
        prev.insert(0, ppr)
    sp = ppr.find(f"{W}sectPr")
    if sp is not None:
        return sp, False
    sp = copy.deepcopy(template_sect)
    ppr.append(sp)
    return sp, True


def _ends_like_toc_entry(text: str) -> bool:
    """目录条目形态：短行且以页码/罗马数字收尾（「参考文献……59」「1.1 研究背景 3」）。"""
    t = _norm_heading(text or "")
    return bool(t) and len(t) <= 60 and bool(re.search(r"[\dIVXivx]+$", t))


def _cut_follows_toc(prev: etree._Element | None) -> bool:
    """切点是否紧跟在「目录」块之后（= 计划把目录并进了前一节）。

    为什么需要：`_product_blocks` 把目录当**透明块**（否则「目录+正文」会被拆成两节，
    本科参考就从 3 节变 4 节）。代价是「正文节」的切点落在目录块**之后**——源稿自带
    目录时，真实分节边界在「目录前」（源稿/引擎的目录分节步骤都放那里），再按计划补
    一刀等于把目录并进封面节：判定层随即把两节都认成「目录」，页边距按角色配对翻红
    （开封标书实测 PASS→FAIL）。认得出目录块就不建这一刀，结构交给目录分节步骤。

    只认「切点往前连续是目录条目/空段，直到目录标题」这一种形态，遇到普通段落或已有
    分节边界立刻放弃——避免正文里提一句「目录」就误跳过整份。
    """
    if prev is None:
        return False
    prev_text = "".join(t.text or "" for t in prev.iter(f"{W}t")).strip()
    if prev_text and _para_role(prev_text) == "toc":
        return True
    p = prev.getprevious()
    seen = 0
    while p is not None and seen < 40:
        seen += 1
        if p.tag == f"{W}sectPr":
            return False
        if p.tag != f"{W}p":
            p = p.getprevious()
            continue
        text = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        if not text:                      # 空段：目录与正文之间的常见间隔
            p = p.getprevious()
            continue
        if _para_role(text) == "toc":
            return True
        if not _ends_like_toc_entry(text):
            return False                  # 普通段落（含声明/摘要标题）：不是目录块
        p = p.getprevious()
    return False


def apply_section_plan(doc_path, template_path, *, dry_run: bool = False) -> dict:
    """按内容角色**补建**产物分节结构并逐节复制模板属性（EZI-0034 A2/A3/A5；回工只增不删）。

    OOXML 语义（别写错位置）：

    * `w:sectPr` 写在**该节最后一段**的 `w:pPr` 里（表示「本节到此结束」）；
    * `body` 末尾那个 `w:sectPr` 属于**最后一节**；
    * 建 N 节 = 在 N-1 个「节末段落」的 `pPr` 里放 `sectPr` + 保留 body 末尾那个。

    **三种落地方式（2026-11-17 回工，按「计划节数 P vs 现有节数 E」选）**：

    * `P == E` → **整份不动**（`skipped="keep-existing"`）。现有结构是按角色建的
      （`build_section_skeleton` / `apply_section_columns` / 目录分节步骤），边界落在
      角色标题上；按计划重锚定只会把边界挪走，判定层按内容判角色再配对就全变
      （北航本科实测：5 节重锚定后 `1134` top 从附录节漂到参考文献节 → PASS→FAIL；
      开封标书 2 节边距一字未改、只是节内区间挪了，verdict 也翻）。
    * `P > E` → **只增不删**（`mode="add"`）：只补缺失的切点，该位置已有 `sectPr`
      就复用（保住页眉页脚引用与首末页标志），没有才新建；绝不删任何边界。
      切点落在「目录块之后」时整份不建（`skipped="toc-boundary"`）——目录是透明块，
      这一刀会把目录并进前一节（源稿自带目录时真实边界在「目录前」，两种放法打架）。
    * `P < E` → **收口**（`mode="rebuild"`）：产物比计划多切了刀，多出来的边界都是
      上游步骤自己插的（`apply_section_columns` 的栏变化节、目录分节步骤的空段分节符），
      留着会让属性抄到错的内容上（CJC 参考：多出的 `Background` 节让附录节停在 2 栏；
      本科参考：多出的「目录」节）。这一步才删，且只删这些计划外边界。

    前两种模式下 `removed` 恒为 0；`skipped` 的 `keep-existing / toc-boundary /
    no-anchor / duplicate-cut` 四个原因都写进 artifacts，可观测。

    计划切点全部落地后「计划节集合」与「产物节集合」一致，产物每一节都正好是计划的
    一节，逐节复制属性才成立。已有节的 pgMar/cols 只在它就是计划的变化点时被改写。

    N==1 时不动结构：没有属性变化点就没有理由拍平源稿自带的节。

    复制项：`w:type`（必复制，`continuous` 写成 `nextPage` 会多空白页）、`w:pgSz`、
    `w:pgMar`、`w:cols`、`w:docGrid` 整元素复制；`w:pgNumType` 按页码规范自定；
    封面/题名节删掉带 PAGE 域的页脚引用（封面无页码）。
    """
    info: dict = {"changed": False, "errors": [], "skipped": None,
                  "sections": 0, "created": 0, "removed": 0, "touched": 0}
    try:
        target_path = Path(doc_path)
        troot = _read(Path(template_path))
        oroot = _read(target_path)
        if troot is None or oroot is None:
            info["errors"].append("document.xml unreadable")
            return info
        regions = _template_regions(troot)
        info["template_sections"] = len(regions)
        if len(regions) <= 1:
            info["skipped"] = "single-template"
            return info
        plan = _plan(oroot, regions)
        info.update(_public_plan(regions, plan))
        sections = plan["sections"]
        info["sections"] = len(sections)
        if plan["skipped"] or len(sections) < 2:
            # 只有一节 = 没有属性变化点：结构保持原样（不做拍平）
            info["skipped"] = plan["skipped"] or "no-change-point"
            return info

        body = oroot.find(f"{W}body")
        body_sect = body.find(f"{W}sectPr") if body is not None else None
        if body_sect is None:
            info["errors"].append("no body sectPr")
            return info

        # 0) 现有分节边界（只读）：段落级 sectPr + body 末尾那个。
        existing = [sp for p in oroot.iter(f"{W}p")
                    if (sp := _para_sectpr(p)) is not None]
        existing_count = len(existing) + 1
        info["existing_sections"] = existing_count

        # 1) 节数相同 → 保持现状。现有结构是按角色建的（骨架/栏变化/目录分节），
        #    按计划重锚定只会把边界从角色标题挪到别处，判定层按内容判角色时配对随之
        #    全变（北航本科实测：5 节重锚定后 `1134` top 漂到参考文献节 → PASS→FAIL）。
        if len(sections) == existing_count:
            info["skipped"] = "keep-existing"
            return info

        # 2) 计划切点落到「节末段落」；落不下 / 切点重复 → 保持现状
        cuts: list[tuple[dict, etree._Element, etree._Element | None]] = []
        for s in sections[1:]:
            prev = _prev_body_para(s["start"])
            if prev is None:
                info["skipped"] = "no-anchor"
                return info
            cuts.append((s, prev, _para_sectpr(prev)))
        if len({id(prev) for _, prev, _ in cuts}) != len(cuts):
            info["skipped"] = "duplicate-cut"
            return info

        mode = "rebuild"
        if len(sections) > existing_count:
            mode = "add"
            # 2b) 加节模式下，切点紧跟在目录块之后 → 不建：目录是透明块，这一刀会把
            #     目录并进前一节（源稿自带目录时真实边界在「目录前」，两种放法打架；
            #     开封标书实测 PASS→FAIL）。
            if any(_cut_follows_toc(prev) for _, prev, _ in cuts):
                info["skipped"] = "toc-boundary"
                return info
            # 2c) 加节模式下还有「计划外」的已有边界 → 也不建：删不掉它，补上新切点
            #     会让节数超过计划（E + 新切点 > P），模板属性就抄到错的内容上。
            #     留着现状比多切一刀安全。
            planned = {id(sp) for _, _, sp in cuts if sp is not None}
            if any(id(sp) not in planned for sp in existing):
                info["skipped"] = "extra-breaks"
                return info
        info["mode"] = mode

        # 3) 落切点：该位置已有 sectPr 就复用，没有才新建
        breaks: list[etree._Element] = []
        created = 0
        for _s, prev, sp in cuts:
            if sp is None:
                sp, was_created = _ensure_para_sectpr(prev, body_sect)
                if was_created and not plan["allow_page_stage"]:
                    # 新节是 body 级 sectPr 的拷贝：它的 pgNumType 会把中间节重编号
                    # （末节的 start=1 被复制到每一节 → 「1…1…1」）。不归我们管页码体系时
                    # 删掉，让它续前节。
                    pn = sp.find(f"{W}pgNumType")
                    if pn is not None:
                        sp.remove(pn)
                created += 1 if was_created else 0
            breaks.append(sp)

        # 3b) 收口（只发生在「产物比计划多切了刀」时）：多出来的边界是上游步骤自己插的
        #     （`apply_section_columns` 的栏变化节 / 目录分节步骤的空段分节符），不是源稿
        #     边界——留着它们，节数就与「属性变化点」不符，属性会抄到错的内容上（CJC 实测：
        #     多出的 Background 节让附录节停在 2 栏；本科参考多出「目录」节）。加节模式
        #     绝不删任何边界。
        removed = 0
        if mode == "rebuild":
            keep = {id(sp) for sp in breaks}
            keep.add(id(body_sect))
            for p in oroot.iter(f"{W}p"):
                sp = _para_sectpr(p)
                if sp is not None and id(sp) not in keep:
                    p.find(f"{W}pPr").remove(sp)
                    removed += 1

        # 4) 逐节复制属性：第 k 节的 sectPr = 第 k+1 个切点（末节 = body 级 sectPr）
        page_footer_rids = _page_footer_rids(target_path)
        manage_pg = bool(plan["allow_page_stage"])
        sect_of = list(breaks) + [body_sect]
        touched = 0
        for s, sect in zip(sections, sect_of):
            changed = _copy_from_region(sect, regions[s["tpl"]])
            if manage_pg and _set_pg_num_type(sect, s["pg_num"]):
                changed = True
            if manage_pg and s["stage"] == "none":
                # 封面/题名无页码。只在「页码体系归我们管」时动手：模板本身不区分前置区
                # 时不动页脚引用（否则等于替用户决定首页不显示页码）。
                for ref in sect.findall(f"{W}footerReference"):
                    rid = ref.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    if rid in page_footer_rids:
                        sect.remove(ref)
                        changed = True
            _reorder_sectpr(sect)
            touched += 1 if changed else 0

        info.update({"created": created, "removed": removed, "touched": touched})
        if dry_run or not (created or removed or touched):
            info["changed"] = False
            return info
        _write(target_path, oroot)
        info["changed"] = True
    except Exception as e:  # noqa: BLE001 — 非阻断
        info["errors"].append(f"{type(e).__name__}: {e}")
    return info
