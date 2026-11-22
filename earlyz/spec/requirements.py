"""模板要求 ↔ 机器值 的对账（2026-10-17 用户拍板 P1「未解析要求清单」）。

背景：注释式模板（「（黑体小2加粗）」）与散文式模板（IEEE/自动化学报）都可能在
正文里写明要求，但只有一部分被 spec 读成机器值。用户看不到「哪些要求没落实」，
只能等成品出来才发现。

本模块做三件事：
1. extract_requirements(template)：把模板里的要求句抽出来（*…* 注释块 / 括号注释 / 散文）；
2. spec_evidence(spec, template)：每类要求是否有机器值（structure / rules / 样式）；
3. unparsed_requirements(...)：没被读成机器值的类别 + 原文例句，供编译结果告警。

口径：只报「类别级」缺口（这类要求完全没有任何机器值），不逐条比对取值——
取值比对见 scripts/audit_template_coverage.py。宁少报不误报：类别有任一机器值就算已覆盖。
"""
from __future__ import annotations

import re
import zipfile
from collections import defaultdict
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

REQ_HINT = re.compile(
    r"字体|字号|加粗|斜体|行距|段距|行间距|缩进|居中|对齐|页边距|边距|"
    r"页眉|页脚|页码|目录|摘要|关键词|参考文献|文献|题注|图题|表题|图\s*\d|表\s*\d|"
    r"表格|三线表|公式|分栏|双栏|栏宽|脚注|尾注|封面|标题|正文|单倍|磅|号字|"
    r"font|typeface|point|bold|italic|spacing|indent|margin|centered|justif|"
    r"header|footer|page number|abstract|keyword|reference|caption|table|figure|"
    r"equation|column|footnote|heading|body text",
    re.I,
)

CATEGORY_RULES: list[tuple[str, re.Pattern]] = [
    ("页边距", re.compile(r"页边距|边距|margin", re.I)),
    ("页眉页脚页码", re.compile(r"页眉|页脚|页码|header|footer|page number", re.I)),
    ("目录", re.compile(r"目录|table of contents", re.I)),
    ("摘要关键词", re.compile(r"摘要|关键词|abstract|key\s*words?", re.I)),
    ("参考文献", re.compile(r"参考文献|文献|reference|citation", re.I)),
    ("题注图表", re.compile(r"题注|图题|表题|caption|图\s*\d|表\s*\d", re.I)),
    ("表格", re.compile(r"表格|三线表|\btable\b", re.I)),
    ("公式", re.compile(r"公式|equation|math", re.I)),
    ("分栏", re.compile(r"分栏|双栏|栏宽|column", re.I)),
    ("脚注", re.compile(r"脚注|尾注|footnote|endnote", re.I)),
    ("封面", re.compile(r"封面|cover", re.I)),
    ("标题", re.compile(r"标题|heading|section title", re.I)),
    ("正文", re.compile(r"正文|body text|字体|字号|font|typeface|point|bold|italic|spacing|indent|justif|居中|对齐|行距|段距|缩进", re.I)),
]


def _ptxt(p) -> str:
    return "".join(t.text or "" for t in p.iter(f"{W}t")).strip()


def _classify(text: str) -> str:
    for name, pat in CATEGORY_RULES:
        if pat.search(text):
            return name
    return "其他"


def _paragraphs(template: Path) -> list[tuple[int, str]]:
    root = etree.fromstring(zipfile.ZipFile(template).read("word/document.xml"))
    body = root.find(f"{W}body")
    out = []
    for i, p in enumerate(body.iter(f"{W}p")):
        t = _ptxt(p)
        if t:
            out.append((i, t))
    return out


_REQ_VERB = re.compile(
    r"应|须|需|要|请|采用|使用|不得|不宜|统一|保持|设置|尺寸|字号|字体|磅|"
    r"居中|对齐|缩进|行距|边距|分栏|按|为|"
    r"\b(shall|should|must|use|used|set|format|formatted|type|typeface|font|size|"
    r"bold|italic|centered|justified|indent|spacing|margin|column|numbered)\b",
    re.I,
)
_SAMPLE_LABEL = re.compile(r"^(TABLE|FIGURE|FIG\.?|EQ\.?)\s*[IVXLC\d\.]*$", re.I)


def _is_requirement(text: str) -> bool:
    """要求句 vs 样例标签/章节名：短标签、全大写样例标号、无要求动词的一律不算。"""
    t = text.strip()
    if len(t) < 6:
        return False
    if _SAMPLE_LABEL.match(t):
        return False
    if re.fullmatch(r"[\u4e00-\u9fff]{2,4}", t):  # 「公式」「表格」这类小节名
        return False
    return bool(_REQ_VERB.search(t))


def extract_requirements(template: str | Path) -> list[dict]:
    """要求单元：*…* 注释块优先，否则整段（含关键词、长度合理、且像要求句）。"""
    template = Path(template)
    units: list[dict] = []
    for idx, text in _paragraphs(template):
        stars = re.findall(r"\*([^*]{3,200})\*", text)
        if stars:
            for s in stars:
                s = s.strip()
                if REQ_HINT.search(s) and _is_requirement(s):
                    units.append({"idx": idx, "text": s, "cat": _classify(s), "kind": "annotation"})
            continue
        if REQ_HINT.search(text) and len(text) <= 400 and _is_requirement(text):
            units.append({"idx": idx, "text": text, "cat": _classify(text), "kind": "prose"})
    return units


def _style_ids(template: Path) -> set[str]:
    try:
        root = etree.fromstring(zipfile.ZipFile(template).read("word/styles.xml"))
    except Exception:
        return set()
    return {s.get(f"{W}styleId", "") for s in root.iter(f"{W}style")}


def _has_toc_field(template: Path) -> bool:
    try:
        root = etree.fromstring(zipfile.ZipFile(template).read("word/document.xml"))
    except Exception:
        return False
    for x in root.iter(f"{W}instrText"):
        t = (x.text or "").strip().upper()
        if t.startswith("PAGEREF") or t.startswith("TOC"):
            return True
    return False


def spec_evidence(spec, template: str | Path) -> dict[str, tuple[bool, str]]:
    """每类要求是否有机器值：structure / rules / 样式 三处任一有就算覆盖。"""
    template = Path(template)
    st = getattr(spec, "structure", {}) or {}
    roles = set()
    for r in getattr(spec, "rules", []) or []:
        sel = getattr(r, "selector", None) or {}
        role = sel.get("semantic_role")
        if role:
            roles.add(str(role))
    styles = _style_ids(template)
    sections = st.get("sections") or []
    ev: dict[str, tuple[bool, str]] = {}

    def mark(cat: str, ok: bool, src: str):
        ev[cat] = (bool(ok), src if ok else "")

    mark("页边距", any((s.get("pg_mar") or {}) for s in sections),
         "structure.sections.pg_mar")
    mark("页眉页脚页码",
         bool(st.get("header_format")) or bool(st.get("header_footer_parts"))
         or any((s.get("fmt") or {}) for s in sections) or bool(st.get("front_matter_roman")),
         "structure.header_format/header_footer_parts/sections.fmt")
    mark("目录", bool(st.get("toc")) or bool(st.get("toc_title")) or _has_toc_field(template),
         "structure.toc/toc_title 或 TOC 域")
    mark("摘要关键词", any(r in roles for r in ("abstract", "abstract_en", "keywords", "keywords_en")),
         "rules.abstract/keywords")
    mark("参考文献", bool(st.get("refs")) or bool(roles & {"refs", "refs_heading", "refs_item"}),
         "structure.refs / rules.refs*")
    mark("题注图表", bool(st.get("caption")), "structure.caption")
    mark("表格", bool(st.get("table_cells")) or bool(st.get("table_width"))
         or bool(st.get("table_borders")), "structure.table_*")
    mark("公式", bool(st.get("equation")), "structure.equation")
    mark("分栏", bool(st.get("body_cols"))
         or any(int(s.get("cols_num") or 1) > 1 for s in sections), "structure.body_cols")
    mark("脚注", bool(st.get("footnote")), "structure.footnote")
    # 封面：spec 不承载封面（cover.py 走模板匹配）；有 structure.cover 才算
    mark("封面", bool(st.get("cover")), "structure.cover")
    mark("标题", any("heading" in r for r in roles)
         or any(s.lower().startswith("heading") for s in styles), "rules.heading*")
    mark("正文", "body" in roles, "rules.body")
    return ev


# ── 机制级检查（2026-10-17）──────────────────────────────────────────────
# 类别级判定会把「有字体字号」当「已覆盖」——机制型要求（每页重编号、奇偶页
# 不同…）就漏了。这里逐条要求句配对「机制是否真的落在结构里」。
_MECHANISM_RULES: list[tuple[str, re.Pattern]] = [
    ("脚注每页重编号",
     re.compile(r"(每\s*页[^。；]{0,16}脚注)|(脚注[^。；]{0,16}每\s*页)")),
    ("奇偶页页眉不同",
     re.compile(r"奇[数偶]页|even and odd|odd and even", re.I)),
    ("首页页眉不同",
     re.compile(r"首页不同|首页页眉|title page|first page (?:header|different)", re.I)),
    ("页码接续/分节重编",
     re.compile(r"页码(?:接续|连续)|接上页|continued page numbering", re.I)),
    ("双语题注",
     re.compile(r"双语题注|中英文题注|bilingual caption", re.I)),
]


def _mech_ok(name: str, spec, template: Path) -> tuple[bool, str]:
    st = getattr(spec, "structure", {}) or {}
    slots = [tuple(x) for x in (st.get("hf_bound_slots") or [])]
    if name == "脚注每页重编号":
        return (st.get("footnote_restart") == "eachPage", "structure.footnote_restart")
    if name == "奇偶页页眉不同":
        return (any(s[:2] == ("header", "even") for s in slots), "hf_bound_slots")
    if name == "首页页眉不同":
        return (any(s[:2] == ("header", "first") for s in slots), "hf_bound_slots")
    if name == "页码接续/分节重编":
        sects = st.get("sections") or []
        return (any((s.get("fmt") or {}) for s in sects) or any(s.get("pg_num") for s in sects),
                "structure.sections[].fmt")
    if name == "双语题注":
        return (bool((st.get("caption") or {}).get("bilingual")), "structure.caption.bilingual")
    return (True, "")


def mechanism_gaps(template: str | Path, spec) -> list[dict]:
    """模板写了机制、结构里没有 → 逐条返回（类别名「机制：xxx」）。"""
    template = Path(template)
    try:
        units = extract_requirements(template)
    except Exception:
        return []
    out: list[dict] = []
    seen: set[str] = set()
    for name, pat in _MECHANISM_RULES:
        if name in seen:
            continue
        hits = [u["text"] for u in units if pat.search(u["text"])]
        if not hits:
            continue
        ok, _src = _mech_ok(name, spec, template)
        if ok:
            continue
        seen.add(name)
        out.append({"category": f"机制：{name}", "count": len(hits),
                    "samples": [h[:160] for h in hits[:2]]})
    return out


def _template_facts(template: str | Path) -> dict:
    """模板里实际有哪些对象（表/图/公式/参考文献章节）。"""
    import zipfile

    from lxml import etree

    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    facts = {"tables": 0, "figures": 0, "equations": 0, "refs": False}
    try:
        with zipfile.ZipFile(str(template)) as z:
            root = etree.fromstring(z.read("word/document.xml"))
    except Exception:  # noqa: BLE001
        return facts
    facts["tables"] = len(list(root.iter(f"{W}tbl")))
    facts["figures"] = len(list(root.iter(f"{W}drawing"))) + len(list(root.iter(f"{W}pict")))
    facts["equations"] = len(list(root.iter(f"{W}object")))
    for para in root.iter(f"{W}p"):
        text = "".join(t.text or "" for t in para.iter(f"{W}t")).strip()
        if text in ("参考文献", "References", "REFERENCES", "参 考 文 献"):
            facts["refs"] = True
            break
    return facts


def structure_gaps(template: str | Path, spec) -> list[dict]:
    """模板里有对象、结构里却没机器值 → 报缺口（比「散文没抽到」更硬的证据）。"""
    st = getattr(spec, "structure", {}) or {}
    facts = _template_facts(template)
    cap = st.get("caption") or {}
    out: list[dict] = []

    def add(cat: str, why: str, advice: str = ""):
        out.append({"category": cat, "count": 1, "samples": [],
                    "why": why, "advice": advice, "no_value": True})

    if facts["tables"]:
        if not st.get("table_borders"):
            add("表格", f"模板里有 {facts['tables']} 个表格，但表线样式没给具体值",
                "表线用哪种、线宽多少")
        if not cap.get("table_sz"):
            add("题注图表", "模板里有表格，但表题字号没给具体值",
                "表题几号字、是否加粗")
    if facts["figures"]:
        if not (cap.get("figure_sz") or st.get("figure_width")):
            add("题注图表", f"模板里有 {facts['figures']} 个图，但图题字号/图宽没给具体值",
                "图题几号字、图占栏宽多少")
    if facts["refs"] and not st.get("refs"):
        add("参考文献", "模板有参考文献章节，但字号/行距/编号方式没给具体值",
            "几号字、什么行距、是否要 [1] 编号")
    if facts["equations"] and not st.get("equation"):
        add("公式", f"模板里有 {facts['equations']} 个公式对象，但字号/编号方式没给具体值",
            "公式几号字、编号靠右还是居中")
    return out


def unparsed_requirements(
    template: str | Path,
    spec,
    *,
    max_items: int = 6,
) -> list[dict]:
    """没被读成机器值的类别（含原文例句）。已覆盖的类别不返回。"""
    template = Path(template)
    try:
        units = extract_requirements(template)
    except Exception:
        units = []
    # 注意：没有散文要求不代表没有缺口——模板里可能有表格/图/参考文献，
    # 但结构里没机器值（structure_gaps 就是为这种情况加的）→ 不能在这里 return。
    ev = spec_evidence(spec, template)
    by_cat: dict[str, list[str]] = defaultdict(list)
    for u in units:
        cat = u["cat"]
        if cat == "其他":
            continue
        if not ev.get(cat, (False, ""))[0]:
            by_cat[cat].append(u["text"][:160])
    out = []
    for cat, samples in by_cat.items():
        out.append({
            "category": cat,
            "count": len(samples),
            "samples": samples[:2],
            "why": _why_unresolved(samples[0] if samples else ""),
        })
    # 模板自己声明了、但引擎做不到的约定（如表格罗马数字编号）
    for item in (getattr(spec, "structure", {}) or {}).get("unsupported_requirements") or []:
        out.append({
            "category": str(item.get("category") or "其他"),
            "count": 1,
            "samples": [str(item.get("text") or "")[:160]],
            "reason": item.get("reason"),
        })

    # 机制级：类别级看起来"有值"，但机制没落地（CJC 脚注每页重编号就是这么漏的）
    for mg in mechanism_gaps(template, spec):
        out.append(mg)

    # 结构级（2026-10-17）：模板里有表格/图/公式/参考文献，结构里却没机器值
    seen = {o["category"] for o in out}
    for gap in structure_gaps(template, spec):
        if gap["category"] not in seen:
            out.append(gap)
            seen.add(gap["category"])
    out.sort(key=lambda x: -x["count"])
    return out[:max_items]


_WRITING_HINTS = (
    "定义", "缩写", "引用", "符号", "单位", "作者", "写作", "表述", "拼写", "不要",
    "应当", "请", "avoid", "use the", "abbreviat", "italicize", "define",
    "should", "spell", "write", "refer to",
)


def _why_unresolved(sample: str) -> str:
    """区分「模板没给样例（版式类）」和「本来就是写作规范（引擎不代写）」。"""
    t = (sample or "").casefold()
    if any(h.casefold() in t for h in _WRITING_HINTS):
        return "模板只写了写作要求，没有版式数值；内容要你自己写"
    return "模板没给可解析的样例/数值"


def warning_text(unparsed: list[dict]) -> str | None:
    if not unparsed:
        return None
    lines = []
    for u in unparsed:
        sample = (u.get("samples") or [""])[0]
        why = u.get("why") or u.get("reason") or ""
        if sample:
            seg = f"{u['category']}（{u['count']} 条）——「{sample[:70]}」"
            if why:
                seg += f"（{why}）"
        else:
            # 结构级缺口没有原话，只有原因
            seg = f"{u['category']}——{why}" if why else f"{u['category']}（{u['count']} 条）"
        if u.get("advice"):
            seg += f"（告诉我{u['advice']}）"
        lines.append(seg)
    tail = "。成品可能和模板不完全一致，建议人工核对一下。"
    if any(u.get("no_value") for u in unparsed):
        tail = "。这些地方模板没给数值，我没有自行编造。"
    return (f"模板里有 {len(unparsed)} 类要求还没有自动落实：" + "；".join(lines) + tail)
