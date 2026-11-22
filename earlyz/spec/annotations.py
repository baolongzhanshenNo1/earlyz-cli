"""模板「注解层」解析：把模板作者写在文本框/批注/正文里的版式说明变成规则。

问题（2026-10-27 交大本科案例）：校内模板把版式要求写在文本框里——
「一级标题三号黑体居中加粗，上下各空一行」「正文: 中文为五号宋体…首行缩进」。
引擎原先用 body.iter(w:p) 统一取样，这些**说明句被当成正文/标题样例**统计，
导致 body 规则丢缩进、heading1 被认成五号宋体，产物版式自然全错。

本模块只做两件事：
1. 把注解文本解析成结构化规则（角色 + 字号/字体/加粗/对齐/缩进/行距）；
2. 按「注解优先于样例」把这些值合并进 FormatSpec（只在注解明确给出时覆盖）。

不做：不猜注解里没写的属性；角色映射不了就原样留在清单里，交给上层确认。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

SIZE_MAP = {
    "初号": 42.0, "小初": 36.0, "一号": 26.0, "小一": 24.0,
    "二号": 22.0, "小二": 18.0, "三号": 16.0, "小三": 15.0,
    "四号": 14.0, "小四": 12.0, "五号": 10.5, "小五": 9.0,
    "六号": 7.5, "小六": 6.5,
}

# 说明句的特征（也是「这不是正文」的判据）。
# 关键：(1) 真标题「1 引言（黑体4号）」不能算说明；(2) 纯格式标签「四号黑体居中」要算说明。
INSTRUCTION_RE = re.compile(
    r"(号(宋体|黑体|仿宋|楷体|隶书|Times)|首行缩进|空一行|单倍行距|固定值|加粗|加黑"
    r"|居中，|页边距|字号|行距|缩进(二|两|2)"
    # 表格类模板说明（「表题允许下页接写…表头应重复书写」）
    r"|下页接写|表头应重复|续表xx|续表 XX)"
)
_HEADING_PREFIX = re.compile(
    r"^\s*[（(]?[①-⑳]?[）)]?\s*(\d+(?:\.\d+)*[\s.．、]|第[一二三四五六七八九十百]+[章节]"
    r"|[一二三四五六七八九十]+、|Chapter\s+\d)"
)
_ROLE_START = re.compile(
    r"^\s*[（(]?[①-⑳\d]*[）)]?\s*(正文|摘要|题目|中文题目|英文题目|一级标题|二级标题|三级标题"
    r"|四级标题|第[一二三四]级标题|关键词|关键字|参考文献|图题|表题|图表题注|题注|页眉|页脚|页码"
    r"|注[:：]|题目|标题)"
)
_DESCRIBE = re.compile(r"(空一行|首行缩进|缩进\s*(二|两|2)\s*个?\s*字?符?|单倍行距|固定值|行距|页边距|段前|段后|字体为|加黑|一律用大写|大写字母)")
_SIZE_OR_FONT = re.compile(r"(号(宋体|黑体|仿宋|楷体|隶书)|宋体|黑体|仿宋|楷体|Times New Roman|\d+\s*(磅|pt))")


FONT_EA = ("黑体", "宋体", "仿宋", "楷体", "隶书")
FONT_LATIN = ("Times New Roman", "Arial", "Calibri")

# 角色识别：具体在前，泛化在后
_ROLE_PATTERNS: list[tuple[re.Pattern, str]] = [
    (re.compile(r"(英文题目|英文标题|英文论文题目)"), "title_en"),
    (re.compile(r"(中文题目|论文题目|毕业设计（论文）题目|题目)"), "title"),
    (re.compile(r"(一级标题|章标题|章名|第[一二三四五六七八九十]+章)"), "heading1"),
    (re.compile(r"(二级标题)"), "heading2"),
    (re.compile(r"(三级标题|四级标题|第三级|第四级|三级和四级)"), "heading3"),
    (re.compile(r"(关键词|关键字|keywords)", re.I), "keywords"),
    # EZI-0065 ②：期刊前区角色（CJC「作者名1)…字体为3号仿宋_GB2312」全被丢弃，
    # 因为词表里根本没有 作者/单位）
    (re.compile(r"(英文作者|英文姓名)"), "authors_en"),
    (re.compile(r"(作者|姓名)"), "authors"),
    (re.compile(r"(单位|学院|学校|机构)"), "affiliation"),
    # EZI-0065 E3：「参考文献（四号黑体居中）」说的是**标题行**不是条目
    # （中南民族实测：refs_item 被改成 黑体14居中，条目本应宋体顶格）。
    # 居中/短标签形态 → refs_heading；带 编号/次序/条目 词的才是条目格式。
    (re.compile(r"(参考文献|文献著录)[^,。；]{0,12}(居中|标题)"), "refs_heading"),
    (re.compile(r"(参考文献|文献著录)"), "references"),
    (re.compile(r"(表题|图题|图表题注|题注|表名|图名|表标题|图标题)"), "caption"),
    (re.compile(r"(页眉|页脚|页码)"), "header_footer"),
    # 「摘要正文」是摘要专用的版式，不等于全文正文：单列角色，避免
    # 「摘要单倍行距」被当成正文行距（交大实测：正文样例是 1.5 倍）
    (re.compile(r"(摘要正文|abstract\s*body)", re.I), "abstract_body"),
    (re.compile(r"(摘要|abstract)", re.I), "abstract_body"),
    (re.compile(r"(正文|论文正文|内容)"), "body"),
]

ROLE_TO_RULE_SUFFIX = {
    # EZI-0065 ③：后缀表对齐 decompiler 实际产出的规则名——期刊域是
    # tpl.article_title/tpl.authors/tpl.affiliation/tpl.refs_item，旧表只有
    # 「title/cover_title」，role 命中了也找不到规则位（CJC/矿大实测全弃）。
    "title": ("title", "cover_title", "article_title"),
    "title_en": ("title_en", "title", "article_title_en"),
    "authors": ("authors",),
    "authors_en": ("authors_en",),
    "affiliation": ("affiliation",),
    "heading1": ("heading1",),
    "heading2": ("heading2",),
    "heading3": ("heading3", "heading4"),
    "body": ("body",),
    "keywords": ("keywords", "keywords_en"),
    "caption": ("caption_table", "caption_figure", "caption"),
    "references": ("references", "refs_item"),
    "refs_heading": ("refs_heading",),
    "header_footer": ("header", "footer", "footer_format", "header_format"),
    # 摘要正文暂无独立规则位（body 是全文正文）——只记录不合并，
    # 防止「摘要单倍行距」污染正文行距。
    "abstract_body": (),
}


@dataclass
class Annotation:
    text: str
    role: str | None = None
    size_pt: float | None = None
    font_ea: str | None = None
    font_latin: str | None = None
    bold: bool | None = None
    jc: str | None = None
    indent2: bool = False
    line: int | None = None
    line_rule: str | None = None
    where: str = "textbox"          # textbox / body / comment / table
    role_pos: int = -1              # 角色词在句中的位置（越靠前越可信）
    label_only: bool = False        # EZI-0065 E1：只描述标签格式的注解（不覆盖内容）
    anchor_idx: int = -1            # EZI-0065 E9：批注锚定的段落序号（-1=非批注/无锚）
    role_hits: int = 0              # 命中的角色关键词个数（>1 说明句子在讲多个角色）
    extra: dict = field(default_factory=dict)

    def as_dict(self) -> dict:
        return {
            "text": self.text, "role": self.role, "size_pt": self.size_pt,
            "font_ea": self.font_ea, "font_latin": self.font_latin,
            "bold": self.bold, "jc": self.jc, "indent2": self.indent2,
            "line": self.line, "line_rule": self.line_rule, "where": self.where,
        }


def is_instruction(text: str) -> bool:
    """这段文字像不像「版式说明」而不是正文/标题。

    判据（按优先级）：
      1. 带标题编号前缀（1 / 1.1 / 第一章 / 一、）→ 不是说明，是真标题
      2. 句首是角色词（正文/摘要/一级标题/题注/页眉…）→ 是说明
      3. 含描述性短语（空一行/首行缩进/行距/页边距）→ 是说明
      4. 短纯标签（四号黑体居中）→ 是说明
    """
    t = (text or "").strip()
    if not t:
        return False
    if _HEADING_PREFIX.match(t) and not _DESCRIBE.search(t):
        return False
    if _ROLE_START.match(t) and (_SIZE_OR_FONT.search(t) or _DESCRIBE.search(t)):
        return True
    # 表格类模板说明（「表题允许下页接写…表头应重复书写」）——没有字号，
    # 但整句就是在讲模板怎么排，属说明句
    if re.search(r"(下页接写|表头应重复|续表\s*[xX]{1,2})", t):
        return True
    # 纯空行说明（「空一行空一行」）——没有字号字体，但同样是模板说明
    if re.fullmatch(r"\s*(空[一两二]?行[，,。；;、]?\s*)+", t):
        return True
    if _DESCRIBE.search(t) and _SIZE_OR_FONT.search(t):
        return True
    if (len(t) <= 20 and _SIZE_OR_FONT.search(t)
            and re.fullmatch(r"[一-鿿A-Za-z\s，,、。；;（）()]{4,20}", t)
            and not re.search(r"\d{2,}", t)):
        return True
    return False


# 数字号 → 汉字号归一（EZI-0065 ①）：期刊模板大量用「2号/4号/小5号」数字形态
# （CJC-Templet_Word2003 实测：15 条注解字号全 None，只因词表只有「二号/四号」）。
# 「小N号」先归一（长词优先，防止「小5号」在位置 1 被「5号」抢走）。
_DIGIT_CN = {"1": "一", "2": "二", "3": "三", "4": "四", "5": "五", "6": "六", "7": "七"}


def _normalize_digit_hao(text: str) -> str:
    t = re.sub(r"小\s*([1-6])\s*号",
               lambda m: "小" + _DIGIT_CN[m.group(1)], text)
    t = re.sub(r"(?<![0-9小])([1-6])\s*号",
               lambda m: _DIGIT_CN[m.group(1)] + "号", t)
    return t


def _pick_size(text: str) -> float | None:
    # 必须长词优先：「小四号」里含「四号」，按 dict 顺序会先命中「四号」=14pt
    # （2026-10-28 大连民族实测：注解写「小四号」，解析成 14pt，正文规则全错）
    # 用正则按「文本位置」匹配：「小四号」在位置 0 命中「小四」，
    # 不会像子串查找那样在位置 1 先命中「四号」
    text = _normalize_digit_hao(text)
    m = re.search("|".join(sorted(SIZE_MAP, key=len, reverse=True)), text)
    if m:
        return SIZE_MAP[m.group(0)]
    m = re.search(r"(\d+(?:\.\d+)?)\s*(?:磅|pt)", text, re.I)
    return float(m.group(1)) if m else None


def _pick_fonts(text: str) -> tuple[str | None, str | None]:
    ea = next((f for f in FONT_EA if f in text), None)
    latin = next((f for f in FONT_LATIN if f.lower() in text.lower()), None)
    if ea is None and latin is None:
        # 「黑体四号」这种把字体写在前面也算
        m = re.search(r"(黑体|宋体|仿宋|楷体|隶书)", text)
        ea = m.group(1) if m else None
    return ea, latin


def _pick_line(text: str) -> tuple[int | None, str | None]:
    if "单倍" in text:
        return 240, "auto"
    m = re.search(r"(\d(?:\.\d+)?)\s*倍行距", text)
    if m:
        try:
            return int(round(float(m.group(1)) * 240)), "auto"
        except ValueError:
            pass
    m = re.search(r"固定值\s*(\d+(?:\.\d+)?)\s*(?:磅|pt)?", text)
    if m:
        try:
            return int(round(float(m.group(1)) * 20)), "exact"
        except ValueError:
            pass
    return None, None


# EZI-0065 E1：带引号的标签形态（"关键词"：黑体，小四…）只描述**标签**，
# 不覆盖整段内容格式（3-2-1 实测：内容本应宋体，被改成黑体整段）。
_LABEL_QUOTED = re.compile(r"[“「『\"]?(关键词|关键字|Key\s*words?)[”」』\"]?[:：]")


def parse_annotation(text: str, where: str = "textbox") -> Annotation | None:
    t = (text or "").strip()
    if not t or not is_instruction(t):
        return None
    ann = Annotation(text=t[:160], where=where)
    if _LABEL_QUOTED.search(t) and "内容" not in t:
        ann.label_only = True
    ann.size_pt = _pick_size(t)
    ann.font_ea, ann.font_latin = _pick_fonts(t)
    if "加粗" in t or "加黑" in t:
        ann.bold = True
    if "居中" in t:
        ann.jc = "center"
    elif "右对齐" in t or "靠右" in t:
        ann.jc = "right"
    elif "两端" in t:
        ann.jc = "both"
    if re.search(r"缩进\s*(二|两|2)\s*个?\s*字?符?", t):
        ann.indent2 = True
    ann.line, ann.line_rule = _pick_line(t)
    # ① 结构前缀优先（「第4章…」「4.1 …」「4.1.1 …」）——这类句子在讲某级标题，
    #    而不是文档题目（中国石油大学模板实测：「4.1 第4章第1节题目」被错认成 title）
    m_pref = re.match(r"^\s*第\s*[一二三四五六七八九十\d]+\s*章", t)
    if m_pref:
        ann.role, ann.role_pos, ann.role_hits = "heading1", 0, 1
    else:
        m_num = re.match(r"^\s*(\d+(?:\.\d+)+)", t)
        if m_num:
            depth = len(m_num.group(1).split("."))
            ann.role = "heading3" if depth >= 3 else "heading2"
            ann.role_pos, ann.role_hits = 0, 1
        elif re.match(r"^\s*\d+\s+\S", t):
            ann.role, ann.role_pos, ann.role_hits = "heading1", 0, 1

    # ② 「标题“ABSTRACT”…标题1」：引用的是某一级标题，不是文档题目
    if ann.role is None and re.match(r"^\s*标题\s*[“\"']", t):
        m_lv = re.search(r"标题\s*([1-4])", t)
        ann.role = f"heading{m_lv.group(1)}" if m_lv else "heading1"
        ann.role_pos, ann.role_hits = 0, 1

    # ③ 「英文/中文…题目」这类主语是题目，别被句中的「摘要」抢走
    head = t[:20]
    if ann.role is None and re.search(r"(?<!一级)(?<!二级)(?<!三级)(?<!四级)(?<!副)(题目|标题)", head)             and not re.match(r"^(正文|摘要正文)", t):
        # EZI-0065 ⑦：英文信号必须在**句首**——CJC「题目*（中英文题目一致）
        # 字体为2号黑体」的括号注里有「英文」二字，search 全句会把中文题目行
        # 误路由成 title_en，22pt 题目被 14pt Title 批注改掉（实测回退）。
        ann.role = ("title_en" if re.match(r"\s*(英文|English|Title)", head, re.I)
                    else "title")
        ann.role_pos, ann.role_hits = 0, 1

    if ann.role is None:
        best: tuple[int, str] | None = None
        roles_hit: set[str] = set()
        for pat, role in _ROLE_PATTERNS:
            m = pat.search(t)
            if not m:
                continue
            roles_hit.add(role)   # 同角色多关键词（「摘要正文」）不算歧义
            if best is None or m.start() < best[0]:
                best = (m.start(), role)
        if best is not None:
            ann.role, ann.role_pos, ann.role_hits = best[1], best[0], len(roles_hit)
    if ann.role is None and not any((ann.size_pt, ann.font_ea, ann.jc, ann.indent2, ann.line)):
        return None
    return ann


def _iter_parts(z) -> list[tuple[str, etree._Element]]:
    out = []
    for name in z.namelist():
        if not name.endswith(".xml"):
            continue
        if not (name == "word/document.xml" or name.startswith("word/header")
                or name.startswith("word/footer") or name == "word/comments.xml"):
            continue
        try:
            out.append((name, etree.fromstring(z.read(name))))
        except etree.XMLSyntaxError:
            continue
    return out


def extract_annotations(path: str | Path) -> list[Annotation]:
    """抽取模板里的版式说明：文本框优先，其次是正文/表格/批注/页眉页脚。"""
    import zipfile

    anns: list[Annotation] = []
    seen: set[str] = set()
    try:
        z = zipfile.ZipFile(path)
    except Exception:  # noqa: BLE001
        return anns

    # EZI-0065 E9：批注锚点表——commentRangeStart/commentReference 所在段的序号
    anchor_map: dict[str, int] = {}
    try:
        droot = etree.fromstring(z.read("word/document.xml"))
        for i, dp in enumerate(droot.iter(f"{W}p")):
            for el in dp.iter(f"{W}commentRangeStart"):
                anchor_map.setdefault(el.get(f"{W}id"), i)
            for el in dp.iter(f"{W}commentReference"):
                anchor_map.setdefault(el.get(f"{W}id"), i)
    except Exception:  # noqa: BLE001 — 无 document.xml/解析失败 → 锚点缺失，不阻断
        anchor_map = {}

    for part, root in _iter_parts(z):
        if part == "word/comments.xml":
            # EZI-0065 ⑥：批注按**整条**解析，不按段——矿大大摘要实测每条批注
            # 首段是角色标签（「论文题目：」），字体字号写在后续段里；按段拆开
            # 后标签段不含格式词被 is_instruction 丢、值段不含角色词 role=None，
            # 18 条批注 0 应用。整条拼接后「论文题目：（1）字体：黑体…字号：
            # 小二…」角色+值俱全。
            _ITEM = re.compile(r"[（(]\s*[0-9①-⑳一二三四五六七八九十]+\s*[)）]")
            for c in root.iter(f"{W}comment"):
                t = " ".join(
                    "".join(x.text or "" for x in p.iter(f"{W}t")).strip()
                    for p in c.iter(f"{W}p")
                ).strip()
                cid_anchor = anchor_map.get(c.get(f"{W}id"), -1)
                # EZI-0065 ⑩：一条批注里塞多条要求时的两种形态——
                #   矿大「论文题目：（1）字体：黑体…（2）…」：编号**前**有角色
                #   标签，整条一个元素 → 不拆，整条解析；
                #   3-3-1「（1）摘要…（2）关键词…（3）一级标题和页码：…」：
                #   编号前没有角色标签，每个编号一条元素 → 按段拆开分别解析
                #   （不拆会丢掉后面元素的角色，heading1 加粗补不上）。
                m0 = _ITEM.search(t)
                head = t[:m0.start()] if m0 else t
                segs = [t]
                # 包含式角色词判定（锚定式 _ROLE_START 要求角色词在句首 0 位，
                # 「论文题目：」的「题目」在第 2 位会漏匹配 → 矿大被误拆）
                _head_role = re.search(
                    r"(题目|标题|摘要|关键词|关键字|作者|姓名|单位|学院|正文"
                    r"|页眉|页脚|页码|参考文献|题注|图题|表题|题名)", head)
                if m0 and not _head_role:
                    parts = _ITEM.split(t)
                    marks = _ITEM.findall(t)
                    segs = [f"{m}{s}".strip() for m, s in zip(marks, parts[1:]) if s.strip()]
                for s in segs:
                    a = parse_annotation(s, "comment")
                    if a and s not in seen:
                        seen.add(s)
                        a.anchor_idx = cid_anchor
                        anns.append(a)
            continue

        # 文本框
        for box in root.iter(f"{W}txbxContent"):
            for p in box.iter(f"{W}p"):
                t = "".join(x.text or "" for x in p.iter(f"{W}t")).strip()
                a = parse_annotation(t, "textbox")
                if a and t not in seen:
                    seen.add(t)
                    anns.append(a)

        # 正文/表格/页眉页脚里直接写的说明
        for p in root.iter(f"{W}p"):
            # 不能用 id(p) 记账：lxml 的段代理是临时的，循环结束后被回收，
            # id 会被后续段复用 → 正文段随机被当成文本框段跳过（实测同一模板
            # 三次调用抽出 10/5/6 条注解）。按祖先判定文本框，结果稳定。
            if any(a.tag == f"{W}txbxContent" for a in p.iterancestors()):
                continue
            # EZI-0065 ⑤：mc:Fallback 下的段是 AlternateContent 的兼容拷贝，
            # 与 Choice 内容完全重复（上交本科实测「四号黑体居中」出现两次粘
            # 成一段）——Fallback 段整段跳过，只认 Choice。
            if any(a.tag.endswith("}Fallback") for a in p.iterancestors()):
                continue
            t = "".join(x.text or "" for x in p.iter(f"{W}t")).strip()
            where = "body"
            if part.startswith("word/header") or part.startswith("word/footer"):
                where = "header_footer"
            elif any(anc.tag == f"{W}tbl" for anc in p.iterancestors()):
                where = "table"
            a = parse_annotation(t, where)
            if a and t not in seen:
                seen.add(t)
                anns.append(a)
    return anns


def _rule_for_role(spec, role: str):
    """按角色找目标规则（优先精确后缀，其次含该角色名的规则）。"""
    suffixes = ROLE_TO_RULE_SUFFIX.get(role, ())
    rules = list(getattr(spec, "rules", []) or [])
    for suf in suffixes:
        for r in rules:
            if r.rule_id.endswith("." + suf) or r.rule_id == f"rule:{suf}":
                return r
    return None


def _mode(values):
    vals = [v for v in values if v is not None]
    if not vals:
        return None
    counts = {}
    for v in vals:
        counts[v] = counts.get(v, 0) + 1
    return max(counts.items(), key=lambda kv: (kv[1], -vals.index(kv[0])))[0]


def merge_annotations(spec, annotations: list[Annotation]) -> list[dict]:
    """把注解值合并进 spec（注解优先，只覆盖它明确给出的属性）。

    返回「实际应用」的清单，供日志/诊断/测试断言。
    """
    applied: list[dict] = []
    conflicts: list[dict] = []
    by_role: dict[str, list[Annotation]] = {}
    for ann in annotations:
        if ann.role:
            by_role.setdefault(ann.role, []).append(ann)

    for role, group in by_role.items():
        rule = _rule_for_role(spec, role)
        if rule is None:
            continue
        # 置信度：角色词必须在句首附近，且整句只讲一个角色。
        # 低置信度的句子往往是「（3）每节中的一级标题…」（实指三级）这类，
        # 直接覆盖会把模板改坏 —— 只记冲突，不动取值。
        _qualified = sum(1 for a in group if a.role_pos <= 3 and a.role_hits <= 1)
        # 半数即可（原为「严格过半」，4 条里 2 条合格时会被判低置信，正文缩进就补不上）
        high = _qualified > 0 and _qualified >= len(group) / 2
        # EZI-0065 E1：纯标签注解组（全部 label_only）不覆盖内容格式，只记冲突
        if all(a.label_only for a in group):
            conflicts.append({
                "role": role, "rule_id": rule.rule_id, "text": group[0].text,
                "reason": "标签格式注解（如 “关键词”：黑体）只约束标签 run，"
                          "整段内容格式未动，待标签级规则位支持",
            })
            continue
        before = dict(rule.expected or {})
        exp = dict(before)

        # EZI-0065 E2/E9：同名角色多条注解值有分歧 = 在描述不同元素（封面题目
        # vs 大摘要页题目；3-3-1 实测小二18 与 三号16 混组 mash）。消歧两档：
        #   ① 有批注锚点 → **早锚优先**（前区主元素：文档第一个该角色元素，
        #      矿大实测 锚段7 的封面题目 18 胜 锚段30 的大摘要页题目 14）；
        #   ② 无锚点 → 一致才应用，分歧记冲突待人工裁定。
        def _unanimous(vals):
            vs = [v for v in vals if v is not None]
            if not vs:
                return None
            return vs[0] if len(set(vs)) == 1 else "CONFLICT"

        _anchored = [a for a in group if a.anchor_idx >= 0]

        def _get_prop(a, key):
            return {"字号": a.size_pt, "中文字体": a.font_ea, "西文字体": a.font_latin,
                    "加粗": a.bold, "对齐": a.jc, "行距": a.line}.get(key)

        def _mode_or_conflict(vals, key):
            # 三档：全一致 → 用；严格多数（> 半）→ 用多数并记少数派；
            # 无多数（平票/碎片）→ 描述不同元素（封面 vs 大摘要页题目），
            # 有批注锚点取早锚（文档首个该角色元素），否则挂起待人工。
            vs = [v for v in vals if v is not None]
            if not vs:
                return None
            if len(set(vs)) == 1:
                return vs[0]
            counts: dict = {}
            for v in vs:
                counts[v] = counts.get(v, 0) + 1
            top_v, top_n = max(counts.items(), key=lambda kv: kv[1])
            if top_n > len(vs) / 2:
                conflicts.append({
                    "role": role, "rule_id": rule.rule_id, "text": group[0].text,
                    "reason": f"注解组内 {key} 多数 {top_v}（{top_n}/{len(vs)} 票），"
                              f"少数派 {[v for v in vs if v != top_v]} 已记冲突",
                })
                return top_v
            if _anchored:
                for a in sorted(_anchored, key=lambda x: x.anchor_idx):
                    v = _get_prop(a, key)
                    if v is not None:
                        conflicts.append({
                            "role": role, "rule_id": rule.rule_id,
                            "text": group[0].text,
                            "reason": f"注解组内 {key} 分歧 {[v for v in vs]}"
                                      f"——早锚优先取 {v}（锚段 {a.anchor_idx}）",
                        })
                        return v
                return None
            conflicts.append({
                "role": role, "rule_id": rule.rule_id, "text": group[0].text,
                "reason": f"注解组内 {key} 分歧 {vs}——"
                          "疑似描述不同元素，无多数无锚点，挂起待人工裁定",
            })
            return None

        size = _mode_or_conflict([a.size_pt for a in group], "字号")
        ea = _mode_or_conflict([a.font_ea for a in group], "中文字体")
        latin = _mode_or_conflict([a.font_latin for a in group], "西文字体")
        bold = _mode_or_conflict([a.bold for a in group], "加粗")
        jc = _mode_or_conflict([a.jc for a in group], "对齐")
        line = _mode_or_conflict([a.line for a in group], "行距")
        rule_ = _mode([a.line_rule for a in group])
        ind2 = sum(1 for a in group if a.indent2) > len(group) / 2

        # ── 覆盖策略（EZI-0065 收编 2026-10-27 版）────────────────────────
        # 三层：
        #   A. 前区角色（题目/作者/单位/关键词/页眉页脚/参考文献条目）——注解
        #      明说就算数（用户裁定「模板原文为准」）；高置信（角色词在句首附近
        #      且整句只讲一个角色）+ 字号字体**双全**才覆盖，防半截注解改坏。
        #      实测案例：矿大大摘要 批注「论文题目：黑体…字号：小二…居中」——
        #      样例池把封面题目读成 10.5 宋体，注解是唯一正确的值来源；
        #      旧策略 allow_override=False 恒弃 → 取值必错。
        #   B. 标题（heading*）——维持 10-27 收紧版：只有「引擎取值明显不合理」
        #      （标题字号 ≤ 正文字号）且一级才整套覆盖（交大本科 heading1 硬错
        #      是这条救回来的；二三级与正文同字号是合法写法，公文伤疤维持）。
        #   C. 正文——永不覆盖，只补缺（矿大/湖南师大 body 加粗伤疤维持）。
        body_size = None
        for r in getattr(spec, "rules", []) or []:
            if r.rule_id.endswith(".body"):
                body_size = (r.expected or {}).get("font_size_pt")
                break
        spec_size = before.get("font_size_pt")
        complete_evidence = size is not None and ea is not None
        if role in ("title", "title_en", "authors", "authors_en", "affiliation",
                    "keywords", "references", "header_footer"):
            allow_override = high and complete_evidence
        elif role.startswith("heading"):
            # 只有「引擎取值明显不合理」（标题字号 ≤ 正文字号）才整套覆盖；
            # 其余只补缺 —— 否则公文模板「三级标题仿宋」会被注解里的宋体改掉。
            # 只认「取值存在且小于正文字号」为不合理；spec 缺字号属信息不足，
            # 只补缺不覆盖（公文模板三级标题 仿宋 曾被误改成宋体）
            implausible = (spec_size is not None and body_size is not None
                           and spec_size <= body_size + 0.1)
            # 只有一级标题允许整套覆盖（交大本科 heading1=五号宋体 这种硬错）；
            # 二三级标题与正文同字号是常见合法写法（公文三级标题 = 正文三号仿宋）
            allow_override = high and implausible and role == "heading1"
        else:
            allow_override = False
        allow_fill = high

        # 正文永不因注解变粗（2026-10-27：中国矿业/湖南师大模板实测会被加粗）
        if role == "body" and bold is True:
            bold = None
        # 一级/二级标题默认顶格：缺缩进也不补（公文标题缩进 = 版式错误）
        if role in ("heading1", "heading2") and ind2:
            ind2 = False

        skipped: dict = {}

        def _set(key, value):
            if value is None:
                return
            if key in before and before[key] is not None and before[key] != value:
                # EZI-0065 ⑧ 家族前缀守卫：spec「仿宋_GB2312」是注解「仿宋」的
                # 更具体同族形态，泛化注解不降级具体取值（视同一致，跳过）。
                if isinstance(value, str) and isinstance(before[key], str):
                    _a, _b = value.replace(" ", ""), before[key].replace(" ", "")
                    if _a != _b and (_a in _b or _b in _a):
                        same_family = (_a.rstrip("_GB2312") == _b.rstrip("_GB2312")
                                       if "_GB2312" in _a or "_GB2312" in _b else False)
                        if same_family:
                            return
                if allow_override:
                    exp[key] = value
                else:
                    skipped[key] = [before[key], value]
                return
            if allow_fill:
                exp[key] = value

        _set("font_size_pt", size)
        _set("font_east_asia", ea)
        _set("font_ascii", latin)
        if bold is not None:
            _set("bold", bold)
        _set("jc", jc)
        # 注意：spec 里该键可能存在但值为 None（提取器写空），不能用 `not in`
        if ind2 and not before.get("ind_first_line_chars"):
            if allow_fill:
                exp["ind_first_line_chars"] = 200
                exp["ind_first_line"] = 0
        if line is not None:
            _set("line", line)
            if not before.get("line_rule") and allow_fill:
                exp["line_rule"] = rule_ or "auto"

        changed = {k: [before.get(k), v] for k, v in exp.items() if before.get(k) != v}
        if not changed:
            continue
        if skipped and high:
            conflicts.append({
                "role": role, "rule_id": rule.rule_id, "text": group[0].text,
                "would_change": skipped,
                "reason": "注解与模板样例不一致（当前策略：正文只补缺、标题仅在不合理时覆盖），待确认",
            })
        if not high:
            conflicts.append({
                "role": role, "rule_id": rule.rule_id,
                "text": group[0].text, "would_change": changed,
                "reason": f"角色词位置 {group[0].role_pos}、命中 {group[0].role_hits} 个角色词，置信度不足",
            })
            continue
        rule.expected = exp
        applied.append({
            "role": role, "rule_id": rule.rule_id, "samples": len(group),
            "text": group[0].text, "changed": changed,
        })
    if conflicts:
        if not isinstance(getattr(spec, "structure", None), dict):
            spec.structure = {}
        spec.structure["annotation_conflicts"] = conflicts
    return applied


def annotation_summary(annotations: list[Annotation]) -> dict:
    by_role: dict[str, int] = {}
    for a in annotations:
        by_role[a.role or "(未映射)"] = by_role.get(a.role or "(未映射)", 0) + 1
    return {
        "total": len(annotations),
        "by_role": by_role,
        "unmapped": [a.text for a in annotations if not a.role][:10],
    }
