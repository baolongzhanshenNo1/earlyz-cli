"""TSD — Template Semantic Decompiler（模板语义反编译器）。

3.2 §5.4 规则来源「规范模板 OOXML > 模板示例段落」的落地实现。

背景（2026-08-15 根因审计）：uploaded_template_docx 适配器只读 styles.xml，
而真实高校模板的规范往往不在样式表里——西安理工高科学院模板 styles.xml
只有 9 个 header/footer 辅助样式——规范写在 document.xml 示例段落的直接
格式里（黑体小2居中章标题、宋体小4正文、逐段封面布局），并伴随
「（黑体小2加粗）」式注释段。旧提取器产出 7 条样式级规则 → lint 零段落
违规 → plan 全是 create_or_update_style → 排版能力为零。

本模块从模板反编译三类规则（全部确定性，零 LLM）：

1. 语义角色规则（selector.semantic_role）
   - 样本源 A：模板段落按 classify_paragraph_role 分类（与目标文档同一
     真相源）后的直接格式；
   - 样本源 B：标题形状检测——模板示例标题常无 heading 样式
     （'1□绪论' 被分类为 body），按「居中+粗+章序号文本」/层级序号
     正则补充 heading1/2/3 样本；
   - 样本源 C：heading/Normal 等样式定义（样式型模板兜底；段落样本
     优先，§5.4 示例段落 > 样式定义）；
   - 主导值 ≥ 阈值才成为期望值（模板自相矛盾时宁缺毋滥：行距 360auto
     与 460exact 并存 → 不出行距规则，保留目标文档现状）。
2. 节级规则（selector.scope=section_all）：模板各节全一致的页边距键。
3. 封面锚点规则（selector.cover_anchor）：封面是布局不是角色——按
   归一化文本精确/字段前缀/标题启发式/日期模式匹配，克隆模板封面格式。

排除项（污染源，全部确定性识别）：
- 注释段：（……黑体/宋体/号/磅……）格式说明文字（楷体_GB2312 居中小字）；
- 标签段：……示例：/……说明： 行；
- 目录行：含 ……/...... 点线导引；
- 参考文献条目：^[n]；
- 关键词行：关键词：/Key words；
- 封面/前置区：首个锚点（摘要/ABSTRACT/目录/结论/…/章标题）之前。
"""

from __future__ import annotations
import copy
import re
import xml.etree.ElementTree as etree
from collections import Counter
from pathlib import Path
from typing import Optional
from earlyz.ir.analyzer import analyze_document
from earlyz.ir.ooxml_format import merge_extra_bundles
from earlyz.spec.models import FormatRule, FormatSpec, SpecProvenance

# ── 文本工具 ────────────────────────────────────────────────────

_COVER_START_WORDS = {"摘要", "ABSTRACT", "Abstract", "目录", "目次"}
_BACK_MATTER_WORDS = {"参考文献", "致谢", "谢辞", "附录", "结论", "结语"}
_ANCHOR_WORDS = {
    "摘要", "ABSTRACT", "Abstract", "目录", "目次", "结论", "结语",
    "参考文献", "致谢", "谢辞", "附录", "引言", "前言", "绪论",
}

_CHAPTER_RE = re.compile(
    r"^(第[一二三四五六七八九十百\d]+[章节]"  # 第1章/第一章
    r"|[\d]{1,2}(?:[.．]\d{1,2}){0,2}\s*[\u4e00-\u9fff])"  # 1 绪论 / 1.1 概述 / 3.2.1 系统
)


def _chapterish_text(n: str) -> bool:
    """章节式文本证据（用于封面边界弱判据收紧，2026-09-25 审计）。

    只认「文本形态」证据：章节序号行 / 语义标题词表（摘要/目录/声明/任务书…）/
    数字或中文序号标题。不含样式推断——封面内的样式残留（作者给校名/题名套了
    标题样式、空段残留 heading）不再把边界拉进封面内部。
    """
    if not n:
        return False
    from earlyz.ir.analyzer import (
        _is_chapter_title_text, _is_cn_numbered_heading, _is_numeric_heading,
    )
    return bool(_CHAPTER_RE.match(n) or _is_chapter_title_text(n)
                or _is_numeric_heading(n) or _is_cn_numbered_heading(n))
# 层级分隔符兼容：1.1□ / 1．4．1□（□ 为全角空格占位显示）
_H2_RE = re.compile(r"^\d+[\.．]\d+[\s□　·]")
_H3_RE = re.compile(r"^\d+[\.．]\d+[\.．]\d+(?:[\.．]\d+)*[\s□　·]")
_TOC_DOTS_RE = re.compile(r"[…\.]{4,}")
_PLACEHOLDER_RE = re.compile(r"^[\s…\.。·□○●]+$")
_REF_RE = re.compile(r"^\[\d+\]")
_KW_RE = re.compile(r"^(关键词|Key\s*words)", re.I)
_ANNOTATION_RE = re.compile(
    r"^[（(][^）)]{0,80}(黑体|宋体|仿宋|楷体|Times|号|磅|加粗|居中|行距|间距|字符|粗体|段前|段后)[^）)]{0,80}[）)]$"
)
_LABEL_RE = re.compile(r"(示例|说明|举例)\s*[:：]\s*$")
_DATE_RE = re.compile(r"^[二〇○零一二三四五六七八九〇\d]{2,4}\s*年\s*[\d一二三四五六七八九十]{1,2}\s*月")
_FIELD_SPLIT = re.compile(r"[：:]")

# 封面字段别名（canonical 匹配用——模板「指导教师：」vs 目标「导师姓名、职称：」
# 2026-08-15 真机回归根因 C：字段名不同 → prefix 规则不命中 → 导师行不排版）
_FIELD_ALIAS_GROUPS: list[list[str]] = [
    ["指导教师", "指导老师", "导师姓名", "导师", "导师姓名职称",
     "指导教师姓名职称", "导师姓名及职称", "指导教师姓名及职称", "指导单位"],
    ["专业名称", "专业"],
    ["学生姓名", "姓名"],
    ["学生学号", "学号"],
    ["培养单位", "单位"],
    ["申请学位", "学位"],
    ["研究方向", "方向"],
    ["完成日期", "日期"],
]
_FIELD_CANON: dict[str, str] = {
    a: g[0] for g in _FIELD_ALIAS_GROUPS for a in g
}


def _canon_field(name: str) -> str:
    """字段名规范化：去空白/顿号逗号后查别名表（未命中返回原值）。"""
    k = re.sub(r"[\s　、，,·]", "", name or "")
    return _FIELD_CANON.get(k, k)


def _field_match(field: str, anchor: str) -> bool:
    """封面字段匹配：精确相等或别名同组（lint/writer 共用语义）。"""
    if field == anchor:
        return True
    return bool(field) and bool(anchor) and _canon_field(field) == _canon_field(anchor)

# 段落属性键（lint/writer 对齐；first_line_indent_chars 为 spec 惯例别名）
_RUN_KEYS = ("font_east_asia", "font_ascii", "font_size_pt", "bold")
_PPR_KEYS = (
    "jc", "ind_first_line_chars", "ind_first_line", "ind_left",
    "spacing_before", "spacing_after", "spacing_before_lines", "spacing_after_lines",
    "line", "line_rule", "line_spacing",
    "keep_next", "keep_lines", "page_break_before",
)
_COVER_KEYS = _RUN_KEYS + _PPR_KEYS


def norm_text(t: str | None) -> str:
    """锚点归一化：去空白与 □○●· 占位符。"""
    return re.sub(r"[\s□○●·.。…]+", "", t or "")


# 已有专属 exact 规则的封面锚点 norm 集合（title 形状匹配的排除表——
# 防止 title 规则后发覆盖 exact 规则；match_cover_paragraphs/writer 同源）
_EXACT_ANCHOR_NORMS = {norm_text(x) for x in
                       ("毕业设计（论文）", "毕业设计(论文)", *_ANCHOR_WORDS)}


def _is_annotation(text: str) -> bool:
    t = (text or "").strip()
    return bool(t) and bool(_ANNOTATION_RE.match(t))


def _is_label(text: str) -> bool:
    t = (text or "").strip()
    return bool(t) and bool(_LABEL_RE.search(t))


# 前置区可采样的角色（模板封面/前置区取值来源）
_FRONT_SAMPLE_ROLES = ("article_title", "article_title_en", "authors",
                       "authors_en", "affiliation", "affiliation_en",
                       "abstract", "abstract_en",
                       "keywords", "keywords_en", "classification")

# 独立说明行 → 上方样本行的角色提示（长词在前，避免「英文题名」被「题名」抢）
_ANNOT_ROLE_HINTS = (
    (("英文题名", "英文标题"), "article_title_en"),
    (("英文作者", "英文姓名"), "authors_en"),
    (("中英文题名", "中英文标题", "题名", "题目", "标题"), "article_title"),
    (("作者", "姓名"), "authors"),
    (("单位", "学院", "学校", "机构"), "affiliation"),
    (("摘要",), "abstract"),
    (("关键词", "关键字"), "keywords"),
    (("分类号",), "classification"),
)


def _role_hint_from_annotation(text: str) -> Optional[str]:
    """独立说明行 → 上方样本行的角色提示（2026-09-25 审计发现 2）。

    模板把「这是什么」写在示例下方的括号里（西安理工：「（2号黑体加粗居中，
    标题行间距为32磅）」写在真题名下方）。取首个命中，命中不到返回 None
    （只描述字体的说明行不给角色提示，不做模板专有词表）。"""
    t = (text or "").strip().strip("（()）")
    if not t:
        return None
    for toks, role in _ANNOT_ROLE_HINTS:
        if any(k in t for k in toks):
            return role
    return None


def _is_template_instruction(text: str) -> bool:
    """模板说明行：模板作者写给填表人的批注（「封面示例：」「学术声明示例：」）。

    这类行不是内容样本——它的格式是批注格式，采成规则会把目标文档同类段刷成
    批注样式（2026-09-25 真机：西安理工模板的首行「封面示例：」12pt →
    article_title 规则 = 12pt，题名规则只剩正文级取值）。
    只认尾词/起首词形态，不做模板专有词表（机制层，无学校常量）。
    """
    t = (text or "").strip().rstrip("：:")
    if not t or len(t) > 16:
        return False
    # 以内容标签起首的行是内容样本（「关键词 内容示例」「摘要 说明」），不是
    # 模板说明行——说明行是「封面示例：」这类页面/表格类提示（2026-09-25）。
    if re.match(r"^(?:关键词|关键字|摘\s*要|ABSTRACT|Abstract|中图|分类号|DOI"
                r"|收稿|基金|作者简介|通信作者|题目|题名|作者|单位|学号|姓名"
                r"|专业|导师|日期|Title|KEYWORDS)", t):
        return False
    return bool(re.search(r"(?:示例|范例|说明|格式|样张|模板)$", t)
                or re.match(r"^(?:注|备注|说明)", t))


def _is_toc_line(text: str) -> bool:
    return bool(_TOC_DOTS_RE.search(text or ""))


# 公式编号行（纯文本）：`(1)` / `（1）` / `(1.2)` / `(1-2)`——与公式正文同角色，
# 但排版是编号行的（前导空格/首行缩进/固定行距），不得学进公式规则。
_EQ_NUMBER_LINE_RE = re.compile(r"^[（(]\s*\d+(?:[.\-–—]\d+)*\s*[)）]$")


def _is_ref_entry(text: str) -> bool:
    return bool(_REF_RE.match((text or "").strip()))


# ── 规范注释解析（示例式模板的「（黑体小4加粗）」类注释 → 期望属性）──
# 2026-08-15 提取层根因修复：注释段此前只被「排除」，从不「解析」——
# 真实模板的正文规范常只写在注释里（（宋体小4，行距23磅）），导致
# body/heading 规则只剩 font_ascii 一个属性。解析后合并进角色样本。
_CN_SIZE_PT = {
    "初号": 42.0, "小初": 36.0, "一号": 26.0, "小一": 24.0, "二号": 22.0,
    "小二": 18.0, "三号": 16.0, "小三": 15.0, "四号": 14.0, "小四": 12.0,
    "五号": 10.5, "小五": 9.0, "六号": 7.5, "小六": 6.5, "七号": 5.5,
}
_SZ_RE = re.compile(r"(初号|小初|一号|小一|二号|小二|三号|小三|四号|小四|五号|小五|六号|小六|七号)")
# 阿拉伯数字字号（真实模板写「小4」「4号」不写「小四」）
_SZ_ARABIC_RE = re.compile(r"小([1-7])(?!\d)|([1-7])\s*号")
_ARABIC_SIZE_PT = {  # 序号→pt：1=初号系退化按 GB 表
    1: 26.0, 2: 22.0, 3: 16.0, 4: 14.0, 5: 10.5, 6: 7.5, 7: 5.5,
}
_ARABIC_Xiao_PT = {1: 24.0, 2: 18.0, 3: 15.0, 4: 12.0, 5: 9.0, 6: 6.5, 7: 5.0}
_PT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*磅")
_LINE_PT_RE = re.compile(r"行(?:间)?距[^，,。；;0-9]{0,6}(\d+(?:\.\d+)?)\s*磅|(\d+(?:\.\d+)?)\s*磅[^，,。；;]{0,4}行(?:间)?距")
_FONT_RE = re.compile(
    r"(仿宋_GB2312|楷体_GB2312|黑体|宋体|仿宋|楷体|Times\s+New\s+Roman)",
    re.I,
)


def parse_annotation_props(text: str) -> dict:
    """解析规范注释文本 → 格式属性 dict（解析失败返回 {}，绝不算错）。

    区分字号 X磅 与行距 X磅（「正文行间距固定为23磅」的 23 是行距不是
    字号——2026-08-15 真机回归曾把 heading3 污染成 23pt 字号）。
    """
    t = (text or "").strip()
    if not t or len(t) > 400:
        return {}
    out: dict = {}
    # 行距优先消化磅值，防止字号误吞
    ml = _LINE_PT_RE.search(t)
    if ml:
        pt = float(ml.group(1) or ml.group(2))
        out["line"] = int(round(pt * 20))
        out["line_rule"] = "exact"
    m = _SZ_RE.search(t)
    if m and m.group(1) in _CN_SIZE_PT:
        out["font_size_pt"] = _CN_SIZE_PT[m.group(1)]
    else:
        ma = _SZ_ARABIC_RE.search(t)
        if ma:
            if ma.group(1):  # 小N
                out["font_size_pt"] = _ARABIC_Xiao_PT.get(int(ma.group(1)))
            else:            # N号
                out["font_size_pt"] = _ARABIC_SIZE_PT.get(int(ma.group(2)))
        elif not ml:  # 无中文/阿拉伯字号时，行距未占用的裸磅值才算字号
            m2 = _PT_RE.search(t)
            if m2:
                try:
                    out["font_size_pt"] = float(m2.group(1))
                except ValueError:
                    pass
    out = {k: v for k, v in out.items() if v is not None}
    mf = _FONT_RE.search(t)
    if mf:
        f = mf.group(1)
        if re.match(r"times\s+new\s+roman", f, re.I):
            out["font_ascii"] = "Times New Roman"
        else:
            # 注释写「仿宋_GB2312」时保留全名，不截成「仿宋」
            if f.lower().startswith("仿宋") and "GB" in f.upper():
                out["font_east_asia"] = "仿宋_GB2312"
            elif f.lower().startswith("楷体") and "GB" in f.upper():
                out["font_east_asia"] = "楷体_GB2312"
            else:
                out["font_east_asia"] = f
    if "加粗" in t or "粗体" in t:
        out["bold"] = True
    if "居中" in t:
        out["jc"] = "center"
    if "两端对齐" in t:
        out["jc"] = "both"
    return out


def _is_keywords(text: str) -> bool:
    return bool(_KW_RE.match((text or "").strip()))


def _exclude_from_stats(text: str) -> bool:
    t = (text or "").strip()
    if not t or _PLACEHOLDER_RE.match(t):
        return True
    if (_is_annotation(t) or _is_label(t) or _is_toc_line(t)
            or _is_ref_entry(t) or _is_keywords(t)):
        return True
    # 短标签/公式号/图号残留（2026-08-15 body 池污染源）：(1.1)、(a)、目录示例:
    n = norm_text(t)
    if len(n) <= 6 and (re.match(r"^[（(\[]?[a-z0-9.]{1,5}[）)\]]?$", n)
                        or n.endswith(("示例:", "示例：", "说明:", "说明："))):
        return True
    return False


def _is_anchor(text: str, role: str | None, jc: str | None) -> bool:
    n = norm_text(text)
    if not n or len(n) > 16:
        return False
    if n.upper() in {a.upper() for a in _ANCHOR_WORDS}:
        return True
    if _CHAPTER_RE.match(n) and (jc == "center" or role == "heading1"):
        return True
    return False


def _field_name(text: str) -> str:
    """字段前缀：'专 业 名 称：计算机' → '专业名称'（无分隔符不是字段）。"""
    t = re.sub(r"[\s　]", "", text or "")
    if not _FIELD_SPLIT.search(t):
        return ""
    m = _FIELD_SPLIT.split(t, 1)
    return m[0] if m and m[0] and len(m[0]) <= 12 else ""


# ── 主导值 ──────────────────────────────────────────────────────

_DOMINANCE_DEFAULT = 0.6      # 常规属性
_DOMINANCE_STRICT = 0.75      # 行距/字体等易矛盾属性
_IND_DOMINANCE = 0.4          # 成对悬挂缩进（结构一致时取主导值，见 _expected_from）
_IND_PAIR_KEYS = ("ind_left", "ind_hanging")


def _mode_of(values: list, dominance: float, coverage_min: float = 0.5) -> tuple[Optional[object], float]:
    """主导值：先要求属性覆盖率（非 None 比例）≥ coverage_min，再在非 None
    值内部算主导占比。修复 2026-08-15 真机回归：None 被算进分母 →
    模板 body 255 段只有 46 段写了 font_size_pt（覆盖率 18%）→ 12.0pt
    在非空值内明明是 85% 主导却被 0.75 门槛杀掉 → body 规则只剩 font_ascii。
    覆盖率低 = 模板没写该属性（示例式模板常只写注释），宁缺勿错。"""
    if not values:
        return None, 0.0
    normed = [round(v, 3) if isinstance(v, float) else v for v in values]
    non_none = [v for v in normed if v is not None]
    if not non_none:
        return None, 0.0
    if len(non_none) / len(normed) < coverage_min:
        return None, len(non_none) / len(normed)
    cnt = Counter(non_none)
    val, n = cnt.most_common(1)[0]
    ratio = n / len(non_none)
    return (val, ratio) if ratio >= dominance else (None, ratio)


# ── 属性键策略（克制原则：宁可少排，不可错排）──────────────────

_ROLE_KEYS = {
    "body": ("font_east_asia", "font_ascii", "font_size_pt",
             "jc", "ind_first_line_chars", "ind_first_line",
             "spacing_before", "spacing_after", "line", "line_rule"),
    # 公式段（2026-09-27 用户拍板方案）：只取段落级版式——对齐/缩进/段距/行距/keep。
    # **不取字体字号**：公式 run 是 OLE（MathType）或 OMML，run 字体对它无意义，
    # 取了只会污染规则、并把公式当正文刷字体。
    "equation_display": ("jc", "ind_first_line_chars", "ind_first_line",
                         "ind_left", "ind_right", "ind_hanging",
                         "spacing_before", "spacing_after", "line", "line_rule",
                         "keep_next", "keep_lines"),
    "heading1": ("font_east_asia", "font_ascii", "font_size_pt", "bold", "italic",
                 "small_caps",
                 "jc", "spacing_before", "spacing_after",
                 "spacing_before_lines", "spacing_after_lines",
                 "line", "line_rule", "keep_next", "keep_lines"),
    "heading2": ("font_east_asia", "font_ascii", "font_size_pt", "bold", "italic",
                 "small_caps",
                 "jc", "spacing_before", "spacing_after",
                 "spacing_before_lines", "spacing_after_lines",
                 "line", "line_rule", "keep_next", "keep_lines"),
    "heading3": ("font_east_asia", "font_ascii", "font_size_pt", "bold", "italic",
                 "jc", "keep_next",
                 "line", "line_rule", "line_spacing",
                 "spacing_before", "spacing_after",
                 "spacing_before_lines", "spacing_after_lines"),
    "caption": ("font_east_asia", "font_ascii", "font_size_pt", "bold", "jc",
                "keep_lines", "spacing_before", "spacing_after",
                "spacing_before_lines", "spacing_after_lines", "line", "line_rule"),
    "caption_table": ("font_east_asia", "font_ascii", "font_size_pt", "bold",
                      "small_caps", "jc",
                      "spacing_before", "spacing_after",
                      "spacing_before_lines", "spacing_after_lines"),
    "caption_figure": ("font_east_asia", "font_ascii", "font_size_pt", "bold", "jc",
                       "spacing_before", "spacing_after",
                       "spacing_before_lines", "spacing_after_lines"),
    "toc1": ("font_east_asia", "font_ascii", "font_size_pt", "bold", "jc"),
    "toc2": ("font_east_asia", "font_ascii", "font_size_pt", "bold", "jc"),
    "toc3": ("font_east_asia", "font_ascii", "font_size_pt", "bold", "jc"),
    "refs_heading": ("font_east_asia", "font_ascii", "font_size_pt", "bold",
                     "italic", "small_caps",
                     "jc", "spacing_before", "spacing_after",
                     "spacing_before_lines", "spacing_after_lines",
                     "line", "line_rule", "outline_lvl", "label_text"),
    "ack_heading": ("font_east_asia", "font_ascii", "font_size_pt", "bold",
                    "italic", "small_caps",
                    "jc", "spacing_before", "spacing_after",
                    "spacing_before_lines", "spacing_after_lines",
                    "line", "line_rule", "label_text"),
    "appendix_heading": ("font_east_asia", "font_ascii", "font_size_pt", "bold",
                         "italic", "small_caps",
                         "jc", "spacing_before", "spacing_after", "label_text"),
    "bio_heading": ("font_east_asia", "font_ascii", "font_size_pt", "bold", "jc",
                    "label_text"),
    "refs_item": ("font_east_asia", "font_ascii", "font_size_pt", "jc",
                  "line", "line_rule", "ind_left", "ind_hanging",
                  "ind_first_line_chars"),
    "appendix_body": ("font_east_asia", "font_ascii", "font_size_pt", "jc",
                      "ind_left", "ind_hanging", "line", "line_rule"),
    "bio_body": ("font_east_asia", "font_ascii", "font_size_pt", "bold",
                 "jc", "ind_left", "ind_hanging"),
    "keywords": ("font_east_asia", "font_ascii", "font_size_pt", "jc",
                 "ind_first_line", "ind_first_line_chars", "ind_left", "ind_hanging",
                 "tab_pos", "line", "line_rule",
                 "label_font_east_asia", "label_font_ascii", "label_font_size_pt",
                 "label_bold", "label_italic", "label_text", "label_after_tab"),
    "keywords_en": ("font_east_asia", "font_ascii", "font_size_pt", "bold", "jc",
                    "ind_left", "ind_hanging", "tab_pos",
                    "label_font_east_asia", "label_font_ascii", "label_font_size_pt",
                    "label_bold", "label_italic", "label_text", "label_after_tab"),
    "article_title": ("font_east_asia", "font_ascii", "font_size_pt", "bold",
                      "italic", "small_caps", "jc",
                      "spacing_before", "spacing_after",
                      "spacing_before_lines", "spacing_after_lines"),
    "article_title_en": ("font_east_asia", "font_ascii", "font_size_pt", "bold",
                         "italic", "small_caps", "jc",
                         "spacing_before", "spacing_after",
                         "spacing_before_lines", "spacing_after_lines"),
    "authors": ("font_east_asia", "font_ascii", "font_size_pt", "bold",
                "italic", "jc",
                "spacing_before", "spacing_after",
                "spacing_before_lines", "spacing_after_lines",
                "line", "line_rule"),
    "authors_en": ("font_east_asia", "font_ascii", "font_size_pt", "bold",
                   "italic", "jc",
                   "spacing_before", "spacing_after",
                   "spacing_before_lines", "spacing_after_lines",
                   "line", "line_rule"),
    "affiliation": ("font_east_asia", "font_ascii", "font_size_pt", "jc",
                    "spacing_before", "spacing_after"),
    "affiliation_en": ("font_east_asia", "font_ascii", "font_size_pt", "jc",
                       "spacing_before", "spacing_after"),
    "abstract": ("font_east_asia", "font_ascii", "font_size_pt", "bold",
                 "italic", "jc", "tab_pos", "ind_first_line", "ind_first_line_chars",
                 "spacing_before", "spacing_after", "line", "line_rule",
                 "label_font_east_asia", "label_font_ascii", "label_font_size_pt",
                 "label_bold", "label_italic", "label_text", "label_after_tab"),
    "abstract_en": ("font_east_asia", "font_ascii", "font_size_pt", "bold",
                    "italic", "jc", "tab_pos",
                    "label_font_east_asia", "label_font_ascii", "label_font_size_pt",
                    "label_bold", "label_italic", "label_text", "label_after_tab"),
    "classification": ("font_east_asia", "font_ascii", "font_size_pt", "bold",
                       "jc", "tab_pos", "spacing_after", "spacing_after_lines",
                       "ind_first_line", "ind_first_line_chars",
                       "label_font_east_asia", "label_font_ascii", "label_font_size_pt",
                       "label_bold", "label_text", "label_after_tab"),
    "figure_para": ("jc", "ind_left", "keep_next", "keep_lines"),
    "drop_cap": ("font_east_asia", "font_ascii", "font_size_pt", "small_caps",
                 "jc", "line", "line_rule", "keep_next",
                 "spacing_before", "spacing_after",
                 "drop_cap", "drop_cap_lines"),
    "list_item": ("font_east_asia", "font_ascii", "font_size_pt", "jc"),
}
# 位置性伪影键：模板示例段的分页/锚定属性来自其在示例文档中的位置，
# 不是角色规范（章标题示例恰好落在分页点 → page_break_before=true
# 会给摘要/目录标题强加分页）——永不进角色规则。
# keep_next/keep_lines 的显式 false 同属伪影（2026-08-17 D2：模板示例区
# 11 个 keepNext 全为 w:val="0"——作者防示例跨页的位置性操作，不是
# "标题不该 keepNext"的规范；曾被推断成 keep_next:false 去关产物标题，
# 违反 08-08 铁律②图片/标题防跨页）。
_POSITIONAL_KEYS = {"page_break_before", "keep_next", "keep_lines"}
# 高矛盾风险属性（模板示例常自相矛盾 → 提高门槛）
_STRICT_KEYS = {"line", "line_rule", "line_spacing", "font_east_asia",
                "spacing_before", "spacing_after"}

_H_LEVEL_BY_NAME = re.compile(r"^(?:heading\s*(\d)|标题\s*(\d))", re.I)


_PORTABLE_TYPED_KEYS = (
    "font_east_asia", "font_ascii", "font_size_pt", "bold", "italic",
    "underline", "color", "small_caps", "jc",
    "ind_first_line", "ind_first_line_chars", "ind_left", "ind_right", "ind_hanging",
    "tab_pos", "spacing_before", "spacing_after", "spacing_before_lines",
    "spacing_after_lines", "line", "line_rule", "line_spacing",
    "keep_next", "keep_lines", "widow_control", "page_break_before",
    "outline_lvl", "snap_to_grid",
)


_EQ_NUM_ANY_RE = re.compile(r"[（(]\s*\d{1,3}(?:\s*[.\-–]\s*\d{1,3})?\s*[)）]")


def _equation_number_style(ir) -> tuple:
    """从**模板的公式样本**读编号形态（2026-09-28 用户拍板「模板里放公式样本」）。

    返回 (format, paren)：`"(N)"` 顺序编号 / `"(C-N)"` 章编号；`"()"`/`"（）"` 括号形态。
    模板里没有带编号的公式样本 → (None, None)：调用方不得发明（保持保守默认）。
    """
    nums = []
    for p in getattr(ir, "paragraphs", None) or []:
        if getattr(p, "semantic_role", None) != "equation_display":
            continue
        text = (getattr(p, "text", "") or "").strip()
        m = _EQ_NUM_ANY_RE.search(text)
        if m:
            nums.append(m.group(0))
    if not nums:
        return None, None
    fmt = "(C-N)" if any(re.search(r"[.\-–]", n) for n in nums) else "(N)"
    paren = "（）" if any(n.startswith("（") for n in nums) else "()"
    return fmt, paren


def _allowed_keys(role: str) -> tuple:
    keys = _ROLE_KEYS.get(role, ())
    # Typed + lossless extras are candidates for every known role. Dominance and
    # positional-artifact gates still decide whether a property becomes a rule.
    return tuple(dict.fromkeys((*keys, *_PORTABLE_TYPED_KEYS,
                                "ppr_extra_xml", "rpr_extra_xml"))) if keys else ()


# ── 段落样本收集 ────────────────────────────────────────────────

def _fp_dict(fp) -> dict:
    if fp is None:
        return {}
    return {k: v for k, v in fp.__dict__.items() if v is not None}


def _merge_fp_dict(base: dict, extra: dict) -> dict:
    """Merge typed properties plus lossless OOXML extras across the cascade."""
    out = dict(base or {})
    for k, v in (extra or {}).items():
        if v is None:
            continue
        if k == "ppr_extra_xml":
            out[k] = merge_extra_bundles(out.get(k), v, "pPr")
        elif k == "rpr_extra_xml":
            out[k] = merge_extra_bundles(out.get(k), v, "rPr")
        else:
            out[k] = v
    return out


def _bundle_has_child(xml_text: Optional[str], local: str) -> bool:
    if not xml_text:
        return False
    try:
        root = etree.fromstring(xml_text)
    except etree.ParseError:
        return False
    return any(child.tag.rsplit("}", 1)[-1] == local for child in list(root))


def _resolved_style_map(ir) -> dict:
    """样式 basedOn 链：子样式覆盖父样式。分类号 sz 在 Date1 上、黑体在自身上。"""
    raw = {st.style_id: _fp_dict(st.style_format) for st in (ir.styles or [])}
    based = {st.style_id: st.based_on for st in (ir.styles or [])}
    resolved: dict = {}

    def walk(sid, seen=None):
        if not sid:
            return {}
        if sid in resolved:
            return resolved[sid]
        seen = seen or set()
        if sid in seen or sid not in raw:
            return {}
        seen.add(sid)
        parent = walk(based.get(sid), seen)
        child = raw.get(sid) or {}
        out = _merge_fp_dict(parent, child)
        resolved[sid] = out
        return out

    for sid in raw:
        walk(sid)
    return resolved


def _default_para_style_merged(ir, style_by_id: dict) -> dict:
    """无 pStyle 段继承的默认段落样式（w:style default=1 paragraph，机制层
    OOXML 语义）沿 basedOn 链合并结果；无 default 样式则返回空。"""
    for st in (ir.styles or []):
        if (getattr(st, "style_type", "paragraph") == "paragraph"
                and getattr(st, "is_default", False)):
            return style_by_id.get(st.style_id) or {}
    return {}


_NAMED_STYLE_ROLES = (
    (re.compile(r"^figure caption$", re.I), "caption_figure"),
    (re.compile(r"^(table title|table caption|表名)$", re.I), "caption_table"),
    (re.compile(r"^references$", re.I), "refs_item"),
    (re.compile(r"^authors$", re.I), "authors"),
)
_STYLE_SEED_KEYS = (
    "font_east_asia", "font_ascii", "font_size_pt", "bold", "italic", "small_caps",
    "spacing_before", "spacing_after", "spacing_before_lines", "spacing_after_lines",
    "drop_cap", "drop_cap_lines", "ppr_extra_xml", "rpr_extra_xml",
)


def _seed_named_styles(ir, para_samples: dict) -> None:
    """段落样本没有 Fig.1 / [1] 时，用模板同名样式补角色取值。"""
    style_by_id = _resolved_style_map(ir)
    for st in ir.styles or []:
        name = (st.name or "").strip()
        for pat, role in _NAMED_STYLE_ROLES:
            if pat.match(name):
                resolved = style_by_id.get(st.style_id) or {}
                own = _fp_dict(st.style_format)
                d = {k: v for k, v in resolved.items()
                     if v is not None and k in _STYLE_SEED_KEYS}
                # jc/悬挂只取样式自己写的，不继承 Normal 的两端对齐
                for k in ("jc", "ind_hanging", "ind_left", "ind_first_line",
                          "ind_first_line_chars"):
                    if own.get(k) is not None:
                        d[k] = own[k]
                if d:
                    para_samples.setdefault(role, []).append(d)
                break


def _label_props_from_para(p, style_extra: dict | None = None) -> dict:
    """摘要/关键词标签 run 的字体，写入 label_* 键（内容 run 走普通键）。

    模板标签 run 常无直接 sz（CJC 摘要标签只有 szCs=18），有效字号在
    样式链（af8→a0→Normal=18）——标签 run 缺字号时回退到段落样式链，
    否则源稿 docDefaults（如 29=14.5pt）会被无端继承。
    """
    from earlyz.ir.analyzer import _is_front_label_run
    out: dict = {}
    for run in getattr(p, "runs", None) or []:
        if not _is_front_label_run(run.text or ""):
            continue
        lf = _fp_dict(getattr(run, "direct_format", None))
        if lf.get("font_east_asia"):
            out["label_font_east_asia"] = lf["font_east_asia"]
        if lf.get("font_ascii"):
            out["label_font_ascii"] = lf["font_ascii"]
        if lf.get("font_size_pt") is not None:
            out["label_font_size_pt"] = lf["font_size_pt"]
        elif style_extra and style_extra.get("font_size_pt") is not None:
            out["label_font_size_pt"] = style_extra["font_size_pt"]
        if lf.get("bold") is not None:
            out["label_bold"] = lf["bold"]
        elif style_extra and style_extra.get("bold") is not None:
            out["label_bold"] = style_extra["bold"]
        if lf.get("italic") is not None:
            out["label_italic"] = lf["italic"]
        raw = (run.text or "").strip()
        if raw:
            out["label_text"] = raw
        # 模板标签 run 与内容之间常有 <w:tab/>（独立 run；Word 显示为
        # 「标签→内容」箭头）。标签完成标志（label_after_tab）由模板段落
        # run 结构决定（有就加、没有就减——2026-08-22 用户裁决）。
        runs = getattr(p, "runs", None) or []
        for idx, rn in enumerate(runs):
            if rn is run:
                nxt = runs[idx + 1] if idx + 1 < len(runs) else None
                if nxt is not None and getattr(nxt, "has_tab", False):
                    out["label_after_tab"] = True
                break
        break
    return out


_INLINE_ANNOT_RE = re.compile(r"[（(]([^（）()]{0,40}(?:黑体|宋体|仿宋|楷体|Times|号|磅|加粗|居中|行距|粗体)[^（）()]{0,40})[）)]")


def _extract_inline_annotation(text: str) -> str:
    """行内注释：段文本中部/尾部的（黑体4号加粗）或 *小5号* → 独立注释串。"""
    m = _INLINE_ANNOT_RE.search(text or "")
    if m:
        return m.group(1)
    m2 = re.search(
        r"[*＊]([^*＊]{0,80}(?:黑体|宋体|仿宋|楷体|Times|号|磅|加粗|居中)[^*＊]{0,80})",
        text or "",
    )
    return m2.group(1) if m2 else ""


def _is_toc_region(p) -> bool:
    """目录区判定（样本收集期）：段前紧邻目录锚点或同行含点 leader。
    简化语义：文本含点 leader（……）或空段。保守起见仅用于排除污染，
    不用于定位。"""
    t = (p.text or "").strip()
    if not t:
        return False
    return bool(_TOC_DOTS_RE.search(t))



def _style_effective_map(tpl: Path) -> tuple[dict[str, dict], dict]:
    """styles.xml 段落样式的**有效格式**（含 basedOn 链继承），返回 ({style_id: dict}, docDefaults dict)。

    采样层原来只读段落**直接格式**，模板把缩进/字体写在样式里时（建设工程合同族
    首行 440 写在段落样式、数据中介合同 200 写在样式）就整键丢失 → 产物没有缩进，
    而 Word 里这些段落看起来是有缩进的。这是模型缺陷，不是个别模板问题。
    """
    import zipfile
    from lxml import etree
    _W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    from earlyz.ir.analyzer import _parse_ppr_format, _parse_rpr_format, _merge_format, _parse_doc_defaults
    try:
        with zipfile.ZipFile(str(tpl)) as z:
            styles_xml = z.read("word/styles.xml")
    except Exception:
        return {}, {}
    try:
        root = etree.fromstring(styles_xml)
    except Exception:
        return {}, {}
    raw: dict[str, dict] = {}
    for st in root.iter(f"{_W}style"):
        if st.get(f"{_W}type") != "paragraph":
            continue
        sid = st.get(f"{_W}styleId")
        if not sid:
            continue
        ppr = st.find(f"{_W}pPr")
        rpr_el = st.find(f"{_W}rPr")
        merged = _merge_format(_parse_ppr_format(ppr),
                               _parse_rpr_format(rpr_el))
        base = st.find(f"{_W}basedOn")
        raw[sid] = {"d": _fp_dict(merged),
                    "base": base.get(f"{_W}val") if base is not None else None}

    def resolve(sid: str, depth: int = 0, seen=None) -> dict:
        seen = seen or set()
        if sid in seen or depth > 6 or sid not in raw:
            return {}
        seen.add(sid)
        out = resolve(raw[sid]["base"], depth + 1, seen) if raw[sid]["base"] else {}
        out.update(raw[sid]["d"])   # 近者优先
        return out

    eff = {sid: resolve(sid) for sid in raw}
    try:
        defaults = dict(_parse_doc_defaults(styles_xml.decode("utf-8", errors="replace")))
    except Exception:
        defaults = {}
    return eff, defaults


def _collect_samples(ir, style_map: dict | None = None,
                     doc_defaults: dict | None = None) -> tuple[dict[str, list[dict]], int, dict]:
    """按角色收集模板段落直接格式样本；返回 (样本表, cover_end, body_norm)。 封面边界索引)。

    2026-08-15 提取层根因修复（真机回归「整篇样式没应用」）：
    - 目录区跟踪：无点 leader 的目录条目（「3□关于海量图像无级缩放」）
      也排除，不再污染 body 池；
    - 注释双向就近合并：规范注释（（宋体小4））修饰紧邻的示例段（前或后），
      注释明说的键覆盖段落实况（模板示例段自身常不规范）；
    - □□ 假缩进推导：模板作者用全角空格伪造首行缩进 → ind_first_line_chars=200；
    - 「正文…行距」注释 = 全局正文行距规范，应用到全部 body 样本。
    """
    cover_end = 0
    _style_map = style_map or {}
    _defaults = doc_defaults or {}

    def _eff(p) -> dict:
        """段落有效格式：直接格式 > 样式链（basedOn 近者优先）> docDefaults。"""
        d = _fp_dict(getattr(p, "direct_format", None))
        sid = getattr(p, "style_id", None)
        for layer in (_style_map.get(sid) or {}, _defaults):
            for k, v in layer.items():
                if d.get(k) is None and v is not None:
                    d[k] = v
        return d

    from earlyz.ir.analyzer import _is_chapter_title_text, heading_title_key
    # 同名标题样本（2026-09-29）：摘要/ABSTRACT/目录/声明/结论… 这类不带编号的章级
    # 标题各自按标题名单独成池（键 `heading1#摘要`），不再与编号章标题合成一条
    # heading1 规则。legacy_pool = 旧口径下本会进 heading 池的那批（有 heading 样式的
    # 同名标题），仅在模板**没有任何编号章标题样本**时回填，保持旧模板行为不变。
    legacy_pool: dict[str, list] = {}
    for i, p in enumerate(ir.paragraphs):
        n = norm_text(p.text)
        if n and n.upper() in {a.upper() for a in _BACK_MATTER_WORDS}:
            continue
        role = p.semantic_role or ""
        # 锚点必须有文字（2026-09-25 审计，与目标侧 find_cover_end 同条件）：
        # 模板首段空段残留 heading 样式 → 此处 cover_end=0 → 模板封面（36pt 校名/
        # 校徽）整段进 body 样本池，取值被封面大字号污染。
        if n and str(role).startswith("heading") and not _is_chapter_title_text(p.text or ""):
            cover_end = i
            break
        if not n or len(n) > 16:
            continue
        d = _eff(p)
        if n.upper() in {a.upper() for a in _COVER_START_WORDS}:
            cover_end = i
            break
        if _is_anchor(p.text, p.semantic_role, d.get("jc")):
            cover_end = i
            break

    samples: dict[str, list[tuple[int, dict]]] = {}
    pending: dict | None = None
    pending_at: int = -1
    body_norm: dict = {}
    region: str | None = None   # 前置区跟踪（摘要/ABSTRACT/目录/致谢…）
    front_run: list[tuple[int, dict]] = []   # 说明行上方的连续样本（角色提示用）
    _ANCHOR_UP = {a.upper() for a in _ANCHOR_WORDS}
    _FRONT_REGIONS = {a.upper() for a in
                      ("摘要", "ABSTRACT", "目录", "目次", "致谢", "谢辞",
                       "参考文献", "附录", "结论", "结语", "引言", "前言")}
    style_by_id = _resolved_style_map(ir)
    style_name_by_id = {st.style_id: (st.name or "")
                        for st in (ir.styles or [])}
    from earlyz.ir.analyzer import _is_authorish_line
    for i, p in enumerate(ir.paragraphs):
        # 模板说明行（「封面示例：」「学术声明示例：」）不是内容样本：其格式是
        # 模板作者的批注格式，采成规则会把目标文档同类段刷成批注样式
        # （2026-09-25 真机：西安理工模板 → article_title 规则 = 12pt，正是
        # 「封面示例：」这行的格式）。两条采样路径（封面/正文）共用此排除。
        if _is_template_instruction(p.text or ""):
            continue
        if i < cover_end:
            role_f = p.semantic_role or ""
            raw = (p.text or "").strip()
            # 独立成段的模板说明行（「（1号宋体居中）」「（2号黑体加粗居中，标题行
            # 间距为32磅）」）不是内容样本：它的格式是说明格式，采成规则会把目标
            # 文档同类段刷成说明样式（2026-09-25 审计实锤：authors_en 规则 =
            # 10.5pt 正是这两行）；同时它对**上方**样本行给角色提示，「标题」→
            # article_title——西安理工模板的真题名（22pt 黑体加粗居中）因此才能
            # 被采成题名取值（此前落在 authors_en，目标题名被套成 10.5pt）。
            if _is_annotation(raw):
                hint = _role_hint_from_annotation(raw)
                if hint:
                    for j, dj, old_role in front_run:
                        # 搬迁（不是复制）：从原角色样本里摘掉，归到提示角色
                        samples[old_role] = [(k, dk) for k, dk in
                                             (samples.get(old_role) or []) if k != j]
                        samples.setdefault(hint, []).append((j, dj))
                front_run = []
                continue
            # 封面区里的前置页标题（学位论文「原创性声明」「版权使用授权书」常在首个
            # 锚点词「摘要」之前）：只采成**同名**标题取值，不进一级标题池（2026-09-29）。
            _tk = (heading_title_key(raw)
                   if str(role_f).startswith("heading") else None)
            if _tk:
                d = _merge_fp_dict(style_by_id.get(p.style_id) or {},
                                   _eff(p))
                if d:
                    samples.setdefault(f"{role_f}#{_tk}", []).append((i, d))
                front_run = []
                continue
            if role_f in _FRONT_SAMPLE_ROLES:
                d = _eff(p)
                extra = style_by_id.get(p.style_id) or {}
                if extra:
                    d = _merge_fp_dict(extra, d)
                d.update(_label_props_from_para(p, style_by_id.get(p.style_id) or {}))
                _apply_star_annotation(raw, d)
                # 形态核对：作者行必须真的像作者行（名字列表/作者标签），否则不学
                # 规则——西安理工模板 [5]「毕业设计（论文）」被判 authors，26pt
                # 居中加粗被采成「作者」取值，目标文档题名/日期被刷成大字号。
                if d and (role_f not in ("authors", "authors_en")
                          or _is_authorish_line(raw)):
                    samples.setdefault(role_f, []).append((i, d))
                    front_run.append((i, d, role_f))
            elif raw:
                front_run = []
            continue  # 封面/前置区不进 body 统计
        text = (p.text or "").strip()
        n = norm_text(text)
        d0 = _eff(p)
        is_anchor_word = (not _is_toc_line(text) and n.upper() in _ANCHOR_UP
                          and len(n) <= 16)
        # ── 区域状态机：锚点词进入对应前置区；真章标题（居中加粗）退出 ──
        if is_anchor_word:
            region = n
        elif (str(p.semantic_role or "").startswith("heading")
              and n and len(n) <= 20 and not is_sentence_like(text)
              and not _is_toc_line(text)):
            # 无编号的章级标题也必须退出前置区：模板用「引言/概述/系统设计」
            # 这类不带序号的标题时，区域状态机原来永远停在「目录/摘要」，
            # 后面整篇正文被当前置区跳过 → body 样本 0 条 → 正文规则退化
            # （2026-10-29 产品安装部署文档：61 段正文，样本池 0）。
            region = None
        elif _CHAPTER_RE.match(n) and (
                d0.get("bold") or d0.get("jc") == "center"
                or str(p.semantic_role or "").startswith("heading")
                or re.search(r"heading", style_name_by_id.get(p.style_id or "", ""), re.I)):
            # 退出前置区：章标题形态即可。原条件要求「居中且加粗」，模板章标题
            # 只加粗不居中（或只居中不加粗）时永远不退出 → 后面正文段全被当
            # 前置区跳过，body 样本 0 条 → 正文规则退化成样式默认值
            # （2026-10-28 大连民族实测：产物正文 10.5pt/无缩进/240 行距）。
            region = None
        in_front = region is not None and region.upper() in _FRONT_REGIONS

        # ── 注释段处理（整段注释 + 行内注释掏空后的纯注释段）──
        whole_ann = _is_annotation(text)
        inline_txt = "" if whole_ann else _extract_inline_annotation(text)
        inline = parse_annotation_props(inline_txt) if inline_txt else {}
        residue = _INLINE_ANNOT_RE.sub("", text).strip() if inline_txt else text
        if whole_ann or (inline and not residue):
            props = parse_annotation_props(text) if whole_ann else inline
            if props:
                # 全局正文行距规范（「正文行间距固定为23磅」）
                if "正文" in text and ("line" in props or "line_spacing" in props):
                    body_norm.update({k: v for k, v in props.items()
                                      if k in ("line", "line_rule", "line_spacing")})
                merged = False
                for role_key, lst in samples.items():
                    if lst and lst[-1][0] == i - 1:
                        if role_key == "keywords":
                            continue  # 关键词分工特判（标签/内容），注释不合并
                        _apply_annotation(lst[-1][1], props)
                        merged = True
                        break
                if not merged:
                    pending, pending_at = props, i
            continue
        role0 = p.semantic_role
        # 图片段（drawing 无文字）：_exclude_from_stats 的空文本排除会误杀，
        # 直接按角色收集（模板图示例段的 jc=center 等对齐规则）
        if role0 in ("figure_para", "keywords", "keywords_en", "abstract", "abstract_en",
                     "classification", "drop_cap",
                     "refs_heading", "ack_heading", "appendix_heading", "bio_heading",
                     "appendix_body", "bio_body"):
            # 附录/作者简介内容：空白占位段与带行内注记的说明段不是内容样本
            # （2026-10-09 用户定案「附录A是章标题，其余内容属于正文」；模板
            # 附录区常只有「正文部分…（宋体小4…）」这类版式注记，其 jc=center
            # 是模板展示效果——学成规则会把内容排成居中。规则改由正文继承，
            # 见角色规则组装段 1b）。
            if role0 in ("appendix_body", "bio_body"):
                _t = (p.text or "").strip()
                if not _t or _extract_inline_annotation(_t):
                    continue
            d = _eff(p)
            extra = style_by_id.get(p.style_id) or {}
            if extra:
                d = _merge_fp_dict(extra, d)
            # 有效字号链补全（2026-08-23 CJC v9 实锤）：附录内容/作者简介段 run 级
            # 无 sz（小五由 Normal 提供、说明段直接无 sz）——补链后规则才有
            # font_size_pt，否则目标段落入源稿 docDefaults（14.5）（同页眉页脚机制）。
            if d.get("font_size_pt") is None:
                _def_sz = _default_para_style_merged(ir, style_by_id).get("font_size_pt")
                if _def_sz is not None:
                    d["font_size_pt"] = _def_sz
            d.update(_label_props_from_para(p, style_by_id.get(p.style_id) or {}))
            _apply_star_annotation(p.text or "", d)
            raw = re.split(r"[*＊]", (p.text or "").strip())[0].strip()
            if raw and role0.endswith("_heading"):
                d["label_text"] = raw
            if d:
                samples.setdefault(role0, []).append((i, d))
            continue
        # 公式段（2026-09-27）：`equation_display` 段**文本为空**（公式本体是 OLE/OMML），
        # 会被 `_exclude_from_stats`（空文本）挡掉 → 模板真带公式样本时也学不到取值。
        # 该角色只取段落级版式，空文本无妨，单独放行。
        # 但**公式编号行**（纯文本 "(1.2)"／"（1）"，与公式正文同角色）必须排除：
        # 它的排版（前导空格 + 首行缩进 + 固定行距）是编号行的，不是公式的——
        # 学进来会把公式推得更右（西安理工模板实测 250/375 + line=300 exact）。
        # 与「图/表题分流」同一原则：取值层不能把不同形态合成一条规则。
        if (p.semantic_role != "refs_item" and _exclude_from_stats(p.text)
                and p.semantic_role != "equation_display"):
            continue
        if (p.semantic_role == "equation_display"
                and _EQ_NUMBER_LINE_RE.match((p.text or "").strip())):
            continue
        d = _eff(p)
        role = p.semantic_role
        title_pool = None
        if str(role or "").startswith("heading"):
            sname = style_name_by_id.get(p.style_id) or ""
            styled = bool(re.search(r"heading", sname, re.I))
            from earlyz.ir.analyzer import _is_numeric_heading
            _tk = heading_title_key(text)
            if _tk:
                # 同名标题：无论有无 heading 样式都采（它的取值就是模板对这个标题的定义），
                # 但只进同名池；有 heading 样式者同时记入 legacy_pool（旧口径回填用）
                title_pool = (f"{role}#{_tk}", styled and not _is_numeric_heading(text))
            elif (_is_chapter_title_text(text) and not styled
                    and not _is_numeric_heading(text)):
                continue
            extra = style_by_id.get(p.style_id) or {}
            if extra:
                d = _merge_fp_dict(extra, d)
        elif role in ("article_title", "article_title_en", "authors",
                      "authors_en", "affiliation", "affiliation_en",
                      "abstract", "abstract_en",
                      "keywords", "keywords_en", "classification"):
            extra = style_by_id.get(p.style_id) or {}
            if extra:
                d = _merge_fp_dict(extra, d)
            d.update(_label_props_from_para(p, style_by_id.get(p.style_id) or {}))
            _apply_star_annotation(p.text or "", d)
        elif role == "body":
            normal_extra = {}
            for st in ir.styles or []:
                if (st.name or "").lower() in ("normal", "正文"):
                    normal_extra = _fp_dict(st.style_format)
                    break
            extra = _merge_fp_dict(normal_extra, style_by_id.get(p.style_id) or {})
            if extra:
                d = _merge_fp_dict(extra, d)
        if not d:
            continue

        role = p.semantic_role
        jc = d.get("jc")
        if role in (None, "body"):
            if is_anchor_word:
                role = "heading1"
            elif _CHAPTER_RE.match(n) and jc == "center" and d.get("bold"):
                role = "heading1"
            elif _CHAPTER_RE.match(n) and region in ("目录", "目次"):
                continue       # 目录条目（无点 leader）不进任何角色池
            elif _H3_RE.match(text) and not in_front:
                role = "heading3"
            elif _H2_RE.match(text) and not _is_toc_line(text) and not in_front:
                role = "heading2"
            elif role == "body" and in_front:
                # 摘要正文必须单独成池（2026-10-29）：大摘要类模板整篇就是摘要，
                # 正文段落全在「摘要」前置区里——一律跳过会导致 body 池 0 条，
                # 正文规则退化成注解值（矿大 1-1-2：模板 12pt，产物 10.5pt）。
                # 只收摘要区、只收有实质内容的段落；目录/致谢/声明区仍然跳过。
                if str(region or "").upper() in ("摘要", "ABSTRACT", "中文摘要", "英文摘要"):
                    if n and len(n) >= 8:
                        samples.setdefault("abstract_body", []).append((i, d))
                continue       # 摘要/目录/致谢等前置区段落不进 body 池
            elif role0 == "keywords" and in_front:
                pass           # 关键词行（摘要前置区）——专门角色，正常收集
        if not role:
            continue
        # 图/表题分流（取值层：不能合成一条 caption=黑体，否则图题被错刷）
        if role == "caption":
            if re.match(r"^表\s*[\dXx]", text):
                role = "caption_table"
            elif re.match(r"^图\s*[\dXx]", text):
                role = "caption_figure"
        # 行内注释（「1.1□概述（黑体4号加粗）」）合并进当前段
        if inline:
            _apply_annotation(d, inline)
        # 紧邻前位注释合并（注释在示例段之后：回头注释）
        # keywords 例外（2026-08-16）：模板「关键词：...（黑体小4加粗）（宋体小4）」
        # 注释描述的是标签/内容分工，已由 analyzer 内容 run 提取+writer 特判实现，
        # 再合并会把整段刷成黑体加粗
        if pending is not None and pending_at == i - 1 and role0 != "keywords":
            _apply_annotation(d, pending)
            pending = None
        # □□ 假缩进推导（模板作者用全角空格伪造缩进）
        lead = re.match(r"^[□　\u3000]+", (p.text or ""))
        if lead and role == "body" and "ind_first_line_chars" not in d:
            chars = len(lead.group(0))
            if 1 <= chars <= 4:
                d["ind_first_line_chars"] = chars * 100
                sz = d.get("font_size_pt") or 12.0
                d.setdefault("ind_first_line", int(round(sz * 20)) * chars)
        if title_pool is not None and title_pool[0].startswith(f"{role}#"):
            samples.setdefault(title_pool[0], []).append((i, d))
            if title_pool[1]:
                legacy_pool.setdefault(role, []).append((i, d))
            continue
        samples.setdefault(role, []).append((i, d))

    # 旧口径回填：模板没有编号章标题样本（heading 池空）时，有 heading 样式的同名标题
    # 仍作该级标题取值——与改前一致（只是不再作用于同名标题本身，见 skip_title_keys）
    for _role, _lst in legacy_pool.items():
        if not samples.get(_role):
            samples[_role] = list(_lst)

    # 全局正文行距规范应用
    if body_norm:
        for i2, d2 in samples.get("body", []):
            _apply_annotation(d2, body_norm)

    # 有效字体链补全（2026-08-23 CJC 标题 Arial 实锤）：模板示例段 run 级
    # 无显式 ascii（Word 显示=docDefaults TNR/宋体），规则不带 font_ascii →
    # 目标标题保留源稿显式 ascii=Arial（源稿 docDefaults 残留）。机制层：
    # docDefaults 是 Word 有效默认值（同页眉页脚字号链解析）。
    _dd = (getattr(ir, "font_table", None) or {}).get("doc_defaults", {}) or {}
    for _prop in ("font_ascii", "font_east_asia"):
        _val = _dd.get(_prop)
        if _val is None:
            continue
        for _lst in samples.values():
            for _item in _lst:
                _d = _item if isinstance(_item, dict) else _item[1]
                if _d.get(_prop) is None:
                    _d[_prop] = _val

    # 正文"长段"子池：只用于 body 取值（封面字段/目录条目照旧留在 samples 里给别的
    # 角色判据用，不从采集层删，避免连带改变样式合并与封面对齐逻辑）。
    _long_body: list[dict] = []
    for _item in (samples.get("body") or []):
        try:
            _idx, _props = _item
            _txt = (ir.paragraphs[_idx].text or "")
        except Exception:  # noqa: BLE001 — 兼容 dict 形态/越界
            continue
        if isinstance(_props, dict) and _looks_like_long_body(_txt):
            _long_body.append(_props)

    return ({r: [d for _, d in lst] for r, lst in samples.items()}, cover_end,
            body_norm, _long_body)


def _looks_like_long_body(text: str) -> bool:
    """正文样本的"像正文"判据（2026-11-06）。

    背景：body 池长期混进封面字段（题目：/学院：/成绩：，行距 760 固定值）、
    目录条目（摘要…………）、作者/单位行、说明括号行。这些短行的行距/字号一旦
    成为众数，整篇正文会被套上封面的 38 磅行距（浙江外国语实测 spec line=760）。
    判据只做"够长 + 不像目录 + 不像字段行"，宁松勿严：长正文里出现的表格/公式
    说明段照样保留，避免把正文池清空。
    """
    t = re.sub(r"\s+", "", text or "")
    if len(t) < 25:
        return False
    if t.count("…") >= 3 or t.count("．") >= 6:
        return False
    if re.fullmatch(r"[一-龥]{1,8}[：:]", t):
        return False
    return True


#: 正文行距"荒谬值"上限（twips）：>30 磅的行距在中文正文里不可能成立，命中即
#: 认为取到了封面字段/目录条目（2026-11-06 护栏口径，EZI-0017 沿用）。
_BODY_LINE_MAX_TWIPS = 600

#: 参与"正文行距/字号"取值的角色。标题/题注/参考文献/关键词/封面字段各有自己的规则，
#: 它们的行距与字号不是正文的；`list_item` 是正文列表段（交大中期 6 段正文被判成
#: list_item，排除它就等于取不到正文行距）。
_BODY_SAMPLE_ROLES = frozenset({"", "body", "list_item", "abstract_body",
                                "appendix_body", "bio_body"})

#: 正文字号的"荒谬值"护栏（EZI-0026）：5–20pt 之外视为封面大标题/DOI/脚注样
#: 误入正文池（IEEE 模板 body 池里混进 22pt PaperTitle 与 6pt DOI 各 1 段），
#: 不参与众数；全是荒谬值时才退回全集（宁取有据的值，行距护栏同口径）。
_BODY_SIZE_MIN_PT = 5.0
_BODY_SIZE_MAX_PT = 20.0


#: 正文字号"证据成立"的门槛（EZI-0026）：正文段**过半自己写了字号**时，众数才算
#: "模板给正文设的字号"；否则字号是继承来的（样式/Word 默认），拿少数派的直接值
#: 当正文主字号会整篇排小/排大 —— IEEE Transactions 模板 110 段里只有 44 段写 8pt
#: （说明性小字），其余 66 段继承 10pt，正文主字号是 10pt 不是 8pt。
#: 注意：这不是"主峰占比"门槛（EZI-0022 反对的那种），而是"这批段落到底有没有
#: 自己写字号"的覆盖率——与 `_mode_of(..., coverage_min=0.5)` 的既有口径一致；
#: IEEE Access 195/195 段都写了，10pt 才站得住。
_BODY_SIZE_MIN_COVERAGE = 0.5


def _body_sample_paras(ir, cover_end: int = 0) -> list:
    """正文取样的**同一批段落**（行距 EZI-0017 / 字号 EZI-0026 共用一套判据）。

    判据：正文类角色（``_BODY_SAMPLE_ROLES``）+ 无 ``outlineLvl`` + 文本 >10 字
    + 非目录条目（``_is_toc_line``）+ 非版式说明行（``_is_annotation`` / 带行内注记的
    「正文部分…（宋体小4，行间距23磅）」）+ 非参考文献格式示例行（「[编号] 作者. 题目…」）
    + 封面区外。两个取值（行距/字号）用**同一个样本池**，否则"正文怎么排"会出现
    两套互不相容的答案。
    """
    out = []
    for idx, p in enumerate(getattr(ir, "paragraphs", None) or []):
        if cover_end and idx < cover_end:
            continue        # 封面区（「专 业 名 称：XXX」这类表单字段带 480/auto）
        if str(getattr(p, "semantic_role", "") or "") not in _BODY_SAMPLE_ROLES:
            continue
        if getattr(p, "is_heading", False) or getattr(p, "outline_lvl", None) is not None:
            continue
        text = getattr(p, "text", "") or ""
        if len(text.strip()) <= 10:
            continue
        if _is_toc_line(text) or _is_annotation(text):
            continue
        if _extract_inline_annotation(text):
            continue        # 「正文部分…（宋体小4，正文行间距固定为23磅）」示例行
        if re.match(r"^[\[［]\s*(编号|序号|\d+)\s*[\]］]", text.strip()):
            continue        # 参考文献格式示例行：「[编号] 作者. 文章题目…」
        out.append(p)
    return out


def _abstract_body_fallback_paras(ir) -> list:
    """EZI-0072：全程无「摘要」region 的大摘要形态兜底样本。

    病灶（3-2-2 实锤）：大摘要模板整篇即摘要、文档里没有「摘要」标题，region
    全程 None——角色收集链 1037 行 ``elif role == "body" and in_front`` 的
    region 判 ``in ("摘要", ...)`` 不成立，52 段 role=body 全 ``continue``，
    body 池 0 条、abstract_body 池 0 条 → ``rule:tpl.body`` 压根不生成，
    正文行距 360/auto 无人执行（产物保持源 480 红）。

    只服务 body 池为空时的**第二兜底**（第一兜底 = 摘要区 abstract_body 池）：
    正常模板 body 池非空，本函数结果不生效，零风险。闸与 ``_body_sample_paras``
    同一套（无标题/无目录/无注记/≥8 字）。说明行（「字体：黑体；字号：三号…」）
    在 analyzer 角色通道已被划走（figure_para 等，3-2-2 实锤 15 条），role=body
    收不到它们；strip/判定层池的说明段排除由 EZI-0071 谓词同源负责。

    返回**格式字典**（与 ``_collect_samples`` 池元素同构：直接格式 > 样式链 >
    docDefaults 的 ``_eff`` 语义，直接格式优先避免封面样式链污染众数）。
    """
    out = []
    for p in getattr(ir, "paragraphs", None) or []:
        if str(getattr(p, "semantic_role", "") or "") != "body":
            continue
        if getattr(p, "is_heading", False) or getattr(p, "outline_lvl", None) is not None:
            continue
        text = getattr(p, "text", "") or ""
        n = re.sub(r"\s+", "", text)
        if len(n) < 8:
            continue
        if _is_toc_line(text) or _is_annotation(text) or _extract_inline_annotation(text):
            continue
        d = _fp_dict(getattr(p, "direct_format", None))
        out.append(d)
    return out


def _template_body_line(ir, cover_end: int = 0,
                        style_map: dict | None = None) -> dict | None:
    """模板正文行距 = 正文段落**自己**的设值众数（直接格式 > 显式段落样式链）。

    口径（EZI-0017，与判定层 ``validation.layout_fidelity.line_spacing_fact``
    同一套，两侧才能对得上）：

    * 只有**正文类角色**参与（``_BODY_SAMPLE_ROLES``，取样判据见
      ``_body_sample_paras``）；
    * 只认段落**自己**的行距：``pPr/w:spacing`` 直接格式，没有才看段落显式挂的
      ``w:pStyle`` 链（``style_map``）。**文档默认样式（docDefaults / Normal 兜底）
      不进众数** —— 那是全局兜底，不是"模板给正文段落设的行距"。修前这里把
      Normal 的 560/exact 折进每个 body 样本，中介/保管两份合同模板的产物整篇被
      写成直接 560/exact；
    * 短行不参与（≤10 字）：表单标签「学生姓名:」500/exact、公式号、签名栏、以及
      参考文献格式示例行、含行内注记的说明行（见 ``_body_sample_paras``）；
    * 众数并列时**直接格式优先**，再按文档序（``Counter`` 稳定序）；>600 twips
      （30 磅）的荒谬值不参与众数（2026-11-06 护栏口径）。

    返回 ``{"line", "line_rule", "n", "set", "dist"}``；正文段落完全没有取值时返回
    ``None``（调用方再试注记 / 默认段落样式 / 单倍行距兜底）。
    """
    _style_map = style_map or {}
    dist: Counter = Counter()
    direct: Counter = Counter()
    for p in _body_sample_paras(ir, cover_end):
        fp = getattr(p, "direct_format", None)
        line = getattr(fp, "line", None) if fp is not None else None
        rule = getattr(fp, "line_rule", None) if fp is not None else None
        from_direct = line is not None
        if not from_direct and getattr(p, "style_id", None):
            style = _style_map.get(p.style_id) or {}
            line = style.get("line")
            rule = style.get("line_rule")
        if line is None:
            continue
        key = (int(line), str(rule or "auto"))
        dist[key] += 1
        if from_direct:
            direct[key] += 1
    if not dist:
        return None
    # 荒谬值（>600 twips = 30 磅）不参与众数；全是荒谬值时才退回全集（宁取有据的值）
    pool = {k: v for k, v in dist.items() if k[0] <= _BODY_LINE_MAX_TWIPS} or dict(dist)
    (line, rule) = max(pool.items(), key=lambda kv: (kv[1], direct.get(kv[0], 0)))[0]
    n = pool[(line, rule)]
    return {"line": line, "line_rule": rule, "n": n, "set": sum(dist.values()),
            "dist": {f"{k[0]}/{k[1]}": v for k, v in dist.items()}}


def _template_body_size(ir, cover_end: int = 0,
                        style_map: dict | None = None) -> dict | None:
    """模板正文主字号 = 正文段落**自己**的设值众数（直接格式 > 显式段落样式链）。

    与 ``_template_body_line`` 同一套取样（``_body_sample_paras``）与同一原则：
    **文档默认样式（Normal / docDefaults）不进众数**。EZI-0026 的 IEEE 实测就是这个
    病根——模板正文段（PARA 样式）10pt、参考文献 8pt×45、H2/H3 9pt×30，而 `Normal`
    样式声明 ``sz=24``（12pt）；旧实现让 ``style_samples["body"]`` 把 Normal 的 12pt
    补进 `rule:tpl.body`，产物 103 段正文被写成直接 12pt（模板 10pt），另一半源稿的
    10pt/9pt 原样保留 → 产物内部混排 12/10/9。

    口径：

    * 只有**正文类角色**参与，短行/目录/注记/参考文献示例行一概不参与（同取样判据）；
    * **不设占比门槛**：众数就是"模板正文段落最常设的字号"（IEEE 10pt 只占 40.6%，
      被 45%/60% 两道门槛挡掉才轮到默认样式开口）；
    * 但要求**过半正文段自己写字号**（``_BODY_SIZE_MIN_COVERAGE``）：覆盖率不足时
      正文主字号是继承来的，返回 ``size=None`` 让调用方走"默认样式 → Word 默认"
      兜底（IEEE Transactions 模板 44/110 段写 8pt，正文其实是继承的 10pt）；
    * 众数并列时**直接格式优先**，再按文档序；``_BODY_SIZE_MIN_PT``–``_BODY_SIZE_MAX_PT``
      之外的荒谬值（封面 22pt 大标题、6pt DOI）不参与众数；
    * 正文段一个字号都没写时返回 ``None``。

    返回 ``{"size", "n", "set", "pool", "coverage", "dist"}``（``size`` 可能为 None，
    见上）或 ``None``。
    """
    _style_map = style_map or {}
    dist: Counter = Counter()
    direct: Counter = Counter()
    pool_paras = _body_sample_paras(ir, cover_end)
    for p in pool_paras:
        fp = getattr(p, "direct_format", None)
        size = getattr(fp, "font_size_pt", None) if fp is not None else None
        from_direct = size is not None
        if not from_direct and getattr(p, "style_id", None):
            size = (_style_map.get(p.style_id) or {}).get("font_size_pt")
        if size is None:
            continue
        key = round(float(size), 2)
        dist[key] += 1
        if from_direct:
            direct[key] += 1
    if not dist:
        return None
    n_pool = len(pool_paras)
    coverage = sum(dist.values()) / n_pool if n_pool else 0.0
    plausible = {k: v for k, v in dist.items()
                 if _BODY_SIZE_MIN_PT <= k <= _BODY_SIZE_MAX_PT} or dict(dist)
    size = max(plausible.items(), key=lambda kv: (kv[1], direct.get(kv[0], 0)))[0]
    out = {"size": size, "n": plausible[size], "set": sum(dist.values()),
           "pool": n_pool, "coverage": round(coverage, 3),
           "dist": {f"{k}pt": v for k, v in dist.items()}}
    if coverage < _BODY_SIZE_MIN_COVERAGE:
        # 正文段的字号绝大多数是继承来的：少数派的直接值不能当正文主字号
        out["size"] = None
        out["reason"] = (f"正文段只有 {out['set']}/{n_pool} 段自己写了字号"
                         f"（{coverage:.0%} < {_BODY_SIZE_MIN_COVERAGE:.0%}），"
                         f"字号是继承来的")
    return out


def _apply_star_annotation(text: str, d: dict) -> None:
    """段内 *小5号宋体* / （黑体4号）注释并入内容属性，不改 label_*。"""
    txt = text or ""
    ann_txt = _extract_inline_annotation(txt)
    props = parse_annotation_props(ann_txt) if ann_txt else {}
    if not props and "*" in txt:
        props = parse_annotation_props(txt.split("*", 1)[-1])
    if props:
        _apply_annotation(d, props)
        if d.get("font_ascii") and not d.get("label_font_ascii"):
            d["label_font_ascii"] = d["font_ascii"]


def _apply_annotation(d: dict, props: dict) -> None:
    """注释属性并入样本：注释明说的键覆盖实况；行距三键互斥清理。"""
    for k, v in props.items():
        d[k] = v
    if "line" in props or "line_spacing" in props:
        d.pop("line_spacing", None) if "line" in props else d.pop("line", None)
        d.pop("line_rule", None) if "line" in props and "line_rule" not in props else None
        if "line" in props:
            d.pop("line_spacing", None)
        if "line_spacing" in props:
            d.pop("line", None)
            d.pop("line_rule", None)


def _style_samples(ir) -> dict[str, list[dict]]:
    """样式定义兜底样本（样式型模板：格式在 Heading/Normal 样式里）。"""
    out: dict[str, list[dict]] = {}
    resolved = _resolved_style_map(ir)
    for st in ir.styles or []:
        name = (st.name or "")
        role = None
        m = _H_LEVEL_BY_NAME.match(name)
        if m:
            lvl = int(m.group(1) or m.group(2) or 1)
            role = f"heading{min(lvl, 3)}"
        elif name.lower() in ("normal", "正文", "text", "body text"):
            role = "body"
        elif name.lower() in ("title", "title1"):
            role = "article_title"
        elif "caption" in name.lower() or "题注" in name:
            role = "caption"
        if role:
            d = dict(resolved.get(st.style_id) or _fp_dict(st.style_format))
            if d:
                out.setdefault(role, []).append(d)
    return out


def _doc_defaults(tpl: Path) -> dict:
    """styles.xml docDefaults → 默认字体规则期望值（宋体/TNR 基线）。"""
    import zipfile
    try:
        with zipfile.ZipFile(tpl) as z:
            if "word/styles.xml" not in z.namelist():
                return {}
            root = etree.fromstring(z.read("word/styles.xml"))
    except Exception:
        return {}
    ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    rpr = root.find(f"{ns}docDefaults/{ns}rPrDefault/{ns}rPr")
    if rpr is None:
        return {}
    out: dict = {}
    fonts = rpr.find(f"{ns}rFonts")
    if fonts is not None:
        ea = fonts.get(f"{ns}eastAsia")
        asc = fonts.get(f"{ns}ascii")
        if ea:
            out["font_east_asia"] = ea
        if asc:
            out["font_ascii"] = asc
    sz = rpr.find(f"{ns}sz")
    if sz is not None:
        try:
            out["font_size_pt"] = int(sz.get(f"{ns}val", "0")) / 2.0
        except ValueError:
            pass
    return out


def _font_baseline(ir, dd: dict) -> dict:
    """Word 有效字体：Normal/正文 basedOn 链 → docDefaults → 未写 sz 则 10pt。"""
    resolved = _resolved_style_map(ir)
    base: dict = {}
    for st in ir.styles or []:
        if (st.name or "").lower() in ("normal", "正文"):
            base = dict(resolved.get(st.style_id) or {})
            break
    out: dict = {}
    for k in ("font_ascii", "font_east_asia", "font_size_pt"):
        if base.get(k) is not None:
            out[k] = base[k]
        elif dd.get(k) is not None:
            out[k] = dd[k]
    if out.get("font_size_pt") is None:
        out["font_size_pt"] = 10.0
    return out


def _expected_from(samples: list[dict], allowed: tuple) -> dict:
    """主导值 → 期望属性（高矛盾属性用高门槛；位置伪影键永不产出）。"""
    expected: dict = {}
    # 成对悬挂缩进（2026-08-23 CJC v9 实锤）：参考文献条目 ind_left/hanging
    # 在模板按条目类型各异（252/359/308/546/672），众数 0.6 阈值下全部丢弃
    # → 候选无悬挂缩进「内容超过序号」。机制：样本中 left==hanging 成对结构
    # 覆盖率≥0.5 即认为「参考文献=悬挂缩进」结构成立，值取占比最高者
    # （_IND_DOMINANCE=0.4，模板示例值分散，308=50% 已是主导）。
    pair_vals = [(s.get("ind_left"), s.get("ind_hanging")) for s in samples]
    pair_cnt = sum(1 for l, h in pair_vals if l is not None and h is not None)
    hang_cnt = sum(1 for l, h in pair_vals if l is not None and h is not None
                   and str(l) == str(h))
    ind_pair_ok = bool(pair_cnt) and hang_cnt / max(1, len(pair_vals)) >= 0.5
    for prop in allowed:
        if prop in _POSITIONAL_KEYS:
            continue
        vals = [s.get(prop) for s in samples]
        if prop in _IND_PAIR_KEYS and ind_pair_ok:
            dom = _IND_DOMINANCE
        else:
            dom = _DOMINANCE_STRICT if prop in _STRICT_KEYS else _DOMINANCE_DEFAULT
        val, _ratio = _mode_of(vals, dom)
        if val is not None:
            expected[prop] = val
    return expected


# ── 页面几何（节级规则）────────────────────────────────────────

_MARGIN_KEYS = ("top", "bottom", "left", "right", "header", "footer", "gutter")
_MARGIN_TOL = 3  # twips 容差（Word UI 取整：1417/1418 等价）


def unanimous_margins(ir) -> dict:
    """模板各节全一致（容差内）的页边距键 → 期望值；不一致的键丢弃。"""
    sets = []
    for s in ir.sections or []:
        if s.margins:
            sets.append(s.margins)
    if not sets:
        return {}
    out: dict = {}
    for k in _MARGIN_KEYS:
        vals = [m.get(k) for m in sets if m.get(k) is not None]
        if len(vals) != len(sets):
            continue  # 有节缺失该键 → 不视为全一致
        lo, hi = min(vals), max(vals)
        if hi - lo <= _MARGIN_TOL:
            out[k] = (lo + hi) // 2
    return out


def margins_diff(actual: dict, expected: dict) -> dict:
    """超出容差的键（lint/planner 共用语义）。"""
    out = {}
    for k, v in expected.items():
        if k == "section_extra_xml":
            continue
        a = actual.get(k)
        if a is None or abs(int(a) - int(v)) > _MARGIN_TOL:
            out[k] = v
    return out


def _common_structure_bundle(values: list[str | None], kind: str) -> str | None:
    """Return child properties present identically in every sample bundle."""
    if not values or any(not value for value in values):
        return None
    roots = [etree.fromstring(value) for value in values if value]
    maps = [{
        child.tag.rsplit("}", 1)[-1]: etree.tostring(child, encoding="unicode")
        for child in list(root)
    } for root in roots]
    common_names = set(maps[0])
    for item in maps[1:]:
        common_names &= set(item)
    root = etree.Element(
        f"{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}{kind}"
    )
    for child in list(roots[0]):
        name = child.tag.rsplit("}", 1)[-1]
        if name in common_names and all(item[name] == maps[0][name] for item in maps[1:]):
            root.append(copy.deepcopy(child))
    if not list(root):
        return None
    return etree.tostring(root, encoding="unicode", short_empty_elements=True)


# ── 封面锚点 ────────────────────────────────────────────────────

def cover_anchor_rules(ir, cover_end: int, tpl_name: str) -> list[FormatRule]:
    """封面不再按字段克隆格式（专业名称/日期/大标题）。

    封面及封面性质页只保证节级：无可见页眉、无 PAGE、无页眉线——
    取值来自 spec.structure.front，由 repair_section_references 执行。
    """
    return []
    rules: list[FormatRule] = []
    tpl_paras = ir.paragraphs[:cover_end]
    # 模板侧候选：有实际文本、非注释/标签、带有效格式
    cands = []
    for p in tpl_paras:
        t = (p.text or "").strip()
        if not t or _is_annotation(t) or _is_label(t):
            continue
        d = _fp_dict(p.direct_format)
        keys = [k for k in _COVER_KEYS if d.get(k) is not None]
        if not keys:
            continue
        cands.append((t, norm_text(t), d))

    # 模板侧标题行（最长中文行，非字段行/章节序号/日期/长文本块）
    tpl_title = None
    for t, n, d in cands:
        if _FIELD_NAME_OK(t) or _CHAPTER_RE.match(n) or _DATE_RE.match(t):
            continue
        # 标题行特征：中文为主、8-30 字（超过 30 字是段落不是标题）
        compact = re.sub(r"[\s]", "", t)
        if _cjk_ratio(t) >= 0.5 and 8 <= len(compact) <= 30:
            if tpl_title is None or len(compact) > len(re.sub(r"[\s]", "", tpl_title[0])):
                tpl_title = (t, n, d)

    used_norms: set[str] = set()
    rid = 0
    for t, n, d in cands:
        match = None
        if tpl_title is not None and n == tpl_title[1]:
            match = "title"
        elif _DATE_RE.match(t):
            match = "date"
        elif _field_name(t) and len(n) > len(_field_name(t)):
            match = "prefix"       # 带值字段行（专业名称：XXX）
        elif n in ("毕业设计（论文）", "毕业设计(论文)") or _is_anchor(t, None, d.get("jc")):
            match = "exact"
        else:
            continue
        expected = {k: d[k] for k in _COVER_KEYS
                    if d.get(k) is not None and k not in _POSITIONAL_KEYS}
        if not expected:
            continue
        # 声明/授权书长文本段不克隆布局（是正文性质，不是封面标题）
        if match == "title" and len(n) > 30:
            continue
        rid += 1
        rules.append(FormatRule(
            rule_id=f"rule:tpl.cover{rid}",
            selector={"cover_anchor": _field_name(t) or n[:24], "match": match,
                      "norm": n[:40]},
            expected=expected,
            severity="warning",
            auto_fix=True,
            fix_strategy="direct",
            risk_ceiling="R1",
            provenance=SpecProvenance(
                source="template_ooxml",
                detail=f"{tpl_name} 封面锚点「{t[:20]}」",
            ),
            confidence=0.8,
        ))
        used_norms.add(n)
    return rules


def _FIELD_NAME_OK(t: str) -> bool:
    return bool(_field_name(t))


def _cjk_ratio(t: str) -> float:
    if not t:
        return 0.0
    cjk = sum(1 for c in t if "\u4e00" <= c <= "\u9fff")
    return cjk / max(1, len(re.sub(r"\s", "", t)))


# ── 主入口 ──────────────────────────────────────────────────────


_STRUCTURE_EXCLUDE = {
    "毕业设计(论文)", "西安理工大学高科学院", "西安理工大学",
    "学位论文原创性声明", "学位论文版权使用授权书", "摘要", "ABSTRACT",
    "目录", "致谢", "参考文献", "结论", "附录",
    "本人签名:日期:", "本人签名:", "日期:", "论文提交日期:",
}  # 结构段（规范化后）——目标文档真实结构，绝不视为示例残留


def _norm_residual(t: str) -> str:
    """示例残留规范化：去全部空白（含全角）、□占位符、标点归一。"""
    t = re.sub(r"[\s\u3000□]+", "", t or "")
    t = t.replace("：", ":").replace("；", ";").replace("（", "(").replace("）", ")")
    t = t.replace("，", ",").replace("。", ".").replace("、", ",")
    return t.strip()


def extract_example_residuals(ir) -> list[str]:
    """从模板提取示例残留清单（2026-08-16，用户拍板：
    「模板是用来提取格式的，照搬模板内容干什么」）。

    策略：定位「××示例：」标记段 → 其后连续段落累积为示例文本块，
    跳过：注释段（以（开头）、锚点/标题段、TOC 行、空段、表格内段。
    每个连续块 → 一条规范化文本（封面示例区 Altera 标题 3 段 → 一条）。
    """
    residuals: list[str] = []
    in_example = False
    buf: list[str] = []

    def flush():
        nonlocal buf
        if buf:
            norm = _norm_residual("".join(buf))
            if len(norm) >= 4:
                residuals.append(norm)
            buf = []

    for p in ir.paragraphs:
        text = (p.text or "").strip()
        if getattr(p, "cell_id", None) is not None:
            continue
        if re.search(r"示例\s*[:：]", text) and len(text) <= 24:
            flush()
            in_example = True
            continue
        if not in_example:
            continue
        if not text:
            flush()
            continue
        if text.startswith("(") or text.startswith("（"):
            flush()
            continue
        if _is_toc_line(text):
            flush()
            continue
        if p.semantic_role and p.semantic_role.startswith("heading"):
            flush()
            continue
        if _norm_residual(text) in _STRUCTURE_EXCLUDE:
            flush()
            continue
        buf.append(text)
    flush()
    seen = set()
    out = []
    for r in residuals:
        if r not in seen:
            seen.add(r)
            out.append(r)
    return out

def _heading_size(sizes: list) -> float:
    """标题字号冲突时的取值：**保留最大值，除非最大值本身是离群值**。

    层级样例互相冲突时旧口径取 `max`（标题不会比正文小）——交大 12/10.5 打平取 12 靠它。
    但 `max` 会被池里的离群样本带走：浙江政采 `{15.5:23, 12:23, 17.5:2, 15:3, 16:1}` 取到
    17.5pt（4% 占比），整篇条款级标题被套成章级字号；上海黄浦取到 36pt。
    因此只在**最大值占比 <15%**（真离群）时改取「占比 ≥25% 的最大值」，其余一律维持旧口径
    ——实测把规则放粗（无视占比一律取众数）会让施工合同 88.9→77.8、南航 100→85.7、
    同济 80→70，净负（2026-10-31）。
    """
    cnt = Counter(round(float(v), 2) for v in sizes if v)
    if not cnt:
        return 0.0
    total = sum(cnt.values())
    top = max(cnt)
    if cnt[top] / total >= 0.15:
        return top
    solid = [v for v, n in cnt.items() if n / total >= 0.25]
    return max(solid) if solid else top


def decompile_template(
    template_path: str | Path,
    domain: str = "thesis",
    min_cluster: int = 2,
    heading_min_cluster: int = 1,
) -> FormatSpec:
    """从模板反编译 FormatSpec（角色规则 + 节级规则 + 封面锚点规则）。"""
    tpl = Path(template_path)
    if not tpl.exists():
        raise FileNotFoundError(f"template not found: {tpl}")

    ir = analyze_document(tpl)
    _style_map, _doc_def = _style_effective_map(tpl)
    para_samples, cover_end, body_norm, _long_body = _collect_samples(
        ir, _style_map, _doc_def)
    _body_from_abstract = False
    _fallback_body = False
    if not para_samples.get("body") and len(para_samples.get("abstract_body") or []) >= 3:
        para_samples["body"] = list(para_samples["abstract_body"])
        _body_from_abstract = True
    elif len(para_samples.get("body") or []) < min_cluster:
        # EZI-0072 第二兜底：全程无摘要 region（大摘要整篇即摘要、无前后部结构，
        # cover_end 实测吞到 66/67 段）→ 重收 role=body 段
        # （_abstract_body_fallback_paras，闸同源）。此时 cover 切分形同虚设，
        # _body_line/_body_size 的取样窗也须同步放开（cover_end 按 0 用）。
        _fb = _abstract_body_fallback_paras(ir)
        if len(_fb) >= 3:
            para_samples["body"] = _fb
            _body_from_abstract = True
            _fallback_body = True
    # 正文行距（EZI-0017）：只认正文段落自己（直接格式/显式 pStyle）的取值，文档默认
    # 样式与模板注记只在正文段完全没取值时兜底；都没有才用单倍行距（旧口径）。
    # EZI-0072：第二兜底触发时 cover_end 失效（整篇即摘要）→ 取样窗传 0。
    _body_line = _template_body_line(ir, 0 if _fallback_body else cover_end, _style_map)
    # 正文主字号（EZI-0026）：同一个样本池、同一个原则——只认正文段落自己的字号，
    # Normal/docDefaults 不进众数（否则 Normal 的 12pt 会被当成"模板正文 12pt"）。
    _body_size = _template_body_size(ir, 0 if _fallback_body else cover_end, _style_map)

    # 摘要样本池只用于上面的正文兜底：IR 没有 abstract_body 角色，留着会生成一条
    # 永远匹配不到段落的规则（计划里多一条空动作）。
    para_samples.pop("abstract_body", None)
    _seed_named_styles(ir, para_samples)
    style_samples = _style_samples(ir)

    rules: list[FormatRule] = []
    prov_base = f"template_ooxml:{tpl.name}"

    # 0) docDefaults 默认字体规则（styles.xml；无段落样本覆盖时的基线）
    dd = _doc_defaults(tpl)
    baseline = _font_baseline(ir, dd)
    if dd:
        rules.append(FormatRule(
            rule_id="rule:tpl.document_default",
            selector={"scope": "document_default"},
            expected=dd,
            severity="warning",
            auto_fix=True,
            fix_strategy="direct",
            risk_ceiling="R2",
            provenance=SpecProvenance(
                source="template_ooxml",
                detail=f"{prov_base} styles.xml docDefaults",
                artifact_id=str(tpl),
            ),
            confidence=0.9,
        ))

    # 1) 语义角色规则
    # 正文首行缩进提示（2026-10-30）：合同类模板里条款段多被 analyzer 判成
    # list_item，body 池里「无缩进」的段落占多数，取众数会把首行缩进抹掉
    # （建设工程勘察合同：body 样本 229 段里 184 段无缩进、25 段 firstLine=440，
    #  产物全文 left=557 无首行缩进）。这里只在**非零缩进段**里取多数，
    # 覆盖率 ≥60% 才采信，供 body 规则在缺该键时补上。
    _body_ind_hint = None
    try:
        _vals = []
        for _i, _p in enumerate(ir.paragraphs):
            if _i < cover_end:
                continue
            _t = (_p.text or "").strip()
            if len(_t) < 40 or str(_p.semantic_role or "").startswith("heading"):
                continue
            _d = _fp_dict(getattr(_p, "direct_format", None))
            for _layer in ((_style_map.get(getattr(_p, "style_id", None)) or {}), _doc_def):
                for _k, _v in _layer.items():
                    if _d.get(_k) is None and _v is not None:
                        _d[_k] = _v
            _fl, _ch = _d.get("ind_first_line"), _d.get("ind_first_line_chars")
            _vals.append((_fl, _ch))
        # 判据分两层（2026-10-31 实测校准）：
        #   ①「有缩进的正文段」必须占全部候选段 ≥60%——否则就是正文本来不缩进
        #     （物理学报：12 段里只有 2 段带缩进，产物不该压缩进）；
        #   ② 缩进值在**非零段内部**过半即可——合同类文档常 440/330 两套并存
        #     （设备监理 440×31 + 330×14），要求单一值占全量 60% 会漏掉正确值。
        _nz = [v for v in _vals if v[0] or v[1]]
        if len(_nz) >= 5 and len(_nz) / len(_vals) >= 0.6:
            _top, _n = Counter(_nz).most_common(1)[0]
            if _n / len(_nz) >= 0.5 and (_top[0] or _top[1]):
                _body_ind_hint = _top
    except Exception:  # noqa: BLE001 — 提示只做增强
        _body_ind_hint = None

    all_roles = set(para_samples) | {r for r in style_samples if r in _ROLE_KEYS}
    # 同名标题池（`heading1#摘要`）：base 角色 → 已有同名规则的标题键
    title_keys_by_base: dict[str, list[str]] = {}
    for _r in sorted(all_roles):
        if "#" in _r:
            _b, _k = _r.split("#", 1)
            title_keys_by_base.setdefault(_b, []).append(_k)
    for role in sorted(all_roles):
        base_role, _, title_key = role.partition("#")
        allowed = _allowed_keys(base_role)
        if not allowed:
            continue
        ps = para_samples.get(role) or []
        ss = style_samples.get(role) or []
        samples = ps or ss  # 段落样本优先（§5.4：示例段落 > 样式定义）
        min_n = heading_min_cluster if (role.startswith("heading")
                                        or role.startswith("caption")
                                        or role in ("keywords", "keywords_en",
                                                    "article_title",
                                                    "article_title_en", "authors",
                                                    "authors_en",
                                                    "affiliation", "affiliation_en",
                                                    "abstract",
                                                    "abstract_en",
                                                    "classification",
                                                    "caption_figure", "caption_table",
                                                    "refs_item", "drop_cap",
                                                    "refs_heading", "ack_heading",
                                                    "appendix_heading", "bio_heading",
                                                    "appendix_body", "bio_body",
                                                    "equation_display")
                                        ) else min_cluster
        if len(samples) < min_n:
            continue
        expected = _expected_from(samples, allowed)
        # 正文字号（EZI-0026）：不再用样本池众数定正文主字号。样本池（`ps`）里混着
        # 表单标签（「学生姓名:」16pt）、封面行、注记合并值——交大中期的 16pt、
        # 旧的 IEEE 12pt 都是这么来的。正文主字号一律由下面的正文阶梯给：
        # ③ `_template_body_size`（正文段落自己的字号众数）→ ④ 默认样式 → Word 10pt。
        # 正文首行缩进：模板里往往「多数段缩进、少数段不缩进」（条款/列表/说明段），
        # 主导占比一卡就整键丢掉 → 产物全篇没有首行缩进（2026-10-28 全库基线：
        # body_indent 68% 失败，合同类模板首行 440 twips 全部丢失）。
        # 只要 ≥30% 的正文样本写了缩进，就取其中最常见的非空值。
        if base_role == "body" and ps and not expected.get("ind_first_line_chars")                 and not expected.get("ind_first_line"):
            _ind = [d for d in ps if d.get("ind_first_line_chars") or d.get("ind_first_line")]
            if len(_ind) >= max(3, int(0.3 * len(ps))):
                _c = Counter((d.get("ind_first_line_chars") or 0,
                              d.get("ind_first_line") or 0) for d in _ind)
                (_ch, _fl), _ = _c.most_common(1)[0]
                if _ch:
                    expected["ind_first_line_chars"] = _ch
                elif _fl:
                    expected["ind_first_line"] = _fl
        # 标题字号：直接样例互相冲突时（如 12pt 真标题 + 10.5pt 杂样）会被整键丢掉，
        # 随后样式层补一个更大的「标题样式字号」——交大三级标题就是这样变成 14pt 的。
        # 取「众数里的最大值」：既保住交大 12/10.5 打平取 12，也不会被池里的离群值带走
        # （浙江政采 {15.5:23, 12:23, 17.5:2, …} 旧逻辑 max=17.5 只占 2/52，整篇条款级
        # 标题被套成章级字号）。无主峰（最频值 <25%）才退到旧口径最大值。
        if base_role.startswith("heading") and "font_size_pt" not in expected and ps:
            _sizes = [d.get("font_size_pt") for d in ps if d.get("font_size_pt")]
            if _sizes:
                expected["font_size_pt"] = _heading_size(_sizes)
        # `line_spacing` 是 lineRule=auto 时由 line/240 派生的便利字段；同一规则
        # 同时携带 line/line_rule 与 line_spacing 会因字典/planner 顺序互相覆盖，
        # exact/atLeast 尤其会被错误改回 auto。模板已有原始 line 时只保留原值。
        if expected.get("line") is not None:
            expected.pop("line_spacing", None)
        # 公式段专属归一（2026-09-27 用户拍板方案）：
        # ① 缩进**一律归零**（不是"模板没写才归零"）——公式段不是正文，正文/编号行的
        #    首行缩进落到公式上永远是错的（真论文实测 firstLineChars=200 / 480 twips，
        #    公式整段被右推）。悬挂缩进同理。左/右缩进（块缩进）保留模板取值。
        # ② 行距禁固定——lineRule=exact 会裁掉分数/积分/大符号，固定行距一律丢弃
        #    （保留 auto/atLeast）。
        if role == "equation_display":
            for _k in ("ind_first_line_chars", "ind_first_line", "ind_hanging"):
                expected[_k] = 0
            if expected.get("line_rule") == "exact":
                for _k in ("line", "line_rule", "line_spacing"):
                    expected.pop(_k, None)
        # 正文/关键词的加粗不能「少数样本说了算」（2026-10-27 交大本科实测：
        # 封面「学生姓名/学院」这类字段段被归进 body 池，7/23 带 bold，
        # 规则取到 bold=True → 产物整篇正文被刷成粗体）。
        if role in ("body", "keywords"):
            flags = [bool(s.get("bold")) for s in samples]
            if not flags or not all(flags):
                expected.pop("bold", None)
        if (role == "body" and _body_ind_hint
                and not expected.get("ind_first_line")
                and not expected.get("ind_first_line_chars")):
            if _body_ind_hint[0]:
                expected["ind_first_line"] = _body_ind_hint[0]
            if _body_ind_hint[1]:
                expected["ind_first_line_chars"] = _body_ind_hint[1]
        # 正文行距（EZI-0017）：三级来源，逐级**只认有据的**来源，都不成立才兜底。
        #   ① 模板正文段落自己的设值众数（`_template_body_line`：直接格式 > 显式
        #      pStyle 链）——「模板怎么排就怎么排」；
        #   ② 模板注记声明的全局正文行距（「正文行间距固定为23磅」）——只补缺不覆盖：
        #      与 `spec/annotations.merge_annotations` 的正文策略（只补缺）一致；旧实现
        #      让注记覆盖样例，本科论文 33 段 360/auto 被刷成 460/exact；
        #   ③ 正文段落一个值都没有 → 模板默认段落样式（styles.xml 正文样式）→ 仍无
        #      才用单倍行距 240/auto（Word 默认）。旧实现把默认样式折进样本池，等于
        #      把"全局兜底"当成"模板给正文设的行距"，中介/保管合同因此整篇 560/exact。
        if role == "body":
            _body_line_value = None
            if _body_line is not None:
                _body_line_value = (_body_line["line"], _body_line["line_rule"])
            elif body_norm.get("line") is not None:
                _body_line_value = (body_norm["line"], body_norm.get("line_rule") or "auto")
            else:
                _style_line = next((s for s in (style_samples.get("body") or [])
                                    if s.get("line") is not None), None)
                if _style_line is not None:
                    _body_line_value = (_style_line["line"],
                                        _style_line.get("line_rule") or "auto")
                else:
                    _body_line_value = (240, "auto")
            expected["line"], expected["line_rule"] = _body_line_value
            expected.pop("line_spacing", None)
        # 正文主字号（EZI-0026）：与行距同一原则、同一取样——样本池众数一律作废
        # （含表单标签/离群值/默认样式继承），只认下面这条阶梯：
        #   ① 模板正文段落**自己的**字号众数（`_template_body_size`：直接格式 >
        #      显式 pStyle 链，Normal/docDefaults 不进众数、不设占比门槛）——
        #      IEEE 模板正文本体 10pt 占 40.6%，旧口径两道门槛都过不去，随后被
        #      `style_samples["body"]` 的 Normal 12pt 顶替，产物 103 段正文写成 12pt；
        #   ② 正文段过半数没写字号（覆盖率 < 50%）→ 字号是继承来的，不采少数派
        #      直接值（IEEE Transactions 模板 44/110 段写 8pt，正文实为继承的 10pt）；
        #   ③ 仍然没有 → 下方默认段落样式兜底 → 最后才是 Word 默认 10pt。
        # 注意 Normal 不再从 `style_samples["body"]` 后门补字号（见该处跳过列表）。
        if role == "body":
            expected.pop("font_size_pt", None)
            if _body_size and _body_size.get("size") is not None:
                expected["font_size_pt"] = float(_body_size["size"])
        # 正文行距"荒谬值"护栏（2026-11-06 口径）已上移到 `_template_body_line`：
        # 那里按 ≤600 twips 过滤后再取众数（浙江外国语 760 twips 那种封面值进不了
        # 正文规则），不再拿 body 池（含样式链继承值）兜底。
        if role == "body" and "jc" not in expected:
            jcs = [s.get("jc") for s in samples if s.get("jc")]
            if jcs:
                expected["jc"] = max(set(jcs), key=jcs.count)
        if role == "body":
            for s in (style_samples.get("body") or []):
                for k, v in s.items():
                    if v is None or k in expected or k not in allowed:
                        continue
                    # 行距/字号不由样式定义补：样式级 line/sz 是"继承"，不是模板给正文
                    # 段落设的值（EZI-0017 行距 / EZI-0026 字号）。补缺就会把 Normal 的
                    # 560/exact、sz=24(12pt) 从后门拉回正文规则——IEEE 的 12pt 正是这里。
                    if k in ("line", "line_rule", "line_spacing", "font_size_pt"):
                        continue
                    expected[k] = v
            if "font_ascii" not in expected and dd.get("font_ascii"):
                expected["font_ascii"] = dd["font_ascii"]
            if "font_east_asia" not in expected and dd.get("font_east_asia"):
                expected["font_east_asia"] = dd["font_east_asia"]
            if "font_size_pt" not in expected:
                # ③ 默认段落样式兜底（EZI-0026：从上面"样式定义补缺"里挪到这儿）。
                #    Word 有效字号 = Normal 样式（再退 docDefaults，最后 10pt），这是
                #    **继承**来的，不是"模板给正文设的字号"，所以只能当最后兜底、不
                #    参与实测众数——IEEE 的 12pt 就是它越位当了正文主字号。
                _def_sz = next((s.get("font_size_pt") for s in (style_samples.get("body") or [])
                                if s.get("font_size_pt") is not None), None)
                if _def_sz is not None:
                    expected["font_size_pt"] = _def_sz
                elif dd.get("font_size_pt") is not None:
                    expected["font_size_pt"] = dd["font_size_pt"]
                else:
                    # Normal/docDefaults 未写 sz 时 Word 解析为 10pt（半磅 20）
                    expected["font_size_pt"] = 10.0
        # 所有角色：样本没写的字体/字号按 Word 有效值一次补完
        # （IEEE Heading 2 无 sz → 10pt；学报 heading1 已有 14pt 不覆盖）
        for s in ss:
            for k in ("font_ascii", "font_east_asia", "font_size_pt"):
                if k in allowed and k not in expected and s.get(k) is not None:
                    expected[k] = s[k]
        for k, v in baseline.items():
            if k in allowed and k not in expected and v is not None:
                expected[k] = v
        # 末尾再兜一次：样式层（Normal/正文样式）合并后可能重新带回 bold。
        if role in ("body", "keywords") and expected.get("bold"):
            if not all(bool(x.get("bold")) for x in samples):
                expected.pop("bold", None)
        if not expected:
            continue
        # 字体名归一：模板里的 Noto Serif SC / Liberation Serif 是 LibreOffice
        # 替代字体，Windows Word 没有 → 产物会挂着一堆「本机不存在的字体」。
        try:
            from earlyz.tools.fonts import normalize_font_fields
            normalize_font_fields(expected)
        except Exception:  # noqa: BLE001 — 字体归一是增强，不阻断提取
            pass
        # 标题层级由角色推导，禁止照抄模板样例的 outlineLvl（2026-10-27 交大实测：
        # 模板把 1.1 样例段标成 outline=2，规则照抄后写回产物，把
        # ensure_heading_outline 刚改对的层级又覆盖回三级，目录/导航全错）。
        if str(role) in ("heading1", "heading2", "heading3", "heading4"):
            try:
                expected["outline_lvl"] = int(str(role)[-1]) - 1
            except ValueError:
                pass
        if str(role).startswith("heading") and "line" not in expected:
            expected["line"] = 240
            expected["line_rule"] = "auto"
        if role == "classification" and "line" not in expected:
            expected["line"] = 240
            expected["line_rule"] = "auto"
        # Word 2013+ Normal 段后 8 磅。模板没写 after 就必须写 0，否则源稿/默认 8 磅还在。
        if role in ("body", "classification", "drop_cap") and "spacing_after" not in expected:
            expected["spacing_after"] = 0
        if role == "drop_cap" and "spacing_before" not in expected:
            expected["spacing_before"] = 0
        # 单样本噪声防护：丢掉位置伪影，保留字体/对齐/段距/行距/keep
        # （期刊模板每级标题常常只有 1 个示例段，旧过滤把段前段后全删了）
        if (role.startswith("heading") or role in ("keywords", "keywords_en",
                                                    "article_title",
                                                    "article_title_en", "authors",
                                                    "authors_en",
                                                    "affiliation", "affiliation_en",
                                                    "abstract",
                                                    "abstract_en",
                                                    "classification",
                                                    "caption_figure", "caption_table",
                                                    "refs_item", "drop_cap",
                                                    "refs_heading", "ack_heading",
                                                    "appendix_heading", "bio_heading",
                                                    "appendix_body", "bio_body",
                                                    "equation_display")
            ) and len(samples) == 1:
            single_keep = set(_PORTABLE_TYPED_KEYS) | {
                "label_font_east_asia", "label_font_ascii", "label_font_size_pt",
                "label_bold", "label_italic", "label_text", "drop_cap",
                "drop_cap_lines", "ppr_extra_xml", "rpr_extra_xml",
                "label_after_tab",
            }
            expected = {k: v for k, v in expected.items() if k in single_keep}
        # Final single source of truth after style/baseline/single-sample enrichment.
        if expected.get("line") is not None:
            expected.pop("line_spacing", None)
        if _bundle_has_child(expected.get("ppr_extra_xml"), "tabs"):
            expected.pop("tab_pos", None)
        # 公式编号形态（2026-09-28 用户拍板「模板里放公式样本」）：模板公式样本带编号时
        # 记下形态，供 generate_equation_numbers **优先采用**（模板没写就不发明）。
        if role == "equation_display":
            _nf, _np = _equation_number_style(ir)
            if _nf:
                expected["number_format"] = _nf
                expected["number_paren"] = _np
        src = "paragraph_cluster" if ps else "style_definition"
        if title_key:
            # 同名标题规则：只作用于目标里同名的标题（analyzer.heading_rule_applies）
            _rid = f"rule:tpl.{base_role}.title.{title_key}"
            _sel = {"semantic_role": base_role, "title_key": title_key}
        else:
            _rid = f"rule:tpl.{role}"
            _sel = {"semantic_role": role}
            if title_keys_by_base.get(role):
                _sel["skip_title_keys"] = sorted(title_keys_by_base[role])
        rules.append(FormatRule(
            rule_id=_rid,
            selector=_sel,
            expected=expected,
            severity="error" if role.startswith("heading") or role == "body" else "warning",
            auto_fix=True,
            fix_strategy="direct",
            risk_ceiling="R2",
            provenance=SpecProvenance(
                source="template_ooxml",
                detail=(f"{prov_base} {src} n={len(samples)}"
                        + (" from=abstract_body" if (role == "body" and _body_from_abstract) else "")),
                artifact_id=str(tpl),
            ),
            confidence=min(0.95, 0.55 + 0.1 * min(len(samples), 4)),
        ))

    # 1b) appendix_body / bio_body 无样本时继承正文（2026-10-09 用户定案）：
    #     模板的附录/作者简介区往往只有版式注记与空白占位段（已在上游剔除），
    #     取不到真实内容样本——注记原文即声明「正文部分…」：内容按正文规则
    #     排版（用户原话「附录A是章标题，其余内容是属于正文」）。模板自带真实
    #     内容样本时各自规则保留，不受影响。
    #
    #     2026-10-17 修正（CJC 实测）：模板若**明确写了附录正文的字体字号**
    #     （「附录内容…字体为小5号宋体」），以注记为准确覆盖继承值——模板定义
    #     怎么排优先于「附录按正文」的旧口径。
    _appendix_ann: dict = {}
    for _p in ir.paragraphs:
        _t = (_p.text or "").strip()
        if "附录" in _t and ("字体" in _t or "字号" in _t or "加粗" in _t):
            _props = parse_annotation_props(_t)
            if _props.get("font_size_pt") or _props.get("font_east_asia"):
                _appendix_ann = _props
                break
    _roles_built = {r.selector.get("semantic_role") for r in rules
                    if isinstance(r.selector, dict)}
    _body_rule = next((r for r in rules if isinstance(r.selector, dict)
                       and r.selector.get("semantic_role") == "body"), None)
    if _body_rule is not None:
        for _derive in ("appendix_body", "bio_body"):
            if _derive in _roles_built:
                continue
            _r = copy.deepcopy(_body_rule)
            _r.rule_id = f"rule:tpl.{_derive}"
            _r.selector = {"semantic_role": _derive}
            _r.provenance = SpecProvenance(
                source="template_ooxml",
                detail=(f"{prov_base} 继承正文规则（{_derive} 区无独立内容"
                        f"样本；注记声明「正文部分…」）"),
                artifact_id=str(tpl),
            )
            rules.append(_r)

    # 1c) 附录正文注记覆盖（2026-10-17）：上面继承或已有样本的 appendix_body，
    #     只要模板注记写了字体/字号，就以注记为准确值。
    #     注记段本身不是内容样本（它是说明文字，常带 jc=center 等演示格式），
    #     所以版式（jc/缩进/行距/段距）沿用正文规则，只采纳注记声明的字体/字号/加粗。
    if _appendix_ann:
        _layout_base = dict(_body_rule.expected) if _body_rule is not None else {}
        for _r in rules:
            if isinstance(_r.selector, dict) and _r.selector.get("semantic_role") == "appendix_body":
                _merged = dict(_layout_base)
                _merged.update({k: v for k, v in _appendix_ann.items() if v is not None})
                _r.expected = _merged
                _r.provenance = SpecProvenance(
                    source="written_standard",
                    detail=(f"{prov_base} 附录正文注记覆盖："
                            f"{', '.join(f'{k}={v}' for k, v in _appendix_ann.items())}"
                            f"（版式沿用正文）"),
                    artifact_id=str(tpl),
                )
                _r.confidence = max(0.9, float(getattr(_r, "confidence", 0.8) or 0.8))

    # 2) 节级公共格式。页边距走 typed 字段；pgSz/docGrid/边框等安全
    # 属性走 portable bundle。type/cols/pgNumType/titlePg/关系引用不在 bundle
    # 中，它们由语义节结构系统处理。
    section_expected = unanimous_margins(ir)
    section_bundle = _common_structure_bundle(
        [s.section_extra_xml for s in (ir.sections or [])], "sectPr"
    )
    if section_bundle:
        section_expected["section_extra_xml"] = section_bundle
    if section_expected:
        rules.append(FormatRule(
            rule_id="rule:tpl.page_setup",
            selector={"scope": "section_all"},
            expected=section_expected,
            severity="warning",
            auto_fix=True,
            fix_strategy="direct",
            risk_ceiling="R1",
            provenance=SpecProvenance(
                source="template_ooxml",
                detail=f"{prov_base} sections={len(ir.sections)} 全一致页边距",
                artifact_id=str(tpl),
            ),
            confidence=0.85,
        ))

    # 2a) 页眉格式规则（2026-08-15：模板页眉=宋体+下划线，目标摘要页眉
    # 无格式 → 用户「页眉没改」。取模板所有 header 非空段的公共键：
    # 字体/字号/下划线；jc 各页眉不同不取）
    def _part_common_rule(part_map: dict, scope: str, rule_id: str) -> None:
        fps = [para["fp"] for part in (part_map or {}).values() for para in part]
        if not fps:
            return
        common: dict = {}
        keys = (
            "font_east_asia", "font_ascii", "font_size_pt", "bold", "italic",
            "underline", "color", "small_caps", "p_bdr_bottom",
            "ppr_extra_xml", "rpr_extra_xml",
        )
        for k in keys:
            vals = [getattr(f, k, None) for f in fps]
            if vals and vals[0] is not None and all(v == vals[0] for v in vals):
                common[k] = vals[0]
        if common:
            rules.append(FormatRule(
                rule_id=rule_id,
                selector={"scope": scope},
                expected=common,
                severity="warning",
                auto_fix=True,
                fix_strategy="direct",
                risk_ceiling="R2",
                provenance=SpecProvenance(
                    source="template_ooxml",
                    detail=f"{prov_base} {scope} paragraphs={len(fps)} 公共格式",
                    artifact_id=str(tpl),
                ),
                confidence=0.85,
            ))


    _HF_TYPED_KEYS = (
        "font_east_asia", "font_ascii", "font_size_pt", "bold", "italic",
        "underline", "color", "small_caps", "jc", "line", "line_rule",
        "spacing_before", "spacing_after", "tab_pos", "p_bdr_bottom",
    )


    def _hf_para_dict(fp) -> dict:
        """header/footer 段直接格式 → dict（typed + extras，None 剔除）。

        只取模板段自身直接属性（含其 pPr/rPr extras 携带的 pBdr/tabs 等）；
        样式链基线不展开——目标文档保留自己的 header/footer 样式（a5/a7），
        跨文档样式 ID（af0/af）不做移植。
        """
        out: dict = {}
        for k in _HF_TYPED_KEYS:
            v = getattr(fp, k, None)
            if v is not None:
                out[k] = v
        pe = getattr(fp, "ppr_extra_xml", None)
        if pe:
            out["ppr_extra_xml"] = pe
        re_ = getattr(fp, "rpr_extra_xml", None)
        if re_:
            out["rpr_extra_xml"] = re_
        return out


    def _text_core(s: str) -> str:
        """页眉/页脚文本核：去空白，供内容相似配对（机制层，不做学校常量）。"""
        return re.sub(r"[\s\u3000]+", "", s or "")


    def _resolve_effective_font_size(ir, style_id) -> float | None:
        """header/footer 段有效字号：段直接 fp → pStyle basedOn 链 → docDefaults。
        （2026-08-23 CJC 字号实锤：模板页眉 sz=18 由 Normal 提供、run 级无 sz；
        若不展开链，应用层 _apply_hf_format 不写 sz，目标段落入源稿 styles.xml
        的 docDefaults（该案例源稿 docDefaults sz=29 → Word 显示 14.5）。）
        无 pStyle 的段（如模板 header3）按 Word 语义落到默认段落样式
        （w:style default=1 paragraph）再沿链解析。
        链解析=机制层，无学校常量：换模板后跟随该模板 styles.xml。"""
        styles = {s.style_id: s for s in (getattr(ir, "styles", None) or [])}

        def _walk(sid: str | None) -> float | None:
            cur = sid
            seen: set[str] = set()
            steps = 0
            while cur and cur not in seen and steps < 16:
                seen.add(cur)
                st = styles.get(cur)
                if st is None:
                    break
                sz = getattr(getattr(st, "style_format", None), "font_size_pt", None)
                if sz is not None:
                    return sz
                cur = getattr(st, "based_on", None)
                steps += 1
            return None

        if style_id:
            sz = _walk(style_id)
            if sz is not None:
                return sz
        else:
            # 无 pStyle → 默认段落样式（机制层：w:style default=1）
            for st in (getattr(ir, "styles", None) or []):
                if getattr(st, "style_type", "paragraph") == "paragraph" \
                        and getattr(st, "is_default", False):
                    sz = _walk(st.style_id)
                    if sz is not None:
                        return sz
        dd = (getattr(ir, "font_table", None) or {}).get("doc_defaults", {}) or {}
        return dd.get("font_size_pt")


    def _header_footer_part_formats(ir) -> list[dict]:
        """模板 header/footer part → 逐 part 段格式（取值来自模板 OOXML）。

        输出：[{kind, ref_type, part, text_core, paragraphs:[{text_core, 格式}]}]
        ref_type 取模板 sectPr 的实际绑定（even/default/first）；某 part 被
        多个节引用时只出一次。
        """
        out: dict[str, dict] = {}
        for sect in (getattr(ir, "sections", None) or []):
            for hf in (getattr(sect, "header_footers", None) or []):
                kind = hf.get("kind")
                part = hf.get("part")
                if not part or (kind not in ("header", "footer")):
                    continue
                if part in out:
                    continue
                cell = {
                    "kind": kind,
                    "ref_type": hf.get("ref_type") or "default",
                    "part": part,
                    "text_core": "",
                    "paragraphs": [],
                }
                part_map = (getattr(ir, "headers", None) or {}) if kind == "header" \
                    else (getattr(ir, "footers", None) or {})
                for pa in part_map.get(part, []) or []:
                    txt = pa.get("text") or ""
                    if not _text_core(txt):
                        continue
                    _para = {
                        "text_core": _text_core(txt),
                        # run 级机制签名（2026-08-23 CJC 实锤）：TEXT/{PAGE}/TAB
                        "run_sig": pa.get("run_sig") or [],
                        "tab_segments": pa.get("tab_segments") or [],
                        # EZI-0067：模板段 pStyle 引用随行——页码段字号/样式
                        # 藏在样式链（pStyle=Footer sz18），池条目带上 style_id
                        # 供应用层移植引用或解析链。
                        "tpl_pstyle": pa.get("style_id"),
                        **_hf_para_dict(pa.get("fp")),
                    }
                    # 有效字号链补全（2026-08-23 CJC 字号实锤）：模板段 run 级
                    # 无 sz（页眉=Normal 链 sz18）时解析 pStyle basedOn 链，
                    # 否则应用层得不到 font_size_pt → 目标落 docDefaults(14.5)。
                    if _para.get("font_size_pt") is None:
                        _sz = _resolve_effective_font_size(ir, pa.get("style_id"))
                        if _sz is not None:
                            _para["font_size_pt"] = _sz
                    cell["paragraphs"].append(_para)
                # EZI-0067：页码段旁路（ir.footers_page）→ cell["page_para"]。
                # 模板纯域 PAGE 段此前被空文字过滤整段丢弃，页码字号/样式
                # （pStyle=Footer sz18=9pt）从未进池 → 产物页码落 docDefaults。
                # 纯页码 part（无文字段）也发cell——页码格式需要载体；
                # 首个页码段即代表（同 part 多 PAGE 段取一致公共口径）。
                if kind == "footer":
                    pmap_page = getattr(ir, "footers_page", None) or {}
                    for pp in pmap_page.get(part, []) or []:
                        page_pa = {**_hf_para_dict(pp.get("fp")),
                                   "tpl_pstyle": pp.get("style_id"),
                                   "page_only": True}
                        if page_pa.get("font_size_pt") is None:
                            _psz = _resolve_effective_font_size(ir, pp.get("style_id"))
                            if _psz is not None:
                                page_pa["font_size_pt"] = _psz
                        cell["page_para"] = page_pa
                        break
                if cell.get("page_para") or cell["paragraphs"]:
                    if cell["paragraphs"]:
                        cell["text_core"] = "".join(p["text_core"] for p in cell["paragraphs"])
                    out[part] = cell
        return list(out.values())

    _part_common_rule(ir.headers or {}, "header_all", "rule:tpl.header_format")
    _part_common_rule(ir.footers or {}, "footer_all", "rule:tpl.footer_format")

    # 2a') 逐 part 页眉/页脚格式（2026-08-22 用户裁决「双栏模板页眉页脚
    # 支持为零」）：模板各 header/footer part 内容/格式本就不同
    # （even/default/first 三态 + 首页刊头双线），「跨 part 公共键」必然
    # 失败。改为按 part 提取：段文本核 + 段直接格式（pps/rpr extras +
    # typed 键），由应用层按内容相似配对落位（模板定义怎么排，目标提供
    # 排什么）。类型化值只取直接格式——样式链基线（af0/a5 等）保留目标
    # 自身样式，不做跨文档样式 ID 移植。
    header_footer_parts = _header_footer_part_formats(ir)

    # 2b) 表格对齐规则（模板各表 alignment 全一致才出——2026-08-15：
    # 模板 tblPr jc=center，目标文档表格不居中 → 用户「表也没有居中」）
    table_expected: dict = {}
    tbl_aligns = {t.alignment for t in (ir.tables or []) if t.alignment}
    if len(tbl_aligns) == 1:
        table_expected["jc"] = next(iter(tbl_aligns))
    table_bundle = _common_structure_bundle(
        [t.tblpr_extra_xml for t in (ir.tables or [])], "tblPr"
    )
    if table_bundle:
        table_expected["tblpr_extra_xml"] = table_bundle
    if table_expected and (ir.tables or []):
        rules.append(FormatRule(
            rule_id="rule:tpl.table_alignment",
            selector={"scope": "table_all"},
            expected=table_expected,
            severity="warning",
            auto_fix=True,
            fix_strategy="direct",
            risk_ceiling="R2",
            provenance=SpecProvenance(
                source="template_ooxml",
                detail=f"{prov_base} tables={len(ir.tables)} 公共 tblPr 格式",
                artifact_id=str(tpl),
            ),
            confidence=0.85,
        ))

    # 2c) 已有编号绑定的语义级别格式。模板 numId/abstractNumId 不进入
    # 规则；只在同一角色的 ilvl/numFmt/lvlText/start 完全一致时发布，目标
    # writer 再解析目标自己的 numId → abstractNumId。
    definitions = {d.num_id: d for d in (ir.numbering or [])}
    numbering_samples: dict[str, list] = {}
    for para in ir.paragraphs or []:
        binding = para.list_binding or {}
        num_id = binding.get("num_id")
        ilvl = binding.get("ilvl")
        if num_id in (None, 0) or ilvl is None:
            continue
        definition = definitions.get(num_id)
        level = next(
            (item for item in (definition.levels if definition else []) if item.level == ilvl),
            None,
        )
        if level is None:
            continue
        numbering_samples.setdefault(para.semantic_role or "list_item", []).append(level)
    for role, samples in numbering_samples.items():
        levels = {sample.level for sample in samples}
        fmts = {sample.num_fmt for sample in samples}
        texts = {sample.lvl_text for sample in samples}
        starts = {sample.start for sample in samples}
        if len(levels) != 1 or len(fmts) != 1 or len(texts) != 1 or len(starts) != 1:
            continue
        expected = {
            "num_fmt": next(iter(fmts)),
            "lvl_text": next(iter(texts)),
            "start": next(iter(starts)),
        }
        expected = {k: v for k, v in expected.items() if v is not None}
        level_bundle = _common_structure_bundle(
            [sample.level_extra_xml for sample in samples], "lvl"
        )
        if level_bundle:
            expected["level_extra_xml"] = level_bundle
        if not expected:
            continue
        level_no = next(iter(levels))
        rules.append(FormatRule(
            rule_id=f"rule:tpl.numbering.{role}.l{level_no}",
            selector={"scope": "numbering_role", "numbering_role": role, "level": level_no},
            expected=expected,
            severity="warning",
            auto_fix=True,
            fix_strategy="direct",
            risk_ceiling="R2",
            provenance=SpecProvenance(
                source="template_ooxml",
                detail=f"{prov_base} numbering role={role} ilvl={level_no} n={len(samples)}",
                artifact_id=str(tpl),
            ),
            confidence=min(0.95, 0.7 + 0.05 * min(len(samples), 5)),
        ))

    # 3) 封面锚点规则
    rules.extend(cover_anchor_rules(ir, cover_end, tpl.name))

    spec = FormatSpec(
        spec_id=f"spec:{domain}-tpl-{tpl.stem[:28]}",
        name=f"{tpl.stem}（语义反编译）",
        domain=domain,
        source_type="uploaded_template_docx",
        source_artifact_id=str(tpl),
        rules=rules,
    )
    spec.example_residuals = extract_example_residuals(ir)
    from earlyz.spec.structure import extract_template_structure
    spec.structure = extract_template_structure(tpl)
    spec.structure["header_footer_parts"] = header_footer_parts
    # 正文行距取值出处（EZI-0017 诊断）：判定层/语料报告可直接引用，不必再猜
    # 「560 从哪来」。`paragraphs` = 正文段落自身设值众数（直接格式 > 显式 pStyle）；
    # `annotation` = 模板注记声明（仅前者缺失时补缺）；两者皆 None = 走默认样式/单倍兜底。
    spec.structure["body_line_source"] = {
        "paragraphs": _body_line,
        "annotation": (dict(body_norm) if body_norm else None),
    }
    # 正文主字号取值出处（EZI-0026 诊断）：`paragraphs` = 正文段落自身字号众数
    # （直接格式 > 显式 pStyle，Normal/docDefaults 不进），`applied` = 规则里实际
    # 落下的值。判定层/语料报告可直接复核"这个字号是从模板哪来的"。
    spec.structure["body_size_source"] = {
        "paragraphs": _body_size,
        "applied": next((r.expected.get("font_size_pt") for r in rules
                         if r.rule_id == "rule:tpl.body"), None),
    }
    # 注释抽出的 toc 规则并入 rules，供 planner 与结构段共用
    toc = spec.structure.get("toc") or {}
    for lvl, expected in toc.items():
        if not expected:
            continue
        rules.append(FormatRule(
            rule_id=f"rule:tpl.toc{lvl}",
            selector={"semantic_role": f"toc{lvl}"},
            expected=expected,
            severity="warning",
            auto_fix=True,
            fix_strategy="direct",
            risk_ceiling="R1",
            provenance=SpecProvenance(
                source="written_standard",
                detail=f"{prov_base} toc annotation l{lvl}",
                artifact_id=str(tpl),
            ),
            confidence=0.9,
        ))
    spec.rules = rules
    spec.delivery_policy.front_matter_roman = bool(
        spec.structure.get("front_matter_roman")
    )
    return spec



# 句子形态（含句号/分号等）不是章节标题——编号列表项「3.请示内容要真实。需…」
# 会被 _CHAPTER_RE 误当成章标题，导致封面边界落在文末、正文规则整篇不生效
# （2026-10-28 燕山大学请示实测：边界 21/21，产物无行距）。
_SENTENCE_MARKS = "。；！？.;!?"


def is_sentence_like(text: str) -> bool:
    t = (text or "").strip()
    return bool(t) and (t[-1] in _SENTENCE_MARKS or "。" in t)

_NUM_LIST_MARK_RE = re.compile(r"^\s*\d+\s*[.、)）]")


def is_chapter_anchor(text: str) -> bool:
    """章节序号行判据（封面边界专用）。

    阿拉伯数字列表项不是章标题：「1.一次性付款，…」「2.分期付款，…：」这类
    段落成排出现在合同/公文的条款里，一旦被当锚点，封面区会一路推进到文中，
    其后正文段落全被当封面跳过（2026-10-29 中介合同实测：边界 42，正文前
    14 段没有缩进）。判据：数字编号项且正文含句读/冒号 → 不是章标题。
    """
    t = (text or "").strip()
    n = norm_text(t)
    if not n or len(n) > 24 or not _CHAPTER_RE.match(n):
        return False
    if is_sentence_like(t):
        return False
    # 注意：norm_text 会吃掉 ASCII 句点（"2.分期" → "2分期"），数字编号
    # 必须在原文上判，否则守卫永远不触发。
    if _NUM_LIST_MARK_RE.match(t.lstrip()) and re.search(r"[，,：:；;]", t):
        return False
    return True



# 前置区角色：学位论文/期刊的摘要、目录、题名、作者等。它们可以很长，
# 但都不属于「正文」——不能用来判定封面到此结束。
_FRONT_COVER_ROLES = {
    "abstract", "abstract_en", "toc", "keywords", "keywords_en",
    "article_title", "article_title_en", "authors", "authors_en",
    "affiliation", "affiliation_en", "classification", "title_en",
    "declaration", "ack_heading",
}


def _cover_like_text(t: str, role: str | None = None) -> bool:
    """该段是否可能属于封面：空段 / 短行（≤24 字脱敏后）/ 日期行 / 前置区角色。

    封面永远是短行堆叠；一旦出现**正文角色**的长段，封面区就结束了。
    前置区（摘要/目录/声明）段落再长也不算正文——否则学位论文的封面边界
    会被摘要截断，声明页、目录标题被当成正文重刷（2026-10-30 东北师大实测）。
    """
    raw = (t or "").strip()
    if not raw:
        return True
    if role and str(role) in _FRONT_COVER_ROLES:
        return True
    if _DATE_RE.match(raw):
        return True
    return len(norm_text(raw)) <= 24


def cover_like_prefix(paragraphs) -> int:
    """从 0 开始连续「封面形态」段落的个数（遇到第一个正文长段即停）。"""
    n = 0
    for p in paragraphs:
        if not _cover_like_text(getattr(p, "text", "") or "",
                                getattr(p, "semantic_role", None)):
            break
        n += 1
    return n


def find_cover_end(ir) -> int:
    """目标文档封面边界（首个锚点索引；lint/writer 判定封面区共用）。

    封面区定义：首个「结构性锚点」之前的段落——锚点=正文章起始章标题
    （第X章/1 绪论 等章节序号行）或已知固定锚点词（郑重声明/目录等）。
    找不到任何锚点 = 文档无封面区（返回 0），role 规则全文生效。
    """
    # 封面区不能超过「封面形态前缀」：合同/说明书类文档第一个章锚点可能在
    # 第 100+ 段（合同节水管理：首个锚点 3.1 在第 142 段），若直接取锚点索引，
    # 前 142 段会被当成封面整段跳过，正文缩进/字号全不应用。
    _cap = cover_like_prefix(ir.paragraphs)
    for i, p in enumerate(ir.paragraphs):
        if i > _cap:
            return _cap
        # 章节序号行（正文首个章标题）→ 封面区终点
        n = norm_text(p.text or "")
        if is_chapter_anchor(p.text):
            return min(i, _cap)
        # 锚点必须有文字：空段（即使带 heading 样式/outlineLvl）不承载
        # 「封面结束」语义——与 new_actions._compute_cover_boundary 同条件
        # （2026-09-25 真机：首段空段残留 heading 1 样式 → 边界=0 → 封面
        # 被 role 批量动作按正文取值重刷）。
        if (n and p.semantic_role and p.semantic_role.startswith("heading")
                and len(n) <= 24 and _chapterish_text(n) and not is_sentence_like(p.text)):
            # 弱判据收紧（2026-09-25 审计）：role=heading 但无「章节/词表」文本
            # 证据的段（封面内样式残留：校名/题名被作者套了标题样式）不作锚点，
            # 否则边界落进封面内部 → 封面被 role 批量动作按正文取值重刷。
            return min(i, _cap)
        d = _fp_dict(p.direct_format)
        if _is_anchor(p.text, p.semantic_role, d.get("jc")):
            return min(i, _cap)
    return 0


def match_cover_paragraphs(ir, rule_selector: dict) -> list:
    """按封面锚点规则匹配目标文档段落（lint 与 writer 共用语义）。

    title/date 类匹配只认封面区内的段落（结构边界之外的正文长行
    即使形似标题也不命中——防止把正文段刷成封面黑体26pt）。
    """
    match = rule_selector.get("match")
    anchor = (rule_selector.get("cover_anchor") or "").strip()
    norm = (rule_selector.get("norm") or "").strip()
    cover_end = find_cover_end(ir)
    limit = min(len(ir.paragraphs), cover_end + 8)
    out = []
    for i, p in enumerate(ir.paragraphs[:limit]):
        t = (p.text or "").strip()
        if not t:
            continue
        n = norm_text(t)
        in_cover = i < cover_end
        if match == "title":
            # 排除已有专属 exact 规则的锚点文本（毕业设计（论文）/摘要/目录…）：
            # 否则 title 规则（黑体22pt）会覆盖 exact 规则（26pt）的排版结果
            # （2026-08-15 真机回归根因 B）。
            if n in _EXACT_ANCHOR_NORMS:
                continue
            if (in_cover and _cjk_ratio(t) >= 0.5 and 8 <= len(n) <= 30
                    and not _field_name(t) and not _CHAPTER_RE.match(n)
                    and not _DATE_RE.match(t)):
                out.append(p)
        elif match == "date":
            if in_cover and _DATE_RE.match(t):
                out.append(p)
        elif match == "prefix":
            if anchor and _field_match(_field_name(t), anchor):
                out.append(p)
        else:  # exact
            if n == norm or (anchor and _field_name(t) == anchor) or n == anchor:
                out.append(p)
    return out
