"""早睡 (EarlyZ) 3.2 — DocumentAnalyzer：DOCX → DocumentIR（§4，ADR-003）。

只读分析，绝不修改文档。P0 覆盖：
- styles.xml → StyleNode（含 basedOn 继承链）
- numbering.xml → ListDefinitionNode
- document.xml → sections/paragraphs/tables/figures/fields/captions
- 包健康 → PackageHealth
- 内容 hash（内容保持不变量）

语义角色（P0 样式驱动，完整语义分类后续扩展）：
heading / body / list_item / figure_para / table_cell / caption
"""

from __future__ import annotations

import hashlib
import re
import zipfile
from collections import OrderedDict
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

from earlyz.ir.models import (
    IR_SCHEMA_VERSION,
    DocumentIR,
    FigureNode,
    FieldNode,
    FormatProperties,
    ListDefinitionNode,
    ListLevelNode,
    ParagraphNode,
    RunNode,
    SectionNode,
    StyleNode,
    TableNode,
)
from earlyz.ir.package import content_hash, inspect_package
from earlyz.ir.selectors import build_node_id, normalize_text
from earlyz.ir.ooxml_format import (
    extract_extra_bundle,
    extract_structure_bundle,
    merge_extra_bundles,
)

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
PIC = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
_EQ_OLE_RE = re.compile(r"^(Equation|MathType)", re.I)


def _has_equation_ole(para) -> bool:
    """MathType / Word Equation OLE，不是 Visio 图。"""
    for el in para.iter():
        pid = el.get("ProgID") or ""
        if pid and _EQ_OLE_RE.match(pid):
            return True
    return False

# ── 纯文本公式行（2026-09-28 夜班） ─────────────────────────────────────────
# 用户产物里 8 个公式段**全是纯文本**（`R_n = B log₂(...)`），而 equation_display
# 判定只看 OMML / OLE → 这些段落落到「正文」角色 → 继承正文首行缩进 2 字符、
# 左右对齐 → 用户看到「公式都没处理」。
# 口径：**宁可漏认，不可误认** —— 正文被误判成公式（居中 + 去缩进）比漏掉一个
# 公式严重得多。判据全部来自用户真机文档的正反例（tests/test_plain_formula_detection.py）。
_PSEUDO_KW_RE = re.compile(
    r"\b(FOR|END|IF|THEN|ELSE|WHILE|RETURN|REPEAT|UNTIL|DO|TO|STEP)\b", re.I)
_ALGO_STEP_RE = re.compile(r"^\s*\d+\s*[:：]\s")
_CJK_ANY_RE = re.compile(r"[\u4e00-\u9fff]")
# 强记号：关系符 / 大运算符 / 希腊字母（真机 `Σₙ` 用的是希腊大写 Σ U+03A3，
# 不是数学记号 ∑ U+2211 —— 两个都要收，否则目标函数行会被漏认）
_STRONG_MATH_RE = re.compile(
    r"[=∑∏∫√∞±×÷·≤≥≠≈≡∝∈∉⊂⊆∪∩∂∇"
    r"αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΥΦΨΩ]")
# 「数学味」：上下标记号 / 希腊字母 / 数学符号 / 字母数字与关系符相邻
_MATHY_RE = re.compile(
    r"[_^][A-Za-z0-9{]"
    r"|[αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΥΦΨΩ]"
    r"|[∑∏∫√∞±×÷·≤≥≠≈∈∂∇]"
    r"|[²³¹⁰₀₁₂₃₄ₙₐₑₓ]"
    r"|(?:\d|[A-Za-z)\]}])[ \t]*[=<>≤≥][ \t]*(?:\d|[A-Za-z(])"
)
# 符号密度：运算符/括号/上下标/希腊/数学符号 占全行比例（英文散文会很低）
_SYM_DENSITY_RE = re.compile(
    r"[=+\-*/^_(){}\[\]<>≤≥≠≈∈∑∏∫√±×÷·∂∇∞"
    r"αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΥΦΨΩ"
    r"²³¹⁰₀₁₂₃₄ₙₐₑₓ⋅∘]")
_PROSE_HEAD_RE = re.compile(r"^(见|由式|其中|式中|代入|整理|即|故|所以|证明|定理|算法|图|表|式)")


def looks_like_plain_formula(text: str) -> bool:
    """纯文本写的公式行（源文档里没有 OMML/OLE 时用）。

    真机正例：`R_n = B log₂(1 + p_n h_n / (N₀ B)) (1)`
    真机反例：`其中，D_n 为任务数据量…本地执行时延为 T_n^loc = C_n / f_n^loc。 (4)`
              `1: FOR episode = 1 TO K do`（伪代码，绝不能居中）
              `约束：C1: a_n ∈ {0,1}；…。`
    """
    t = (text or "").strip()
    if not t or len(t) > 120:
        return False
    if _ALGO_STEP_RE.match(t) or _PSEUDO_KW_RE.search(t):
        return False
    if _PROSE_HEAD_RE.match(t):
        return False
    if len(_CJK_ANY_RE.findall(t)) > 2:
        return False
    if re.search(r"[。；！？，、.;]$", t):
        return False
    if not _STRONG_MATH_RE.search(t):
        return False
    if not _MATHY_RE.search(t):
        return False
    return len(_SYM_DENSITY_RE.findall(t)) / max(1, len(t)) >= 0.12


DOCUMENT_PART = "word/document.xml"
STYLES_PART = "word/styles.xml"
NUMBERING_PART = "word/numbering.xml"

_HEADING_STYLE_RE = None  # lazy import re


def _local(tag: str) -> str:
    return tag.split("}")[-1]


def _text_of(elem) -> str:
    """元素内全部 w:t 文本拼接。"""
    out = []
    for t in elem.iter():
        if _local(t.tag) == "t" and t.text:
            out.append(t.text)
    return "".join(out)


# ────────────────────────────────────────────────────────────────
# 样式解析
# ────────────────────────────────────────────────────────────────

def _parse_rpr_format(rpr) -> FormatProperties:
    fp = FormatProperties()
    if rpr is None:
        return fp
    fp.rpr_extra_xml = extract_extra_bundle(rpr, "rPr")
    fonts = rpr.find(f"{W}rFonts")
    if fonts is not None:
        fp.font_east_asia = fonts.get(f"{W}eastAsia")
        fp.font_ascii = fonts.get(f"{W}ascii")
    sz = rpr.find(f"{W}sz")
    if sz is not None and sz.get(f"{W}val"):
        try:
            fp.font_size_pt = int(sz.get(f"{W}val")) / 2.0
        except ValueError:
            pass
    def _onoff(name: str) -> Optional[bool]:
        el = rpr.find(f"{W}{name}")
        if el is None:
            return None
        return el.get(f"{W}val") not in ("0", "false", "off", "none")
    fp.bold = _onoff("b")
    fp.italic = _onoff("i")
    fp.small_caps = _onoff("smallCaps")
    underline = rpr.find(f"{W}u")
    if underline is not None:
        fp.underline = underline.get(f"{W}val") or "single"
    color = rpr.find(f"{W}color")
    if color is not None:
        fp.color = color.get(f"{W}val")
    return fp


def _parse_ppr_format(ppr) -> FormatProperties:
    fp = FormatProperties()
    if ppr is None:
        return fp
    fp.ppr_extra_xml = extract_extra_bundle(ppr, "pPr")
    jc = ppr.find(f"{W}jc")
    if jc is not None:
        fp.jc = jc.get(f"{W}val")
    ind = ppr.find(f"{W}ind")
    if ind is not None:
        fp.ind_first_line = _twips(ind.get(f"{W}firstLine"))
        fp.ind_first_line_chars = _twips(ind.get(f"{W}firstLineChars"))
        fp.ind_left = _twips(ind.get(f"{W}left"))
        fp.ind_right = _twips(ind.get(f"{W}right"))
        fp.ind_hanging = _twips(ind.get(f"{W}hanging"))
    tabs = ppr.find(f"{W}tabs")
    if tabs is not None:
        tab = tabs.find(f"{W}tab")
        if tab is not None and tab.get(f"{W}pos"):
            fp.tab_pos = _twips(tab.get(f"{W}pos"))
    pbdr = ppr.find(f"{W}pBdr")
    if pbdr is not None and pbdr.find(f"{W}bottom") is not None:
        fp.p_bdr_bottom = True
    spacing = ppr.find(f"{W}spacing")
    if spacing is not None:
        fp.spacing_before = _twips(spacing.get(f"{W}before"))
        fp.spacing_after = _twips(spacing.get(f"{W}after"))
        # 行单位（2026-08-18 用户彻查「行就是行」）：模板标题段前/段后
        # 用 beforeLines/afterLines（0.8行/0.5行），磅值仅是 Word 换算值
        # ——只读磅值会把行单位丢成 12.55磅/7.85磅。行值优先（用户说
        # 模板用什么单位 earlyz 就用什么单位，不给歧义）。
        bl = spacing.get(f"{W}beforeLines")
        al = spacing.get(f"{W}afterLines")
        if bl is not None:
            try:
                fp.spacing_before_lines = float(bl) / 100.0
            except (ValueError, TypeError):
                fp.spacing_before_lines = None
        if al is not None:
            try:
                fp.spacing_after_lines = float(al) / 100.0
            except (ValueError, TypeError):
                fp.spacing_after_lines = None
        fp.line = _twips(spacing.get(f"{W}line"))
        fp.line_rule = spacing.get(f"{W}lineRule")
        if fp.line and spacing.get(f"{W}lineRule") == "auto":
            fp.line_spacing = round(fp.line / 240, 4)
    # keepNext/keepLines 三态（2026-08-17 投诉彻查 D2）：
    # 无元素=未设；w:val="0"/"false"=显式关闭；无 val 或 val≠0=开启。
    # 旧代码 find() is not None 把显式关闭也读成 True——模板示例区
    # 11 个 keepNext 全为 val=0，被误读后 spec 反向推断出 keep_next:false
    # 去关产物标题（08-08 铁律② 违规）。
    def _tri_state(el) -> Optional[bool]:
        if el is None:
            return None
        v = el.get(f"{W}val")
        return v not in ("0", "false", "off")
    fp.keep_next = _tri_state(ppr.find(f"{W}keepNext"))
    fp.keep_lines = _tri_state(ppr.find(f"{W}keepLines"))
    fp.widow_control = _tri_state(ppr.find(f"{W}widowControl"))
    fp.page_break_before = _tri_state(ppr.find(f"{W}pageBreakBefore"))
    fp.snap_to_grid = _tri_state(ppr.find(f"{W}snapToGrid"))
    ol = ppr.find(f"{W}outlineLvl")
    if ol is not None:
        try:
            fp.outline_lvl = int(ol.get(f"{W}val"))
        except (ValueError, TypeError):
            pass
    frame = ppr.find(f"{W}framePr")
    if frame is not None:
        dc = frame.get(f"{W}dropCap")
        if dc:
            fp.drop_cap = dc
            try:
                fp.drop_cap_lines = int(frame.get(f"{W}lines") or 0) or None
            except (ValueError, TypeError):
                fp.drop_cap_lines = None
    return fp


def _twips(v: Optional[str]) -> Optional[int]:
    if v is None:
        return None
    try:
        return int(v)
    except ValueError:
        return None


def _parse_styles(styles_xml: str, rel_map: dict) -> list[StyleNode]:
    root = ET.fromstring(styles_xml)
    nodes: list[StyleNode] = []
    for style in root.findall(f"{W}style"):
        style_id = style.get(f"{W}styleId")
        if not style_id:
            continue
        name_el = style.find(f"{W}name")
        style_name = name_el.get(f"{W}val") if name_el is not None else style_id
        based = style.find(f"{W}basedOn")
        fp = FormatProperties()
        fp = _merge_format(fp, _parse_ppr_format(style.find(f"{W}pPr")))
        fp = _merge_format(fp, _parse_rpr_format(style.find(f"{W}rPr")))
        nodes.append(
            StyleNode(
                node_id=f"style:{style_id}",
                style_id=style_id,
                style_type=style.get(f"{W}type", "paragraph"),
                name=style_name,
                based_on=based.get(f"{W}val") if based is not None else None,
                next=style.find(f"{W}next").get(f"{W}val") if style.find(f"{W}next") is not None else None,
                link=style.find(f"{W}link").get(f"{W}val") if style.find(f"{W}link") is not None else None,
                hidden=style.get(f"{W}hidden") in ("1", "true"),
                is_default=style.get(f"{W}default") in ("1", "true"),
                style_format=fp,
            )
        )
    return nodes


def _parse_doc_defaults(styles_xml: str) -> dict:
    """提取 docDefaults/rPrDefault → {font_east_asia, font_ascii, font_size_pt, ...}。"""
    root = ET.fromstring(styles_xml)
    out: dict = {}
    doc_defaults = root.find(f"{W}docDefaults")
    if doc_defaults is None:
        return out
    rpr_default = doc_defaults.find(f"{W}rPrDefault")
    if rpr_default is None:
        return out
    rpr = rpr_default.find(f"{W}rPr")
    if rpr is None:
        return out
    fp = _parse_rpr_format(rpr)
    for k, v in fp.__dict__.items():
        if v is not None:
            out[k] = v
    return out


def _merge_format(base: FormatProperties, extra: FormatProperties) -> FormatProperties:
    for k, v in extra.__dict__.items():
        if v is None:
            continue
        if k == "ppr_extra_xml":
            base.ppr_extra_xml = merge_extra_bundles(base.ppr_extra_xml, v, "pPr")
        elif k == "rpr_extra_xml":
            base.rpr_extra_xml = merge_extra_bundles(base.rpr_extra_xml, v, "rPr")
        else:
            setattr(base, k, v)
    return base


# ────────────────────────────────────────────────────────────────
# 编号解析
# ────────────────────────────────────────────────────────────────

def _parse_numbering(numbering_xml: str) -> list[ListDefinitionNode]:
    root = ET.fromstring(numbering_xml)
    abstract: dict[int, list[ListLevelNode]] = {}
    for abs_num in root.findall(f"{W}abstractNum"):
        aid = int(abs_num.get(f"{W}abstractNumId", -1))
        levels = []
        for lvl in abs_num.findall(f"{W}lvl"):
            ilvl = int(lvl.get(f"{W}ilvl", 0))
            fmt_el = lvl.find(f"{W}numFmt")
            text_el = lvl.find(f"{W}lvlText")
            start_el = lvl.find(f"{W}start")
            pstyle = lvl.find(f"{W}pStyle")
            levels.append(
                ListLevelNode(
                    node_id=f"list_level:{aid}:{ilvl}",
                    level=ilvl,
                    num_fmt=fmt_el.get(f"{W}val") if fmt_el is not None else None,
                    lvl_text=text_el.get(f"{W}val") if text_el is not None else None,
                    start=int(start_el.get(f"{W}val")) if start_el is not None else None,
                    style_link=pstyle.get(f"{W}val") if pstyle is not None else None,
                    level_extra_xml=extract_structure_bundle(lvl, "lvl"),
                )
            )
        abstract[aid] = levels

    out: list[ListDefinitionNode] = []
    for num in root.findall(f"{W}num"):
        num_id = int(num.get(f"{W}numId", -1))
        aid_el = num.find(f"{W}abstractNumId")
        aid = int(aid_el.get(f"{W}val")) if aid_el is not None else -1
        levels = abstract.get(aid, [])
        out.append(
            ListDefinitionNode(
                node_id=f"list_definition:{num_id}",
                num_id=num_id,
                abstract_num_id=aid,
                levels=levels,
                multi_level=len(levels) > 1,
            )
        )
    return out


# ────────────────────────────────────────────────────────────────
# 文档解析
# ────────────────────────────────────────────────────────────────

def _resolve_rel_target(rel_map: dict | None, rid: str) -> str | None:
    """rId → target part 名（word/header1.xml 等；/word/... 或相对路径归一）。"""
    if not rel_map or not rid:
        return None
    tgt = rel_map.get(rid) or ""
    if not tgt:
        return None
    if tgt.startswith("word/"):
        return tgt
    return "word/" + tgt.lstrip("/")


def _parse_sectpr(sectpr, index: int, rel_map: dict | None = None, part_name: str = "word/document.xml") -> SectionNode:
    s = SectionNode(node_id=f"section:{index}", index=index)
    s.section_extra_xml = extract_structure_bundle(sectpr, "sectPr")
    typ = sectpr.find(f"{W}type")
    if typ is not None:
        s.break_type = typ.get(f"{W}val")
    pg = sectpr.find(f"{W}pgSz")
    if pg is not None:
        s.page_size = {k: v for k, v in {
            "w": _twips(pg.get(f"{W}w")),
            "h": _twips(pg.get(f"{W}h")),
            "orient": pg.get(f"{W}orient"),
        }.items() if v is not None}
    mar = sectpr.find(f"{W}pgMar")
    if mar is not None:
        s.margins = {k: v for k, v in {
            "top": _twips(mar.get(f"{W}top")),
            "bottom": _twips(mar.get(f"{W}bottom")),
            "left": _twips(mar.get(f"{W}left")),
            "right": _twips(mar.get(f"{W}right")),
            "header": _twips(mar.get(f"{W}header")),
            "footer": _twips(mar.get(f"{W}footer")),
            "gutter": _twips(mar.get(f"{W}gutter")),
        }.items() if v is not None}
    cols = sectpr.find(f"{W}cols")
    if cols is not None:
        s.columns = {
            "num": _twips(cols.get(f"{W}num")) or 1,
            "space": _twips(cols.get(f"{W}space")),
            "equal_width": cols.get(f"{W}equalWidth") != "0",
            "separator": cols.get(f"{W}sep") in ("1", "true", "on"),
            "widths": [
                {"w": _twips(c.get(f"{W}w")), "space": _twips(c.get(f"{W}space"))}
                for c in cols.findall(f"{W}col")
            ],
        }
    pgtype = sectpr.find(f"{W}pgNumType")
    if pgtype is not None:
        s.pg_num_type = {k: v for k, v in {
            "fmt": pgtype.get(f"{W}fmt"),
            "start": _twips(pgtype.get(f"{W}start")),
        }.items() if v is not None}
    # 页眉/页脚引用（rId → part 名），供"模板 header/footer 格式"规则配对
    for ref in sectpr:
        tag = ref.tag.replace(W, "")
        if tag not in ("headerReference", "footerReference"):
            continue
        rid = ref.get(f"{R}id")
        kind = tag[:-len("Reference")]  # header / footer
        ref_type = ref.get(f"{W}type") or "default"
        s.header_footers.append({
            "kind": kind,
            "ref_type": ref_type,
            "rel_id": rid,
            "part": _resolve_rel_target(rel_map, rid) if rel_map else None,
        })
    return s


def _field_type_from_instr(instr: str) -> tuple[str, list[str]]:
    parts = instr.strip().split()
    if not parts:
        return "UNKNOWN", []
    ftype = parts[0].upper()
    return ftype, parts[1:]


def _parse_drawing(elem, rel_map: dict) -> Optional[dict]:
    """提取 inline/anchor 图片信息。返回 dict 或 None。"""
    for inline in elem.iter(f"{WP}inline"):
        return _extract_drawing_common(inline, "inline", rel_map)
    for anchor in elem.iter(f"{WP}anchor"):
        return _extract_drawing_common(anchor, "float", rel_map)
    return None


def _extract_drawing_common(container, wrap: str, rel_map: dict) -> dict:
    d = {"wrap": wrap}
    ext = container.find(f"{WP}extent")
    if ext is not None:
        d["extent"] = {"cx": _twips(ext.get("cx")), "cy": _twips(ext.get("cy"))}
    docpr = container.find(f"{WP}docPr")
    if docpr is not None:
        d["docpr_id"] = docpr.get("id")
        d["name"] = docpr.get("name")
    blip = container.find(f".//{A}blip")
    if blip is not None:
        d["embed"] = blip.get(f"{R}embed")
    return d


class _FieldTracker:
    """fldChar 状态机：begin → instrText → separate → result → end。"""

    def __init__(self):
        self.open: list[FieldNode] = []
        self.fields: list[FieldNode] = []
        self._instr_parts: dict[int, list[str]] = {}
        self._result_parts: dict[int, list[str]] = {}

    def feed(self, run_elem) -> None:
        """按文档顺序处理 run 内的全部子元素（fldChar 是流，不是"每 run 一个"）。

        2026-10-12 事故：旧实现用 ``run_elem.find(fldChar)`` 只看第一个 fldChar
        就 return——引用联动写出的 REF 域是单 run 形态（begin+instrText+separate+
        结果+end 同 run），end 永远读不到 ⇒ 字段全被判残缺 ⇒ 交付验证硬拦整篇编译。
        OOXML 的 run 内容本就是有序流，必须逐个走。
        """
        for child in run_elem:
            tag = child.tag
            if tag == f"{W}fldChar":
                ftype = child.get(f"{W}fldCharType")
                if ftype == "begin":
                    node = FieldNode(
                        node_id=f"field:{len(self.fields)}:{len(self.open)}",
                        field_type="UNKNOWN",
                        complete=False,
                    )
                    self.open.append(node)
                    self._instr_parts[id(node)] = []
                    self._result_parts[id(node)] = []
                elif ftype == "separate":
                    if self.open:
                        self._result_parts[id(self.open[-1])] = []
                elif ftype == "end":
                    if self.open:
                        node = self.open.pop()
                        instr = " ".join(self._instr_parts.get(id(node), [])).strip()
                        node.field_type, node.targets = _field_type_from_instr(instr)
                        node.instruction = instr
                        node.result = "".join(self._result_parts.get(id(node), []))
                        node.complete = True
                        self.fields.append(node)
            elif tag == f"{W}instrText":
                if child.text and self.open:
                    self._instr_parts[id(self.open[-1])].append(child.text)
            elif tag == f"{W}t":
                if self.open and child.text:
                    self._result_parts[id(self.open[-1])].append(child.text)
            elif tag == f"{W}tab":
                if self.open:
                    self._result_parts[id(self.open[-1])].append("\t")

    def finish(self) -> list[FieldNode]:
        # 未闭合的 begin → 残缺字段（complete=False）
        for node in self.open:
            node.field_type, node.targets = _field_type_from_instr(
                " ".join(self._instr_parts.get(id(node), [])).strip()
            )
            node.instruction = " ".join(self._instr_parts.get(id(node), [])).strip()
            self.fields.append(node)
        self.open = []
        return self.fields


def _heading_level(ppr, pstyle_name: Optional[str]) -> Optional[int]:
    ol = ppr.find(f"{W}outlineLvl") if ppr is not None else None
    if ol is not None:
        try:
            _v = int(ol.get(f"{W}val"))
        except (ValueError, TypeError):
            return None
        # outlineLvl 深到 7 级以下（val>=6）实际都是「正文文本」的误设：Word 的
        # 「正文文本」是 val=9，但源文档常写成 val=8（劳动合同实测 107 段），
        # 被判成 heading9 → 不走正文规则、没有首行缩进。真标题不会到这么深。
        return None if _v >= 6 else _v + 1
    if pstyle_name:
        import re
        m = re.search(r"heading\s*(\d+)", pstyle_name, re.IGNORECASE)
        if m:
            return int(m.group(1))
    return None


def _semantic_role(ppr, has_drawing: bool, in_table: bool, numpr: bool, heading_level: Optional[int]) -> str:
    if heading_level is not None:
        return "heading"
    if in_table:
        return "table_cell"
    if has_drawing:
        return "figure_para"
    if numpr:
        return "list_item"
    return "body"


# ── 语义角色单一真相源（analyzer 与 action handler 共享，禁止两套分类）──

_CAPTION_RE = re.compile(
    r"^\s*(表|图|公式|式|Fig(?:ure)?\.?|TABLE|Table)\s*[（(]?\s*"
    r"(?:\d+(?:\.\d+)*|[Xx]|[IVXLCM]+)\s*[)）]?\s*[:：.—–-]?\s*",
    re.I,
)
_REFS_ITEM_RE = re.compile(r"^\s*\[\d+\]")


def _strip_caption_instruction(text: str) -> str:
    """去掉模板夹在图/表题里的 *小5号* / （黑体）说明，只留核心题名。"""
    t = (text or "").strip()
    t = re.sub(r"[*＊].*$", "", t)
    t = re.sub(
        r"[（(][^）)]{0,80}(?:黑体|宋体|仿宋|楷体|号|磅|加粗)[^）)]{0,80}[）)]\s*$",
        "",
        t,
    )
    return t.strip()
_KEYWORDS_RE = re.compile(
    r"^\s*(?:关键词|Key(?:\s+words?|words)|KEY\s*WORDS?|Index\s*Terms?)\s*[:：—–-]?",
    re.IGNORECASE,
)
_CLC_RE = re.compile(r"^\s*(中图法分类号|中图分类号|CLC)\s*[:：]?", re.I)
_ABSTRACT_RE = re.compile(
    r"^\s*(?:摘\s+要|摘要\s*[:：]|Abstract\b)",
    re.I,
)
_TITLE_STYLE_RE = re.compile(r"^(?:title|title1|subtitle)\b", re.I)
_AUTHORS_RE = re.compile(
    r"^\s*(?:authors?|作者名|作者)\s*[:：]",
    re.I,
)
_AFFIL_RE = re.compile(
    r"大学|学院|研究所|University|Institute|Department|ZipCode|邮政编码|单位全名",
    re.I,
)
_CHROME_RE = re.compile(
    r"^REPLACE THIS LINE|^第\?+卷|Vol\s*\?+",
    re.I,
)
FRONT_ROLES = {
    "article_title", "article_title_en", "authors", "authors_en",
    "affiliation", "affiliation_en",
    "abstract", "abstract_en", "keywords", "keywords_en",
    "classification",
}
# 期刊题名区的「非题名」强证据：学位论文封面首行也可能被判成题名，孤立题名不算
_ARTICLE_HEAD_STRONG = frozenset({
    "authors", "authors_en", "affiliation", "affiliation_en",
    "abstract", "abstract_en", "keywords", "keywords_en",
    "classification", "article_title_en",
})


def zone_is_article_head(roles) -> bool:
    """封面区是否为期刊题名区（article head）——交付门 / writer / lint 共用单一谓词。

    roles：封面区**非空段**的语义角色序列。判据只看**开头连续的前部角色段**
    （遇到第一个非前部角色即停）：必须以题名开头，且至少有一个非题名强角色
    （作者/单位/摘要/关键词）。是 → 前部角色按模板规则下发格式；否 → 学位论文封面。

    2026-11-10 修正（英文期刊 IEEE 实测）：原判据要求「封面区里**全部**角色都是
    前部角色」。期刊题名区后面紧跟的就是 "I. INTRODUCTION" 这类章节标题，而
    封面边界 `find_cover_end` 会把短标题一起算进封面区（cap 按「≤24 字」推），
    于是 heading1 混进序列 → 判假 → lint 把题名/作者/摘要/关键词全部当封面跳过
    → 模板的题名 24pt、作者 11pt 一条都不下发。补充第二分支：题名区被章节标题
    截断的期刊——开头连续前部角色段以题名打头且含强角色，区内的非前部角色
    **全部**是章节标题/首字下沉。语料面 A/B（199 份）：学位论文封面（题名+
    作者之后跟 body/日期行）不翻转，全部维持原判。
    """
    rs = [r or "" for r in roles]
    if not rs:
        return False
    if all(r in FRONT_ROLES for r in rs):
        return any(r in _ARTICLE_HEAD_STRONG for r in rs)
    lead: list[str] = []
    for r in rs:
        if r not in FRONT_ROLES:
            break
        lead.append(r)
    if len(lead) < 2 or lead[0] not in ("article_title", "article_title_en"):
        return False
    if not any(r in _ARTICLE_HEAD_STRONG for r in lead):
        return False
    # 题名区里的非前部角色只允许是章节标题/首字下沉（期刊第一章节标题）；
    # 出现 body（学位论文封面的日期行/校名行/说明长段）即维持学位论文口径。
    return all(r.startswith("heading") or r == "drop_cap" for r in rs[len(lead):])


_FRONT_LABEL_RUN_RE = re.compile(
    r"^(摘\s*要|摘要|关键词|Abstract|Key(?:\s+words?|words)|Index\s*Terms?|中图法分类号|中图分类号)\s*[:：—–\*]?\s*$",
    re.I,
)


def _is_front_label_run(text: str) -> bool:
    t = (text or "").strip()
    if not t:
        return False
    if _FRONT_LABEL_RUN_RE.match(t):
        return True
    return bool(re.match(
        r"^(摘\s*要|摘要|关键词|Abstract|Key(?:\s+words?|words))\s*\*",
        t, re.I,
    ))


class FrontRoleAssigner:
    """前部题名/作者/单位/摘要：标题之前按顺序补位，避免整页只剩双栏。"""

    def __init__(self) -> None:
        self.seen_heading = False
        self.phase = "title"
        self.seen_abstract = False
        self.seen_keywords = False
        self.seen_authors = False
        self.seen_affiliation = False

    def feed(self, text: str, pstyle_name: str, role: str) -> str:
        if self.seen_heading:
            return role
        if str(role or "").startswith("heading"):
            # Direct outlineLvl is mutable formatting: applying an article-title rule
            # can turn the title into raw heading1. The front/body boundary must use
            # stable evidence only — the original heading style name or heading text.
            # 空段带标题样式（作者遗留/模板残留）不作「已见章节」证据：否则前置角色
            # 机器在文档开头即被静默关闭，题名/单位永不识别（2026-09-25 真机：第 0
            # 段空段带 heading 1 样式 → 报告题名/校名全部落 body）。与封面边界
            # 「空段不作锚点」同一判据。
            _t0 = (text or "").strip()
            style_heading = bool(_t0) and bool(
                re.search(r"(?:heading|标题)\s*\d*", pstyle_name or "", re.I))
            text_heading = bool(
                _is_numeric_heading(text) or _is_chapter_title_text(text)
                or _is_cn_numbered_heading(text))
            if style_heading or text_heading:
                self.seen_heading = True
                return role
            if (self.phase == "en_title" and pstyle_name
                    and _TITLE_STYLE_RE.match(pstyle_name.strip())):
                self.phase = "authors"
                return "article_title_en"
            role = "body"
        if role in ("list_item", "table_cell", "figure_para", "equation_display",
                    "caption", "caption_figure", "caption_table", "caption_figure_en",
                    "caption_table_en", "toc_entry"):
            return role
        t = (text or "").strip()
        if not t or _CHROME_RE.match(t):
            return role
        if role in FRONT_ROLES:
            if role == "article_title" and self.phase == "en_title":
                self.phase = "authors"
                return "article_title_en"
            if role == "authors":
                if self.seen_authors:
                    return "authors_en"
                self.seen_authors = True
                self._advance(role)
                return "authors"
            if role == "affiliation":
                if self.seen_affiliation:
                    return "affiliation_en"
                self.seen_affiliation = True
                self._advance(role)
                return "affiliation"
            if role == "abstract":
                if self.seen_abstract:
                    return "abstract_en"
                self.seen_abstract = True
                self._advance(role)
                return "abstract"
            if role == "keywords":
                if self.seen_keywords:
                    return "keywords_en"
                self.seen_keywords = True
                self._advance(role)
                return "keywords"
            self._advance(role)
            return role
        if self.phase == "en_title":
            letters = sum(1 for c in t if c.isascii() and c.isalpha())
            compact = re.sub(r"\s+", "", t)
            if letters >= 8 and compact and letters / max(len(compact), 1) >= 0.6 and len(t) <= 180:
                self.phase = "authors"
                return "article_title_en"
            return role
        if _AFFIL_RE.search(t) and len(t) <= 160 and self.phase in ("title", "authors", "affil"):
            self.phase = "affil"
            return "affiliation"
        if self.phase == "title":
            self.phase = "authors"
            return "article_title"
        if self.phase == "authors":
            if re.search(r"[：:]", t) and not _AUTHORS_RE.search(t):
                return role
            # 正文形态不作作者（2026-09-25 批3）：纯报告题名之后直接是正文，
            # 期刊式相位机器会把正文长段判成 authors → 套模板「作者」取值
            # （居中加粗大字号）。作者行是名字列表：短行且无句读收尾。
            # 注意只在此分支收紧，不全局终止机器——投稿须知/声明等前部长段
            # 之后仍有 摘要/英文题名 等前部段落（CJC 模板即如此）。
            if len(t) <= 120 and not _looks_like_body_text(t):
                if _is_authorish_line(t):
                    if self.seen_authors:
                        return "authors_en"
                    self.seen_authors = True
                    return "authors"
                # 封面版式（2026-09-25 审计发现 2）：题名之前先出现封面头
                # （「毕业设计（论文）」）时，相位机器会把真题名顺次判成
                # authors → 套模板「作者」取值（真机 26pt）。作者行必须有
                # 「名字列表/作者标签」证据，否则按题名形态归 article_title。
                # 这一条必须在 feed 本体里：写动作定位时会重新跑一遍本机器，
                # 只放在 assign_front_roles 的后置修正里会被绕过（真机实锤：
                # 题名段拿不到题名取值，只有封面头段拿到）。
                if _looks_like_cover_title(t):
                    if _ascii_dominant(t):
                        return "article_title_en"
                    self._advance("article_title")
                    return "article_title"
            return role
        if self.phase == "affil":
            if _AFFIL_RE.search(t) and len(t) <= 160:
                if self.seen_affiliation:
                    return "affiliation_en"
                self.seen_affiliation = True
                return "affiliation"
            return role
        return role

    def _advance(self, role: str) -> None:
        if role == "article_title":
            self.phase = "authors"
        elif role in ("authors", "affiliation"):
            self.phase = "affil" if role == "affiliation" else "authors"
        elif role == "abstract":
            self.phase = "keywords"
        elif role == "keywords":
            self.phase = "en_title"


def _looks_like_body_text(t: str) -> bool:
    """正文段落形态（不应被前置机器当作作者/单位，2026-09-25 批3）。

    判据用「句读 + 章节标记 + 长句」，先剥掉模板行内说明（`*` 之后的批注、
    括号注记）再判——中文期刊模板的示例行常把说明写在行内（CJC：
    「作者名1) 作者名2),3) *字体为3号仿宋_GB2312」），名字列表因此必须
    在剔除批注后仍是「多短词」形态才算作者行。
    """
    s = (t or "").strip()
    if not s:
        return False
    core = re.split(r"[*＊]", s)[0].strip()
    stripped = re.sub(r"[（(][^（()）]{0,60}[)）]", "", core).strip()
    core = stripped or core
    if re.search(r"[。；！？]", core):
        return True
    if re.search(r"\.\s|\.$", core):
        return True     # 英文句子收尾（期刊模板前部英文段）
    if _is_cn_numbered_heading(core) or _is_numeric_heading(core):
        return True
    if len(core) >= 40:
        toks = [x for x in re.split(r"[,，、;；\s]+", core) if x]
        if len(toks) >= 3 and max(len(x) for x in toks) <= 12:
            return False     # 名字列表（多位作者，中英文）
        return True
    return False


def _median_font_size_pt(paragraphs) -> Optional[float]:
    """全文非空段的中位数字号（判「首页大字号题名」用）。"""
    sizes = []
    for p in paragraphs:
        if not (getattr(p, "text", "") or "").strip():
            continue
        d = getattr(p, "direct_format", None)
        s = getattr(d, "font_size_pt", None) if d is not None else None
        if s:
            sizes.append(float(s))
    if not sizes:
        return None
    sizes.sort()
    return sizes[len(sizes) // 2]


def _looks_like_title_block(paragraphs) -> bool:
    """文档是否「像带题名」——无标题样式时也跑前置角色机器。

    首个非空段 4-40 字、不以句号结尾，且（居中 或 字号 ≥ 全文非空段中位数
    1.25 倍）。2026-09-25 实测：纯报告（无摘要/目录、无标题样式）此前闸门
    不过 → 整篇落 body，题名拿不到 article_title → 被套模板正文取值。
    """
    first = None
    for p in paragraphs:
        if (getattr(p, "text", "") or "").strip():
            first = p
            break
    if first is None:
        return False
    t = (first.text or "").strip()
    if not (4 <= len(t) <= 40) or t.endswith(("。", "；")):
        return False
    d = getattr(first, "direct_format", None)
    jc = getattr(d, "jc", None) if d is not None else None
    size = getattr(d, "font_size_pt", None) if d is not None else None
    med = _median_font_size_pt(paragraphs)
    big = bool(size and med and float(size) >= med * 1.25)
    return jc == "center" or big


def assign_front_roles(paragraphs, style_name_by_id: Optional[dict] = None) -> None:
    has_heading = any(str(getattr(p, "semantic_role", "") or "").startswith("heading")
                      for p in paragraphs)
    if not (has_heading or _looks_like_title_block(paragraphs)):
        return
    box = FrontRoleAssigner()
    names = style_name_by_id or {}
    for p in paragraphs:
        style_id = getattr(p, "style_id", None) or ""
        style_name = names.get(style_id, style_id)
        p.semantic_role = box.feed(p.text or "", style_name, p.semantic_role or "body")
    _fix_cover_misassigned_authors(paragraphs)


# ── 封面题名/日期被套「作者」取值：角色归位（2026-09-25 横向机制审计发现 2）──────
# 相位机器是期刊式假设（题名→作者→单位→摘要），封面版式下会漂移：真机
# 王福勋论文封面 [3]「毕业设计（论文）」被当题名后，真题名 [5] 顺次判成
# authors → 套模板「作者」取值（26pt 居中加粗），日期 [18] 判成 authors_en
# → 10.5pt。作者行的证据是「名字列表/作者类标签」，不是「前部第 N 行」。
_AUTHOR_LABEL_RE = re.compile(
    r"(?:姓\s*名|名\s*字|作\s*者|学\s*号|指导教?师|通信作者|Author|Name)")
_DATEISH_RE = re.compile(
    r"(?:[二〇○零一二三四五六七八九Ｏ0-9]{2,4}\s*年|(?<!\d)(?:19|20)\d{2}\s*[-/年])")
# 前部区域终止词（作者/题名只出现在此之前）
_FRONT_STOP_WORDS = ("摘要", "ABSTRACT", "目录", "目次", "致谢", "谢辞", "参考文献")
_CJK_RE = re.compile(r"^[\u4e00-\u9fff]+$")


def _ascii_dominant(t: str) -> bool:
    """英文行判据（与 en_title 分支同口径）：字母占比 ≥ 0.6 且 ≥ 8 字母。"""
    letters = sum(1 for c in t if c.isascii() and c.isalpha())
    compact = re.sub(r"\s+", "", t)
    return letters >= 8 and bool(compact) and letters / max(len(compact), 1) >= 0.6


def _is_authorish_line(t: str) -> bool:
    """作者行形态：带作者类标签，或「名字列表」（中/英）。

    单个 2-6 字纯中文名（「张三」「王福勋」）也算；脚注序号先剥掉
    （「作者名1)」→「作者名」）。非作者行（题名、日期）返回 False。
    """
    s = (t or "").strip()
    if not s:
        return False
    if _AUTHOR_LABEL_RE.search(s):
        return True
    core = re.split(r"[*＊]", s)[0].strip()
    toks = [x for x in re.split(r"[,，、;；/\s]+", core) if x]
    toks = [re.sub(r"[0-9,，)）(（*＊#]", "", x) for x in toks]
    toks = [x for x in toks if x]
    if not toks:
        return False
    if len(toks) == 1:
        # 单个中文名（2-6 字）；带括号残迹的短行不算——「毕业设计（论文）」
        # 剥掉括号后剩「毕业设计」，靠原串非纯中文把它挡在作者行之外。
        return bool(_CJK_RE.match(core)) and 2 <= len(core) <= 6
    # 名字列表：每个词 2-8 字。拆字名字（「张 三1)  李 四2)」）是 1 字词 +
    # 脚注序号，凭序号把它和「郑 重 声 明」这类拆字标题区分开。
    if all(len(x) <= 2 for x in toks) and re.search(r"[0-9]", core):
        return True
    return all(2 <= len(x) <= 8 for x in toks)


def _looks_like_cover_title(t: str) -> bool:
    """封面题名形态：短、无句读、非日期、非标签行、非模板说明行。"""
    s = (t or "").strip()
    if not (4 <= len(s) <= 40):
        return False
    if re.search(r"[。；！？，、：:]", s):
        return False
    if _DATEISH_RE.search(s):
        return False
    if _AUTHOR_LABEL_RE.search(s):
        return False
    if re.match(r"^(?:注|备注|说明|示例|封面|格式)", s):
        return False
    # 拆字行（「张 三  李 四」「郑 重 声 明」）：词都很短，不当题名
    parts = [x for x in re.split(r"\s+", s) if x]
    if len(parts) >= 2 and all(len(x) <= 2 for x in parts):
        return False
    return True


def _is_cover_header_line(t: str, following) -> bool:
    """封面头行（「毕业设计（论文）」「毕业论文」）：短 + 括号后缀 + 后面跟着更长的题名。

    封面头不是文档题名。判成 article_title 会把它的字号（真机 26pt）混进
    题名取值投票（模板侧），也会让它在目标文档上被套题名格式；判成 body
    即「保持原稿」（用户 2026-09-25 拍板 B：封面段落格式以原稿为准）。
    """
    s = (t or "").strip()
    if not (4 <= len(s) <= 10) or "（" not in s:
        return False
    return any(_looks_like_cover_title(x) and len(x) > len(s) for x in following)


def _fix_cover_misassigned_authors(paragraphs) -> None:
    """前部区角色归位（2026-09-25 审计发现 2）。

    ① 封面头行（「毕业设计（论文）」）不是题名：后面跟着更长的题名行时降为
       body（保持原稿）。判成 article_title 会把它的字号（真机 26pt）混进
       模板题名取值投票；
    ② 日期/说明类行不套「作者」取值（真机「二○二四年五月」被套 authors_en
       10.5pt 居中）；
    ③ 题名形态却带 authors/authors_en 的段归 article_title（feed 本体已处理，
       这里兜底）。

    只在前部区动作（见到 heading 或 摘要/目录 等终止词即停）。
    """
    rows = list(paragraphs)
    for pos, p in enumerate(rows):
        role = str(getattr(p, "semantic_role", "") or "")
        text = (getattr(p, "text", "") or "").strip()
        if role.startswith("heading") or (text and len(text) <= 16
                                          and text.upper() in
                                          {w.upper() for w in _FRONT_STOP_WORDS}):
            break
        if role not in ("authors", "authors_en", "article_title",
                        "article_title_en"):
            continue
        if not text or _is_authorish_line(text):
            continue
        nxt = [(x.text or "").strip() for x in rows[pos + 1:pos + 4]
               if (getattr(x, "text", "") or "").strip()]
        if _is_cover_header_line(text, nxt):
            p.semantic_role = "body"      # 封面头：保持原稿
        elif role in ("authors", "authors_en") and _looks_like_cover_title(text):
            p.semantic_role = ("article_title_en" if _ascii_dominant(text)
                               else "article_title")
        elif role in ("authors", "authors_en"):
            p.semantic_role = "body"      # 日期/说明等：保持原稿


def assign_back_roles(paragraphs) -> None:
    """后部区域上下文角色（2026-08-23 CJC v9 实锤）：
    appendix_heading 之后的非空 body/equation_display 段 → appendix_body，
    直到遇到其他非收集类角色段（heading/refs/bio/表格等）即止；空段跳过但
    保持区域（模板「附录X.」后是分节空段，再接说明段——空段不是边界）。
    模板侧：附录 X. 后的说明段成为附录内容格式样本；目标侧：附录标题后的
    定理证明/公式段获得附录内容格式。区域语义=标题后连排正文，机制层，无学校常量。
    Author 段由 classify 的 bio_body 规则截断（同一区域的后半段互斥）。"""
    in_appendix = False
    for p in paragraphs:
        role = getattr(p, "semantic_role", None) or ""
        text = (getattr(p, "text", "") or "").strip()
        if role == "appendix_heading":
            in_appendix = True
            continue
        if not in_appendix:
            continue
        if role in ("body", "equation_display") and text:
            p.semantic_role = "appendix_body"
            continue
        if not text:
            continue  # 空段跳过但保持区域（分节空段不是边界）
        in_appendix = False
_HEADING_NAME_RE = re.compile(r"heading\s*(\d+)", re.IGNORECASE)

# ── 内容推断标题的形态闸门 ──
# 没有 Word 标题样式时，「是不是标题」只能靠文本形态判。真标题是短语：
# 短、无段内句读、无尾随句读。反例（2026-09-25 张鑫佳报告 × 论文模板实锤）：
# 「2. 基层司法工作者：包括海珠区司法局…共8名，重点了解…供给情况；」只有
# 「≤80 字」一道闸时被判 heading1 → 整段套上模板标题格式（18pt 加粗居中 +
# 段前分页 + 进大纲）→ 用户看到「正文被改成一级标题」。中文序号判据当时已收紧，
# 数字判据漏了——两条路径现在共用本闸门。
_HEADING_INNER_PUNCT_RE = re.compile(r"[。；，、！？：]")
_HEADING_TAIL_PUNCT = ("。", "；", "，", "、", "！", "？", "：")
_MAX_INFERRED_HEADING_CHARS = 40
# ── 编号标题的文本形态判据（2026-10-31 手标 126 段校准） ────────────────
# 「同一个『一、』形态，哪些是标题、哪些是正文条款」的实测结论（P/R/F1）：
#   `_looks_like_heading_shape`（≤40 字 + 禁 。；，、：）  P=1.00 R=0.82（保守，漏 13/71）
#   本判据（允「、」与尾冒号 + 排目录条目 + 排字段行）        P=0.96 R=0.99
#   「比相邻正文短 50%/60%」「同文档短簇」                  无增益（F1 不变、召回反降）
# 目录条目（页码尾/点线）与字段行（「标签：值」）才是主要干扰源，不是长度比。
_HEAD_TAIL_PAGE_RE = re.compile(r"[\s\u3000]*[\d０-９]{1,4}\s*$")
_HEAD_LEADER_RE = re.compile(r"([.…·]{3,}|[－—]{3,})")
_HEAD_FIELD_RE = re.compile(r"^[^：:\s]{1,20}[：:]\s*\S")
_HEAD_SENTENCE_CHARS = "。；，"


def _looks_like_numbered_heading(text: str) -> bool:
    """编号段是否成标题形态（比 `_looks_like_heading_shape` 宽一档）。

    拒：>40 字、目录条目（页码尾／点线）、句子形态（含 。；，）、字段行
    （冒号后还有内容）。允：段内「、」与**尾冒号**——「四、提交投标文件截止
    时间、开标时间和地点」「三、获取招标文件：」都是真标题，旧闸门把它们
    整批判成正文（实测 13 处漏判全部落在这两类）。
    """
    t = (text or "").strip()
    if not t or len(t) > _MAX_INFERRED_HEADING_CHARS:
        return False
    if _HEAD_TAIL_PAGE_RE.search(t) or _HEAD_LEADER_RE.search(t):
        return False
    body = _LEAD_NUMBER_RE.sub("", t).strip()
    body = re.sub(r"[（(][^（()）]{0,40}[）)]\s*$", "", body).strip()
    if not body:
        return True
    if any(ch in body for ch in _HEAD_SENTENCE_CHARS):
        return False
    if _HEAD_FIELD_RE.match(body):
        return False
    # 尾冒号也不放行：合同「七、名词解释：」「四、服务时间：」是字段/条款标签，
    # 放行后这类段被抬成标题套上章标题字号（实测数据委托/数据融合合同 100→88.9）。
    # 代价是漏掉招标公告的「三、获取招标文件：」——手标 126 段里 F1 反而更高
    # （拒绝尾冒号 P=1.00 R=0.96 F1=0.978；放行 P=0.96 R=0.99 F1=0.972）。
    return body[-1] not in tuple(_HEAD_SENTENCE_CHARS + "：:")
# 序号前缀：序号里的「、」「.」是编号分隔符，不是句读，判形态前先摘掉
_LEAD_NUMBER_RE = re.compile(
    r"^(?:第[一二三四五六七八九十百\d]{1,4}[章节篇条]|"
    r"[一二三四五六七八九十]{1,3}\s*[、．.]|"
    r"[（(]\s*[一二三四五六七八九十]{1,3}\s*[）)]|"
    r"[\d]{1,2}\s*[、)）]|"
    r"[\d]{1,2}[.．](?:\d{1,2}[.．])*|"
    r"[IVXivx]{1,4}[.．]|[A-Za-z][.．]|[①②③④⑤⑥⑦⑧⑨⑩])\s*"
)


def _looks_like_heading_shape(text: str) -> bool:
    """无样式时判断文本是否具备标题形态（短 + 摘掉序号后无段内/尾随句读）。

    纯序号段（「第一章」「1.」）不在此闸门否定——它们是否成标题由各正则决定，
    本闸门只负责排除「句子形态」。
    """
    t = (text or "").strip()
    if not t or len(t) > _MAX_INFERRED_HEADING_CHARS:
        return False
    body = _LEAD_NUMBER_RE.sub("", t).strip()
    # 模板说明行常带尾部括号标注（「1 引言（一级标题：黑体五号）」）——
    # 括号内是标注文本，不是标题句子本身，判形态前摘掉。
    body = re.sub(r"[（(][^（()）]{0,40}[）)]\s*$", "", body).strip()
    if not body:
        return True
    if _HEADING_INNER_PUNCT_RE.search(body):
        return False
    return not body.endswith(_HEADING_TAIL_PUNCT)


# 语义章标题词表：无 heading 样式但独立成段的章级标题（摘要/目录/致谢/声明等）
# 匹配条件：去除全部空白后的整段文本 ∈ 词表（独立成段 + 短文本，正文长段落天然排除）
# 数字章节识别（2026-08-15 全篇样式修复）：模板/论文的章标题常是
# 「第1章 绪论」「1 绪论」「1.1 研究背景」「1.1.1 小节」——
# 无 outlineLvl 无 pStyle 时此前被 classify 为 body，heading 规则永不命中
# （真机回归实锤：章标题被 body 规则刷成宋体小四无粗无居中）。
_NUM_CHAP_RE = re.compile(
    r"^(?:第\s*[一二三四五六七八九十百\d]+\s*[章节](?:[\s□　]|$)"  # 第1章 / 第 5 章 绪论
    r"|[\d]{1,2}[.．]?\s(?:[\u4e00-\u9fff]|[A-Za-z]{4,}))")  # 「1 绪论」「1. 引言」单级（日期/
    # 小数如「3.14 是」的 . 后仍有数字→\s 不匹配，安全）；1.1 / 7.1.1 不是章。
    # 2026-10-08：①「第 X 章」章号两侧允许空格（模板常写「第 5 章」，此前漏识别→整章
    # 标题被当正文套小四）；②英文词须 ≥4 字母——「10 MHz」「4 GHz」「2 MB」等单位行
    # 此前命中单字母分支被误判章级标题（真机：表数据行被套黑体小2加粗居中）。
_NUM_H2_RE = re.compile(r"^\d+[.．]\d+(?![.．\d])\s*[\u4e00-\u9fffA-Za-z□　·]")
_NUM_H3_RE = re.compile(r"^\d+[.．]\d+[.．]\d+(?:[.．]\d+)*\s*[\u4e00-\u9fffA-Za-z□　·]")
# 中文序号标题（报告/公文/总结常见，常无 Word 标题样式）
_CN_H1_RE = re.compile(r"^[一二三四五六七八九十]{1,3}\s*[、．.]\s*\S")
_CN_H2_RE = re.compile(r"^[（(]\s*[一二三四五六七八九十]{1,3}\s*[）)]\s*\S")
_CN_H3_RE = re.compile(r"^(?:\d{1,2}\s*[、]|[①②③④⑤⑥⑦⑧⑨⑩])\s*\S")
# 期刊节标题：I. INTRODUCTION / A. Network Model（无 heading 样式）
_ROMAN_H1_RE = re.compile(r"^(?:[IVXLCM]{2,}|[IVX])\.\s+\S")
_LETTER_H2_RE = re.compile(r"^[A-Z]\.\s+\S")
# 目录条目尾特征：点 leader/箭头/页码尾随（「第1章  绪论1」「1.1 研究背景1」「致  谢I」）
_TOC_ENTRY_RE = re.compile(r"(…|\.\.|→|—)\s*\d+\s*$|[\dIVX]{1,4}\s*$")


def _is_numeric_heading(text: str) -> str | None:
    """数字章节标题 → heading 级别（1/2/3），非标题返回 None。

    TOC 目录条目（带页码尾/点 leader）排除——目录条目不是标题。
    2026-10-31：多级小节（1.1 / 1.1.1）保留旧闸门（禁「、：」）——放宽后合同
    条款段的「4.4.1 …为：」会整批变成三级标题（实测 121 段），那类是条款不是
    标题；「第X章」「1 绪论」这类章级标题改用放宽判据（「第3章 系统设计、实现」
    此前被误判成正文）。
    """
    t = (text or "").strip()
    if _TOC_ENTRY_RE.search(t):
        return None
    if _NUM_H3_RE.match(t) or _NUM_H2_RE.match(t):
        if not _looks_like_heading_shape(t):
            return None
        return "heading3" if _NUM_H3_RE.match(t) else "heading2"
    if not _looks_like_numbered_heading(t):
        return None
    if _NUM_CHAP_RE.match(t):
        return "heading1"
    if _ROMAN_H1_RE.match(t):
        return "heading1"
    if _LETTER_H2_RE.match(t):
        # J. K. Author… 是参考文献，不是 A. Network Model
        if re.match(r"^[A-Z]\.\s+[A-Z]\.", t):
            return None
        return "heading2"
    return None


_BACK_HEADING_NORM = {
    "参考文献": "refs_heading",
    "references": "refs_heading",
    "致谢": "ack_heading",
    "acknowledgment": "ack_heading",
    "acknowledgement": "ack_heading",
    "作者简介": "bio_heading",
    "background": "bio_heading",
}


def _back_heading_role(text: str) -> str | None:
    """后部短标题：参考文献/致谢/附录跟模板自己的段取值，不当 heading1。"""
    t = (text or "").strip()
    if not t:
        return None
    t = re.sub(r"^\[\d+\]\s*", "", t)
    # 附录带描述形态（2026-10-08 真机：用户论文「附录 A 公式(3-5) 的 NP-hard 性证明
    # （详细证明置于此处，正文从略）」此前判 body）——行尾批注先摘，再按「附录 +
    # 字母/数字 + 空格 + 短短语」判 heading；句子形态（句读/动词开头）不接受。
    t2 = strip_trailing_annotation(t)
    m = re.match(
        r"^附录[\t \u3000]*([A-Za-z0-9一二三四五六七八九十]{1,4})[\t \u3000]+"
        r"(?!(?:是|为|由|在|的|所|中|上|下|里|内))(\S[\s\S]{0,32})$", t2)
    if m and len(t2) <= 40 and not _HEADING_INNER_PUNCT_RE.search(t2):
        return "appendix_heading"
    core = _strip_caption_instruction(t)
    core = re.split(r"[*＊]", core)[0].strip()
    norm = re.sub(r"[\s\u3000]+", "", core)
    norm = re.sub(r"[.。:：]+$", "", norm)
    if not norm or len(norm) > 16:
        return None
    key = norm.lower() if all(ord(c) < 128 for c in norm) else norm
    if key in _BACK_HEADING_NORM:
        return _BACK_HEADING_NORM[key]
    if re.match(r"^附录[A-Za-z0-9X零一二三四五六七八九十]*$", norm):
        return "appendix_heading"
    if re.match(r"^appendix[a-z0-9]*$", key):
        return "appendix_heading"
    return None


_CHAPTER_TITLE_WORDS = {
    "摘要", "ABSTRACT", "目录", "致谢", "参考文献", "附录", "结论", "前言",
    "引言", "后记", "符号说明", "术语表", "索引", "原创性声明", "版权使用授权书",
    "学位论文原创性声明", "学位论文版权使用授权书", "开题报告", "任务书",
    "答辩委员会", "答辩记录", "评阅意见", "指导教师意见",
}


# ── 文档级层级事实：一级层级是否已被「第X章」占据（EZI-0023）───────────────
# 「一、调查背景」是标题还是条款，差别不在文本，而在**这份文档的一级层级有没有
# 被「第X章」占掉**：报告里「一、」就是一级；东台招标里「第X章 ×3」才是一级，
# 「一、项目基本情况」是公告条目/合同条款，模板给它们是正文。纯文本函数看不见
# 文档，所以这个事实单独算一次、按文档树缓存（`classify_paragraph_role` 逐段调用）。
_CHAPTER_TIER_RE = re.compile(r"^第\s*[一二三四五六七八九十百千零〇0-9]{1,4}\s*章")

#: 文档根 → tier。键用 `id(root)` 且值里**持住 root 强引用**：不持住的话树被回收后
#: id 会复用，另一份文档可能拿到上一份的结论。上限 2：一次编译同时活着的树就一两棵，
#: 超了淘汰最旧的（下次重扫一遍，几毫秒）；强引用只留 2 棵，长语料跑批不会攒内存。
_CHAPTER_TIER_CACHE: "OrderedDict[int, tuple[object, bool]]" = OrderedDict()
_CHAPTER_TIER_CACHE_MAX = 2


def _doc_body(root):
    """取 `w:body`：调用方可能传 document 根，也可能传 body 本身。"""
    if root is None:
        return None
    if getattr(root, "tag", None) == f"{W}body":
        return root
    return root.find(f"{W}body")


def _scan_chapter_tier(root) -> bool:
    """扫全篇：一级层级是否已被「第X章」占据。

    判据与标题取样同源：正文流（只看 body 的直接子块——表格里的「1.6」「60天」
    不是层级证据）、非目录条目（页码尾/点 leader 的「第一章 绪论……1」不是层级
    证据）、成标题形态（`_looks_like_numbered_heading`：≤40 字、无句读、非字段行）。
    ≥1 个即算占据：一份「第一章 总则」的合同，其后「一、」也是条款不是一级。
    """
    body = _doc_body(root)
    if body is None:
        return False
    for block in body:                 # 直接子块：w:p 才看，w:tbl 整块跳过
        if block.tag != f"{W}p":
            continue
        text = "".join(t.text or "" for t in block.iter(f"{W}t")).strip()
        if not text or not _CHAPTER_TIER_RE.match(text):
            continue
        if _TOC_ENTRY_RE.search(text):
            continue
        if not _looks_like_numbered_heading(text):
            continue
        return True
    return False


def _is_cn_numbered_heading(text: str, *, chapter_tier: bool = False) -> str | None:
    """中文序号标题 → heading 级别（报告/公文常见，无 Word 标题样式）。

    「一、调查背景」→ heading1；「（一）调查对象」→ heading2；
    「1、调查方法」「①…」→ heading3。

    收紧条件（防把正文句子当标题）：≤40 字、不以句号/分号/逗号结尾、非目录
    条目形态（带页码尾/点 leader）。2026-09-25 实测：报告类文档的「一、（一）」
    式标题若没挂 Word 标题样式，此前一律落 body → 被套模板正文取值（首行缩进/
    字号），而模板的标题取值落不到（用户场景：报告 × 论文模板）。

    ``chapter_tier``（EZI-0023）：调用方所在文档的一级层级**已被「第X章」占据**
    时，「一、」只能是公告条目/合同条款，不是一级标题 → 返回 ``None``。东台招标
    实测：34 个「一级标题」样本里 16 段就是被这条认成一级的「一、项目基本情况」
    /「一、甲方职责」，模板对它们的处理是「Normal 样式 + 没有大纲级别」。
    默认 ``False``：纯文本函数看不见文档，而报告/公文类文档（无章标题样式、
    「一、」就是一级）必须保持旧行为——把默认值改成 ``None`` 会让整类报告掉标题。
    文档级事实由 :func:`_chapter_tier` 供给（与 ``classify_paragraph_role`` 同源）。
    """
    t = (text or "").strip()
    if _TOC_ENTRY_RE.search(t):
        return None
    if not _looks_like_numbered_heading(t):
        return None
    if not re.search(r"[\u4e00-\u9fff]", t):
        return None
    if _CN_H2_RE.match(t):
        return "heading2"
    if _CN_H1_RE.match(t):
        return None if chapter_tier else "heading1"
    if _CN_H3_RE.match(t):
        return "heading3"
    return None


def _chapter_tier_for_root(root) -> bool:
    """`_scan_chapter_tier` 的缓存层（同一棵树只扫一次）。"""
    key = id(root)
    hit = _CHAPTER_TIER_CACHE.get(key)
    if hit is not None and hit[0] is root:
        _CHAPTER_TIER_CACHE.move_to_end(key)
        return hit[1]
    tier = _scan_chapter_tier(root)
    _CHAPTER_TIER_CACHE[key] = (root, tier)
    _CHAPTER_TIER_CACHE.move_to_end(key)
    while len(_CHAPTER_TIER_CACHE) > _CHAPTER_TIER_CACHE_MAX:
        _CHAPTER_TIER_CACHE.popitem(last=False)
    return tier


def _chapter_tier(ppr) -> bool:
    """段落所在文档是否「章层级」文档；取不到文档树 → False（= 旧行为）。

    供给两条路：① `analyze_document` 用 stdlib ET 解析，显式传给
    `_parse_paragraph`/`classify_paragraph_role`；② 动作层（`ensure_heading_outline`
    等，lxml）只拿到段落，靠 `getroottree()` 反查文档再走缓存。反查不到（传进来的
    是 `pPr` 的游离副本，例如 EZI-0014 的标题证据判定）就按 False 处理——宁可退回
    旧行为，也不要凭半截树猜文档结构。
    """
    get = getattr(ppr, "getroottree", None)
    if get is None:          # xml.etree.ElementTree 没有父指针
        return False
    try:
        root = get().getroot()
    except Exception:        # noqa: BLE001 — 游离元素/异常树一律按旧行为
        return False
    if root is None or getattr(root, "tag", None) != f"{W}document":
        return False
    return _chapter_tier_for_root(root)


# 标题文本归一：空白 + 全角空格 + `□`（模板拿 □ 当占位空格用：「摘□□要」「第 5 章」）。
# 2026-10-08：□ 此前不参与归一 → 「摘□□要」既不识别为「摘要」同名标题（格式规则生
# 不出来），也不命中分页词表（摘要不另起页）——一个字符两处症状。
_TITLE_WS_RE = re.compile(r"[\s\u3000\u25a1]+")

# 行尾括号标注（模板说明/作者批注）：「Abstract（Times New Roman 三号加粗，居中）」。
# 2026-10-08 真机缺陷（用户报「英文摘要也没处理」直查）：标题判定只认纯文本——带
# 行尾标注的标题段既不进同名标题池（规则生不出来），命中不了分页词表（不另起页）。
# 与 `_looks_like_heading_shape` 既有口径一致：判形态前摘掉行尾括号注记。
# 只作**判定回退**用（原文不删——文字内容是用户资产，删除走「先报列表、点名才改」）。
_ANNO_TAIL_RE = re.compile(r"[（(][^（()）]{0,40}[）)]\s*$")


def strip_trailing_annotation(text: str) -> str:
    """摘掉行尾括号标注（可连续多组）——\u4ec5判定用，绝不用于改写原文。"""
    t = (text or "").strip()
    while True:
        t2 = _ANNO_TAIL_RE.sub("", t).strip()
        if t2 == t:
            return t
        t = t2
# 声明类兜底（2026-10-08 用户拍板「按类兜底」）：原创性声明/郑重声明/诚信声明/
# 学位论文独创性声明……统一进「声明」池，对齐模板的声明示例取值。
_DECLARATION_RE = re.compile(r"^[\u4e00-\u9fff]{0,6}声明$")


def _chapter_title_core(text: str) -> str | None:
    """章级标题词判定核心（含行尾标注剥离回退）——返回规范化核心串，不中返回 None。

    2026-10-08：「Abstract（Times New Roman 三号加粗，居中）」这类带行尾注记的
    标题段，原文直接归一会超词表/长度判定——先按原文判、不中时摘掉行尾括号注记
    重判（顺序保证干净文本行为不变，注记形态是纯新增识别面）。
    """
    t = text or ""
    for cand in (t, strip_trailing_annotation(t)):
        norm = _TITLE_WS_RE.sub("", cand)
        if not norm or len(norm) > 20:
            continue
        canon = norm.upper() if norm.isascii() else norm
        # 声明/授权书等带"学位论文"前缀的变体
        core = canon[4:] if canon.startswith("学位论文") else canon
        if core in _CHAPTER_TITLE_WORDS or _DECLARATION_RE.match(core):
            return core
    return None


def _is_chapter_title_text(text: str) -> bool:
    return _chapter_title_core(text) is not None


def _is_top_tier_text(text: str) -> bool:
    """文本自身能否作一级标题（章层级文档的唯一合法一级证据）。"""
    t = (text or "").strip()
    if not t:
        return False
    return bool(_CHAPTER_TIER_RE.match(t)) or _is_chapter_title_text(t)


def _trusted_heading_level(ppr, pstyle_name: str, text: str, chapter_tier: bool) -> int | None:
    """段落自带的一级标题级别；**章层级文档里只信「章级文本」的 outlineLvl=0**。

    EZI-0023 A3（不得照搬错级别）：东台招标源文档把 14 段条款项（「2.采购代理机构
    信息」「1.供货时间及地点」…）直接标成 `outlineLvl=0`，同文档「60天」「1.6」
    「2.2」这类表格数字也带 0 级——源大纲级别不可信。它与 EZI-0014（不得**新增**）
    互补：那条管引擎别自己加，这条管别把源文档标的错级别原样继承成一级标题。

    只卡 0 级：文档一级层级已被「第X章」占掉时，`outlineLvl=0` 只有在文本自己
    是章级时才作数（「第X章」/摘要·目录·参考文献等章级词/标题样式）。1、2 级不动
    ——那是小节层级，「编号深度优先于直接 outlineLvl」的既有校正（2026-10-27 交大）
    继续管。
    """
    hlevel = _heading_level(ppr, pstyle_name)
    if hlevel is None or not chapter_tier or hlevel != 1:
        return hlevel
    if re.search(r"heading\s*\d+", pstyle_name or "", re.IGNORECASE):
        return hlevel          # 标题样式是作者意图（样式声明的层级优先）
    if _is_top_tier_text(text):
        return hlevel          # 文本自己就是章级（第X章 / 摘要 / 目录 / 参考文献…）
    return None                # 其余 0 级不照搬 → 交给文本判据（可能是 body）


# `_CHAPTER_TITLE_WORDS` 里结构上属于「章」序列的词（正文起点 / 正文之后）；词表其余
# （摘要/ABSTRACT/目录/声明/授权书…）是前置页标题。分类层：只分组，不带任何取值。
_CHAPTER_CLASS_TITLE_WORDS = {"结论", "前言", "引言", "后记", "索引", "致谢", "参考文献", "附录"}


def heading_title_key(text: str) -> str | None:
    """不带编号的章级标题（`_is_chapter_title_text` 词表）→ 同名匹配键；否则 None。

    模板样本与目标段落两侧共用：去空白/全角空格/□、英文大写、去「学位论文」前缀
    （「学位论文原创性声明」≡「原创性声明」）；声明类（2026-10-08 用户拍板按类兜底）
    统一归到「声明」——原创性声明/郑重声明/诚信声明与模板的声明示例同池。
    行尾括号注记不参与词表判定（回退剥离，见 `_chapter_title_core`）。
    """
    core = _chapter_title_core(text)
    if core is None:
        return None
    if core.endswith("声明"):
        return "声明"
    return core


def heading_rule_applies(selector: dict | None, role: str | None, text: str) -> bool:
    """heading 角色规则是否作用于这一段（lint / writer / 建议层同一谓词）。

    2026-09-29 真论文 × CJC：摘要/ABSTRACT/目录/声明/授权书全判 heading1，被套期刊
    一级标题取值（居中→左对齐 14pt）。取值口径改为：
    - 同名规则（selector.title_key，取自模板同名段）：只作用于同名标题；
    - 普通 heading 规则：不作用于前置页标题（模板没有同名段 = 模板没定义它怎么排，
      保持原稿），也不作用于已有同名规则的标题（selector.skip_title_keys）；
      结论/引言等章级标题模板无同名段时仍走普通规则（结构上就是一章）。
    """
    if not str(role or "").startswith("heading"):
        return True
    sel = selector or {}
    want = sel.get("title_key")
    key = heading_title_key(text)
    if want is not None:
        return key == want
    if key is None:
        return True
    if key in set(sel.get("skip_title_keys") or ()):
        return False
    return key in _CHAPTER_CLASS_TITLE_WORDS


def classify_paragraph_role(
    ppr,
    pstyle_name: str,
    has_drawing: bool,
    in_table: bool,
    numpr: bool,
    text: str,
    in_hyperlink: bool = False,
    has_omath_para: bool = False,
    has_omath: bool = False,
    has_equation_ole: bool = False,
    chapter_tier: bool | None = None,
) -> str:
    """与 analyze_document 完全一致的段落角色分类（handler 复用入口）。

    在 analyzer 既有规则基础上补充 caption 识别（表/图/公式标题文本模式），
    使表标题/图标题不被误分类为 body 而遭批量正文规则改写。

    2026-08-15 全篇样式修复：
    - 数字章节识别（第1章 绪论/1.1 概述/1.1.1 小节 → heading1/2/3）
    - hyperlink 段落 = 目录条目（toc_entry），隔离 body/heading 规则
      （否则 body 规则会给目录条目强加首行缩进/行距——用户「目录没改」
      的观感来源；目录条目格式由模板目录规则或保持原样）

    ``chapter_tier``（EZI-0023）：文档一级层级是否已被「第X章」占据。``None`` =
    从段落所在文档树推算（动作层 lxml 走 `_chapter_tier`），取不到就是 ``False``
    = 旧行为（报告/公文类的「一、」仍是一级标题）。`analyze_document` 用 stdlib
    ET 解析，显式传值（ET 元素没有父指针）。
    """
    if chapter_tier is None:
        chapter_tier = _chapter_tier(ppr)
    if in_hyperlink:
        # Index Terms / Key words 有时带超链接，仍是关键词行不是目录条目
        if _KEYWORDS_RE.match(text or ""):
            return "keywords"
        if _ABSTRACT_RE.match(text or ""):
            return "abstract"
        return "toc_entry"   # 目录条目（TOC 域生成的超链接段落）
    back = _back_heading_role(text)
    if back:
        return back
    hlevel = _trusted_heading_level(ppr, pstyle_name, text, chapter_tier)
    if hlevel is not None:
        # 编号深度优先于「直接格式里遗留的 outlineLvl」（2026-10-27 交大实测）：
        # 源文档是模板副本时，1.1 段落带着模板的 outline=2，分类器信了它 →
        # 判成三级标题；ensure_heading_outline 又按分类结果「校正」成同一个值，
        # 形成自证循环，动作层永远修不动。样式声明的层级仍然优先（那是作者意图）。
        from_style = bool(re.search(r"heading\s*\d+", pstyle_name or "", re.IGNORECASE))
        if not from_style:
            nrole = _is_numeric_heading(text or "")
            if nrole:
                nlevel = int(nrole[-1])
                if chapter_tier and nlevel == 1 and not _is_top_tier_text(text):
                    # 一级已被「第X章」占据：「6. 投标函」这类「数字+空格+标题」
                    # 单级形态（`_NUM_CHAP_RE` 的「1 绪论」分支）只是章内条目，
                    # 不升一级 → 保留段落自带级别（此处必 ≥2 级）。
                    return f"heading{hlevel}" if hlevel >= 1 else "heading"
                if nlevel != hlevel:
                    return nrole
        return f"heading{hlevel}" if hlevel >= 1 else "heading"
    if in_table:
        return "table_cell"
    jc = None
    if ppr is not None:
        j = ppr.find(f"{W}jc")
        if j is not None:
            jc = j.get(f"{W}val")
    if (has_omath_para
            or (has_equation_ole and (not (text or "").strip()
                                      or re.match(r"^[（(]\s*\d+\s*[)）]$", (text or "").strip())
                                      or len((text or "").strip()) <= 6))
            or (has_omath and not (text or "").strip())
            or (jc == "right" and re.match(r"^[（(]\s*\d+\s*[)）]$", (text or "").strip()))
            # 2026-09-28：纯文本公式行（源文档没把公式做成 OMML/OLE 时）
            or looks_like_plain_formula(text or "")):
        return "equation_display"
    if has_drawing:
        return "figure_para"
    if ppr is not None:
        fr = ppr.find(f"{W}framePr")
        if fr is not None and fr.get(f"{W}dropCap"):
            return "drop_cap"
    if numpr:
        return "list_item"
    if _is_chapter_title_text(text):
        return "heading1"  # 语义章标题（摘要/目录/致谢/声明等，视为一级标题）
    nlevel = _is_numeric_heading(text)
    if nlevel:
        if chapter_tier and nlevel == "heading1" and not _is_top_tier_text(text):
            # 章层级文档里没有第二个一级：「6. 投标函」这种单级编号是章内条目
            return "heading2"
        return nlevel    # 数字章节：第1章 绪论 / 1.1 概述 / 1.1.1 小节
    clevel = _is_cn_numbered_heading(text, chapter_tier=chapter_tier)
    if clevel:
        return clevel    # 中文序号章节：一、/（一）/1、/①（报告类无样式标题）
    if _KEYWORDS_RE.match(text):
        return "keywords"
    if _CLC_RE.match(text or ""):
        return "classification"
    if _ABSTRACT_RE.match(text):
        return "abstract"
    if _REFS_ITEM_RE.match(text or ""):
        return "refs_item"
    # 英文作者简介段（2026-08-23 CJC v9 实锤：模板 AuthorX 说明段/目标 Author1
    # 简介 = 作者简介区内容，格式=模板对应段）；与附录内容区边界互斥。
    # 形态：Author1 (…/AuthorX, …/Author1(…——需排除 "Authors:" 类普通词。
    if re.match(r"^Author(?:[Xx]\b|\d+\b|(?=\s*[（(]))", (text or "").strip()):
        return "bio_body"
    if pstyle_name and _TITLE_STYLE_RE.match(pstyle_name.strip()):
        return "article_title"
    if _AUTHORS_RE.search(text or "") and len((text or "").strip()) <= 160:
        return "authors"
    if _CAPTION_RE.match(text) or _CAPTION_RE.match(_strip_caption_instruction(text)):
        # 标题段收紧：用去掉模板说明（*小5号*）后的核心文本判长短/句读
        core = _strip_caption_instruction(text)
        stripped = (core or text).strip()
        lead = stripped.lstrip()
        # 「图6-1为…」是正文句子，不是题注
        if re.match(r"^(图|表)\S{0,16}为", lead):
            return "body"
        if re.match(r"^(TABLE|Table)\s+(?:\d+|[IVXLCM]+|[Xx])\b", lead, re.I):
            return "caption_table"
        if re.match(r"^Fig(?:ure)?\.?\s*(?:\d+|[Xx])", lead, re.I):
            return "caption_figure"
        if lead.startswith("表"):
            return "caption_table"
        if lead.startswith("图"):
            return "caption_figure"
        if len(stripped) <= 80 and not re.search(r"[，。；、]", stripped):
            return "caption"
        return "body"
    return "body"


# ────────────────────────────────────────────────────────────────
# 主入口
# ────────────────────────────────────────────────────────────────

def _hf_run_sig(pel):
    """段内 run 序列机制签名（2026-08-23 CJC 页眉实锤）：
    返回 ['TEXT'|'{PAGE}'|'TAB'|'BR'] token 序列。
    {PAGE} = 页眉/页脚页码域（instrText 含 PAGE）；TEXT = 文本 run；
    TAB/BR = 制表位/换行（run 内 w:tab/w:br）。
    用途：页眉页脚排版时按模板将目标「死页码文本/空白分隔」桥接为域/制表位。
    """
    sig = []
    for r in pel.iter(f"{W}r"):
        for child in r:
            tag = child.tag.split("}")[-1]
            if tag == "t":
                sig.append("TEXT")
            elif tag == "tab":
                sig.append("TAB")
            elif tag == "br":
                sig.append("BR")
            elif tag == "instrText":
                txt = (child.text or "").upper()
                sig.append("{PAGE}" if "PAGE" in txt else "INSTR")
            # fldChar begin/separate/end 不计 token（instrText 已标）
    return sig


def _hf_tab_segments(pel) -> list:
    """段内行×列结构：[[行1列1, 行1列2...], [行2列1...]]
    （2026-08-23 CJC 首页页眉实锤：header3 两行「第??卷… TAB XXX TAB Vol…」BR
    「年/月 TAB 英文刊名…」——按行比例在目标段近似插入 w:tab，机制层布局，
    不拆不改文本。BR 分隔行、TAB 分隔列。"""
    lines: list = []
    line: list = []
    cur: list = []
    for r in pel.iter(f"{W}r"):
        for child in r:
            tag = child.tag.split("}")[-1]
            if tag == "t":
                cur.append(child.text or "")
            elif tag == "tab":
                line.append(re.sub(r"[\s\u3000]+", "", "".join(cur)))
                cur = []
            elif tag == "br":
                line.append(re.sub(r"[\s\u3000]+", "", "".join(cur)))
                cur = []
                lines.append([s for s in line if s])
                line = []
    if cur or line:
        line.append(re.sub(r"[\s\u3000]+", "", "".join(cur)))
        lines.append([s for s in line if s])
    return [ln for ln in lines if ln]


def analyze_document(path: str | Path, include_render: bool = False) -> DocumentIR:
    """把 DOCX 解析为 DocumentIR（§8.3 analyze_document 的内核）。"""
    p = Path(path)
    ir = DocumentIR(document_id=f"doc:{p.stem}")

    with zipfile.ZipFile(p) as z:
        names = set(z.namelist())
        ir.package = inspect_package(str(p)).to_dict()

        # 关系映射
        rel_map: dict[str, str] = {}
        if "word/_rels/document.xml.rels" in names:
            rels_root = ET.fromstring(
                z.read("word/_rels/document.xml.rels").decode("utf-8", errors="replace")
            )
            for rel in rels_root:
                rel_map[rel.get("Id")] = rel.get("Target", "")

        document_xml = z.read("word/document.xml").decode("utf-8", errors="replace")

        if "word/styles.xml" in names:
            styles_xml = z.read("word/styles.xml").decode("utf-8", errors="replace")
            ir.styles = _parse_styles(styles_xml, rel_map)
            ir.font_table = {"doc_defaults": _parse_doc_defaults(styles_xml)}
        if "word/numbering.xml" in names:
            ir.numbering = _parse_numbering(
                z.read("word/numbering.xml").decode("utf-8", errors="replace")
            )

        # 页眉/页脚只采集格式，不复制模板文字。pPr/rPr extras 与正文走同一
        # 无损通道，PAGE 域和关系 ID 仍由结构子系统持有。
        def _part_formats(prefix: str) -> dict:
            result: dict = {}
            result_page: dict = {}
            for part_name in sorted(names):
                if not (part_name.startswith(f"word/{prefix}") and part_name.endswith(".xml")):
                    continue
                part_root = ET.fromstring(z.read(part_name))
                paras = []
                # EZI-0067：纯域段（无 w:t 文字但 run 链含 PAGE 域）旁路单列——
                # 页脚页码段的字号/样式藏在 pStyle 样式链（10 案实锤：模板
                # PAGE 段 pStyle=Footer sz18=9pt，旧逻辑空文字即跳过，页码段
                # 格式从未进池 → 产物页码段落 docDefaults 10.5/12/14）。
                # 旁路进 result_page，不进 paras：ir.footers 主池与其消费面
                # （_part_common_rule 等）零扰动。
                page_paras = []
                for pel in part_root.iter(f"{W}p"):
                    texts = "".join(t.text or "" for t in pel.iter(f"{W}t"))
                    if not texts.strip():
                        if prefix == "footer" and "{PAGE}" in _hf_run_sig(pel):
                            page_paras.append(pel)
                        continue
                    hppr = pel.find(f"{W}pPr")
                    first_run = pel.find(f"{W}r")
                    hrpr = first_run.find(f"{W}rPr") if first_run is not None else None
                    hstyle = hppr.find(f"{W}pStyle") if hppr is not None else None
                    hfp = _merge_format(_parse_ppr_format(hppr), _parse_rpr_format(hrpr))
                    # 页脚「第 {PAGE} 页」这类段落，**首个 run 常常没有 sz/rFonts**
                    # （是空白或纯文本），真正的字号写在后面的 run 上。只取首 run 会让
                    # fp.font_size_pt=None，退化到样式链/docDefaults，把模板的 10pt
                    # 读成 9pt（EZI-0006 实锤）。逐 run 补齐缺口，先到先得。
                    for _r in pel.iter(f"{W}r"):
                        if _r is first_run:
                            continue
                        _rpr = _r.find(f"{W}rPr")
                        if _rpr is None:
                            continue
                        for _k, _v in _parse_rpr_format(_rpr).__dict__.items():
                            if _v is None or _k.endswith("_extra_xml"):
                                continue
                            if getattr(hfp, _k, None) is None:
                                setattr(hfp, _k, _v)
                    paras.append({
                        "text": texts.strip(),
                        "style_id": hstyle.get(f"{W}val") if hstyle is not None else None,
                        "fp": hfp,
                        # run 级机制签名（2026-08-23 CJC 实锤）：TEXT/{PAGE}/TAB
                        # ——页眉页脚排版时按模板把目标「死页码/空白分隔」桥接为
                        # 域/制表位（机制层布局，绝不动中文文本 run）。
                        "run_sig": _hf_run_sig(pel),
                        "tab_segments": _hf_tab_segments(pel),
                    })
                if paras:
                    result[part_name] = paras
                if page_paras:
                    # 页码段 fmt 构建与文字段同构（逐 run 补缺字号），
                    # text 置空 + page_only 标志供下游路由。
                    pplist = []
                    for ppel in page_paras:
                        pppr = ppel.find(f"{W}pPr")
                        pfirst = ppel.find(f"{W}r")
                        prpr = pfirst.find(f"{W}rPr") if pfirst is not None else None
                        pstyle = pppr.find(f"{W}pStyle") if pppr is not None else None
                        pfp = _merge_format(_parse_ppr_format(pppr), _parse_rpr_format(prpr))
                        for _r in ppel.iter(f"{W}r"):
                            _rpr = _r.find(f"{W}rPr")
                            if _rpr is None:
                                continue
                            for _k, _v in _parse_rpr_format(_rpr).__dict__.items():
                                if _v is None or _k.endswith("_extra_xml"):
                                    continue
                                if getattr(pfp, _k, None) is None:
                                    setattr(pfp, _k, _v)
                        pplist.append({
                            "text": "", "page_only": True,
                            "style_id": pstyle.get(f"{W}val") if pstyle is not None else None,
                            "fp": pfp, "run_sig": _hf_run_sig(ppel),
                            "tab_segments": _hf_tab_segments(ppel),
                        })
                    result_page[part_name] = pplist
            return result, result_page

        ir.headers, _hdr_page = _part_formats("header")
        ir.footers, _ftr_page = _part_formats("footer")
        # EZI-0067：纯域页码段旁路（页眉页码由 _mech bridge 管，只挂 footer）
        ir.footers_page = _ftr_page

    root = ET.fromstring(document_xml)
    body = root.find(f"{W}body")
    if body is None:
        raise ValueError("missing w:body in document.xml")

    ir.metadata = {
        "part_name": DOCUMENT_PART,
    }
    # 内容 hash 在 paragraphs 填充完成后计算（_normalized_content_hash 依赖段落文本）
    ir.metadata["content_hash"] = _normalized_content_hash(ir.paragraphs)

    style_name_by_id = {s.style_id: s.name for s in ir.styles}

    # 文档级层级事实（EZI-0023）：一级层级是否已被「第X章」占据。stdlib ET 元素
    # 没有父指针，动作层能靠 `getroottree()` 反查，这里必须显式算一次再往下传。
    chapter_tier = _chapter_tier_for_root(root)
    ir.metadata["chapter_tier"] = chapter_tier

    # 节：段落内嵌 sectPr（按出现顺序）在前，body 尾部 sectPr 最后；统一编号
    embedded: list[SectionNode] = []
    for para_elem in body.iter(f"{W}p"):
        ppr = para_elem.find(f"{W}pPr")
        if ppr is not None:
            inner = ppr.find(f"{W}sectPr")
            if inner is not None:
                embedded.append(_parse_sectpr(inner, 0, rel_map))
    body_sects = [_parse_sectpr(sp, 0, rel_map) for sp in body.findall(f"{W}sectPr")]
    all_sects = embedded + body_sects
    for i, s in enumerate(all_sects):
        s.index = i
        s.node_id = f"section:{i}"
    ir.sections = all_sects

    # 遍历 body 块
    table_counter = 0
    figure_counter = 0
    field_tracker = _FieldTracker()
    body_blocks: list = []
    # TOC 域结果区状态机（2026-09-21 P3 实锤）：TOC 控件（sdt）展开后目
    # 录条目若失去 hyperlink，classify 会把「摘  要II」判成 abstract 漏进
    # content hash。TOC 域（instrText 含 TOC，非 PAGEREF）separate..end 之间
    # 的段落一律是派生目录条目，角色强制 toc_entry（机制层，与学校无关）。
    toc_result = False
    toc_stack: list[bool] = []  # 域栈：该域是否为 TOC 域
    toc_entry_indexes: list[int] = []
    pageref_toc_paras: set[int] = set()  # 含 PAGEREF _Toc 域的段（目录条目证据）

    for child in body:
        tag = _local(child.tag)
        if tag == "p":
            _para_idx = len(ir.paragraphs)
            _entry_before = toc_result
            # 先推进本段内的域状态（begin/separate/end/instrText）
            for run_el in child.iter(f"{W}r"):
                fc = run_el.find(f"{W}fldChar")
                if fc is not None:
                    ftype = fc.get(f"{W}fldCharType")
                    if ftype == "begin":
                        toc_stack.append(False)
                    elif ftype == "separate":
                        if toc_stack and toc_stack[-1]:
                            toc_result = True
                    elif ftype == "end":
                        if toc_stack:
                            if toc_stack.pop() and toc_result:
                                toc_result = False
                    continue
                it = run_el.find(f"{W}instrText")
                if it is not None and toc_stack:
                    txt = (it.text or "").strip().upper()
                    # TOC 域指令必须**整词**识别（2026-09-25 张鑫佳实锤）：
                    # 条目内的超链接域指令「HYPERLINK \l _Toc12345」也含
                    # "TOC" 子串——子串匹配把超链接域误标为 TOC 域，该域
                    # 的 end 又把 toc_result 关掉，使每个条目段在边界恰好
                    # 落回 False，24 条目录全部漏标 → 页码刷新顶翻 content
                    # hash → 交付被误拦。整词正则 ^TOC(\s|$) 只认真正的
                    # TOC 域指令。
                    if re.match(r"^TOC(?:\s|$)", txt) and "PAGEREF" not in txt:
                        toc_stack[-1] = True
                    elif _TOC_ANCHOR_FIELD_RE.match(txt):
                        # 目录锚点域（PAGEREF _Toc… 页码 / HYPERLINK \l _Toc…
                        # 条目跳转）：Word 目录条目特征，正文交叉引用极少命中
                        # → 记段，供「展开后无域结构」条目做区域法二次识别。
                        pageref_toc_paras.add(_para_idx)
            # fldSimple 形态的域（w:instr 属性，无 instrText run）一并取证
            for _fs in child.iter(f"{W}fldSimple"):
                _fsi = (_fs.get(f"{W}instr") or "").strip()
                if _fsi and _TOC_ANCHOR_FIELD_RE.match(_fsi):
                    pageref_toc_paras.add(_para_idx)
            # 段首已在目录结果区、或段内才进入结果区（TOC 域 separate 与
            # 首个条目同段）→ 本段属派生目录条目。判定放在段末，覆盖
            # 「首条与域结构同段」的漏网情形。
            if _entry_before or toc_result:
                toc_entry_indexes.append(_para_idx)
            para, figures = _parse_paragraph(
                child, None, style_name_by_id, field_tracker, figure_counter,
                chapter_tier=chapter_tier,
            )
            figure_counter += len(figures)
            ir.figures.extend(figures)
            ir.paragraphs.append(para)
            body_blocks.append(para)
        elif tag == "tbl":
            table = _parse_table(child, table_counter)
            table_counter += 1
            ir.tables.append(table)
            body_blocks.append(table)
        elif tag == "sectPr":
            continue

    ir.fields = field_tracker.finish()
    ir.blocks = body_blocks
    assign_front_roles(ir.paragraphs, style_name_by_id)
    assign_back_roles(ir.paragraphs)
    # TOC 域结果区段落强制 toc_entry（2026-09-21 P3）：在 front/back 角色
    # 分配之后覆写——否则「摘  要II」会被 assign_front_roles 判成 abstract
    # 而绕过 hash 的 toc_entry 豁免，展开后的目录条目混入正文文本。
    for _idx in toc_entry_indexes:
        if _idx < len(ir.paragraphs):
            ir.paragraphs[_idx].semantic_role = "toc_entry"
    # 目录区二次识别（2026-09-25 张鑫佳实锤第二轮）：unfold_toc_field 拆掉
    # TOC 外层域后，条目段只剩条目内的 HYPERLINK/PAGEREF 域对（无外层域、
    # 无 w:hyperlink 元素），域状态机与 classify 均判不出 toc_entry →
    # 条目文本混入 content hash，编译刷新页码后「内容保全」误拦交付。
    # 区域法补判：目录标题段起、第一个正文章节段止。
    for _idx in _toc_region_entry_indexes(ir.paragraphs, pageref_toc_paras):
        ir.paragraphs[_idx].semantic_role = "toc_entry"
    # 规范化内容 hash：依赖已填充的 paragraphs（跳过分页控制空白段，§10.1）
    ir.metadata["content_hash"] = _normalized_content_hash(ir.paragraphs)
    ir.diagnostics = [
        {"level": "warning", "code": "BROKEN_FIELD", "node_id": f.node_id, "detail": "unclosed fldChar begin"}
        for f in ir.fields if not f.complete
    ]
    return ir


_TOC_TITLE_CORES = {"目录", "目次", "CONTENTS", "TABLEOFCONTENTS"}
# 目录锚点域证据：PAGEREF _Toc…（页码）或 HYPERLINK \l _Toc…（条目跳转），
# 锚点名以 _Toc 开头即 Word 目录专属命名。
_TOC_ANCHOR_FIELD_RE = re.compile(
    r"^(?:PAGEREF|HYPERLINK)\s+(?:\\L\s+)?[\"']?_TOC", re.IGNORECASE)
_TOC_REGION_STOP_ROLES = {
    "abstract", "abstract_en", "keywords", "keywords_en", "classification",
    "refs_heading", "ack_heading", "bio_heading", "appendix_heading",
}


def _toc_region_entry_indexes(paragraphs: list, pageref_toc_paras: set) -> list[int]:
    """目录区条目索引：区域 + 结构证据/条目形态（机制层，2026-09-25）。

    适用 TOC 域被静态化（unfold_toc_field）后的条目形态——此时条目仅剩
    条目内的 HYPERLINK/PAGEREF 域对（首条甚至丢域），域状态机与
    classify_paragraph_role 都判不出 toc_entry。判定（从目录标题段起）：
    ① 带锚点域证据（PAGEREF/HYPERLINK _Toc…）的段 → 条目；
    ② 空段跳过；③ 标题/摘要/关键词/参考文献等章节 role → 终止区域；
    ④ 其余实文本段按条目形态兜底（短行 ≤80 字或含制表位 → 条目），
    长段落视作正文一并终止区域（防吞正文）。
    命中集（不含目录标题段，标题保留自身角色）属派生目录条目。
    """
    out: list[int] = []
    n = len(paragraphs)
    for i, p in enumerate(paragraphs):
        core = re.sub(r"[\s\u3000]+", "", (getattr(p, "text", "") or "")).upper()
        if core not in _TOC_TITLE_CORES:
            continue
        # 目录标题段保留原角色（heading1/body）：两侧分析口径一致，其文本
        # 计入 content hash 不产生漂移；只有条目段改判 toc_entry
        # （2026-09-25 test_chapter_semantics 回归：标题必须仍是标题）。
        j = i + 1
        lenient = 0     # 首条丢域兜底配额（仅紧邻目录标题的最多 1 行）
        while j < n:
            q = paragraphs[j]
            role = getattr(q, "semantic_role", None) or ""
            text = getattr(q, "text", "") or ""
            if j in pageref_toc_paras:
                # 结构证据优先于角色分类：「关键词…1」「摘要…1」这类条目文本
                # 会被 classify 误判成 keywords/abstract 角色，但锚点域/超链接
                # 证据在场即属派生目录条目。
                out.append(j)
                j += 1
                continue
            if not text.strip():
                j += 1
                continue
            if str(role).startswith("heading") or role in _TOC_REGION_STOP_ROLES:
                break
            # 无域证据的条目（如首条被静态化时丢域）：按条目形态兜底——
            # 页码尾/点 leader/制表位/宽空白+页码（条目「标题……页码」特征）。
            # 收紧（2026-09-25）：必须先是「非章节标题形态」；章节序号标题、带
            # 句读的句子都不再被吞（报告目录后紧跟同形短标题+短正文，此前
            # 「短行 ≤80」判据把真标题与正文一并判成 toc_entry）。
            if _is_cn_numbered_heading(text) or _is_numeric_heading(text) \
                    or _is_chapter_title_text(text):
                break
            if (_TOC_ENTRY_RE.search(text) or "\t" in text
                    or re.search(r"[\s\u3000]{2,}\d{1,4}\s*$", text)):
                out.append(j)
                j += 1
                lenient = 0
                continue
            _flat = re.sub(r"\s+", "", text)
            # 首条丢域兜底：紧邻目录标题的最多 1 行短行（无句读、无标题形态）
            if lenient < 1 and len(_flat) <= 40 \
                    and not re.search(r"[。；，、：！？]$", _flat):
                out.append(j)
                j += 1
                lenient += 1
                continue
            break
    return out


# ── EZI-0071：格式说明段谓词（三消费同源）──────────────────────────
# 真机形态库（3-2-1 源 18 条实锤）：「字体：黑体；字号：三号；间距：1.5倍
# 行距，段前12磅」「标题"摘要"：黑体，三号，居中」「存档论文打印装订…书脊
# 文字从上向下为」「下划线所填信息：（1）字体：楷体GB2312…」——冒号/分号
# 分隔的格式规格串。真文对照片（80 段）零命中校准。
_FORMAT_SPEC_TERMS = (
    "字体", "字号", "间距", "行距", "黑体", "宋体", "楷体", "仿宋",
    "Times New Roman", "居中", "左对齐", "右对齐", "段前", "段后",
    "磅", "页边距", "页眉", "页脚", "首行缩进", "缩进", "粗体", "加粗",
    "下划线", "倍行距", "五号", "小四", "三号", "四号", "二号", "小五",
    # EZI-0071 校准补充（3-2-1 十八条实锤的漏网形态）：
    "标题", "装订", "书脊", "分类号", "UDC", "小一", "黑体，", "缩进2字符",
)
_FORMAT_SPEC_LABELS = (
    "字体", "字号", "间距", "标题", "页眉", "页脚", "英文题目", "下划线",
    "正文", "封面", "存档", "装订", "中图分类号", "书脊", "表", "图",
    "关键词", "英文摘要", "具体要求", "填写",
)


def is_format_instruction_text(text: str) -> bool:
    """格式说明段谓词（EZI-0071）——引擎剥除/内容hash豁免/判定层正文池
    排除**同一谓词**（同源铁律）。

    形态：冒号/分号分隔的格式规格串（「字体：黑体；字号：三号；间距：…」）。
    保守闸（误删正文=不可挽回，宁可漏删）：
      ①长度 20~500（太短信息不足、太长多半真文）；
      ②格式术语命中 ≥3（字体/字号/行距/段前/磅/号/黑体…）；
      ③结构闸：前 36 字内出现说明标签（字体/标题/页眉/存档…）或前 30 字
        内有冒号；
      ④密度闸：每 40 字术语命中数 ≥1.2（真文提及一两个格式词不命中）。
    """
    raw = text or ""
    s = re.sub(r"\s+", "", raw)
    n = len(s)
    if n < 20 or n > 800:
        return False
    # 术语在**原文**上数（Times New Roman 带空格，剥空白后永不命中——
    # 校准实锤的漏网根因）；中文术语不受空白影响。
    hits = sum(1 for k in _FORMAT_SPEC_TERMS if k in raw)
    if hits < 3:
        return False
    lead = s[:36]
    labeled = any(lead.startswith(x) for x in _FORMAT_SPEC_LABELS)
    coloned = "：" in s[:30] or ":" in s[:30]
    if not (labeled or coloned):
        return False
    # 长段分支（校准实锤）：说明行拼接成的巨段（432/1134 字实测）密度被
    # 巨量文字稀释（0.39~0.65），但 hits≥6 + 冒号结构仍是强说明信号；
    # 跨语料真文负检（矿大/人社部/汕尾 287 长段）hits 上限 0 → 零误判。
    if n > 300 and hits >= 6 and coloned:
        return True
    density = hits * 40.0 / n
    return density >= 0.9


def _normalized_content_hash(paragraphs: list) -> str:
    """规范化内容 hash：所有有意义段落的规范化文本按序拼接。

    分页控制空白段（规范化后为空）不计入——打印版 finalizer 合法插入
    的空白页段落不构成内容变化。段落边界也不进入文字 hash：首字
    下沉会把同一段拆成首字段+续段，字符序列不变时仍属于格式修复。
    段落/节结构由 IR 和 Package 门禁独立验证。TOC 条目是正文标题与
    PAGEREF 的派生缓存，不属于作者正文，刷新后不进入内容 hash。TOC
    展开后，紧邻 toc_entry 块之前的标题同样属于派生结构。参考
    文献区的前导 [N] 是列表格式，编号变化不算正文改写；条目文字仍计入。
    """
    parts = []
    in_references = False
    for index, p in enumerate(paragraphs):
        role = getattr(p, "semantic_role", None)
        if role == "toc_entry":
            continue
        # EZI-0065 A4：整段【…】/*…* 包裹的模板说明段不是作者正文——
        # 与 actions.strip_instruction_paragraphs 同一谓词（源侧不计入、
        # 产物侧已剥除，两侧口径一致交付门才不误拦；真机案例：源里粘贴的
        # 「【表说明采用黑体；表头文字采用中文】」）。
        _s = re.sub(r"\s+", "", normalize_text(p.text) or "")
        if (( _s.startswith("【") and _s.endswith("】") and _s.count("【") == 1
                and _s.count("】") == 1 and len(_s) <= 120)
                or (_s.startswith("*") and _s.endswith("*") and _s.count("*") == 2
                    and 4 <= len(_s) <= 160)
                # EZI-0071：格式说明段（与 strip 第四闸同一谓词）
                or is_format_instruction_text(p.text or "")):
            continue
        # 目录条目（派生缓存）豁免：unfold_toc_field 展开后条目可能失去
        # hyperlink，若文本以「摘要」开头会被 classify 成 abstract 而绕过
        # toc_entry 跳过（2026-08-22 单栏文档「摘要III」误计入）。TOC1-3
        # 是 Word 内置目录样式，正文段落不会使用——按样式无条件豁免。
        sid = getattr(p, "style_id", None) or ""
        if str(sid).upper() in ("TOC1", "TOC2", "TOC3"):
            continue
        if str(role or "").startswith("heading"):
            next_role = None
            for following in paragraphs[index + 1:]:
                if normalize_text(getattr(following, "text", "")):
                    next_role = getattr(following, "semantic_role", None)
                    break
            if next_role == "toc_entry":
                continue
        norm = normalize_text(p.text)
        if role == "refs_heading":
            in_references = True
        elif in_references and (
            role in ("ack_heading", "bio_heading", "appendix_heading")
            or str(role or "").startswith("heading")
        ):
            in_references = False
        if in_references and role != "refs_heading":
            norm = re.sub(r"^\[\d+\]", "", norm)
        # 摘要/关键词/分类号标签：模板对齐会把「摘要：」规范为「摘  要」
        # （2026-08-21 用户裁决：用户笔误的空格/标点向模板对齐）。标签
        # 核心（去空格去冒号）一致即视为未改写，正文内容仍严格计入。
        if role in ("abstract", "abstract_en", "keywords", "keywords_en",
                    "classification"):
            m = re.match(
                r"^((?:摘\s*要|关\s*键\s*词|中图法?分类号|Abstract|Key\s*words)\s*[：:]?)",
                norm,
            )
            if m:
                lab = re.sub(r"[\s:：]+", "", m.group(1))
                # 「中图法分类号/中图分类号」为同一标签（漏/多一个"法"字
                # 属用户笔误，对齐模板修补不影响 hash）
                lab = lab.replace("中图法分类号", "中图分类号")
                norm = lab + norm[m.end():]
        if norm:
            parts.append(norm)
    return hashlib.sha1("".join(parts).encode("utf-8")).hexdigest()


def _parse_paragraph(
    para_elem, cell_id: Optional[str], style_name_by_id: dict, field_tracker: _FieldTracker,
    figure_counter: int, chapter_tier: bool = False,
) -> tuple[ParagraphNode, list[FigureNode]]:
    ppr = para_elem.find(f"{W}pPr")
    pstyle_id = None
    numpr = None
    if ppr is not None:
        pstyle_el = ppr.find(f"{W}pStyle")
        if pstyle_el is not None:
            pstyle_id = pstyle_el.get(f"{W}val")
        num_el = ppr.find(f"{W}numPr")
        if num_el is not None:
            numpr = {
                "num_id": _twips(num_el.find(f"{W}numId").get(f"{W}val")) if num_el.find(f"{W}numId") is not None else None,
                "ilvl": _twips(num_el.find(f"{W}ilvl").get(f"{W}val")) if num_el.find(f"{W}ilvl") is not None else 0,
            }

    text = _text_of(para_elem)
    norm = normalize_text(text)
    pstyle_name = style_name_by_id.get(pstyle_id) if pstyle_id else None
    hlevel = _trusted_heading_level(ppr, pstyle_name or "", text, chapter_tier)

    figures: list[FigureNode] = []
    has_drawing = False
    for drawing in para_elem.iter(f"{W}drawing"):
        info = _parse_drawing(drawing, {})
        if info:
            has_drawing = True
            figures.append(
                FigureNode(
                    node_id=build_node_id(DOCUMENT_PART, "figure", "figure", norm,
                                          None, f"media:{info.get('embed', '')}"),
                    media_rel=info.get("embed"),
                    docpr_id=info.get("docpr_id"),
                    extent=info.get("extent"),
                    wrap=info.get("wrap"),
                    index=figure_counter + len(figures),
                    content_hash=hashlib.sha1(
                        f"{info.get('embed')}|{info.get('extent')}".encode()
                    ).hexdigest()[:16],
                )
            )
    # OLE / VML 预览（Visio.Drawing 等）也是图段——漏判会被 body 规则
    # 打上 line=460 exact，把图裁成一条（2026-08-19 用户截图图4.2）。
    if not has_drawing:
        if (para_elem.find(f".//{W}object") is not None
                or para_elem.find(f".//{W}pict") is not None):
            has_drawing = True

    in_hyperlink = para_elem.find(f"{W}hyperlink") is not None
    has_omath_para = para_elem.find(f".//{M}oMathPara") is not None
    has_omath = para_elem.find(f".//{M}oMath") is not None
    has_equation_ole = _has_equation_ole(para_elem)
    role = classify_paragraph_role(ppr, pstyle_name, has_drawing,
                                   cell_id is not None, numpr is not None,
                                   text, in_hyperlink=in_hyperlink,
                                   has_omath_para=has_omath_para,
                                   has_omath=has_omath,
                                   has_equation_ole=has_equation_ole,
                                   chapter_tier=chapter_tier)
    parent_id = cell_id
    extra_fp = f"style:{pstyle_id}" if pstyle_id else None
    # run 直接格式：段落级 rPr（罕见）或第一个 run 的 rPr（常规写法）
    _rpr_el = para_elem.find(f"{W}rPr")
    if _rpr_el is None:
        _first_run = para_elem.find(f"{W}r")
        if _first_run is not None:
            _rpr_el = _first_run.find(f"{W}rPr")
    # 关键词行特判（2026-08-16）：模板示例「关键词：」标签 run 黑体加粗、
    # 内容 run 常规——direct 取「内容 run」格式（标签格式由 writer 特判），
    # 否则提取出的规则会把整段刷成黑体
    if role in ("keywords", "keywords_en", "abstract", "abstract_en", "classification"):
        for _r in para_elem.findall(f"{W}r"):
            _rt = "".join(t.text or "" for t in _r.iter(f"{W}t"))
            if _is_front_label_run(_rt) or _rt.strip() in ("*",):
                continue
            _rpr_el = _r.find(f"{W}rPr")
            break
    elif hlevel is not None:
        # 标题内容 run 特判（2026-08-23 CJC v9 实锤）：模板二级标题数字 run
        # 「1.1 」b=加粗但标题文字 run 不加粗（注释只写「字体为5号黑体」）。
        # 首 run（数字）格式被当成整段属性 → 应用时标题文字被误加粗。
        # 取第一个内容 run（跳过数字前缀）为段格式来源，与关键词行同机制。
        for _r in para_elem.findall(f"{W}r"):
            _rt = "".join(t.text or "" for t in _r.iter(f"{W}t")).strip()
            if re.match(r"^\d+(?:[.．]\d+)*$", _rt):
                continue
            _rpr_el = _r.find(f"{W}rPr")
            break
    para = ParagraphNode(
        node_id=build_node_id(DOCUMENT_PART, "paragraph", role, norm, parent_id, extra_fp),
        text=text,
        text_hash=norm,
        semantic_role=role,
        style_id=pstyle_id,
        is_heading=hlevel is not None,
        heading_level=hlevel,
        list_binding=numpr,
        parent_id=parent_id,
        direct_format=_merge_format(_parse_ppr_format(ppr), _parse_rpr_format(_rpr_el)),
        content_hash=hashlib.sha1(norm.encode("utf-8")).hexdigest()[:16],
    )

    # runs
    for r in para_elem.findall(f"{W}r"):
        field_tracker.feed(r)
        t = r.find(f"{W}t")
        has_tab = r.find(f"{W}tab") is not None
        # tab-only run（<w:tab/> 无 <w:t>，如模板「关键词→内容」箭头）也必须
        # 进 runs，否则标签后 tab 检测/重建丢失（2026-08-22 用户裁决）。
        if t is None and not has_tab:
            continue
        rpr = r.find(f"{W}rPr")
        para.runs.append(
            RunNode(
                node_id=build_node_id(DOCUMENT_PART, "run", role, norm, para.node_id,
                                      f"text:{normalize_text(_text_of(r))[:32]}"),
                text=_text_of(r),
                rstyle_id=rpr.find(f"{W}rStyle").get(f"{W}val") if rpr is not None and rpr.find(f"{W}rStyle") is not None else None,
                direct_format=_parse_rpr_format(rpr),
                has_tab=has_tab,
                parent_id=para.node_id,
            )
        )
    return para, figures


def _parse_table(tbl_elem, index: int) -> TableNode:
    tblpr = tbl_elem.find(f"{W}tblPr")
    alignment = None
    if tblpr is not None:
        jcel = tblpr.find(f"{W}jc")
        if jcel is not None:
            alignment = jcel.get(f"{W}val")
        layout_el = tblpr.find(f"{W}tblLayout")
        layout = layout_el.get(f"{W}type") if layout_el is not None else None
    else:
        layout = None
    grid = tbl_elem.find(f"{W}tblGrid")
    widths = []
    if grid is not None:
        for gc in grid.findall(f"{W}gridCol"):
            w = _twips(gc.get(f"{W}w"))
            if w is not None:
                widths.append(w)

    rows = tbl_elem.findall(f"{W}tr")
    row_count = len(rows)
    col_count = len(widths)
    merged: list[dict] = []
    header_rows: list[int] = []
    for ri, tr in enumerate(rows):
        trpr = tr.find(f"{W}trPr")
        if trpr is not None and trpr.find(f"{W}tblHeader") is not None:
            header_rows.append(ri)
        for ci, tc in enumerate(tr.findall(f"{W}tc")):
            tcpr = tc.find(f"{W}tcPr")
            vmerge = None
            if tcpr is not None:
                vm = tcpr.find(f"{W}vMerge")
                if vm is not None:
                    vmerge = vm.get(f"{W}val") or "continue"
            if vmerge:
                merged.append({"row": ri, "col": ci, "type": vmerge})
            if col_count == 0:
                col_count = max(col_count, ci + 1)

    node_id = build_node_id(DOCUMENT_PART, "table", "table",
                            hashlib.sha1(f"tbl:{index}".encode()).hexdigest()[:16], None)
    return TableNode(
        node_id=node_id,
        index=index,
        row_count=row_count,
        col_count=col_count,
        grid_widths=widths,
        layout=layout,
        alignment=alignment,
        tblpr_extra_xml=extract_structure_bundle(tblpr, "tblPr"),
        merged_cells=merged,
        header_rows=header_rows,
        content_hash=hashlib.sha1(_text_of(tbl_elem).encode("utf-8")).hexdigest()[:16],
    )
