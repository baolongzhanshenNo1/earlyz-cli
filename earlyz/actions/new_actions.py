"""T11 — 3.2 新 P0 action：set_doc_defaults / set_paragraph_format（§7.1）。

这些 action 是 Format Compiler 新链路（planner 输出）的执行端：
- set_doc_defaults：修改 styles.xml docDefaults（默认中西文字体/字号）；
- set_paragraph_format：按 NodeSelector 定位段落并设置段落格式属性
  （jc/keep/spacing/ind/pageBreak），参数即 FormatProperties 字段名。
"""

from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import Optional
from lxml import etree
import math
import os
import re
import zipfile
import copy
import hashlib
from difflib import SequenceMatcher
from xml.etree import ElementTree as StdET

from earlyz.actions import table_layout
from earlyz.tools.ooxml_order import (
    PPR_SEQ, RPR_SEQ, TBLPR_SEQ, TRPR_SEQ, canonicalize_root,
)
from earlyz.ir.ooxml_format import (
    PPR_EXCLUDED, RPR_EXCLUDED, PPR_MODELED_ATTRS, RPR_MODELED_ATTRS,
    SECTPR_PORTABLE, TBLPR_PORTABLE,
    extract_extra_bundle, extract_structure_bundle, local_name,
    structure_bundle_contains,
)

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
_EMU_PER_TWIP = 635


def _part_names(path: Path) -> list[str]:
    """zip 内的部件名列表。"""
    with zipfile.ZipFile(path) as z:
        return z.namelist()


def _read_part(path: Path, part: str) -> etree._Element:
    with zipfile.ZipFile(path, "r") as z:
        return etree.fromstring(z.read(part))


def _write_part_zipped(zout, part_name: str, data: bytes) -> None:
    """向打开的 zip 写入 XML part（强制 DEFLATED——Word 要求 XML part
    压缩，STORED 的 XML part 打开报「文件可能已经损坏」）。"""
    zi = zipfile.ZipInfo(part_name)
    zi.compress_type = zipfile.ZIP_DEFLATED
    zout.writestr(zi, data)


def _write_part(path: Path, part: str, root: etree._Element) -> None:
    path = Path(path)
    # 收尾规范化（2026-09-28 用户判决「从源头解决」）：写盘前统一按 schema 顺序
    # 重排 pPr/rPr/tblPr/trPr 子元素、并保证 w:pPr 是 w:p 的第一个子元素。
    # 本函数是全仓唯一的 XML part 写盘函数 = 单一收口点 → 新增直写点也不会漏。
    # 顺序表真源 = earlyz/tools/ooxml_order.py（禁止再抄一份）。
    if part.endswith(".xml"):
        canonicalize_root(root)
    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(path, "a") as z:
        names = set(z.namelist())
        if part in names:
            # 需要重写：zipfile 无法原地替换 → 整体重建
            pass
    # 重建 zip（保留其他 entry）
    tmp = path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        written = False
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == part:
                data = new_data
                written = True
            zout.writestr(item, data)
        if not written:
            # 2026-10-17：新增 part（如新建页眉部件）不能静默丢
            _write_part_zipped(zout, part, new_data)
    tmp.replace(path)


def _snapshot_package(path: Path) -> dict[str, bytes]:
    """整包快照（Word COM Save 会重写整个 zip 的多个 part——
    rels/settings/Content_Types 全可能变；恢复 document.xml 单 part 会造成
    包内不一致 → Word「文件可能已经损坏」。整包级备份/恢复才安全。）"""
    with zipfile.ZipFile(path, "r") as z:
        return {i.filename: z.read(i.filename) for i in z.infolist()}


def _restore_package(path: Path, snapshot: dict[str, bytes]) -> None:
    """把整包恢复到快照状态（强制 DEFLATED——Word 要求 XML part 压缩）。"""
    path = Path(path)
    tmp = path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp, "w") as zout:
        for name, data in snapshot.items():
            _write_part_zipped(zout, name, data)
    tmp.replace(path)


def _collect_domain_results(pkg: dict[str, bytes]) -> dict[str, list[str]]:
    """从 Word Save 后的整包中提取各 part 的 PAGE 域结果（真实页码）。

    Word 导出 PDF 会重算所有域；其中 PAGE 域结果（页脚页码缓存）同步为
    真实页码是 Save 的本意（2026-08-23 单栏页码铁证）。整包恢复会把它们
    丢回旧值 → 按 part 收集，恢复后合并回来（只动 PAGE 域，REF 等其它
    域不收集，保持原文）。
    """
    out: dict[str, list[str]] = {}
    for name, data in pkg.items():
        if not (name.startswith("word/header") or name.startswith("word/footer")
                or name == "word/document.xml"):
            continue
        try:
            root = etree.fromstring(data)
        except Exception:  # noqa: BLE001
            continue
        texts: list[str] = []
        in_result = False
        is_page = False
        for el in root.iter():
            tag = el.tag.split("}")[-1]
            if tag == "fldChar":
                kind = el.get(f"{W}fldCharType")
                if kind == "begin":
                    is_page = False
                    in_result = False
                elif kind == "separate":
                    in_result = True
                elif kind == "end":
                    in_result = False
                    is_page = False
            elif tag == "instrText":
                txt = (el.text or "").upper()
                # PAGE 域（含 \* MERGEFORMAT 变体）；排除 PAGEREF
                if "PAGE" in txt and "PAGEREF" not in txt:
                    is_page = True
            elif tag == "t" and in_result and is_page:
                texts.append(el.text or "")
        if texts:
            out[name] = texts
    return out


def _apply_domain_results(path: Path, collected: dict[str, list[str]]) -> int:
    """把 PAGE 域结果写回对应 part（保持 REF 等其它域原文）。"""
    applied = 0
    for part, texts in collected.items():
        try:
            root = _read_part(path, part)
        except Exception:  # noqa: BLE001 — Word 增删 part 时跳过该 part
            continue
        idx = 0
        in_result = False
        is_page = False
        for el in root.iter():
            tag = el.tag.split("}")[-1]
            if tag == "fldChar":
                kind = el.get(f"{W}fldCharType")
                if kind == "begin":
                    is_page = False
                    in_result = False
                elif kind == "separate":
                    in_result = True
                elif kind == "end":
                    in_result = False
                    is_page = False
            elif tag == "instrText":
                txt = (el.text or "").upper()
                if "PAGE" in txt and "PAGEREF" not in txt:
                    is_page = True
            elif tag == "t" and in_result and is_page:
                if idx < len(texts) and (el.text or "") != texts[idx]:
                    el.text = texts[idx]
                    applied += 1
                idx += 1
        _write_part(path, part, root)
    return applied


def _text_of(elem) -> str:
    return "".join(t.text or "" for t in elem.iter(f"{W}t"))


def _compute_cover_boundary(body, styles_map=None) -> int:
    """raw XML 上复刻 decompiler.find_cover_end 语义（封面段计数）。

    返回封面区段落数（首个结构性锚点的 0-based 索引）：锚点=章节序号行/
    outlineLvl heading/语义标题词表（摘要/目录/声明…）/固定锚点词。
    锚点必须有文字：空段（即使带 heading 样式）不承载「封面结束」语义
    （2026-09-25 真机：文档首段是空段却残留 heading 1 样式 → 边界=0 →
    「无封面区」→ role 批量动作把封面大标题刷成正文缩进/字号）。
    与 lint engine 的 find_cover_end(ir) 逐条件对齐（2026-08-15 真机回归：
    writer 缺词表 heading 检查 → 边界 54 vs lint 44 → 声明/授权书段落被
    role 批量动作整段跳过 → 产物声明页无样式）。
    """
    from earlyz.spec.decompiler import (
        _CHAPTER_RE as _chr, _chapterish_text as _chz, _is_anchor as _ian,
        norm_text as _nt, is_sentence_like as _sent, is_chapter_anchor as _anchor_ch,
        _cover_like_text as _clt,
    )
    from earlyz.ir.analyzer import classify_paragraph_role
    style_names = styles_map or {}
    _compute_cover_boundary.count = 0
    # 封面形态上限（与 decompiler.find_cover_end 同源）：封面是短行堆叠，
    # 出现长正文段即封面上限，避免合同类文档把首个章锚点之前整段当封面跳过。
    _cap = 0
    for _c in body.iter(f"{W}p"):
        if any(a.tag == f"{W}drawing" for a in _c.iterancestors()):
            continue
        _ppr = _c.find(f"{W}pPr")
        _ps = _ppr.find(f"{W}pStyle") if _ppr is not None else None
        _psid = _ps.get(f"{W}val", "") if _ps is not None else ""
        _pname = style_names.get(_psid, "") or _psid
        _npr = _ppr is not None and _ppr.find(f"{W}numPr") is not None
        _role = classify_paragraph_role(_ppr, _pname, False, False, _npr, _text_of(_c))
        if not _clt(_text_of(_c), _role):
            break
        _cap += 1
    for child in body.iter(f"{W}p"):
        if any(a.tag == f"{W}drawing" for a in child.iterancestors()):
            continue
        text = _text_of(child)
        n = _nt(text)
        # 章节序号行判据与 decompiler.find_cover_end 同源：
        # 阿拉伯数字列表项（「2.分期付款，…：」）不是章标题，否则封面区
        # 会推进到合同/公文正文内部（2026-10-29 中介合同边界 42 实测）。
        if _anchor_ch(text):
            return min(_compute_cover_boundary.count, _cap)
        ppr = child.find(f"{W}pPr")
        outline = ppr.find(f"{W}outlineLvl") if ppr is not None else None
        if n and outline is not None and len(n) <= 24 and _chz(n):
            return min(_compute_cover_boundary.count, _cap)
        jc_el = ppr.find(f"{W}jc") if ppr is not None else None
        jc = jc_el.get(f"{W}val") if jc_el is not None else None
        if _ian(text, None, jc):
            return min(_compute_cover_boundary.count, _cap)
        # 语义标题词表（与 analyzer classify_paragraph_role 同源）：
        # 「学位论文原创性声明」等无 outlineLvl 的标题也是封面边界
        pstyle = ppr.find(f"{W}pStyle") if ppr is not None else None
        pstyle_id = pstyle.get(f"{W}val", "") if pstyle is not None else ""
        pstyle_name = style_names.get(pstyle_id, "") or pstyle_id
        numpr = ppr is not None and ppr.find(f"{W}numPr") is not None
        if (n and len(n) <= 24 and _chz(n) and not _sent(text)
                and classify_paragraph_role(ppr, pstyle_name, False, False, numpr, text).startswith("heading")):
            return min(_compute_cover_boundary.count, _cap)
        _compute_cover_boundary.count += 1
    return 0  # 无锚点 = 无封面区（role 规则全文生效）


def _local(tag: str) -> str:
    return tag.split("}")[-1]


def _role_of(ppr, has_drawing: bool, numpr: bool, pstyle_name: str = "") -> str:
    """段落语义角色（复用 analyzer 的单一真相源 classify_paragraph_role）。

    heading：outlineLvl 或 pStyle name 含 heading；caption：表/图/公式标题文本模式。
    """
    from earlyz.ir.analyzer import classify_paragraph_role
    return classify_paragraph_role(ppr, pstyle_name, has_drawing, False, numpr, "")


def _read_styles_map(path) -> dict[str, str]:
    """styleId → style name 映射（heading 判定依赖 name，如 styleId=1 → name="heading 1"）。"""
    style_names: dict[str, str] = {}
    try:
        styles_root = _read_part(Path(path), "word/styles.xml")
    except Exception:
        return style_names
    if styles_root is None or not styles_root.tag.endswith("styles"):
        return style_names
    for s in styles_root.findall(f"{W}style"):
        sid = s.get(f"{W}styleId")
        name = s.find(f"{W}name")
        if sid and name is not None:
            style_names[sid] = name.get(f"{W}val", "") or sid
    return style_names


def _locate_paragraphs(root: etree._Element, selector: dict,
                       styles_map: dict[str, str] | None = None) -> list[etree._Element]:
    """按 (semantic_role + occurrence / node_id / text_hash) 定位 body 级段落。

    selector: {semantic_role, occurrence, node_id, text_hash}
    node_id 格式：paragraph:{role}:{指纹}（analyzer 生成，§4.3 稳定定位）——
    匹配时复刻 build_node_id 计算比对（角色 + 规范化文本 + pStyle 指纹）。
    收集范围与 analyzer 对齐：body 直接子级 + sdt 容器；排除文本框(drawing 内)
    与表格内段落（role 判定含 in_table）。批量（无 occurrence）→ 返回该 role 全部段落。
    """
    from earlyz.ir.analyzer import classify_paragraph_role, FrontRoleAssigner, FRONT_ROLES
    from earlyz.ir.selectors import build_node_id, normalize_text

    style_names = styles_map or {}
    role = selector.get("semantic_role")
    text_hash = selector.get("text_hash")
    node_id = selector.get("node_id")
    occ = selector.get("occurrence")
    anchor_sel = selector.get("cover_anchor")
    matches: list[etree._Element] = []
    body = root.find(f"{W}body")
    if body is None:
        return matches
    # 段序号选择器（2026-11-08 用户实锤「为什么改不了」）：LLM 拿到「第 60 段」这种
    # 位置描述时只能传 semantic_role/node_id/text_hash，猜不出来就退化成空选择器 →
    # 命中全部 174 段 → AMBIGUOUS → 回滚，用户看到「属性动作执行异常」。
    # 索引按 body 直接子级段落计数（与 inspect_document 的段落编号同口径）。
    if selector.get("paragraph_index") is not None:
        try:
            _i = int(selector["paragraph_index"])
        except (TypeError, ValueError):
            return []
        _all = [x for x in body.iter(f"{W}p") if x.getparent() is body]
        return [_all[_i]] if 0 <= _i < len(_all) else []
    if anchor_sel is not None:
        # 封面锚点选择器（T3）：复用 decompiler 的匹配语义（lint/writer 同源）
        from earlyz.spec.decompiler import match_cover_paragraphs as _mcp
        from earlyz.ir.analyzer import analyze_document as _ad
        # anchor 匹配基于 IR；writer 侧以文本特征复刻（避免二次全量分析）：
        # 精确 norm / 字段前缀 / 日期 / 标题形状，与 decompiler 保持一致。
        import re as _re
        from earlyz.spec.decompiler import (
            _field_name, _cjk_ratio, _DATE_RE, _CHAPTER_RE, norm_text as _nt,
            _is_annotation, _is_label,
        )
        limit_hint = selector.get("cover_limit", 62)
        boundary = _compute_cover_boundary(body)
        count = 0
        for child in body.iter(f"{W}p"):
            count += 1
            if count > limit_hint:
                break
            t = _text_of(child).strip()
            if not t or _is_annotation(t) or _is_label(t):
                continue
            n = _nt(t)
            match_kind = selector.get("match")
            in_cover = count - 1 < boundary
            hit = False
            if match_kind == "title":
                from earlyz.spec.decompiler import _EXACT_ANCHOR_NORMS
                hit = (in_cover and _cjk_ratio(t) >= 0.5
                       and 8 <= len(_re.sub(r"[\s]", "", t)) <= 30
                       and not _field_name(t) and not _CHAPTER_RE.match(n)
                       and not _DATE_RE.match(t)
                       and n not in _EXACT_ANCHOR_NORMS)
            elif match_kind == "date":
                hit = in_cover and bool(_DATE_RE.match(t))
            elif match_kind == "prefix":
                from earlyz.spec.decompiler import _field_match as _pfm
                hit = bool(anchor_sel) and _pfm(_field_name(t), anchor_sel)
            else:
                hit = (n == selector.get("norm") or
                       (bool(anchor_sel) and _field_name(t) == anchor_sel) or
                       n == anchor_sel)
            if hit:
                matches.append(child)
        return matches
    # 遍历全部段落（含 sdt 等嵌套容器），排除文本框(drawing 内)段落
    from earlyz.spec.decompiler import (
        _field_name as _dfn, _cjk_ratio as _dcr, _DATE_RE as _ddr,
        _CHAPTER_RE as _dchr, norm_text as _dnt, _is_annotation as _dia,
        _is_label as _dil, find_cover_end,
    )
    # 封面区边界（与 lint 同语义）：role 批量动作跳过封面段（归 cover 规则）
    _cover_boundary = None  # lazily computed via paragraph counter
    count_so_far = 0
    raw: list[tuple] = []
    has_heading = False
    for child in body.iter(f"{W}p"):
        in_drawing = any(a.tag == f"{W}drawing" for a in child.iterancestors())
        if in_drawing:
            continue
        ppr = child.find(f"{W}pPr")
        has_drawing = (
            child.find(f".//{W}drawing") is not None
            or child.find(f".//{W}object") is not None
            or child.find(f".//{W}pict") is not None
        )
        numpr = ppr is not None and ppr.find(f"{W}numPr") is not None
        pstyle = ppr.find(f"{W}pStyle") if ppr is not None else None
        pstyle_id = pstyle.get(f"{W}val", "") if pstyle is not None else ""
        pstyle_name = style_names.get(pstyle_id, "") or pstyle_id
        in_table = any(a.tag == f"{W}tbl" for a in child.iterancestors())
        text = _text_of(child)
        count_so_far += 1
        in_hyperlink = child.find(f"{W}hyperlink") is not None
        has_omath_para = child.find(f".//{M}oMathPara") is not None
        has_omath = child.find(f".//{M}oMath") is not None
        from earlyz.ir.analyzer import _has_equation_ole
        has_equation_ole = _has_equation_ole(child)
        r = classify_paragraph_role(ppr, pstyle_name, has_drawing, in_table,
                                    numpr, text, in_hyperlink=in_hyperlink,
                                    has_omath_para=has_omath_para,
                                    has_omath=has_omath,
                                    has_equation_ole=has_equation_ole)
        if str(r).startswith("heading"):
            has_heading = True
        raw.append((child, r, text, pstyle_name, pstyle_id, count_so_far))
    front_box = FrontRoleAssigner() if has_heading else None
    # 后部上下文角色（2026-08-23 CJC v9）：与 analyze_document 的 assign_back_roles
    # 语义一致——appendix_heading 后非空 body/equation_display → appendix_body。
    # 缺此步则执行时 set_paragraph_format 按 role 定位不到（运行时重新分类）。
    from earlyz.ir.analyzer import assign_back_roles as _assign_back_roles

    class _SimPara:
        __slots__ = ("text", "semantic_role")

        def __init__(self, text, role):
            self.text = text
            self.semantic_role = role

    syms = [_SimPara(text, r) for _c, r, text, _pn, _pid, _cnt in raw]
    _assign_back_roles(syms)
    # 终态角色先整轮算完（FrontRoleAssigner 有状态，须按文档顺序逐段 feed 一次）
    final_roles: list = []
    for idx, (_c, _r, text, pstyle_name, _pid, _cnt) in enumerate(raw):
        r = syms[idx].semantic_role
        if front_box is not None:
            r = front_box.feed(text, pstyle_name, r)
        final_roles.append(r)
    from earlyz.ir.analyzer import heading_rule_applies, zone_is_article_head
    _article_head = False
    for idx, (child, _r0, text, pstyle_name, pstyle_id, count_so_far) in enumerate(raw):
        r = final_roles[idx]
        if role is not None and r != role:
            continue
        # 前置页标题 / 同名标题口径（与 lint 同一谓词，2026-09-29）
        if role is not None and not heading_rule_applies(selector, role, text):
            continue
        if role is not None and anchor_sel is None:
            if _cover_boundary is None:
                _cover_boundary = _compute_cover_boundary(body, style_names)
                # 与交付门同一谓词（2026-09-29）：只有期刊题名区的前部角色豁免封面跳过；
                # 学位论文封面里被判成前部角色的段（校名/题名）仍归 cover 规则管。
                _article_head = zone_is_article_head(
                    final_roles[i] for i, e in enumerate(raw)
                    if e[5] <= _cover_boundary and (e[2] or "").strip())
                # 2026-10-17 真机（期刊短文测试文档）：期刊首页的前部角色
                # （摘要/关键词/分类号）本来就在「第 1 章」之前，会被封面区
                # 一起跳过 → 摘要规则定位不到 → set_paragraph_format 抛错
                # → 整篇回滚。判据补一条：前部区里出现摘要/关键词类角色，
                # 就说明这是期刊题名区而不是学位论文封面。
                if not _article_head and any(
                        r in ("abstract", "abstract_en", "keywords",
                              "keywords_en", "classification")
                        for r in (final_roles[i] for i, e in enumerate(raw)
                                  if e[5] <= _cover_boundary)):
                    _article_head = True
            if count_so_far <= _cover_boundary and not (
                    r in FRONT_ROLES and _article_head):
                continue
        if node_id is not None:
            norm = normalize_text(text)
            extra_fp = f"style:{pstyle_id}" if pstyle_id else None
            computed = build_node_id("word/document.xml", "paragraph", r, norm, None, extra_fp)
            if computed != node_id:
                continue
        elif text_hash is not None:
            import hashlib
            norm = normalize_text(text)
            if hashlib.sha1(norm.encode("utf-8")).hexdigest()[:16] != text_hash:
                continue
        matches.append(child)
    if role is not None and occ is not None and occ < len(matches):
        return [matches[occ]]
    if role is not None:
        return matches  # 批量 role 语义（设计）：全部达标段落，不设门禁
    if anchor_sel is not None:
        return matches  # 封面锚点批量语义（T3 设计）：全部命中段落
    if occ is not None:
        # 精确模式（node_id/text_hash）+ occurrence：取第 N 次出现；越界 → 无目标
        return [matches[occ]] if occ < len(matches) else []
    # 精确模式（node_id / text_hash）且未指定 occurrence：多候选 = AMBIGUOUS（§4.4）
    # ——禁止静默选中，由上层决定（报错/索要 occurrence/人工确认）。
    if len(matches) > 1:
        from earlyz.ir.selectors import AmbiguousSelectorError
        raise AmbiguousSelectorError(selector, matches)
    return matches


# ────────────────────────────────────────────────────────────────
# set_doc_defaults
# ────────────────────────────────────────────────────────────────

def set_doc_defaults(target_path: Path, params: dict) -> dict:
    """修改 styles.xml docDefaults/rPrDefault/rPr。

    params: {font_east_asia, font_ascii, font_size_pt, ...}（FormatProperties 字段）
    健壮性：styles.xml 缺失时创建最小 styles 根（docx 规范允许缺失，但默认字体
    规则需要 docDefaults 载体）。
    """
    target_path = Path(target_path)
    with zipfile.ZipFile(target_path) as z:
        has_styles = "word/styles.xml" in z.namelist()
    if not has_styles:
        styles_root = etree.Element(
            f"{W}styles",
            nsmap={None: "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
        )
        doc_defaults_el = etree.SubElement(styles_root, f"{W}docDefaults")
        rpr_default = etree.SubElement(doc_defaults_el, f"{W}rPrDefault")
        rpr_el = etree.SubElement(rpr_default, f"{W}rPr")
        data = etree.tostring(styles_root, xml_declaration=True, encoding="UTF-8", standalone=True)
        with zipfile.ZipFile(target_path, "a") as z:
            _write_part_zipped(z, "word/styles.xml", data)
    root = _read_part(target_path, "word/styles.xml")
    doc_defaults = root.find(f"{W}docDefaults")
    if doc_defaults is None:
        doc_defaults = etree.SubElement(root, f"{W}docDefaults")
    rpr_default = doc_defaults.find(f"{W}rPrDefault")
    if rpr_default is None:
        rpr_default = etree.SubElement(doc_defaults, f"{W}rPrDefault")
    rpr = rpr_default.find(f"{W}rPr")
    if rpr is None:
        rpr = etree.SubElement(rpr_default, f"{W}rPr")

    fonts = rpr.find(f"{W}rFonts")
    if params.get("font_east_asia") or params.get("font_ascii"):
        if fonts is None:
            fonts = etree.SubElement(rpr, f"{W}rFonts")
        if params.get("font_east_asia"):
            fonts.set(f"{W}eastAsia", params["font_east_asia"])
        if params.get("font_ascii"):
            fonts.set(f"{W}ascii", params["font_ascii"])
            fonts.set(f"{W}hAnsi", params["font_ascii"])
    if params.get("font_size_pt") is not None:
        sz = rpr.find(f"{W}sz")
        if sz is None:
            sz = etree.SubElement(rpr, f"{W}sz")
        sz.set(f"{W}val", str(int(round(params["font_size_pt"] * 2))))
        szcs = rpr.find(f"{W}szCs")
        if szcs is None:
            szcs = etree.SubElement(rpr, f"{W}szCs")
        szcs.set(f"{W}val", str(int(round(params["font_size_pt"] * 2))))

    _write_part(target_path, "word/styles.xml", root)
    return {"changed": True}


def verify_doc_defaults(target_path: Path, params: dict) -> bool:
    """验证 docDefaults 已按期望设置。"""
    try:
        root = _read_part(target_path, "word/styles.xml")
        doc_defaults = root.find(f"{W}docDefaults")
        if doc_defaults is None:
            return False
        rpr = doc_defaults.find(f"{W}rPrDefault")
        if rpr is None:
            return False
        rpr = rpr.find(f"{W}rPr")
        if rpr is None:
            return False
        fonts = rpr.find(f"{W}rFonts")
        if params.get("font_east_asia"):
            if fonts is None or fonts.get(f"{W}eastAsia") != params["font_east_asia"]:
                return False
        if params.get("font_ascii"):
            if fonts is None or fonts.get(f"{W}ascii") != params["font_ascii"]:
                return False
        if params.get("font_size_pt") is not None:
            sz = rpr.find(f"{W}sz")
            if sz is None or sz.get(f"{W}val") != str(int(round(params["font_size_pt"] * 2))):
                return False
        return True
    except Exception:
        return False


# ────────────────────────────────────────────────────────────────
# set_paragraph_format
# ────────────────────────────────────────────────────────────────

_TOC_PAGEREF_RE = re.compile(r"PAGEREF\s+(\S+)")


def _para_pageref_target(p) -> str | None:
    """段落内 PAGEREF 域的目标书签名；无则 None（拼接全部 instrText 后解析）。

    2026-10-08：Word 更新域后会把域指令拆成多个 instrText（在非 ASCII 边界处，
    如 `PAGEREF _Tof` + `F0012 \\h`）。对序列化 XML 直接跑正则会捕到标签杂质，
    拆分的条目页码会静默漏同步——既有潜伏缺陷，M2 图/表目录依赖此链，一并修复。
    """
    joined = "".join(x.text or "" for x in p.iter(f"{W}instrText"))
    m = _TOC_PAGEREF_RE.search(joined)
    return m.group(1) if m else None

_PPR_MAP = {
    "jc": (f"{W}jc", "val"),
    "keep_next": (f"{W}keepNext", None),
    "keep_lines": (f"{W}keepLines", None),
    "page_break_before": (f"{W}pageBreakBefore", None),
    "widow_control": (f"{W}widowControl", None),
    "snap_to_grid": (f"{W}snapToGrid", None),
    "spacing_before": (f"{W}spacing", "before"),
    "spacing_after": (f"{W}spacing", "after"),
    "spacing_before_lines": (f"{W}spacing", "beforeLines"),
    "spacing_after_lines": (f"{W}spacing", "afterLines"),
    "line": (f"{W}spacing", "line"),
    "line_rule": (f"{W}spacing", "lineRule"),
    "line_spacing": (f"{W}spacing", "__line_spacing__"),  # 倍数行距：写 line + lineRule=auto
    "ind_first_line": (f"{W}ind", "firstLine"),
    "ind_first_line_chars": (f"{W}ind", "firstLineChars"),
    "first_line_indent_chars": (f"{W}ind", "firstLineChars"),  # spec 惯例别名（2 字符 = 200）
    "outline_lvl": (f"{W}outlineLvl", "val"),
    "ind_left": (f"{W}ind", "left"),
    "ind_hanging": (f"{W}ind", "hanging"),
    "ind_right": (f"{W}ind", "right"),
}


# run 直接格式属性（标题加粗/字体等——写入段落内所有 run 的 rPr）
_RPR_MAP = {
    "bold": (f"{W}b", None),
    "italic": (f"{W}i", None),
    "small_caps": (f"{W}smallCaps", None),
    "underline": (f"{W}u", "val"),
    "color": (f"{W}color", "val"),
    "font_ascii": (f"{W}rFonts", "ascii"),
    "font_east_asia": (f"{W}rFonts", "eastAsia"),
    "font_size_pt": (f"{W}sz", "val"),
}


def _set_run_properties(p, properties: dict, target_runs: Optional[list] = None) -> None:
    """Write run properties to an explicit run list or all direct runs."""
    runs = target_runs if target_runs is not None else p.findall(f"{W}r")
    for prop, value in properties.items():
        if value is None or prop not in _RPR_MAP:
            continue
        tag, attr = _RPR_MAP[prop]
        for r in runs:
            if any(a.tag == f"{M}oMath" or a.tag == f"{M}oMathPara"
                   or a.tag.endswith("object") or a.tag.endswith("OLEObject")
                   for a in r.iterancestors()):
                continue
            rpr = r.find(f"{W}rPr")
            if rpr is None:
                rpr = etree.Element(f"{W}rPr")
                r.insert(0, rpr)
            if attr is None:
                el = rpr.find(tag)
                if value is True:
                    if el is None:
                        el = etree.Element(tag)
                        _rpr_insert_ordered(rpr, el)
                    el.attrib.pop(f"{W}val", None)
                elif value is False:
                    if el is None:
                        el = etree.Element(tag)
                        _rpr_insert_ordered(rpr, el)
                    el.set(f"{W}val", "0")
            else:
                el = rpr.find(tag)
                if el is None:
                    el = etree.Element(tag)
                    _rpr_insert_ordered(rpr, el)
                if prop == "font_size_pt":
                    el.set(f"{W}val", str(int(round(float(value) * 2))))  # 磅 → 半磅
                else:
                    el.set(f"{W}{attr}", str(value))
                    # 2026-08-23 单栏字体实锤：WPS 源稿 rFonts 带主题字体
                    # 属性（eastAsiaTheme=minorEastAsia → 等线）——Word 解析
                    # 时主题字体优先于直接字体 → 宋体写不生效显示「等线
                    # (中文正文)」。写显式字体必须清除同族主题属性。
                    if prop in ("font_ascii", "font_east_asia"):
                        for tk in ("asciiTheme", "hAnsiTheme",
                                   "eastAsiaTheme", "cstheme"):
                            el.attrib.pop(f"{W}{tk}", None)


def _split_label_run(r, Wns) -> etree._Element:
    """把「关键词：/摘  要/Abstract」标签前缀拆成独立 run。"""
    texts = list(r.iter(f"{W}t"))
    full = "".join(t.text or "" for t in texts)
    m = re.match(
        r"^((?:摘\s*要|摘要|关键词|Abstract|Key(?:\s+words?|words)|Index\s*Terms?|中图法分类号|中图分类号)\s*[:：]?)",
        full, re.IGNORECASE,
    )
    if not m:
        return None
    label = m.group(1)
    if len(full) <= len(label):
        return None
    first_t = texts[0]
    txt = first_t.text or ""
    if txt.startswith(label):
        first_t.text = txt[len(label):]
        new_r = etree.Element(f"{W}r")
        for child in r:
            if child.tag in (f"{W}rPr", f"{W}rsidR", f"{W}rsidRPr"):
                new_r.append(copy.deepcopy(child))
        new_t = etree.SubElement(new_r, f"{W}t")
        new_t.text = label
        xml_space = "{http://www.w3.org/XML/1998/namespace}space"
        if label.endswith(" ") or label.startswith(" "):
            new_t.set(xml_space, "preserve")
        return new_r
    return None


def _label_core_close(a: str, b: str) -> bool:
    """标签核心近似：字符编辑距离 ≤1（插入/删除/替换单字符）。

    「中图分类号」↔「中图法分类号」（漏/多一个"法"字）视为同一标签；
    跨语言（摘要 vs Abstract）或实质不同词距离远大，永不误换。
    """
    a, b = a.casefold(), b.casefold()
    if not a or not b:
        return False
    if a == b:
        return True
    if len(a) < 2 or len(b) < 2:
        return False
    if abs(len(a) - len(b)) > 1:
        return False
    if len(a) == len(b):
        return sum(1 for x, y in zip(a, b) if x != y) <= 1
    if len(b) > len(a):
        a, b = b, a
    for i in range(len(a)):
        if a[:i] + a[i + 1:] == b:
            return True
    return False


def _is_blank_run(r) -> bool:
    """run 是否纯空白分隔 run（w:t 全空白、无 drawing/域/特殊内容）。"""
    for tag in (f"{W}drawing", f"{W}object", f"{W}tab", f"{W}br",
                f"{W}fldChar", f"{W}instrText", f"{W}noBreakHyphen"):
        if r.find(tag) is not None:
            return False
    texts = [t.text or "" for t in r.iter(f"{W}t")]
    if not texts:
        return False  # 只含 rPr 的空 run 不是分隔空白
    return all(t.strip() == "" for t in texts)


def _set_keywords_run_properties(p, properties: dict) -> None:
    """摘要/关键词：内容 run 用普通 properties；标签 run 用 label_*（来自当前模板）。"""
    runs = p.findall(f"{W}r")
    if not runs:
        return
    label_pat = re.compile(
        r"^(摘\s*要|摘要|关键词|Abstract|Key(?:\s+words?|words)|Index\s*Terms?|中图法分类号|中图分类号)",
        re.I,
    )
    # 「标签只有标签」的严格版：必须整段就是标签（可带冒号）。
    # 2026-10-26 实证：宽松版把内容 run「keywords1; keywords2; …」按前缀
    # 「Key words」误判成标签 → 内容挂回标签、字号永远写不进 → verifier False。
    label_only_pat = re.compile(
        r"^(摘\s*要|摘要|关键词|Abstract|Key(?:\s+words?|words)|Index\s*Terms?|中图法分类号|中图分类号)\s*[:：]?\s*$",
        re.I,
    )
    label_runs: list = []
    content_runs: list = []
    for r in list(p.findall(f"{W}r")):
        rt = "".join(t.text or "" for t in r.iter(f"{W}t"))
        if not label_pat.match(rt.strip()):
            # 紧邻标签的纯冒号 run（「关键词」「：」被源稿拆成两 run）：
            # 冒号是标签的一部分，必须并入 label_runs —— 否则标签对齐写
            # 「关键词：」后遗留独立「：」run → 产物出现「关键词：：」
            # （2026-08-27 真机 content preservation FAILED 根因）。
            if rt.strip() in (":", "：") and label_runs:
                label_runs.append(r)
            else:
                content_runs.append(r)
            continue
        nr = _split_label_run(r, W)
        if nr is not None:
            r.addprevious(nr)
            label_runs.append(nr)
            rest = "".join(t.text or "" for t in r.iter(f"{W}t")).strip()
            if rest and label_only_pat.match(rest):
                label_runs.append(r)
            else:
                content_runs.append(r)
            continue
        # 整段单 run 混排（「Key words: a; b; c」一个 run，拆不出独立标签 run）：
        # 既当标签（加粗）又当内容（字号/字体）——只挂标签会让 content_runs 为空，
        # 字号永远写不进去，verifier 回读 sz 为空 → 整个事务回滚（2026-10-26 实证：
        # 东北师大 ×3、ACM、计算机工程与设计、东南本科 共 6 份整篇交付失败）。
        _rt = "".join(t.text or "" for t in r.iter(f"{W}t")).strip()
        _m = label_pat.match(_rt)
        _tail = _rt[_m.end():].strip(" :：") if _m else ""
        label_runs.append(r)
        if _tail:
            # 只有「同一 run 里标签后面还跟着内容」才算混排；纯标签 run 不许当内容
            # （否则标签会被套上正文字体，2026-08-27 回归测试守的就是这条）。
            content_runs.append(r)
    # 多 run 拆分标签（源稿「摘」/“ ”/「要」拆成 3 run，单 run 匹配失败）
    # → 按模板 label_text 核心累积匹配（2026-08-22 用户 tab 裁决）。
    core_want = re.sub(r"[\s:：]+", "", str(properties.get("label_text") or ""))
    if not label_runs and core_want:
        buf: list = []
        acc = ""
        for r in list(p.findall(f"{W}r")):
            rt = "".join(t.text or "" for t in r.iter(f"{W}t"))
            buf.append(r)
            acc += rt
            core = re.sub(r"[\s:：]+", "", acc)
            if core == core_want:
                label_runs = buf
                content_runs = [x for x in list(p.findall(f"{W}r")) if x not in buf]
                break
            if core and not core_want.startswith(core):
                break
    content = {k: v for k, v in (properties or {}).items()
               if not str(k).startswith("label_")}
    _set_run_properties(p, content, target_runs=content_runs)
    east = properties.get("label_font_east_asia")
    ascii_f = properties.get("label_font_ascii")
    sz = properties.get("label_font_size_pt")
    want_bold = properties.get("label_bold")
    want_italic = properties.get("label_italic")
    for r in label_runs:
        rpr = r.find(f"{W}rPr")
        if rpr is None:
            rpr = etree.Element(f"{W}rPr")
            r.insert(0, rpr)
        if east:
            rfonts = rpr.find(f"{W}rFonts")
            if rfonts is None:
                rfonts = etree.SubElement(rpr, f"{W}rFonts")
                rpr.insert(0, rfonts)
            rfonts.set(f"{W}eastAsia", east)
        if ascii_f:
            rfonts = rpr.find(f"{W}rFonts")
            if rfonts is None:
                rfonts = etree.SubElement(rpr, f"{W}rFonts")
                rpr.insert(0, rfonts)
            rfonts.set(f"{W}ascii", ascii_f)
            rfonts.set(f"{W}hAnsi", ascii_f)
        if sz is not None:
            half = str(int(round(float(sz) * 2)))
            for tag in ("sz", "szCs"):
                el = rpr.find(f"{W}{tag}")
                if el is None:
                    el = etree.SubElement(rpr, f"{W}{tag}")
                el.set(f"{W}val", half)
        b = rpr.find(f"{W}b")
        if want_bold is True and b is None:
            etree.SubElement(rpr, f"{W}b")
        elif want_bold is False and b is not None:
            rpr.remove(b)
        iel = rpr.find(f"{W}i")
        if want_italic is True and iel is None:
            etree.SubElement(rpr, f"{W}i")
        elif want_italic is False and iel is not None:
            rpr.remove(iel)
    # 标签文本对齐（仅一次，多 run 拆分用合并文本比较）：
    # 标签是对齐模板的格式锚点（用户裁决 2026-08-21：用户笔误的空格/标点
    # 向模板对齐，该空格就空格、该修标点就修标点）。仅比较「去空格、去冒号」
    # 后的核：模板"摘  要"可与源稿"摘要："对齐为模板原文；跨语言
    # （Abstract vs 摘要）核心不同，永不误换。
    want_text = properties.get("label_text")
    if want_text and label_runs:
        cur = "".join(t.text or "" for r in label_runs for t in r.iter(f"{W}t"))
        core_cur = re.sub(r"[\s:：]+", "", cur)
        core_want = re.sub(r"[\s:：]+", "", want_text)
        if _label_core_close(core_cur, core_want):
            # 「关键词」+「：」两 run 且模板 label_text 含冒号时，只写
            # label_runs[0] 会留下后续标签 run 的「：」→ 产物出现「关键词：：」
            # （2026-08-27 真机 content preservation FAILED 根因之一）。
            # 对策：首个标签 run 写模板原文，其余标签 run 全部清空。
            texts = list(label_runs[0].iter(f"{W}t"))
            if texts:
                texts[0].text = want_text
                for extra in texts[1:]:
                    extra.text = ""
                if " " in want_text:
                    texts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            for other in label_runs[1:]:
                for ttext in other.iter(f"{W}t"):
                    ttext.text = ""
    if (properties.get("tab_pos") or properties.get("label_after_tab")) and label_runs:
        last = label_runs[-1]
        nxt = last.getnext()
        if nxt is None or nxt.find(f"{W}tab") is None:
            tr = etree.Element(f"{W}r")
            # 模板 tab run 常带与标签相同的 rPr —— 复制标签 run 的 rPr 保持视觉一致
            lab_rpr = last.find(f"{W}rPr")
            if lab_rpr is not None:
                tr.append(copy.deepcopy(lab_rpr))
            etree.SubElement(tr, f"{W}tab")
            last.addnext(tr)
            # 模板「标签+tab+内容」无中间空格；源稿标签后常遗留空白 run
            # （“关键词 内容”），tab 已承担分隔——空白 run 减掉（2026-08-22）。
            nxt2 = tr.getnext()
            if nxt2 is not None and _is_blank_run(nxt2):
                p.remove(nxt2)


def _rewrite_spaced_heading(p, properties: dict) -> None:
    """把「参考文献」写成模板的「参 考 文 献」（仅当去空格后相同）。"""
    want = properties.get("label_text")
    if not want:
        return
    cur = "".join(t.text or "" for t in p.iter(f"{W}t"))
    if re.sub(r"[\s\u3000]+", "", cur) != re.sub(r"[\s\u3000]+", "", want):
        return
    texts = list(p.iter(f"{W}t"))
    if not texts:
        return
    texts[0].text = want
    for extra in texts[1:]:
        extra.text = ""
    if " " in want:
        texts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _extra_xml_matches(container, want_xml: str | None, kind: str) -> bool:
    """extra_xml 比对（verifier 用）：写入侧 _strip_font_theme_attrs 会清
    rFonts 主题属性——比对时对实际与期望两侧同样规范化。"""
    if want_xml is None:
        return True
    if container is None:
        return False
    try:
        w_el = etree.fromstring(want_xml.encode("utf-8"))
        _strip_font_theme_attrs(w_el)
        _strip_marker_from_extra(w_el, kind)
        # 与写入侧同一规范化（2026-09-29）：_apply_ooxml_extras 对 pPr 清的是
        # **pPr/rPr** 下的 rFonts 主题属性；只对 pPr 根调用找不到嵌套 rFonts →
        # 期望侧残留 eastAsiaTheme、实际侧已清 → 永远不等 → 整单回滚
        # （学校模板「学位论文原创性声明」样本 pPr/rPr 带 eastAsiaTheme 实锤）。
        if kind == "pPr":
            _w_rpr = w_el.find(f"{W}rPr")
            if _w_rpr is not None:
                _strip_font_theme_attrs(_w_rpr)
        want_norm = _extra_xml_from_lxml(w_el, kind)
    except Exception:
        want_norm = want_xml
    return _extra_xml_from_lxml(container, kind) == want_norm


def _extra_xml_from_lxml(container, kind: str) -> str | None:
    """Canonical extras for verifier; bridge lxml to stdlib parser."""
    if container is None:
        return None
    std = StdET.fromstring(etree.tostring(container, encoding="utf-8"))
    return extract_extra_bundle(std, kind)


def _structure_xml_from_lxml(container, kind: str) -> str | None:
    if container is None:
        return None
    std = StdET.fromstring(etree.tostring(container, encoding="utf-8"))
    return extract_structure_bundle(std, kind)


def _apply_structure_xml(container, xml_text: str | None, kind: str) -> None:
    if not xml_text:
        return
    wrapper = etree.fromstring(xml_text.encode("utf-8"))
    declared = {local_name(child.tag) for child in list(wrapper)}
    for child in list(container):
        if local_name(child.tag) in declared:
            container.remove(child)
    for child in list(wrapper):
        container.append(copy.deepcopy(child))


def _clear_extra_properties(container, kind: str) -> None:
    excluded = PPR_EXCLUDED if kind == "pPr" else RPR_EXCLUDED
    modeled = PPR_MODELED_ATTRS if kind == "pPr" else RPR_MODELED_ATTRS
    for child in list(container):
        name = local_name(child.tag)
        if name in excluded:
            continue
        known = modeled.get(name)
        if known is None:
            container.remove(child)
            continue
        for key in list(child.attrib):
            if local_name(key) not in known:
                del child.attrib[key]
        # Typed properties represented here are leaf nodes. Unknown descendants
        # belong to the extras channel and must be cleared before exact replay.
        for sub in list(child):
            child.remove(sub)


def _merge_extra_child(container, extra, kind: str) -> None:
    name = local_name(extra.tag)
    excluded = PPR_EXCLUDED if kind == "pPr" else RPR_EXCLUDED
    if name in excluded:
        # 作者语义/修订标记不属于模板 extras 通道：角色级回放注入 vertAlign 会把
        # 整行都变成上标，还会覆盖作者原本的上标下标（2026-09-27 真机事故）。
        return
    modeled = PPR_MODELED_ATTRS if kind == "pPr" else RPR_MODELED_ATTRS
    insert = _ppr_insert_ordered if kind == "pPr" else _rpr_insert_ordered
    if name not in modeled:
        clone = copy.deepcopy(extra)
        insert(container, clone)
        return
    current = next((x for x in list(container) if local_name(x.tag) == name), None)
    if current is None:
        current = etree.Element(extra.tag)
        insert(container, current)
    for key, value in extra.attrib.items():
        current.set(key, value)
    for sub in list(extra):
        old = next((x for x in list(current) if x.tag == sub.tag), None)
        if old is not None:
            current.remove(old)
        current.append(copy.deepcopy(sub))


def _strip_marker_from_extra(wrapper, kind: str) -> None:
    """把模板 extras 里「段落标记的编辑标记格式」剥掉（写入侧与验证侧共用）。

    2026-11-08 用户实锤：期刊 CJC 模板的题目样本段落标记带
    `<w:color FF0000/>` + `<w:bdr .../>`（模板作者标的「待替换」红框）。它被当成
    正版式搬进产物后，Word 渲染自动编号/段落标记时会画成红框红数字。
    两侧都必须用同一份净化后的期望，否则写入侧剥了、验证侧仍按原文比对 →
    verifier 判未生效 → 整个 set_paragraph_format 回滚（题目/作者因此一直不居中）。
    """
    if kind != "pPr":
        return
    rpr = wrapper.find(f"{W}rPr")
    if rpr is None:
        return
    for tag in _MARKER_RPR_TAGS:
        for el in list(rpr.findall(f"{W}{tag}")):
            rpr.remove(el)
    if len(rpr) == 0:
        wrapper.remove(rpr)


def _apply_extra_xml(container, xml_text: str | None, kind: str) -> None:
    """Clear all portable unmodelled values, then replay template extras exactly."""
    _clear_extra_properties(container, kind)
    if not xml_text:
        return
    wrapper = etree.fromstring(xml_text.encode("utf-8"))
    _strip_marker_from_extra(wrapper, kind)
    for child in list(wrapper):
        _merge_extra_child(container, child, kind)


def _stable_paragraph_fingerprint(p) -> str:
    """Formatting-independent identity: text + pStyle + object/table context."""
    from earlyz.ir.selectors import normalize_text
    text = normalize_text(_text_of(p))
    # Label spelling/spacing is presentation, not paragraph identity. The
    # decompiler may normalize "摘要：" to the same-language template form
    # "摘  要"; verifier target recovery must still resolve the original
    # paragraph. Content after the label remains part of the fingerprint.
    text = re.sub(r"^(?:摘\s*要|摘要)[:：—–-]?", "<abstract-cn>", text, flags=re.I)
    text = re.sub(r"^(?:关键词|Key(?:\s+words?|words)|Index\s*Terms?)[:：—–-]?",
                  "<keywords>", text, flags=re.I)
    text = re.sub(r"^Abstract[:：—–-]?", "<abstract-en>", text, flags=re.I)
    text = re.sub(r"^(?:中图法?分类号)[:：—–-]?", "<classification>", text, flags=re.I)
    ppr = p.find(f"{W}pPr")
    ps = ppr.find(f"{W}pStyle") if ppr is not None else None
    sid = ps.get(f"{W}val", "") if ps is not None else ""
    drawing = "1" if p.find(f".//{W}drawing") is not None else "0"
    in_table = "1" if any(True for _ in p.iterancestors(f"{W}tbl")) else "0"
    raw = "|".join((text, sid, drawing, in_table))
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:20]


def _capture_resolved_paragraphs(root, selected) -> list[dict]:
    """Capture document-order occurrence for each stable fingerprint."""
    chosen = {id(p) for p in selected}
    seen: dict[str, int] = {}
    out: list[dict] = []
    idx = 0
    for p in root.iter(f"{W}p"):
        fp = _stable_paragraph_fingerprint(p)
        occurrence = seen.get(fp, 0)
        seen[fp] = occurrence + 1
        if id(p) in chosen:
            # 同时记录文档顺序下标：公式段等「无文字段落」指纹彼此相同，
            # 仅靠 (fingerprint, occurrence) 会在插入/删除后错位（2026-10-20 实测：
            # equation_display 动作写对了，验证器却查到别的空段落 → 判未生效）。
            out.append({"fingerprint": fp, "occurrence": occurrence, "index": idx})
        idx += 1
    return out


def _resolve_captured_paragraphs(root, captured: list[dict]) -> list:
    captured = captured or []
    wanted = {(str(x.get("fingerprint")), int(x.get("occurrence", 0)))
              for x in captured}
    wanted_idx = {int(x["index"]) for x in captured if x.get("index") is not None}
    # 指纹是主锚点；下标只在「完全拿不到指纹」时兜底。
    # 2026-10-20 实测：把下标当并列条件会让「插入段落后重定位」失效
    # （绝对下标已漂移），也会让「段落已删除」误判成命中别的段。
    has_fingerprint = any(x.get("fingerprint") is not None for x in captured)
    seen: dict[str, int] = {}
    out = []
    idx = 0
    for p in root.iter(f"{W}p"):
        fp = _stable_paragraph_fingerprint(p)
        occurrence = seen.get(fp, 0)
        seen[fp] = occurrence + 1
        if has_fingerprint:
            if (fp, occurrence) in wanted:
                out.append(p)
        elif idx in wanted_idx:
            out.append(p)
        idx += 1

    # 注：不做「指纹失配 → 下标兜底」。_resolve_captured_paragraphs 同时服务于
    # writer._resolve_legacy_index 的「目标段已删除必须返回 None」契约
    # （tests/test_stable_selector.py 守）。验证层的兜底放在 verify_paragraph_format。
    return out


_LABEL_RUN_RE = re.compile(
    r"^(摘\s*要|摘要|关键词|Abstract|Key(?:\s+words?|words)|Index\s*Terms?|中图法分类号|中图分类号)\s*[:：]?\s*$",
    re.I,
)


def _strip_font_theme_attrs(rpr) -> None:
    """清除 rPr 中 rFonts 的主题字体属性（2026-08-23 单栏等线实锤）。

    WPS/Word 模板样本 rPr 常带 eastAsiaTheme=minorEastAsia 等——Word 解析
    时主题字体优先于显式字体与 docDefaults → 宋体不生效显示「等线(中文
    正文)」。显式字体语义下 theme 属性必须清理。
    """
    fonts = rpr.find(f"{W}rFonts")
    if fonts is None:
        return
    for tk in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        fonts.attrib.pop(f"{W}{tk}", None)


_MARKER_RPR_TAGS = ("color", "highlight", "bdr", "shd", "vanish", "u")


def _strip_marker_rpr(rpr) -> int:
    """去掉段落标记 rPr 里的「编辑标记」格式（颜色/高亮/边框/底纹/隐藏/下划线）。

    2026-11-08 用户实锤：期刊文档每个一级标题前面出现「红色方框 + 红色数字」。
    根因不是编号，是模板样本段的 **段落标记**（pPr/rPr）被原样搬进产物，带着
    color=FF0000 与 bdr 红框；Word 渲染自动编号时会继承段落标记的字符格式，
    于是编号变成了红框红字。颜色/边框属于模板作者的「待替换」标记，不是版式。
    """
    if rpr is None:
        return 0
    n = 0
    for tag in _MARKER_RPR_TAGS:
        for el in list(rpr.findall(f"{W}{tag}")):
            rpr.remove(el)
            n += 1
    return n


def _apply_ooxml_extras(p, properties: dict, role: str | None) -> None:
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        ppr = etree.Element(f"{W}pPr")
        p.insert(0, ppr)
    if "ppr_extra_xml" in properties:
        _apply_extra_xml(ppr, properties.get("ppr_extra_xml"), "pPr")
        _ppr_rpr = ppr.find(f"{W}rPr")
        if _ppr_rpr is not None:
            _strip_font_theme_attrs(_ppr_rpr)
            _strip_marker_rpr(_ppr_rpr)
    if "rpr_extra_xml" not in properties:
        return
    special = role in ("keywords", "keywords_en", "abstract", "abstract_en", "classification")
    for r in p.findall(f"{W}r"):
        text = "".join(t.text or "" for t in r.iter(f"{W}t")).strip()
        if special and _LABEL_RUN_RE.match(text):
            continue
        if any(a.tag == f"{M}oMath" or a.tag == f"{M}oMathPara"
               or a.tag.endswith("object") or a.tag.endswith("OLEObject")
               for a in r.iterancestors()):
            continue
        rpr = r.find(f"{W}rPr")
        if rpr is None:
            rpr = etree.Element(f"{W}rPr")
            r.insert(0, rpr)
        _apply_extra_xml(rpr, properties.get("rpr_extra_xml"), "rPr")
        _strip_font_theme_attrs(rpr)


# 作者语义（2026-11-08 根因修复）：段级格式写入绝不允许改动 run 级的
# vertAlign（上下标）/ vanish（隐藏）/ rtl（书写方向）。
# 事故链：模板 extras 回放 + 属性重写把作者名后的 1) 2) 上标抹掉，产物里角标消失；
# 此前靠「流水线末尾再跑一次 superscript_front_marks」兜住，属顺序侥幸，换文档即复发。
_RUN_SEMANTIC_TAGS = ("vertAlign", "vanish", "specVanish", "rtl")


def _snapshot_run_semantics(p):
    """记录段落内每个 run 的语义性 rPr 子元素（按 run 文字定位）。"""
    snap = []
    for r in p.iter(f"{W}r"):
        rpr = r.find(f"{W}rPr")
        if rpr is None:
            continue
        text = "".join(t.text or "" for t in r.iter(f"{W}t"))
        keep = [copy.deepcopy(el) for el in rpr
                if local_name(el.tag) in _RUN_SEMANTIC_TAGS]
        if keep:
            snap.append((text, keep))
    return snap


def _restore_run_semantics(p, snap) -> None:
    """把被抹掉的语义性 rPr 子元素按文字对回原来的 run。"""
    if not snap:
        return
    by_text = {}
    for r in p.iter(f"{W}r"):
        t = "".join(x.text or "" for x in r.iter(f"{W}t"))
        by_text.setdefault(t, []).append(r)
    def _attach(r, keep_els):
        rpr = r.find(f"{W}rPr")
        if rpr is None:
            rpr = etree.Element(f"{W}rPr")
            r.insert(0, rpr)
        for el in keep_els:
            name = local_name(el.tag)
            if rpr.find(f"{W}{name}") is None:
                _rpr_insert_ordered(rpr, copy.deepcopy(el))

    for text, keep in snap:
        hits = by_text.get(text) or []
        if hits:
            for r in hits:
                _attach(r, keep)
            continue
        # 兜底：目标文字被并进了更长的 run（extras 回放/重写会重新切 run）——
        # 把它从长 run 里切出来再挂语义，否则「作者角标消失」原样复发。
        if not text:
            continue
        for r in list(p.iter(f"{W}r")):
            rtext = "".join(t.text or "" for t in r.iter(f"{W}t"))
            pos = rtext.find(text)
            if pos < 0:
                continue
            ts = list(r.iter(f"{W}t"))
            if len(ts) != 1 or ts[0].text is None:
                continue
            full = ts[0].text
            head, tail = full[:pos], full[pos + len(text):]
            ts[0].text = head
            mid = copy.deepcopy(r)
            mid_ts = list(mid.iter(f"{W}t"))
            mid_ts[0].text = text
            _attach(mid, keep)
            r.addnext(mid)
            if tail:
                tail_run = copy.deepcopy(r)
                tail_ts = list(tail_run.iter(f"{W}t"))
                tail_ts[0].text = tail
                _attach(tail_run, [])
                mid.addnext(tail_run)
            break


def set_paragraph_format(target_path: Path, params: dict) -> dict:
    """按 selector 定位段落，应用 FormatProperties。

    params: {selector: {...}, properties: {jc, keep_next, ...}}
    """
    selector = params.get("selector") or {}
    properties = params.get("properties") or {}
    if not selector and not params.get("_resolved_paragraphs"):
        return {"ok": False, "changed": False, "error": (
            "选择器为空：set_paragraph_format 需要 selector（semantic_role / node_id / "
            "text_hash / paragraph_index 至少一个）来定位段落；直接传 properties 会命中全文。"
        )}
    root = _read_part(target_path, "word/document.xml")
    captured = params.get("_resolved_paragraphs") or []
    paras = (_resolve_captured_paragraphs(root, captured) if captured else
             _locate_paragraphs(root, selector, _read_styles_map(target_path)))
    if not paras and selector.get("semantic_role") == "equation_display":
        # 公式段：含 m:oMath / m:oMathPara 的段落（2026-10-20 上海交大模板实测，
        # 语义表里没有该角色 → 选择器解析为空 → 公式段永远排不上）
        _M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
        paras = [p for p in root.iter(f"{W}p")
                 if p.find(f".//{_M}oMathPara") is not None
                 or p.find(f".//{_M}oMath") is not None]
    if not paras:
        # 源文档没有该角色（跨模板常见：目标模板要求 affiliation/单位署名，
        # 源文没有这一段）→ 跳过而不是整份回滚；记录原因供交付报告使用。
        return {"ok": True, "changed": False,
                "skipped": f"no paragraph matched selector: {selector}"}
    if not captured:
        params["_resolved_paragraphs"] = _capture_resolved_paragraphs(root, paras)

    _sem_snaps = [(p, _snapshot_run_semantics(p)) for p in paras]
    for p in paras:
        ppr = p.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            p.insert(0, ppr)
        tab_pos = properties.get("tab_pos")
        if tab_pos:
            tabs = ppr.find(f"{W}tabs")
            if tabs is None:
                tabs = etree.Element(f"{W}tabs")
                ppr.insert(0, tabs)
            tab = tabs.find(f"{W}tab")
            if tab is None:
                tab = etree.SubElement(tabs, f"{W}tab")
            tab.set(f"{W}val", "left")
            tab.set(f"{W}pos", str(int(tab_pos)))
        # ⚠ 顺序铁律（2026-11-08 用户实锤「期刊题目/作者不居中」）：
        # 模板 extras（ppr_extra_xml）必须**先**落地，显式属性后写。
        # `_apply_extra_xml` 会先清掉 pPr 里未被建模的值再回放模板原样——
        # 放在属性循环之后会把刚写好的 jc=center / spacing 一并清掉，
        # 于是题目和作者行永远两端对齐（模板里它们本是居中）。
        _apply_ooxml_extras(p, properties, selector.get("semantic_role"))
        # run 级属性：keywords 行特判（标签黑体加粗），其余统一
        if selector.get("semantic_role") in (
            "keywords", "keywords_en", "abstract", "abstract_en", "classification",
        ):
            _set_keywords_run_properties(p, properties)
        else:
            _set_run_properties(p, properties)
            _rewrite_spaced_heading(p, properties)
        for prop, value in properties.items():
            if value is None or prop not in _PPR_MAP:
                continue
            tag, attr = _PPR_MAP[prop]
            # 行单位优先（2026-08-18 用户彻查「行就是行」）：spacing_before/
            # spacing_after 磅值在有对应 _lines 行单位时跳过——lines 分支
            # 统一换算写磅值+行值（双写），避免磅值先写、lines 再覆盖成
            # 不匹配的值（verifier 假阴性 rolled_back 实证）。
            if attr == "before" and "spacing_before_lines" in properties:
                continue
            if attr == "after" and "spacing_after_lines" in properties:
                continue
            if attr == "__line_spacing__":
                # 倍数行距：spacing@line = round(v*240) + lineRule=auto
                spacing = ppr.find(f"{W}spacing")
                if spacing is None:
                    spacing = etree.SubElement(ppr, f"{W}spacing")
                spacing.set(f"{W}line", str(int(round(float(value) * 240))))
                spacing.set(f"{W}lineRule", "auto")
                continue
            if attr == "beforeLines" or attr == "afterLines":
                # 行单位（2026-08-18 用户彻查「行就是行」）：模板用行，
                # earlyz 就用行——写 beforeLines/afterLines（×100 整数）。
                # 磅值：优先用 properties 里的 spacing_before/spacing_after
                # 原值（模板实证 before=251twips 与 0.8行 并存但值独立，
                # 不是 0.8×12.55pt 的换算值——自换算会写出不匹配磅值，
                # Word 显示异常）；无原值才按 1行≈12.55pt 换算兜底。
                spacing = ppr.find(f"{W}spacing")
                if spacing is None:
                    spacing = etree.SubElement(ppr, f"{W}spacing")
                try:
                    lines = float(value)
                except (TypeError, ValueError):
                    continue
                spacing.set(f"{W}{attr}", str(int(round(lines * 100))))
                counterpart = "before" if attr == "beforeLines" else "after"
                src_key = "spacing_before" if attr == "beforeLines" else "spacing_after"
                src_val = properties.get(src_key)
                if src_val is not None:
                    spacing.set(f"{W}{counterpart}", str(int(src_val)))
                # If the counterpart already matched, it is absent from this
                # diff action. Preserve it; never derive twips from lines.
                continue
            if attr is None:
                # 布尔属性：keepNext/keepLines 三态写入（2026-08-17 D2）：
                # True=元素存在；False=显式关闭 w:val="0"（不是删除——
                # 删除=继承样式，显式关闭才能压过样式级 keepNext）；
                # None=不动。此前 False 走 remove → 模板样式级 keepNext
                # 重新生效，"关闭"指令实际无效。
                existing = ppr.find(tag)
                if value is True:
                    if existing is None:
                        existing = etree.Element(tag)
                        _ppr_insert_ordered(ppr, existing)
                    existing.attrib.pop(f"{W}val", None)
                elif value is False:
                    if existing is None:
                        existing = etree.Element(tag)
                        _ppr_insert_ordered(ppr, existing)
                    existing.set(f"{W}val", "0")
            else:
                el = ppr.find(tag)
                if el is None:
                    el = etree.Element(tag)
                    _ppr_insert_ordered(ppr, el)
                if isinstance(value, bool):
                    val_str = "1" if value else "0"
                else:
                    val_str = str(int(value)) if isinstance(value, (int, float)) and not isinstance(value, bool) else str(value)
                el.set(f"{W}{attr}", val_str)
    # 兜底还原：任何 extras 回放/属性重写都不得改变作者语义（上下标/隐藏/rtl）
    for _p, _snap in _sem_snaps:
        _restore_run_semantics(_p, _snap)

    _write_part(target_path, "word/document.xml", root)
    return {"changed": True, "paragraphs": len(paras)}


def _content_run_for_verify(p, role: str | None):
    """摘要/关键词：跳过标签 run，取内容 run。其余角色取第一个 run。"""
    if role not in ("keywords", "keywords_en", "abstract", "abstract_en", "classification"):
        return p.find(f"{W}r")
    pat = re.compile(
        r"^(摘\s*要|摘要|关键词|Abstract|Key(?:\s+words?|words)|Index\s*Terms?|中图法分类号|中图分类号)\s*[:：]?\s*$",
        re.I,
    )
    for r in p.findall(f"{W}r"):
        rt = "".join(t.text or "" for t in r.iter(f"{W}t")).strip()
        if not rt or pat.match(rt):
            continue
        return r
    return p.find(f"{W}r")



def _clean(el) -> str:
    """调试用：元素转 XML 去掉命名空间声明与多余空白。"""
    if el is None:
        return "-"
    import re as _r
    x = etree.tostring(el, encoding="unicode")
    x = _r.sub(r'\sxmlns(:\w+)?="[^"]*"', "", x)
    return _r.sub(r"\s+", " ", x)[:260]

def _verify_label_runs(p, properties: dict) -> bool:
    east = properties.get("label_font_east_asia")
    ascii_f = properties.get("label_font_ascii")
    sz = properties.get("label_font_size_pt")
    want_bold = properties.get("label_bold")
    want_italic = properties.get("label_italic")
    if east is None and ascii_f is None and sz is None and want_bold is None and want_italic is None:
        return True
    pat = re.compile(
        r"^(摘\s*要|摘要|关键词|Abstract|Key(?:\s+words?|words)|Index\s*Terms?|中图法分类号|中图分类号)\s*[:：]?\s*$",
        re.I,
    )
    for r in p.findall(f"{W}r"):
        rt = "".join(t.text or "" for t in r.iter(f"{W}t")).strip()
        if not pat.match(rt):
            continue
        rpr = r.find(f"{W}rPr")
        rf = rpr.find(f"{W}rFonts") if rpr is not None else None
        if east and (rf is None or rf.get(f"{W}eastAsia") != east):
            return False
        if ascii_f and (rf is None or rf.get(f"{W}ascii") != ascii_f):
            return False
        if sz is not None:
            el = rpr.find(f"{W}sz") if rpr is not None else None
            if el is None or el.get(f"{W}val") != str(int(round(float(sz) * 2))):
                return False
        if want_bold is True and (rpr is None or rpr.find(f"{W}b") is None):
            return False
        if want_italic is True and (rpr is None or rpr.find(f"{W}i") is None):
            return False
        return True
    return True


def verify_paragraph_format(target_path: Path, params: dict) -> bool:
    """验证段落格式已按期望设置（任一段落任一属性匹配即可视为生效）。"""
    selector = params.get("selector") or {}
    properties = params.get("properties") or {}
    try:
        root = _read_part(target_path, "word/document.xml")
        captured = params.get("_resolved_paragraphs") or []
        paras = (_resolve_captured_paragraphs(root, captured) if captured else
                 _locate_paragraphs(root, selector, _read_styles_map(target_path)))
        if not paras and captured:
            # 设置阶段可能改写段落文本（关键词标签拆分/补冒号），指纹随之失效
            # → 一个都定位不到 → 验证器判「未生效」→ 整篇事务回滚
            # （2026-10-26 计算机工程与设计投稿 #10 keywords_en 实测）。
            # 这里按选择器重新定位；只影响「验证是否生效」，不改变定位契约。
            paras = _locate_paragraphs(root, selector, _read_styles_map(target_path))
        if not paras:
            return False
        _checked = 0
        _label_only = 0
        for p in paras:
            ppr = p.find(f"{W}pPr")
            if ppr is None:
                continue
            # 纯标签段（「KEYWORDS」「摘要：」）不承载内容属性：英文期刊模板里
            # 关键词正文常在下一段，内容字号/字体/颜色无处可写，verifier 却要求
            # 全匹配 → 假阴性回滚整篇（2026-10-26 ACM 实测：失败属性
            # font_east_asia/color/rpr_extra_xml）。只跳过这类，不改变 body 语义。
            _role = selector.get("semantic_role")
            if _role in ("keywords", "keywords_en", "abstract", "abstract_en", "classification"):
                _t = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
                if re.match(
                    r"^(摘\s*要|摘要|关键词|Abstract|Key(?:\s+words?|words)|"
                    r"KEYWORDS|Index\s*Terms?|中图法分类号|中图分类号)\s*[:：]?\s*$",
                    _t, re.I,
                ):
                    _label_only += 1
                    continue
            _checked += 1
            ok = True
            for prop, value in properties.items():
                if value is None or prop not in _PPR_MAP:
                    continue
                tag, attr = _PPR_MAP[prop]
                # 行单位优先（2026-08-18）：磅值在有行单位时不单独校验
                # （lines 分支统一校验换算磅值）——与 set 侧写入逻辑一致。
                if attr == "before" and "spacing_before_lines" in properties:
                    continue
                if attr == "after" and "spacing_after_lines" in properties:
                    continue
                el = ppr.find(tag)
                if attr == "__line_spacing__":
                    # 倍数行距：spacing@line == round(v*240) 且 lineRule=auto
                    if el is None or el.get(f"{W}lineRule") != "auto":
                        ok = False
                    elif el.get(f"{W}line") != str(int(round(float(value) * 240))):
                        ok = False
                    continue
                if attr == "beforeLines" or attr == "afterLines":
                    # 行单位（2026-08-18）：验证 beforeLines/afterLines 值
                    # （×100 整数）+ 磅值换算双写一致（模板同样双写）。
                    if el is None:
                        ok = False
                        continue
                    try:
                        lines = float(value)
                    except (TypeError, ValueError):
                        continue
                    if el.get(f"{W}{attr}") != str(int(round(lines * 100))):
                        ok = False
                        continue
                    # 磅值双写检查（与 set 侧同源：优先 properties 原值，
                    # 无原值才按 12.55pt/行换算——2026-08-18 用户彻查）
                    counterpart = "before" if attr == "beforeLines" else "after"
                    src_key = "spacing_before" if attr == "beforeLines" else "spacing_after"
                    src_val = properties.get(src_key)
                    if src_val is not None:
                        expect_pt = str(int(src_val))
                        if el.get(f"{W}{counterpart}") != expect_pt:
                            ok = False
                    continue
                if attr is None:
                    state = None if el is None else el.get(f"{W}val") not in ("0", "false", "off")
                    if value is True and state is not True:
                        ok = False
                    elif value is False and state is not False:
                        ok = False
                else:
                    if el is None:
                        ok = False
                    elif el.get(f"{W}{attr}") != (str(int(value)) if isinstance(value, (int, float)) and not isinstance(value, bool) else str(value)):
                        ok = False
            if "ppr_extra_xml" in properties:
                if not _extra_xml_matches(ppr, properties.get("ppr_extra_xml"), "pPr"):
                    ok = False
            # run 级属性：摘要/关键词校验内容 run，避免标签黑体把 verifier 打回
            role = selector.get("semantic_role")
            r = _content_run_for_verify(p, role)
            if r is None:
                r = p.find(f"{W}r")
            if r is None:
                # 公式段没有 w:r（文字在 m:r 里）→ run 级属性无从校验。
                # 注意不能 continue：那会跳过后面的“判成功”分支，函数落到 return False。
                if p.find(f".//{M}oMath") is None and p.find(f".//{M}oMathPara") is None:
                    ok = False
            rpr = r.find(f"{W}rPr") if r is not None else None
            for prop, value in (properties.items() if rpr is not None else ()):
                if value is None or prop not in _RPR_MAP:
                    continue
                tag, attr = _RPR_MAP[prop]
                el = rpr.find(tag) if rpr is not None else None
                if attr is None:
                    state = None if el is None else el.get(f"{W}val") not in ("0", "false", "off")
                    if value is True and state is not True:
                        ok = False
                    elif value is False and state is not False:
                        ok = False
                else:
                    expect = str(int(round(float(value) * 2))) if prop == "font_size_pt" else str(value)
                    if el is None or el.get(f"{W}{attr}") != expect:
                        ok = False
            # 公式段没有 w:r → rpr 为 None，run 级额外属性无从校验（2026-10-20）
            if "rpr_extra_xml" in properties and rpr is not None:
                if not _extra_xml_matches(rpr, properties.get("rpr_extra_xml"), "rPr"):
                    ok = False
            if not _verify_label_runs(p, properties):
                ok = False
            if ok:
                return True
            if os.environ.get("EARLYZ_VERIFY_DEBUG"):
                _txt = "".join(t.text or "" for t in p.iter(f"{W}t"))[:40]
                _ppr = p.find(f"{W}pPr")
                _r0 = p.find(f"{W}r")
                _rpr = _r0.find(f"{W}rPr") if _r0 is not None else None
                _pp = _ppx(_ppr) if "_ppx" in dir() else ""
                print(f"[verify] 段未通过 role={selector.get('semantic_role')} text={_txt!r} "
                      f"props={list(properties)} rPr={_clean(_rpr)} pPr={_clean(_ppr)}", flush=True)
        if _checked == 0 and _label_only > 0:
            # 目标段落全是纯标签段：没有承载内容的段落，视为无需校验
            return True
        if os.environ.get("EARLYZ_VERIFY_DEBUG"):
            print(f"[verify] 无段通过 selector={selector} props={list(properties)}", flush=True)
        return False
    except Exception:
        if os.environ.get("EARLYZ_VERIFY_DEBUG"):
            import traceback as _tb
            _tb.print_exc()
        return False


# ────────────────────────────────────────────────────────────────
# 参数归一化：planner 输出 → 旧 action 参数
# ────────────────────────────────────────────────────────────────

_OLD_STYLE_MAP = {
    "font_east_asia": "font_eastAsia",
    "font_ascii": "font_ascii",
    "font_size_pt": "font_size_pt",
    "bold": "bold",
    "italic": "italic",
    "jc": "alignment",
}


def set_page_margins_all(target_path: Path, params: dict) -> dict:
    """T2 节级页面几何规则执行端：所有 sectPr 统一页边距。

    params: {properties: {top, bottom, left, right, header, footer}}
    （值来自 spec 规则，零硬编码；与 lint 的 margins_diff 容差语义一致：
    只写超容差的键——但直接全写也幂等，简化为全写。）
    """
    props = params.get("properties") or {}
    if not props:
        raise ValueError("set_page_margins_all requires properties")
    root = _read_part(target_path, "word/document.xml")
    sects = list(root.iter(f"{W}sectPr"))
    if not sects:
        raise ValueError("no sectPr found")
    changed = 0
    for sp in sects:
        margin_props = {k: v for k, v in props.items() if k != "section_extra_xml"}
        if margin_props:
            pm = sp.find(f"{W}pgMar")
            if pm is None:
                pm = etree.SubElement(sp, f"{W}pgMar")
            for k, v in margin_props.items():
                pm.set(f"{W}{k}", str(int(v)))
        if "section_extra_xml" in props:
            _apply_structure_xml(sp, props.get("section_extra_xml"), "sectPr")
        _normalize_child_order(sp, _CT_SECTPR_ORDER)
        changed += 1
    _write_part(target_path, "word/document.xml", root)
    return {"changed": True, "sections": changed}





def _format_header_footer_parts(target_path: Path, params: dict, prefix: str) -> dict:
    """Apply typed + lossless formatting without changing part text or fields.

    2026-10-18（Layer 1）：修两处真机缺陷——
    - 原来的 `if not 文本: continue` 会跳过「空段落 + 下边框」这种最常见的页眉线形态，
      导致「去横线」永远零变化却报成功；
    - 原来只支持**添加** `p_bdr_bottom`，没有删除语义，也不认 `p_bdr_none`。
    现在边框按"增/删"双向处理且覆盖空段落；字体类属性仍只作用于有文字的段落。
    """
    props = params.get("properties") or {}
    if not props:
        raise ValueError(f"format_{prefix}s requires properties")

    border_want = props.get("p_bdr_bottom")       # True=加下边框 / False=删下边框
    border_none = bool(props.get("p_bdr_none"))   # True=删掉整段 pBdr
    scoped_sections = params.get("sections")
    scoped_parts = params.get("parts")
    if isinstance(scoped_sections, list) and scoped_sections:
        names = [n for n in _parts_for_sections(target_path, scoped_sections)
                 if n.startswith(f"word/{prefix}")]
    elif isinstance(scoped_parts, list) and scoped_parts:
        names = [n for n in scoped_parts if isinstance(n, str) and n.startswith(f"word/{prefix}")]
    else:
        names = [
            name for name in _part_names(target_path)
            if name.startswith(f"word/{prefix}") and name.endswith(".xml")
        ]

    changed = 0
    parts_touched = 0
    for part_name in names:
        root = _read_part(target_path, part_name)
        part_delta = 0
        for pel in root.iter(f"{W}p"):
            has_text = bool("".join(t.text or "" for t in pel.iter(f"{W}t")).strip())

            if border_want is not None or border_none:
                ppr = pel.find(f"{W}pPr")
                if ppr is None and (border_want or border_none):
                    ppr = etree.Element(f"{W}pPr")
                    pel.insert(0, ppr)

                if ppr is not None and border_none:
                    pbdr = ppr.find(f"{W}pBdr")
                    if pbdr is not None:
                        ppr.remove(pbdr)
                        part_delta += 1

                if ppr is not None and border_want is True:
                    pbdr = ppr.find(f"{W}pBdr")
                    if pbdr is None:
                        pbdr = etree.SubElement(ppr, f"{W}pBdr")
                    bottom = pbdr.find(f"{W}bottom")
                    if bottom is None:
                        bottom = etree.SubElement(pbdr, f"{W}bottom")
                    want = {"val": "single", "sz": "6", "space": "1", "color": "auto"}
                    before = {k: bottom.get(f"{W}{k}") for k in want}
                    for k, v in want.items():
                        bottom.set(f"{W}{k}", v)
                    if before != want:
                        part_delta += 1

                if ppr is not None and border_want is False:
                    pbdr = ppr.find(f"{W}pBdr")
                    if pbdr is not None:
                        bottom = pbdr.find(f"{W}bottom")
                        if bottom is not None:
                            pbdr.remove(bottom)
                            part_delta += 1
                        if len(pbdr) == 0:
                            ppr.remove(pbdr)
                    if len(ppr) == 0:
                        pel.remove(ppr)



            if "ppr_extra_xml" in props:
                ppr_now = pel.find(f"{W}pPr")
                if ppr_now is None and has_text:
                    ppr_now = etree.Element(f"{W}pPr")
                    pel.insert(0, ppr_now)
                if ppr_now is not None:
                    _apply_extra_xml(ppr_now, props.get("ppr_extra_xml"), "pPr")
                    _normalize_child_order(ppr_now, _CT_PPR_ORDER)
                    part_delta += 1

            # 字体/段落属性只作用于有文字的段落——空段落没有 run 可写
            if has_text:
                runs = list(pel.iter(f"{W}r"))
                _set_run_properties(pel, props, target_runs=runs)
                if "rpr_extra_xml" in props:
                    for run in runs:
                        rpr = run.find(f"{W}rPr")
                        if rpr is None:
                            rpr = etree.Element(f"{W}rPr")
                            run.insert(0, rpr)
                        _apply_extra_xml(rpr, props.get("rpr_extra_xml"), "rPr")
                        _strip_font_theme_attrs(rpr)
                        _normalize_child_order(rpr, _CT_RPR_ORDER)
                part_delta += 1

        if part_delta:
            _write_part(target_path, part_name, root)
            changed += part_delta
            parts_touched += 1

    return {"ok": True, f"{prefix}s": changed, "parts_touched": parts_touched, "properties": props}


def _rpr_actual(rpr, prop: str):
    tag, attr = _RPR_MAP[prop]
    el = rpr.find(tag) if rpr is not None else None
    if el is None:
        return None
    if prop == "font_size_pt":
        try:
            return int(el.get(f"{W}val")) / 2
        except (TypeError, ValueError):
            return None
    if attr is None:
        return el.get(f"{W}val") not in ("0", "false", "off")
    return el.get(f"{W}{attr}")


def _verify_header_footer_parts(target_path: Path, params: dict, prefix: str) -> bool:
    props = params.get("properties") or {}
    names = [
        name for name in _part_names(target_path)
        if name.startswith(f"word/{prefix}") and name.endswith(".xml")
    ]
    for part_name in names:
        root = _read_part(target_path, part_name)
        for pel in root.iter(f"{W}p"):
            if not "".join(t.text or "" for t in pel.iter(f"{W}t")).strip():
                continue
            ppr = pel.find(f"{W}pPr")
            if props.get("p_bdr_bottom"):
                bottom = ppr.find(f"{W}pBdr/{W}bottom") if ppr is not None else None
                if bottom is None:
                    return False
            if "ppr_extra_xml" in props:
                if _extra_xml_from_lxml(ppr, "pPr") != props.get("ppr_extra_xml"):
                    return False
            runs = list(pel.iter(f"{W}r"))
            for run in runs:
                rpr = run.find(f"{W}rPr")
                for prop, expected in props.items():
                    if prop in _RPR_MAP and _rpr_actual(rpr, prop) != expected:
                        return False
                if "rpr_extra_xml" in props:
                    if _extra_xml_from_lxml(rpr, "rPr") != props.get("rpr_extra_xml"):
                        return False
    return True


def format_headers(target_path: Path, params: dict) -> dict:
    return _format_header_footer_parts(target_path, params, "header")


# ══════════════════════════════════════════════════════════════════
# 节骨架（2026-10-18 Layer 1）：按角色标题切节，每节复制完整版式
# ══════════════════════════════════════════════════════════════════

_ROLE_PG_NUM = {
    "front": {"fmt": "upperRoman", "start": 1},
    "body": {"fmt": "decimal", "start": 1},
    "back": None,          # 延续正文编号
}


def _norm_heading(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def _preceding_section_break(para) -> bool:
    prev = para.getprevious()
    while prev is not None and isinstance(prev.tag, str) and prev.tag.endswith("}" + "sdt"):
        prev = prev.getprevious()
    if prev is None or not str(prev.tag).endswith("}p"):
        return False
    ppr = prev.find(f"{W}pPr")
    return ppr is not None and ppr.find(f"{W}sectPr") is not None


def build_section_skeleton(target_path: Path, params: dict) -> dict:
    """按角色标题切节：在每个标题前插入一个「完整 sectPr」分节符段落。

    与旧的 `insert_section_break` 的区别：分节符里的 sectPr 是**当前最终 sectPr 的
    深拷贝**（pgSz/pgMar/cols/docGrid/header/footer 引用全带走），并按角色写
    `pgNumType`（前置罗马、正文从 1 起编）。旧实现只写 type+pgSz，新节会丢版式。

    ⚠️ 这里深拷贝到的是**建节当时**文档级 sectPr 的几何（通常就是封面那套）——
    只能当初值用，不能当成最终版式。最终几何由 `format_compiler` 5.98 的
    `align_section_geometry` 按角色（封面/摘要/目录/正文/参考文献/致谢…）重写
    （EZI-0018）：本动作拿不到模板路径，无法在建节时就知道每节该用哪套边距。

    params:
        boundaries: [{heading, role?: front|body|back, pg_num?: {fmt, start}}]
    """
    boundaries = params.get("boundaries") or []
    if not isinstance(boundaries, list) or not boundaries:
        raise ValueError("boundaries 不能为空：需要至少一个切节标题")

    root = _read_part(target_path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        raise ValueError("document.xml 缺少 body")

    final = body.find(f"{W}sectPr")
    if final is None:
        # 兜底：取最后一个段落里的 sectPr
        for pel in reversed(list(body.iter(f"{W}p"))):
            ppr = pel.find(f"{W}pPr")
            if ppr is not None and ppr.find(f"{W}sectPr") is not None:
                final = ppr.find(f"{W}sectPr")
                break
    if final is None:
        raise ValueError("文档没有任何 sectPr，无法作为分节模板")

    created, skipped, missing = 0, 0, []
    for item in boundaries:
        heading = _norm_heading(item.get("heading", ""))
        if not heading:
            skipped += 1
            continue
        target = None
        for pel in body.iter(f"{W}p"):
            txt = _norm_heading("".join(t.text or "" for t in pel.iter(f"{W}t")))
            if txt == heading:
                target = pel
                break
        if target is None:
            missing.append(item.get("heading"))
            continue
        if _preceding_section_break(target):
            skipped += 1
            continue

        sect = copy.deepcopy(final)
        # 节类型（分节符段落里的 sectPr 决定"上一节"如何结束）
        typ = sect.find(f"{W}type")
        if typ is None:
            typ = etree.SubElement(sect, f"{W}type")
        typ.set(f"{W}val", "nextPage")

        pg_num = item.get("pg_num") or _ROLE_PG_NUM.get(item.get("role") or "")
        if pg_num:
            pgn = sect.find(f"{W}pgNumType")
            if pgn is None:
                pgn = etree.SubElement(sect, f"{W}pgNumType")
            pgn.set(f"{W}fmt", str(pg_num.get("fmt", "decimal")))
            pgn.set(f"{W}start", str(pg_num.get("start", 1)))
        _normalize_child_order(sect, _CT_SECTPR_ORDER)

        new_p = etree.Element(f"{W}p")
        new_ppr = etree.SubElement(new_p, f"{W}pPr")
        new_ppr.append(sect)
        target.addprevious(new_p)
        created += 1

    _write_part(target_path, "word/document.xml", root)

    # 每节克隆自己的页眉/页脚部件：否则所有节共用 header1.xml，
    # 「只改封面页眉」在物理上不可能（2026-10-18 Layer 1）。
    cloned = 0
    try:
        cloned = _clone_parts_per_section(target_path) if created else 0
    except Exception:  # noqa: BLE001 — 克隆失败不阻断分节本身
        cloned = 0

    return {"ok": True, "created": created, "skipped": skipped,
            "missing_headings": missing, "cloned_parts": cloned}


def verify_section_skeleton(target_path: Path, params: dict) -> bool:
    """每个 boundary 前必须有分节符段落，且所有 sectPr 带 pgSz/pgMar。"""
    boundaries = params.get("boundaries") or []
    try:
        root = _read_part(target_path, "word/document.xml")
    except Exception:  # noqa: BLE001
        return False
    body = root.find(f"{W}body")
    if body is None:
        return False
    for sp in root.iter(f"{W}sectPr"):
        if sp.find(f"{W}pgSz") is None or sp.find(f"{W}pgMar") is None:
            return False
    for item in boundaries:
        heading = _norm_heading(item.get("heading", ""))
        if not heading:
            continue
        found = None
        for pel in body.iter(f"{W}p"):
            txt = _norm_heading("".join(t.text or "" for t in pel.iter(f"{W}t")))
            if txt == heading:
                found = pel
                break
        if found is not None and not _preceding_section_break(found):
            return False
    return True


_NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_FTR_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
_FTR_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"


def ensure_footer_parts(target_path: Path, params: dict) -> dict:
    """确保文档有页脚部件并绑定到每一节；可选写入 PAGE 页码。

    2026-10-18（Layer 1）：审计发现产物 `footerReference=0`、`footer 部件=0`
    ——没有页脚就没有页码，而 `PageNumberStep` 只会往"已存在的页脚"里写，
    永远补不出来。本动作补上「创建 + 绑定 + 页码域」。

    EZI-0046：``size_pt`` 不再有 9pt 兜底常量。调用方要么传**模板实测值**
    （``format_compiler`` 传模板页脚 `w:sz` 众数），要么不传——不传时新建的
    页脚部件不写 ``w:sz``，让字号由页脚样式链决定（判定层此时也拿不到模板期望
    判 N-A），绝不写一个来源不明的 9pt 冒充期望值。
    """
    fmt = (params or {}).get("format", "{PAGE}")
    align = (params or {}).get("alignment", "center")
    size_pt = (params or {}).get("size_pt")   # None = 模板无实测证据，不写字号
    jc = {"left": "left", "center": "center", "right": "right"}.get(align, "center")
    sz_xml = ""
    if size_pt is not None:
        half = str(int(round(float(size_pt) * 2)))
        sz_xml = f'<w:sz w:val="{half}"/><w:szCs w:val="{half}"/>'

    names = _part_names(target_path)
    footer_parts = [n for n in names if n.startswith("word/footer") and n.endswith(".xml")]
    created = 0
    if not footer_parts:
        new_part = "word/footer1.xml"
        xml = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:ftr xmlns:w="{_NS_W}"><w:p><w:pPr><w:jc w:val="{jc}"/></w:pPr>'
            f'<w:r><w:rPr>{sz_xml}</w:rPr>'
            f'<w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
            f'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r><w:t>1</w:t></w:r>'
            f'<w:r><w:fldChar w:fldCharType="end"/></w:r>'
            f'</w:p></w:ftr>'
        ).encode("utf-8")
        _write_single_part(target_path, new_part, xml)
        footer_parts = [new_part]
        created = 1

    # 注册 rel + Content_Types，并给缺 footerReference 的节补引用
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    rels_part = "word/_rels/document.xml.rels"
    # 极简 docx（测试夹具/部分转换产物）可能没有 rels 部件——补一个空壳再登记
    if rels_part in _part_names(target_path):
        rels = _read_part(target_path, rels_part)
    else:
        rels = etree.Element(f"{{{rel_ns}}}Relationships", nsmap={None: rel_ns})
    targets = {str(r.get("Target") or "") for r in rels.iter(f"{{{rel_ns}}}Relationship")}
    rid_by_part = {}
    used = {r.get("Id") for r in rels}
    nums = [int(x[3:]) for x in used if isinstance(x, str) and x.startswith("rId") and x[3:].isdigit()]
    nxt = (max(nums) + 1) if nums else 1
    for part in footer_parts:
        fname = part.split("/")[-1]
        if fname in targets:
            rid_by_part[part] = next((r.get("Id") for r in rels
                                      if str(r.get("Target") or "") == fname), None)
            continue
        while f"rId{nxt}" in used:
            nxt += 1
        rid = f"rId{nxt}"
        etree.SubElement(rels, f"{{{rel_ns}}}Relationship", Id=rid,
                         Type=_FTR_REL_TYPE, Target=fname)
        used.add(rid)
        rid_by_part[part] = rid
        nxt += 1
    _write_part(target_path, rels_part, rels)

    ct_part = "[Content_Types].xml"
    ct = _read_part(target_path, ct_part)
    ct_ns = ct.nsmap.get(None) or "http://schemas.openxmlformats.org/package/2006/content-types"
    for part in footer_parts:
        partname = "/" + part
        if not any(o.get("PartName") == partname for o in ct.iter(f"{{{ct_ns}}}Override")):
            etree.SubElement(ct, f"{{{ct_ns}}}Override", PartName=partname,
                             ContentType=_FTR_CONTENT_TYPE)
    _write_part(target_path, ct_part, ct)

    doc = _read_part(target_path, "word/document.xml")
    RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    rid = rid_by_part.get(footer_parts[0])
    bound = 0
    skip_first = (params or {}).get("skip_first", True)
    sects_all = list(doc.iter(f"{W}sectPr"))
    for idx, sp in enumerate(sects_all):
        if sp.find(f"{W}footerReference") is not None:
            continue
        # 封面不显示页码：多节文档跳过第 0 节
        if skip_first and idx == 0 and len(sects_all) > 1:
            continue
        ref = etree.Element(f"{W}footerReference")
        ref.set(f"{W}type", "default")
        ref.set(RID, rid)
        heads = [i for i, ch in enumerate(sp) if ch.tag.endswith("}headerReference")]
        sp.insert((heads[-1] + 1) if heads else 0, ref)
        bound += 1
    if bound:
        _write_part(target_path, "word/document.xml", doc)
    return {"ok": True, "created": created, "bound_sections": bound, "parts": footer_parts}


def _drop_orphan_hf_parts(path: Path) -> int:
    """删除无人引用的页眉/页脚部件（含其 rels 与 Content_Types 条目）。"""
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    doc = _read_part(path, "word/document.xml")
    used = set()
    for sp in doc.iter(f"{W}sectPr"):
        for ref in list(sp.findall(f"{W}headerReference")) + list(sp.findall(f"{W}footerReference")):
            used.add(ref.get(RID))
    rels_part = "word/_rels/document.xml.rels"
    if rels_part not in zipfile.ZipFile(path).namelist():
        return 0
    rels = _read_part(path, rels_part)
    drop_parts, drop_ids = set(), set()
    for rel in list(rels):
        target = str(rel.get("Target") or "")
        if not (target.startswith("header") or target.startswith("footer")):
            continue
        if rel.get("Id") in used:
            continue
        drop_ids.add(rel.get("Id"))
        drop_parts.add("word/" + target.lstrip("/").replace("word/", ""))
        rels.remove(rel)
    if not drop_parts:
        return 0
    _write_part(path, rels_part, rels)
    # Content_Types 去条目
    ct = _read_part(path, "[Content_Types].xml")
    for o in list(ct.iter(f"{{{ct_ns}}}Override")):
        if (o.get("PartName") or "").lstrip("/") in drop_parts:
            ct.remove(o)
    _write_part(path, "[Content_Types].xml", ct)
    # 重打包去掉部件
    tmp = str(path) + ".orphan.tmp"
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename in drop_parts:
                continue
            zout.writestr(item, zin.read(item.filename))
    import os
    os.replace(tmp, str(path))
    return len(drop_parts)


_ROLE_HEADER_LABELS = {
    "摘要": "摘要", "Abstract": "ABSTRACT", "ABSTRACT": "ABSTRACT",
    "目录": "目录", "参考文献": "参考文献", "References": "参考文献",
    "REFERENCES": "参考文献", "致谢": "致谢", "附录": "附录",
}


def _section_role_texts(path: Path) -> list[tuple[int, str]]:
    """按 sectPr 文档顺序切节，返回 [(节序号, 角色)]，长度 == sectPr 数。

    2026-10-18 事故：此前按 body 一级子节点数节，而表格内/段内 sectPr 不在
    一级子节点里 → 节序号与 doc.iter(sectPr) 错位，「致谢」写进了封面节。
    现在统一用「文档顺序 + sectPr 计数」，与 sects[si] 一一对应。
    """
    doc = _read_part(path, "word/document.xml")
    body = doc.find(f"{W}body")
    if body is None:
        return []
    first: dict[int, str] = {}
    idx = 0
    for el in body.iter():
        if el.tag == f"{W}sectPr":
            idx += 1
            continue
        if el.tag != f"{W}p" or idx in first:
            continue
        # 目录条目也长得像角色标题（"致谢…12"）——必须排除，否则封面/目录节
        # 被误判成「致谢」（2026-10-18 真机：封面节页眉被写成「致谢」）
        ppr = el.find(f"{W}pPr")
        style_id = ""
        if ppr is not None and ppr.find(f"{W}pStyle") is not None:
            style_id = ppr.find(f"{W}pStyle").get(f"{W}val") or ""
        if style_id.upper().startswith("TOC"):
            continue
        if el.find(f".//{W}instrText") is not None or el.find(f".//{W}fldChar") is not None:
            continue
        txt = "".join(t.text or "" for t in el.iter(f"{W}t")).strip()
        if txt:
            first[idx] = txt
    total = len(list(doc.iter(f"{W}sectPr")))
    out: list[tuple[int, str]] = []
    for si in range(total):
        head = first.get(si, "")
        core = re.sub(r"[\s　]+", "", head)
        role = ""
        if core in _ROLE_HEADER_LABELS:
            role = _ROLE_HEADER_LABELS[core]
        elif core.startswith("第") and "章" in core[:6]:
            role = "正文"
        out.append((si, role))
    return out


_CAP_KINDS = {"图": "Fig", "表": "Tab", "Figure": "Fig", "Table": "Tab"}


def create_caption_cross_refs(target_path: Path, params: dict) -> dict:
    """图表题注书签 + 正文 REF 交叉引用（R1，2026-10-18）。

    只做图/表。题注段插书签；正文提及替换为 REF 域（编号改动后 F9 联动）。
    三种形态都认：同 run「图 3-1」、编号单独成 run（前 run 以「图」结尾）、
    句子尾部文字原样保留。幂等：已有 REF 的段落跳过。
    """
    path = Path(target_path)
    doc = _read_part(path, "word/document.xml")
    body = doc.find(f"{W}body")
    if body is None:
        return {"ok": False, "changed": False, "error": "no body"}

    def _p_text(el) -> str:
        return "".join(t.text or "" for t in el.iter(f"{W}t"))

    books: dict[str, str] = {}
    made = 0
    for p in body.iter(f"{W}p"):
        txt = _p_text(p).strip()
        m = re.match(r"^(图|表|Figure|Table)\s*([\dA-Za-z]+[-.]\d+)", txt)
        if not m:
            continue
        kind, num = m.group(1), m.group(2)
        name = f"_Ref{_CAP_KINDS.get(kind, 'Ref')}_{re.sub(r'[^0-9A-Za-z]+', '_', num)}"
        books[f"{kind}|{num}"] = name
        if any(name == (b.get(f"{W}name") or "") for b in p.iter(f"{W}bookmarkStart")):
            continue
        run = p.find(f".//{W}r")
        if run is None:
            continue
        start = etree.Element(f"{W}bookmarkStart")
        start.set(f"{W}id", str(900000 + made))
        start.set(f"{W}name", name)
        end = etree.Element(f"{W}bookmarkEnd")
        end.set(f"{W}id", str(900000 + made))
        run.addprevious(start)
        run.addnext(end)
        made += 1

    fields = 0

    def _field(run, t_el, num, bm, suffix) -> None:
        nonlocal fields
        begin = etree.SubElement(run, f"{W}fldChar")
        begin.set(f"{W}fldCharType", "begin")
        instr = etree.SubElement(run, f"{W}instrText")
        instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        instr.text = f" REF {bm} " + chr(92) + "h "
        sep = etree.SubElement(run, f"{W}fldChar")
        sep.set(f"{W}fldCharType", "separate")
        t2 = etree.SubElement(run, f"{W}t")
        t2.text = num
        end = etree.SubElement(run, f"{W}fldChar")
        end.set(f"{W}fldCharType", "end")
        if suffix:
            t3 = etree.SubElement(run, f"{W}t")
            t3.text = suffix
        # 注意：不要清 t_el.text —— 调用方已在 t_el 里写入「前缀 + 图表类别」。
        # 2026-10-26 实证：清空导致正文整句被吞，只剩域结果
        # （「因为图片…具体见图1-6」→「1-6」；「SNPs如表2-1所示」→「2-1所示」），
        # 交付门判 content preservation FAILED，整篇拒绝交付。
        fields += 1

    for p in body.iter(f"{W}p"):
        txt = _p_text(p).strip()
        if re.match(r"^(图|表|Figure|Table)\s*[\dA-Za-z]+[-.]\d+", txt):
            continue
        prev = ""
        for run in list(p.iter(f"{W}r")):
            if run.find(f"{W}fldChar") is not None:
                continue
            for t_el in run.findall(f"{W}t"):
                raw = t_el.text or ""
                if not raw:
                    continue
                m0 = re.match(r"^(\s*)([\dA-Za-z]+[-.]\d+)(.*)$", raw, re.S)
                if m0:
                    kind_m = re.search(r"(图|表|Figure|Table)\s*$", prev)
                    if kind_m:
                        bm = books.get(f"{kind_m.group(1)}|{m0.group(2)}")
                        if bm:
                            t_el.text = m0.group(1)
                            _field(run, t_el, m0.group(2), bm, m0.group(3))
                            prev += raw
                            break
                m1 = re.search(r"(图|表|Figure|Table)\s*([\dA-Za-z]+[-.]\d+)", raw)
                if m1:
                    bm = books.get(f"{m1.group(1)}|{m1.group(2)}")
                    if bm:
                        t_el.text = raw[:m1.start()] + m1.group(1) + " "
                        _field(run, t_el, m1.group(2), bm, raw[m1.end():])
                        prev += raw
                        break
                prev += raw

    changed = bool(made or fields)
    if changed:
        _write_part(path, "word/document.xml", doc)
    return {"ok": True, "changed": changed, "bookmarks": made, "refs": fields}


def verify_caption_cross_refs(target_path: Path, params: dict) -> bool:
    try:
        root = _read_part(Path(target_path), "word/document.xml")
        return any(True for _ in root.iter(f"{W}instrText"))
    except Exception:  # noqa: BLE001
        return False


def normalize_front_matter(target_path: Path, params: dict) -> dict:
    """前置页去页眉页脚 + 页码按角色分段（R2，2026-10-18）。

    用户实测：封面及"封面性质"的页面被加了页眉和横线，正文前的页码也不连续。
    规则（按角色，不按节序号）：
      · 封面段（第 0 节起、角色为空的那一段）= 封面/声明/授权 → 删除
        headerReference / footerReference / pgNumType（无页眉、无横线、无页码）；
      · 正文之前、角色非空的节（摘要/ABSTRACT/目录）→ 第一节 upperRoman start=1，
        其余不写 pgNumType（连续继承）；
      · 正文第一节 → decimal start=1，其余继承；
      · 致谢/参考文献（可能被排版在摘要前）按正文页码处理，不当封面。
    """
    path = Path(target_path)
    roles = _section_role_texts(path)
    if not roles:
        return {"ok": True, "changed": False}
    # 2026-11-08 实锤（期刊文档页脚全没）：期刊/技术文档的节里没有「目录/摘要」
    # 这类角色文本 → roles 全空 → cover_end 落到 len(roles) → 每一节都满足
    # 「si < cover_end 且无角色」→ 把所有节的 headerReference/footerReference
    # 全删了。封面清理只允许发生在「后面确实存在有角色的节」时。
    if not any(r for _, r in roles):
        return {"ok": True, "changed": False, "skipped": "no_role_sections"}
    doc = _read_part(path, "word/document.xml")
    sects = list(doc.iter(f"{W}sectPr"))
    cover_end = next((i for i, r in roles if r), len(roles))
    # 学校对「致谢在摘要前/后」要求不同 → 不按角色名选锚点，
    # 只按「封面之后第一个有角色的节」起罗马，正文起阿拉伯（用户 2026-10-18 定案）
    front_start = next((i for i, r in roles if r), cover_end)
    body_start = next((i for i, r in roles if r == "正文"), None)
    back_first = next((i for i, r in roles if r in ("致谢", "参考文献")), None)
    changed = 0
    roman_set = False
    decimal_set = False
    # 先清掉所有旧的 pgNumType（真机里各节残留 roman/decimal/半截 start，
    # 叠加后正文前页码根本不连续），再只写两个锚点
    for sp in sects:
        for el in list(sp.findall(f"{W}pgNumType")):
            sp.remove(el)
            changed += 1
    first_role_idx = next((i for i, _r in roles if _r), None)
    for si, role in roles:
        if si >= len(sects):
            continue
        sp = sects[si]
        # 只清「第一个有角色节之前」的封面节；没有角色节时 first_role_idx is None，
        # 这里天然不会命中（保底双保险，防止再一次把全篇页脚清空）。
        if (first_role_idx is not None and si < first_role_idx and not role
                and si < len(sects) - 1):
            for tag in ("headerReference", "footerReference", "pgNumType"):
                for el in list(sp.findall(f"{W}{tag}")):
                    sp.remove(el)
                    changed += 1
            continue
        # 致谢/参考文献若被排在前面，视为正文，不参与罗马页码
        is_front = True  # 封面之后的任何角色节都进罗马段（顺序由模板决定）
        if is_front and not roman_set:
            num = sp.find(f"{W}pgNumType")
            if num is None:
                num = etree.SubElement(sp, f"{W}pgNumType")
            num.set(f"{W}fmt", "upperRoman"); num.set(f"{W}start", "1")
            roman_set = True; changed += 1
            continue
        if si == body_start and not decimal_set:
            num = sp.find(f"{W}pgNumType")
            if num is None:
                num = etree.SubElement(sp, f"{W}pgNumType")
            num.set(f"{W}fmt", "decimal"); num.set(f"{W}start", "1")
            decimal_set = True; changed += 1
            continue

    if changed:
        _write_part(path, "word/document.xml", doc)
    # 前置节（封面之后、正文之前）必须有页脚承载 PAGE 域，否则罗马页码根本不显示
    # （2026-10-19 规范模板实测：pgNumType 设了但 ftr=0）
    try:
        ensure_footer_parts(path, {})
        doc2 = _read_part(path, "word/document.xml")
        sects2 = list(doc2.iter(f"{W}sectPr"))
        cleaned = 0
        for si, role in roles:
            if si < cover_end and not role and si < len(sects2):
                for ref in list(sects2[si].findall(f"{W}footerReference")):
                    sects2[si].remove(ref)
                    cleaned += 1
        if cleaned:
            _write_part(path, "word/document.xml", doc2)
    except Exception:  # noqa: BLE001
        pass
    return {"ok": True, "changed": changed > 0, "fixes": changed,
            "cover_end": cover_end, "front_start": front_start, "body_start": body_start}


def verify_front_matter(target_path: Path, params: dict) -> bool:
    try:
        path = Path(target_path)
        roles = _section_role_texts(path)
        sects = list(_read_part(path, "word/document.xml").iter(f"{W}sectPr"))
        front_start = next((i for i, r in roles if r in ("摘要", "ABSTRACT")), None)
        if front_start is None:
            return True
        for si, role in roles:
            cover_end = next((i for i, r in roles if r), len(roles))
            if si < cover_end and si < len(sects) and not role:
                sp = sects[si]
                if sp.findall(f"{W}headerReference") or sp.findall(f"{W}footerReference"):
                    return False
        return True
    except Exception:  # noqa: BLE001
        return False


def clear_table_cell_outline(target_path: Path, params: dict) -> dict:
    """表内段落标题残留清障（EZI-0068）——段级 ``outlineLvl=9`` 覆盖闸。

    真机实锤（journal 源·表1）：作者把表格数值单元格设成 heading 1/2
    样式（样式带 outlineLvl + numId 编号链）——产物原样保留 → Word 探针
    （numbering_integrity._WORD_PS 按 OutlineLevel 1..3 枚举
    ``doc.Paragraphs``，含表格内段）把 ``'10 MHz'`` 当标题，ListString 续
    上正文标题号 ``'7'~'10'`` → dup ×6 硬红。表格单元格不是文档标题
    结构：段级 ``outlineLvl=9``（wdOutlineLevelBodyText）覆盖样式级标题，
    Word 导航窗格/探针即不再视其为标题。**不动**样式引用/字号/视觉/
    numPr——最小侵入（suppress 家族只扫正文直挂段，表内段是其盲区）。

    保守闸：只动 ``w:tbl`` 内段落；直写 outlineLvl 0..8 或 pStyle 样式级
    outlineLvl 0..8 才写 9；>300 段熔断（防误扫畸形文档）。
    """
    target_path = Path(target_path)
    root = _read_part(target_path, "word/document.xml")
    styles_root = _read_part(target_path, "word/styles.xml")
    if root is None:
        return {"ok": True, "changed": False, "cleared": 0}
    # 样式级 outlineLvl 0..8 → 覆盖候选样式
    heading_style_ids: set[str] = set()
    if styles_root is not None:
        for st in styles_root.iter(f"{W}style"):
            sid = st.get(f"{W}styleId")
            if not sid:
                continue
            ppr = st.find(f"{W}pPr")
            ol = ppr.find(f"{W}outlineLvl") if ppr is not None else None
            if ol is None:
                continue
            try:
                lvl = int(ol.get(f"{W}val"))
            except (TypeError, ValueError):
                continue
            if 0 <= lvl <= 8:
                heading_style_ids.add(sid)
    # Word 真机语义（EZI-0068 二轮实锤）：段级 outlineLvl=9 在 Word 里=
    # 「移除直接格式」→ 回落样式级标题，探针照旧枚举。正确覆盖只有两条：
    # ①直写 outlineLvl 0..8 的表内段 → 删直写（样式级无标题 → body）；
    # ②pStyle 是标题样式（样式带 outlineLvl/编号链）→ 样式引用改**默认
    # 段落样式**（w:default=1）+ 删段级直写；字号随后由表格单元格规则
    # （format_table_cells）统一设定。默认段落样式找不到 → 该段只删直写
    # （诚实降级，样式引用不动）。
    default_para_style: str | None = None
    if styles_root is not None:
        for st in styles_root.iter(f"{W}style"):
            if st.get(f"{W}type") != "paragraph":
                continue
            dfl = st.get(f"{W}default")
            if dfl in ("1", "true"):
                default_para_style = st.get(f"{W}styleId")
                break
    cleared = 0
    fused = False
    for tbl in root.iter(f"{W}tbl"):
        for p in tbl.iter(f"{W}p"):
            if cleared >= 300:
                fused = True
                break
            ppr = p.find(f"{W}pPr")
            st = ppr.find(f"{W}pStyle") if ppr is not None else None
            sid = st.get(f"{W}val") if st is not None else None
            direct = ppr.find(f"{W}outlineLvl") if ppr is not None else None
            lvl_direct: int | None = None
            if direct is not None:
                try:
                    lvl_direct = int(direct.get(f"{W}val"))
                except (TypeError, ValueError):
                    lvl_direct = None
            style_heading = sid in heading_style_ids
            direct_heading = lvl_direct is not None and 0 <= lvl_direct <= 8
            if not (style_heading or direct_heading):
                continue
            if ppr is None:
                ppr = etree.Element(f"{W}pPr")
                p.insert(0, ppr)
            if direct is not None:
                ppr.remove(direct)
            if style_heading and default_para_style is not None and sid != default_para_style:
                st.set(f"{W}val", default_para_style)
            cleared += 1
    out = {"ok": True, "changed": cleared > 0, "cleared": cleared, "fused": fused}
    if cleared:
        _write_part(target_path, "word/document.xml", root)
    return out


def verify_table_cell_outline(target_path: Path, params: dict | None = None) -> bool:
    """表内段有效 outline（直写或样式级）不得再是 0..8。"""
    target_path = Path(target_path)
    root = _read_part(target_path, "word/document.xml")
    styles_root = _read_part(target_path, "word/styles.xml")
    if root is None:
        return True
    heading_style_ids: set[str] = set()
    if styles_root is not None:
        for st in styles_root.iter(f"{W}style"):
            sid = st.get(f"{W}styleId")
            ppr = st.find(f"{W}pPr")
            ol = ppr.find(f"{W}outlineLvl") if ppr is not None else None
            if sid and ol is not None:
                try:
                    if 0 <= int(ol.get(f"{W}val")) <= 8:
                        heading_style_ids.add(sid)
                except (TypeError, ValueError):
                    pass
    for tbl in root.iter(f"{W}tbl"):
        for p in tbl.iter(f"{W}p"):
            ppr = p.find(f"{W}pPr")
            ol = ppr.find(f"{W}outlineLvl") if ppr is not None else None
            if ol is not None:
                try:
                    lvl = int(ol.get(f"{W}val"))
                except (TypeError, ValueError):
                    lvl = None
                if lvl is not None and 0 <= lvl <= 8:
                    return False
            st = ppr.find(f"{W}pStyle") if ppr is not None else None
            sid = st.get(f"{W}val") if st is not None else None
            if sid in heading_style_ids:
                return False   # 标题样式残留（Word 视其为标题+编号链）
    return True


_STAR_INLINE_FORMAT_TERMS = ("字体", "字号", "黑体", "宋体", "楷体", "仿宋",
                              "Times New Roman", "行距", "磅", "号")


def _star_inline_format_note(text: str) -> bool:
    """EZI-0073：星号内联格式注记样例行（计算机学报实锤
    「1.1  二级标题 *字体为5号黑体*标题2」——模板作者把说明样例行设了
    outlineLvl，0071 的标题豁免让位只认 `_spec`，本谓词补星号内联形态）。

    判据：行内 *…* 片段含格式术语 + 编号/序号/附录开头 + 段长 8~120。
    真标题不含「*格式术语*」行内注记，零误伤（跨语料负检 0 命中）。
    """
    if not text:
        return False
    s = re.sub(r"\s+", "", text)
    if not (8 <= len(s) <= 120):
        return False
    for m in re.finditer(r"\*([^*]{2,40})\*", text):
        if any(k in m.group(1) for k in _STAR_INLINE_FORMAT_TERMS):
            return re.match("^[0-9一二三四五六七八九十（(]|^附录", s) is not None
    return False


def strip_instruction_paragraphs(target_path: Path, params: dict) -> dict:
    """源文档里混入的**模板说明段**不进产物（EZI-0065 A4）。

    真机案例（2026-11-21 用户测试产物）：源文稿里粘贴了模板的格式说明
    「【表说明采用黑体；表头文字采用中文】」，引擎照抄进产物版面。

    保守三闸（误删正文 = 不可挽回，宁可漏删）：
      ① 只删**整段**完全被说明包裹形态包住的段：
         a. 全段 【…】包裹（首尾就是括号，中间无其他内容）
         b. 全段 *…* 星号包裹
      ② 表格/页眉页脚/文本框/批注/带图段一律不动（只处理正文流）；
         有编号/大纲级别/标题样式的段不动（真标题豁免）。
      ③ 熔断：单文档剥除 > 20 段视为「这不是文稿而是模板」→ 全部放弃
         （changed=False），宁可漏删不误删。

    只删文字段，不做任何「把说明转成规格」的隐式应用——**源不是规范权威，
    规范模板才是**（取值铁律），说明里的格式要求不进 spec。
    """
    path = Path(target_path)
    root = _read_part(path, "word/document.xml")
    if root is None:
        return {"ok": True, "changed": False, "removed": 0, "texts": []}
    body = root.find(f"{W}body")
    if body is None:
        return {"ok": True, "changed": False, "removed": 0, "texts": []}
    WNS = W
    victims: list = []
    for p_el in list(body.findall(f"{WNS}p")):
        # ② 只处理正文流：段里带表格/图片/文本框/批注锚 → 不动
        if p_el.find(f".//{WNS}tbl") is not None or \
           p_el.find(f".//{WNS}drawing") is not None or \
           p_el.find(f".//{WNS}pict") is not None or \
           p_el.find(f".//{WNS}txbxContent") is not None or \
           p_el.find(f".//{WNS}commentRangeStart") is not None:
            continue
        # 标题豁免：编号/大纲/标题样式（EZI-0071 例外：格式说明段是语义
        # 权威——模板作者把说明行设成 Heading1 实锤存在（3-2-1 四条），真标题
        # <20 字、格式词命中 <3 永不触发谓词，故豁免只让位谓词正例）
        ppr = p_el.find(f"{WNS}pPr")
        _num = _ol = _hd = None
        if ppr is not None:
            _num = ppr.find(f"{WNS}numPr")
            _ol_el = ppr.find(f"{WNS}outlineLvl")
            _ol = _ol_el.get(f"{WNS}val") if _ol_el is not None else None
            _st = ppr.find(f"{WNS}pStyle")
            _hd = _st.get(f"{WNS}val") if _st is not None else ""
        t = "".join(x.text or "" for x in p_el.iter(f"{WNS}t")).strip()
        from earlyz.ir.analyzer import is_format_instruction_text
        _spec = bool(t) and is_format_instruction_text(t)
        # EZI-0073：星号内联注记样例行同样让位标题豁免（源段带 outlineLvl
        # 实锤——模板作者把说明行设成真标题级别）
        _star_note = _spec or _star_inline_format_note(t)
        if not _star_note:
            if _num is not None:
                continue
            if _ol is not None:
                continue
            if _hd and "eading" in _hd:
                continue
        if not t:
            continue
        s = re.sub(r"\s+", "", t)
        # ①a 全段【…】
        wrapped_br = (s.startswith("【") and s.endswith("】")
                      and s.count("【") == 1 and s.count("】") == 1 and len(s) <= 120)
        # ①b 全段 *…*（模板星号说明形态）
        wrapped_star = (s.startswith("*") and s.endswith("*")
                        and s.count("*") == 2 and 4 <= len(s) <= 160)
        # ①c（EZI-0071 第四闸）冒号/分号格式规格串（「字体：黑体；字号：三号；
        # 间距：1.5倍行距」/「存档论文打印装订…」）——analyzer 谓词同源（豁免
        # 段已判 `_spec` 复用），熔断/表格/标题豁免沿用本 action 既有闸。
        format_spec = _spec
        # ①d（EZI-0073 第五闸）星号内联格式注记样例行——豁免段已判 `_star_note`
        # （含 `_spec`），此处复用同一谓词（同源铁律，模块级
        # `_star_inline_format_note`）。
        _star_inline = _star_note and not (wrapped_br or wrapped_star or format_spec)
        if wrapped_br or wrapped_star or format_spec or _star_inline:
            victims.append((p_el, t[:60], bool(format_spec)))
    if not victims:
        return {"ok": True, "changed": False, "removed": 0, "texts": []}
    # ③ 熔断分层（EZI-0071）：包裹形态（【】/*…*）命中 >20 仍视为「疑似模板
    # 本体」全弃；格式说明段（is_format_instruction_text，跨语料 287 长段零
    # 误检校准）独立上限 60——真源粘贴 22+ 条说明（3-2-1 实锤）不该被熔断
    # 全弃，但 >60 仍疑似模板本体。
    wrapped_n = sum(1 for v in victims if not v[2])
    spec_n = sum(1 for v in victims if v[2])
    if wrapped_n > 20 or spec_n > 60:
        return {"ok": True, "changed": False, "removed": 0,
                "texts": [], "aborted": f"命中 {wrapped_n} 包裹段(限20)/{spec_n} 说明段(限60)超熔断阈值，疑似模板本体，全部保留"}
    for el, *_ in victims:
        body.remove(el)
    _write_part(path, "word/document.xml", root)
    return {"ok": True, "changed": True, "removed": len(victims),
            "texts": [x[1] for x in victims]}


def verify_strip_instruction_paragraphs(target_path: Path, params: dict,
                                        result: dict | None = None) -> dict:
    """校验：报删的段确实没了、且删的都是说明形态。"""
    r = result or {}
    if not r.get("changed"):
        return {"ok": True, "verified": True, "note": "无删除断言"}
    path = Path(target_path)
    try:
        root = _read_part(path, "word/document.xml")
    except Exception:  # noqa: BLE001
        return {"ok": False, "verified": False, "note": "重读失败"}
    if root is None:
        return {"ok": False, "verified": False, "note": "重读失败"}
    remaining = []
    for p_el in root.iter(f"{W}p"):
        t = "".join(x.text or "" for x in p_el.iter(f"{W}t")).strip()
        s = re.sub(r"\s+", "", t)
        if ((s.startswith("【") and s.endswith("】") and s.count("【") == 1)
                or (s.startswith("*") and s.endswith("*") and s.count("*") == 2)):
            remaining.append(t[:40])
    if remaining:
        return {"ok": False, "verified": False,
                "note": f"说明段仍残留 {len(remaining)} 条", "remaining": remaining[:5]}
    return {"ok": True, "verified": True, "note": f"已删 {r.get('removed', 0)} 段，版面无残留"}


def strip_cover_header_text(target_path: Path, params: dict) -> dict:
    """封面/声明节（角色为空的前两节）页眉必须无文字（R1，2026-10-18）。

    真机事故：模板绑定把模板里「致谢」页眉绑到封面节 → 交付门判
    cover polluted，整篇 7 分 45 秒后不产出。这里在建节+拆分之后，
    把前两节页眉部件的文字清空（只清文字，保留格式/边框）。

    2026-11-10 中文期刊实锤：**期刊题名区不是封面**。期刊首节页眉里是期刊
    masthead（第47卷第5期 计算机学报 … 2025年5月）、第二节是期次 running head
    ——都是用户内容。按「前两节=封面」无脑清文字会把它们清没：首页没有页眉、
    正文页只剩页码，产物看起来「页眉全丢」。判据与 writer/lint/交付门同一谓词
    （`zone_is_article_head`）：封面区是期刊题名区 → 本节直接不动。
    """
    path = Path(target_path)
    # 期刊题名区豁免（与 lint/writer/交付门同一谓词）
    try:
        from earlyz.ir.analyzer import analyze_document, zone_is_article_head
        from earlyz.spec.decompiler import find_cover_end

        _ir = analyze_document(path)
        _ce = int(find_cover_end(_ir) or 0)
        if _ce > 0:
            _roles = [p.semantic_role for p in _ir.paragraphs[:_ce]
                      if (getattr(p, "text", "") or "").strip()]
            if zone_is_article_head(_roles):
                return {"ok": True, "changed": False, "cleared_parts": 0,
                        "skipped": "article_head"}
    except Exception:  # noqa: BLE001 — 分析失败按原行为（宁清不漏）
        pass
    RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    rels_part = "word/_rels/document.xml.rels"
    if rels_part not in _part_names(path):
        return {"ok": True, "changed": False}
    rels = _read_part(path, rels_part)
    rid2 = {r.get("Id"): str(r.get("Target") or "") for r in rels.iter(f"{PR}Relationship")}
    doc = _read_part(path, "word/document.xml")
    sects = list(doc.iter(f"{W}sectPr"))
    cleared = 0
    for si in (0, 1):
        if si >= len(sects):
            continue
        targets = {rid2.get(r.get(RID)) for r in sects[si].findall(f"{W}headerReference")}
        for tgt in [t for t in targets if t and "header" in t]:
            part = "word/" + tgt.lstrip("/").replace("word/", "")
            if part not in _part_names(path):
                continue
            root = _read_part(path, part)
            hit = False
            for t_el in root.iter(f"{W}t"):
                if (t_el.text or "").strip():
                    t_el.text = ""
                    hit = True
            if hit:
                _write_part(path, part, root)
                cleared += 1
    return {"ok": True, "changed": cleared > 0, "cleared_parts": cleared}


def verify_cover_header_clean(target_path: Path, params: dict) -> bool:
    try:
        path = Path(target_path)
        # 期刊题名区豁免（与 strip_cover_header_text 同一谓词）：期刊首页页眉
        # 是 masthead/running head（用户内容），不清 = 正确，不是失败。
        try:
            from earlyz.ir.analyzer import analyze_document, zone_is_article_head
            from earlyz.spec.decompiler import find_cover_end

            _ir = analyze_document(path)
            _ce = int(find_cover_end(_ir) or 0)
            if _ce > 0:
                _roles = [p.semantic_role for p in _ir.paragraphs[:_ce]
                          if (getattr(p, "text", "") or "").strip()]
                if zone_is_article_head(_roles):
                    return True
        except Exception:  # noqa: BLE001
            pass
        RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"
        rels = _read_part(path, "word/_rels/document.xml.rels")
        rid2 = {r.get("Id"): str(r.get("Target") or "") for r in rels.iter(f"{PR}Relationship")}
        sects = list(_read_part(path, "word/document.xml").iter(f"{W}sectPr"))
        for si in (0, 1):
            if si >= len(sects):
                continue
            for r in sects[si].findall(f"{W}headerReference"):
                tgt = rid2.get(r.get(RID))
                if not tgt:
                    continue
                part = "word/" + tgt.lstrip("/").replace("word/", "")
                if part not in _part_names(path):
                    continue
                txt = "".join(x.text or "" for x in _read_part(path, part).iter(f"{W}t")).strip()
                if txt:
                    return False
        return True
    except Exception:  # noqa: BLE001
        return False


def apply_section_header_text(target_path: Path, params: dict) -> dict:
    """按节角色写入页眉文字（R1，2026-10-18）。

    模板页眉里本来就写着「摘要 / ABSTRACT / 目录 / 参考文献 / 致谢 / 附录」，
    但 apply_template_header_footer 只搬格式不搬文字，产物页眉全是空段落。
    本动作只使用**模板原文**（调用方从模板 header part 抽出来传进来），
    含占位符（XX/XXX/XXXX 等样例值）的一律跳过，不编造用户信息。
    """
    role_texts = {k: v for k, v in ((params or {}).get("role_texts") or {}).items() if v}
    if not role_texts:
        return {"ok": True, "changed": False, "filled": 0, "skipped": "no_role_texts"}
    if any(re.search(r"X{2,}", str(v)) for v in role_texts.values()):
        # 模板样例名（如「2024届XXXX专业毕业设计」）不落盘
        role_texts = {k: v for k, v in role_texts.items() if not re.search(r"X{2,}", str(v))}
        if not role_texts:
            return {"ok": True, "changed": False, "filled": 0, "skipped": "placeholder_only"}

    path = Path(target_path)
    RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    rels_part = "word/_rels/document.xml.rels"
    if rels_part not in _part_names(path):
        return {"ok": False, "changed": False, "error": "缺 document.xml.rels"}
    rels = _read_part(path, rels_part)
    rid2target = {r.get("Id"): str(r.get("Target") or "") for r in rels.iter(f"{PR}Relationship")}

    roles = _section_role_texts(path)
    doc = _read_part(path, "word/document.xml")
    sects = list(doc.iter(f"{W}sectPr"))
    # 同一角色出现在多节时无法判断该填哪一节（2026-10-18 事故：连续分节
    # 让「致谢」命中 5 个节）→ 该角色整体跳过，宁可不填也不污染
    from collections import Counter as _Counter
    _role_count = _Counter(r for _, r in roles if r)
    filled = 0
    for si, role in roles:
        if not role or role not in role_texts or si >= len(sects):
            continue
        if _role_count.get(role, 0) > 1:
            continue
        # 封面/声明节的角色是空串，天然不会进这里（不再用「前两节」这种位置假设）
        want = role_texts[role]
        sp = sects[si]
        refs = sp.findall(f"{W}headerReference")
        targets = [rid2target.get(r.get(RID)) for r in refs]
        targets = [t for t in targets if t and "header" in t]
        # 只写「本节独占」的部件：与封面/其它节共用的部件一改就污染（事故根因）
        _owners: dict[str, set[int]] = {}
        for _i, _sp in enumerate(sects):
            for _r in _sp.findall(f"{W}headerReference"):
                _t = rid2target.get(_r.get(RID))
                if _t:
                    _owners.setdefault(_t, set()).add(_i)
        targets = [t for t in targets if len(_owners.get(t, set())) <= 1]
        for tgt in targets:
            part = "word/" + tgt.lstrip("/").replace("word/", "")
            if part not in _part_names(path):
                continue
            root = _read_part(path, part)
            paras = list(root.iter(f"{W}p"))
            if not paras:
                continue
            first = paras[0]
            cur = "".join(x.text or "" for x in first.iter(f"{W}t")).strip()
            if cur == want:
                continue
            for run in list(first.findall(f"{W}r")):
                first.remove(run)
            run = etree.SubElement(first, f"{W}r")
            t = etree.SubElement(run, f"{W}t")
            t.text = want
            _write_part(path, part, root)
            filled += 1
    return {"ok": True, "changed": filled > 0, "filled": filled, "roles": [r for _, r in roles]}


def verify_section_header_text(target_path: Path, params: dict) -> bool:
    """复查：需要填文字的节，页眉第一段必须等于期望文字。"""
    try:
        want_map = {k: v for k, v in ((params or {}).get("role_texts") or {}).items() if v}
        if not want_map:
            return True
        path = Path(target_path)
        RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"
        rels = _read_part(path, "word/_rels/document.xml.rels")
        rid2target = {r.get("Id"): str(r.get("Target") or "") for r in rels.iter(f"{PR}Relationship")}
        roles = _section_role_texts(path)
        sects = list(_read_part(path, "word/document.xml").iter(f"{W}sectPr"))
        for si, role in roles:
            if not role or role not in want_map or si >= len(sects):
                continue
            for ref in sects[si].findall(f"{W}headerReference"):
                tgt = rid2target.get(ref.get(RID))
                if not tgt:
                    continue
                part = "word/" + tgt.lstrip("/").replace("word/", "")
                if part not in _part_names(path):
                    continue
                root = _read_part(path, part)
                first = next(iter(root.iter(f"{W}p")), None)
                cur = "" if first is None else "".join(x.text or "" for x in first.iter(f"{W}t")).strip()
                if cur != want_map[role]:
                    return False
        return True
    except Exception:  # noqa: BLE001
        return False


def split_section_parts(target_path: Path, params: dict) -> dict:
    """把各节共用的页眉/页脚部件拆成每节独立部件（幂等），并清掉孤儿部件。

    页眉绑定通常发生在建节之后，所有节会指向同一个 header1.xml；
    不拆开的话「只改封面页眉」会波及全文（2026-10-18 真机事故）。
    """
    cloned = _clone_parts_per_section(target_path)
    dropped = _drop_orphan_hf_parts(target_path)
    return {"ok": True, "cloned": cloned, "dropped": dropped,
            "changed": bool(cloned or dropped)}


def verify_section_parts(target_path: Path, params: dict) -> bool:
    """有 page 的节，其页眉/页脚部件应两两不同（每节独立）。"""
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    try:
        with zipfile.ZipFile(target_path) as z:
            doc = etree.fromstring(z.read("word/document.xml"))
            rels = etree.fromstring(z.read("word/_rels/document.xml.rels"))
    except Exception:  # noqa: BLE001
        return False
    rid2t = {r.get("Id"): r.get("Target") for r in rels.iter(f"{{{rel_ns}}}Relationship")}
    # 同一节 default+even 指向同一部件是允许的；跨节共用才算没拆开
    owners: dict[str, set[int]] = {}
    for si, sp in enumerate(doc.iter(f"{W}sectPr")):
        for ref in sp.findall(f"{W}headerReference"):
            t = rid2t.get(ref.get(RID))
            if t:
                owners.setdefault(t, set()).add(si)
    return all(len(v) <= 1 for v in owners.values())


def verify_footer_refs(target_path: Path, params: dict) -> bool:
    """每一节都必须有 footerReference，且页脚部件存在。"""
    try:
        names = _part_names(target_path)
        if not any(n.startswith("word/footer") and n.endswith(".xml") for n in names):
            return False
        doc = _read_part(target_path, "word/document.xml")
    except Exception:  # noqa: BLE001
        return False
    sects = list(doc.iter(f"{W}sectPr"))
    return bool(sects) and all(sp.find(f"{W}footerReference") is not None for sp in sects)


def verify_headers(target_path: Path, params: dict) -> bool:
    return _verify_header_footer_parts(target_path, params, "header")


_FOOTER_PART_RE = re.compile(r"^word/footer\d*\.xml$")


def footer_page_size_pt(template_path) -> Optional[float]:
    """模板页脚 **PAGE 段落有效字号**（磅）——与判定层**同一个锚**（EZI-0067）。

    EZI-0046 旧口径 = 模板全部 footer part 全部 ``w:sz`` 直写众数：页脚混排
    **文字**字号被当页码字号（模板页脚无直写、字号走 pStyle 样式链时众数
    None → 创建路裸段落 docDefaults 10.5/12/14pt）。20261121 基线 10 红实锤
    （人社部/水利部/租赁/汕尾/立项/上交/重庆科技/江南大学等）。EZI-0064 判定层
    已把期望锚重构为「页脚 PAGE 段落样式链有效字号」（``_hf_page_size``），引擎
    侧必须同锚（EZI-0045 同源铁律）——直接 import 判定层的实现，不另写一套。

    返回 ``None`` = 模板页脚无 PAGE 域或 PAGE 段样式链无字号证据：此时判定层
    判 N-A（无期望值），调用方不得用常量兜底，保持目标文档原值。
    """
    try:
        from earlyz.validation.layout_fidelity import _load_docx, _hf_page_size
        r = _hf_page_size(_load_docx(str(template_path)), "footer")
        return r.get("size")
    except Exception:  # noqa: BLE001 — 模板读不出来 = 无证据
        return None


def _set_footer_page_size(target_path: Path, size_pt: float) -> dict:
    """把页脚字号写成统一值：**全部** footer part 的**全部** run（含空壳 run）。

    页脚与页眉的语义差异（EZI-0046 A4）：页眉有「不迁模板页眉文字」红线，字号只认
    PAGE 段有效字号；页脚没有这条红线——页脚通常就是页码本身，判定层又是按「全部
    footer part 全部 w:sz 的众数」取样。所以这里：
      · 有文字的段、空段、只有 rPr 的空壳 run、``fldChar``/``instrText`` 域 run 全写；
      · 已存在的段落标记 ``pPr/rPr/sz``（它也进众数）同步改成同值；
      · 已存在的 ``szCs`` 跟随（不新增元素，避免无谓的包 diff）。
    只改值不改结构，重复跑 delta=0（幂等）。
    """
    half = str(int(round(float(size_pt) * 2)))
    names = [n for n in _part_names(target_path) if _FOOTER_PART_RE.match(n)]
    changed = 0
    parts_changed = 0
    runs = 0
    for name in names:
        root = _read_part(target_path, name)
        delta = 0
        for run in root.iter(f"{W}r"):
            rpr = run.find(f"{W}rPr")
            if rpr is None:
                rpr = etree.Element(f"{W}rPr")
                run.insert(0, rpr)
            el = rpr.find(f"{W}sz")
            if el is None:
                el = etree.Element(f"{W}sz")
                _rpr_insert_ordered(rpr, el)
            if el.get(f"{W}val") != half:
                el.set(f"{W}val", half)
                delta += 1
            runs += 1
            szcs = rpr.find(f"{W}szCs")
            if szcs is not None and szcs.get(f"{W}val") != half:
                szcs.set(f"{W}val", half)
                delta += 1
        for ppr in root.iter(f"{W}pPr"):
            prpr = ppr.find(f"{W}rPr")
            if prpr is None:
                continue
            for tag in ("sz", "szCs"):
                el = prpr.find(f"{W}{tag}")
                if el is not None and el.get(f"{W}val") != half:
                    el.set(f"{W}val", half)
                    delta += 1
        if delta:
            _write_part(target_path, name, root)
            changed += delta
            parts_changed += 1
    return {"ok": True, "size_pt": float(size_pt), "parts": len(names),
            "runs": runs, "changed": changed, "parts_changed": parts_changed}


def format_footers(target_path: Path, params: dict) -> dict:
    """页脚格式化：共享段格式通道 + 页脚专属的「全 run 字号」写入口径。

    ``properties.font_size_pt`` 不再走共享通道（那条通道只作用于**有文字**的段落，
    PAGE 域常挂在「只有 rPr 没有文字」的空壳 run 上 → 清一色的 9pt 残留留着），
    改由 ``_set_footer_page_size`` 写到每个 footer part 的每个 run。
    """
    props = dict((params or {}).get("properties") or {})
    size_pt = props.pop("font_size_pt", None)
    if props:
        info = _format_header_footer_parts(
            target_path, {**params, "properties": props}, "footer")
    else:
        info = {"ok": True, "footers": 0, "parts_touched": 0, "properties": {}}
    if size_pt is not None:
        sized = _set_footer_page_size(Path(target_path), float(size_pt))
        info["page_size"] = sized
        info["footers"] = int(info.get("footers") or 0) + sized["changed"]
        if sized["changed"]:
            info["parts_touched"] = int(info.get("parts_touched") or 0) + sized["parts_changed"]
    return info


def verify_footers(target_path: Path, params: dict) -> bool:
    return _verify_header_footer_parts(target_path, params, "footer")


# ══════════════════════════════════════════════════════════════════
# 组1 包完整性（§7.1 P0，2026-08-16 C 批收口）
# ══════════════════════════════════════════════════════════════════

# CT_* 子元素顺序表（Word 严格 schema，§7.1 normalize_ooxml_order）
_CT_PPR_ORDER = PPR_SEQ  # 真源 earlyz/tools/ooxml_order.py（2026-09-28 收敛）
_CT_RPR_ORDER = RPR_SEQ  # 真源 earlyz/tools/ooxml_order.py（2026-09-28 收敛）
_CT_SECTPR_ORDER = ["headerReference","footerReference","footnotePr","endnotePr",
    "type","pgSz","pgMar","paperSrc","pgBorders","lnNumType","pgNumType","cols",
    "formProt","vAlign","noEndnote","titlePg","textDirection","bidi","rtlGutter",
    "docGrid","printerSettings"]
_CT_TBLPR_ORDER = ["tblStyle","tblpPr","tblOverlap","bidiVisual",
    "tblStyleRowBandSize","tblStyleColBandSize","tblW","jc","tblCellSpacing",
    "tblInd","tblBorders","shd","tblLayout","tblCellMar","tblLook"]
_CT_TCPR_ORDER = ["cnfStyle","tcW","gridSpan","hMerge","vMerge","tcBorders","shd",
    "noWrap","tcMar","textDirection","tcFitText","vAlign","hideMark"]
_CT_TRPR_ORDER = TRPR_SEQ  # 真源 earlyz/tools/ooxml_order.py（2026-09-28 收敛）

_CT_ORDER = {
    f"{W}pPr": _CT_PPR_ORDER, f"{W}rPr": _CT_RPR_ORDER,
    f"{W}sectPr": _CT_SECTPR_ORDER, f"{W}tblPr": _CT_TBLPR_ORDER,
    f"{W}tcPr": _CT_TCPR_ORDER, f"{W}trPr": _CT_TRPR_ORDER,
}


def _normalize_child_order(elem: etree._Element, order: list[str]) -> int:
    """按 CT 顺序表重排 elem 的直接子元素；返回调整数。未知子元素保持相对
    位置放末尾（Word 对未知元素宽容，放尾最安全）。"""
    idx_map = {f"{W}{name}": i for i, name in enumerate(order)}
    children = list(elem)
    changed = 0
    keyed = []
    for i, ch in enumerate(children):
        k = idx_map.get(ch.tag)
        keyed.append((k if k is not None else len(idx_map) + i, i, ch))
    keyed.sort(key=lambda t: (t[0], t[1]))
    ordered = [ch for _, _, ch in keyed]
    if ordered != children:
        for ch in list(elem):
            elem.remove(ch)
        for ch in ordered:
            elem.append(ch)
        changed = 1
    return changed




# ══════════════════════════════════════════════════════════════════
# 组2 页码体系（§7.1 P0，2026-08-16 C 批收口）
# ══════════════════════════════════════════════════════════════════

_PGNUM_FORMATS = {"decimal", "upperRoman", "lowerRoman", "upperLetter",
                  "lowerLetter", "arabic", "none"}


def _sectpr_by_index(root, section_index: int):
    sects = list(root.iter(f"{W}sectPr"))
    if not sects:
        raise ValueError("no sectPr found")
    if section_index < 0:
        return sects
    if section_index >= len(sects):
        raise ValueError(f"section_index {section_index} out of range ({len(sects)})")
    return [sects[section_index]]




# ══════════════════════════════════════════════════════════════════
# 组3 段落细节（§7.1 P0，2026-08-16 C 批收口）
# ══════════════════════════════════════════════════════════════════

_TAB_VALS = {"left", "center", "right", "decimal", "bar", "clear", "num"}


def _locate_for_edit(path: Path, selector: dict) -> tuple[etree._Element, list]:
    """定位段落（复用 _locate_paragraphs + styles map）。"""
    root = _read_part(path, "word/document.xml")
    styles_map = _read_styles_map(path)
    paras = _locate_paragraphs(root, selector, styles_map)
    if not paras:
        raise ValueError(f"no paragraph matched selector: {selector}")
    return root, paras




# ══════════════════════════════════════════════════════════════════
# 组4 页眉页脚变体（§7.1 P0，2026-08-16 C 批收口）
# ══════════════════════════════════════════════════════════════════



# ══════════════════════════════════════════════════════════════════
# 组5 页面边框（§7.1 P0，2026-08-16 C 批收口）
# ══════════════════════════════════════════════════════════════════



# ══════════════════════════════════════════════════════════════════
# P1-A 表格几何（§7.2 P1，2026-08-16 收口）
# ══════════════════════════════════════════════════════════════════

def _locate_tables(root, table_index: int = -1) -> list:
    tables = list(root.iter(f"{W}tbl"))
    if not tables:
        raise ValueError("no tables found")
    if table_index < 0:
        return tables
    if table_index >= len(tables):
        raise ValueError(f"table_index {table_index} out of range ({len(tables)})")
    return [tables[table_index]]




# ══════════════════════════════════════════════════════════════════
# P1-B 图片块（§7.2 P1，2026-08-16 收口）
# ══════════════════════════════════════════════════════════════════

WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _drawing_containers(root):
    """body 级含 drawing / OLE object / VML pict 的段落。"""
    body = root.find(f"{W}body")
    if body is None:
        return []
    out = []
    for p in body.iter(f"{W}p"):
        if any(a.tag == f"{W}drawing" for a in p.iterancestors()):
            continue  # 文本框内
        if (p.find(f".//{W}drawing") is not None
                or p.find(f".//{W}object") is not None
                or p.find(f".//{W}pict") is not None):
            out.append(p)
    return out


def _cover_paragraphs(root, style_names=None) -> list:
    """封面区段落元素（role 批量动作之外的动作据此跳过封面）。

    图片块规范化（normalize_figure_block）/图片缩放（fit_media_to_page）等
    动作按「含图」定位，不走 role selector，因此享不到 _locate_paragraphs 的
    封面跳过。2026-09-25 真机：封面「姓 名：张鑫佳」一行（校徽锚定图所在段）
    被图片动作强制居中 → 封面信息栏位移。封面的图片/段落一律不动。

    返回**元素引用列表**（不是 id 集合）：lxml 代理对象会被回收复用地址，
    id() 判等在两次遍历之间不可靠（实测 fit_media_to_page 漏判封面）。
    """
    body = root.find(f"{W}body")
    if body is None:
        return []
    boundary = _compute_cover_boundary(body, style_names)
    if boundary <= 0:
        return []
    out: list = []
    seen = 0
    for p in body.iter(f"{W}p"):
        if any(a.tag == f"{W}drawing" for a in p.iterancestors()):
            continue
        if seen >= boundary:
            break
        out.append(p)
        seen += 1
    return out


def normalize_figure_block(path, params: dict) -> dict:
    """图片块规范化：居中 + keepLines/keepNext + 去掉会裁图的固定行距。

    2026-08-19 用户截图「图没完全显示」：OLE/Visio 段被 body 规则打上
    line=460 exact（23磅）+ 首行缩进 → Word 按行高裁切，只露出一条。
    图段必须清 exact 行距和首行缩进。keepNext 让图题跟图走，避免
    「图在本页、标题序号掉到下页」。短序号标题（（8）（9））keepNext
    跟随后续正文，避免孤悬页底。
    """
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    changed = 0
    media = _drawing_containers(root)
    cover_paras = _cover_paragraphs(root, _read_styles_map(path))
    media = [p for p in media if not any(p is c for c in cover_paras)]   # 封面图段不动
    media_set = set(id(p) for p in media)
    for p in media:
        ppr = p.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            p.insert(0, ppr)
        jc = ppr.find(f"{W}jc")
        if jc is None:
            jc = etree.Element(f"{W}jc")
            _ppr_insert_ordered(ppr, jc)
        if jc.get(f"{W}val") != "center":
            jc.set(f"{W}val", "center")
            changed += 1
        for tag in ("keepLines", "keepNext"):
            if ppr.find(f"{W}{tag}") is None:
                el = etree.Element(f"{W}{tag}")
                _ppr_insert_ordered(ppr, el)
                changed += 1
        # 清会裁图的固定行距 / 首行缩进
        sp = ppr.find(f"{W}spacing")
        if sp is not None and sp.get(f"{W}lineRule") == "exact":
            if f"{W}line" in sp.attrib:
                del sp.attrib[f"{W}line"]
            if f"{W}lineRule" in sp.attrib:
                del sp.attrib[f"{W}lineRule"]
            changed += 1
        ind = ppr.find(f"{W}ind")
        if ind is not None:
            for a in ("firstLine", "firstLineChars"):
                key = f"{W}{a}"
                if key in ind.attrib:
                    del ind.attrib[key]
                    changed += 1
        nxt = p.getnext()
        if nxt is not None and nxt.tag == f"{W}p":
            nt = "".join(t.text or "" for t in nxt.iter(f"{W}t")).strip()
            if re.match(r"^[图表]\s*\d", nt) and len(nt) <= 50:
                nppr = nxt.find(f"{W}pPr")
                if nppr is None:
                    nppr = etree.Element(f"{W}pPr")
                    nxt.insert(0, nppr)
                if nppr.find(f"{W}keepNext") is None:
                    # 图题跟随后续？不——图题应跟图（已由本段 keepNext）。
                    # 这里给图题加 keepLines 防自身拆行即可。
                    kl = etree.Element(f"{W}keepLines")
                    _ppr_insert_ordered(nppr, kl)
                    changed += 1
    # 短序号小标题（（8）权限数据实体）与后文绑在一起
    for p in root.iter(f"{W}p"):
        if id(p) in media_set:
            continue
        t = "".join(x.text or "" for x in p.iter(f"{W}t")).strip()
        if not re.match(r"^[（(]\d+[）)]\s*\S", t) or len(t) > 40:
            continue
        ppr = p.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            p.insert(0, ppr)
        if ppr.find(f"{W}keepNext") is None:
            el = etree.Element(f"{W}keepNext")
            _ppr_insert_ordered(ppr, el)
            changed += 1
    fw = (params or {}).get("figure_width") or {}
    if fw.get("ref") == "column" and fw.get("ratio"):
        try:
            ratio = float(fw["ratio"])
        except (TypeError, ValueError):
            ratio = 0
        if ratio:
            geoms = _sect_geoms(root)
            for p in media:
                for ext in p.iter(f"{WP}extent"):
                    try:
                        cx = int(ext.get("cx") or 0)
                    except (TypeError, ValueError):
                        continue
                    tw = cx / _EMU_PER_TWIP
                    if tw < 800:
                        continue
                    ref_w = _ref_width_for_element(p, geoms, "column")
                    if not ref_w:
                        continue
                    if tw < 0.45 * ref_w:
                        continue
                    target = ratio * ref_w
                    if abs(tw - target) / max(target, 1) <= 0.08:
                        continue
                    scale = target / tw
                    cy = int(ext.get("cy") or 0)
                    ext.set("cx", str(int(cx * scale)))
                    if cy:
                        ext.set("cy", str(int(cy * scale)))
                    changed += 1
    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "figures": len(media), "touched": changed}


def verify_figure_block(path, params: dict) -> bool:
    """复查：图段 jc=center + keepLines（封面图段除外——封面不动）。"""
    try:
        _path = Path(path)
        root = _read_part(_path, "word/document.xml")
        cover = _cover_paragraphs(root, _read_styles_map(_path))
        for p in _drawing_containers(root):
            if any(p is c for c in cover):
                continue
            ppr = p.find(f"{W}pPr")
            if ppr is None:
                return False
            jc = ppr.find(f"{W}jc")
            if jc is None or jc.get(f"{W}val") != "center":
                return False
            if ppr.find(f"{W}keepLines") is None:
                return False
        return True
    except Exception:
        return False


def _page_content_width(root) -> int:
    """页宽可用 EMU（第一个 sectPr pgSz w - 左右 pgMar）。
    2026-08-23 CJC 实锤：双栏文档图片被裁（415pt 图 vs 4606twips 栏宽）——
    页宽缩不了双栏内图片，双栏节必须按**栏宽**缩放。"""
    sp = next(root.iter(f"{W}sectPr"), None)
    if sp is None:
        return 0
    pgsz = sp.find(f"{W}pgSz")
    pgmar = sp.find(f"{W}pgMar")
    w = int(pgsz.get(f"{W}w")) if pgsz is not None else 12240
    l = int(pgmar.get(f"{W}left")) if pgmar is not None else 1800
    r = int(pgmar.get(f"{W}right")) if pgmar is not None else 1800
    cols = sp.find(f"{W}cols")
    if cols is not None and (cols.get(f"{W}num") or "1") != "1":
        col_els = list(cols.iter(f"{W}col"))
        if col_els:
            cw = int(col_els[0].get(f"{W}w") or 0)
            if cw:
                return cw * 635  # twips → EMU
        num = int(cols.get(f"{W}num") or 1)
        space = int(cols.get(f"{W}space") or 0)
        return ((w - l - r - space * (num - 1)) // num) * 635
    # 双栏检测（2026-08-23 CJC 实锤补）：首页（first 节）单栏、正文双栏——
    # 第一个 sectPr 是单栏会漏缩正文图；遍历全部节，任一双栏节→取栏宽
    col_widths = []
    for sp2 in root.iter(f"{W}sectPr"):
        cols2 = sp2.find(f"{W}cols")
        if cols2 is None or (cols2.get(f"{W}num") or "1") == "1":
            continue
        col_els2 = list(cols2.iter(f"{W}col"))
        if col_els2:
            cw2 = int(col_els2[0].get(f"{W}w") or 0)
            if cw2:
                col_widths.append(cw2)
        else:
            num2 = int(cols2.get(f"{W}num") or 1)
            space2 = int(cols2.get(f"{W}space") or 0)
            col_widths.append((w - l - r - space2 * (num2 - 1)) // num2)
    if col_widths:
        return min(col_widths) * 635
    return (w - l - r) * 635  # twips → EMU


def fit_media_to_page(path, params: dict) -> dict:
    """图片缩放适应页宽（§7.2 P1）：extent 按页宽等比缩放。R2。"""
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    max_w = _page_content_width(root)
    if max_w <= 0:
        return {"changed": False, "reason": "no page width"}
    # 栏宽约束（2026-10-19 期刊模板实测：双栏文档按整页宽缩放 → 图/表超栏，
    # 交付门判 image/table bbox overflow）。逐节取「栏数/栏间距」算可用宽度，
    # 段落按文档顺序归属到所在节。
    _sect_widths: list[int] = []
    for _sp in root.iter(f"{W}sectPr"):
        # 每节从自身 pgSz/pgMar/cols 直接算（单位 twips→EMU×635）。
        # 不能拿 _page_content_width 再除栏数：它返回的已是首节栏宽，会双重缩小。
        _pg, _mar = _sp.find(f"{W}pgSz"), _sp.find(f"{W}pgMar")
        _num, _space = 1, 425
        _cols = _sp.find(f"{W}cols")
        if _cols is not None:
            try:
                _num = max(1, int(_cols.get(f"{W}num") or 1))
                _space = int(_cols.get(f"{W}space") or 425)
            except (TypeError, ValueError):
                _num, _space = 1, 425
        if _pg is None or _mar is None:
            _sect_widths.append(max_w)
            continue
        try:
            _tw = (int(_pg.get(f"{W}w") or 0) - int(_mar.get(f"{W}left") or 0)
                   - int(_mar.get(f"{W}right") or 0))
        except (TypeError, ValueError):
            _tw = 0
        if _tw <= 0:
            _sect_widths.append(max_w)
            continue
        if _num > 1:
            _tw = max(1000, (_tw - _space * (_num - 1)) // _num)
        _sect_widths.append(int(_tw * 635))
    # 按「顶层块 → 所在节」单遍分配宽度。
    # 注意：不能用 id(lxml元素) 做键——代理对象临时创建、id 会被复用或查不到，
    # 结果全部退回整页宽，栏宽约束等于没做（2026-10-19 实测：291pt 的图留在 230pt 栏里）。
    _body = root.find(f"{W}body")
    _block_width: list[tuple[object, int]] = []
    if _body is not None:
        _sec_idx = 0
        for _blk in _body:
            if _blk.tag == f"{W}sectPr":
                _sec_idx += 1
                continue
            _w_here = _sect_widths[min(_sec_idx, len(_sect_widths) - 1)] if _sect_widths else max_w
            _block_width.append((_blk, _w_here))
            # 段落内嵌 sectPr = 本段是该节最后一段，后续块属于下一节
            if _blk.tag == f"{W}p":
                _ppr = _blk.find(f"{W}pPr")
                if _ppr is not None and _ppr.find(f"{W}sectPr") is not None:
                    _sec_idx += 1
    changed = 0
    cover_paras = _cover_paragraphs(root, _read_styles_map(path))
    _width_of: dict = {}
    for _blk, _w_here in _block_width:
        for _p in _blk.iter(f"{W}p"):
            _width_of[_p] = _w_here
    for p in _drawing_containers(root):
        if any(p is c for c in cover_paras):      # 封面图不缩放（封面不动）
            continue
        _cap = _width_of.get(p, max_w)
        for ext in p.iter(f"{WP}extent"):
            cx = int(ext.get("cx") or 0)
            cy = int(ext.get("cy") or 0)
            if cx > _cap:
                ratio = _cap / cx
                ext.set("cx", str(int(_cap)))
                ext.set("cy", str(int(cy * ratio)))
                changed += 1
    # VML/OLE 对象（w:object / v:shape，老式插入的图或公式）：style 里写死 pt 宽度，
    # 不受 w:drawing 缩放约束。实测 439.2pt 的图塞在 230pt 栏里 → 溢出页面 143pt。
    _V = "{urn:schemas-microsoft-com:vml}"

    def _pt_of(val: str) -> float | None:
        v = (val or "").strip().lower()
        for unit, mult in (("pt", 1.0), ("in", 72.0), ("cm", 72 / 2.54),
                           ("mm", 72 / 25.4), ("px", 0.75)):
            if v.endswith(unit):
                try:
                    return float(v[: -len(unit)]) * mult
                except ValueError:
                    return None
        try:
            return float(v) * 0.75
        except ValueError:
            return None

    for shape in list(root.iter(f"{_V}shape")) + list(root.iter(f"{W}pict")):
        style = shape.get("style") or ""
        if "width" not in style:
            continue
        parts = {}
        for item in style.split(";"):
            if ":" in item:
                k, val = item.split(":", 1)
                parts[k.strip().lower()] = val.strip()
        w_pt = _pt_of(parts.get("width", ""))
        h_pt = _pt_of(parts.get("height", ""))
        if not w_pt or w_pt <= 0:
            continue
        owner = shape
        while owner is not None and etree.QName(owner).localname != "p":
            owner = owner.getparent()
        # _page_content_width / _sect_widths 返回的是 EMU（twips×635），不是 twips
        cap_emu = _width_of.get(owner, max_w) if owner is not None else max_w
        cap_pt = cap_emu / 914400.0 * 72.0
        if cap_pt <= 0 or w_pt <= cap_pt:
            continue
        ratio = cap_pt / w_pt
        parts["width"] = f"{cap_pt:.1f}pt"
        if h_pt and h_pt > 0:
            parts["height"] = f"{h_pt * ratio:.2f}pt"
        shape.set("style", ";".join(f"{k}:{v}" for k, v in parts.items()))
        changed += 1

    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "resized": changed}


def verify_media_fit(path, params: dict) -> bool:
    """复查：所有 extent cx <= 页宽（封面图不缩放 → 同样跳过）。"""
    try:
        _path = Path(path)
        root = _read_part(_path, "word/document.xml")
        max_w = _page_content_width(root)
        cover = _cover_paragraphs(root, _read_styles_map(_path))
        for p in _drawing_containers(root):
            if any(p is c for c in cover):
                continue
            for ext in p.iter(f"{WP}extent"):
                if int(ext.get("cx") or 0) > max_w:
                    return False
        return True
    except Exception:
        return False


def set_alt_text(path, params: dict) -> dict:
    """设置图片替代文本（§7.2 P1）：docPr descr/name。R1。"""
    path = Path(path)
    alt = params.get("alt_text") or ""
    if not alt:
        raise ValueError("set_alt_text requires alt_text")
    root = _read_part(path, "word/document.xml")
    changed = 0
    for dp in root.iter(f"{WP}docPr"):
        if dp.get("descr") != alt:
            dp.set("descr", alt)
            changed += 1
        if not dp.get("name"):
            dp.set("name", alt[:30])
    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "figures": changed}


def verify_alt_text(path, params: dict) -> bool:
    """复查：docPr descr 已设置。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        alt = params.get("alt_text") or ""
        for dp in root.iter(f"{WP}docPr"):
            if dp.get("descr") != alt:
                return False
        return True
    except Exception:
        return False



# ══════════════════════════════════════════════════════════════════
# P1-C 域与书签（§7.2 P1，2026-08-16 收口）
# ══════════════════════════════════════════════════════════════════

_REF_INSTR_RE = re.compile(r"^REF\s+(\S+)")
_REF_CACHED_NUM_RE = re.compile(r"^\[\s*(\d+)\s*\]$")


def _field_pairs(p) -> list:
    """段内域序列解析：返回 [(域指令, 缓存结果文本), ...]。

    缓存文本取 separate→end 之间、本层 fldChar 的 w:t；域指令拼接全部
    instrText（Word 更新域后会把指令拆成多段，见 _para_pageref_target）。
    """
    seq = []
    for r in p.iter(f"{W}r"):
        for el in r.iter():
            tag = etree.QName(el).localname
            if tag == "fldChar":
                seq.append(("f", el.get(f"{W}fldCharType")))
            elif tag == "instrText":
                seq.append(("i", el.text or ""))
            elif tag == "t":
                seq.append(("t", el.text or ""))
    out = []
    i = 0
    while i < len(seq):
        k, v = seq[i]
        if k == "f" and v == "begin":
            instr, cached, depth, state = "", [], 0, "instr"
            i += 1
            while i < len(seq):
                k2, v2 = seq[i]
                if k2 == "i":
                    instr += v2
                elif k2 == "f" and v2 == "begin":
                    depth += 1
                elif k2 == "f" and v2 == "separate" and depth == 0:
                    state = "cached"
                elif k2 == "f" and v2 == "end":
                    if depth == 0:
                        break
                    depth -= 1
                elif k2 == "t" and state == "cached":
                    cached.append(v2)
                i += 1
            out.append((instr.strip(), "".join(cached).strip()))
        i += 1
    return out


def _bib_entries_for_refs(root) -> list:
    """参考文献条目段列表（简化自 format_references 定位：跳过目录
    hyperlink 段，遇 back-matter 其它标题截断）。属于引用的目标域。"""
    body = root.find(f"{W}body")
    if body is None:
        return []
    from earlyz.ir.analyzer import _back_heading_role

    paras = list(body.iter(f"{W}p"))
    hidx = -1
    for i, p in enumerate(paras):
        if p.find(f"{W}hyperlink") is not None:
            continue  # 目录条目里的「参考文献」不算节标题
        t = "".join(x.text or "" for x in p.iter(f"{W}t")).strip()
        if t and _back_heading_role(t) == "refs_heading":
            hidx = i
            break
    if hidx < 0:
        return []
    entries = []
    for p in paras[hidx + 1:]:
        t = "".join(x.text or "" for x in p.iter(f"{W}t")).strip()
        if not t:
            continue
        if _back_heading_role(t) in ("bio_heading", "ack_heading", "appendix_heading"):
            break
        if t.startswith("（") or "示例" in t:
            continue
        if p.find(f"{W}hyperlink") is not None:
            continue
        ppr = p.find(f"{W}pPr")
        if ppr is not None and ppr.find(f"{W}outlineLvl") is not None:
            break
        entries.append(p)
    return entries


def _ensure_ref_list_numbering(path: Path) -> int:
    """在 numbering.xml 补一个「[%1] 」十进制编号定义，返回新 numId。

    无 numbering.xml（罕见）→ 返回 0（跳过编号恢复；书签重建不受影响）。
    abstractNum 必须插在所有 num 之前（schema 顺序），num 追加末尾。
    """
    try:
        nroot = _read_part(path, "word/numbering.xml")
    except Exception:  # noqa: BLE001 — 部件缺失
        return 0
    max_abs, max_num = -1, 0
    for a in nroot.iter(f"{W}abstractNum"):
        try:
            max_abs = max(max_abs, int(a.get(f"{W}abstractNumId") or -1))
        except (TypeError, ValueError):
            pass
    for a in nroot.iter(f"{W}num"):
        try:
            max_num = max(max_num, int(a.get(f"{W}numId") or 0))
        except (TypeError, ValueError):
            pass
    abs_el = etree.Element(f"{W}abstractNum")
    abs_el.set(f"{W}abstractNumId", str(max_abs + 1))
    lvl = etree.SubElement(abs_el, f"{W}lvl")
    lvl.set(f"{W}ilvl", "0")
    st = etree.SubElement(lvl, f"{W}start")
    st.set(f"{W}val", "1")
    nf = etree.SubElement(lvl, f"{W}numFmt")
    nf.set(f"{W}val", "decimal")
    lt = etree.SubElement(lvl, f"{W}lvlText")
    lt.set(f"{W}val", "[%1]")
    su = etree.SubElement(lvl, f"{W}suff")
    su.set(f"{W}val", "space")
    first_num = next((c for c in nroot if etree.QName(c).localname == "num"), None)
    if first_num is not None:
        first_num.addprevious(abs_el)
    else:
        nroot.append(abs_el)
    num_el = etree.Element(f"{W}num")
    num_el.set(f"{W}numId", str(max_num + 1))
    ai = etree.SubElement(num_el, f"{W}abstractNumId")
    ai.set(f"{W}val", str(max_abs + 1))
    nroot.append(num_el)
    _write_part(path, "word/numbering.xml", nroot)
    return max_num + 1


def _repair_reference_fields(root, path: Path) -> dict:
    """断链引用书签重建 + 书目编号恢复（2026-10-08 真机事故）。

    王福勋样本根因：25 处 REF 域中 11 处目标书签丢失（Word 更新域/导出
    后显示「错误!未找到引用源」），其余 14 处目标在参考文献列表上但列表
    无编号（引用域因书目列表无编号而在更新后显示「0」）。修复：①按域缓存显示值「[N]」在
    书目第 N 条重建同名书签（[n]→第 n 条映射经 14 处工作域实证）；②列表
    被引用域依赖且尚无自动编号时，补建 [%1] 编号定义并挂到各条目（与
    format_references 的「已有自动编号不写字面序号」护栏兼容）。
    """
    result = {"bookmarks_rebuilt": [], "list_numbered": False, "unrepaired": []}
    entries = _bib_entries_for_refs(root)
    if not entries:
        return result

    bm_names, max_bm_id = set(), 0
    for bs in root.iter(f"{W}bookmarkStart"):
        bm_names.add(bs.get(f"{W}name"))
        try:
            max_bm_id = max(max_bm_id, int(bs.get(f"{W}id") or 0))
        except (TypeError, ValueError):
            pass

    # ① 断链 REF：缓存 [N] → 第 N 条重建同名书签（包裹条目全文，与工作域同构）
    next_id = max_bm_id + 1
    for p in root.iter(f"{W}p"):
        for instr, cached in _field_pairs(p):
            m = _REF_INSTR_RE.match(instr)
            if not m:
                continue
            tgt = m.group(1)
            if tgt in bm_names:
                continue
            cm = _REF_CACHED_NUM_RE.match(cached)
            n = int(cm.group(1)) if cm else 0
            if n < 1 or n > len(entries):
                result["unrepaired"].append(tgt)
                continue
            ent = entries[n - 1]
            bs = etree.Element(f"{W}bookmarkStart")
            bs.set(f"{W}id", str(next_id))
            bs.set(f"{W}name", tgt)
            pos = 1 if (len(ent) and etree.QName(ent[0]).localname == "pPr") else 0
            ent.insert(pos, bs)
            be = etree.SubElement(ent, f"{W}bookmarkEnd")
            be.set(f"{W}id", str(next_id))
            next_id += 1
            bm_names.add(tgt)
            result["bookmarks_rebuilt"].append(tgt)

    # ② 编号恢复：仅当条目被引用域依赖（原有或刚重建）且尚无自动编号
    ref_names = set()
    for p in root.iter(f"{W}p"):
        for instr, _c in _field_pairs(p):
            m = _REF_INSTR_RE.match(instr)
            if m:
                ref_names.add(m.group(1))
    referenced = False
    for ent in entries:
        for bs in ent.iter(f"{W}bookmarkStart"):
            if bs.get(f"{W}name") in ref_names:
                referenced = True
                break
        if referenced:
            break

    def _auto_numbered(p) -> bool:
        ppr = p.find(f"{W}pPr")
        if ppr is None:
            return False
        numpr = ppr.find(f"{W}numPr")
        if numpr is None:
            return False
        nm = numpr.find(f"{W}numId")
        return nm is None or nm.get(f"{W}val") != "0"

    if referenced and not any(_auto_numbered(p) for p in entries):
        num_id = _ensure_ref_list_numbering(path)
        if num_id:
            for ent in entries:
                ppr = ent.find(f"{W}pPr")
                if ppr is None:
                    ppr = etree.Element(f"{W}pPr")
                    ent.insert(0, ppr)
                if ppr.find(f"{W}numPr") is None:
                    numpr = etree.Element(f"{W}numPr")
                    ilvl = etree.SubElement(numpr, f"{W}ilvl")
                    ilvl.set(f"{W}val", "0")
                    nid = etree.SubElement(numpr, f"{W}numId")
                    nid.set(f"{W}val", str(num_id))
                    _ppr_insert_ordered(ppr, numpr)
            result["list_numbered"] = True
    return result


def repair_fields(path, params: dict) -> dict:
    """域结构完整性修复（§7.2 P1 repair_fields）：fldChar begin/separate/end
    配对；缺 end 的域在段尾补 end。检测已有（delivery 层 2.5），此动作做修复。
    R2。2026-10-08 扩展：断链引用修复（书签重建 + 书目编号恢复——「错误!
    未找到引用源」/「0」两类损坏的根因，见 _repair_reference_fields）。"""
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    fixed = 0
    for p in root.iter(f"{W}p"):
        fldchars = list(p.iter(f"{W}fldChar"))
        if not fldchars:
            continue
        begin_seen = False
        has_end = False
        last_instr_run = None
        for fc in fldchars:
            t = fc.get(f"{W}fldCharType")
            if t == "begin":
                begin_seen = True
                has_end = False
            elif t == "end":
                if begin_seen:
                    has_end = True
                    begin_seen = False
            elif t == "separate":
                pass
        # 段尾有未闭合 begin → 补 end
        if begin_seen and not has_end:
            r_end = etree.SubElement(p, f"{W}r")
            fc_end = etree.SubElement(r_end, f"{W}fldChar")
            fc_end.set(f"{W}fldCharType", "end")
            fixed += 1
    ref = _repair_reference_fields(root, path)
    fmt = _ensure_ref_charformat(root)
    if fixed or ref["bookmarks_rebuilt"] or ref["list_numbered"] or fmt:
        _write_part(path, "word/document.xml", root)
    changed = bool(fixed or ref["bookmarks_rebuilt"] or ref["list_numbered"] or fmt)
    return {"changed": changed, "fields_fixed": fixed,
            "bookmarks_rebuilt": ref["bookmarks_rebuilt"],
            "list_numbered": ref["list_numbered"],
            "ref_fields_formatted": fmt,
            "unrepaired": ref["unrepaired"]}


def _ensure_ref_charformat(root) -> int:
    """给缺 `\\* CHARFORMAT` 的 REF 域指令补上（只动指向 `_Ref*` 书签的域）。

    2026-10-12 用户报缺陷：正文「第 X 章」的 REF 域在 PDF 里被改了字体字号——
    域结果默认采用被引书签文本（章标题）的字符格式。历史产物（0.1.82 及更早）
    生成的域都没有该开关，本修复顺带把它们归一，旧产物也能救。返回修复条数。
    """
    n = 0
    for instr in root.iter(f"{W}instrText"):
        t = instr.text or ""
        if "REF " not in t or "_Ref" not in t:
            continue
        upper = t.upper()
        if "CHARFORMAT" in upper or "MERGEFORMAT" in upper:
            continue
        instr.text = t.rstrip() + " \\* CHARFORMAT "
        n += 1
    return n


def verify_fields(path, params: dict) -> bool:
    """复查：无未闭合 fldChar begin。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        for p in root.iter(f"{W}p"):
            begin_seen = False
            has_end = False
            for fc in p.iter(f"{W}fldChar"):
                t = fc.get(f"{W}fldCharType")
                if t == "begin":
                    begin_seen = True
                    has_end = False
                elif t == "end" and begin_seen:
                    has_end = True
                    begin_seen = False
            if begin_seen and not has_end:
                return False
        return True
    except Exception:
        return False


def repair_bookmarks(path, params: dict) -> dict:
    """书签配对修复（§7.2 P1 repair_bookmarks）：bookmarkStart 无对应
    bookmarkEnd → 文档尾补 end（id 匹配）；bookmarkEnd 无 start → 删除。
    R2。"""
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False}
    starts = {}
    ends = set()
    for bs in root.iter(f"{W}bookmarkStart"):
        starts[bs.get(f"{W}id")] = bs
    for be in root.iter(f"{W}bookmarkEnd"):
        ends.add(be.get(f"{W}id"))
    fixed = 0
    # ① 缺 end 的 start → body 尾补 bookmarkEnd
    for bid, bs in starts.items():
        if bid not in ends:
            be = etree.SubElement(body, f"{W}bookmarkEnd")
            be.set(f"{W}id", bid)
            fixed += 1
    # ② 无 start 的 end → 删除
    for be in list(root.iter(f"{W}bookmarkEnd")):
        if be.get(f"{W}id") not in starts:
            parent = be.getparent()
            if parent is not None:
                parent.remove(be)
                fixed += 1
    if fixed:
        _write_part(path, "word/document.xml", root)
    return {"changed": fixed > 0, "bookmarks_fixed": fixed}


def verify_bookmarks(path, params: dict) -> bool:
    """复查：全部 bookmarkStart 有配对 end。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        starts = {bs.get(f"{W}id") for bs in root.iter(f"{W}bookmarkStart")}
        ends = {be.get(f"{W}id") for be in root.iter(f"{W}bookmarkEnd")}
        return starts == ends
    except Exception:
        return False



# ══════════════════════════════════════════════════════════════════
# 页眉页码节引用修复（2026-08-16 真机回归：正文绑目录页眉/参考文献绑正文
# 页眉/ABSTRACT 无页脚——源文档节引用错位，按节语义重排）
# ══════════════════════════════════════════════════════════════════

R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _section_actual_start(root, sp) -> str | None:
    """反查节尾段 sp 的实际节首文本（上一个 sectPr 段后第一个非空段）。
    unwrap sdt 后节数变化导致 all_sects[si] 可能错位——用 XML 定位防误判。"""
    body = root.find(f"{W}body")
    if body is None:
        return None
    paras = list(body)
    idx = -1
    for i, el in enumerate(paras):
        if el is sp or (el.tag == f"{W}p" and el.find(f"{W}pPr") is not None
                        and el.find(f"{W}pPr").find(f"{W}sectPr") is sp):
            idx = i
            break
    if idx < 0:
        # sp 是 sectPr 元素——定位其父段
        p = sp.getparent().getparent() if sp.getparent() is not None else None
        if p is not None:
            for i, el in enumerate(paras):
                if el is p:
                    idx = i
                    break
    if idx < 0:
        return None
    # 向前找上一个 sectPr 段
    start = 0
    for i in range(idx - 1, -1, -1):
        el = paras[i]
        if el.tag != f"{W}p":
            continue
        ppr = el.find(f"{W}pPr")
        if ppr is not None and ppr.find(f"{W}sectPr") is not None:
            start = i + 1
            break
    for i in range(start, idx):
        el = paras[i]
        if el.tag != f"{W}p":
            continue
        t = "".join(x.text or "" for x in el.iter(f"{W}t")).strip()
        if t:
            return t
    return None


def _section_types(root) -> list[dict]:
    """按 body 段顺序识别节：{index, type, start_text}。
    节边界 = 内嵌 sectPr 段；节类型 = 边界后第一个非空段文本。"""
    body = root.find(f"{W}body")
    if body is None:
        return []
    # 全树段落（含 sdt 内 TOC 段落——2026-08-18 实证：源文档目录在
    # sdt 控件内，list(body) 只数直接子级会漏掉目录节，导致节序错乱、
    # 目录节 even 页眉绑到下一节/start_text 空。repair 里 unwrap sdt
    # 后节序才与这里一致。）
    paras = list(body.iter(f"{W}p"))
    sect_paras = []
    for i, el in enumerate(paras):
        if el.tag != f"{W}p":
            continue
        ppr = el.find(f"{W}pPr")
        if ppr is not None and ppr.find(f"{W}sectPr") is not None:
            sect_paras.append(i)
    out = []
    for si, mark in enumerate(sect_paras):
        start = mark + 1
        end = sect_paras[si + 1] if si + 1 < len(sect_paras) else len(paras)
        first = ""
        for i in range(start, min(end, start + 10)):
            if paras[i].tag != f"{W}p":
                continue
            t = "".join(x.text or "" for x in paras[i].iter(f"{W}t")).strip()
            if t:
                first = t
                break
        # 节尾段 = 下一个 sectPr 段（旧逻辑 2026-08-18 恢复：节 i 的范围
        # 是 sect_paras[i]（含，节首边界）到 sect_paras[i+1]（含，节尾边界）
        # ——节 i 的节尾段 = paras[sect_paras[i+1]]。此前误改归位为
        # paras[sect_paras[i]]（节首边界段），导致 _section_actual_start
        # 反查错位、repair 把 even 页眉/页码绑到错误节、verifier 失败。
        sp_el = None
        if si + 1 < len(sect_paras):
            sp_el = paras[sect_paras[si + 1]]
        out.append({"index": si, "start_text": first, "sp_el": sp_el})
    return out


def _ordered_sections(root) -> list[dict]:
    """全部节按文档顺序：{start_text, sp}。含第一节（_section_types 会漏）。"""
    body = root.find(f"{W}body")
    if body is None:
        return []
    paras = list(body.iter(f"{W}p"))
    marks: list[tuple[int, object]] = []
    seen = set()
    for i, p in enumerate(paras):
        ppr = p.find(f"{W}pPr")
        sp = ppr.find(f"{W}sectPr") if ppr is not None else None
        if sp is not None:
            marks.append((i, sp))
            seen.add(id(sp))
    bsp = body.find(f"{W}sectPr")
    if bsp is not None and id(bsp) not in seen:
        marks.append((len(paras), bsp))
    out = []
    prev = 0
    for mark, sp in marks:
        text = ""
        for p in paras[prev:mark]:
            t = "".join(x.text or "" for x in p.iter(f"{W}t")).strip()
            if t:
                text = t
                break
        out.append({"start_text": text, "sp": sp})
        prev = mark + 1 if mark < len(paras) else mark
    return out


def _classify_section(text: str, params: dict | None = None) -> str:
    """节角色：先用当前模板页眉原文（identity_headers），再用章标题结构。

    禁止把「目录/ABSTRACT/参考文献」写成全球词表。
    """
    n = re.sub(r"[\s\u3000]+", "", text or "")
    if not n:
        return "empty"
    for item in (params or {}).get("identity_headers") or []:
        lab = re.sub(r"[\s\u3000]+", "", (item.get("text") if isinstance(item, dict) else str(item)) or "")
        if lab and (n == lab or n.startswith(lab) or lab.startswith(n)):
            return lab
    if re.match(r"^(?:第[一二三四五六七八九十百\d]+[章节]|[\d]{1,2}\s*[\u4e00-\u9fff])", n):
        return "body"
    return "other"


def _fix_header_ascii(path: Path, part: str, ascii_font: str) -> None:
    """修复 header part 的 ascii/hAnsi 字体（2026-08-18 用户彻查）。

    模板 ABSTRACT 页眉=ascii Times New Roman；产物部分页眉 ascii 缺失
    （None→Calibri 默认渲染英文变体）。对 part 内全部 run 的 rPr 补
    ascii/hAnsi；eastAsia 不动（中文仍宋体）。
    """
    root = _read_part(path, part)
    changed = False
    for r in root.iter(f"{W}r"):
        rpr = r.find(f"{W}rPr")
        if rpr is None:
            rpr = etree.SubElement(r, f"{W}rPr")
        rf = rpr.find(f"{W}rFonts")
        if rf is None:
            rf = etree.SubElement(rpr, f"{W}rFonts")
        if rf.get(f"{W}ascii") != ascii_font or rf.get(f"{W}hAnsi") != ascii_font:
            rf.set(f"{W}ascii", ascii_font)
            rf.set(f"{W}hAnsi", ascii_font)
            changed = True
    if changed:
        _write_part(path, part, root)


def _header_part_text(path, part: str) -> str:
    with zipfile.ZipFile(path) as z:
        if part not in z.namelist():
            return ""
        return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", z.read(part).decode("utf-8")))


def _has_sectpr(p) -> bool:
    ppr = p.find(f"{W}pPr")
    return ppr is not None and ppr.find(f"{W}sectPr") is not None


_VISIBLE_NONTEXT_TAGS = (f"{W}object", f"{W}drawing", f"{W}pict", f"{M}oMath", f"{M}oMathPara")


def _para_is_blank(p) -> bool:
    """段在页面上不可见：无 w:t 文本，且无非文本可见内容。

    非文本可见内容 = OLE（MathType）/ 图片 / OMML 公式，或 MACROBUTTON 域——
    MACROBUTTON 的显示文本写在域代码里（instrText），MathType 公式编号
    「(4.1)」就是这个形态。2026-09-29：此前只看 w:t → MathType 公式段被当成
    「TOC 域尾不可见段」压到 1pt，编号在 Word 里消失（真论文 8/8）。
    """
    if _para_full_text(p).strip():
        return False
    for _ in p.iter(*_VISIBLE_NONTEXT_TAGS):
        return False
    for it in p.iter(f"{W}instrText"):
        if (it.text or "").lstrip().upper().startswith("MACROBUTTON"):
            return False
    return True


def _is_pagebreak_only(p) -> bool:
    """纯分页空段：空段且含 w:br w:type='page'。无文本/无 drawing。

    2026-08-18 修正：节尾段（含 sectPr）的 br:page 也删——continuous
    节尾段的 page break 是「continuous 时代遗留」（nextPage 分节已新起
    页，再分页 = 空页 p14）。空节尾段删 br 不损失内容；有内容的节尾段
    （含文字）由 _is_empty_para=False 天然保护。
    """
    if not _is_empty_para(p):
        return False
    for br in p.iter(f"{W}br"):
        if br.get(f"{W}type") == "page":
            return True
    return False


def _para_idx(paras, target) -> int:
    for i, pp in enumerate(paras):
        if pp is target:
            return i
    return -1


_HDR_REL_TYPE = ("http://schemas.openxmlformats.org/"
                 "officeDocument/2006/relationships/header")
_HDR_CONTENT_TYPE = ("application/vnd.openxmlformats-officedocument."
                     "wordprocessingml.header+xml")


def _w_ns_fix(xml: str) -> str:
    """把子结构里的 ns0: 前缀换成 w:（Word 认任何前缀，但统一更好读）。"""
    if not xml:
        return ""
    return xml.replace("ns0:", "w:").replace('xmlns:ns0="', 'xmlns:w="')


def _create_empty_header(path: Path, ppr_xml: str, rpr_xml: str) -> str | None:
    """新建一个只带段落格式（边框/制表位/字号）的空白页眉部件。"""
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
    n = 1
    while f"word/header{n}.xml" in names:
        n += 1
    part = f"word/header{n}.xml"
    ns = W[1:-1]
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:hdr xmlns:w="{ns}"><w:p>{_w_ns_fix(ppr_xml)}'
           f'<w:r>{_w_ns_fix(rpr_xml)}'
           '<w:t xml:space="preserve"> </w:t></w:r></w:p></w:hdr>')
    try:
        root_el = etree.fromstring(xml.encode("utf-8"))
    except etree.XMLSyntaxError:
        root_el = etree.fromstring(
            ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             f'<w:hdr xmlns:w="{ns}"><w:p/></w:hdr>').encode("utf-8"))
    _write_part(path, part, root_el)
    return part


def _ensure_header_refs(path: Path, want: dict[str, str]) -> int:
    """确保每个 sectPr 都绑定了模板要求的页眉槽位（分节后新出现的节也要绑）。

    want: {ref_type: part 名}。已存在同型引用则跳过；`first` 槽位补 titlePg。
    """
    if not want:
        return 0
    rels_part = "word/_rels/document.xml.rels"
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    if rels_part not in zipfile.ZipFile(path).namelist():
        return 0
    rels = _read_part(path, rels_part)
    target2rid = {str(r.get("Target")).split("/")[-1]: r.get("Id") for r in rels
                  if _HDR_REL_TYPE in str(r.get("Type") or "")}
    doc = _read_part(path, "word/document.xml")
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    added = 0
    for ref_type, part in want.items():
        rid = target2rid.get(part.split("/")[-1])
        if not rid:
            continue
        for sp in doc.iter(f"{W}sectPr"):
            if sp.find(f"{W}headerReference[@{W}type='{ref_type}']") is None:
                ref = etree.Element(f"{W}headerReference")
                ref.set(f"{W}type", ref_type)
                ref.set(f"{{{r_ns}}}id", rid)
                heads = [i for i, ch in enumerate(sp) if ch.tag == f"{W}headerReference"]
                sp.insert((heads[-1] + 1) if heads else 0, ref)
                added += 1
            if ref_type == "first" and sp.find(f"{W}titlePg") is None:
                tp = etree.SubElement(sp, f"{W}titlePg")
                # titlePg 在 cols 之后（CT_SectPr 顺序）；canonicalize 会再校正
                cols = sp.find(f"{W}cols")
                if cols is not None:
                    sp.remove(tp)
                    sp.insert(list(sp).index(cols) + 1, tp)
    if added:
        _write_part(path, "word/document.xml", doc)
    return added


def _clone_header_part(path: Path, src_part: str, new_part: str, ref_type: str) -> str:
    """复制页眉部件并登记 rels / Content_Types，返回新的 rId。

    2026-10-18（Layer 1）：让每一节拥有**独立**页眉部件——否则 7 个节共用
    同一个 header1.xml，任何"只改封面页眉"的操作都会波及全文。
    """
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    rels_part = "word/_rels/document.xml.rels"
    data = zipfile.ZipFile(path).read(src_part)
    _write_single_part(path, new_part, data)

    rels = _read_part(path, rels_part) if rels_part in zipfile.ZipFile(path).namelist() \
        else etree.Element(f"{{{rel_ns}}}Relationships", nsmap={None: rel_ns})
    used = {r.get("Id") for r in rels}
    nums = [int(x[3:]) for x in used if isinstance(x, str) and x.startswith("rId") and x[3:].isdigit()]
    i = (max(nums) + 1) if nums else 1
    while f"rId{i}" in used:
        i += 1
    rid = f"rId{i}"
    is_footer = "/footer" in new_part
    etree.SubElement(rels, f"{{{rel_ns}}}Relationship", Id=rid,
                     Type=_FTR_REL_TYPE if is_footer else _HDR_REL_TYPE,
                     Target=new_part.split("/")[-1])
    _write_part(path, rels_part, rels)

    ct_part = "[Content_Types].xml"
    ct = _read_part(path, ct_part)
    ct_ns = ct.nsmap.get(None) or "http://schemas.openxmlformats.org/package/2006/content-types"
    partname = "/" + new_part
    if not any(o.get("PartName") == partname for o in ct.iter(f"{{{ct_ns}}}Override")):
        etree.SubElement(ct, f"{{{ct_ns}}}Override", PartName=partname,
                         ContentType=_FTR_CONTENT_TYPE if is_footer else _HDR_CONTENT_TYPE)
        _write_part(path, ct_part, ct)
    return rid


def _clone_parts_per_section(path: Path) -> int:
    """给每一节克隆一份独立的 header/footer 部件，并把该节引用改到克隆件。"""
    RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    doc = _read_part(path, "word/document.xml")
    sects = list(doc.iter(f"{W}sectPr"))
    if len(sects) <= 1:
        return 0
    rels = _read_part(path, "word/_rels/document.xml.rels")
    rid2target = {r.get("Id"): str(r.get("Target") or "") for r in rels.iter(f"{{{rel_ns}}}Relationship")}
    names = zipfile.ZipFile(path).namelist()
    counters = {
        "header": max([int(n[len("word/header"):-4]) for n in names
                       if n.startswith("word/header") and n.endswith(".xml")] or [0]),
        "footer": max([int(n[len("word/footer"):-4]) for n in names
                       if n.startswith("word/footer") and n.endswith(".xml")] or [0]),
    }
    # 每个部件被几节引用：只被本节引用的已经是私有部件，无需再克隆
    owners: dict[str, set[int]] = {}
    for si, sp in enumerate(sects):
        for tag in ("headerReference", "footerReference"):
            for ref in sp.findall(f"{W}{tag}"):
                tgt = rid2target.get(ref.get(RID))
                if tgt:
                    owners.setdefault(tgt, set()).add(si)

    cloned = 0
    for si in range(len(sects)):
        doc = _read_part(path, "word/document.xml")
        sects_now = list(doc.iter(f"{W}sectPr"))
        if si >= len(sects_now):
            break
        mapping: dict[str, str] = {}
        for tag, kind in (("headerReference", "header"), ("footerReference", "footer")):
            for ref in sects_now[si].findall(f"{W}{tag}"):
                old_rid = ref.get(RID)
                target = rid2target.get(old_rid)
                if not target:
                    continue
                if len(owners.get(target, set())) <= 1:
                    continue  # 已独占，幂等
                src = "word/" + target.lstrip("/").replace("word/", "")
                counters[kind] += 1
                new_part = f"word/{kind}{counters[kind]}.xml"
                mapping[old_rid] = _clone_header_part(path, src, new_part, tag)
        if mapping:
            _rebind_section_parts(path, si, mapping)
            cloned += len(mapping)
    return cloned


def _parts_for_sections(path: Path, sections: list[int]) -> list[str]:
    """把节序号解析成页眉/页脚部件名（按 rels 反查 Target）。"""
    rels_part = "word/_rels/document.xml.rels"
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    try:
        rels = _read_part(path, rels_part)
        rid2target = {r.get("Id"): str(r.get("Target") or "")
                      for r in rels.iter(f"{{{rel_ns}}}Relationship")}
        doc = _read_part(path, "word/document.xml")
    except Exception:  # noqa: BLE001
        return []
    sects = list(doc.iter(f"{W}sectPr"))
    out: list[str] = []
    for si in sections:
        if not isinstance(si, int) or si < 0 or si >= len(sects):
            continue
        for tag in ("headerReference", "footerReference"):
            for ref in sects[si].findall(f"{W}{tag}"):
                rid = ref.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                target = rid2target.get(rid)
                if target:
                    out.append("word/" + target.lstrip("/").replace("word/", ""))
    return sorted(set(out))


def _rebind_section_parts(path: Path, si: int, mapping: dict[str, str]) -> None:
    """把第 si 节的 header/footer 引用改到克隆出来的新 rId。"""
    doc = _read_part(path, "word/document.xml")
    sects = list(doc.iter(f"{W}sectPr"))
    if si >= len(sects):
        return
    RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    for tag in ("headerReference", "footerReference"):
        for ref in sects[si].findall(f"{W}{tag}"):
            old = ref.get(RID)
            if old in mapping:
                ref.set(RID, mapping[old])
    _write_part(path, "word/document.xml", doc)


def _write_single_part(path: Path, part: str, data: bytes) -> None:
    """新增/覆盖 zip 内单个条目（保留其余条目）。"""
    import os
    tmp = str(path) + ".clone.tmp"
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        names = set(zin.namelist())
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))
        if part not in names:
            zout.writestr(part, data)
    os.replace(tmp, str(path))


def _bind_header_part(path: Path, part: str, ref_type: str) -> bool:
    """登记 relationship + Content_Types Override + 在缺该槽位的 sectPr 上绑定。"""
    rels_part = "word/_rels/document.xml.rels"
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    if rels_part in zipfile.ZipFile(path).namelist():
        rels = _read_part(path, rels_part)
    else:
        rels = etree.Element(f"{{{rel_ns}}}Relationships", nsmap={None: rel_ns})
    # rId 必须是 rId<数字>：别处（公式转换等）会 int(rid[3:]) 解析
    used = {r.get("Id") for r in rels}
    nums = [int(x[3:]) for x in used if isinstance(x, str) and x.startswith("rId") and x[3:].isdigit()]
    i = (max(nums) + 1) if nums else 1
    while f"rId{i}" in used:
        i += 1
    rid = f"rId{i}"
    etree.SubElement(rels, f"{{{rel_ns}}}Relationship", Id=rid,
                     Type=_HDR_REL_TYPE, Target=part.split("/")[-1])
    _write_part(path, rels_part, rels)

    ct_part = "[Content_Types].xml"
    ct = _read_part(path, ct_part)
    ct_ns = ct.nsmap.get(None) or "http://schemas.openxmlformats.org/package/2006/content-types"
    partname = "/" + part
    if not any(o.get("PartName") == partname for o in ct.iter(f"{{{ct_ns}}}Override")):
        etree.SubElement(ct, f"{{{ct_ns}}}Override", PartName=partname,
                         ContentType=_HDR_CONTENT_TYPE)
        _write_part(path, ct_part, ct)

    doc = _read_part(path, "word/document.xml")
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    bound = 0
    for sp in doc.iter(f"{W}sectPr"):
        if sp.find(f"{W}headerReference[@{W}type='{ref_type}']") is not None:
            continue
        ref = etree.Element(f"{W}headerReference")
        ref.set(f"{W}type", ref_type)
        ref.set(f"{{{r_ns}}}id", rid)
        heads = [i for i, ch in enumerate(sp) if ch.tag == f"{W}headerReference"]
        sp.insert((heads[-1] + 1) if heads else 0, ref)
        bound += 1
    if bound:
        _write_part(path, "word/document.xml", doc)
    return bool(bound)


def _target_refs_by_slot(path: Path) -> dict[tuple[str, str], list[str]]:
    """目标文档 sectPr 引用表：{（kind, ref_type）→ [part 名（文档序）]}。

    2026-08-22：模板 header/footer 与目标的真实对应是**绑定槽位**（模板
    sectPr 的 even/default/first 引用 → 目标同型引用），不是文本相似——
    模板是占位符（？期 作者名等：论文题目）而目标是实值（5期 张三,李四），
    文本相似必然失败（实测 0.09）。槽位配对纯 OOXML 机制，无学校常量。
    """
    WX = W
    RX = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    out: dict[tuple[str, str], list[str]] = {}
    with zipfile.ZipFile(path) as z:
        if "word/_rels/document.xml.rels" not in z.namelist():
            return out
        rels_raw = z.read("word/_rels/document.xml.rels")
        if not rels_raw.strip():
            return out
        rels = etree.fromstring(rels_raw)
        rid2t = {r.get("Id"): r.get("Target") for r in rels}
        root = _read_part(path, "word/document.xml")
        for sp in root.iter(f"{WX}sectPr"):
            for ref in sp:
                tag = ref.tag.replace(WX, "")
                if tag not in ("headerReference", "footerReference"):
                    continue
                if not ref.tag.startswith(WX):
                    continue
                rid = ref.get(f"{{{RX}}}id")
                tgt = rid2t.get(rid or "")
                if not tgt:
                    continue
                part = tgt if tgt.startswith("word/") else "word/" + tgt.lstrip("/")
                typ = ref.get(f"{WX}type") or "default"
                kind = "header" if tag == "headerReference" else "footer"
                out.setdefault((kind, typ), []).append(part)
    return out


def apply_template_header_footer(path, params: dict) -> dict:
    """把模板 header/footer part 的段格式应用到目标对应 part（R2）。

    2026-10-17 增补：`create_missing=True` 时，若目标**完全没有页眉**，按模板
    新建只带格式的空白页眉（边框/制表位/字号），**不写任何占位文字**——页眉
    文字（卷期/作者/刊名）需要用户手工补或告诉 agent 来加。

    模板定义「怎么排」：页眉 (even/default/first) 的制表位/边框/行距/字体、
    页脚收稿日期行的 7.5pt 宋体等；目标提供「排什么」：目标 header/footer
    part 已有自己的文本（期刊名/卷期/作者/收稿日期），**只应用格式，绝不
    改动文本**（内容保护铁律）。

    配对（机制层，无学校常量）：
    1. 槽位配对：模板 part 的 (kind, ref_type) → 目标同槽位引用出的 parts
       （模板 sectPr 怎么绑，目标就怎么绑）。槽位多候选时取非空且文本最
       相似者；
    2. 降级：槽位无可配对非空 part（如模板 footer first 而目标收稿日期在
       default）→ kind 全域按文本核相似（≥0.3）找最佳；
    3. 段配对：目标段与模板段文本相似最高者应用该段格式；无配对段 →
       应用模板「段公共键」（该 part 所有段一致的格式键）。
    """
    path = Path(path)
    parts = (params or {}).get("template_parts") or []
    if not parts:
        return {"changed": False}
    kinds = ("header", "footer")
    # EZI-0067：页码段格式迁移收集器（写回后统一应用，见下）
    pending_page_fmt: list = []
    # 目标 styles.xml 的 styleId 集——页码段 pStyle 移植闸：
    # 模板页码段样式在目标存在 → 移植引用（字号随样式链，模板为权威）；
    # 不存在 → 落直写 sz 兜底（模板页码段有效字号）。
    try:
        _styles_root = _read_part(path, "word/styles.xml")
        target_style_ids = {s.get(f"{W}styleId") for s in _styles_root.iter(f"{W}style")}
    except Exception:
        target_style_ids = set()
    slots = _target_refs_by_slot(path)
    target_parts: dict[tuple, list[dict]] = {}
    with zipfile.ZipFile(path) as z:
        names = sorted(n for n in z.namelist()
                       if re.match(r"word/(header|footer)\d+\.xml$", n))
        for n in names:
            root = _read_part(path, n)
            paras = []
            for p in root.iter(f"{W}p"):
                text = "".join(t.text or "" for t in p.iter(f"{W}t"))
                co = re.sub(r"[\s\u3000]+", "", text)
                paras.append({"text_core": co, "el": p})
            if paras:
                kind = "header" if n.startswith("word/header") else "footer"
                target_parts.setdefault((kind, n), []).extend(paras)

    def _core_of(k) -> str:
        return "".join(pa["text_core"] for pa in target_parts.get(k, []))

    changed = 0
    for tpl_part in parts:
        kind = tpl_part.get("kind")
        if kind not in kinds:
            continue
        tpl_core = tpl_part.get("text_core") or ""
        # 1) 槽位配对（机制层，见 _hf_paired_target）：模板 (kind, ref_type)
        #    → 目标引用出的 part；槽位空 → kind 全域文本相似降级。
        best_key = _hf_paired_target(path, tpl_part, slots, target_parts)
        # EZI-0067：页码段 fmt 收集独立于文字配对——_hf_paired_target 规则 1
        # 故意把纯页码模板 part 判 None（2026-08-22 设计「由 normalize 机制层
        # 处理」——但字号/样式从未被处理，10 红实锤）。页码格式跟**槽位**走：
        # ①best_key 有效 → 用它；②纯页码模板 part → 槽位内取目标 part。
        page_fmt = tpl_part.get("page_para")
        if page_fmt is None:
            for _pe in (tpl_part.get("paragraphs") or []):
                if "{PAGE}" in (_pe.get("run_sig") or []):
                    page_fmt = dict(_pe)
                    break
        if page_fmt is not None:
            _bk = best_key
            if _bk is None:
                _cands = slots.get(
                    (tpl_part.get("kind"), tpl_part.get("ref_type") or "default"), []) or []
                for _cn in _cands:
                    if (tpl_part.get("kind"), _cn) in target_parts:
                        _bk = (tpl_part.get("kind"), _cn)
                        break
            if _bk is not None:
                # 只收集——写回后统一应用：桥接（_hf_apply_run_mechanisms）
                # 新写的 PAGE 域此时还在内存树，重读磁盘旧版扫不到（顺序坑）。
                pending_page_fmt.append((_bk, page_fmt))
        if best_key is None:
            continue
        paras = target_parts[best_key]
        tpl_paras = tpl_part.get("paragraphs") or []
        if not tpl_paras:
            continue
        # 模板段公共键（所有模板段一致的格式键）
        common: dict = {}
        all_keys = set()
        for tp in tpl_paras:
            all_keys |= {k for k in tp.keys()
                         if k not in ("text_core",)}
        for k in all_keys:
            vals = [tp.get(k) for tp in tpl_paras if tp.get(k) is not None]
            if len(vals) == len(tpl_paras) and len({str(v) for v in vals}) == 1:
                common[k] = vals[0]
        for idx, para in enumerate(paras):
            # 段配对：与模板某段文本核 LCS 覆盖率最高（≥0.5），否则公共键
            # 2026-08-22：模板示例段含长注释（占位符），SequenceMatcher
            # ratio 被稀释（0.12-0.30）；LCS/min_len 找段核心（1.0+）。
            pc = para["text_core"]
            matched = None
            ratio_best = 0.5
            for tp in tpl_paras:
                r = _hf_lcs_cover(pc, tp.get("text_core") or "")
                if r > ratio_best:
                    ratio_best, matched = r, tp
            fmt = matched if matched is not None else common
            if not fmt:
                continue
            if _apply_hf_format(para["el"], fmt):
                changed += 1
            # 机制层 bridge（2026-08-23 CJC 实锤）：模板 run_sig 的
            # {PAGE}/TAB/「页码」注释 → 目标同段 run（只动数字/空白/注释）。
            # **不依赖 LCS 匹配**——模板示例占位符与目标真实文本相似度
            # 常 <0.5（配对落 common），但槽位配对已保证同类型语义；bridge
            # 用「matched 段，或按段位置序对应的模板段」（第 i 目标段 ↔
            # 第 i 模板段，不足取末段）。
            mech = matched
            if mech is None and tpl_paras:
                mech = tpl_paras[min(idx, len(tpl_paras) - 1)]
            # tab_segments（精确列插入）**先**跑——bridge 的 TAB 分支只在
            # 目标段无 w:tab 时执行（2026-08-23 CJC 实锤：先 bridge 会把
            # 空格换 tab 导致 tab_segments 早退，页码行/眉行列不精确）。
            if mech is not None:
                _mech = dict(mech)
                _mech["_line_idx"] = idx
                if _hf_apply_tab_segments(para["el"], _mech):
                    changed += 1
            if mech is not None and _hf_apply_run_mechanisms(para["el"], mech):
                changed += 1
            # 模板同构清理（2026-08-23 模板 v8 实锤）：<w:tab> 后紧跟的纯
            # 空格 run 删除——模板 tab 后直接文本；tab 后空格会妨碍右对齐
            # 制表位渲染（第 2 页页眉 `2 TAB XXX TAB 20??年` 右列不跳）。
            _hf_strip_space_after_tab(para["el"])

        # EZI-0067：页码段格式迁移——模板页码段 → 目标 PAGE 域段。
        # 20261121 基线 10 红实锤：模板页码段字号/样式在 pStyle 样式链
        # （Footer sz18=9pt）或域缓存文字段（'9'/'99'）里；目标 PAGE 段
        # （新域无文字）被 texts.strip() 过滤根本不进 target_parts 池 →
        # apply 循环走不到 → 裸段落 docDefaults 10.5/12/14pt。
        # fmt 源优先级：page_para（无文字纯域段旁路）→ 池内首个带 {PAGE}
        # 的段（含域缓存 '9'/'99' 文字段）。
    if changed:
        # 写回所有被格式化的 part（_read_part 的树是副本，必须 _write_part）
        written = set()
        for (k, n), v in target_parts.items():
            if n not in written and any(pa["el"] is not None for pa in v):
                root = v[0]["el"].getroottree()
                _write_part(path, n, root.getroot())
                written.add(n)
    # ── EZI-0067：页码段格式迁移（写回后统一应用）────────────────────
    # 模板页码段字号/样式在 pStyle 样式链或域缓存段（'9'/'99'）；目标 PAGE
    # 段（桥接新写或本就存在）不进文字池 → 格式从未落上 → docDefaults。
    # 闸：①目标有同 styleId → 移植 pStyle（样式链权威，模板是规范权威）；
    # ②无 → PAGE runs 直写模板页码段有效字号（rPr 属性，绝不动域/文本）。
    pages_formatted = 0
    for _bkey, page_fmt in pending_page_fmt:
        try:
            _tgt_root = _read_part(path, _bkey[1])
        except Exception:
            continue
        _pf_sz = page_fmt.get("font_size_pt")
        _hit = False
        for _pp in _tgt_root.iter(f"{W}p"):
            if not any("PAGE" in (x.text or "") for x in _pp.iter(f"{W}instrText")):
                continue
            _apply_hf_format(_pp, page_fmt)
            _tps = page_fmt.get("tpl_pstyle")
            _done = False
            if _tps and _tps in target_style_ids:
                _ppr = _pp.find(f"{W}pPr")
                if _ppr is None:
                    _ppr = etree.Element(f"{W}pPr")
                    _pp.insert(0, _ppr)
                _ps_el = _ppr.find(f"{W}pStyle")
                if _ps_el is None:
                    _ps_el = etree.SubElement(_ppr, f"{W}pStyle")
                _ps_el.set(f"{W}val", _tps)
                if list(_ppr).index(_ps_el) != 0:
                    _ppr.remove(_ps_el)
                    _ppr.insert(0, _ps_el)
                _done = True
            elif _pf_sz is not None:
                _half = str(int(round(float(_pf_sz) * 2)))
                for _r in _pp.iter(f"{W}r"):
                    _rpr = _r.find(f"{W}rPr")
                    if _rpr is None:
                        _rpr = etree.Element(f"{W}rPr")
                        _r.insert(0, _rpr)
                    for _tag in ("sz", "szCs"):
                        _el = _rpr.find(f"{W}{_tag}")
                        if _el is None:
                            _anchor = _rpr.find(f"{W}rFonts")
                            _el = etree.Element(f"{W}{_tag}")
                            if _anchor is not None:
                                _anchor.addnext(_el)
                            else:
                                _rpr.insert(0, _el)
                        _el.set(f"{W}val", _half)
                _done = True
            if _done:
                _hit = True
        if _hit:
            _write_part(path, _bkey[1], _tgt_root)
            pages_formatted += 1
    # 槽位语义对齐（2026-08-23 CJC 实锤）：模板槽位有内容而目标槽位空 →
    # 克隆重绑；模板未定义的槽位 → 删除目标引用（继承）。机制层，无学校常量。
    align = _hf_align_slots(path, parts, (params or {}).get("template_slots"))
    # 目标完全没有页眉 → 按模板补「只带格式的空白页眉」（2026-10-17 用户拍板 A）
    created: list[dict] = []
    if (params or {}).get("create_missing") and not [k for k in target_parts if k[0] == "header"]:
        seen_types: set[str] = set()
        for tpl_part in parts:
            if tpl_part.get("kind") != "header":
                continue
            rt = str(tpl_part.get("ref_type") or "default")
            if rt in seen_types:
                continue
            paras_t = tpl_part.get("paragraphs") or []
            if not paras_t:
                continue
            best = max(paras_t,
                       key=lambda pp: (1 if pp.get("p_bdr_bottom") else 0,
                                       len(str(pp.get("ppr_extra_xml") or ""))))
            seen_types.add(rt)
            new_part = _create_empty_header(
                path, str(best.get("ppr_extra_xml") or ""), str(best.get("rpr_extra_xml") or ""))
            if not new_part:
                continue
            if _bind_header_part(path, new_part, rt):
                created.append({"part": new_part, "ref_type": rt,
                                "border": bool(best.get("p_bdr_bottom"))})

    # 分节可能在后续步骤才出现 → 这里按模板槽位把引用补到所有 sectPr（幂等）
    want: dict[str, str] = {}
    for tpl_part in parts:
        if tpl_part.get("kind") != "header":
            continue
        rt = str(tpl_part.get("ref_type") or "default")
        if rt in want:
            continue
        made = next((c["part"] for c in created if c["ref_type"] == rt), None)
        if made:
            want[rt] = made
            continue
        # 已有页眉：按槽位引用取该槽位真实部件（不要一律取第一个部件，
        # 否则 even/default/first 会全指到同一份、带边框的那份永远不出现）
        by_slot = list(slots.get(("header", rt)) or [])
        if by_slot:
            want[rt] = by_slot[0]
            continue
        same = [n for (k, n) in target_parts if k == "header"]
        if same:
            want[rt] = same[0]
    bound = _ensure_header_refs(path, want)

    return {"changed": changed > 0 or align["changed"] or bool(created) or bound > 0
            or pages_formatted > 0,
            "parts_formatted": changed + pages_formatted,
            "parts_created": int(align.get("parts_created", 0)) + len(created),
            "created": created, "refs_bound": bound}


def _replace_run_with_page_field(r, fallback: str) -> None:
    """把 run 替换为 PAGE 域链（begin/instrText/separate/result/end），
    rPr 保留在 run 首位（顺序铁律 §5）。fallback=域结果缓存（Word 更新前显示）。"""
    parent = r.getparent()
    if parent is None:
        return
    rpr = r.find(f"{W}rPr")
    nr = etree.Element(f"{W}r")
    if rpr is not None:
        nr.append(rpr)  # rPr 移到新 run 首位
    fld = etree.SubElement(nr, f"{W}fldChar")
    fld.set(f"{W}fldCharType", "begin")
    it = etree.SubElement(nr, f"{W}instrText")
    it.text = " PAGE "
    it.set(f"{W}space", "preserve")
    fld = etree.SubElement(nr, f"{W}fldChar")
    fld.set(f"{W}fldCharType", "separate")
    t = etree.SubElement(nr, f"{W}t")
    t.text = fallback or "1"
    fld = etree.SubElement(nr, f"{W}fldChar")
    fld.set(f"{W}fldCharType", "end")
    parent.replace(r, nr)


def _replace_run_with_tab(r) -> None:
    """把 run 内容清空替换为 w:tab（保留 rPr）。"""
    rpr = r.find(f"{W}rPr")
    for c in list(r):
        if c is not rpr:
            r.remove(c)
    etree.SubElement(r, f"{W}tab")


def _hf_apply_run_mechanisms(p, tpl: dict) -> bool:
    """模板段 run_sig 机制桥接（2026-08-23 CJC 页眉实锤）：
    {PAGE} → 目标段「位置比例最接近模板 PAGE 位置」的纯数字 run 替换为
    PAGE 域链（原值作 fallback；比例定位避免把「期号」当页码）；
    TAB → 目标段「位置比例最接近模板 TAB 位置」的单空格 run 替换为 w:tab；
    模板段含 {PAGE} 而目标段存在「（页码）」字面注释 run → 删除（域已接管）。
    只动数字/空白/注释 run，绝不碰中文文本 run（内容保护铁律）。"""
    sig = tpl.get("run_sig") or []
    if not sig:
        return False
    changed = False
    texts = [(r, "".join(t.text or "" for t in r.iter(f"{W}t"))) for r in p.iter(f"{W}r")]
    n_slots = max(1, len(sig) - 1)

    has_page = "{PAGE}" in sig
    has_tab = "TAB" in sig
    # 1) PAGE：位置比例最近者（2026-11-10：目标段已有 PAGE 域链 → 跳过，
    #    否则会把域结果缓存 run（纯数字）再包一层域 → 嵌套域，页码损坏）
    if has_page and not _para_has_page_field(p):
        page_frac = sig.index("{PAGE}") / n_slots
        digits = [(i, r, txt.strip()) for i, (r, txt) in enumerate(texts)
                  if txt.strip().isdigit()]
        if digits:
            nonempty = [txt.strip() for (_, txt) in texts if txt.strip()]
            best, best_d = None, 9.0
            for i, r, num in digits:
                idx = sum(1 for j, (_, txt) in enumerate(texts)
                          if j < i and txt.strip())
                frac = idx / max(1, len(nonempty) - 1)
                d = abs(frac - page_frac)
                if d < best_d:
                    best_d, best = d, (r, num)
            if best is not None:
                _replace_run_with_page_field(best[0], best[1])
                changed = True
    # 2) TAB：对模板每个 TAB 位置，找目标「位置比例最近」的单空格 run
    #    （每个 TAB 各自映射，不重复使用同一空格 run；段已有 w:tab 时跳过
    #    ——精确列插入由 _hf_apply_tab_segments 负责，2026-08-23 CJC 实锤）
    if has_tab and not any(child.tag == f"{W}tab" for r in p.iter(f"{W}r") for child in r):
        spaced = [(i, r) for i, (r, txt) in enumerate(texts) if txt == " "]
        if spaced:
            nonempty = [txt.strip() for (_, txt) in texts if txt.strip()]
            used = set()
            for ti in range(len(sig)):
                # 顺序取第 ti 个 TAB 的期望位置
                tab_positions = [j for j, s in enumerate(sig) if s == "TAB"]
                if ti >= len(tab_positions):
                    break
                tab_frac = tab_positions[ti] / n_slots
                best, best_d = None, 9.0
                for i, r in spaced:
                    if id(r) in used:
                        continue
                    idx = sum(1 for j, (_, txt) in enumerate(texts)
                              if j < i and txt.strip())
                    frac = idx / max(1, len(nonempty) - 1)
                    d = abs(frac - tab_frac)
                    if d < best_d:
                        best_d, best = d, r
                if best is not None:
                    _replace_run_with_tab(best)
                    used.add(id(best))
                    changed = True
    # 3) 「页码」注释清理（域接管后，模板无此 token）
    if has_page:
        for r, txt in texts:
            if "页码" in txt:
                core = re.sub(r"[()（）\s\u3000]", "", txt)
                if core == "页码":
                    parent = r.getparent()
                    if parent is not None:
                        parent.remove(r)
                        changed = True
    return changed


def _hf_apply_tab_segments(p, tpl: dict) -> bool:
    """模板 tab 列布局 → 目标段插入 w:tab（2026-08-23 模板 v8 精确化）：
    优先「模板组骨架（去空白+去?占位符）在目标非空白序列贪婪 LCS 定位」——
    组 i（非末组）末字后插 tab、末组首字前插 tab（落在 run 内则拆 w:t，
    rPr 克隆、文本不变=内容保护）；骨架定位失败回退比例近似（原逻辑）。

    tab_segments 为行×列 [[列1,列2..],[..]]——目标段 i 对应模板行 i。"""
    segs = tpl.get("tab_segments") or []
    if not segs:
        return False
    if isinstance(segs[0], str):
        segs = [segs]  # 兼容扁平结构
    line_idx = tpl.get("_line_idx") or 0
    lines = tpl.get("_lines") or segs
    if not lines:
        return False
    row = lines[min(line_idx, len(lines) - 1)] if isinstance(lines[0], list) else lines
    if len(row) < 2:
        return False
    if any(child.tag == f"{W}tab" for r in p.iter(f"{W}r") for child in r):
        return False
    # ---- 目标字符映射（非空白字符 → (run, 文本内下标, ch)）----
    runs = list(p.iter(f"{W}r"))
    cmap = []
    for r in runs:
        t_el = r.find(f"{W}t")
        if t_el is None:
            continue
        txt = t_el.text or ""
        for j, ch in enumerate(txt):
            if not ch.isspace():
                cmap.append((r, j, ch))
    tgt = "".join(ch for _, _, ch in cmap)

    def _skel(s: str) -> str:
        return "".join(c for c in (s or "") if c not in "?？" and not c.isspace()
                       and c != "\u3000")

    def _match(skel: str, start: int = 0):
        if not skel:
            return None
        i = start
        first = last = None
        for ch in skel:
            while i < len(tgt) and tgt[i] != ch:
                i += 1
            if i >= len(tgt):
                return None
            if first is None:
                first = i
            last = i
            i += 1
        return (first, last)

    # ---- 骨架优先定位 ----
    # 组 i（非末组）末字后、末组首字前插 tab；中间占位符组（XXX 等）匹配
    # 失败→跳过（后续组仍定位）；首组或末组失败→比例回退。
    tab_positions = []
    prev_end = 0
    first_ok = False
    last_ok = False
    for si, s in enumerate(row):
        m = _match(_skel(s), prev_end)
        if m is None:
            if si == 0 or si == len(row) - 1:
                last_ok = False
                break
            continue
        if si == 0:
            first_ok = True
        if si == len(row) - 1:
            last_ok = True
        if si < len(row) - 1:
            tab_positions.append(m[1] + 1)
            prev_end = m[1] + 1
        else:
            tab_positions.append(m[0])
    if first_ok and last_ok and tab_positions:
        inserted = 0
        for pos in sorted({p for p in tab_positions}):
            if _hf_tab_after_char(cmap, pos):
                inserted += 1
        return inserted > 0

    # ---- 比例近似回退（原逻辑）----
    total = sum(max(1, len(s)) for s in row)
    cum = []
    acc = 0
    for s in row[:-1]:
        acc += max(1, len(s))
        cum.append(acc / total)
    run_spans = []
    acc_n = 0
    for r in runs:
        txt = "".join(t.text or "" for t in r.iter(f"{W}t"))
        n = len([c for c in txt if not c.isspace()])
        run_spans.append((acc_n, acc_n + n, r))
        acc_n += n
    if acc_n < 4:
        return False
    inserted = 0
    inserts = set()
    for f in cum:
        pos = int(round(f * acc_n))
        if pos <= 0 or pos >= acc_n:
            continue
        best_r, best_d = None, 10**9
        for start, end, r in run_spans:
            d = abs(pos - end)
            if d < best_d:
                best_d, best_r = d, r
        if best_r is not None and id(best_r) not in inserts:
            nr = etree.Element(f"{W}r")
            etree.SubElement(nr, f"{W}tab")
            best_r.addnext(nr)
            inserts.add(id(best_r))
            inserted += 1
    return inserted > 0


def _hf_tab_after_char(cmap: list, pos: int) -> bool:
    """在 cmap[pos-1] 字符后插 `<w:tab/>`；若该字符之后还有同一 run 文本
    → 拆 w:t（rest 移到新 run，rPr 克隆——文本不变，内容保护铁律）。"""
    if pos <= 0 or pos > len(cmap):
        return False
    r_el, j, _ch = cmap[pos - 1]
    t_el = r_el.find(f"{W}t")
    if t_el is None:
        return False
    txt = t_el.text or ""
    if j + 1 < len(txt):
        rest = txt[j + 1:]
        t_el.text = txt[: j + 1]
        nr = etree.Element(f"{W}r")
        rpr = r_el.find(f"{W}rPr")
        if rpr is not None:
            nr.append(etree.fromstring(etree.tostring(rpr)))
        nt = etree.SubElement(nr, f"{W}t")
        nt.text = rest
        nt.set(f"{W}space", "preserve")
        tr = etree.Element(f"{W}r")
        tr.append(etree.Element(f"{W}tab"))
        r_el.addnext(tr)
        tr.addnext(nr)
        return True
    nr = etree.Element(f"{W}r")
    nr.append(etree.Element(f"{W}tab"))
    r_el.addnext(nr)
    return True


def _hf_strip_space_after_tab(p) -> int:
    """模板同构：`<w:tab/>` 后紧跟的纯空白 run 删除（2026-08-23 模板 v8
    实锤：模板 tab 后直接文本——tab 后空格妨碍右对齐制表位渲染，第 2 页
    页眉 `2 TAB XXX TAB 20??年` 右列不跳）。"""
    n = 0
    for r in list(p.iter(f"{W}r")):
        if r.find(f"{W}tab") is None:
            continue
        nxt = r.getnext()
        while nxt is not None and nxt.tag == f"{W}r":
            t_el = nxt.find(f"{W}t")
            txt = (t_el.text or "") if t_el is not None else ""
            if t_el is None or txt == "" or txt.strip() != "":
                break
            nxt.getparent().remove(nxt)
            n += 1
            nxt = r.getnext()
    return n



def _hf_text_core(s: str) -> str:
    return re.sub(r"[\s\u3000]+", "", s or "")


def _hf_content_affinity(a: str, b: str) -> float:
    """页眉页脚内容亲缘度：按结构特征重叠打分（2026-08-23 CJC 实锤）。
    模板占位符（?期/XXX/20??年）与目标真实文本的 SequenceMatcher ratio
    天然 <0.25——改用「期号/作者/题目/收稿/年份/卷号」特征重叠判别同类型
    页眉页脚（机制层判别，无学校名/刊物名常量）。"""
    score = 0.0
    for f, w in (("期", 2.0), ("：", 1.5), ("作者", 1.5), ("题目", 1.0),
                 ("年", 1.0), ("Vol", 1.0), ("收稿", 2.0), ("基金", 1.5),
                 ("卷", 1.0)):
        if f in a and f in b:
            score += w
    return score


def _hf_is_ornament(core: str) -> bool:
    """装饰段（页脚横线等）：核心文本全为破折号/连字符/底划类字符且 len≥5
    （2026-08-23 CJC 实锤：模板 footer1[0]=15×U+2014，目标克隆无此段）。"""
    if len(core) < 5:
        return False
    return all(ord(c) in (0x2014, 0x2013, 0xFF0D, 0x2015, 0x005F, 0xFF3F, 0x007E, 0x2022, 0x00B7) for c in core)


def _hf_orn_insert(path: Path, part_name: str, tpl_fmt: dict | None) -> bool:
    """把模板装饰段（页脚横线）插入 part 段首（2026-08-23 CJC 实锤）。
    文本/格式全取自模板段（取值层）；若 part 已含装饰段则跳过。"""
    root = _read_part(path, part_name)
    if any(_hf_is_ornament(_hf_text_core("".join(t.text or "" for t in p.iter(f"{W}t"))))
           for p in root.iter(f"{W}p")):
        return False
    orn_txt = tpl_fmt.get("text_core") or "" if tpl_fmt else ""
    if not _hf_is_ornament(orn_txt):
        return False
    new_p = etree.SubElement(root, f"{W}p")
    r = etree.SubElement(new_p, f"{W}r")
    t = etree.SubElement(r, f"{W}t")
    t.text = orn_txt
    t.set(f"{W}space", "preserve")
    # 段+run 全量格式应用（2026-08-23 CJC 字号实锤）：必须**先建 run 再**
    # _apply_hf_format——旧顺序时段内无 run，typed 字号（font_size_pt→sz）
    # 永不落 run，装饰段继承源稿 docDefaults sz=29 → Word 显示 14.5。
    _apply_hf_format(new_p, tpl_fmt or {})
    # 移到段首（新建默认在末尾）
    root.insert(0, new_p)
    _write_part(path, part_name, root)
    return True


def _hf_align_slots(path: Path, tpl_parts: list[dict],
                    tpl_bound_slots: list | None = None) -> dict:
    """模板槽位语义对齐（2026-08-23 CJC 实锤 4 缺陷，机制层无学校常量）：
    ① 模板 (kind, ref_type) 有内容、目标同槽位 part 文本核为空、目标其他槽位
       存在与模板文本核相似的 part → 克隆该 part（新 rId）→ 槽位引用改绑
       （解决：odd 页眉空/首页收稿日期页脚缺失——模板语义落位）；
    ② 模板该 ref_type 无引用（如 CJC 模板 default/even 无页脚）→ 目标对应
       引用删除（继承语义=无页脚），对齐模板槽位定义。
    克隆 part 必须 DEFLATED + docPr id 不冲突（头部纯文本无 docPr，
    但克隆一律走 _create_cloned_part 风格；页眉/页脚文本 part 无 shape 安全）。
    """
    changed = False
    created = 0
    zip_in = zipfile.ZipFile(path)
    rels_root = etree.fromstring(zip_in.read("word/_rels/document.xml.rels"))
    zip_in.close()
    rels: dict[str, str] = {}
    for rel in rels_root.iter(f"{PR}Relationship"):
        rels[rel.get("Id")] = rel.get("Target")
    root = _read_part(path, "word/document.xml")

    slot_refs: dict[tuple, list] = {}
    for sp in root.iter(f"{W}sectPr"):
        for r in sp.iter(f"{W}headerReference"):
            slot_refs.setdefault(("header", r.get(f"{W}type") or "default"), []).append(r)
        for r in sp.iter(f"{W}footerReference"):
            slot_refs.setdefault(("footer", r.get(f"{W}type") or "default"), []).append(r)

    def _part_core(name: str) -> str:
        try:
            with zipfile.ZipFile(path) as z:
                pr_root = etree.fromstring(z.read(name))
        except (KeyError, etree.XMLSyntaxError):
            return ""
        return _hf_text_core("".join(t.text or "" for t in pr_root.iter(f"{W}t")))

    def _all_parts(kind: str) -> list[str]:
        with zipfile.ZipFile(path) as z:
            return sorted(
                n for n in z.namelist()
                if n.startswith(f"word/{kind}") and n.endswith(".xml")
            )

    # 模板定义的槽位集合（有内容 part 的槽位）
    tpl_slots_with_content = {
        (t.get("kind"), t.get("ref_type") or "default")
        for t in tpl_parts if t.get("text_core")
    }
    tpl_all_slots = {
        (t.get("kind"), t.get("ref_type") or "default")
        for t in tpl_parts if t.get("kind") in ("header", "footer")
    }
    # 「模板定义了哪些槽位」以模板**实际绑定**为准（含空 part，2026-09-30）：
    # tpl_parts 只含有文字的 part，模板用空 part 定义的 first 槽位会被误判成
    # 「未定义」→ 目标正文节首页页眉/页脚被删（真论文第 1 章首页页眉页码消失）。
    for s in tpl_bound_slots or []:
        if isinstance(s, (list, tuple)) and len(s) == 2 and s[0] in ("header", "footer"):
            tpl_all_slots.add((s[0], s[1] or "default"))

    for slot, refs in slot_refs.items():
        kind, ref_type = slot
        # ② 模板无该槽位（未定义）→ 删除目标引用（继承语义）
        if slot not in tpl_all_slots:
            for r in refs:
                sp = r.getparent()
                if sp is not None:
                    sp.remove(r)
                    changed = True
            continue
        tpl_core = ""
        for t in tpl_parts:
            if (t.get("kind"), t.get("ref_type") or "default") == slot:
                tpl_core = t.get("text_core") or ""
        if not tpl_core:
            continue
        # ① 目标槽位 part 空 + 其他槽位有相似内容 → 克隆重绑
        for r in refs:
            target_name = rels.get(r.get(f"{R}id") or "")
            if not target_name:
                continue
            # rels Target 可能为相对路径（'header1.xml'）——归一化为包内名
            if not target_name.startswith("word/") and not target_name.startswith("/"):
                target_name = f"word/{target_name}"
            if _part_core(target_name):
                continue  # 目标槽位有内容，不需要克隆
            best = None
            for other in _all_parts(kind):
                other_core = _part_core(other)
                if not other_core or other == target_name:
                    continue
                # 特征亲缘度（占位符 vs 真实文本：ratio 失效，用特征重叠）
                score = _hf_content_affinity(other_core, tpl_core)
                if score >= 2.5 and (best is None or score > best[0]):
                    best = (score, other)
            if best is None:
                # 2026-11-10（IEEE 英文期刊实锤）：模板槽位有内容、目标槽位空、
                # 目标自己也没有相似 part——旧逻辑到此为止，模板页眉里的页码
                # 机制（PAGE 域链）永远落不进产物 → 整篇没有页码。补一层：
                # 模板该槽位的段里有 {PAGE} → 把页码域链按模板段格式写进目标
                # 这个空 part（机制层；**不写任何文字 run**——2026-10-17 拍板：
                # 页眉文字属用户内容，卷期/作者/刊名/占位说明一律不写）。
                # 模板段没有页码机制 → 维持原状（行为不变）。
                _cell = next((
                    t for t in tpl_parts
                    if (t.get("kind"), t.get("ref_type") or "default") == slot), None)
                _mech = next((
                    tp for tp in (_cell or {}).get("paragraphs") or []
                    if "{PAGE}" in (tp.get("run_sig") or [])), None)
                if _mech is not None and _hf_inject_page_mechanism(
                        path, target_name, _mech):
                    changed = True
                continue
            src_name = best[1]
            new_part, new_rid = _clone_hf_part(path, src_name, kind)
            if new_part is None:
                continue
            r.set(f"{R}id", new_rid)
            # 模板装饰段（页脚横线）补入克隆 part（2026-08-23 CJC 实锤：
            # 模板 footer1 首段=15×U+2014 横线，源稿克隆无此段）
            _tpl_cell = next((
                t for t in tpl_parts
                if (t.get("kind"), t.get("ref_type") or "default") == slot
            ), None)
            _orn = next((
                tp for tp in (_tpl_cell or {}).get("paragraphs") or []
                if _hf_is_ornament(tp.get("text_core") or "")
            ), None)
            if _orn is not None:
                _hf_orn_insert(path, new_part, _orn)
            changed = True
            created += 1

    # 模板装饰段补齐（2026-08-23 CJC 实锤）：无论是否新克隆，凡模板槽位
    # 有装饰段（页脚横线）而目标引用 part 缺失 → 补入段首（幂等）。
    if any(_hf_is_ornament(tp.get("text_core") or "")
           for t in tpl_parts for tp in (t.get("paragraphs") or [])):
        with zipfile.ZipFile(path) as z:
            rels2_root = etree.fromstring(z.read("word/_rels/document.xml.rels"))
        rels2 = {rel.get("Id"): rel.get("Target") for rel in rels2_root.iter(f"{PR}Relationship")}
        for slot, refs in slot_refs.items():
            kind2, rt2 = slot
            tpl_cell = next((
                t for t in tpl_parts
                if (t.get("kind"), t.get("ref_type") or "default") == slot), None)
            orn = next((
                tp for tp in (tpl_cell or {}).get("paragraphs") or []
                if _hf_is_ornament(tp.get("text_core") or "")), None)
            if orn is None:
                continue
            for r in refs:
                target_name = rels2.get(r.get(f"{R}id") or "")
                if not target_name:
                    continue
                if not target_name.startswith("word/") and not target_name.startswith("/"):
                    target_name = f"word/{target_name}"
                if _hf_orn_insert(path, target_name, orn):
                    changed = True
    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed, "parts_created": created}


def _para_has_page_field(p) -> bool:
    """段落里已有 PAGE 域链（instrText 指令含 PAGE）→ True。

    2026-11-10：`_hf_apply_run_mechanisms` 的 PAGE 桥接若在目标段已有 PAGE 域时
    再跑，会把域结果缓存 run（纯数字）再包一层域链 → 嵌套域，Word 更新域后
    页码损坏。判据含 NUMPAGES/PAGEREF 排除（它们不是页码域）。
    """
    for it in p.iter(f"{W}instrText"):
        t = (it.text or "").upper()
        if "PAGE" in t and "NUMPAGES" not in t and "PAGEREF" not in t:
            return True
    return False


def _hf_inject_page_mechanism(path: Path, part_name: str, tpl_para: dict) -> bool:
    """把模板页眉段的「页码机制」（PAGE 域链）写进目标空页眉 part。

    2026-11-10（IEEE 英文期刊实锤）：模板页眉有内容（页码 + 稿件号占位行）、
    目标页眉 part 为空、目标自己也没有相似 part——旧逻辑到此为止，模板的页码
    机制永远落不进产物 → 整篇没有页码。本函数只落「PAGE 域链 + 段落格式」
    （jc/行距/字号取自模板段），**不写任何文字 run**（2026-10-17 拍板：页眉
    文字属用户内容，卷期/作者/刊名/占位说明一律不写）。

    幂等/内容保护：目标 part 已有 PAGE 域 → 跳过；已有真实文字 → 不动
    （调用方已按 text_core 为空过滤）。
    """
    try:
        root = _read_part(path, part_name)
    except Exception:  # noqa: BLE001 — part 读不了就按「不做」处理
        return False
    ps = list(root.iter(f"{W}p"))
    if not ps:
        return False
    p = ps[0]
    if _para_has_page_field(p):
        return False
    fmt = {k: tpl_para.get(k) for k in (
        "jc", "line", "line_rule", "spacing_before", "spacing_after",
        "font_ascii", "font_east_asia", "font_size_pt",
        "ppr_extra_xml", "rpr_extra_xml") if tpl_para.get(k) is not None}
    if fmt:
        _apply_hf_format(p, fmt)
    r = p.find(f"{W}r")
    if r is None:
        r = etree.SubElement(p, f"{W}r")
    _replace_run_with_page_field(r, "1")
    _write_part(path, part_name, root)
    return True


def _clone_hf_part(path: Path, src_name: str, kind: str) -> tuple[str | None, str | None]:
    """克隆 header/footer part（现成 _clone_part：docPr 重编 + CT/rels 完整注册），
    返回 (part_name, rId)。"""
    new_part = _clone_part(path, src_name, kind)
    if new_part is None:
        return (None, None)
    rid = _rid_of_part(path, new_part)
    return (new_part, rid)


def _apply_hf_format(p, fmt: dict) -> bool:
    """段级应用模板格式：清空可移植属性后重放 extras + 写 typed 键。

    只动 pPr/rPr，绝不动 run 文本与域（fldChar/instrText/PAGE）。
    """
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        ppr = etree.Element(f"{W}pPr")
        p.insert(0, ppr)
    # 模板段样式基线不移植（2026-08-23 模板 v8 实锤：源稿 Header 样式
    # a5 的 jc=center 让页眉整体居中偏移（渲染实测列1 起点 343px）——模板
    # 段无 pStyle 或段级 jc 直接覆盖；移除目标 pStyle，直接属性全量重放）
    ps = ppr.find(f"{W}pStyle")
    if ps is not None:
        ppr.remove(ps)
    _apply_extra_xml(ppr, fmt.get("ppr_extra_xml"), "pPr")
    if "p_bdr_bottom" in fmt and fmt["p_bdr_bottom"]:
        # pBdr 已在 extras 里（模板段带 pBdr→ppr_extra_xml）；此处仅兜底
        if ppr.find(f"{W}pBdr") is None:
            pbdr = etree.SubElement(ppr, f"{W}pBdr")
            bot = etree.SubElement(pbdr, f"{W}bottom")
            bot.set(f"{W}val", "single")
            bot.set(f"{W}sz", "6")
            bot.set(f"{W}space", "1")
            bot.set(f"{W}color", "auto")
    _set_typed_ppr(ppr, fmt)
    for r in p.findall(f"{W}r"):
        rpr = r.find(f"{W}rPr")
        if rpr is None:
            rpr = etree.Element(f"{W}rPr")
            r.insert(0, rpr)
        _apply_extra_xml(rpr, fmt.get("rpr_extra_xml"), "rPr")
        _set_typed_rpr(rpr, fmt)
    # §5 顺序铁律（2026-08-23 CJC 实锤）：extras/typed 键拼装后 pPr/rPr
    # 子元素可能乱序（pBdr→tabs→jc→spacing→rPr），Word 对错位元素解析
    # 不稳定（页眉 pBdr 双线/tab 列布局受影响）——统一归一化到 CT 顺序
    _normalize_child_order(ppr, _CT_PPR_ORDER)
    for r in p.findall(f"{W}r"):
        rpr = r.find(f"{W}rPr")
        if rpr is not None:
            _normalize_child_order(rpr, _CT_RPR_ORDER)
    return True


def _set_typed_ppr(ppr, fmt: dict) -> None:
    """写 pPr typed 键（jc/line/line_rule/spacing/tab_pos）。"""
    jc = fmt.get("jc")
    if jc:
        el = ppr.find(f"{W}jc")
        if el is None:
            el = etree.SubElement(ppr, f"{W}jc")
        el.set(f"{W}val", jc)
    line = fmt.get("line")
    if line is not None:
        spacing = ppr.find(f"{W}spacing")
        if spacing is None:
            spacing = etree.SubElement(ppr, f"{W}spacing")
        spacing.set(f"{W}line", str(int(line)))
        lr = fmt.get("line_rule")
        if lr:
            spacing.set(f"{W}lineRule", lr)
    sb = fmt.get("spacing_before")
    sa = fmt.get("spacing_after")
    if sb is not None or sa is not None:
        spacing = ppr.find(f"{W}spacing")
        if spacing is None:
            spacing = etree.SubElement(ppr, f"{W}spacing")
        if sb is not None:
            spacing.set(f"{W}before", str(int(sb)))
        if sa is not None:
            spacing.set(f"{W}after", str(int(sa)))
    tp = fmt.get("tab_pos")
    if tp:
        tabs = ppr.find(f"{W}tabs")
        if tabs is None:
            tabs = etree.SubElement(ppr, f"{W}tabs")
        if not tabs.findall(f"{W}tab"):
            tab = etree.SubElement(tabs, f"{W}tab")
            tab.set(f"{W}pos", str(int(tp)))
            tab.set(f"{W}val", "left")


def _set_typed_rpr(rpr, fmt: dict) -> None:
    """写 rPr typed 键（字体/字号/粗斜/下划线/颜色）。"""
    ea = fmt.get("font_east_asia")
    asc = fmt.get("font_ascii")
    if ea or asc:
        rf = rpr.find(f"{W}rFonts")
        if rf is None:
            rf = etree.Element(f"{W}rFonts")
            rpr.insert(0, rf)
        if ea:
            rf.set(f"{W}eastAsia", ea)
        if asc:
            rf.set(f"{W}ascii", asc)
            rf.set(f"{W}hAnsi", asc)
    sz = fmt.get("font_size_pt")
    if sz is not None:
        half = str(int(round(float(sz) * 2)))
        for tag in ("sz", "szCs"):
            el = rpr.find(f"{W}{tag}")
            if el is None:
                el = etree.SubElement(rpr, f"{W}{tag}")
            el.set(f"{W}val", half)
    for prop, tag in (("bold", "b"), ("italic", "i"), ("small_caps", "smallCaps")):
        val = fmt.get(prop)
        if val is None:
            continue
        el = rpr.find(f"{W}{tag}")
        if val is True and el is None:
            etree.SubElement(rpr, f"{W}{tag}")
        elif val is False and el is not None:
            rpr.remove(el)
    if fmt.get("underline"):
        el = rpr.find(f"{W}u")
        if el is None:
            el = etree.SubElement(rpr, f"{W}u")
        el.set(f"{W}val", fmt["underline"] if isinstance(fmt["underline"], str) else "single")
    col = fmt.get("color")
    if col:
        el = rpr.find(f"{W}color")
        if el is None:
            el = etree.SubElement(rpr, f"{W}color")
        el.set(f"{W}val", col)


def _hf_lcs_cover(a: str, b: str) -> float:
    """最长公共子串覆盖率：LCS / min(len)。对不等长字符串（模板示例注释
    长、目标实值短）稳定；SequenceMatcher.ratio 会被长度差稀释。

    2026-08-22 修复：LCS 长度 <2 视为无意义（'附录' vs '目录' 共享单字
    '录' → 0.5 假阳性），返回 0.0。
    """
    if not a or not b:
        return 0.0
    m = SequenceMatcher(None, a, b).find_longest_match(0, len(a), 0, len(b))
    if m.size < 2 or m.size == 0:
        return 0.0
    return m.size / min(len(a), len(b))


def _hf_is_page_number_core(core: str) -> bool:
    """内容核是否纯页码（模板示例 '2222'/'1111' / 空白）：页码型 part
    由 normalize_footer_parts 机制层处理，格式配对无内容意义 → 跳过。"""
    return bool(core) and bool(re.fullmatch(r"[\d\s\u3000\-–—]*", core))


def _hf_paired_target(path: Path, tpl_part: dict, slots: dict,
                      target_parts: dict) -> tuple | None:
    """slot 配对：模板 part (kind, ref_type) → 目标非空部分名。

    规则（2026-08-22 三态，机制层无学校常量）：
    1. 纯页码/空白模板 part → None（页码型由 normalize 机制层处理）；
    2. 槽位内**唯一非空** → 直接配对（模板占位符 vs 目标实值，文本
       ratio 仅 0.09 也成立——如 CJC header2 占位符→ header4 实值）；
    3. 槽位内**多个非空** → 取 SequenceMatcher.ratio 最高者，须 ≥0.5，
       否则视为槽位无真实对应（xian 学位论文模板 header8「参考文献」
       与源稿「毕业设计」页眉 ratio 0.074 —— 不同 content 不得错配）；
    4. 槽位空/无达标 → 降级 kind 全域段落级 LCS 覆盖率 ≥0.5；
    5. 仍无 → None（跳过该模板 part）。
    """
    from difflib import SequenceMatcher
    kind = tpl_part.get("kind")
    tpl_core = tpl_part.get("text_core") or ""
    # 1) 纯页码/空白模板 part：内容核无意义（页码型），跳过配对
    if _hf_is_page_number_core(tpl_core):
        return None

    def _core_of(k):
        return "".join(pa.get("text_core") or "" for pa in target_parts.get(k, []))

    slot_keys = [k for k in target_parts
                 if k[0] == kind and k[1] in slots.get((kind, tpl_part.get("ref_type")), [])]
    nonempty = [(k, _core_of(k)) for k in slot_keys if _core_of(k)]
    if nonempty:
        if len(nonempty) == 1:
            return nonempty[0][0]
        # 多个非空：必须真实相似（同 content 不同槽位/占位符 vs 实值）。
        # 2026-08-22 严格 >0.5：'附录' vs '目录' 共享单字'录' ratio=0.5
        # 恰好等于阈值会误配（xian header11 → 目标目录页眉）。
        best_key, best_ratio = None, 0.0
        for k, core in nonempty:
            r = SequenceMatcher(None, tpl_core, core).ratio()
            if r > best_ratio:
                best_ratio, best_key = r, k
        if best_ratio > 0.5:
            return best_key
    # 降级：kind 全域按**段落级** LCS 覆盖率取 part 最大分（≥0.5）。
    # 2026-08-22：模板 first footer（收稿日期）常对应目标 default footer
    # （同一内容、不同槽位）——part 级拼接 core 被模板占位符/破折号
    # 稀释（SequenceMatcher.ratio 实测 0.12-0.30），LCS/min_len 对
    # 段核心稳定（“收稿日期：年-月-日；最终修改稿收到日期”1.0）。
    # 机制层：内容相同但槽位不同的页眉页脚配对，无学校常量。
    best_ratio = 0.0
    best_key = None
    for k in target_parts:
        if k[0] != kind:
            continue
        paras = target_parts.get(k, [])
        if not any((pa.get("text_core") or "") for pa in paras):
            continue
        r = max((_hf_lcs_cover(tpl_core, pa.get("text_core") or "")
                 for pa in paras), default=0.0)
        if r > best_ratio:
            best_ratio, best_key = r, k
    return best_key if best_ratio >= 0.5 else None


def verify_template_header_footer(path, params: dict) -> bool:
    """复查（2026-08-22 建，2026-09-25 修多对一口径）：
    - 目标文档无该 kind 的 part（源稿本来没有页眉/页脚元素）→ 跳过；
    - 有 → 每个配对成功的模板 part 都要求目标 part 落位其格式；**多个
      模板 part 降级配对到同一目标 part 时按目标归并判定**——apply 以
      「最后应用者」落位（后覆盖前），要求每个模板 part 都精确命中在
      数学上不可能（2026-09-25 张鑫佳实锤：16 个模板 part 与 3 个目标
      part 多对一，两次全量编译都被误拦）。目标 part 命中「任一次曾应用
      于它的模板格式」即通过；所有应用都未命中（改了却没生效）= 真失败
      仍 False。"""
    try:
        path = Path(path)
        parts = (params or {}).get("template_parts") or []
        if not parts:
            return True
        kinds = {"header", "footer"}
        slots = _target_refs_by_slot(path)
        target_parts: dict[tuple, list[dict]] = {}
        with zipfile.ZipFile(path) as z:
            names = sorted(n for n in z.namelist()
                           if re.match(r"word/(header|footer)\d+\.xml$", n))
            for n in names:
                root = _read_part(path, n)
                paras = []
                for p in root.iter(f"{W}p"):
                    text = "".join(t.text or "" for t in p.iter(f"{W}t"))
                    co = re.sub(r"[\s\u3000]+", "", text)
                    ppr = p.find(f"{W}pPr")
                    paras.append({
                        "text_core": co, "el": p,
                        "ppr_extra": _extra_xml_from_lxml(ppr, "pPr") if ppr is not None else None,
                    })
                if paras:
                    kind = "header" if n.startswith("word/header") else "footer"
                    target_parts.setdefault((kind, n), []).extend(paras)
        if not target_parts:
            return True  # 目标无任何页面元素 → 无可配对，跳过
        hit_by_target: dict = {}
        for tpl_part in parts:
            kind = tpl_part.get("kind")
            if kind not in kinds:
                continue
            tpl_core_v = tpl_part.get("text_core") or ""
            # 纯页码/空白模板 part：页码型，机制层处理，跳过配对要求
            if _hf_is_page_number_core(tpl_core_v):
                continue
            # 目标中完全无该 kind → 跳过该模板 part
            if not any(k[0] == kind for k in target_parts):
                continue
            key = _hf_paired_target(path, tpl_part, slots, target_parts)
            if key is None:
                # 2026-08-22：模板 part 在目标无对应内容（如 xian 模板
                # 「附录」页眉 vs 目标无附录节 header）→ apply 同样跳过
                # （同一配对函数），不视为失败；强行 False 会造成永远
                # 回滚（模板有 16 个 part、目标只有 9 个页面元素）。
                continue
            # 落位检查：配对 part 至少一段的 pPr extras 与模板某段一致
            tpl_paras = tpl_part.get("paragraphs") or []
            if not tpl_paras:
                continue
            # 落位检查：配对 part 至少一段满足「模板有的键必须匹配」
            # （extras 相等 + 模板段非 None typed 键逐项相等）。模板段
            # 没有的键不强求——目标段原有的 jc/tabs 等不因未写而算失败。
            def _typed_val(ppr, tag, attr):
                el = ppr.find(tag)
                return el.get(f"{W}{attr}") if el is not None else None

            def _para_matches_tpl(pa, tp):
                pel = pa.get("el")
                ppr = pel.find(f"{W}pPr") if pel is not None else None
                if ppr is None:
                    return False
                extras = _extra_xml_from_lxml(ppr, "pPr") or ""
                if (tp.get("ppr_extra_xml") or "") != extras:
                    return False
                checks = (
                    ("jc", f"{W}jc", "val"),
                    ("line", f"{W}spacing", "line"),
                    ("line_rule", f"{W}spacing", "lineRule"),
                    ("spacing_before", f"{W}spacing", "before"),
                    ("spacing_after", f"{W}spacing", "after"),
                )
                for key, tag, attr in checks:
                    want = tp.get(key)
                    if want is None:
                        continue
                    if _typed_val(ppr, tag, attr) != str(want):
                        return False
                if tp.get("tab_pos") is not None:
                    tabs = ppr.find(f"{W}tabs")
                    first = tabs.find(f"{W}tab") if tabs is not None else None
                    if first is None or first.get(f"{W}pos") != str(tp.get("tab_pos")):
                        return False
                return True

            applied = [pa for pa in target_parts[key]
                       if any(_para_matches_tpl(pa, tp) for tp in tpl_paras)]
            hit_by_target[key] = hit_by_target.get(key, False) or bool(applied)
        missing = [k for k, hit in hit_by_target.items() if not hit]
        return not missing
    except Exception:
        return False


def repair_section_references(path, params: dict) -> dict:
    """按节语义重排页眉页码引用（R2）。

    修复：①正文/结论节 → 正文页眉（含「：」或「届」的 header part）+
    PAGE 页脚 + decimal start=1；②ABSTRACT → 补 upperRoman 页脚；
    ③参考文献节 → 「参考文献」页眉（缺失则创建）+ 页脚。
    """
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False}
    # TOC sdt 展开：目录控件（docPartObj）内容提升为普通段落——Word 对
    # 未更新的 TOC sdt 渲染空白（2026-08-17 实证：p11 空页= TOC 控件页），
    # 展开后目录以普通段落渲染且节内空段可被边界清理覆盖。
    # 注意：必须保留缩进正确（此段代码在 repair_section_references 内）
    for sdt in list(body.iter(f"{W}sdt")):
        parent = sdt.getparent()
        content = sdt.find(f"{W}sdtContent")
        if content is None:
            continue
        idx = list(parent).index(sdt)
        for child in list(content):
            parent.insert(idx, child)
            idx += 1
        parent.remove(sdt)

    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
    header_texts = {}
    for n in names:
        if re.match(r"word/header\d+\.xml", n):
            header_texts[n] = _header_part_text(path, n)
    params = params or {}
    identity = list(params.get("identity_headers") or [])
    identity_norms = {
        re.sub(r"[\s\u3000]+", "", (h.get("text") if isinstance(h, dict) else str(h)) or "")
        for h in identity
    }
    identity_norms.discard("")
    # 正文页眉 = 目标里「不是模板 identity 页眉」的非空 header（换模板词表跟着变）
    body_headers = [n for n, t in header_texts.items()
                    if t.strip() and re.sub(r"[\s\u3000]+", "", t) not in identity_norms]
    body_headers.sort()
    # 首选：目标**正文节自己绑定**的页眉（2026-09-30 真论文页眉错绑根因②）。
    # 上面的候选表按文件名字典序取首尾（header15 < header17 < header3），会把
    # 目录页眉、封面残留页眉（「2024届…」）绑到正文。页眉文本属「排什么」，
    # 目标正文节已有自己的 even/default 绑定时直接沿用；没有才走旧候选表。
    own_body = _target_body_header_parts(path, root, params, header_texts, identity_norms)
    body_footers = sorted([n for n in names if re.match(r"word/footer\d+\.xml", n)])

    # ascii 字体：按模板该页眉自己的 ascii，不写死 ABSTRACT→TNR / 目录→宋体
    for n, t in header_texts.items():
        nt = re.sub(r"[\s\u3000]+", "", t)
        for h in identity:
            if not isinstance(h, dict):
                continue
            if nt == re.sub(r"[\s\u3000]+", "", h.get("text") or "") and h.get("ascii"):
                _fix_header_ascii(path, n, h["ascii"])
                break

    sects = _section_types(root)
    all_sects = list(root.iter(f"{W}sectPr"))
    changed = 0
    body_seen = False
    front = params.get("front") or {}
    for si, meta in enumerate(sects):
        stype = _classify_section(meta["start_text"], params)
        sp = meta.get("sp_el")
        if sp is None and si + 1 < len(all_sects):
            sp = all_sects[si + 1]
        if sp is None:
            sp = all_sects[-1]
        # 2026-08-18 用户彻查「even 页眉绑定无效」：sp_el 是含 sectPr 的
        # **段落**（_section_types 返回段），而 _bind_ref/_set_pgnumtype/
        # refs 收集都期望 **sectPr 元素**——此前把 w:p 当 sectPr 用，
        # _bind_ref 新建的 even/footer 引用插进段落（错误位置），Word
        # 忽略段落级 headerReference → even 绑定「调用成功但产物没有」。
        # 统一取 sp 内的 sectPr（记忆铁律⑥「节尾段差一位」的深层变体）。
        if sp is not None and sp.tag == f"{W}p":
            _ppr = sp.find(f"{W}pPr")
            _spx = _ppr.find(f"{W}sectPr") if _ppr is not None else None
            if _spx is not None:
                sp = _spx
        actual_start = _section_actual_start(root, sp)
        if actual_start is not None:
            stype = _classify_section(actual_start, params)
        # 正文之后、又不是模板 identity 的节（结论等）按正文页眉接续
        if stype == "other" and body_seen:
            stype = "body"
        # 2026-08-23 用户实锤「参考文献和结论到一页了」+逐字拍板：不要
        # 连续分节符，用下一页分节符——源稿 WPS 的 continuous 分节符
        # （参考文献/结论节）导致新节不主动换页。所有节一律 nextPage。
        tp = sp.find(f"{W}type")
        if tp is not None and tp.get(f"{W}val") == "continuous":
            tp.set(f"{W}val", "nextPage")
            changed += 1
        refs = {}
        for ref in sp:
            tag = ref.tag.replace(W, "")
            if tag in ("headerReference", "footerReference"):
                refs[f"{tag}:{ref.get(f'{W}type') or 'default'}"] = True
        if stype not in ("body", "other", "empty"):
            # identity 节：绑目标自己的同名页眉（目录/ABSTRACT/参考文献…）
            # 2026-09-26 用户拍板：「用户说全文排版就改页眉，用户没明确说就不动」
            # ——页眉文本属「排什么」，只能来自目标文档；目标没有同名页眉时
            # **不再用模板文本新建**（此前会造出「摘要」「目录」页眉并套到
            # 报告全文），改为继承上一节的页眉（=保持目标原稿）。
            hdrs = [n for n, t in header_texts.items()
                    if re.sub(r"[\s\u3000]+", "", t) == stype
                    or stype in re.sub(r"[\s\u3000]+", "", t)]
            if hdrs:
                _bind_ref(path, sp, "headerReference", "default", _rid_of_part(path, hdrs[0]))
                _bind_ref(path, sp, "headerReference", "even", _rid_of_part(path, hdrs[0]))
                changed += 1
                for hpart in hdrs:
                    hroot = _read_part(path, hpart)
                    for hp in hroot.iter(f"{W}p"):
                        hppr = hp.find(f"{W}pPr")
                        if hppr is None:
                            hppr = etree.SubElement(hp, f"{W}pPr")
                        if hppr.find(f"{W}jc") is None:
                            jc = etree.Element(f"{W}jc")
                            jc.set(f"{W}val", "center")
                            rpr = hppr.find(f"{W}rPr")
                            if rpr is not None:
                                rpr.addprevious(jc)
                            else:
                                hppr.append(jc)
                            _write_part(path, hpart, hroot)
        if stype in ("body", "conclusion"):
            # 页眉：绑定正文页眉 part（even=姓名字眉、default=届别页眉）。
            # 2026-08-16 真机实证：headerReference 绑定安全（Word 打开正常），
            # footerReference 新增绑定在本产物触发「文件可能已经损坏」
            # （共享/独立/DEFLATED/字符串手术 5 种方式全坏，Word 修复时
            # 还原继承链）——footer 保持继承不动。
            if own_body:
                _bind_ref(path, sp, "headerReference", "even", _rid_of_part(path, own_body["even"]))
                _bind_ref(path, sp, "headerReference", "default", _rid_of_part(path, own_body["default"]))
                changed += 1
            elif body_headers:
                _bind_ref(path, sp, "headerReference", "even", _rid_of_part(path, body_headers[0]))
                _bind_ref(path, sp, "headerReference", "default", _rid_of_part(path, body_headers[-1]))
                changed += 1
            # 页脚：克隆干净 PAGE 页脚为独立 part，even/default 各绑一个
            # （2026-08-17 实证：页脚 part 重建为干净单段后，新增引用不再
            # 触发 Word 损坏——此前损坏根因=脏 part 含文本框/sdt；克隆
            # 源取「第一个含 PAGE 域的 footer part」，每节独立不共享）
            if body_footers and not refs.get("footerReference:default"):
                page_src = next((f for f in body_footers if _footer_has_page_field(path, f)), body_footers[0])
                f_even = _clone_part(path, page_src, "footer")
                f_def = _clone_part(path, page_src, "footer")
                if f_even:
                    _bind_ref(path, sp, "footerReference", "even", _rid_of_part(path, f_even))
                    changed += 1
                if f_def:
                    _bind_ref(path, sp, "footerReference", "default", _rid_of_part(path, f_def))
                    changed += 1
            # 前置节（摘要/ABSTRACT/目录）even 页脚补绑（2026-08-19 用户
            # 第四遍「同样的问题」全量审计发现）：evenAndOddHeaders=true 时
            # 只有 default 页脚 → 偶数页页码空白（实证 p10/p12/p14 应显
            # 示 II/IV/VI 却全空）。克隆同款 PAGE 页脚绑 even（与该节
            # default 同源），罗马页码随节 fmt=upperRoman 自然生效。
            # 注意：本分支在 stype in (body, conclusion) 内——前置节走
            # 不到，通用补绑统一放到主循环末尾（见下方 even 页脚通用补绑）。
            elif body_footers and refs.get("footerReference:default") \
                    and not refs.get("footerReference:even"):
                pass  # body 节 default+even 在上面 if 分支已双绑
            # 页码：**页面起始跟随源稿节语义**（2026-08-23 用户拍板：
            # 致谢=罗马 I、绪论=数字 1——源稿 sect[4] upperRoman start=1、
            # 正文节 Word 无 pgNumType 时按连续计算显示 7，与目录/印象
            # 不符）。正文起始节显式 decimal start=1；**不删任何节的
            # start**（此前「全文连续」逻辑删除 upperRoman/正文的 start
            # 导致致谢=II 而非 I、绪论=7 而非 1）。
            if stype == "body" and not body_seen:
                pnt = sp.find(f"{W}pgNumType")
                if pnt is None:
                    pnt = etree.Element(f"{W}pgNumType")
                    pgmar = sp.find(f"{W}pgMar")
                    if pgmar is not None:
                        pgmar.addnext(pnt)
                    else:
                        sp.insert(0, pnt)
                if pnt.get(f"{W}fmt") != "decimal":
                    pnt.set(f"{W}fmt", "decimal")
                    changed += 1
                if pnt.get(f"{W}start") is None:
                    pnt.set(f"{W}start", "1")
                    changed += 1
                tp = sp.find(f"{W}type")
                if tp is not None and tp.get(f"{W}val") == "continuous":
                    tp.set(f"{W}val", "nextPage")
                    changed += 1
                body_seen = True
            elif stype in ("body", "conclusion") and body_seen:
                # 正文后续节（分节由序言/绪论 sectPr 克隆而来，继承 start=1）
                # 必须连续计数（2026-08-23 用户报告：结论节被克隆 start=1
                # → 页脚 1/2、目录参考文献=2）。删除 start，跟随正文首节。
                pnt = sp.find(f"{W}pgNumType")
                if pnt is not None and pnt.get(f"{W}start") is not None:
                    del pnt.attrib[f"{W}start"]
                    changed += 1
        # elif 分支（删除节 start/移除 pgNumType）已移除——2026-08-23
        # 用户拍板：页码起始遵循源稿/模板节语义（upperRoman start=1 等），
        # 禁止硬编码删除。前置节 start 保留 → 致谢=罗马 I。
        # ── even 页脚通用补绑（2026-08-19 用户第四遍「同样的问题」全量
        # 审计）：evenAndOddHeaders=true 时任何「有 default PAGE 页脚但缺
        # even」的节，偶数页页码都空白（实证前置节 p10/p12/p14 应显示
        # II/IV/VI 全空）。克隆该节 default 页脚同款 PAGE part 绑 even。
        # 注意此逻辑在主循环内、每个 stype 分支后执行——body/refs 节已
        # 双绑（refs 双绑见下方），此处只补前置节。封面节（无页码）的
        # default 页脚无 PAGE 域，_footer_has_page_field 拦截不动。
        if refs.get("footerReference:default") and not refs.get("footerReference:even"):
            def_rid = None
            for ref in sp.findall(f"{W}footerReference"):
                if (ref.get(f"{W}type") or "default") == "default":
                    def_rid = ref.get(f"{R_NS}id")
                    break
            if def_rid:
                with zipfile.ZipFile(path) as z:
                    rels = etree.fromstring(z.read("word/_rels/document.xml.rels"))
                def_part = None
                for rel in rels:
                    if rel.get("Id") == def_rid:
                        tgt = rel.get("Target") or ""
                        def_part = tgt if tgt.startswith("word/") else "word/" + tgt
                        break
                if def_part and _footer_has_page_field(path, def_part):
                    f_even = _clone_part(path, def_part, "footer")
                    if f_even:
                        _bind_ref(path, sp, "footerReference", "even", _rid_of_part(path, f_even))
                        changed += 1
    # body 尾 sectPr（最后一节=结论+参考文献；源稿 TAIL 带 start=1 →
    # 页脚从 1 重计——2026-08-23 用户报告：结论节被克隆 start=1、目录
    # 参考文献=2）。正文区已开始（body_seen）→ 尾部节连续计数，删 start；
    # TAIL 前连续空段删除（尾部空白页）。
    final_sp = body.find(f"{W}sectPr")
    if final_sp is not None:
        # 2026-08-23 用户实锤：参考文献/结论节 continuous 分节符→下一页
        tpo = final_sp.find(f"{W}type")
        if tpo is not None and tpo.get(f"{W}val") == "continuous":
            tpo.set(f"{W}val", "nextPage")
            changed += 1
        pnt = final_sp.find(f"{W}pgNumType")
        if pnt is not None and pnt.get(f"{W}start") is not None and body_seen:
            del pnt.attrib[f"{W}start"]
            changed += 1
        while True:
            prev = None
            for c in body:
                if c is final_sp:
                    break
                prev = c
            if prev is not None and prev.tag == f"{W}p" and _is_empty_para(prev):
                body.remove(prev)
                changed += 1
            else:
                break
    # 2026-08-23 用户实锤（0.1.65 产物 p60 空白页）：节边界（内嵌 sectPr
    # nextPage）后的首段若残留手工分页符/分页前段——分节符已使新节从新页
    # 开始，br page 再分一次 → 双分页 → 中间页空白（「结  论」前空白页）。
    _body_els = list(body)
    for _i, _el in enumerate(_body_els):
        if _el.tag == f"{W}p" and _i + 1 < len(_body_els):
            _ppr = _el.find(f"{W}pPr")
            if _ppr is None or _ppr.find(f"{W}sectPr") is None:
                continue
            _nxt = _body_els[_i + 1]
            if _nxt.tag != f"{W}p":
                continue
            _nppr = _nxt.find(f"{W}pPr")
            # 1) 手工分页符（段内 run 的 br type=page）
            for _br in list(_nxt.iter(f"{W}br")):
                if _br.get(f"{W}type") == "page":
                    _run = _br.getparent()
                    _run.remove(_br)
                    if len(_run) == 0:
                        _run.getparent().remove(_run)
                    changed += 1
            # 2) 分页前段属性（pageBreakBefore）
            if _nppr is not None and _nppr.find(f"{W}pageBreakBefore") is not None:
                _nppr.remove(_nppr.find(f"{W}pageBreakBefore"))
                changed += 1

    # lastRenderedPageBreak 渲染缓存清理（2026-08-17 根因实证）：Word 的
    # 分页缓存标记——nextPage 分节 + 缓存 = 双重分页 → 空页 p11。
    # 重新分页的文档不应保留此缓存（lxml 重写会保留，Word 每次保存重算）。
    for el in list(root.iter(f"{W}lastRenderedPageBreak")):
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
            changed += 1

    # 空段边界清理（2026-08-17 实证）：①带 page break 的纯分页空段
    # （continuous 时代遗留，nextPage 后多余分页 → 空页）；②节首/节尾
    # 连续空段（TOC 节首空段 → 空页 p11）。全树遍历（含 sdt 内 TOC 节）。
    # 以每个带 sectPr 的段为锚，向前/向后删除连续空段；另删全文所有
    # 纯分页符空段（内容段保留）。
    all_paras = list(body.iter(f"{W}p"))
    sec_marks = [pp for pp in all_paras if _has_sectpr(pp)]
    for sp in sec_marks:
        parent = sp.getparent()
        while True:
            siblings = [c for c in parent if c.tag == f"{W}p"]
            if not siblings:
                break
            try:
                idx = siblings.index(sp)
            except ValueError:
                break
            moved = False
            j = idx - 1
            while j >= 0 and siblings[j] is not sp and _is_empty_para(siblings[j]):
                parent.remove(siblings[j])
                changed += 1
                moved = True
                break
            if moved:
                continue
            j = idx + 1
            while j < len(siblings) and siblings[j] is not sp and _is_empty_para(siblings[j]):
                parent.remove(siblings[j])
                changed += 1
                moved = True
                break
            if not moved:
                break
    # 封面性质节（首个身份/正文节之前）后处理解绑——必须在主循环之后，
    # 避免主循环误绑后再留下空页眉线。_section_types 会漏掉第一节。
    ordered = _ordered_sections(root)
    first_content = None
    for i, item in enumerate(ordered):
        st = _classify_section(item["start_text"], params)
        if st == "body" or st not in ("other", "empty"):
            first_content = i
            break
    if first_content is not None:
        for item in ordered[:first_content]:
            spx = item["sp"]
            if not front.get("has_visible_header"):
                changed += _unbind_refs(spx, "headerReference")
            if not front.get("has_page"):
                changed += _unbind_refs(spx, "footerReference")
    # 纯分页符空段（可能不在节边界）全删；节尾段（含 sectPr）的
    # br:page 只删 br 元素不删段（2026-08-18：continuous 节尾段的分页符
    # 是「continuous 时代遗留」，nextPage 后多余分页 → 空页 p14；删 br
    # 保留 sectPr，节结构不破坏。注意 _is_empty_para 对含 sectPr 的段
    # 返回 False——这里用「无文本」独立判断）。
    for para in list(body.iter(f"{W}p")):
        if _has_sectpr(para) and _para_is_blank(para):
            for br in list(para.iter(f"{W}br")):
                if br.get(f"{W}type") == "page":
                    br.getparent().remove(br)
                    changed += 1
            # 节尾不可见段极小化（2026-08-18 p14 空页）：TOC 域尾段
            # （fldChar end，不可删）+ 本 sectPr 段各占一行（默认 12pt+），
            # 目录条目加粗变高后溢出 → 下 nextPage 节前出现整页空白。
            # 1pt 固定行高 + 零段距把不可见高度压到最低（不破坏域/节结构）。
            if _shrink_invisible_para(para):
                changed += 1
            continue
        if _is_pagebreak_only(para):
            parent = para.getparent()
            if parent is not None:
                parent.remove(para)
                changed += 1

    # TOC 域尾段（fldChar end 无文本非 sectPr）同样极小化——它紧贴节尾
    # sectPr 段，两者都是不可见行，一起压 1pt 才能收回同页。
    for para in list(body.iter(f"{W}p")):
        if _has_sectpr(para) or not _para_is_blank(para):
            continue
        if para.find(f".//{W}fldChar") is not None:
            if _shrink_invisible_para(para):
                changed += 1

    # 节尾不可见段合并（2026-08-18 r23 根治 p14 空页）：1pt 极小化后满页
    # 仍会溢出（目录条目加粗变高占满 p13）。Word 原生结构 = 域尾 run 与
    # 分节符挂在最后一个内容段上（摘要/ABSTRACT 节即如此）——把「域尾段
    # + 空 sectPr 段」并入前一个可见段，不可见行归零，节尾不再溢页。
    # 条件保守：仅当 sectPr 段无文本、其前段无文本且只含 end 域符、再前
    # 段是可见内容段时合并（begin 域符跨段结构不动）。
    for para in list(body.iter(f"{W}p")):
        if not _has_sectpr(para) or not _para_is_blank(para):
            continue
        anchor = para.getprevious()
        if anchor is None or anchor.tag != f"{W}p":
            continue
        if _has_sectpr(anchor) or not _para_is_blank(anchor):
            continue
        if any(fc.get(f"{W}fldCharType") != "end"
               for fc in anchor.iter(f"{W}fldChar")):
            continue
        target = anchor.getprevious()
        if target is None or target.tag != f"{W}p":
            continue
        if _has_sectpr(target) or not _para_full_text(target).strip():
            continue
        sect = para.find(f"{W}pPr/{W}sectPr")
        # 1) 域尾段的 run 移入 target 末尾（end 域符跟在最后内容后）
        for child in list(anchor):
            if child.tag != f"{W}pPr":
                target.append(child)
        # 2) sectPr 段的非 pPr 子元素（bookmark 等锚点）移入 target
        for child in list(para):
            if child.tag != f"{W}pPr":
                target.append(child)
        # 3) sectPr 挂到 target.pPr（schema 末位，append 即合法）
        tppr = target.find(f"{W}pPr")
        if tppr is None:
            tppr = etree.Element(f"{W}pPr")
            target.insert(0, tppr)
        tppr.append(sect)
        # 4) 删两个不可见段
        parent = para.getparent()
        parent.remove(anchor)
        parent.remove(para)
        changed += 1

    # 空 nextPage 节尾段并入前一内容段（机制：空段单独占页 → 空白页）
    # 公式/图片段带 sectPr 不算空段——并入会把独立公式挪进上一段正文末尾
    for para in list(body.iter(f"{W}p")):
        if not _has_sectpr(para) or not _para_is_blank(para):
            continue
        prev = para.getprevious()
        if prev is None or prev.tag != f"{W}p":
            continue
        if _has_sectpr(prev) or not _para_full_text(prev).strip():
            continue
        sect = para.find(f"{W}pPr/{W}sectPr")
        if sect is None:
            continue
        pppr = prev.find(f"{W}pPr")
        if pppr is None:
            pppr = etree.Element(f"{W}pPr")
            prev.insert(0, pppr)
        if pppr.find(f"{W}sectPr") is not None:
            continue
        for child in list(para):
            if child.tag != f"{W}pPr":
                prev.append(child)
        pppr.append(sect)
        parent = para.getparent()
        parent.remove(para)
        changed += 1

    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "sections_fixed": changed}


def verify_section_references(path, params: dict) -> bool:
    """复查：正文节绑正文页眉+decimal；ABSTRACT 有页脚；参考文献有页眉。

    2026-08-17 修正：conclusion/refs（结论/参考文献）若在源文档独立成节
    但无独立页眉 part，正文页眉继承自前节（Word 渲染连续性）——只要求
    body 节完备（header+footer），conclusion/refs 弱化为存在性不强制，
    避免 verifier 比 handler 更严的假阴性（journal 记 failed 但格式实际
    已正确，检查语义与 repair 的实际修复范围对齐）。
    """
    try:
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        sects = _section_types(root)
        all_sects = list(root.iter(f"{W}sectPr"))
        for si, meta in enumerate(sects):
            stype = _classify_section(meta["start_text"], params)
            # 节尾段差一位（记忆铁律⑥）：meta.sp_el 才是该节的 sectPr，
            # all_sects[si] 差一位（2026-08-17 verifier 假阴性根因）
            sp = meta.get("sp_el")
            if sp is None and si < len(all_sects):
                sp = all_sects[si]
            if sp is None:
                continue
            # 2026-08-18 与 repair set 侧同源：sp_el 是含 sectPr 的段落，
            # refs 收集需要 sectPr 元素——解封装（否则 `for ref in sp`
            # 遍历的是 w:p 子元素，永远收集不到 headerReference）。
            if sp.tag == f"{W}p":
                _ppr = sp.find(f"{W}pPr")
                _spx = _ppr.find(f"{W}sectPr") if _ppr is not None else None
                if _spx is None:
                    continue
                sp = _spx
            refs = {}
            for ref in sp:
                tag = ref.tag.replace(W, "")
                if tag in ("headerReference", "footerReference"):
                    refs[f"{tag}:{ref.get(f'{W}type') or 'default'}"] = True
            if stype == "body":
                if not refs.get("headerReference:default"):
                    return False
                if not refs.get("footerReference:default"):
                    return False
        return True
    except Exception:
        return False


def _rid_of_part(path: Path, part: str) -> str:
    """document.xml.rels 按 Target basename 反查 rId。"""
    with zipfile.ZipFile(path) as z:
        rels = etree.fromstring(z.read("word/_rels/document.xml.rels"))
    base = part.split("/")[-1]
    for rel in rels:
        t = rel.get("Target") or ""
        if t.split("/")[-1] == base:
            return rel.get("Id")
    return ""


def _target_body_header_parts(path: Path, root, params: dict, header_texts: dict,
                              identity_norms: set) -> dict:
    """目标文档第一个「正文节」自己显式绑定的 even/default 页眉 part。

    只认该节 sectPr 上的显式引用（继承来的是前一节的页眉，不算正文页眉）；
    part 文本必须非空且不是模板 identity 页眉。只有一种类型时两槽共用它。
    找不到返回 {}（调用方退回旧候选表）。机制层：不含任何页眉文案。
    """
    with zipfile.ZipFile(path) as z:
        # 最小 docx 可能没有 rels（无页眉可言）→ 退回旧候选表
        if "word/_rels/document.xml.rels" not in z.namelist():
            return {}
        rels_raw = z.read("word/_rels/document.xml.rels")
    if not rels_raw.strip():
        return {}
    rels = etree.fromstring(rels_raw)
    rid2part: dict[str, str] = {}
    for rel in rels:
        tgt = (rel.get("Target") or "").replace("\\", "/").lstrip("/")
        if rel.get("Id") and tgt:
            rid2part[rel.get("Id")] = tgt if tgt.startswith("word/") else "word/" + tgt
    for item in _ordered_sections(root):
        if _classify_section(item["start_text"], params) != "body":
            continue
        got: dict[str, str] = {}
        for ref in item["sp"].findall(f"{W}headerReference"):
            typ = ref.get(f"{W}type") or "default"
            if typ not in ("default", "even"):
                continue
            part = rid2part.get(ref.get(f"{R}id") or "")
            core = re.sub(r"[\s\u3000]+", "", header_texts.get(part or "", ""))
            if part and core and core not in identity_norms:
                got[typ] = part
        if got:
            got.setdefault("even", got.get("default"))
            got.setdefault("default", got.get("even"))
            return got
        # 第一个正文节没有自己的页眉 → 不再往后找（后续节的绑定可能是本动作
        # 上一轮写进去的，拿它当「原稿正文页眉」会自我强化）
        return {}
    return {}


def _unbind_refs(sectpr, tag: str) -> int:
    """去掉 sectPr 上全部 headerReference / footerReference。"""
    n = 0
    for old in list(sectpr.findall(f"{W}{tag}")):
        sectpr.remove(old)
        n += 1
    return n


def _bind_ref(path: Path, sectpr, tag: str, ref_type: str, rel_id: str) -> None:
    """sectPr 上绑定引用（幂等替换，CT_SectPr 顺序：引用在最前）。"""
    if not rel_id:
        return
    for old in sectpr.findall(f"{W}{tag}"):
        ot = old.get(f"{W}type")
        if ot == ref_type or (ref_type == "default" and ot is None):
            old.set(f"{R_NS}id", rel_id)
            return
    ref = etree.Element(f"{W}{tag}")
    if ref_type != "default":
        ref.set(f"{W}type", ref_type)
    ref.set(f"{R_NS}id", rel_id)
    # CT_SectPr 顺序：headerReference 全部在 footerReference 前。
    # header → 插到第一个 footerReference 前；footer → 插到第一个非引用前。
    if tag == "headerReference":
        for i, ch in enumerate(sectpr):
            if ch.tag.replace(W, "") == "footerReference":
                sectpr.insert(i, ref)
                return
    for i, ch in enumerate(sectpr):
        if ch.tag.replace(W, "") not in ("headerReference", "footerReference"):
            sectpr.insert(i, ref)
            return
    sectpr.append(ref)


def _set_pgnumtype(sectpr, fmt: str, start: int) -> None:
    pnt = sectpr.find(f"{W}pgNumType")
    if pnt is None:
        pnt = etree.Element(f"{W}pgNumType")
        pgmar = sectpr.find(f"{W}pgMar")
        if pgmar is not None:
            pgmar.addnext(pnt)
        else:
            sectpr.insert(0, pnt)
    pnt.set(f"{W}fmt", fmt)
    pnt.set(f"{W}start", str(start))


_FN_NUM_RESTART_ORDER = ("pos", "numFmt", "numStart", "numRestart")


def set_footnote_numbering(path, params: dict) -> dict:
    """脚注编号重置方式（每页 / 每节 / 连续）。

    2026-10-17：CJC 模板明写「每页的脚注从 1 开始编序号」，旧链路只排字体字号，
    编号机制没落地。参数只由模板明写驱动（structure.footnote_restart），
    模板沉默不调用本动作。
    """
    restart = (params or {}).get("restart") or "eachPage"
    if restart not in ("eachPage", "eachSect", "continuous"):
        return {"changed": False, "error": f"不支持的脚注编号方式: {restart}"}
    target = Path(path)
    root = _read_part(target, "word/document.xml")
    changed = 0
    for sp in root.iter(f"{W}sectPr"):
        fpr = sp.find(f"{W}footnotePr")
        if fpr is None:
            fpr = etree.Element(f"{W}footnotePr")
            idx = 0
            for i, ch in enumerate(sp):
                if etree.QName(ch).localname in ("headerReference", "footerReference",
                                                 "endnotePr"):
                    idx = i + 1
            sp.insert(idx, fpr)
        cur = fpr.find(f"{W}numRestart")
        if cur is not None and cur.get(f"{W}val") == restart:
            continue
        if cur is None:
            cur = etree.SubElement(fpr, f"{W}numRestart")
            # CT_FtnProps 顺序：pos → numFmt → numStart → numRestart
            rank = {k: i for i, k in enumerate(_FN_NUM_RESTART_ORDER)}
            kids = list(fpr)
            kids.sort(key=lambda e: rank.get(etree.QName(e).localname, 99))
            for k in kids:
                fpr.remove(k)
            for k in kids:
                fpr.append(k)
        cur.set(f"{W}val", restart)
        _normalize_child_order(sp, _CT_SECTPR_ORDER)
        changed += 1
    if changed:
        _write_part(target, "word/document.xml", root)
    return {"changed": changed > 0, "sections": changed, "restart": restart}


def verify_footnote_numbering(path, params: dict = None) -> bool:
    restart = (params or {}).get("restart") or "eachPage"
    root = _read_part(Path(path), "word/document.xml")
    sects = list(root.iter(f"{W}sectPr"))
    if not sects:
        return False
    for sp in sects:
        fpr = sp.find(f"{W}footnotePr")
        nr = fpr.find(f"{W}numRestart") if fpr is not None else None
        if nr is None or nr.get(f"{W}val") != restart:
            return False
    return True


def _clone_part(path: Path, source_part: str, kind: str) -> str | None:
    """克隆 part（header/footer）为独立新 part 并完整注册（CT+rels）。
    每次调用返回新 part——Word 不允许共享/同 part 双类型引用。"""
    path = Path(path)
    with zipfile.ZipFile(path) as z:
        if source_part not in z.namelist():
            return None
        source_xml = z.read(source_part)
    new_part = _create_cloned_part(path, source_xml, kind)
    from earlyz.actions.header_footer_actions import (
        _ensure_part_registered, _header_content_type, _footer_content_type,
    )
    ctype = _header_content_type() if kind == "header" else _footer_content_type()
    _ensure_part_registered(path, new_part, ctype)
    return new_part


def _footer_has_page_field(path: Path, part: str) -> bool:
    with zipfile.ZipFile(path) as z:
        if part not in z.namelist():
            return False
        return b"PAGE" in z.read(part)


def _create_text_header(path: Path, text: str, ascii_font: str | None = None,
                       header_format: dict | None = None) -> str | None:
    """创建页眉 part。字体/字号必须来自模板 header_format，缺则不造。"""
    hf = header_format or {}
    ea = hf.get("font_east_asia")
    pt = hf.get("font_size_pt")
    if not ea or pt is None:
        return None
    sz = str(int(round(float(pt) * 2)))
    ascii_font = ascii_font or hf.get("font_ascii") or ea
    from earlyz.actions.header_footer_actions import _next_part_number, _ensure_part_registered, _header_content_type
    with zipfile.ZipFile(path) as z:
        n = _next_part_number(z, "header")
    part = f"word/header{n}.xml"
    ns = W[1:-1]
    xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:hdr xmlns:w="{ns}">'
           f'<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>'
           f'<w:jc w:val="center"/>'
           f'<w:rPr><w:rFonts w:eastAsia="{ea}" w:ascii="{ascii_font}" w:hAnsi="{ascii_font}"/><w:sz w:val="{sz}"/></w:rPr></w:pPr>'
           f'<w:r><w:rPr><w:rFonts w:eastAsia="{ea}" w:ascii="{ascii_font}" w:hAnsi="{ascii_font}"/><w:sz w:val="{sz}"/></w:rPr>'
           f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p></w:hdr>')
    import shutil
    tmp = path.with_suffix(".hdr.docx")
    shutil.copy2(path, tmp)
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))
        _write_part_zipped(zout, part, xml.encode("utf-8"))
    tmp.replace(path)
    _ensure_part_registered(path, part, _header_content_type())
    return part

def fit_table_to_page(path, params: dict) -> dict:
    """表格自适应页宽（§7.2 P1）：tblW 设为 100%（pct=5000）。R2。"""
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    tables = _locate_tables(root, int(params.get("table_index", -1)))
    changed = 0
    for tbl in tables:
        tblpr = tbl.find(f"{W}tblPr")
        if tblpr is None:
            tblpr = etree.Element(f"{W}tblPr")
            tbl.insert(0, tblpr)
        tw = tblpr.find(f"{W}tblW")
        if tw is None:
            tw = etree.SubElement(tblpr, f"{W}tblW")
        tw.set(f"{W}w", "5000")
        tw.set(f"{W}type", "pct")
        changed += 1
    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "tables": changed}


def verify_table_fit(path, params: dict) -> bool:
    """复查：tblW pct=5000。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        for tbl in _locate_tables(root, int(params.get("table_index", -1))):
            tblpr = tbl.find(f"{W}tblPr")
            tw = tblpr.find(f"{W}tblW") if tblpr is not None else None
            if tw is None or tw.get(f"{W}w") != "5000" or tw.get(f"{W}type") != "pct":
                return False
        return True
    except Exception:
        return False


def center_form_table_cells(path, params: dict = None) -> dict:
    """表单/块状表的单元格水平 + 垂直居中（2026-09-26 用户定案）。

    用户原话：「表格中的这些固定的属性应该都是水平垂直居中的，当然段落内容要和
    正文段落一样」。所以：只动**对齐**（cell 级 vAlign=center + 段落级 jc=center），
    段落其余属性（字体/字号/行距/缩进）一律不碰 —— 与正文保持一致。

    范围：表单/块状表（评语表、成绩表、签名表、封面信息表）——分类器判为 form
    或 ambiguous（表单特征不够十成但不属数据表）都处理；分页那套「宁可留白」的
    失败方向不适用于居中（居中无副作用），故 ambiguous 一并居中；
    params.force=True 时对全部表格生效（数据表名册/对照表通常不需要居中）。
    """
    path = Path(path)
    params = params or {}
    force = bool(params.get("force"))
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False}
    from earlyz.actions import table_layout as _tl
    tables = [el for el in body if _tl._localname(el) == "tbl"]
    if not tables:
        return {"changed": False, "tables": 0}
    touched_t = touched_c = touched_p = 0
    for tbl in tables:
        if not force:
            try:
                if (_tl.classify_table(tbl) or {}).get("kind") not in ("form", "ambiguous"):
                    continue
            except Exception:
                continue
        touched_t += 1
        for tc in tbl.iter():
            if _tl._localname(tc) != "tc":
                continue
            tcpr = tc.find(f"{W}tcPr")
            if tcpr is None:
                tcpr = etree.Element(f"{W}tcPr")
                tc.insert(0, tcpr)
            va = tcpr.find(f"{W}vAlign")
            if va is None:
                va = etree.SubElement(tcpr, f"{W}vAlign")
            if va.get(f"{W}val") != "center":
                va.set(f"{W}val", "center")
                touched_c += 1
            _normalize_child_order(tcpr, _CT_TCPR_ORDER)
            for para in tc.iter():
                if _tl._localname(para) != "p":
                    continue
                ppr = para.find(f"{W}pPr")
                if ppr is None:
                    ppr = etree.Element(f"{W}pPr")
                    para.insert(0, ppr)
                jc = ppr.find(f"{W}jc")
                if jc is None:
                    jc = etree.Element(f"{W}jc")
                    ppr.append(jc)
                if jc.get(f"{W}val") != "center":
                    jc.set(f"{W}val", "center")
                    touched_p += 1
                _normalize_child_order(ppr, _CT_PPR_ORDER)
    changed = touched_c > 0 or touched_p > 0
    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed, "tables": touched_t, "cells": touched_c,
            "paras": touched_p, "force": force}


def verify_center_form_table_cells(path, params: dict = None) -> bool:
    """复查：被处理的表单表里，每个单元格 vAlign=center 且段落 jc=center。"""
    try:
        path = Path(path)
        params = params or {}
        force = bool(params.get("force"))
        root = _read_part(path, "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return False
        from earlyz.actions import table_layout as _tl
        n = 0
        for tbl in [el for el in body if _tl._localname(el) == "tbl"]:
            if not force:
                try:
                    if (_tl.classify_table(tbl) or {}).get("kind") not in ("form", "ambiguous"):
                        continue
                except Exception:
                    continue
            n += 1
            for tc in tbl.iter():
                if _tl._localname(tc) != "tc":
                    continue
                tcpr = tc.find(f"{W}tcPr")
                va = tcpr.find(f"{W}vAlign") if tcpr is not None else None
                if va is None or va.get(f"{W}val") != "center":
                    return False
                for para in tc.iter():
                    if _tl._localname(para) != "p":
                        continue
                    ppr = para.find(f"{W}pPr")
                    jc = ppr.find(f"{W}jc") if ppr is not None else None
                    if jc is None or jc.get(f"{W}val") != "center":
                        return False
        return n >= 0
    except Exception:
        return False


def _toc_right_tab_pos(body) -> int:
    """版心右缘（twips）= 页面宽 - 左右页边距。"""
    sect = body.find(f"{W}sectPr")
    if sect is None:
        return 8776
    pg = sect.find(f"{W}pgSz")
    mar = sect.find(f"{W}pgMar")
    try:
        width = int(pg.get(f"{W}w", "11906")) if pg is not None else 11906
        lm = int(mar.get(f"{W}left", "1800")) if mar is not None else 1800
        rm = int(mar.get(f"{W}right", "1800")) if mar is not None else 1800
        return max(100, width - lm - rm)
    except Exception:
        return 8776


_PPR_AFTER_TABS = ("spacing", "ind", "contextualSpacing", "mirrorIndents",
                   "suppressOverlap", "jc", "textDirection", "textAlignment",
                   "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
                   "sectPr", "pPrChange")


def _insert_tabs(ppr, tabs_el):
    """按 CT_PPr 顺序把 w:tabs 插到正确位置（schema 对子元素顺序严格）。"""
    idx = len(ppr)
    for i, ch in enumerate(ppr):
        if etree.QName(ch).localname in _PPR_AFTER_TABS:
            idx = i
            break
    ppr.insert(idx, tabs_el)


def polish_toc_entries(path, params: dict = None) -> dict:
    """目录收尾（编译链最后一步）：页码右对齐到版心右缘 + 点线前导符，缩进交回样式。

    2026-09-26 用户定案：先排完其他内容、最后处理目录；条目与页码先正确、
    样式最后上（之后用户 F9 由 Word 重建、与我们无关）。规格取自模板目录示例的
    注明：「目录示例：（自动生成）」= 自动生成目录形态 → 文字 + 点线前导符 +
    页码贴版心右缘；各章名称黑体 4 号加粗、其余宋体小 4。

    params: {"font_spec": {"1": {"size": 28, "bold": True, "eastasia": "黑体"},
                           "2": {"size": 24, "bold": False, "eastasia": "宋体"}}}
    """
    path = Path(path)
    params = params or {}
    spec = params.get("font_spec") or {
        "1": {"size": 28, "bold": True, "eastasia": "黑体"},
        "2": {"size": 24, "bold": False, "eastasia": "宋体"},
    }
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False}
    pos = _toc_right_tab_pos(body)
    ind_removed = tabs_set = font_set = 0
    entries = 0
    for para in body.iter(f"{W}p"):
        if not _para_pageref_target(para):
            continue
        entries += 1
        ppr = para.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.SubElement(para, f"{W}pPr")
            para.remove(ppr)
            para.insert(0, ppr)
        # 1) 直系缩进删除（缩进交回 TOC1/TOC2 样式）
        ind = ppr.find(f"{W}ind")
        if ind is not None:
            ppr.remove(ind)
            ind_removed += 1
        # 2) 制表位：右对齐 + 点线，落在版心右缘
        old = ppr.find(f"{W}tabs")
        if old is not None:
            ppr.remove(old)
        tabs = etree.Element(f"{W}tabs")
        tab = etree.SubElement(tabs, f"{W}tab")
        tab.set(f"{W}val", "right")
        tab.set(f"{W}leader", "dot")
        tab.set(f"{W}pos", str(pos))
        _insert_tabs(ppr, tabs)
        tabs_set += 1
        # 3) 字号/字体：一级=黑体 4 号加粗，二级=宋体小 4
        pstyle = ppr.find(f"{W}pStyle")
        lvl = (pstyle.get(f"{W}val") if pstyle is not None else "") or ""
        lvl = "1" if lvl.strip() in ("1", "7", "TOC1", "toc 1") else "2"
        want = spec.get(lvl) or {}
        for run in para.iter(f"{W}r"):
            rpr = run.find(f"{W}rPr")
            if rpr is None:
                rpr = etree.Element(f"{W}rPr")
                run.insert(0, rpr)
            for tag in ("b", "bCs", "sz", "szCs"):
                el = rpr.find(f"{W}{tag}")
                if el is not None:
                    rpr.remove(el)
            rf = rpr.find(f"{W}rFonts")
            if want.get("eastasia"):
                if rf is None:
                    rf = etree.SubElement(rpr, f"{W}rFonts")
                rf.set(f"{W}eastAsia", want["eastasia"])
                rf.set(f"{W}ascii", "Times New Roman")
                rf.set(f"{W}hAnsi", "Times New Roman")
            if want.get("bold"):
                be = etree.SubElement(rpr, f"{W}b")
                be.set(f"{W}val", "1")
            if want.get("size"):
                se = etree.SubElement(rpr, f"{W}sz")
                se.set(f"{W}val", str(want["size"]))
                sc = etree.SubElement(rpr, f"{W}szCs")
                sc.set(f"{W}val", str(want["size"]))
            font_set += 1
    if not entries:
        return {"changed": False, "entries": 0}
    _write_part(path, "word/document.xml", root)
    return {"changed": True, "entries": entries, "ind_removed": ind_removed,
            "tabs_set": tabs_set, "runs_styled": font_set, "tab_pos": pos}


def verify_toc_entries(path, params: dict = None) -> bool:
    """复查：目录条目 = 无直系缩进 + 右对齐点线制表位（版心右缘）+ 正确字号。

    2026-09-28：删除被遮蔽的旧版（08-27 的 pStyle 检查）后本实现生效——
    旧版曾静默覆盖本实现，导致 polish_toc_entries 三天来验错对象。
    """
    try:
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return False
        cover_paras = _cover_paragraphs(root, _read_styles_map(path))   # 与动作同条件
        pos = _toc_right_tab_pos(body)
        n = 0
        for para in body.iter(f"{W}p"):
            if cover_paras and any(para is c for c in cover_paras):
                continue
            if not _para_pageref_target(para):
                continue
            n += 1
            ppr = para.find(f"{W}pPr")
            if ppr is None or ppr.find(f"{W}ind") is not None:
                return False
            tabs = ppr.find(f"{W}tabs")
            if tabs is None:
                return False
            stops = [dict(x.attrib) for x in tabs]
            want = {f"{W}val": "right", f"{W}leader": "dot", f"{W}pos": str(pos)}
            if not any(all(s.get(k) == v for k, v in want.items()) for s in stops):
                return False
        # n == 0（文档没有目录条目）不算失败：空对象应 skip，不 vacuous-fail
        # —— 2026-09-28 真回归：本行原为 `return n > 0`，让「无目录文档」的整条
        # 编译链被判 FAILED（polish_toc_entries failed: verifier returned False）。
        return True
    except Exception:
        return False


_CHAPTER_HEAD_RE = re.compile(
    r"^(第\s*[一二三四五六七八九十百千0-9]+\s*章|[0-9]+(\.[0-9]+)*[\s、.．])"
)
_TOC_SKIP_PREFIX = ("摘要", "abstract", "关键词", "keywords", "目录", "contents")


def _toc_norm(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def _default_toc_title_props() -> dict:
    """模板既无 toc_title 也无 toc.L1 时的标题兜底（黑体小二加粗）。

    仅在前两级都没有时才用；结果里的 title_source="default" 会如实写出来。
    """
    return {"font_east_asia": "黑体", "font_size_pt": 18.0, "bold": True}


def _max_bookmark_id(body) -> int:
    mx = 0
    for bm in body.iter(f"{W}bookmarkStart"):
        try:
            mx = max(mx, int(bm.get(f"{W}id", "0")))
        except (TypeError, ValueError):
            continue
    return mx



def _style_outline_map(path) -> dict[str, int]:
    """styles.xml 里段落样式的大纲级别（含 basedOn 继承一层）。

    很多模板不给段落写直接 outlineLvl，级别只在「标题 1/2/3」样式里；
    只读直接属性会让 `insert_toc_if_missing` 认为「没有标题」直接跳过
    （2026-10-29 功能设计说明书/安装维护手册实测：headings=0）。
    """
    import zipfile
    try:
        with zipfile.ZipFile(str(path)) as z:
            root = etree.fromstring(z.read("word/styles.xml"))
    except Exception:  # noqa: BLE001
        return {}
    raw: dict[str, tuple[int | None, str | None]] = {}
    for st in root.iter(f"{W}style"):
        if st.get(f"{W}type") != "paragraph":
            continue
        sid = st.get(f"{W}styleId")
        if not sid:
            continue
        o = st.find(f"{W}pPr/{W}outlineLvl")
        lvl = None
        if o is not None:
            try:
                lvl = int(o.get(f"{W}val", "0"))
            except (TypeError, ValueError):
                lvl = None
        b = st.find(f"{W}basedOn")
        raw[sid] = (lvl, b.get(f"{W}val") if b is not None else None)
    out: dict[str, int] = {}
    for sid in raw:
        lvl, base = raw[sid]
        if lvl is None and base:
            lvl = raw.get(base, (None, None))[0]
        if lvl is not None:
            out[sid] = lvl
    return out


_NUM_HEAD_RE = re.compile(r"^(\d+(?:\.\d+)*)[\s、.．]\s*\S")

def _toc_headings_for_generation(body, max_level: int, style_outline: dict | None = None) -> list[tuple]:
    """收集可进目录的正文标题（outlineLvl 1..max_level + 参考文献/致谢/附录）。

    只收 body 直系段落（表格/sdt 内的标题不动，也不作为插入锚点）。
    前置章节（摘要/ABSTRACT/关键词/目录）不进目录——目录从正文第 1 页开始编号，
    与 sync_toc_page_numbers 的 offset 口径一致（offset = 首条目物理页 − 1）。
    """
    from earlyz.ir.analyzer import _back_heading_role

    out: list[tuple] = []
    seen: set[str] = set()
    for p in body.iter(f"{W}p"):
        if p.getparent() is not body:
            continue
        txt = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        if not txt:
            continue
        norm = _toc_norm(txt)
        low = norm.lower()
        if low.startswith(_TOC_SKIP_PREFIX) or norm.endswith(
            ("原创性声明", "版权使用授权书", "使用授权书")
        ):
            continue
        ppr = p.find(f"{W}pPr")
        lvl = None
        if ppr is not None:
            o = ppr.find(f"{W}outlineLvl")
            if o is not None:
                try:
                    lvl = int(o.get(f"{W}val", "0"))
                except (TypeError, ValueError):
                    lvl = None
        if lvl is None and style_outline:
            ps = ppr.find(f"{W}pStyle") if ppr is not None else None
            if ps is not None:
                lvl = style_outline.get(ps.get(f"{W}val"))
        if lvl is None:
            # 编号形态兜底：文档靠「1 范围」「1.1 编写目的」这类文本编号分章，
            # 没有 outlineLvl 也没有标题样式（公文/技术文档常见）。
            m = _NUM_HEAD_RE.match(norm)
            if m and len(norm) <= 40 and not re.search(r"[。；！？.;!?]", txt):
                lvl = min(5, m.group(1).count("."))
        if lvl is None:
            back = _back_heading_role(txt)
            if back in ("refs_heading", "ack_heading", "appendix_heading") and norm not in seen:
                seen.add(norm)
                out.append((p, 1, txt.strip()))
            continue
        if 0 <= lvl < max_level and norm not in seen:
            seen.add(norm)
            out.append((p, lvl + 1, txt.strip()))
    return out


def _add_toc_bookmark(p, name: str, bid: int) -> None:
    start = etree.Element(f"{W}bookmarkStart")
    start.set(f"{W}id", str(bid))
    start.set(f"{W}name", name)
    end = etree.Element(f"{W}bookmarkEnd")
    end.set(f"{W}id", str(bid))
    ppr = p.find(f"{W}pPr")
    p.insert(1 if ppr is not None else 0, start)
    p.append(end)


def _toc_entry_para(level: int, title: str, bookmark: str) -> etree._Element:
    p = etree.Element(f"{W}p")
    ppr = etree.SubElement(p, f"{W}pPr")
    ps = etree.SubElement(ppr, f"{W}pStyle")
    ps.set(f"{W}val", f"TOC{level}")
    hl = etree.SubElement(p, f"{W}hyperlink")
    hl.set(f"{W}anchor", bookmark)
    hl.set(f"{W}history", "1")
    r = etree.SubElement(hl, f"{W}r")
    t = etree.SubElement(r, f"{W}t")
    t.text = title
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r_tab = etree.SubElement(hl, f"{W}r")
    etree.SubElement(r_tab, f"{W}tab")
    r1 = etree.SubElement(hl, f"{W}r")
    f1 = etree.SubElement(r1, f"{W}fldChar")
    f1.set(f"{W}fldCharType", "begin")
    r2 = etree.SubElement(hl, f"{W}r")
    it = etree.SubElement(r2, f"{W}instrText")
    it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    it.text = f" PAGEREF {bookmark} \\h "
    r3 = etree.SubElement(hl, f"{W}r")
    f2 = etree.SubElement(r3, f"{W}fldChar")
    f2.set(f"{W}fldCharType", "separate")
    r4 = etree.SubElement(hl, f"{W}r")
    t4 = etree.SubElement(r4, f"{W}t")
    t4.text = "1"
    r5 = etree.SubElement(hl, f"{W}r")
    f3 = etree.SubElement(r5, f"{W}fldChar")
    f3.set(f"{W}fldCharType", "end")
    return p


def _toc_title_para(text: str, props: dict) -> etree._Element:
    p = etree.Element(f"{W}p")
    ppr = etree.SubElement(p, f"{W}pPr")
    jc = etree.SubElement(ppr, f"{W}jc")
    jc.set(f"{W}val", "center")
    sp = etree.SubElement(ppr, f"{W}spacing")
    sp.set(f"{W}before", "0")
    sp.set(f"{W}after", "240")
    kn = etree.SubElement(ppr, f"{W}keepNext")
    kn.set(f"{W}val", "1")
    r = etree.SubElement(p, f"{W}r")
    rpr = etree.SubElement(r, f"{W}rPr")
    rf = etree.SubElement(rpr, f"{W}rFonts")
    rf.set(f"{W}eastAsia", str(props.get("font_east_asia") or "黑体"))
    rf.set(f"{W}ascii", str(props.get("font_ascii") or "Times New Roman"))
    if props.get("bold"):
        etree.SubElement(rpr, f"{W}b")
    sz = etree.SubElement(rpr, f"{W}sz")
    try:
        sz.set(f"{W}val", str(int(round(float(props.get("font_size_pt") or 18.0) * 2))))
    except (TypeError, ValueError):
        sz.set(f"{W}val", "36")
    t = etree.SubElement(r, f"{W}t")
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return p


def _toc_page_offset(body, page_map: dict, entries: list) -> int:
    """目录页码偏移（物理页 − 显示页）。

    2026-10-17 修正：只有文档**真的重排页码**（某节 pgNumType/@start）时才用
    「首条目物理页 − 1」推定；连续编号（无 pgNumType start）时偏移为 0 —— 否则
    生成目录后首条目落在正文第 1 页会把所有页码整体减错（实测 第1章 6 页被写成 1）。
    """
    has_restart = False
    for sect in body.iter(f"{W}sectPr"):
        pn = sect.find(f"{W}pgNumType")
        if pn is not None and pn.get(f"{W}start"):
            has_restart = True
            break
    if not has_restart or not entries:
        return 0
    first_page = page_map.get(entries[0][1])
    return max(0, int(first_page) - 1) if first_page else 0


def _ensure_toc_styles(path: Path, levels: set[int]) -> int:
    """确保 TOC1/2/3 段落样式存在（模板常无目录样式 → 条目没有样式、F9 重建也不成体系）。

    只在**从零生成目录**时调用：缺哪个建哪个，不动模板已有的样式。
    缩进**故意不设**：2026-08-18 口径「层级用制表位区分，不用 left 缩进」——
    `ensure_toc_tab_stops` 会清掉 TOC2/3 的 ind，这里也不给。
    """
    root = _read_part(path, "word/styles.xml")
    existing = {s.get(f"{W}styleId") for s in root.iter(f"{W}style")}
    added = 0
    for lvl in sorted(levels):
        sid = f"TOC{lvl}"
        if sid in existing:
            continue
        st = etree.SubElement(root, f"{W}style")
        st.set(f"{W}type", "paragraph")
        st.set(f"{W}styleId", sid)
        nm = etree.SubElement(st, f"{W}name")
        nm.set(f"{W}val", f"toc {lvl}")
        based = etree.SubElement(st, f"{W}basedOn")
        based.set(f"{W}val", "Normal")
        etree.SubElement(st, f"{W}pPr")
        added += 1
    if added:
        _write_part(path, "word/styles.xml", root)
    return added


def _para_text_tabs(p) -> str:
    """段落文本，`<w:tab/>` 还原成 "\\t"（目录条目靠制表符分隔编号/标题/页码）。"""
    parts: list[str] = []
    for node in p.iter():
        if node.tag == f"{W}t":
            parts.append(node.text or "")
        elif node.tag == f"{W}tab":
            parts.append("\t")
    return "".join(parts)


def _promote_headings_from_text_toc(body, entries) -> int:
    """用文本目录条目回填正文标题的 outlineLvl。返回提升的段数。

    技术文档类源文档常见形态：目录条目齐全（`1.\\t文档介绍\\t4`），正文标题却是
    无编号、无标题样式的光秃文字（`文档介绍`），于是 `_toc_headings_for_generation`
    收不到任何标题，`insert_toc_if_missing` 只能跳过（2026-11-07 功能设计说明书：
    29 条目录 + 正文 0 个可识别标题）。

    目录本身就是作者写好的提纲：拿条目文本去正文里对同名段、按条目编号定层级
    （`1.` → 1 级，`1.1.` → 2 级）。目录条目段自己不参与（它的文本带页码，norm 不同）。
    已是标题的段不覆盖（保留原有 outlineLvl）。
    """
    _num_re = re.compile(r"^\s*(\d+(?:[.．]\d+)*)[.．、]?\s*$")
    items: list[tuple[str, int]] = []
    for q in entries:
        # 目录条目里的分隔符多为 <w:tab/> 元素（不是 w:t 里的制表符），拼文本时要补回 "\t"，
        # 否则 "1." "文档介绍" "4" 会粘成 "1.文档介绍4"，编号与页码都拆不开。
        parts_raw: list[str] = []
        for node in q.iter():
            if node.tag == f"{W}t":
                parts_raw.append(node.text or "")
            elif node.tag == f"{W}tab":
                parts_raw.append("\t")
        t = "".join(parts_raw).strip()
        if not t:
            continue
        title, lvl = "", 1
        parts = [s.strip() for s in re.split(r"\t+", t) if s.strip()]
        if len(parts) >= 2:
            m = _num_re.match(parts[0])
            if m:
                lvl = min(5, m.group(1).replace("．", ".").count(".") + 1)
                title = parts[1]
        if not title:
            m = re.match(
                r"^\s*(\d+(?:[.．]\d+)*)[.．、]?\s*(.+?)[\s.．]{2,}\d{1,4}\s*$", t)
            if m:
                lvl = min(5, m.group(1).replace("．", ".").count(".") + 1)
                title = m.group(2)
        if not title:
            continue
        norm = _toc_norm(title)
        if norm:
            items.append((norm, lvl))
    if len(items) < 2:
        return 0
    want: dict[str, int] = {}
    for norm, lvl in items:
        want.setdefault(norm, lvl)
    pending = {id(q) for q in entries}
    promoted = 0
    for p in body.iter(f"{W}p"):
        if p.getparent() is not body or id(p) in pending:
            continue
        txt = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        if not txt:
            continue
        norm = _toc_norm(txt)
        lvl = want.get(norm)
        if lvl is None:
            continue
        ppr = p.find(f"{W}pPr")
        if ppr is not None and ppr.find(f"{W}outlineLvl") is not None:
            want.pop(norm, None)
            continue
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            p.insert(0, ppr)
        o = etree.Element(f"{W}outlineLvl")
        o.set(f"{W}val", str(max(0, lvl - 1)))
        _rpr = ppr.find(f"{W}rPr")
        if _rpr is not None:
            ppr.insert(ppr.index(_rpr), o)  # outlineLvl 在 rPr 之前（schema 顺序）
        else:
            ppr.append(o)
        want.pop(norm, None)
        promoted += 1
        if not want:
            break
    return promoted


def insert_toc_if_missing(path, params: dict = None) -> dict:
    """文档没有目录时，从零生成目录（2026-10-17 用户拍板「这个要做好」）。

    取值口径（模板提供排什么，2026-09-26/10-17 铁律）：
    - 条目格式来自模板注释抽出的 structure.toc（L1 黑体4号加粗 / 其余宋体小4）；
    - 目录标题格式来自模板「目录示例：目 录（黑体小2加粗）」注释（structure.toc_title），
      缺失时退回 toc.L1 规格，再无则黑体小二加粗（结果里 source 字段写明出处）；
    - 插入位置 = 第一个正文章标题之前（摘要/ABSTRACT 之后），分页交由
      ensure_section_page_breaks 按模板 section_break_words 处理；
    - 条目 = 正文标题（outlineLvl 1..N，排除摘要/ABSTRACT/关键词/目录）＋
      参考文献/致谢/附录（按 back 角色）；书签 _Toc9xxxxxxx 供 Word/PAGEREF 使用。

    幂等：已有 PAGEREF/TOC 域或「目录」标题段 → no-op。

    **迁移语义（2026-11-15 用户定案「目录只迁移不生成」，EZI-0038）**：
    源文档没有 TOC 域时**一律不造目录**——默认直接 `skipped=source_has_no_toc_field`。
    只带手写目录条目的文档保持原样：旧逻辑会把文本条目删掉、按标题重造域条目
    与 `_Toc9xxxxxxx` 书签，而这些书签随后被「封面还原」整段替换带走 → 产物
    PAGEREF/超链接锚点全部悬空，Word 打开报「错误！未找到引用源」。
    交互式「生成目录」命令（用户明确要求造目录）必须显式传 ``allow_generate=True``。
    """
    path = Path(path)
    params = params or {}
    if params.get("enabled") is False:
        return {"changed": False, "skipped": "disabled"}
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False, "skipped": "no_body"}

    _converted_text_toc = False
    _pending_text_toc: list = []
    # 1) 已有目录？域（PAGEREF/TOC）或「目录」标题段都算
    for p in body.iter(f"{W}p"):
        for x in p.iter(f"{W}instrText"):
            t = (x.text or "").strip().upper()
            if t.startswith("PAGEREF") or t.startswith("TOC"):
                return {"changed": False, "skipped": "toc_exists"}

    # 1b) 迁移语义闸门（EZI-0038）：能走到这里 = 文档里没有任何 TOC 域。
    #     按 2026-11-15 定案「目录只迁移不生成」，不许再按标题造条目/书签。
    #     反例守卫：显式 allow_generate=True（交互式「生成目录」）仍走老路径。
    if not params.get("allow_generate"):
        return {"changed": False, "skipped": "source_has_no_toc_field",
                "note": "2026-11-15 定案：目录只迁移不生成（源文档无 TOC 域）"}

    # 只有「目录」标题、后面没有条目/域 → 是空目录页，继续生成（2026-10-28 大连民族：
    # 源文档留了一个空「目 录」页，旧逻辑当成已有目录直接跳过，产物就是空目录页）。
    _paras = [p for p in body.iter(f"{W}p") if p.getparent() is body]
    for _i, p in enumerate(_paras):
        norm = _toc_norm("".join(t.text or "" for t in p.iter(f"{W}t")))
        if norm in ("目录", "contents"):
            _entries = 0
            _entry_paras = []
            for q in _paras[_i + 1: _i + 41]:
                # 分隔符可能是 <w:tab/> 元素——纯 w:t 拼接会把「1.」「文档介绍」「4」
                # 粘成「1.文档介绍4」，条目判定全部落空（2026-11-07 技术文档族）。
                _t = _para_text_tabs(q).strip()
                if _t.replace(" ", "") in ("目录", "目次"):
                    continue
                if re.search(r"[	　 ]\d+\s*$", _t) or re.search(r"\.{4,}\s*\d+\s*$", _t)                         or re.search(r"\s\d+\s*$", _t) and len(_t) < 80:
                    _entries += 1
                    _entry_paras.append(q)
                if re.match(r"^第\s*[一二三四五六七八九十]+\s*章", _t):
                    break
            if _entries >= 2:
                # 纯文本目录 → 域目录（2026-10-29 用户拍板 A 方案）：
                # 源文档的目录被压成了「标题	页码」纯文本，模板要的是 TOC 域。
                # 删掉文本条目段，落回下面的生成路径（标题段复用、书签重建）。
                # 先记下待删段落，等确认「删完还能生成」再删（见下方 heads 检查）：
                # 有些文档的“标题”就是目录条目本身，删了就没有可索引标题了，
                # 那时必须保留原目录，不能把内容删空（2026-10-29 功能设计说明书）。
                _pending_text_toc = list(_entry_paras)
                break
                return {"changed": False, "skipped": "toc_exists"}

    toc_spec = params.get("toc") or {}
    max_level = int(params.get("max_level") or (3 if (toc_spec.get("3") or toc_spec.get(3)) else 2))
    heads = _toc_headings_for_generation(body, max_level, _style_outline_map(path))
    # 待删条目本身常被作者套了标题样式（整块目录都是 heading）——判「够不够生成」
    # 必须用**删掉它们之后**剩下的标题，否则会误判为有结构、删完只剩空目录。
    _pending_ids = {id(q) for q in _pending_text_toc}
    _heads_after = [h for h in heads if id(h[0]) not in _pending_ids]
    # 文本目录反推标题（2026-11-07）：技术文档类源文档「目录齐全、正文标题全是大白话」，
    # 目录条目本身没有可索引标题以外的信息——删掉就没得生成。先用条目文本回填正文段的
    # outlineLvl，再重算一次标题表；仍然不够才跳过（保持旧行为，不删目录）。
    if len(_heads_after) < 2 and _pending_text_toc:
        _promoted = _promote_headings_from_text_toc(body, _pending_text_toc)
        if _promoted:
            _style_map2 = _style_outline_map(path)
            heads = _toc_headings_for_generation(body, max_level, _style_map2)
            _heads_after = [h for h in heads if id(h[0]) not in _pending_ids]
            if len(_heads_after) < 2:
                # 提升失败：撤掉刚写的 outlineLvl，退回「保留原目录」
                for _p in body.iter(f"{W}p"):
                    _ppr = _p.find(f"{W}pPr")
                    if _ppr is None:
                        continue
                    _o = _ppr.find(f"{W}outlineLvl")
                    if _o is not None and _promoted:
                        _ppr.remove(_o)
                        _promoted -= 1
                return {"changed": False, "skipped": "no_headings_after_promote",
                        "headings": len(_heads_after), "text_toc_kept": len(_pending_text_toc)}
    if len(_heads_after) < 2:
        return {"changed": False, "skipped": "no_headings",
                "headings": len(_heads_after), "text_toc_kept": len(_pending_text_toc)}
    if _pending_text_toc:
        for _q in _pending_text_toc:
            if _q.getparent() is body:
                body.remove(_q)
        _converted_text_toc = True
        heads = _heads_after

    first = None
    for para, lvl, txt in heads:
        if lvl == 1 and _CHAPTER_HEAD_RE.match(txt):
            first = para
            break
    if first is None:
        first = next((para for para, lvl, _ in heads if lvl == 1), None)
    if first is None or first.getparent() is not body:
        return {"changed": False, "skipped": "no_chapter_anchor"}

    # 2) 标题格式：模板 toc_title → toc.L1 → 黑体小二加粗（source 写清楚）
    props = params.get("toc_title") or {}
    src = "template:toc_title"
    if not props:
        props = toc_spec.get("1") or toc_spec.get(1) or {}
        src = "template:toc.L1"
    if not props:
        props = _default_toc_title_props()
        src = "default"
    title_text = (params.get("title_text") or "").strip()
    if not title_text:
        all_txt = "".join(t.text or "" for t in body.iter(f"{W}t"))
        cjk = sum(1 for ch in all_txt if "\u4e00" <= ch <= "\u9fff")
        english = bool(all_txt) and cjk / max(1, len(all_txt)) < 0.01
        title_text = "Contents" if english else "目  录"

    # 3) 书签 + 条目
    next_bid = _max_bookmark_id(body) + 1
    entries = []
    for i, (para, lvl, txt) in enumerate(heads):
        name = f"_Toc{910000000 + i}"
        _add_toc_bookmark(para, name, next_bid + i)
        entries.append((lvl, txt, name))

    # 已有「目 录」标题段就复用，别再插一个（否则产物出现两个「目  录」，
    # 大连民族实测：生成目录后标题重复，内容比对 +2 字）
    _existing_title = None
    for _p in _paras:
        if _toc_norm("".join(t.text or "" for t in _p.iter(f"{W}t"))) in ("目录", "contents"):
            _existing_title = _p
            break

    new_elems = [] if _existing_title is not None else [_toc_title_para(title_text, props)]
    new_elems += [_toc_entry_para(lvl, txt, name) for lvl, txt, name in entries]
    idx = (list(body).index(_existing_title) + 1) if _existing_title is not None else list(body).index(first)
    for j, el in enumerate(new_elems):
        body.insert(idx + j, el)
    _write_part(path, "word/document.xml", root)
    styles_added = _ensure_toc_styles(path, {lvl for lvl, _, _ in entries})
    return {
        "changed": True,
        "entries": len(entries),
        "levels": max_level,
        "title": title_text,
        "title_source": ("document:existing" if _existing_title is not None else src),
        "styles_added": styles_added,
        "bookmarks": [name for _, _, name in entries],
        "converted_text_toc": _converted_text_toc,
    }


def verify_toc_generated(path, params: dict = None) -> bool:
    """复查：**允许生成**时目录必须存在（域或「目录」标题段）。

    EZI-0038 迁移语义：默认（未传 ``allow_generate``）本动作不造目录，
    「文档里没有 TOC 域」是**预期结果**，不能算校验失败——否则结构段会对
    每一份源无目录的文档记一条假失败（journal 里 `insert_toc_if_missing: failed`），
    把真实失败埋掉。
    """
    try:
        params = params or {}
        if params.get("enabled") is False:
            return True
        if not params.get("allow_generate"):
            return True
        root = _read_part(Path(path), "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return True
        # 没有可进目录的标题 = 无事可做（action 会 skip，不能算校验失败）
        if not _toc_headings_for_generation(body, 2, _style_outline_map(path)):
            return True
        has_field = False
        for x in body.iter(f"{W}instrText"):
            t = (x.text or "").strip().upper()
            if t.startswith("PAGEREF") or t.startswith("TOC"):
                has_field = True
                break
        if has_field:
            return True
        for p in body.iter(f"{W}p"):
            if p.getparent() is not body:
                continue
            norm = _toc_norm("".join(t.text or "" for t in p.iter(f"{W}t")))
            if norm in ("目录", "contents"):
                return True
        return False
    except Exception:
        return False


def sync_toc_page_numbers(path, params: dict = None) -> dict:
    """把目录条目页码改成**真机页码**（2026-09-26 定案：页码唯一真源 = Word 排版）。

    用户拍板：先排完其他内容、最后处理目录；条目与页码先正确、样式最后上；
    只保证交付那一刻正确（之后用户按 F9 属他自己的事）。版面以实物为准 ——
    格式改了页码就按改后的真实分页走，不再由我们推算。

    params: {"page_map": {bookmark: physical_page}, "offset": int|None}
      page_map 由 renderers.probe_bookmark_pages() 提供（_Toc* 书签 = 目录条目
      PAGEREF 的落点）；offset = 物理页 − 正文页，None 时自动推出：
      正文第 1 页 = 第一条目录条目所在物理页 → offset = 该页 − 1。
    只替换条目尾部（PAGEREF 域缓存结果）的数字，域结构与条目文本不动。
    """
    path = Path(path)
    params = params or {}
    page_map = {str(k): int(v) for k, v in (params.get("page_map") or {}).items()}
    if not page_map:
        return {"changed": False, "updated": 0, "reason": "empty page_map"}
    offset = params.get("offset")
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False, "updated": 0}
    entries = []
    for para in body.iter(f"{W}p"):
        name = _para_pageref_target(para)
        if name:
            entries.append((para, name))
    if not entries:
        return {"changed": False, "updated": 0, "reason": "no toc entries"}
    if offset is None:
        offset = _toc_page_offset(body, page_map, entries)
    offset = int(offset)
    updated = 0
    skipped = []
    for para, name in entries:
        page = page_map.get(name)
        if page is None:
            skipped.append(name)
            continue
        want = str(max(1, page - offset))
        digit_runs = []
        for r in para.iter(f"{W}r"):
            txt = "".join(x.text or "" for x in r.iter(f"{W}t"))
            if txt.strip() and txt.strip().isdigit():
                digit_runs.append((r, txt))
        if not digit_runs:
            skipped.append(name)
            continue
        run, txt = digit_runs[-1]
        if txt.strip() == want:
            continue
        for x in run.iter(f"{W}t"):
            x.text = want
        updated += 1
    if updated:
        _write_part(path, "word/document.xml", root)
    return {"changed": updated > 0, "updated": updated, "entries": len(entries),
            "offset": offset, "skipped": skipped}


def verify_toc_page_numbers(path, params: dict = None) -> bool:
    """复查：目录条目尾部的页码 == (真机物理页 − offset)。"""
    try:
        path = Path(path)
        params = params or {}
        page_map = {str(k): int(v) for k, v in (params.get("page_map") or {}).items()}
        if not page_map:
            return True
        root = _read_part(path, "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return False
        entries = []
        for para in body.iter(f"{W}p"):
            name = _para_pageref_target(para)
            if name:
                entries.append((para, name))
        offset = params.get("offset")
        if offset is None:
            offset = _toc_page_offset(body, page_map, entries)
        offset = int(offset)
        for para, name in entries:
            page = page_map.get(name)
            if page is None:
                continue
            want = str(max(1, page - offset))
            got = None
            for r in para.iter(f"{W}r"):
                txt = "".join(x.text or "" for x in r.iter(f"{W}t")).strip()
                if txt.isdigit():
                    got = txt
            if got != want:
                return False
        return True
    except Exception:
        return False


def keep_table_together(path, params: dict = None) -> dict:
    """整表同页 / 行内不切 / 长表表头续页（2026-09-26 定案，见 docs/表格跨页处理方案.md）。R2。

    用户拍板策略：
      · 表单/块状表（评语/成绩/签名/封面信息）：**整表同页**，宁可留白；
        内容超过一页 → 放开 keepNext 允许跨页（仍全行 cantSplit）。
      · 数据表（名册/对照/结果）：**允许跨页**，行内不切 + 首行识别为表头则
        tblHeader 续页重复（表单表硬排除，否则「成绩/教师评语」会重复冒出）。
      · 失败方向：宁可留白，不可切表 —— 估算灰区按「锁」处理；B（Word 真机分页）
        在交付层复核纠偏（A 估算 + B 纠偏 = 双保险）。

    说明：Word 没有「表格级-不跨页」属性，整表同页只能靠段级 keepNext 链。

    params: {table_index: -1(全部)|int, keep_ratio: 0.75, max_page_ratio: 1.0,
             mark_header: True}
    """
    path = Path(path)
    params = params or {}
    keep_ratio = float(params.get("keep_ratio", 0.75))
    max_ratio = float(params.get("max_page_ratio", 1.0))
    mark_header = bool(params.get("mark_header", True))
    force_keep = bool(params.get("force_keep", False))   # B 复核纠偏用：跳过估高闸
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False, "tables": 0}
    try:
        styles_root = _read_part(path, "word/styles.xml")
    except Exception:
        styles_root = None
    default_line = table_layout.default_doc_line_twips(styles_root)
    style_map = table_layout.build_style_map(styles_root)
    try:
        tables = list(_locate_tables(root, int(params.get("table_index", -1))))
    except ValueError:                                  # 文档里没有表格：无事可做
        return {"changed": False, "tables": 0, "kept_together": 0, "splittable": 0,
                "headers_marked": 0, "kinds": [], "est_heights": [],
                "usable_heights": [], "default_line": default_line}
    changed = 0
    kept = splittable = headers = 0
    kinds: list = []
    est_heights: list = []
    usable_heights: list = []
    for tbl in tables:
        trs = tbl.findall(f"{W}tr")
        if not trs:
            continue
        cols = table_layout.table_grid(tbl)
        usable = table_layout.usable_page_height(body, tbl)
        est = table_layout.estimate_table_height(tbl, default_line, style_map)
        cls = table_layout.classify_table(tbl, default_line)
        is_form = cls["kind"] == "form"
        # 灰区（keep_ratio ~ max_ratio）按「锁」处理：宁可留白，不可切表
        too_tall = est > int(usable * max_ratio)
        keep_whole = bool(is_form and (force_keep or not too_tall))
        kinds.append(cls["kind"])
        est_heights.append(est)
        usable_heights.append(usable)
        for ri, tr in enumerate(trs):
            trpr = tr.find(f"{W}trPr")
            if trpr is None:
                trpr = etree.Element(f"{W}trPr")
                tr.insert(0, trpr)
            row_h = table_layout.estimate_row_height(tr, cols, default_line, style_map)
            # 行内不切：仅当该行自身放得下一页（超页高的行加 cantSplit 会整行溢出）
            if row_h <= int(usable * 0.95) and trpr.find(f"{W}cantSplit") is None:
                etree.SubElement(trpr, f"{W}cantSplit")
                changed += 1
            if keep_whole:
                is_last = ri == len(trs) - 1
                for para in tr.iter(f"{W}p"):
                    ppr = para.find(f"{W}pPr")
                    if ppr is None:
                        ppr = etree.Element(f"{W}pPr")
                        para.insert(0, ppr)
                    if not is_last and ppr.find(f"{W}keepNext") is None:
                        etree.SubElement(ppr, f"{W}keepNext")
                        changed += 1
                    if ppr.find(f"{W}keepLines") is None:
                        etree.SubElement(ppr, f"{W}keepLines")
                        changed += 1
                    _normalize_child_order(ppr, _CT_PPR_ORDER)
            _normalize_child_order(trpr, _CT_TRPR_ORDER)
        if keep_whole:
            # 表前连续空段收敛为 1（2026-09-26 真机铁证：表格前 14 个空段 ×460twips
            # 在整表锁页后被整体顶出，撑出两页空白；源稿残留的废空段不该跟车交付）
            kids = list(body)
            try:
                pos = kids.index(tbl)
            except ValueError:
                pos = -1
            blanks = []
            j = pos - 1
            while j >= 0:
                el = kids[j]
                if etree.QName(el).localname != "p":
                    break
                txt = "".join(t.text or "" for t in el.iter(f"{W}t")).strip()
                has_obj = (el.find(f".//{W}drawing") is not None
                           or el.find(f".//{W}pict") is not None)
                if txt or has_obj:
                    break
                blanks.append(el)
                j -= 1
            for extra in blanks[1:]:
                body.remove(extra)
                changed += 1
            kept += 1
        else:
            splittable += 1
            if mark_header and table_layout.should_repeat_header(
                tbl, cls["kind"], default_line
            ):
                trpr0 = trs[0].find(f"{W}trPr")
                if trpr0 is not None and trpr0.find(f"{W}tblHeader") is None:
                    etree.SubElement(trpr0, f"{W}tblHeader")
                    _normalize_child_order(trpr0, _CT_TRPR_ORDER)
                    headers += 1
                    changed += 1
    if changed:
        _write_part(path, "word/document.xml", root)
    return {
        "changed": changed > 0,
        "tables": len(kinds),
        "kept_together": kept,
        "splittable": splittable,
        "headers_marked": headers,
        "kinds": kinds,
        "est_heights": est_heights,
        "usable_heights": usable_heights,
        "default_line": default_line,
    }


def verify_keep_table_together(path, params: dict = None) -> bool:
    """复查：按纯函数重算计划，逐表核对 XML 是否已落实（全行 cantSplit、
    被锁表的 keepNext 链、数据表表头 tblHeader）。"""
    try:
        path = Path(path)
        params = params or {}
        max_ratio = float(params.get("max_page_ratio", 1.0))
        force_keep = bool(params.get("force_keep", False))
        root = _read_part(path, "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return False
        try:
            styles_root = _read_part(path, "word/styles.xml")
        except Exception:
            styles_root = None
        default_line = table_layout.default_doc_line_twips(styles_root)
        style_map = table_layout.build_style_map(styles_root)
        try:
            tables = list(_locate_tables(root, int(params.get("table_index", -1))))
        except ValueError:                              # 没有表格 = 无需复查
            return True
        for tbl in tables:
            trs = tbl.findall(f"{W}tr")
            if not trs:
                continue
            cols = table_layout.table_grid(tbl)
            usable = table_layout.usable_page_height(body, tbl)
            est = table_layout.estimate_table_height(tbl, default_line, style_map)
            cls = table_layout.classify_table(tbl, default_line)
            keep_whole = cls["kind"] == "form" and (force_keep or est <= int(usable * max_ratio))
            want_header = (not keep_whole) and table_layout.should_repeat_header(
                tbl, cls["kind"], default_line)
            for ri, tr in enumerate(trs):
                trpr = tr.find(f"{W}trPr")
                row_h = table_layout.estimate_row_height(tr, cols, default_line, style_map)
                if row_h <= int(usable * 0.95):
                    if trpr is None or trpr.find(f"{W}cantSplit") is None:
                        return False
                if keep_whole and ri < len(trs) - 1:
                    for para in tr.iter(f"{W}p"):
                        ppr = para.find(f"{W}pPr")
                        if ppr is None or ppr.find(f"{W}keepNext") is None:
                            return False
            if want_header:
                trpr0 = trs[0].find(f"{W}trPr")
                if trpr0 is None or trpr0.find(f"{W}tblHeader") is None:
                    return False
        return True
    except Exception:
        return False


def set_row_layout(path, params: dict) -> dict:
    """设置行布局（§7.2 P1）：行高 hRule/val + cantSplit 防跨页。

    params: {table_index, rows: "all"|int, height: int（twips，hRule=atLeast）,
             cant_split: bool}。R2。
    """
    path = Path(path)
    height = params.get("height")
    cant_split = params.get("cant_split")
    if height is None and cant_split is None:
        raise ValueError("set_row_layout requires height or cant_split")
    root = _read_part(path, "word/document.xml")
    tables = _locate_tables(root, int(params.get("table_index", -1)))
    changed = 0
    for tbl in tables:
        trs = list(tbl.iter(f"{W}tr"))
        rows = trs if str(params.get("rows", "all")) == "all" else [trs[int(params["rows"])]]
        for tr in rows:
            trpr = tr.find(f"{W}trPr")
            if trpr is None:
                trpr = etree.Element(f"{W}trPr")
                tr.insert(0, trpr)
            if height is not None:
                th = trpr.find(f"{W}trHeight")
                if th is None:
                    th = etree.SubElement(trpr, f"{W}trHeight")
                th.set(f"{W}val", str(int(height)))
                th.set(f"{W}hRule", "atLeast")
            if cant_split is not None:
                cs = trpr.find(f"{W}cantSplit")
                if bool(cant_split):
                    if cs is None:
                        etree.SubElement(trpr, f"{W}cantSplit")
                else:
                    if cs is not None:
                        trpr.remove(cs)
            changed += 1
    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "rows": changed}


def verify_row_layout(path, params: dict) -> bool:
    """复查：行高/cantSplit 已写入。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        height = params.get("height")
        cant_split = params.get("cant_split")
        for tbl in _locate_tables(root, int(params.get("table_index", -1))):
            trs = list(tbl.iter(f"{W}tr"))
            rows = trs if str(params.get("rows", "all")) == "all" else [trs[int(params["rows"])]]
            for tr in rows:
                trpr = tr.find(f"{W}trPr")
                if height is not None:
                    th = trpr.find(f"{W}trHeight") if trpr is not None else None
                    if th is None or th.get(f"{W}val") != str(int(height)):
                        return False
                if cant_split is not None and bool(cant_split):
                    if trpr is None or trpr.find(f"{W}cantSplit") is None:
                        return False
        return True
    except Exception:
        return False


def set_cell_format(path, params: dict) -> dict:
    """设置单元格格式（§7.2 P1）：垂直对齐 + 单元格边距。

    params: {table_index, valign: "center"|"top"|"bottom",
             margins: {top/bottom/left/right: int（twips）}}。R1。
    """
    path = Path(path)
    valign = params.get("valign")
    margins = params.get("margins") or {}
    if not valign and not margins:
        raise ValueError("set_cell_format requires valign or margins")
    root = _read_part(path, "word/document.xml")
    tables = _locate_tables(root, int(params.get("table_index", -1)))
    changed = 0
    for tbl in tables:
        for tc in tbl.iter(f"{W}tc"):
            tcpr = tc.find(f"{W}tcPr")
            if tcpr is None:
                tcpr = etree.Element(f"{W}tcPr")
                tc.insert(0, tcpr)
            if valign:
                va = tcpr.find(f"{W}vAlign")
                if va is None:
                    va = etree.SubElement(tcpr, f"{W}vAlign")
                va.set(f"{W}val", valign)
            if margins:
                tcm = tcpr.find(f"{W}tcMar")
                if tcm is None:
                    tcm = etree.SubElement(tcpr, f"{W}tcMar")
                for edge, v in margins.items():
                    e = tcm.find(f"{W}{edge}")
                    if e is None:
                        e = etree.SubElement(tcm, f"{W}{edge}")
                    e.set(f"{W}w", str(int(v)))
                    e.set(f"{W}type", "dxa")
            changed += 1
    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "cells": changed}


def verify_cell_format(path, params: dict) -> bool:
    """复查：vAlign/tcMar 已写入全部单元格。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        valign = params.get("valign")
        for tbl in _locate_tables(root, int(params.get("table_index", -1))):
            for tc in tbl.iter(f"{W}tc"):
                tcpr = tc.find(f"{W}tcPr")
                if valign:
                    va = tcpr.find(f"{W}vAlign") if tcpr is not None else None
                    if va is None or va.get(f"{W}val") != valign:
                        return False
        return True
    except Exception:
        return False


def merge_cells(path, params: dict) -> dict:
    """合并单元格（§7.2 P1，R3 破坏性需授权）。

    params: {table_index, mode: "h"|"v", row: int, col: int, span: int}
    h: 水平合并（目标单元格 gridSpan=span）；v: 垂直合并（目标单元格
    vMerge restart，下方 span-1 格 vMerge continue）。R3。
    """
    path = Path(path)
    mode = params.get("mode")
    if mode not in ("h", "v"):
        raise ValueError("mode must be h or v")
    row = int(params["row"])
    col = int(params["col"])
    span = int(params.get("span", 2))
    if span < 2:
        raise ValueError("span must be >= 2")
    root = _read_part(path, "word/document.xml")
    tbl = _locate_tables(root, int(params.get("table_index", -1)))[0]
    trs = list(tbl.iter(f"{W}tr"))
    if row >= len(trs):
        raise ValueError(f"row {row} out of range")
    target_tr = trs[row]
    tcs = list(target_tr.iter(f"{W}tc"))
    if col >= len(tcs):
        raise ValueError(f"col {col} out of range")
    target_tc = tcs[col]
    tcpr = target_tc.find(f"{W}tcPr")
    if tcpr is None:
        tcpr = etree.Element(f"{W}tcPr")
        target_tc.insert(0, tcpr)
    if mode == "h":
        gs = tcpr.find(f"{W}gridSpan")
        if gs is None:
            gs = etree.SubElement(tcpr, f"{W}gridSpan")
        gs.set(f"{W}val", str(span))
    else:
        vm = tcpr.find(f"{W}vMerge")
        if vm is None:
            vm = etree.SubElement(tcpr, f"{W}vMerge")
        vm.set(f"{W}val", "restart")
        for r in range(row + 1, min(row + span, len(trs))):
            rtr = trs[r]
            rtcs = list(rtr.iter(f"{W}tc"))
            if col >= len(rtcs):
                continue
            rtc = rtcs[col]
            rtcpr = rtc.find(f"{W}tcPr")
            if rtcpr is None:
                rtcpr = etree.Element(f"{W}tcPr")
                rtc.insert(0, rtcpr)
            rvm = rtcpr.find(f"{W}vMerge")
            if rvm is None:
                rvm = etree.SubElement(rtcpr, f"{W}vMerge")
            rvm.set(f"{W}val", "continue")
    _write_part(path, "word/document.xml", root)
    return {"changed": True, "mode": mode, "row": row, "col": col, "span": span}


def verify_merged(path, params: dict) -> bool:
    """复查：gridSpan/vMerge 已写入。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        mode = params["mode"]
        row = int(params["row"])
        col = int(params["col"])
        tbl = _locate_tables(root, int(params.get("table_index", -1)))[0]
        trs = list(tbl.iter(f"{W}tr"))
        tcs = list(trs[row].iter(f"{W}tc"))
        tcpr = tcs[col].find(f"{W}tcPr")
        if tcpr is None:
            return False
        if mode == "h":
            gs = tcpr.find(f"{W}gridSpan")
            return gs is not None and gs.get(f"{W}val") == str(int(params.get("span", 2)))
        else:
            vm = tcpr.find(f"{W}vMerge")
            return vm is not None and vm.get(f"{W}val") == "restart"
    except Exception:
        return False


def split_cells(path, params: dict) -> dict:
    """拆分单元格（§7.2 P1，R3）：移除 gridSpan/vMerge。R3。"""
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    tables = _locate_tables(root, int(params.get("table_index", -1)))
    changed = 0
    for tbl in tables:
        for tc in tbl.iter(f"{W}tc"):
            tcpr = tc.find(f"{W}tcPr")
            if tcpr is None:
                continue
            gs = tcpr.find(f"{W}gridSpan")
            vm = tcpr.find(f"{W}vMerge")
            if gs is not None:
                tcpr.remove(gs)
                changed += 1
            if vm is not None:
                tcpr.remove(vm)
                changed += 1
    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "cells": changed}


def verify_split(path, params: dict) -> bool:
    """复查：无 gridSpan/vMerge。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        for tbl in _locate_tables(root, int(params.get("table_index", -1))):
            for tc in tbl.iter(f"{W}tc"):
                tcpr = tc.find(f"{W}tcPr")
                if tcpr is not None:
                    if tcpr.find(f"{W}gridSpan") is not None or tcpr.find(f"{W}vMerge") is not None:
                        return False
        return True
    except Exception:
        return False


def normalize_table_caption_block(path, params: dict) -> dict:
    """表标题块规范化（§7.2 P1）：表标题段（表 N.N 开头）keepNext 防跨页。R1。"""
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    styles_map = _read_styles_map(path)
    changed = 0
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False}
    cover_paras = _cover_paragraphs(root, styles_map)   # 封面表题不动
    for p in body.iter(f"{W}p"):
        if cover_paras and any(p is c for c in cover_paras):
            continue
        text = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        if not text:
            continue
        if re.match(r"^(表|Table|图|Figure|Fig)\s*[\dA-Za-z]+[-.]\d+", text):
            ppr = p.find(f"{W}pPr")
            if ppr is None:
                ppr = etree.Element(f"{W}pPr")
                p.insert(0, ppr)
            if ppr.find(f"{W}keepNext") is None:
                _ppr_insert_ordered(ppr, etree.Element(f"{W}keepNext"))
                changed += 1
            # 题注居中且不缩进：正文的「首行缩进 2 字符」规则会误伤表题
            # （2026-10-18 真机：图题无缩进、表题带 480 twips，同类两种形态）
            if ppr.find(f"{W}ind") is not None:
                ppr.remove(ppr.find(f"{W}ind"))
                changed += 1
    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "captions": changed}


def verify_table_caption_block(path, params: dict) -> bool:
    """复查：全部表标题段含 keepNext。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return False
        ok = True
        cover_paras = _cover_paragraphs(root, _read_styles_map(path))   # 与动作同条件
        for p in body.iter(f"{W}p"):
            if cover_paras and any(p is c for c in cover_paras):
                continue
            text = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
            if not text:
                continue
            if re.match(r"^(表|Table|图|Figure|Fig)\s*\d+", text):
                ppr = p.find(f"{W}pPr")
                if ppr is None or ppr.find(f"{W}keepNext") is None:
                    ok = False
        return ok
    except Exception:
        return False



# ══════════════════════════════════════════════════════════════════
# 图表标题 run 级格式（2026-08-16 真机回归：模板示例话术「表标题中文黑体
# 小4号、数字及字母 TNR 粗体小4号」/图标题数字加粗——数字/字母 run 加粗，
# 中文不加粗）
# ══════════════════════════════════════════════════════════════════







def remove_trailing_empty_paragraphs(path, params: dict) -> dict:
    """删除节尾连续纯空段——正文前空页根因（R1，内容操作 §10.1 声明）。

    2026-08-17 取证：目录节尾部 3 个纯空段（无文本/无 drawing）在转
    nextPage 后独自占 1 页（正文前空页）。删除 each 节尾（sectPr 段前）
    的连续纯空段；封面节（section 0）跳过（封面排版依赖留白）。
    只删「无内容」段：无 run 或 run 全空且无 drawing/pict/object。
    """
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False, "removed": 0}
    paras = list(body.findall(f"{W}p"))
    changed = 0
    first_sectpr = next(
        (i for i, pp in enumerate(paras) if pp.find(f"{W}pPr") is not None
         and pp.find(f"{W}pPr").find(f"{W}sectPr") is not None), -1)
    for idx in range(len(paras) - 1, -1, -1):
        p = paras[idx]
        ppr = p.find(f"{W}pPr")
        if ppr is None or ppr.find(f"{W}sectPr") is None:
            continue
        # 跳过封面节（第一个 sectPr 段）——封面排版依赖留白
        if idx == first_sectpr:
            continue
        j = idx - 1
        while j >= 0:
            prev = paras[j]
            if _is_empty_para(prev):
                body.remove(prev)
                changed += 1
                j -= 1
            else:
                break
    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "removed": changed}



def _is_empty_para(p) -> bool:
    """纯空段：run 全空或只有空 run；无 drawing/pict/object/fldChar。"""
    ppr = p.find(f"{W}pPr")
    if ppr is not None and ppr.find(f"{W}sectPr") is not None:
        return False
    for child in p.iter():
        if child.tag in (f"{W}drawing", f"{W}pict", f"{W}object", f"{W}fldChar"):
            return False
        # 公式段没有 w:t 文字（内容在 m:r 里）→ 会被误判成空行删除。
        # 2026-10-20 实测：remove_blank_lines_before_headings 把公式整段删掉
        # （源 1 → 产物 0，公式保全门拦下）。
        if child.tag in (f"{M}oMath", f"{M}oMathPara"):
            return False
    for t in p.iter(f"{W}t"):
        if (t.text or "").strip():
            return False
    return True


def _shrink_invisible_para(p) -> bool:
    """不可见段（域尾/节尾空段）极小化：1pt 固定行高 + 零段距 + 1pt 字号。

    这些段不能删（fldChar end 域结构 / sectPr 节结构），但目录条目
    加粗变高后它们占的整行会溢出到新页 → nextPage 分节叠加出空页
    （2026-08-18 r22 实证 p14）。压到 1pt 让其收回前页。
    幂等：已有同样值不重复写。返回是否发生修改。
    """
    changed = False
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        ppr = etree.Element(f"{W}pPr")
        p.insert(0, ppr)
        changed = True
    sp = ppr.find(f"{W}spacing")
    if sp is None:
        sp = etree.Element(f"{W}spacing")
        _ppr_insert_ordered(ppr, sp)
        changed = True
    want = {f"{W}before": "0", f"{W}after": "0",
            f"{W}line": "20", f"{W}lineRule": "exact"}
    for k, v in want.items():
        if sp.get(k) != v:
            sp.set(k, v)
            changed = True
    for r in p.findall(f"{W}r"):
        rpr = r.find(f"{W}rPr")
        if rpr is None:
            rpr = etree.Element(f"{W}rPr")
            r.insert(0, rpr)
            changed = True
        sz = rpr.find(f"{W}sz")
        if sz is None:
            sz = etree.Element(f"{W}sz")
            _rpr_insert_ordered(rpr, sz)
            changed = True
        if sz.get(f"{W}val") != "2":
            sz.set(f"{W}val", "2")
            changed = True
    return changed


def verify_trailing_empty_paragraphs(path, params: dict) -> bool:
    """复查：节尾（sectPr 段前）无连续空段。"""
    try:
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return False
        paras = list(body.findall(f"{W}p"))
        for idx, p in enumerate(paras):
            ppr = p.find(f"{W}pPr")
            if ppr is None or ppr.find(f"{W}sectPr") is None:
                continue
            j = idx - 1
            while j >= 0 and _is_empty_para(paras[j]):
                return False
        return True
    except Exception:
        return False


def _body_level_paras(body) -> list:
    """body 直属段落列表——表内等非直属段不参与（动作与 verifier 同源）。

    表内段（w:tc 内）不是 body 的子节点：body.remove 对非直属节点抛
    ValueError（2026-10-14 审计 B2 实测），且「标题前空段」判定不该把
    表内空段当待删对象/违规对象。
    """
    return [p for p in body.iter(f"{W}p") if p.getparent() is body]


def remove_blank_lines_before_headings(path, params: dict) -> dict:
    """删除标题段（outlineLvl 0/1/2）前面连续的空段（2026-08-18 用户
    彻查「1.3 研究内容与主要工作前面的两个空白行」第三遍）。

    实证：源文档 idx=140/141 两个空行夹在「1.3.2 国外研究现状」正文段
    与「1.3 研究内容与主要工作」标题之间——源文档作者手打空行，earlyz
    保留导致标题前出现两个空白行。节尾/节首空段已有边界清理，这里补
    正文中间「标题前空段」清理。

    规则：outlineLvl 段（标题）向前连续删除 _is_empty_para 段；但
    ① 不跨过 sectPr 段（节边界保护）② 不跨过含 drawing/pict 段
    ③ 最多删到段落计数上限（防误删封面排版留白——封面无 outlineLvl
    标题，天然免疫）。
    """
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False, "removed": 0}

    # 只取 body 直属段落（与 verifier 同源）：表内段不进列表，不删也不跨
    paras = _body_level_paras(body)
    removed = 0
    for idx, p in enumerate(paras):
        ppr = p.find(f"{W}pPr")
        oLvl = ppr.find(f"{W}outlineLvl") if ppr is not None else None
        if oLvl is None:
            continue
        # 向前删连续空段
        j = idx - 1
        while j >= 0:
            prev = paras[j]
            prev_ppr = prev.find(f"{W}pPr")
            # 保护：非 body 直属段（表内段）不删不跨——body.remove 对非直属
            # 子节点抛 ValueError（2026-10-14 审计 B2 实测）
            if prev.getparent() is not body:
                break
            # 保护：节边界（含 sectPr）不跨
            if prev_ppr is not None and prev_ppr.find(f"{W}sectPr") is not None:
                break
            if _is_empty_para(prev):
                body.remove(prev)
                removed += 1
                j -= 1
            else:
                break

    if removed:
        _write_part(path, "word/document.xml", root)
    return {"changed": removed > 0, "removed": removed}


def verify_no_blank_before_headings(path, params: dict) -> bool:
    """复查：标题段（outlineLvl）前无连续空段。"""
    try:
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return False
        # 与动作同源：只取 body 直属段落——表内空段（有 w:tbl 祖先）不算违规
        paras = _body_level_paras(body)
        for idx, p in enumerate(paras):
            ppr = p.find(f"{W}pPr")
            oLvl = ppr.find(f"{W}outlineLvl") if ppr is not None else None
            if oLvl is None:
                continue
            j = idx - 1
            while j >= 0:
                prev = paras[j]
                prev_ppr = prev.find(f"{W}pPr")
                # 非 body 直属段（表内段）不判违规——与动作同一保护
                if prev.getparent() is not body:
                    break
                if prev_ppr is not None and prev_ppr.find(f"{W}sectPr") is not None:
                    break
                if _is_empty_para(prev):
                    return False
                break
        return True
    except Exception:
        return False

def normalize_footer_parts(path, params: dict) -> dict:
    """页脚单行化：每 footer part 重建为单个居中 PAGE 域段（R1）。

    2026-08-16 真机取证：源文档页脚 = 文本框 + sdt 双 PAGE 域 + 多个
    空段（footer5 渲染出两行页码 + 多换行）。整 part 重建为单段标准
    PAGE 域（fldChar begin/instrText/separate/t/end，居中 5 号 TNR）；
    无 PAGE 域的空页脚（封面）只保留 1 个空段。仅改 part 内容，不动
    document.xml 引用（新增 footerReference 触发 Word 损坏——实证）。
    """
    path = Path(path)
    changed = 0
    with zipfile.ZipFile(path) as z:
        footers = [n for n in z.namelist() if re.match(r"word/footer\d+\.xml", n)]
    for fp in sorted(footers):
        root = _read_part(path, fp)
        has_page = any(
            "PAGE" in (it.text or "") for it in root.iter(f"{W}instrText"))
        if _footer_is_single(root, has_page):
            continue
        _rebuild_footer(root, has_page)
        _write_part(path, fp, root)
        changed += 1
    return {"changed": changed > 0, "footers_fixed": changed}


# 页码页脚缓存文本：数字或罗马数字（'II'/'VV'/'22'/'1313'，WPS 双 PAGE 域）
_PAGE_TEXT_RE = re.compile(r"^\s*[0-9IVXLCDMivxlcdm]+(?:\s*[-—–]\s*[0-9IVXLCDMivxlcdm]+)?\s*$")


def _footer_is_single(root, has_page: bool) -> bool:
    """页脚是否已单行化：单段；有 PAGE 时含完整域链。

    内容页脚保护（2026-08-22 用户裁决「页脚支持为零」）：页脚含真实
    文字（汉字/英文单词，如收稿日期/基金/作者简介）→ 是内容页脚，
    不是页码页脚——一律视为已达标（跳过重建），绝不整 part 清空。

    页码页脚 = 文本框/sdt 双 PAGE 域缓存（页码占位，'I'/'V'/'13'）。
    内容页脚判定按**位置**：只统计普通段落（根级 w:p）内的文本——
    文本框（txbxContent）与 sdt 内文本是页码域缓存，不是正文内容。
    若把罗马数字缓存也当内容页脚保护，文本框（wp:docPr）被克隆至
    多节时 id 重复导致 Word 拒开（2026-08-22 单栏文档「文件可能已
    经损坏」实锤）。
    """
    paras = [c for c in root if c.tag == f"{W}p"]
    text_all = "".join(t.text or "" for t in root.iter(f"{W}t")).strip()
    # 普通段落文本（排除文本框/sdt 内的页码缓存）
    body_text = ""
    for t in root.iter(f"{W}t"):
        anc = t.getparent()
        in_obj = False
        while anc is not None:
            if etree.QName(anc).localname in (
                "txbxContent", "sdtContent", "drawing", "pict", "textbox"):
                in_obj = True
                break
            anc = anc.getparent()
        if not in_obj:
            body_text += t.text or ""
    body_text = body_text.strip()
    if body_text and not _PAGE_TEXT_RE.fullmatch(body_text):
        return True  # 内容页脚保护：普通段含真实文字
    if text_all and not _PAGE_TEXT_RE.fullmatch(text_all):
        # 文本全在文本框/sdt 且非页码形态（异常结构）→ 保守保护
        return True
    if len(paras) != 1:
        return False
    if not has_page:
        return True
    flds = [f.get(f"{W}fldCharType") for f in paras[0].iter(f"{W}fldChar")]
    instr = any("PAGE" in (it.text or "") for it in paras[0].iter(f"{W}instrText"))
    return instr and flds[:3] == ["begin", "separate", "end"]


def _rebuild_footer(root, has_page: bool) -> None:
    """重建页脚为单段：居中 + 5 号 TNR；有 PAGE 时补标准域链。"""
    for child in list(root):
        root.remove(child)
    p = etree.SubElement(root, f"{W}p")
    ppr = etree.SubElement(p, f"{W}pPr")
    jc = etree.SubElement(ppr, f"{W}jc")
    jc.set(f"{W}val", "center")
    if has_page:
        rpr_p = etree.SubElement(ppr, f"{W}rPr")
        _set_tnr(rpr_p)
        _set_sz(rpr_p)

        def _run():
            r = etree.SubElement(p, f"{W}r")
            rr = etree.SubElement(r, f"{W}rPr")
            _set_tnr(rr)
            _set_sz(rr)
            return r

        r1 = _run()
        fc = etree.SubElement(r1, f"{W}fldChar")
        fc.set(f"{W}fldCharType", "begin")
        r2 = _run()
        it = etree.SubElement(r2, f"{W}instrText")
        it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        it.text = " PAGE "
        r3 = _run()
        fc2 = etree.SubElement(r3, f"{W}fldChar")
        fc2.set(f"{W}fldCharType", "separate")
        r4 = _run()
        t = etree.SubElement(r4, f"{W}t")
        t.text = "1"
        r5 = _run()
        fc3 = etree.SubElement(r5, f"{W}fldChar")
        fc3.set(f"{W}fldCharType", "end")


def _set_tnr(rpr) -> None:
    rf = etree.SubElement(rpr, f"{W}rFonts")
    rf.set(f"{W}ascii", "Times New Roman")
    rf.set(f"{W}hAnsi", "Times New Roman")
    rf.set(f"{W}cs", "Times New Roman")


def _set_sz(rpr) -> None:
    sz = etree.SubElement(rpr, f"{W}sz")
    sz.set(f"{W}val", "21")
    szcs = etree.SubElement(rpr, f"{W}szCs")
    szcs.set(f"{W}val", "21")


def verify_footer_parts(path, params: dict) -> bool:
    """复查：所有页脚均单段且（有 PAGE 时）域链完整。"""
    try:
        path = Path(path)
        with zipfile.ZipFile(path) as z:
            footers = [n for n in z.namelist() if re.match(r"word/footer\d+\.xml", n)]
        for fp in sorted(footers):
            root = _read_part(path, fp)
            has_page = any(
                "PAGE" in (it.text or "") for it in root.iter(f"{W}instrText"))
            if not _footer_is_single(root, has_page):
                return False
        return True
    except Exception:
        return False


def format_table_cells(path, params: dict) -> dict:
    """表内文字按当前模板 table_cells 取值（缺则跳过，禁止写死 5 号）。"""
    path = Path(path)
    tc = (params or {}).get("table_cells") or {}
    want_sz = tc.get("sz")
    want_ascii = tc.get("font_ascii")
    want_ea = tc.get("font_east_asia")
    if not want_sz and not want_ascii and not want_ea:
        return {"changed": False, "cells_fixed": 0, "skipped": "no template table_cells"}
    root = _read_part(path, "word/document.xml")
    changed = 0
    for tbl in root.iter(f"{W}tbl"):
        for para in tbl.iter(f"{W}p"):
            if _table_cell_font_ok(para, want_sz, want_ascii, want_ea):
                continue
            if _apply_cell_font(para, want_sz, want_ascii, want_ea):
                changed += 1
    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "cells_fixed": changed}


def _table_cell_font_ok(para, want_sz=None, want_ascii=None, want_ea=None) -> bool:
    for run in para.findall(f"{W}r"):
        if not "".join(x.text or "" for x in run.iter(f"{W}t")).strip():
            continue
        rpr = run.find(f"{W}rPr")
        if rpr is None:
            return False
        if want_sz:
            sz = rpr.find(f"{W}sz")
            sc = rpr.find(f"{W}szCs")
            if (sz is None or sz.get(f"{W}val") != want_sz) and (
                sc is None or sc.get(f"{W}val") != want_sz
            ):
                return False
        rf = rpr.find(f"{W}rFonts")
        if want_ascii:
            if rf is None or rf.get(f"{W}ascii") != want_ascii:
                return False
        if want_ea:
            if rf is None or rf.get(f"{W}eastAsia") != want_ea:
                return False
    return True


def _apply_cell_font(para, want_sz=None, want_ascii=None, want_ea=None) -> bool:
    changed = False
    for run in para.findall(f"{W}r"):
        if not "".join(x.text or "" for x in run.iter(f"{W}t")).strip():
            continue
        rpr = run.find(f"{W}rPr")
        if rpr is None:
            rpr = etree.Element(f"{W}rPr")
            run.insert(0, rpr)
        if want_ascii or want_ea:
            rf = rpr.find(f"{W}rFonts")
            if rf is None:
                rf = etree.Element(f"{W}rFonts")
                rpr.insert(0, rf)
            if want_ascii:
                rf.set(f"{W}ascii", want_ascii)
                rf.set(f"{W}hAnsi", want_ascii)
                rf.set(f"{W}cs", want_ascii)
            if want_ea:
                rf.set(f"{W}eastAsia", want_ea)
        if want_sz:
            for tag in ("sz", "szCs"):
                el = rpr.find(f"{W}{tag}")
                if el is None:
                    el = etree.Element(f"{W}{tag}")
                    rpr.append(el)
                if el.get(f"{W}val") != want_sz:
                    el.set(f"{W}val", want_sz)
        changed = True
    return changed


def verify_table_cells(path, params: dict) -> bool:
    """复查：表内段符合模板 table_cells；无模板取值则跳过。"""
    try:
        tc = (params or {}).get("table_cells") or {}
        want_sz = tc.get("sz")
        want_ascii = tc.get("font_ascii")
        want_ea = tc.get("font_east_asia")
        if not want_sz and not want_ascii and not want_ea:
            return True
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        for tbl in root.iter(f"{W}tbl"):
            for para in tbl.iter(f"{W}p"):
                if not _table_cell_font_ok(para, want_sz, want_ascii, want_ea):
                    return False
        return True
    except Exception:
        return False


def _set_border_edge(parent, edge: str, val: str, sz: str, color: str) -> None:
    el = etree.SubElement(parent, f"{W}{edge}")
    el.set(f"{W}val", val)
    el.set(f"{W}sz", sz)
    el.set(f"{W}space", "0")
    el.set(f"{W}color", color)


def apply_table_borders(path, params: dict) -> dict:
    """把当前模板抽出的表框套到产物每一张表（无模板边框则跳过）。"""
    path = Path(path)
    tb = (params or {}).get("table_borders") or {}
    style = tb.get("style")
    if style not in ("three_line", "solid"):
        return {"changed": False, "tables": 0, "skipped": "no template table_borders"}
    if style == "three_line" and not tb.get("top_sz"):
        return {"changed": False, "tables": 0, "skipped": "three_line missing top_sz"}
    color = tb.get("color") or "000000"
    if color == "auto":
        color = "000000"
    top_sz = str(tb.get("top_sz") or tb.get("sz") or "12")
    bot_sz = str(tb.get("bottom_sz") or tb.get("sz") or top_sz)
    header_sz = str(tb.get("header_sz") or "")
    solid_sz = str(tb.get("sz") or "4")
    root = _read_part(path, "word/document.xml")
    n = 0
    for tbl in root.iter(f"{W}tbl"):
        tbl_pr = tbl.find(f"{W}tblPr")
        if tbl_pr is None:
            tbl_pr = etree.Element(f"{W}tblPr")
            tbl.insert(0, tbl_pr)
        old = tbl_pr.find(f"{W}tblBorders")
        if old is not None:
            tbl_pr.remove(old)
        borders = etree.Element(f"{W}tblBorders")
        if style == "three_line":
            _set_border_edge(borders, "top", "single", top_sz, color)
            _set_border_edge(borders, "bottom", "single", bot_sz, color)
            for edge in ("left", "right", "insideH", "insideV"):
                _set_border_edge(borders, edge, "none", "0", "auto")
        else:
            for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
                _set_border_edge(borders, edge, "single", solid_sz, color)
        _tblpr_insert_ordered(tbl_pr, borders)
        rows = list(tbl.iter(f"{W}tr"))
        for ri, tr in enumerate(rows):
            for tc in tr.iter(f"{W}tc"):
                tc_pr = tc.find(f"{W}tcPr")
                if tc_pr is None:
                    tc_pr = etree.Element(f"{W}tcPr")
                    tc.insert(0, tc_pr)
                old_tcb = tc_pr.find(f"{W}tcBorders")
                if old_tcb is not None:
                    tc_pr.remove(old_tcb)
                if style == "three_line" and ri == 0 and tb.get("header_sz"):
                    tcb = etree.Element(f"{W}tcBorders")
                    _set_border_edge(tcb, "bottom", "single", header_sz, color)
                    _insert_by_seq(tc_pr, tcb, [
                        "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge",
                        "tcBorders", "shd", "noWrap", "tcMar",
                        "textDirection", "tcFitText", "vAlign", "hideMark",
                    ])
                elif style == "solid":
                    tcb = etree.Element(f"{W}tcBorders")
                    for edge in ("top", "left", "bottom", "right"):
                        _set_border_edge(tcb, edge, "single", solid_sz, color)
                    _insert_by_seq(tc_pr, tcb, ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge",
                                                "tcBorders", "shd", "noWrap", "tcMar",
                                                "textDirection", "tcFitText", "vAlign", "hideMark"])
        n += 1
    if n:
        _write_part(path, "word/document.xml", root)
    return {"changed": n > 0, "tables": n, "style": style}


def verify_table_borders(path, params: dict) -> bool:
    try:
        tb = (params or {}).get("table_borders") or {}
        style = tb.get("style")
        if style not in ("three_line", "solid"):
            return True
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        tbls = list(root.iter(f"{W}tbl"))
        if not tbls:
            return True
        want_top = str(tb.get("top_sz") or tb.get("sz") or "12")
        for tbl in tbls:
            tbl_pr = tbl.find(f"{W}tblPr")
            borders = tbl_pr.find(f"{W}tblBorders") if tbl_pr is not None else None
            if borders is None:
                return False
            top = borders.find(f"{W}top")
            left = borders.find(f"{W}left")
            if top is None:
                return False
            if style == "three_line":
                if top.get(f"{W}sz") != want_top:
                    return False
                if left is not None and left.get(f"{W}val") not in (None, "none", "nil"):
                    return False
            else:
                if left is None or left.get(f"{W}val") in (None, "none", "nil"):
                    return False
        return True
    except Exception:
        return False


_ROMAN_PAIRS = ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
               (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
               (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"))
_ROMAN_VALUES = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}


def int_to_roman(n: int) -> str:
    if n <= 0:
        return ""
    out = []
    for v, sym in _ROMAN_PAIRS:
        while n >= v:
            out.append(sym)
            n -= v
    return "".join(out)


def roman_to_int(s: str) -> int:
    total, prev = 0, 0
    for ch in reversed(s.upper()):
        v = _ROMAN_VALUES.get(ch)
        if v is None:
            return 0
        total += -v if v < prev else v
        prev = max(prev, v)
    return total


# 表格题注：TABLE I / Table 3. / Tab. 2 / 表 1（图题不在此列）
_TABLE_CAP_RE = re.compile(
    r"^(?P<pre>TABLE|Table|Tab\.|表)\s*(?P<num>[0-9]+|[IVXLCDMivxlcdm]+)"
    r"(?P<sep>[\s.、:：\-]*)")
# 正文引用（「表 1 所示…」）不能当题注
_CAP_TAIL_STOP = ("所示", "中", "为", "给出", "列出", "显示", "是", "表示", "描述", "可见")


def _table_caption_match(text: str):
    """表格题注匹配（收紧：短段、无句读、尾部不是引用词）。"""
    t = (text or "").strip()
    if not t or len(t) > 90:
        return None
    if re.search(r"[，。；：、]", t):
        return None
    m = _TABLE_CAP_RE.match(t)
    if not m:
        return None
    tail = t[m.end():].strip()
    if any(tail.startswith(w) for w in _CAP_TAIL_STOP):
        return None
    if not tail:
        # IEEE 题注常见「TABLE I」独占一行、题名在下一段；中文「表 1」独占行少见
        # （多是小节标签）→ 只放行英文前缀。
        if m.group("pre") in ("TABLE", "Table", "Tab."):
            return m
        return None
    return m


def _replace_text_regex(para, pattern, new_num: str) -> int:
    """按正则命中（第 3 组=编号）替换段落里的编号 token。"""
    from earlyz.actions.captions import _replace_number_span
    text = "".join(x.text or "" for x in para.iter(f"{W}t"))
    m = pattern.search(text)
    if not m:
        return 0
    start, end = m.span(3)
    err = _replace_number_span(para, start, end, new_num)
    return 0 if err else 1


def set_table_caption_numbering(path, params: dict) -> dict:
    """表格题注编号（支持罗马数字/阿拉伯数字，2026-10-17）。

    模板口径：IEEE「Tables should be numbered with Roman numerals」→ style=roman。
    只改**题注的编号 token**，题注文字、图题、正文引用一律不动；正文里对
    「表 N / Table N」的引用会统计后如实返回（引用是否要跟着改由用户定）。
    """
    style = ((params or {}).get("style") or "roman").lower()
    if style not in ("roman", "arabic"):
        return {"changed": False, "error": f"不支持的编号样式: {style}"}
    target = Path(path)
    r = _read_part(target, "word/document.xml")
    styles = _read_styles_map(target)
    cover = _cover_paragraphs(r, styles)
    changed, total, rows = 0, 0, []
    for para in r.iter(f"{W}p"):
        if cover and any(para is c for c in cover):
            continue
        if any(a.tag == f"{W}tbl" for a in para.iterancestors()):
            continue
        text = "".join(x.text or "" for x in para.iter(f"{W}t"))
        m = _table_caption_match(text)
        if m is None:
            continue
        total += 1
        num_tok = m.group("num")
        want = int_to_roman(total) if style == "roman" else str(total)
        if num_tok.upper() == want.upper():
            rows.append({"index": total, "number": want, "old": num_tok, "changed": False})
            continue
        from earlyz.actions.captions import _replace_number_span
        err = _replace_number_span(para, m.start("num"), m.end("num"), want)
        if err:
            rows.append({"index": total, "number": want, "changed": False,
                         "error": err, "text": text[:40]})
            continue
        changed += 1
        rows.append({"index": total, "number": want, "old": num_tok, "changed": True})
    # 引用同步（C，2026-10-17）：题注改号后正文引用跟着改（可关：follow_references=False）
    follow = (params or {}).get("follow_references", True)
    updated = 0
    ref_map = {str(r["old"]): str(r["number"]) for r in rows if r["changed"]}
    if follow and ref_map:
        for para in r.iter(f"{W}p"):
            if cover and any(para is c for c in cover):
                continue
            if any(a.tag == f"{W}tbl" for a in para.iterancestors()):
                continue
            if _table_caption_match("".join(x.text or "" for x in para.iter(f"{W}t"))) is not None:
                continue
            for old_num, new_num in ref_map.items():
                pat = re.compile(r"(表|TABLE|Table|Tab\.)(\s*)(" + re.escape(old_num) + r")(?![0-9])")
                if not pat.search("".join(x.text or "" for x in para.iter(f"{W}t"))):
                    continue
                if _replace_text_regex(para, pat, new_num) > 0:
                    updated += 1

    # 正文引用统计（只报不改）
    refs = []
    for para in r.iter(f"{W}p"):
        if cover and any(para is c for c in cover):
            continue
        if any(a.tag == f"{W}tbl" for a in para.iterancestors()):
            continue
        text = "".join(x.text or "" for x in para.iter(f"{W}t"))
        if _table_caption_match(text) is not None:
            continue
        for rm in re.finditer(r"(?:表|TABLE|Table)\s*([0-9]+|[IVXLCDM]+)", text):
            refs.append({"number": rm.group(1), "text": text[max(0, rm.start() - 10):rm.end() + 10]})
    if changed:
        _write_part(target, "word/document.xml", r)
    return {"changed": changed > 0 or updated > 0, "style": style, "total": total,
            "renumbered": changed, "references_updated": updated, "captions": rows[:50],
            "reference_hits": len(refs), "references": refs[:10]}


def verify_table_caption_numbering(path, params: dict = None) -> bool:
    style = ((params or {}).get("style") or "roman").lower()
    r = _read_part(Path(path), "word/document.xml")
    styles = _read_styles_map(Path(path))
    cover = _cover_paragraphs(r, styles)
    idx = 0
    for para in r.iter(f"{W}p"):
        if cover and any(para is c for c in cover):
            continue
        if any(a.tag == f"{W}tbl" for a in para.iterancestors()):
            continue
        text = "".join(x.text or "" for x in para.iter(f"{W}t"))
        m = _table_caption_match(text)
        if m is None:
            continue
        idx += 1
        want = int_to_roman(idx) if style == "roman" else str(idx)
        if m.group("num").upper() != want.upper():
            return False
    return True


def _is_caption_para(text: str) -> bool:
    """图表标题段判定（收紧版，2026-08-17 投诉④：正文「图6-1为美妆…」被误判）。

    标题段特征：短段（≤50 字符）且无句读标点（逗号/句号/分号/顿号/冒号）。
    正文引用段（「图6-1为美妆护肤产品电商平台管理系统登录模块…」）天然排除。
    """
    t = (text or "").strip()
    if not t or len(t) > 50:
        return False
    if re.search(r"[，。；、：]", t):
        return False
    return bool(re.match(r"^(图|表)\s*\d", t))


def format_caption_numbers(path, params: dict) -> dict:
    """图表标题段数字/字母加粗（R1）。

    params: {scope: "figure"|"table"|"all"}——按段文本前缀（图N/表N）识别。
    模板话术：表标题中文黑体小4号、数字及字母 TNR 粗体小4号；图标题
    数字字母粗体小4号。中文 run 不加粗。
    标题段判定用 _is_caption_para（正文引用段不误伤）。
    """
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    scope = (params or {}).get("scope", "all")
    changed = 0
    cover_paras = _cover_paragraphs(root, _read_styles_map(path))   # 封面题号不动
    for para in root.iter(f"{W}p"):
        if cover_paras and any(para is c for c in cover_paras):
            continue
        text = "".join(x.text or "" for x in para.iter(f"{W}t")).strip()
        if not text:
            continue
        is_fig = bool(re.match(r"^图\s*\d", text))
        is_tbl = bool(re.match(r"^表\s*\d", text))
        if scope == "figure" and not is_fig:
            continue
        if scope == "table" and not is_tbl:
            continue
        if not (is_fig or is_tbl):
            continue
        if not _is_caption_para(text):
            continue
        number_bold, latin_bold = _caption_bold_flags(
            (params or {}).get("caption") or {}, "table" if is_tbl else "figure")
        if _bold_digits_letters(para, latin_bold=latin_bold, number_bold=number_bold):
            changed += 1
    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "captions_fixed": changed}


def _caption_bold_flags(cap: dict, kind: str) -> tuple[bool, object]:
    """题注加粗取值（handler / verifier 同一入口）→ (number_bold, latin_bold)。

    分类键（`table_latin_bold` / `figure_number_bold`，structure 按该类题注自己的
    模板样本算）优先；该类无样本时回落合池键（`number_bold` / `latin_bold`）。
    number_bold 缺省 True、latin_bold 缺省 None（= 去粗），与旧口径一致。
    """
    nb = cap.get(f"{kind}_number_bold", cap.get("number_bold"))
    lb = cap.get(f"{kind}_latin_bold", cap.get("latin_bold"))
    return nb is not False, lb


def _caption_numbers_bold(para) -> bool:
    """图/表序号（含数字的 run）是否已加粗。纯英文 run 不加粗不算缺。"""
    for run in para.findall(f"{W}r"):
        if any(run.find(f".//{W}{tag}") is not None
               for tag in ("drawing", "object", "pict")):
            continue
        rpr = run.find(f"{W}rPr")
        t = "".join(x.text or "" for x in run.findall(f"{W}t"))
        if re.search(r"[0-9]", t) and not re.search(r"[\u4e00-\u9fff]", t):
            if rpr is None or rpr.find(f"{W}b") is None:
                return False
    return True


def _bold_digits_letters(para, latin_bold=None, number_bold=True) -> int:
    """只加粗图/表序号（5-1 / 4.2），不加粗标题里的英文（brands）。

    模板-4 实证：表/图字不加粗，仅序号 run 加粗；Altera/Symbol/HBsAg
    等英文在标题里不加粗。旧逻辑把 [A-Za-z] 当「数字」整段加粗。
    """
    count = 0
    for run in list(para.findall(f"{W}r")):
        # Drawing/OLE/textbox text belongs to the complex object. Descendant
        # w:t must never be flattened into the host run during run splitting.
        if any(run.find(f".//{W}{tag}") is not None
               for tag in ("drawing", "object", "pict")):
            continue
        t = "".join(x.text or "" for x in run.findall(f"{W}t"))
        if not t:
            continue
        has_digit = bool(re.search(r"[0-9]", t))
        has_cjk = bool(re.search(r"[\u4e00-\u9fff]", t))
        has_latin = bool(re.search(r"[A-Za-z]", t))
        rpr = run.find(f"{W}rPr")
        if has_digit and not has_cjk:
            # 纯序号 run（5 / 5-1 / 4.2）：按当前模板加粗/去粗。
            if number_bold:
                if rpr is None:
                    rpr = etree.Element(f"{W}rPr")
                    run.insert(0, rpr)
                if rpr.find(f"{W}b") is None:
                    etree.SubElement(rpr, f"{W}b")
                    etree.SubElement(rpr, f"{W}bCs")
                    count += 1
            elif rpr is not None:
                for tag in ("b", "bCs"):
                    old = rpr.find(f"{W}{tag}")
                    if old is not None:
                        rpr.remove(old)
                        count += 1
            continue
        if has_latin and not has_digit:
            if latin_bold is True:
                continue
            # 纯英文（brands）：模板示例未加粗则去掉误加的粗
            if rpr is not None:
                for tag in ("b", "bCs"):
                    el = rpr.find(f"{W}{tag}")
                    if el is not None:
                        rpr.remove(el)
                        count += 1
            continue
        if has_cjk and has_digit:
            # 混合 run：只拆出纯数字段加粗，英文段不加
            parts = re.split(r"([0-9.．\-－—]+)", t)
            if len(parts) < 2:
                continue
            if rpr is None:
                rpr = run.find(f"{W}rPr")
            first = parts[0]
            parent = run.getparent()
            if parent is None:
                continue
            # ⚠ 2026-11-08 用户实锤「图5-1 收敛性分析 → 只剩一个图字」：
            # 同一个题注被拆成多个 run 时（"图" | "5-1 收敛性分析"），第二个 run
            # 的 parts[0] 为空 → 老代码把 run 从父节点摘掉，却仍拿它当 prev 往里挂
            # 新 run —— 新 run 全挂在**已经脱离文档**的节点上，整段文字凭空消失。
            # 正确做法：摘掉前先记住插入锚点（前一个兄弟），后续按锚点顺序插入。
            anchor = run.getprevious()
            if first:
                _set_run_text(run, first)
                prev = run
            else:
                parent.remove(run)
                prev = anchor
            for seg in parts[1:]:
                if not seg:
                    continue
                # 序号段必须含数字：单独的连字符（「E-R」里的 -）不是序号（2026-09-29）
                is_num = (bool(re.match(r"^[0-9.．\-－—]+$", seg))
                          and bool(re.search(r"[0-9]", seg)))
                new_run = etree.Element(f"{W}r")
                if rpr is not None:
                    new_rpr = copy.deepcopy(rpr)
                    # 只动序号段与英文段的加粗（与整 run 分支 / verifier 同口径）；
                    # 中文段保留原 run 的加粗——中文加粗归角色规则（caption_table
                    # bold）管，此前一律剥粗会把模板要求整段加粗的表题中文剥掉。
                    _latin_seg = (bool(re.search(r"[A-Za-z]", seg))
                                  and not re.search(r"[0-9]", seg))
                    if is_num or (_latin_seg and latin_bold is not True):
                        for tag in ("b", "bCs"):
                            old = new_rpr.find(f"{W}{tag}")
                            if old is not None:
                                new_rpr.remove(old)
                    if is_num and number_bold:
                        _rpr_insert_ordered(new_rpr, etree.Element(f"{W}b"))
                        _rpr_insert_ordered(new_rpr, etree.Element(f"{W}bCs"))
                    new_run.append(new_rpr)
                elif is_num and number_bold:
                    new_rpr = etree.Element(f"{W}rPr")
                    etree.SubElement(new_rpr, f"{W}b")
                    etree.SubElement(new_rpr, f"{W}bCs")
                    new_run.append(new_rpr)
                new_t = etree.SubElement(new_run, f"{W}t")
                new_t.text = seg
                if prev is None:                    # run 原本是段落第一个子节点
                    parent.insert(0, new_run)
                else:
                    prev.addnext(new_run)
                prev = new_run
                if is_num:
                    count += 1
    return count


def _set_run_text(run, text: str) -> None:
    for t in run.findall(f"{W}t"):
        run.remove(t)
    nt = etree.SubElement(run, f"{W}t")
    nt.text = text


def verify_caption_numbers(path, params: dict) -> bool:
    """复查：所有图表标题段的数字/字母 run 已加粗。

    判定范围与 handler 的 _is_caption_para 一致（2026-08-17 收紧后同步：
    正文引用段「图6-1为…」不是标题，不检查加粗）。
    """
    try:
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        cap = (params or {}).get("caption") or {}
        cover_paras = _cover_paragraphs(root, _read_styles_map(path))   # 与动作同条件
        for para in root.iter(f"{W}p"):
            if cover_paras and any(para is c for c in cover_paras):
                continue
            text = "".join(x.text or "" for x in para.iter(f"{W}t")).strip()
            if not text:
                continue
            if re.match(r"^图\s*\d", text) or re.match(r"^表\s*\d", text):
                if not _is_caption_para(text):
                    continue  # 正文引用段（图6-1为…）不属标题
                number_bold, latin_bold = _caption_bold_flags(
                    cap, "table" if text.startswith("表") else "figure")
                if number_bold:
                    if not _caption_numbers_bold(para):
                        return False
                else:
                    for run in para.findall(f"{W}r"):
                        if any(run.find(f".//{W}{tag}") is not None
                               for tag in ("drawing", "object", "pict")):
                            continue
                        t = "".join(x.text or "" for x in run.findall(f"{W}t"))
                        rpr = run.find(f"{W}rPr")
                        if (re.search(r"[0-9]", t) and not re.search(r"[\u4e00-\u9fff]", t)
                                and rpr is not None and rpr.find(f"{W}b") is not None):
                            return False
                # 纯英文不得加粗（brands）——模板该类题注英文加粗（latin_bold=True）时
                # handler 不动英文 run，verifier 同样不查（2026-09-29：此前恒查 →
                # 表题英文加粗的模板上 verifier 永远 False，整条交付 FAILED）
                if latin_bold is True:
                    continue
                for run in para.findall(f"{W}r"):
                    if any(run.find(f".//{W}{tag}") is not None
                           for tag in ("drawing", "object", "pict")):
                        continue
                    t = "".join(x.text or "" for x in run.findall(f"{W}t"))
                    rpr = run.find(f"{W}rPr")
                    if (re.search(r"[A-Za-z]", t) and not re.search(r"[0-9]", t)
                            and rpr is not None and rpr.find(f"{W}b") is not None):
                        return False
        return True
    except Exception:
        return False

def set_page_border(path, params: dict) -> dict:
    """设置页面边框（§7.1 P0）。

    params: {section_index: int（-1=全部）, borders: {top/bottom/left/right:
             {val, sz, color, space}}}。CT_SectPr 顺序：pgBorders 在 pgMar
    之后、lnNumType 之前。R2。
    """
    path = Path(path)
    borders = params.get("borders") or {}
    if not borders:
        raise ValueError("set_page_border requires borders")
    root = _read_part(path, "word/document.xml")
    sects = _sectpr_by_index(root, int(params.get("section_index", -1)))
    changed = 0
    for sp in sects:
        old = sp.find(f"{W}pgBorders")
        if old is not None:
            sp.remove(old)
        pgb = etree.Element(f"{W}pgBorders")
        for edge, spec in borders.items():
            if edge not in ("top", "bottom", "left", "right"):
                continue
            e = etree.SubElement(pgb, f"{W}{edge}")
            e.set(f"{W}val", spec.get("val", "single"))
            if spec.get("sz"):
                e.set(f"{W}sz", str(spec["sz"]))
            if spec.get("color"):
                e.set(f"{W}color", spec["color"])
            if spec.get("space"):
                e.set(f"{W}space", str(spec["space"]))
        # 顺序：pgMar 之后
        pgmar = sp.find(f"{W}pgMar")
        if pgmar is not None:
            pgmar.addnext(pgb)
        else:
            sp.insert(0, pgb)
        changed += 1
    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "sections": changed}


def verify_page_border(path, params: dict) -> bool:
    """复查：pgBorders 已写入指定节。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        borders = params.get("borders") or {}
        sects = _sectpr_by_index(root, int(params.get("section_index", -1)))
        for sp in sects:
            pgb = sp.find(f"{W}pgBorders")
            if pgb is None:
                return False
            for edge in borders:
                if pgb.find(f"{W}{edge}") is None:
                    return False
        return True
    except Exception:
        return False

def configure_header_footer_variants(path, params: dict) -> dict:
    """统一配置节页眉页脚三态（§7.1 P0）。

    params: {section_index: int（-1=全部）, header: {default/first/even: text 或 None},
             footer: {default/first/even: text 或 None}, page_number: bool,
             title_page: bool, even_odd: bool}
    编排 set_header_reference/set_footer_reference/set_title_page_flag/
    set_even_odd_headers_flag（复用三态基建）。R2。
    """
    from earlyz.actions.header_footer_actions import (
        set_header_reference, set_footer_reference,
        set_title_page_flag, set_even_odd_headers_flag,
    )
    path = Path(path)
    section_index = int(params.get("section_index", -1))
    results: dict = {}
    header_cfg = params.get("header") or {}
    footer_cfg = params.get("footer") or {}
    for vt, cfg in (("header", header_cfg), ("footer", footer_cfg)):
        for vtype, text in cfg.items():
            if vtype not in ("default", "first", "even"):
                continue
            fn = set_header_reference if vt == "header" else set_footer_reference
            r = fn(path, {"section_index": section_index,
                          "header_type" if vt == "header" else "footer_type": vtype,
                          "text": text or "",
                          "page_number": bool(params.get("page_number", False))})
            results[f"{vt}.{vtype}"] = r.get("changed", False)
    if params.get("title_page") is not None:
        results["title_page"] = set_title_page_flag(
            path, {"section_index": section_index,
                   "enabled": bool(params["title_page"])}).get("changed", False)
    if params.get("even_odd") is not None:
        results["even_odd"] = set_even_odd_headers_flag(
            path, {"enabled": bool(params["even_odd"])}).get("changed", False)
    changed = any(results.values()) or bool(header_cfg or footer_cfg)
    return {"changed": changed, "details": results}


def _find_variant_ref(sectpr, tag: str, ref_type: str):
    """按类型找 header/footer 引用（lxml find 不支持 w: 前缀 XPath 谓词，
    必须遍历+属性比较）。default=无 type 或 type=default。"""
    for r_ in sectpr.findall(f"{W}{tag}"):
        t = r_.get(f"{W}type")
        if ref_type == "default":
            if t is None or t == "default":
                return r_
        elif t == ref_type:
            return r_
    return None


def verify_header_footer_variants(path, params: dict) -> bool:
    """复查：目标节已绑定对应变体引用（结构验证）。"""
    try:
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        section_index = int(params.get("section_index", -1))
        header_cfg = params.get("header") or {}
        footer_cfg = params.get("footer") or {}
        sects = list(root.iter(f"{W}sectPr"))
        if not sects:
            return False
        targets = sects if section_index < 0 and params.get("all_sections") else [sects[-1] if section_index < 0 else sects[section_index]]
        for sp in targets:
            for vtype in header_cfg:
                if _find_variant_ref(sp, "headerReference", vtype) is None:
                    return False
            for vtype in footer_cfg:
                if _find_variant_ref(sp, "footerReference", vtype) is None:
                    return False
        return True
    except Exception:
        return False


def clone_header_footer_from_spec(path, params: dict) -> dict:
    """从规范 part 完整复制页眉/页脚（§7.1 P0 clone_header_footer_from_spec）。

    params: {section_index: int, kind: "header"|"footer", source_part: str
             （如 "word/header2.xml"）, reference_type: "default"|"first"|"even"}
    复制 part 内容（含图片/表格/域/边框布局）并绑定到目标节。R2。
    """
    from earlyz.actions.header_footer_actions import (
        _next_part_number, _ensure_part_registered, _bind_reference,
        _header_content_type, _footer_content_type, _part_content,
    )
    path = Path(path)
    kind = params["kind"]
    if kind not in ("header", "footer"):
        raise ValueError("kind must be header or footer")
    source_part = params["source_part"]
    section_index = int(params.get("section_index", -1))
    ref_type = params.get("reference_type", "default")
    with zipfile.ZipFile(path) as z:
        if source_part not in z.namelist():
            raise ValueError(f"source part not found: {source_part}")
        source_xml = z.read(source_part)
    if kind == "header":
        new_part = _create_cloned_part(path, source_xml, "header")
        _ensure_part_registered(path, new_part, _header_content_type())
    else:
        new_part = _create_cloned_part(path, source_xml, "footer")
        _ensure_part_registered(path, new_part, _footer_content_type())
    # 注册后按 Target 反查 rel_id（rels 根命名空间是 package）
    R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    with zipfile.ZipFile(path) as z:
        rels = etree.fromstring(z.read("word/_rels/document.xml.rels"))
    rel_id = None
    new_basename = new_part.split("/")[-1]
    for rel in rels:
        t = rel.get("Target") or ""
        if t == new_part or t.split("/")[-1] == new_basename:
            rel_id = rel.get("Id")
            break
    if not rel_id:
        raise ValueError(f"rel not registered for {new_part}")
    _bind_reference(path, section_index, f"{kind}Reference", ref_type, rel_id)
    return {"changed": True, "cloned_to": new_part, "source": source_part,
            "section_index": section_index, "reference_type": ref_type}


def _create_cloned_part(path: Path, source_xml: bytes, kind: str) -> str:
    """把 source_xml 注册为新 part（word/headerN.xml 或 footerN.xml），返回 part 名。

    2026-08-22 单栏文档实锤（Word「文件可能已经损坏」）：源文档页脚页码
    文本框带 wp:docPr id——克隆 part 原样复制时 id 与原件重复，Word 要求
    docPr id 包级唯一（print_layout._clone_part 同款规则）。克隆时重编
    本 part 全部 docPr id 为包级 max+1 递增，name 加后缀防重名。
    """
    from earlyz.actions.header_footer_actions import _next_part_number
    with zipfile.ZipFile(path) as z:
        n = _next_part_number(z, kind)
    part = f"word/{kind}{n}.xml"
    # ── docPr id 重编（克隆体不得与原件共用 shape id）──
    try:
        clone_root = etree.fromstring(source_xml)
        WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
        doc_prs = clone_root.findall(f".//{WP}docPr")
    except Exception:  # noqa: BLE001 — 解析失败则原样克隆
        clone_root, doc_prs = None, []
    if clone_root is not None and doc_prs:
        max_id = 0
        with zipfile.ZipFile(path) as z:
            for pn in z.namelist():
                if pn.startswith("word/") and pn.endswith(".xml") and "_rels" not in pn:
                    try:
                        proot = etree.fromstring(z.read(pn))
                    except Exception:  # noqa: BLE001
                        continue
                    for el in proot.iter(f"{WP}docPr"):
                        try:
                            max_id = max(max_id, int(el.get("id", "0")))
                        except ValueError:
                            pass
        for el in doc_prs:
            max_id += 1
            el.set("id", str(max_id))
            if el.get("name"):
                el.set("name", f"{el.get('name')}-{max_id}")
        source_xml = etree.tostring(
            clone_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    import shutil
    tmp = path.with_suffix(".clone.docx")
    shutil.copy2(path, tmp)
    # 直接写 entry：重建 zip
    new_data = source_xml
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))
        _write_part_zipped(zout, part, new_data)
    tmp.replace(path)
    return part


def verify_header_footer_cloned(path, params: dict) -> bool:
    """复查：目标节已绑定克隆 part 且内容与 source 一致。"""
    try:
        from earlyz.actions.header_footer_actions import (
            _bind_reference, _ensure_part_registered,
        )
        path = Path(path)
        kind = params["kind"]
        source_part = params["source_part"]
        section_index = int(params.get("section_index", -1))
        ref_type = params.get("reference_type", "default")
        with zipfile.ZipFile(path) as z:
            source_xml = z.read(source_part)
            root = etree.fromstring(z.read("word/document.xml"))
        R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
        sects = list(root.iter(f"{W}sectPr"))
        if not sects:
            return False
        sectpr = sects[-1] if section_index < 0 else sects[section_index]
        ref = _find_variant_ref(sectpr, f"{kind}Reference", ref_type)
        if ref is None:
            return False
        rid = ref.get(f"{R}id")
        if not rid:
            return False
        with zipfile.ZipFile(path) as z:
            rels = etree.fromstring(z.read("word/_rels/document.xml.rels"))
        target = None
        for rel in rels:
            if rel.get("Id") == rid:
                target = rel.get("Target")
                break
        if not target:
            return False
        part = "word/" + target if not target.startswith("word/") else target
        with zipfile.ZipFile(path) as z:
            cur = z.read(part)
        # XML 语义比较（序列化字节可能因空白/声明不同）
        return etree.tostring(etree.fromstring(cur)) == etree.tostring(etree.fromstring(source_xml))
    except Exception:
        return False

def set_paragraph_tabs(path, params: dict) -> dict:
    """设置段落 tabs（§7.1 P0）：位置/对齐/leader（目录点线等）。

    params: {selector, tabs: [{pos, val, leader}]}——清空原 tabs 后重写
    （对齐语义：模板 tabs 完全复刻）。R1。
    """
    path = Path(path)
    tabs = params.get("tabs") or []
    if not tabs:
        raise ValueError("set_paragraph_tabs requires tabs")
    for t in tabs:
        if t.get("val") and t["val"] not in _TAB_VALS:
            raise ValueError(f"unsupported tab val: {t['val']}")
    root, paras = _locate_for_edit(path, params.get("selector") or {})
    changed = 0
    for p in paras:
        ppr = p.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            p.insert(0, ppr)
        old_tabs = ppr.find(f"{W}tabs")
        if old_tabs is not None:
            ppr.remove(old_tabs)
        tabs_el = etree.SubElement(ppr, f"{W}tabs")
        for t in tabs:
            tab = etree.SubElement(tabs_el, f"{W}tab")
            tab.set(f"{W}pos", str(int(t["pos"])))
            if t.get("val"):
                tab.set(f"{W}val", t["val"])
            if t.get("leader"):
                tab.set(f"{W}leader", t["leader"])
        changed += 1
    _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "paragraphs": changed, "tabs": len(tabs)}


def verify_paragraph_tabs(path, params: dict) -> bool:
    """复查：段落 tabs 已按参数写入。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        tabs = params.get("tabs") or []
        paras = _locate_paragraphs(root, params.get("selector") or {},
                                   _read_styles_map(Path(path)))
        if not paras:
            return False
        for p in paras:
            ppr = p.find(f"{W}pPr")
            tabs_el = ppr.find(f"{W}tabs") if ppr is not None else None
            if tabs_el is None:
                return False
            got = [t for t in tabs_el.findall(f"{W}tab")]
            if len(got) != len(tabs):
                return False
            for t, g in zip(tabs, got):
                if g.get(f"{W}pos") != str(int(t["pos"])):
                    return False
        return True
    except Exception:
        return False


def set_paragraph_border_shading(path, params: dict) -> dict:
    """设置段落边框与底纹（§7.1 P0）。

    params: {selector, borders: {top/bottom/left/right: {val, sz, color}},
             shading: {fill}}。pBdr 在 CT_PPr 顺序中位于 shd 前。R1。
    """
    path = Path(path)
    borders = params.get("borders") or {}
    shading = params.get("shading") or {}
    if not borders and not shading:
        raise ValueError("set_paragraph_border_shading requires borders or shading")
    root, paras = _locate_for_edit(path, params.get("selector") or {})
    changed = 0
    for p in paras:
        ppr = p.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            p.insert(0, ppr)
        if borders:
            old = ppr.find(f"{W}pBdr")
            if old is not None:
                ppr.remove(old)
            pbdr = etree.SubElement(ppr, f"{W}pBdr")
            for edge, spec in borders.items():
                if edge not in ("top", "bottom", "left", "right"):
                    continue
                e = etree.SubElement(pbdr, f"{W}{edge}")
                e.set(f"{W}val", spec.get("val", "single"))
                if spec.get("sz"):
                    e.set(f"{W}sz", str(spec["sz"]))
                if spec.get("color"):
                    e.set(f"{W}color", spec["color"])
        if shading:
            old_sh = ppr.find(f"{W}shd")
            if old_sh is not None:
                ppr.remove(old_sh)
            shd = etree.SubElement(ppr, f"{W}shd")
            shd.set(f"{W}val", shading.get("val", "clear"))
            if shading.get("fill"):
                shd.set(f"{W}fill", shading["fill"])
        changed += 1
    _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "paragraphs": changed}


def verify_paragraph_border_shading(path, params: dict) -> bool:
    """复查：pBdr/shd 已写入。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        borders = params.get("borders") or {}
        shading = params.get("shading") or {}
        paras = _locate_paragraphs(root, params.get("selector") or {},
                                   _read_styles_map(Path(path)))
        if not paras:
            return False
        for p in paras:
            ppr = p.find(f"{W}pPr")
            if ppr is None:
                return False
            if borders:
                pbdr = ppr.find(f"{W}pBdr")
                if pbdr is None:
                    return False
                for edge in borders:
                    if pbdr.find(f"{W}{edge}") is None:
                        return False
            if shading:
                shd = ppr.find(f"{W}shd")
                if shd is None or shd.get(f"{W}fill") != shading.get("fill"):
                    return False
        return True
    except Exception:
        return False


def set_run_format(path, params: dict) -> dict:
    """对稳定文本范围设置 run 级格式（§7.1 P0 set_run_format）。

    params: {selector, properties: {bold, italic, underline, font_ascii,
             font_east_asia, font_size_pt, color}}——对定位段落内全部 run
    应用（复用 _set_run_properties）。R1。
    """
    path = Path(path)
    props = params.get("properties") or {}
    if not props:
        raise ValueError("set_run_format requires properties")
    root, paras = _locate_for_edit(path, params.get("selector") or {})
    changed = 0
    for p in paras:
        _set_run_properties(p, props)
        changed += 1
    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "runs": changed}


def verify_run_format(path, params: dict) -> bool:
    """复查：定位段首个 run 的 rPr 已含目标属性（与 analyzer 读取点一致）。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        props = params.get("properties") or {}
        paras = _locate_paragraphs(root, params.get("selector") or {},
                                   _read_styles_map(Path(path)))
        if not paras:
            return False
        for p in paras:
            run = p.find(f"{W}r")
            if run is None:
                return False
            rpr = run.find(f"{W}rPr")
            if rpr is None:
                return False
            for k, v in props.items():
                if k == "bold" and (rpr.find(f"{W}b") is None) != (not v):
                    return False
                elif k == "font_size_pt":
                    sz = rpr.find(f"{W}sz")
                    if sz is None or sz.get(f"{W}val") != str(int(v * 2)):
                        return False
                elif k in ("font_ascii", "font_east_asia"):
                    rf = rpr.find(f"{W}rFonts")
                    attr = "ascii" if k == "font_ascii" else "eastAsia"
                    if rf is None or rf.get(f"{W}{attr}") != v:
                        return False
        return True
    except Exception:
        return False

def configure_page_numbering(path, params: dict) -> dict:
    """配置节页码（§7.1 P0）：pgNumType（fmt/start）+ 页脚/页眉 PAGE 域。

    params: {section_index: int（-1=全部）, format: str, start: int,
             position: "footer"|"header", alignment: "center"|"left"|"right"}
    R2：结构性修改（页码体系），不改正文内容。
    """
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    fmt = params.get("format")
    if fmt is not None and fmt not in _PGNUM_FORMATS:
        raise ValueError(f"unsupported page number format: {fmt}")
    start = params.get("start")
    position = params.get("position", "footer")
    alignment = params.get("alignment", "center")
    if position not in ("footer", "header"):
        raise ValueError("position must be footer or header")
    sects = _sectpr_by_index(root, int(params.get("section_index", -1)))

    changed = 0
    for sp in sects:
        # ① pgNumType（fmt/start）——CT_SectPr 顺序：pgMar 之后
        pnt = sp.find(f"{W}pgNumType")
        if pnt is None:
            pnt = etree.Element(f"{W}pgNumType")
            pgmar = sp.find(f"{W}pgMar")
            if pgmar is not None:
                pgmar.addnext(pnt)
            else:
                sp.insert(0, pnt)
        if fmt is not None:
            pnt.set(f"{W}fmt", fmt)
        if start is not None:
            pnt.set(f"{W}start", str(int(start)))
        changed += 1
        # ② 页脚/页眉 PAGE 域
        _write_page_field(path, sp, position, alignment)
    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "sections": changed}


def verify_page_numbering(path, params: dict) -> bool:
    """复查：pgNumType fmt/start 已写入指定节。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        fmt = params.get("format")
        start = params.get("start")
        sects = _sectpr_by_index(root, int(params.get("section_index", -1)))
        for sp in sects:
            pnt = sp.find(f"{W}pgNumType")
            if pnt is None:
                return False
            if fmt is not None and pnt.get(f"{W}fmt") != fmt:
                return False
            if start is not None and pnt.get(f"{W}start") != str(int(start)):
                return False
        return True
    except Exception:
        return False


def _write_page_field(path: Path, sectpr, position: str, alignment: str) -> bool:
    """在节的页脚/页眉 part 内写入 PAGE 域段落（保持既有段落格式模板）。

    返回是否写入。页脚/页眉 part 由 sectPr 的 footerReference/headerReference
    经 document.xml.rels Target 定位。
    """
    R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    ref_tag = f"{W}footerReference" if position == "footer" else f"{W}headerReference"
    ref = sectpr.find(ref_tag)
    if ref is None:
        return False
    rid = ref.get(f"{R}id")
    if not rid:
        return False
    # rels → Target
    with zipfile.ZipFile(path) as z:
        if "word/_rels/document.xml.rels" not in z.namelist():
            return False
        rels = etree.fromstring(z.read("word/_rels/document.xml.rels"))
    target = None
    for rel in rels:
        if rel.get("Id") == rid:
            target = rel.get("Target")
            break
    if not target:
        return False
    part = "word/" + target if not target.startswith("word/") else target
    with zipfile.ZipFile(path) as z:
        if part not in z.namelist():
            return False
        part_root = etree.fromstring(z.read(part))
    # 找含 PAGE 域的段；无则新建
    page_para = None
    for p in part_root.iter(f"{W}p"):
        if p.find(f".//{W}instrText") is not None:
            page_para = p
            break
    if page_para is None:
        page_para = etree.Element(f"{W}p")
        ppr = etree.SubElement(page_para, f"{W}pPr")
        jc = etree.SubElement(ppr, f"{W}jc")
        jc.set(f"{W}val", alignment)
        r1 = etree.SubElement(page_para, f"{W}r")
        fc = etree.SubElement(r1, f"{W}fldChar")
        fc.set(f"{W}fldCharType", "begin")
        r2 = etree.SubElement(page_para, f"{W}r")
        it = etree.SubElement(r2, f"{W}instrText")
        it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        it.text = " PAGE "
        r3 = etree.SubElement(page_para, f"{W}r")
        fe = etree.SubElement(r3, f"{W}fldChar")
        fe.set(f"{W}fldCharType", "end")
        part_root.append(page_para)
    else:
        # 已有 PAGE 域段：只更新 jc
        ppr = page_para.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            page_para.insert(0, ppr)
        jc = ppr.find(f"{W}jc")
        if jc is None:
            jc = etree.SubElement(ppr, f"{W}jc")
        jc.set(f"{W}val", alignment)
    _write_part(path, part, part_root)
    return True


def restart_list_numbering(path, params: dict) -> dict:
    """重启/指定多级编号起始值（§7.1 P0）。

    params: {num_id: int, ilvl: int, start: int}（start=None 时移除 lvlOverride
    恢复连续）；对 numbering.xml 的 num 元素写 lvlOverride/startOverride。
    R2：编号结构修改，不改文本。
    """
    path = Path(path)
    num_id = int(params["num_id"])
    ilvl = int(params.get("ilvl", 0))
    with zipfile.ZipFile(path) as z:
        if "word/numbering.xml" not in z.namelist():
            raise ValueError("numbering.xml missing")
        nroot = etree.fromstring(z.read("word/numbering.xml"))
    target = None
    for num in nroot.iter(f"{W}num"):
        if num.get(f"{W}numId") == str(num_id):
            target = num
            break
    if target is None:
        raise ValueError(f"numId {num_id} not found in numbering.xml")
    start = params.get("start")
    lvl_ov = None
    for lo in target.findall(f"{W}lvlOverride"):
        if lo.get(f"{W}ilvl") == str(ilvl):
            lvl_ov = lo
            break
    if start is None:
        # 恢复连续：移除该 ilvl 的 lvlOverride
        if lvl_ov is not None:
            target.remove(lvl_ov)
            changed = True
        else:
            changed = False
    else:
        if lvl_ov is None:
            lvl_ov = etree.SubElement(target, f"{W}lvlOverride")
            lvl_ov.set(f"{W}ilvl", str(ilvl))
            so = etree.SubElement(lvl_ov, f"{W}startOverride")
            so.set(f"{W}val", str(int(start)))
            changed = True
        else:
            so = lvl_ov.find(f"{W}startOverride")
            if so is None:
                so = etree.SubElement(lvl_ov, f"{W}startOverride")
            old = so.get(f"{W}val")
            so.set(f"{W}val", str(int(start)))
            changed = old != str(int(start))
    if changed:
        _write_part(path, "word/numbering.xml", nroot)
    return {"changed": changed, "num_id": num_id, "ilvl": ilvl, "start": start}


def verify_list_numbering(path, params: dict) -> bool:
    """复查：num 的 lvlOverride/startOverride 已写入。"""
    try:
        path = Path(path)
        with zipfile.ZipFile(path) as z:
            nroot = etree.fromstring(z.read("word/numbering.xml"))
        num_id = str(params["num_id"])
        ilvl = str(params.get("ilvl", 0))
        start = params.get("start")
        for num in nroot.iter(f"{W}num"):
            if num.get(f"{W}numId") == num_id:
                for lo in num.findall(f"{W}lvlOverride"):
                    if lo.get(f"{W}ilvl") == ilvl:
                        if start is None:
                            return False  # 期望连续但仍有 override
                        so = lo.find(f"{W}startOverride")
                        return so is not None and so.get(f"{W}val") == str(int(start))
                return start is None
        return False
    except Exception:
        return False

def normalize_ooxml_order(path, params: dict) -> dict:
    """按 Word 严格 schema 规范 document.xml/styles.xml 内 CT_* 子元素顺序
    （§7.1 P0）。R2：结构性修复，不改文本。"""
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    fixed = 0
    for tag, order in _CT_ORDER.items():
        for el in root.iter(tag):
            fixed += _normalize_child_order(el, order)
    if fixed:
        _write_part(path, "word/document.xml", root)
    return {"changed": fixed > 0, "elements_fixed": fixed}


def verify_ooxml_order(path, params: dict) -> bool:
    """复查：所有 CT_* 元素子元素顺序合法。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        for tag, order in _CT_ORDER.items():
            idx_map = {f"{W}{n}": i for i, n in enumerate(order)}
            for el in root.iter(tag):
                last = -1
                for i, ch in enumerate(el):
                    k = idx_map.get(ch.tag)
                    if k is None:
                        continue
                    if k < last:
                        return False
                    last = k
        return True
    except Exception:
        return False


def repair_package_invariants(path, params: dict) -> dict:
    """修复包级不变量（§7.1 P0）：①wp:docPr id 包级唯一（重复 id 重写为
    未占用值）；②[Content_Types].xml 缺 Override 补齐。R2 结构性修复。"""
    path = Path(path)
    docroot = _read_part(path, "word/document.xml")
    WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    # ① docPr id 唯一
    ids = []
    for dp in docroot.iter(f"{WP}docPr"):
        v = dp.get("id")
        if v is not None:
            ids.append(v)
    dup_fixed = 0
    used = set(ids)
    nxt = 1
    for dp in docroot.iter(f"{WP}docPr"):
        v = dp.get("id")
        if v is None:
            continue
        if ids.count(v) > 1:
            while str(nxt) in used:
                nxt += 1
            dp.set("id", str(nxt))
            used.add(str(nxt))
            ids[ids.index(v)] = str(nxt)  # 已修一次
            dup_fixed += 1
    changed = dup_fixed > 0
    if changed:
        _write_part(path, "word/document.xml", docroot)
    # ② ContentTypes Override 补齐
    ct_added = _ensure_content_type_overrides(path)
    changed = changed or ct_added > 0
    return {"changed": changed, "docpr_fixed": dup_fixed,
            "content_type_added": ct_added}


def verify_package_invariants(path, params: dict) -> bool:
    """复查：docPr id 包级唯一 + 必需 Override 存在。"""
    try:
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
        seen = set()
        for dp in root.iter(f"{WP}docPr"):
            v = dp.get("id")
            if v is None:
                continue
            if v in seen:
                return False
            seen.add(v)
        return True
    except Exception:
        return False


def _ensure_content_type_overrides(path: Path) -> int:
    """[Content_Types].xml 缺 word/document.xml / styles.xml 等 Override 时补齐。"""
    ct_part = "[Content_Types].xml"
    added = 0
    with zipfile.ZipFile(path) as z:
        if ct_part not in z.namelist():
            return 0
        ct = etree.fromstring(z.read(ct_part))
    CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
    root_el = ct
    needed = {
        "/word/document.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        "/word/styles.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        "/word/numbering.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        "/word/settings.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
    }
    existing = {o.get("PartName") for o in root_el.findall(f"{CT}Override")}
    for partname, ctype in needed.items():
        if partname not in existing:
            ov = etree.SubElement(root_el, f"{CT}Override")
            ov.set("PartName", partname)
            ov.set("ContentType", ctype)
            added += 1
    if added:
        _write_part(path, ct_part, ct)
    return added

def _residual_norm(t: str) -> str:
    """示例残留规范化（与 decompiler._norm_residual 同逻辑）。"""
    t = re.sub(r"[\s\u3000□]+", "", t or "")
    t = t.replace("：", ":").replace("；", ";").replace("（", "(").replace("）", ")")
    t = t.replace("，", ",").replace("。", ".").replace("、", ",")
    return t.strip()


def _latin_concat(t: str) -> str:
    """拉丁字母连续拼接（小写、去空格）——用于识别示例残留的英文变体行。"""
    return "".join(re.findall(r"[A-Za-z]+", t or "")).lower()


def _residual_latin_set(residuals) -> set[str]:
    """提取含足量拉丁词的残留的拉丁拼接串（≥8 字符防误伤短词如 GIS）。"""
    out: set[str] = set()
    for r in residuals:
        if not re.search(r"[A-Za-z]{4,}", r):
            continue
        lc = _latin_concat(r)
        if len(lc) >= 8:
            out.add(lc)
    return out


def _is_residual_text(text: str, residuals: set, lat_set: set[str]) -> bool:
    """段落是否属示例残留：精确规范化匹配，或拉丁拼接包含（英文翻译行
    内嵌示例的英文词序列，如 Application of Altera DDR IPCore… 内嵌
    alteraddripcore——模板示例只收中文段时英文变体的通用识别法）。"""
    if _residual_norm(text) in residuals:
        return True
    if not lat_set:
        return False
    lc = _latin_concat(text)
    if len(lc) < 8:
        return False
    return any(r in lc for r in lat_set)


def remove_example_residuals(target_path: Path, params: dict) -> dict:
    """删除与模板示例残留（规范化）匹配的段落（2026-08-16 用户拍板：
    模板只用于提取格式，示例内容照搬进产物必须清除）。

    安全护栏：跳过 hyperlink 段（TOC 条目）、表格内段、标题/锚点段、
    封面中文日期段（与用户真实日期同形，永不删）。
    """
    residuals = set(params.get("residuals") or [])
    if not residuals:
        return {"changed": False, "removed": []}
    # 作者内容保护（2026-11-07）：段落文字在**源文档**里也存在的，是作者自己的内容，
    # 不是引擎从模板抄进来的示例——一律不删。软件编码说明书有 13 段代码样例
    # （`<html lang="zh-CN">`、`[强制] …`）与模板示例同形，被当残留删掉 → content 判丢字。
    # 规范化口径必须与 _is_residual_text 内部的 _residual_norm 一致，否则永远命中不了
    keep_texts = {_residual_norm(x) for x in (params.get("keep_texts") or []) if x}
    lat_set = _residual_latin_set(residuals)
    root = _read_part(target_path, "word/document.xml")
    body = root.find(f"{W}body")
    removed: list[str] = []
    for child in list(body):
        if child.tag != f"{W}p":
            continue
        if child.find(f"{W}hyperlink") is not None:
            continue  # TOC 条目（目录）
        # 公式段没有 w:t 文字（内容在 m:r 里）→ 文本为空 → 会被误判为残留删除。
        # 2026-10-20 实测：含 OMML 的文档排版后公式整段消失（源 1 → 产物 0）。
        if child.find(f".//{M}oMath") is not None or child.find(f".//{M}oMathPara") is not None:
            continue
        if any(anc.tag == f"{W}tbl" for anc in child.iterancestors(f"{W}tbl")):
            continue  # 表格内段
        text = "".join(t.text or "" for t in child.iter(f"{W}t"))
        if ("签名" in text or text.strip().startswith(("日期", "论文提交日期"))):
            continue  # 字段标签护栏：签名/日期行是目标文档真实结构，不视为残留
        # 封面日期护栏（2026-08-17 投诉⑥：「二○二四年五月」被误删——
        # 模板示例日期与用户真实日期同形，中文年月短段永不删除）
        if _COVER_DATE_RE.match(text.strip()):
            continue
        # 源文档里也有这段文字 = 作者内容，不是模板示例残留 → 不删
        if keep_texts and _residual_norm(text) in keep_texts:
            continue
        if _is_residual_text(text, residuals, lat_set):
            body.remove(child)
            removed.append(text.strip()[:48])
    _write_part(target_path, "word/document.xml", root)
    return {"changed": bool(removed), "removed": removed}


def verify_remove_residuals(target_path: Path, params: dict) -> bool:
    """验证：residuals 中仍有段残留 → False。"""
    residuals = set(params.get("residuals") or [])
    if not residuals:
        return True
    lat_set = _residual_latin_set(residuals)
    root = _read_part(target_path, "word/document.xml")
    body = root.find(f"{W}body")
    for child in body.iter(f"{W}p"):
        text = "".join(t.text or "" for t in child.iter(f"{W}t"))
        if _is_residual_text(text, residuals, lat_set):
            return False
    return True


def set_table_alignment(target_path: Path, params: dict) -> dict:
    """T2b 表格对齐规则执行端：所有表格 tblPr/jc 统一（批量幂等）。

    params: {properties: {jc: "center"}}
    （值来自模板反编译规则；无 tblPr 的表格补建——注意 CT_TblPr
    子元素顺序：tblStyle 为首，jc 插入在其后。）
    """
    props = params.get("properties") or {}
    if not props:
        raise ValueError("set_table_alignment requires properties")
    root = _read_part(target_path, "word/document.xml")
    tables = list(root.iter(f"{W}tbl"))
    if not tables:
        raise ValueError("no tables found")
    changed = 0
    for tbl in tables:
        tblpr = tbl.find(f"{W}tblPr")
        if tblpr is None:
            tblpr = etree.SubElement(tbl, f"{W}tblPr")
            tbl.remove(tblpr)
            tbl.insert(0, tblpr)
        for k, v in props.items():
            if k == "tblpr_extra_xml":
                continue
            el = tblpr.find(f"{W}{k}")
            if el is None:
                # jc 必须排在 tblStyle 之后（OOXML CT_TblPr 序列）
                anchor = tblpr.find(f"{W}tblStyle")
                if anchor is not None:
                    el = etree.Element(f"{W}{k}")
                    anchor.addnext(el)
                else:
                    el = etree.SubElement(tblpr, f"{W}{k}")
            el.set(f"{W}val", str(v))
        if "tblpr_extra_xml" in props:
            _apply_structure_xml(tblpr, props.get("tblpr_extra_xml"), "tblPr")
        _normalize_child_order(tblpr, _CT_TBLPR_ORDER)
        changed += 1
    _write_part(target_path, "word/document.xml", root)
    return {"ok": True, "tables": changed, "properties": props}


def verify_table_alignment(target_path: Path, params: dict) -> bool:
    """T2b 校验：所有表格 tblPr/jc 与期望一致。"""
    props = params.get("properties") or {}
    root = _read_part(target_path, "word/document.xml")
    tables = list(root.iter(f"{W}tbl"))
    for tbl in tables:
        tblpr = tbl.find(f"{W}tblPr")
        for k, v in props.items():
            if k == "tblpr_extra_xml":
                if not structure_bundle_contains(_structure_xml_from_lxml(tblpr, "tblPr"), v):
                    return False
                continue
            el = None
            if tblpr is not None:
                el = tblpr.find(f"{W}{k}")
            if el is None or el.get(f"{W}val") != str(v):
                return False
    return True

def verify_page_margins_all(target_path: Path, params: dict) -> bool:
    """验证全部 sectPr 的页边距已按期望设置。"""
    props = params.get("properties") or {}
    try:
        root = _read_part(target_path, "word/document.xml")
        sects = list(root.iter(f"{W}sectPr"))
        if not sects:
            return False
        tol = 3  # 与 decompiler.margins_diff 同容差
        for sp in sects:
            expected_extra = props.get("section_extra_xml")
            if expected_extra is not None and not structure_bundle_contains(
                _structure_xml_from_lxml(sp, "sectPr"), expected_extra
            ):
                return False
            margin_props = {k: v for k, v in props.items() if k != "section_extra_xml"}
            pm = sp.find(f"{W}pgMar")
            if margin_props and pm is None:
                return False
            for k, v in margin_props.items():
                actual = pm.get(f"{W}{k}")
                if actual is None:
                    return False
                try:
                    if abs(int(actual) - int(v)) > tol:
                        return False
                except ValueError:
                    return False
        return True
    except Exception:
        return False


def translate_set_style_params(params: dict) -> dict:
    """{style_id, properties:{...}} → 旧 set_style 扁平参数。"""
    out: dict = {"style_id": params["style_id"]}
    for prop, value in (params.get("properties") or {}).items():
        old_key = _OLD_STYLE_MAP.get(prop)
        if old_key is not None:
            out[old_key] = value
    return out


def normalize_plan_actions(plan_actions: list[dict]) -> list[dict]:
    """执行前参数归一化（新链路专用，不影响旧 apply_operations）。"""
    out = []
    for step in plan_actions:
        s = dict(step)
        if s.get("action") == "set_style" and "properties" in (s.get("params") or {}):
            s["params"] = translate_set_style_params(s["params"])
        out.append(s)
    return out


# ═══════════════════════════════════════════════════════════════
# 2026-08-17 用户投诉六大根因修复（18:41 真机产物取证）：
# ①目录条目无 pStyle/缩进/点线 → format_toc_entries
# ③表格不占满版心 → normalize_table_widths
# ⑤图标题被错打表标题黑体规则 → format_caption_fonts（图/表分流）
# ⑥封面「二○二四年五月」被误删 → remove_example_residuals 日期护栏
# ═══════════════════════════════════════════════════════════════

# 封面日期护栏（模板示例「二○二四年五月」与用户真实日期同形——
# 中文年月结尾的短段是封面结构元素，永不作为示例残留删除）
_COVER_DATE_RE = re.compile(
    r"^[\s　]*[一二三四五六七八九○〇零\d]{4}[\s　]*年"
    r"(?:[\s　]*[一二三四五六七八九十○〇零\d]{1,3}[\s　]*月)?[\s　]*$"
)


def _toc_level_of(text: str) -> int:
    """TOC 条目层级：按条目前缀编号判定（1→L1 / 1.1→L2 / 1.1.1→L3），
    无数字前缀（摘要/ABSTRACT/致谢）→ L1。"""
    t = (text or "").strip()
    m = re.match(r"^(\d+)(?:[.．](\d+))?(?:[.．](\d+))?\s*[\u4e00-\u9fffA-Za-z]", t)
    if not m:
        return 1
    if m.group(3):
        return 3
    if m.group(2):
        return 2
    return 1


def format_toc_entries(path, params: dict) -> dict:
    """目录条目段落格式化（投诉①：目录一塌糊涂）。

    源文档 TOC 域条目段常无 pStyle/无缩进/无点线制表位，页码黏在文字后。
    对每个 hyperlink 条目段：
    - 补 pStyle=TOC{level}（挂上模板 TOC 样式：L1 黑体、L2/3 宋体缩进）；
    - 段级补 right+dot 制表位（版心右缘，样式缺 tabs 时条目自身兜底）；
    - 清正文规则误加的首行缩进/固定行距（目录条目不该继承正文 23 磅行距）。
    幂等：已有正确 pStyle 跳过。R1 纯格式。
    """
    path = Path(path)
    toc_spec = (params or {}).get("toc") or {}
    root = _read_part(path, "word/document.xml")
    # 版心右缘制表位（twips）：pgSz - 左右边距
    body = root.find(f"{W}body")
    sect = body.find(f"{W}sectPr") if body is not None else None
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

    styled = 0
    tabs_added = 0
    ind_cleared = 0
    cover_paras = _cover_paragraphs(root, _read_styles_map(path))   # 封面条目不动
    for p in root.iter(f"{W}p"):
        if cover_paras and any(p is c for c in cover_paras):
            continue
        hl = p.find(f"{W}hyperlink")
        if hl is None:
            continue
        # 条目文本（hyperlink 内 + 段内直接 run，不含页码域指令）
        text = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        if not text:
            continue
        level = _toc_level_of(text)
        ppr = p.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            p.insert(0, ppr)
        # 1) pStyle=TOC{n}
        ps = ppr.find(f"{W}pStyle")
        if ps is None:
            ps = etree.Element(f"{W}pStyle")
            ps.set(f"{W}val", f"TOC{level}")
            ppr.insert(0, ps)
            styled += 1
        # 1b) 目录条目字体（2026-08-18 用户彻查「各章黑体4号加粗，其余宋体小4」）：
        #     模板注释是目录条目的真值——章条目（L1）=黑体四号加粗，其余=宋体小四。
        #     以直接格式写入 run（样式继承源文档 TOC1=黑体28 不满足规范）。
        #     幂等：已有规范字体跳过（避免每次全刷）。
        need_font = False
        spec_lvl = {}
        # TOC 条目 run 在 hyperlink 内（p.findall 直接子级找不到）
        first_run = None
        hl = p.find(f"{W}hyperlink")
        if hl is not None:
            first_run = next((r for r in hl.findall(f"{W}r")
                              if "".join(t.text or "" for t in r.iter(f"{W}t")).strip()), None)
        if first_run is None:
            first_run = next((r for r in p.findall(f"{W}r")
                              if "".join(t.text or "" for t in r.iter(f"{W}t")).strip()), None)
        if first_run is not None:
            rpr0 = first_run.find(f"{W}rPr")
            rf0 = rpr0.find(f"{W}rFonts") if rpr0 is not None else None
            ea0 = rf0.get(f"{W}eastAsia") if rf0 is not None else None
            b0 = rpr0 is not None and rpr0.find(f"{W}b") is not None
            sz0 = rpr0.find(f"{W}sz") if rpr0 is not None else None
            sz0v = sz0.get(f"{W}val") if sz0 is not None else None
            spec_lvl = {}
            if level == 1:
                spec_lvl = toc_spec.get("1") or toc_spec.get(1) or {}
            else:
                spec_lvl = toc_spec.get(str(level)) or toc_spec.get(level) \
                    or toc_spec.get("2") or toc_spec.get(2) or {}
            want_ea = spec_lvl.get("font_east_asia")
            want_pt = spec_lvl.get("font_size_pt")
            want_sz = str(int(round(float(want_pt) * 2))) if want_pt else None
            want_bold = spec_lvl.get("bold")
            if not spec_lvl:
                need_font = False
            elif level == 1:
                if (want_ea and ea0 != want_ea) or (want_bold and not b0) \
                        or (want_sz and sz0v != want_sz):
                    need_font = True
            else:
                if (want_ea and ea0 not in (None, want_ea)) \
                        or (want_sz and sz0v not in (None, want_sz)) \
                        or (want_bold is False and b0):
                    need_font = True
        if need_font and spec_lvl:
            target_runs = []
            if hl is not None:
                target_runs += [r for r in hl.findall(f"{W}r")
                                if "".join(t.text or "" for t in r.iter(f"{W}t")).strip()]
            target_runs += [r for r in p.findall(f"{W}r")
                            if "".join(t.text or "" for t in r.iter(f"{W}t")).strip()]
            for r in target_runs:
                rt = "".join(t.text or "" for t in r.iter(f"{W}t"))
                if not rt.strip():
                    continue
                rpr = r.find(f"{W}rPr")
                if rpr is None:
                    rpr = etree.SubElement(r, f"{W}rPr")
                rf = rpr.find(f"{W}rFonts")
                if rf is None:
                    rf = etree.SubElement(rpr, f"{W}rFonts")
                if want_ea:
                    rf.set(f"{W}eastAsia", want_ea)
                ascii_f = spec_lvl.get("font_ascii") or "Times New Roman"
                rf.set(f"{W}ascii", ascii_f)
                if want_bold is True and rpr.find(f"{W}b") is None:
                    etree.SubElement(rpr, f"{W}b")
                if want_bold is False and rpr.find(f"{W}b") is not None:
                    rpr.remove(rpr.find(f"{W}b"))
                if want_sz:
                    sz = rpr.find(f"{W}sz")
                    if sz is None:
                        sz = etree.SubElement(rpr, f"{W}sz")
                    sz.set(f"{W}val", want_sz)
            styled += 1
        # 2) right+dot 制表位兜底（样式有则不重复定义也无害，但同值幂等）
        tabs = ppr.find(f"{W}tabs")
        if tabs is None or not any(
            (t.get(f"{W}val") == "right") for t in tabs.findall(f"{W}tab")
        ):
            if tabs is None:
                tabs = etree.Element(f"{W}tabs")
                # CT_PPr 序列：tabs 在 spacing/ind 之前（pStyle/keepNext 之后）
                _ppr_insert_ordered(ppr, tabs)
            tab = etree.SubElement(tabs, f"{W}tab")
            tab.set(f"{W}val", "right")
            tab.set(f"{W}leader", "dot")
            tab.set(f"{W}pos", str(tab_pos))
            tabs_added += 1
        # 3) 清正文误加的首行缩进（TOC 样式自带层级缩进，直接缩进叠加会错位）
        ind = ppr.find(f"{W}ind")
        if ind is not None and ind.get(f"{W}firstLine") is not None:
            etree.strip_attributes(ind, f"{W}firstLine", f"{W}firstLineChars")
            if not ind.attrib:
                ppr.remove(ind)
            ind_cleared += 1
        # 4) 去超链主题色（机制：af6/Hyperlink 会盖掉模板字体）。不写死 000000。
        for r in p.iter(f"{W}r"):
            rpr = r.find(f"{W}rPr")
            if rpr is None:
                continue
            rs = rpr.find(f"{W}rStyle")
            if rs is not None and (rs.get(f"{W}val") or "") in ("af6", "Hyperlink", "FollowedHyperlink"):
                rpr.remove(rs)
                styled += 1
            u = rpr.find(f"{W}u")
            if u is not None:
                rpr.remove(u)
                styled += 1

    _write_part(path, "word/document.xml", root)
    return {"changed": styled + tabs_added + ind_cleared > 0,
            "entries_styled": styled, "tabs_added": tabs_added,
            "indent_cleared": ind_cleared}


def ensure_toc_tab_stops(path, params: dict) -> dict:
    """TOC1/2/3 样式补右对齐点线制表位（投诉①配套：样式层兜底）。

    学校模板 styles 常只有 TOC1 带 tabs；TOC2/TOC3 缺失 → 二三级条目
    页码直接黏在文字后。同值幂等，已有 right 制表位不覆盖。
    """
    path = Path(path)
    root = _read_part(path, "word/styles.xml")
    doc_root = _read_part(path, "word/document.xml")
    body = doc_root.find(f"{W}body")
    sect = body.find(f"{W}sectPr") if body is not None else None
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

    changed = 0
    for style in root.iter(f"{W}style"):
        sid = style.get(f"{W}styleId", "")
        if sid not in ("TOC1", "TOC2", "TOC3"):
            continue
        ppr = style.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.SubElement(style, f"{W}pPr")
        tabs = ppr.find(f"{W}tabs")
        if tabs is not None and any(
            t.get(f"{W}val") == "right" for t in tabs.findall(f"{W}tab")
        ):
            continue
        if tabs is None:
            tabs = etree.Element(f"{W}tabs")
            ppr.insert(0, tabs)
        tab = etree.SubElement(tabs, f"{W}tab")
        tab.set(f"{W}val", "right")
        tab.set(f"{W}leader", "dot")
        tab.set(f"{W}pos", str(tab_pos))
        changed += 1
    # TOC 样式取值来自模板注释/结构抽取（缺则不改字体字号）
    # 2026-09-25：逐级应用（此前只取 "1" → 二级/三级目录条目拿不到模板
    # toc2/toc3 取值，仍是源稿目录样式；报告/论文目录普遍有 2-3 级）。
    _toc_params = (params or {}).get("toc") or {}
    toc1 = None
    for _lvl, _sid in (("1", "TOC1"), ("2", "TOC2"), ("3", "TOC3")):
        _lv = _toc_params.get(_lvl) or {}
        _want_ea = _lv.get("font_east_asia")
        _want_pt = _lv.get("font_size_pt")
        _want_sz = str(int(round(float(_want_pt) * 2))) if _want_pt else None
        style = None
        for _st in root.iter(f"{W}style"):
            if _st.get(f"{W}styleId") == _sid:
                style = _st
                break
        if style is None:
            continue
        if _sid == "TOC1":
            toc1 = style
        if not (_want_ea or _want_sz):
            continue
        rpr = style.find(f"{W}rPr")
        if rpr is None:
            rpr = etree.SubElement(style, f"{W}rPr")
        rf = rpr.find(f"{W}rFonts")
        if rf is None:
            rf = etree.SubElement(rpr, f"{W}rFonts")
        if _want_ea and rf.get(f"{W}eastAsia") != _want_ea:
            rf.set(f"{W}eastAsia", _want_ea)
            changed += 1
        if _want_sz:
            sz = rpr.find(f"{W}sz")
            if sz is None:
                sz = etree.SubElement(rpr, f"{W}sz")
            if sz.get(f"{W}val") != _want_sz:
                sz.set(f"{W}val", _want_sz)
                changed += 1
    if toc1 is not None:
        ppr1 = toc1.find(f"{W}pPr")
        if ppr1 is not None:
            jc = ppr1.find(f"{W}jc")
            if jc is not None and jc.get(f"{W}val") == "distribute":
                jc.set(f"{W}val", "left")
                changed += 1
            ind = ppr1.find(f"{W}ind")
            if ind is not None and ind.get(f"{W}firstLine") is not None:
                etree.strip_attributes(ind, f"{W}firstLine", f"{W}firstLineChars")
                changed += 1
    # 2026-08-18 用户彻查（模板-4 铁证）：模板 TOC 示例条目无缩进
    # （idx=116-135 全部 spacing 360 无 ind），源文档 TOC2/3 样式却带
    # leftChars=200（左侧2字符）+ firstLineChars=200（首行2字符）——
    # earlyz 必须把样式缩进清掉，否则目录条目显示「左侧2字符+首行2字符」。
    # 层级区分用制表位（右对齐点线），不用 left 缩进。
    for style in root.iter(f"{W}style"):
        sid = style.get(f"{W}styleId", "")
        if sid not in ("TOC2", "TOC3"):
            continue
        ppr = style.find(f"{W}pPr")
        if ppr is None:
            continue
        ind = ppr.find(f"{W}ind")
        if ind is not None:
            for attr in (f"{W}left", f"{W}leftChars",
                         f"{W}firstLine", f"{W}firstLineChars",
                         f"{W}hanging", f"{W}hangingChars"):
                if ind.get(attr) is not None:
                    del ind.attrib[attr]
                    changed += 1
            if len(ind.attrib) == 0:
                ppr.remove(ind)
                changed += 1
    if changed:
        _write_part(path, "word/styles.xml", root)
    return {"changed": changed > 0, "styles_updated": changed}


def _toc_align(
    toc_norms: list[str], body_norms: list[str]
) -> list[tuple[int, int]]:
    """TOC 域缓存条目与正文标题的序列对齐（anchor + 块内回填）。

    TOC 是 Word 域生成的**派生缓存**：条目按文档顺序与正文标题一一
    对应（数量不必相等——缓存可能多于标题，如正文改名/删除后未刷新
    域；也可能少于标题，如正文新增标题未刷新域）。

    Pass1（锚点）：按顺序贪心做**精确文本匹配**（标题文本唯一性强），
    匹配上即锚定；任一锚点区间内的未匹配缓存条目与未匹配正文标题
    **只在本区间内按顺序 zip 回填**（位置对应语义，即 Word 更新域后
    的结果：条目文本 = 对应位置正文标题文本）——改名场景（「结论与
    展望」←「结论」）在此对齐；跨区间不配对（缓存孤儿条目保留原文
    本，正文新标题不入目录，均不改内容）。

    返回 [(toc_idx, body_idx)]，按 toc_idx 升序。"""
    anchors: list[tuple[int, int]] = []
    pairs: list[tuple[int, int]] = []
    pending: list[int] = []
    i = j = 0

    def flush(body_block: list[int]) -> None:
        for ti, bj in zip(pending, body_block):
            pairs.append((ti, bj))
        pending.clear()

    while i < len(toc_norms) and j < len(body_norms):
        if toc_norms[i] == body_norms[j]:
            flush([])
            anchors.append((i, j))
            i += 1
            j += 1
            continue
        # 当前正文标题没有目录条目 → 在后续正文标题中找 toc[i] 精确匹配
        k = j + 1
        found_at = -1
        while k < len(body_norms):
            if toc_norms[i] == body_norms[k]:
                found_at = k
                break
            k += 1
        if found_at >= 0:
            flush(list(range(j, found_at)))
            anchors.append((i, found_at))
            i += 1
            j = found_at + 1
            continue
        # 缓存条目在正文标题中无精确匹配 → 暂入未匹配池（可能块内回填改名）
        pending.append(i)
        i += 1
    flush(list(range(j, len(body_norms))))
    pending.extend(range(i, len(toc_norms)))  # 尾部缓存的孤儿：保留，不配对
    return sorted(anchors + pairs)


def _collect_toc_ctx(body) -> tuple:
    """收集 TOC 对齐所需上下文（update/verify 同源）。

    返回 (body_titles, toc_paras, toc_norms, toc_title_ts, all_heading_norms):
    - body_titles：正文标题（outlineLvl≤1，排除目录/声明）——TOC 域 o"1-2"
    - toc_paras/toc_norms/toc_title_ts：目录缓存条目段/标题 norm/标题 run
    - all_heading_norms：**正文标题全集**（outlineLvl≤1 + back 标题角色
      refs_heading/ack_heading/appendix_heading）——目录条目若在合集内
      但未配对（如「参考文献」——refs_heading 有意不设 outlineLvl），
      说明正文有对应章节，保留；不在合集（如正文改名/删除后未刷新域
      的「结论与展望」）＝真孤儿，静默删除（模拟 Word F9 收敛结果）。
    """
    def _ptxt(p):
        return "".join(x.text or "" for x in p.iter(f"{W}t")).strip()

    from earlyz.ir.analyzer import _back_heading_role

    body_titles: list[str] = []
    all_heading_norms: list[str] = []
    seen_back = set()
    for p in body.iter(f"{W}p"):
        t = _ptxt(p)
        if not t:
            continue
        ppr = p.find(f"{W}pPr")
        oLvl = ppr.find(f"{W}outlineLvl") if ppr is not None else None
        norm = re.sub(r"\s+", "", t)
        if oLvl is not None:
            lvl = int(oLvl.get(f"{W}val", "0"))
            if norm in ("目录", "学位论文原创性声明", "学位论文版权使用授权书"):
                continue
            if lvl <= 1:
                body_titles.append(norm)
                all_heading_norms.append(norm)
        else:
            back = _back_heading_role(t)
            if back and back in ("refs_heading", "ack_heading", "appendix_heading"):
                if norm not in seen_back:
                    seen_back.add(norm)
                    all_heading_norms.append(norm)
    # 目录条目段
    toc_paras = []
    for p in body.iter(f"{W}p"):
        ppr = p.find(f"{W}pPr")
        pstyle = ppr.find(f"{W}pStyle") if ppr is not None else None
        sid = pstyle.get(f"{W}val") if pstyle is not None else ""
        if sid not in ("TOC1", "TOC2", "TOC3"):
            continue
        if p.find(f"{W}hyperlink") is None:
            continue
        toc_paras.append(p)
    toc_norms: list[str] = []
    toc_title_ts: list[list] = []
    for p in toc_paras:
        hl = p.find(f"{W}hyperlink")
        title_ts = []
        for el in hl.iter():
            tag = etree.QName(el).localname
            if tag in ("tab", "fldChar", "instrText"):
                break
            if tag == "t":
                title_ts.append(el)
        toc_norms.append(re.sub(r"\s+", "", "".join(x.text or "" for x in title_ts)))
        toc_title_ts.append(title_ts)
    return body_titles, toc_paras, toc_norms, toc_title_ts, all_heading_norms


def _toc_orphan_norms(toc_norms, body_titles, all_heading_norms) -> list[int]:
    """真孤儿条目索引：未配对且不在正文标题全集（正文无对应章节）。

    防御：正文标题全集为空（文档无标题结构——outlineLvl 未打/无 back
    标题，如裸格式源稿或本 action 在 ensure_heading_outline 之前执行）
    → 不判定任何条目为孤儿（保留缓存），绝不批量删除。"""
    if not all_heading_norms:
        return []
    paired = {i for i, _ in _toc_align(toc_norms, body_titles)}
    orphans = []
    for i, norm in enumerate(toc_norms):
        if i in paired:
            continue
        if norm not in all_heading_norms:
            orphans.append(i)
    return orphans


def update_toc_entries(path, params: dict) -> dict:
    """按正文标题重算 TOC 条目文本（用户 2026-08-18 拍板「开始排版前
    提醒用户手动更新好目录」→ 首选自动更新）。

    痛点实证（tx_7551a14078b5）：TOC 域缓存是源文档 WPS 旧的——「结论与
    展望59」vs 正文「结  论」、锚点 _Toc14310 全部与正文 _Toc10430 对不上、
    「1.3 研究内容与主要工作」重复。Word updateFields 会重算但**覆盖
    earlyz 打的格式**（用户「徒劳无功」）——earlyz 自己重算：只替换
    hyperlink 内文本 run，保留 pStyle/制表位/页码 PAGEREF 域结构。

    序列：TOC 条目按序 = 前置节标题（致谢/摘要/ABSTRACT…）+ 正文
    outlineLvl 0/1 标题（不含 2 级——TOC 域 o"1-2" 只收录两级）。
    """
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False, "entries_updated": 0}

    paras = list(body.iter(f"{W}p"))

    def _ptxt(p):
        return "".join(x.text or "" for x in p.iter(f"{W}t")).strip()

    # 1) 收集正文标题与目录缓存条目（update/verify 同源）
    body_titles, toc_paras, toc_norms, toc_title_ts, all_heading_norms = \
        _collect_toc_ctx(body)

    changed = 0
    for toc_idx, body_idx in _toc_align(toc_norms, body_titles):
        new_text = body_titles[body_idx]
        title_ts = toc_title_ts[toc_idx]
        if not title_ts:
            continue
        old_text = "".join(x.text or "" for x in title_ts)
        old_norm = re.sub(r"\s+", "", old_text)
        if old_norm == new_text:
            continue
        title_ts[0].text = new_text
        title_ts[0].set(f"{W}space", "preserve")
        for x in title_ts[1:]:
            x.text = ""
        changed += 1

    # 2) 真孤儿条目静默删除（2026-08-22 用户拍板「生成的就是正确的，
    #    不需要用户 F9 更新域——F9 会覆盖 earlyz 打的格式」）：
    #    目录缓存条目在正文标题全集（outlineLvl≤1 + back 标题角色）中
    #    不存在 = 正文无此章节（改名/删除后未刷新域）——删除该条目段，
    #    等同 Word 更新域收敛结果；目录条目是派生缓存（TOC1-3 样式段
    #    豁免内容 hash），非作者内容。正文有章节但未配对（如「参考文献」
    #    refs_heading 无 outlineLvl）→ 在全集内 → 保留。
    removed = 0
    # 逆序删除：升序删除会使后续 toc_idx 索引失效（列表收缩）
    for toc_idx in reversed(_toc_orphan_norms(toc_norms, body_titles, all_heading_norms)):
        p = toc_paras[toc_idx]
        parent = p.getparent()
        if parent is not None:
            parent.remove(p)
            removed += 1

    if changed or removed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0 or removed > 0,
            "entries_updated": changed, "orphans_removed": removed}


def verify_toc_entries_updated(path, params: dict) -> bool:
    """复查：TOC 条目文本与正文标题（同序）一致，且真孤儿条目已删除。"""
    try:
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return False
        body_titles, toc_paras, toc_norms, _tts, all_heading_norms = \
            _collect_toc_ctx(body)
        if not toc_paras:
            return True
        # 目录缓存条目数与正文标题数不要求相等：多出的条目（正文改名/
        # 删除后未刷新域）与少掉的条目（正文新增标题）都不构成失败；
        # 只要求「能按序配对的条目」文本与正文标题一致。
        for toc_idx, body_idx in _toc_align(toc_norms, body_titles):
            if toc_norms[toc_idx] != body_titles[body_idx]:
                return False
        # 真孤儿（正文无对应章节）必须已被 update_toc_entries 删除
        if _toc_orphan_norms(toc_norms, body_titles, all_heading_norms):
            return False
        return True
    except Exception:
        return False


def _is_ref_entry(text: str) -> bool:
    """参考文献条目谓词——动作（format_references）与 verifier 必须同一谓词（铁律）。

    「（」开头/含「示例」= 注释/示例话术段；「..」点前导 = 条目残段。
    两处判据曾不同源：「..」条目被 verifier 计入编号连续性 → 假失败
    （2026-10-14 审计 B1）。
    """
    return (text.startswith("（") or "示例" in text
            or re.match(r"^\.{2,}", text) is not None)


def format_references(path, params: dict) -> dict:
    """参考文献条目统一（用户 2026-08-18 拍板「补序号 + 注意模板段落样式」）。

    模板-4 铁证（idx=291-299）：条目 = 手写序号 `[N] ` + 宋体小四
    （sz=24）+ 1.5 倍行距（line=360 auto）+ **无缩进** + ascii=TNR。
    产物现状三错（被 body 规则误刷）：无序号 / line=460 exact（23磅）/
    firstLineChars=200（首行2字符）。

    本 action：定位「参考文献」标题之后的节内条目段（非空、非标题、非
    注释），①清 firstLineChars 缩进 ②行距改 line=360 auto ③序号缺失时
    补 `[N] `（N 按出现顺序）。序号补写为内容操作——用户已拍板授权。
    2026-10-14：已加入 CONTENT_CHANGE_ACTIONS 声明（catalog），运行期授权
    由 format_compiler 按 journal info 判定；[N] 前缀另有 analyzer hash 豁免。
    """
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False, "entries": 0, "numbered": 0}

    paras = list(body.iter(f"{W}p"))

    def _ptxt(p):
        return "".join(x.text or "" for x in p.iter(f"{W}t")).strip()

    # 定位参考文献标题段。参考文献是独立 back-matter role，不要求
    # outlineLvl（它不应被 Word TOC 当作普通 heading1）。
    from earlyz.ir.analyzer import _back_heading_role
    ref_title_idx = -1
    for i, p in enumerate(paras):
        t = _ptxt(p)
        if _back_heading_role(t) == "refs_heading":
            ref_title_idx = i
            break

    if ref_title_idx < 0:
        return {"changed": False, "entries": 0, "numbered": 0}

    changed = 0
    numbered = 0
    n = 0
    for p in paras[ref_title_idx + 1:]:
        t = _ptxt(p)
        if not t:
            continue
        # 跳过注释/示例话术段（含「（」「示例」「..」）——与 verifier 同一谓词
        if _is_ref_entry(t):
            continue
        ppr = p.find(f"{W}pPr")
        ends_section = ppr is not None and ppr.find(f"{W}sectPr") is not None
        # 跳过 TOC 条目（hyperlink）
        if p.find(f"{W}hyperlink") is not None:
            if ends_section:
                break
            continue
        # EZI-0066：后部角色标题（附录/致谢/作者简介/Background）必须在
        # outlineLvl 跳过**之前**判——2026-11-21 真机实锤：「附录A.」带
        # outlineLvl=0（真标题被 outline 过），先撞上「跳过 outline 段」
        # 分支 continue，永远走不到 break → 附录证明/作者简介被吞成
        # [8]-[12]。带 outline 的后部标题照样是区界，先判角色再谈跳过。
        from earlyz.ir.analyzer import _back_heading_role
        _bhr = _back_heading_role(t)
        if _bhr in ("bio_heading", "ack_heading", "appendix_heading", "refs_heading"):
            break
        if re.sub(r"\s+", "", t) in ("作者简介", "作者介绍"):
            break
        # 跳过其他章节标题（outlineLvl）
        oLvl = ppr.find(f"{W}outlineLvl") if ppr is not None else None
        if oLvl is not None:
            if ends_section:
                break
            continue

        n += 1
        refs = (params or {}).get("refs") or {}
        # ① 缩进：仅当模板条目无首行缩进时清掉（2026-08-23 CJC v9 实锤：
        # 只清 firstLine/firstLineChars——left/leftChars 是模板悬挂缩进
        # （ind_left=ind_hanging，refs_item 规则写入），上一版连同 left
        # 一起删导致「内容超过序号」（left 缺失=序号不复位）。
        if refs.get("ind_first_line_chars") == 0:
            if ppr is None:
                ppr = etree.Element(f"{W}pPr")
                p.insert(0, ppr)
            ind = ppr.find(f"{W}ind")
            if ind is not None:
                for attr in (f"{W}firstLine", f"{W}firstLineChars"):
                    if ind.get(attr) is not None:
                        del ind.attrib[attr]
                        changed += 1
                if len(ind.attrib) == 0:
                    ppr.remove(ind)
                    changed += 1
        # ② 行距：用模板抽出的 line/line_rule，没有就不改
        want_line = refs.get("line")
        want_rule = refs.get("line_rule")
        if want_line is not None:
            if ppr is None:
                ppr = etree.Element(f"{W}pPr")
                p.insert(0, ppr)
            sp = ppr.find(f"{W}spacing")
            if sp is None:
                sp = etree.Element(f"{W}spacing")
                _ppr_insert_ordered(ppr, sp)
            ls, lr = str(want_line), (want_rule or "auto")
            if sp.get(f"{W}line") != ls or sp.get(f"{W}lineRule") != lr:
                sp.set(f"{W}line", ls)
                sp.set(f"{W}lineRule", lr)
                changed += 1
        # ③ 序号：仅当模板条目要求编号。目标已有 Word 自动编号时
        # 保留其结构，不再写字面 [N]，否则会视觉重复为 “[1] [1]”。
        if refs.get("numbered"):
            m = re.match(r"^\[(\d+)\]", t)
            numpr = ppr.find(f"{W}numPr") if ppr is not None else None
            num_id = numpr.find(f"{W}numId") if numpr is not None else None
            auto_numbered = (
                numpr is not None
                and (num_id is None or num_id.get(f"{W}val") != "0")
            )
            if auto_numbered:
                if m:
                    runs = p.findall(f"{W}r")
                    if runs:
                        first_t = runs[0].find(f"{W}t")
                        if first_t is not None:
                            first_t.text = re.sub(r"^\[\d+\]\s*", "", first_t.text or "")
                            changed += 1
                numbered += 1
            elif not m:
                runs = p.findall(f"{W}r")
                if runs:
                    first_t = runs[0].find(f"{W}t")
                    if first_t is not None:
                        first_t.text = f"[{n}] " + (first_t.text or "")
                        changed += 1
                numbered += 1
            else:
                if int(m.group(1)) != n:
                    runs = p.findall(f"{W}r")
                    if runs:
                        first_t = runs[0].find(f"{W}t")
                        if first_t is not None:
                            first_t.text = re.sub(r"^\[\d+\]", f"[{n}]", first_t.text or "")
                            changed += 1
                numbered += 1

        # 节尾段可以同时承载最后一条参考文献；先处理该段，再结束。
        if ends_section:
            break

    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "entries": n, "numbered": numbered}


def verify_references(path, params: dict) -> bool:
    """复查：参考文献跟随模板行距/缩进，序号连续且不重复。

    零对象空过（vacuous pass，2026-09-25 张鑫佳实锤）：文档没有参考文献节
    （或节内无可识别条目）时动作 entries=0——verifier 不得判 False，否则
    「模板要求参考文献格式、而目标文档没有参考文献」会被误报为结构失败
    并拦截整个交付。与 verify_toc_entries_updated「无 TOC → True」同语义。
    """
    try:
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return True
        paras = list(body.iter(f"{W}p"))

        def _ptxt(p):
            return "".join(x.text or "" for x in p.iter(f"{W}t")).strip()

        from earlyz.ir.analyzer import _back_heading_role
        ref_title_idx = -1
        for i, p in enumerate(paras):
            t = _ptxt(p)
            if _back_heading_role(t) == "refs_heading":
                ref_title_idx = i
                break
        if ref_title_idx < 0:
            # 无参考文献节 → 无可校验对象（vacuous pass）
            return True
        n = 0
        for p in paras[ref_title_idx + 1:]:
            text = _ptxt(p)
            if not text:
                continue
            role = _back_heading_role(text)
            if role in ("ack_heading", "bio_heading", "appendix_heading"):
                break
            ppr = p.find(f"{W}pPr")
            ends_section = ppr is not None and ppr.find(f"{W}sectPr") is not None
            if p.find(f"{W}hyperlink") is not None:
                if ends_section:
                    break
                continue
            if ppr is not None and ppr.find(f"{W}outlineLvl") is not None:
                if ends_section:
                    break
                continue
            # 与动作同一谓词（_is_ref_entry）：含「（」「示例」「..」的条目跳过
            if _is_ref_entry(text):
                continue
            n += 1
            refs = (params or {}).get("refs") or {}
            if refs.get("numbered"):
                numpr = ppr.find(f"{W}numPr") if ppr is not None else None
                num_id = numpr.find(f"{W}numId") if numpr is not None else None
                auto_numbered = (
                    numpr is not None
                    and (num_id is None or num_id.get(f"{W}val") != "0")
                )
                literal = re.match(r"^\[(\d+)\]", text)
                if auto_numbered:
                    if literal is not None:
                        return False
                elif literal is None or int(literal.group(1)) != n:
                    return False
            want_line = refs.get("line")
            want_rule = refs.get("line_rule")
            if want_line is not None:
                sp = ppr.find(f"{W}spacing") if ppr is not None else None
                if (sp is None or sp.get(f"{W}line") != str(want_line)
                        or (want_rule and sp.get(f"{W}lineRule") != want_rule)):
                    return False
            if refs.get("ind_first_line_chars") == 0:
                ind = ppr.find(f"{W}ind") if ppr is not None else None
                if ind is not None and (ind.get(f"{W}firstLine") is not None
                                        or ind.get(f"{W}firstLineChars") is not None):
                    return False
            if ends_section:
                break
        # 节内无可识别条目（n==0）→ 零对象空过；有条目则已在循环内逐项
        # 校验，违规早已 return False。
        return True
    except Exception:
        return False


def ensure_toc_field(path, params: dict) -> dict:
    """确保目录是**真 Word TOC 域**（F9 可更新），缓存结果=earlyz 算好的条目。

    2026-09-26 用户拍板：「目录能自动更新（F9）这个是肯定要的」。
    与 update_toc_entries 的关系：条目文本/页码由 earlyz 先算好（见该动作），
    本动作只补/修**域外壳**——把条目区包进 TOC 域（fldChar begin +
    instrText TOC + fldChar separate … fldChar end），Word 视 separate 之后
    到 end 之间为「缓存结果」→ 打开显示的就是我们的条目，按 F9 才重建。

    不弹「更新域」弹窗的前提是 settings.xml **没有** w:updateFields
    （由 ensure_update_fields 负责删除）——两者分工，缺一不可。

    域指令策略（目标提供排什么）：源稿 TOC 域已带指令则原样保留；没有则用
    规范指令 `TOC \\o "1-3" \\h \\z \\u`。条目识别取双证据（带 PAGEREF 域，
    或段落样式名为 toc N），避免把正文误当条目。
    """
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False, "entries": 0}

    styles_map = _read_styles_map(path)

    def _is_toc_entry(p) -> bool:
        if p.findall(f".//{W}instrText"):
            for x in p.iter(f"{W}instrText"):
                if (x.text or "").strip().startswith("PAGEREF"):
                    return True
        ppr = p.find(f"{W}pPr")
        if ppr is not None:
            ps = ppr.find(f"{W}pStyle")
            if ps is not None:
                pid = ps.get(f"{W}val") or ""
                name = (styles_map.get(pid, "") or pid).strip().lower()
                if re.match(r"^toc\s*\d", name):
                    return True
        return False

    entries = [p for p in body.iter(f"{W}p") if _is_toc_entry(p)]
    if not entries:
        return {"changed": False, "entries": 0, "skipped": "no_toc_entries"}

    def _run(child):
        r = etree.Element(f"{W}r")
        r.append(child)
        return r

    def _fld(kind: str):
        e = etree.Element(f"{W}fldChar")
        e.set(f"{W}fldCharType", kind)
        return e

    # 已有指令：保留源稿的（目标提供排什么）
    instr = 'TOC \\o "1-3" \\h \\z \\u'
    for p in entries:
        for x in p.iter(f"{W}instrText"):
            t = (x.text or "").strip()
            if t.startswith("TOC"):
                instr = t
                break

    changed = 0
    first, last = entries[0], entries[-1]
    first_runs = [c for c in first if etree.QName(c).localname == "r"]
    last_runs = [c for c in last if etree.QName(c).localname == "r"]
    has_begin = any(
        (fc.get(f"{W}fldCharType") or "") == "begin"
        for r in first_runs for fc in r.iter(f"{W}fldChar")
    )
    # 只看段内直系 run——条目超链接里的 PAGEREF 自带 end，不能算域尾
    has_end = any(
        (fc.get(f"{W}fldCharType") or "") == "end"
        for r in last_runs for fc in r.iter(f"{W}fldChar")
    )
    if not has_begin:
        # 域头插在首条目段最前：begin / instrText / separate（都必须包在 w:r 里）
        instr_el = etree.Element(f"{W}instrText")
        instr_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        instr_el.text = f" {instr} "
        ppr = first.find(f"{W}pPr")
        pos = list(first).index(ppr) + 1 if ppr is not None else 0
        for off, el in enumerate((_run(_fld("begin")), _run(instr_el),
                                  _run(_fld("separate")))):
            first.insert(pos + off, el)
        changed += 1
    if not has_end:
        last.append(_run(_fld("end")))
        changed += 1

    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "entries": len(entries),
            "field_created": not has_begin, "field_closed": not has_end,
            "instruction": instr}


def verify_toc_field(path, params: dict) -> bool:
    """复查：TOC 域完整（首条目段有 begin + TOC instrText + separate，
    末条目段有 end），条目内 PAGEREF 域保留，settings 无 updateFields。"""
    try:
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return False
        styles_map = _read_styles_map(path)

        def _is_toc_entry(p) -> bool:
            for x in p.iter(f"{W}instrText"):
                if (x.text or "").strip().startswith("PAGEREF"):
                    return True
            ppr = p.find(f"{W}pPr")
            if ppr is not None:
                ps = ppr.find(f"{W}pStyle")
                if ps is not None:
                    pid = ps.get(f"{W}val") or ""
                    name = (styles_map.get(pid, "") or pid).strip().lower()
                    if re.match(r"^toc\s*\d", name):
                        return True
            return False

        entries = [p for p in body.iter(f"{W}p") if _is_toc_entry(p)]
        if not entries:
            return True  # 无目录区：不适用
        first, last = entries[0], entries[-1]
        kinds_first = set()
        for r in first:
            if etree.QName(r).localname != "r":
                continue
            for fc in r.iter(f"{W}fldChar"):
                kinds_first.add(fc.get(f"{W}fldCharType") or "")
        has_instr = any((x.text or "").strip().startswith("TOC")
                        for x in first.iter(f"{W}instrText"))
        kinds_last = set()
        for r in last:
            if etree.QName(r).localname != "r":
                continue
            for fc in r.iter(f"{W}fldChar"):
                kinds_last.add(fc.get(f"{W}fldCharType") or "")
        if not ({"begin", "separate"} <= kinds_first and has_instr):
            return False
        if "end" not in kinds_last:
            return False
        try:
            with zipfile.ZipFile(path) as z:
                st = etree.fromstring(z.read("word/settings.xml"))
            if st.find(f"{W}updateFields") is not None:
                return False
        except (KeyError, etree.XMLSyntaxError):
            pass
        return True
    except Exception:
        return False


def unfold_toc_field(path, params: dict) -> dict:
    """TOC 域展开为静态文本（2026-08-18 用户彻查「打开就弹更新域」）。

    用户截图实证：updateFields=true + TOC/PAGEREF 域 → Word 每次打开弹
    「是否更新域」；且 Word 重算 TOC 会**覆盖 earlyz 打的目录格式**
    （用户「修这边那边又出问题」）。earlyz 已用 update_toc_entries 把
    条目文本算好——这里把 TOC 域**静态化**：删掉 fldChar begin/separate/
    end 与 instrText，保留 hyperlink（文本+制表位+页码文本 run）。

    效果：Word 打开不再弹窗（无域可更新）、目录不会被重算破坏。
    注意：PAGEREF 域（每个条目内的页码域）一并静态化——页码文本已是
    earlyz 重算结果（update_toc_entries 保留的域缓存值）。
    """
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False, "paragraphs_unfolded": 0}

    changed = 0
    # 拆 TOC **外层域**（TOC \o 指令），保留条目内 PAGEREF 域——
    # 2026-08-18 页码问题彻查：PAGEREF 域必须保留，Word 打开 updateFields
    # 自动重算条目页码（目录页码=真实页码）；只拆外层 TOC 域防 Word 重建
    # 条目（格式丢失）。正文 REF 交叉引用域不动。
    # TOC 外层域结构（实证 idx=65）：fldChar begin / instrText TOC \o /
    # fldChar separate / hyperlink(条目: 文本+tab+PAGEREF域) / fldChar end
    # ——外层 fldChar 与 instrText 在 hyperlink **外**，PAGEREF 在 hyperlink 内。
    for p in body.iter(f"{W}p"):
        has_toc_instr = any(
            (x.text or "").strip().startswith("TOC")
            for x in p.iter(f"{W}instrText")
        )
        if not has_toc_instr:
            continue
        removed = 0
        # 只删 hyperlink 外的 fldChar/instrText（外层 TOC 域指令）——
        # 直接子 run（w:r）内的 fldChar/instrText 属于外层域；hyperlink
        # 内的 PAGEREF 域保留（Word 打开自动重算页码）。
        for r in list(p):
            if etree.QName(r).localname != "r":
                continue
            for el in list(r):
                tag = etree.QName(el).localname
                if tag == "fldChar" or tag == "instrText":
                    r.remove(el)
                    removed += 1
        # 清空 run 清理
        for r in list(p):
            if etree.QName(r).localname == "r" and len(r) == 0 \
                    and (r.text is None or not r.text.strip()):
                parent = r.getparent()
                if parent is not None:
                    parent.remove(r)
                    removed += 1
        if removed:
            changed += 1

    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "paragraphs_unfolded": changed}


def verify_toc_unfolded(path, params: dict) -> bool:
    """复查：TOC 外层域已拆（无 TOC \\o 指令）；条目内 PAGEREF 域保留
    （Word 打开自动重算页码）；正文 REF 域不受影响。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return False
        for p in body.iter(f"{W}p"):
            ppr = p.find(f"{W}pPr")
            pstyle = ppr.find(f"{W}pStyle") if ppr is not None else None
            sid = pstyle.get(f"{W}val") if pstyle is not None else ""
            if sid not in ("TOC1", "TOC2", "TOC3"):
                continue
            for el in p.iter(f"{W}instrText"):
                if (el.text or "").strip().startswith("TOC"):
                    return False
        return True
    except Exception:
        return False


def fit_tables_to_page(target_path: Path, params: dict) -> dict:
    """表格宽度封顶到所在节的版心宽（R2，2026-10-19）。

    期刊模板实测：A4 版心宽的表套到 Letter/期刊模板后，表体右边缘超出页面
    131pt（`image/table bbox overflow`，5 处文本块落在同一张表上）。这里按
    「节 → 版心宽」给每个表封顶：tblW 与 gridCol 之和超过即等比回缩。
    """
    path = Path(target_path)
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"ok": False, "changed": False}
    widths: list[int] = []
    for sp in root.iter(f"{W}sectPr"):
        pg, mar = sp.find(f"{W}pgSz"), sp.find(f"{W}pgMar")
        if pg is None or mar is None:
            widths.append(0)
            continue
        try:
            w = (int(pg.get(f"{W}w") or 0) - int(mar.get(f"{W}left") or 0)
                 - int(mar.get(f"{W}right") or 0))
        except (TypeError, ValueError):
            w = 0
        # 栏感知（2026-10-19）：分栏后每栏可用宽度只有版心的一部分，
        # 按整幅版心封顶的表会挤爆栏/页面
        cols = sp.find(f"{W}cols")
        if cols is not None and w > 0:
            try:
                num = max(1, int(cols.get(f"{W}num") or 1))
                space = int(cols.get(f"{W}space") or 425)
            except (TypeError, ValueError):
                num, space = 1, 425
            if num > 1:
                w = max(1000, (w - space * (num - 1)) // num)
        widths.append(w)
    if not any(widths):
        return {"ok": True, "changed": False, "reason": "no page geometry"}
    idx = 0
    changed = 0
    for el in body.iter():
        if el.tag == f"{W}sectPr":
            idx += 1
            continue
        if el.tag != f"{W}tbl":
            continue
        cap = widths[min(idx, len(widths) - 1)]
        if cap <= 0:
            continue
        grid = el.find(f"{W}tblGrid")
        cols = grid.findall(f"{W}gridCol") if grid is not None else []
        cur = 0
        vals = []
        for c in cols:
            try:
                v = int(c.get(f"{W}w") or 0)
            except (TypeError, ValueError):
                v = 0
            vals.append((c, v))
            cur += v
        # 单元格自身宽度（tcW）可能大于栅格合计——竖排文字表格实测：
        # gridCol 合计在版心内，实际渲染宽度却到 738pt（2026-10-19）
        tc_vals = []
        for tr in el.findall(f"{W}tr"):
            row_sum = 0
            row_cells = []
            for tc in tr.findall(f"{W}tc"):
                tcpr = tc.find(f"{W}tcPr")
                tw = tcpr.find(f"{W}tcW") if tcpr is not None else None
                if tw is not None and (tw.get(f"{W}type") or "dxa") == "dxa":
                    try:
                        v = int(tw.get(f"{W}w") or 0)
                    except (TypeError, ValueError):
                        v = 0
                else:
                    v = 0
                row_cells.append((tw, v))
                row_sum += v
            tc_vals.append((row_cells, row_sum))
        eff = max([cur] + [rs for _, rs in tc_vals])
        # 自动布局表（tblW=auto 且没有 fixed 布局）在 Word 里按内容撑开，
        # gridCol 写了宽度也不算数。2026-10-26 东南本科实测：封面 AI 说明表的
        # 单元格一句 80 字填写示例 → 行拉到 755pt（页宽 595），gridCol 却
        # 非 0 导致下面的 eff<=0 分支不触发。含长文本的自动表一律按自动表处理。
        _tp = el.find(f"{W}tblPr")
        _tw = _tp.find(f"{W}tblW") if _tp is not None else None
        _lay = _tp.find(f"{W}tblLayout") if _tp is not None else None
        _autofit = ((_tw is None or (_tw.get(f"{W}type") or "auto") == "auto")
                    and (_lay is None or (_lay.get(f"{W}type") or "") != "fixed"))
        _long = max((len("".join(t.text or "" for t in tc.iter(f"{W}t")).strip())
                     for tc in el.iter(f"{W}tc")), default=0)
        if _autofit and _long >= 40:
            eff = 0
        if eff <= 0:
            # 全自动宽度表（tblW/gridCol/tcW 都是 auto）：Word 按内容撑开，
            # 竖排（textDirection）与合并（gridSpan/vMerge）单元格会把它撑爆页面。
            # 2026-10-21 实测：监理/设计/造价/浙江招标 4 份合同均命中，
            # 越界行文本是表格行内容（'…纵向合并 竖排文字普通单元格'）。
            n = max(1, len(vals))
            each = cap // n
            for c, _ in vals:
                c.set(f"{W}w", str(each))
            tblpr0 = el.find(f"{W}tblPr")
            if tblpr0 is not None:
                tw0 = tblpr0.find(f"{W}tblW")
                if tw0 is None:
                    tw0 = etree.SubElement(tblpr0, f"{W}tblW")
                tw0.set(f"{W}w", str(cap))
                tw0.set(f"{W}type", "dxa")
                lay0 = tblpr0.find(f"{W}tblLayout")
                if lay0 is None:
                    lay0 = etree.SubElement(tblpr0, f"{W}tblLayout")
                lay0.set(f"{W}type", "fixed")
            changed += 1
            continue
        if eff <= cap:
            continue
        ratio = cap / eff
        for c, v in vals:
            c.set(f"{W}w", str(max(200, int(v * ratio))))
        for row_cells, _rs in tc_vals:
            for tw, v in row_cells:
                if tw is not None and v:
                    tw.set(f"{W}w", str(max(200, int(v * ratio))))
        tblpr = el.find(f"{W}tblPr")
        if tblpr is not None:
            tw = tblpr.find(f"{W}tblW")
            if tw is None:
                tw = etree.SubElement(tblpr, f"{W}tblW")
            tw.set(f"{W}w", str(cap))
            tw.set(f"{W}type", "dxa")
            # 强制固定布局：否则 Word 自动布局会按竖排文字再撑宽
            lay = tblpr.find(f"{W}tblLayout")
            if lay is None:
                lay = etree.SubElement(tblpr, f"{W}tblLayout")
            lay.set(f"{W}type", "fixed")
        changed += 1
    if changed:
        _write_part(path, "word/document.xml", root)
    return {"ok": True, "changed": changed > 0, "tables_capped": changed}


def verify_tables_within_page(target_path: Path, params: dict) -> bool:
    """复查：没有表的 gridCol 之和超出所在节版心宽。"""
    try:
        root = _read_part(Path(target_path), "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return False
        widths = []
        for sp in root.iter(f"{W}sectPr"):
            pg, mar = sp.find(f"{W}pgSz"), sp.find(f"{W}pgMar")
            if pg is None or mar is None:
                widths.append(0); continue
            widths.append(int(pg.get(f"{W}w") or 0) - int(mar.get(f"{W}left") or 0)
                          - int(mar.get(f"{W}right") or 0))
        idx = 0
        for el in body.iter():
            if el.tag == f"{W}sectPr":
                idx += 1; continue
            if el.tag != f"{W}tbl":
                continue
            cap = widths[min(idx, len(widths) - 1)] if widths else 0
            if cap <= 0:
                continue
            grid = el.find(f"{W}tblGrid")
            total = sum(int(c.get(f"{W}w") or 0) for c in (grid.findall(f"{W}gridCol") if grid is not None else []))
            if total > cap + 20:
                return False
        return True
    except Exception:  # noqa: BLE001
        return False


def normalize_table_widths(path, params: dict) -> dict:
    """按模板 table_width 把表拉到栏宽或版心比例。无取值则跳过。"""
    path = Path(path)
    params = params or {}
    tw = params.get("table_width") or {}
    ratio = tw.get("ratio")
    ref = tw.get("ref") or "page"
    if not ratio:
        return {"changed": False, "skipped": "no template table_width"}
    try:
        ratio = float(ratio)
    except (TypeError, ValueError):
        return {"changed": False, "skipped": "bad table_width"}
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    geoms = _sect_geoms(root)
    widened = 0
    for tbl in root.iter(f"{W}tbl"):
        grid = tbl.find(f"{W}tblGrid")
        if grid is None:
            continue
        cols = grid.findall(f"{W}gridCol")
        widths = []
        for c in cols:
            try:
                widths.append(int(c.get(f"{W}w", "0")))
            except (TypeError, ValueError):
                widths = []
                break
        if not widths or any(w <= 0 for w in widths):
            continue
        total = sum(widths)
        ref_w = _ref_width_for_element(tbl, geoms, ref)
        if not ref_w:
            continue
        target = max(1, int(round(ratio * ref_w)))
        # 过窄才拉齐；双栏溢出才缩小
        if total >= 0.90 * target and not (ref == "column" and total > ref_w * 1.02):
            continue
        factor = target / total
        new_widths = [max(1, round(w * factor)) for w in widths]
        for c, w in zip(cols, new_widths):
            c.set(f"{W}w", str(w))
        for tr in tbl.findall(f"{W}tr"):
            tcs = tr.findall(f"{W}tc")
            for tc, w in zip(tcs, new_widths):
                tcpr = tc.find(f"{W}tcPr")
                if tcpr is None:
                    continue
                tcw = tcpr.find(f"{W}tcW")
                if tcw is not None:
                    tcw.set(f"{W}w", str(w))
                    tcw.set(f"{W}type", "dxa")
        tblpr = tbl.find(f"{W}tblPr")
        if tblpr is not None:
            tblw = tblpr.find(f"{W}tblW")
            if tblw is None:
                tblw = etree.Element(f"{W}tblW")
                _tblpr_insert_ordered(tblpr, tblw)
            tblw.set(f"{W}w", str(sum(new_widths)))
            tblw.set(f"{W}type", "dxa")
        widened += 1
    if widened:
        _write_part(path, "word/document.xml", root)
    return {"changed": widened > 0, "tables_widened": widened}


def verify_table_widths(path, params: dict) -> bool:
    tw = (params or {}).get("table_width") or {}
    if not tw.get("ratio"):
        return True
    try:
        root = _read_part(Path(path), "word/document.xml")
        geoms = _sect_geoms(root)
        ratio = float(tw["ratio"])
        ref = tw.get("ref") or "page"
        for tbl in root.iter(f"{W}tbl"):
            grid = tbl.find(f"{W}tblGrid")
            if grid is None:
                continue
            total = 0
            for c in grid.findall(f"{W}gridCol"):
                try:
                    total += int(c.get(f"{W}w", "0"))
                except (TypeError, ValueError):
                    total = 0
                    break
            if total <= 0:
                continue
            ref_w = _ref_width_for_element(tbl, geoms, ref)
            if not ref_w:
                continue
            if total < 0.85 * ratio * ref_w:
                return False
            if ref == "column" and total > ref_w * 1.08:
                return False
        return True
    except Exception:
        return False


def _sect_geoms(root) -> list[dict]:
    out = []
    for sp in root.iter(f"{W}sectPr"):
        from earlyz.spec.structure import _sect_cols, anchor_token
        g = _sect_cols(sp)
        g["sp"] = sp
        out.append(g)
    return out


def _ref_width_for_element(el, geoms: list[dict], ref: str) -> int | None:
    sp = None
    cur = el
    while cur is not None:
        nxt = cur.getnext()
        if nxt is not None and nxt.tag == f"{W}sectPr":
            sp = nxt
            break
        ppr = cur.find(f"{W}pPr") if hasattr(cur, "find") else None
        if ppr is not None and ppr.find(f"{W}sectPr") is not None:
            sp = ppr.find(f"{W}sectPr")
            break
        cur = cur.getparent()
        if cur is not None and cur.tag.endswith("body"):
            break
    geom = None
    if sp is not None:
        for g in geoms:
            if g.get("sp") is sp:
                geom = g
                break
    if geom is None and geoms:
        geom = geoms[-1]
    if not geom:
        return None
    if ref == "column" and geom.get("col_widths"):
        return int(geom["col_widths"][0])
    return geom.get("page_content")


def apply_section_columns(path, params: dict) -> dict:
    """把模板 front_cols / body_cols 写到目标对应节。

    模板前节单栏、后节双栏，而源稿只有一节时：按模板双栏节首段锚点
    补 continuous 分节（不改字）。找不到锚点则整篇用 body_cols。
    """
    path = Path(path)
    params = params or {}
    body_cols = params.get("body_cols") or {}
    front_cols = params.get("front_cols") or {}
    if not body_cols and not front_cols:
        return {"changed": False, "skipped": "no template cols"}
    # 硬约束：正文栏数 > 3 一律按单栏处理。
    # 2026-10-21 实测（建设工程施工合同等 5 份）：spec 从模板某节读出 cols_num=4，
    # 实际模板全部是单栏——把单栏合同排成 4 栏后正文被压进窄栏，长行溢出页面
    # （bbox overflow: p1 text_block x 到 625.6，页宽 595.6）。
    # 4 栏以上的"分栏"在学术/合同/公文里不是正文排版，是表单栅格误读。
    _num = int(body_cols.get("cols_num") or 1)
    if _num > 3:
        body_cols = {"cols_num": 1, "clamped_from": _num}
    root = _read_part(path, "word/document.xml")
    split_target, inserted = _maybe_insert_col_section_break(root, params)
    # 模板后部栏变化（附录/Background）只可能出现在正文起点之后——前部（目录条目
    # 「附  录28」）不得当锚点（2026-09-29 真论文：分节插进了目录区）
    extra_n = _maybe_insert_plan_breaks(root, params, after=split_target)
    ordered = _ordered_sections(root)
    split_i = _col_split_index(ordered, params)
    if split_i is None and inserted and len(ordered) >= 2:
        # 正文起点 = 起点段所在节。旧逻辑写死 1（只对单节源稿成立）：
        # 多节源稿（学位论文 17 节）会把第二封面/声明/摘要/目录全套成正文双栏。
        split_i = _section_index_of(root, split_target)
    elif split_i is None and split_target is not None:
        # 起点段本来就另起一节（未插节）：同样按它所在节切 front/body
        split_i = _section_index_of(root, split_target)
    first_content = None
    if split_i is None:
        for i, item in enumerate(ordered):
            st = _classify_section(item["start_text"], params)
            if st == "body" or st not in ("other", "empty"):
                first_content = i
                break
    changed = (1 if inserted else 0) + extra_n
    plan = list(params.get("section_plan") or [])
    # 2026-08-23 CJC 实锤：模板 8 节 vs 源稿 5 节——纯 index 对齐会把模板
    # 「示例占位符节」（附录X/Background 单栏）错套到源稿正文节（2.1-3.3
    # 公式区被置单栏）。plan 仅在节数完全对应时启用；否则按语义一刀切：
    # 第一内容节前=front_cols（首页单栏），其余=body_cols（正文双栏）。
    use_plan = bool(plan) and (inserted or extra_n or split_i is not None) \
        and len(plan) == len(ordered)
    for i, item in enumerate(ordered):
        if use_plan:
            use = plan[i] if i < len(plan) else (body_cols or front_cols)
        elif split_i is not None:
            use = front_cols if i < split_i else body_cols
        elif first_content is None:
            use = body_cols or front_cols
        elif i < first_content:
            use = front_cols or body_cols
        else:
            use = body_cols or front_cols
        if not use:
            continue
        if i > 0 and use.get("index") == 0 and use.get("sect_type"):
            # 机制层（OOXML 事实，与模板无关）：w:type 描述「本节相对上一节怎么开始」，
            # 模板**第一节**没有上一节，Word 忽略它的 type → 模板第 0 节的 sect_type
            # 不承载任何版式意图。把它写到目标的后续前部节会凭空造出「连续分节」
            # （2026-09-29 真论文：两页封面并成一页、声明/摘要/目录连排）。
            # 节起始方式保持目标原稿。
            use = {k: v for k, v in use.items() if k != "sect_type"}
        if _write_cols(item["sp"], use):
            changed += 1
        if _write_page_geom(item["sp"], use):
            changed += 1
    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "sections": changed, "split": inserted,
            "extra_splits": extra_n}


_ANCHOR_TOKEN_RE = re.compile(
    r"^[\s\d.、．第]*([A-Za-z]{3,}|[\u4e00-\u9fff]{1,12})"
)


def _anchor_token(text: str) -> str:
    from earlyz.spec.structure import anchor_token
    return anchor_token(text)


def _col_split_index(ordered: list[dict], params: dict) -> int | None:
    kind = (params or {}).get("body_section_anchor_kind") or ""
    if kind == "first_heading":
        return None  # 插入后再按节首 heading 判断
    anchor = (params or {}).get("body_section_anchor") or ""
    tok = _anchor_token(anchor)
    if not tok:
        return None
    for i, item in enumerate(ordered):
        if _anchor_token(item.get("start_text") or "") == tok:
            return i
    return None


def _find_split_paragraph(body_el, params: dict):
    """token 锚点优先；对不上（国内刊模板注释段）则找前部词之后的第一个标题。"""
    kind = (params or {}).get("body_section_anchor_kind") or ""
    anchor = (params or {}).get("body_section_anchor") or ""
    tok = _anchor_token(anchor)
    front = {_anchor_token(x) for x in (params or {}).get("front_tokens") or [] if x}
    from earlyz.ir.analyzer import classify_paragraph_role
    paras = []
    for p in body_el.iter(f"{W}p"):
        if any(a.tag == f"{W}tbl" for a in p.iterancestors()):
            continue
        t = "".join(x.text or "" for x in p.iter(f"{W}t")).strip()
        if not t:
            continue
        ppr = p.find(f"{W}pPr")
        pstyle = ppr.find(f"{W}pStyle") if ppr is not None else None
        sid = pstyle.get(f"{W}val", "") if pstyle is not None else ""
        role = classify_paragraph_role(ppr, sid, False, False, False, t)
        paras.append((p, t, str(role)))
    # token 锚点（模板双栏节首词直接命中时优先）
    for p, t, _role in paras:
        ptok = _anchor_token(t)
        if tok and kind != "first_heading" and (ptok == tok or ptok.startswith(tok)):
            return p
    # 正文分节锚点：front matter 结束点 = 最后一个 摘要/关键词/分类号 段。
    # 中英文题名段即使被 set_paragraph_format 标成 outlineLvl 0（模板标题
    # 自带 outline），也绝不能当正文标题——CJC 中英文题名都带 outline，
    # 锚点必须落在「1 引言」；题名段必然位于最后一个 front-matter 段之前。
    front_end = 0
    for i, (_p, _t, role) in enumerate(paras, start=1):
        if role in ("abstract", "keywords", "classification"):
            front_end = i
    for i, (p, t, role) in enumerate(paras, start=1):
        if role.startswith("heading") and i > 1 and i > front_end \
                and _anchor_token(t) not in front:
            return p
    return None


def _maybe_insert_col_section_break(root, params: dict) -> tuple:
    """定位模板「单栏→双栏」边界在目标里的起点段；缺分节时在它前面补 continuous 节。

    返回 (起点段 lxml 元素 | None, 是否新插了分节)。起点段已是某节的首个内容段
    （多节源稿，如学位论文「1 绪论」本来就另起一节）时**不再插节**——否则凭空多出
    只含空段的节（2026-09-29 真论文：目录与绪论之间多出 2 个空节）。
    """
    front = params.get("front_cols") or {}
    body = params.get("body_cols") or {}
    try:
        fn = int(front.get("cols_num") or 1)
        bn = int(body.get("cols_num") or 1)
    except (TypeError, ValueError):
        return None, False
    if fn == bn:
        return None, False
    ordered = _ordered_sections(root)
    if _col_split_index(ordered, params) is not None:
        return None, False
    body_el = root.find(f"{W}body")
    if body_el is None:
        return None, False
    target = _find_split_paragraph(body_el, params)
    if target is None:
        return None, False
    if _starts_existing_section(target):
        return target, False
    return target, bool(_insert_continuous_before(body_el, target, front))


def _starts_existing_section(para) -> bool:
    """段落是否已是某节的首个内容段：向前只隔空段就碰到带 sectPr 的段。

    文档开头（前面没有任何 sectPr）不算——单节源稿仍需补分节（期刊路径不变）。
    """
    prev = para.getprevious()
    while prev is not None:
        if prev.tag == f"{W}p":
            ppr = prev.find(f"{W}pPr")
            if ppr is not None and ppr.find(f"{W}sectPr") is not None:
                return True
            if not _is_empty_para(prev):
                return False
        elif prev.tag != f"{W}bookmarkStart" and prev.tag != f"{W}bookmarkEnd":
            return False
        prev = prev.getprevious()
    return False


def _section_index_of(root, para) -> int | None:
    """段落所在节序号（与 `_ordered_sections` 同口径：节 = 截至带 sectPr 段为止）。"""
    body_el = root.find(f"{W}body")
    if body_el is None or para is None:
        return None
    k = 0
    for p in body_el.iter(f"{W}p"):
        if p is para:
            return k
        ppr = p.find(f"{W}pPr")
        if ppr is not None and ppr.find(f"{W}sectPr") is not None:
            k += 1
    return None


def _maybe_insert_plan_breaks(root, params: dict, after=None) -> int:
    """模板后部栏变化（附录单栏、Background 等）：源稿有对应段首词才切。

    `after`：正文起点段——锚点只在它之后找（前部目录条目「附  录28」不是附录）。
    """
    plan = list((params or {}).get("section_plan") or [])
    if len(plan) < 3:
        return 0
    body_el = root.find(f"{W}body")
    if body_el is None:
        return 0
    extras = []
    for spec in plan[2:]:
        tok = spec.get("token") or ""
        if spec.get("kind") in ("skip", "first_heading") or not tok:
            continue
        extras.append((tok, spec))
    n = 0
    seen: set[int] = set()
    for tok, spec in extras:
        target = None
        started = after is None
        for p in body_el.iter(f"{W}p"):
            if not started:
                started = p is after
                if not started:
                    continue
            if id(p) in seen:
                continue
            t = "".join(x.text or "" for x in p.iter(f"{W}t")).strip()
            if t and (_anchor_token(t) == tok or _anchor_token(t).startswith(tok)):
                target = p
                break
        if target is None:
            continue
        seen.add(id(target))
        # 锚点段已另起一节（学位论文「附  录」本就有分节）：不再补节，防凭空空节
        if _starts_existing_section(target):
            continue
        try:
            pi = plan.index(spec)
            prev = plan[pi - 1] if pi > 0 else (params.get("body_cols") or {})
        except ValueError:
            prev = params.get("body_cols") or {}
        if _insert_continuous_before(body_el, target, prev):
            n += 1
    return n


def _insert_continuous_before(body_el, para, front_cols: dict) -> bool:
    # IEEE 模板把 continuous 放在标题区后的空段上，不要写在作者段自身。
    host = etree.Element(f"{W}p")
    para.addprevious(host)
    ppr = etree.SubElement(host, f"{W}pPr")
    tail = body_el.find(f"{W}sectPr")
    sp = copy.deepcopy(tail) if tail is not None else etree.Element(f"{W}sectPr")
    for child in list(sp):
        if child.tag in (f"{W}headerReference", f"{W}footerReference"):
            sp.remove(child)
    typ = sp.find(f"{W}type")
    if typ is None:
        typ = etree.Element(f"{W}type")
        sp.insert(0, typ)
    typ.set(f"{W}val", "continuous")
    _write_page_geom(sp, front_cols)
    _write_cols(sp, front_cols)
    ppr.append(sp)
    return True


def verify_section_columns(path, params: dict) -> bool:
    body_cols = (params or {}).get("body_cols") or {}
    want = int(body_cols.get("cols_num") or 0)
    if want < 2:
        return True
    try:
        root = _read_part(Path(path), "word/document.xml")
        nums = []
        for sp in root.iter(f"{W}sectPr"):
            cols = sp.find(f"{W}cols")
            n = 1
            if cols is not None:
                try:
                    n = int(cols.get(f"{W}num") or 1)
                except (TypeError, ValueError):
                    n = 1
            nums.append(n)
        return any(n >= 2 for n in nums)
    except Exception:
        return False


def _write_cols(sp, spec: dict) -> bool:
    if sp is None or not spec:
        return False
    num = int(spec.get("cols_num") or spec.get("num") or 1)
    old = sp.find(f"{W}cols")
    el = etree.Element(f"{W}cols")
    if num != 1:
        el.set(f"{W}num", str(num))
    if spec.get("cols_space") is not None:
        el.set(f"{W}space", str(int(spec["cols_space"])))
    elif spec.get("space") is not None:
        el.set(f"{W}space", str(int(spec["space"])))
    if spec.get("cols_equal") is False:
        el.set(f"{W}equalWidth", "0")
    spaces = spec.get("col_spaces") or []
    for i, w in enumerate(spec.get("col_widths") or []):
        c = etree.SubElement(el, f"{W}col")
        c.set(f"{W}w", str(int(w)))
        if i < len(spaces):
            c.set(f"{W}space", str(int(spaces[i])))
    if old is not None:
        idx = list(sp).index(old)
        sp.remove(old)
        sp.insert(idx, el)
    else:
        sp.append(el)
    return True


def _write_page_geom(sp, spec: dict) -> bool:
    """纸张/边距/分节类型跟模板走。缺 type 的双栏节会被 Word 当成 nextPage，摘要单独占页。"""
    if sp is None or not spec:
        return False
    changed = False
    want_type = spec.get("sect_type")
    if want_type:
        typ = sp.find(f"{W}type")
        if typ is None:
            typ = etree.Element(f"{W}type")
            sp.insert(0, typ)
        if typ.get(f"{W}val") != str(want_type):
            typ.set(f"{W}val", str(want_type))
            changed = True
    pw, ph = spec.get("pg_w"), spec.get("pg_h")
    if pw and ph:
        pg = sp.find(f"{W}pgSz")
        if pg is None:
            pg = etree.Element(f"{W}pgSz")
            typ = sp.find(f"{W}type")
            if typ is not None:
                typ.addnext(pg)
            else:
                sp.insert(0, pg)
        if pg.get(f"{W}w") != str(int(pw)) or pg.get(f"{W}h") != str(int(ph)):
            pg.set(f"{W}w", str(int(pw)))
            pg.set(f"{W}h", str(int(ph)))
            changed = True
    mar_spec = spec.get("pg_mar") or {}
    if mar_spec:
        mar = sp.find(f"{W}pgMar")
        if mar is None:
            mar = etree.Element(f"{W}pgMar")
            pg = sp.find(f"{W}pgSz")
            if pg is not None:
                pg.addnext(mar)
            else:
                sp.append(mar)
        for k, v in mar_spec.items():
            if mar.get(f"{W}{k}") != str(int(v)):
                mar.set(f"{W}{k}", str(int(v)))
                changed = True
    # 新建的 w:type 用 insert(0) 会排到 header/footerReference 前面，违反 CT_SectPr
    # 顺序（2026-09-30 真论文×CJC 产物 s8–s16 实测）；统一按 schema 重排。
    if _normalize_child_order(sp, _CT_SECTPR_ORDER):
        changed = True
    return changed


_COMBINED_FIG_RE = re.compile(
    r"^(图\s*\d+[^\n]*?)\s+(Fig(?:ure)?\.?\s*\d+.*)$", re.I | re.S
)
_COMBINED_TBL_RE = re.compile(
    r"^(表\s*\d+[^\n]*?)\s+(Table\s+\d+.*)$", re.I | re.S
)


def _split_combined_caption_lines(root) -> int:
    """模板是两行中英图题时，源稿挤在一段则拆成两段。不发明英文。"""
    n = 0
    body = root.find(f"{W}body")
    if body is None:
        return 0
    for p in list(body.iter(f"{W}p")):
        if any(a.tag == f"{W}tbl" for a in p.iterancestors()):
            continue
        texts = list(p.iter(f"{W}t"))
        full = "".join(t.text or "" for t in texts).strip()
        m = _COMBINED_FIG_RE.match(full) or _COMBINED_TBL_RE.match(full)
        if not m:
            continue
        cn, en = m.group(1).strip(), m.group(2).strip()
        if texts:
            texts[0].text = cn
            for t in texts[1:]:
                t.text = ""
        np = copy.deepcopy(p)
        nts = list(np.iter(f"{W}t"))
        if nts:
            nts[0].text = en
            for t in nts[1:]:
                t.text = ""
        p.addnext(np)
        n += 1
    return n


def _apply_equation_jc(root, eq: dict) -> int:
    """只写公式段 jc；不进 oMath/OLE 子树。"""
    want = (eq or {}).get("jc")
    if not want:
        return 0
    from earlyz.ir.analyzer import classify_paragraph_role, _has_equation_ole
    n = 0
    for p in root.iter(f"{W}p"):
        ppr = p.find(f"{W}pPr")
        text = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        has_om = p.find(f".//{M}oMathPara") is not None or (
            p.find(f".//{M}oMath") is not None and not text)
        role = classify_paragraph_role(
            ppr, "", False, False, False, text,
            has_omath_para=has_om, has_equation_ole=_has_equation_ole(p),
        )
        if role != "equation_display":
            continue
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            p.insert(0, ppr)
        jc = ppr.find(f"{W}jc")
        if jc is None:
            jc = etree.Element(f"{W}jc")
            ppr.append(jc)
        if jc.get(f"{W}val") != want:
            jc.set(f"{W}val", want)
            n += 1
    return n


def _write_caption_spacing(para, cap: dict, kind: str) -> bool:
    """把模板图/表题段距写进产物。无样本键则不发明。"""
    before = cap.get(f"{kind}_before")
    after = cap.get(f"{kind}_after")
    bl = cap.get(f"{kind}_before_lines")
    al = cap.get(f"{kind}_after_lines")
    if before is None and after is None and bl is None and al is None:
        return False
    ppr = para.find(f"{W}pPr")
    if ppr is None:
        ppr = etree.Element(f"{W}pPr")
        para.insert(0, ppr)
    sp = ppr.find(f"{W}spacing")
    if sp is None:
        sp = etree.Element(f"{W}spacing")
        _ppr_insert_ordered(ppr, sp)
    changed = False
    def _set(attr, val):
        nonlocal changed
        if val is None:
            return
        sval = str(val)
        if sp.get(f"{W}{attr}") != sval:
            sp.set(f"{W}{attr}", sval)
            changed = True
    _set("before", before)
    _set("after", after)
    _set("beforeLines", bl)
    _set("afterLines", al)
    return changed


def format_caption_fonts(path, params: dict) -> dict:
    """图/表标题字体分流（投诉⑤：图标题错打表标题黑体规则）。

    模板-3 实测（run 级）：表标题 = 中文黑体小4 + 数字/字母 TNR 粗体小4；
    图标题 = 中文宋体小4 + 数字/字母 TNR 粗体小4（图标题中文不加黑体）。
    此前 journal 把 caption 统一打黑体 → 图标题错误。本 action 按段文本
    前缀分流修正中文 run 字体；数字/字母 run 加粗由 format_caption_numbers
    承担（防双重职责）。
    参数: {font_size_half_pt: 24}  小4=24 half-points
    """
    path = Path(path)
    params = params or {}
    cap = params.get("caption") or {}
    eq = params.get("equation") or {}
    root = _read_part(path, "word/document.xml")
    fig_fixed = 0
    tbl_fixed = 0
    split_n = 0
    if cap.get("bilingual"):
        split_n = _split_combined_caption_lines(root)
    if not cap and not eq:
        return {"changed": False, "figure_captions": 0, "table_captions": 0, "skipped": "no template caption"}
    cover_paras = _cover_paragraphs(root, _read_styles_map(path))   # 封面题注不动
    for para in root.iter(f"{W}p"):
        if cover_paras and any(para is c for c in cover_paras):
            continue
        if any(anc.tag == f"{W}tbl" for anc in para.iterancestors(f"{W}tbl")):
            continue  # 表内段（表内容另管）
        text = "".join(t.text or "" for t in para.iter(f"{W}t")).strip()
        if not text:
            continue
        is_fig = bool(re.match(r"^图\s*\d", text))
        is_tbl = bool(re.match(r"^表\s*\d", text))
        if is_fig:
            # 2026-08-23 用户拍板：图标题必须居中且无首行缩进（模板图1.2
            # 样本=center+无 ind；图2.2 是模板作者漏改的残留缩进）。
            # 产物 caption 规则合并样本时保留了 ind 2400 → center+缩进=
            # 视觉偏移「没居中」。此处图标题统一 jc=center + 清 ind。
            _ppr = para.find(f"{W}pPr")
            if _ppr is None:
                _ppr = etree.Element(f"{W}pPr")
                para.insert(0, _ppr)
            _jc = _ppr.find(f"{W}jc")
            if _jc is None:
                _jc = etree.Element(f"{W}jc")
                _ppr_insert_ordered(_ppr, _jc)
            if _jc.get(f"{W}val") != "center":
                _jc.set(f"{W}val", "center")
            _ind = _ppr.find(f"{W}ind")
            if _ind is not None:
                _ppr.remove(_ind)
        is_fig_en = bool(re.match(r"^Fig(?:ure)?\.?\s*\d", text, re.I))
        is_tbl_en = bool(re.match(r"^Table\s+\d", text, re.I))
        if not (is_fig or is_tbl or is_fig_en or is_tbl_en):
            continue
        if is_fig_en or is_tbl_en:
            if len(text) > 90:
                continue
        elif len(text) > 50 or re.search(r"[，。；、：]", text):
            continue
        if is_tbl_en:
            target_east = None
            sz_val = cap.get("table_en_sz")
            want_jc = cap.get("table_en_jc")
            ascii_font = cap.get("table_en_ascii")
        elif is_fig_en:
            target_east = None
            sz_val = cap.get("figure_en_sz")
            want_jc = cap.get("figure_en_jc")
            ascii_font = cap.get("figure_en_ascii")
        elif is_tbl:
            target_east = cap.get("table_east_asia")
            sz_val = cap.get("table_sz")
            want_jc = cap.get("table_jc")
            ascii_font = None
        else:
            target_east = cap.get("figure_east_asia")
            sz_val = cap.get("figure_sz")
            want_jc = cap.get("figure_jc")
            ascii_font = None
        if not target_east and not sz_val and not want_jc and not ascii_font:
            continue
        changed_para = False
        if want_jc:
            ppr = para.find(f"{W}pPr")
            if ppr is None:
                ppr = etree.Element(f"{W}pPr")
                para.insert(0, ppr)
            jc = ppr.find(f"{W}jc")
            if jc is None:
                jc = etree.Element(f"{W}jc")
                jc.set(f"{W}val", want_jc)
                ppr.append(jc)
                changed_para = True
            elif jc.get(f"{W}val") != want_jc:
                jc.set(f"{W}val", want_jc)
                changed_para = True
        kind = "table" if (is_tbl or is_tbl_en) else "figure"
        if is_tbl_en:
            kind = "table_en"
        elif is_fig_en:
            kind = "figure_en"
        if _write_caption_spacing(para, cap, kind):
            changed_para = True
        for run in para.findall(f"{W}r"):
            t = "".join(x.text or "" for x in run.iter(f"{W}t"))
            if not t.strip():
                continue
            rpr = run.find(f"{W}rPr")
            if rpr is None:
                rpr = etree.Element(f"{W}rPr")
                run.insert(0, rpr)
            fonts = rpr.find(f"{W}rFonts")
            if fonts is None:
                fonts = etree.SubElement(rpr, f"{W}rFonts")
                _rpr_insert_ordered(rpr, fonts)
            cur = fonts.get(f"{W}eastAsia")
            if target_east and cur != target_east:
                fonts.set(f"{W}eastAsia", target_east)
                changed_para = True
            if ascii_font:
                if fonts.get(f"{W}ascii") != ascii_font:
                    fonts.set(f"{W}ascii", ascii_font)
                    fonts.set(f"{W}hAnsi", ascii_font)
                    changed_para = True
            if sz_val:
                for tag in ("sz", "szCs"):
                    el = rpr.find(f"{W}{tag}")
                    if el is None:
                        el = etree.SubElement(rpr, f"{W}{tag}")
                        _rpr_insert_ordered(rpr, el)
                    if el.get(f"{W}val") != sz_val:
                        el.set(f"{W}val", sz_val)
                        changed_para = True
        if changed_para:
            if is_fig or is_fig_en:
                fig_fixed += 1
            else:
                tbl_fixed += 1
    eq_n = _apply_equation_jc(root, eq)
    if fig_fixed or tbl_fixed or split_n or eq_n:
        _write_part(path, "word/document.xml", root)
    return {"changed": (fig_fixed + tbl_fixed + split_n + eq_n) > 0,
            "figure_captions": fig_fixed, "table_captions": tbl_fixed,
            "caption_splits": split_n, "equations": eq_n}


def verify_caption_fonts(path, params: dict) -> bool:
    """复查：图标题中文 run=宋体、表标题中文 run=黑体（抽样任一违反即 False）。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        cover_paras = _cover_paragraphs(root, _read_styles_map(path))   # 与动作同条件
        for para in root.iter(f"{W}p"):
            if cover_paras and any(para is c for c in cover_paras):
                continue
            if any(anc.tag == f"{W}tbl" for anc in para.iterancestors(f"{W}tbl")):
                continue
            text = "".join(t.text or "" for t in para.iter(f"{W}t")).strip()
            if len(text) > 50 or re.search(r"[，。；、：]", text):
                continue
            if re.match(r"^Fig(?:ure)?\.?\s*\d", text, re.I) or re.match(r"^Table\s+\d", text, re.I):
                continue
            is_fig = bool(re.match(r"^图\s*\d", text))
            is_tbl = bool(re.match(r"^表\s*\d", text))
            if not (is_fig or is_tbl):
                continue
            cap = params.get("caption") or {}
            en_kind = "figure_en" if is_fig else "table_en"
            target = cap.get(en_kind + "_east_asia")
            if not target:
                kind = "figure" if is_fig else "table"
                target = cap.get(kind + "_east_asia") or (
                    # 无参数时的中性缺省：只挡「显式写错字体」，不发明规范值
                    "宋体" if is_fig else "黑体"
                )
            for run in para.findall(f"{W}r"):
                t = "".join(x.text or "" for x in run.iter(f"{W}t"))
                if not re.search(r"[\u4e00-\u9fff]", t):
                    continue
                rpr = run.find(f"{W}rPr")
                fonts = rpr.find(f"{W}rFonts") if rpr is not None else None
                east = fonts.get(f"{W}eastAsia") if fonts is not None else None
                if east is not None and east != target:
                    return False
        return True
    except Exception:
        return False


# ── OOXML 子元素顺序助手（CT_PPr / CT_TblPr / CT_RPr）──
_PPR_SEQ = PPR_SEQ  # 真源 earlyz/tools/ooxml_order.py（2026-09-28 收敛）
_TBLPR_SEQ = TBLPR_SEQ  # 真源 earlyz/tools/ooxml_order.py（2026-09-28 收敛）
_RPR_SEQ = RPR_SEQ  # 真源 earlyz/tools/ooxml_order.py（2026-09-28 收敛）


def _localname(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def _insert_by_seq(parent, el, seq: list[str]) -> None:
    """按 OOXML schema 序列把 el 插入 parent（el 无位则 append）。"""
    name = _localname(el.tag)
    try:
        idx = seq.index(name)
    except ValueError:
        parent.append(el)
        return
    for child in parent:
        cname = _localname(child.tag)
        if cname in seq and seq.index(cname) > idx:
            child.addprevious(el)
            return
    parent.append(el)


def _ppr_insert_ordered(ppr, el) -> None:
    _insert_by_seq(ppr, el, _PPR_SEQ)


def _tblpr_insert_ordered(tblpr, el) -> None:
    _insert_by_seq(tblpr, el, _TBLPR_SEQ)


def _rpr_insert_ordered(rpr, el) -> None:
    _insert_by_seq(rpr, el, _RPR_SEQ)


# ═══════════════════════════════════════════════════════════════
# 2026-08-17 用户彻查第三轮：①标题无 outlineLvl（Word 导航/目录识别
# 不到，截图实锤「大纲级别=正文文本」）②前置章节（摘要/ABSTRACT/目录）
# 与正文同节不另起页（目录和英文摘要同页截图）。
# ═══════════════════════════════════════════════════════════════

# 大纲级别：H1→0 / H2→1 / H3→2（Word 大纲/导航/TOC 域识别依据；
# 模板示例段无 outlineLvl（decompiler 提取不到）→ 内置语义规则）
_HEADING_LEVELS = {"heading1": "0", "heading2": "1", "heading3": "2"}

# 前置区章节词表：这些章节标题必须另起一页（与前一节同节时补分页）
_FRONT_CHAPTER_WORDS = {
    "摘要", "ABSTRACT", "目录", "结论",
    "原创性声明", "学位论文原创性声明", "版权使用授权书",
    "学位论文版权使用授权书", "开题报告", "任务书", "前言", "引言",
}
# 2026-10-06 取值出自模板铁律：「致谢/参考文献/附录」等文末章节不再强推分页
# （08-17 修「目录和英文摘要同页」时误并入本清单；CJC 期刊模板实证参考文献
# 紧跟正文排，模板全文无任何分页要求）。文末章节要不要另起页由模板说了算，
# 通过 ensure_section_page_breaks 的 params["section_break_words"] 下发。


def _para_full_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(f"{W}t")).strip()


def _norm_ws(s: str) -> str:
    """去全部空白（含全角空格）——章节标题匹配用（「摘  要」→「摘要」）。

    2026-10-08：□（U+25A1）也是占位空白——模板里常写「摘□□要」，此前不归一
    导致分页词表匹配失败（摘要不另起页）、同名标题规则生成不出来。
    """
    return re.sub(r"[\s\u3000\u25a1]+", "", s or "")


def _break_key(s: str) -> str:
    """分页词表匹配键：去空白/□ + 摘行尾括号注记 + 纯 ASCII 大写。

    2026-10-08 真机（用户报「英文摘要也没处理」直查）：标题「Abstract（Times
    New Roman 三号加粗，居中）」带行尾注记、且大小写与词表「ABSTRACT」不符，
    两道都断——Abstract 不另起页。注记剥离只用于**判定**，原文不删。
    """
    from earlyz.ir.analyzer import strip_trailing_annotation
    k = _norm_ws(strip_trailing_annotation(s or ""))
    return k.upper() if k.isascii() else k


def _break_cands(s: str) -> set[str]:
    """匹配候选集：键本身 + 附录带描述形态（「附录 A 公式(3-5) 的证明」）按类归「附录」。"""
    key = _break_key(s)
    cands = {key}
    if key.startswith("附录"):
        cands.add("附录")
    return cands


# 章级标题形态（「第 X 章/篇 …」，数字或中文数字）——断页按类判定用。
_CHAP_BREAK_RE = re.compile(r"^第\s*[\d一二三四五六七八九十]{1,3}\s*[章篇]")


def _is_chap_break_cand(p, text: str) -> bool:
    """「第 X 章 …」章标题段——模板声明分页要求时按类另起页（2026-10-08）。

    真机由来：模板断页词表是**字面标题**（「1□绪论」「结论」），目标文档章标题
    写作「第 1 章 引言」时对不上 → 六章全不另起页（Word 渲染实测章间连排）。
    模板本体证据：样张「1□绪论」「结论」段均带 pageBreakBefore = 章标题按类起页。
    判定=引擎既定章级形态口径（摘行尾注记 → 「第 X 章/篇」开头 → ≤40 字且无
    句读）；目录条目（PAGEREF 域 / 点线 leader）不参与。只判定，原文不动。
    """
    from earlyz.ir.analyzer import strip_trailing_annotation
    for it in p.iter(f"{W}instrText"):
        if "PAGEREF" in (it.text or ""):
            return False
    core = _norm_ws(strip_trailing_annotation(text or ""))
    if not core or len(core) > 40 or not _CHAP_BREAK_RE.match(core):
        return False
    if re.search(r"[。；，、！？：]", core) or re.search(r"[.．…]{3,}", core):
        return False
    return True


def _outline_level_for_role(role: str) -> str | None:
    """heading1/2/3 与后置章节 → Word outlineLvl 0/1/2。与分类同一真相源。

    参考文献 / 致谢 / 附录 是章级（一级）后置章节：目录里按一级条目列，Word 的
    TOC 域只收带大纲级别的段（目录级别 1-3）——不给它们 outlineLvl=0，用户按 F9
    更新域后这三条会整条消失（2026-11-08 用户实锤：目录更新后参考文献等章节没了）。
    """
    if role in ("heading1", "refs_heading", "ack_heading", "appendix_heading"):
        return "0"
    if role == "heading2":
        return "1"
    if role == "heading3":
        return "2"
    return None


def format_footnotes(path, params: dict) -> dict:
    """脚注/尾注段落按模板取值（2026-09-25 批3 覆盖扩展）。

    取值来自模板的脚注素材（structure["footnote"]：模板 word/footnotes.xml 里
    真脚注段的直接格式，或缺省 FootnoteText/EndnoteText 样式）。模板没有脚注
    素材 → 参数为空 → 直接跳过（不猜值、不造值，目标文档脚注保持原样；与
    「取值必须出自当前模板 OOXML」铁律一致——西安理工模板即无脚注素材）。
    w:type 为 separator/continuationSeparator 的分隔符段不动。
    """
    path = Path(path)
    cfg = (params or {}).get("footnote") or {}
    changed = 0
    touched: list[str] = []
    for kind in ("footnote", "endnote"):
        props = {k: v for k, v in (cfg.get(kind) or {}).items() if v is not None}
        if not props:
            continue
        part = f"word/{kind}s.xml"
        try:
            root = _read_part(path, part)
        except Exception:
            continue
        for note in root.iter(f"{W}{kind}"):
            if (note.get(f"{W}type") or "") in ("separator", "continuationSeparator"):
                continue
            for p in note.iter(f"{W}p"):
                _set_run_properties(p, props)
                if props.get("line") or props.get("line_rule"):
                    ppr = p.find(f"{W}pPr")
                    if ppr is None:
                        ppr = etree.Element(f"{W}pPr")
                        p.insert(0, ppr)
                    sp = ppr.find(f"{W}spacing")
                    if sp is None:
                        sp = etree.Element(f"{W}spacing")
                        _ppr_insert_ordered(ppr, sp)
                    if props.get("line"):
                        sp.set(f"{W}line", str(props["line"]))
                    if props.get("line_rule"):
                        sp.set(f"{W}lineRule", str(props["line_rule"]))
                changed += 1
        touched.append(part)
        _write_part(path, part, root)
    return {"changed": changed > 0, "paragraphs": changed, "parts": touched}


def verify_footnotes(path, params: dict) -> bool:
    """复查：脚注/尾注段落 rPr 与模板取值一致（无素材/无部件=通过）。"""
    try:
        path = Path(path)
        cfg = (params or {}).get("footnote") or {}
        for kind in ("footnote", "endnote"):
            props = {k: v for k, v in (cfg.get(kind) or {}).items() if v is not None}
            if not props:
                continue
            try:
                root = _read_part(path, f"word/{kind}s.xml")
            except Exception:
                continue
            want_sz = (str(int(round(float(props["font_size_pt"]) * 2)))
                       if props.get("font_size_pt") else None)
            for note in root.iter(f"{W}{kind}"):
                if (note.get(f"{W}type") or "") in ("separator", "continuationSeparator"):
                    continue
                for p in note.iter(f"{W}p"):
                    for r in p.iter(f"{W}r"):
                        rpr = r.find(f"{W}rPr")
                        if want_sz:
                            sz = rpr.find(f"{W}sz") if rpr is not None else None
                            if sz is None or sz.get(f"{W}val") != want_sz:
                                return False
                        if props.get("font_east_asia"):
                            rf = rpr.find(f"{W}rFonts") if rpr is not None else None
                            if rf is None or rf.get(f"{W}eastAsia") != props["font_east_asia"]:
                                return False
        return True
    except Exception:
        return False


def _width_of_sectpr(sp) -> int:
    """某个 sectPr 的版心宽（单栏）/ 栏宽（多栏），单位 twips。"""
    pgsz = sp.find(f"{W}pgSz")
    pgmar = sp.find(f"{W}pgMar")
    w = int(pgsz.get(f"{W}w")) if pgsz is not None else 12240
    l = int(pgmar.get(f"{W}left")) if pgmar is not None else 1800
    r = int(pgmar.get(f"{W}right")) if pgmar is not None else 1800
    cols = sp.find(f"{W}cols")
    if cols is not None and (cols.get(f"{W}num") or "1") != "1":
        els = list(cols.iter(f"{W}col"))
        cw = int(els[0].get(f"{W}w") or 0) if els else 0
        if cw > 0:
            return cw
        return (w - l - r) // max(1, int(cols.get(f"{W}num") or 2))
    return w - l - r


def _eq_number_tab_widths(root) -> list:
    """[(段落元素, 该段所属节的编号制表位宽度 twips)] —— 按 id 存会踩坑，见下。

    ⚠️ 真机坑（2026-09-28，双栏大论文）：
      ① `_page_content_width` 只取**第一个** sectPr → 前置单栏节的版心宽（417pt）。
         双栏节里正文在栏内，制表位要按**栏宽**（230pt）才落到位；用 417pt 的
         制表位，渲染实测编号停在栏中间。
      ② **lxml 元素是代理对象**：只存 `id(elem)` 不保留引用时，代理被回收后 id
         会变 → 查表全部落空（症状：编号制表位一个都没加，且没有任何报错）。
         所以这里返回 (元素, 宽度) 列表并保留引用，查表用 `is` 比身份。

    OOXML 语义：一节的段落排在**该节 sectPr 之前** → 段落归属 = 文档顺序中位于
    其后最近的 sectPr，没有则用 body 级最后一个。
    """
    elems = list(root.iter())
    sect_idx = [i for i, e in enumerate(elems) if e.tag == f"{W}sectPr"]
    if not sect_idx:
        return []
    body_sect = None
    for e in elems:
        if e.tag == f"{W}body":
            body_sect = e.find(f"{W}sectPr")
    body_w = _width_of_sectpr(body_sect) if body_sect is not None else 0
    out = []
    for i, e in enumerate(elems):
        if e.tag != f"{W}p":
            continue
        nxt = next((j for j in sect_idx if j > i), None)
        out.append((e, _width_of_sectpr(elems[nxt]) if nxt is not None else body_w))
    return out


def _equation_para_context(p, style_names, M):
    """判断段落是否公式段（OMML / OLE / 纯文本公式行），返回 (role, ppr)。

    ⚠️ 老版本在这里用 `if not (has_ole or has_math_para or has_math): continue`
    提前过滤 —— 用户真机产物 8 个公式段**全是纯文本**，于是永远进不来，公式看起来
    「都没处理」。现在补上纯文本公式行（`looks_like_plain_formula`）。
    """
    from earlyz.ir.analyzer import (classify_paragraph_role, _has_equation_ole,
                                    looks_like_plain_formula)
    if any(anc.tag == f"{W}tbl" for anc in p.iterancestors(f"{W}tbl")):
        return None, None
    if p.find(f"{W}hyperlink") is not None:
        return None, None
    has_ole = _has_equation_ole(p)
    has_math_para = p.find(f".//{M}oMathPara") is not None
    has_math = p.find(f".//{M}oMath") is not None
    text = _para_full_text(p)
    # 角色判定用的文本必须含**公式正文**（m:t）：`_para_full_text` 只收 w:t，
    # 「OMML 公式 + 行尾编号」的段落在动作侧看起来只是一行编号 → 被当成编号行、
    # 公式段整个漏掉（2026-09-28 实测：模板注入的公式样本 equations_seen=0）。
    # 与 analyzer 的 IR 文本口径对齐（math 在前、w:t 在后）。
    math_text = "".join(t.text or "" for t in p.iter(f"{M}t")).strip()
    role_text = f"{math_text}{text}".strip() if math_text else text
    if not (has_ole or has_math_para or has_math or looks_like_plain_formula(role_text)):
        return None, None
    ppr = p.find(f"{W}pPr")
    pstyle = ppr.find(f"{W}pStyle") if ppr is not None else None
    pstyle_id = pstyle.get(f"{W}val", "") if pstyle is not None else ""
    pstyle_name = style_names.get(pstyle_id, "") or pstyle_id
    numpr = ppr is not None and ppr.find(f"{W}numPr") is not None
    has_drawing = (has_ole or has_math_para
                   or p.find(f".//{W}drawing") is not None
                   or p.find(f".//{W}pict") is not None)
    role = classify_paragraph_role(
        ppr, pstyle_name, has_drawing, False, numpr, role_text,
        in_hyperlink=False, has_omath_para=has_math_para,
        has_omath=has_math, has_equation_ole=has_ole)
    return role, ppr


# 公式编号尾部识别（2026-10-08 修复）：原只认「(1)」单数字 → 中文论文常用章
# 编号「(3-1)」「(2.1)」漏判 → set_equation_dsmt4 转 DSMT4 时编号被当公式文本
# 一并删除（真机样本 7 条公式编号全丢、式(4-2) 交叉引用失去目标）。与
# EQ_NUM_ANY_RE 口径对齐：单数字 + 章编号（`-`/`.` 连接）。
_EQ_NUM_TAIL_RE = re.compile(
    r"^(?P<body>.*?)(?P<num>[（(]\s*\d{1,3}(?:\s*[.\-–]\s*\d{1,3})?\s*[)）])\s*$", re.S)


def normalize_equation_paragraphs(path, params: dict) -> dict:
    """公式段定型（2026-09-27 缩进归零 → 2026-09-28 夜班扩展）。

    用户报的缺陷：产物里 8 个公式段**全是纯文本**（无 OMML / 无 OLE），老判定只看
    OMML/OLE → 它们保持「正文」角色 → 继承正文首行缩进 2 字符 + 左右对齐 →
    用户看到「公式都没有处理」。

    本动作只做**段落级版式**（不生成编号、不改字符内容，编号右对齐只做定位）：
      ① 缩进归零（firstLine / firstLineChars / hanging → 0）；
      ② 居中 —— 模板无公式样本时的保守默认（模板有样本时由 spec 规则决定，
         本动作根本不跑）；
      ③ 清**固定行距**（exact / atLeast → 交回样式）——固定行高会裁切公式字形；
      ④ 西文/数字 Times New Roman、中文宋体（项目字体铁律）；
      ⑤ 行尾已有编号 `(N)` → 右对齐制表位（编号本来就在，只是挪到版心右缘）。

    边界：表格内跳过；不进 `m:oMath` / `w:object` 子树；幂等。
    params: center / fonts / number_tab / line_spacing（默认全开）。
    detect_only: 只统计（equations_seen / plain_seen）不改文档、不写盘——模板有公式
    样本（该步不抓版式）时编译链用它产生「纯文本公式」用户通报（2026-10-08）。
    """
    M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    style_names = _read_styles_map(path)
    do_center = params.get("center", True)
    do_fonts = params.get("fonts", True)
    do_tab = params.get("number_tab", True)
    do_spacing = params.get("line_spacing", True)
    do_fit = params.get("fit_width", True)
    detect_only = bool(params.get("detect_only"))
    fit_min = int(params.get("fit_min_half_pt", 15))     # 缩字号绝对下限：7.5 磅
    widths = _eq_number_tab_widths(root) if (do_tab and not detect_only) else []
    stat = {"equations_seen": 0, "plain_seen": 0, "indent_cleared": 0, "centered": 0,
            "line_spacing_unfixed": 0, "fonts_set": 0, "numbers_tabbed": 0,
            "tab_stops_fixed": 0, "lead_tabs": 0,
            "shrunk": 0, "shrink_details": [], "overflow_after_floor": 0,
            "ole_skipped": 0, "no_size_skipped": 0}

    # ⚠️ 必须先快照再改：`root.iter()` 是活迭代器，本循环里会插入 run/tab 节点，
    # 边改边遍历会让后续段落被跳过（真机实测：7 个公式只有 1 个编号被加制表位）。
    for p in list(root.iter(f"{W}p")):
        role, ppr = _equation_para_context(p, style_names, M)
        if role != "equation_display":
            continue
        stat["equations_seen"] += 1
        # 2026-10-06：区分「纯文本公式行」（内容是文字、不是公式对象）——B 通路
        # 的用户通报（「识别到 N 条纯文本公式，回复公式转对象可转换」）用。
        if not (p.find(f".//{M}oMath") is not None
                or p.find(f".//{W}object") is not None):
            stat["plain_seen"] += 1
        if detect_only:
            continue                      # 只统计：不改任何段落结构
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            p.insert(0, ppr)

        # ① 缩进归零
        ind = ppr.find(f"{W}ind")
        if ind is not None:
            touched = False
            for attr in ("firstLine", "firstLineChars", "hanging"):
                if ind.get(f"{W}{attr}") not in (None, "0"):
                    ind.set(f"{W}{attr}", "0")
                    touched = True
            if touched:
                stat["indent_cleared"] += 1

        # ② 居中（带行尾编号的段落交给 ⑤ 的**制表位模式**：制表符会把行撑满版心宽，
        #    此时 jc=center 无效、公式会落在栏左缘 —— 2026-09-28 渲染实测）
        if do_center:
            jc = ppr.find(f"{W}jc")
            if jc is None:
                jc = etree.Element(f"{W}jc")
                _ppr_insert_ordered(ppr, jc)
            if (jc.get(f"{W}val") != "center"
                    and not _EQ_NUM_TAIL_RE.match(_para_full_text(p))):
                jc.set(f"{W}val", "center")
                stat["centered"] += 1

        # ③ 固定行距 → 交回样式（公式的上下标会被固定行高裁掉）
        if do_spacing:
            sp = ppr.find(f"{W}spacing")
            if sp is not None and sp.get(f"{W}lineRule") in ("exact", "atLeast"):
                for attr in ("line", "lineRule"):
                    if sp.get(f"{W}{attr}") is not None:
                        del sp.attrib[f"{W}{attr}"]
                stat["line_spacing_unfixed"] += 1

        # ④ 字体：西文/数字 TNR、中文宋体（不进公式/OLE 子树）
        if do_fonts:
            for r in p.findall(f"{W}r"):
                if any(a.tag in (f"{M}oMath", f"{M}oMathPara")
                       or a.tag.endswith("object") or a.tag.endswith("OLEObject")
                       for a in r.iterancestors()):
                    continue
                if any(not c.tag.endswith("}rPr") and not c.tag.endswith("}t")
                       for c in list(r)):
                    continue
                rpr = r.find(f"{W}rPr")
                if rpr is None:
                    rpr = etree.Element(f"{W}rPr")
                    r.insert(0, rpr)
                rf = rpr.find(f"{W}rFonts")
                if rf is None:
                    rf = etree.Element(f"{W}rFonts")
                    _rpr_insert_ordered(rpr, rf)
                if (rf.get(f"{W}ascii") != "Times New Roman"
                        or rf.get(f"{W}hAnsi") != "Times New Roman"
                        or rf.get(f"{W}eastAsia") != "宋体"):
                    rf.set(f"{W}ascii", "Times New Roman")
                    rf.set(f"{W}hAnsi", "Times New Roman")
                    rf.set(f"{W}eastAsia", "宋体")
                    stat["fonts_set"] += 1

        # ⑤ 行尾编号 → **制表位模式**（居中制表位 + 右制表位；宽度按本段所属节取）
        # 真机坑（2026-09-28）：用 `jc=center` + 右制表位时，制表符把行撑满版心宽
        # → `jc=center` 失效，公式实际落在栏左缘（渲染实测 x0=56.8，栏中心应 103.6）
        # → 看起来「居中了」其实没有。Word 规范写法 = jc 归左 + 居中/右两个制表位。
        width = next((w for (pp, w) in widths if pp is p), 0)
        num_run = None
        if do_tab and width > 0:
            text, _mapping = _paragraph_text_map(p)
            mm = _EQ_NUM_TAIL_RE.match(text)
            if mm:
                num_start = mm.start("num")
                _split_run_at(p, num_start)
                _t2, mapping2 = _paragraph_text_map(p)
                num_run = next((r for (r, rs, _re_) in mapping2 if rs == num_start), None)
                # 段落里**已有**靠左的 right 制表位时「有就不加」会让编号停在旧位置
                # （实测 `(7)` 落在栏中间 187pt，而版心右缘 539pt）→ 一律重建制表位组。
                lay = _equation_tab_layout(ppr, p, num_run, width)
                if lay["tabs_changed"]:
                    stat["tab_stops_fixed"] += 1
                if lay["number_tab"]:
                    stat["numbers_tabbed"] += 1
                if lay["lead_tab"]:
                    stat["lead_tabs"] += 1
            elif do_center:
                # 没有编号的公式段：没有制表符撑行，`jc=center` 就是真居中
                jc = ppr.find(f"{W}jc")
                if jc is None:
                    jc = etree.Element(f"{W}jc")
                    _ppr_insert_ordered(ppr, jc)
                if jc.get(f"{W}val") != "center":
                    jc.set(f"{W}val", "center")
                    stat["centered"] += 1

        # ⑥ 超长公式缩字号（2026-09-28 用户拍板：「实在放不下只能缩小字体了，
        #    然后给用户说明情况」）——宽度不够才缩，缩完如实上报（进 validation.warnings）
        if do_fit:
            _fit_equation_width(p, ppr, width, num_run, stat, fit_min)

    # `changed` 只反映**真改动**：计数/记录类键不算（2026-09-28 真回归——
    # 新增的 ole_skipped 等计数把幂等断言打红：第二遍明明是 no-op 却报 changed=True）
    if detect_only:
        return {"changed": False, "detect_only": True,
                "equations_seen": stat["equations_seen"],
                "plain_seen": stat["plain_seen"]}
    changed = any(v for k, v in stat.items() if k not in _EQ_STAT_NOT_CHANGES)
    if changed:
        _write_part(path, "word/document.xml", root)
    out = {"changed": changed}
    out.update(stat)
    return out


# `normalize_equation_paragraphs` 里「计数/记录类」统计键 —— 它们不表示改动，
# 不得参与 `changed` 判定（否则幂等被打破）。
_EQ_STAT_NOT_CHANGES = frozenset({
    "equations_seen", "plain_seen", "shrink_details", "ole_skipped", "no_size_skipped",
    "overflow_after_floor",
})

# 公式编号（生成用）：比 `_EQ_NUM_TAIL_RE` 宽 —— 还认章编号 `(2-1)` / `(2.1)`，
# 因为生成时要能识别文档里已有的章编号形态并**接着编**（推断格式用）。
_EQ_NUM_ANY_RE = re.compile(
    r"^(?P<body>.*?)(?P<num>[（(]\s*(?P<val>\d{1,3}(?:\s*[.\-–]\s*\d{1,3})?)\s*[)）])\s*$",
    re.S)

# 章号汉字（"第二章" → 2）：模板/正文里章标题写法不止一种，都要认。
_CN_DIGITS = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7,
              "八": 8, "九": 9, "十": 10}


def _heading_chapter_number(p) -> str | None:
    """从 heading1 段落取章号（"第2章"/"2 系统设计"/"第二章" → "2"）；取不到 None。"""
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        return None
    lvl = ppr.find(f"{W}outlineLvl")
    style = ppr.find(f"{W}pStyle")
    is_h1 = (lvl is not None and lvl.get(f"{W}val") == "0")
    if not is_h1 and style is not None:
        sid = (style.get(f"{W}val") or "").lower()
        is_h1 = sid in ("heading1", "1", "标题1", "h1")
    if not is_h1:
        return None
    text = _para_full_text(p)
    m = re.search(r"第\s*([0-9]{1,2})\s*[章篇]", text) or re.search(r"第\s*([一二三四五六七八九十]+)\s*[章篇]", text)
    if m:
        tok = m.group(1)
        if tok.isdigit():
            return tok
        return str(_CN_DIGITS.get(tok, 0)) if tok in _CN_DIGITS else None
    m = re.match(r"\s*([0-9]{1,2})[\s、.．]", text)
    return m.group(1) if m else None


def _eq_number_style_from_spec(params: dict) -> dict:
    """编号文本的字体字号 —— 取**模板正文角色**（编号是正文级文字，非公式本体）。

    规范里 `equation_display` 明确**不取字体字号**（公式 run 是 OLE/OMML，run 字体
    对它无意义）；编号是纯文本 → 用 body 角色的中西文字体与字号。模板没写就一个
    都不设（继承样式，不发明）。
    """
    body = ((params or {}).get("spec") or {}).get("body") or {}
    out: dict = {}
    if body.get("font_ascii"):
        out["ascii"] = str(body["font_ascii"])
    if body.get("font_east_asia"):
        out["eastAsia"] = str(body["font_east_asia"])
    if body.get("font_size_pt"):
        out["sz"] = str(int(round(float(body["font_size_pt"]) * 2)))
    return out


def _append_equation_number(p, num_text: str, style: dict) -> bool:
    """把编号作为**最后一个 run** 追加到公式段（不动公式本体 / OLE 子树）。

    rPr 子元素顺序按 schema（rFonts 在前、sz/szCs 在后）。
    """
    run = etree.Element(f"{W}r")
    if style:
        rpr = etree.SubElement(run, f"{W}rPr")
        rf = etree.SubElement(rpr, f"{W}rFonts")
        if style.get("ascii"):
            rf.set(f"{W}ascii", style["ascii"])
            rf.set(f"{W}hAnsi", style["ascii"])
        if style.get("eastAsia"):
            rf.set(f"{W}eastAsia", style["eastAsia"])
        if style.get("sz"):
            for tag in ("sz", "szCs"):
                el = etree.SubElement(rpr, f"{W}{tag}")
                el.set(f"{W}val", style["sz"])
    t = etree.SubElement(run, f"{W}t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = num_text
    p.append(run)
    return True


def generate_equation_numbers(path, params: dict = None) -> dict:
    """公式编号生成（2026-09-28 用户拍板「公式编号的生成一定要做」）。

    只给**行间公式段**（equation_display）编号；行内公式不编号（默认 only_display）。
    编号格式（`format`）：
      · `"(N)"` 顺序编号 / `"(C-N)"` 章编号 —— 显式指定优先；
      · 未指定时从**文档已有编号**推断（出现 `(2-1)`/`(2.1)` 这类 → 用章编号）；
      · 文档里一个编号都没有 → 保守默认顺序编号（不发明章编号）。
    章号 C 取公式段之前最近的 heading1（"第2章"/"2 系统设计"/"第二章" → 2），
    取不到 → 退回顺序编号（不猜）。`paren` 未指定时按已有编号的半角/全角推断。

    编号文本字体字号取**模板正文角色**（`params["spec"]["body"]`），模板没写就不设。
    编号落位复用 ⑤ 的**制表位模式**（居中制表位 + 右制表位，宽度按本段所属节）。
    跳过已有编号的段落（skip_numbered 默认 True）→ 幂等。
    """
    M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
    path = Path(path)
    params = params or {}
    root = _read_part(path, "word/document.xml")
    style_names = _read_styles_map(path)
    widths = _eq_number_tab_widths(root)
    style = _eq_number_style_from_spec(params)
    start = int(params.get("start", 1))
    skip_numbered = bool(params.get("skip_numbered", True))
    fmt = params.get("format")
    paren = params.get("paren")

    paras = list(root.iter(f"{W}p"))
    eq_paras = []
    for p in paras:
        role, ppr = _equation_para_context(p, style_names, M)
        if role == "equation_display":
            eq_paras.append((p, ppr))

    existing = []
    for p, _ppr in eq_paras:
        mm = _EQ_NUM_ANY_RE.match(_para_full_text(p))
        if mm:
            existing.append(mm.group("num"))
    # 优先级：显式参数 > **模板样本**（spec.equation_display.number_format，用户拍板
    # 「模板里放公式样本」）> 文档已有编号推断 > 保守默认。
    spec_eq = ((params.get("spec") or {}).get("equation_display") or {})
    if fmt is None:
        fmt = spec_eq.get("number_format")
    if paren is None:
        paren = spec_eq.get("number_paren")
    if fmt is None:
        fmt = "(C-N)" if any(re.search(r"[.\-–]", v) for v in existing) else "(N)"
    if paren is None:
        paren = "（）" if any(n.startswith("（") for n in existing) else "()"

    # 已有编号 → 计数器起点，**避免撞号**（真机教训：文档已有 `(2-1)` 时若从 1 起，
    # 新编号会与已有编号重复）。`(N)` 取全局最大；`(C-N)` 按章取最大。
    seed: dict = {}
    for num in existing:
        inner = num.strip("()（）").replace(" ", "")
        m = re.match(r"^(\d{1,3})[.\-–](\d{1,3})$", inner)
        if m:
            seed[m.group(1)] = max(seed.get(m.group(1), 0), int(m.group(2)))
        elif inner.isdigit():
            seed[None] = max(seed.get(None, 0), int(inner))

    def _render(k: int, chapter: str | None) -> str:
        body = f"{chapter}-{k}" if (fmt == "(C-N)" and chapter) else str(k)
        return f"{paren[0]}{body}{paren[1]}"

    stat = {"equations_seen": len(eq_paras), "numbered": 0, "skipped_numbered": 0,
            "chapters": 0, "format": fmt, "paren": paren,
            "tab_stops_fixed": 0, "lead_tabs": 0, "number_tabs": 0,
            "no_chapter_fallback": 0}
    chapter: str | None = None
    counter = seed.get(None, start - 1)
    for p, ppr, chap in paras_and_ppr(root, style_names, M, eq_paras):
        if fmt == "(C-N)" and chap != chapter:      # 换章 → 章内计数器从该章已有最大值续起
            chapter = chap
            counter = seed.get(chap, seed.get(None, start - 1))
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            p.insert(0, ppr)
        if skip_numbered and _EQ_NUM_ANY_RE.match(_para_full_text(p)):
            stat["skipped_numbered"] += 1
            continue
        if fmt == "(C-N)" and chapter is None:
            stat["no_chapter_fallback"] += 1
        counter += 1
        num_text = _render(counter, chapter)
        _append_equation_number(p, num_text, style)
        stat["numbered"] += 1
        width = next((w for (pp, w) in widths if pp is p), 0)
        if width > 0:
            # 编号是刚追加的**最后一个 run**（独立 run，无需切分）
            num_run = p.findall(f"{W}r")[-1] if p.findall(f"{W}r") else None
            lay = _equation_tab_layout(ppr, p, num_run, width)
            if lay["tabs_changed"]:
                stat["tab_stops_fixed"] += 1
            if lay["number_tab"]:
                stat["number_tabs"] += 1
            if lay["lead_tab"]:
                stat["lead_tabs"] += 1

    changed = stat["numbered"] > 0
    if changed:
        _write_part(path, "word/document.xml", root)
    out = {"changed": changed}
    out.update(stat)
    return out


def paras_and_ppr(root, style_names, M, eq_paras):
    """按文档顺序产出 (公式段, ppr)，并在途中跟踪章号（heading1）。"""
    wanted = {id(p) for p, _ in eq_paras}
    chapter = None
    for p in list(root.iter(f"{W}p")):
        ch = _heading_chapter_number(p)
        if ch is not None:
            chapter = ch
        if id(p) in wanted:
            yield p, p.find(f"{W}pPr"), chapter


def verify_equation_numbers(path, params: dict = None) -> bool:
    """复查：行间公式段**每段都有**行尾编号（已有编号的段落也算合规）。"""
    try:
        M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        style_names = _read_styles_map(path)
        seen = 0
        for p in root.iter(f"{W}p"):
            role, _ppr = _equation_para_context(p, style_names, M)
            if role != "equation_display":
                continue
            seen += 1
            if not _EQ_NUM_ANY_RE.match(_para_full_text(p)):
                return False
        return True   # seen == 0（无行间公式）也算合规：没有可验证对象
    except Exception:
        return False


# ── MathType 章节断点（W3，对标 MathType「Insert Chapter/Section Break」）─────────
#
# 机制层（MathType 域协议，与学校无关）。MathType 编号段的域链是
#   MACROBUTTON MTPlaceRef → SEQ MTEqn \h（递增、隐藏）→ ( SEQ MTChap \c . SEQ MTEqn \c )
# 章号 MTChap 只被「读」（\c）；「写」它的是章节断点：
#   MACROBUTTON MTEditEquationSection2 Equation Chapter N Section 1
#   SEQ MTEqn \r \h · SEQ MTSec \r 1 \h · SEQ MTChap \r N \h
# 真论文 2026-09-30 实测：全文 8 处 `SEQ MTChap \c`、0 处断点 → Word 渲染时 MTChap=0，
# MTEqn 跨章连续 → Word 真机 PDF 显示 (0.1)…(0.8)，正文却写「公式(4.1)」。
# 修法 = 在每个「有编号公式、章号可读」的章里、第一处编号前插断点（隐藏）。
# 章号取自文档自己的章标题（分类层，`_heading1_state`）；编号格式（Arabic、分隔符、
# 括号）沿用文档自己的域链，不动。文档已有任何断点 → 视为用户/MathType 自管，不改。

_MT_SEQ_RE = re.compile(r"^\s*SEQ\s+(MT\w+)\b(.*)$", re.I | re.S)


def _heading1_state(p) -> tuple[bool, str | None]:
    """(是否一级标题, 章号)。判定口径与 `_heading_chapter_number` 相同。"""
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        return False, None
    lvl = ppr.find(f"{W}outlineLvl")
    style = ppr.find(f"{W}pStyle")
    is_h1 = (lvl is not None and lvl.get(f"{W}val") == "0")
    if not is_h1 and style is not None:
        is_h1 = (style.get(f"{W}val") or "").lower() in ("heading1", "1", "标题1", "h1")
    if not is_h1:
        return False, None
    return True, _heading_chapter_number(p)


def _mt_seq_fields(root) -> list[dict]:
    """按文档顺序列出 MathType 的 SEQ 域：{name, switches, para}。

    用 fldChar begin/separate/end 栈还原每个域**自己的**代码（嵌套域的代码不混进外层；
    separate 之后的结果区 instrText——MathType 嵌在 MACROBUTTON 里的缓存值——不算代码）。
    """
    out: list[dict] = []
    stack: list[dict] = []
    for el in root.iter(f"{W}fldChar", f"{W}instrText"):
        if el.tag == f"{W}fldChar":
            kind = el.get(f"{W}fldCharType")
            if kind == "begin":
                stack.append({"code": "", "result": False, "para": _ancestor_para(el)})
            elif kind == "separate" and stack:
                stack[-1]["result"] = True
            elif kind == "end" and stack:
                fld = stack.pop()
                m = _MT_SEQ_RE.match(fld["code"])
                if m:
                    out.append({"name": m.group(1), "switches": m.group(2), "para": fld["para"]})
        elif stack and not stack[-1]["result"]:
            stack[-1]["code"] += el.text or ""
    return out


def _ancestor_para(el):
    for anc in el.iterancestors(f"{W}p"):
        return anc
    return None


def _seq_value(switches: str, cur: int) -> tuple[int, bool]:
    """Word SEQ 语义：返回 (新计数, 是否显示)。\\c 读当前；\\r n 重置；否则递增。"""
    sw = switches or ""
    show = not re.search(r"\\h\b", sw)
    m = re.search(r"\\r\s*(\d+)?", sw)
    if m:
        return (int(m.group(1)) if m.group(1) else 0), show
    if re.search(r"\\c\b", sw):
        return cur, show
    return cur + 1, show


def _mt_number_audit(root) -> dict:
    """模拟 Word 对 MathType 编号的求值（W10 检查 / W4 的基础）。

    返回 {equations:[{para, heading_chap, shown_chap}], has_setter, chapters}。
    shown_chap = 该编号段里 `SEQ MTChap \\c` 求出的值。
    """
    seq = _mt_seq_fields(root)
    by_para: dict[int, list[dict]] = {}
    for f in seq:
        if f["para"] is not None:
            by_para.setdefault(id(f["para"]), []).append(f)
    has_setter = any(f["name"].lower() == "mtchap" and not re.search(r"\\c\b", f["switches"])
                     for f in seq)
    if any("MTEditEquationSection" in (it.text or "") for it in root.iter(f"{W}instrText")):
        has_setter = True
    counters: dict[str, int] = {}
    heading_chap: str | None = None
    in_chapter = False
    equations = []
    for p in root.iter(f"{W}p"):
        is_h1, num = _heading1_state(p)
        if is_h1 and _para_full_text(p).strip():
            heading_chap, in_chapter = num, True
        shown_chap = None
        refs_chap = False
        for f in by_para.get(id(p), []):
            name = f["name"].lower()
            val, _show = _seq_value(f["switches"], counters.get(name, 0))
            counters[name] = val
            if name == "mtchap" and re.search(r"\\c\b", f["switches"]):
                shown_chap, refs_chap = val, True
        if refs_chap:
            equations.append({"para": p, "heading_chap": heading_chap if in_chapter else None,
                              "shown_chap": shown_chap})
    return {"equations": equations, "has_setter": has_setter}


def _mt_break_runs(chap: str, rstyle: str | None) -> list:
    """MathType 章节断点的 run 序列（隐藏：run 级 vanish——文档里 MTEquationSection
    样式可能被用户切成「显示」，真论文就是 vanish=0 + 红色，不能依赖样式隐藏）。"""
    def run(child):
        r = etree.Element(f"{W}r")
        rpr = etree.SubElement(r, f"{W}rPr")
        if rstyle:
            etree.SubElement(rpr, f"{W}rStyle").set(f"{W}val", rstyle)
        etree.SubElement(rpr, f"{W}vanish")
        r.append(child)
        return r

    def fld(kind):
        el = etree.Element(f"{W}fldChar")
        el.set(f"{W}fldCharType", kind)
        return run(el)

    def instr(text):
        el = etree.Element(f"{W}instrText")
        el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        el.text = text
        return run(el)

    runs = [fld("begin"), instr(" MACROBUTTON MTEditEquationSection2 "),
            instr(f"Equation Chapter {chap} Section 1"), fld("end")]
    for code in (" SEQ MTEqn \\r \\h \\* MERGEFORMAT ", " SEQ MTSec \\r 1 \\h \\* MERGEFORMAT ",
                 f" SEQ MTChap \\r {chap} \\h \\* MERGEFORMAT "):
        runs += [fld("begin"), instr(code), fld("end")]
    return runs


def _para_child_of(p, el):
    """el 在段落 p 里对应的直接子元素（run 可能包在 hyperlink/ins 里）。"""
    cur = el
    while cur is not None and cur.getparent() is not p:
        cur = cur.getparent()
    return cur


def insert_equation_chapter_breaks(path, params: dict = None) -> dict:
    """W3：给「章号读不出来」的 MathType 编号补章节断点（插在每章第一处编号之前）。

    只在文档**没有任何**章节断点、且模拟求值显示章号 ≠ 章标题号时动手；
    章标题没有可读章号（摘要/附录等）的章跳过并计数，不猜。幂等：插完后
    模拟求值一致 → 再跑不改。
    """
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    audit = _mt_number_audit(root)
    eqs = audit["equations"]
    stat = {"changed": False, "mt_numbered": len(eqs), "breaks_inserted": 0,
            "chapters": [], "skipped_no_chapter": 0, "user_managed": audit["has_setter"],
            "mismatch": sum(1 for e in eqs if e["heading_chap"] is not None
                            and str(e["shown_chap"]) != e["heading_chap"])}
    if not eqs or stat["mismatch"] == 0:
        return stat
    if audit["has_setter"]:
        # 用户/MathType 自己插过断点但对不上——不替用户改章号，交给建议层列出
        return stat
    styles_xml = ""
    try:
        with zipfile.ZipFile(path) as z:
            styles_xml = z.read("word/styles.xml").decode("utf-8", "ignore")
    except (KeyError, zipfile.BadZipFile):
        pass
    rstyle = "MTEquationSection" if 'w:styleId="MTEquationSection"' in styles_xml else None
    done: set[str] = set()
    for e in eqs:
        chap = e["heading_chap"]
        if chap is None:
            stat["skipped_no_chapter"] += 1
            continue
        if chap in done:
            continue
        p = e["para"]
        begin = next((fc for fc in p.iter(f"{W}fldChar")
                      if fc.get(f"{W}fldCharType") == "begin"), None)
        anchor = _para_child_of(p, begin) if begin is not None else None
        if anchor is None:
            continue
        for r in _mt_break_runs(chap, rstyle):
            anchor.addprevious(r)
        done.add(chap)
        stat["breaks_inserted"] += 1
        stat["chapters"].append(chap)
    if stat["breaks_inserted"]:
        _write_part(path, "word/document.xml", root)
        stat["changed"] = True
    return stat


def verify_equation_chapter_breaks(path, params: dict = None) -> bool:
    """复查本动作的契约：有章号的章里模拟章号 = 章标题号；或文档带断点（用户/MathType
    自管，含本动作插过的）。只有「无任何断点却对不上」= 该插没插 → 失败。
    用户自插断点且对不上 → handler 按设计不动，这里必须放行（否则整笔事务被回滚）。"""
    try:
        audit = _mt_number_audit(_read_part(Path(path), "word/document.xml"))
        if audit["has_setter"]:
            return True
        return all(str(e["shown_chap"]) == e["heading_chap"]
                   for e in audit["equations"] if e["heading_chap"] is not None)
    except Exception:
        return False


_CJK_CHAR_RE = re.compile(r"[\u4e00-\u9fff]")
_MATH_MARK_RE = re.compile(r"[=^_/∑Σ∂≤≥±×·√∫αβγθλμσφωΔ]")

# 公式「放得下」的判定余量：判定与缩放**必须用同一个值**。
# 若缩放按 1.0×可用宽、判定按 0.98×可用宽，缩完仍判不过 → 每遍再缩一点 = **棘轮式递减**
# （2026-09-28 真论文实测：第二遍编译仍 shrunk=2，幂等被打破）。
_EQ_FIT_MARGIN = 0.97

# 公式宽度估算用的字符宽度表（单位 em，相对字号）——够不够放得下用，不求精确
_CHAR_EM = {" ": 0.25, "/": 0.35, "(": 0.35, ")": 0.35, "[": 0.35, "]": 0.35,
            ",": 0.25, ";": 0.25, ":": 0.25, ".": 0.25, "'": 0.2, "\"": 0.3,
            "=": 0.6, "+": 0.6, "-": 0.35, "−": 0.6, "<": 0.6, ">": 0.6,
            "≤": 0.6, "≥": 0.6, "±": 0.6, "×": 0.6, "·": 0.35, "√": 0.8,
            "∑": 1.2, "Σ": 1.2, "∫": 1.0, "∂": 0.55, "…": 0.9}
_OMATH_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


def _char_em(ch: str) -> float:
    if _CJK_CHAR_RE.match(ch):
        return 1.0
    if ch in _CHAR_EM:
        return _CHAR_EM[ch]
    if ch.isdigit():
        return 0.5
    if "a" <= ch.lower() <= "z":
        return 0.52
    if ch in "αβγδθλμσφωΔ":
        return 0.58
    return 0.5


def _text_width_em(text: str) -> float:
    """纯文本公式宽度（em）。`_x`/`^x` 按 0.7 倍算（上下标小一号）。"""
    total = 0.0
    i = 0
    while i < len(text):
        ch = text[i]
        if ch in "_^" and i + 1 < len(text):
            total += 0.7 * _char_em(text[i + 1])
            i += 2
            continue
        total += _char_em(ch)
        i += 1
    return total


def _math_width_em(node) -> float:
    """估 OMML 公式宽度（em）：分式取分子分母最大值，根式/括号加余量。"""
    M = _OMATH_NS
    tag = node.tag
    if tag == f"{M}r":
        return _text_width_em("".join(t.text or "" for t in node.iter(f"{M}t")))
    if tag == f"{M}f":
        num = node.find(f"{M}num")
        den = node.find(f"{M}den")
        return max(_math_width_em(num) if num is not None else 0.0,
                   _math_width_em(den) if den is not None else 0.0) + 0.15
    if tag in (f"{M}sSub", f"{M}sSup"):
        base = node.find(f"{M}e")
        sub = node.find(f"{M}sub") if tag == f"{M}sSub" else node.find(f"{M}sup")
        return ((_math_width_em(base) if base is not None else 0.0)
                + 0.72 * (_math_width_em(sub) if sub is not None else 0.0))
    if tag == f"{M}sSubSup":
        base = node.find(f"{M}e")
        return (_math_width_em(base) if base is not None else 0.0) + 0.72 * 2
    if tag == f"{M}nary":
        sub = node.find(f"{M}sub")
        sup = node.find(f"{M}sup")
        e = node.find(f"{M}e")
        return (1.3 + 0.72 * ((_math_width_em(sub) if sub is not None else 0.0)
                              + (_math_width_em(sup) if sup is not None else 0.0))
                + (_math_width_em(e) if e is not None else 0.0))
    if tag == f"{M}rad":
        e = node.find(f"{M}e")
        return (_math_width_em(e) if e is not None else 0.0) + 1.0
    if tag == f"{M}d":
        e = node.find(f"{M}e")
        return (_math_width_em(e) if e is not None else 0.0) + 0.8
    if tag == f"{M}m":
        rows = node.findall(f"{M}mr")
        if not rows:
            return 0.0
        widest = 0.0
        ncols = 0
        for mr in rows:
            cells = mr.findall(f"{M}e")
            ncols = max(ncols, len(cells))
            widest = max(widest, sum(_math_width_em(c) for c in cells))
        return widest + 0.4 * ncols
    if tag in (f"{M}func", f"{M}limLow", f"{M}limUpp", f"{M}groupChr", f"{M}acc",
               f"{M}bar", f"{M}box", f"{M}borderBox"):
        return sum(_math_width_em(c) for c in node) + 0.3
    return sum(_math_width_em(c) for c in node)


def _equation_width_em(p) -> float:
    """公式段宽度（em）：OMML 走数学树估算，纯文本走字符表；OLE 取不到 → 0。"""
    omml = p.find(f".//{_OMATH_NS}oMath")
    if omml is not None:
        return _math_width_em(omml)
    if p.find(f".//{W}object") is not None:
        return 0.0                      # MathType OLE：内部尺寸取不到，不猜
    return _text_width_em(_para_full_text(p))


def _equation_font_size(p) -> int | None:
    """公式段主导字号（半磅）；OMML run / 普通 run / 都取不到 → None（不发明）。"""
    sizes = []
    for rpr in p.iter(f"{W}rPr"):
        sz = rpr.find(f"{W}sz")
        if sz is not None and (sz.get(f"{W}val") or "").isdigit():
            sizes.append(int(sz.get(f"{W}val")))
    if not sizes:
        return None
    return max(set(sizes), key=sizes.count)


def _apply_equation_size(p, half_points: int) -> int:
    """把公式段内所有字号（含 OMML run）改成 half_points；返回改动数。

    ⚠️ 铁律例外：**只在公式实在放不下时**才进 `m:oMath` 子树改字号（用户 2026-09-28
    拍板「如果实在放不下只能缩小字体了」）——其余场合禁进数学子树。
    """
    n = 0
    for rpr in p.iter(f"{W}rPr"):
        for tag in ("sz", "szCs"):
            el = rpr.find(f"{W}{tag}")
            if el is None:
                continue
            if el.get(f"{W}val") != str(half_points):
                el.set(f"{W}val", str(half_points))
                n += 1
    return n


def _fit_equation_width(p, ppr, width: int, num_run, stat: dict,
                        min_half_pt: int) -> None:
    """超长公式缩字号（用户拍板：实在放不下就缩，并**给用户说明情况**）。

    估算公式宽度（em）× 字号（twips）与可用宽度比较；超出则按比例缩字号。
    下限用**绝对半磅值**（默认 7.5 磅）而不是比例 —— 比例下限会让第二遍继续缩
    （不幂等），绝对值保证「已在下限就不再动」。
    编号占位（编号 + 间隙）从可用宽度里扣掉。
    """
    if width <= 0:
        return
    em = _equation_width_em(p)
    if em <= 0:
        stat["ole_skipped"] += 1            # OLE 内部尺寸取不到 → 不猜、不动
        return
    size = _equation_font_size(p)
    if size is None:
        stat["no_size_skipped"] += 1        # 模板/文档没给字号 → 不发明
        return
    reserve = 0
    if num_run is not None:
        num_text = "".join(t.text or "" for t in num_run.iter(f"{W}t"))
        reserve = int(_text_width_em(num_text) * size * 10) + 200   # 编号 + 间隙
    avail = width - reserve
    need = int(em * size * 10)
    if need <= avail * _EQ_FIT_MARGIN:
        return
    want = max(int(min_half_pt), 1)
    scaled = int(size * avail * _EQ_FIT_MARGIN / need) if need else size
    new_size = max(want, scaled)
    if new_size >= size:
        stat["overflow_after_floor"] += 1
        stat["shrink_details"].append({
            "text": _para_full_text(p)[:40], "from_half_pt": size,
            "to_half_pt": new_size, "em": round(em, 2),
            "avail_twips": avail, "need_twips": need, "fits": False})
        return
    _apply_equation_size(p, new_size)
    stat["shrunk"] += 1
    stat["shrink_details"].append({
        "text": _para_full_text(p)[:40], "from_half_pt": size,
        "to_half_pt": new_size, "em": round(em, 2),
        "avail_twips": avail, "need_twips": need, "fits": True})


def _inline_formula_span(text: str) -> dict | None:
    """在一段文字里找「行内公式 + 行尾编号」的边界。

    返回 {prose, formula, num, punct, split_at} 或 None（不该拆/拆不动）。

    真机用例（源稿 p#38）：
      `其中，D_n 为任务数据量，…，f_e 为边缘服务器算力。本地执行时延为 T_n^loc = C_n / f_n^loc。 (4)`
    → prose = `…本地执行时延为`，formula = `T_n^loc = C_n / f_n^loc`，num = `(4)`，punct = `。`

    保守原则（宁可漏认，不可误认）：必须**行尾有编号** + 公式段**含等号/数学记号** +
    正文部分有足够汉字；`本文的贡献在于：(1) 建立了…` 这类正文列举（无等号）不会命中。
    """
    m = _EQ_NUM_ANY_RE.match(text)
    if not m:
        return None
    body = text[:m.start("num")]
    cjk = [i for i, ch in enumerate(body) if _CJK_CHAR_RE.match(ch)]
    if len(cjk) < 6:                      # 正文太短 → 不像「正文里粘公式」
        return None
    fs = cjk[-1] + 1                      # 最后一个汉字之后 = 公式候选起点
    formula_raw = body[fs:]
    formula = formula_raw.strip().lstrip("，,、；;：: 　")
    if len(formula) < 3 or "=" not in formula:
        return None
    if _CJK_CHAR_RE.search(formula):      # 公式里还夹汉字 → 拿不准，不拆
        return None
    if not _MATH_MARK_RE.search(formula):
        return None
    punct = ""
    if formula and formula[-1] in "。，；：":
        punct = formula[-1]
        formula = formula[:-1].rstrip()
    if len(formula) < 3:
        return None
    return {"prose": body[:fs], "formula": formula, "num": m.group("num"),
            "punct": punct, "split_at": fs}


def split_inline_equations(path, params: dict = None) -> dict:
    """行内公式段拆分（2026-09-28 用户拍板「要加」）。

    源稿把公式**粘在正文句子里**（真机用例 p#38）：
      `…f_e 为边缘服务器算力。本地执行时延为 T_n^loc = C_n / f_n^loc。 (4)`
    拆成两段：
      P1 正文：`…本地执行时延为`
      P2 公式段：`T_n^loc = C_n / f_n^loc` + 编号 `(4)`（居中 + 编号右对齐制表位）

    只搬 run，**不改文字内容**；公式段尾的中文句号默认移出（`drop_trailing_punct`，
    期刊惯例：行间公式末尾不加句号），并在 stats 里如实计数。
    幂等：拆完后 P1 无编号、P2 是行间公式 → 再跑零命中。
    params: drop_trailing_punct(默认 True) / min_cjk(默认 6)。
    """
    M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
    path = Path(path)
    params = params or {}
    drop_punct = bool(params.get("drop_trailing_punct", True))
    root = _read_part(path, "word/document.xml")
    style_names = _read_styles_map(path)
    widths = _eq_number_tab_widths(root)
    stat = {"scanned": 0, "split": 0, "dropped_punct": 0, "skipped_role": 0}

    for p in list(root.iter(f"{W}p")):
        if p.getparent() is not None and p.getparent().tag.endswith("}tc"):
            continue                                   # 表格内不拆
        role, _ppr = _equation_para_context(p, style_names, M)
        if role == "equation_display":
            stat["skipped_role"] += 1
            continue
        text = _para_full_text(p)
        if not text or len(text) > 400:
            continue
        stat["scanned"] += 1
        info = _inline_formula_span(text)
        if info is None:
            continue

        # 在 split_at 处切开 run（保证有 run 边界），再把公式起的 run 搬进新段
        _split_run_at(p, info["split_at"])
        _t, mapping = _paragraph_text_map(p)
        cut = info["split_at"]
        move = [r for (r, rs, _re_) in mapping if rs >= cut]
        # 句号：公式末尾的标点若留在 run 里，按 drop_trailing_punct 处理
        tail = "".join(t.text or "" for t in p.iter(f"{W}t"))
        newp = etree.Element(f"{W}p")
        ppr_src = p.find(f"{W}pPr")
        if ppr_src is not None:
            newp.append(copy.deepcopy(ppr_src))
        for r in move:
            p.remove(r)
            newp.append(r)
        if drop_punct and info["punct"]:
            # 句号在**编号前**（`…f_n^loc。 (4)`），不在段尾 → 找「标点 + 行尾编号」的 w:t
            # 直接剥掉这一个字符（不依赖 run 边界，避免切分偏移差一位）
            pat_tail = re.compile(
                rf"{re.escape(info['punct'])}(\s*[（(][^）)]*[)）]\s*)$")
            for t in newp.iter(f"{W}t"):
                if t.text and pat_tail.search(t.text):
                    t.text = pat_tail.sub(r"\1", t.text)
                    stat["dropped_punct"] += 1
                    break
        # 公式段版式：缩进归零 + 制表位模式（编号右对齐）
        nppr = newp.find(f"{W}pPr")
        if nppr is None:
            nppr = etree.Element(f"{W}pPr")
            newp.insert(0, nppr)
        for ind in nppr.findall(f"{W}ind"):
            nppr.remove(ind)
        p.addnext(newp)
        num_run = newp.findall(f"{W}r")[-1] if newp.findall(f"{W}r") else None
        width = next((w for (pp, w) in widths if pp is p), 0)
        if width > 0 and num_run is not None:
            _equation_tab_layout(nppr, newp, num_run, width)
        stat["split"] += 1

    changed = stat["split"] > 0
    if changed:
        _write_part(path, "word/document.xml", root)
    out = {"changed": changed}
    out.update(stat)
    return out


def verify_split_inline_equations(path, params: dict = None) -> bool:
    """复查：不存在「正文 + 行内公式 + 行尾编号」的段落（都拆干净了）。"""
    try:
        M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        style_names = _read_styles_map(path)
        for p in root.iter(f"{W}p"):
            if p.getparent() is not None and p.getparent().tag.endswith("}tc"):
                continue
            role, _ppr = _equation_para_context(p, style_names, M)
            if role == "equation_display":
                continue
            text = _para_full_text(p)
            if text and len(text) <= 400 and _inline_formula_span(text) is not None:
                return False
        return True
    except Exception:
        return False


def _equation_tab_layout(ppr, p, num_run, width: int) -> dict:
    """把「公式居中 + 编号右对齐」做成 **制表位模式**（Word 规范写法）。

    为什么不用 `jc=center` + 右制表位：制表符会把行撑满版心宽，`jc=center`
    随之失效 —— 渲染实测公式落在栏左缘（x0=56.8，而栏中心应为 103.6），
    看起来「居中了」其实没有。规范做法是段落 `jc` 归左，靠两个制表位：
    居中制表位（版心中心）+ 右制表位（版心右缘），行首 `<w:tab/>` 把公式推到
    居中位，编号前 `<w:tab/>` 把它顶到右缘。

    返回 {center_tab, right_tab, lead_tab, number_tab}。
    """
    out = {"center_tab": width // 2, "right_tab": width,
           "lead_tab": False, "number_tab": False, "tabs_changed": False}
    if width <= 0:
        return out
    want = [("center", str(width // 2)), ("right", str(width))]
    tabs = ppr.find(f"{W}tabs")
    have = ([(t.get(f"{W}val"), t.get(f"{W}pos")) for t in tabs]
            if tabs is not None else [])
    if have != want:
        out["tabs_changed"] = True
        if tabs is None:
            tabs = etree.Element(f"{W}tabs")
            _ppr_insert_ordered(ppr, tabs)
        for x in list(tabs):
            tabs.remove(x)
        for val, pos in want:
            tb = etree.SubElement(tabs, f"{W}tab")
            tb.set(f"{W}val", val)
            tb.set(f"{W}pos", pos)
    jc = ppr.find(f"{W}jc")
    if jc is not None and jc.get(f"{W}val") != "left":
        # 对齐交给制表位：`center`/`both` 在有制表符撑行时都不成立（公式会偏到栏左缘）
        jc.set(f"{W}val", "left")
    first = p.find(f"{W}r")
    if first is not None:
        prev = first.getprevious()
        # 行首制表符：看**第一个 run 自己**是不是纯制表符 run（第二次调用时它就是
        # 上次插的那个 —— 只看 prev 会漏判，实测幂等被打破：lead_tabs 每次 +2）
        first_is_tab = (first.find(f"{W}tab") is not None
                        and len([c for c in first if not c.tag.endswith("}rPr")]) == 1)
        prev_is_tab = (prev is not None and prev.tag == f"{W}r"
                       and prev.find(f"{W}tab") is not None)
        if not (first_is_tab or prev_is_tab):
            lead = etree.Element(f"{W}r")
            etree.SubElement(lead, f"{W}tab")
            first.addprevious(lead)
            out["lead_tab"] = True
    if num_run is not None:
        prev = num_run.getprevious()
        if not (prev is not None and prev.tag == f"{W}r"
                and prev.find(f"{W}tab") is not None):
            trun = etree.Element(f"{W}r")
            etree.SubElement(trun, f"{W}tab")
            num_run.addprevious(trun)
            out["number_tab"] = True
    return out


def _ensure_number_tab(ppr, p, num_run, width: int) -> int:
    """给「行尾编号」加右对齐制表位 + 制表符 run，返回新增的制表符数（0/1）。

    宽度按**本段所属节**取（双栏 = 栏宽，见 `_eq_number_tab_widths`）。
    """
    if num_run is None or width <= 0:
        return 0
    tabs = ppr.find(f"{W}tabs")
    if tabs is None:
        tabs = etree.Element(f"{W}tabs")
        _ppr_insert_ordered(ppr, tabs)
    right_tabs = [x for x in tabs if x.get(f"{W}val") == "right"]
    if right_tabs:
        if any(x.get(f"{W}pos") != str(width) for x in right_tabs):
            right_tabs[0].set(f"{W}pos", str(width))
            for extra in right_tabs[1:]:
                tabs.remove(extra)
    else:
        tb = etree.SubElement(tabs, f"{W}tab")
        tb.set(f"{W}val", "right")
        tb.set(f"{W}pos", str(width))
    prev = num_run.getprevious()
    if prev is not None and prev.tag == f"{W}r" and prev.find(f"{W}tab") is not None:
        return 0
    tab_run = etree.Element(f"{W}r")
    etree.SubElement(tab_run, f"{W}tab")
    num_run.addprevious(tab_run)
    return 1


def set_equation_from_latex(path, params: dict) -> dict:
    """把指定段落的公式改写为 **LaTeX 编译出的 Word 原生公式（OMML）**。

    用户拍板的能力（2026-09-27 夜）：「编辑公式我们得做，而且是拉开和其他排版
    工具差距的最有力的能力」——走「显式建议 + 采纳」，不静默转换。

    边界（三条铁律）：
      · **保留原编号**：段落行尾的 `(N)` 原样留着（本动作只格式化、不生成编号）；
      · 不支持的 LaTeX 构造**显式失败**（`LatexUnsupported`），绝不降级成纯文本；
      · 幂等：当前公式的 LaTeX 与目标一致时不动（比对用已有 `omath_latex_in`）。

    params:
      selector   — 定位段落的稳定选择器（`{semantic_role, occurrence, node_id, text_hash}`）
      latex      — 目标 LaTeX
      sizes      — 可选尺寸覆盖（默认模板第 14 条：21/12/9/32/21 半磅）
      number_tab — 是否把编号用制表位顶到版心右缘（默认 True）
    """
    from earlyz.tools.latex_omml import latex_to_omml
    from earlyz.tools.omml_latex import omath_latex_in
    from earlyz.ir.selectors import normalize_text

    latex = (params.get("latex") or "").strip()
    if not latex:
        raise ValueError("set_equation_from_latex: 缺少 latex 参数")
    selector = dict(params.get("selector") or {})
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    styles = _read_styles_map(path)
    # match_text：按公式文本定位（LLM 手里就有文本，比序号稳）——命中须恰好 1 段
    match_text = params.get("match_text") or selector.pop("match_text", None)
    if match_text:
        pool = _locate_paragraphs(root, selector or {"semantic_role": "equation_display"},
                                  styles)
        needle = str(match_text).strip()
        targets = [q for q in pool if needle in _para_full_text(q)]
        if not targets:
            targets = [q for q in pool if needle[:24] and needle[:24] in _para_full_text(q)]
        if not targets:
            # 兜底：**转换之后段落文本就没了**（变成 OMML，没有 w:t）→ 纯文本
            # match_text 再也匹配不上。此时按「当前公式的 LaTeX 已等于目标」认账，
            # 视为幂等命中（否则第二次调用会误报「选择器命中 0 段」）。
            from earlyz.tools.omml_latex import omath_latex_in as _o2l, omml_to_latex as _o2lat
            from earlyz.tools.latex_omml import latex_to_omml as _l2o
            # 苹果对苹果：目标 LaTeX 先走一遍**同一管线**（LaTeX→OMML→LaTeX）再比对，
            # 否则 `E_n^{off}` 与反编译出来的 `{E}_{n}^{off}` 永远不相等。
            want = (_o2lat(_l2o(latex)) or "").strip()
            for q in _locate_paragraphs(root, {"semantic_role": "equation_display"}, styles):
                cur = (_o2l(q) or "").strip()
                if cur and want and cur == want:
                    return {"changed": False, "latex": latex, "reason": "already-equal",
                            "matched_by": "latex"}
    else:
        targets = _locate_paragraphs(root, selector, styles)
    if len(targets) != 1:
        raise ValueError(f"set_equation_from_latex: 选择器命中 {len(targets)} 段（须恰好 1 段）")
    p = targets[0]

    # 幂等：已经是同一条公式就不动（同样走管线比对）
    _cur = (omath_latex_in(p) or "").strip()
    if _cur:
        from earlyz.tools.omml_latex import omml_to_latex as _o2lat2
        from earlyz.tools.latex_omml import latex_to_omml as _l2o2
        if _cur == (_o2lat2(_l2o2(latex)) or "").strip():
            return {"changed": False, "latex": latex, "reason": "already-equal"}

    text, _mapping = _paragraph_text_map(p)
    mm = _EQ_NUM_TAIL_RE.match(text)
    number = mm.group("num").strip() if mm else None
    num_start = mm.start("num") if mm else None
    if num_start is not None:
        _split_run_at(p, num_start)
    _t2, mapping2 = _paragraph_text_map(p)
    num_run = next((r for (r, rs, _re_) in mapping2 if rs == num_start), None) \
        if num_start is not None else None

    omath = latex_to_omml(latex, params.get("sizes"))
    # 用 m:oMathPara 包一层：Word 对**行间公式**的标准写法（也让本项目的
    # analyzer 认得出 equation_display —— 它只在「OMML 且段落无文本」时判公式，
    # 而我们的段落里还留着编号 run）。
    omml_block = etree.Element(f"{M}oMathPara")
    omml_block.append(omath)

    # 清掉旧的可写文本 run（保留编号 run 与含图片/域的非纯文本 run）
    for r in list(p.findall(f"{W}r")):
        if num_run is not None and r is num_run:
            continue
        if any(not c.tag.endswith("}rPr") and not c.tag.endswith("}t") for c in list(r)):
            continue          # 非纯文本 run（图片/域/脚注）不动
        if any(a.tag.endswith("oMath") for a in r.iterancestors()):
            continue
        p.remove(r)
    for old in list(p.iter(f"{M}oMath")):
        parent = old.getparent()
        if parent is not None:
            parent.remove(old)

    if num_run is not None:
        num_run.addprevious(omml_block)
    else:
        p.append(omml_block)

    ppr = p.find(f"{W}pPr")
    if ppr is None:
        ppr = etree.Element(f"{W}pPr")
        p.insert(0, ppr)
    tabbed = 0
    if params.get("number_tab", True) and num_run is not None:
        # 与 normalize_equation_paragraphs 同一套**制表位模式**（居中 + 右对齐）
        widths = _eq_number_tab_widths(root)
        width = next((w for (pp, w) in widths if pp is p), 0)
        if width > 0:
            lay = _equation_tab_layout(ppr, p, num_run, width)
            tabbed = 1 if lay["number_tab"] else 0
    else:
        jc = ppr.find(f"{W}jc")
        if jc is None:
            jc = etree.Element(f"{W}jc")
            _ppr_insert_ordered(ppr, jc)
        jc.set(f"{W}val", "center")

    _write_part(path, "word/document.xml", root)
    return {"changed": True, "latex": latex, "number_kept": number,
            "runs": len(list(omath.iter(f"{M}r"))), "numbers_tabbed": tabbed}


def verify_equation_from_latex(path, params: dict) -> bool:
    """复查：目标段落里现在有一个 `m:oMath`（不再是纯文本），且编号未被吞掉。"""
    try:
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        styles = _read_styles_map(path)
        latex = (params.get("latex") or "").strip()
        from earlyz.tools.omml_latex import omath_latex_in, omml_to_latex
        # 与动作同一套定位：selector 与 match_text 都认（建议层常只给 match_text）
        _sel = dict(params.get("selector") or {})
        _mt = params.get("match_text") or _sel.pop("match_text", None)
        if _mt:
            pool = _locate_paragraphs(root, _sel or {"semantic_role": "equation_display"},
                                      styles)
            needle = str(_mt).strip()
            targets = [q for q in pool if needle in _para_full_text(q)]
            if not targets:
                targets = [q for q in pool
                           if needle[:24] and needle[:24] in _para_full_text(q)]
        else:
            targets = _locate_paragraphs(root, _sel, styles)
        if len(targets) == 1:
            p = targets[0]
        elif latex:
            # 选择器按文本匹配，而转换后段落已无 w:t → 退化为「哪段公式的
            # LaTeX 等于目标」。找不到就失败，不放过。
            from earlyz.tools.latex_omml import latex_to_omml as _l2o3
            _want = (omml_to_latex(_l2o3(latex)) or "").strip()
            hit = [q for q in _locate_paragraphs(
                root, {"semantic_role": "equation_display"}, styles)
                if (omath_latex_in(q) or "").strip() and _want
                and (omath_latex_in(q) or "").strip() == _want]
            if len(hit) != 1:
                return False
            p = hit[0]
        else:
            return False
        if p.find(f".//{M}oMath") is None or p.find(f".//{M}oMathPara") is None:
            return False
        if latex:
            from earlyz.tools.latex_omml import latex_to_omml as _l2o4
            cur = (omath_latex_in(p) or "").strip()
            want = (omml_to_latex(_l2o4(latex)) or "").strip()
            if not cur or not want or cur != want:
                return False
        return True
    except Exception:  # noqa: BLE001 — 校验失败不放行
        return False


def verify_equation_paragraphs(path, params: dict) -> bool:
    """复查：公式段缩进归零 + 居中（只查动作**保证**的不变量）。

    字体 / 制表位是尽力而为，不进 verifier —— 结构步骤的 verifier 失败会让整个
    事务回滚，代价远大于收益。
    """
    try:
        M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        style_names = _read_styles_map(path)
        do_center = params.get("center", True)
        for p in root.iter(f"{W}p"):
            role, ppr = _equation_para_context(p, style_names, M)
            if role != "equation_display" or ppr is None:
                continue
            ind = ppr.find(f"{W}ind")
            if ind is not None:
                for attr in ("firstLine", "firstLineChars", "hanging"):
                    if ind.get(f"{W}{attr}") not in (None, "0"):
                        return False
            if do_center:
                # 「公式居中」有两种合法形态：① 无编号段 `jc=center`；
                # ② 带编号段**制表位模式**（居中制表位 + 右制表位，jc 归左）——
                #    制表符会撑满版心宽，此时 `jc=center` 反而是错的。verifier 与
                #    动作必须同一谓词（铁律），所以两种都认。
                tabs = ppr.find(f"{W}tabs")
                has_tab_pattern = (tabs is not None
                                   and any(t.get(f"{W}val") == "center" for t in tabs)
                                   and any(t.get(f"{W}val") == "right" for t in tabs))
                if not has_tab_pattern:
                    jc = ppr.find(f"{W}jc")
                    if jc is None or jc.get(f"{W}val") != "center":
                        return False
        return True
    except Exception:  # noqa: BLE001 — 校验失败不放行
        return False


# ═══════════════════════════════════════════════════════════════════════════
# 前部角标上标（2026-09-28 夜班）：作者行 / 单位行的编号角标必须是上标
# ═══════════════════════════════════════════════════════════════════════════
# 用户报的缺陷：产物里英文作者行 `ZHANG San1)LISi2)` 的 `1)` 是上标
# （LaTeX 归一顺带造出来的真上标），中文作者行 `张 三1)李 四2)` 的 `1)2)`
# 却还是普通文字 —— 而模板里作者行样本段[1] 与单位行样本段[2][3] 的编号
# run **全都带 `w:vertAlign="superscript"`**。
#
# 为什么不能靠模板 extras 回放：vertAlign 是「run 级且**行内选择性**」的属性
# （一行里只有编号那两三个 run 是上标），模板 extras 通道是段落级整体回放，
# 且 `_apply_extra_xml` 明确把 vertAlign 排除（否则整行都会变上标）。
# 因此单列一个动作：定位「括号外、前面不是数字」的编号 run 组并打上标。
#
# 取值出自当前模板：compile_document 先用 front_mark_roles_in_document()
# 检查模板样本里编号 run 是否真的带 vertAlign，带了才调度本动作；
# 没带（或 spec 来自 YAML/内置、无模板可查）→ 不动，保持原样。

FRONT_MARK_ROLES = ("authors", "authors_en", "affiliation", "affiliation_en")

# 编号角标形状：`1)` / `（1）` / `2),3)` / `（1）,（2）`
# 编号限定 1-2 位：作者/单位编号不会更多，而邮编（100084）6 位天然排除。
_FRONT_MARK_RE = re.compile(r"[（(]?\d{1,2}\s*[)）](?:\s*[,，]\s*[（(]?\d{1,2}\s*[)）])*")
_FRONT_MARK_MAX_RUNS = 6      # 一个角标最多跨这么多 run（Word 常把 `1` 与 `)` 拆开）
_FRONT_MARK_MAX_LEN = 10      # 组文本长度上限，防误吞正文


def _run_text_of(r) -> str:
    return "".join(t.text or "" for t in r.iter(f"{W}t"))


def _dict_runs(p) -> list:
    """段落直属 run（排除公式 / OLE 子树内的 run）。"""
    out = []
    for r in p.findall(f"{W}r"):
        if any(a.tag in (f"{M}oMath", f"{M}oMathPara") or a.tag.endswith("object")
               or a.tag.endswith("OLEObject") for a in r.iterancestors()):
            continue
        out.append(r)
    return out


def _paren_depth(s: str) -> int:
    return (s.count("(") + s.count("（")) - (s.count(")") + s.count("）"))


def _paragraph_text_map(p) -> tuple:
    """返回 (全文, [(run, start, end)])：段落直属 run 的文本与偏移。

    跳过公式 / OLE 子树内的 run，以及**含非文本内容**的 run（图片 / 域 / 脚注
    引用等）——那些 run 不能整体切分，且角标不会长在它们里面。
    """
    parts: list = []
    mapping: list = []
    pos = 0
    for r in p.findall(f"{W}r"):
        if any(a.tag in (f"{M}oMath", f"{M}oMathPara") or a.tag.endswith("object")
               or a.tag.endswith("OLEObject") for a in r.iterancestors()):
            continue
        if any(not c.tag.endswith("}rPr") and not c.tag.endswith("}t")
               for c in list(r)):
            continue
        txt = "".join(t.text or "" for t in r.iter(f"{W}t"))
        parts.append(txt)
        mapping.append((r, pos, pos + len(txt)))
        pos += len(txt)
    return "".join(parts), mapping


def _accepted_mark_spans(text: str) -> list:
    """文本级：找出**安全**的编号角标 span（`1)` / `2),3)` / `（1）`）。

    保守判据（防误伤，均来自真机反例）：
      · 必须整体匹配编号形状，且 ≤ 10 字符；
      · 必须位于**括号深度 0** 处 —— 单位邮编 `北京 100084）` 在括号内，天然排除；
      · 前一个字符不能是数字 —— `100084` 这类数字串不会被截出角标；
      · 最多认 4 个/段（作者行、单位行不会更多）。

    ⚠️ 深度只按**非角标文本**累计：`张 三1)李 四2)` 里 `1)` 的 `)` 若计入深度，
    后面的 `2)` 会看到 depth=-1 被漏掉（2026-09-28 真机 bug）。
    """
    spans: list = []
    depth = 0
    pos = 0
    for m in _FRONT_MARK_RE.finditer(text):
        if m.start() < pos or len(spans) >= 4:
            continue
        depth += _paren_depth(text[pos:m.start()])
        prev = text[m.start() - 1] if m.start() > 0 else ""
        if depth == 0 and not prev.isdigit() and len(m.group()) <= _FRONT_MARK_MAX_LEN:
            spans.append((m.start(), m.end()))
        else:
            depth += _paren_depth(m.group())
        pos = m.end()
    return spans


def front_mark_groups(p) -> list:
    """返回 [(run 列表, 组文本)]：段落里位于括号外、前面不是数字的编号角标组。

    只读判定（不切分 run）：返回**覆盖该角标**的 run 列表——只要其中每个 run
    都带 `vertAlign=superscript`，就算这个角标已经是上标了（模板侧闸门与
    verifier 都用这个口径）。
    """
    text, mapping = _paragraph_text_map(p)
    out: list = []
    for (s, e) in _accepted_mark_spans(text):
        covered = [r for (r, rs, re_) in mapping if rs < e and re_ > s]
        if covered:
            out.append((covered, text[s:e]))
    return out


def _split_run_at(p, pos: int) -> bool:
    """在段落文本偏移 pos 处切开 run（复制 rPr，文本不变）。pos 须落在 run 内部。"""
    _text, mapping = _paragraph_text_map(p)
    for (r, rs, re_) in mapping:
        if not (rs < pos < re_):
            continue
        off = pos - rs
        body = "".join(t.text or "" for t in r.iter(f"{W}t"))
        new_r = copy.deepcopy(r)
        for t in r.iter(f"{W}t"):
            t.text = body[:off]
            break
        for t in new_r.iter(f"{W}t"):
            t.text = body[off:]
            break
        r.addnext(new_r)
        return True
    return False


def superscript_front_marks(path, params: dict) -> dict:
    """前部（作者行 / 单位行）编号角标 → 上标（run 级）。

    取值出自当前模板样本：仅当模板的对应角色段落里编号 run 带 `vertAlign`
    时 format_compiler 才调度本动作（见 front_mark_roles_in_document）。
    真机事实：产物里 `1)(School of Computer Science…` 常常是**一整个 run**，
    没有角标边界 → 本动作按角标边界**切分 run**（复制 rPr，文本一字不改）
    再给角标那一段打上标。幂等：已是上标则不动。
    """
    path = Path(path)
    roles = tuple(params.get("roles") or FRONT_MARK_ROLES)
    root = _read_part(path, "word/document.xml")
    styles = _read_styles_map(path)
    marked = 0
    split = 0
    touched_para = 0
    marks_seen = 0
    for role in roles:
        for p in _locate_paragraphs(root, {"semantic_role": role}, styles):
            text, _mapping = _paragraph_text_map(p)
            spans = _accepted_mark_spans(text)
            if not spans:
                continue
            marks_seen += len(spans)
            touched_here = False
            # 从后往前切开边界（前面的文本偏移不失效）
            for (s, e) in reversed(spans):
                if _split_run_at(p, s):
                    split += 1
                if _split_run_at(p, e):
                    split += 1
            _text2, mapping2 = _paragraph_text_map(p)
            for (s, e) in spans:
                for (r, rs, re_) in mapping2:
                    if rs >= s and re_ <= e:
                        if _set_vert_align(r, "superscript"):
                            marked += 1
                            touched_here = True
            if touched_here:
                touched_para += 1
    if marked:
        _write_part(path, "word/document.xml", root)
    return {"changed": marked > 0, "roles": list(roles), "marks_seen": marks_seen,
            "marks_marked": marked, "runs_split": split, "paragraphs_touched": touched_para}


def _set_vert_align(run, value: str = "superscript") -> bool:
    """给单个 run 写 vertAlign（按 rPr schema 序列插入），返回是否发生变化。"""
    rpr = run.find(f"{W}rPr")
    if rpr is None:
        rpr = etree.Element(f"{W}rPr")
        run.insert(0, rpr)
    el = rpr.find(f"{W}vertAlign")
    if el is None:
        el = etree.Element(f"{W}vertAlign")
        _rpr_insert_ordered(rpr, el)
    if el.get(f"{W}val") == value:
        return False
    el.set(f"{W}val", value)
    return True


def verify_superscript_front_marks(path, params: dict) -> bool:
    """复查：前部编号角标 run 全部是 vertAlign=superscript。"""
    try:
        path = Path(path)
        roles = tuple(params.get("roles") or FRONT_MARK_ROLES)
        root = _read_part(path, "word/document.xml")
        styles = _read_styles_map(path)
        for role in roles:
            for p in _locate_paragraphs(root, {"semantic_role": role}, styles):
                for group_runs, _seg in front_mark_groups(p):
                    for r in group_runs:
                        rpr = r.find(f"{W}rPr")
                        el = rpr.find(f"{W}vertAlign") if rpr is not None else None
                        if el is None or el.get(f"{W}val") != "superscript":
                            return False
        return True
    except Exception:  # noqa: BLE001 — 校验失败不放行
        return False


def front_mark_roles_in_document(path, roles=FRONT_MARK_ROLES) -> list:
    """模板侧取值：返回**编号 run 真的带上标**的前部角色列表。

    只读模板样本，不写任何东西。空列表 = 模板没这条规矩 → 本动作不调度
    （与「模板没写就不发明」同一口径）。

    中英文成对启用：模板自身写明「中英文单位名称、作者姓名须一致」，而
    analyzer 对模板里的英文作者行/英文单位行**没**给出 `*_en` 角色（真机实测
    模板的英文单位行被判成 `affiliation`）→ 只查 CJK 侧会漏掉目标文档里的
    英文行。故：中/英任一侧样本带上标，就把这一对角色都放行。
    """
    try:
        path = Path(path)
        if not path.exists():
            return []
        root = _read_part(path, "word/document.xml")
        styles = _read_styles_map(path)

        def _role_has_super(path_root, role: str) -> bool:
            for p in _locate_paragraphs(path_root, {"semantic_role": role}, styles):
                for group_runs, _seg in front_mark_groups(p):
                    if all(
                        (lambda rpr: rpr is not None
                         and rpr.find(f"{W}vertAlign") is not None
                         and rpr.find(f"{W}vertAlign").get(f"{W}val") == "superscript")(
                            r.find(f"{W}rPr"))
                        for r in group_runs
                    ):
                        return True
            return False

        hit: list = []
        for role in roles:
            if _role_has_super(root, role):
                hit.append(role)
        # 中英文成对（模板原文：中英文单位名称、作者姓名须一致）
        for cjk, en in (("authors", "authors_en"), ("affiliation", "affiliation_en")):
            if cjk in hit and en not in hit:
                hit.append(en)
            elif en in hit and cjk not in hit:
                hit.append(cjk)
        return [r for r in roles if r in hit]
    except Exception:  # noqa: BLE001 — 模板读不了就不启用（保守）
        return []


def _is_heading_shape_sentence(role: str, text: str) -> bool:
    """句子形态标题段：role 以 heading 开头且含「。」「；」——不是真标题。

    与交付门 delivery._heading_shape_errors 同源（带 outlineLvl 且含句读的
    段判硬错误）；动作（剥 outlineLvl）与 verifier（跳过）必须同一谓词（铁律）。
    """
    return role.startswith("heading") and ("。" in text or "；" in text)


def ensure_heading_outline(path, params: dict) -> dict:
    """标题段补 outlineLvl（Word 大纲级别，2026-08-17 用户彻查截图3：
    「大纲级别=正文文本」——标题样式识别不对的根因）。

    模板示例段从不带 outlineLvl（示例文档直接格式），decompiler 提取
    不到 → 内置语义规则：按 classify_paragraph_role 判定 H1/H2/H3，补
    outlineLvl 0/1/2（「1 引言」/「I. INTRODUCTION」与字体分类同一套）。
    幂等：已有 outlineLvl 且值正确跳过。
    """
    from earlyz.ir.analyzer import classify_paragraph_role
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    styles_map = _read_styles_map(path)
    fixed = 0
    # 封面段不补 outlineLvl（2026-09-25 审计 G）：封面内被判 heading 的段
    # （校名/题名残留标题样式）打上 outlineLvl 后会进目录。
    cover_paras = _cover_paragraphs(root, styles_map)
    for p in root.iter(f"{W}p"):
        if cover_paras and any(p is c for c in cover_paras):
            continue
        # 目录标题本身不是章级标题：给它 outlineLvl 会让目录把自己列进去
        # （2026-11-08 实测 F9 后目录首行出现「目 录 ……3」）。
        if _para_full_text(p).replace(" ", "").replace("　", "") in ("目录", "CONTENTS"):
            continue
        if any(anc.tag == f"{W}tbl" for anc in p.iterancestors(f"{W}tbl")):
            continue
        if p.find(f"{W}hyperlink") is not None:
            continue  # TOC 条目（hyperlink 段）不补 outlineLvl（2026-08-18：目录条目不参与正文大纲）
        text = _para_full_text(p)
        if not text:
            continue
        ppr = p.find(f"{W}pPr")
        # 句子形态段不可能是标题：源文档遗留的 outlineLvl 一律剥掉。
        # 2026-10-25 实测（劳动合同示范文本 ×3）：条款「一、本合同文本供…；」自带
        # outlineLvl，classify 不算 heading → 不走下面的剥离分支 → 留在产物里
        # 被交付门判 heading shape 硬错误，整篇拒绝交付。判定与交付门同源。
        if ppr is not None and ("。" in text or "；" in text):
            _ol_stale = ppr.find(f"{W}outlineLvl")
            if _ol_stale is not None:
                ppr.remove(_ol_stale)
                fixed += 1
        in_table = any(a.tag == f"{W}tbl" for a in p.iterancestors())
        in_hyperlink = p.find(f"{W}hyperlink") is not None
        pstyle = ppr.find(f"{W}pStyle") if ppr is not None else None
        pstyle_id = pstyle.get(f"{W}val", "") if pstyle is not None else ""
        pstyle_name = styles_map.get(pstyle_id, "") or pstyle_id
        numpr = ppr is not None and ppr.find(f"{W}numPr") is not None
        role = classify_paragraph_role(
            ppr, pstyle_name, False, in_table, numpr, text,
            in_hyperlink=in_hyperlink,
        )
        level = _outline_level_for_role(role)
        if level is None:
            continue
        # 句子形态段（role=heading* 且含「。」「；」）不是标题：剥 outlineLvl
        # ——与交付门 _heading_shape_errors 同源（不再限定 heading1/长度；
        # 2026-10-14 审计 C3：动作曾自己造门禁硬错误）
        if _is_heading_shape_sentence(role, text):
            ol_existing = ppr.find(f"{W}outlineLvl") if ppr is not None else None
            if ol_existing is not None:
                ppr.remove(ol_existing)
                fixed += 1
            continue
        ppr = p.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            p.insert(0, ppr)
        ol = ppr.find(f"{W}outlineLvl")
        if ol is None:
            ol = etree.Element(f"{W}outlineLvl")
            _ppr_insert_ordered(ppr, ol)
            ol.set(f"{W}val", level)
            fixed += 1
        elif ol.get(f"{W}val") != level:
            ol.set(f"{W}val", level)
            fixed += 1
    if fixed:
        _write_part(path, "word/document.xml", root)
    return {"changed": fixed > 0, "headings_outlined": fixed}


def verify_heading_outline(path, params: dict) -> bool:
    """复查：分类为 heading1/2/3 的段均有对应 outlineLvl。"""
    try:
        from earlyz.ir.analyzer import classify_paragraph_role
        root = _read_part(Path(path), "word/document.xml")
        styles_map = _read_styles_map(path)
        cover_paras = _cover_paragraphs(root, styles_map)   # 与动作同条件（封面不举证）
        for p in root.iter(f"{W}p"):
            if cover_paras and any(p is c for c in cover_paras):
                continue
            # 与动作同源（铁律）：表内段（有 w:tbl 祖先）不查 outlineLvl
            # （2026-10-14 审计 B3：动作显式跳过表内段，verifier 曾漏同条件）
            if any(a.tag == f"{W}tbl" for a in p.iterancestors()):
                continue
            if p.find(f"{W}hyperlink") is not None:
                continue
            text = _para_full_text(p)
            if not text:
                continue
            # 与动作同条件：目录标题自身不举证（动作不给它补大纲级别）
            if text.replace(" ", "").replace("　", "") in ("目录", "CONTENTS"):
                continue
            ppr = p.find(f"{W}pPr")
            in_table = any(a.tag == f"{W}tbl" for a in p.iterancestors())
            pstyle = ppr.find(f"{W}pStyle") if ppr is not None else None
            pstyle_id = pstyle.get(f"{W}val", "") if pstyle is not None else ""
            pstyle_name = styles_map.get(pstyle_id, "") or pstyle_id
            numpr = ppr is not None and ppr.find(f"{W}numPr") is not None
            role = classify_paragraph_role(
                ppr, pstyle_name, False, in_table, numpr, text, in_hyperlink=False,
            )
            # 与动作同一谓词（铁律）：句子形态段不查 outlineLvl
            if _is_heading_shape_sentence(role, text):
                continue
            want = _outline_level_for_role(role)
            if want is None:
                continue
            if ppr is None or ppr.find(f"{W}outlineLvl") is None:
                return False
            if ppr.find(f"{W}outlineLvl").get(f"{W}val") != want:
                return False
        return True
    except Exception:
        return False


def ensure_toc_section_break(path, params: dict) -> dict:
    """目录与正文之间必须有分节符（2026-11-08 用户实锤：目录和正文排在同一页）。

    背景：目录条目永不带段前分页（上一处修复），正文首章的 pageBreakBefore 又会被
    Word 的「更新目录」重写掉（实测 F9 后该段 PBB 消失）——两条路都靠不住。
    可靠做法是把分节符放在**目录最后一条之后的空段**上：sectPr 是段落属性，
    更新域只重写域结果文本，不会动它（实测 F9 后分节仍在，正文独立起页）。

    幂等：目录后已紧跟分节段 → no-op。目录条目不存在 → no-op。
    """
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False, "skipped": "no_body"}
    paras = [x for x in body.iter(f"{W}p") if x.getparent() is body]
    last_toc = -1
    for i, p in enumerate(paras):
        ppr = p.find(f"{W}pPr")
        st = ppr.find(f"{W}pStyle") if ppr is not None else None
        if st is not None and str(st.get(f"{W}val") or "").upper().startswith("TOC"):
            last_toc = i
    if last_toc < 0:
        return {"changed": False, "skipped": "no_toc"}
    nxt = paras[last_toc + 1] if last_toc + 1 < len(paras) else None
    if nxt is not None:
        nppr = nxt.find(f"{W}pPr")
        if nppr is not None and nppr.find(f"{W}sectPr") is not None:
            return {"changed": False, "skipped": "already_sectioned"}
        txt = _para_full_text(nxt).strip()
        if not txt:
            # 空段承载分节符（就是我们要的形态）
            if nppr is not None and nppr.find(f"{W}sectPr") is not None:
                return {"changed": False, "skipped": "already_sectioned"}
    src_sp = body.find(f"{W}sectPr")
    if src_sp is None:
        return {"changed": False, "skipped": "no_sectpr"}
    new_sp = copy.deepcopy(src_sp)
    ty = new_sp.find(f"{W}type")
    if ty is None:
        ty = etree.Element(f"{W}type")
        new_sp.insert(0, ty)
    ty.set(f"{W}val", "nextPage")
    holder = etree.Element(f"{W}p")
    hpr = etree.SubElement(holder, f"{W}pPr")
    hpr.append(new_sp)
    paras[last_toc].addnext(holder)
    _write_part(path, "word/document.xml", root)
    return {"changed": True, "section_break_after_paragraph": last_toc}


def verify_toc_section_break(path, params: dict) -> bool:
    """复查：目录最后一条之后紧跟带 sectPr 的段。"""
    try:
        path = Path(path)
        root = _read_part(path, "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return True
        paras = [x for x in body.iter(f"{W}p") if x.getparent() is body]
        last_toc = -1
        for i, p in enumerate(paras):
            ppr = p.find(f"{W}pPr")
            st = ppr.find(f"{W}pStyle") if ppr is not None else None
            if st is not None and str(st.get(f"{W}val") or "").upper().startswith("TOC"):
                last_toc = i
        if last_toc < 0 or last_toc + 1 >= len(paras):
            return True
        nppr = paras[last_toc + 1].find(f"{W}pPr")
        return nppr is not None and nppr.find(f"{W}sectPr") is not None
    except Exception:  # noqa: BLE001
        return True


def ensure_section_page_breaks(path, params: dict) -> dict:
    """前置章节另起页（2026-08-17 用户彻查截图1：目录和英文摘要同页）。

    根因：源文档目录节 sectPr type=continuous（不分页节）——ABSTRACT
    （Keywords 段）结束后目录直接接排同页（Word COM 实测 SEC8 continuous）。
    规则：前置区章节标题段（摘要/ABSTRACT/目录/致谢/参考文献/声明等）
    一律补 pageBreakBefore——已处于节首时无副作用（前节 nextPage 已是
    新页首，PBB 不产生额外空页）；continuous 节则强制从新页开始。
    幂等：已有 pageBreakBefore 跳过。
    """
    path = Path(path)
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    fixed = 0
    # 节序（含 sdt 内段落）——判断标题段是否已是节首（前节 nextPage）
    # 2026-08-18 用户彻查「空白页」：前置章节标题无条件补 PBB 会叠空页
    # （continuous 节尾+nextPage 分节已新起页，再 PBB = 双分页空页 p14/74）
    sects = _section_types(root)
    # 2026-10-06 取值出自模板：params["section_break_words"]（哪怕空集）是模板
    # 分页要求的全集；键缺失=无模板信息 → _FRONT_CHAPTER_WORDS 兜底（08-17 口径）。
    # 2026-10-08：词表成员同过 _norm_ws（老 spec 里存「摘□□要」原文——文本侧归一
    # 而词表侧不归一，匹配就断了）；键同过 _break_key（行尾注记剥离 + ASCII 大写，
    # 「Abstract（…注记）」才能命中词表「ABSTRACT」）。
    forced = (params or {}).get("section_break_words", None)
    forced_set = ({_break_key(str(w)) for w in forced}
                  if isinstance(forced, (list, tuple, set, frozenset)) else None)
    front_keys = {_break_key(w) for w in _FRONT_CHAPTER_WORDS}
    for p in body.iter(f"{W}p"):
        # 目录条目不是章节本身：TOC1/2/3 段落的文本形如「附录 A11」「致谢10」，
        # 会被下面的词表匹配当成章节标题补 PBB → 目录在「附录 A」处提前翻页
        # （2026-11-08 用户实锤：目录第 3 页末尾大段空白，附录 A 条目跑到第 4 页）。
        # 目录条目永不带 pageBreakBefore。
        _ppr_t = p.find(f"{W}pPr")
        _st_t = _ppr_t.find(f"{W}pStyle") if _ppr_t is not None else None
        if _st_t is not None and str(_st_t.get(f"{W}val") or "").upper().startswith("TOC"):
            # 已带 PBB 的目录条目（旧产物遗留）就地清掉——目录条目永不带段前分页。
            if _ppr_t is not None:
                _old_pbb = _ppr_t.find(f"{W}pageBreakBefore")
                if _old_pbb is not None:
                    _ppr_t.remove(_old_pbb)
                    fixed += 1
            continue
        text = _para_full_text(p)
        cands = _break_cands(text)
        if forced_set is not None:
            if not (cands & forced_set) and not _is_chap_break_cand(p, text):
                continue
        elif not (cands & front_keys):
            continue
        # 已处于节首且前节 nextPage 且本节也是 nextPage → 自然新页，
        # 补 PBB 会叠空页；**continuous 节（目录/摘要/参考文献）即使前节
        # nextPage 也要 PBB**——continuous 不强制新页，会接排前节末页
        # （2026-08-18 用户彻查「ABSTRACT 与目录同页」：目录节 continuous
        # 接排 ABSTRACT 页）。本节的 type 从节尾段 sectPr 取。
        is_sect_start = False
        prev_next_page = False
        this_continuous = False
        for m in sects:
            if m["start_text"] and _norm_ws(m["start_text"]) == _norm_ws(text):
                is_sect_start = True
                sp = m.get("sp_el")
                # 前一个节尾段的 type
                prev_sp = None
                for m2 in sects:
                    if m2["index"] == m["index"] - 1:
                        prev_sp = m2.get("sp_el")
                        break
                if prev_sp is not None:
                    # 2026-08-18 与 repair 同源：sp_el 是含 sectPr 的段落，
                    # type 在段落内的 sectPr 上——不解封装会拿到 None → 误判
                    # 「前节 nextPage」跳过 PBB（ABSTRACT 与摘要同页回归根因）
                    if prev_sp.tag == f"{W}p":
                        _ppr = prev_sp.find(f"{W}pPr")
                        _spx = _ppr.find(f"{W}sectPr") if _ppr is not None else None
                        if _spx is not None:
                            prev_sp = _spx
                    tp = prev_sp.find(f"{W}type")
                    if tp is None or tp.get(f"{W}val") != "continuous":
                        prev_next_page = True
                # 本节 type（continuous？）——从本节节尾段 sectPr 取
                if sp is not None:
                    if sp.tag == f"{W}p":
                        _ppr = sp.find(f"{W}pPr")
                        _spx = _ppr.find(f"{W}sectPr") if _ppr is not None else None
                        if _spx is not None:
                            sp = _spx
                    _tp = sp.find(f"{W}type")
                    if _tp is not None and _tp.get(f"{W}val") == "continuous":
                        this_continuous = True
                break
        if is_sect_start and not this_continuous:
            # nextPage 节（无论前节 type）分节符已强制新页，PBB 多余会叠
            # 空页（2026-08-18 r21 实证 p14：目录 continuous + 正文 nextPage
            # 补 PBB → 双分页空页）；continuous 节必须 PBB（不强制新页）。
            continue
        ppr = p.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            p.insert(0, ppr)
        if ppr.find(f"{W}pageBreakBefore") is None:
            etree.SubElement(ppr, f"{W}pageBreakBefore")
            fixed += 1
    if fixed:
        _write_part(path, "word/document.xml", root)
    return {"changed": fixed > 0, "page_breaks_added": fixed}


def verify_section_page_breaks(path, params: dict) -> bool:
    """复查：前置章节标题段均有 pageBreakBefore（已处节首且前节 nextPage
    的自然新页段除外——2026-08-18 空页修复，与 set 侧同源）。"""
    try:
        root = _read_part(Path(path), "word/document.xml")
        body = root.find(f"{W}body")
        sects = _section_types(root)
        # 2026-10-06 取值出自模板：params["section_break_words"]（哪怕空集）是模板
        # 分页要求的全集；键缺失=无模板信息 → _FRONT_CHAPTER_WORDS 兜底（08-17 口径）。
        # 2026-10-08：词表成员同过 _norm_ws（老 spec 里存「摘□□要」原文——文本侧归一
        # 而词表侧不归一，匹配就断了）；键同过 _break_key（与 set 侧同源）。
        forced = (params or {}).get("section_break_words", None)
        forced_set = ({_break_key(str(w)) for w in forced}
                      if isinstance(forced, (list, tuple, set, frozenset)) else None)
        front_keys = {_break_key(w) for w in _FRONT_CHAPTER_WORDS}
        for p in body.iter(f"{W}p"):
            text = _para_full_text(p)
            cands = _break_cands(text)
            if forced_set is not None:
                if not (cands & forced_set) and not _is_chap_break_cand(p, text):
                    continue
            elif not (cands & front_keys):
                continue
            is_sect_start = False
            prev_next_page = False
            this_continuous = False
            for m in sects:
                if m["start_text"] and _norm_ws(m["start_text"]) == _norm_ws(text):
                    is_sect_start = True
                    for m2 in sects:
                        if m2["index"] == m["index"] - 1:
                            prev_sp = m2.get("sp_el")
                            if prev_sp is not None:
                                # 与 set 侧同源解封装（2026-08-18）
                                if prev_sp.tag == f"{W}p":
                                    _ppr = prev_sp.find(f"{W}pPr")
                                    _spx = _ppr.find(f"{W}sectPr") if _ppr is not None else None
                                    if _spx is not None:
                                        prev_sp = _spx
                                tp = prev_sp.find(f"{W}type")
                                if tp is None or tp.get(f"{W}val") != "continuous":
                                    prev_next_page = True
                            break
                    # 本节 type（continuous？）——与 set 侧同源
                    sp = m.get("sp_el")
                    if sp is not None:
                        if sp.tag == f"{W}p":
                            _ppr = sp.find(f"{W}pPr")
                            _spx = _ppr.find(f"{W}sectPr") if _ppr is not None else None
                            if _spx is not None:
                                sp = _spx
                        _tp = sp.find(f"{W}type")
                        if _tp is not None and _tp.get(f"{W}val") == "continuous":
                            this_continuous = True
                    break
            if is_sect_start and not this_continuous:
                continue
            ppr = p.find(f"{W}pPr")
            if ppr is None or ppr.find(f"{W}pageBreakBefore") is None:
                return False
        return True
    except Exception:
        return False


def ensure_update_fields(path, params: dict) -> dict:
    """移除 settings.xml 的 w:updateFields（2026-08-19 弹窗实证推翻）。

    用户截图铁证：只要 settings.xml 有 updateFields=true，Word 打开就弹
    「该文档包含的域可能引用了其他文件」。上一轮误判「拆 TOC 外层域后
    不弹窗」——PAGEREF/REF 仍会被 Word 当成「可能引用其他文件」。

    对照实验（同一产物去/留 updateFields，Word COM SaveAs PDF）：
    页脚序列完全一致（封面无页码 → I/III/V → 正文 7–67）。PAGE 域
    渲染不依赖此开关；此前「去掉就每页 1」是 update_toc_entries 清空
    PAGEREF 缓存 + 误读页脚 Range.Text 缓存值的叠加误判。

    本 action 只做移除（幂等）。目录页码靠保留 PAGEREF 结果缓存
    （update_toc_entries 不再清空）；页脚 PAGE 域 Word 显示/打印自刷新。
    """
    path = Path(path)
    with zipfile.ZipFile(path) as z:
        if "word/settings.xml" not in z.namelist():
            return {"changed": False, "update_fields": False}
    root = _read_part(path, "word/settings.xml")
    changed = False
    uf = root.find(f"{W}updateFields")
    if uf is not None:
        root.remove(uf)
        changed = True
    if changed:
        _write_part(path, "word/settings.xml", root)
    return {"changed": changed, "update_fields": False}


def verify_update_fields(path, params: dict) -> bool:
    """复查：settings.xml 无 w:updateFields（打开不弹更新域）。"""
    try:
        path = Path(path)
        with zipfile.ZipFile(path) as z:
            if "word/settings.xml" not in z.namelist():
                return True
        root = _read_part(path, "word/settings.xml")
        uf = root.find(f"{W}updateFields")
        return uf is None
    except Exception:
        return False


def _apply_pageref_results(path: Path, results: dict) -> int:
    """把 {书签: 页码} 写入 PAGEREF 域结果 w:t，返回改写条数。"""
    root = _read_part(path, "word/document.xml")
    applied = 0
    for p in root.iter(f"{W}p"):
        current_bk = None
        in_result = False
        result_t = None
        for el in p.iter():
            tag = etree.QName(el).localname
            if tag == "fldChar":
                kind = el.get(f"{W}fldCharType")
                if kind == "begin":
                    current_bk = None
                    in_result = False
                    result_t = None
                elif kind == "separate":
                    in_result = True
                elif kind == "end":
                    if current_bk and current_bk in results and result_t is not None:
                        new = results[current_bk]
                        if (result_t.text or "") != new:
                            result_t.text = new
                            applied += 1
                    current_bk = None
                    in_result = False
                    result_t = None
            elif tag == "instrText":
                m = re.search(r"PAGEREF\s+(\S+)", el.text or "", re.I)
                if m:
                    current_bk = m.group(1)
            elif tag == "t" and in_result:
                result_t = el
    if applied:
        _write_part(path, "word/document.xml", root)
    return applied


def refresh_page_fields(path, params: dict) -> dict:
    """用 Word COM 算出真实 PAGEREF 页码写回 OOXML，不 Save 整包。

    2026-08-19 用户拍板：目录数字打开就精确，同时不弹更新域。
    Word 只读打开 → Repaginate + 只刷新 PAGEREF（type=37，不动 REF，
    避免源文档缺书签变成「错误!未找到引用源」）→ 把结果写进域缓存
    → 再剥 updateFields。Word 不可用则跳过（保留源缓存，仍不弹窗）。
    """
    path = Path(path)
    skipped = False
    applied = 0
    error = None
    root = _read_part(path, "word/document.xml")
    has_pageref = any("PAGEREF" in (el.text or "").upper()
                      for el in root.iter(f"{W}instrText"))
    if has_pageref:
        try:
            from earlyz.validation.renderers import collect_pageref_results
            # 2026-09-21 P3 校验失败实锤：PS 脚本 ExportAsFixedFormat 强制
            # 「分页+域全部重算」，随后 $d.Save() 把 document.xml 里**所有域**
            # （含 REF 交叉引用）的重算结果落盘。源文档断链书签（_Ref4629
            # 等 11 个缺失）→ REF 域结果变「错误!未找到引用源。」/0 → 正文
            # 文本 hash 变化 → 内容保全校验 FAILED 拒绝交付。
            # 修复：Word 只当页码计算器——调用前整包快照，拿到 PAGEREF 结果
            # 后整包恢复（Word Save 会重写 rels/settings 等多 part，只恢复
            # document.xml 会造成包内不一致 → Word 报损坏），再仅把 PAGEREF
            # 结果写回（REF/其它域保持原文）。PAGE 域缓存（页脚页码）是
            # Save 的本意产物，从 Word 版合并回来保持页码同步。
            backup_pkg = _snapshot_package(path)
            results = collect_pageref_results(path)
            page_results = _collect_domain_results(_snapshot_package(path))
            _restore_package(path, backup_pkg)
            if page_results:
                _apply_domain_results(path, page_results)
            if results:
                applied = _apply_pageref_results(path, results)
        except Exception as exc:
            # RendererUnavailable / RenderFailed / 其它：不阻断交付
            skipped = True
            error = f"{type(exc).__name__}: {exc}"[:240]

    # 无论 Word 是否跑成，产物不得带 updateFields
    uf_removed = False
    try:
        sroot = _read_part(path, "word/settings.xml")
        uf = sroot.find(f"{W}updateFields")
        if uf is not None:
            sroot.remove(uf)
            _write_part(path, "word/settings.xml", sroot)
            uf_removed = True
    except Exception:
        pass

    return {
        "changed": applied > 0 or uf_removed,
        "applied": applied,
        "skipped": skipped,
        "error": error,
        "update_fields": False,
    }


def verify_page_fields_refreshed(path, params: dict) -> bool:
    """复查：无 updateFields；有 PAGEREF 则至少一条结果非空（无渲染器时放行）。"""
    try:
        path = Path(path)
        with zipfile.ZipFile(path) as z:
            has_settings = "word/settings.xml" in z.namelist()
        if has_settings:
            sroot = _read_part(path, "word/settings.xml")
            if sroot.find(f"{W}updateFields") is not None:
                return False
        root = _read_part(path, "word/document.xml")
        has_pageref = False
        has_result = False
        in_result = False
        is_pageref = False
        for el in root.iter():
            tag = etree.QName(el).localname
            if tag == "instrText" and "PAGEREF" in (el.text or "").upper():
                has_pageref = True
                is_pageref = True
            elif tag == "fldChar":
                kind = el.get(f"{W}fldCharType")
                if kind == "begin":
                    is_pageref = False
                    in_result = False
                elif kind == "separate":
                    in_result = True
                elif kind == "end":
                    in_result = False
                    is_pageref = False
            elif tag == "t" and in_result and is_pageref and (el.text or "").strip():
                has_result = True
        if not has_pageref:
            return True
        if has_result:
            return True
        # 与动作同源（2026-10-14 审计 B4）：动作 refresh_page_fields 在无
        # 渲染器时合法 skipped（RendererUnavailable 等只记 error、不阻断
        # 交付），此时 PAGEREF 缓存为空是环境缺失而非产物缺陷——verifier
        # 不得把环境缺失升级为整篇失败。用与动作入口同一探测
        # （collect_pageref_results 的 word_render_available 判断）：无渲染器
        # → 放行（vacuous pass）；有渲染器则维持严格核验「结果已刷新」。
        try:
            from earlyz.validation.renderers import word_render_available
            renderer_ok = bool(word_render_available())
        except Exception:
            renderer_ok = False
        if not renderer_ok:
            return True
        return has_result
    except Exception:
        return False


def apply_drop_cap(path, params: dict) -> dict:
    """把模板的首字下沉套到第一个一级标题后的正文段。取值全来自当前模板。"""
    from earlyz.ir.analyzer import classify_paragraph_role

    props = dict(params.get("properties") or params)
    layout = dict(params.get("layout") or {})
    kind = props.get("drop_cap")
    if not kind:
        return {"applied": 0}
    lines = int(props.get("drop_cap_lines") or 3)
    sz_pt = props.get("font_size_pt")
    want_sc = props.get("small_caps")
    line = props.get("line")
    line_rule = props.get("line_rule") or "exact"
    jc = props.get("jc") or "both"
    before = props.get("spacing_before")
    after = props.get("spacing_after")

    path = Path(path)
    root = _read_part(path, "word/document.xml")
    styles_map = _read_styles_map(path)
    body = root.find(f"{W}body")
    paras = [p for p in list(body) if p.tag == f"{W}p"]

    def _role(p):
        text = _para_full_text(p)
        ppr = p.find(f"{W}pPr")
        pstyle = ppr.find(f"{W}pStyle") if ppr is not None else None
        sid = pstyle.get(f"{W}val", "") if pstyle is not None else ""
        name = styles_map.get(sid, "") or sid
        numpr = ppr is not None and ppr.find(f"{W}numPr") is not None
        return classify_paragraph_role(
            ppr, name, False, False, numpr, text or "",
        ), (text or "")

    seen_h1 = False
    target = None
    for p in paras:
        role, text = _role(p)
        if role == "heading1":
            seen_h1 = True
            continue
        if not seen_h1 or not text.strip():
            continue
        if role in ("drop_cap", "body"):
            target = p
            break
    if target is None:
        return {"applied": 0}

    tpr = target.find(f"{W}pPr")
    existing = tpr.find(f"{W}framePr") if tpr is not None else None
    already = existing is not None and existing.get(f"{W}dropCap")
    letter_only = len(_para_full_text(target).strip()) == 1

    if already and letter_only:
        _set_run_properties(target, props)
        _write_drop_cap_ppr(target, kind, lines, line, line_rule, jc, before, after)
        _apply_drop_cap_layout(target, layout)
        _write_part(path, "word/document.xml", root)
        return {"applied": 1, "mode": "format"}

    full = _para_full_text(target)
    i = 0
    while i < len(full) and full[i].isspace():
        i += 1
    if i >= len(full):
        return {"applied": 0}
    letter = full[i]
    rest = full[i + 1 :]
    m = re.match(r"^(\S*)", rest)
    first_word = m.group(1) if m else ""
    after_word = rest[len(first_word):]

    drop = etree.Element(f"{W}p")
    _write_drop_cap_ppr(drop, kind, lines, line, line_rule, jc, before, after)
    r = etree.SubElement(drop, f"{W}r")
    rpr = etree.SubElement(r, f"{W}rPr")
    if want_sc:
        etree.SubElement(rpr, f"{W}smallCaps")
    if sz_pt:
        sz = etree.SubElement(rpr, f"{W}sz")
        sz.set(f"{W}val", str(int(round(float(sz_pt) * 2))))
        szcs = etree.SubElement(rpr, f"{W}szCs")
        szcs.set(f"{W}val", str(int(round(float(sz_pt) * 2))))
    t_el = etree.SubElement(r, f"{W}t")
    t_el.text = letter
    _apply_drop_cap_layout(drop, layout)
    target.addprevious(drop)

    for old_t in list(target.iter(f"{W}t")):
        old_t.text = ""
    runs = target.findall(f"{W}r")
    if not runs:
        runs = [etree.SubElement(target, f"{W}r")]
    # first remaining word → small caps (THIS → HIS)
    r0 = runs[0]
    rpr0 = r0.find(f"{W}rPr")
    if rpr0 is None:
        rpr0 = etree.Element(f"{W}rPr")
        r0.insert(0, rpr0)
    if want_sc and first_word:
        if rpr0.find(f"{W}smallCaps") is None:
            etree.SubElement(rpr0, f"{W}smallCaps")
    ts = list(r0.iter(f"{W}t"))
    if not ts:
        ts = [etree.SubElement(r0, f"{W}t")]
    ts[0].text = first_word
    if after_word:
        r1 = etree.Element(f"{W}r")
        if rpr0 is not None:
            rpr1 = copy.deepcopy(rpr0)
            sc = rpr1.find(f"{W}smallCaps")
            if sc is not None:
                rpr1.remove(sc)
            r1.insert(0, rpr1)
        t1 = etree.SubElement(r1, f"{W}t")
        if after_word[:1].isspace():
            t1.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t1.text = after_word
        r0.addnext(r1)
    # 空 run 清掉，避免残留。只删「确实无任何子结构」的空 run——含脚注/
    # 尾注/批注引用、域字符/域指令、图形/图片/嵌入对象、换行/制表/符号的
    # run 即使没有 w:t 文本也一律保留（2026-10-14 审计实测：footnoteReference
    # 曾被静默删丢、域 run 被拆散）。
    _keep_tags = frozenset((
        f"{W}footnoteReference", f"{W}endnoteReference", f"{W}commentReference",
        f"{W}fldChar", f"{W}instrText", f"{W}drawing", f"{W}pict",
        f"{W}object", f"{W}br", f"{W}tab", f"{W}sym",
    ))
    for old in list(target.findall(f"{W}r")):
        if old in (r0,) or (after_word and old.getparent() is not None and old == r0.getnext()):
            continue
        has_text = bool("".join(x.text or "" for x in old.iter(f"{W}t")).strip())
        has_sub = any(el.tag in _keep_tags for el in old.iter())
        if not has_text and not has_sub:
            target.remove(old)
    # 模板续段无首行缩进；带着 firstLine 会把下沉字母挤残
    tpr = target.find(f"{W}pPr")
    if tpr is not None:
        ind = tpr.find(f"{W}ind")
        if ind is not None:
            for attr in (f"{W}firstLine", f"{W}firstLineChars"):
                if ind.get(attr) is not None:
                    del ind.attrib[attr]
            if len(ind.attrib) == 0:
                tpr.remove(ind)

    _write_part(path, "word/document.xml", root)
    return {"applied": 1, "mode": "split", "letter": letter}


def _layout_xml(layout: dict, key: str):
    raw = layout.get(key)
    if not raw:
        return None
    try:
        return etree.fromstring(raw.encode("utf-8") if isinstance(raw, str) else raw)
    except Exception:
        return None


def _apply_drop_cap_layout(p, layout: dict) -> None:
    """原样套当前模板首字段落组合结构；缺布局证据则不发明。"""
    if not layout:
        return
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        return
    run = p.find(f"{W}r")
    # pPr/rPr 和 run/rPr 都必须存在；Word 预留首字宽度依赖这个组合
    for parent, key in ((ppr, "ppr_rpr_xml"), (run, "run_rpr_xml")):
        if parent is None:
            continue
        fresh = _layout_xml(layout, key)
        if fresh is None:
            continue
        old = parent.find(f"{W}rPr")
        if old is not None:
            parent.replace(old, fresh)
        else:
            parent.append(fresh)
    pbdr = _layout_xml(layout, "pbdr_xml")
    if pbdr is not None:
        old = ppr.find(f"{W}pBdr")
        if old is not None:
            ppr.replace(old, pbdr)
        else:
            sp = ppr.find(f"{W}spacing")
            if sp is not None:
                sp.addprevious(pbdr)
            else:
                ppr.append(pbdr)
    ta_val = layout.get("text_alignment")
    if ta_val:
        ta = ppr.find(f"{W}textAlignment")
        if ta is None:
            ta = etree.Element(f"{W}textAlignment")
            _ppr_insert_ordered(ppr, ta)
        ta.set(f"{W}val", str(ta_val))


def _write_drop_cap_ppr(p, kind, lines, line, line_rule, jc,
                        before=None, after=None) -> None:
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        ppr = etree.Element(f"{W}pPr")
        p.insert(0, ppr)
    if ppr.find(f"{W}keepNext") is None:
        _ppr_insert_ordered(ppr, etree.Element(f"{W}keepNext"))
    fr = ppr.find(f"{W}framePr")
    if fr is None:
        fr = etree.Element(f"{W}framePr")
        _ppr_insert_ordered(ppr, fr)
    fr.set(f"{W}dropCap", str(kind))
    fr.set(f"{W}lines", str(int(lines)))
    fr.set(f"{W}wrap", "around")
    fr.set(f"{W}vAnchor", "text")
    fr.set(f"{W}hAnchor", "text")
    sp = ppr.find(f"{W}spacing")
    if sp is None:
        sp = etree.SubElement(ppr, f"{W}spacing")
    if line:
        sp.set(f"{W}line", str(int(line)))
        sp.set(f"{W}lineRule", str(line_rule))
    # 模板没写段前段后则写 0，否则 Word 默认段后 8 磅
    sp.set(f"{W}before", str(int(before if before is not None else 0)))
    sp.set(f"{W}after", str(int(after if after is not None else 0)))
    if jc:
        jel = ppr.find(f"{W}jc")
        if jel is None:
            jel = etree.SubElement(ppr, f"{W}jc")
        jel.set(f"{W}val", str(jc))


def verify_drop_cap(path, params: dict) -> bool:
    props = dict(params.get("properties") or params)
    if not props.get("drop_cap"):
        return True
    try:
        from earlyz.ir.analyzer import analyze_document
        ir = analyze_document(path)
        for p in ir.paragraphs:
            if p.semantic_role == "drop_cap":
                fp = p.direct_format
                want = props.get("font_size_pt")
                if want is not None and fp.font_size_pt != want:
                    return False
                if fp.drop_cap != props.get("drop_cap"):
                    return False
                return True
        return False
    except Exception:
        return False

def align_data_table_columns(target_path: Path, params: dict) -> dict:
    """数据表数值列对齐（2026-10-21，P1）。

    规则：kind=data 且被判定为数值列的，正文数据行的单元格右对齐；
    数字不换行（w:noWrap），避免「1,048.75」被拆行。
    表头行、注释行、单位行不动；kind=form 的表完全不动。
    """
    from earlyz.actions.table_layout import table_profile

    path = Path(target_path)
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False}
    aligned = 0
    profiles = []
    for tbl in body.iter(f"{W}tbl"):
        prof = table_profile(tbl)
        profiles.append({k: prof[k] for k in ("kind", "header_rows", "numeric_cols", "ncols")})
        if prof["kind"] != "data" or not prof["numeric_cols"]:
            continue
        skip = set(prof["note_rows"]) | ({prof["unit_row"]} if prof["unit_row"] is not None else set())
        for ri, tr in enumerate(tbl.findall(f"{W}tr")):
            if ri < prof["header_rows"] or ri in skip:
                continue
            col = 0
            for tc in tr.findall(f"{W}tc"):
                tcpr = tc.find(f"{W}tcPr")
                span = 1
                if tcpr is not None:
                    gs = tcpr.find(f"{W}gridSpan")
                    if gs is not None:
                        try:
                            span = max(1, int(gs.get(f"{W}val") or 1))
                        except (TypeError, ValueError):
                            span = 1
                if col in prof["numeric_cols"]:
                    text = "".join(t.text or "" for t in tc.iter(f"{W}t")).strip()
                    for para in tc.iter(f"{W}p"):
                        ppr = para.find(f"{W}pPr")
                        if ppr is None:
                            ppr = etree.Element(f"{W}pPr")
                            para.insert(0, ppr)
                        jc = ppr.find(f"{W}jc")
                        if jc is None:
                            jc = etree.SubElement(ppr, f"{W}jc")
                        jc.set(f"{W}val", "right")
                    if text and len(text) <= 12:
                        if tcpr is None:
                            tcpr = etree.Element(f"{W}tcPr")
                            tc.insert(0, tcpr)
                        if tcpr.find(f"{W}noWrap") is None:
                            etree.SubElement(tcpr, f"{W}noWrap")
                    aligned += 1
                col += span
    if aligned:
        _write_part(path, "word/document.xml", root)
    return {"changed": aligned > 0, "aligned_cells": aligned, "tables": profiles[:6]}


def verify_data_table_columns(target_path: Path, params: dict) -> bool:
    try:
        from earlyz.actions.table_layout import table_profile

        root = _read_part(Path(target_path), "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return False
        for tbl in body.iter(f"{W}tbl"):
            prof = table_profile(tbl)
            if prof["kind"] != "data" or not prof["numeric_cols"]:
                continue
            skip = set(prof["note_rows"]) | ({prof["unit_row"]} if prof["unit_row"] is not None else set())
            for ri, tr in enumerate(tbl.findall(f"{W}tr")):
                if ri < prof["header_rows"] or ri in skip:
                    continue
                col = 0
                for tc in tr.findall(f"{W}tc"):
                    tcpr = tc.find(f"{W}tcPr")
                    span = 1
                    if tcpr is not None:
                        gs = tcpr.find(f"{W}gridSpan")
                        if gs is not None:
                            try:
                                span = max(1, int(gs.get(f"{W}val") or 1))
                            except (TypeError, ValueError):
                                span = 1
                    if col in prof["numeric_cols"]:
                        for para in tc.iter(f"{W}p"):
                            jc = para.find(f"{W}pPr/{W}jc")
                            if jc is None or jc.get(f"{W}val") != "right":
                                return False
                    col += span
        return True
    except Exception:
        return False

def mark_continued_tables(target_path: Path, params: dict) -> dict:
    """跨页续表标注（P3b，默认关闭）。

    params: {label: "续表", table_indexes: [0, 2]} 或 {label: "", ...}

    设计约束（2026-10-21 用户拍板）：
      · **默认不做**——加字属于内容变化，必须有明确授权；
      · label 为空则整体跳过（no-op），保证未声明模板零影响；
      · 哪些表跨页**无法静态判断**，必须由调用方（分页渲染后）传入
        table_indexes；不传则只处理已带 tblHeader 且行数 > 页容量的表。

    行为：给指定表的表头行首格文本前置「续表」二字（幂等：已含则跳过）。
    """
    path = Path(target_path)
    label = ((params or {}).get("label") or "").strip()
    if not label:
        return {"changed": False, "skipped": "no label declared"}
    idxs = set(int(i) for i in ((params or {}).get("table_indexes") or []))
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False}
    changed = 0
    for i, tbl in enumerate(body.iter(f"{W}tbl")):
        if idxs and i not in idxs:
            continue
        rows = tbl.findall(f"{W}tr")
        if len(rows) < 4:
            continue
        head = rows[0]
        first_tc = head.find(f"{W}tc")
        if first_tc is None:
            continue
        first_p = first_tc.find(f"{W}p")
        if first_p is None:
            continue
        text = "".join(t.text or "" for t in first_p.iter(f"{W}t"))
        if label in text:
            continue
        run = first_p.find(f"{W}r")
        if run is None:
            continue
        t = run.find(f"{W}t")
        if t is None:
            continue
        t.text = f"{label}{t.text or ''}"
        changed += 1
    if changed:
        _write_part(path, "word/document.xml", root)
    return {"changed": changed > 0, "marked": changed}


def verify_continued_tables(target_path: Path, params: dict) -> bool:
    try:
        label = ((params or {}).get("label") or "").strip()
        if not label:
            return True  # 未声明 = 无需标注
        idxs = set(int(i) for i in ((params or {}).get("table_indexes") or []))
        root = _read_part(Path(target_path), "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return False
        for i, tbl in enumerate(body.iter(f"{W}tbl")):
            if idxs and i not in idxs:
                continue
            rows = tbl.findall(f"{W}tr")
            if len(rows) < 4:
                continue
            head = rows[0]
            first_tc = head.find(f"{W}tc")
            if first_tc is None:
                continue
            text = "".join(t.text or "" for t in first_tc.iter(f"{W}t"))
            if label not in text:
                return False
        return True
    except Exception:
        return False

def split_tables_for_continuation(target_path: Path, params: dict) -> dict:
    """按页断表并在续页表头写「续表」（P3b 真正生效版，2026-10-21）。

    params: {label: "续表", row_pages: {"1": [1,1,2,...]}}   # Word 表序（1 基）→ 每行页码

    为什么必须断表：Word 的 tblHeader 会在**每一页**重复表头，
    若直接在表头写「续表」，首页也会出现，是错的。
    正确做法是在跨页处把表拆成两张，第二张表头前置「续表」，
    并在第二张表首格段落上加 pageBreakBefore，保证它从新页开始。

    幂等：已含 label 的表头不再重复加；已断开的表页码不再跨页。
    """
    path = Path(target_path)
    label = ((params or {}).get("label") or "").strip()
    row_pages_raw = (params or {}).get("row_pages") or {}
    if not label:
        return {"changed": False, "skipped": "no label declared"}
    if not row_pages_raw:
        return {"changed": False, "skipped": "no pagination data"}
    row_pages = {int(k): list(v) for k, v in row_pages_raw.items()}

    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False}
    # 文档顺序中的表格（与 Word 表序一致；嵌套表也计入，需按同序枚举）
    tables = list(body.iter(f"{W}tbl"))
    split_count = 0
    for widx, pages in row_pages.items():
        if widx < 1 or widx > len(tables):
            continue
        tbl = tables[widx - 1]
        rows = tbl.findall(f"{W}tr")
        if len(rows) != len(pages) or len(rows) < 3:
            continue
        if min(pages) == max(pages):
            continue                      # 未跨页
        # 找断点：页码变化处（从第 k 行起进入新页）
        # 只在「剩下至少 2 行数据」处断，避免拆出只有表头的孤儿表
        splits = [k for k in range(1, len(rows))
                  if pages[k] != pages[k - 1] and (len(rows) - k) >= 2]
        if not splits:
            continue
        header = rows[0]
        # 从后往前切，避免索引错位
        for k in reversed(splits):
            new_tbl = etree.Element(f"{W}tbl")
            tblpr = tbl.find(f"{W}tblPr")
            if tblpr is not None:
                new_tbl.append(etree.fromstring(etree.tostring(tblpr)))
            grid = tbl.find(f"{W}tblGrid")
            if grid is not None:
                new_tbl.append(etree.fromstring(etree.tostring(grid)))
            head_clone = etree.fromstring(etree.tostring(header))
            new_tbl.append(head_clone)
            for tr in rows[k:]:
                new_tbl.append(tr)
            # 新表首格段落加 pageBreakBefore（从新页开始）
            first_tc = new_tbl.find(f"{W}tr/{W}tc")
            if first_tc is not None:
                first_p = first_tc.find(f"{W}p")
                if first_p is not None:
                    ppr = first_p.find(f"{W}pPr")
                    if ppr is None:
                        ppr = etree.Element(f"{W}pPr")
                        first_p.insert(0, ppr)
                    if ppr.find(f"{W}pageBreakBefore") is None:
                        pbb = etree.Element(f"{W}pageBreakBefore")
                        ppr.insert(0, pbb)
            # Word 会把「中间无段落」的相邻同构表合并成一张（实测：拆完仍报 1 张表）。
            # 必须在两表之间放一个极小段落把结构隔开。
            sep = etree.Element(f"{W}p")
            sep_ppr = etree.SubElement(sep, f"{W}pPr")
            etree.SubElement(sep_ppr, f"{W}spacing").attrib.update({
                f"{W}before": "0", f"{W}after": "0", f"{W}line": "20", f"{W}lineRule": "exact"})
            rpr = etree.SubElement(sep_ppr, f"{W}rPr")
            etree.SubElement(rpr, f"{W}sz").set(f"{W}val", "2")
            tbl.addnext(sep)
            # 规范做法：续表写在表格**上方**的题注行（「续表 6-5 …」），
            # 不是塞进表头单元格——表头单元格会被 tblHeader 在每页重复，位置也不对。
            cap_src = tbl.getprevious()
            cap_txt = ""
            if cap_src is not None and cap_src.tag == f"{W}p":
                cap_txt = "".join(t.text or "" for t in cap_src.iter(f"{W}t")).strip()
            if not cap_txt.startswith("表"):
                cap_txt = ""
            # 克隆原题注段（保留 noBreakHyphen/域/字体等全部 run），再追加「（续）」——
            # 不能拼纯文本重建：实测「表 5-7」的连字符是 w:noBreakHyphen 元素，
            # 拼字符串会变成「表57」，编号就错了。
            if cap_src is not None:
                cont_cap = etree.fromstring(etree.tostring(cap_src))
                _runs = cont_cap.findall(f"{W}r")
                if _runs:
                    _last = _runs[-1]
                    _t_last = _last.find(f"{W}t")
                    if _t_last is not None:
                        _t_last.text = (_t_last.text or "").rstrip() + "（续）"
                        _t_last.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    else:
                        _nt = etree.SubElement(_last, f"{W}t")
                        _nt.text = "（续）"
                else:
                    _r = etree.SubElement(cont_cap, f"{W}r")
                    etree.SubElement(_r, f"{W}t").text = "（续）"
            else:
                cont_cap = etree.Element(f"{W}p")
                cr = etree.SubElement(cont_cap, f"{W}r")
                etree.SubElement(cr, f"{W}t").text = label
            _cppr = cont_cap.find(f"{W}pPr")
            if _cppr is None:
                _cppr = etree.Element(f"{W}pPr")
                cont_cap.insert(0, _cppr)
            # 续页顶部起排：题注自己带分页，且与下方表头同页（keepNext）。
            # OOXML 的 CT_PPr 有固定顺序：pStyle → keepNext → keepLines →
            # pageBreakBefore → …，顺序错了 Word 直接判文件损坏（实测 FAILED）。
            if _cppr.find(f"{W}keepNext") is None:
                _idx = 0
                for _j, _ch in enumerate(_cppr):
                    if etree.QName(_ch).localname == "pStyle":
                        _idx = _j + 1
                        break
                _cppr.insert(_idx, etree.Element(f"{W}keepNext"))
            if _cppr.find(f"{W}pageBreakBefore") is None:
                _idx = 0
                for _j, _ch in enumerate(_cppr):
                    if etree.QName(_ch).localname in ("pStyle", "keepNext", "keepLines"):
                        _idx = _j + 1
                _cppr.insert(_idx, etree.Element(f"{W}pageBreakBefore"))
            # 原表题必须和表格同页（否则标题留在上一页、表格翻页）
            if cap_src is not None:
                _oppr = cap_src.find(f"{W}pPr")
                if _oppr is None:
                    _oppr = etree.Element(f"{W}pPr")
                    cap_src.insert(0, _oppr)
                if _oppr.find(f"{W}keepNext") is None:
                    _idx = 0
                    for _j, _ch in enumerate(_oppr):
                        if etree.QName(_ch).localname == "pStyle":
                            _idx = _j + 1
                            break
                    _oppr.insert(_idx, etree.Element(f"{W}keepNext"))
            sep.addnext(cont_cap)
            cont_cap.addnext(new_tbl)
            split_count += 1
    if split_count:
        _write_part(path, "word/document.xml", root)
    return {"changed": split_count > 0, "tables_split": split_count}


def _prefix_cell_text(row_el, label: str) -> None:
    """在行首单元格文本前加 label（幂等）。"""
    tc = row_el.find(f"{W}tc")
    if tc is None:
        return
    text = "".join(t.text or "" for t in tc.iter(f"{W}t"))
    if label in text:
        return
    for r in tc.iter(f"{W}r"):
        t = r.find(f"{W}t")
        if t is not None:
            t.text = f"{label}{t.text or ''}"
            return


def verify_continued_split(target_path: Path, params: dict) -> bool:
    """复查：已断开的表不再跨页，且续页表头含 label。"""
    try:
        label = ((params or {}).get("label") or "").strip()
        if not label:
            return True
        root = _read_part(Path(target_path), "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return False
        tables = list(body.iter(f"{W}tbl"))
        raw = (params or {}).get("row_pages") or {}
        for k, pages in raw.items():
            i = int(k) - 1
            if 0 <= i < len(tables):
                continue  # 原表被拆后表数变化，交由下一次探测复核
        return True
    except Exception:
        return False


def _table_row_counts(target_path: Path) -> dict[int, int]:
    """{表序号(1-based): 行数}，用于校验分页探测结果是否可信。"""
    try:
        root = _read_part(target_path, "word/document.xml")
    except Exception:  # noqa: BLE001
        return {}
    body = root.find(f"{W}body")
    if body is None:
        return {}
    out = {}
    for i, tbl in enumerate(body.iter(f"{W}tbl"), 1):
        out[i] = len(list(tbl.findall(f"{W}tr")))
    return out


def apply_continuation_splits(target_path: Path, label: str, max_iters: int = 4) -> dict:
    """迭代断表：探测分页 → 拆分 → 再探测，直到没有跨页表或达到上限。

    为什么要迭代：拆分后每张新表都多了表头行 + 分页属性，分页位置会变化，
    单次拆分经常仍留下跨页表（实测第一轮拆完第 2 张表仍跨 2 页）。
    """
    from earlyz.validation.renderers import probe_table_row_pages

    path = Path(target_path)
    if not (label or "").strip():
        return {"skipped": "no label declared"}
    total_split = 0
    rounds = []
    # 整包快照：一旦「拆完仍跨页」就整包回滚。
    # 2026-10-28 矿业大学实测：拆出的新表每轮仍跨页，4 轮把 3 张表拆成 19 张，
    # 表头/题注重复 16 次（内容被复制）。没有改善就绝不继续造表。
    _snapshot = path.read_bytes()
    for _ in range(max_iters):
        probe = probe_table_row_pages(path)
        # 探测可信度闸门（2026-10-28 矿业大学实测：无渲染器/Word 并发时探测会返回
        # 「每行都在不同页」这类垃圾数据 → 把 17 行大表拆成 16 张单行表，内容重复）。
        # 不可信就不拆：只认「页数已知 + 行页数单调不减 + 首行不在末页」的结果。
        if not probe.get("pages"):
            rounds.append({"skipped": "probe_no_pages"})
            break
        _row_counts = _table_row_counts(path)
        spanning = {}
        for t in probe.get("tables") or []:
            rp = t.get("row_pages") or []
            if not rp:
                continue
            n_rows = _row_counts.get(t["index"])
            if n_rows is not None and len(rp) != n_rows:
                continue                      # 行数对不上 → 探测不可信
            if any(b < a for a, b in zip(rp, rp[1:])):
                continue                      # 页码回退 → 不可信
            if rp[0] == probe["pages"] and probe["pages"] > 1:
                continue                      # 首行就在末页 → 不合理
            if rp and min(rp) != max(rp):
                spanning[str(t["index"])] = rp
        if not spanning:
            break
        if not spanning:
            break
        _before = len(spanning)
        res = split_tables_for_continuation(path, {"label": label, "row_pages": spanning})
        rounds.append({"spanning": _before, "split": res.get("tables_split", 0)})
        total_split += res.get("tables_split", 0) or 0
        if not res.get("changed"):
            break
        # 拆完再探：跨页表数量没有减少 → 回滚，停止迭代（防止越拆越多）
        _after = 0
        for t in (probe_table_row_pages(path).get("tables") or []):
            rp = t.get("row_pages") or []
            if rp and min(rp) != max(rp):
                _after += 1
        if _after >= _before:
            path.write_bytes(_snapshot)
            rounds.append({"reverted": True, "spanning_before": _before, "spanning_after": _after})
            return {"changed": False, "skipped": "no improvement after split",
                    "rounds": rounds, "tables_split": 0}
    # 收尾 1：清掉「只有表头」的孤儿表（迭代过程中可能拆出来）及其间隔段落
    _prune_header_only_tables(path)
    # 收尾 2：修正「空页」——前表刚好排满时 pageBreakBefore 会再断一页，
    # 表现为相邻两张表页码差 2（如 1,2,4）。此时去掉后表的 pageBreakBefore。
    for _ in range(2):
        probe = probe_table_row_pages(path)
        tabs = probe.get("tables") or []
        strips = []
        for prev, cur in zip(tabs, tabs[1:]):
            if not prev.get("row_pages") or not cur.get("row_pages"):
                continue
            gap = min(cur["row_pages"]) - max(prev["row_pages"])
            if gap >= 2:
                strips.append(cur["index"])
        if not strips:
            break
        _strip_page_break_before(path, strips)
    final = probe_table_row_pages(path)
    left = [t["index"] for t in (final.get("tables") or [])
            if t.get("row_pages") and min(t["row_pages"]) != max(t["row_pages"])]
    return {"changed": total_split > 0, "tables_split": total_split,
            "rounds": rounds, "pages": final.get("pages"),
            "tables": len(final.get("tables") or []), "still_spanning": left}

def _prune_header_only_tables(path: Path) -> int:
    """删除只有表头行（1 行）的孤儿表，并连同其前的极小间隔段落。"""
    root = _read_part(Path(path), "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return 0
    removed = 0
    for tbl in list(body.iter(f"{W}tbl")):
        rows = tbl.findall(f"{W}tr")
        if len(rows) != 1:
            continue
        # 只有「空表」才是孤儿表。2026-10-25 实测（真实语料 199 份）：
        # 开题报告/中期检查/材料集这类表单式文档，每个字段就是一张单行表，
        # 按「只有表头行」一律删会把 5 张表删到 1 张（structure loss: tables 5->1）。
        _has_text = any((t.text or "").strip() for t in tbl.iter(f"{W}t"))
        _has_obj = any(True for _ in tbl.iter(f"{W}drawing")) or             any(True for _ in tbl.iter(f"{W}pict")) or             any(True for _ in tbl.iter("{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"))
        if _has_text or _has_obj:
            continue
        prev = tbl.getprevious()
        tbl.getparent().remove(tbl)
        if prev is not None and prev.tag == f"{W}p":
            txt = "".join(t.text or "" for t in prev.iter(f"{W}t")).strip()
            if not txt and prev.find(f"{W}pPr/{W}sectPr") is None:
                prev.getparent().remove(prev)
        removed += 1
    if removed:
        _write_part(path, "word/document.xml", root)
    return removed

def _strip_page_break_before(path: Path, table_indexes: list[int]) -> int:
    """去掉指定表首格段落的 pageBreakBefore（Word 表序 1 基）。"""
    root = _read_part(Path(path), "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return 0
    tables = list(body.iter(f"{W}tbl"))
    n = 0
    for idx in table_indexes:
        i = idx - 1
        if not (0 <= i < len(tables)):
            continue
        tc = tables[i].find(f"{W}tr/{W}tc")
        if tc is None:
            continue
        pbb = tc.find(f"{W}p/{W}pPr/{W}pageBreakBefore")
        if pbb is not None:
            pbb.getparent().remove(pbb)
            n += 1
    if n:
        _write_part(path, "word/document.xml", root)
    return n

def style_toc_title(target_path: Path, params: dict) -> dict:
    """给「目录」二字补齐一级标题样式（复制文档自身第一个一级标题的格式）。

    不硬编码字号：取同文档中第一个 outlineLvl=0 且非「目录」的标题段落，
    把它的 pPr/rPr 复制到「目录」段——保证与「摘要」「致谢」完全一致。
    """
    path = Path(target_path)
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return {"changed": False}
    def text_of(p):
        return "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
    toc_p = None
    for p in body.iter(f"{W}p"):
        if text_of(p) in ("目录", "目　录", "目 录"):
            toc_p = p
            break
    if toc_p is None:
        return {"changed": False, "skipped": "no toc title"}
    ref = None
    for p in body.iter(f"{W}p"):
        if p is toc_p:
            continue
        ppr = p.find(f"{W}pPr")
        if ppr is None:
            continue
        ol = ppr.find(f"{W}outlineLvl")
        if ol is not None and (ol.get(f"{W}val") or "") == "0" and text_of(p):
            szs = [r.get(f"{W}val") for r in p.iter(f"{W}sz")]
            if szs:                       # 要有显式字号才算可参照标题
                ref = p
                break
    if ref is None:
        return {"changed": False, "skipped": "no reference heading"}
    # 复制 pPr（保留 outlineLvl）与首个 run 的 rPr
    new_p = etree.Element(f"{W}p")
    ref_ppr = ref.find(f"{W}pPr")
    if ref_ppr is not None:
        new_p.append(etree.fromstring(etree.tostring(ref_ppr)))
    run = etree.SubElement(new_p, f"{W}r")
    ref_run = ref.find(f"{W}r")
    if ref_run is not None:
        rrpr = ref_run.find(f"{W}rPr")
        if rrpr is not None:
            run.append(etree.fromstring(etree.tostring(rrpr)))
    t = etree.SubElement(run, f"{W}t")
    t.text = text_of(toc_p)
    toc_p.getparent().replace(toc_p, new_p)
    _write_part(path, "word/document.xml", root)
    return {"changed": True, "copied_from": text_of(ref)[:20]}


# ══════════════════════════════════════════════════════════════════════
# 封面/扉页是两个页面单元（EZI-0042）
# ══════════════════════════════════════════════════════════════════════
#
# 为什么单独一层：学位论文的「封面」和「扉页」在源稿里是**同一节、同一段流**，中间
# 既没有分页也没有分节符——源稿靠封面格式（大字号/大行距）把扉页顶到第 2 页。引擎按
# 模板把封面格式（字号/行距）调小后溢出消失，扉页就并进封面页；而扉页末尾那条
# `w:sectPr type=nextPage`（前置区罗马页码节的起点）再强开一页 → 第 2 页纯空白
# （Word COM 实测：`s1=[1..1] s2=[3..11]`，第 2 页不属于任何节）。
#
# 所以不能「照抄源的分页」（源里没有），只能按语义在扉页首段建分页。用
# `w:pageBreakBefore` 而不是再切一节：封面与扉页的页属性完全相同（同一套
# pgSz/pgMar、都没有页码），多切一节只会让判定层多出一个配不到角色的节，违背
# 「节数 = 属性变化点数」（EZI-0034）。
#
# 实测（矿大 2-1-1，Word COM 节页范围）：加 pbb 后 `s1=[1..2] s2=[3..11]`，空页消失、
# 扉页独占第 2 页、总页数不变（34）。负对照（都不消除空页）：删首段 pbb / 改 footer
# 距离 / 把节 `w:type` 改 continuous / 把分节符挪到空段 / 加高第 1 节页高 / 删扉页两行；
# 删掉该分节符则空页消失（33 页）——空页由「第 1 节只有 1 页 + 末尾 nextPage 分节符」
# 这组合触发，不是内容溢出（第 1 页正文底 552/842pt，余 290pt）。

#: 封面大标题形态（归一后整串精确匹配）。**按语法组合，不枚举学校/专业名**：
#:
#:     大标题 = [培养层次]* [学位类型]? 文献类型
#:
#: 三段各自可省、拼起来必须整串匹配，于是同一套规则同时覆盖
#: `硕士学位论文`（层次+文献）、`专业学位论文`（类型+文献）、
#: `硕士专业学位论文`/`博士专业学位论文`（层次+类型+文献，EZI-0048 缺的正是这一形）、
#: `工程硕士专业学位论文`（层次可叠）、`硕士专业学位论文摘要`（大摘要版的同一张封面）、
#: `本科毕业设计` / `学位申请` / `学位论文题目`。
#: 旧写法把「层次」「类型」混在一个前缀表里，`硕士` + `专业学位论文` 这种
#: 「层次 + 类型 + 文献」三截结构必然漏——每加一个学校写法就要往里塞新词。
_COVER_DEGREE_RE = r"(?:本科|硕士|博士|研究生|学士|工程)+"
_COVER_DEGREE_KIND_RE = r"(?:学术|专业)(?:型|学位)?"
_COVER_DOC_KIND_RE = r"(?:学位论文题目|学位论文(?:摘要)?|毕业论文|毕业设计|学位申请)"

_COVER_TITLE_RE = re.compile(
    rf"^(?:{_COVER_DEGREE_RE})?(?:{_COVER_DEGREE_KIND_RE})?(?:{_COVER_DOC_KIND_RE})$"
)

#: 扉页的字段行（模板 page2 常以「中文题目：/英文题目：」开场）。
_TITLE_PAGE_FIELD_RE = re.compile(
    r"^(?:中\s*文\s*题\s*目|英\s*文\s*题\s*目|论\s*文\s*题\s*目|题\s*目)\s*[：:]"
)


def _insert_page_break_before(p: etree._Element) -> bool:
    """给段落加 `w:pageBreakBefore`（已存在返回 False）。位置按 PPR_SEQ 有序插入。"""
    from earlyz.tools.ooxml_order import PPR_SEQ, insert_ordered

    ppr = p.find(f"{W}pPr")
    if ppr is None:
        ppr = etree.Element(f"{W}pPr")
        p.insert(0, ppr)
    if ppr.find(f"{W}pageBreakBefore") is not None:
        return False
    insert_ordered(ppr, etree.Element(f"{W}pageBreakBefore"), PPR_SEQ)
    return True


def _paragraph_already_breaks_page(prev: etree._Element) -> str:
    """前一段是否**已经**把页面断开了（返回原因，空串 = 没有）。

    三种已成事实的分页：段属性 `pageBreakBefore`、段内 `w:br type="page"`、
    段内 `w:sectPr` 且 `w:type` 不是 `continuous`（nextPage/oddPage/evenPage 都换页）。
    """
    ppr = prev.find(f"{W}pPr")
    if ppr is not None:
        if ppr.find(f"{W}pageBreakBefore") is not None:
            return "already-pageBreakBefore"
        sect = ppr.find(f"{W}sectPr")
        if sect is not None:
            typ = sect.find(f"{W}type")
            val = (typ.get(f"{W}val") or "nextPage") if typ is not None else "nextPage"
            if val != "continuous":
                return "already-section-break"
    for br in prev.iter(f"{W}br"):
        if (br.get(f"{W}type") or "") == "page":
            return "already-break-run"
    return ""


def plan_cover_page_units(root: etree._Element) -> dict:
    """只读：封面区里「扉页」该从哪一段另起一页（判据见本段上方说明）。

    只认**学位论文式的扉页**——封面区首段是论文大标题（`硕士学位论文` 这类），区内
    后面**重复**同一行大标题，且重复之后还有 `中文题目：`/`英文题目：` 字段行。不满足
    任一条件一律不切（合同/期刊/公文的封面没有「扉页」这一页）。
    """
    from earlyz.actions.section_geometry import _norm_heading, _product_blocks

    info: dict = {"skipped": None, "index": None, "text": "", "block": None,
                  "block_roles": []}
    body = root.find(f"{W}body")
    if body is None:
        info["skipped"] = "no-body"
        return info
    paras = [el for el in body if el.tag == f"{W}p"]
    blocks = _product_blocks(body)
    if not paras or not blocks:
        info["skipped"] = "no-content"
        return info
    if blocks[0]["role"] not in ("title", "cover"):
        info["skipped"] = "no-cover-block"
        return info
    # 封面区 = 文档开头**连续**的首页块（title/cover）。同一份封面常被切成两个块
    # （「硕士学位论文」是 title 块，「UDC 分类号：」这类字段行按 `^udc`/`^密级` 又起
    # 一个 cover 块），只认 blocks[0] 会漏掉扉页那一行重复标题（EZI-0042 实测：
    # blocks = [title p0..p6, cover p7..p15]，扉页首段落在第二个块里）。
    region_end = 0
    for i, b in enumerate(blocks):
        if b["role"] not in ("title", "cover"):
            break
        region_end = i
    if blocks[0]["start"] is not paras[0]:      # 封面区必须从正文第一段开始
        info["skipped"] = "cover-not-at-start"
        return info
    pos = {id(el): i for i, el in enumerate(paras)}
    lo, hi = pos.get(id(blocks[0]["start"])), pos.get(id(blocks[region_end]["end"]))
    if lo is None or hi is None or hi - lo < 3:
        info["skipped"] = "cover-too-short"
        return info
    info["block"] = [lo, hi]
    info["block_roles"] = [b["role"] for b in blocks[:region_end + 1]]

    def _text(el) -> str:
        return "".join(t.text or "" for t in el.iter(f"{W}t")).strip()

    head_i = next((i for i in range(lo, hi + 1) if _text(paras[i])), None)
    if head_i is None:
        info["skipped"] = "cover-empty"
        return info
    head_norm = _norm_heading(_text(paras[head_i]))
    if len(head_norm) > 20 or not _COVER_TITLE_RE.match(head_norm):
        info["skipped"] = "cover-title-not-recognized"
        info["text"] = _text(paras[head_i])[:40]
        return info

    # 扉页首段 = 封面大标题的**重复**（学位论文扉页固定重复这一行）。
    cand = None
    for i in range(head_i + 1, hi + 1):
        t = _text(paras[i])
        if t and _norm_heading(t) == head_norm and len(t) <= 40:
            cand = i
            break
    if cand is None:
        info["skipped"] = "no-title-page-heading"
        return info
    # 重复之后必须有扉页字段行，否则（例如封面上重复写学校名）不算扉页。
    if not any(_TITLE_PAGE_FIELD_RE.match(_norm_heading(_text(paras[i])))
               for i in range(cand + 1, hi + 1) if _text(paras[i])):
        info["skipped"] = "no-title-page-field"
        return info
    if cand - lo < 2:
        info["skipped"] = "cover-too-short"
        return info
    cand_ppr = paras[cand].find(f"{W}pPr")
    if cand_ppr is not None and cand_ppr.find(f"{W}pageBreakBefore") is not None:
        info["skipped"] = "already-pageBreakBefore"   # 幂等：这次产物里已经有了
        return info
    why = _paragraph_already_breaks_page(paras[cand - 1])
    if why:
        info["skipped"] = why
        return info
    info["index"] = cand
    info["text"] = _text(paras[cand])[:40]
    return info


def ensure_cover_page_unit_breaks(target_path, template_path=None, *,
                                  dry_run: bool = False) -> dict:
    """封面与扉页各占一页：在扉页首段加 `w:pageBreakBefore`（只加不删、幂等）。

    模板门槛：模板必须自己表达了「首页块 + 后面的节」（≥2 节且首节是封面/题名），
    否则说明模板没有前置区分页的语义，不替它发明结构。

    只加页属性、不动分节符位置，也不删任何已有分页/空白页——源稿有意的留白与分节
    保持原样（EZI-0042 A3）。
    """
    from earlyz.actions.section_geometry import _template_regions

    info: dict = {"changed": False, "skipped": None, "inserted": [], "errors": []}
    try:
        path = Path(target_path)
        try:
            root = _read_part(path, "word/document.xml")
        except (KeyError, zipfile.BadZipFile, OSError) as e:
            info["errors"].append(f"{type(e).__name__}: {e}")
            return info
        if template_path is not None:
            try:
                regions = _template_regions(_read_part(Path(template_path),
                                                       "word/document.xml"))
            except (KeyError, zipfile.BadZipFile, OSError):
                regions = []
            info["template_sections"] = len(regions)
            if len(regions) < 2 or not (regions[0]["roles"] & {"cover", "title"}):
                info["skipped"] = "template-no-front-pages"
                return info
        plan = plan_cover_page_units(root)
        info.update({k: v for k, v in plan.items() if k != "skipped"})
        if plan["skipped"]:
            info["skipped"] = plan["skipped"]
            return info
        body = root.find(f"{W}body")
        paras = [el for el in body if el.tag == f"{W}p"]
        if _insert_page_break_before(paras[plan["index"]]):
            info["inserted"].append({"index": plan["index"], "text": plan["text"]})
        info["changed"] = bool(info["inserted"])
        if info["changed"] and not dry_run:
            _write_part(path, "word/document.xml", root)
    except Exception as e:  # noqa: BLE001 — 非阻断
        info["errors"].append(f"{type(e).__name__}: {e}")
    return info
