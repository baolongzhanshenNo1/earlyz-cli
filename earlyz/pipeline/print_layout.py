"""早睡 (EarlyZ) 3.2 — 打印版最终化（新链路入口，§12 Phase 1 工作包 9）。

在既有 enforce_print_layout（Word COM 分页循环，已验证）之上：
- find_body_start_section：用 Analyzer IR 的 heading 语义识别"正文第一个章节"
  所在 section，兼容 `第1章绪论` / `第一章` / `1 绪论` / `1.1 标题` 等格式，
  替代旧链路的 `^第X章` 硬正则（uploads 真实论文均为 `1 绪论` 格式）；
- apply_print_layout：传入 IR 识别的 body_start，其余复用验证过的
  奇数页循环 / 空白页插入 / 页码体系 / titlePg 逻辑。
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Optional
from lxml import etree

from earlyz.ir.analyzer import analyze_document
from earlyz.ir.models import DocumentIR
from earlyz.ir.selectors import normalize_text
from pipeline.print_finalizer import (
    _insert_sectpr_child,
    _section_start_paragraphs,
    _text,
    enforce_print_layout,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# 数字标题格式：`1 绪论` / `1.1 研究背景` / `1、绪论`
_NUM_TITLE_RE = re.compile(r"^\d+(\.\d+)*[\s、.．]+[^\d\s]")


def find_body_start_section(docx_path: str | Path, ir: Optional[DocumentIR] = None) -> Optional[int]:
    """返回正文第一个章节所在 section index；找不到返回 None。

    识别顺序（§4.3 身份恢复精神）：
    1. IR：第一个 heading_level==1 段落；
    2. 文本正则：`第X章` / `第一章` / 数字标题（`1 绪论` / `1.1 x`）。
    """
    docx_path = Path(docx_path)
    if ir is None:
        ir = analyze_document(docx_path)

    # 1) IR 语义：第一个一级标题段落
    h1 = next((p for p in ir.paragraphs if p.heading_level == 1), None)
    if h1 is None:
        # 宽松兜底：任意 heading 且文本匹配章节模式
        h1 = next(
            (p for p in ir.paragraphs if p.is_heading and _looks_like_chapter(p.text)),
            None,
        )

    with zipfile.ZipFile(docx_path) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    starts = _section_start_paragraphs(root)
    if not starts:
        return None

    if h1 is not None:
        norm_h1 = normalize_text(h1.text)
        for i, p in enumerate(starts):
            if normalize_text(_text(p)) == norm_h1:
                return i
        # 章节标题不是 section 起始段：找其前面最近的 section
        # （正文第一章可能从"1 绪论"标题段落开始，section 起始在其前一段）
        return _nearest_section_before(root, starts, h1.text)

    # 2) 文本正则兜底
    for i, p in enumerate(starts):
        if _looks_like_chapter(_text(p)) and i > 0:
            return i
    return None


def _looks_like_chapter(text: str) -> bool:
    stripped = re.sub(r"\s+", "", text)  # 第X章格式去空白匹配
    if re.match(r"^第[一二三四五六七八九十百\d]+章", stripped):
        return True
    # 数字标题格式必须在原始文本上匹配（`1 绪论` 保留空格分隔）
    if _NUM_TITLE_RE.match(text.strip()):
        return True
    return False


def _nearest_section_before(root: etree._Element, starts: list, heading_text: str) -> Optional[int]:
    """heading 段落之前最近的 section 起始段（body_start 取该节）。"""
    body = root.find(f"{{{W}}}body")
    if body is None:
        return None
    norm = normalize_text(heading_text)
    section_of: dict[int, int] = {}  # 段落对象 id → section index
    for i, p in enumerate(starts):
        section_of[id(p)] = i
    last_section = 0
    for child in body:
        if child.tag == f"{W}p":
            p = child
            if id(p) in section_of:
                last_section = section_of[id(p)]
            if normalize_text(_text(p)) == norm:
                return last_section
    return None


def _align_abstract_headers(docx_path: str | Path, body_start: Optional[int] = None) -> dict:
    """通用页眉对齐（判断能力，不硬编码学校要求）。

    规则（从文档自身结构推导）：
    1. evenAndOddHeaders 开启时，任何节缺 even 引用 → 克隆 default part 为独立
       even part（同一 footer part 被 default+even 双引用时 Word 判文档损坏，
       header 同 part 双引用虽被容忍但非 canonical——Word 自修复亦建独立 part）。
    2. 非正文节的节首段为语义章标题（摘要/ABSTRACT/目录/致谢/参考文献/附录/结论等）
       → 确保 default 页眉 = 节标题文本；摘要节中英文标题分居奇偶页时
       even='ABSTRACT'/default='摘要'（与原反置交换逻辑等价且更通用）。
    3. 封面性质节（声明/授权书/封面/任务书等）→ 页眉内容清空并移除下划线边框。
    正文节（body_start 之后）只补 even，不设标题页眉（页眉由模板决定）。
    """
    from earlyz.ir.analyzer import _is_chapter_title_text
    from earlyz.actions.header_footer_actions import (
        _create_part, _bind_reference, _part_content, _header_content_type,
        _footer_content_type,
    )
    from earlyz.actions.new_actions import _read_part, _write_part

    WX = "{" + W + "}"
    RX = "{" + R + "}"

    # 封面性质词：页眉应为空且无下划线
    _FRONT_COVER_WORDS = {
        "封面", "学位论文原创性声明", "学位论文版权使用授权书", "任务书",
        "开题报告", "评阅意见", "答辩记录", "答辩委员会", "学校代码",
        "密级", "中图分类号",
    }
    # 封面特征词：节内任一段含其一即视为封面性质（节首段可能不是封面词）
    _COVER_FEATURE_WORDS = (
        "专业名称", "学生姓名", "学生学号", "导师姓名", "申请学位",
        "论文提交日期", "毕业设计（论文）", "分类号",
    )

    root = _read_part(docx_path, "word/document.xml")
    body = root.find(f"{WX}body")
    if body is None:
        return {"changed": False, "sections": 0}
    rels_root = _read_part(docx_path, "word/_rels/document.xml.rels")
    rel_map: dict[str, str] = {}
    if rels_root is not None:
        for rel in rels_root:
            rel_map[rel.get("Id")] = rel.get("Target", "")

    def header_text(rel_id: str) -> str:
        tgt = rel_map.get(rel_id, "")
        full = f"word/{tgt}" if tgt and not tgt.startswith("word/") else tgt
        try:
            h = _read_part(docx_path, full)
        except Exception:
            return ""
        return "".join(t.text or "" for t in h.iter(f"{WX}t"))

    # 节首段文本：Word 节标记段在节末（段落 pPr 内 sectPr），故取边界**后**
    # 第一个非空段落文本；若节内无文本则回退边界前文本
    section_first: dict[int, str] = {}
    section_texts: dict[int, list[str]] = {}  # 节内全部语义标题候选
    pending_idx: int = 0  # 第 0 节从 body 开头开始
    sec_idx = 0
    cur_sec = 0
    for b in body:
        if b.tag == f"{WX}p":
            t = "".join(x.text or "" for x in b.iter(f"{WX}t")).strip()
            if t:
                if pending_idx is not None:
                    section_first.setdefault(pending_idx, t)
                    pending_idx = None
                section_texts.setdefault(cur_sec, []).append(t)
                last_text = t
            ppr = b.find(f"{WX}pPr")
            if ppr is not None and ppr.find(f"{WX}sectPr") is not None:
                section_first.setdefault(sec_idx, last_text)
                sec_idx += 1
                pending_idx = sec_idx
                cur_sec = sec_idx
        elif b.tag == f"{WX}sectPr":
            section_first.setdefault(sec_idx, last_text)
            sec_idx += 1
            pending_idx = sec_idx
            cur_sec = sec_idx

    # evenAndOddHeaders 是否开启（决定偶数页是否用 even 页眉）
    try:
        settings = _read_part(docx_path, "word/settings.xml")
        even_odd = settings.find(f"{WX}evenAndOddHeaders") is not None
    except Exception:
        even_odd = False

    def rel_of(sp, typ):
        for hr in sp.findall(f"{WX}headerReference"):
            if hr.get(f"{WX}type") == typ:
                return hr.get(f"{RX}id")
        return None

    def _rel_id_of(part_name: str):
        """按 part 文件名找 rel_id（_create_part 返回 part_name，需查 rels）。"""
        PRX = "{http://schemas.openxmlformats.org/package/2006/relationships}"
        rels2 = _read_part(docx_path, "word/_rels/document.xml.rels")
        base = part_name.split("/")[-1]
        for r in rels2.findall(f"{PRX}Relationship"):
            if (r.get("Target") or "").endswith(base):
                return r.get("Id")
        return None

    def set_part_text(rel_id, text, keep_border=True) -> bool:
        """更新页眉文本，保留原段落格式（pStyle 下划线/字号）。

        keep_border=False：封面性质页——移除下划线来源（显式 pBdr 或 header 样式引用）。
        """
        tgt = rel_map.get(rel_id, "")
        full = f"word/{tgt}" if tgt and not tgt.startswith("word/") else tgt
        if not full:
            return False
        h = _read_part(docx_path, full)
        p = h.find(f"{WX}p")
        if p is None:
            p = etree.SubElement(h, f"{WX}p")
        ppr = p.find(f"{WX}pPr")
        if not keep_border and ppr is not None:
            if ppr.find(f"{WX}pBdr") is not None:
                ppr.remove(ppr.find(f"{WX}pBdr"))
            ps = ppr.find(f"{WX}pStyle")
            if ps is not None:
                ppr.remove(ps)
        # 保留 pPr；复制第一个 run 的 rPr（字号）；替换文本
        old_rpr = None
        for r in list(p.findall(f"{WX}r")):
            if old_rpr is None:
                rpr = r.find(f"{WX}rPr")
                if rpr is not None:
                    old_rpr = rpr
            p.remove(r)
        if text:
            r = etree.SubElement(p, f"{WX}r")
            if old_rpr is not None:
                r.append(old_rpr)
            t = etree.SubElement(r, f"{WX}t")
            t.text = text
        _write_part(docx_path, full, h)
        return True

    def _clone_part(rel_id: str, prefix: str) -> str | None:
        """克隆 header/footer part 为独立新 part，返回新 rel_id。

        Word 判损坏规则（实测二分定位）：
        1. 同一 footer part 被 default+even 双引用 → 损坏（header 双引用被
           容忍但非 canonical）——故必须克隆独立 part；
        2. 克隆部件的 wp:docPr id 与原件重复 → 损坏（docPr id 须包级唯一，
           页脚页码文本框 shape 各带 docPr id）——故克隆时重编 docPr id。
        """
        tgt = rel_map.get(rel_id, "")
        full = f"word/{tgt}" if tgt and not tgt.startswith("word/") else tgt
        if not full:
            return None
        try:
            content = zipfile.ZipFile(docx_path).read(full)
        except Exception:
            return None
        # 重编 docPr id：克隆体不得与原件共用 shape id
        WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
        try:
            clone_root = etree.fromstring(content)
            doc_prs = clone_root.findall(f".//{WP}docPr")
        except Exception:
            clone_root, doc_prs = None, []
        if clone_root is not None and doc_prs:
            max_id = 0
            with zipfile.ZipFile(docx_path) as z:
                for n in z.namelist():
                    if n.startswith("word/") and n.endswith(".xml") and "_rels" not in n:
                        for el in etree.fromstring(z.read(n)).iter(f"{WP}docPr"):
                            try:
                                max_id = max(max_id, int(el.get("id", "0")))
                            except ValueError:
                                pass
            for el in doc_prs:
                max_id += 1
                el.set("id", str(max_id))
                if el.get("name"):
                    el.set("name", f"{el.get('name')}-{max_id}")
            content = etree.tostring(clone_root, xml_declaration=True,
                                     encoding="UTF-8", standalone=True)
        is_header = full.rsplit("/", 1)[-1].startswith("header")
        part_name = _create_part(
            Path(docx_path), "header" if is_header else "footer",
            _header_content_type() if is_header else _footer_content_type(),
            content,
        )
        new_rid = _rel_id_of(part_name)
        if new_rid is None:
            raise RuntimeError(f"cloned part 未注册 rel: {part_name}")
        # rel_map 增量更新（后续 header_text 等查询用）
        rel_map[new_rid] = part_name.split("/")[-1]
        return new_rid

    def _build_header_part(text: str) -> bytes:
        """新建页眉 part：以同文档现有页眉段落为格式模板（下划线/字号），仅换文本。"""
        template = None
        with zipfile.ZipFile(docx_path) as z:
            for n in sorted(z.namelist()):
                if n.startswith("word/header") and n.endswith(".xml"):
                    h = _read_part(docx_path, n)
                    if h.find(f"{WX}p") is not None:
                        template = h
                        break
        if template is None:
            return _part_content(text, False)
        tp = template.find(f"{WX}p")
        new_h = etree.Element(f"{WX}hdr")
        new_p = etree.SubElement(new_h, f"{WX}p")
        tppr = tp.find(f"{WX}pPr")
        if tppr is not None:
            new_p.append(etree.fromstring(etree.tostring(tppr)))
        new_r = etree.SubElement(new_p, f"{WX}r")
        tr = tp.find(f"{WX}r")
        if tr is not None:
            trpr = tr.find(f"{WX}rPr")
            if trpr is not None:
                new_r.append(etree.fromstring(etree.tostring(trpr)))
        t = etree.SubElement(new_r, f"{WX}t")
        t.text = text
        return etree.tostring(new_h, xml_declaration=True, encoding="UTF-8", standalone=True)

    def clear_footer_part(rel_id) -> bool:
        """把页脚 part 清空（无页码、无内容）——封面性质页要求。"""
        tgt = rel_map.get(rel_id, "")
        full = f"word/{tgt}" if tgt and not tgt.startswith("word/") else tgt
        if not full:
            return False
        ftr = etree.Element(f"{WX}ftr")
        etree.SubElement(ftr, f"{WX}p")
        _write_part(docx_path, full, ftr)
        return True

    def footer_has_page_number(rel_id) -> bool:
        """页脚 part 是否含页码域（PAGE）。"""
        tgt = rel_map.get(rel_id, "")
        full = f"word/{tgt}" if tgt and not tgt.startswith("word/") else tgt
        if not full:
            return False
        try:
            f = _read_part(docx_path, full)
        except Exception:
            return False
        return (f.find(f"{WX}fldSimple") is not None or
                f.find(f".//{WX}instrText") is not None)

    changed = 0
    sec_idx = 0
    for sp in body.iter(f"{WX}sectPr"):
        refs = {hr.get(f"{WX}type"): hr for hr in sp.findall(f"{WX}headerReference")}
        first = section_first.get(sec_idx, "")
        first_norm = re.sub(r"\s+", "", first)
        # 封面判定：节首段 ∈ 封面词表，或（正文前节）节内含封面特征词
        sec_texts = section_texts.get(sec_idx, [])
        is_cover = (
            first_norm in _FRONT_COVER_WORDS
            or (body_start is None or sec_idx < body_start)
            and any(w in t for t in sec_texts for w in _COVER_FEATURE_WORDS)
        )
        is_chapter_title = (not is_cover and
                            (_is_chapter_title_text(first) or
                             (body_start is not None and sec_idx < body_start and
                              re.match(r"^\d+(\.\d+)*\s", first))))
        even = refs.get("even")
        default = refs.get("default")

        if is_cover:
            # 封面性质节：页眉置空（去下划线）+ 页脚清空（无页码）
            for hr in (even, default):
                if hr is not None and header_text(hr.get(f"{RX}id")) != "":
                    if set_part_text(hr.get(f"{RX}id"), "", keep_border=False):
                        changed += 1
            for fr in sp.findall(f"{WX}footerReference"):
                if footer_has_page_number(fr.get(f"{RX}id")):
                    if clear_footer_part(fr.get(f"{RX}id")):
                        changed += 1
            sec_idx += 1
            continue

        if is_chapter_title:
            # 非正文节：页眉 = 节标题文本（每页一致）
            title = first_norm
            # 摘要节：中英文标题同节 → even = 与 default 不同的那个标题
            even_title = None
            if first_norm in ("摘要", "ABSTRACT"):
                texts = section_texts.get(sec_idx, [])
                for t in texts:
                    tn = re.sub(r"\s+", "", t)
                    if tn in ("摘要", "ABSTRACT") and tn != first_norm:
                        even_title = tn
                        break
            # 确保 default = 节标题
            if default is None:
                if even is not None and header_text(even.get(f"{RX}id")) == title:
                    # even part 内容已匹配 → 克隆为独立 default part（禁止同 rId 双引用）
                    new_rid = _clone_part(even.get(f"{RX}id"), "header")
                    if new_rid:
                        _bind_reference(docx_path, sec_idx, "headerReference", "default", new_rid)
                        changed += 1
                else:
                    part_name = _create_part(Path(docx_path), "header",
                                             _header_content_type(), _build_header_part(title))
                    rel_id = _rel_id_of(part_name)
                    if rel_id is None:
                        raise RuntimeError(f"header part 未注册 rel: {part_name}")
                    _bind_reference(docx_path, sec_idx, "headerReference", "default", rel_id)
                    changed += 1
            elif header_text(default.get(f"{RX}id")) != title:
                if set_part_text(default.get(f"{RX}id"), title):
                    changed += 1
            # even：摘要节 → 另一语言标题；其他节 → 与 default 一致
            expect_even = even_title if even_title is not None else title
            if even_odd:
                if even is None:
                    if even_title is not None:
                        part_name = _create_part(Path(docx_path), "header",
                                                 _header_content_type(), _build_header_part(even_title))
                        rel_id = _rel_id_of(part_name)
                        if rel_id is None:
                            raise RuntimeError(f"header part 未注册 rel: {part_name}")
                        _bind_reference(docx_path, sec_idx, "headerReference", "even", rel_id)
                        changed += 1
                    else:
                        default_rid = rel_of(sp, "default")
                        if default_rid:
                            # 克隆 default part 为独立 even part（偶数页页眉一致但部件独立）
                            new_rid = _clone_part(default_rid, "header")
                            if new_rid:
                                _bind_reference(docx_path, sec_idx, "headerReference", "even", new_rid)
                                changed += 1
                elif header_text(even.get(f"{RX}id")) != expect_even:
                    if set_part_text(even.get(f"{RX}id"), expect_even):
                        changed += 1
        elif even_odd and default is not None and even is None:
            # 其他节（含正文节）：缺 even → 克隆 default 为独立 even part
            new_rid = _clone_part(default.get(f"{RX}id"), "header")
            if new_rid:
                _bind_reference(docx_path, sec_idx, "headerReference", "even", new_rid)
                changed += 1
        # 页脚 even 补齐（与页眉同机制：缺 even 页脚的节偶数页无页码）
        if even_odd:
            frefs = {fr.get(f"{WX}type"): fr for fr in sp.findall(f"{WX}footerReference")}
            if "default" in frefs and "even" not in frefs:
                # 克隆 default footer 为独立 even part——同一 footer part 双引用
                # 被 Word 判损坏（实测 12 变体二分定位），必须独立部件
                new_rid = _clone_part(frefs["default"].get(f"{RX}id"), "footer")
                if new_rid:
                    _bind_reference(docx_path, sec_idx, "footerReference", "even", new_rid)
                    changed += 1
        sec_idx += 1
    return {"changed": changed > 0, "sections": changed}


def apply_print_layout(
    input_docx: str | Path,
    output_docx: str | Path,
    libreoffice_path: str,
    force_odd_from_section: int = 1,
    ir: Optional[DocumentIR] = None,
) -> dict:
    """新链路打印版最终化：IR 语义识别 body_start + 验证过的 Word 分页循环。"""
    body_start = find_body_start_section(input_docx, ir=ir)
    result = enforce_print_layout(
        input_docx,
        output_docx,
        libreoffice_path,
        force_odd_from_section=force_odd_from_section,
        body_start=body_start,
    )
    # 页眉对齐（通用规则：偶数页页眉补齐 + 非正文节页眉=节标题 + 封面节去下划线）
    result["header_align"] = _align_abstract_headers(output_docx, body_start=body_start)
    return result
