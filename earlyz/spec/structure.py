"""从模板 OOXML / 规范注释抽取结构取值（目录、图表题、参考文献、节页码）。

取值层唯一入口：structural_params(spec)。action 禁止再写学校常量。
机制层（去超链主题色、even 页脚补绑）不走这里。
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Any, Optional

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
_EMU_PER_TWIP = 635
_CHAPTER_RE = re.compile(
    r"^(?:第[一二三四五六七八九十百\d]+[章节]|[\d]{1,2}\s+[\u4e00-\u9fff])"
)
_H_NUM_RE = re.compile(r"^\d+[.．]\d+")  # 1.1 概述 / 1.1□概述


def _ptxt(p) -> str:
    return "".join(t.text or "" for t in p.iter(f"{W}t")).strip()


def _norm(t: str) -> str:
    # 2026-10-08：□ 视为占位空白归一（模板「摘□□要」→「摘要」；分页词表比对用）
    return re.sub(r"[\s\u3000\u25a1]+", "", t or "")


def _style_names(z: zipfile.ZipFile) -> dict[str, str]:
    """styleId → style name（TOC 样式判定用 name：模板 styleId 常是数字）。"""
    try:
        root = etree.fromstring(z.read("word/styles.xml"))
    except (KeyError, etree.XMLSyntaxError):
        return {}
    out: dict[str, str] = {}
    for s in root.findall(f"{W}style"):
        sid = s.get(f"{W}styleId")
        name = s.find(f"{W}name")
        if sid:
            out[sid] = (name.get(f"{W}val") if name is not None else None) or sid
    return out


def _is_toc_entry_para(p, style_names: dict[str, str]) -> bool:
    """目录条目段（Word 机制判据）：TOC n 样式，或段内 PAGEREF/TOC 域。"""
    ppr = p.find(f"{W}pPr")
    ps = ppr.find(f"{W}pStyle") if ppr is not None else None
    if ps is not None:
        sid = ps.get(f"{W}val") or ""
        if re.match(r"^toc\s*\d", (style_names.get(sid) or sid).strip().lower()):
            return True
    for it in p.iter(f"{W}instrText"):
        if re.match(r"^\s*(?:PAGEREF|TOC)(?:\s|$)", it.text or ""):
            return True
    for fs in p.iter(f"{W}fldSimple"):
        if re.match(r"^\s*(?:PAGEREF|TOC)(?:\s|$)", fs.get(f"{W}instr") or ""):
            return True
    return False


def _fill_header_vocab(z: zipfile.ZipFile, doc, out: dict) -> None:
    """从模板各节绑定的 header part 抽页眉词表（分类标签=模板原文，不是学校常量）。"""
    names = set(z.namelist())
    rid_to_part: dict[str, str] = {}
    if "word/_rels/document.xml.rels" in names:
        rels = etree.fromstring(z.read("word/_rels/document.xml.rels"))
        for rel in rels:
            rid, tgt = rel.get("Id"), rel.get("Target")
            if rid and tgt and "header" in (tgt or ""):
                rid_to_part[rid] = "word/" + tgt.replace("\\", "/").lstrip("/")

    def _hdr_info(part: str) -> dict:
        if part not in names:
            return {"text": "", "ascii": None, "east_asia": None, "part": part,
                    "has_page": False}
        root = etree.fromstring(z.read(part))
        text = "".join(t.text or "" for t in root.iter(f"{W}t")).strip()
        run = next((r for r in root.iter(f"{W}r")
                    if "".join(x.text or "" for x in r.iter(f"{W}t")).strip()), None)
        ascii_f = ea = None
        if run is not None:
            rpr = run.find(f"{W}rPr")
            rf = rpr.find(f"{W}rFonts") if rpr is not None else None
            if rf is not None:
                ascii_f = rf.get(f"{W}ascii")
                ea = rf.get(f"{W}eastAsia")
        return {"text": text, "ascii": ascii_f, "east_asia": ea, "part": part,
                "has_page": _part_has_page(root)}

    def _part_has_page(root) -> bool:
        for el in root.iter(f"{W}instrText"):
            if "PAGE" in (el.text or "").upper() and "NUMPAGES" not in (el.text or "").upper():
                return True
        return False

    body = doc.find(f"{W}body")
    if body is None:
        return
    style_names = _style_names(z)
    paras = list(body.iter(f"{W}p"))
    sect_pr = list(body.iter(f"{W}sectPr"))
    # 每节段落范围（节尾 sectPr 在段内或 body 末）
    marks = []
    for i, p in enumerate(paras):
        ppr = p.find(f"{W}pPr")
        if ppr is not None and ppr.find(f"{W}sectPr") is not None:
            marks.append(i)
    ranges = []
    prev = 0
    for m in marks:
        ranges.append((prev, m))
        prev = m + 1
    if prev < len(paras) or (sect_pr and len(ranges) < len(sect_pr)):
        ranges.append((prev, len(paras)))

    identity: list[dict] = []
    body_headers: dict[str, dict] = {}
    seen_id = set()
    for i, sp in enumerate(sect_pr):
        a, b = ranges[i] if i < len(ranges) else (0, 0)
        has_chapter = False
        for p in paras[a:b + 1]:
            t = _ptxt(p)
            if ".." in t or "…" in t:
                continue
            # 目录条目不是章节（2026-09-30 真论文页眉错绑根因①）：模板目录页的
            # 「1.1 概述1」条目命中 _H_NUM_RE → 目录节被当正文节 → 「目录」进不了
            # identity_headers，正文页眉筛选把目标的「目录」页眉当成正文页眉。
            # 判据是 Word 机制（TOC 样式 / PAGEREF 域），与学校无关。
            if _is_toc_entry_para(p, style_names):
                continue
            n = _norm(t).replace("□", "")
            if n and (_CHAPTER_RE.match(n) or _H_NUM_RE.match(n)):
                has_chapter = True
                break
        bound: dict[str, dict] = {}
        for ref in sp.findall(f"{W}headerReference"):
            typ = ref.get(f"{W}type") or "default"
            part = rid_to_part.get(ref.get(REL + "id") or "")
            if part:
                bound[typ] = _hdr_info(part)
        if has_chapter:
            if bound.get("default"):
                body_headers["default"] = bound["default"]
            if bound.get("even"):
                body_headers["even"] = bound["even"]
            elif bound.get("default"):
                body_headers.setdefault("even", bound["default"])
        else:
            for info in bound.values():
                t = (info.get("text") or "").strip()
                nt = _norm(t)
                if t and nt not in seen_id:
                    seen_id.add(nt)
                    identity.append(info)
    out["identity_headers"] = identity
    out["body_headers"] = body_headers
    # 模板实际绑定的页眉/页脚槽位（含空 part）：「模板定义了哪些槽位」的取值依据。
    # header_footer_parts 只收有文字的 part，拿它判「模板未定义该槽位」会把模板
    # 用空 part 定义的槽位（学校模板封面节 first）当成未定义 → 删掉目标正文节的
    # 首页页眉/页脚（2026-09-30 真论文第 1 章首页页眉页码消失，根因③）。
    all_rels: dict[str, str] = {}
    if "word/_rels/document.xml.rels" in names:
        for rel in etree.fromstring(z.read("word/_rels/document.xml.rels")):
            if rel.get("Id"):
                all_rels[rel.get("Id")] = rel.get("Target") or ""
    slots: list[list[str]] = []
    for sp in sect_pr:
        for tag, kind in (("headerReference", "header"), ("footerReference", "footer")):
            for ref in sp.findall(f"{W}{tag}"):
                if not all_rels.get(ref.get(REL + "id") or ""):
                    continue
                slot = [kind, ref.get(f"{W}type") or "default"]
                if slot not in slots:
                    slots.append(slot)
    out["hf_bound_slots"] = slots
    # 封面性质页 = 首个身份页眉/章标题节之前。取值：这些节有没有可见页眉/PAGE。
    front = {"has_visible_header": False, "has_page": False}
    id_norms = {_norm(h.get("text") or "") for h in identity}
    id_norms.discard("")
    for i, sp in enumerate(sect_pr):
        a, b = ranges[i] if i < len(ranges) else (0, 0)
        start_n = ""
        for p in paras[a:b + 1]:
            t = _ptxt(p)
            if t:
                start_n = _norm(t).replace("□", "")
                break
        is_id = any(start_n == lab or start_n.startswith(lab) or lab.startswith(start_n)
                    for lab in id_norms)
        bound_texts = []
        for ref in sp.findall(f"{W}headerReference"):
            part = rid_to_part.get(ref.get(REL + "id") or "")
            if not part:
                continue
            bt = _norm((_hdr_info(part).get("text") or ""))
            if bt:
                bound_texts.append(bt)
        if any(any(bt == lab or bt.startswith(lab) or lab.startswith(bt) for lab in id_norms)
               for bt in bound_texts):
            is_id = True
        is_ch = bool(start_n and (_CHAPTER_RE.match(start_n) or _H_NUM_RE.match(start_n)))
        if is_id or is_ch:
            break
        for ref in list(sp.findall(f"{W}headerReference")) + list(sp.findall(f"{W}footerReference")):
            part = rid_to_part.get(ref.get(REL + "id") or "")
            if not part:
                continue
            info = _hdr_info(part)
            if (info.get("text") or "").strip() and ref.tag.endswith("headerReference"):
                front["has_visible_header"] = True
            if info.get("has_page"):
                front["has_page"] = True
    out["front"] = front
    out["provenance"].append(
        f"front has_visible_header={front['has_visible_header']} has_page={front['has_page']}"
    )
    if identity:
        out["provenance"].append(
            "identity_headers from template: "
            + ",".join(h["text"] for h in identity)
        )
    if body_headers:
        out["provenance"].append("body_headers from template chapter section")


def _extract_footnote_format(tpl: Path) -> dict[str, Any]:
    """模板脚注/尾注取值（2026-09-25 批3 覆盖扩展；机制层，无学校常量）。

    取值来源按可信度：
    ① word/footnotes.xml、word/endnotes.xml 里真脚注段的直接格式（模板作者
       自排的脚注样式）——跳过分隔符段（w:type=separator/continuationSeparator）；
    ② styles.xml 里的 FootnoteText/EndnoteText（脚注文本/尾注文本）样式。
    两者都没有 → 返回 {}（不猜、不造值），目标文档脚注保持原样。
    """
    out: dict[str, Any] = {}
    try:
        with zipfile.ZipFile(tpl) as z:
            names = set(z.namelist())
            for part, kind in (("word/footnotes.xml", "footnote"),
                               ("word/endnotes.xml", "endnote")):
                if part not in names:
                    continue
                try:
                    root = etree.fromstring(z.read(part))
                except Exception:
                    continue
                for note in root.iter(f"{W}{kind}"):
                    ntype = note.get(f"{W}type") or ""
                    if ntype in ("separator", "continuationSeparator"):
                        continue
                    for p in note.iter(f"{W}p"):
                        props: dict[str, Any] = {}
                        for r in p.iter(f"{W}r"):
                            rpr = r.find(f"{W}rPr")
                            if rpr is None:
                                continue
                            rf = rpr.find(f"{W}rFonts")
                            if rf is not None:
                                if rf.get(f"{W}eastAsia") and "font_east_asia" not in props:
                                    props["font_east_asia"] = rf.get(f"{W}eastAsia")
                                if rf.get(f"{W}ascii") and "font_ascii" not in props:
                                    props["font_ascii"] = rf.get(f"{W}ascii")
                            sz = rpr.find(f"{W}sz")
                            if sz is not None and "font_size_pt" not in props:
                                try:
                                    props["font_size_pt"] = round(
                                        int(sz.get(f"{W}val")) / 2.0, 1)
                                except (TypeError, ValueError):
                                    pass
                        ppr = p.find(f"{W}pPr")
                        if ppr is not None:
                            sp = ppr.find(f"{W}spacing")
                            if sp is not None:
                                if sp.get(f"{W}line"):
                                    try:
                                        props["line"] = int(sp.get(f"{W}line"))
                                    except (TypeError, ValueError):
                                        pass
                                if sp.get(f"{W}lineRule"):
                                    props["line_rule"] = sp.get(f"{W}lineRule")
                        if props:
                            out[kind] = props
                            break
                    if out.get(kind):
                        break
            if not out:
                styles_name = "word/styles.xml"
                if styles_name in names:
                    try:
                        sroot = etree.fromstring(z.read(styles_name))
                    except Exception:
                        sroot = None
                    if sroot is not None:
                        for st in sroot.iter(f"{W}style"):
                            nm = (st.find(f"{W}name").get(f"{W}val")
                                  if st.find(f"{W}name") is not None else "")
                            sid = st.get(f"{W}styleId") or ""
                            if not (nm.startswith(("FootnoteText", "EndnoteText",
                                                   "脚注文本", "尾注文本"))
                                    or sid.startswith(("FootnoteText", "EndnoteText"))):
                                continue
                            rpr = st.find(f"{W}rPr")
                            props: dict[str, Any] = {}
                            if rpr is not None:
                                rf = rpr.find(f"{W}rFonts")
                                if rf is not None:
                                    if rf.get(f"{W}eastAsia"):
                                        props["font_east_asia"] = rf.get(f"{W}eastAsia")
                                    if rf.get(f"{W}ascii"):
                                        props["font_ascii"] = rf.get(f"{W}ascii")
                                sz = rpr.find(f"{W}sz")
                                if sz is not None:
                                    try:
                                        props["font_size_pt"] = round(
                                            int(sz.get(f"{W}val")) / 2.0, 1)
                                    except (TypeError, ValueError):
                                        pass
                            if props:
                                kind = "endnote" if "ndnote" in (nm + sid) else "footnote"
                                out.setdefault(kind, props)
    except Exception:
        return out
    return out


def extract_template_structure(tpl: Path) -> dict[str, Any]:
    """读模板 zip，返回 structure dict（可进 FormatSpec.structure）。"""
    tpl = Path(tpl)
    out: dict[str, Any] = {
        "toc": {},
        "caption": {},
        "refs": {},
        "sections": [],
        "header_format": {},
        "identity_headers": [],
        "body_headers": {},
        "front": {"has_visible_header": False, "has_page": False},
        "front_cols": {},
        "body_cols": {},
        "body_cols_source": {},
        "body_section_anchor": "",
        "body_section_anchor_kind": "",
        "front_tokens": [],
        "section_plan": [],
        "equation": {},
        "figure_width": {},
        "table_width": {},
        "table_cells": {},
        "table_borders": {},
        "drop_cap_layout": {},
        "footnote": {},
        "even_and_odd": False,
        "provenance": [],
    }
    settings = None
    try:
        with zipfile.ZipFile(tpl) as z:
            doc = etree.fromstring(z.read("word/document.xml"))
            if "word/settings.xml" in z.namelist():
                settings = etree.fromstring(z.read("word/settings.xml"))
            _fill_header_vocab(z, doc, out)
    except Exception:
        return out

    if settings is not None:
        out["even_and_odd"] = settings.find(f"{W}evenAndOddHeaders") is not None

    from earlyz.spec.decompiler import parse_annotation_props

    body = doc.find(f"{W}body")
    paras = list(body.iter(f"{W}p")) if body is not None else []

    # ── 目录：规范注释「各章的名称黑体4号加粗，其余宋体小4」──
    for p in paras:
        t = _ptxt(p)
        if "各章" in t and "号" in t:
            if "其余" in t:
                left, right = t.split("其余", 1)
            else:
                left, right = t, ""
            l1 = parse_annotation_props(left)
            l2 = parse_annotation_props(right) if right else {}
            if l1:
                out["toc"]["1"] = l1
                out["provenance"].append(f"toc.l1 from annotation: {t[:40]}")
            if l2:
                out["toc"]["2"] = l2
                out["toc"]["3"] = dict(l2)
                out["provenance"].append(f"toc.l2 from annotation: {t[:40]}")
            break

    # ── 目录存在性：模板里有 TOC 域即视为「要求出目录」（2026-10-28 大连民族）──
    # 旧逻辑只在模板写了「各章的名称黑体4号…」这类注释时才填 toc，
    # 没有注释但有 TOC 域的模板 → toc={} → 编译链不触发 insert_toc_if_missing，
    # 产物就没有目录（模板明确要求目录的文档被判「无要求」）。
    if not out["toc"]:
        for _p in paras:
            _instr = "".join(x.text or "" for x in _p.iter(f"{W}instrText"))
            if _instr.strip().upper().startswith("TOC"):
                out["toc"]["present"] = True
                out["provenance"].append("toc.present from TOC field")
                break

    # ── 目录标题格式：目录示例区块里的「目□□录」+ 下一段「（黑体小2加粗）」（2026-10-17）──
    #    模板把标题与格式拆成两段（实证：id113「目□□录」/ id114「（黑体小2加粗）」）
    for idx, p in enumerate(paras):
        tc = re.sub(r"[□\u25a1\u25a0\s\u3000]+", "", _ptxt(p))
        if tc not in ("目录", "目次"):
            continue
        for nxt in paras[idx + 1: idx + 3]:
            t2 = _ptxt(nxt).strip()
            m = re.fullmatch(r"[（(]([^）)]+)[）)]", t2)
            if not m:
                continue
            cand = parse_annotation_props(m.group(1))
            if cand:
                out["toc_title"] = cand
                out["provenance"].append(f"toc.title from annotation: {m.group(1)[:40]}")
                break
        break

    # ── 图表题：示例段 run 级（注释只作线索，示例 OOXML 为准）──
    tbl_east, fig_east = [], []
    num_bold, latin_bold, latin_seen = [], [], 0
    # 表题 / 图题分池（2026-09-29）：同一模板里两类题注的加粗形态可以不同
    # （西安理工 2026 版：表题整段加粗含 Altera/HBsAg；图题只序号加粗、Symbol 不粗）。
    # 合池多数票会把表题取值套到图题上 → 与公式编号行、前置页标题同一原则：
    # 取值层不能把不同形态合成一条规则。
    kind_num_bold: dict[str, list] = {"table": [], "figure": []}
    kind_latin_bold: dict[str, list] = {"table": [], "figure": []}
    tbl_sz, fig_sz = [], []
    tbl_jc, fig_jc = [], []
    tbl_before, fig_before = [], []
    tbl_after, fig_after = [], []
    tbl_bl, fig_bl = [], []
    tbl_al, fig_al = [], []
    _CJK_FACE = ("宋体", "黑体", "仿宋", "楷体")
    for p in paras:
        t = _ptxt(p)
        if not t:
            continue
        core = re.sub(r"[*＊].*$", "", t).strip()
        core = re.sub(
            r"[（(][^）)]{0,80}(?:黑体|宋体|仿宋|号|磅|加粗)[^）)]{0,80}[）)]\s*$",
            "",
            core,
        ).strip()
        if not core or len(core) > 60:
            continue
        is_tbl = bool(re.match(r"^表\s*[\dX]", core))
        is_fig = bool(re.match(r"^图\s*[\dX]", core))
        if not (is_tbl or is_fig):
            continue
        if re.search(r"[，。；]", core):
            continue
        ppr = p.find(f"{W}pPr")
        jc_el = ppr.find(f"{W}jc") if ppr is not None else None
        if jc_el is not None and jc_el.get(f"{W}val"):
            (tbl_jc if is_tbl else fig_jc).append(jc_el.get(f"{W}val"))
        sp = ppr.find(f"{W}spacing") if ppr is not None else None
        if sp is not None:
            if sp.get(f"{W}before"):
                (tbl_before if is_tbl else fig_before).append(int(sp.get(f"{W}before")))
            if sp.get(f"{W}after"):
                (tbl_after if is_tbl else fig_after).append(int(sp.get(f"{W}after")))
            if sp.get(f"{W}beforeLines"):
                (tbl_bl if is_tbl else fig_bl).append(sp.get(f"{W}beforeLines"))
            if sp.get(f"{W}afterLines"):
                (tbl_al if is_tbl else fig_al).append(sp.get(f"{W}afterLines"))
        from earlyz.spec.decompiler import parse_annotation_props
        ann = parse_annotation_props(t)
        if not ann and "*" in t:
            ann = parse_annotation_props(t.split("*", 1)[-1])
        if ann.get("font_size_pt"):
            half = str(int(round(float(ann["font_size_pt"]) * 2)))
            (tbl_sz if is_tbl else fig_sz).append(half)
        if ann.get("font_east_asia"):
            (tbl_east if is_tbl else fig_east).append(ann["font_east_asia"])
        for r in p.iter(f"{W}r"):
            rt = "".join(x.text or "" for x in r.iter(f"{W}t"))
            if not rt.strip():
                continue
            rpr = r.find(f"{W}rPr")
            b = rpr is not None and rpr.find(f"{W}b") is not None
            sz = rpr.find(f"{W}sz") if rpr is not None else None
            rf = rpr.find(f"{W}rFonts") if rpr is not None else None
            ea = rf.get(f"{W}eastAsia") if rf is not None else None
            asc = rf.get(f"{W}ascii") if rf is not None else None
            has_digit = bool(re.search(r"[0-9]", rt))
            has_cjk = bool(re.search(r"[\u4e00-\u9fff]", rt))
            has_latin = bool(re.search(r"[A-Za-z]", rt))
            face = ea
            if has_cjk and not face and asc and any(k in asc for k in _CJK_FACE):
                face = asc
            if has_cjk and face:
                (tbl_east if is_tbl else fig_east).append(face)
            if sz is not None:
                (tbl_sz if is_tbl else fig_sz).append(sz.get(f"{W}val"))
            _kind = "table" if is_tbl else "figure"
            if has_digit and not has_cjk:
                num_bold.append(b)
                kind_num_bold[_kind].append(b)
            if has_latin and not has_digit:
                latin_seen += 1
                latin_bold.append(b)
                kind_latin_bold[_kind].append(b)
    cap: dict[str, Any] = {}
    if tbl_east:
        cap["table_east_asia"] = max(set(tbl_east), key=tbl_east.count)
    if fig_east:
        cap["figure_east_asia"] = max(set(fig_east), key=fig_east.count)
    if tbl_sz:
        cap["table_sz"] = max(set(tbl_sz), key=tbl_sz.count)
    if fig_sz:
        cap["figure_sz"] = max(set(fig_sz), key=fig_sz.count)
    if tbl_jc:
        cap["table_jc"] = max(set(tbl_jc), key=tbl_jc.count)
    if fig_jc:
        cap["figure_jc"] = max(set(fig_jc), key=fig_jc.count)
    if tbl_before:
        cap["table_before"] = max(set(tbl_before), key=tbl_before.count)
    if fig_before:
        cap["figure_before"] = max(set(fig_before), key=fig_before.count)
    if tbl_after:
        cap["table_after"] = max(set(tbl_after), key=tbl_after.count)
    if fig_after:
        cap["figure_after"] = max(set(fig_after), key=fig_after.count)
    if tbl_bl:
        cap["table_before_lines"] = max(set(tbl_bl), key=tbl_bl.count)
    if fig_bl:
        cap["figure_before_lines"] = max(set(fig_bl), key=fig_bl.count)
    if tbl_al:
        cap["table_after_lines"] = max(set(tbl_al), key=tbl_al.count)
    if fig_al:
        cap["figure_after_lines"] = max(set(fig_al), key=fig_al.count)
    if num_bold:
        cap["number_bold"] = sum(num_bold) >= (len(num_bold) + 1) // 2
    elif tbl_jc or fig_jc or tbl_east or fig_east:
        cap["number_bold"] = False
    if latin_seen:
        cap["latin_bold"] = sum(latin_bold) > len(latin_bold) / 2
    else:
        cap["latin_bold"] = False
    # 分类取值：该类题注自己有样本才写（`table_latin_bold` / `figure_number_bold` …）；
    # 没有样本的一类由动作侧回落到上面的合池值（旧行为不变）。
    for _kind in ("table", "figure"):
        if kind_num_bold[_kind]:
            _v = kind_num_bold[_kind]
            cap[f"{_kind}_number_bold"] = sum(_v) >= (len(_v) + 1) // 2
        if kind_latin_bold[_kind]:
            _v = kind_latin_bold[_kind]
            cap[f"{_kind}_latin_bold"] = sum(_v) > len(_v) / 2
    if cap:
        out["caption"] = cap
        out["provenance"].append(f"caption from template runs: {cap}")
    _fill_bilingual_captions(paras, out)
    _fill_equation_layout(paras, out)
    _fill_equation_from_annotation(paras, out)
    _fill_cover_slots(paras, out)

    # ── 参考文献：条目 OOXML + 条目注释（跳过标题「黑体小2加粗」）──
    in_refs = False
    ref_anns: list[tuple[str, dict]] = []
    saw_entry = False
    for p in paras:
        t = _ptxt(p)
        n = re.sub(r"\s+", "", t)
        if n in ("参考文献", "参 考 文 献".replace(" ", ""), "References", "REFERENCES"):
            in_refs = True
            continue
        if not in_refs:
            continue
        if n in ("致谢", "附录", "外文资料", "外文译文"):
            break
        if t.startswith("（") or t.startswith("(") or t.startswith("…"):
            props = parse_annotation_props(t)
            if props:
                ref_anns.append((t, props))
            continue
        if re.match(r"^\[\d+\]", t) or re.match(r"^\d+\.", t):
            if not saw_entry:
                out["refs"]["numbered"] = True
                ppr = p.find(f"{W}pPr")
                sp = ppr.find(f"{W}spacing") if ppr is not None else None
                if sp is not None:
                    if sp.get(f"{W}line"):
                        out["refs"]["line"] = int(sp.get(f"{W}line"))
                    if sp.get(f"{W}lineRule"):
                        out["refs"]["line_rule"] = sp.get(f"{W}lineRule")
                ind = ppr.find(f"{W}ind") if ppr is not None else None
                if ind is not None and ind.get(f"{W}firstLineChars"):
                    out["refs"]["ind_first_line_chars"] = int(ind.get(f"{W}firstLineChars"))
                else:
                    out["refs"]["ind_first_line_chars"] = 0
                saw_entry = True
                out["provenance"].append("refs first entry ooxml")
            continue
        # 2026-10-17：模板只给「格式说明 + 一条样例」时（IEEE/自动化学报），
        # 样例不带 [n] → 用第一条像文献的段落兜底抽字体/字号/行距。
        if not saw_entry and not out["refs"].get("numbered"):
            _maybe_ref_sample(p, t, out)
    # 条目注释：不加粗的才是条目规范；标题注释（黑体加粗）丢掉
    for t, props in reversed(ref_anns):
        if props.get("bold"):
            continue
        out["refs"].update(props)
        out["provenance"].append(f"refs annotation: {t[:40]}")
        break

    # ── 节页码：模板 sectPr 原样 ──
    has_roman = False
    starts = _section_start_texts(body) if body is not None else []
    if body is not None:
        for i, sp in enumerate(body.iter(f"{W}sectPr")):
            pn = sp.find(f"{W}pgNumType")
            rec = {"index": i, "fmt": None, "start": None,
                   "start_text": starts[i] if i < len(starts) else ""}
            if pn is not None:
                rec["fmt"] = pn.get(f"{W}fmt")
                rec["start"] = pn.get(f"{W}start")
                if rec["fmt"] and "roman" in rec["fmt"].lower():
                    has_roman = True
            rec.update(_sect_cols(sp))
            out["sections"].append(rec)
    out["front_matter_roman"] = has_roman
    _fill_column_object_widths(doc, out)
    _fill_table_cells(doc, paras, parse_annotation_props, out)
    _fill_table_borders(doc, out)
    _fill_drop_cap_layout(doc, out)
    _fill_section_break_words(paras, out)
    out["footnote"] = _extract_footnote_format(tpl)
    # 页眉格式（2026-10-17）：只在 decompiler 没抽到 header_all 规则时兜底。
    # 取模板 header part 里**有文字的段**的字体/字号/对齐，多部件取多数值——
    # 只取格式不取文本（文本属于目标文档，内容保护铁律）。
    _numbering_notes(paras, out)
    _hf = _header_format_from_parts(tpl)
    if _hf:
        out["header_format"] = _hf
        out["provenance"].append("header_format from template header parts")
    # 脚注编号机制（2026-10-17）：模板明写「每页的脚注从 1 开始编序号」才落键。
    # 只认明写——模板沉默不猜编号方式。
    for _p in paras:
        _t = _ptxt(_p)
        if ("脚注" in _t and re.search(r"每\s*页", _t)
                and re.search(r"从\s*1\s*开始|重新编号|重新编|从\s*1\s*编", _t)):
            out["footnote_restart"] = "eachPage"
            out["provenance"].append("footnote numRestart=eachPage from annotation")
            break
    return out


def _section_start_texts(body) -> list[str]:
    """每节第一个非空段文本。按 body 子节点顺序切节，避免 find(sectPr) 对上错节。"""
    starts: list[str] = []
    current = ""
    seen = False

    def flush() -> None:
        nonlocal current, seen
        starts.append(current)
        current = ""
        seen = False

    def consider(p) -> None:
        nonlocal current, seen
        if seen:
            return
        t = "".join(x.text or "" for x in p.iter(f"{W}t")).strip()
        if t:
            current = t
            seen = True

    for child in list(body):
        if child.tag == f"{W}sectPr":
            flush()
            continue
        if child.tag == f"{W}p":
            consider(child)
            ppr = child.find(f"{W}pPr")
            if ppr is not None and ppr.find(f"{W}sectPr") is not None:
                flush()
            continue
        if child.tag == f"{W}tbl":
            for p in child.iter(f"{W}p"):
                consider(p)
            continue
        for p in child.iter(f"{W}p"):
            consider(p)
            ppr = p.find(f"{W}pPr")
            if ppr is not None and ppr.find(f"{W}sectPr") is not None:
                flush()
    return starts


def _sect_cols(sp) -> dict[str, Any]:
    """从 sectPr 读栏数/栏宽/版心。缺 num 视为 1 栏。"""
    rec: dict[str, Any] = {"cols_num": 1, "cols_space": None, "cols_equal": True,
                           "col_widths": [], "page_content": None,
                           "pg_w": None, "pg_h": None, "sect_type": None, "pg_mar": {}}
    pg = sp.find(f"{W}pgSz")
    mar = sp.find(f"{W}pgMar")
    typ = sp.find(f"{W}type")
    if typ is not None and typ.get(f"{W}val"):
        rec["sect_type"] = typ.get(f"{W}val")
    try:
        if pg is not None:
            rec["pg_w"] = int(pg.get(f"{W}w") or 0) or None
            rec["pg_h"] = int(pg.get(f"{W}h") or 0) or None
        if mar is not None:
            rec["pg_mar"] = {}
            for k in ("top", "right", "bottom", "left", "header", "footer", "gutter"):
                if mar.get(f"{W}{k}") is not None:
                    rec["pg_mar"][k] = int(mar.get(f"{W}{k}"))
        if pg is not None and mar is not None:
            rec["page_content"] = (int(pg.get(f"{W}w") or 0)
                                   - int(mar.get(f"{W}left") or 0)
                                   - int(mar.get(f"{W}right") or 0))
    except (TypeError, ValueError):
        rec["page_content"] = None
    cols = sp.find(f"{W}cols")
    if cols is None:
        return rec
    try:
        rec["cols_num"] = int(cols.get(f"{W}num") or 1)
    except (TypeError, ValueError):
        rec["cols_num"] = 1
    if cols.get(f"{W}space"):
        try:
            rec["cols_space"] = int(cols.get(f"{W}space"))
        except (TypeError, ValueError):
            pass
    rec["cols_equal"] = cols.get(f"{W}equalWidth") != "0"
    widths = []
    col_spaces = []
    for c in cols.findall(f"{W}col"):
        try:
            widths.append(int(c.get(f"{W}w") or 0))
        except (TypeError, ValueError):
            pass
        if c.get(f"{W}space"):
            try:
                col_spaces.append(int(c.get(f"{W}space")))
            except (TypeError, ValueError):
                pass
    # 分栏间距以 col/@space 为准（IEEE：cols@space=720 但 col@space=288 才等于版心）
    if col_spaces:
        rec["cols_space"] = col_spaces[0]
        rec["col_spaces"] = col_spaces
    if not widths and rec["page_content"] and rec["cols_num"] >= 1:
        gap = rec["cols_space"] or 0
        avail = rec["page_content"] - gap * (rec["cols_num"] - 1)
        widths = [max(1, avail // rec["cols_num"])] * rec["cols_num"]
    rec["col_widths"] = widths
    return rec


_FIG_EN_RE = re.compile(r"^Fig(?:ure)?\.?\s*\d", re.I)
_TBL_EN_RE = re.compile(r"^Table\s+\d", re.I)
_EQ_NUM_RE = re.compile(r"^[（(]\s*\d+\s*[)）]$")
_EQ_OLE_RE = re.compile(r"^(Equation|MathType)", re.I)


def _para_jc_sz_font(p) -> dict[str, Any]:
    rec: dict[str, Any] = {}
    ppr = p.find(f"{W}pPr")
    if ppr is not None:
        j = ppr.find(f"{W}jc")
        if j is not None and j.get(f"{W}val"):
            rec["jc"] = j.get(f"{W}val")
    for r in p.iter(f"{W}r"):
        rpr = r.find(f"{W}rPr")
        if rpr is None:
            continue
        sz = rpr.find(f"{W}sz")
        rf = rpr.find(f"{W}rFonts")
        if sz is not None and sz.get(f"{W}val"):
            rec["sz"] = sz.get(f"{W}val")
        if rf is not None:
            if rf.get(f"{W}ascii"):
                rec["ascii"] = rf.get(f"{W}ascii")
            if rf.get(f"{W}eastAsia"):
                rec["east_asia"] = rf.get(f"{W}eastAsia")
        if rec.get("sz"):
            break
    return rec


def _fill_bilingual_captions(paras, out: dict) -> None:
    """连续「图N / Fig.N」「表N / Table N」才收英文题。无英文样本则不发明。"""
    cap = dict(out.get("caption") or {})
    n = len(paras)
    fig_en_jc, fig_en_sz, fig_en_ascii = [], [], []
    tbl_en_jc, tbl_en_sz, tbl_en_ascii = [], [], []
    bilingual = False
    for i, p in enumerate(paras):
        t = _ptxt(p)
        if not t:
            continue
        nxt = ""
        if i + 1 < n:
            nxt = _ptxt(paras[i + 1])
        fmt_n = _para_jc_sz_font(paras[i + 1]) if nxt else {}
        if re.match(r"^图\s*[\dX]", t) and _FIG_EN_RE.match(nxt):
            bilingual = True
            if fmt_n.get("jc"):
                fig_en_jc.append(fmt_n["jc"])
            if fmt_n.get("sz"):
                fig_en_sz.append(fmt_n["sz"])
            if fmt_n.get("ascii"):
                fig_en_ascii.append(fmt_n["ascii"])
        if re.match(r"^表\s*[\dX]", t) and _TBL_EN_RE.match(nxt):
            bilingual = True
            if fmt_n.get("jc"):
                tbl_en_jc.append(fmt_n["jc"])
            if fmt_n.get("sz"):
                tbl_en_sz.append(fmt_n["sz"])
            if fmt_n.get("ascii"):
                tbl_en_ascii.append(fmt_n["ascii"])
    if fig_en_jc:
        cap["figure_en_jc"] = max(set(fig_en_jc), key=fig_en_jc.count)
    if fig_en_sz:
        cap["figure_en_sz"] = max(set(fig_en_sz), key=fig_en_sz.count)
    if fig_en_ascii:
        cap["figure_en_ascii"] = max(set(fig_en_ascii), key=fig_en_ascii.count)
    if tbl_en_jc:
        cap["table_en_jc"] = max(set(tbl_en_jc), key=tbl_en_jc.count)
    if tbl_en_sz:
        cap["table_en_sz"] = max(set(tbl_en_sz), key=tbl_en_sz.count)
    if tbl_en_ascii:
        cap["table_en_ascii"] = max(set(tbl_en_ascii), key=tbl_en_ascii.count)
    if bilingual:
        cap["bilingual"] = True
        out["caption"] = cap
        out["provenance"].append("caption bilingual CN+EN lines from template")


def _fill_equation_layout(paras, out: dict) -> None:
    """公式段对齐：优先取含 OLE 公式的段落 jc；没有 OLE 才退回编号行 jc。"""
    ole_jcs: list[str] = []
    num_jcs: list[str] = []
    tab_pos = None
    for p in paras:
        t = _ptxt(p)
        ole = False
        for el in p.iter():
            pid = el.get("ProgID") or ""
            if pid and _EQ_OLE_RE.match(pid):
                ole = True
                break
        if not (ole or _EQ_NUM_RE.match(t)):
            continue
        if ole and t and not _EQ_NUM_RE.match(t) and len(t) > 6:
            continue
        ppr = p.find(f"{W}pPr")
        if ppr is None:
            continue
        j = ppr.find(f"{W}jc")
        jval = j.get(f"{W}val") if j is not None else None
        if jval:
            (ole_jcs if ole else num_jcs).append(jval)
        # 编号右对齐靠制表位：记下右制表位位置（twips）
        tabs = ppr.find(f"{W}tabs")
        if tabs is not None:
            for tab in tabs.findall(f"{W}tab"):
                if (tab.get(f"{W}val") or "") == "right" and tab.get(f"{W}pos"):
                    try:
                        tab_pos = int(tab.get(f"{W}pos"))
                    except (TypeError, ValueError):
                        pass
    jcs = ole_jcs or num_jcs
    if jcs:
        out["equation"] = {"jc": max(set(jcs), key=jcs.count), "protect": True}
        if tab_pos is not None:
            out["equation"]["number_tab_pos"] = tab_pos
            out["equation"]["number_align"] = "right"
        out["provenance"].append(
            f"equation jc={out['equation']['jc']}"
            + ("（取自 OLE 公式段）" if ole_jcs else "（取自编号行）"))


_EQ_SIZE_LABELS = (
    ("font_size_pt", re.compile(r"(?:标准|正文|公式)\s*[:：]\s*([0-9]+(?:\.[0-9]+)?)\s*磅")),
    ("script_size_pt", re.compile(r"(?<!次)\s*下标\s*/\s*上标\s*[:：]\s*([0-9]+(?:\.[0-9]+)?)\s*磅")),
    ("sub_script_size_pt", re.compile(r"次下标\s*/\s*上标\s*[:：]\s*([0-9]+(?:\.[0-9]+)?)\s*磅")),
)


def _fill_equation_from_annotation(paras, out: dict) -> None:
    """没有 OLE/编号样本时，从「公式尺寸：标准 10.5 磅 / 下标上标 5.8 磅」这类注记取值。

    2026-10-17：CJC 模板没有任何公式对象，但明写了公式三级字号——旧实现
    `structure.equation` 为空，公式格式只能凭默认值。
    """
    if out.get("equation"):
        return
    for i, p in enumerate(paras):
        t = _ptxt(p)
        if "公式" not in t or not re.search(r"尺寸|字号|磅", t):
            continue
        vals: dict[str, float] = {}
        for q in paras[i:i + 8]:
            qt = _ptxt(q)
            for key, pat in _EQ_SIZE_LABELS:
                if key in vals:
                    continue
                m = pat.search(qt)
                if m:
                    try:
                        vals[key] = float(m.group(1))
                    except (TypeError, ValueError):
                        pass
        if vals:
            vals["source"] = "annotation"
            vals["protect"] = True
            out["equation"] = vals
            out["provenance"].append(
                "equation sizes from annotation: "
                + ", ".join(f"{k}={v}" for k, v in vals.items() if k.endswith("_pt"))
            )
            return


_COVER_LABELS = (
    "题目|论文题目|专业|专业名称|学院|系别|学生姓名|姓名|学号|学生学号|"
    "指导教师|教师|日期|年月日|班级|学位|作者|学制|研究方向"
)
_COVER_LABEL_RE = re.compile(rf"^(?:{_COVER_LABELS})$")
_COVER_SLOT_RE = re.compile(rf"^({_COVER_LABELS})\s*[:：]")


_EQ_NUM_ARABIC = re.compile(r"公式[^。；\n]{0,24}阿拉伯[^。；\n]{0,16}编号")
_TABLE_ROMAN = re.compile(
    r"(?:表格?|tables?)[^。；.\n]{0,40}(?:罗马数字|roman numerals?)"
    r"|(?:罗马数字|roman numerals?)[^。；.\n]{0,24}(?:表格?|tables?)", re.I)


def _numbering_notes(paras, out: dict) -> None:
    """模板明写的编号约定（2026-10-17）。

    能落实的进结构（公式阿拉伯数字连续编号 = 引擎现行行为）；
    引擎做不到的（表格罗马数字编号）进 unsupported_requirements，
    由编译期「未落实要求」告警如实告诉用户，而不是静默忽略。
    """
    for p in paras:
        t = _ptxt(p)
        if not t:
            continue
        if _EQ_NUM_ARABIC.search(t) or re.search(
                r"equations?[^.]{0,60}numbered consecutively", t, re.I):
            eq = out.setdefault("equation", {})
            if isinstance(eq, dict):
                eq.setdefault("numbering", "arabic")
                eq.setdefault("numbering_scope", "continuous")
                out["provenance"].append(
                    "equation numbering: arabic continuous from annotation")
        if _TABLE_ROMAN.search(t):
            cap = out.setdefault("caption", {})
            if isinstance(cap, dict):
                cap.setdefault("table_number_style", "roman")
                out["provenance"].append(
                    "caption.table_number_style=roman from annotation")


def _header_format_from_parts(tpl: Path) -> dict:
    """模板 header 部件的字体/字号/对齐（多部件取多数）。模板没有页眉 → {}。"""
    try:
        z = zipfile.ZipFile(tpl)
    except Exception:
        return {}
    votes: list[dict] = []
    try:
        names = sorted(n for n in z.namelist()
                       if re.match(r"word/header\d+\.xml$", n))
        for name in names:
            try:
                root = etree.fromstring(z.read(name))
            except Exception:
                continue
            for p in root.iter(f"{W}p"):
                txt = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
                if not txt:
                    continue
                props: dict = {}
                rpr = None
                for r in p.iter(f"{W}r"):
                    te = r.find(f"{W}t")
                    if te is not None and (te.text or "").strip():
                        rpr = r.find(f"{W}rPr")
                        break
                if rpr is None:
                    ppr0 = p.find(f"{W}pPr")
                    rpr = ppr0.find(f"{W}rPr") if ppr0 is not None else None
                if rpr is not None:
                    rf = rpr.find(f"{W}rFonts")
                    if rf is not None:
                        if rf.get(f"{W}eastAsia"):
                            props["font_east_asia"] = rf.get(f"{W}eastAsia")
                        if rf.get(f"{W}ascii"):
                            props["font_ascii"] = rf.get(f"{W}ascii")
                    sz = rpr.find(f"{W}sz")
                    if sz is not None:
                        try:
                            props["font_size_pt"] = round(int(sz.get(f"{W}val")) / 2.0, 1)
                        except (TypeError, ValueError):
                            pass
                    if rpr.find(f"{W}b") is not None:
                        props["bold"] = True
                ppr = p.find(f"{W}pPr")
                if ppr is not None:
                    jc = ppr.find(f"{W}jc")
                    if jc is not None and jc.get(f"{W}val"):
                        props["jc"] = jc.get(f"{W}val")
                    if ppr.find(f"{W}pBdr") is not None:
                        props["p_bdr_bottom"] = True
                if props:
                    votes.append(props)
                    break  # 每个 header 部件只取首个有文字的段
    finally:
        z.close()
    if not votes:
        return {}
    out: dict = {}
    for key in ("font_east_asia", "font_ascii", "font_size_pt", "jc", "bold",
                "p_bdr_bottom"):
        vals = [v.get(key) for v in votes if v.get(key) is not None]
        if not vals:
            continue
        # 多数值；并列取先出现者（稳定）
        out[key] = max(vals, key=lambda x: (vals.count(x), -vals.index(x)))
    return out


def _fill_cover_slots(paras, out: dict) -> None:
    """封面槽位标签（2026-10-17，P2）：把 cover.py 的模板匹配结果并入 spec 可见产物。

    只在模板前部找（封面总在文档开头），标签去空白后匹配——
    「专 业 名 称：XXX」这种字间加空格 + 带示例值是模板常态，
    去空白后按「标签：」前缀识别，不能要求整段都是标签。
    """
    labels: list[str] = []
    for p in paras[:80]:
        t = re.sub(r"[\s\u3000]+", "", _ptxt(p)).strip()
        if not t:
            continue
        m = _COVER_SLOT_RE.match(t)
        lab = m.group(1) if m else (t if _COVER_LABEL_RE.match(t) else "")
        if lab and lab not in labels:
            labels.append(lab)
    if len(labels) >= 2:
        out["cover"] = {"slots": labels, "source": "template_labels"}
        out["provenance"].append("cover slots from template labels: " + "、".join(labels))


def _plan_token(raw: str) -> str:
    if not raw or _ANN_START_RE.search(raw):
        return ""
    tok = anchor_token(raw)
    if re.fullmatch(r"[a-z]{4,}", tok) or re.fullmatch(r"[\u4e00-\u9fff]{2,6}", tok):
        return tok
    return ""


def _fill_section_plan(out: dict) -> None:
    """栏数变化点序列（连续同栏合并）。供源稿按锚点多段补节。"""
    plan = []
    prev_n = None
    for s in out.get("sections") or []:
        n = int(s.get("cols_num") or 1)
        if prev_n == n and plan:
            continue
        spec = dict(s)
        raw = (s.get("start_text") or "").strip()
        spec["token"] = _plan_token(raw)
        if _ANN_START_RE.search(raw):
            already = any(x.get("kind") == "first_heading" for x in plan)
            if n >= 2 and not already:
                spec["kind"] = "first_heading"
            else:
                spec["kind"] = "skip"
                spec["token"] = ""
        else:
            spec["kind"] = "token"
        plan.append(spec)
        prev_n = n
    out["section_plan"] = plan
    if plan:
        out["provenance"].append(f"section_plan {len(plan)} col-changes")


_ANCHOR_TOKEN_RE = re.compile(
    r"^[\s\d.、．第]*([A-Za-z]{3,}|[\u4e00-\u9fff]{1,12})"
)
_ANN_START_RE = re.compile(r"字体|置于此处|本模板|样本|请用|\*[^*]{3,}\*")


def anchor_token(text: str) -> str:
    t = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", "", (text or "").strip())
    m = _ANCHOR_TOKEN_RE.match(t)
    return (m.group(1) if m else "").casefold()


def _median(xs: list[float]) -> float:
    xs = sorted(xs)
    return xs[len(xs) // 2]


def _clustered(xs: list[float], tol: float = 0.12) -> bool:
    if not xs:
        return False
    m = _median(xs)
    return all(abs(x - m) <= tol for x in xs)


def _section_para_counts(body) -> list[int]:
    """按 body 子节点顺序切节，数每节段落数（与 ``_section_start_texts`` 同一套切节规则）。

    为什么需要：节记录（``out["sections"]``）里只有栏数/栏宽/版心，没有「这一节占了
    全文多少段」。EZI-0035 的覆盖率判据要的正是这个比例，只能在切节层再数一遍。
    body 缺失 → 返回 ``[]``，调用方退回旧口径（不报错）。
    """
    counts: list[int] = []
    if body is None:
        return counts
    cur = 0

    def flush() -> None:
        nonlocal cur
        counts.append(cur)
        cur = 0

    for child in list(body):
        if child.tag == f"{W}sectPr":
            flush()
            continue
        if child.tag == f"{W}p":
            cur += 1
            ppr = child.find(f"{W}pPr")
            if ppr is not None and ppr.find(f"{W}sectPr") is not None:
                flush()
            continue
        if child.tag == f"{W}tbl":
            for _p in child.iter(f"{W}p"):
                cur += 1
            continue
        for p in child.iter(f"{W}p"):
            cur += 1
            ppr = p.find(f"{W}pPr")
            if ppr is not None and ppr.find(f"{W}sectPr") is not None:
                flush()
    return counts


# EZI-0035 覆盖率判据的两个门槛（只判「这段 2 栏是不是前部装饰」）：
# 2 栏节段落占全文 <5%，且它**后面**另有一个 1 栏节占全文 >50% → 那个 2 栏节是
# 封面/前部装饰，真正的正文在后面那节。实例（国管局 2026 版服务合同）：节 1 是
# 「国家机关事务管理局 / 国家市场监督管理总局 / 制定」三行装饰（5 段，1.5%），
# 正文在节 4（317 段，92.2%，1 栏）——旧口径取节 1 → 整篇正文排成双栏。
_BODY_COLS_DECOR_SHARE = 0.05
_BODY_COLS_BODY_SHARE = 0.50


def _body_cols_source(pick: dict, counts: list[int], total: int, rule: str) -> dict:
    """正文栏样板的来源节（index + 段落数 + 段落占比 + 取值规则），供判定器对照。"""
    idx = int(pick.get("index") or 0)
    n = counts[idx] if 0 <= idx < len(counts) else 0
    return {
        "index": idx,
        "paras": n,
        "share": round(n / total, 4) if total else 0.0,
        "cols_num": int(pick.get("cols_num") or 1),
        "rule": rule,
    }


def _resolve_body_cols(sects: list[dict], two: Optional[dict], counts: list[int]) -> tuple[dict, dict]:
    """正文栏样板节：默认「第一个非首节 ≥2 栏节」，覆盖率判据可改认正文节。

    ``counts`` 是 ``_section_para_counts`` 的逐节段落数（与 ``sects`` 同序）。

    默认口径（2026-10-17 自动化学报短文）：正文分栏只看非首节 2 栏节；整篇一节时
    才认第 0 节。**该口径本身会错**（EZI-0035）：模板前部另有一个 2 栏装饰节时，
    「第一个非首节 2 栏」是装饰而不是正文。加覆盖率判据——见
    ``_BODY_COLS_DECOR_SHARE`` / ``_BODY_COLS_BODY_SHARE``。

    取不到合法正文节（无段落占比 / 占比长度与节数对不上 / 后面没有 >50% 的 1 栏节）
    时**退回默认口径**：不抛错、不丢值，行为与改动前一致。
    """
    total = sum(counts)
    if two is None:
        pick = sects[-1]
        return pick, _body_cols_source(pick, counts, total, "last_section")
    if len(counts) != len(sects) or total <= 0:
        return two, _body_cols_source(two, counts, total, "first_multi_col")
    pos = int(two.get("index") or 0)
    if not (0 <= pos < len(counts)):
        return two, _body_cols_source(two, counts, total, "first_multi_col")
    share = counts[pos] / total
    if share < _BODY_COLS_DECOR_SHARE:
        cand = None
        for i, s in enumerate(sects):
            if i <= pos or int(s.get("cols_num") or 1) != 1:
                continue
            if counts[i] / total <= _BODY_COLS_BODY_SHARE:
                continue
            if cand is None or counts[i] > counts[int(cand.get("index") or 0)]:
                cand = s
        if cand is not None:
            src = _body_cols_source(cand, counts, total, "coverage")
            src["decor_index"] = pos
            src["decor_share"] = round(share, 4)
            return cand, src
    return two, _body_cols_source(two, counts, total, "first_multi_col")


def _fill_column_object_widths(doc, out: dict) -> None:
    """body/front 栏规格 + 图/表相对栏宽或版心的比例（样本聚类才收）。"""
    sects = list(out.get("sections") or [])
    if not sects:
        return
    # 2026-10-17 真机（自动化学报短文）：模板第 0 节是 2 栏但那是**封面前部**，
    # 正文（末节）其实是 1 栏——按「第一个 2 栏节」取会整篇排成双栏。
    # 口径：正文分栏只看「非首个节」的 2 栏节；整篇只有一节时才认第 0 节。
    # 2026-11-17（EZI-0035）：这个口径仍可能取到**前部装饰** 2 栏节（国管局服务
    # 合同节 1 是三行「制定单位」装饰），由 _resolve_body_cols 的覆盖率判据改认
    # 真正的正文节——否则整篇正文会被排成双栏。
    two = next((s for s in sects
                if int(s.get("cols_num") or 1) >= 2 and int(s.get("index") or 0) > 0), None)
    if two is None and len(sects) == 1 and int(sects[0].get("cols_num") or 1) >= 2:
        two = sects[0]
    counts = _section_para_counts(doc.find(f"{W}body"))
    body_pick, body_src = _resolve_body_cols(sects, two, counts)
    out["body_cols"] = dict(body_pick or sects[-1])
    out["body_cols_source"] = body_src
    out["provenance"].append(
        "body_cols <- section[{}] cols={} {}/{} paras ({:.1%}) rule={}".format(
            body_src["index"], body_src["cols_num"], body_src["paras"],
            sum(counts), body_src["share"],
            body_src["rule"]))
    out["front_cols"] = dict(sects[0])
    # 双栏从第 2+ 节才开始时，记下该节首段原文，供源稿补 continuous 分节
    if two and int(two.get("index") or 0) > 0 and (two.get("start_text") or "").strip():
        raw = two["start_text"].strip()
        out["body_section_anchor"] = raw
        out["provenance"].append(f"body_section_anchor: {raw[:40]}")
        if _ANN_START_RE.search(raw):
            out["body_section_anchor_kind"] = "first_heading"
            out["provenance"].append("body_section_anchor_kind=first_heading (template start is annotation)")
    # 1 栏前部出现过的段首词：分节时跳过（中英摘要/关键词仍算前部）
    front_toks: list[str] = []
    two_start = (two.get("start_text") if two else "") or ""
    for p in doc.iter(f"{W}p"):
        t = "".join(x.text or "" for x in p.iter(f"{W}t")).strip()
        if two_start and t and t[:24] == two_start[:24]:
            break
        tok = anchor_token(t)
        if tok and tok not in front_toks:
            front_toks.append(tok)
        if len(front_toks) >= 24:
            break
    out["front_tokens"] = front_toks
    _fill_section_plan(out)
    page = sects[-1].get("page_content") or 0
    col_w = 0
    if two and two.get("col_widths"):
        col_w = two["col_widths"][0]
    elif sects[-1].get("col_widths"):
        col_w = sects[-1]["col_widths"][0]
    elif page:
        col_w = page

    fig_page, fig_col = [], []
    for ext in doc.iter(f"{WP}extent"):
        try:
            tw = int(ext.get("cx") or 0) / _EMU_PER_TWIP
        except (TypeError, ValueError):
            continue
        if tw < 800:
            continue
        if col_w and int(out["body_cols"].get("cols_num") or 1) >= 2:
            if abs(tw / col_w - 1) <= abs(tw / page - 1) if page else True:
                if 0.4 <= tw / col_w <= 1.15:
                    fig_col.append(tw / col_w)
                    continue
        if page and 0.4 <= tw / page <= 1.15:
            fig_page.append(tw / page)
    if fig_col and _clustered(fig_col):
        out["figure_width"] = {"ref": "column", "ratio": round(_median(fig_col), 3)}
        out["provenance"].append(f"figure_width column ratio={out['figure_width']['ratio']}")
    elif fig_page and _clustered(fig_page) and len(fig_page) >= 2:
        out["figure_width"] = {"ref": "page", "ratio": round(_median(fig_page), 3)}
        out["provenance"].append(f"figure_width page ratio={out['figure_width']['ratio']}")

    tbl_page, tbl_col = [], []
    for tbl in doc.iter(f"{W}tbl"):
        grid = tbl.find(f"{W}tblGrid")
        if grid is None:
            continue
        try:
            total = sum(int(c.get(f"{W}w") or 0) for c in grid.findall(f"{W}gridCol"))
        except (TypeError, ValueError):
            continue
        if total <= 0:
            continue
        if col_w and int(out["body_cols"].get("cols_num") or 1) >= 2:
            if 0.4 <= total / col_w <= 1.15:
                tbl_col.append(total / col_w)
                continue
        if page and 0.4 <= total / page <= 1.15:
            tbl_page.append(total / page)
    if tbl_col and _clustered(tbl_col):
        out["table_width"] = {"ref": "column", "ratio": round(_median(tbl_col), 3)}
        out["provenance"].append(f"table_width column ratio={out['table_width']['ratio']}")
    elif tbl_page and _clustered(tbl_page):
        out["table_width"] = {"ref": "page", "ratio": round(_median(tbl_page), 3)}
        out["provenance"].append(f"table_width page ratio={out['table_width']['ratio']}")


def _fill_table_cells(doc, paras, parse_annotation_props, out: dict) -> None:
    """表内字号/字体：注释「表内容…N号」优先（书面规范），否则示例单元格主导值。"""
    tc: dict[str, Any] = {}
    for p in paras:
        t = _ptxt(p)
        if "表内容" not in t:
            continue
        props = parse_annotation_props(t.split("表内容", 1)[1])
        if props.get("font_size_pt"):
            tc["sz"] = str(int(round(float(props["font_size_pt"]) * 2)))
            out["provenance"].append(f"table_cells sz from annotation: {t[:50]}")
        break
    szs, ascs, eas = [], [], []
    for tbl in doc.iter(f"{W}tbl"):
        for p in tbl.iter(f"{W}p"):
            t = _ptxt(p)
            if not t or t.startswith("（") or t.startswith("("):
                continue
            for r in p.findall(f"{W}r"):
                if not "".join(x.text or "" for x in r.iter(f"{W}t")).strip():
                    continue
                rpr = r.find(f"{W}rPr")
                if rpr is None:
                    continue
                sz = rpr.find(f"{W}sz")
                sc = rpr.find(f"{W}szCs")
                if sz is not None and sz.get(f"{W}val"):
                    szs.append(sz.get(f"{W}val"))
                elif sc is not None and sc.get(f"{W}val"):
                    szs.append(sc.get(f"{W}val"))
                rf = rpr.find(f"{W}rFonts")
                if rf is not None:
                    if rf.get(f"{W}ascii"):
                        ascs.append(rf.get(f"{W}ascii"))
                    if rf.get(f"{W}eastAsia"):
                        eas.append(rf.get(f"{W}eastAsia"))
    if "sz" not in tc and szs:
        tc["sz"] = max(set(szs), key=szs.count)
        out["provenance"].append(f"table_cells sz from cell runs: {tc['sz']}")
    if ascs:
        tc["font_ascii"] = max(set(ascs), key=ascs.count)
    if eas:
        tc["font_east_asia"] = max(set(eas), key=eas.count)
    if tc:
        out["table_cells"] = tc


def _tbl_edge(tbl_borders, name: str) -> dict | None:
    if tbl_borders is None:
        return None
    el = tbl_borders.find(f"{W}{name}")
    if el is None:
        return None
    val = el.get(f"{W}val")
    if val in (None, "nil", "none"):
        return None
    return {
        "val": val,
        "sz": el.get(f"{W}sz") or "4",
        "color": el.get(f"{W}color") or "000000",
    }


def _fill_section_break_words(paras, out: dict) -> None:
    """模板中「标题自带另起页」的章节词集合（2026-10-06 取值出自模板铁律）。

    判定「另起页」= 本段 pageBreakBefore / 前段含显式分页符 / 前段段尾
    sectPr 非 continuous（nextPage/oddPage/evenPage）。只收标题样短文本
    （≤25 字且无句读）——长正文段落即使处于节首也不算章节标题。
    模板全文无分页要求时收空集：下游据此「不发明分页」。
    """
    words: set[str] = set()
    prev = None
    for p in paras:
        starts_new_page = False
        ppr = p.find(f"{W}pPr")
        if ppr is not None and ppr.find(f"{W}pageBreakBefore") is not None:
            starts_new_page = True
        if prev is not None and not starts_new_page:
            for br in prev.iter(f"{W}br"):
                if br.get(f"{W}type") == "page":
                    starts_new_page = True
                    break
            if not starts_new_page:
                ppr_prev = prev.find(f"{W}pPr")
                sp = ppr_prev.find(f"{W}sectPr") if ppr_prev is not None else None
                if sp is not None:
                    tp = sp.find(f"{W}type")
                    starts_new_page = tp is None or tp.get(f"{W}val") != "continuous"
        if starts_new_page:
            key = _norm(_ptxt(p))
            if key and len(key) <= 25 and not any(c in key for c in "。，；？！、："):
                words.add(key)
        prev = p
    out["section_break_words"] = sorted(words)


def _run_props(p) -> dict:
    """段落首个有效 run 的字体/字号（半磅 → 磅）。"""
    from earlyz.spec.decompiler import parse_annotation_props  # noqa: F401
    out: dict = {}
    for r in p.iter(f"{W}r"):
        rpr = r.find(f"{W}rPr")
        if rpr is None:
            continue
        fonts = rpr.find(f"{W}rFonts")
        if fonts is not None:
            if fonts.get(f"{W}eastAsia"):
                out["font_east_asia"] = fonts.get(f"{W}eastAsia")
            if fonts.get(f"{W}ascii"):
                out["font_ascii"] = fonts.get(f"{W}ascii")
        sz = rpr.find(f"{W}sz")
        if sz is not None and sz.get(f"{W}val"):
            try:
                out["font_size_pt"] = round(float(sz.get(f"{W}val")) / 2, 1)
            except ValueError:
                pass
        if out:
            break
    return out


def _maybe_ref_sample(p, text: str, out: dict) -> None:
    """够长、像一条文献（含作者/标点）的段落 → 当作参考文献样例。"""
    t = (text or "").strip()
    if len(t) < 40 or len(t) > 400:
        return
    if t.endswith(":") or t.endswith("："):
        return                      # 「Basic format for periodicals:」这类说明行
    if not any(ch in t for ch in ("，", ",", "“", "”", ".", "、")):
        return
    props = _run_props(p)
    ppr = p.find(f"{W}pPr")
    sp = ppr.find(f"{W}spacing") if ppr is not None else None
    if sp is not None:
        if sp.get(f"{W}line"):
            try:
                out["refs"]["line"] = int(sp.get(f"{W}line"))
            except ValueError:
                pass
        if sp.get(f"{W}lineRule"):
            out["refs"]["line_rule"] = sp.get(f"{W}lineRule")
    if props:
        out["refs"].update(props)
        out["refs"].setdefault("numbered", False)
        out["provenance"].append(f"refs sample paragraph: {t[:40]}")


def _fill_table_borders(doc, out: dict) -> None:
    """从模板表示例表抽边框：无左右/无竖线 + 有顶底 = 三线表，字号跟当前表。"""
    votes: list[dict] = []
    for tbl in doc.iter(f"{W}tbl"):
        tbl_pr = tbl.find(f"{W}tblPr")
        tb = tbl_pr.find(f"{W}tblBorders") if tbl_pr is not None else None
        top = _tbl_edge(tb, "top")
        bot = _tbl_edge(tb, "bottom")
        left = _tbl_edge(tb, "left")
        right = _tbl_edge(tb, "right")
        ih = _tbl_edge(tb, "insideH")
        iv = _tbl_edge(tb, "insideV")
        header_sz = None
        header_color = None
        first = tbl.find(f"{W}tr")
        if first is not None:
            for tc in first.iter(f"{W}tc"):
                tc_pr = tc.find(f"{W}tcPr")
                tcb = tc_pr.find(f"{W}tcBorders") if tc_pr is not None else None
                hb = _tbl_edge(tcb, "bottom") if tcb is not None else None
                if hb:
                    header_sz = hb["sz"]
                    header_color = hb["color"]
                    break
        if top and bot and not left and not right and not iv:
            rec = {
                "style": "three_line",
                "top_sz": top["sz"],
                "bottom_sz": bot["sz"],
                "color": top["color"] if top["color"] != "auto" else "000000",
            }
            hs = header_sz or (ih["sz"] if ih else None)
            if hs:
                rec["header_sz"] = hs
            votes.append(rec)
        elif left or right or iv:
            votes.append({"style": "solid", "sz": (left or right or iv or top or {}).get("sz") or "4",
                          "color": "000000"})
    if not votes:
        return
    styles = [v["style"] for v in votes]
    kind = max(set(styles), key=styles.count)
    chosen = [v for v in votes if v["style"] == kind][0]
    out["table_borders"] = chosen
    out["provenance"].append(f"table_borders {kind} from {len(votes)} template table(s)")


def _fill_drop_cap_layout(doc, out: dict) -> None:
    """首字下沉需保真的 OOXML 组合；只从当前模板首字段落读取。"""
    for p in doc.iter(f"{W}p"):
        ppr = p.find(f"{W}pPr")
        frame = ppr.find(f"{W}framePr") if ppr is not None else None
        if frame is None or not frame.get(f"{W}dropCap"):
            continue
        rec: dict[str, Any] = {
            "frame": {k.split("}")[-1]: v for k, v in frame.attrib.items()},
        }
        ppr_rpr = ppr.find(f"{W}rPr")
        run_rpr = p.find(f"{W}r/{W}rPr")
        if ppr_rpr is not None:
            rec["ppr_rpr_xml"] = etree.tostring(ppr_rpr, encoding="unicode")
        if run_rpr is not None:
            rec["run_rpr_xml"] = etree.tostring(run_rpr, encoding="unicode")
        pbdr = ppr.find(f"{W}pBdr")
        if pbdr is not None:
            rec["pbdr_xml"] = etree.tostring(pbdr, encoding="unicode")
        ta = ppr.find(f"{W}textAlignment")
        if ta is not None and ta.get(f"{W}val"):
            rec["text_alignment"] = ta.get(f"{W}val")
        out["drop_cap_layout"] = rec
        out["provenance"].append("drop_cap_layout from template paragraph OOXML")
        return


def structural_params(spec) -> dict[str, Any]:
    """FormatSpec → 结构段 action 参数（缺键=不做取值，禁止学校默认）。"""
    st = getattr(spec, "structure", None) or {}
    header = {}
    body_ea = None
    for r in getattr(spec, "rules", []) or []:
        sel = r.selector or {}
        if sel.get("scope") == "header_all":
            header = dict(r.expected or {})
        if sel.get("semantic_role") == "body":
            body_ea = (r.expected or {}).get("font_east_asia")
    cap = dict(st.get("caption") or {})
    if cap and not cap.get("figure_east_asia") and body_ea:
        cap["figure_east_asia"] = body_ea
    for r in getattr(spec, "rules", []) or []:
        role = (r.selector or {}).get("semantic_role")
        if role == "caption_table":
            if (r.expected or {}).get("jc"):
                cap["table_jc"] = r.expected["jc"]
            if (r.expected or {}).get("font_size_pt"):
                cap["table_sz"] = str(int(round(float(r.expected["font_size_pt"]) * 2)))
            if (r.expected or {}).get("font_east_asia"):
                cap["table_east_asia"] = r.expected["font_east_asia"]
        if role == "caption_figure" and (r.expected or {}).get("jc"):
            cap["figure_jc"] = r.expected["jc"]
        if role == "caption_figure" and (r.expected or {}).get("font_size_pt") and not cap.get("figure_sz"):
            cap["figure_sz"] = str(int(round(float(r.expected["font_size_pt"]) * 2)))
    out = {
        "toc": st.get("toc") or {},
        "toc_title": dict(st.get("toc_title") or {}),
        "caption": cap,
        "footnote": dict(st.get("footnote") or {}),
        "footnote_restart": st.get("footnote_restart") or None,
        "refs": st.get("refs") or {},
        "sections": st.get("sections") or [],
        # 规则（decompiler header_all）优先，structure 兜底补缺（jc/边框等）
        "header_format": ({**(st.get("header_format") or {}), **header}
                          if (header or st.get("header_format")) else {}),
        "header_footer_parts": list(st.get("header_footer_parts") or []),
        "hf_bound_slots": [list(s) for s in (st.get("hf_bound_slots") or [])],
        "identity_headers": list(st.get("identity_headers") or []),
        "body_headers": dict(st.get("body_headers") or {}),
        "front": dict(st.get("front") or {}),
        "front_cols": dict(st.get("front_cols") or {}),
        "body_cols": dict(st.get("body_cols") or {}),
        "body_cols_source": dict(st.get("body_cols_source") or {}),
        "body_section_anchor": st.get("body_section_anchor") or "",
        "body_section_anchor_kind": st.get("body_section_anchor_kind") or "",
        "front_tokens": list(st.get("front_tokens") or []),
        "section_plan": list(st.get("section_plan") or []),
        "equation": dict(st.get("equation") or {}),
        "cover": dict(st.get("cover") or {}),
        "figure_width": dict(st.get("figure_width") or {}),
        "table_width": dict(st.get("table_width") or {}),
        "table_cells": dict(st.get("table_cells") or {}),
        "table_borders": dict(st.get("table_borders") or {}),
        "drop_cap_layout": dict(st.get("drop_cap_layout") or {}),
        "even_and_odd": bool(st.get("even_and_odd")),
        "front_matter_roman": bool(st.get("front_matter_roman")),
        "provenance": list(st.get("provenance") or []),
    }
    # 分页要求只从模板取（2026-10-06 铁律）：键缺失=无模板信息（下游保持
    # 旧口径兜底），空集=模板没要求=不得发明分页——两者语义不同，禁 or[] 合并。
    if "section_break_words" in st:
        out["section_break_words"] = list(st.get("section_break_words") or [])
    return out


def pt_to_sz(pt: Optional[float]) -> Optional[str]:
    if pt is None:
        return None
    return str(int(round(float(pt) * 2)))
