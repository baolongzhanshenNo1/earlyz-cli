"""早睡 (EarlyZ) 诊断指纹导出（2026-08-22 用户提议：用户不用脱敏文件，跑一次排版，通过 log/指纹反解）。

用途：用户（Word 小白）把文档在 早睡 (EarlyZ) 排一次版或直接跑本工具，产出
diagnose.json —— 纯结构化属性指纹（**不含任何正文文本**，隐私安全），
发回给开发侧反解串页/排版根因。

运行：python -m earlyz.tools.diagnose <docx|earlyz 结果目录> [--out diagnose.json]

输出结构见 build_diagnose() 的 docstring。机制层统计，无学校/刊物常量。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
W = NS["w"]


def _x(root, tag):
    return root.find(f"{{{NS['w']}}}{tag}")


def _ptxt(p):
    return "".join(t.text or "" for t in p.iter(f"{{{W}}}t"))


def _sig(text: str) -> str:
    """文本 → 指纹哈希（防原文外泄；长度也隐藏到 8 位 hash）。"""
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:8]


def _read_part(zipf: zipfile.ZipFile, name: str) -> bytes:
    try:
        return zipf.read(name)
    except KeyError:
        return b""


def _spacing_sig(ppr) -> str | None:
    """段落行距/段距指纹（机制层：exact 240 / auto / atLeast 280 等组合签名）。"""
    sp = ppr.find(f"{{{W}}}spacing")
    if sp is None:
        return None
    parts = []
    for attr in ("w:line", "w:lineRule", "w:before", "w:after", "w:firstLine"):
        val = sp.get(attr.replace("w:", f"{{{W}}}"), sp.get(f"{{{W}}}{attr.split(':')[1]}"))
        if val is not None:
            parts.append(f"{attr.split(':')[1]}={val}")
    return ";".join(parts) if parts else None


def _table_border_profile(tbl) -> str:
    """表格边框模式指纹：tblBorders 的六向+线条类型 → 模式归类。

    归类（机制层）：
    - three_line   : 顶/底+表头线（top/bottom/insideH 非 none）
    - bottom_only  : 仅 bottom 非 none
    - full_frame   : 六向全非 none（single/double 等）
    - all_none     : 全部 none
    - mixed_other  : 其他组合
    """
    tblpr = _x(tbl, "tblPr")
    if tblpr is None:
        return "no_tblPr"
    borders = tblpr.find(f"{{{W}}}tblBorders")
    if borders is None:
        return "no_borders_tag"
    vals = {}
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.find(f"{{{W}}}{side}")
        if el is not None:
            val = el.get(f"{{{W}}}val") or ""
            if val == "none":
                vals[side] = "none"
                continue
            sz = el.get(f"{{{W}}}sz")
            vals[side] = f"{val}:{sz}" if val else "none"
        else:
            vals[side] = "inherit"
    non_none = [s for s, v in vals.items() if v not in ("none", "inherit")]
    if not non_none:
        return "all_none"
    if non_none == ["top", "bottom"] or "insideH" in non_none and len(non_none) <= 3:
        return "three_line"
    if non_none == ["bottom"]:
        return "bottom_only"
    if set(non_none) == {"top", "bottom", "left", "right", "insideH", "insideV"}:
        return "full_frame"
    return "mixed_other"


def _redacted_name(docx_path: Path) -> str:
    """真实文件名 → 脱敏标识 `doc-<6位哈希><扩展名>`（同一份文档稳定、不可反解）。

    2026-09-27 用户拍板：诊断包会上传到服务端，真实文件名可能含客户名/案由/人名，
    且对排查无用（定位文档靠短码 + 用户描述 + 截图）。
    """
    name = docx_path.name
    base = name[: -len("-candidate.docx")] if name.endswith("-candidate.docx") else docx_path.stem
    digest = hashlib.sha256(base.encode("utf-8")).hexdigest()[:6]
    return f"doc-{digest}{docx_path.suffix or '.docx'}"


def build_diagnose(
    docx_path: Path,
    journal: dict | None = None,
    template_path: Path | None = None,
    skip_pagination: bool = False,
) -> dict:
    """构建诊断指纹（不含正文文本）。
    2026-08-23 v3：template_path 提供时输出模板指纹 + differences（产物 vs
    模板差异：节数/栏宽/图片超栏/页眉页脚槽位/行距/表格 cantSplit）。"""
    docx_path = Path(docx_path)
    # 2026-08-23 v3.1：journal 未显式传入时自动从产物同名 journal 附带
    # （compile 产物与 tx_xxx-journal.json 同目录——诊断包带 compile.actions）
    if journal is None:
        _jf = docx_path.with_name(
            docx_path.name.replace("-candidate.docx", "-journal.json"))
        if _jf.exists():
            try:
                journal = json.loads(_jf.read_text(encoding="utf-8"))
            except Exception:
                journal = None
    out: dict = {
        "schema": "earlyz-diagnose-v3.2",  # 2026-09-27 v3.2：文件名脱敏（只留扩展名 + 稳定短哈希）
        # 文件名脱敏（2026-09-27 用户拍板）：诊断包会被上传到服务端，真实文件名可能含客户名/案由/人名，
        # 而它对排查无用（定位文档靠短码 + 用户描述 + 截图）。保留「扩展名 + 名称稳定短哈希」→
        # 同一份文档多次上传仍可关联，但不可反解。
        "file": {
            "name": _redacted_name(docx_path),
            "bytes": docx_path.stat().st_size,
            "redacted": True,
        },
    }
    with zipfile.ZipFile(docx_path) as z:
        doc_bytes = _read_part(z, "word/document.xml")
        root = None
        if doc_bytes:
            from lxml import etree
            root = etree.fromstring(doc_bytes)

        # ---------- 节 ----------
        sectPrs = list(root.iter(f"{{{W}}}sectPr")) if root is not None else []
        pg_sz_sigs = []
        pg_mar_sigs = []
        for sp in sectPrs:
            pgsz = sp.find(f"{{{W}}}pgSz")
            pgmar = sp.find(f"{{{W}}}pgMar")
            pg_sz_sigs.append(
                f"{pgsz.get(f'{{{W}}}w')}x{pgsz.get(f'{{{W}}}h')}" if pgsz is not None else ""
            )
            pg_mar_sigs.append(
                "|".join(
                    str(pgmar.get(f"{{{W}}}{a}") or "")
                    for a in ("top", "right", "bottom", "left", "header", "footer")
                )
                if pgmar is not None else ""
            )
        out["sections"] = {
            "count": len(sectPrs),
            "pgSz_signatures": counts(pg_sz_sigs),
            "pgMar_signatures": counts(pg_mar_sigs),
        }
        # 栏宽指纹（2026-08-23 v3：图片是否超栏的判定依据）
        col_widths = []
        for sp in sectPrs:
            cols = sp.find(f"{{{W}}}cols")
            if cols is None:
                continue
            num = int(cols.get(f"{{{W}}}num") or "1")
            if num == 1:
                continue
            col_els = list(cols.iter(f"{{{W}}}col"))
            if col_els:
                for ce in col_els:
                    col_widths.append(int(ce.get(f"{{{W}}}w") or 0))
            else:
                space = int(cols.get(f"{{{W}}}space") or 0)
                pgsz = sp.find(f"{{{W}}}pgSz")
                pgmar = sp.find(f"{{{W}}}pgMar")
                if pgsz is not None and pgmar is not None:
                    cw = (int(pgsz.get(f"{{{W}}}w") or 0)
                          - int(pgmar.get(f"{{{W}}}left") or 0)
                          - int(pgmar.get(f"{{{W}}}right") or 0)
                          - space * (num - 1)) // num
                    if cw > 0:
                        col_widths.append(cw)
        out["sections"]["col_widths"] = col_widths

        # ---------- 段落（正文） ----------
        if root is not None:
            paras = list(root.iter(f"{{{W}}}p"))
            spacing_prof = Counter()
            outline = Counter()
            style_usage = Counter()
            empty = 0
            for p in paras:
                ppr = p.find(f"{{{W}}}pPr")
                t = _ptxt(p)
                if not t:
                    empty += 1
                sig = _spacing_sig(ppr) if ppr is not None else None
                spacing_prof[sig if sig is not None else "<no-spacing>"] += 1
                if ppr is not None:
                    olv = ppr.find(f"{{{W}}}outlineLvl")
                    if olv is not None:
                        outline[olv.get(f"{{{W}}}val", "0")] += 1
                    ps = ppr.find(f"{{{W}}}pStyle")
                    if ps is not None:
                        style_usage[ps.get(f"{{{W}}}val", "?")] += 1
            out["paragraphs"] = {
                "count": len(paras),
                "empty": empty,
                "spacing_profile": dict(spacing_prof.most_common(25)),
                "outline_lvl": dict(outline),
                "pStyle_usage": dict(style_usage.most_common(25)),
                "text_bytes": sum(len(t) for t in (_ptxt(p) for p in paras)),
                "text_hash": _sig("".join(sort_hashes(_ptxt(p) for p in paras if _ptxt(p)))),
                # ---- 2026-08-23 v3.1 串页场景加强：分页控制/分页符/分区块 ----
                # 串页主嫌疑①=第 2-100 页与首页行距/段距不同（粘贴回落 Word
                # 默认 1.08/1.15 → 每页多 2-3 行溢出）；②=缺 keepNext/keepLines
                # 分页控制；⑥=手工分页符散布（pPr pageBreakBefore / br type=page）。
                "keep_profile": {
                    "keep_next": sum(
                        1 for p in paras
                        if (ppr := p.find(f"{{{W}}}pPr")) is not None
                        and ppr.find(f"{{{W}}}keepNext") is not None),
                    "keep_lines": sum(
                        1 for p in paras
                        if (ppr := p.find(f"{{{W}}}pPr")) is not None
                        and ppr.find(f"{{{W}}}keepLines") is not None),
                    "page_break_before": sum(
                        1 for p in paras
                        if (ppr := p.find(f"{{{W}}}pPr")) is not None
                        and ppr.find(f"{{{W}}}pageBreakBefore") is not None),
                },
                "manual_page_breaks": sum(
                    1 for br in root.iter(f"{{{W}}}br")
                    if br.get(f"{{{W}}}type") == "page"),
                # 分区块行距指纹：每 100 段一桶——桶间分布不同 = 格式不统一
                # 区块（页级模板场景：0-99 桶=首页块，其余=后续页）。
                "spacing_zones": [
                    {
                        "seg": f"{i}-{min(i + 99, len(paras) - 1)}",
                        "profile": dict(Counter(
                            _spacing_sig(p.find(f"{{{W}}}pPr"))
                            if p.find(f"{{{W}}}pPr") is not None else None
                            for p in paras[i:i + 100]
                        ).most_common(8)),
                    }
                    for i in range(0, len(paras), 100)
                ],
            }

            # ---------- 域预检（2026-08-23 v3.1：域缺陷提前发现） ----------
            # PAGEREF/REF 书签缺失 → Word 界面靠缓存、导出/打印才爆
            # 「错误!未找到引用源」——诊断预置 REF/PAGE 域扫描。
            _field_instr = [
                re.sub(r"[^A-Za-z0-9 _\\&:]", "", (t.text or ""))
                for t in root.iter(f"{{{W}}}instrText")
            ]
            _ref_kinds = Counter()
            for txt in _field_instr:
                for kind in re.findall(r"^(PAGEREF|REF|PAGE|NUMPAGES|TOC|SEQ|SYMBOL)", txt.strip()):
                    _ref_kinds[kind] += 1
                    break
            _bookmark_starts = sum(
                1 for b in root.iter(f"{{{W}}}bookmarkStart"))
            _bookmark_ends = sum(1 for b in root.iter(f"{{{W}}}bookmarkEnd"))
            _bad_ref_text = len(re.findall(
                r"未找到引用源|错误!|錯誤!|Error!Bookmark", _ptxt(root) if False else
                "".join(t.text or "" for t in root.iter(f"{{{W}}}t"))))
            out["fields"] = {
                "census": dict(_ref_kinds.most_common(12)),
                "bookmark_start": _bookmark_starts,
                "bookmark_end": _bookmark_ends,
                "bookmark_imbalance": _bookmark_starts - _bookmark_ends,
                "dirty_fields": sum(
                    1 for fc in root.iter(f"{{{W}}}fldChar")
                    if fc.get(f"{{{W}}}dirty") == "true"),
                "broken_ref_cache": _bad_ref_text,
            }

            # ---------- 表格 ----------
            tbls = list(root.iter(f"{{{W}}}tbl"))
            border_prof = Counter()
            grid_zero = 0
            cant_split_missing = 0
            tbl_count = len(tbls)
            for tb in tbls:
                border_prof[_table_border_profile(tb)] += 1
                for gc in tb.iter(f"{{{W}}}gridCol"):
                    if gc.get(f"{{{W}}}w") == "0":
                        grid_zero += 1
                for tr in tb.iter(f"{{{W}}}tr"):
                    trpr = tr.find(f"{{{W}}}trPr")
                    if trpr is None or trpr.find(f"{{{W}}}cantSplit") is None:
                        cant_split_missing += 1
            out["tables"] = {
                "count": tbl_count,
                "border_profiles": dict(border_prof),
                "gridCol_zero": grid_zero,
                "tr_count": sum(
                    1 for tb in tbls for _ in tb.iter(f"{{{W}}}tr")),
                "tr_without_cantSplit": cant_split_missing,
            }

            # ---------- 图片 ----------
            _wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            docpr_ids = []
            extents = Counter()
            inline_count = 0
            anchor_count = 0
            for dp in root.iter(f"{{{_wp}}}docPr"):
                docpr_ids.append(dp.get("id"))
            for d in root.iter(f"{{{_wp}}}inline"):
                inline_count += 1
                ext = d.find(f"{{{_wp}}}extent")
                if ext is not None:
                    w = int(ext.get("cx", "0")) // 6350  # EMU→0.01cm 粗指纹
                    h = int(ext.get("cy", "0")) // 6350
                    extents[f"{w / 100:.2f}x{h / 100:.2f}cm"] += 1
            for d in root.iter(f"{{{_wp}}}anchor"):
                anchor_count += 1
            dup_ids = [i for i, c in Counter(docpr_ids).items() if c > 1]
            out["images"] = {
                "inline": inline_count,
                "anchor": anchor_count,
                "extent_fingerprints": dict(extents.most_common(15)),
                "docPr_ids_unique": not dup_ids,
                "docPr_id_dups": dup_ids[:10],
            }

            # ---------- 页眉页脚 ----------
            # 2026-08-23 修复：headerReference/footerReference 的 r:id 在
            # officeDocument/relationships ns（r:），旧代码取 w: 命名空间 →
            # 全部 "?" 丢失。升级为槽位语义化：type 分布 + part 内容摘要
            # （文本核 hash+长度，无原文）+ 空 part 判定——诊断包可直接看出
            # 「模板有内容、产物空槽」类问题。
            _r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            hf_entries = []
            _r_pkg = "http://schemas.openxmlformats.org/package/2006/relationships"
            rels_map = {}
            try:
                rels_root = etree.fromstring(
                    _read_part(z, "word/_rels/document.xml.rels"))
                for rel in rels_root.iter(f"{{{_r_pkg}}}Relationship"):
                    rels_map[rel.get("Id")] = rel.get("Target")
            except Exception:
                pass
            for ref_el in root.iter(f"{{{W}}}headerReference"):
                _rid = ref_el.get(f"{{{_r}}}id") or "?"
                _typ = ref_el.get(f"{{{W}}}type") or "default"
                hf_entries.append(("header", _typ, _rid))
            for ref_el in root.iter(f"{{{W}}}footerReference"):
                _rid = ref_el.get(f"{{{_r}}}id") or "?"
                _typ = ref_el.get(f"{{{W}}}type") or "default"
                hf_entries.append(("footer", _typ, _rid))
            hf_slots = []
            for kind, typ, rid in hf_entries:
                part = rels_map.get(rid) or ""
                if part and not part.startswith("word/"):
                    part = f"word/{part}"
                signed = ""
                empty = False
                if part:
                    try:
                        proot = etree.fromstring(_read_part(z, part))
                        core = re.sub(
                            r"[\s\u3000]+", "",
                            "".join(t.text or "" for t in proot.iter(f"{{{W}}}t")),
                        )
                        # 2026-11-10（IEEE 英文期刊实锤）：只有页码域的页眉
                        # 不算空——判空按「文字」时会把「模板有页码、产物有页码」
                        # 误报成「模板有内容、产物为空」。页码域=页眉内容。
                        _has_page_fld = any(
                            (t.text or "").strip().upper().startswith("PAGE")
                            and "NUMPAGES" not in (t.text or "").upper()
                            for t in proot.iter(f"{{{W}}}instrText")
                        )
                        if not core and not _has_page_fld:
                            empty = True
                        elif not core and _has_page_fld:
                            signed = "page-field"
                        else:
                            signed = (
                                "h"
                                + hashlib.sha256(core.encode()).hexdigest()[:6]
                                + "c"
                                + str(len(core))
                            )
                    except Exception:
                        signed = "unreadable"
                hf_slots.append({
                    "kind": kind,
                    "type": typ,
                    "rid": _sig(rid),
                    "empty": empty,
                    "text_sig": signed,
                })
            # ---------- 页眉页脚 part 结构指纹（2026-08-23 v3.1） ----------
            # 页级模板/学报首页页眉常为两行结构（文本行+br+文本框/PAGE 域）
            # ——part 级结构摘要帮助判定页眉是否需要 normalize 重建。
            _hf_struct = {}
            for kind, typ, rid in hf_entries:
                part = rels_map.get(rid) or ""
                if part and not part.startswith("word/"):
                    part = f"word/{part}"
                if part and part not in _hf_struct:
                    try:
                        proot = etree.fromstring(_read_part(z, part))
                        _hf_struct[part] = {
                            "paras": sum(1 for el in proot.iter(f"{{{W}}}p")),
                            "textboxes": sum(1 for el in proot.iter(f"{{{W}}}txbxContent")),
                            "tables": sum(1 for el in proot.iter(f"{{{W}}}tbl")),
                            "page_fields": sum(
                                1 for t in proot.iter(f"{{{W}}}instrText")
                                if (t.text or "").strip() in ("PAGE", "PAGE \\* MERGEFORMAT")),
                            "brs": sum(1 for br in proot.iter(f"{{{W}}}br")),
                            "signed": _sig(part),
                        }
                    except Exception:
                        _hf_struct[part] = {"paras": -1, "note": "unreadable"}
            out["headers_footers"] = {
                "ref_count": len(hf_entries),
                "slot_summary": Counter(f"{k}/{t}" for k, t, _ in hf_entries),
                "slots": hf_slots[:40],
                "part_structure": _hf_struct,
            }

        # ---------- 分页（物理页数，Word COM 可选；模板对照时跳过） ----------
        out["pagination"] = {"physical_pages": None, "note": "Word COM 检测未运行"}
        if not skip_pagination:
            try:
                from earlyz.validation.renderers import count_doc_pages
                out["pagination"]["physical_pages"] = count_doc_pages(docx_path)
                out["pagination"]["note"] = "Word COM Repaginate"
            except Exception as exc:  # noqa: BLE001
                out["pagination"]["note"] = f"Word COM 不可用: {type(exc).__name__}"

    # ---------- 编译结果（若提供 journal） ----------
    if journal:
        acts = []
        for a in journal.get("actions", []):
            acts.append({
                "idx": a.get("index"),
                "action": a.get("action"),
                "status": "ok" if not a.get("error") else "failed",
                "info": str(a.get("info") or "")[:160],
            })
        out["compile"] = {
            "transaction_status": journal.get("status"),
            "actions": acts,
            "validation_level": journal.get("validation_level"),
            "render_pages": journal.get("render_pages"),
        }

    # ---------- 模板对照（2026-08-23 v3：诊断包 = 差异报告） ----------
    if template_path:
        try:
            tf = build_diagnose(template_path, skip_pagination=True)
            out["template"] = {
                "sections_count": (tf.get("sections") or {}).get("count"),
                "sections": tf.get("sections"),
                "spacing_profile": (tf.get("paragraphs") or {}).get("spacing_profile"),
                "images": tf.get("images"),
                "headers_footers": tf.get("headers_footers"),
                "tables": tf.get("tables"),
            }
            diffs = []
            doc_s = out.get("sections") or {}
            tpl_s = (tf.get("sections") or {})
            doc_sc = doc_s.get("count")
            tpl_sc = tpl_s.get("count")
            if doc_sc != tpl_sc:
                diffs.append(f"节数: doc={doc_sc} tpl={tpl_sc}")
            # 图片超栏（模板栏宽 vs 图片宽；无模板栏宽时用自身栏宽）
            tcols = tpl_s.get("col_widths") or doc_s.get("col_widths")
            col_cm = None
            if isinstance(tcols, list) and tcols:
                col_cm = max(tcols) / 567  # 567 twips = 1cm
            for fp, cnt in (out.get("images") or {}).get("extent_fingerprints", {}).items():
                try:
                    w_cm = float(fp.split("x")[0])
                except Exception:
                    continue
                if col_cm and w_cm > col_cm:
                    diffs.append(
                        f"图片 {fp} 宽 {w_cm:.2f}cm > 栏宽 {col_cm:.2f}cm ×{cnt}（可能被裁）")
            # 槽位：模板有、产物缺失/空
            t_hf = (tf.get("headers_footers") or {}).get("slot_summary") or {}
            d_hf = (out.get("headers_footers") or {}).get("slot_summary") or {}
            for slot in sorted(t_hf):
                if slot not in d_hf:
                    diffs.append(f"模板槽位 {slot} 在产物缺失")
            t_slots = (tf.get("headers_footers") or {}).get("slots") or []
            d_slots = (out.get("headers_footers") or {}).get("slots") or []
            for ts in t_slots:
                if ts.get("empty"):
                    continue
                # 产物同 type 槽位是否全空
                same = [s for s in d_slots if s.get("type") == ts.get("type")]
                if same and all(s.get("empty") for s in same):
                    diffs.append(
                        f"页眉页脚槽位 {ts.get('kind')}/{ts.get('type')} 模板有内容、产物为空")
            # 表格 cantSplit 比例对比（绝对数不可比——模板表行数≠产物表行数；
            # 按「有 cantSplit 行占比」比较：模板设计无 cantSplit 时产物同构=正常）
            t_cant = ((tf.get("tables") or {}).get("tr_without_cantSplit"))
            d_cant = ((out.get("tables") or {}).get("tr_without_cantSplit"))
            t_trs = ((tf.get("tables") or {}).get("tr_count") or 1)
            d_trs = ((out.get("tables") or {}).get("tr_count") or 1)
            if isinstance(t_cant, int) and isinstance(d_cant, int) and t_trs and d_trs:
                t_ratio = (t_trs - t_cant) / t_trs
                d_ratio = (d_trs - d_cant) / d_trs
                if d_ratio < t_ratio - 0.3:
                    diffs.append(
                        f"表格 cantSplit 行占比: doc={d_ratio:.0%} < tpl={t_ratio:.0%}（模板行防跨页未传播）")
            # 行距概要差异（排除 <no-spacing>——模板正文行距常在样式层，
            # 段级无显式 spacing 属模板设计，不构成差异）
            t_sp = ((tf.get("paragraphs") or {}).get("spacing_profile") or {})
            d_sp = ((out.get("paragraphs") or {}).get("spacing_profile") or {})
            t_sp = {k: v for k, v in t_sp.items() if k != "<no-spacing>"}
            d_sp = {k: v for k, v in d_sp.items() if k != "<no-spacing>"}
            if t_sp and d_sp:
                t_ratio = max(t_sp.values()) / max(1, sum(t_sp.values()))
                d_ratio = max(d_sp.values()) / max(1, sum(d_sp.values()))
                if d_ratio < t_ratio - 0.2:
                    diffs.append("正文行距分布与模板差异明显（模板行距未充分套用）")
            out["differences"] = diffs
        except Exception as exc:  # noqa: BLE001
            out["template"] = None
            out["differences"] = [f"模板对照失败: {type(exc).__name__}: {exc}"]

    return out


WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


def counts(values: list) -> dict:
    return dict(Counter(v for v in values if v).most_common(15))


def sort_hashes(texts) -> list:
    """有序文本的哈希列表（用于整体指纹，文本不落盘）。"""
    return sorted(_sig(t) for t in texts)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="早睡 (EarlyZ) 诊断指纹导出（不含正文文本）")
    ap.add_argument("path", help="docx 或包含 tx_*-journal.json 的结果目录")
    ap.add_argument("--out", default=None, help="输出 json 路径（默认 stdout）")
    ap.add_argument("--template", default=None,
                    help="规范模板 docx（可选）：输出模板指纹+差异段（诊断包 v3）")
    args = ap.parse_args(argv)

    p = Path(args.path)
    journal = None
    if p.is_dir():
        js = sorted(p.glob("tx_*-journal.json"), key=lambda x: x.stat().st_mtime)
        if not js:
            print(json.dumps({"schema": "earlyz-diagnose-v1", "error": "目录中无 journal"},
                             ensure_ascii=False, indent=2))
            return 1
        journal = json.loads(js[-1].read_text(encoding="utf-8"))
        from lxml import etree
        candidates = sorted(p.glob("tx_*-candidate.docx"), key=lambda x: x.stat().st_mtime)
        docx = candidates[-1] if candidates else None
        if docx is None:
            print(json.dumps({"schema": "earlyz-diagnose-v1", "error": "目录中无候选文档"},
                             ensure_ascii=False, indent=2))
            return 1
        p = docx

    try:
        diag = build_diagnose(p, journal, template_path=args.template)
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"schema": "earlyz-diagnose-v1", "error": f"{type(exc).__name__}: {exc}"},
                         ensure_ascii=False, indent=2))
        return 1

    text = json.dumps(diag, ensure_ascii=False, indent=2)
    if args.out:
        Path(args.out).write_text(text, encoding="utf-8")
        print(f"diagnose saved: {args.out}")
    else:
        print(text)
    return 0


def attach_explain(out: dict, docx_path: Path | None = None,
                   template_path: Path | None = None,
                   timeout_s: float = 90.0) -> bool:
    """v3.2：LLM 撰写 explain/engine_report（解释层——只读分析，铁律 5）。

    智能触发（成本铁律）：仅当发现「可解释疑点」才调 LLM：
      - differences 非空（模板差异）/ fields 域缺陷 / 分页异常 / keep_profile 异常
    正常产物零 LLM 调用。

    隐私约束（用户拍板 2026-08-23）：LLM 输入不含文件名/真实模板名；
    repro 字段仅「通用词 + 内容 hash」（不可逆指纹，开发侧靠 trigger+evidence+指纹复现）。

    失败/超时/无配置：返回 False（诊断包保持完整，explain 缺省）。
    """
    diff = out.get("differences") or {}
    fields = out.get("fields") or {}
    pag = out.get("pagination") or {}
    paras = out.get("paragraphs") or {}
    # 触发判定：模板差异 / 域缺陷（未找到引用源、书签失衡） / 物理页异常 / keep 异常
    triggered = bool(diff if isinstance(diff, dict) else {})
    if fields:
        if fields.get("broken_ref_cache", 0) or fields.get("dirty", 0) > 2 or \
           fields.get("bookmark_mismatch", 0):
            triggered = True
    if pag.get("physical_pages", 0) > 0 and (pag.get("physical_pages", 0) > 400 or
                                             pag.get("page_index_overrun", False)):
        triggered = True
    if paras.get("keep_profile") and paras["keep_profile"].get("keep_next", 0) == 0 and \
       paras.get("count", 0) > 60:
        triggered = True
    if not triggered:
        return False

    # ── 紧凑特征（无任何可识别文件名；repro 用内容 hash）──
    def _hash_of(p: Path | None) -> str:
        try:
            import hashlib
            h = hashlib.sha256()
            with open(p, "rb") as f:
                for chunk in iter(lambda: f.read(262144), b""):
                    h.update(chunk)
            return h.hexdigest()[:16]
        except Exception:
            return "unknown"

    doc_hash = _hash_of(docx_path)
    tpl_hash = _hash_of(template_path)
    compact = {
        "schema": out.get("schema"),
        "file": {"bytes": (out.get("file") or {}).get("bytes")},
        "sections": {"count": (out.get("sections") or {}).get("count"),
                     "pgSz_margins_diversity": len((out.get("sections") or {}).get("pgSz", []))},
        "paragraphs": {"count": paras.get("count"),
                       "keep_profile": paras.get("keep_profile"),
                       "spacing_zone_diff": (paras.get("spacing_zones") or [{}])[0].get("profile")},
        "tables": {"count": (out.get("tables") or {}).get("count"),
                   "border_profiles": (out.get("tables") or {}).get("border_profiles"),
                   "tr_without_cantSplit": (out.get("tables") or {}).get("tr_without_cantSplit")},
        "images": {"count": (out.get("images") or {}).get("count"),
                   "docPr_dup": (out.get("images") or {}).get("docPr_ids_unique") is False},
        "fields": {"broken_ref_cache": fields.get("broken_ref_cache"),
                   "bookmark_mismatch": fields.get("bookmark_mismatch"),
                   "dirty": fields.get("dirty")},
        "headers_footers": {"slot_summary": (out.get("headers_footers") or {}).get("slot_summary"),
                            "part_count": len((out.get("headers_footers") or {}).get("part_structure", {}))},
        "compile": (out.get("compile") or {}).get("actions") and
                   {"status": (out.get("compile") or {}).get("status")},
        "differences": diff,
        "pagination": pag,
    }
    compact["file"]["doc_hash"] = doc_hash
    compact["file"]["tpl_hash"] = tpl_hash

    sys_prompt = (
        "你是「早睡」(EarlyZ) 排版诊断分析师。输入是文档的机构化诊断指纹（如实事实，"
        "无文件名）。你的任务：基于指纹撰写解释——但只引用指纹中存在的证据，"
        "禁止臆测；判断每个发现属于：engine_bug（早睡 (EarlyZ) 引擎缺陷）/ "
        "template_mismatch（模板差异）/ user_doc（用户文档问题）/ ok（正常）。"
        "输出严格 JSON，仅两个键：\n"
        '{"explain": {"model": "...", "summary": "一句话人话总结", "severity": '
        '"error|warning|info|ok", "findings": [{"issue": "...", "root_cause": "...", '
        '"evidence": ["..."], "classification": "...", "confidence": 0.9, '
        '"template_delta": "..."}]}, '
        '"engine_report": {"classification": "...", "confidence": 0.9, "title": "...", '
        '"root_cause": "...", "evidence": ["..."], "suspected_module": "...", '
        '"repro": {"template": "用户上传模板（tpl-hash: ...）", "source": "用户文档'
        '（doc-hash: ...）", "trigger": "触发条件描述"}, "suggested_fix": "...", '
        '"related_version": "..."}}。'
        "若证据不足，classification 用 user_doc/ok，confidence 小于 0.8。"
    )
    user_prompt = "诊断指纹：\n" + json.dumps(compact, ensure_ascii=False, default=str)

    try:
        sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent / "api"))
        from llm_client import LLMClient
        msg = LLMClient(timeout=timeout_s).chat(
            [{"role": "system", "content": sys_prompt},
             {"role": "user", "content": user_prompt}],
            temperature=0, model=None)
        content = (msg or {}).get("content", "") or ""
        # 提取 JSON（LLM 可能包 ```json 围栏）
        m = re.search(r"\{.*\}", content, re.S)
        if not m:
            return False
        parsed = json.loads(m.group(0))
        if isinstance(parsed, dict):
            if isinstance(parsed.get("explain"), dict):
                parsed["explain"]["model"] = parsed["explain"].get("model") or "llm"
                out["explain"] = parsed["explain"]
            if isinstance(parsed.get("engine_report"), dict):
                er = parsed["engine_report"]
                er.setdefault("repro", {})["template"] = f"用户上传模板（tpl-hash: {tpl_hash}…）"
                er.setdefault("repro", {})["source"] = f"用户文档（doc-hash: {doc_hash}…）"
                out["engine_report"] = er
            return "explain" in out or "engine_report" in out
    except Exception:
        return False
    return False


if __name__ == "__main__":
    sys.exit(main())
