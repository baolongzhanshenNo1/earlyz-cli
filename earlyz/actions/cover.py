"""封面构建（A3）——M3，2026-10-08。

机制（对齐 docs/方案-格式检测报告接入与A档七项-2026-10-07.md §A3，已定形）：
- **块级克隆而非重画**：模板封面区段落整体深度拷贝进目标文档——字体随 run
  的 rFonts 携带、logo 图片原字节拷贝（媒体部件注入 + rels/Content_Types 登记），
  不存在「识别再重排」环节；
- **槽位回填**：「标签：值」段的值区原位替换（run 级手术，复用 cross_refs 的
  `_replace_number_span` 坐标系）；下划线值区按字体度量补空格凑回原宽；题目区
  多行 → 值写入首行、其余行清空（保留段落稳定版面）；
- **示例区解析**：模板的「XX示例：」标记段与字体注记段（楷体「（…号…）」）是
  模板元信息，不属交付内容 → 克隆时剔除并如实列出；模板沉默处不发明；
- **替换用户原封面**：目标前部按「承载 sectPr 的段」切块；封面块=首个识别块；
  内容段删除、承载 sectPr 的段落保留（提供分页与节属性）；replace=auto/none/显式；
- **保真断言**：①XML 级——克隆段 vs 模板原段（r:embed 归一化后）逐字节一致，
  不过则回退报错；②资产级——目标媒体逐字节 == 模板媒体；③渲染级——真机渲染
  复核（输出报告供验收，不做阻断）。
- 边界：只做模板已有内容；槽位值缺省留空并如实列出；学号数字形态校验；模板
  缺字体只通报不分发。

入口：`build_cover_from_template(path, params)` / `verify_cover_built(path, params)`
"""
from __future__ import annotations

import copy
import re
import zipfile
from pathlib import Path

from lxml import etree

from earlyz.actions.new_actions import W, _read_part, _part_names
from earlyz.actions.cross_refs import _drtext, _replace_number_span

R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
R_IMG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
# 包级不变量（§7.1 P0）：wp:docPr id 必须包级唯一——克隆体必须换号防撞
_WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"

_ZONES = ("cover", "statement")
_ZONE_ALIAS = {"封面": "cover", "学术声明": "statement", "声明": "statement"}
_ANNOT_KEYWORDS = re.compile(r"磅|宋体|黑体|楷体|仿宋|居中|加粗|行距|Times|New Roman|pt")
_SLOT_RE = re.compile(r"^\s*([^：:]{2,16})[：:]\s*(.*)$")
_TITLE_SZ = "44"  # 2号=22pt=sz44（模板注记「2号黑体加粗居中」实测）
_DATE_RE = re.compile(r"^[\s　]*(二[○〇零]|20\d\d)[\s　]*[^\n]{0,10}[年月]")
_PH = "\ue020CVR\ue021"
_XMLSPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _para_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(f"{W}t"))


def _has_sectpr(p) -> bool:
    return p.find(f"{W}pPr/{W}sectPr") is not None


def _read_bytes(path: Path, name: str) -> bytes:
    with zipfile.ZipFile(path, "r") as z:
        return z.read(name)


# ── 模板解析 ────────────────────────────────────────────────────

def _is_annotation(text: str, p) -> bool:
    t = (text or "").strip()
    if not t:
        return False
    if re.search(r"示例\s*[:：]\s*$", t):
        return True
    if re.fullmatch(r"[（(][^）)]{1,40}[）)]", t) and _ANNOT_KEYWORDS.search(t):
        return True
    fonts = {f.get(f"{W}eastAsia") for f in p.iter(f"{W}rFonts")
             if f.get(f"{W}eastAsia")}
    if fonts and fonts <= {"楷体_GB2312", "楷体"} and len(t) < 60:
        return True
    return False


def _zone_paras(body_kids):
    """按「XX示例：」标记切区 → [(marker_idx, [content_idx...])]。"""
    markers = []
    for i, el in enumerate(body_kids):
        if etree.QName(el).localname != "p":
            continue
        if re.search(r"示例\s*[:：]\s*$", _para_text(el).strip()):
            markers.append(i)
    zones = []
    for mi, m in enumerate(markers):
        end = markers[mi + 1] - 1 if mi + 1 < len(markers) else len(body_kids) - 1
        zones.append((m, list(range(m + 1, end + 1))))
    return zones


def _zone_name(marker_text: str) -> str:
    m = re.match(r"([\u4e00-\u9fa5]{1,6})示例", marker_text.strip())
    return m.group(1) if m else ""


def _zone_filter(body_kids, idxs):
    """区内容段筛选：剔除注记段与分节控制段。"""
    kept, dropped = [], []
    for i in idxs:
        el = body_kids[i]
        if etree.QName(el).localname != "p":
            continue
        t = _para_text(el)
        if _is_annotation(t, el):
            dropped.append({"index": i, "text": t.strip()[:60], "reason": "模板注记（非交付内容）"})
            continue
        if _has_sectpr(el):
            dropped.append({"index": i, "text": t.strip()[:40], "reason": "分节控制段（结构）"})
            continue
        kept.append(i)
    return kept, dropped


def _slot_of(p):
    t = _para_text(p)
    m = _SLOT_RE.match(t)
    if not m:
        return None
    label, value = m.group(1).strip(), m.group(2)
    if not re.search(r"[\u4e00-\u9fa5]", label) or _ANNOT_KEYWORDS.search(label):
        return None
    return {"label": label, "label_ns": re.sub(r"\s+", "", label), "value": value}


def _date_of(p):
    t = _para_text(p)
    if _DATE_RE.match(t) and len(t) < 24:
        return {"label": "日期", "label_ns": "日期", "value": t.strip()}
    return None


def _norm_slot_label(s: object) -> str:
    return re.sub(r"[\s\u3000]+", "", str(s or "")).lower()


def _match_slot_values(provided: dict, labels: list[str]) -> tuple[dict[str, str], list[dict]]:
    """把「用户/LLM 给的槽位键」映射到模板标签（2026-10-17 封面回填事故）。

    真机：模板标签是「专 业 名 称 / 学 生 姓 名 / 学 生 学 号」，而提示/LLM 给的是
    「专业 / 姓名 / 学号」→ 精确匹配全部失败 → 重建出的封面留 XXX 占位（半成品，
    且只在 notes 里写「未提供值」，用户看到的是一张没填信息的封面）。
    规则（确定性，无 LLM）：
      1) 完全同名（去空白、忽略大小写）优先；
      2) 双向包含匹配：提供的键是标签子串（或反之），取长度差最小的一个；
      3) 每个标签、每个提供的键最多用一次；未匹配的键原样返回（调用方响亮报告）。
    """
    norm_labels = {lab: _norm_slot_label(lab) for lab in labels}
    matched: dict[str, str] = {}
    used: set[str] = set()
    for k, v in provided.items():
        if not v or k in used:
            continue
        for lab, nl in norm_labels.items():
            if nl == k and lab not in matched:
                matched[lab] = str(v)
                used.add(k)
                break
    for lab, nl in norm_labels.items():
        if lab in matched:
            continue
        best: tuple[int, str] | None = None
        for k, v in provided.items():
            if not v or k in used:
                continue
            if k in nl or nl in k:
                gap = abs(len(k) - len(nl))
                if best is None or gap < best[0]:
                    best = (gap, k)
        if best is not None:
            matched[lab] = str(provided[best[1]])
            used.add(best[1])
    leftovers = [{"label": k, "value": str(v)[:30]}
                 for k, v in provided.items() if v and k not in used]
    return matched, leftovers


def _title_lines(zone_kept, body_kids):
    out = []
    for i in zone_kept:
        p = body_kids[i]
        if not _para_text(p).strip():
            continue
        for r in p.findall(f"{W}r"):
            rpr = r.find(f"{W}rPr")
            if rpr is None:
                continue
            sz = rpr.find(f"{W}sz")
            b = rpr.find(f"{W}b")
            if (sz is not None and (sz.get(f"{W}val") or "") == _TITLE_SZ
                    and b is not None and (b.get(f"{W}val") or "1") not in ("0", "false")):
                out.append(i)
                break
    return out


def _assets_of(paras, tpl_zip):
    rels = etree.fromstring(tpl_zip.read("word/_rels/document.xml.rels"))
    rid_map = {x.get("Id"): x.get("Target") for x in rels}
    out = []
    for p in paras:
        for el in p.iter():
            tag = etree.QName(el).localname
            rid = el.get(f"{R}embed") if tag == "blip" else (
                el.get(f"{R}id") if tag == "imagedata" else None)
            if not rid:
                continue
            target = rid_map.get(rid)
            if target:
                out.append("word/" + target.lstrip("/"))
    return sorted(set(out))


def _analyze_template(tpl_path: Path):
    tpl_root = _read_part(tpl_path, "word/document.xml")
    tpl_kids = list(tpl_root.find(f"{W}body"))
    zones = []
    with zipfile.ZipFile(tpl_path) as z:
        for mi, idxs in _zone_paras(tpl_kids):
            mtext = _para_text(tpl_kids[mi]).strip()
            name = _ZONE_ALIAS.get(_zone_name(mtext))
            if name is None:
                continue
            kept, dropped = _zone_filter(tpl_kids, idxs)
            if not kept:
                continue
            slots = []
            for i in kept:
                s = _slot_of(tpl_kids[i]) or _date_of(tpl_kids[i])
                if s:
                    slots.append({"index": i, **s})
            # 题目行只在封面/扉页语义区生效：声明区首段（如「郑 重 声 明」）是固定
            # 文案，绝不能被论文题目替换（2026-10-08 双区实测缺陷）
            titles = (_title_lines(kept, tpl_kids)
                      if name in ("cover", "title_page") else [])
            paras = [tpl_kids[i] for i in kept]
            fonts = set()
            for p in paras:
                for f in p.iter(f"{W}rFonts"):
                    for k in ("eastAsia", "ascii", "hAnsi"):
                        v = f.get(f"{W}{k}")
                        if v:
                            fonts.add(v)
            zones.append({
                "name": name, "marker_index": mi, "kept": kept, "dropped": dropped,
                "slots": slots, "title_lines": titles,
                "assets": _assets_of(paras, z),
                "fonts": sorted(fonts),
                "preview": [_para_text(tpl_kids[i]).strip()[:42]
                            for i in kept if _para_text(tpl_kids[i]).strip()][:12],
            })
    return tpl_root, tpl_kids, zones


def _installed_fonts():
    """已装字体名集合。2026-11-21 改为共用 earlyz.engine.fonts：
    旧实现只读注册表，而中文家族名（宋体/黑体/微软雅黑/楷体/仿宋）并不出现在
    注册表键名里 → 模板里每个中文字体都被误报「缺失」。返回 None = 判断不了。
    """
    from earlyz.engine.fonts import installed_font_names

    return installed_font_names()


# ── 目标前部块 ─────────────────────────────────────────────────

def _resolve_template_path(tpl_s: str, target: Path) -> Path:
    """模板路径解析（2026-10-07；2026-10-09 扩搜索域：数据根兄弟目录）。

    绝对路径原样返回；裸文件名（如 LLM 汇报的上传名 <uid>_模板.docx）解析域 =
    目标同目录 + 数据根同级目录（output/ ↔ uploads/）：先各目录精确匹配，
    再按「_+文件名」后缀匹配取最新（<uid>_ 前缀命名）；都找不到返回原样，
    由调用方照常报「模板文件不存在」。与下载/预览指针解析口径一致。
    2026-10-09 事故：目标=output/ 的工作副本，模板=uploads/（本会话刚上传），
    旧实现只查 target.parent 的两级 → 封面重建两次「模板文件不存在」。
    2026-10-12 事故：LLM 把「我的模板」包 id 形态（user:<sha12>）的残段
    「user:」当 template_path 传进来 → 直接判「模板文件不存在」、封面没重建。
    现补：user:<sha12> 形态直接查 packs_user；仍找不到才按原样返回给调用方报错。
    """
    if isinstance(tpl_s, str) and tpl_s.startswith("user:"):
        try:
            from earlyz.spec.packs import find_user_pack_file
            _f = find_user_pack_file(tpl_s)
            if _f is not None and Path(_f).is_file():
                return Path(_f)
        except Exception:  # noqa: BLE001 — 解析失败落到下面的常规路径
            pass
    p = Path(tpl_s)
    if p.is_absolute() or p.exists():
        return p
    dirs: list[Path] = []
    parent = target.parent
    for cand in (parent, parent.parent / "uploads", parent.parent / "output"):
        try:
            is_dir = cand.is_dir()
        except OSError:
            is_dir = False
        if is_dir and cand not in dirs:
            dirs.append(cand)
    for d in dirs:
        exact = d / p.name
        if exact.is_file():
            return exact
    cands: list[Path] = []
    for d in dirs:
        try:
            cands += [q for q in d.iterdir()
                      if q.is_file() and q.suffix.lower() == ".docx"
                      and q.name.endswith("_" + p.name)]
        except OSError:
            continue
    if cands:
        return max(cands, key=lambda q: q.stat().st_mtime)
    return p


def _target_blocks(path: Path):
    root = _read_part(path, "word/document.xml")
    body = root.find(f"{W}body")
    if body is None:
        return None, []
    kids = list(body)
    blocks, start = [], 0
    for i, el in enumerate(kids):
        tag = etree.QName(el).localname
        if (tag == "p" and _has_sectpr(el)) or tag == "sectPr":
            blocks.append((start, i))
            start = i + 1
    if start < len(kids):
        blocks.append((start, len(kids) - 1))
    out = []
    for (a, b) in blocks:
        texts = []
        for el in kids[a:b + 1]:
            if etree.QName(el).localname != "p":
                continue
            t = _para_text(el).strip()
            if t:
                texts.append(t[:48])
        out.append({"range": [a, b], "texts": texts[:14]})
    return root, out


def _classify_block(block: dict) -> str:
    # 中文标题常写成「郑 重 声 明」「毕 业 设 计（论文）」——先压掉空白再匹配，
    # 否则声明块会漏判，导致只替换其中一段、另一段原样留下（实测重复）。
    joined = "".join("".join(block["texts"]).split())
    if re.search(r"郑重声明|原创性声明|版权使用授权书|学位论文.{0,4}声明", joined):
        return "statement"
    if re.search(r"毕业设计（论文）|学士学位论文|毕业设计\(论文\)", joined):
        return "cover"
    return ""


def _find_target_map(blocks):
    out = {}
    for name in _ZONES:
        hits = [b["range"] for b in blocks if _classify_block(b) == name]
        if hits:
            # 同类区可能被分节切成多块（封面页 / 声明页各有一块）——取首尾合并区间，
            # 中间的分节载体由删除逻辑保留，不会破坏页边界。
            out[name] = [hits[0][0], hits[-1][1]]
    return out


def _cover_span_fallback(root, blocks):
    """封面区兜底定位（2026-10-09）：文档无分节（整篇为一块）且声明标题把
    首块分类吸收为 statement 时，以首个声明类标题为界——其前即封面区。
    实测案例「本科论文测试文档」：无分节、整篇 [0,124] 一块且被「原创性
    声明」吸收 → 封面应为 [0,8]（声明标题前一版段）。
    """
    body = root.find(f"{W}body")
    if body is None:
        return None
    kids = list(body)
    start = blocks[0]["range"][0] if blocks else 0
    for i in range(start, len(kids)):
        el = kids[i]
        if etree.QName(el).localname != "p":
            continue
        t = _para_text(el).strip()
        if not t:
            continue
        if re.match(r"^(原创性声明|郑重声明|学位论文声明|版权使用授权书|学术声明)", t):
            if 4 <= i - start <= 40:
                return [start, i - 1]
            return None
    return None


# ── 克隆 → 回填 ─────────────────────────────────────────────────

def _slot_span(p):
    """槽位值区（直接 run 坐标）[start, end) = 冒号后一位到段尾。

    注意：起点取 m.end(1)+1（冒号本身保留；空白值区如「：␣␣␣」也可替换）。
    """
    text = _drtext(p)
    m = _SLOT_RE.match(text)
    if not m:
        return None, text
    return (m.end(1) + 1, len(text)), text


def _font_em(ch: str) -> float:
    o = ord(ch)
    if (0x4E00 <= o <= 0x9FFF or 0x3000 <= o <= 0x303F or 0xFF00 <= o <= 0xFFEF
            or 0x3400 <= o <= 0x4DBF):
        return 1.0
    return 0.5


def _value_region_em(p) -> float:
    """值区宽度估算（em）——仅当值区含下划线时用于补宽。"""
    span, text = _slot_span(p)
    if span is None:
        return 0.0
    for r in [c for c in p if etree.QName(c).localname == "r"]:
        rpr = r.find(f"{W}rPr")
        if rpr is not None:
            u = rpr.find(f"{W}u")
            if u is not None and (u.get(f"{W}val") or "") not in ("", "none"):
                return sum(_font_em(c) for c in text[span[0]:span[1]])
    return 0.0


def _replace_whole_text(p, value: str) -> None:
    """整段文本替换：值写入首个「含文字」的 run，其余文字 run 清空（保格式结构）。

    run 选择（2026-10-07 修复「日期 18pt→14pt」事故）：模板日期段形如
    [空格 run sz28][文字 run sz36][文字 run sz36]……——值必须落在「原来是文字」的
    run 上才能带走正确字号字体；纯空白 run（前导空格/缩进占位）原样保留。
    全段皆无文字时退回首 run（旧行为）。
    """
    pairs = [t for r in p if etree.QName(r).localname == "r"
             for t in r.iter(f"{W}t")]
    if not pairs:
        return
    carrier = next((t for t in pairs if (t.text or "").strip()), pairs[0])
    for t in pairs:
        if t is carrier:
            t.text = value
            t.set(_XMLSPACE, "preserve")
        elif (t.text or "").strip():
            t.text = ""


def _clear_text(p) -> None:
    for t in p.iter(f"{W}t"):
        t.text = ""


def _fill_value(p, value: str, pad_to_em: float | None = None):
    """值区原位替换（run 级手术 +（可选）补宽）。"""
    span, _text = _slot_span(p)
    if span is None:
        return False, "未找到值区（无冒号）"
    a, b = span
    if pad_to_em and pad_to_em > 0:
        deficit = pad_to_em - sum(_font_em(c) for c in value)
        if deficit > 0.49:
            value = value + " " * int(round(deficit / 0.5))
    err = _replace_number_span(p, a, b, _PH)
    if err:
        return False, err
    for r in [c for c in p if etree.QName(c).localname == "r"]:
        for t in r.iter(f"{W}t"):
            if _PH in (t.text or ""):
                t.text = t.text.replace(_PH, value)
                t.set(_XMLSPACE, "preserve")
                return True, None
    return False, "占位符丢失"


# ── 媒体重映射 ─────────────────────────────────────────────────

def _remap_assets(clones, tpl_zip, tgt_parts: set, rels_root, new_entries: dict,
                  copied: list, existing_media_bytes: dict | None = None):
    rels = etree.fromstring(tpl_zip.read("word/_rels/document.xml.rels"))
    rid_target = {x.get("Id"): x.get("Target") for x in rels}
    max_rid = 0
    for x in rels_root:
        rid = x.get("Id") or ""
        if rid.startswith("rId"):
            try:
                max_rid = max(max_rid, int(rid[3:]))
            except ValueError:
                pass
    used = {n for n in tgt_parts if n.startswith("word/media/")}
    existing_media_bytes = existing_media_bytes or {}
    exts = set()
    for cl in clones:
        for el in cl.iter():
            tag = etree.QName(el).localname
            attr = f"{R}embed" if tag == "blip" else (
                f"{R}id" if tag == "imagedata" else None)
            if attr is None:
                continue
            rid = el.get(attr)
            target = rid_target.get(rid) if rid else None
            if not target:
                continue
            tpl_part = "word/" + target.lstrip("/")
            # 同段同图去重（同 tpl_part 且已搬过 → 复用已建关系）
            reused = None
            for c in copied:
                if c["from"] == tpl_part and c.get("zone_drawing") is not None:
                    reused = c.get("new_rid")
                    break
            data = tpl_zip.read(tpl_part)
            ext = tpl_part.rsplit(".", 1)[-1].lower()
            if reused:
                el.set(attr, reused)
                continue
            # 字节复用：目标包内已有同字节 cover* 部件 → 直接复用（幂等重建不膨胀）
            reuse_part = None
            for c in copied:
                if c["from"] == tpl_part:
                    reuse_part = c["part"]
                    break
            if reuse_part is None:
                for n_ in range(1, 60):
                    cand = f"word/media/cover{n_}.{ext}"
                    if cand in existing_media_bytes and existing_media_bytes[cand] == data:
                        reuse_part = cand
                        break
            if reuse_part is not None:
                max_rid += 1
                new_rid = f"rId{max_rid}"
                rels_root.append(rels_root.makeelement("Relationship", {
                    "Id": new_rid, "Type": R_IMG,
                    "Target": reuse_part.replace("word/", ""), "TargetMode": "Internal"}))
                el.set(attr, new_rid)
                copied.append({"part": reuse_part, "from": tpl_part, "bytes": len(data),
                               "reused": True, "new_rid": new_rid})
                exts.add(ext)
                continue
            n = 1
            while f"word/media/cover{n}.{ext}" in used:
                n += 1
            name = f"word/media/cover{n}.{ext}"
            used.add(name)
            new_entries[name] = data
            max_rid += 1
            new_rid = f"rId{max_rid}"
            rels_root.append(rels_root.makeelement("Relationship", {
                "Id": new_rid, "Type": R_IMG,
                "Target": name.replace("word/", ""), "TargetMode": "Internal"}))
            el.set(attr, new_rid)
            exts.add(ext)
            copied.append({"part": name, "from": tpl_part, "bytes": len(data),
                           "zone_drawing": True, "new_rid": new_rid})
    return exts


# ── 保真断言 ① / 结构归一 ────────────────────────────────────────

def _norm_xml(el) -> bytes:
    clone = copy.deepcopy(el)
    for x in clone.iter():
        for k in list(x.attrib):
            if k in (f"{R}embed", f"{R}id", f"{R}link"):
                x.set(k, "RID")
    # docPr id 换号（克隆体必须重编防撞，§7.1）后仍应视为保真——与关系 id 同理屏蔽
    for dp in clone.iter(f"{_WP}docPr"):
        if dp.get("id") is not None:
            dp.set("id", "DOCPR")
    return etree.tostring(clone)


def _max_docpr_id(root) -> int:
    """目标文档现有 wp:docPr 的最大数值 id（非数值忽略）——克隆体从这里递增。"""
    mx = 0
    for dp in root.iter(f"{_WP}docPr"):
        try:
            mx = max(mx, int(dp.get("id") or 0))
        except (TypeError, ValueError):
            continue
    return mx


def _renumber_docpr_ids(els, counter: list) -> int:
    """克隆体里全部 wp:docPr 重编 id（counter=[目标现有最大]），防与目标撞号。

    2026-10-12 事故根因：封面克隆原样携带模板图片的 docPr id，与目标文档
    已有图片撞号 ⇒ lint 报 package.docpr_duplicate ⇒ planner 崩「编译异常: 'property'」。
    """
    n = 0
    for el in els:
        for dp in el.iter(f"{_WP}docPr"):
            counter[0] += 1
            dp.set("id", str(counter[0]))
            n += 1
    return n


def _assert_fidelity(clones, tpl_kids, zone, slot_idx, title_idx):
    skip = set(slot_idx) | set(title_idx)
    for pos, i in enumerate(zone["kept"]):
        if i in skip:
            continue
        if _norm_xml(clones[pos]) != _norm_xml(tpl_kids[i]):
            return False, f"段 {i} 克隆失真"
    return True, None


# ── 槽位值校验 ─────────────────────────────────────────────────

def _validate_slot(label_ns: str, value: str):
    if not value:
        return None
    if "学号" in label_ns:
        if re.search(r"[\u4e00-\u9fa5]", value):
            return f"学号「{value}」含中文，需为数字形态"
        if len(re.sub(r"\D", "", value)) < 4:
            return f"学号「{value}」数字位数不足"
    return None


# ── 写包 ───────────────────────────────────────────────────────

def _write_package(path: Path, updates: dict, new_entries: dict):
    tmp = path.with_suffix(".cover.tmp.docx")
    with zipfile.ZipFile(path, "r") as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = updates.pop(item.filename, None)
            if data is None:
                data = zin.read(item.filename)
            zout.writestr(item, data)
        for name, data in {**updates, **new_entries}.items():
            zi = zipfile.ZipInfo(name)
            zi.compress_type = (zipfile.ZIP_STORED if name.startswith("word/media/")
                                else zipfile.ZIP_DEFLATED)
            zout.writestr(zi, data)
    tmp.replace(path)


def _ensure_ct_default(ct_root, ext: str) -> bool:
    for d in ct_root.findall(f"{CT}Default"):
        if (d.get("Extension") or "").lower() == ext.lower():
            return False
    default_map = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
                   "gif": "image/gif", "bmp": "image/bmp",
                   "emf": "image/x-emf", "wmf": "image/x-wmf", "tiff": "image/tiff"}
    ct_root.append(ct_root.makeelement(f"{CT}Default", {
        "Extension": ext,
        "ContentType": default_map.get(ext.lower(), "application/octet-stream")}))
    return True


def _ensure_pbb(p) -> None:
    """给段加 pageBreakBefore（schema 序：pStyle→keepNext→keepLines 之后插入）。"""
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        ppr = etree.Element(f"{W}pPr")
        p.insert(0, ppr)
    if ppr.find(f"{W}pageBreakBefore") is not None:
        return
    el = etree.Element(f"{W}pageBreakBefore")
    idx = 0
    for j, ch in enumerate(ppr):
        if etree.QName(ch).localname in ("pStyle", "keepNext", "keepLines"):
            idx = j + 1
    ppr.insert(idx, el)


# ── 主入口 ─────────────────────────────────────────────────────

def _norm_kinds(value):
    if not value:
        return ["cover"]
    v = [value] if isinstance(value, str) else list(value)
    out = [k for k in _ZONES if k in v]
    return out or ["cover"]


def build_cover_from_template(path, params: dict | None = None) -> dict:
    """用模板封面页集合构建目标文档封面（块级克隆 + 槽位回填）。

    params:
      template_path: str 必填——规范模板 .docx 绝对路径
      apply: bool 必填——false 只出分析/计划预览（不落盘）
      zones: ["cover"] | ["cover","statement"]（默认 ["cover"]）
      slots: {"专业名称": "…", "学生姓名": "…", "学生学号": "…", "指导教师": "…",
              "日期": "…", "题目": "…", "英文题目": "…（可选，写入标题区第二行）"}
              —— 键=模板标签（去空格）或「题目」「日期」「英文题目」
      replace: "auto"（默认）/ "none" / {"cover": [start, end], …}
    """
    path = Path(path)
    params = params or {}
    tpl_s = params.get("template_path")
    if not tpl_s:
        return {"changed": False, "error": "template_path 必填（规范模板 .docx 路径）"}
    tpl_path = _resolve_template_path(str(tpl_s), path)
    if not tpl_path.exists() or tpl_path.suffix.lower() != ".docx":
        return {"changed": False, "error": f"模板文件不存在或不是 .docx：{tpl_s}"}
    apply_ = bool(params.get("apply"))
    sel_names = _norm_kinds(params.get("zones"))
    # 槽位键归一化（2026-10-07）：模板标签形如「专 业 名 称」，agent 若原样
    # 复制（带空格）会全部匹配不上且静默留空——真机实测事故。统一去空白。
    slots_in = {re.sub(r"[\s\u3000]+", "", str(k)): v
                for k, v in (params.get("slots") or {}).items()}
    # 日期零字符归一（2026-10-07）：符号 ○(U+25CB) → 汉字 〇(U+3007)。○ 走
    # 西文字体取形（真机渲染实测：TNR 圆环、高 19.9 vs 汉字 18.0、基线偏 0.6pt，
    # 即「文字不是同一字体」），且日期规范写法本就是〇。仅作用于日期槽值。
    _dv = slots_in.get("日期")
    if isinstance(_dv, str) and "\u25cb" in _dv:
        slots_in["日期"] = _dv.replace("\u25cb", "\u3007")
    replace = params.get("replace", "auto")
    # force=True：跳过「已是当前模板克隆」的幂等判定，强制按 replace 区段重建。
    # 2026-10-22 实测：文档里混入过一份旧克隆时，_already_clone 会认成"已建过"
    # 而整体跳过，源封面原样留着 → 用户看到「封面没改」。封面属于产品托管区，
    # 允许强制重建。
    force = bool(params.get("force"))

    _tpl_root, tpl_kids, zones = _analyze_template(tpl_path)
    sel = [z for z in zones if z["name"] in sel_names]
    if not sel:
        return {"changed": False,
                "error": "模板内未识别到所请求的示例区（封面/声明）",
                "zones_found": [z["name"] for z in zones]}

    # 槽位键 → 模板标签的语义匹配（2026-10-17）：「专业」→「专业名称」、
    # 「姓名」→「学生姓名」、「学号」→「学生学号」等。
    _labels = [s["label_ns"] for z in sel for s in z["slots"]]
    _matched, _unmatched = _match_slot_values(slots_in, _labels)
    for _k in ("题目", "英文题目"):
        if slots_in.get(_k):
            _matched.setdefault(_k, str(slots_in[_k]))

    for z in sel:
        for s in z["slots"]:
            err = _validate_slot(s["label_ns"], _matched.get(s["label_ns"], ""))
            if err:
                return {"changed": False, "error": err,
                        "slots": [x["label_ns"] for z2 in sel for x in z2["slots"]]}

    fonts = set()
    for z in sel:
        fonts.update(z["fonts"])
    installed = _installed_fonts()
    missing_fonts = []
    if installed is not None:
        missing_fonts = sorted(f for f in fonts if f not in installed)

    tgt_root, blocks = _target_blocks(path)
    if tgt_root is None:
        return {"changed": False, "error": "目标文档没有 body"}
    suggested = _find_target_map(blocks)
    if "cover" not in suggested:
        # 无分节整块文档兜底（2026-10-09）：整篇为一块且含声明时，块分类被
        # 「声明」吸收、封面自动定位失败——以声明类标题为界切出首部封面区。
        _fb = _cover_span_fallback(tgt_root, blocks)
        if _fb:
            suggested["cover"] = _fb

    # 题目兜底（2026-10-17）：LLM 没传「题目」时，沿用目标文档封面里的大标题行——
    # 否则克隆出来的封面会顶着模板的示例标题（「XXX管理系统…」），比不重建更糟。
    if not _matched.get("题目") and suggested.get("cover"):
        try:
            _kids_all = list(tgt_root.find(f"{W}body"))
            _a, _b = int(suggested["cover"][0]), int(suggested["cover"][1])
            for _i in _title_lines(list(range(_a, _b + 1)), _kids_all):
                _t = _para_text(_kids_all[_i]).strip()
                if _t:
                    _matched["题目"] = _t
                    break
            if not _matched.get("题目"):
                # 其次：目标封面里的「题目：xxx」槽（很多论文封面是这种写法）
                for _i in range(_a, _b + 1):
                    _el = _kids_all[_i]
                    if etree.QName(_el).localname != "p":
                        continue
                    _m = _SLOT_RE.match(_para_text(_el).strip())
                    if _m and re.sub(r"[\s\u3000]+", "", _m.group(1)) == "题目":
                        if _m.group(2).strip():
                            _matched["题目"] = _m.group(2).strip()
                        break
            # 槽位兜底（2026-10-17）：模板槽没拿到值时沿用目标封面里的同义值——
            # 否则交付出来的封面顶着模板占位符（「XXX」「二○二四年五月」）。
            _tgt_slots: dict[str, str] = {}
            for _i in range(_a, _b + 1):
                _el = _kids_all[_i]
                if etree.QName(_el).localname != "p":
                    continue
                _s = _slot_of(_el) or _date_of(_el)
                if _s and _s["value"].strip() and _s["label_ns"] not in _tgt_slots:
                    _tgt_slots[_s["label_ns"]] = _s["value"].strip()
            if _tgt_slots:
                _m2, _ = _match_slot_values(_tgt_slots, _labels)
                for _lab, _val in _m2.items():
                    if _val and not _matched.get(_lab):
                        _matched[_lab] = _val
        except Exception:  # noqa: BLE001 — 兜底失败就按「题目未提供」处理
            pass

    if not apply_:
        plan = []
        for z in sel:
            rng = None
            if replace == "auto":
                rng = suggested.get(z["name"])
            elif isinstance(replace, dict):
                rng = replace.get(z["name"])
            plan.append({
                "zone": z["name"],
                "template_paragraphs": len(z["kept"]),
                "dropped": len(z["dropped"]),
                "slots": [{"label": s["label"],
                           "current": (_matched.get(s["label_ns"])
                                       or s["value"]).strip()[:30]}
                          for s in z["slots"]],
                "title_lines": len(z["title_lines"]),
                "assets": z["assets"],
                "replace_range": rng,
                "preview": z["preview"],
            })
        return {
            "changed": False, "applied": False,
            "zones": plan,
            "zones_found": [
                {"name": z["name"], "paragraphs": len(z["kept"]),
                 "slots": [s["label"] for s in z["slots"]]}
                for z in zones],
            "target_blocks": blocks[:6],
            "suggested_map": suggested,
            "missing_fonts": missing_fonts,
            "message": ("封面构建预览（未落盘）：确认 zones/槽位值与 replace 后 "
                        "apply=true 落盘。"
                        + (f" 模板字体缺失（只通报）：{'、'.join(missing_fonts)}。"
                           if missing_fonts else "")),
        }

    with zipfile.ZipFile(tpl_path) as tpl_zip:
        rels_root = etree.fromstring(_read_bytes(path, "word/_rels/document.xml.rels"))
        ct_root = etree.fromstring(_read_bytes(path, "[Content_Types].xml"))
        tgt_parts = set(_part_names(path))
        new_entries: dict = {}
        copied: list = []
        cloned_regions = []
        replaced_ranges = []
        removed_total = 0
        slots_filled: dict = {}
        slots_left: list = []
        zones_skipped: list = []
        # 未匹配的槽位键响亮报告（不静默丢值，2026-10-07）
        for _u in _unmatched:
            if _u["label"] in ("题目", "英文题目"):
                continue
            slots_left.append({"label": _u["label"],
                               "note": "槽位键不匹配模板标签（未写入）"})
        existing_media_bytes: dict = {}
        for n_ in _part_names(path):
            if n_.startswith("word/media/cover"):
                existing_media_bytes[n_] = _read_bytes(path, n_)

        # 各区的克隆/回填/校验先行（不动目标）
        prepared = []
        docpr_ctr = [_max_docpr_id(tgt_root)]
        for z in sel:
            clones = [copy.deepcopy(tpl_kids[i]) for i in z["kept"]]
            if not force and _already_clone(tgt_root, clones, z, _matched):
                zones_skipped.append(z["name"])
                continue
            # 包级不变量（§7.1 P0）：克隆体 docPr 换号防撞（在已构建判定之后、
            # 保真断言之前——断言侧已按 DOCPR 归一，见 _norm_xml/_sig）
            _renumber_docpr_ids(clones, docpr_ctr)
            slot_map = {s["index"]: s for s in z["slots"]}
            first_title = z["title_lines"][0] if z["title_lines"] else None
            for pos, i in enumerate(z["kept"]):
                p = clones[pos]
                if i in slot_map:
                    s = slot_map[i]
                    val = _matched.get(s["label_ns"], "")
                    if val:
                        if _slot_span(p)[0] is None:
                            # 无冒号槽（日期行等）→ 整段替换
                            _replace_whole_text(p, val)
                            slots_filled[s["label"]] = val
                        else:
                            pad = _value_region_em(p)
                            ok, err = _fill_value(p, val, pad_to_em=pad or None)
                            if ok:
                                slots_filled[s["label"]] = val
                            else:
                                slots_left.append({"label": s["label"], "error": err})
                    else:
                        slots_left.append({"label": s["label"], "note": "未提供值（留空）"})
                elif i in z["title_lines"]:
                    title = _matched.get("题目", "")
                    en = str(_matched.get("英文题目") or "").strip()
                    tls = z["title_lines"]
                    if title and i == first_title:
                        _replace_whole_text(p, title)
                        slots_filled["题目"] = title
                    elif title and en and len(tls) > 1 and i == tls[1]:
                        # 英文题目（可选槽，2026-10-07）：写入标题区第二行，
                        # 格式随模板标题行；不提供则不占位（保持整行清空）。
                        _replace_whole_text(p, en)
                        slots_filled["英文题目"] = en
                    elif title:
                        _clear_text(p)
                    else:
                        slots_left.append({"label": "题目",
                                           "note": "未提供（标题行保留模板示例文字）"})
            exts = _remap_assets(clones, tpl_zip, tgt_parts, rels_root,
                                 new_entries, copied, existing_media_bytes)
            for e in exts:
                _ensure_ct_default(ct_root, e)
            ok, why = _assert_fidelity(clones, tpl_kids, z,
                                       [s["index"] for s in z["slots"]], z["title_lines"])
            if not ok:
                return {"changed": False,
                        "error": f"保真校验未过（{z['name']}）：{why}——已回退，未落盘"}
            prepared.append((z, clones))

        # 落盘：替换目标块——先把全部区范围解引用为**元素引用快照**
        # （多区连续替换时，前一个区的增删会让后续区的数字索引失效 → 必须
        #   在任何结构编辑之前把 [a,b] 绑定到元素；2026-10-08 双区实测缺陷）
        body_pre = tgt_root.find(f"{W}body")
        kids_snapshot = list(body_pre)
        for z, clones in prepared:
            rng = None
            if replace == "auto":
                rng = suggested.get(z["name"])
            elif isinstance(replace, dict):
                rng = replace.get(z["name"])
            if rng is not None:
                a, b = int(rng[0]), int(rng[1])
                a = max(0, min(a, len(kids_snapshot) - 1))
                b = max(a, min(b, len(kids_snapshot) - 1))
                region = kids_snapshot[a:b + 1]
                tail = region[-1].getnext() if region else None
                keep_sect = None
                for el in reversed(region):
                    ln = etree.QName(el).localname
                    # 分节载体（段内 sectPr / body 级 sectPr）一律不删：
                    # body 级 sectPr 是文档级必需件，删掉即非法文档
                    if ln == "sectPr" or (ln == "p" and _has_sectpr(el)):
                        if keep_sect is None:
                            keep_sect = el
                        # 分节载体段的**文字**必须清掉：它是上一节最后一段，
                        # 常携带「二○二四年五月」「本人签名：…」这类行尾文字，
                        # 只删正文段不清它 → 重建后这些行原样留着（实测重复）。
                        if ln == "p":
                            _clear_text(el)
                        continue
                    el.getparent().remove(el)
                    removed_total += 1
                replaced_ranges.append(list(rng))
                # 回插锚：优先保留下来的分节段（页边界随之保留）；若无分节段，
                # 插回原区段位置（区段后首个兄弟之前），并对后随段加段前分页保
                # 持原分页语义。绝不许插到文首——那会让中部的区段错位到封面位。
                anchor = keep_sect if keep_sect is not None else tail
                if keep_sect is None and anchor is not None \
                        and etree.QName(anchor).localname == "p" \
                        and not _has_sectpr(anchor):
                    _ensure_pbb(anchor)
                if anchor is not None:
                    for p in clones:
                        anchor.addprevious(p)
                else:
                    body = tgt_root.find(f"{W}body")
                    for p in clones:
                        body.append(p)
            else:
                # replace="none"：不删任何内容，克隆块插到文首（沿用旧封面之下）
                body = tgt_root.find(f"{W}body")
                first = body[0] if len(body) else None
                if first is None:
                    for p in clones:
                        body.append(p)
                else:
                    for p in clones:
                        first.addprevious(p)
                    _ensure_pbb(first)
            cloned_regions.append({"zone": z["name"], "count": len(clones)})

    if not prepared and zones_skipped:
        return {"changed": False, "applied": False, "reason": "already_built",
                "skipped": zones_skipped,
                "message": "封面已是当前模板克隆（槽位值一致），无需重建。"}

    doc_data = etree.tostring(tgt_root, xml_declaration=True, encoding="UTF-8",
                              standalone=True)
    rels_data = etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8",
                               standalone=True)
    ct_data = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8",
                             standalone=True)
    _write_package(path, {"word/document.xml": doc_data,
                          "word/_rels/document.xml.rels": rels_data,
                          "[Content_Types].xml": ct_data}, new_entries)

    with zipfile.ZipFile(path) as zt, zipfile.ZipFile(tpl_path) as zp:
        for c in copied:
            if c["part"] not in zt.namelist() or zt.read(c["part"]) != zp.read(c["from"]):
                return {"changed": True, "applied": True, "verified": False,
                        "error": f"资产断言未过：{c['part']}"}

    return {
        "changed": True, "applied": True, "verified": True,
        "zones": cloned_regions,
        "zones_skipped": zones_skipped,
        "replaced_ranges": replaced_ranges,
        "removed_paragraphs": removed_total,
        "slots_filled": slots_filled,
        "slots_not_filled": slots_left,
        "assets_copied": copied,
        "missing_fonts": missing_fonts,
        "render_check": "建议真机渲染复核（本函数只做 XML/资产两级断言）",
        "message": ("封面已按模板克隆落盘。"
                    + (f" 模板字体缺失（只通报）：{'、'.join(missing_fonts)}。"
                       if missing_fonts else "")),
    }


# ── 幂等检测 ───────────────────────────────────────────────────

def _already_clone(tgt_root, clones, zone, slots_in) -> bool:
    """目标前部已有与克隆结构一致的段序列且槽位值一致 → 视为已构建。"""
    body = tgt_root.find(f"{W}body")
    if body is None:
        return False
    kids = list(body)[:160]
    want = [_sig(c) for c in clones]
    n = len(want)
    for s in range(0, max(1, len(kids) - n + 1)):
        if not all(_sig(kids[s + t]) == want[t]
                   for t in range(n) if s + t < len(kids)):
            continue
        slot_map = {x["index"]: x for x in zone["slots"]}
        ok = True
        for pos, i in enumerate(zone["kept"]):
            if s + pos >= len(kids):
                ok = False
                break
            txt = _drtext(kids[s + pos])
            if i in slot_map:
                want_v = slots_in.get(slot_map[i]["label_ns"], "")
                if want_v and want_v not in txt:
                    ok = False
                    break
            elif (zone["title_lines"] and i == zone["title_lines"][0]
                    and slots_in.get("题目")):
                if slots_in["题目"] not in txt:
                    ok = False
                    break
            elif (zone["title_lines"] and len(zone["title_lines"]) > 1
                    and i == zone["title_lines"][1]
                    and str(slots_in.get("英文题目") or "").strip()):
                # 已建封面但缺英文行 → 不算已构建，走重建补上（2026-10-07）
                if str(slots_in["英文题目"]).strip() not in txt:
                    ok = False
                    break
        if ok:
            return True
        return False
    return False


# ── verify ─────────────────────────────────────────────────────

_RID_ATTRS = (f"{R}embed", f"{R}id", f"{R}link")
_XMLSPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _sig(el) -> bytes:
    """结构签名：标签+属性（rid/xml:space 归一）——**去文本节点**。

    槽位回填会合法地重分布 run 内文本（甚至清空某行），文本不参与复查；
    结构（runs/rPr/pPr/图片/对齐）必须一致。元素重建保证命名空间净化（模板
    与目标 nsmap 不同，deepcopy 直序列化会携带各自文档的 ns 声明）。
    """
    def build(x):
        ln = etree.QName(x).localname
        if ln == "t":
            return None
        e = etree.Element(x.tag)
        for k in sorted(x.attrib):
            if k in _RID_ATTRS:
                e.set(k, "RID")
            elif x.tag == f"{_WP}docPr" and k == "id":
                # 克隆体 docPr 必须换号（包级唯一，§7.1）——与关系 id 同理，
                # 换号不算结构差异，签名/复查必须归一
                e.set(k, "DOCPR")
            elif k == _XMLSPACE:
                continue
            else:
                e.set(k, x.attrib[k])
        if ln != "t" and x.text:
            e.text = x.text
        for c in x:
            b = build(c)
            if b is not None:
                e.append(b)
        return e
    return etree.tostring(build(el))


def _text_agnostic_xml(el) -> bytes:
    return _sig(el)


def verify_cover_built(path, params: dict | None = None) -> bool:
    """复查：目标前部存在与模板克隆「结构一致」的段序列（文本与 rid 归一）。"""
    try:
        params = params or {}
        if not params.get("apply"):
            return True
        tpl_s = params.get("template_path")
        if not tpl_s:
            return False
        tpl_path = _resolve_template_path(str(tpl_s), Path(path))
        if not tpl_path.exists():
            return False
        _tpl_root, tpl_kids, zones = _analyze_template(tpl_path)
        sel = [z for z in zones if z["name"] in _norm_kinds(params.get("zones"))]
        if not sel:
            return False
        root = _read_part(Path(path), "word/document.xml")
        body = root.find(f"{W}body")
        if body is None:
            return False
        kids = list(body)[:160]
        for z in sel:
            want = [_text_agnostic_xml(tpl_kids[i]) for i in z["kept"]]
            n = len(want)
            found = False
            for s in range(0, max(1, len(kids) - n + 1)):
                if all(_text_agnostic_xml(kids[s + t]) == want[t]
                       for t in range(n) if s + t < len(kids)):
                    found = True
                    break
            if not found:
                return False
        return True
    except Exception:  # noqa: BLE001
        return False


# ── 编译交付通报扫描（2026-10-09）─────────────────────────────

def scan_cover_candidates(path, template_path) -> dict:
    """封面重建候选扫描（只读；编译交付时提示「封面可按模板重建」）。

    返回 {ok, reason?, span, slots, already_template}：
    - ok=True：模板含封面区 + 目标前部可定位封面块（含无分节兜底）；
    - slots：目标封面块「标签：值」候选（通报展示 / 执行前核对用）；
    - already_template：目标封面已含模板封面首行文本（已重建 → 不提示）。
    全部异常吞掉返回 ok=False——通报环节绝不影响主流程。
    """
    try:
        tgt = Path(path)
        tpl = _resolve_template_path(str(template_path), tgt)
        if not tpl.exists():
            return {"ok": False, "reason": "template_missing"}
        _tpl_root, _tpl_kids, zones = _analyze_template(tpl)
        cover_zones = [z for z in zones if z["name"] in ("cover", "title_page")]
        if not cover_zones:
            return {"ok": False, "reason": "template_no_cover"}
        tgt_root, blocks = _target_blocks(tgt)
        if tgt_root is None:
            return {"ok": False, "reason": "target_unreadable"}
        span = _find_target_map(blocks).get("cover") \
            or _cover_span_fallback(tgt_root, blocks)
        if not span:
            return {"ok": False, "reason": "no_target_cover"}
        tpl_first = ""
        for z in cover_zones:
            for t in z.get("preview") or []:
                if t.strip():
                    tpl_first = t.strip()
                    break
            if tpl_first:
                break
        kids = list(tgt_root.find(f"{W}body"))
        joined = "".join(_para_text(el).strip()
                         for el in kids[span[0]: span[1] + 1])
        already = bool(tpl_first) and tpl_first[:8] in joined
        slots = {}
        for el in kids[span[0]: span[1] + 1]:
            if etree.QName(el).localname != "p":
                continue
            t = _para_text(el).strip()
            m = _SLOT_RE.match(t)
            if not m:
                continue
            label = re.sub(r"[\s\u3000]+", "", m.group(1))
            val = m.group(2).strip()
            if label and val and label not in slots:
                slots[label] = val
        # 2026-10-17：提示里的槽位必须用**模板的标签**（build_cover_from_template
        # 按模板标签回填）。旧实现直接报目标文档的标签（学院/专业/学号/姓名），
        # 与模板标签（专 业 名 称/学 生 学 号/学 生 姓 名）对不上 → 重建出的封面
        # 三个字段留 XXX 占位。这里先按语义匹配成「模板标签 → 目标里的值」。
        tpl_labels = [s["label_ns"] for z in cover_zones for s in z["slots"]]
        matched, _ = _match_slot_values(slots, tpl_labels)
        # 题目/英文题目不走模板标签（写标题区），但提示里必须带上——否则 LLM 只
        # 填槽位、标题行仍留模板示例文字（2026-10-17 真机：重建后标题是模板的）。
        for _k in ("题目", "英文题目"):
            if slots.get(_k):
                matched.setdefault(_k, slots[_k])
        return {"ok": True, "span": [span[0], span[1]],
                # slots 用模板标签（回填口径），没匹配上就留空——绝不回落到目标
                # 文档标签，否则提示给出的是模板填不进去的键（2026-10-17）。
                "slots": matched,
                "template_slots": tpl_labels,
                "target_slots": slots,
                "already_template": already}
    except Exception:  # noqa: BLE001
        return {"ok": False, "reason": "scan_error"}
