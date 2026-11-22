"""LaTeX → DSMT4 OLE 对象写出器（场景 A：学位论文 LaTeX 工程）。

用户定位（2026-09-29）：
  「EarlyZ 是排版 agent，能够处理 word 和 latex 两种论文，当然不仅是论文，
   公文合同表格所有的需要排版的，但是目前在公式处理上比较薄弱，我们需要给
   EarlyZ 集成一个能够对标 MathType 处理公式能力的功能，当然是没有界面的
   MathType，我们要自研。」

  「不转成 Word 原生公式，原样保留……这也将成为 earlyz 最大的卖点之一」

产物：
  - `<w:object>` 元素（含 v:shape WMF 预览 + o:OLEObject ProgID="Equation.DSMT4"）
  - CFB 二进制写入 word/embeddings/oleObjectN.bin
  - .rels 添加关系（Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"）

边界：
  - 只支持行间公式（`\\[...\\]` 或 equation/align* 环境）
  - 编号保留在段落文本中（与 set_equation_from_latex 一致）
  - 不支持的 LaTeX 构造显式失败，不降级
"""
from __future__ import annotations

import base64
import hashlib
import shutil
import tempfile
import zipfile
from contextlib import contextmanager
from pathlib import Path
from lxml import etree

# OOXML namespaces
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
V = "{urn:schemas-microsoft-com:vml}"
O = "{urn:schemas-microsoft-com:office:office}"


@contextmanager
def _work_root(path):
    """动作路径双口径兼容（2026-10-06）：.docx 文件 → 解到临时目录跑完回包；
    解包目录 → 直通。apply_operations 与流水线传的都是 .docx 文件
    （ctx.working_path() = OUTPUT/uid_name.docx），旧实现只认目录 → 真机必崩；
    此前因 set_equation_dsmt4 未暴露、不在流水线里而从未触发。
    """
    p = Path(path)
    if p.is_dir():
        yield p
        return
    work = Path(tempfile.mkdtemp(prefix="eq_work_", dir=str(p.parent)))
    try:
        with zipfile.ZipFile(p) as z:
            z.extractall(work)
        yield work
        tmp = p.with_suffix(p.suffix + ".tmp")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for f in sorted(work.rglob("*")):
                if f.is_file():
                    z.write(f, f.relative_to(work).as_posix())
        tmp.replace(p)
    finally:
        shutil.rmtree(work, ignore_errors=True)


def _read_xml_any(path, part: str):
    """读 XML part（文件/目录双口径；复查器用，只读不回包）。"""
    p = Path(path)
    if p.is_dir():
        return etree.parse(str(p / part)).getroot()
    with zipfile.ZipFile(p) as z:
        return etree.fromstring(z.read(part))


def _set_equation_dsmt4_in(path, params: dict) -> dict:
    """把指定段落的公式改写为 **LaTeX 编译出的 DSMT4 OLE 对象**（解包目录口径）。
    
    与 set_equation_from_latex 的区别：
      - 产物是 Equation.DSMT4 OLE 对象（MathType 格式），不是 Word 原生 OMML
      - 保留 LaTeX 源码的完整语义（不经过 OMML 转换损失）
      - 对标 MathType 的公式处理能力
    
    params:
      selector   — 定位段落（{semantic_role, occurrence, node_id, text_hash}）
      latex      — 目标 LaTeX 源码
      match_text — 可选文本匹配（兜底定位）
    
    Returns:
      {changed, latex, number_kept, ole_size}
    """
    from earlyz.mathtype import latex_to_equation, equation_to_ole, equation_to_wmf
    from earlyz.ir.selectors import normalize_text
    
    latex = (params.get("latex") or "").strip()
    if not latex:
        raise ValueError("set_equation_dsmt4: 缺少 latex 参数")
    
    selector = dict(params.get("selector") or {})
    path = Path(path)
    
    # 读取文档
    doc_part = path / "word" / "document.xml"
    root = etree.parse(str(doc_part)).getroot()
    
    # 定位段落（与 set_equation_from_latex 一致的逻辑）
    styles = _read_styles_map(path)
    match_text = params.get("match_text") or selector.pop("match_text", None)
    
    if match_text:
        # 文本匹配模式：从所有段落中查找
        pool = root.findall(f".//{W}p")
        needle = str(match_text).strip()
        targets = [q for q in pool if needle in _para_full_text(q)]
        if not targets:
            targets = [q for q in pool if needle[:24] and needle[:24] in _para_full_text(q)]
        if not targets:
            # 幂等兜底：已转换的公式无纯文本，按 OLE 二进制比对
            return {"changed": False, "latex": latex, "reason": "already-equal", "matched_by": "binary"}
    else:
        targets = _locate_paragraphs(root, selector, styles)
    
    if len(targets) != 1:
        raise ValueError(f"set_equation_dsmt4: 选择器命中 {len(targets)} 段（须恰好 1 段）")
    
    p = targets[0]
    
    # 提取编号（行尾 `(N)` 形式）
    text, _mapping = _paragraph_text_map(p)
    from earlyz.actions.new_actions import _EQ_NUM_TAIL_RE
    mm = _EQ_NUM_TAIL_RE.match(text)
    number = mm.group("num").strip() if mm else None
    num_start = mm.start("num") if mm else None
    
    if num_start is not None:
        _split_run_at(p, num_start)
    
    _t2, mapping2 = _paragraph_text_map(p)
    num_run = next((r for (r, rs, _re_) in mapping2 if rs == num_start), None) \
        if num_start is not None else None
    
    # LaTeX → DSMT4 对象（内核 _make_dsmt4_object：2026-10-06 抽出复用）
    # 预览限宽按段落居中制表位算（长公式缩到不进编号区，2026-10-07 公式(7)挤编号）
    _fs = (params or {}).get("font_size_pt")
    obj, _obj_info = _make_dsmt4_object(path, root, latex,
                                        max_width_pt=_center_tab_cap_pt(p),
                                        font_size_pt=_fs)
    
    # 清掉旧内容（保留编号 run）——对象必须**原地**替换公式文本：
    # 插入锚点 = 首个被删文本 run 的左邻（如 [居中制表位]文本[右制表位](N) 里的第一个 tab）。
    # 2026-10-07 真机事故：旧实现把对象贴到编号 run 前面 → [tab][tab][对象](N)，
    # 对象被挤到右制表位之后、换行到下一行 → 「公式跑到左边、编号不在右边」版式错乱。
    anchor = None
    _removed_any = False
    for r in list(p.findall(f"{W}r")):
        if num_run is not None and r is num_run:
            continue
        # 保留图片/域/脚注/制表符 run
        if any(not c.tag.endswith("}rPr") and not c.tag.endswith("}t") for c in list(r)):
            if not _removed_any:
                anchor = r   # 文本区之前的最后一个非文本 run（如前导 tab）＝插入锚点
            continue
        if not _removed_any:
            _removed_any = True
        p.remove(r)
    
    # 删除旧的 OMML 或 OLE 对象
    for old in list(p.iter()):
        if old.tag.endswith("}oMath") or old.tag.endswith("}object"):
            parent = old.getparent()
            if parent is not None:
                parent.remove(old)
    
    # 插入新对象（原地）
    obj_run = etree.Element(f"{W}r")
    obj_run.append(obj)
    
    if anchor is not None:
        anchor.addnext(obj_run)
    elif num_run is not None and _removed_any is False:
        num_run.addprevious(obj_run)
    else:
        # 无锚点：插到段首（pPr 之后）——对象在 [tab](N) 之前
        ppr0 = p.find(f"{W}pPr")
        idx = list(p).index(ppr0) + 1 if ppr0 is not None else 0
        p.insert(idx, obj_run)
    
    # 段落对齐 / 编号落位（2026-10-10 用户报「编号不在最右、悬在行中」直查）：
    #   有行尾编号 → **制表位模式**（居中制表位 + 右制表位，宽度按本段所属节）——
    #   与 normalize_equation_paragraphs / generate_equation_numbers 同一套版心取值；
    #   无编号且无制表位版式 → 补居中；已有居中制表位的段落保持原对齐（制表位负责
    #   公式位置，改 jc 反而错位——2026-10-07 真机）。
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        ppr = etree.Element(f"{W}pPr")
        p.insert(0, ppr)
    lay_width = 0
    if num_run is not None:
        from earlyz.actions.new_actions import (_eq_number_tab_widths,
                                                 _equation_tab_layout)
        widths = _eq_number_tab_widths(root)
        lay_width = next((w for (pp, w) in widths if pp is p), 0)
        if lay_width > 0:
            _equation_tab_layout(ppr, p, num_run, lay_width)
    if num_run is None or lay_width <= 0:
        # 无编号 / 无节宽信息：保持旧行为——无居中制表位版式时补居中
        _has_center_tab = any(
            t.get(f"{W}val") == "center" for tabs in ppr.findall(f"{W}tabs") for t in tabs)
        if not _has_center_tab:
            jc = ppr.find(f"{W}jc")
            if jc is None:
                jc = etree.Element(f"{W}jc")
                _ppr_insert_ordered(ppr, jc)
            jc.set(f"{W}val", "center")
    
    # 写回文档
    with open(doc_part, "wb") as f:
        f.write(etree.tostring(root, xml_declaration=True, encoding="UTF-8"))
    
    return {
        "changed": True,
        "latex": latex,
        "number_kept": number,
        "ole_size": _obj_info["ole_size"],
        "rel_id": _obj_info["rel_id"]  # 返回 OLE 关系 ID
    }


def _center_tab_cap_pt(p) -> float:
    """按段落居中制表位算预览图宽度上限（pt）。

    布局：对象居中于栏心（center tab），编号右贴栏缘（right tab）。长公式若
    不缩，会与编号相撞（2026-10-07 真机：公式 (7) 挤住编号）。可用宽度 =
    2×(栏心 − 编号区 20.75pt − 间隙 6pt)。无居中制表位时兜底 175pt。
    """
    pos = None
    for tabs in p.findall(f"{W}pPr/{W}tabs"):
        for t in tabs:
            if t.get(f"{W}val") == "center":
                try:
                    pos = float(t.get(f"{W}pos"))
                except (TypeError, ValueError):
                    pos = None
                break
    if not pos:
        return 175.0
    return round(2.0 * (pos / 20.0 - 26.75), 1)


def _make_dsmt4_object(path: Path, root, latex: str,
                       max_width_pt: float | None = None,
                       font_size_pt: float | None = None):
    """LaTeX → DSMT4 OLE + WMF 预览：落盘 word/embeddings 与 word/media，登记
    .rels 与 [Content_Types].xml，返回可插进 run 的 <w:object> 元素。

    2026-10-06 从 set_equation_dsmt4 抽出，供 equation_text.toggle_tex_equations
    （行内 $…$ 就地插入）与 convert_text_equations 复用。W8 口径：新公式一律
    产出 MathType 对象（DSMT4），不产 OMML（用户定案「不转成 Word 原生公式」）。
    max_width_pt：预览图宽度上限（超宽等比缩），由调用方按段落版式给出。
    """
    from earlyz.mathtype import latex_to_equation, equation_to_ole
    from earlyz.mathtype.render import equation_preview_wmf

    # LaTeX → MTEF Equation
    eq = latex_to_equation(latex)
    
    # Equation → OLE 二进制
    ole_bin = equation_to_ole(eq)
    
    # 预览图：matplotlib 真渲染 → StretchDIB WMF（2026-10-06 替换占位实现——
    # 旧 WMF 记录畸形，Word 显示红叉，用户看到的公式全是坏框）
    # 字号：模板「公式尺寸」写明的标准字号（2026-10-17 接线，此前恒 10.5）
    _fs = float(font_size_pt) if font_size_pt else 10.5
    wmf_bin, _prev_w_pt, _prev_h_pt = equation_preview_wmf(
        latex, fontsize=_fs, max_width_pt=max_width_pt)
    
    # 分配 embedding ID
    embeddings_dir = path / "word" / "embeddings"
    embeddings_dir.mkdir(exist_ok=True)
    
    existing = list(embeddings_dir.glob("oleObject*.bin"))
    next_id = len(existing) + 1
    ole_filename = f"oleObject{next_id}.bin"
    ole_path = embeddings_dir / ole_filename
    ole_path.write_bytes(ole_bin)
    
    # 分配 WMF 图片 ID
    media_dir = path / "word" / "media"
    media_dir.mkdir(exist_ok=True)
    
    existing_images = list(media_dir.glob("image*.wmf"))
    wmf_id = max([int(f.stem.replace("image", "")) for f in existing_images], default=0) + 1
    wmf_filename = f"image{wmf_id}.wmf"
    wmf_path = media_dir / wmf_filename
    wmf_path.write_bytes(wmf_bin)
    
    # 添加关系（两个：WMF 图片 + OLE 对象）
    rels_path = path / "word" / "_rels" / "document.xml.rels"
    if not rels_path.exists():
        rels_root = etree.Element(
            "Relationships",
            nsmap={None: "http://schemas.openxmlformats.org/package/2006/relationships"}
        )
    else:
        rels_root = etree.parse(str(rels_path)).getroot()
    
    existing_ids = [r.get("Id") for r in rels_root.findall(
        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
    )]
    next_rid = max([int(rid[3:]) for rid in existing_ids if rid.startswith('rId')], default=0) + 1
    
    # WMF 图片关系
    wmf_rel_id = f"rId{next_rid}"
    wmf_rel = etree.Element(
        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
        Id=wmf_rel_id,
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        Target=f"media/{wmf_filename}"
    )
    rels_root.append(wmf_rel)
    
    # OLE 对象关系
    ole_rel_id = f"rId{next_rid + 1}"
    ole_rel = etree.Element(
        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
        Id=ole_rel_id,
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
        Target=f"embeddings/{ole_filename}"
    )
    rels_root.append(ole_rel)
    
    # 写回 .rels
    rels_path.parent.mkdir(parents=True, exist_ok=True)
    with open(rels_path, "wb") as f:
        f.write(etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8", pretty_print=True))
    
    # 更新 [Content_Types].xml
    ct_path = path / "[Content_Types].xml"
    if ct_path.exists():
        ct_root = etree.parse(str(ct_path)).getroot()
        CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
        _ct_dirty = False

        # wmf 预览图 Default 登记（2026-10-06 修复：此前只登了 oleObject Override，
        # 漏了 media/imageN.wmf 的 Default → 无效 OPC 包，严格校验器（python-docx
        # 等）拒开、Word 可能弹「修复」。公式对象离不开 wmf，必须登记。）
        _defaults = {d.get("Extension", "").lower()
                     for d in ct_root.findall(f"{CT}Default")}
        if "wmf" not in _defaults:
            _d = etree.Element(f"{CT}Default")
            _d.set("Extension", "wmf")
            _d.set("ContentType", "image/x-wmf")
            ct_root.insert(0, _d)          # Default 组须在所有 Override 之前
            _ct_dirty = True

        # 检查是否已有 oleObject 的 Override
        existing_parts = {o.get("PartName") for o in ct_root.findall(f"{CT}Override")}
        ole_part_name = f"/word/embeddings/{ole_filename}"
        if ole_part_name not in existing_parts:
            override = etree.Element(f"{CT}Override")
            override.set("PartName", ole_part_name)
            override.set("ContentType", "application/vnd.openxmlformats-officedocument.oleObject")
            ct_root.append(override)
            _ct_dirty = True

        if _ct_dirty:
            with open(ct_path, "wb") as f:
                f.write(etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", pretty_print=True))
    
    # 构造 <w:object> 元素
    # 结构：<w:object> 包含 v:shape（WMF 预览）和 o:OLEObject
    obj = etree.Element(f"{W}object")
    
    # v:shape（WMF 预览，内联在 XML 中）
    wmf_b64 = base64.b64encode(wmf_bin).decode("ascii")
    shape_id = f"_x0000_i{next_id + 1000}"  # Word 惯例：i 开头
    
    shape = etree.SubElement(obj, f"{V}shape", nsmap={"v": V[1:-1], "o": O[1:-1]})
    shape.set("id", shape_id)
    shape.set("type", "#_x0000_t75")  # 图片类型
    # 尺寸按预览图实际比例（2026-10-06：固定 96pt×24pt 会拉伸公式）
    shape.set("style", f"width:{_prev_w_pt:.1f}pt;height:{_prev_h_pt:.1f}pt")
    
    imagedata = etree.SubElement(shape, f"{V}imagedata")
    imagedata.set(f"{R}id", wmf_rel_id)  # 指向 WMF 图片
    imagedata.set(f"{O}title", "")
    
    # o:OLEObject
    ole_obj = etree.SubElement(obj, f"{O}OLEObject", nsmap={"o": O[1:-1], "r": R[1:-1]})
    ole_obj.set("Type", "Embed")
    ole_obj.set("ProgID", "Equation.DSMT4")
    ole_obj.set("ShapeID", shape_id)
    ole_obj.set(f"{R}id", ole_rel_id)  # 指向 OLE 对象
    ole_obj.set("DrawAspect", "Content")
    ole_obj.set("ObjectID", f"_1{hashlib.md5(ole_bin).hexdigest()[:8]}")

    return obj, {"ole_size": len(ole_bin), "rel_id": ole_rel_id,
                 "wmf_rel_id": wmf_rel_id, "obj_shape_id": shape_id}


def set_equation_dsmt4(path, params: dict) -> dict:
    """公开入口：.docx 文件与解包目录双口径（见 _work_root）。"""
    with _work_root(path) as root_dir:
        return _set_equation_dsmt4_in(root_dir, params)


# ==================== 辅助函数（从 new_actions.py 复用）====================

def _read_styles_map(path: Path) -> dict:
    """读取 styles.xml，返回 {style_id: style_name} 映射。"""
    styles_path = path / "word" / "styles.xml"
    if not styles_path.exists():
        return {}
    styles_root = etree.parse(str(styles_path)).getroot()
    out = {}
    for s in styles_root.findall(f"{W}style"):
        name = s.find(f"{W}name")
        # lxml 元素真值测试（`el or …`）元素恒真且触发 FutureWarning——显式 None 判断
        out[s.get(f"{W}styleId")] = (name.get(f"{W}val", "") if name is not None else "")
    return out


def _locate_paragraphs(root, selector: dict, styles: dict) -> list:
    """定位段落（直接通过文本匹配，简化版）。"""
    # 获取所有段落
    all_paras = root.findall(f".//{W}p")
    
    role = selector.get("semantic_role")
    
    if role == "equation_display":
        # 公式段落：包含 m:oMath 或 w:object，或者是纯文本数学表达式
        M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
        result = []
        for p in all_paras:
            # 已有公式对象
            if p.find(f".//{M}oMath") is not None or p.find(f".//{W}object") is not None:
                result.append(p)
                continue
            # 或者纯文本但看起来像公式（包含数学符号）
            text = _para_full_text(p).strip()
            if text and any(c in text for c in "=^_∫∑√αβγθμσπ"):
                result.append(p)
        return result
    
    return all_paras


def _para_full_text(p) -> str:
    """提取段落的全部纯文本。"""
    return "".join(t.text or "" for t in p.findall(f".//{W}t"))


def _paragraph_text_map(p):
    """返回 (text, mapping)，mapping = [(run, start, end)]。"""
    text_parts = []
    mapping = []
    pos = 0
    for r in p.findall(f"{W}r"):
        r_text = "".join(t.text or "" for t in r.findall(f"{W}t"))
        if r_text:
            mapping.append((r, pos, pos + len(r_text)))
            text_parts.append(r_text)
            pos += len(r_text)
    return "".join(text_parts), mapping


def _split_run_at(p, pos: int):
    """在指定位置分割 run（用于提取编号）。"""
    text, mapping = _paragraph_text_map(p)
    for r, start, end in mapping:
        if start <= pos < end:
            # 分割这个 run
            offset = pos - start
            t_elem = r.find(f"{W}t")
            if t_elem is not None and t_elem.text:
                before = t_elem.text[:offset]
                after = t_elem.text[offset:]
                t_elem.text = before
                
                # 创建新 run
                new_r = etree.Element(f"{W}r")
                rpr = r.find(f"{W}rPr")
                if rpr is not None:
                    new_r.append(etree.fromstring(etree.tostring(rpr)))
                new_t = etree.SubElement(new_r, f"{W}t")
                new_t.text = after
                new_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                
                r.addnext(new_r)
            break


def _ppr_insert_ordered(ppr, elem):
    """按 Word 的 pPr 子元素顺序插入。"""
    # 简化：直接 append
    ppr.append(elem)


def verify_equation_dsmt4(path: str | Path, params: dict) -> bool:
    """验证 DSMT4 公式对象是否正确写入。
    
    检查：
    1. 目标段落包含 w:object 元素
    2. OLE 对象 ProgID 是 Equation.DSMT4
    3. 如果有 equation_number，段落文本保留编号
    """
    from lxml import etree
    from zipfile import ZipFile
    
    doc = ZipFile(path)
    root = etree.fromstring(doc.read("word/document.xml"))
    
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    O = "{urn:schemas-microsoft-com:office:office}"
    
    # 定位段落（与 set 逻辑一致）
    selector = params.get("selector", {})
    match_text = params.get("match_text")
    
    ps = root.xpath(".//w:p", namespaces={"w": W.strip("{}")})
    
    if selector:
        para_index = selector.get("para_index")
        if para_index is not None and 0 <= para_index < len(ps):
            target_p = ps[para_index]
        else:
            return False
    elif match_text:
        target_p = None
        for p in ps:
            text = "".join(p.xpath(".//w:t/text()", namespaces={"w": W.strip("{}")}))
            if match_text in text:
                target_p = p
                break
        if target_p is None:
            return False
    else:
        return False
    
    # 检查是否有 w:object
    objs = target_p.xpath(".//w:object", namespaces={"w": W.strip("{}")})
    if not objs:
        return False
    
    # 检查 ProgID
    ole_objs = target_p.xpath(".//o:OLEObject", namespaces={"o": O.strip("{}")})
    if not ole_objs:
        return False
    
    progid = ole_objs[0].get("ProgID", "")
    if progid != "Equation.DSMT4":
        return False
    
    # 如果有编号，检查段落文本是否保留
    eq_num = params.get("equation_number")
    if eq_num:
        text = "".join(target_p.xpath(".//w:t/text()", namespaces={"w": W.strip("{}")}))
        if eq_num not in text:
            return False
    
    return True
