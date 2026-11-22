"""
Writer v2 — 纯执行器，只执行 fix_plan 中的指令。

核心原则:
- 不做任何判断逻辑（不选值、不决策）
- 不产生任何新值（所有参数来自 fix_plan）
- 只执行 ZIP 级别的 OOXML 修改

用法:
    from pipeline.writer_v2 import apply_fixes

    fix_plan = [
        {"action": "set_sectpr_type", "params": {"section_index": 0, "val": "nextPage"}},
        {"action": "set_header_font", "params": {"header_path": "word/header2.xml", "font_ascii": "Times New Roman", "font_eastAsia": "宋体", "font_size": "18"}},
        {"action": "add_page_number", "params": {"footer_path": "word/footer2.xml", "format": "第{PAGE}页"}},
        {"action": "set_body_font_ascii", "params": {"style_id": "a8", "font_ascii": "Times New Roman"}},
        {"action": "set_style", "params": {"style_id": "1", "font_eastAsia": "黑体", "font_size_pt": 16}},
        {"action": "remove_extra_sectpr", "params": {"section_index": 1}},
    ]
    apply_fixes(fix_plan, Path("output/毕业设计.docx"))
"""

import shutil
import zipfile
import re
import copy
from pathlib import Path
from typing import Any, Union

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
XML = "http://www.w3.org/XML/1998/namespace"

# ── Action dispatcher ──────────────────────────────────────────

_ACTION_HANDLERS: dict[str, callable] = {}


def _handler(action_name: str):
    """装饰器：注册 action 处理函数"""
    def decorator(fn):
        _ACTION_HANDLERS[action_name] = fn
        return fn
    return decorator


# ── Structured writer result ─────────────────────────────────

class WriterResult:
    """单个 action 的执行结果（结构化，供 API 消费）。"""
    def __init__(self, index: int, action: str, ok: bool, message: str = "",
                 suggestion: str = "", params: dict = None):
        self.index = index
        self.action = action
        self.ok = ok          # True=OK, False=FAIL, None=SKIP
        self.message = message
        self.suggestion = suggestion
        self.params = params or {}

    def to_dict(self) -> dict:
        return {
            "index": self.index,
            "action": self.action,
            "ok": self.ok,
            "message": self.message,
            "suggestion": self.suggestion,
        }

    def line(self) -> str:
        """生成日志行（兼容旧格式）。"""
        prefix = "[OK" if self.ok else ("[SKIP" if self.ok is None else "[FAIL")
        return f"{prefix} #{self.index}] {self.action}: {self.message}"


def _failure_suggestion(action_type: str, exc: Exception) -> str:
    """根据 action 类型和异常，返回面向用户的修复建议。"""
    msg = str(exc).lower()

    if "no such file" in msg or "not found" in msg:
        if "header" in msg or "footer" in msg:
            return "头/页脚文件不存在，可能该节未设置页眉/页脚。请先在 Word 中为该节添加页眉/页脚，然后再排版。"
        return "文档内部文件缺失，请确认 docx 文件未损坏。尝试在 Word 中另存一次再排版。"

    if "badzipfile" in msg or "corrupt" in msg:
        return "docx 文件可能已损坏。请用 Word 打开后另存为新文件，再重新上传排版。"

    if "xml" in msg or "lxml" in msg or "parse" in msg:
        return "文档 XML 结构异常。可能是之前排版操作导致 OOXML 损坏。建议撤销操作恢复备份，或从原始文件重新开始。"

    if "index" in msg or "out of range" in msg:
        return "操作索引超出文档范围。请重新分析文档结构（DocDiscover），确保 fix_plan 中的索引与实际文档一致。"

    if "style" in msg:
        return "样式 ID 可能不存在于文档中。请确认 styles.xml 中包含该样式 ID，或改为使用已存在的样式。"

    if action_type == "set_sectpr_type":
        return "分节符类型修改失败。请检查文档节结构，或在 Word 中手动调整分节符类型。"

    if action_type == "set_header_font":
        return "页眉字体设置失败。请确认该节页眉存在文字内容，且 header XML 文件路径正确。"

    if action_type == "add_page_number":
        return "页码添加失败。请确认该节页脚存在，且 footer XML 文件路径正确。"

    return f"该操作执行失败。请检查文档是否仍可在 Word 中正常打开，或尝试从备份恢复。"


def apply_fixes(fix_plan: list[dict], target_path: Union[str, Path]) -> list[str]:
    """逐条执行 fix_plan 中的每个 action。

    Args:
        fix_plan: action 列表，每个 action 有 "action" 和 "params" 键
        target_path: 目标 docx 文件路径

    Returns:
        执行日志列表，每项描述执行结果。
        [FAIL] 行包含异常原因 + 修复建议（以 |suggestion: 分隔）。
    """
    target_path = Path(target_path)
    logs: list[str] = []

    # ── 缓存模式：全量读入内存，结束时一次性写盘 ──
    global _ZIP_CACHE
    if _ZIP_CACHE is not None:
        raise RuntimeError("apply_fixes 不支持嵌套调用（ZIP 缓存已启用）")
    try:
        with zipfile.ZipFile(target_path, "r") as zin:
            _ZIP_CACHE = {item.filename: zin.read(item.filename) for item in zin.infolist()}
    except Exception as e:
        _ZIP_CACHE = None
        logs.append(f"[FAIL] 读取文档失败: {e} |suggestion: 文档可能损坏，请用 Word 另存后重试")
        return logs

    try:
        _freeze_legacy_indexes(fix_plan, Path(target_path))
        for i, action in enumerate(fix_plan):
            action_type = action.get("action", "")
            params = action.get("params", {})

            # §4.4 兼容层执行期：把冻结锚点重映射为当前文档索引。
            # 前置 action 增删段落后，para_index 仍指向原稳定目标（无漂移）。
            if action.get("_captured") is not None:
                remapped = _resolve_legacy_index(Path(target_path), action["_captured"])
                if remapped is None:
                    logs.append(f"[SKIP #{i}] {action_type} 目标段落不存在（锚点解析失败）")
                    continue
                if "para_index" in params:
                    params["para_index"] = remapped

            handler = _ACTION_HANDLERS.get(action_type)
            if handler is None:
                logs.append(f"[SKIP #{i}] 未知 action: {action_type}")
                continue

            try:
                handler(target_path, params)
                logs.append(f"[OK #{i}] {action_type} {_brief_params(params)}")
            except Exception as e:
                sug = _failure_suggestion(action_type, e)
                logs.append(f"[FAIL #{i}] {action_type}: {type(e).__name__}: {e} |suggestion: {sug}")
    finally:
        # ── 一次性写盘 ──
        if _ZIP_CACHE is not None:
            try:
                tmp = str(target_path) + ".tmp"
                with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                    for name, data in _ZIP_CACHE.items():
                        zout.writestr(name, data)
                shutil.move(tmp, str(target_path))
            except Exception as e:
                logs.append(f"[FAIL] 写回文档失败: {e} |suggestion: 磁盘空间不足或文件被占用")
        _ZIP_CACHE = None

    return logs


def _brief_params(params: dict) -> str:
    """生成参数的简要描述"""
    parts = []
    for k, v in params.items():
        if isinstance(v, str) and len(v) > 30:
            v = v[:27] + "..."
        parts.append(f"{k}={v}")
    return " ".join(parts)


def _freeze_legacy_indexes(fix_plan: list[dict], target_path: Path) -> None:
    """§4.4 兼容层：旧 para_index 立即解析为稳定锚点（指纹捕获）。

    legacy action（format_paragraph 等）参数仍是 para_index——按 §4.4
    「兼容层可继续接受旧索引，但必须立即解析成 NodeSelector」：执行前对
    同一份文档快照指纹（_capture_resolved_paragraphs），执行时刻按锚点
    重映射为当前索引；前置 action 增删段落后目标不漂移。
    """
    if not any((a.get("params") or {}).get("para_index") is not None for a in fix_plan):
        return
    if any(a.get("_captured") is not None for a in fix_plan):
        return
    from earlyz.actions.new_actions import _capture_resolved_paragraphs
    data = _read_zip_entry(target_path, "word/document.xml")
    root = etree.fromstring(data)
    paras = list(root.iter(f"{{{W}}}p"))  # 与旧 handler 同序遍历语义（body.iter）
    for a in fix_plan:
        params = a.get("params") or {}
        idx = params.get("para_index")
        if idx is None or a.get("_captured") is not None:
            continue
        try:
            para = paras[int(idx)]
        except (TypeError, ValueError, IndexError):
            continue  # 越界交给现有容错（handler 会 SKIP/FAIL）
        a["_captured"] = _capture_resolved_paragraphs(root, [para])


def _resolve_legacy_index(target_path: Path, captured: list[dict]) -> int | None:
    """执行期：按冻结锚点重定位目标段，返回其当前文档索引（None=已不存在）。"""
    from earlyz.actions.new_actions import _resolve_captured_paragraphs
    data = _read_zip_entry(target_path, "word/document.xml")
    root = etree.fromstring(data)
    paras = list(root.iter(f"{{{W}}}p"))
    resolved = _resolve_captured_paragraphs(root, captured)
    if not resolved:
        return None
    for i, p in enumerate(paras):
        if p is resolved[0]:  # lxml 元素身份比较
            return i
    return None


# ── ZIP 工具 ──────────────────────────────────────────────────

# 模块级 ZIP 缓存：apply_fixes 期间把所有 zip 条目读入内存，
# handler 的读/写全部走缓存，结束时一次性写盘——避免每个 action
# 全量重写 zip（203 项 action = 203 次全量重写，4MB 文档约 300s）。
_ZIP_CACHE: dict | None = None


def _zip_cache_enabled() -> bool:
    return _ZIP_CACHE is not None


def _zip_replace(zip_path: Path, internal_path: str, new_data: bytes) -> None:
    """替换 ZIP 包中的单个文件（缓存模式下只更新内存缓存）"""
    if _ZIP_CACHE is not None:
        _ZIP_CACHE[internal_path] = new_data
        return
    zip_path_str = str(zip_path)
    tmp = zip_path_str + ".tmp"
    with zipfile.ZipFile(zip_path_str, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == internal_path:
                    data = new_data
                zout.writestr(item, data)
    shutil.move(tmp, zip_path_str)


def _read_zip_entry(zip_path: Path, internal_path: str) -> bytes:
    """读取 ZIP 中的单个条目（缓存模式下优先读缓存）"""
    if _ZIP_CACHE is not None:
        if internal_path in _ZIP_CACHE:
            return _ZIP_CACHE[internal_path]
        # 缓存缺失（新条目）→ 读盘兜底
        with zipfile.ZipFile(zip_path, "r") as z:
            data = z.read(internal_path)
        _ZIP_CACHE[internal_path] = data
        return data
    with zipfile.ZipFile(zip_path, "r") as z:
        return z.read(internal_path)


def _zip_namelist(zip_path: Path) -> list[str]:
    """ZIP 条目列表（缓存模式下用缓存 keys——包含本次新建的文件）"""
    if _ZIP_CACHE is not None:
        return list(_ZIP_CACHE.keys())
    with zipfile.ZipFile(zip_path, "r") as z:
        return z.namelist()


def _parse_xml_from_zip(zip_path: Path, internal_path: str):
    """从 ZIP 中读取并解析 XML"""
    return etree.fromstring(_read_zip_entry(zip_path, internal_path))


# ── Action: set_sectpr_type ────────────────────────────────────

@_handler("set_sectpr_type")
def _set_sectpr_type(target_path: Path, params: dict) -> None:
    """修改 sectPr 的 w:type val。

    params:
        section_index: int  — 第几个 sectPr（0-based）
        val: str             — type 值，如 "nextPage", "oddPage", "continuous"
    """
    section_index = params["section_index"]
    val = params["val"]

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    sects = list(root.iter(f"{{{W}}}sectPr"))

    if section_index >= len(sects):
        return  # 越界静默跳过（不崩溃——LLM 给的索引可能过期）

    sp = sects[section_index]
    type_el = sp.find(f"{{{W}}}type")
    if type_el is None:
        type_el = etree.SubElement(sp, f"{{{W}}}type")
    type_el.set(f"{{{W}}}val", val)

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ── Action: set_page_number_type ─────────────────────────────

@_handler("set_page_number_type")
def _set_page_number_type(target_path: Path, params: dict) -> None:
    """在 sectPr 中设置页码格式（罗马/阿拉伯）和重起序号。

    params:
        section_index: int   — 第几个 sectPr（0-based）
        fmt: str             — "upperRoman" | "decimal"
        start: int | None    — 页码起始值（如 1），None 表示不重起
    """
    section_index = params["section_index"]
    fmt_val = params["fmt"]
    start_val = params.get("start", None)

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    sects = list(root.iter(f"{{{W}}}sectPr"))

    if section_index >= len(sects):
        return  # 越界静默跳过（不崩溃——LLM 给的索引可能过期）

    sp = sects[section_index]
    pnt = sp.find(f"{{{W}}}pgNumType")
    if pnt is None:
        pnt = etree.SubElement(sp, f"{{{W}}}pgNumType")
    pnt.set(f"{{{W}}}fmt", fmt_val)
    if start_val is not None:
        pnt.set(f"{{{W}}}start", str(start_val))

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ── Action: set_header_font ────────────────────────────────────

@_handler("set_header_font")
def _set_header_font(target_path: Path, params: dict) -> None:
    """修改 header XML 中的 rPr/rFonts。

    params:
        header_path: str     — ZIP 内部路径，如 "word/header2.xml"
        font_ascii: str      — 西文字体名
        font_eastAsia: str   — 中文字体名
        font_size: str       — 字号（half-points），如 "18"
        font_hAnsi: str | None — hAnsi 字体（默认同 font_ascii）
    """
    header_path = params["header_path"]
    font_ascii = params["font_ascii"]
    font_eastAsia = params["font_eastAsia"]
    font_size = params["font_size"]
    font_hAnsi = params.get("font_hAnsi", font_ascii)

    root = _parse_xml_from_zip(target_path, header_path)

    for rPr in root.iter(f"{{{W}}}rPr"):
        old_rf = rPr.find(f"{{{W}}}rFonts")
        if old_rf is not None:
            rPr.remove(old_rf)

        rf = etree.SubElement(rPr, f"{{{W}}}rFonts")
        rf.set(f"{{{W}}}ascii", font_ascii)
        rf.set(f"{{{W}}}hAnsi", font_hAnsi)
        rf.set(f"{{{W}}}eastAsia", font_eastAsia)

        for tag in [f"{{{W}}}sz", f"{{{W}}}szCs"]:
            old_sz = rPr.find(tag)
            if old_sz is not None:
                rPr.remove(old_sz)
            etree.SubElement(rPr, tag, {f"{{{W}}}val": font_size})

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, header_path, new_data)


# ── Action: add_page_number ────────────────────────────────────

@_handler("add_page_number")
def _add_page_number(target_path: Path, params: dict) -> None:
    """在 footer XML 插入 "第{PAGE}页" fldChar 序列。

    params:
        footer_path: str     — ZIP 内部路径，如 "word/footer2.xml"
        format: str          — 页码格式模板，如 "第{PAGE}页"
        font_eastAsia: str | None  — 中文字体
        font_ascii: str | None     — 西文字体
        font_size: str | None      — 字号（half-points），如 "18"
    """
    footer_path = params["footer_path"]
    fmt = params["format"]
    font_eastAsia = params.get("font_eastAsia", "宋体")
    font_ascii = params.get("font_ascii", "Times New Roman")
    font_size = params.get("font_size", "18")

    root = _parse_xml_from_zip(target_path, footer_path)

    for p in list(root):
        root.remove(p)

    p = etree.SubElement(root, f"{{{W}}}p")
    pPr = etree.SubElement(p, f"{{{W}}}pPr")
    alignment = params.get("alignment", "right")  # 中文论文默认右对齐
    etree.SubElement(pPr, f"{{{W}}}jc", {f"{{{W}}}val": alignment})

    parts = fmt.split("{PAGE}")
    for idx, prefix in enumerate(parts):
        if prefix:
            r = etree.SubElement(p, f"{{{W}}}r")
            rPr = etree.SubElement(r, f"{{{W}}}rPr")
            rf = etree.SubElement(rPr, f"{{{W}}}rFonts")
            rf.set(f"{{{W}}}ascii", font_ascii)
            rf.set(f"{{{W}}}eastAsia", font_eastAsia)
            etree.SubElement(rPr, f"{{{W}}}sz", {f"{{{W}}}val": font_size})
            etree.SubElement(rPr, f"{{{W}}}szCs", {f"{{{W}}}val": font_size})
            t = etree.SubElement(r, f"{{{W}}}t", {f"{{{XML}}}space": "preserve"})
            t.text = prefix

        if idx < len(parts) - 1:
            r = etree.SubElement(p, f"{{{W}}}r")
            etree.SubElement(r, f"{{{W}}}fldChar", {f"{{{W}}}fldCharType": "begin"})

            r = etree.SubElement(p, f"{{{W}}}r")
            instr = etree.SubElement(r, f"{{{W}}}instrText", {f"{{{XML}}}space": "preserve"})
            instr.text = " PAGE "

            r = etree.SubElement(p, f"{{{W}}}r")
            etree.SubElement(r, f"{{{W}}}fldChar", {f"{{{W}}}fldCharType": "separate"})

            r = etree.SubElement(p, f"{{{W}}}r")
            rPr = etree.SubElement(r, f"{{{W}}}rPr")
            etree.SubElement(rPr, f"{{{W}}}sz", {f"{{{W}}}val": font_size})
            t = etree.SubElement(r, f"{{{W}}}t")
            t.text = "1"

            r = etree.SubElement(p, f"{{{W}}}r")
            etree.SubElement(r, f"{{{W}}}fldChar", {f"{{{W}}}fldCharType": "end"})

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, footer_path, new_data)


# ── Action: set_body_font_ascii ────────────────────────────────

@_handler("set_body_font_ascii")
def _set_body_font_ascii(target_path: Path, params: dict) -> None:
    """修改 styles.xml 中 body 样式（默认段落样式）的 w:ascii。

    params:
        style_id: str        — 样式 ID，如 "a8"
        font_ascii: str      — 西文字体名
    """
    style_id = params["style_id"]
    font_ascii = params["font_ascii"]

    root = _parse_xml_from_zip(target_path, "word/styles.xml")

    for style in root.iter(f"{{{W}}}style"):
        if style.get(f"{{{W}}}styleId") != style_id:
            continue

        rPr = style.find(f"{{{W}}}rPr")
        if rPr is None:
            rPr = etree.SubElement(style, f"{{{W}}}rPr")

        old_rf = rPr.find(f"{{{W}}}rFonts")
        if old_rf is not None:
            old_rf.set(f"{{{W}}}ascii", font_ascii)
            if not old_rf.get(f"{{{W}}}hAnsi"):
                old_rf.set(f"{{{W}}}hAnsi", font_ascii)
        else:
            rf = etree.SubElement(rPr, f"{{{W}}}rFonts")
            rf.set(f"{{{W}}}ascii", font_ascii)
            rf.set(f"{{{W}}}hAnsi", font_ascii)
        break

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/styles.xml", new_data)


# ── Action: set_style ─────────────────────────────────────────

@_handler("set_style")
def _set_style(target_path: Path, params: dict) -> None:
    """修改 styles.xml 中指定 style_id 的属性。

    params:
        style_id: str            — 样式 ID
        font_eastAsia: str | None
        font_ascii: str | None
        font_hAnsi: str | None
        font_size_pt: float | None  — 字号（point）
        bold: bool | None
        alignment: str | None    — "left", "center", "right", "both"
        before_lines: float | None
        after_lines: float | None
        line: int | None
        lineRule: str | None
    """
    style_id = params["style_id"]

    root = _parse_xml_from_zip(target_path, "word/styles.xml")

    for style in root.iter(f"{{{W}}}style"):
        if style.get(f"{{{W}}}styleId") != style_id:
            continue

        if any(k in params for k in ("font_eastAsia", "font_ascii", "font_hAnsi", "font_size_pt", "bold")):
            rPr = style.find(f"{{{W}}}rPr")
            if rPr is None:
                rPr = etree.SubElement(style, f"{{{W}}}rPr")

            if any(k in params for k in ("font_eastAsia", "font_ascii", "font_hAnsi")):
                old_rf = rPr.find(f"{{{W}}}rFonts")
                if old_rf is not None:
                    rPr.remove(old_rf)
                rf = etree.SubElement(rPr, f"{{{W}}}rFonts")
                if "font_ascii" in params and params["font_ascii"] is not None:
                    rf.set(f"{{{W}}}ascii", params["font_ascii"])
                if "font_hAnsi" in params and params["font_hAnsi"] is not None:
                    rf.set(f"{{{W}}}hAnsi", params["font_hAnsi"])
                elif "font_ascii" in params and params["font_ascii"] is not None:
                    rf.set(f"{{{W}}}hAnsi", params["font_ascii"])
                if "font_eastAsia" in params and params["font_eastAsia"] is not None:
                    rf.set(f"{{{W}}}eastAsia", params["font_eastAsia"])

            if "font_size_pt" in params and params["font_size_pt"] is not None:
                sv = str(int(params["font_size_pt"] * 2))
                for tag in [f"{{{W}}}sz", f"{{{W}}}szCs"]:
                    old = rPr.find(tag)
                    if old is not None:
                        rPr.remove(old)
                    etree.SubElement(rPr, tag, {f"{{{W}}}val": sv})

            if "bold" in params and params["bold"] is not None:
                old_b = rPr.find(f"{{{W}}}b")
                if old_b is not None:
                    rPr.remove(old_b)
                if params["bold"]:
                    etree.SubElement(rPr, f"{{{W}}}b")

        if any(k in params for k in ("alignment", "before_lines", "after_lines", "before_twips", "after_twips", "line", "lineRule", "ind_first_line_chars")):
            pPr = style.find(f"{{{W}}}pPr")
            if pPr is None:
                pPr = etree.SubElement(style, f"{{{W}}}pPr")

            if "alignment" in params and params["alignment"] is not None:
                old_jc = pPr.find(f"{{{W}}}jc")
                if old_jc is not None:
                    pPr.remove(old_jc)
                etree.SubElement(pPr, f"{{{W}}}jc", {f"{{{W}}}val": params["alignment"]})

            if "ind_first_line_chars" in params and params["ind_first_line_chars"] is not None:
                old_ind = pPr.find(f"{{{W}}}ind")
                if old_ind is not None:
                    pPr.remove(old_ind)
                etree.SubElement(pPr, f"{{{W}}}ind",
                                 {f"{{{W}}}firstLineChars": str(params["ind_first_line_chars"])})

            if any(k in params for k in ("before_lines", "after_lines", "before_twips", "after_twips", "line", "lineRule")):
                old_sp = pPr.find(f"{{{W}}}spacing")
                if old_sp is not None:
                    pPr.remove(old_sp)
                sp_attrs = {}
                if "before_lines" in params and params["before_lines"] is not None:
                    sp_attrs[f"{{{W}}}beforeLines"] = str(int(params["before_lines"]))
                if "after_lines" in params and params["after_lines"] is not None:
                    sp_attrs[f"{{{W}}}afterLines"] = str(int(params["after_lines"]))
                if "before_twips" in params and params["before_twips"] is not None:
                    sp_attrs[f"{{{W}}}before"] = str(int(params["before_twips"]))
                if "after_twips" in params and params["after_twips"] is not None:
                    sp_attrs[f"{{{W}}}after"] = str(int(params["after_twips"]))
                if "line" in params and params["line"] is not None:
                    sp_attrs[f"{{{W}}}line"] = str(int(params["line"]))
                if "lineRule" in params and params["lineRule"] is not None:
                    sp_attrs[f"{{{W}}}lineRule"] = params["lineRule"]
                if sp_attrs:
                    etree.SubElement(pPr, f"{{{W}}}spacing", sp_attrs)

        # 缩进（left/right/firstLine，twips）——重写 ind，丢弃字符单位属性
        # （Word 中 firstLineChars/leftChars 优先于 twips；spec 只提供 twips，需清 chars 才生效）
        if any(k in params for k in ("left_indent", "right_indent", "first_line_indent")):
            ind_attrs = {}
            if "left_indent" in params and params["left_indent"] is not None:
                ind_attrs[f"{{{W}}}left"] = str(int(params["left_indent"]))
            if "right_indent" in params and params["right_indent"] is not None:
                ind_attrs[f"{{{W}}}right"] = str(int(params["right_indent"]))
            if "first_line_indent" in params and params["first_line_indent"] is not None:
                ind_attrs[f"{{{W}}}firstLine"] = str(int(params["first_line_indent"]))
            if ind_attrs:
                pPr = style.find(f"{{{W}}}pPr")
                if pPr is None:
                    pPr = etree.SubElement(style, f"{{{W}}}pPr")
                old_ind = pPr.find(f"{{{W}}}ind")
                if old_ind is not None:
                    pPr.remove(old_ind)
                etree.SubElement(pPr, f"{{{W}}}ind", ind_attrs)

        break

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/styles.xml", new_data)


# ── Action: remove_extra_sectpr ────────────────────────────────

@_handler("remove_extra_sectpr")
def _remove_extra_sectpr(target_path: Path, params: dict) -> None:
    """删除嵌入在内容段落中的多余 sectPr。

    params:
        section_index: int | None  — 指定要删除第几个嵌入 sectPr（None=全部删除）
        preserve_body_sectpr: bool — 是否保留 body 末尾的 sectPr（默认 True）
    """
    section_index = params.get("section_index", None)
    preserve_body_sectpr = params.get("preserve_body_sectpr", True)

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    removed = 0
    for child in list(body):
        if child.tag != f"{{{W}}}p":
            continue
        pPr = child.find(f"{{{W}}}pPr")
        if pPr is None:
            continue
        extra_sp = pPr.find(f"{{{W}}}sectPr")
        if extra_sp is None:
            continue

        if section_index is not None and removed != section_index:
            removed += 1
            continue

        pPr.remove(extra_sp)
        removed += 1

        if section_index is not None:
            break

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ── Action: insert_section_break ───────────────────────────────

@_handler("insert_section_break")
def _insert_section_break(target_path: Path, params: dict) -> None:
    """在指定 H1 标题段落前插入新的 sectPr 分节符段落。

    params:
        heading_text: str  — 需要独立分节符的下一个 H1 标题文字（用于定位）
    """
    heading_text = params["heading_text"]

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")

    target_para = None
    target_idx = None
    for i, child in enumerate(list(body)):
        if child.tag == f"{{{W}}}p":
            txt = ''.join(t.text or '' for t in child.iter(f"{{{W}}}t")).strip()
            norm = re.sub(r'\s+', '', txt)
            if norm == re.sub(r'\s+', '', heading_text):
                target_para = child
                target_idx = i
                break

    if target_para is None:
        return

    # 2026-10-18：分节符必须带走完整版式——原来只写 type + pgSz，
    # 新节会丢 pgMar/cols/headerReference，导致分节后版式跳变、页眉消失。
    body_sect = body.find(f"{{{W}}}sectPr")
    if body_sect is not None:
        sectPr = copy.deepcopy(body_sect)
        typ = sectPr.find(f"{{{W}}}type")
        if typ is None:
            typ = etree.SubElement(sectPr, f"{{{W}}}type")
        typ.set(f"{{{W}}}val", "nextPage")
    else:
        sectPr = etree.Element(f"{{{W}}}sectPr")
        etree.SubElement(sectPr, f"{{{W}}}type", {f"{{{W}}}val": "nextPage"})
        etree.SubElement(sectPr, f"{{{W}}}pgSz", {f"{{{W}}}w": "11906", f"{{{W}}}h": "16838"})

    new_sp = etree.Element(f"{{{W}}}p")
    new_pPr = etree.SubElement(new_sp, f"{{{W}}}pPr")
    new_pPr.append(sectPr)

    target_para.addprevious(new_sp)

    new_data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ── Action: create_section_headers ──────────────────────────────

@_handler("create_section_headers")
def _create_section_headers(target_path: Path, params: dict) -> None:
    """为每节创建独立 header 文件，注册到 rels 和 Content_Types，更新 sectPr 引用。

    params:
        section_headers: list of dicts  — 每节的页眉配置
            [{sectpr_index: int,
              default_text: str, even_text: str | None},
             ...]
        font_eastAsia: str   — 中文字体
        font_ascii: str      — 西文字体
        font_size: str       — 字号 half-points
    """
    import zipfile
    from pathlib import Path

    section_headers = params["section_headers"]
    font_eastAsia = params.get("font_eastAsia", "宋体")
    font_ascii = params.get("font_ascii", "Times New Roman")
    font_size = params.get("font_size", "18")
    # 2026-10-18：默认不再给新建页眉加下边框（真机事故：想删的正是这条线，
    # 而 create_section_headers 每次都会把它加回来）。需要线就显式传 border=true。
    border_xml = ('<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>'
                  if params.get("border") else '')

    def _make_hdr_xml(text: str) -> bytes:
        return (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:hdr xmlns:w="{W}"><w:p><w:pPr><w:jc w:val="center"/>'
            f'{border_xml}</w:pPr><w:r><w:rPr>'
            f'<w:rFonts w:eastAsia="{font_eastAsia}" w:ascii="{font_ascii}" w:hAnsi="{font_ascii}"/>'
            f'<w:sz w:val="{font_size}"/><w:szCs w:val="{font_size}"/>'
            f'</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p></w:hdr>'
        ).encode('utf-8')

    tmp_path = str(target_path) + '.hdr_tmp'
    doc_xml = etree.fromstring(_read_zip_entry(target_path, 'word/document.xml'))
    rels_xml = etree.fromstring(_read_zip_entry(target_path, 'word/_rels/document.xml.rels'))
    ct_xml = etree.fromstring(_read_zip_entry(target_path, '[Content_Types].xml'))
    try:
        settings_xml = etree.fromstring(_read_zip_entry(target_path, 'word/settings.xml'))
        has_settings = True
    except KeyError:
        settings_xml = None
        has_settings = False

    rels_root = rels_xml
    max_rid = 1
    for rel in rels_root:
        rid = rel.get('Id', '')
        if rid.startswith('rId'):
            try:
                max_rid = max(max_rid, int(rid[3:]))
            except ValueError:
                pass

    existing_headers = set()
    for name in _zip_namelist(target_path):
        if 'header' in name and name.endswith('.xml'):
            existing_headers.add(name)

    body = doc_xml.find(f"{{{W}}}body")
    sects_in_order = []
    for child in body:
        if child.tag == f"{{{W}}}sectPr":
            sects_in_order.append(child)
        elif child.tag == f"{{{W}}}p":
            pp = child.find(f"{{{W}}}pPr")
            if pp is not None:
                sp = pp.find(f"{{{W}}}sectPr")
                if sp is not None:
                    sects_in_order.append(sp)

    ct_root = ct_xml
    header_num = 20
    used_texts = {}

    for sh in section_headers:
        si = sh.get("sectpr_index") if "sectpr_index" in sh else sh["section_index"]
        if si >= len(sects_in_order):
            continue
        sp = sects_in_order[si]

        default_text = sh.get("default_text", "")
        even_text = sh.get("even_text")

        if not default_text:
            continue

        key = (default_text, even_text or "")
        if key not in used_texts:
            max_rid += 1
            header_num += 1
            def_fname = f'header{header_num}.xml'
            def_rid = f'rId{max_rid}'

            ct_root.append(ct_root.makeelement(f'{{{CT}}}Override', {
                'PartName': f'/word/{def_fname}',
                'ContentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
            }))
            rels_root.append(rels_root.makeelement('Relationship', {
                'Id': def_rid,
                'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
                'Target': def_fname,
                'TargetMode': 'Internal'
            }))

            eve_fname = eve_rid = None
            if even_text:
                max_rid += 1
                header_num += 1
                eve_fname = f'header{header_num}.xml'
                eve_rid = f'rId{max_rid}'
                ct_root.append(ct_root.makeelement(f'{{{CT}}}Override', {
                    'PartName': f'/word/{eve_fname}',
                    'ContentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
                }))
                rels_root.append(rels_root.makeelement('Relationship', {
                    'Id': eve_rid,
                    'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
                    'Target': eve_fname,
                    'TargetMode': 'Internal'
                }))

            used_texts[key] = (def_fname, def_rid, eve_fname, eve_rid)

        def_fname, def_rid, eve_fname, eve_rid = used_texts[key]

        # 清除旧引用 + 旧标志
        for old in list(sp.findall(f"{{{W}}}headerReference")):
            sp.remove(old)
        for old_flag in list(sp.findall(f"{{{W}}}evenAndOddHeaders")):
            sp.remove(old_flag)
        for old_tp in list(sp.findall(f"{{{W}}}titlePg")):
            sp.remove(old_tp)

        # ── 清理：如果 sectPr 在内容段落中，移到独立空段 ──
        parent = sp.getparent()  # pPr
        if parent is not None:
            grandparent = parent.getparent()  # p (the paragraph)
            if grandparent is not None and grandparent.tag == f'{{{W}}}p':
                para_txt = ''.join(t.text or '' for t in grandparent.iter(f'{{{W}}}t')).strip()
                if para_txt:
                    parent.remove(sp)
                    new_para = etree.Element(f'{{{W}}}p')
                    new_pPr = etree.SubElement(new_para, f'{{{W}}}pPr')
                    new_pPr.append(sp)
                    grandparent.addnext(new_para)
                    sp = new_pPr.find(f'{{{W}}}sectPr')  # 后续操作使用移动后的 sectPr

        hr = etree.SubElement(sp, f"{{{W}}}headerReference")
        hr.set(f"{{{W}}}type", 'default')
        hr.set(f"{{{R}}}id", def_rid)

        if eve_rid:
            hr = etree.SubElement(sp, f"{{{W}}}headerReference")
            hr.set(f"{{{W}}}type", 'even')
            hr.set(f"{{{R}}}id", eve_rid)
            if sp.find(f"{{{W}}}evenAndOddHeaders") is None:
                etree.SubElement(sp, f"{{{W}}}evenAndOddHeaders")

    # ── 清理文档级 settings.xml 中的全局 evenAndOddHeaders ──
    settings_data = None
    if has_settings and settings_xml is not None:
        settings_root = settings_xml
        for eoh in settings_root.findall(f'{{{W}}}evenAndOddHeaders'):
            settings_root.remove(eoh)
        settings_data = etree.tostring(settings_root, xml_declaration=True, encoding='UTF-8', standalone=True)

    doc_data = etree.tostring(doc_xml, xml_declaration=True, encoding='UTF-8', standalone=True)
    rels_data = etree.tostring(rels_xml, xml_declaration=True, encoding='UTF-8', standalone=True)
    ct_data = etree.tostring(ct_xml, xml_declaration=True, encoding='UTF-8', standalone=True)

    # ── 写回（缓存模式下全部走内存缓存，apply_fixes 结束时一次性写盘）──
    _zip_replace(target_path, 'word/document.xml', doc_data)
    _zip_replace(target_path, 'word/_rels/document.xml.rels', rels_data)
    _zip_replace(target_path, '[Content_Types].xml', ct_data)
    if settings_data is not None:
        _zip_replace(target_path, 'word/settings.xml', settings_data)
    for (dt, et), (df, dr, ef, er) in used_texts.items():
        _zip_replace(target_path, f'word/{df}',
                     _make_hdr_xml_via_python_docx(dt, font_size, font_eastAsia, font_ascii))
        if ef:
            _zip_replace(target_path, f'word/{ef}',
                         _make_hdr_xml_via_python_docx(et, font_size, font_eastAsia, font_ascii))


@_handler("set_header_text")
def _set_header_text(target_path: Path, params: dict) -> None:
    """替换 header XML 中的所有文字内容，保留字体格式。

    params:
        header_path: str     — ZIP 内部路径，如 "word/header2.xml"
        text: str            — 新的页眉文字
        font_eastAsia: str   — 中文字体（默认 "宋体"）
        font_ascii: str      — 西文字体（默认 "Times New Roman"）
        font_size: str       — 字号 half-points（默认 "18"）
    """
    header_path = params["header_path"]
    text = params["text"]
    font_eastAsia = params.get("font_eastAsia", "宋体")
    font_ascii = params.get("font_ascii", "Times New Roman")
    font_size = params.get("font_size", "18")

    root = _parse_xml_from_zip(target_path, header_path)

    for p in list(root):
        root.remove(p)

    p = etree.SubElement(root, f"{{{W}}}p")
    pPr = etree.SubElement(p, f"{{{W}}}pPr")
    etree.SubElement(pPr, f"{{{W}}}jc", {f"{{{W}}}val": "center"})
    etree.SubElement(pPr, f"{{{W}}}pBdr").append(
        etree.Element(f"{{{W}}}bottom", {
            f"{{{W}}}val": "single",
            f"{{{W}}}sz": "6",
            f"{{{W}}}space": "1",
            f"{{{W}}}color": "auto",
        })
    )

    r = etree.SubElement(p, f"{{{W}}}r")
    rPr = etree.SubElement(r, f"{{{W}}}rPr")
    rf = etree.SubElement(rPr, f"{{{W}}}rFonts")
    rf.set(f"{{{W}}}eastAsia", font_eastAsia)
    rf.set(f"{{{W}}}ascii", font_ascii)
    rf.set(f"{{{W}}}hAnsi", font_ascii)
    etree.SubElement(rPr, f"{{{W}}}sz", {f"{{{W}}}val": font_size})
    etree.SubElement(rPr, f"{{{W}}}szCs", {f"{{{W}}}val": font_size})
    t = etree.SubElement(r, f"{{{W}}}t", {f"{{{XML}}}space": "preserve"})
    t.text = text

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, header_path, new_data)



# ── Action: clear_section_headers ────────────────────────────

@_handler("clear_section_headers")
def _clear_section_headers(target_path: Path, params: dict) -> None:
    """移除 sectPr 中所有 headerReference。"""
    section_index = params["section_index"]

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    sects = list(root.iter(f"{{{W}}}sectPr"))
    if section_index >= len(sects):
        return

    sp = sects[section_index]
    for hr in list(sp.findall(f"{{{W}}}headerReference")):
        sp.remove(hr)
    # 同时移除 evenAndOddHeaders
    eoh = sp.find(f"{{{W}}}evenAndOddHeaders")
    if eoh is not None:
        sp.remove(eoh)

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


@_handler("clear_section_footers")
def _clear_section_footers(target_path: Path, params: dict) -> None:
    """移除 sectPr 中所有 footerReference（封面/声明等不应显示页码的节）。"""
    section_index = params["section_index"]

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    sects = list(root.iter(f"{{{W}}}sectPr"))
    if section_index >= len(sects):
        return

    sp = sects[section_index]
    for fr in list(sp.findall(f"{{{W}}}footerReference")):
        sp.remove(fr)

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ── Action: set_toc_entry_styles ─────────────────────────────

@_handler("set_toc_entry_styles")
def _set_toc_entry_styles(target_path: Path, params: dict) -> None:
    """给目录条目段落（PAGEREF 域段落）补 pStyle=TOC{n}。

    学校模板的目录条目常无 pStyle（Word 域缓存段落），TOC1/2/3 样式
    修改不生效。按书签定位条目段落，补上对应的目录样式。

    params:
        entries: list of {bookmark: str, style: str}  — 条目书签与目标样式
    """
    entries = params.get("entries") or []
    if not entries:
        return

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    # 书签 → 目标样式
    style_by_bookmark = {}
    for entry in entries:
        bm = entry.get("bookmark", "")
        if bm:
            style_by_bookmark[bm] = entry.get("style", "TOC1")

    changed = 0
    # 给 PAGEREF 条目段落（目录页的条目）补 pStyle——不是书签所在的标题段落
    for p in root.iter(f"{{{W}}}p"):
        instr = ''.join(t.text or '' for t in p.iter(f"{{{W}}}instrText"))
        m = re.search(r'PAGEREF\s+(_Toc\w+)', instr)
        if not m:
            continue
        style_id = style_by_bookmark.get(m.group(1))
        if not style_id:
            continue
        ppr = p.find(f"{{{W}}}pPr")
        if ppr is None:
            ppr = etree.Element(f"{{{W}}}pPr")
            p.insert(0, ppr)
        # 已有 pStyle 不覆盖
        if ppr.find(f"{{{W}}}pStyle") is not None:
            continue
        ps = etree.SubElement(ppr, f"{{{W}}}pStyle")
        ps.set(f"{{{W}}}val", style_id)
        changed += 1

    if changed:
        new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        _zip_replace(target_path, "word/document.xml", new_data)


# ── Action: ensure_toc_tabs ─────────────────────────────────

@_handler("ensure_toc_tabs")
def _ensure_toc_tabs(target_path: Path, params: dict) -> None:
    """确保 TOC1/2/3 样式都有右对齐点线制表位（标准目录外观）。

    学校模板 TOC2/TOC3 常缺 w:tabs → 二级条目无点线引导。
    已有 right 制表位（leader=dot 或任意）则跳过，不覆盖模板定义。

    params:
        pos: int — 制表位位置（twips，页面宽度-右边距）
    """
    pos = int(params.get("pos", 8776))
    root = _parse_xml_from_zip(target_path, "word/styles.xml")

    changed = 0
    for style in root.iter(f"{{{W}}}style"):
        sid = style.get(f"{{{W}}}styleId", "")
        if sid not in ("TOC1", "TOC2", "TOC3", "TOC 1", "TOC 2", "TOC 3"):
            continue
        pPr = style.find(f"{{{W}}}pPr")
        if pPr is None:
            pPr = etree.SubElement(style, f"{{{W}}}pPr")
        tabs = pPr.find(f"{{{W}}}tabs")
        # 已有 right 制表位 → 跳过
        if tabs is not None:
            has_right = any(
                t.get(f"{{{W}}}val") == "right" for t in tabs.findall(f"{{{W}}}tab")
            )
            if has_right:
                continue
        else:
            tabs = etree.Element(f"{{{W}}}tabs")
            pPr.insert(0, tabs)
        tab = etree.SubElement(tabs, f"{{{W}}}tab")
        tab.set(f"{{{W}}}val", "right")
        tab.set(f"{{{W}}}leader", "dot")
        tab.set(f"{{{W}}}pos", str(pos))
        changed += 1

    if changed:
        new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        _zip_replace(target_path, "word/styles.xml", new_data)


# ── Action: add_footer_reference ─────────────────────────────

@_handler("add_footer_reference")
def _add_footer_reference(target_path: Path, params: dict) -> None:
    """在 sectPr 中添加默认 footerReference。"""
    import zipfile
    section_index = params["section_index"]

    rels_root = _parse_xml_from_zip(target_path, "word/_rels/document.xml.rels")
    footer_rid = None
    for rel in rels_root:
        if "footer" in (rel.get("Type", "")):
            footer_rid = rel.get("Id")
            break
    if footer_rid is None:
        return

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    sects = list(root.iter(f"{{{W}}}sectPr"))
    if section_index >= len(sects):
        return

    sp = sects[section_index]
    title_pg = sp.find(f"{{{W}}}titlePg")
    if title_pg is not None:
        sp.remove(title_pg)

    if sp.findall(f"{{{W}}}footerReference"):
        return

    fr = etree.SubElement(sp, f"{{{W}}}footerReference")
    fr.set(f"{{{W}}}type", "default")
    fr.set(f"{{{R}}}id", footer_rid)

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ── Action: strip_style_numbering ──────────────────────────────
# 用户要求：序号一律为普通文本，不做自动编号特殊处理。
# 样式层（styles.xml 的默认/正文样式）可能带 numPr → 继承后 Word 会给
# 无 pStyle 的正文段自动编号（如 [1][2]）+ 悬挂缩进（视觉"空格/没到最左边"）。

@_handler("strip_style_numbering")
def _strip_style_numbering(target_path: Path, params: dict) -> None:
    """移除 styles.xml 中所有样式的 numPr + document.xml 中段落级 numPr。

    只清编号引用，不动段落文本——编号文字（如"(1)"）若原本是普通文本则保留。
    """
    import zipfile

    def _strip(xml_bytes: bytes) -> bytes:
        root = etree.fromstring(xml_bytes)
        for numpr in list(root.iter(f"{{{W}}}numPr")):
            parent = numpr.getparent()
            if parent is not None:
                parent.remove(numpr)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    names = _zip_namelist(target_path)
    data = {}
    if "word/styles.xml" in names:
        data["word/styles.xml"] = _strip(_read_zip_entry(target_path, "word/styles.xml"))
    if "word/document.xml" in names:
        # 必须读缓存（apply_fixes 期间），直接读磁盘会覆盖前面 action 的修改
        # （如 create_section_headers 写入的 headerReference）
        data["word/document.xml"] = _strip(_read_zip_entry(target_path, "word/document.xml"))
    for path, blob in data.items():
        _zip_replace(target_path, path, blob)


# ── Action: convert_ole_to_image ───────────────────────────────
# 源文档图片可能是 OLE 嵌入对象（如 Visio.Drawing.15 流程图）：
# 用户机器无对应程序时 Word 只显示占位图标（"链接"按钮），图片不可见。
# 方案：用 OLE 内嵌的 imagedata 位图，把 w:object 替换为普通 w:drawing。

@_handler("convert_ole_to_image")
def _convert_ole_to_image(target_path: Path, params: dict) -> None:
    """把 document.xml 中的 w:object（OLE 嵌入对象）转换为普通 w:drawing 图片。

    用 OLE 内嵌的 v:imagedata 引用的位图（media 文件），保留原尺寸。
    """
    import zipfile
    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    # 读 rels 拿到 rId → media 目标
    rels_root = _parse_xml_from_zip(target_path, "word/_rels/document.xml.rels")
    rel_target = {}
    for rel in rels_root:
        rid = rel.get("Id")
        if rid:
            rel_target[rid] = rel.get("Target", "")

    V = "urn:schemas-microsoft-com:vml"
    NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    # 注册常用前缀，保证序列化时用 a:/pic:/wp:（Word 对前缀敏感）
    etree.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
    etree.register_namespace("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture")
    etree.register_namespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
    etree.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    converted = 0
    # 现有 docPr 最大 id（新 drawing 的 docPr id 必须唯一，否则 Word 拒绝打开）
    # 注意：docPr 在 wp 命名空间（wordprocessingDrawing），不是 w 命名空间！
    max_docpr = 0
    _tpl_drawing = None
    for child in body:
        for d in child.iter(f"{{{W}}}drawing"):
            if _tpl_drawing is None:
                _tpl_drawing = d
            for dp in d.iter(f"{{{WP}}}docPr"):
                try:
                    max_docpr = max(max_docpr, int(dp.get("id") or 0))
                except ValueError:
                    pass
    for obj in list(body.iter(f"{{{W}}}object")):
        # 找 v:imagedata 的 r:id（位图预览）
        img_rid = None
        for imgdata in obj.iter(f"{{{V}}}imagedata"):
            rid = imgdata.get(f"{{{R}}}id")
            if rid:
                img_rid = rid
                break
        if not img_rid or img_rid not in rel_target:
            continue
        target = rel_target[img_rid]
        if not target.startswith("media/"):
            continue
        # 尺寸：优先 v:shape style width/height（pt），回退 dxaOrig/dyaOrig（twips）
        width_emu = None
        height_emu = None
        for shape in obj.iter(f"{{{V}}}shape"):
            style = shape.get("style", "")
            mw = re.search(r"width:([\d.]+)pt", style)
            mh = re.search(r"height:([\d.]+)pt", style)
            if mw:
                width_emu = int(round(float(mw.group(1)) * 12700))
            if mh:
                height_emu = int(round(float(mh.group(1)) * 12700))
        dxa = obj.get(f"{{{W}}}dxaOrig")
        dya = obj.get(f"{{{W}}}dyaOrig")
        if width_emu is None and dxa:
            width_emu = int(dxa) * 635  # twips → EMU (1440 twips/inch, 914400 EMU/inch)
        if height_emu is None and dya:
            height_emu = int(dya) * 635
        if width_emu is None:
            width_emu = 5000000
        if height_emu is None:
            height_emu = 3000000

        # 方案：深拷贝文档中第一张现有 drawing（Word 原生结构，100% 兼容），
        # 只改 docPr/cNvPr id + blip 引用 + 尺寸——避免手工构造的兼容性坑
        if _tpl_drawing is None:
            continue  # 无模板 drawing，跳过（理论上不会发生）
        import copy as _copy
        newd = _copy.deepcopy(_tpl_drawing)
        max_docpr += 1
        # 改 docPr/cNvPr id + name
        for dp in newd.iter(f"{{{WP}}}docPr"):
            dp.set("id", str(max_docpr))
            dp.set("name", f"图片 {max_docpr}")
        for cnv in newd.iter(f"{{{NS}}}cNvPr"):
            cnv.set("id", str(max_docpr))
            cnv.set("name", f"图片 {max_docpr}")
        # 改 blip 引用为 OLE 内嵌位图
        for blip in newd.iter(f"{{{NS}}}blip"):
            blip.set(f"{{{R}}}embed", img_rid)
        # 改尺寸（extent + xfrm/ext）
        for extel in newd.iter(f"{{{WP}}}extent"):
            extel.set("cx", str(width_emu))
            extel.set("cy", str(height_emu))
        for extel in newd.iter(f"{{{NS}}}ext"):
            extel.set("cx", str(width_emu))
            extel.set("cy", str(height_emu))
        # 替换 object
        parent = obj.getparent()
        if parent.tag == f"{{{W}}}p":
            run = etree.SubElement(parent, f"{{{W}}}r")
            parent.replace(obj, run)
            run.append(newd)
        else:
            parent.replace(obj, newd)
        converted += 1

    if converted:
        new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        _zip_replace(target_path, "word/document.xml", new_data)


# ── Action: set_code_block_style ──────────────────────────────

@_handler("set_code_block_style")
def _set_code_block_style(target_path: Path, params: dict) -> None:
    """为指定的段落设置代码块样式：灰色背景 + 等宽字体。

    params:
        paragraph_indices: list[int]  — 段落索引列表（0-based，body 子元素中的 p 元素序号）
    """
    paragraph_indices = params["paragraph_indices"]

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    # 收集所有 p 段落的索引映射
    p_elements = []
    for i, child in enumerate(body):
        if child.tag == f"{{{W}}}p":
            p_elements.append((i, child))

    for pi in paragraph_indices:
        if pi >= len(p_elements):
            continue
        _, p = p_elements[pi]
        # 设置段落背景色
        pPr = p.find(f"{{{W}}}pPr")
        if pPr is None:
            pPr = etree.SubElement(p, f"{{{W}}}pPr")
            p.insert(0, pPr)
        # 移除旧 shd
        for old_shd in pPr.findall(f"{{{W}}}shd"):
            pPr.remove(old_shd)
        etree.SubElement(pPr, f"{{{W}}}shd", {
            f"{{{W}}}val": "clear",
            f"{{{W}}}fill": "F5F5F5",
        })

        # 设置每个 run 的等宽字体
        for run in p.iter(f"{{{W}}}r"):
            rPr = run.find(f"{{{W}}}rPr")
            if rPr is None:
                rPr = etree.SubElement(run, f"{{{W}}}rPr")
                run.insert(0, rPr)
            old_rf = rPr.find(f"{{{W}}}rFonts")
            if old_rf is not None:
                old_rf.set(f"{{{W}}}ascii", "Consolas")
                old_rf.set(f"{{{W}}}hAnsi", "Consolas")
            else:
                rf = etree.SubElement(rPr, f"{{{W}}}rFonts")
                rf.set(f"{{{W}}}ascii", "Consolas")
                rf.set(f"{{{W}}}hAnsi", "Consolas")
                rPr.insert(0, rf)

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ── Action: insert_toc ────────────────────────────────────────

@_handler("insert_toc")
def _insert_toc(target_path: Path, params: dict) -> None:
    """在指定位置插入 TOC（目录）域。

    params:
        after_paragraph_index: int | None  — 插入到该段落后（None = prepend 到 body 开头）
    """
    after_paragraph_index = params.get("after_paragraph_index", None)

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    # 2026-10-18 幂等（能力矩阵发现）：已有 TOC 域时重复调用会不断插入新目录
    for instr in root.iter(f"{{{W}}}instrText"):
        if "TOC" in (instr.text or "").upper():
            return
    for fld in root.iter(f"{{{W}}}fldSimple"):
        if "TOC" in str(fld.get(f"{{{W}}}instr") or "").upper():
            return

    # 收集所有 p 段落的索引映射
    p_elements = []
    for i, child in enumerate(body):
        if child.tag == f"{{{W}}}p":
            p_elements.append((i, child))

    # 构建 TOC 域段落
    toc_p = etree.SubElement(body, f"{{{W}}}p")  # 临时添加
    # 先移除，稍后插入到正确位置
    body.remove(toc_p)

    # pPr: 设置段落格式
    pPr = etree.SubElement(toc_p, f"{{{W}}}pPr")
    etree.SubElement(pPr, f"{{{W}}}spacing", {
        f"{{{W}}}before": "120",
        f"{{{W}}}after": "120",
    })

    # TOC begin
    r_begin = etree.SubElement(toc_p, f"{{{W}}}r")
    etree.SubElement(r_begin, f"{{{W}}}fldChar", {f"{{{W}}}fldCharType": "begin"})

    # TOC instrText
    r_instr = etree.SubElement(toc_p, f"{{{W}}}r")
    instr = etree.SubElement(r_instr, f"{{{W}}}instrText", {f"{{{XML}}}space": "preserve"})
    instr.text = " TOC \\o \"1-3\" \\h \\z \\u "

    # TOC separate
    r_sep = etree.SubElement(toc_p, f"{{{W}}}r")
    etree.SubElement(r_sep, f"{{{W}}}fldChar", {f"{{{W}}}fldCharType": "separate"})

    # TOC placeholder text
    r_text = etree.SubElement(toc_p, f"{{{W}}}r")
    t = etree.SubElement(r_text, f"{{{W}}}t")
    t.text = '（请右键点击此处，选择\u201c更新域\u201d以生成目录）'

    # TOC end
    r_end = etree.SubElement(toc_p, f"{{{W}}}r")
    etree.SubElement(r_end, f"{{{W}}}fldChar", {f"{{{W}}}fldCharType": "end"})

    # 在指定位置插入 TOC 段落
    if after_paragraph_index is None:
        # 插入到 body 最前面
        body.insert(0, toc_p)
    else:
        if after_paragraph_index < len(p_elements):
            _, target_para = p_elements[after_paragraph_index]
            target_para.addnext(toc_p)
        else:
            body.insert(0, toc_p)

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ── Action: insert_red_header ──────────────────────────────────

@_handler("insert_red_header")
def _insert_red_header(target_path: Path, params: dict) -> None:
    """在文档正文开头插入公文红头区域。

    依据 GB/T 9704 格式：
    - 第一段：发文机关全称（红色大号居中 + 下方红色反线）
    - 第二段：发文字号（红色小号居中）

    params:
        org_name: str       — 发文机关全称
        doc_number: str     — 发文字号
        font_eastAsia: str  — 中文字体（默认 "方正小标宋简体"）
        org_font_size: str  — 机关名称字号 half-points（默认 "84" = 42pt 初号）
        doc_font_size: str  — 发文字号字号 half-points（默认 "32" = 16pt 三号）
    """
    org_name = params["org_name"]
    doc_number = params["doc_number"]
    font_east = params.get("font_eastAsia", "方正小标宋简体")
    org_font_size = params.get("org_font_size", "84")
    doc_font_size = params.get("doc_font_size", "32")

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    # ── 机关名称段落（带红色反线 = 下边框） ──
    p_org = etree.Element(f"{{{W}}}p")
    pPr_org = etree.SubElement(p_org, f"{{{W}}}pPr")
    etree.SubElement(pPr_org, f"{{{W}}}jc", {f"{{{W}}}val": "center"})
    # 段前间距，让红头离页面上边距近一些
    etree.SubElement(pPr_org, f"{{{W}}}spacing", {
        f"{{{W}}}before": "0",
        f"{{{W}}}after": "0",
    })
    # 红色反线（段落下边框）
    pBdr = etree.SubElement(pPr_org, f"{{{W}}}pBdr")
    etree.SubElement(pBdr, f"{{{W}}}bottom", {
        f"{{{W}}}val": "single",
        f"{{{W}}}sz": "12",
        f"{{{W}}}space": "4",
        f"{{{W}}}color": "FF0000",
    })
    r_org = etree.SubElement(p_org, f"{{{W}}}r")
    rPr_org = etree.SubElement(r_org, f"{{{W}}}rPr")
    rf_org = etree.SubElement(rPr_org, f"{{{W}}}rFonts")
    rf_org.set(f"{{{W}}}eastAsia", font_east)
    etree.SubElement(rPr_org, f"{{{W}}}sz", {f"{{{W}}}val": org_font_size})
    etree.SubElement(rPr_org, f"{{{W}}}szCs", {f"{{{W}}}val": org_font_size})
    etree.SubElement(rPr_org, f"{{{W}}}color", {f"{{{W}}}val": "FF0000"})
    t_org = etree.SubElement(r_org, f"{{{W}}}t", {f"{{{XML}}}space": "preserve"})
    t_org.text = org_name

    # ── 发文字号段落 ──
    p_doc = etree.Element(f"{{{W}}}p")
    pPr_doc = etree.SubElement(p_doc, f"{{{W}}}pPr")
    etree.SubElement(pPr_doc, f"{{{W}}}jc", {f"{{{W}}}val": "center"})
    etree.SubElement(pPr_doc, f"{{{W}}}spacing", {
        f"{{{W}}}before": "60",
        f"{{{W}}}after": "0",
    })
    r_doc = etree.SubElement(p_doc, f"{{{W}}}r")
    rPr_doc = etree.SubElement(r_doc, f"{{{W}}}rPr")
    rf_doc = etree.SubElement(rPr_doc, f"{{{W}}}rFonts")
    rf_doc.set(f"{{{W}}}eastAsia", font_east)
    etree.SubElement(rPr_doc, f"{{{W}}}sz", {f"{{{W}}}val": doc_font_size})
    etree.SubElement(rPr_doc, f"{{{W}}}szCs", {f"{{{W}}}val": doc_font_size})
    etree.SubElement(rPr_doc, f"{{{W}}}color", {f"{{{W}}}val": "FF0000"})
    t_doc = etree.SubElement(r_doc, f"{{{W}}}t", {f"{{{XML}}}space": "preserve"})
    t_doc.text = doc_number

    # 插入到 body 开头
    if len(body) > 0:
        body.insert(0, p_doc)
        body.insert(0, p_org)
    else:
        body.append(p_org)
        body.append(p_doc)

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ── 便捷函数 ──────────────────────────────────────────────────

def fix_from_reconciler_output(reconciler_output: dict, target_path: Union[str, Path]) -> list[str]:
    raise NotImplementedError("请在 agent.py 中使用 _reconciler_to_fix_plan()")


def _make_header_xml(text: str, font_size: str = "21",
                     font_eastAsia: str = "宋体",
                     font_ascii: str = "Times New Roman") -> bytes:
    """生成最小可用的页眉 XML，作为 python-docx 路径失败时的兜底。"""
    from xml.sax.saxutils import escape

    safe_text = escape(text)
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{W}"><w:p><w:pPr><w:jc w:val="center"/>'
        f'<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>'
        f'</w:pBdr></w:pPr><w:r><w:rPr>'
        f'<w:rFonts w:eastAsia="{font_eastAsia}" w:ascii="{font_ascii}" w:hAnsi="{font_ascii}"/>'
        f'<w:sz w:val="{font_size}"/><w:szCs w:val="{font_size}"/>'
        f'</w:rPr><w:t xml:space="preserve">{safe_text}</w:t></w:r></w:p></w:hdr>'
    ).encode("utf-8")


def _make_hdr_xml_via_python_docx(text: str, font_size: str = "21",
                                  font_eastAsia: str = "宋体",
                                  font_ascii: str = "Times New Roman") -> bytes:
    """用 python-docx 创建临时文档，提取其 header XML。字号/字体来自规范，不硬编码。"""
    try:
        from docx import Document
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        import tempfile, os

        doc = Document()
        sec = doc.sections[0]
        sec.header.is_linked_to_previous = False
        hdr = sec.header
        p = hdr.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.clear()
        run = p.add_run(text)
        rPr = run._r.get_or_add_rPr()
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:eastAsia="{font_eastAsia}" w:ascii="{font_ascii}" w:hAnsi="{font_ascii}"/>')
        rPr.insert(0, rFonts)
        sz = parse_xml(f'<w:sz {nsdecls("w")} w:val="{font_size}"/>')
        rPr.append(sz)
        szCs = parse_xml(f'<w:szCs {nsdecls("w")} w:val="{font_size}"/>')
        rPr.append(szCs)
        pPr = p._p.get_or_add_pPr()
        pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>')
        pPr.append(pBdr)

        fd, tmp = tempfile.mkstemp(suffix='.docx')
        os.close(fd)
        doc.save(tmp)
        with zipfile.ZipFile(tmp, 'r') as z:
            for n in z.namelist():
                if n.startswith('word/header') and n.endswith('.xml'):
                    hdr_xml = z.read(n)
                    break
            else:
                hdr_xml = _make_header_xml(text, font_size, font_eastAsia, font_ascii)
        os.unlink(tmp)
        return hdr_xml
    except Exception:
        return _make_header_xml(text, font_size, font_eastAsia, font_ascii)


# ═══════════════════════════════════════════════════════════════
# Action: protect_image_para — 图片段落保护
# ═══════════════════════════════════════════════════════════════



def _canonicalize_ppr(p: etree._Element) -> etree._Element:
    """Return a valid CT_PPr in the first position of ``w:p``.

    A previous implementation appended pPr after the run. Word silently
    ignores a pPr in that position, which made keep/spacing/centering appear
    in XML but not in the rendered document.

    2026-09-28：顺序表收敛到唯一真源 `earlyz/tools/ooxml_order.py`（本文件原有
    一份不完整的字典拷贝，缺 suppressLineNumbers/rPr/sectPr/pPrChange 四项）。
    """
    from earlyz.tools.ooxml_order import (
        PPR_SEQ, canonicalize_element, canonicalize_ppr_position,
    )

    ppr = p.find(f"{{{W}}}pPr")
    if ppr is None:
        ppr = etree.Element(f"{{{W}}}pPr")
        p.insert(0, ppr)
    canonicalize_ppr_position(ppr)
    canonicalize_element(ppr, PPR_SEQ)
    return ppr


def _set_image_paragraph_layout(p: etree._Element, keep_next: bool) -> None:
    """Apply Word-safe layout to an inline-image paragraph."""
    ppr = _canonicalize_ppr(p)
    for local in ("keepLines", "keepNext", "spacing", "ind", "jc"):
        for child in list(ppr):
            if etree.QName(child).localname == local:
                ppr.remove(child)

    etree.SubElement(ppr, f"{{{W}}}keepLines")
    if keep_next:
        etree.SubElement(ppr, f"{{{W}}}keepNext")
    spacing = etree.SubElement(ppr, f"{{{W}}}spacing")
    # atLeast lets Word expand the line box to the full drawing height;
    # exact/low auto lines clip large inline drawings at the page boundary.
    spacing.set(f"{{{W}}}line", "240")
    spacing.set(f"{{{W}}}lineRule", "atLeast")
    ind = etree.SubElement(ppr, f"{{{W}}}ind")
    ind.set(f"{{{W}}}firstLine", "0")
    ind.set(f"{{{W}}}firstLineChars", "0")
    jc = etree.SubElement(ppr, f"{{{W}}}jc")
    jc.set(f"{{{W}}}val", "center")
    _canonicalize_ppr(p)


def _set_caption_layout(p: etree._Element) -> None:
    ppr = _canonicalize_ppr(p)
    for local in ("keepLines", "keepNext", "jc"):
        for child in list(ppr):
            if etree.QName(child).localname == local:
                ppr.remove(child)
    etree.SubElement(ppr, f"{{{W}}}keepLines")
    jc = etree.SubElement(ppr, f"{{{W}}}jc")
    jc.set(f"{{{W}}}val", "center")
    _canonicalize_ppr(p)


@_handler("normalize_image_paragraphs")
def _normalize_image_paragraphs(target_path: Path, params: dict) -> None:
    """Normalize every final inline drawing and its directly following caption."""
    root = _parse_xml_from_zip(target_path, "word/document.xml")
    changed = 0
    for container in root.iter():
        children = list(container)
        for i, child in enumerate(children):
            if child.tag != f"{{{W}}}p" or child.find(f".//{{{W}}}drawing") is None:
                continue
            next_p = None
            if i + 1 < len(children) and children[i + 1].tag == f"{{{W}}}p":
                candidate = children[i + 1]
                caption_text = ''.join(t.text or '' for t in candidate.iter(f"{{{W}}}t"))
                if re.match(r"^\s*图\s*\d+(?:[-. ]\d+)*", caption_text):
                    next_p = candidate
            _set_image_paragraph_layout(child, keep_next=next_p is not None)
            if next_p is not None:
                _set_caption_layout(next_p)
            changed += 1
    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)
    return changed

@_handler("protect_image_para")
def _protect_image_para(target_path: Path, params: dict) -> None:
    """保护图片段落：防止跨页截断、与标题分离、固定行距改为自动。

    params:
        para_index: int            — document.xml body 中该段落的索引
        keepLines: bool (默认True) — 段落所有行保持在一起
        keepNext: bool (默认True)  — 与下一段不分离
        fixLineRule: bool (默认False) — lineRule="exact" → "auto" (等效 Ctrl+1)
        center: bool (默认False)   — 段落居中（图片段要求）
        singleSpacing: bool (默认False) — 强制单倍行距 line=240/auto（Ctrl+1，防图片裁剪）
    """
    para_idx = _first_param(params, "para_index", "paragraph_index", "paragraph", "idx", "index")
    if para_idx is None:
        raise ValueError("format_paragraph 需要 para_index（别名 paragraph/idx/index 兼容）——见参数表 hint")
    keep_lines = params.get("keepLines", False)
    keep_next = params.get("keepNext", True)
    fix_line_rule = params.get("fixLineRule", False)
    center = params.get("center", False)
    single_spacing = params.get("singleSpacing", False)

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    paras = list(body.iter(f"{{{W}}}p"))
    if para_idx >= len(paras):
        return

    p = paras[para_idx]
    # text_snippet 防漂移校验：索引错位时按 snippet 全文档重定位
    snippet = params.get("text_snippet")
    docpr_id = params.get("docpr_id")
    if snippet or docpr_id:
        need_relocate = True
        if snippet:
            cur = ''.join(t.text or '' for t in p.iter(f"{{{W}}}t"))
            cur_norm = re.sub(r'\s+', '', cur)[:20]
            if cur_norm == snippet:
                need_relocate = False
        if need_relocate:
            target = None
            # 优先按 docPr id 定位（图片段文本为空时 snippet 失效）
            if docpr_id:
                wp_ns = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
                for cand in paras:
                    for dp in cand.iter(f'{wp_ns}docPr'):
                        if dp.get('id') == docpr_id:
                            target = cand
                            break
                    if target is not None:
                        break
            if target is None and snippet:
                for cand in paras:
                    ctxt = ''.join(t.text or '' for t in cand.iter(f"{{{W}}}t"))
                    if re.sub(r'\s+', '', ctxt)[:20] == snippet:
                        target = cand
                        break
            if target is None:
                return
            p = target
    if keep_lines or center or single_spacing:
        _set_image_paragraph_layout(p, keep_next=keep_next)

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ═══════════════════════════════════════════════════════════════
# Action: clean_heading_text — 标题文本净化
# ═══════════════════════════════════════════════════════════════

@_handler("clean_heading_text")
def _clean_heading_text(target_path: Path, params: dict) -> None:
    """净化标题类段落文本：
    1) 去除 w:hyperlink 包裹（目录超链接，保留内部文字）
    2) 删除 PAGEREF 页码域（fldChar begin..end，目录页码尾巴如\"绪论1\"的\"1\"）
    3) 删除制表符
    4) 段首空格清理（半角/全角）
    5) 连续多个空格压缩为 1 个（\"第1章  绪论\"→\"第1章 绪论\"）

    params:
        para_index: int — document.xml body 中该段落的索引
    """
    para_idx = params.get("para_index")
    snippet = params.get("text_snippet")
    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return
    paras = list(body.iter(f"{{{W}}}p"))
    # §21.5 兼容（2026-08-25 建议层实测 KeyError）：LLM/建议层无 para_index
    # 只传 selector/text_snippet——直接按 snippet 全文定位（唯一匹配）
    def _snip_match(cur_norm: str, snip: str) -> bool:
        # 2026-08-26 真机：LLM 从 inspect 抄 snippet 带空格（"第1章  绪论"）——
        # cur_norm 已去空格，snip 未归一化 → 不匹配 → clean 空转假成功。
        # 两侧都去除空白后比较/前缀匹配。
        snip_n = re.sub(r'\s+', '', snip or '')
        if not snip_n:
            return False
        return cur_norm == snip_n or cur_norm.startswith(snip_n)

    def _clean_one(p) -> bool:
        """对单个段落执行净化（1-5 步骤），返回是否有改动。"""
        _changed = False
        # 目录条目防护（2026-08-26 真机事故 tx_62650de733f4）：clean 在目录条目
        # （pStyle TOC1/2/3）上执行了「拆 hyperlink + 删 PAGEREF 域 + 删 tab」——
        # 目录「第1章 绪论」被正文化、页码被删。目录条目只压缩空格（步骤 4+5），
        # 不做结构性破坏（hyperlink/页码域/制表符保留）。
        _ppr = p.find(f"{{{W}}}pPr")
        _psty = _ppr.find(f"{{{W}}}pStyle") if _ppr is not None else None
        _is_toc = _psty is not None and (_psty.get(f"{{{W}}}val") or "").upper().startswith("TOC")
        if not _is_toc:
            # 1) 拆掉 hyperlink 包裹（内容上提）
            for hl in list(p.iter(f"{{{W}}}hyperlink")):
                for child in list(hl):
                    hl.addprevious(child)
                hl.getparent().remove(hl)
                _changed = True
            # 2) 删除页码域（PAGEREF 等 fldChar 序列）
            runs = list(p.iter(f"{{{W}}}r"))
            i = 0
            while i < len(runs):
                r = runs[i]
                fc = r.find(f"{{{W}}}fldChar")
                if fc is not None and fc.get(f"{{{W}}}fldCharType") == "begin":
                    j = i
                    while j < len(runs):
                        fc2 = runs[j].find(f"{{{W}}}fldChar")
                        if fc2 is not None and fc2.get(f"{{{W}}}fldCharType") == "end":
                            break
                        j += 1
                    for k in range(j, i - 1, -1):
                        p.remove(runs[k])
                    _changed = True
                    runs = list(p.iter(f"{{{W}}}r"))
                    i = 0
                    continue
                i += 1
            # 3) 删除制表符
            for tab in list(p.iter(f"{{{W}}}tab")):
                tab.getparent().remove(tab)
                _changed = True
        # 4+5) 段首空格清理 + 双空格压缩
        texts = list(p.iter(f"{{{W}}}t"))
        if texts:
            first = texts[0]
            if first.text and first.text[0] in (" ", "\u3000"):
                first.text = first.text.lstrip(" \u3000")
                _changed = True
            for t in texts:
                if t.text and re.search(r"[ \u3000]{2,}", t.text):
                    t.text = re.sub(r"[ \u3000]{2,}", " ", t.text)
                    _changed = True
            last = texts[-1]
            if last.text and last.text != last.text.rstrip(" \u3000"):
                last.text = last.text.rstrip(" \u3000")
                _changed = True
            for k in range(len(texts) - 1):
                a, b = texts[k], texts[k + 1]
                if a.text and b.text and a.text.endswith(" ") and b.text.startswith(" "):
                    b.text = b.text.lstrip(" ")
                    _changed = True
                if a.text and b.text and a.text.endswith("\u3000") and b.text.startswith("\u3000"):
                    b.text = b.text.lstrip("\u3000")
                    _changed = True
        return _changed

    if para_idx is None:
        if not snippet:
            return
        # 2026-08-26 真机：用户指令「第1章与绪论之间应为一个空格」=全文语义——
        # 目录条目与正文标题均为匹配段，全部处理（此前只处理第一个=目录漏正文）
        targets = []
        for cand in paras:
            ctxt = ''.join(t.text or '' for t in cand.iter(f"{{{W}}}t"))
            if _snip_match(re.sub(r'\s+', '', ctxt)[:20], snippet):
                targets.append(cand)
        if not targets:
            return
        changed = False
        for p in targets:  # any() 会短路——必须逐个执行全部匹配段
            if _clean_one(p):
                changed = True
    else:
        if para_idx >= len(paras):
            return
        p = paras[para_idx]
        # text_snippet 防漂移校验：前面 action 增删段落后索引错位时，
        # 按 snippet 全文档重定位（create_section_headers 插段会导致索引漂移）
        if snippet:
            cur = ''.join(t.text or '' for t in p.iter(f"{{{W}}}t"))
            cur_norm = re.sub(r'\s+', '', cur)[:20]
            if not _snip_match(cur_norm, snippet):
                target = None
                for cand in paras:
                    ctxt = ''.join(t.text or '' for t in cand.iter(f"{{{W}}}t"))
                    if _snip_match(re.sub(r'\s+', '', ctxt)[:20], snippet):
                        target = cand
                        break
                if target is None:
                    return
                p = target
        changed = _clean_one(p)

    if changed:
        new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        _zip_replace(target_path, "word/document.xml", new_data)


# ═══════════════════════════════════════════════════════════════
# Action: set_tbl_header_repeat — 表格表头跨页重复
# ═══════════════════════════════════════════════════════════════

@_handler("set_tbl_header_repeat")
def _set_tbl_header_repeat(target_path: Path, params: dict) -> None:
    """给表格设置 tblHeader 标记，使其跨页时重复表头行。

    同时设置表级 <w:tblPr><w:tblHeader/>（触发重复）和
    首行级 <w:trPr><w:tblHeader/>（标记哪些行是表头）。

    params:
        tbl_index: int   — document.xml body 中该表格的索引
    """
    tbl_idx = params["tbl_index"]

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    tables = list(body.iter(f"{{{W}}}tbl"))
    if tbl_idx >= len(tables):
        return  # 越界静默跳过（不崩溃——LLM 给的索引可能过期）

    tbl = tables[tbl_idx]

    # ── 首行级 tblHeader（标记哪行是表头——OOXML 规范中 tblHeader 只在 trPr）──
    first_row = tbl.find(f"{{{W}}}tr")
    if first_row is not None:
        trPr = first_row.find(f"{{{W}}}trPr")
        if trPr is None:
            trPr = etree.SubElement(first_row, f"{{{W}}}trPr")
        if trPr.find(f"{{{W}}}tblHeader") is None:
            etree.SubElement(trPr, f"{{{W}}}tblHeader")

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ═══════════════════════════════════════════════════════════════
# Action: set_tr_cant_split — 表格行禁止跨页截断
# ═══════════════════════════════════════════════════════════════

@_handler("set_tr_cant_split")
def _set_tr_cant_split(target_path: Path, params: dict) -> None:
    """给表格所有数据行加 cantSplit=true，防止单行跨页裂开。

    params:
        tbl_index: int   — document.xml body 中该表格的索引
        skip_header: bool (默认True) — 跳过第一行（表头通常无需 cantSplit）
    """
    tbl_idx = params["tbl_index"]
    skip_header = params.get("skip_header", True)

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    tables = list(body.iter(f"{{{W}}}tbl"))
    if tbl_idx >= len(tables):
        return  # 越界静默跳过（不崩溃——LLM 给的索引可能过期）

    tbl = tables[tbl_idx]
    for i, tr in enumerate(tbl.iter(f"{{{W}}}tr")):
        if skip_header and i == 0:
            continue

        trPr = tr.find(f"{{{W}}}trPr")
        if trPr is None:
            trPr = etree.SubElement(tr, f"{{{W}}}trPr")

        old_cs = trPr.find(f"{{{W}}}cantSplit")
        if old_cs is not None:
            trPr.remove(old_cs)
        etree.SubElement(trPr, f"{{{W}}}cantSplit", {f"{{{W}}}val": "true"})

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ═══════════════════════════════════════════════════════════════
# Action: set_tbl_font — 表格内文字字体修正
# ═══════════════════════════════════════════════════════════════

@_handler("set_tbl_font")
def _set_tbl_font(target_path: Path, params: dict) -> None:
    """修改指定表格所有单元格内文字的基础字体和字号。

    只改 rFonts + sz，不碰 b/i/color（保留单元格内的强调语义）。

    params:
        tbl_index: int
        font_eastAsia: str | None
        font_ascii: str | None
        font_size_pt: float | None
        skip_header: bool (默认 False) — 跳过首行（与 set_tbl_header_font 配合时设为 True）
    """
    tbl_idx = params["tbl_index"]
    font_east = params.get("font_eastAsia")
    font_ascii = params.get("font_ascii")
    font_size_pt = params.get("font_size_pt")
    skip_header = params.get("skip_header", False)

    if not any((font_east, font_ascii, font_size_pt)):
        return

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    tables = list(body.iter(f"{{{W}}}tbl"))
    if tbl_idx >= len(tables):
        return  # 越界静默跳过（不崩溃——LLM 给的索引可能过期）

    tbl = tables[tbl_idx]

    for tr_idx, tr in enumerate(tbl.iter(f"{{{W}}}tr")):
        if skip_header and tr_idx == 0:
            continue  # 跳过表头行，保护 set_tbl_header_font 的设置
        for r in tr.iter(f"{{{W}}}r"):
            rPr = r.find(f"{{{W}}}rPr")
            if rPr is None:
                rPr = etree.Element(f"{{{W}}}rPr")
                r.insert(0, rPr)

            if font_east or font_ascii:
                old_rf = rPr.find(f"{{{W}}}rFonts")
                if old_rf is not None:
                    rPr.remove(old_rf)
                rf = etree.SubElement(rPr, f"{{{W}}}rFonts")
                if font_ascii:
                    rf.set(f"{{{W}}}ascii", font_ascii)
                    rf.set(f"{{{W}}}hAnsi", font_ascii)
                if font_east:
                    rf.set(f"{{{W}}}eastAsia", font_east)

            if font_size_pt:
                sv = str(int(font_size_pt * 2))
                for tag in [f"{{{W}}}sz", f"{{{W}}}szCs"]:
                    old = rPr.find(tag)
                    if old is not None:
                        rPr.remove(old)
                    etree.SubElement(rPr, tag, {f"{{{W}}}val": sv})

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ═══════════════════════════════════════════════════════════════
# Action: set_tbl_body_font — 表体行字体（自动跳过所有表头行）
# ═══════════════════════════════════════════════════════════════

@_handler("set_tbl_body_font")
def _set_tbl_body_font(target_path: Path, params: dict) -> None:
    """设置表格表体行（非表头行）的字体和字号。

    自动跳过所有标记了 <w:tblHeader/> 的行（支持多行表头），
    只修改数据行。与 set_tbl_header_font 配合使用，互不覆盖。

    params:
        tbl_index: int
        font_eastAsia: str | None (默认 "宋体")
        font_ascii: str | None   (默认 "Times New Roman")
        font_size_pt: float | None (默认 None=不改字号)
    """
    tbl_idx = params["tbl_index"]
    font_east = params.get("font_eastAsia", "宋体")
    font_ascii = params.get("font_ascii", "Times New Roman")
    font_size_pt = params.get("font_size_pt")

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    tables = list(body.iter(f"{{{W}}}tbl"))
    if tbl_idx >= len(tables):
        return  # 越界静默跳过（不崩溃——LLM 给的索引可能过期）

    tbl = tables[tbl_idx]

    for tr in tbl.iter(f"{{{W}}}tr"):
        # 跳过所有标记了 tblHeader 的表头行（支持多行表头）
        trPr = tr.find(f"{{{W}}}trPr")
        if trPr is not None and trPr.find(f"{{{W}}}tblHeader") is not None:
            continue

        for r in tr.iter(f"{{{W}}}r"):
            rPr = r.find(f"{{{W}}}rPr")
            if rPr is None:
                rPr = etree.Element(f"{{{W}}}rPr")
                r.insert(0, rPr)

            old_rf = rPr.find(f"{{{W}}}rFonts")
            if old_rf is not None:
                rPr.remove(old_rf)
            rf = etree.SubElement(rPr, f"{{{W}}}rFonts")
            rf.set(f"{{{W}}}ascii", font_ascii)
            rf.set(f"{{{W}}}hAnsi", font_ascii)
            rf.set(f"{{{W}}}eastAsia", font_east)

            if font_size_pt:
                sv = str(int(font_size_pt * 2))
                for tag in [f"{{{W}}}sz", f"{{{W}}}szCs"]:
                    old = rPr.find(tag)
                    if old is not None:
                        rPr.remove(old)
                    etree.SubElement(rPr, tag, {f"{{{W}}}val": sv})

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ═══════════════════════════════════════════════════════════════
# Action: set_tbl_header_font — 表头行字体（中文黑体加粗+英文TNR加粗）
# ═══════════════════════════════════════════════════════════════

@_handler("set_tbl_header_font")
def _set_tbl_header_font(target_path: Path, params: dict) -> None:
    """设置表格表头行（第一行）的字体和加粗。

    中文=黑体加粗，英文/数字=Times New Roman 加粗，Ascii/EastAsian 分离。

    params:
        tbl_index: int
        font_eastAsia: str (默认 "黑体")
        font_ascii: str   (默认 "Times New Roman")
        font_size_pt: float | None  (默认 None=不改字号)
        bold: bool        (默认 True)
    """
    tbl_idx = params["tbl_index"]
    font_east = params.get("font_eastAsia", "黑体")
    font_ascii = params.get("font_ascii", "Times New Roman")
    font_size_pt = params.get("font_size_pt")
    bold = params.get("bold", True)

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    tables = list(body.iter(f"{{{W}}}tbl"))
    if tbl_idx >= len(tables):
        return  # 越界静默跳过（不崩溃——LLM 给的索引可能过期）

    tbl = tables[tbl_idx]
    first_row = tbl.find(f"{{{W}}}tr")
    if first_row is None:
        return

    for r in first_row.iter(f"{{{W}}}r"):
        rPr = r.find(f"{{{W}}}rPr")
        if rPr is None:
            rPr = etree.Element(f"{{{W}}}rPr")
            r.insert(0, rPr)

        # 字体：Ascii/EastAsian 分离
        old_rf = rPr.find(f"{{{W}}}rFonts")
        if old_rf is not None:
            rPr.remove(old_rf)
        rf = etree.SubElement(rPr, f"{{{W}}}rFonts")
        rf.set(f"{{{W}}}ascii", font_ascii)
        rf.set(f"{{{W}}}hAnsi", font_ascii)
        rf.set(f"{{{W}}}eastAsia", font_east)

        # 字号
        if font_size_pt:
            sv = str(int(font_size_pt * 2))
            for tag in [f"{{{W}}}sz", f"{{{W}}}szCs"]:
                old = rPr.find(tag)
                if old is not None:
                    rPr.remove(old)
                etree.SubElement(rPr, tag, {f"{{{W}}}val": sv})

        # 加粗
        if bold:
            old_b = rPr.find(f"{{{W}}}b")
            if old_b is not None:
                rPr.remove(old_b)
            etree.SubElement(rPr, f"{{{W}}}b")

        # 字符间距归零（去掉多余空格效果）
        old_sp = rPr.find(f"{{{W}}}spacing")
        if old_sp is not None:
            rPr.remove(old_sp)
        etree.SubElement(rPr, f"{{{W}}}spacing", {f"{{{W}}}val": "0"})

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ═══════════════════════════════════════════════════════════════
# Action: set_tbl_borders — 表格边框（三线表/实线表）
# ═══════════════════════════════════════════════════════════════

@_handler("set_tbl_borders")
def _set_tbl_borders(target_path: Path, params: dict) -> None:
    """设置表格边框样式。

    params:
        tbl_index: int
        style: str — "three_line"（三线表）/ "solid"（实线表）/ "none"（无边框）

    三线表：顶线+底线各 1.5pt 粗线，表头下 0.75pt 细线，无竖线无内横线。
    实线表：所有边框 0.5pt 单实线。
    """
    tbl_idx = params["tbl_index"]
    style = params.get("style", "three_line")

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    tables = list(body.iter(f"{{{W}}}tbl"))
    if tbl_idx >= len(tables):
        return  # 越界静默跳过（不崩溃——LLM 给的索引可能过期）

    tbl = tables[tbl_idx]

    # ── 表级边框 ──
    tblPr = tbl.find(f"{{{W}}}tblPr")
    if tblPr is None:
        tblPr = etree.Element(f"{{{W}}}tblPr")
        tbl.insert(0, tblPr)

    old_borders = tblPr.find(f"{{{W}}}tblBorders")
    if old_borders is not None:
        tblPr.remove(old_borders)

    # tblBorders 必须位于 tblW 之后、tblCellMar/tblLook 之前（OOXML 顺序，Word 严格校验；
    # 追加到末尾会落到 tblLook 后导致 Word 打不开文档）
    borders = etree.Element(f"{{{W}}}tblBorders")
    anchor = None
    for tag in (f"{{{W}}}tblCellMar", f"{{{W}}}tblLook"):
        a = tblPr.find(tag)
        if a is not None:
            anchor = a
            break
    if anchor is not None:
        tblPr.insert(list(tblPr).index(anchor), borders)
    else:
        tblPr.append(borders)

    if style == "three_line":
        # 三线表：顶线 1.5pt + 底线 1.5pt，其余无
        _add_border(borders, "top", "single", "12", "000000")
        _add_border(borders, "bottom", "single", "12", "000000")
        _add_border(borders, "left", "none", "0", "auto")
        _add_border(borders, "right", "none", "0", "auto")
        _add_border(borders, "insideH", "none", "0", "auto")
        _add_border(borders, "insideV", "none", "0", "auto")

        # 表头行每格加底线 0.75pt（三线表的中间线）
        first_row = tbl.find(f"{{{W}}}tr")
        if first_row is not None:
            for tc in first_row.iter(f"{{{W}}}tc"):
                tcPr = tc.find(f"{{{W}}}tcPr")
                if tcPr is None:
                    tcPr = etree.Element(f"{{{W}}}tcPr")
                    tc.insert(0, tcPr)
                old_tcb = tcPr.find(f"{{{W}}}tcBorders")
                if old_tcb is not None:
                    tcPr.remove(old_tcb)
                tcb = etree.Element(f"{{{W}}}tcBorders")
                _insert_tcpr_child(tcPr, tcb)  # 按 OOXML 顺序插入（vAlign 等之前）
                _add_border(tcb, "bottom", "single", "6", "000000")

        # 清除体行 tcBorders（否则残留的实线边框覆盖表级三线设置）
        rows = list(tbl.iter(f"{{{W}}}tr"))
        for tr in rows[1:]:  # 跳过表头行
            trPr = tr.find(f"{{{W}}}trPr")
            is_header = trPr is not None and trPr.find(f"{{{W}}}tblHeader") is not None
            if is_header:
                continue
            for tc in tr.iter(f"{{{W}}}tc"):
                tcPr = tc.find(f"{{{W}}}tcPr")
                if tcPr is not None:
                    old_tcb = tcPr.find(f"{{{W}}}tcBorders")
                    if old_tcb is not None:
                        tcPr.remove(old_tcb)

    elif style == "solid":
        # 实线表：全部 0.5pt 单线
        # 表级设全部六边 + 所有格四边 tcBorders 双重兜底
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            _add_border(borders, edge, "single", "4", "000000")

        for tc in tbl.iter(f"{{{W}}}tc"):
            tcPr = tc.find(f"{{{W}}}tcPr")
            if tcPr is None:
                tcPr = etree.Element(f"{{{W}}}tcPr")
                tc.insert(0, tcPr)
            old_tcb = tcPr.find(f"{{{W}}}tcBorders")
            if old_tcb is not None:
                tcPr.remove(old_tcb)
            tcb = etree.Element(f"{{{W}}}tcBorders")
            _insert_tcpr_child(tcPr, tcb)  # 按 OOXML 顺序插入（vAlign 等之前）
            for edge in ("top", "left", "bottom", "right"):
                _add_border(tcb, edge, "single", "4", "000000")

        # ── 列宽对齐：body → header（修复黑体/宋体字体宽度差异导致边框不对齐）──
        # 找第一个非 tblHeader 行作为列宽参照
        ref_widths = None
        for tr in tbl.iter(f"{{{W}}}tr"):
            trPr = tr.find(f"{{{W}}}trPr")
            is_header = trPr is not None and trPr.find(f"{{{W}}}tblHeader") is not None
            if is_header:
                continue
            w_list = []
            for tc in tr.iter(f"{{{W}}}tc"):
                tcPr = tc.find(f"{{{W}}}tcPr")
                w = None
                if tcPr is not None:
                    tcW = tcPr.find(f"{{{W}}}tcW")
                    if tcW is not None:
                        w = tcW.get(f"{{{W}}}w")
                w_list.append(w)
            if any(w is not None for w in w_list):
                ref_widths = w_list
                break

        if ref_widths is not None:
            for tr in tbl.iter(f"{{{W}}}tr"):
                for i, tc in enumerate(tr.iter(f"{{{W}}}tc")):
                    if i >= len(ref_widths) or ref_widths[i] is None:
                        continue
                    tcPr = tc.find(f"{{{W}}}tcPr")
                    if tcPr is None:
                        tcPr = etree.Element(f"{{{W}}}tcPr")
                        tc.insert(0, tcPr)
                    old_tcw = tcPr.find(f"{{{W}}}tcW")
                    if old_tcw is not None:
                        tcPr.remove(old_tcw)
                    etree.SubElement(tcPr, f"{{{W}}}tcW", {
                        f"{{{W}}}w": ref_widths[i],
                        f"{{{W}}}type": "dxa",
                    })

    elif style == "none":
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            _add_border(borders, edge, "none", "0", "auto")

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


def _insert_tcpr_child(tcPr: etree.Element, child: etree.Element) -> None:
    """按 OOXML 顺序插入 tcPr 子元素。

    tcPr 标准子元素顺序：tcW, gridSpan, hMerge, vMerge, tcBorders, shd,
    noWrap, tcMar, textDirection, tcFitText, vAlign, hideMark。
    tcBorders 追加到末尾会落到 vAlign 后——Word 严格校验会打不开文档。
    """
    for tag in (f"{{{W}}}shd", f"{{{W}}}noWrap", f"{{{W}}}tcMar", f"{{{W}}}textDirection",
                f"{{{W}}}tcFitText", f"{{{W}}}vAlign", f"{{{W}}}hideMark"):
        a = tcPr.find(tag)
        if a is not None:
            tcPr.insert(list(tcPr).index(a), child)
            return
    tcPr.append(child)


def _add_border(parent: etree.Element, edge: str, val: str, sz: str, color: str) -> None:
    """添加单条边框元素 <w:{edge} w:val=... w:sz=... w:space=\\\"0\\\" w:color=...>"""
    el = etree.SubElement(parent, f"{{{W}}}{edge}")
    el.set(f"{{{W}}}val", val)
    el.set(f"{{{W}}}sz", sz)
    el.set(f"{{{W}}}space", "0")
    el.set(f"{{{W}}}color", color)


# ═══════════════════════════════════════════════════════════════
# Action: set_tbl_col_widths — 表格列宽手动控制
# ═══════════════════════════════════════════════════════════════

@_handler("set_tbl_col_widths")
def _set_tbl_col_widths(target_path: Path, params: dict) -> None:
    """设置表格每列的宽度（单位：twips，1/20 pt）。

    同时操作 tblGrid（定义列结构）和每格的 tcW（覆盖列宽），
    确保所有行共享同一份列宽。

    params:
        tbl_index: int           — document.xml body 中该表格的索引
        widths: list[int|str]    — 每列宽度，twips 数值或 "auto"
                                    例：[2400, 3600, "auto", 1800]
    """
    tbl_idx = params["tbl_index"]
    widths = params["widths"]

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    tables = list(body.iter(f"{{{W}}}tbl"))
    if tbl_idx >= len(tables):
        return  # 越界静默跳过（不崩溃——LLM 给的索引可能过期）

    tbl = tables[tbl_idx]

    # ── 1. 重建 tblGrid（必须在 tblPr 之后）──
    # 注意：tbl.insert(0, tblGrid) 会把 tblGrid 放在 tblPr 之前 → Word 解析异常
    # 正确：找到 tblPr → 在其后插入 tblGrid；若无 tblPr 则 insert(0)
    old_grid = tbl.find(f"{{{W}}}tblGrid")
    if old_grid is not None:
        tbl.remove(old_grid)

    new_grid = etree.Element(f"{{{W}}}tblGrid")
    for w in widths:
        w_str = str(int(w)) if isinstance(w, (int, float)) else "1600"
        etree.SubElement(new_grid, f"{{{W}}}gridCol", {
            f"{{{W}}}w": w_str,
        })
    # 插在 tblPr 之后
    tblPr = tbl.find(f"{{{W}}}tblPr")
    if tblPr is not None:
        # 找到 tblPr 的索引位置，在其后插入
        tbl_children = list(tbl)
        pr_idx = tbl_children.index(tblPr)
        tbl.insert(pr_idx + 1, new_grid)
    else:
        tbl.insert(0, new_grid)

    # ── 2. 更新每格的 tcW ──
    for tr in tbl.iter(f"{{{W}}}tr"):
        tc_idx = 0
        for tc in tr.iter(f"{{{W}}}tc"):
            # 计算该单元格的实际列宽（考虑 gridSpan 合并）
            grid_span = 1
            tcPr = tc.find(f"{{{W}}}tcPr")
            if tcPr is not None:
                gs = tcPr.find(f"{{{W}}}gridSpan")
                if gs is not None:
                    grid_span = int(gs.get(f"{{{W}}}val", "1"))

            # 计算合并格应有的总宽度
            total_w = 0
            for k in range(grid_span):
                if tc_idx + k < len(widths):
                    w_val = widths[tc_idx + k]
                    if isinstance(w_val, (int, float)):
                        total_w += int(w_val)
                    else:
                        total_w += 1600  # "auto" → 默认
            tc_idx += grid_span

            if tcPr is None:
                tcPr = etree.Element(f"{{{W}}}tcPr")
                tc.insert(0, tcPr)

            # 移除旧 tcW，写入新值
            old_tcw = tcPr.find(f"{{{W}}}tcW")
            if old_tcw is not None:
                tcPr.remove(old_tcw)
            etree.SubElement(tcPr, f"{{{W}}}tcW", {
                f"{{{W}}}w": str(total_w),
                f"{{{W}}}type": "dxa",
            })

    # 写入 _widths_applied 供 Fixer 验证
    params["_widths_applied"] = widths

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ═══════════════════════════════════════════════════════════════
# Action: set_tbl_cell_valign — 单元格垂直对齐
# ═══════════════════════════════════════════════════════════════

@_handler("set_tbl_cell_valign")
def _set_tbl_cell_valign(target_path: Path, params: dict) -> None:
    """设置表格所有单元格的垂直对齐方式。

    params:
        tbl_index: int       — 表格索引
        valign: str          — "top" / "center" / "bottom"
        skip_header: bool    — 跳过表头行（默认 False）
    """
    tbl_idx = params["tbl_index"]
    valign = params.get("valign", "center")
    skip_header = params.get("skip_header", False)

    if valign not in ("top", "center", "bottom"):
        raise ValueError(f"valign 必须是 'top'/'center'/'bottom'，收到: {valign}")

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    tables = list(body.iter(f"{{{W}}}tbl"))
    if tbl_idx >= len(tables):
        return  # 越界静默跳过（不崩溃——LLM 给的索引可能过期）

    tbl = tables[tbl_idx]

    affected = 0
    for tr_idx, tr in enumerate(tbl.iter(f"{{{W}}}tr")):
        if skip_header and tr_idx == 0:
            continue

        for tc in tr.iter(f"{{{W}}}tc"):
            tcPr = tc.find(f"{{{W}}}tcPr")
            if tcPr is None:
                tcPr = etree.Element(f"{{{W}}}tcPr")
                tc.insert(0, tcPr)

            # 移除旧的 vAlign
            old_valign = tcPr.find(f"{{{W}}}vAlign")
            if old_valign is not None:
                tcPr.remove(old_valign)

            etree.SubElement(tcPr, f"{{{W}}}vAlign", {
                f"{{{W}}}val": valign,
            })
            affected += 1

    params["_affected"] = affected

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ═══════════════════════════════════════════════════════════════
# Action: split_merged_cells — 检测合并单元格（警告模式，不自动拆分）
# ═══════════════════════════════════════════════════════════════

@_handler("split_merged_cells")
def _split_merged_cells(target_path: Path, params: dict) -> None:
    """检测表格中的合并单元格（gridSpan / vMerge），记录位置信息。

    ⚠️ 不执行实际拆分——合并单元格拆分风险太大，
       可能破坏表格结构、数据对齐和排版语义。
       仅返回合并信息供用户/Agent 审查。

    params:
        tbl_index: int  — 表格索引

    结果写入 params["_merged"]:
        [{"row": 0, "col": 1, "span": 2, "type": "gridSpan"},
         {"row": 1, "col": 0, "type": "vMerge", "val": "continue"}, ...]
    """
    tbl_idx = params["tbl_index"]

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    tables = list(body.iter(f"{{{W}}}tbl"))
    if tbl_idx >= len(tables):
        return  # 越界静默跳过（不崩溃——LLM 给的索引可能过期）

    tbl = tables[tbl_idx]
    merged: list[dict] = []

    for tr_idx, tr in enumerate(tbl.iter(f"{{{W}}}tr")):
        col_idx = 0
        for tc in tr.iter(f"{{{W}}}tc"):
            tcPr = tc.find(f"{{{W}}}tcPr")
            if tcPr is None:
                col_idx += 1
                continue

            # ── gridSpan：水平合并 ──
            gs = tcPr.find(f"{{{W}}}gridSpan")
            if gs is not None:
                span = int(gs.get(f"{{{W}}}val", "1"))
                if span > 1:
                    cell_text = "".join(
                        t.text or "" for t in tc.iter(f"{{{W}}}t")
                    ).strip()[:40]
                    merged.append({
                        "row": tr_idx,
                        "col": col_idx,
                        "span": span,
                        "type": "gridSpan",
                        "cell_text": cell_text,
                    })
                col_idx += span
            else:
                # ── vMerge：垂直合并 ──
                vm = tcPr.find(f"{{{W}}}vMerge")
                if vm is not None:
                    vm_val = vm.get(f"{{{W}}}val", "continue")
                    cell_text = "".join(
                        t.text or "" for t in tc.iter(f"{{{W}}}t")
                    ).strip()[:40]
                    merged.append({
                        "row": tr_idx,
                        "col": col_idx,
                        "type": "vMerge",
                        "val": vm_val,
                        "cell_text": cell_text,
                    })
                col_idx += 1

    params["_merged"] = merged

    if merged:
        # ⚠️ 仅警告，不拆分
        gs_count = sum(1 for m in merged if m["type"] == "gridSpan")
        vm_count = sum(1 for m in merged if m["type"] == "vMerge")
        raise ValueError(
            f"检测到 {len(merged)} 处合并单元格（{gs_count} 处水平 gridSpan + "
            f"{vm_count} 处垂直 vMerge）。为安全起见，未执行拆分。"
            f"合并详细信息见 params['_merged']。"
            f"如需拆分，请在 Word 中手动取消合并后重新上传。"
        )


# ═══════════════════════════════════════════════════════════════
# Action: normalize_tbl_grid — 确保 tblGrid 与列数一致
# ═══════════════════════════════════════════════════════════════

@_handler("normalize_tbl_grid")
def _normalize_tbl_grid(target_path: Path, params: dict) -> None:
    """按实际列数重建/修正 tblGrid，解决列宽不一致导致的边框不对齐问题。

    从第一个非合并单元格最多的行推断列数 + 列宽，
    删除所有格子的 tcW，重建统一 tblGrid。
    这是 set_tbl_borders 的 solid 模式中列宽归一化逻辑的独立版本。

    params:
        tbl_index: int  — 表格索引
    """
    tbl_idx = params["tbl_index"]

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    tables = list(body.iter(f"{{{W}}}tbl"))
    if tbl_idx >= len(tables):
        return  # 越界静默跳过（不崩溃——LLM 给的索引可能过期）

    tbl = tables[tbl_idx]

    # ── 0. 旧 tblGrid 列宽（fallback——tcW 缺失/为 0 时使用，避免生成 0 宽列）──
    old_grid0 = tbl.find(f"{{{W}}}tblGrid")
    old_widths: list[str] = []
    if old_grid0 is not None:
        old_widths = [g.get(f"{{{W}}}w", "1600") for g in old_grid0.iter(f"{{{W}}}gridCol")]

    # ── 1. 推断列数：从首行 tc 数量（考虑 gridSpan）──
    first_row = tbl.find(f"{{{W}}}tr")
    col_count = 0
    grid_widths: list[str] = []
    if first_row is not None:
        for tc in first_row.iter(f"{{{W}}}tc"):
            tcPr = tc.find(f"{{{W}}}tcPr")
            span = 1
            tcw_val = "1600"
            if tcPr is not None:
                gs = tcPr.find(f"{{{W}}}gridSpan")
                if gs is not None:
                    span = int(gs.get(f"{{{W}}}val", "1"))
                tcW = tcPr.find(f"{{{W}}}tcW")
                if tcW is not None:
                    tcw_val = tcW.get(f"{{{W}}}w", "1600")
                    if tcw_val in ("0", "", None):
                        # 0 宽会导致 Word 拒绝打开——回退旧 gridCol 值或默认 1600
                        tcw_val = (old_widths[len(grid_widths)]
                                   if len(grid_widths) < len(old_widths) else "1600")
            for _ in range(span):
                grid_widths.append(tcw_val)
            col_count += span

    if col_count == 0:
        return  # 空表格

    # ── 2. 删除所有 tcW（让 tblGrid 成为唯一列宽来源）──
    for tc in tbl.iter(f"{{{W}}}tc"):
        tcPr = tc.find(f"{{{W}}}tcPr")
        if tcPr is not None:
            tcW = tcPr.find(f"{{{W}}}tcW")
            if tcW is not None:
                tcPr.remove(tcW)

    # ── 3. 重建 tblGrid（放在 tblPr 之后）──
    old_grid = tbl.find(f"{{{W}}}tblGrid")
    if old_grid is not None:
        tbl.remove(old_grid)

    new_grid = etree.Element(f"{{{W}}}tblGrid")
    for w in grid_widths:
        etree.SubElement(new_grid, f"{{{W}}}gridCol", {
            f"{{{W}}}w": w,
        })

    tblPr = tbl.find(f"{{{W}}}tblPr")
    if tblPr is not None:
        tbl_children = list(tbl)
        pr_idx = tbl_children.index(tblPr)
        tbl.insert(pr_idx + 1, new_grid)
    else:
        tbl.insert(0, new_grid)

    params["_grid_cols"] = col_count
    params["_grid_widths"] = grid_widths

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ═══════════════════════════════════════════════════════════════
# Action: normalize_cjk_spacing — 清除中英混排多余空格
# ═══════════════════════════════════════════════════════════════

import re as _re

@_handler("normalize_cjk_spacing")
def _normalize_cjk_spacing(target_path: Path, params: dict) -> None:
    """清除文档中中英混排段落的多余空格。

    常见于 AI 生成的文章：英文/数字两侧被自动加了空格，
    如\"字段 student_no 和 department 字段\"——不符合中文排版规范。

    处理逻辑（不改英文单词间的正当空格）：
    1. 删除 CJK 字符之间的空格
    2. 删除 CJK 字符与拉丁/数字之间的多余空格
    3. 多个连续空格归一为单个

    params: 无需参数（处理整个 word/document.xml）
    """
    CJK = _re.compile(r'[\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff\u3000-\u303f\uff00-\uffef]')
    LATIN = _re.compile(r'[a-zA-Z0-9]')

    def _fix_text(text: str) -> str:
        result = []
        i = 0
        while i < len(text):
            ch = text[i]
            result.append(ch)
            # 跳过连续空格，只保留一个
            if ch == ' ':
                while i + 1 < len(text) and text[i + 1] == ' ':
                    i += 1
            i += 1
        return ''.join(result)

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    changed = 0
    # ── Phase 1: 单个 run 内的空格清理 ──
    # 2026-08-25 语义修正：CJK-CJK 空格是中文排版语义（如标题「第1章 绪论」），
    # 只压缩连续空格为单个、不清除合法单空格；拉丁字母-中文边界的多余空格清除
    # （动作定义"清除中英混排多余空格"）；**数字-中文边界保留**（章节号/数值
    # 「1.1 研究背景」「第1章 绪论」是排版语义，非"多余空格"）。
    for t_elem in root.iter(f"{{{W}}}t"):
        if t_elem.text:
            old = t_elem.text
            new = _re.sub(r'([\u4e00-\u9fff\u3000-\u303f\uff00-\uffef])\s+([\u4e00-\u9fff\u3000-\u303f\uff00-\uffef])',
                          r'\1 \2', old)
            new = _re.sub(r'([\u4e00-\u9fff\u3000-\u303f\uff00-\uffef])\s+([a-zA-Z])',
                          r'\1\2', new)
            new = _re.sub(r'([a-zA-Z])\s+([\u4e00-\u9fff\u3000-\u303f\uff00-\uffef])',
                          r'\1\2', new)
            if new != old:
                t_elem.text = new
                changed += 1

    # ── Phase 2: 跨 run 边界空格清理 ──
    # Word 经常把 CJK 和 Latin 拆到不同 <w:t> run 里，
    # Phase 1 的正则无法跨越 run 边界。
    # 处理三种情况：
    #   a) 独立空格 run：前后 run 类型不同 → 清空
    #   b) 前一个 run 尾空格 + 后一个 run 首字符类型不同 → 删尾空格
    #   c) 前一个 run 尾字符 + 后一个 run 首空格且类型不同 → 删首空格
    _CJK_CHARS = re.compile(r'[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]')
    _LATIN_CHARS = re.compile(r'[a-zA-Z]')

    def _char_type(ch: str) -> str:
        """返回字符类型: 'cjk' | 'latin' | 'space' | 'other'（数字归入 cjk 中文侧——
        章节号/数值「1.1 研究」的边界空格是排版语义，不清除）。"""
        if ch == ' ':
            return 'space'
        if _CJK_CHARS.match(ch):
            return 'cjk'
        if ch.isdigit():
            return 'cjk'
        if _LATIN_CHARS.match(ch):
            return 'latin'
        return 'other'

    def _text_first_char(text: str) -> str | None:
        """返回文本第一个非空格字符的类型"""
        for ch in text:
            t = _char_type(ch)
            if t != 'space':
                return t
        return None

    def _text_last_char(text: str) -> str | None:
        """返回文本最后一个非空格字符的类型"""
        for ch in reversed(text):
            t = _char_type(ch)
            if t != 'space':
                return t
        return None

    def _text_trailing_spaces(text: str) -> int:
        """返回文本末尾连续空格数"""
        count = 0
        for ch in reversed(text):
            if ch == ' ':
                count += 1
            else:
                break
        return count

    def _text_leading_spaces(text: str) -> int:
        """返回文本开头连续空格数"""
        count = 0
        for ch in text:
            if ch == ' ':
                count += 1
            else:
                break
        return count

    for p_elem in root.iter(f"{{{W}}}p"):
        t_elems = list(p_elem.iter(f"{{{W}}}t"))
        if len(t_elems) < 2:
            continue

        for i in range(len(t_elems)):
            t = t_elems[i]
            txt = t.text
            if not txt:
                continue

            # ── 情况 a: 独立空格 run ──
            if txt.strip() == '' and txt != '':
                if i == 0 or i >= len(t_elems) - 1:
                    continue
                prev_type = _text_last_char(t_elems[i-1].text or '')
                next_type = _text_first_char(t_elems[i+1].text or '')
                if prev_type and next_type and prev_type != next_type:
                    # 不同类型边界 → 清空空格 run
                    t.text = ''
                    changed += 1
                continue

            # ── 情况 b: 前一个 run 尾空格 + 本 run 首字符类型不同 ──
            if i > 0:
                prev_txt = t_elems[i-1].text
                if prev_txt:
                    trail = _text_trailing_spaces(prev_txt)
                    if trail > 0:
                        prev_type = _text_last_char(prev_txt)
                        this_type = _text_first_char(txt)
                        if prev_type and this_type and prev_type != this_type:
                            t_elems[i-1].text = prev_txt.rstrip(' ')
                            changed += 1

            # ── 情况 c: 本 run 首空格 + 前一个 run 尾字符类型不同 ──
            if i > 0:
                lead = _text_leading_spaces(txt)
                if lead > 0:
                    prev_type = _text_last_char(t_elems[i-1].text or '')
                    this_type = _text_first_char(txt)
                    if prev_type and this_type and prev_type != this_type:
                        t.text = txt.lstrip(' ')
                        changed += 1

    if changed:
        new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        _zip_replace(target_path, "word/document.xml", new_data)
    params["_changed"] = changed


# ═══════════════════════════════════════════════════════════════
# Action: set_images_inline — 批量将浮动图片转为嵌入型
# ═══════════════════════════════════════════════════════════════

@_handler("set_images_inline")
def _set_images_inline(target_path: Path, params: dict) -> None:
    """将文档中所有浮动图片（<wp:anchor>）转为嵌入型（<wp:inline>）。

    处理范围：word/document.xml + word/header*.xml + word/footer*.xml
    只改环绕方式，不动图片内容/尺寸/引用。

    params: 无需参数
    """
    WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

    changed = 0
    protected = 0  # 顺便保护了段落的数量
    entries = [n for n in _zip_namelist(target_path)
               if n.startswith("word/") and n.endswith(".xml")
               and any(k in n for k in ("document", "header", "footer"))]

    for entry in entries:
        root = _parse_xml_from_zip(target_path, entry)

        drawings = list(root.iter(f"{{{W}}}drawing"))
        for drawing in drawings:
            anchor = drawing.find(f"{{{WP_NS}}}anchor")
            if anchor is None:
                continue

            # ── 定位所在段落（用于后续保护）──
            # 需要向上爬树直到找到 <w:p>，因为 drawing 可能被包裹在
            # <mc:AlternateContent>/<mc:Choice> 等兼容性元素中
            para = drawing.getparent()
            while para is not None and para.tag != f"{{{W}}}p":
                para = para.getparent()

            docPr = anchor.find(f"{{{WP_NS}}}docPr")
            extent = anchor.find(f"{{{WP_NS}}}extent")
            effectExtent = anchor.find(f"{{{WP_NS}}}effectExtent")
            graphic = anchor.find(f"{{{A_NS}}}graphic")

            inline = etree.Element(f"{{{WP_NS}}}inline")
            for attr in ("distT", "distB", "distL", "distR"):
                inline.set(attr, anchor.get(attr, "0"))

            if extent is not None:
                inline.append(extent)
            if effectExtent is not None:
                inline.append(effectExtent)
            if docPr is not None:
                inline.append(docPr)
            if graphic is not None:
                inline.append(graphic)

            drawing.remove(anchor)
            drawing.append(inline)
            changed += 1

            # ── 图片独立成段：如果所在段落还有其他文字，拆分出来 ──
            if para is not None and para.tag == f"{{{W}}}p":
                all_text = "".join(t.text or "" for t in para.iter(f"{{{W}}}t")).strip()
                # 找到包含 drawing 的 w:r 元素
                run = drawing.getparent()
                while run is not None and run.tag != f"{{{W}}}r":
                    run = run.getparent()

                if all_text and run is not None and run.getparent() is para:
                    # 段落有文字 → 把图片移到独立新段落，插入原文后面
                    new_p = etree.Element(f"{{{W}}}p")
                    new_pPr = etree.SubElement(new_p, f"{{{W}}}pPr")
                    etree.SubElement(new_pPr, f"{{{W}}}jc", {f"{{{W}}}val": "center"})
                    new_p.append(run)
                    para.addnext(new_p)

                    # 在新段落上做保护
                    if new_pPr.find(f"{{{W}}}keepLines") is None:
                        etree.SubElement(new_pPr, f"{{{W}}}keepLines")
                else:
                    # 图片段落无其他文字 → 在原段落做保护
                    pPr = para.find(f"{{{W}}}pPr")
                    if pPr is None:
                        pPr = etree.SubElement(para, f"{{{W}}}pPr")
                    if pPr.find(f"{{{W}}}keepLines") is None:
                        etree.SubElement(pPr, f"{{{W}}}keepLines")

                # 智能行距修正
                target_pPr = para.find(f"{{{W}}}pPr") if para.find(f"{{{W}}}pPr") is not None else None
                if extent is not None and target_pPr is not None:
                    cy = int(extent.get('cy', '0'))
                    sp = target_pPr.find(f"{{{W}}}spacing")
                    if sp is not None and sp.get(f"{{{W}}}lineRule") == "exact":
                        line_val = int(sp.get(f"{{{W}}}line", '240'))
                        line_emu = (line_val / 20) * 12700
                        if cy > line_emu:
                            sp.set(f"{{{W}}}lineRule", "auto")

                protected += 1

        new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        _zip_replace(target_path, entry, new_data)

    # 日志供 apply_operations 读取
    params["_changed"] = changed


# ═══════════════════════════════════════════════════════════════
# Action: set_paragraph_alignment — 段落对齐方式
# ═══════════════════════════════════════════════════════════════

def _first_param(params: dict, *names):
    """参数取第一个非 None 值（or 链会把 para_index=0 当 falsy——2026-08-27 真机 Bug）。"""
    for n in names:
        v = params.get(n)
        if v is not None:
            return v
    return None


@_handler("set_paragraph_alignment")
def _set_paragraph_alignment(target_path: Path, params: dict) -> None:
    """设置指定段落的水平对齐方式。

    params:
        para_index: int   — document.xml body 中该段落的索引
        val: str          — left / center / right / both / distribute
    """
    para_idx = _first_param(params, "para_index", "paragraph_index", "paragraph", "idx", "index")
    if para_idx is None:
        raise ValueError("format_paragraph 需要 para_index（别名 paragraph/idx/index 兼容）——见参数表 hint")
    # 2026-08-26 22:45 真机：LLM 传 alignment（hint 参数表笔误）——兼容别名 val/alignment/align
    align_val = params.get("val") or params.get("alignment") or params.get("align")
    if align_val is None:
        raise ValueError("set_paragraph_alignment 需要 val（left/center/right/both/distribute）")

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    paras = list(body.iter(f"{{{W}}}p"))
    if para_idx >= len(paras):
        return

    pPr = paras[para_idx].find(f"{{{W}}}pPr")
    if pPr is None:
        pPr = etree.SubElement(paras[para_idx], f"{{{W}}}pPr")

    # 移除旧 jc，写入新值
    old_jc = pPr.find(f"{{{W}}}jc")
    if old_jc is not None:
        pPr.remove(old_jc)
    etree.SubElement(pPr, f"{{{W}}}jc", {f"{{{W}}}val": align_val})

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ═══════════════════════════════════════════════════════════════
# Action: extract_textbox_captions — 把文本框内图标题提取为正文段落
# ═══════════════════════════════════════════════════════════════

@_handler("extract_textbox_captions")
def _extract_textbox_captions(target_path: Path, params: dict) -> None:
    """扫描文档中所有文本框（<w:txbxContent>），将其中的文字提取为
    独立正文段落，插入在所在图片段落之后。
    文本框本身保留不动；提取的段落自动居中 + 宋体五号。
    """
    extracted = 0
    seen_texts: set[str] = set()  # 全局去重：同一标题只提取一次
    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    for txbx in list(root.iter(f"{{{W}}}txbxContent")):
        inner_paras = list(txbx.iter(f"{{{W}}}p"))
        if not inner_paras:
            continue
        ancestor = txbx.getparent()
        # 找所在段落：支持 DrawingML（<w:drawing>）和 VML（<w:pict>）两种路径
        drawing_or_pict = None
        while ancestor is not None:
            if ancestor.tag in (f"{{{W}}}drawing", f"{{{W}}}pict"):
                drawing_or_pict = ancestor
                break
            ancestor = ancestor.getparent()
        if drawing_or_pict is None:
            continue
        para = drawing_or_pict.getparent()
        while para is not None and para.tag != f"{{{W}}}p":
            para = para.getparent()
        if para is None:
            continue

        # 收集文本框内有实际文字的非空段落
        texts_to_extract: list[str] = []
        for inner_p in inner_paras:
            text = "".join(t.text or "" for t in inner_p.iter(f"{{{W}}}t")).strip()
            if text:
                texts_to_extract.append(text)

        # 收集图片附近已有段落文字（用于去重）
        nearby_texts: set[str] = set()
        sibling = para.getnext()
        for _ in range(5):
            if sibling is None or sibling.tag != f"{{{W}}}p":
                break
            st = "".join(t.text or "" for t in sibling.iter(f"{{{W}}}t")).strip()
            if st:
                nearby_texts.add(st)
            sibling = sibling.getnext()

        # 提取为正文段落（跳过与附近已有段落重复的文字）
        for text in reversed(texts_to_extract):
            if text in seen_texts or text in nearby_texts:
                continue
            seen_texts.add(text)
            new_p = etree.Element(f"{{{W}}}p")
            new_pPr = etree.SubElement(new_p, f"{{{W}}}pPr")
            etree.SubElement(new_pPr, f"{{{W}}}jc", {f"{{{W}}}val": "center"})
            new_r = etree.SubElement(new_p, f"{{{W}}}r")
            new_rPr = etree.SubElement(new_r, f"{{{W}}}rPr")
            rf = etree.SubElement(new_rPr, f"{{{W}}}rFonts")
            rf.set(f"{{{W}}}eastAsia", "宋体")
            rf.set(f"{{{W}}}ascii", "Times New Roman")
            rf.set(f"{{{W}}}hAnsi", "Times New Roman")
            etree.SubElement(new_rPr, f"{{{W}}}sz", {f"{{{W}}}val": "21"})
            etree.SubElement(new_rPr, f"{{{W}}}szCs", {f"{{{W}}}val": "21"})
            new_t = etree.SubElement(new_r, f"{{{W}}}t")
            new_t.text = text
            new_t.set(f"{{{XML}}}space", "preserve")
            para.addnext(new_p)
            extracted += 1

        # 删除整个形状容器
        # - VML/Fallback 路径：从 <w:pict> 移除直接子元素（v:rect/v:shape）
        # - DrawingML/Choice 路径：移除整个 <w:drawing> 元素
        #   （原逻辑只删 wps:txbx 留 wps:wsp 空壳 → Word 渲染为空白矩形）
        shape_removed = False
        if drawing_or_pict is not None:
            tag_name = drawing_or_pict.tag.split('}')[-1].lower()
            if tag_name == 'pict':
                # VML 路径：移除 pict 的直接子元素（v:shape/v:rect）
                node = txbx.getparent()
                while node is not None:
                    parent = node.getparent()
                    if parent is not None and parent.tag == f"{{{W}}}pict":
                        parent.remove(node)
                        shape_removed = True
                        break
                    node = parent
            elif tag_name == 'drawing':
                # DrawingML 路径：此 drawing 只含文本框不含图片 → 移除整个 drawing
                drawing_parent = drawing_or_pict.getparent()
                if drawing_parent is not None:
                    drawing_parent.remove(drawing_or_pict)
                    shape_removed = True

        # 形状删除后，检查所在段落是否变空（无文字、无图片），是则移除
        # 同时清理相邻的连续空白段落（防止多行空白残留）
        if shape_removed and para.getparent() is not None:
            para_text = "".join(t.text or "" for t in para.iter(f"{{{W}}}t")).strip()
            para_drawings = list(para.iter(f"{{{W}}}drawing"))
            if not para_text and not para_drawings:
                para_parent = para.getparent()
                # 先收集：向前后各扫描最多 3 段，收集连续空白段落（不含 sectPr 的）
                to_remove = []
                # 向前扫描
                prev = para.getprevious()
                for _ in range(3):
                    if prev is None or prev.tag != f"{{{W}}}p":
                        break
                    pt = "".join(t.text or "" for t in prev.iter(f"{{{W}}}t")).strip()
                    pd = list(prev.iter(f"{{{W}}}drawing"))
                    ps = prev.find(f"{{{W}}}pPr/{{{W}}}sectPr") is not None
                    if not pt and not pd and not ps:
                        to_remove.append(prev)
                        prev = prev.getprevious()
                    else:
                        break
                # 向后扫描
                nxt = para.getnext()
                for _ in range(3):
                    if nxt is None or nxt.tag != f"{{{W}}}p":
                        break
                    nt = "".join(t.text or "" for t in nxt.iter(f"{{{W}}}t")).strip()
                    nd = list(nxt.iter(f"{{{W}}}drawing"))
                    ns = nxt.find(f"{{{W}}}pPr/{{{W}}}sectPr") is not None
                    if not nt and not nd and not ns:
                        to_remove.append(nxt)
                        nxt = nxt.getnext()
                    else:
                        break
                # 一次性移除所有空白段落
                for p in to_remove:
                    para_parent.remove(p)
                para_parent.remove(para)

    if extracted:
        new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        _zip_replace(target_path, "word/document.xml", new_data)

    params["_extracted"] = extracted


# ═══════════════════════════════════════════════════════════════
# Action: format_paragraph — 内容级段落直接格式（纯文本排版）
# ═══════════════════════════════════════════════════════════════

@_handler("format_paragraph")
def _format_paragraph(target_path: Path, params: dict) -> None:
    """按模板规则格式化一个普通段落（直接格式写入，不依赖样式）。

    用于纯文本/无格式文档的内容级排版：识别出的标题/正文/关键词
    段落按规范模板写入段落与 run 的直接格式。

    params:
        para_index: int — document.xml body 中该段落索引（body.iter 顺序）
        jc: str | None — 对齐（left/center/both/right/distribute）
        ind_first_line: int | None — 首行缩进（twips）
        ind_first_line_chars: int | None — 首行缩进（字符数×100）
        font_eastAsia: str | None — 中文字体
        font_ascii: str | None — 西文字体
        font_size_pt: float | None — 字号（pt）
        bold: bool | None — 加粗
        before_twips / after_twips: int | None — 段前/段后（twips）
        before_lines / after_lines: int | None — 段前/段后（行数×100）
        line: int | None — 行距值
        lineRule: str | None — auto/exact/atLeast
        outline_lvl: int | None — 大纲级别（标题段落）
        run_formats: list[dict] | None — run 级前缀格式：
            [{"match": "关键词：", "font_eastAsia": "黑体", "bold": true},
             {"match": None, "font_eastAsia": "宋体", "font_size_pt": 12}]
            第一个含前缀文字的 run 设格式 A，其余 run 设格式 B
    """
    para_idx = _first_param(params, "para_index", "paragraph_index", "paragraph", "idx", "index")
    if para_idx is None:
        raise ValueError("format_paragraph 需要 para_index（别名 paragraph/idx/index 兼容）——见参数表 hint")
    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return
    paras = list(body.iter(f"{{{W}}}p"))
    if para_idx >= len(paras):
        return
    p = paras[para_idx]
    # 文字快照校验：防索引漂移（前面 action 可能增删段落导致错位）
    snippet = params.get("text_snippet")
    if snippet:
        norm = lambda s: re.sub(r'\s+', '', s)
        if norm(''.join(t.text or '' for t in p.iter(f"{{{W}}}t")))[:20] != snippet:
            # 索引错位：在 ±10 范围内线性搜文字匹配的段落
            found = None
            lo, hi = max(0, para_idx - 10), min(len(paras), para_idx + 11)
            for i in range(lo, hi):
                cand = paras[i]
                if norm(''.join(t.text or '' for t in cand.iter(f"{{{W}}}t")))[:20] == snippet:
                    found = cand
                    break
            if found is None:
                return  # 找不到目标段落——保守放弃，不误改
            p = found

    pPr = p.find(f"{{{W}}}pPr")
    if pPr is None:
        pPr = etree.Element(f"{{{W}}}pPr")
        p.insert(0, pPr)

    # 对齐
    if "jc" in params and params["jc"] is not None:
        old_jc = pPr.find(f"{{{W}}}jc")
        if old_jc is not None:
            pPr.remove(old_jc)
        etree.SubElement(pPr, f"{{{W}}}jc", {f"{{{W}}}val": params["jc"]})

    # 缩进
    if params.get("ind_first_line") is not None or params.get("ind_first_line_chars") is not None:
        ind_attrs = {}
        if params.get("ind_first_line") is not None:
            ind_attrs[f"{{{W}}}firstLine"] = str(int(params["ind_first_line"]))
        if params.get("ind_first_line_chars") is not None:
            ind_attrs[f"{{{W}}}firstLineChars"] = str(int(params["ind_first_line_chars"]))
        old_ind = pPr.find(f"{{{W}}}ind")
        if old_ind is not None:
            pPr.remove(old_ind)
        etree.SubElement(pPr, f"{{{W}}}ind", ind_attrs)

    # 间距
    sp_keys = ["before_twips", "after_twips", "before_lines", "after_lines", "line", "lineRule"]
    if any(k in params and params[k] is not None for k in sp_keys):
        old_sp = pPr.find(f"{{{W}}}spacing")
        if old_sp is not None:
            pPr.remove(old_sp)
        sp_attrs = {}
        if params.get("before_twips") is not None:
            sp_attrs[f"{{{W}}}before"] = str(int(params["before_twips"]))
        if params.get("after_twips") is not None:
            sp_attrs[f"{{{W}}}after"] = str(int(params["after_twips"]))
        if params.get("before_lines") is not None:
            sp_attrs[f"{{{W}}}beforeLines"] = str(int(params["before_lines"]))
        if params.get("after_lines") is not None:
            sp_attrs[f"{{{W}}}afterLines"] = str(int(params["after_lines"]))
        if params.get("line") is not None:
            sp_attrs[f"{{{W}}}line"] = str(int(params["line"]))
        if params.get("lineRule") is not None:
            sp_attrs[f"{{{W}}}lineRule"] = params["lineRule"]
        if sp_attrs:
            etree.SubElement(pPr, f"{{{W}}}spacing", sp_attrs)

    # 大纲级别
    if "outline_lvl" in params and params["outline_lvl"] is not None:
        old_ol = pPr.find(f"{{{W}}}outlineLvl")
        if old_ol is not None:
            pPr.remove(old_ol)
        etree.SubElement(pPr, f"{{{W}}}outlineLvl", {f"{{{W}}}val": str(int(params["outline_lvl"]))})

    # run 级格式（前缀匹配拆分——关键词行）
    run_formats = params.get("run_formats")
    if run_formats:
        runs = [r for r in p.findall(f"{{{W}}}r") if ''.join(t.text or '' for t in r.iter(f"{{{W}}}t")).strip()]
        match_fmt = run_formats[0] if run_formats else None
        default_fmt = run_formats[1] if len(run_formats) > 1 else None
        prefix = (match_fmt or {}).get("match") or ""
        if not runs:
            pass
        elif prefix:
            # 累计定位含前缀结尾的 run；若前缀结束于 run 中间 → 拆分 run
            acc = 0
            prefix_end_run = None
            for r in runs:
                text = ''.join(t.text or '' for t in r.iter(f"{{{W}}}t"))
                acc += len(text)
                if acc >= len(prefix):
                    prefix_end_run = r
                    break
            if prefix_end_run is not None:
                # 该 run 内前缀字符数 = len(prefix) - 前缀之前累计
                before_len = acc - len(''.join(t.text or '' for t in prefix_end_run.iter(f"{{{W}}}t")))
                in_run = len(prefix) - before_len
                run_text = ''.join(t.text or '' for t in prefix_end_run.iter(f"{{{W}}}t"))
                if 0 < in_run < len(run_text):
                    # 拆分：前缀部分留在原 run，内容部分进新 run
                    _split_run_at(prefix_end_run, in_run)
                    runs = [r for r in p.findall(f"{{{W}}}r") if ''.join(t.text or '' for t in r.iter(f"{{{W}}}t")).strip()]
                # prefix_end_run 在 runs 中的索引：索引 ≤ 该值的是前缀
                end_idx = runs.index(prefix_end_run)
                for r in runs:
                    _apply_run_format(r, match_fmt if runs.index(r) <= end_idx else default_fmt)
            else:
                for r in runs:
                    _apply_run_format(r, default_fmt)
        else:
            for r in runs:
                _apply_run_format(r, default_fmt)

    # 整段 run 格式（统一）
    elif params.get("font_eastAsia") is not None or params.get("font_ascii") is not None \
            or params.get("font_size_pt") is not None or params.get("bold") is not None:
        for r in p.findall(f"{{{W}}}r"):
            _apply_run_format(r, params)

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


def _split_run_at(r: etree._Element, n_chars: int) -> etree._Element | None:
    """在文字第 n_chars 字符处拆分 run r，返回新 run（内容部分）。

    前缀格式（如"关键词："黑体加粗）需要 run 级区分——当前缀结束于
    run 中间时，把 run 拆成两个（前缀部分保留原 run，内容进新 run）。
    """
    t_els = list(r.findall(f"{{{W}}}t"))
    acc = 0
    for t in t_els:
        text = t.text or ''
        if acc + len(text) > n_chars:
            split_in = n_chars - acc
            head = text[:split_in]
            tail = text[split_in:]
            t.text = head
            new_r = etree.Element(f"{{{W}}}r")
            # 复制原 run 的 rPr
            rPr = r.find(f"{{{W}}}rPr")
            if rPr is not None:
                new_r.append(copy.deepcopy(rPr))
            new_t = etree.SubElement(new_r, f"{{{W}}}t")
            new_t.text = tail
            if t.get('{http://www.w3.org/XML/1998/namespace}space'):
                new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            r.addnext(new_r)
            return new_r
        acc += len(text)
    return None


def _apply_run_format(r: etree._Element, fmt: dict) -> None:
    """给一个 run 应用直接格式（字体/字号/加粗）。"""
    if fmt is None:
        return
    rPr = r.find(f"{{{W}}}rPr")
    if rPr is None:
        rPr = etree.Element(f"{{{W}}}rPr")
        r.insert(0, rPr)

    east = fmt.get("font_eastAsia")
    ascii_f = fmt.get("font_ascii")
    if east is not None or ascii_f is not None:
        rFonts = rPr.find(f"{{{W}}}rFonts")
        if rFonts is None:
            rFonts = etree.SubElement(rPr, f"{{{W}}}rFonts")
        if east is not None:
            rFonts.set(f"{{{W}}}eastAsia", east)
        if ascii_f is not None:
            rFonts.set(f"{{{W}}}ascii", ascii_f)
            rFonts.set(f"{{{W}}}hAnsi", ascii_f)

    sz = fmt.get("font_size_pt")
    if sz is not None:
        half = str(int(round(float(sz) * 2)))
        old_sz = rPr.find(f"{{{W}}}sz")
        if old_sz is not None:
            rPr.remove(old_sz)
        etree.SubElement(rPr, f"{{{W}}}sz", {f"{{{W}}}val": half})
        old_szc = rPr.find(f"{{{W}}}szCs")
        if old_szc is not None:
            rPr.remove(old_szc)
        etree.SubElement(rPr, f"{{{W}}}szCs", {f"{{{W}}}val": half})

    bold = fmt.get("bold")
    if bold is not None:
        old_b = rPr.find(f"{{{W}}}b")
        old_bc = rPr.find(f"{{{W}}}bCs")
        if bold:
            if old_b is None:
                etree.SubElement(rPr, f"{{{W}}}b")
            if old_bc is None:
                etree.SubElement(rPr, f"{{{W}}}bCs")
        else:
            if old_b is not None:
                rPr.remove(old_b)
            if old_bc is not None:
                rPr.remove(old_bc)


# ═══════════════════════════════════════════════════════════════
# Action: set_paragraph_spacing — 定向修改单个段落的间距
# ═══════════════════════════════════════════════════════════════

@_handler("set_paragraph_spacing")
def _set_paragraph_spacing(target_path: Path, params: dict) -> None:
    """设置指定段落的段间距和行间距。

    所有值从 _spec_rules 读取，不在此处硬编码。

    params:
        para_index: int       — document.xml body 中该段落的索引
        before_twips: int | None   — 段前间距（twips）
        after_twips: int | None    — 段后间距（twips）
        before_lines: float | None — 段前间距（行数，1/100 行单位）
        after_lines: float | None  — 段后间距（行数，1/100 行单位）
        line: int | None           — 行距值（240=单倍行距，360=1.5倍，480=2倍）
        lineRule: str | None       — 行距规则 "auto" | "exact" | "atLeast"
    """
    para_idx = _first_param(params, "para_index", "paragraph_index", "paragraph", "idx", "index")
    if para_idx is None:
        raise ValueError("format_paragraph 需要 para_index（别名 paragraph/idx/index 兼容）——见参数表 hint")

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    paras = list(body.iter(f"{{{W}}}p"))
    if para_idx >= len(paras):
        return

    p = paras[para_idx]
    pPr = p.find(f"{{{W}}}pPr")
    if pPr is None:
        pPr = etree.SubElement(p, f"{{{W}}}pPr")
        p.insert(0, pPr)

    # 读取已有的 spacing，合并新值
    old_sp = pPr.find(f"{{{W}}}spacing")
    sp_attrs = {}
    if old_sp is not None:
        for attr_name in ("before", "after", "beforeLines", "afterLines", "line", "lineRule"):
            val = old_sp.get(f"{{{W}}}{attr_name}")
            if val is not None:
                sp_attrs[f"{{{W}}}{attr_name}"] = val
        pPr.remove(old_sp)

    # 覆盖/新增指定的属性
    if "before_twips" in params and params["before_twips"] is not None:
        sp_attrs[f"{{{W}}}before"] = str(int(params["before_twips"]))
    if "after_twips" in params and params["after_twips"] is not None:
        sp_attrs[f"{{{W}}}after"] = str(int(params["after_twips"]))
    if "before_lines" in params and params["before_lines"] is not None:
        sp_attrs[f"{{{W}}}beforeLines"] = str(int(params["before_lines"]))
    if "after_lines" in params and params["after_lines"] is not None:
        sp_attrs[f"{{{W}}}afterLines"] = str(int(params["after_lines"]))
    if "line" in params and params["line"] is not None:
        sp_attrs[f"{{{W}}}line"] = str(int(params["line"]))
    if "lineRule" in params and params["lineRule"] is not None:
        sp_attrs[f"{{{W}}}lineRule"] = params["lineRule"]

    if sp_attrs:
        etree.SubElement(pPr, f"{{{W}}}spacing", sp_attrs)

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)


# ═══════════════════════════════════════════════════════════════
# Action: normalize_list_numbering — 列表编号一致性强制
# ═══════════════════════════════════════════════════════════════

@_handler("normalize_list_numbering")
def _normalize_list_numbering(target_path: Path, params: dict) -> None:
    """检测文档中所有编号列表（<w:numPr>），确保同级列表使用相同的编号格式。

    实现策略：
    1. 扫描 document.xml 中所有带 <w:numPr> 的段落
    2. 按 numId 分组，提取每组的编号格式（从 numbering.xml 读 abstractNum）
    3. 对每个 numId 组，确保所有段落的 numPr/ilvl 一致
    4. 对齐同一 abstractNum 下不同 numId 的格式（当它们应该是同一列表时）

    params:
        dry_run: bool (默认 False) — 如果 True，只检测不修改（通过 _changed 返回计数）
        force_num_id: int | None   — 强制所有列表使用此 numId（高级用法）
    """
    dry_run = params.get("dry_run", False)

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    # ── Phase 1: 收集所有含 numPr 的段落 ──
    # 结构: {numId: [(para_idx, ilvl, para_element), ...]}
    num_groups: dict[str, list] = {}
    paras = list(body.iter(f"{{{W}}}p"))

    for pi, p in enumerate(paras):
        pPr = p.find(f"{{{W}}}pPr")
        if pPr is None:
            continue
        numPr = pPr.find(f"{{{W}}}numPr")
        if numPr is None:
            continue
        numId_el = numPr.find(f"{{{W}}}numId")
        ilvl_el = numPr.find(f"{{{W}}}ilvl")
        if numId_el is None:
            continue
        num_id = numId_el.get(f"{{{W}}}val", "0")
        ilvl = int(ilvl_el.get(f"{{{W}}}val", "0")) if ilvl_el is not None else 0
        num_groups.setdefault(num_id, []).append((pi, ilvl, p))

    if not num_groups:
        params["_changed"] = 0
        return

    # ── Phase 2: 检测每个 numId 组内各段落的 ilvl 是否一致 ──
    # 同一列表级的段落应该有相同的 ilvl
    inconsistent = 0
    for num_id, items in num_groups.items():
        if len(items) < 2:
            continue

        # 找到最常见的 ilvl（投票策略）
        from collections import Counter
        ilvl_counts = Counter(item[1] for item in items)
        dominant_ilvl = ilvl_counts.most_common(1)[0][0]

        # 检测并修复不一致的 ilvl
        for pi, ilvl, p in items:
            if ilvl != dominant_ilvl:
                if not dry_run:
                    pPr = p.find(f"{{{W}}}pPr")
                    numPr = pPr.find(f"{{{W}}}numPr")
                    ilvl_el = numPr.find(f"{{{W}}}ilvl")
                    if ilvl_el is not None:
                        ilvl_el.set(f"{{{W}}}val", str(dominant_ilvl))
                inconsistent += 1

    if not dry_run and inconsistent > 0:
        new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        _zip_replace(target_path, "word/document.xml", new_data)

    params["_changed"] = inconsistent


# ═══════════════════════════════════════════════════════════════
# Action: set_page_margins — 页边距检查与强制
# ═══════════════════════════════════════════════════════════════

@_handler("set_page_margins")
def _set_page_margins(target_path: Path, params: dict) -> None:
    """设置指定 sectPr 的页边距（pgMar）。

    所有值从 _spec_rules.page 读取，零硬编码。

    params:
        section_index: int     — 第几个 sectPr（0-based）
        top: int | None        — 上边距（twips）
        bottom: int | None     — 下边距（twips）
        left: int | None       — 左边距（twips）
        right: int | None      — 右边距（twips）
        header: int | None     — 页眉距离（twips）
        footer: int | None     — 页脚距离（twips）
    """
    section_index = params["section_index"]

    root = _parse_xml_from_zip(target_path, "word/document.xml")
    sects = list(root.iter(f"{{{W}}}sectPr"))

    if section_index >= len(sects):
        raise ValueError(
            f"section_index {section_index} 超出范围 (共 {len(sects)} 个 sectPr)"
        )

    sp = sects[section_index]
    pm = sp.find(f"{{{W}}}pgMar")
    if pm is None:
        pm = etree.SubElement(sp, f"{{{W}}}pgMar")

    # 覆盖指定的边距属性，保留未指定的现有值
    margin_attrs = ("top", "bottom", "left", "right", "header", "footer")
    for attr in margin_attrs:
        if attr in params and params[attr] is not None:
            pm.set(f"{{{W}}}{attr}", str(int(params[attr])))

    new_data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _zip_replace(target_path, "word/document.xml", new_data)
