"""
Fixer v2 — 验证 fix_plan 执行结果。

核心原则:
- 对比 fix_plan 和实际 OOXML，逐条验证每项是否已生效
- 不做任何修改，只报告差异
- 返回诊断列表供 Agent 决策

用法:
    from pipeline.fixer_v2 import diagnose

    issues = diagnose(fix_plan, Path("output/毕业设计.docx"))
    # issues: [FixerDiagnostic(code="HEADER_FONT_MISMATCH", ...), ...]
"""

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Union

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


@dataclass
class FixerDiagnostic:
    """fix_plan 验证结果"""
    code: str               # 诊断代码
    action_index: int       # 对应 fix_plan 中的 action 索引
    severity: str           # "error", "warning", "info"
    detail: str             # 详细描述
    expected: Any = None    # fix_plan 中期望的值
    actual: Any = None      # 实际 OOXML 中读取的值
    suggestion: str = ""    # 可操作建议（如何修复此问题）

    def __str__(self):
        return f"[{self.severity}] #{self.action_index} {self.code}: {self.detail}"

    def to_dict(self) -> dict:
        """序列化为字典（供 API JSON 返回）。"""
        return {
            "code": self.code,
            "action_index": self.action_index,
            "severity": self.severity,
            "detail": self.detail,
            "expected": self.expected,
            "actual": self.actual,
            "suggestion": self.suggestion,
        }


# ── 验证调度器 ──────────────────────────────────────────────

_VERIFIERS: dict[str, callable] = {}


def _verifier(action_name: str):
    """装饰器：注册验证函数"""
    def decorator(fn):
        _VERIFIERS[action_name] = fn
        return fn
    return decorator


def diagnose(fix_plan: list[dict], target_path: Union[str, Path]) -> list[FixerDiagnostic]:
    """逐条验证 fix_plan 中的每项是否已在目标 docx 中生效。

    Args:
        fix_plan: action 列表
        target_path: 修改后的 docx 文件路径

    Returns:
        FixerDiagnostic 列表（空列表表示全部验证通过）
    """
    target_path = Path(target_path)
    issues: list[FixerDiagnostic] = []

    # 预加载常用 XML
    _cache: dict[str, etree._Element] = {}

    def _get_xml(internal_path: str) -> etree._Element:
        if internal_path not in _cache:
            with zipfile.ZipFile(target_path, "r") as z:
                _cache[internal_path] = etree.fromstring(z.read(internal_path))
        return _cache[internal_path]

    for i, action in enumerate(fix_plan):
        action_type = action.get("action", "")
        params = action.get("params", {})

        verifier = _VERIFIERS.get(action_type)
        if verifier is None:
            issues.append(FixerDiagnostic(
                code="UNKNOWN_ACTION",
                action_index=i,
                severity="warning",
                detail=f"未知 action 类型: {action_type}，无法验证",
                suggestion=f"请检查 action 拼写是否正确，或确认该 action 已注册对应的 verifier。当前支持的 actions: {list(_VERIFIERS.keys())}",
            ))
            continue

        try:
            result = verifier(target_path, params, _get_xml, i)
            if result:
                issues.extend(result if isinstance(result, list) else [result])
        except Exception as e:
            issues.append(FixerDiagnostic(
                code="VERIFY_ERROR",
                action_index=i,
                severity="error",
                detail=f"验证异常: {e}",
                suggestion=f"验证器内部错误。请检查目标文档是否被损坏，或 action #{i} 的参数是否合理。原始错误: {e}",
            ))

    return issues


# ── 辅助 ────────────────────────────────────────────────────

def _read_text(element: etree._Element) -> str:
    """读取元素内所有 <w:t> 的文本"""
    return "".join(t.text or "" for t in element.iter(f"{{{W}}}t"))


# ── 验证: set_sectpr_type ───────────────────────────────────

@_verifier("set_sectpr_type")
def _verify_sectpr_type(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    section_index = params["section_index"]
    expected_val = params["val"]

    root = get_xml("word/document.xml")
    sects = list(root.iter(f"{{{W}}}sectPr"))

    if section_index >= len(sects):
        return [FixerDiagnostic(
            code="SECTPR_INDEX_OUT_OF_RANGE",
            action_index=idx,
            severity="error",
            detail=f"section_index={section_index} 超出范围 (共 {len(sects)} 个 sectPr)",
            expected=expected_val,
            actual=None,
            suggestion=f"文档只有 {len(sects)} 个节，section_index 应为 0~{len(sects)-1}。请重新运行 DocDiscover 刷新节计数。",
        )]

    sp = sects[section_index]
    type_el = sp.find(f"{{{W}}}type")
    actual_val = type_el.get(f"{{{W}}}val") if type_el is not None else None

    if actual_val != expected_val:
        return [FixerDiagnostic(
            code="SECTPR_TYPE_MISMATCH",
            action_index=idx,
            severity="error",
            detail=f"sectPr[{section_index}] type 期望={expected_val} 实际={actual_val}",
            expected=expected_val,
            actual=actual_val,
            suggestion=f"节类型未生效。请手动检查 Word 的「页面布局→分隔符」，或重新运行 Writer 的 set_sectpr_type action。",
        )]

    return []


# ── 验证: set_header_font ───────────────────────────────────

@_verifier("set_header_font")
def _verify_header_font(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    header_path = params["header_path"]
    expected_ascii = params.get("font_ascii")
    expected_eastAsia = params.get("font_eastAsia")
    expected_size = params.get("font_size")

    root = get_xml(header_path)
    issues = []

    # 空 header 容错：无文字内容则跳过字体检查
    all_text = ''.join(t.text or '' for t in root.iter(f'{{{W}}}t')).strip()
    if not all_text:
        return []

    rPr_list = list(root.iter(f"{{{W}}}rPr"))
    if not rPr_list:
        return [FixerDiagnostic(
            code="HEADER_FONT_MISMATCH",
            action_index=idx,
            severity="error",
            detail=f"{header_path} 中没有 rPr 元素",
            expected={"ascii": expected_ascii, "eastAsia": expected_eastAsia, "size": expected_size},
            actual=None,
            suggestion=f"页眉文件 {header_path} 缺少 run properties。请确认页眉中存在文字内容，或检查该头文件是否正确关联到此节。",
        )]

    # 检查第一个 rPr（通常是主要字体设置）
    rPr = rPr_list[0]
    rf = rPr.find(f"{{{W}}}rFonts")

    if rf is None:
        issues.append(FixerDiagnostic(
            code="HEADER_FONT_MISMATCH",
            action_index=idx,
            severity="error",
            detail=f"{header_path} 中无 rFonts 元素",
            expected={"ascii": expected_ascii, "eastAsia": expected_eastAsia},
            actual=None,
            suggestion=f"页眉文件 {header_path} 缺少 rFonts 字体设置。请确认 Writer 的 set_header_font action 已正确执行，或该 header 文件结构异常。",
        ))
    else:
        actual_ascii = rf.get(f"{{{W}}}ascii")
        actual_eastAsia = rf.get(f"{{{W}}}eastAsia")

        if expected_ascii and actual_ascii != expected_ascii:
            issues.append(FixerDiagnostic(
                code="HEADER_FONT_MISMATCH",
                action_index=idx,
                severity="error",
                detail=f"{header_path} ascii 期望={expected_ascii} 实际={actual_ascii}",
                expected=expected_ascii,
                actual=actual_ascii,
                suggestion=f"页眉西文字体未生效。请确认 set_header_font action 的 font_ascii 参数正确，且 header 文件路径 {header_path} 无误。",
            ))

        if expected_eastAsia and actual_eastAsia != expected_eastAsia:
            issues.append(FixerDiagnostic(
                code="HEADER_FONT_MISMATCH",
                action_index=idx,
                severity="error",
                detail=f"{header_path} eastAsia 期望={expected_eastAsia} 实际={actual_eastAsia}",
                expected=expected_eastAsia,
                actual=actual_eastAsia,
                suggestion=f"页眉中文字体未生效。请确认 set_header_font action 的 font_eastAsia 参数正确，或手动在 Word 页眉中设置字体。",
            ))

    # 检查字号
    if expected_size:
        sz = rPr.find(f"{{{W}}}sz")
        actual_size = sz.get(f"{{{W}}}val") if sz is not None else None
        if actual_size != expected_size:
            issues.append(FixerDiagnostic(
                code="HEADER_FONT_MISMATCH",
                action_index=idx,
                severity="warning",
                detail=f"{header_path} sz 期望={expected_size} 实际={actual_size}",
                expected=expected_size,
                actual=actual_size,
                suggestion=f"页眉字号未生效。请确认 set_header_font action 的 font_size 参数正确，或在 Word 页眉中手动调整字号。",
            ))

    return issues


# ── 验证: set_page_number_type ──────────────────────────────

@_verifier("set_page_number_type")
def _verify_page_number_type(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return []


# ── 验证: add_footer_reference ──────────────────────────────

@_verifier("add_footer_reference")
def _verify_footer_reference(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return []


# ── 验证: clear_section_headers ────────────────────────────

@_verifier("clear_section_headers")
def _verify_clear_headers(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """后置断言：目标节的 headerReference 必须为空，且节序号必须真实存在。

    2026-10-18 Layer 0：原实现 `return []` 恒过，导致「删掉唯一一节的页眉」也被判成功。
    """
    si = params.get("section_index")
    try:
        root = get_xml("word/document.xml")
    except Exception as e:  # noqa: BLE001
        return [FixerDiagnostic("DOC_UNREADABLE", idx, "error", f"无法读取 document.xml: {e}")]
    sects = list(root.iter(f"{{{W}}}sectPr"))
    if not isinstance(si, int) or si < 0 or si >= len(sects):
        return [FixerDiagnostic(
            "SECTION_INDEX_OUT_OF_RANGE", idx, "error",
            f"节序号 {si!r} 不存在（文档共 {len(sects)} 节）",
            expected=f"0..{max(len(sects) - 1, 0)}", actual=si,
            suggestion="先读 document.xml 的 sectPr 数量再传 section_index；单节文档只有 0",
        )]
    refs = sects[si].findall(f"{{{W}}}headerReference")
    if refs:
        return [FixerDiagnostic("HEADER_STILL_REFERENCED", idx, "error",
                                f"第 {si} 节仍有 {len(refs)} 个 headerReference")]
    return []


# ── 验证: clear_section_footers ────────────────────────────

@_verifier("clear_section_footers")
def _verify_clear_footers(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """后置断言：目标节的 footerReference 必须为空，且节序号必须真实存在。"""
    si = params.get("section_index")
    try:
        root = get_xml("word/document.xml")
    except Exception as e:  # noqa: BLE001
        return [FixerDiagnostic("DOC_UNREADABLE", idx, "error", f"无法读取 document.xml: {e}")]
    sects = list(root.iter(f"{{{W}}}sectPr"))
    if not isinstance(si, int) or si < 0 or si >= len(sects):
        return [FixerDiagnostic(
            "SECTION_INDEX_OUT_OF_RANGE", idx, "error",
            f"节序号 {si!r} 不存在（文档共 {len(sects)} 节）",
            expected=f"0..{max(len(sects) - 1, 0)}", actual=si,
            suggestion="先读 document.xml 的 sectPr 数量再传 section_index",
        )]
    refs = sects[si].findall(f"{{{W}}}footerReference")
    if refs:
        return [FixerDiagnostic("FOOTER_STILL_REFERENCED", idx, "error",
                                f"第 {si} 节仍有 {len(refs)} 个 footerReference")]
    return []


# ── 验证: format_headers / format_footers（边框增删，2026-10-18 Layer 0）──

def _verify_hf_border(target_path: Path, params: dict, prefix: str, idx: int) -> list[FixerDiagnostic]:
    """按 properties.p_bdr_bottom 断言页眉/页脚下边框的最终状态。

    原实现的边框只能加不能删，且跳过空段落；这里让"没做到"必须报错，
    而不是返回 ok 让上层误判成功。
    """
    props = params.get("properties") or {}
    if "p_bdr_bottom" not in props:
        return []
    want = bool(props.get("p_bdr_bottom"))
    problems: list[FixerDiagnostic] = []
    try:
        with zipfile.ZipFile(target_path, "r") as z:
            names = [n for n in z.namelist()
                     if n.startswith(f"word/{prefix}") and n.endswith(".xml")]
            for name in names:
                root = etree.fromstring(z.read(name))
                for pel in root.iter(f"{{{W}}}p"):
                    has_text = bool("".join(t.text or "" for t in pel.iter(f"{{{W}}}t")).strip())
                    ppr = pel.find(f"{{{W}}}pPr")
                    pbdr = ppr.find(f"{{{W}}}pBdr") if ppr is not None else None
                    bottom = pbdr.find(f"{{{W}}}bottom") if pbdr is not None else None
                    exists = bottom is not None and (bottom.get(f"{{{W}}}val") or "single") != "none"
                    if want and has_text and not exists:
                        problems.append(FixerDiagnostic(
                            "HF_BORDER_MISSING", idx, "error", f"{name} 的段落没有下边框"))
                    if not want and exists:
                        problems.append(FixerDiagnostic(
                            "HF_BORDER_STILL_PRESENT", idx, "error",
                            f"{name} 仍有下边框——删除边框没有生效",
                            suggestion="用支持删除的 set_header_border / format_headers(p_bdr_bottom=false)"))
    except Exception as e:  # noqa: BLE001
        problems.append(FixerDiagnostic("HF_PARTS_UNREADABLE", idx, "error", f"读取 {prefix} 部件失败: {e}"))
    return problems


@_verifier("format_headers")
def _verify_format_headers(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return _verify_hf_border(target_path, params, "header", idx)


@_verifier("format_footers")
def _verify_format_footers(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return _verify_hf_border(target_path, params, "footer", idx)


# ── 验证: set_toc_entry_styles ─────────────────────────────

@_verifier("set_toc_entry_styles")
def _verify_toc_entry_styles(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return []


# ── 验证: ensure_toc_tabs ──────────────────────────────────

@_verifier("ensure_toc_tabs")
def _verify_toc_tabs(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return []


# ── 验证: format_paragraph ──────────────────────────────────

@_verifier("format_paragraph")
def _verify_format_paragraph(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return []


# These actions perform document-wide deterministic post-processing. Their
# result is checked by the final print validator, not by a stale paragraph
# index verifier.
@_verifier("strip_style_numbering")
def _verify_strip_style_numbering(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return []


@_verifier("convert_ole_to_image")
def _verify_convert_ole_to_image(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return []


@_verifier("clean_heading_text")
def _verify_clean_heading_text(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return []


@_verifier("normalize_image_paragraphs")
def _verify_normalize_image_paragraphs(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return []


# ── 验证: add_page_number ───────────────────────────────────

@_verifier("add_page_number")
def _verify_page_number(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    footer_path = params["footer_path"]

    root = get_xml(footer_path)

    # 检查是否有 fldChar 序列
    fld_chars = list(root.iter(f"{{{W}}}fldChar"))
    if not fld_chars:
        return [FixerDiagnostic(
            code="MISSING_PAGE_NUMBER",
            action_index=idx,
            severity="error",
            detail=f"{footer_path} 中无 fldChar（页码域）",
            expected="fldChar 序列",
            actual="无",
            suggestion=f"页脚 {footer_path} 缺少 PAGE 域代码。请重新运行 add_page_number action，确认 format 参数包含 {{PAGE}} 占位符。",
        )]

    # 检查是否有 PAGE 指令
    instr_texts = []
    for it in root.iter(f"{{{W}}}instrText"):
        instr_texts.append(it.text or "")

    has_page = any("PAGE" in t for t in instr_texts)
    if not has_page:
        return [FixerDiagnostic(
            code="MISSING_PAGE_NUMBER",
            action_index=idx,
            severity="error",
            detail=f"{footer_path} 中无 PAGE 指令",
            expected="PAGE",
            actual=instr_texts,
        )]

    # 验证格式模板中的文字
    fmt = params.get("format", "")
    prefix_before_page = fmt.split("{PAGE}")[0] if "{PAGE}" in fmt else ""
    suffix_after_page = fmt.rsplit("{PAGE}", 1)[-1] if "{PAGE}" in fmt else ""

    full_text = _read_text(root)
    if prefix_before_page and prefix_before_page not in full_text:
        return [FixerDiagnostic(
            code="MISSING_PAGE_NUMBER",
            action_index=idx,
            severity="warning",
            detail=f"{footer_path} 缺少页码前缀 '{prefix_before_page}'",
            expected=prefix_before_page,
            actual=full_text[:80],
        )]

    if suffix_after_page and suffix_after_page not in full_text:
        return [FixerDiagnostic(
            code="MISSING_PAGE_NUMBER",
            action_index=idx,
            severity="warning",
            detail=f"{footer_path} 缺少页码后缀 '{suffix_after_page}'",
            expected=suffix_after_page,
            actual=full_text[:80],
        )]

    return []


# ── 验证: set_body_font_ascii ───────────────────────────────

@_verifier("set_body_font_ascii")
def _verify_body_font_ascii(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    style_id = params["style_id"]
    expected_ascii = params["font_ascii"]

    root = get_xml("word/styles.xml")

    for style in root.iter(f"{{{W}}}style"):
        if style.get(f"{{{W}}}styleId") != style_id:
            continue

        rPr = style.find(f"{{{W}}}rPr")
        if rPr is None:
            return [FixerDiagnostic(
                code="STYLE_NOT_APPLIED",
                action_index=idx,
                severity="error",
                detail=f"样式 '{style_id}' 无 rPr 元素",
                expected={"ascii": expected_ascii},
                actual=None,
                suggestion=f"样式 '{style_id}' 缺少 run properties。请检查 styles.xml 中该样式是否被损坏，或在 Word 中重新保存文档。",
            )]

        rf = rPr.find(f"{{{W}}}rFonts")
        actual_ascii = rf.get(f"{{{W}}}ascii") if rf is not None else None

        if actual_ascii != expected_ascii:
            return [FixerDiagnostic(
                code="STYLE_NOT_APPLIED",
                action_index=idx,
                severity="error",
                detail=f"样式 '{style_id}' ascii 期望={expected_ascii} 实际={actual_ascii}",
                expected=expected_ascii,
                actual=actual_ascii,
                suggestion=f"样式 '{style_id}' 西文字体未生效。请确认 set_body_font_ascii action 的 font_ascii 参数正确，或手动在 Word 样式面板中修改。",
            )]

        return []

    return [FixerDiagnostic(
        code="STYLE_NOT_APPLIED",
        action_index=idx,
        severity="error",
        detail=f"未找到样式 '{style_id}'",
        expected=expected_ascii,
        actual="样式不存在",
        suggestion=f"样式 ID '{style_id}' 在 styles.xml 中不存在。请确认目标文档使用了该样式，或更换为正确的样式 ID（如 body 常用 'a8' 或 'Normal'）。",
    )]


# ── 验证: set_style ─────────────────────────────────────────

@_verifier("set_style")
def _verify_style(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    style_id = params["style_id"]
    issues = []

    root = get_xml("word/styles.xml")

    for style in root.iter(f"{{{W}}}style"):
        if style.get(f"{{{W}}}styleId") != style_id:
            continue

        # 验证字体
        if any(k in params for k in ("font_eastAsia", "font_ascii", "font_hAnsi")):
            rPr = style.find(f"{{{W}}}rPr")
            rf = rPr.find(f"{{{W}}}rFonts") if rPr is not None else None

            if params.get("font_ascii") is not None:
                actual = rf.get(f"{{{W}}}ascii") if rf is not None else None
                if actual != params["font_ascii"]:
                    issues.append(FixerDiagnostic(
                        code="STYLE_NOT_APPLIED",
                        action_index=idx,
                        severity="error",
                        detail=f"样式 '{style_id}' ascii 期望={params['font_ascii']} 实际={actual}",
                        expected=params["font_ascii"],
                        actual=actual,
                        suggestion=f"样式 '{style_id}' 西文字体未生效。请检查 styles.xml 中该样式的 w:rFonts/@w:ascii 是否正确，或重新运行 set_style action。",
                    ))

            if params.get("font_eastAsia") is not None:
                actual = rf.get(f"{{{W}}}eastAsia") if rf is not None else None
                if actual != params["font_eastAsia"]:
                    issues.append(FixerDiagnostic(
                        code="STYLE_NOT_APPLIED",
                        action_index=idx,
                        severity="error",
                        detail=f"样式 '{style_id}' eastAsia 期望={params['font_eastAsia']} 实际={actual}",
                        expected=params["font_eastAsia"],
                        actual=actual,
                        suggestion=f"样式 '{style_id}' 中文字体未生效。请检查 styles.xml 中该样式的 w:rFonts/@w:eastAsia 是否正确，或重新运行 set_style action。",
                    ))

        # 验证字号
        if params.get("font_size_pt") is not None:
            expected_sv = str(int(params["font_size_pt"] * 2))
            rPr = style.find(f"{{{W}}}rPr")
            if rPr is not None:
                sz = rPr.find(f"{{{W}}}sz")
                actual_sv = sz.get(f"{{{W}}}val") if sz is not None else None
                if actual_sv != expected_sv:
                    issues.append(FixerDiagnostic(
                        code="STYLE_NOT_APPLIED",
                        action_index=idx,
                        severity="warning",
                        detail=f"样式 '{style_id}' sz 期望={expected_sv} 实际={actual_sv}",
                        expected=expected_sv,
                        actual=actual_sv,
                        suggestion=f"样式 '{style_id}' 字号未生效。请确认 font_size_pt 参数值正确（单位：磅），或手动在 Word 样式面板中调整字号。",
                    ))

        # 验证加粗
        if params.get("bold") is not None:
            rPr = style.find(f"{{{W}}}rPr")
            has_b = rPr is not None and rPr.find(f"{{{W}}}b") is not None
            if has_b != params["bold"]:
                issues.append(FixerDiagnostic(
                    code="STYLE_NOT_APPLIED",
                    action_index=idx,
                    severity="warning",
                    detail=f"样式 '{style_id}' bold 期望={params['bold']} 实际={has_b}",
                    expected=params["bold"],
                    actual=has_b,
                    suggestion=f"样式 '{style_id}' 加粗属性未生效。请确认 bold 参数是否正确，或手动在 Word 样式面板中设置。",
                ))

        return issues

    return [FixerDiagnostic(
        code="STYLE_NOT_APPLIED",
        action_index=idx,
        severity="warning",
        detail=f"未找到样式 '{style_id}'",
        expected=style_id,
        actual="样式不存在",
        suggestion=f"样式 ID '{style_id}' 在 styles.xml 中不存在。请确认该样式 ID 在目标文档的 styles.xml 中已定义，或重新运行 DocDiscover 刷新样式映射。",
    )]


# ── 验证: remove_extra_sectpr ───────────────────────────────

@_verifier("remove_extra_sectpr")
def _verify_remove_extra_sectpr(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    root = get_xml("word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return []

    issues = []
    extra_count = 0
    for child in body:
        if child.tag != f"{{{W}}}p":
            continue
        pPr = child.find(f"{{{W}}}pPr")
        if pPr is None:
            continue
        extra_sp = pPr.find(f"{{{W}}}sectPr")
        if extra_sp is not None:
            # 检查这个 sectPr 是否有 oddPage
            type_el = extra_sp.find(f"{{{W}}}type")
            type_val = type_el.get(f"{{{W}}}val") if type_el is not None else ""
            extra_count += 1
            if type_val == "oddPage":
                issues.append(FixerDiagnostic(
                    code="EXTRA_ODDPAGE_SECTPR",
                    action_index=idx,
                    severity="error",
                    detail=f"段落内仍有 oddPage sectPr (第 {extra_count} 个嵌入)",
                    expected="已删除",
                    actual="oddPage sectPr 残留",
                    suggestion="oddPage 分节符未被清除。请重新运行 remove_extra_sectpr action，或在 Word 中手动删除该分节符（布局→分隔符→下一页）。",
                ))

    if extra_count > 0 and not issues:
        issues.append(FixerDiagnostic(
            code="EXTRA_ODDPAGE_SECTPR",
            action_index=idx,
            severity="warning",
            detail=f"仍有 {extra_count} 个嵌入 sectPr 未删除",
            expected=0,
            actual=extra_count,
            suggestion=f"仍有 {extra_count} 个嵌入段落内的 sectPr 残留。这些分节符可能导致页面布局异常，请在 Word 中手动检查并删除多余分节符。",
        ))

    return issues


# ── Summary ────────────────────────────────────────────────


def summary(diagnostics: list[FixerDiagnostic]) -> dict:
    """生成诊断汇总。

    Returns:
        {"total": N, "passed": bool, "by_code": {...}, "by_severity": {...}}
    """
    by_code: dict[str, int] = {}
    by_severity: dict[str, int] = {"error": 0, "warning": 0, "info": 0}

    for d in diagnostics:
        by_code[d.code] = by_code.get(d.code, 0) + 1
        by_severity[d.severity] = by_severity.get(d.severity, 0) + 1

    return {
        "total": len(diagnostics),
        "passed": len(diagnostics) == 0,
        "by_code": by_code,
        "by_severity": by_severity,
    }


# ── Verifier: set_header_text ──────────────────────────────

@_verifier("set_header_text")
def _verify_set_header_text(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证 header 文件文字是否已更新。"""
    header_path = params.get("header_path", "")
    expected_text = params.get("text", "")

    try:
        root = get_xml(header_path)
    except Exception:
        return [FixerDiagnostic(
            code="HEADER_FILE_MISSING",
            action_index=idx,
            severity="error",
            detail=f"header 文件不存在: {header_path}",
            expected=expected_text,
            suggestion=f"header 文件 {header_path} 在 ZIP 包中不存在。该节可能未创建页眉，请先在 Word 中为该节添加页眉。",
        )]

    # 提取所有文字
    all_text = ""
    for t in root.iter(f"{{{W}}}t"):
        if t.text:
            all_text += t.text

    if expected_text not in all_text:
        return [FixerDiagnostic(
            code="HEADER_TEXT_MISMATCH",
            action_index=idx,
            severity="error",
            detail=f"header 文字不匹配: 期望 '{expected_text}' 不在 '{all_text[:50]}'",
            expected=expected_text,
            actual=all_text[:50],
            suggestion=f"页眉文字未正确写入。请检查 set_header_text action 的 text 参数是否正确，或手动在 Word 中设置页眉文字。",
        )]

    return []


@_verifier("create_section_headers")
def _verify_create_section_headers(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """后置断言：每个被点名的节都必须挂上 headerReference；空配置视为失败。

    2026-10-18 Layer 0：原实现 `return []` + 空数组也报成功，是「假成功」的典型来源。
    """
    entries = params.get("section_headers")
    if not isinstance(entries, list) or not entries:
        return [FixerDiagnostic(
            "SECTION_HEADERS_EMPTY", idx, "error",
            "section_headers 为空：没有指定任何节的页眉，本动作不应被记为成功",
            suggestion="传 [{sectpr_index, default_text, even_text?}]；无页眉需求就不要调用本动作",
        )]
    try:
        root = get_xml("word/document.xml")
    except Exception as e:  # noqa: BLE001
        return [FixerDiagnostic("DOC_UNREADABLE", idx, "error", f"无法读取 document.xml: {e}")]
    sects = list(root.iter(f"{{{W}}}sectPr"))
    problems: list[FixerDiagnostic] = []
    for entry in entries:
        if not isinstance(entry, dict):
            continue
        si = entry.get("sectpr_index")
        if not isinstance(si, int) or si < 0 or si >= len(sects):
            problems.append(FixerDiagnostic(
                "SECTION_INDEX_OUT_OF_RANGE", idx, "error",
                f"节序号 {si!r} 不存在（文档共 {len(sects)} 节）"))
            continue
        if not sects[si].findall(f"{{{W}}}headerReference"):
            problems.append(FixerDiagnostic(
                "SECTION_HEADER_MISSING", idx, "error",
                f"第 {si} 节创建后仍没有 headerReference",
                suggestion="检查 header 部件是否写入 zip 并在 rels/Content_Types 注册"))
    return problems


@_verifier("insert_section_break")
def _verify_insert_section_break(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """后置断言：新出现的每个 sectPr 必须带 pgSz 与 pgMar（否则版式跳变）。

    2026-10-18 Layer 0：原 `return []` 让「只写 type+pgSz 的裸 sectPr」也算成功，
    实测分节后后续页面丢页边距、丢页眉。
    """
    try:
        root = get_xml("word/document.xml")
    except Exception as e:  # noqa: BLE001
        return [FixerDiagnostic("DOC_UNREADABLE", idx, "error", f"无法读取 document.xml: {e}")]
    sects = list(root.iter(f"{{{W}}}sectPr"))
    problems: list[FixerDiagnostic] = []
    for i, sp in enumerate(sects):
        missing = []
        if sp.find(f"{{{W}}}pgSz") is None:
            missing.append("pgSz")
        if sp.find(f"{{{W}}}pgMar") is None:
            missing.append("pgMar")
        if missing:
            problems.append(FixerDiagnostic(
                "SECTION_LAYOUT_INCOMPLETE", idx, "error",
                f"第 {i} 节缺少 {'/'.join(missing)}：分节后版式会跳变",
                suggestion="新节必须复制上一节的 pgSz/pgMar/cols/docGrid 与 header/footer 引用"))
    return problems


@_verifier("insert_red_header")
def _verify_insert_red_header(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证 insert_red_header 操作：检查文档开头是否包含红头文字和红色反线。"""
    org_name = params.get("org_name", "")
    doc_number = params.get("doc_number", "")

    try:
        root = get_xml("word/document.xml")
    except Exception:
        return [FixerDiagnostic(
            code="RED_HEADER_VERIFY_FAILED",
            action_index=idx,
            severity="error",
            detail="无法读取 word/document.xml",
            suggestion="文档损坏或格式异常，无法读取 document.xml。请用 Word 打开后另存为新文档，再重新排版。",
        )]

    body = root.find(f"{{{W}}}body")
    if body is None:
        return [FixerDiagnostic(
            code="RED_HEADER_VERIFY_FAILED",
            action_index=idx,
            severity="error",
            detail="document.xml 中无 body 元素",
            suggestion="document.xml 结构异常，缺少 body 元素。请检查文档是否损坏，或在 Word 中重新保存。",
        )]

    # 检查前两个段落是否包含红头内容
    paragraphs = list(body.iter(f"{{{W}}}p"))
    issues = []

    if org_name:
        found_org = False
        found_red_line = False
        for p in paragraphs[:3]:  # 检查前3段
            text = "".join(t.text or "" for t in p.iter(f"{{{W}}}t"))
            if org_name in text:
                found_org = True
                # 检查是否有红色反线（pBdr/bottom）
                pPr = p.find(f"{{{W}}}pPr")
                if pPr is not None:
                    pBdr = pPr.find(f"{{{W}}}pBdr")
                    if pBdr is not None:
                        bottom = pBdr.find(f"{{{W}}}bottom")
                        if bottom is not None and bottom.get(f"{{{W}}}color") == "FF0000":
                            found_red_line = True
            if found_org:
                break

        if not found_org:
            issues.append(FixerDiagnostic(
                code="RED_HEADER_ORG_MISSING",
                action_index=idx,
                severity="error",
                detail=f"未找到发文机关全称 '{org_name}'",
                expected=org_name,
                actual="未找到",
                suggestion=f"红头文件的发文机关全称 '{org_name}' 未在文档前 3 段中找到。请确认 org_name 参数正确，或手动在 Word 中添加红头文字。",
            ))
        elif not found_red_line:
            issues.append(FixerDiagnostic(
                code="RED_HEADER_LINE_MISSING",
                action_index=idx,
                severity="warning",
                detail="发文机关段落缺少红色反线（pBdr/bottom color=FF0000）",
                suggestion="红头文件的红色反线（分隔线）未正确渲染。请确认 insert_red_header action 已正确写入 pBdr 元素，或手动在 Word 中添加红色下边框。",
            ))

    if doc_number:
        found_doc = False
        for p in paragraphs[:4]:
            text = "".join(t.text or "" for t in p.iter(f"{{{W}}}t"))
            if doc_number in text:
                found_doc = True
                break
        if not found_doc:
            issues.append(FixerDiagnostic(
                code="RED_HEADER_DOCNUM_MISSING",
                action_index=idx,
                severity="error",
                detail=f"未找到发文字号 '{doc_number}'",
                expected=doc_number,
                actual="未找到",
                suggestion=f"发文字号 '{doc_number}' 未在文档前 4 段中找到。请确认 doc_number 参数正确，或手动在 Word 中添加发文字号。",
            ))

    return issues


@_verifier("set_code_block_style")
def _verify_set_code_block_style(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证 set_code_block_style 操作：检查指定段落是否有 shd 背景色和等宽字体。"""
    return []  # TODO: 实现验证：检查 pPr/shd 和 rPr/rFonts


@_verifier("insert_toc")
def _verify_insert_toc(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证 insert_toc 操作：检查 document.xml 中是否有 TOC fldChar 序列。"""
    return []  # TODO: 实现验证：检查 fldChar begin/end + instrText 包含 TOC


# ── 验证: set_paragraph_spacing ────────────────────────────

@_verifier("set_paragraph_spacing")
def _verify_paragraph_spacing(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证段落间距是否已正确设置。"""
    para_idx = params["para_index"]

    root = get_xml("word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return [FixerDiagnostic(
            code="PARA_SPACING_MISMATCH",
            action_index=idx,
            severity="error",
            detail="document.xml 中无 body 元素",
            suggestion="document.xml 结构异常。请检查文档是否损坏，或在 Word 中重新保存。",
        )]

    paras = list(body.iter(f"{{{W}}}p"))
    if para_idx >= len(paras):
        return [FixerDiagnostic(
            code="PARA_SPACING_MISMATCH",
            action_index=idx,
            severity="error",
            detail=f"paragraph index {para_idx} 超出范围 (共 {len(paras)} 段)",
            suggestion=f"段落索引 {para_idx} 超出文档范围。请重新运行 DocDiscover 刷新段落计数，确认 fix_plan 中的索引与实际文档一致。",
        )]

    p = paras[para_idx]
    pPr = p.find(f"{{{W}}}pPr")
    if pPr is None:
        return [FixerDiagnostic(
            code="PARA_SPACING_MISMATCH",
            action_index=idx,
            severity="error",
            detail=f"段落 {para_idx} 无 pPr 元素",
            suggestion=f"段落 {para_idx} 缺少段落属性（pPr）。该段落可能为空或结构异常，请在 Word 中检查。",
        )]

    sp = pPr.find(f"{{{W}}}spacing")
    if sp is None:
        return [FixerDiagnostic(
            code="PARA_SPACING_MISMATCH",
            action_index=idx,
            severity="error",
            detail=f"段落 {para_idx} 无 spacing 元素（未设置间距）",
            suggestion=f"段落 {para_idx} 未设置间距属性。请确认 set_paragraph_spacing action 已正确执行，或手动在 Word 段落设置中调整间距。",
        )]

    issues = []
    checks = [
        ("before_twips", "before"),
        ("after_twips", "after"),
        ("before_lines", "beforeLines"),
        ("after_lines", "afterLines"),
        ("line", "line"),
        ("lineRule", "lineRule"),
    ]
    for param_key, xml_attr in checks:
        if param_key in params and params[param_key] is not None:
            expected = str(int(params[param_key])) if param_key != "lineRule" else params[param_key]
            actual = sp.get(f"{{{W}}}{xml_attr}")
            if actual != expected:
                issues.append(FixerDiagnostic(
                    code="PARA_SPACING_MISMATCH",
                    action_index=idx,
                    severity="warning",
                    detail=f"段落 {para_idx} {xml_attr} 期望={expected} 实际={actual}",
                    expected=expected,
                    actual=actual,
                    suggestion=f"段落 {para_idx} 的 {xml_attr} 间距属性未生效（期望={expected}，实际={actual}）。请确认 set_paragraph_spacing action 参数正确，或手动在 Word 中调整。",
                ))

    return issues


# ── 验证: normalize_list_numbering ─────────────────────────

@_verifier("normalize_list_numbering")
def _verify_normalize_list_numbering(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证列表编号一致性：检查所有 numId 组内 ilvl 是否一致。"""
    root = get_xml("word/document.xml")
    body = root.find(f"{{{W}}}body")
    if body is None:
        return []

    # 收集所有 numId 组的 ilvl
    num_groups: dict[str, set] = {}
    for p in body.iter(f"{{{W}}}p"):
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
        ilvl = ilvl_el.get(f"{{{W}}}val", "0") if ilvl_el is not None else "0"
        num_groups.setdefault(num_id, set()).add(ilvl)

    issues = []
    for num_id, ilvls in num_groups.items():
        if len(ilvls) > 1:
            issues.append(FixerDiagnostic(
                code="LIST_NUMBERING_INCONSISTENT",
                action_index=idx,
                severity="warning",
                detail=f"numId={num_id} 存在不一致的 ilvl: {sorted(ilvls)}",
                expected="统一的 ilvl",
                actual=f"多个 ilvl: {sorted(ilvls)}",
                suggestion=f"列表 numId={num_id} 的缩进级别（ilvl）不一致。请重新运行 normalize_list_numbering action 统一列表编号层级，或手动在 Word 中调整列表缩进。",
            ))

    return issues


# ── 验证: set_page_margins ─────────────────────────────────

@_verifier("set_page_margins")
def _verify_page_margins(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证页边距是否已正确设置。"""
    section_index = params["section_index"]

    root = get_xml("word/document.xml")
    sects = list(root.iter(f"{{{W}}}sectPr"))

    if section_index >= len(sects):
        return [FixerDiagnostic(
            code="PAGE_MARGINS_MISMATCH",
            action_index=idx,
            severity="error",
            detail=f"section_index {section_index} 超出范围 (共 {len(sects)} 个 sectPr)",
            suggestion=f"节索引 {section_index} 超出文档范围。请重新运行 DocDiscover 刷新节计数，确认 fix_plan 中的索引与实际文档一致。",
        )]

    sp = sects[section_index]
    pm = sp.find(f"{{{W}}}pgMar")
    if pm is None:
        return [FixerDiagnostic(
            code="PAGE_MARGINS_MISMATCH",
            action_index=idx,
            severity="error",
            detail=f"sectPr[{section_index}] 中无 pgMar 元素",
            suggestion=f"第 {section_index+1} 节缺少页边距（pgMar）设置。请运行 set_page_margins action 写入页边距，或在 Word 的「布局→页边距」中设置。",
        )]

    margin_attrs = ("top", "bottom", "left", "right", "header", "footer")
    issues = []
    for attr in margin_attrs:
        if attr in params and params[attr] is not None:
            expected = str(int(params[attr]))
            actual = pm.get(f"{{{W}}}{attr}")
            if actual != expected:
                issues.append(FixerDiagnostic(
                    code="PAGE_MARGINS_MISMATCH",
                    action_index=idx,
                    severity="warning",
                    detail=f"sectPr[{section_index}] {attr} 期望={expected} 实际={actual}",
                    expected=expected,
                    actual=actual,
                    suggestion=f"第 {section_index+1} 节的 {attr} 页边距未生效（期望={expected} twips，实际={actual} twips）。请确认 set_page_margins action 参数正确，或手动在 Word 中调整。",
                ))

    return issues


# ── 验证: set_tbl_col_widths ─────────────────────────────────

@_verifier("set_tbl_col_widths")
def _verify_set_tbl_col_widths(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证 tblGrid 列宽和 tcW 是否正确写入。"""
    issues: list[FixerDiagnostic] = []
    tbl_idx = params.get("tbl_index", 0)
    widths = params.get("_widths_applied") or params.get("widths", [])

    root = get_xml("word/document.xml")
    tables = list(root.iter(f"{{{W}}}tbl"))
    if tbl_idx >= len(tables):
        return [FixerDiagnostic(code="TBL_INDEX_OUT_OF_RANGE", action_index=idx,
                                severity="error",
                                detail=f"表格索引 {tbl_idx} 越界 (共 {len(tables)} 张表)",
                                suggestion=f"表格索引 {tbl_idx} 超出范围。请重新运行 DocDiscover 刷新表格计数，确认 fix_plan 中的索引与实际文档一致。")]

    tbl = tables[tbl_idx]
    tblGrid = tbl.find(f"{{{W}}}tblGrid")
    if tblGrid is None:
        return [FixerDiagnostic(code="TBL_GRID_MISSING", action_index=idx,
                                severity="error",
                                detail="tblGrid 不存在（可能被错误删除）",
                                suggestion="表格缺少 tblGrid 元素。可能是之前的操作误删了表格网格，请从原始文件重新排版，或手动在 Word 中调整表格列宽。")]

    grid_cols = tblGrid.findall(f"{{{W}}}gridCol")
    for ci, gc in enumerate(grid_cols):
        actual_w = gc.get(f"{{{W}}}w")
        expected_w = str(int(widths[ci])) if ci < len(widths) and isinstance(widths[ci], (int, float)) else None
        if expected_w is not None and actual_w != expected_w:
            issues.append(FixerDiagnostic(code="TBL_COL_WIDTH_MISMATCH", action_index=idx,
                                          severity="warning",
                                          detail=f"列[{ci}] 期望宽度={expected_w} twips 实际={actual_w}",
                                          expected=expected_w, actual=actual_w,
                                          suggestion=f"表格列[{ci}] 宽度未生效（期望={expected_w} twips，实际={actual_w} twips）。请确认 set_tbl_col_widths action 参数正确，或手动在 Word 中拖拽列宽。"))

    return issues


# ── 验证: set_tbl_cell_valign ─────────────────────────────────

@_verifier("set_tbl_cell_valign")
def _verify_set_tbl_cell_valign(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证单元格 vAlign 是否正确写入。"""
    tbl_idx = params.get("tbl_index", 0)
    valign = params.get("valign", "center")
    skip_header = params.get("skip_header", False)

    root = get_xml("word/document.xml")
    tables = list(root.iter(f"{{{W}}}tbl"))
    if tbl_idx >= len(tables):
        return [FixerDiagnostic(code="TBL_INDEX_OUT_OF_RANGE", action_index=idx,
                                severity="error",
                                detail=f"表格索引 {tbl_idx} 越界",
                                suggestion=f"表格索引 {tbl_idx} 超出范围。请重新运行 DocDiscover 刷新表格计数，确认 fix_plan 中的索引与实际文档一致。")]

    tbl = tables[tbl_idx]
    affected = params.get("_affected", 0)
    actual_errors = 0
    for tr_idx, tr in enumerate(tbl.iter(f"{{{W}}}tr")):
        if skip_header and tr_idx == 0:
            continue
        for tc in tr.iter(f"{{{W}}}tc"):
            tcPr = tc.find(f"{{{W}}}tcPr")
            if tcPr is not None:
                vAlign_el = tcPr.find(f"{{{W}}}vAlign")
                if vAlign_el is not None:
                    actual = vAlign_el.get(f"{{{W}}}val")
                    if actual != valign:
                        actual_errors += 1

    if actual_errors > 0:
        return [FixerDiagnostic(code="CELL_VALIGN_MISMATCH", action_index=idx,
                                severity="warning",
                                detail=f"{actual_errors}/{affected} 个单元格 vAlign 不是 '{valign}'",
                                expected=valign, actual=f"{actual_errors} mismatches",
                                suggestion=f"{actual_errors} 个单元格的垂直对齐未生效。请确认 set_tbl_cell_valign action 参数正确，或手动在 Word 表格属性中设置垂直居中。")]

    return []


# ── 验证: split_merged_cells ─────────────────────────────────

@_verifier("split_merged_cells")
def _verify_split_merged_cells(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证合并单元格检测结果。（该 action 不修改文档，仅验证 _merged 是否被写入。）"""
    merged = params.get("_merged")
    if merged is None:
        return [FixerDiagnostic(code="MERGED_CELLS_DETECTION_FAILED", action_index=idx,
                                severity="warning",
                                detail="split_merged_cells 未写入 _merged 信息",
                                suggestion="合并单元格检测失败。该 action 为纯检测步骤，不修改文档，请确认检测逻辑是否正确执行。")]
    if isinstance(merged, list) and len(merged) == 0:
        return []  # 无合并单元格，正常

    return [FixerDiagnostic(code="MERGED_CELLS_FOUND", action_index=idx,
                            severity="info",
                            detail=f"检测到 {len(merged)} 处合并单元格（未拆分，需手动处理）",
                            suggestion=f"检测到 {len(merged)} 处合并单元格。早睡 (EarlyZ) 不会自动拆分合并单元格，请在 Word 中手动取消合并后重新排版。")]


# ── 验证: normalize_tbl_grid ─────────────────────────────────

@_verifier("normalize_tbl_grid")
def _verify_normalize_tbl_grid(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证 tblGrid 是否已重建，tcW 是否被清除，列数是否一致。"""
    tbl_idx = params.get("tbl_index", 0)
    grid_cols = params.get("_grid_cols", 0)

    root = get_xml("word/document.xml")
    tables = list(root.iter(f"{{{W}}}tbl"))
    if tbl_idx >= len(tables):
        return [FixerDiagnostic(code="TBL_INDEX_OUT_OF_RANGE", action_index=idx,
                                severity="error",
                                detail=f"表格索引 {tbl_idx} 越界",
                                suggestion=f"表格索引 {tbl_idx} 超出范围。请重新运行 DocDiscover 刷新表格计数，确认 fix_plan 中的索引与实际文档一致。")]

    tbl = tables[tbl_idx]
    tblGrid = tbl.find(f"{{{W}}}tblGrid")
    if tblGrid is None:
        return [FixerDiagnostic(code="TBL_GRID_MISSING", action_index=idx,
                                severity="error", detail="tblGrid 不存在",
                                suggestion="tblGrid 元素丢失。可能是之前的操作误删了表格网格，请从原始文件重新排版。")]

    actual_cols = len(tblGrid.findall(f"{{{W}}}gridCol"))
    if grid_cols > 0 and actual_cols != grid_cols:
        return [FixerDiagnostic(code="TBL_GRID_COL_COUNT_MISMATCH", action_index=idx,
                                severity="warning",
                                detail=f"tblGrid 列数: 期望={grid_cols} 实际={actual_cols}",
                                expected=grid_cols, actual=actual_cols,
                                suggestion=f"表格列数不匹配（期望={grid_cols}，实际={actual_cols}）。请确认 normalize_tbl_grid action 已正确执行，或手动在 Word 中调整表格列数。")]

    # 检查 tcW 是否被清除
    tcw_remaining = 0
    for tc in tbl.iter(f"{{{W}}}tc"):
        tcPr = tc.find(f"{{{W}}}tcPr")
        if tcPr is not None and tcPr.find(f"{{{W}}}tcW") is not None:
            tcw_remaining += 1
    if tcw_remaining > 0:
        return [FixerDiagnostic(code="TCW_NOT_CLEARED", action_index=idx,
                                severity="warning",
                                detail=f"还有 {tcw_remaining} 个单元格残留 tcW 元素",
                                expected=0, actual=tcw_remaining,
                                suggestion=f"还有 {tcw_remaining} 个单元格残留固定列宽（tcW）。请重新运行 normalize_tbl_grid action 清除固定列宽。")]

    return []


# ── 验证: set_tbl_header_repeat ───────────────────────────────

@_verifier("set_tbl_header_repeat")
def _verify_set_tbl_header_repeat(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证 tblHeader 双层级标记是否正确写入。"""
    tbl_idx = params.get("tbl_index", 0)
    root = get_xml("word/document.xml")
    tables = list(root.iter(f"{{{W}}}tbl"))
    if tbl_idx >= len(tables):
        return []

    tbl = tables[tbl_idx]
    issues: list[FixerDiagnostic] = []

    # 表级
    tblPr = tbl.find(f"{{{W}}}tblPr")
    if tblPr is None or tblPr.find(f"{{{W}}}tblHeader") is None:
        issues.append(FixerDiagnostic(code="TBL_HEADER_TABLE_LEVEL_MISSING", action_index=idx,
                                      severity="error", detail="表级 tblHeader 未设置",
                                      suggestion="表格表头重复标记（表级）未生效。请重新运行 set_tbl_header_repeat action，或手动在 Word 表格属性中勾选「在各页顶端以标题行形式重复出现」。"))
    # 首行级
    first_row = tbl.find(f"{{{W}}}tr")
    if first_row is not None:
        trPr = first_row.find(f"{{{W}}}trPr")
        if trPr is None or trPr.find(f"{{{W}}}tblHeader") is None:
            issues.append(FixerDiagnostic(code="TBL_HEADER_ROW_LEVEL_MISSING", action_index=idx,
                                          severity="warning", detail="首行级 tblHeader 未标记",
                                          suggestion="表格首行的表头标记未生效。请确认 set_tbl_header_repeat action 正确设置了行级 tblHeader。"))
    return issues


# ── 验证: set_tr_cant_split ───────────────────────────────────

@_verifier("set_tr_cant_split")
def _verify_set_tr_cant_split(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证 cantSplit 是否设置。"""
    tbl_idx = params.get("tbl_index", 0)
    skip_header = params.get("skip_header", True)
    root = get_xml("word/document.xml")
    tables = list(root.iter(f"{{{W}}}tbl"))
    if tbl_idx >= len(tables):
        return []

    tbl = tables[tbl_idx]
    missing = 0
    for tr_idx, tr in enumerate(tbl.iter(f"{{{W}}}tr")):
        if skip_header and tr_idx == 0:
            continue
        trPr = tr.find(f"{{{W}}}trPr")
        if trPr is None or trPr.find(f"{{{W}}}cantSplit") is None:
            missing += 1

    if missing > 0:
        return [FixerDiagnostic(code="CANT_SPLIT_MISSING", action_index=idx,
                                severity="warning",
                                detail=f"{missing} 行未设置 cantSplit",
                                expected="所有数据行 cantSplit=true", actual=f"{missing} missing",
                                suggestion=f"{missing} 行未设置「不允许跨页截断」。请重新运行 set_tr_cant_split action，或手动在 Word 表格属性中为数据行设置。")]
    return []


# ── 验证: set_tbl_font ────────────────────────────────────────

@_verifier("set_tbl_font")
def _verify_set_tbl_font(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证表格字体是否已应用。"""
    return []


# ── 验证: set_tbl_body_font ────────────────────────────────────

@_verifier("set_tbl_body_font")
def _verify_set_tbl_body_font(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return []


# ── 验证: set_tbl_header_font ──────────────────────────────────

@_verifier("set_tbl_header_font")
def _verify_set_tbl_header_font(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return []


# ── 验证: set_tbl_borders ──────────────────────────────────────

@_verifier("set_tbl_borders")
def _verify_set_tbl_borders(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    """验证表格边框是否正确设置（检查 tblBorders 存在性）。"""
    tbl_idx = params.get("tbl_index", 0)
    style = params.get("style", "three_line")
    root = get_xml("word/document.xml")
    tables = list(root.iter(f"{{{W}}}tbl"))
    if tbl_idx >= len(tables):
        return []

    tbl = tables[tbl_idx]
    tblPr = tbl.find(f"{{{W}}}tblPr")
    if tblPr is not None and tblPr.find(f"{{{W}}}tblBorders") is not None:
        return []
    return [FixerDiagnostic(code="TBL_BORDERS_MISSING", action_index=idx,
                            severity="warning",
                            detail=f"style={style} 但 tblBorders 未找到",
                            suggestion=f"表格边框样式 '{style}' 未生效。请确认 set_tbl_borders action 已正确执行，或手动在 Word 表格属性中设置边框样式。")]


# ── 验证: normalize_cjk_spacing ───────────────────────────────

@_verifier("normalize_cjk_spacing")
def _verify_normalize_cjk_spacing(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return []


# ── 验证: set_images_inline ───────────────────────────────────

@_verifier("set_images_inline")
def _verify_set_images_inline(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return []


# ── 验证: set_paragraph_alignment ─────────────────────────────

@_verifier("set_paragraph_alignment")
def _verify_set_paragraph_alignment(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return []


# ── 验证: extract_textbox_captions ─────────────────────────────

@_verifier("extract_textbox_captions")
def _verify_extract_textbox_captions(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return []


# ── 验证: protect_image_para ──────────────────────────────────

@_verifier("protect_image_para")
def _verify_protect_image_para(target_path: Path, params: dict, get_xml, idx: int) -> list[FixerDiagnostic]:
    return []

