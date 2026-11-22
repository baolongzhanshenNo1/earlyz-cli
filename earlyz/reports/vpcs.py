"""格式检测报告解析（M1，2026-10-07 边界定稿）。

首批平台 = 维普（泛语 vpcs.fanyu.com）「格式批注 docx」——用户用平台查完
格式后导出的批注版论文：正文 + Word 批注（word/comments.xml）。

- 批注 0 = 检测结果摘要（检测模板 / 错误数 / 差错率 / 严重·一般·提醒计数）
- 批注 1..N = 逐条问题：`<组>-<子类…>-<问题名>：<现状>，规范：<规范>`

产出统一 IR（dict，JSON 安全）供 read_format_report 工具与修复/复核链路使用。
铁律：解析失败/字段缺失 → parse_warnings 显式记录，不猜、不吞。
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

PLATFORM_ID = "vpcs"
PLATFORM_LABEL = "维普格式检测（泛语）"

# ── 问题组 → 引擎处置能力（M1 首批；未收录 → unsupported 只报告）──
# capability: auto_fix（自动修复）/ check（引擎核查+报告）/ unsupported（仅说明）


class ReportParseError(Exception):
    """报告解析失败（格式不识别/结构损坏）——显式报错，不猜。"""


def _comment_text(c_el) -> str:
    parts: list[str] = []
    for pel in c_el.iter(f"{W}p"):
        parts.append("".join(t.text or "" for t in pel.iter(f"{W}t")))
    return "".join(parts).strip()


def _read_comments(z: zipfile.ZipFile):
    if "word/comments.xml" not in z.namelist():
        return None
    try:
        return etree.fromstring(z.read("word/comments.xml"))
    except etree.XMLSyntaxError as e:  # noqa: PERF203
        raise ReportParseError(f"comments.xml 损坏: {e}") from e


def _parse_summary(text: str) -> dict:
    """批注 0：检测结果摘要。取不到 → None 字段（不猜）。"""
    out: dict = {
        "template_name": None,
        "error_count": None,
        "rate": None,
        "serious": None,
        "normal": None,
        "reminder": None,
    }
    m = re.search(r"检测模板[：:]\s*[《]?([^》\n]+?)[》]?(?=检测错误数|$)", text)
    if m:
        out["template_name"] = m.group(1).strip()
    m = re.search(r"检测错误数[：:]\s*(\d+)", text)
    if m:
        out["error_count"] = int(m.group(1))
    m = re.search(r"差错率[：:]\s*([\d.]+)", text)
    if m:
        out["rate"] = m.group(1)
    for key, label in (("serious", "严重错误"), ("normal", "一般错误"), ("reminder", "提醒")):
        m = re.search(rf"{label}[：:]\s*(\d+)", text)
        if m:
            out[key] = int(m.group(1))
    return out


_ISSUE_RE = re.compile(r"^(?P<head>[^：:]{4,80})[：:](?P<rest>.+)$")


def _parse_issue(text: str) -> dict:
    """单条问题批注 → 结构化。解析不了的字段留 None + 记录 raw。"""
    issue: dict = {
        "group": None,
        "sub": None,
        "name": None,
        "severity": "一般错误",
        "current": None,
        "standard": None,
        "raw": text,
    }
    m = _ISSUE_RE.match(text)
    if not m:
        return issue
    head, rest = m.group("head"), m.group("rest")
    parts = [p.strip() for p in head.split("-") if p.strip()]
    if parts:
        issue["group"] = parts[0]
        issue["name"] = parts[-1]
        if len(parts) > 2:
            issue["sub"] = ".".join(parts[1:-1])
    if "严重" in head or "严重错误" in text:
        issue["severity"] = "严重错误"
    elif "提醒" in head:
        issue["severity"] = "提醒"
    # 「，规范：」右侧=规范；两侧都可能含逗号 → 从右取一次切分
    if "，规范：" in rest:
        cur, std = rest.rsplit("，规范：", 1)
        issue["current"] = cur.strip()
        issue["standard"] = std.strip()
    else:
        issue["current"] = rest.strip()
    return issue


def _extract_anchors(doc_root) -> dict[int, str]:
    """批注 id → 锚点文本。

    commentRangeStart/End 范围内文本优先；仅有 commentReference（点锚）时
    取所在段落文本。超长截断 120 字。
    """
    anchors: dict[int, str] = {}
    active: dict[int, list[str]] = {}
    for el in doc_root.iter():
        tag = el.tag
        if tag == f"{W}commentRangeStart":
            try:
                active[int(el.get(f"{W}id"))] = []
            except (TypeError, ValueError):
                continue
        elif tag == f"{W}commentRangeEnd":
            try:
                cid = int(el.get(f"{W}id"))
            except (TypeError, ValueError):
                continue
            if cid in active:
                anchors[cid] = "".join(active.pop(cid))
        elif tag == f"{W}t":
            txt = el.text or ""
            if txt and active:
                for cid in active:
                    active[cid].append(txt)
        elif tag == f"{W}commentReference":
            try:
                cid = int(el.get(f"{W}id"))
            except (TypeError, ValueError):
                continue
            if cid in anchors:
                continue
            p = el.getparent()
            while p is not None and p.tag != f"{W}p":
                p = p.getparent()
            if p is not None:
                anchors[cid] = "".join(t.text or "" for t in p.iter(f"{W}t"))
    for cid, parts in active.items():
        anchors.setdefault(cid, "".join(parts))
    return {
        cid: re.sub(r"\s+", " ", txt).strip()[:120]
        for cid, txt in anchors.items()
    }


def detect(path: str | Path) -> dict | None:
    """轻量检测：文件是否是维普格式检测批注 docx。

    任何异常 → None（本函数在文件上传链路上调用，绝不抛出）。
    返回 {platform, platform_label, issue_comment_count} 或 None。
    """
    try:
        p = Path(path)
        if p.suffix.lower() != ".docx" or not p.exists():
            return None
        with zipfile.ZipFile(p) as z:
            root = _read_comments(z)
            if root is None:
                return None
            comments = list(root.iter(f"{W}comment"))
            if not comments:
                return None
            texts = [_comment_text(c) for c in comments]
            joined = "\n".join(texts)
            # 维普签名：批注含「，规范：」逐条格式 + 摘要含「检测结果/检测模板」
            if "，规范：" not in joined:
                return None
            if not any(("检测结果" in t or "检测模板" in t or "检测错误数" in t) for t in texts):
                return None
            return {
                "platform": PLATFORM_ID,
                "platform_label": PLATFORM_LABEL,
                "issue_comment_count": len(comments) - 1,
            }
    except Exception:  # noqa: BLE001 —— 检测链路绝不因异常中断
        return None


def parse(path: str | Path) -> dict:
    """完整解析维普批注 docx → 统一 IR。

    Returns dict:
      {platform, platform_label, source_file, template_name,
       summary{...}, issues[{id, group, sub, name, severity,
                            current, standard, anchor, raw}],
       parse_warnings[...]}
    Raises ReportParseError：文件不是可识别的维普批注 docx。
    """
    p = Path(path)
    if not p.exists():
        raise ReportParseError(f"文件不存在: {p}")
    with zipfile.ZipFile(p) as z:
        c_root = _read_comments(z)
        if c_root is None:
            raise ReportParseError("文件不含批注（word/comments.xml）——不是检测报告批注版")
        try:
            doc_root = etree.fromstring(z.read("word/document.xml"))
        except KeyError as e:
            raise ReportParseError(f"docx 包缺少 document.xml: {e}") from e

    comments = list(c_root.iter(f"{W}comment"))
    if not comments:
        raise ReportParseError("批注为空——不是检测报告批注版")

    anchors = _extract_anchors(doc_root)
    warnings: list[str] = []

    summary_text = None
    issues: list[dict] = []
    for c in comments:
        try:
            cid = int(c.get(f"{W}id"))
        except (TypeError, ValueError):
            cid = -1
        text = _comment_text(c)
        if not text:
            continue
        if summary_text is None and ("检测结果" in text or "检测错误数" in text
                                     or "检测模板" in text):
            summary_text = text
            continue
        issue = _parse_issue(text)
        issue["id"] = cid
        issue["anchor"] = anchors.get(cid, "")
        if issue["group"] is None or issue["current"] is None:
            warnings.append(f"批注 {cid} 结构不完整（保留 raw 原文）")
        if issue["standard"] is None:
            warnings.append(f"批注 {cid} 未给出「规范」值")
        issues.append(issue)

    summary = _parse_summary(summary_text or "")
    if summary_text is None:
        warnings.append("未找到检测结果摘要批注（批注 0）")

    return {
        "platform": PLATFORM_ID,
        "platform_label": PLATFORM_LABEL,
        "source_file": str(p),
        "template_name": summary.get("template_name"),
        "summary": summary,
        "summary_raw": summary_text,
        "issues": issues,
        "parse_warnings": warnings,
    }
