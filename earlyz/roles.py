"""角色探测：判断一份 .docx 内容结构更像「排版规范（模板）」还是「待排版文档」。

2026-10-10「发送时角色判定方案」的引擎侧实现（方案文档：
D:/Desktop/EarlyZ/EarlyZ-发送时角色判定方案-2026-10-10.md）：

- 用途：文件名无信号时的兜底判定信号之一。前端在发送时合并各信号
  （显式动作 / 指纹 / 消息话语 / 本探测 / 文件名）后，决定「直接发 / 问一次 / 拦一次」。
- 只做「建议」：输出 role_hint/strength + evidence；不落盘、不写任何状态。
- fail-open：任何异常 → ok=False / unknown / none（调用方回退既有行为，绝不阻塞上传或发送）。

阈值以真实文件校准（2026-10-10 实测，见 evidence 字段同名计数）：
    本科论文模板  2877字 / 14占位 / 53样式标注 / 0域  → template/strong
    期刊模板      3463字 /  2占位 / 49样式标注 / 0域  → template/strong
    真论文       31498字 /  1占位 /  1样式标注 / 247域 → document/strong
    格式批注报告  31749字 /  1占位 /  1样式标注 / 298域 → document/strong
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

__all__ = ["probe_role"]

# 占位/说明句：模板里大量出现、正文里几乎不出现
_PLACEHOLDER_RE = re.compile(
    r"此处填写|请填写|请输入|填写说明|本模板|示例[:：]|说明[:：]|注意[:：]"
    r"|（说明|\(说明|请按|请勿|删除本|删去本|按照本"
)
# 样式标注句：模板的「字体为2号黑体」「1号宋体居中」这类排版指示
_STYLE_ANNOT_RE = re.compile(
    r"字体为|字号|号黑体|号宋体|号楷体|号仿宋|Times New Roman|居中|倍行距|行距|段前|段后"
    r"|页边距|页码|页眉|小四|小五|五号|四号|三号|二号|一号|黑体|宋体|仿宋|楷体|加粗|斜体|对齐|缩进|首行"
)
_WT_RE = re.compile(r"<w:t[^>]*>([^<]*)</w:t>")
_HEADING_STYLE_RE = re.compile(r'<w:pStyle w:val="(?:[^"]*?(?:Heading|heading|标题|TOC)[^"]*)"')
_OUTLINE_RE = re.compile(r'<w:outlineLvl w:val="[0-8]"')
_INSTR_RE = re.compile(r"<w:instrText")
_FLD_RE = re.compile(r'<w:fldChar w:fldCharType="begin"')

_TEMPLATE_STRONG = 3
_TEMPLATE_WEAK = 1
_DOCUMENT_WEAK = -1
_DOCUMENT_STRONG = -3


def _classify(document_xml: str) -> tuple[str, str, dict]:
    """纯统计分类（不做 IO，便于单测）：(role_hint, strength, evidence)。"""
    text = "".join(_WT_RE.findall(document_xml))
    chars = len(re.sub(r"\s+", "", text))
    placeholder_hits = len(_PLACEHOLDER_RE.findall(text))
    style_annotations = len(_STYLE_ANNOT_RE.findall(text))
    fields = len(_INSTR_RE.findall(document_xml)) + len(_FLD_RE.findall(document_xml))
    heading_paras = len(_HEADING_STYLE_RE.findall(document_xml)) + len(_OUTLINE_RE.findall(document_xml))

    score = 0
    # 正文体量：真论文几万字；模板多为骨架
    if chars < 2500:
        score += 2
    elif chars < 8000:
        score += 1
    elif chars >= 30000:
        score -= 2
    elif chars >= 12000:
        score -= 1
    # 说明/标注密度：模板的核心特征
    instruction_hits = placeholder_hits + style_annotations
    if instruction_hits >= 10:
        score += 2
    elif instruction_hits >= 4:
        score += 1
    # 域（目录/交叉引用）密集 → 正文特征
    if fields >= 20:
        score -= 1
    # 标题骨架 + 少量占位 → 轻微模板倾向
    if heading_paras >= 8 and chars < 12000 and placeholder_hits >= 1:
        score += 1

    if score >= _TEMPLATE_STRONG:
        role_hint, strength = "template", "strong"
    elif score >= _TEMPLATE_WEAK:
        role_hint, strength = "template", "weak"
    elif score <= _DOCUMENT_STRONG:
        role_hint, strength = "document", "strong"
    elif score <= _DOCUMENT_WEAK:
        role_hint, strength = "document", "weak"
    else:
        role_hint, strength = "unknown", "none"

    evidence = {
        "text_chars": chars,
        "placeholder_hits": placeholder_hits,
        "style_annotations": style_annotations,
        "fields": fields,
        "heading_paras": heading_paras,
        "score": score,
    }
    return role_hint, strength, evidence


def probe_role(path: str | Path) -> dict:
    """探测一份文件更像模板还是正文。

    返回 {ok, role_hint: template|document|unknown, strength: strong|weak|none, evidence}。
    异常安全：任何失败 → ok=False / unknown / none（fail-open，调用方不因此阻塞）。
    """
    p = Path(path)
    if p.suffix.lower() != ".docx":
        return {"ok": True, "role_hint": "unknown", "strength": "none",
                "evidence": {"reason": "non-docx"}}
    try:
        with zipfile.ZipFile(p) as z:
            document_xml = z.read("word/document.xml").decode("utf-8", errors="ignore")
    except Exception as e:  # noqa: BLE001 —— 探测链路绝不抛出
        return {"ok": False, "role_hint": "unknown", "strength": "none",
                "evidence": {"error": str(e)}}
    role_hint, strength, evidence = _classify(document_xml)
    return {"ok": True, "role_hint": role_hint, "strength": strength, "evidence": evidence}
