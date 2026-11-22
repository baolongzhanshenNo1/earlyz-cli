"""源文档卫生检查：找出「源稿里本来就有的模板说明文字」。

背景（2026-10-27 交大本科）：用户的「裸内容」源文档其实是模板副本，正文里混着
「三号黑体居中，上下各空一行」「空一行空一行」「摘要正文五号宋体…」这类写作说明。
产品铁律是**只改格式不动内容**，所以引擎原样保留是对的；但**不提示**就是失职 ——
用户看到产物会以为排版没做对。

本模块只做检测与清单，不删除。删除链路必须是：
    检测 → 列清单 → 用户确认 → 删除（动作 remove_example_residuals）
"""
from __future__ import annotations

from dataclasses import dataclass

from earlyz.spec.annotations import is_instruction


@dataclass
class InstructionHit:
    index: int
    text: str
    role: str | None
    reason: str

    def as_dict(self) -> dict:
        return {"index": self.index, "text": self.text, "role": self.role, "reason": self.reason}


def detect_instruction_paragraphs(ir, limit: int = 200) -> dict:
    """扫描已解析的 DocumentIR，返回说明文字清单。

    只认「明显在讲版式」的句子（字号+字体 / 缩进 / 行距 / 空一行 / 页边距），
    正常论文正文几乎不会命中。
    """
    hits: list[InstructionHit] = []
    for i, p in enumerate(ir.paragraphs):
        text = (p.text or "").strip()
        if not text or len(text) > 400:
            continue
        if not is_instruction(text):
            continue
        hits.append(InstructionHit(
            index=i, text=text[:160], role=getattr(p, "semantic_role", None),
            reason="含版式说明用语（字号/字体/缩进/行距/空一行）",
        ))
        if len(hits) >= limit:
            break

    return {
        "count": len(hits),
        "items": [h.as_dict() for h in hits],
        "samples": [h.text for h in hits[:10]],
        "note": (
            "源文档自带模板说明文字，引擎按「只改格式不动内容」原样保留。"
            "如确认要删，请先看清单再执行删除。"
        ) if hits else "",
    }


def summarize(result: dict) -> str:
    n = result.get("count", 0)
    if not n:
        return "源文档未发现模板说明文字"
    return f"源文档含 {n} 处模板说明文字（未删除，仅在诊断包里列出）"
