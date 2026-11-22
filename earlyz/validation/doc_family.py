"""文档族判定：决定是否启用「续表」等只在特定文类成立的排版规则。

2026-10-21 用户拍板：续表只对**学位论文**和**期刊**开放——
这两类数据表通常长、跨页常见；合同/公文的数据表一般短，开启反而带来无谓断表风险。

判定只看模板正文结构（不看文件名、不硬编码学校）：
  · thesis（学位论文）：有「目录」+（摘要 或 Abstract）
  · journal（期刊）  ：有（摘要 或 Abstract）+（关键词 或 Key words），但没有目录
  · other            ：其余（合同/公文/标书/表单等）

被判定为 thesis/journal 时才返回 CONTINUED_LABEL，其余返回空字符串 = 不启用。
"""
from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

_TOC_RE = ("目录", "目 录", "Contents", "CONTENTS")
_ABS_RE = ("摘要", "摘 要", "Abstract", "ABSTRACT")
_KW_RE = ("关键词", "关键字", "Key words", "Keywords", "KEY WORDS", "Index Terms")

CONTINUED_LABEL = "续表"
CONTINUED_FAMILIES = ("thesis", "journal")


def _template_text(template_path: Path) -> str:
    try:
        with zipfile.ZipFile(template_path) as z:
            root = etree.fromstring(z.read("word/document.xml"))
    except Exception:
        return ""
    return "\n".join("".join(t.text or "" for t in p.iter(f"{W}t"))
                     for p in root.iter(f"{W}p"))


def detect_doc_family(template_path: str | Path) -> str:
    """返回 "thesis" / "journal" / "other"。"""
    path = Path(template_path)
    if not path.exists() or path.suffix.lower() != ".docx":
        return "other"
    text = _template_text(path)
    if not text:
        return "other"
    # 全角空格/制表符会把「目　录」「摘　要」拆开，先压掉全部空白再匹配
    compact = "".join(text.split())
    has_toc = any("".join(k.split()) in compact for k in _TOC_RE)
    has_abs = any("".join(k.split()) in compact for k in _ABS_RE)
    has_kw = any("".join(k.split()) in compact for k in _KW_RE)
    if has_toc and has_abs:
        return "thesis"
    if has_abs and has_kw and not has_toc:
        return "journal"
    return "other"


def continued_label_for(template_path: str | Path) -> str:
    """该模板是否应启用续表：是则返回标签，否则空串。"""
    return CONTINUED_LABEL if detect_doc_family(template_path) in CONTINUED_FAMILIES else ""
