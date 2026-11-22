"""早睡 (EarlyZ) 3.2 — 复杂对象分析（Phase 2 脚注/尾注/修订/批注/书签，§7.2 / §11）。

只读分析 + 交付门禁检查：
- word/footnotes.xml / word/endnotes.xml 存在性 + 引用完整性
  （document.xml 中每个 footnoteReference/endnoteReference 的 id
   必须在对应 part 中有定义，否则 Word 打开即报错）
- 修订标记（w:ins / w:del / w:rPrChange 等）与批注（commentReference）
  —— 排版前未接受修订会导致格式修改与修订叠加，交付时作为 warning
- 书签（bookmarkStart/bookmarkEnd）配对检查
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Optional


def analyze_complex_objects(docx_path: str | Path) -> dict:
    """返回复杂对象统计与缺陷（供 IR metadata / 交付门禁使用）。"""
    src = Path(docx_path)
    with zipfile.ZipFile(str(src)) as z:
        names = set(z.namelist())
        document_xml = z.read("word/document.xml").decode("utf-8", errors="replace")
        footnotes_xml = (
            z.read("word/footnotes.xml").decode("utf-8", errors="replace")
            if "word/footnotes.xml" in names else None
        )
        endnotes_xml = (
            z.read("word/endnotes.xml").decode("utf-8", errors="replace")
            if "word/endnotes.xml" in names else None
        )

    result: dict = {
        "footnotes": {"part_exists": footnotes_xml is not None, "defined": 0, "refs": 0, "missing_refs": []},
        "endnotes": {"part_exists": endnotes_xml is not None, "defined": 0, "refs": 0, "missing_refs": []},
        "revisions": {"tracked_changes": 0, "has_tracked_changes": False},
        "comments": {"comment_refs": 0, "has_comments": False},
        "bookmarks": {"pairs": 0, "broken": []},
        "issues": [],
    }

    if footnotes_xml is not None:
        ids = set(int(m) for m in re.findall(r'<w:footnote[^>]*w:id="(-?\d+)"', footnotes_xml))
        refs = [int(m) for m in re.findall(r'<w:footnoteReference[^>]*w:id="(\d+)"', document_xml)]
        result["footnotes"]["defined"] = len(ids)
        result["footnotes"]["refs"] = len(refs)
        missing = [r for r in refs if r not in ids]
        result["footnotes"]["missing_refs"] = missing
        if missing:
            result["issues"].append(
                f"脚注引用缺失: {len(missing)} 处 footnoteReference 指向不存在的脚注定义（Word 可能报损坏）"
            )

    if endnotes_xml is not None:
        ids = set(int(m) for m in re.findall(r'<w:endnote[^>]*w:id="(-?\d+)"', endnotes_xml))
        refs = [int(m) for m in re.findall(r'<w:endnoteReference[^>]*w:id="(\d+)"', document_xml)]
        result["endnotes"]["defined"] = len(ids)
        result["endnotes"]["refs"] = len(refs)
        missing = [r for r in refs if r not in ids]
        result["endnotes"]["missing_refs"] = missing
        if missing:
            result["issues"].append(
                f"尾注引用缺失: {len(missing)} 处 endnoteReference 指向不存在的尾注定义"
            )

    # 修订标记（ins/del/moveFrom/moveTo/rPrChange/pPrChange）
    n_rev = len(re.findall(r'<w:(ins|del|moveFrom|moveTo|rPrChange|pPrChange)\b', document_xml))
    if n_rev:
        result["revisions"] = {"tracked_changes": n_rev, "has_tracked_changes": True}
        result["issues"].append(
            f"检测到 {n_rev} 处修订标记（tracked changes）——排版前建议先接受/拒绝全部修订，"
            "否则格式修改会与修订叠加"
        )

    # 批注引用
    n_com = len(re.findall(r'<w:commentReference\b', document_xml))
    if n_com:
        result["comments"] = {"comment_refs": n_com, "has_comments": True}

    # 书签配对（属性顺序无关）
    starts = []
    for m in re.finditer(r'<w:bookmarkStart\b([^>]*)/>', document_xml):
        attrs = m.group(1)
        bid = re.search(r'w:id="(\d+)"', attrs)
        bname = re.search(r'w:name="([^"]*)"', attrs)
        if bid:
            starts.append((bname.group(1) if bname else f"id{bid.group(1)}", bid.group(1)))
    ends = set(m for m in re.findall(r'<w:bookmarkEnd\b[^>]*w:id="(\d+)"', document_xml))
    broken = [name for name, bid in starts if bid not in ends]
    result["bookmarks"]["pairs"] = len(starts)
    result["bookmarks"]["broken"] = broken
    if broken:
        result["issues"].append(f"书签未闭合: {len(broken)} 个 bookmarkStart 无对应 bookmarkEnd")

    return result


def complex_object_gate_errors(docx_path: str | Path) -> list[str]:
    """交付门禁用：返回必须阻断的硬错误（引用缺失类）。"""
    info = analyze_complex_objects(docx_path)
    errors = []
    if info["footnotes"]["missing_refs"]:
        errors.append(f"footnote reference integrity failed: {len(info['footnotes']['missing_refs'])} missing")
    if info["endnotes"]["missing_refs"]:
        errors.append(f"endnote reference integrity failed: {len(info['endnotes']['missing_refs'])} missing")
    return errors


def complex_object_gate_warnings(docx_path: str | Path) -> list[str]:
    """交付门禁用：非阻断警告（修订/批注——影响人工复核判断）。"""
    info = analyze_complex_objects(docx_path)
    warnings = []
    if info["revisions"]["has_tracked_changes"]:
        warnings.append(f"document contains {info['revisions']['tracked_changes']} tracked changes; "
                        "建议排版前接受全部修订")
    if info["comments"]["has_comments"]:
        warnings.append(f"document contains {info['comments']['comment_refs']} comment references; "
                        "批注内容不影响排版，交付时请复核")
    if info["bookmarks"]["broken"]:
        warnings.append(f"bookmark integrity: {len(info['bookmarks']['broken'])} broken bookmarks")
    return warnings
