
# ── 2026-08-25 建议层（§22.5 P2）新增 action ──────────────────────

from pathlib import Path

from lxml import etree

def bold_toc_page_numbers(path, params: dict) -> dict:
    """目录条目页码加粗（R1 纯格式，幂等，建议层「目录中章标题页码应加粗」）。

    对 TOC 条目段落（含 w:hyperlink 或 pStyle=TOC{n}）中 PAGEREF 域的
    result 数字 run 加 <w:b/>（无 rPr 补建并置首位——rPr 须在 r 首位）；
    已有 b 跳过。不碰条目文字与域结构（只加粗）。
    """
    from earlyz.actions.new_actions import _read_part, _write_part, W as _W

    path = Path(path)
    root = _read_part(path, "word/document.xml")
    done = 0
    for p in root.iter(f"{_W}p"):
        ppr = p.find(f"{_W}pPr")
        pstyle = ppr.find(f"{_W}pStyle") if ppr is not None else None
        # 只加粗**章标题**（TOC1 一级条目）页码（2026-08-26 用户纠错：节标题
        # TOC2：1.1/1.2…不能加粗——TOC 域 \o "1-2" 含两级条目）
        tval = (pstyle.get(f"{_W}val") or "") if pstyle is not None else ""
        if not tval.startswith("TOC1"):
            continue
        # PAGEREF 域 result 数字 run：域内（begin..end 之间）且文本为纯数字
        runs = list(p.iter(f"{_W}r"))
        in_field = False
        for r in runs:
            fc = r.find(f"{_W}fldChar")
            if fc is not None:
                typ = fc.get(f"{_W}fldCharType")
                if typ == "begin":
                    in_field = True
                    continue
                if typ == "separate":
                    continue
                if typ == "end":
                    in_field = False
                    continue
            if not in_field:
                continue
            txt = "".join(t.text or "" for t in r.iter(f"{_W}t")).strip()
            # 2026-08-26 用户拍板：目录一级条目页码全部加粗——含罗马数字
            # 前导页（致  谢 I / 摘  要 II / ABSTRACT III——用户明确「这三个是章标题」）
            if not txt:
                continue
            rpr = r.find(f"{_W}rPr")
            if rpr is None:
                rpr = etree.Element(f"{_W}rPr")
                r.insert(0, rpr)
            if rpr.find(f"{_W}b") is None:
                b = etree.SubElement(rpr, f"{_W}b")
                b.set(f"{_W}val", "1")
                done += 1
    if done:
        _write_part(path, "word/document.xml", root)
    return {"page_numbers_bolded": done}


def verify_toc_page_numbers(path, params: dict) -> bool:
    """verifier：至少一个 PAGEREF 数字 run 带加粗（幂等验证，不校验全量）。"""
    from earlyz.actions.new_actions import _read_part, W as _W

    try:
        root = _read_part(Path(path), "word/document.xml")
        for p in root.iter(f"{_W}p"):
            ppr = p.find(f"{_W}pPr")
            pstyle = ppr.find(f"{_W}pStyle") if ppr is not None else None
            tval = (pstyle.get(f"{_W}val") or "") if pstyle is not None else ""
            if not tval.startswith("TOC1"):
                continue
            in_field = False
            for r in p.iter(f"{_W}r"):
                fc = r.find(f"{_W}fldChar")
                if fc is not None:
                    typ = fc.get(f"{_W}fldCharType")
                    if typ == "begin":
                        in_field = True
                        continue
                    if typ == "end":
                        in_field = False
                        continue
                if in_field and r.find(f"{_W}rPr") is not None \
                        and r.find(f"{_W}rPr").find(f"{_W}b") is not None:
                    txt = "".join(t.text or "" for t in r.iter(f"{_W}t")).strip()
                    # 罗马页码（I/II/III…）同样加粗——用户明确致谢/摘要/ABSTRACT 是章标题
                    if txt:
                        return True
        return False
    except Exception:  # noqa: BLE001
        return False
