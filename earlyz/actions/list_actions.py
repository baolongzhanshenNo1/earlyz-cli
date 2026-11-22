"""早睡 (EarlyZ) 3.2 — Phase 2 多级编号系统 action（合同/标书核心）。

set_list_level_format    — 修改 numbering.xml 中 abstractNum 指定级（ilvl）的格式
                           （num_fmt / lvl_text / start），确定性、幂等、带 verifier
map_paragraph_to_list    — 按 NodeSelector 定位段落，设置 numPr（num_id + ilvl）
                           挂到正确的编号层级

复用 new_actions 的 zip 读写基础设施。
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional
from lxml import etree

from earlyz.actions.new_actions import (
    _apply_structure_xml,
    _locate_paragraphs,
    _normalize_child_order,
    _read_part,
    _read_styles_map,
    _structure_xml_from_lxml,
    _write_part,
)
from earlyz.ir.ooxml_format import structure_bundle_contains

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_LVL_ORDER = [
    "start", "numFmt", "lvlRestart", "pStyle", "isLgl", "suff", "lvlText",
    "lvlPicBulletId", "lvlJc", "pPr", "rPr",
]


def _find_level(root: etree._Element, abstract_num_id: int, ilvl: int) -> Optional[etree._Element]:
    """定位 abstractNum 中指定 ilvl 的 <w:lvl> 元素（缺失则创建）。"""
    abs_el = None
    for a in root.findall(f"{W}abstractNum"):
        if int(a.get(f"{W}abstractNumId", -1)) == abstract_num_id:
            abs_el = a
            break
    if abs_el is None:
        raise ValueError(f"abstractNum {abstract_num_id} 不存在")
    lvl = None
    for l in abs_el.findall(f"{W}lvl"):
        if int(l.get(f"{W}ilvl", -1)) == ilvl:
            lvl = l
            break
    if lvl is None:
        raise ValueError(f"abstractNum {abstract_num_id} 无 ilvl={ilvl} 级别")
    return lvl


# ────────────────────────────────────────────────────────────────
# set_list_level_format
# ────────────────────────────────────────────────────────────────

def set_list_level_format(target_path: Path, params: dict) -> dict:
    """修改编号级别格式。

    params: {abstract_num_id, ilvl, num_fmt?, lvl_text?, start?}
    """
    target_path = Path(target_path)
    aid = int(params["abstract_num_id"])
    ilvl = int(params["ilvl"])
    root = _read_part(target_path, "word/numbering.xml")
    lvl = _find_level(root, aid, ilvl)

    if params.get("num_fmt"):
        el = lvl.find(f"{W}numFmt")
        if el is None:
            el = etree.SubElement(lvl, f"{W}numFmt")
        el.set(f"{W}val", params["num_fmt"])
    if params.get("lvl_text"):
        el = lvl.find(f"{W}lvlText")
        if el is None:
            el = etree.SubElement(lvl, f"{W}lvlText")
        el.set(f"{W}val", params["lvl_text"])
    if params.get("start") is not None:
        el = lvl.find(f"{W}start")
        if el is None:
            el = etree.SubElement(lvl, f"{W}start")
        el.set(f"{W}val", str(int(params["start"])))
    if "level_extra_xml" in params:
        _apply_structure_xml(lvl, params.get("level_extra_xml"), "lvl")
    _normalize_child_order(lvl, _LVL_ORDER)

    _write_part(target_path, "word/numbering.xml", root)
    return {"changed": True, "abstract_num_id": aid, "ilvl": ilvl}


def verify_list_level_format(target_path: Path, params: dict) -> bool:
    """验证指定级格式与期望一致。"""
    try:
        target_path = Path(target_path)
        aid = int(params["abstract_num_id"])
        ilvl = int(params["ilvl"])
        root = _read_part(target_path, "word/numbering.xml")
        lvl = _find_level(root, aid, ilvl)
        if params.get("num_fmt"):
            el = lvl.find(f"{W}numFmt")
            if el is None or el.get(f"{W}val") != params["num_fmt"]:
                return False
        if params.get("lvl_text"):
            el = lvl.find(f"{W}lvlText")
            if el is None or el.get(f"{W}val") != params["lvl_text"]:
                return False
        if params.get("start") is not None:
            el = lvl.find(f"{W}start")
            if el is None or el.get(f"{W}val") != str(int(params["start"])):
                return False
        if "level_extra_xml" in params:
            if not structure_bundle_contains(
                _structure_xml_from_lxml(lvl, "lvl"), params.get("level_extra_xml")
            ):
                return False
        return True
    except Exception:
        return False


# ────────────────────────────────────────────────────────────────
# map_paragraph_to_list
# ────────────────────────────────────────────────────────────────

def map_paragraph_to_list(target_path: Path, params: dict) -> dict:
    """按 selector 定位段落并设置 numPr（挂编号层级）。

    params: {selector: {...}, num_id, ilvl}
    """
    target_path = Path(target_path)
    selector = params["selector"]
    num_id = int(params["num_id"])
    ilvl = int(params["ilvl"])
    root = _read_part(target_path, "word/document.xml")
    paras = _locate_paragraphs(root, selector, _read_styles_map(target_path))
    if not paras:
        raise ValueError(f"no paragraph matched selector: {selector}")

    for p in paras:
        ppr = p.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            p.insert(0, ppr)
        numpr = ppr.find(f"{W}numPr")
        if numpr is None:
            numpr = etree.SubElement(ppr, f"{W}numPr")
        ilvl_el = numpr.find(f"{W}ilvl")
        if ilvl_el is None:
            ilvl_el = etree.SubElement(numpr, f"{W}ilvl")
        ilvl_el.set(f"{W}val", str(ilvl))
        numid_el = numpr.find(f"{W}numId")
        if numid_el is None:
            numid_el = etree.SubElement(numpr, f"{W}numId")
        numid_el.set(f"{W}val", str(num_id))

    _write_part(target_path, "word/document.xml", root)
    return {"changed": True, "paragraphs": len(paras)}


def verify_paragraph_list(target_path: Path, params: dict) -> bool:
    """验证目标段落 numPr 与期望一致。"""
    try:
        target_path = Path(target_path)
        selector = params["selector"]
        num_id = int(params["num_id"])
        ilvl = int(params["ilvl"])
        root = _read_part(target_path, "word/document.xml")
        paras = _locate_paragraphs(root, selector, _read_styles_map(target_path))
        if not paras:
            return False
        for p in paras:
            numpr = p.find(f"{W}pPr/{W}numPr")
            if numpr is None:
                return False
            ilvl_el = numpr.find(f"{W}ilvl")
            numid_el = numpr.find(f"{W}numId")
            if ilvl_el is None or numid_el is None:
                return False
            if int(ilvl_el.get(f"{W}val", -1)) != ilvl:
                return False
            if int(numid_el.get(f"{W}val", -1)) != num_id:
                return False
        return True
    except Exception:
        return False
