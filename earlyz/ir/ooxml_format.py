"""Lossless OOXML extras for paragraph/run formatting.

The semantic model keeps common properties as typed fields.  This module keeps
all remaining portable pPr/rPr formatting as canonical XML so a template
property cannot silently disappear merely because FormatProperties has no
field for it.  Identity/relationship/structure nodes are deliberately routed
to their own subsystems instead of being copied as paragraph formatting.
"""
from __future__ import annotations

import copy
import xml.etree.ElementTree as ET
from typing import Optional

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{" + W_NS + "}"

# These children are identity, numbering, revision or section structure.  They
# are inventoried elsewhere and must never be transplanted as role formatting.
PPR_EXCLUDED = frozenset({
    "pStyle", "numPr", "sectPr", "framePr", "pPrChange", "cnfStyle",
})
RPR_EXCLUDED = frozenset({"rStyle", "rPrChange", "ins", "del", "moveFrom", "moveTo",
                          # 2026-09-27 真机事故：模板 extras 回放（`_apply_extra_xml` 先清未建模
                          # 属性再精确回放）把**作者语义**当 extras 清掉 → 作者角标 1)/2)、
                          # 参考文献引用上标全部消失（源 9 上标/65 下标 → 产物 1/5）。
                          # 边界：extras 通道管版式（w/kern/szCs…），不管作者语义——
                          # vertAlign 上标下标 / oMath 数学 run / vanish·specVanish 隐藏文本 /
                          # rtl 书写方向：既不清除，也不由角色级 extras 注入。
                          "vertAlign", "oMath", "vanish", "specVanish", "rtl"})

# Portable structural properties.  Relationship IDs, section topology,
# page-number restarts, columns and table geometry remain owned by dedicated
# structural serializers and are deliberately absent from these allowlists.
SECTPR_PORTABLE = frozenset({
    "pgSz", "paperSrc", "pgBorders", "lnNumType", "formProt", "vAlign",
    "noEndnote", "textDirection", "bidi", "rtlGutter", "docGrid",
})
TBLPR_PORTABLE = frozenset({
    "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
    "tblCellSpacing", "tblBorders", "shd", "tblLayout", "tblCellMar",
    "tblLook",
})
LVL_PORTABLE = frozenset({
    "lvlRestart", "isLgl", "suff", "lvlJc", "legacy", "pPr", "rPr",
})

# Attributes represented by typed FormatProperties.  Any other attribute on
# the same element is preserved in the extras bundle (for example rFonts/hAnsi,
# rFonts/cs, spacing/*Autospacing and ind/*Chars).
PPR_MODELED_ATTRS = {
    "jc": frozenset({"val"}),
    "ind": frozenset({"firstLine", "firstLineChars", "left", "right", "hanging"}),
    "spacing": frozenset({"before", "after", "beforeLines", "afterLines", "line", "lineRule"}),
    "keepNext": frozenset({"val"}),
    "keepLines": frozenset({"val"}),
    "pageBreakBefore": frozenset({"val"}),
    "widowControl": frozenset({"val"}),
    "snapToGrid": frozenset({"val"}),
    "outlineLvl": frozenset({"val"}),
}
PPR_FULLY_MODELED = frozenset(PPR_MODELED_ATTRS)

RPR_MODELED_ATTRS = {
    "rFonts": frozenset({"ascii", "eastAsia"}),
    "b": frozenset({"val"}),
    "i": frozenset({"val"}),
    "smallCaps": frozenset({"val"}),
    "u": frozenset({"val"}),
    "color": frozenset({"val"}),
    "sz": frozenset({"val"}),
}
RPR_FULLY_MODELED = frozenset(RPR_MODELED_ATTRS)

PPR_SEQUENCE = (
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
    "sectPr", "pPrChange",
)
RPR_SEQUENCE = (
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
    "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof",
    "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
    "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
    "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
    "specVanish", "oMath", "textFill", "textOutline",
)
SECTPR_SEQUENCE = (
    "headerReference", "footerReference", "footnotePr", "endnotePr", "type",
    "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType",
    "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection",
    "bidi", "rtlGutter", "docGrid", "printerSettings",
)
TBLPR_SEQUENCE = (
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
    "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
    "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
)
LVL_SEQUENCE = (
    "start", "numFmt", "lvlRestart", "pStyle", "isLgl", "suff", "lvlText",
    "lvlPicBulletId", "lvlJc", "pPr", "rPr",
)


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _stdlib_clone(el) -> ET.Element:
    """Clone either stdlib or lxml input into a stdlib Element.

    Analyzer normally supplies ``xml.etree`` nodes, while the explain-format
    cascade intentionally uses lxml.  Deep-copying preserves the source node
    implementation, which cannot be appended across those two libraries.
    Serializing at this boundary also makes canonicalization implementation-
    independent.
    """
    try:
        payload = ET.tostring(el, encoding="utf-8")
    except TypeError:
        from lxml import etree as LET
        payload = LET.tostring(el, encoding="utf-8")
    return ET.fromstring(payload)


def _canonicalize(el: ET.Element) -> None:
    if el.attrib:
        items = sorted(el.attrib.items(), key=lambda kv: kv[0])
        el.attrib.clear()
        el.attrib.update(items)
    children = list(el)
    parent = local_name(el.tag)
    seq = (
        PPR_SEQUENCE if parent == "pPr"
        else RPR_SEQUENCE if parent == "rPr"
        else SECTPR_SEQUENCE if parent == "sectPr"
        else TBLPR_SEQUENCE if parent == "tblPr"
        else LVL_SEQUENCE if parent == "lvl"
        else ()
    )
    if seq and children:
        rank = {name: i for i, name in enumerate(seq)}
        indexed = list(enumerate(children))
        indexed.sort(key=lambda pair: (rank.get(local_name(pair[1].tag), len(rank)), pair[0]))
        for child in children:
            el.remove(child)
        for _, child in indexed:
            el.append(child)
        children = list(el)
    for child in children:
        _canonicalize(child)


def _clone_extra_child(child: ET.Element, modeled_attrs: dict[str, frozenset[str]]) -> Optional[ET.Element]:
    """Clone only the unmodelled portion of a partially modelled child."""
    name = local_name(child.tag)
    known = modeled_attrs.get(name)
    if known is None:
        return _stdlib_clone(child)
    out = ET.Element(child.tag)
    for key, value in child.attrib.items():
        if local_name(key) not in known:
            out.set(key, value)
    # Modelled elements in these maps are leaf elements in WordprocessingML.
    # Retain unexpected descendants fail-safe so extensions are never dropped.
    for sub in list(child):
        out.append(_stdlib_clone(sub))
    if not out.attrib and not list(out):
        return None
    return out


def extract_extra_bundle(container: Optional[ET.Element], kind: str) -> Optional[str]:
    """Return canonical XML containing every portable unmodelled property."""
    if container is None:
        return None
    if kind == "pPr":
        excluded, modeled = PPR_EXCLUDED, PPR_MODELED_ATTRS
    elif kind == "rPr":
        excluded, modeled = RPR_EXCLUDED, RPR_MODELED_ATTRS
    else:
        raise ValueError(f"unsupported format bundle kind: {kind}")
    out = ET.Element(W + kind)
    for child in list(container):
        if local_name(child.tag) in excluded:
            continue
        extra = _clone_extra_child(child, modeled)
        if extra is not None:
            out.append(extra)
    if not list(out):
        return None
    _canonicalize(out)
    return ET.tostring(out, encoding="unicode", short_empty_elements=True)


def extract_structure_bundle(container: Optional[ET.Element], kind: str) -> Optional[str]:
    """Return canonical portable structural formatting for one container.

    The result contains formatting only.  It never contains relationship IDs,
    target-content geometry, numbering identity or section topology.
    """
    if container is None:
        return None
    if kind == "sectPr":
        allowed = SECTPR_PORTABLE
    elif kind == "tblPr":
        allowed = TBLPR_PORTABLE
    elif kind == "lvl":
        allowed = LVL_PORTABLE
    else:
        raise ValueError(f"unsupported structure bundle kind: {kind}")
    out = ET.Element(W + kind)
    for child in list(container):
        if local_name(child.tag) in allowed:
            out.append(_stdlib_clone(child))
    if not list(out):
        return None
    _canonicalize(out)
    return ET.tostring(out, encoding="unicode", short_empty_elements=True)


def structure_bundle_contains(actual_xml: Optional[str], expected_xml: Optional[str]) -> bool:
    """Return whether actual contains every child declared by expected.

    A common structural bundle is an intersection, not a complete sectPr or
    tblPr.  Role-specific children absent from expected remain valid.
    """
    if not expected_xml:
        return True
    if not actual_xml:
        return False

    def child_map(xml_text: str) -> dict[str, str]:
        root = ET.fromstring(xml_text)
        result = {}
        for child in list(root):
            clone = copy.deepcopy(child)
            _canonicalize(clone)
            result[local_name(clone.tag)] = ET.tostring(
                clone, encoding="unicode", short_empty_elements=True
            )
        return result

    actual = child_map(actual_xml)
    expected = child_map(expected_xml)
    return all(actual.get(name) == value for name, value in expected.items())


def _merge_element(base: ET.Element, override: ET.Element) -> ET.Element:
    """Child/property cascade: override attrs and recursively merge same tags."""
    out = copy.deepcopy(base)
    for key, value in override.attrib.items():
        out.set(key, value)
    for och in list(override):
        oname = local_name(och.tag)
        match = next((x for x in list(out) if local_name(x.tag) == oname), None)
        if match is None:
            out.append(copy.deepcopy(och))
        else:
            idx = list(out).index(match)
            out.remove(match)
            out.insert(idx, _merge_element(match, och))
    return out


def merge_extra_bundles(base_xml: Optional[str], override_xml: Optional[str], kind: str) -> Optional[str]:
    """Merge style-chain/direct extras without losing parent-only properties."""
    if not base_xml:
        return override_xml
    if not override_xml:
        return base_xml
    base = ET.fromstring(base_xml)
    over = ET.fromstring(override_xml)
    out = ET.Element(W + kind)
    by_name = {local_name(ch.tag): copy.deepcopy(ch) for ch in list(base)}
    order = [local_name(ch.tag) for ch in list(base)]
    for och in list(over):
        name = local_name(och.tag)
        if name in by_name:
            by_name[name] = _merge_element(by_name[name], och)
        else:
            by_name[name] = copy.deepcopy(och)
            order.append(name)
    for name in order:
        out.append(by_name[name])
    _canonicalize(out)
    return ET.tostring(out, encoding="unicode", short_empty_elements=True)
