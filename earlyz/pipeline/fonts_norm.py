"""文档字体名归一（pre-analyze，作用于工作副本）。

模板/源文档常见两种脏字体名：
1. **替代字体**：`Noto Serif SC`、`Liberation Serif`（LibreOffice 环境产物，
   Windows Word 没装 → 打开时自行替换，字形和行宽都变）；
2. **带分号的字体串**：`宋体;SimSun`、`黑体;SimHei`（作者把「中文名;英文名」
   一起写进了 `w:eastAsia`，Word 当成一个不存在的字体名 → 回退默认字体）。

两者都不改文字，只改 `w:rFonts` 的属性值，属纯格式归一。
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

from lxml import etree

from earlyz.tools.fonts import normalize_font_field

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_FONT_PARTS = ("word/document.xml", "word/styles.xml")


def _iter_font_parts(z: zipfile.ZipFile) -> list[str]:
    names = [n for n in z.namelist()
             if n in _FONT_PARTS or re.match(r"word/(header|footer)\d*\.xml$", n)]
    return names


def normalize_fonts_in_copy(src: str | Path, workdir: str | Path) -> tuple[Path, dict]:
    """把字体名归一后的副本写到 workdir，返回 (副本路径, 统计)。"""
    src = Path(src)
    work = Path(workdir)
    work.mkdir(parents=True, exist_ok=True)
    dst = work / (src.stem + "-fonts.docx")
    if dst.exists():
        dst.unlink()

    fixed = 0
    changed_files: list[str] = []
    with zipfile.ZipFile(src) as zin:
        parts = _iter_font_parts(zin)
        new_xml: dict[str, bytes] = {}
        for name in parts:
            try:
                root = etree.fromstring(zin.read(name))
            except etree.XMLSyntaxError:
                continue
            touched = 0
            for rf in root.iter(f"{W}rFonts"):
                for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
                    key = f"{W}{attr}"
                    val = rf.get(key)
                    if not val:
                        continue
                    new = normalize_font_field(val)
                    if new != val:
                        rf.set(key, new)
                        touched += 1
            if touched:
                new_xml[name] = etree.tostring(root, xml_declaration=True,
                                               encoding="UTF-8", standalone=True)
                fixed += touched
                changed_files.append(name)

        if not new_xml:
            return src, {"changed": False, "fonts_fixed": 0, "files": []}

        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = new_xml.get(item.filename)
                if data is not None:
                    zout.writestr(item, data)
                else:
                    zout.writestr(item, zin.read(item.filename))

    return dst, {"changed": True, "fonts_fixed": fixed, "files": changed_files}


def normalize_fonts_in_place(path: str | Path) -> dict:
    """对已存在的 docx 原地归一字体名（用于动作落盘后的收尾）。

    动作会把模板的 ppr_extra_xml / rpr_extra_xml 原样写进产物，里面可能带着
    「宋体;SimSun」这类脏字体名——规则字段归一管不到，必须再扫一遍文档。
    """
    path = Path(path)
    fixed = 0
    with zipfile.ZipFile(path) as zin:
        parts = _iter_font_parts(zin)
        new_xml: dict[str, bytes] = {}
        for name in parts:
            try:
                root = etree.fromstring(zin.read(name))
            except etree.XMLSyntaxError:
                continue
            touched = 0
            for rf in root.iter(f"{W}rFonts"):
                for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
                    key = f"{W}{attr}"
                    val = rf.get(key)
                    if not val:
                        continue
                    new = normalize_font_field(val)
                    if new != val:
                        rf.set(key, new)
                        touched += 1
            if touched:
                new_xml[name] = etree.tostring(root, xml_declaration=True,
                                               encoding="UTF-8", standalone=True)
                fixed += touched
        if not new_xml:
            return {"changed": False, "fonts_fixed": 0}
        tmp = path.with_suffix(".fonttmp.docx")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = new_xml.get(item.filename)
                zout.writestr(item, data if data is not None else zin.read(item.filename))
    shutil.move(str(tmp), str(path))
    return {"changed": True, "fonts_fixed": fixed}
