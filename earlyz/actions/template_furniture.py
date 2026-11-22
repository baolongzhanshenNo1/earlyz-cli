"""模板版式构件继承：按节复制模板的页眉部件（含 logo 图片）。

背景（2026-10-30 实测）：产品版式从源文档起笔，源文档没有页眉；模板的页眉
里常放着单位 logo。产物缺失这些部件，视觉上就是「页眉整块消失」。

**关键教训**：不能把模板第 2 节的页眉一股脑套到产物所有节。功能设计说明书
模板第 1 页页眉本来空白、第 2 节才放 logo；早期做法让首页显示出一个灰框
占位图，比不继承更糟。所以必须**按节对应**：产物第 i 节沿用模板第 i 节的
headerReference 集合（类型 + 部件），模板没有的节就不加。

其余约束：
- 仅当产物**一个页眉部件都没有**时才继承（源自带页眉时不动）；
- 部件字节原样拷贝，其 rels（图片）一并搬进 word/media，**保留原 rId**
  （页眉 XML 里 r:embed="rId1" 是硬编码的，换 id 会让 Word 画空框）；
- CT_SectPr 要求 headerReference 排在段属性最前；
- 写入标记 part `word/earlyz_furniture.xml` 记录已继承，保证幂等。
"""

from __future__ import annotations

import json
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
MARKER = "word/earlyz_furniture.xml"
HEADER_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
HEADER_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"


def _has_content(blob: bytes) -> bool:
    """页眉部件是否有可看见的内容：文字 / 图片 / 边框线。"""
    try:
        root = etree.fromstring(blob)
    except Exception:  # noqa: BLE001
        return False
    if "".join(t.text or "" for t in root.iter(W + "t")).strip():
        return True
    if list(root.iter(W + "drawing")) or list(root.iter(W + "pict")):
        return True
    return bool(list(root.iter(W + "pBdr")))


def _read(z: zipfile.ZipFile, name: str) -> bytes:
    try:
        return z.read(name)
    except KeyError:
        return b""


# ── 渲染门控 ────────────────────────────────────────────────────────────
# 「模板页眉部件带图」≠「模板上看得见页眉」：功能设计说明书模板的 header1/2
# 里有 inline 图片，但 Word 渲染出来页眉整块是空白（实测）。所以继承之前
# 必须**看渲染结果**：把模板转 PDF，量页眉带里有没有真内容。
# 判据要把「横线」和「内容」分开：横线是一整行暗像素（>70% 页宽），
# 校徽/文字则是局部像素。只数非横线行的暗像素。

def _page_geometry(template_path: str | Path) -> dict:
    """模板第一节的页面几何：页高、上边距、页眉距（twips）。"""
    geo = {"page_h": 16838, "top": 1440, "header": 851}
    try:
        with zipfile.ZipFile(template_path) as tz:
            doc = etree.fromstring(_read(tz, "word/document.xml"))
        sect = next(iter(doc.iter(W + "sectPr")), None)
        if sect is None:
            return geo
        pgsz = sect.find(W + "pgSz")
        pgmar = sect.find(W + "pgMar")
        if pgsz is not None and pgsz.get(W + "h"):
            geo["page_h"] = int(pgsz.get(W + "h"))
        if pgmar is not None:
            if pgmar.get(W + "top"):
                geo["top"] = int(pgmar.get(W + "top"))
            if pgmar.get(W + "header"):
                geo["header"] = int(pgmar.get(W + "header"))
    except Exception:  # noqa: BLE001
        pass
    return geo


def _header_band_has_content(pdf_path: str | Path, page_index: int = 0,
                             band_frac: float = 0.10, scale: float = 1.0) -> tuple[bool, int]:
    """渲染该页，数「页面顶部到上边距之间」的暗像素（页眉只能出现在这一段）。

    正文从上边距开始，所以这一段里出现的任何非白像素都是页眉内容（校徽/
    校名/页眉横线）。band_frac 由调用方按 pgMar/top ÷ pgSz/h 算好传入。
    """
    try:
        from earlyz.validation.pdf_backend import render_rgb

        w, h, rgb = render_rgb(pdf_path, page_index, scale)
    except Exception:  # noqa: BLE001
        return False, 0
    band = max(1, int(h * band_frac))
    total = 0
    for y in range(band):
        base = y * w * 3
        for x in range(w):
            i = base + x * 3
            if rgb[i] < 200 or rgb[i + 1] < 200 or rgb[i + 2] < 200:
                total += 1
    return total >= 120, total


_GATE_CACHE = Path(tempfile.gettempdir()) / "earlyz_header_gate.json"


def _cache_key(path: Path) -> str:
    try:
        st = path.stat()
        return f"{path.resolve()}|{int(st.st_mtime)}|{st.st_size}"
    except Exception:  # noqa: BLE001
        return str(path)


def _cache_load() -> dict:
    try:
        return json.loads(_GATE_CACHE.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return {}


def _cache_store(key: str, visible: bool, info: dict) -> None:
    try:
        data = _cache_load()
        data[key] = {"visible": visible, "info": info}
        _GATE_CACHE.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass


def template_header_visible(template_path: str | Path, render_fn=None) -> tuple[bool, dict]:
    """门控：模板渲染出来的页眉是否真有内容（默认看第 1 页，再看第 2 页）。

    render_fn：与 compile_document 同款 `render(docx_path) -> {"pdf_path": ...}`。
    没有 render_fn 时尝试用 Word COM（Windows）。
    """
    tpl = Path(template_path)
    key = _cache_key(tpl)
    cached = _cache_load().get(key)
    if cached:
        return bool(cached.get("visible")), {**cached.get("info", {}), "cached": True}
    info: dict = {"checked": False, "dark": 0, "pdf": None}
    try:
        if render_fn is None:
            from earlyz.validation.renderers import build_word_render_callback

            render_fn = build_word_render_callback()
        res = render_fn(str(template_path)) or {}
        pdf = res.get("pdf_path")
        if not pdf or not Path(pdf).exists():
            info["reason"] = "no_pdf"
            return False, info
        info["checked"] = True
        info["pdf"] = str(pdf)
        geo = _page_geometry(template_path)
        # 页眉可见区 = 页顶到上边距（twips → 页高占比），留 2% 余量避免压到首行正文
        frac = max(0.02, min(0.35, (geo["top"] / max(1, geo["page_h"])) - 0.01))
        info["band_frac"] = round(frac, 3)
        for idx in (0, 1):
            ok, dark = _header_band_has_content(pdf, idx, band_frac=frac)
            info[f"dark_p{idx + 1}"] = dark
            if ok:
                info["dark"] = dark
                _cache_store(key, True, info)
                return True, info
        info["reason"] = "header_band_blank"
        _cache_store(key, False, info)
        return False, info
    except Exception as e:  # noqa: BLE001
        info["reason"] = f"render_failed: {e}"
        return False, info


def _load_part(tz: zipfile.ZipFile, part: str) -> tuple:
    """读一个页眉部件：字节 + rels 条目 + 可搬的图片（rid, 原路径, 字节）。"""
    blob = _read(tz, part)
    rels = _read(tz, f"word/_rels/{Path(part).name}.rels")
    entries: list = []
    media: list = []
    if rels:
        for rel in etree.fromstring(rels):
            rid0 = rel.get("Id")
            if not rid0:
                continue
            typ0 = rel.get("Type") or ""
            tgt = rel.get("Target") or ""
            mode = rel.get("TargetMode")
            if mode == "External":
                entries.append((rid0, typ0, tgt, mode))
                continue
            norm = _normalize(part, tgt)
            mblob = _read(tz, norm)
            if mblob:
                media.append((rid0, norm, mblob))
            entries.append((rid0, typ0, tgt, None))
    return blob, entries, media


def inherit_template_headers(product_path: str | Path, template_path: str | Path,
                             render_fn=None, gate: bool = True) -> dict:
    """把模板页眉（含图片）按节继承到产物。

    gate=True（默认）先做渲染门控：模板渲染出来页眉是空白就不继承。
    render_fn：`render(docx_path) -> {"pdf_path": ...}`；缺省时用 Word COM。
    """
    product_path = Path(product_path)
    template_path = Path(template_path)
    stats: dict = {"inherited": 0, "parts": [], "skipped": None}
    if not product_path.exists() or not template_path.exists():
        stats["skipped"] = "missing_file"
        return stats
    if gate:
        visible, info = template_header_visible(template_path, render_fn)
        stats["gate"] = info
        if not visible:
            stats["skipped"] = "template_header_not_visible"
            return stats

    with zipfile.ZipFile(template_path) as tz:
        tpl_doc = etree.fromstring(_read(tz, "word/document.xml"))
        if tpl_doc is None:
            stats["skipped"] = "template_no_document"
            return stats
        tpl_sects = list(tpl_doc.iter(W + "sectPr"))
        if not tpl_sects:
            stats["skipped"] = "template_no_sectpr"
            return stats
        per_section: list[list] = []      # 每节 [(type, part)]
        cache: dict[str, tuple] = {}
        for sect in tpl_sects:
            refs: list = []
            for ref in sect.findall(W + "headerReference"):
                typ = ref.get(W + "type", "default")
                tgt = _rels_target(tz, "word/_rels/document.xml.rels", ref.get(R + "id"))
                if not tgt:
                    continue
                if tgt not in cache:
                    cache[tgt] = _load_part(tz, tgt)
                if _has_content(cache[tgt][0]):
                    refs.append((typ, tgt))
            per_section.append(refs)
        if not any(per_section):
            stats["skipped"] = "template_headers_empty"
            return stats

    with zipfile.ZipFile(product_path) as pz:
        names = set(pz.namelist())
        if any(n.startswith("word/header") for n in names):
            stats["skipped"] = "product_has_headers"
            return stats
        entries = {n: pz.read(n) for n in names}

    doc = etree.fromstring(entries["word/document.xml"])
    rels_root = etree.fromstring(entries["word/_rels/document.xml.rels"])
    ct_root = etree.fromstring(entries["[Content_Types].xml"])
    used_ids = {x.get("Id") for x in rels_root}
    nid = 1

    def new_rid() -> str:
        nonlocal nid
        while f"rIdFurn{nid}" in used_ids:
            nid += 1
        rid = f"rIdFurn{nid}"
        used_ids.add(rid)
        return rid

    added: dict[str, bytes] = {}
    part_rid: dict[str, str] = {}
    for part, (blob, rels_entries, media) in cache.items():
        idx = len(part_rid) + 1
        part_name = f"word/headerFurn{idx}.xml"
        added[part_name] = blob
        ct_root.append(ct_root.makeelement(CT + "Override", {
            "PartName": "/" + part_name, "ContentType": HEADER_CT}))
        # 图片搬运 + 该页眉自己的 rels。
        # 关键：rels 里必须保留模板原来的 rId——页眉 XML 里的 r:embed="rId1"
        # 是硬编码的；换成自定义 id 会让 Word 只画一个空框（实测）。
        media_target: dict[str, str] = {}
        for k, (rid0, norm, mblob) in enumerate(media, start=1):
            ext = Path(norm).suffix or ".png"
            mname = f"word/media/furniture{idx}_{k}{ext}"
            added[mname] = mblob
            _ensure_default_ext(ct_root, ext)
            media_target[rid0] = f"media/{Path(mname).name}"
        rels_attrs = []
        for rid0, typ0, tgt, mode in rels_entries:
            attrs = {"Id": rid0, "Type": typ0,
                     "Target": media_target.get(rid0, tgt)}
            if mode:
                attrs["TargetMode"] = mode
            rels_attrs.append(attrs)
        if rels_attrs:
            rr = etree.Element(PR + "Relationships", nsmap={None: PR[1:-1]})
            for attrs in rels_attrs:
                rr.append(rr.makeelement(PR + "Relationship", attrs))
            added[f"word/_rels/{Path(part_name).name}.rels"] = etree.tostring(
                rr, xml_declaration=True, encoding="UTF-8", standalone=True)
        rid = new_rid()
        rels_root.append(rels_root.makeelement(PR + "Relationship", {
            "Id": rid, "Type": HEADER_REL, "Target": Path(part_name).name}))
        part_rid[part] = rid
        stats["parts"].append(part_name)

    # 按节绑定：产物第 i 节沿用模板第 i 节的页眉集合（模板没有就不加）。
    # 产物节多于模板时，沿用模板最后一节的集合。
    for i, sect in enumerate(doc.iter(W + "sectPr")):
        refs = per_section[min(i, len(per_section) - 1)]
        for typ, part in refs:
            rid = part_rid.get(part)
            if not rid:
                continue
            if sect.find(f"{W}headerReference[@{W}type='{typ}']") is not None:
                continue
            sect.insert(0, sect.makeelement(W + "headerReference",
                                            {W + "type": typ, R + "id": rid}))
            stats["inherited"] += 1

    entries["word/document.xml"] = etree.tostring(
        doc, xml_declaration=True, encoding="UTF-8", standalone=True)
    entries["word/_rels/document.xml.rels"] = etree.tostring(
        rels_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    entries["[Content_Types].xml"] = etree.tostring(
        ct_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    entries[MARKER] = f"headers={stats['parts']}".encode()

    tmp = product_path.with_suffix(".furn.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as oz:
        for name, blob in entries.items():
            oz.writestr(name, blob)
        for name, blob in added.items():
            oz.writestr(name, blob)
    shutil.move(str(tmp), str(product_path))
    return stats


def _rels_target(z: zipfile.ZipFile, rels_name: str, rid: str | None) -> str | None:
    if not rid:
        return None
    blob = _read(z, rels_name)
    if not blob:
        return None
    for rel in etree.fromstring(blob):
        if rel.get("Id") == rid:
            tgt = rel.get("Target") or ""
            return _normalize("word/document.xml", tgt)
    return None


def _normalize(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base_dir = Path(base_part).parent
    return str((base_dir / target).as_posix()).replace("./", "")


def _ensure_default_ext(ct_root, ext: str) -> None:
    ext = ext.lstrip(".").lower()
    if not ext:
        return
    for d in ct_root.findall(CT + "Default"):
        if (d.get("Extension") or "").lower() == ext:
            return
    mime = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
            "gif": "image/gif", "emf": "image/x-emf", "wmf": "image/x-wmf",
            "xml": "application/xml"}.get(ext, "application/octet-stream")
    ct_root.append(ct_root.makeelement(CT + "Default", {"Extension": ext, "ContentType": mime}))
