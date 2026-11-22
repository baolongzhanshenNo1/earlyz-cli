"""包级清理：删掉没有任何元素引用的 `word/embeddings/*` OLE 部件及其 rels 条目。

**为什么需要**（2026-11-07 实测）：42/199 份产物带着孤儿 OLE 部件——正文里已经
没有 `w:object`（公式被改成文字/图片），但 `word/embeddings/oleObjectN.bin` 与
`word/_rels/document.xml.rels` 里的 OLEObject 关系还留着。大连民族 30 个、矿大 13 个，
包体白白虚胖数 MB；更麻烦的是这些关系指向的部件一旦缺失，Word 会弹「需要修复」。

**口径**：
- 只处理 `word/embeddings/` 下的部件，不碰 `word/media/`（媒体可能被 VML/图片引用）。
- 「被引用」的判定是保守的：扫描包内**所有** XML 部件里关系命名空间下的属性值
  （`r:id` / `r:embed` / `r:link` / `r:pict` …），出现在任何一个里就算被引用。
- 先删引用不上的关系，再删对应部件；`TargetMode="External"` 一律跳过。
- 幂等：没有孤儿时 `removed_parts == 0`，zip 不重写。
"""
from __future__ import annotations

import posixpath
import re
import zipfile
from pathlib import Path

from lxml import etree

R = "{http://schemas.openxmlformats.org/package/2006/relationships}"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_REF_ATTR = re.compile(r"\{http://schemas\.openxmlformats\.org/officeDocument/2006/"
                       r"relationships\}[a-zA-Z]+$")
_EMB_PREFIX = "word/embeddings/"


def _referenced_ids(z: zipfile.ZipFile) -> set[str]:
    """包内所有 XML 部件引用到的关系 Id（保守：任何 r:* 属性值都算）。"""
    out: set[str] = set()
    for name in z.namelist():
        if not name.lower().endswith((".xml", ".rels")) or name.startswith("_rels/"):
            continue
        if "/_rels/" in name:      # 关系文件本身不是引用方
            continue
        try:
            root = etree.fromstring(z.read(name))
        except Exception:  # noqa: BLE001 — 非 XML/损坏部件跳过
            continue
        for el in root.iter():
            for k, v in el.attrib.items():
                if _REF_ATTR.match(k) and v:
                    out.add(v)
    return out


def _resolve(part_name: str, target: str) -> str:
    base = posixpath.dirname(part_name)          # 例：word/_rels/document.xml.rels → word/_rels
    base = posixpath.dirname(base)               # → word
    return posixpath.normpath(posixpath.join(base, target))


def drop_orphan_embeddings(path) -> dict:
    """删除未被引用的 OLE 部件与关系。非阻断：异常只记在返回值里。"""
    info: dict = {"removed_parts": 0, "removed_rels": 0, "kept": 0,
                  "bytes_saved": 0, "changed": False, "errors": []}
    path = Path(path)
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            refs = _referenced_ids(z)
            emb = [n for n in names if n.startswith(_EMB_PREFIX) and not n.endswith("/")]
            if not emb:
                return info
            # 每个 embeddings 部件被哪些关系指向
            rels_files = [n for n in names if n.endswith(".rels")]
            drop_rel: dict[str, set[str]] = {}
            drop_part: set[str] = set()
            for rf in rels_files:
                try:
                    root = etree.fromstring(z.read(rf))
                except Exception:  # noqa: BLE001
                    continue
                ids: set[str] = set()
                for rel in root.findall(f"{R}Relationship"):
                    if (rel.get("TargetMode") or "").lower() == "external":
                        continue
                    target = rel.get("Target") or ""
                    part = _resolve(rf, target)
                    if not part.startswith(_EMB_PREFIX):
                        continue
                    rid = rel.get("Id") or ""
                    if rid and rid in refs:
                        info["kept"] += 1
                        continue
                    ids.add(rid)
                    drop_part.add(part)
                if ids:
                    drop_rel[rf] = ids
            if not drop_part:
                return info
            payload: dict[str, bytes | None] = {}
            saved = 0
            for n in names:
                if n in drop_part:
                    payload[n] = None
                    saved += z.getinfo(n).file_size
                    continue
                if n in drop_rel:
                    root = etree.fromstring(z.read(n))
                    for rel in list(root.findall(f"{R}Relationship")):
                        if (rel.get("Id") or "") in drop_rel[n]:
                            root.remove(rel)
                    payload[n] = etree.tostring(root, xml_declaration=True,
                                                encoding="UTF-8", standalone=True)
                    continue
                payload[n] = z.read(n)
        # 必须先关掉源 zip 再替换：Windows 下仍持有句柄时 replace 会 PermissionError
        tmp = path.with_suffix(".cleanup")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            for n, data in payload.items():
                if data is None:
                    continue
                zo.writestr(n, data)
        tmp.replace(path)
        info.update(removed_parts=len(drop_part),
                    removed_rels=sum(len(v) for v in drop_rel.values()),
                    bytes_saved=saved, changed=True)
    except Exception as e:  # noqa: BLE001 — 非阻断
        info["errors"].append(f"{type(e).__name__}: {e}")
    return info


def dangling_relationship_ids(path) -> list[str]:
    """自检：关系文件里登记的 Id，若被正文引用但部件不存在 = 会导致 Word「需要修复」。"""
    bad: list[str] = []
    path = Path(path)
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            refs = _referenced_ids(z)
            for rf in [n for n in names if n.endswith(".rels")]:
                root = etree.fromstring(z.read(rf))
                for rel in root.findall(f"{R}Relationship"):
                    if (rel.get("TargetMode") or "").lower() == "external":
                        continue
                    target = rel.get("Target") or ""
                    part = _resolve(rf, target)
                    if part in names:
                        continue
                    rid = rel.get("Id") or ""
                    if rid in refs:
                        bad.append(f"{rf}:{rid}->{part}")
    except Exception:  # noqa: BLE001
        return bad
    return bad
