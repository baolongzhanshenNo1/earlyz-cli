"""早睡 (EarlyZ) 3.2 — 包健康检查（§4.5 package / §11.1 Static Package Gate）。

检查：
- ZIP 可读性；
- [Content_Types].xml 存在；
- 重复 ZIP entry；
- rels 引用的 part 是否存在；
- 主要 part 是否缺失。
"""

from __future__ import annotations

import hashlib
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

import re

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT_NS = "{http://schemas.openxmlformats.org/package/2006/content-types}"

REQUIRED_PARTS = {"[Content_Types].xml", "_rels/.rels", "word/document.xml", "word/styles.xml"}


@dataclass
class PackageHealth:
    valid: bool = True
    errors: list[str] = field(default_factory=list)
    duplicate_entries: list[str] = field(default_factory=list)
    missing_referenced_parts: list[str] = field(default_factory=list)
    parts: list[str] = field(default_factory=list)
    relationships: dict[str, str] = field(default_factory=dict)

    def to_dict(self) -> dict:
        return {
            "valid": self.valid,
            "errors": self.errors,
            "duplicate_entries": self.duplicate_entries,
            "missing_referenced_parts": self.missing_referenced_parts,
            "part_count": len(self.parts),
        }


def inspect_package(path: str | Path) -> PackageHealth:
    health = PackageHealth()
    p = Path(path)
    try:
        z = zipfile.ZipFile(p)
    except (zipfile.BadZipFile, OSError) as e:
        raise ValueError(f"not a valid zip package: {p} ({e})") from e

    with z:
        names = z.namelist()
        health.parts = names
        seen: set[str] = set()
        for name in names:
            if name in seen:
                health.duplicate_entries.append(name)
                health.errors.append(f"duplicate zip entry: {name}")
            seen.add(name)

        for required in REQUIRED_PARTS:
            if required not in seen:
                health.errors.append(f"missing required part: {required}")

        # 解析 _rels/.rels 与 word/_rels/document.xml.rels 引用的 part
        try:
            root_rels = ET.fromstring(z.read("_rels/.rels").decode("utf-8", errors="replace"))
            for rel in root_rels:
                target = rel.get("Target", "")
                if not target.startswith("/") and not target.startswith(("http://", "https://")):
                    part = "/" + target
                else:
                    part = target
                if part.startswith("/"):
                    part = part[1:]
                if target and not target.startswith(("http://", "https://")):
                    health.relationships[rel.get("Id", "")] = part
                    if part not in seen:
                        health.missing_referenced_parts.append(part)
        except (KeyError, ET.ParseError):
            health.errors.append("missing or malformed _rels/.rels")

        try:
            doc_rels = ET.fromstring(
                z.read("word/_rels/document.xml.rels").decode("utf-8", errors="replace")
            )
            base = "word/"
            for rel in doc_rels:
                target = rel.get("Target", "")
                if target.startswith(("http://", "https://")):
                    continue
                part = base + target if not target.startswith("/") else target[1:]
                health.relationships[rel.get("Id", "")] = part
                if part not in seen:
                    health.missing_referenced_parts.append(part)
        except KeyError:
            pass  # 文档关系缺失单独由 OOXML gate 报告
        except ET.ParseError:
            health.errors.append("malformed word/_rels/document.xml.rels")

    health.valid = not health.errors and not health.duplicate_entries
    return health


# ────────────────────────────────────────────────────────────────
# 文本与内容 hash（内容保持不变量，§10.1）
# ────────────────────────────────────────────────────────────────

_WT_RE = re.compile(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>")
# 只保留基础 XML 解码（实体），避免引入额外依赖
_ENTITIES = {"&amp;": "&", "&lt;": "<", "&gt;": ">", "&quot;": '"', "&apos;": "'"}


def content_text(document_xml: str) -> str:
    """提取 w:t 文本（简单正则，适用于 hash 与统计；完整解析在 analyzer）。"""
    out = []
    for m in _WT_RE.finditer(document_xml):
        t = m.group(1)
        for k, v in _ENTITIES.items():
            t = t.replace(k, v)
        out.append(t)
    return "".join(out)


def content_hash(document_xml: str) -> str:
    """包级快照 hash：w:t 拼接 → sha1。仅测试/包级快照用（tests/test_package_health.py）。

    口径 ≠ 交付门禁的 analyzer 版：本函数不豁免 TOC/refs（目录条目、参考文献
    编号），含全部空白；交付判定请用早期解析 IR 的 metadata['content_hash']
    （analyzer._normalized_content_hash 口径）。
    """
    return hashlib.sha1(content_text(document_xml).encode("utf-8")).hexdigest()
