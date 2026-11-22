"""Full-package OOXML inventory and disposition matrix.

Every package member and every XML element/attribute is discoverable.  The
matrix deliberately separates portable formatting from content, identity,
relationships and producer history: "100% inventoried" never means blindly
copying every byte from a template into a target document.
"""
from __future__ import annotations

import hashlib
import json
import re
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from lxml import etree

_NS_ALIAS = {
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main": "w",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships": "r",
    "http://schemas.openxmlformats.org/package/2006/relationships": "pr",
    "http://schemas.openxmlformats.org/package/2006/content-types": "ct",
    "http://schemas.openxmlformats.org/drawingml/2006/main": "a",
    "http://schemas.openxmlformats.org/drawingml/2006/picture": "pic",
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing": "wp",
    "http://schemas.openxmlformats.org/wordprocessingml/2006/wordprocessingDrawing": "wp14",
    "http://schemas.openxmlformats.org/officeDocument/2006/math": "m",
    "urn:schemas-microsoft-com:vml": "v",
    "urn:schemas-microsoft-com:office:office": "o",
    "http://schemas.openxmlformats.org/markup-compatibility/2006": "mc",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas": "wpc",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape": "wps",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup": "wpg",
}


def _qname(name: str) -> str:
    if not isinstance(name, str) or not name.startswith("{"):
        return str(name)
    uri, local = name[1:].split("}", 1)
    alias = _NS_ALIAS.get(uri)
    if alias is None:
        alias = "ns_" + hashlib.sha1(uri.encode("utf-8")).hexdigest()[:7]
    return f"{alias}:{local}"


def _local(name: str) -> str:
    return name.rsplit("}", 1)[-1] if "}" in name else name.split(":")[-1]


def part_class(name: str) -> str:
    """Normalize numbered parts while retaining their engineering role."""
    if name == "[Content_Types].xml":
        return "[Content_Types].xml"
    if name == "_rels/.rels":
        return "package.rels"
    rules = (
        (r"word/header\d+\.xml", "word/header*.xml"),
        (r"word/footer\d+\.xml", "word/footer*.xml"),
        (r"word/_rels/header\d+\.xml\.rels", "word/header*.xml.rels"),
        (r"word/_rels/footer\d+\.xml\.rels", "word/footer*.xml.rels"),
        (r"word/charts/chart\d+\.xml", "word/charts/chart*.xml"),
        (r"word/diagrams/[^/]+\d+\.xml", "word/diagrams/*.xml"),
        (r"customXml/item\d+\.xml", "customXml/item*.xml"),
        (r"customXml/_rels/item\d+\.xml\.rels", "customXml/item*.xml.rels"),
        (r"customXml/itemProps\d+\.xml", "customXml/itemProps*.xml"),
    )
    for pattern, label in rules:
        if re.fullmatch(pattern, name):
            return label
    return name


def _content_types(z: zipfile.ZipFile) -> dict[str, str]:
    result: dict[str, str] = {}
    if "[Content_Types].xml" not in z.namelist():
        return result
    root = etree.fromstring(z.read("[Content_Types].xml"))
    defaults: dict[str, str] = {}
    for el in root:
        local = _local(el.tag)
        if local == "Default":
            defaults[(el.get("Extension") or "").lower()] = el.get("ContentType") or ""
        elif local == "Override":
            result[(el.get("PartName") or "").lstrip("/")] = el.get("ContentType") or ""
    for name in z.namelist():
        extension = "rels" if name.endswith(".rels") else Path(name).suffix.lstrip(".").lower()
        result.setdefault(name, defaults.get(extension, ""))
    return result


def scan_package(path: str | Path, value_limit: int = 24) -> dict[str, Any]:
    """Inventory every member plus every parseable XML element and attribute."""
    src = Path(path)
    paths: dict[str, dict[str, Any]] = {}
    members: list[dict[str, Any]] = []

    def record(key: str, value: str | None, member: str) -> None:
        row = paths.setdefault(key, {"count": 0, "values": Counter(), "parts": Counter()})
        row["count"] += 1
        row["parts"][member] += 1
        if value is not None and len(row["values"]) < value_limit:
            row["values"][str(value)] += 1

    with zipfile.ZipFile(src) as z:
        ctypes = _content_types(z)
        for info in z.infolist():
            name = info.filename
            data = z.read(name)
            cls = part_class(name)
            is_xml = name.endswith((".xml", ".rels")) or name == "[Content_Types].xml"
            members.append({
                "name": name,
                "class": cls,
                "size": len(data),
                "compress_type": info.compress_type,
                "content_type": ctypes.get(name, ""),
                "sha256": hashlib.sha256(data).hexdigest(),
                "kind": "xml" if is_xml else "binary",
            })
            record(f"member::{cls}/#present", None, name)
            record(f"member::{cls}/@contentType", ctypes.get(name, ""), name)
            if not is_xml:
                continue
            try:
                root = etree.fromstring(data)
            except (etree.XMLSyntaxError, ValueError):
                record(f"member::{cls}/#unparseable", None, name)
                continue

            def walk(el, parents: tuple[str, ...] = ()) -> None:
                here = (*parents, _qname(el.tag))
                xpath = "/".join(here)
                record(f"xml::{cls}::{xpath}/#present", None, name)
                for attr, value in sorted(el.attrib.items(), key=lambda kv: _qname(kv[0])):
                    record(f"xml::{cls}::{xpath}/@{_qname(attr)}", value, name)
                # Field instructions are engineering structure; ordinary w:t is content.
                if _local(el.tag) in {"instrText", "fldSimple"} and (el.text or "").strip():
                    record(f"xml::{cls}::{xpath}/#instruction", (el.text or "").strip(), name)
                for child in el:
                    if isinstance(child.tag, str):
                        walk(child, here)

            walk(root)

    serial_paths = {}
    for key, row in sorted(paths.items()):
        serial_paths[key] = {
            "count": row["count"],
            "values": dict(row["values"]),
            "parts": dict(row["parts"]),
        }
    return {
        "file": str(src),
        "size": src.stat().st_size,
        "member_count": len(members),
        "path_count": len(serial_paths),
        "members": members,
        "paths": serial_paths,
    }


def disposition(path_key: str) -> dict[str, str]:
    """Classify one path. Every return is deliberate; no silent fall-through."""
    low = path_key.lower()
    part = ""
    xpath = ""
    if path_key.startswith("xml::"):
        _, part, xpath = path_key.split("::", 2)
    elif path_key.startswith("member::"):
        part = path_key.split("::", 1)[1].split("/", 1)[0]

    attr = path_key.rsplit("/@", 1)[-1] if "/@" in path_key else ""
    attr_local = attr.split(":")[-1]

    # Producer/edit history is inventoried but never a formatting rule.
    if ("/w:rsids/" in low or attr_local.lower().startswith("rsid")
            or attr_local in {"paraId", "textId", "durableId"}):
        return _d("producer_history", "intentionally_ignore", "closed",
                  "Word revision/session identity; copying would forge edit history")

    # Relationship/object identities must be remapped, not copied.
    if part.endswith(".rels") or part in {"package.rels"}:
        return _d("relationships", "regenerate_and_remap", "closed",
                  "relationship IDs and targets are package-local")
    if attr_local in {"id", "embed", "link", "anchorId", "editId", "docId"}:
        return _d("object_identity", "regenerate_and_remap", "closed",
                  "object/relationship identifiers must be unique in the target package")

    if path_key.startswith("member::"):
        if any(x in low for x in ("media", "embedding", "oleobject", "image")):
            return _d("binary_content", "preserve_target_content", "closed",
                      "template sample payload is content, not a formatting rule")
        return _d("package_member", "preserve_or_rebuild_by_owner", "closed",
                  "member lifecycle is owned by its structural serializer")

    if part == "[Content_Types].xml":
        return _d("package_contract", "regenerate_from_parts", "closed",
                  "content types follow the actual target package members")

    # Drawing/VML/text-box geometry belongs to each target object. IDs are handled above.
    if any(token in xpath for token in ("w:drawing", "wp:anchor", "wp:inline", "v:shape",
                                         "w:object", "wps:", "wpc:", "pic:pic")):
        return _d("content_object_geometry", "preserve_target_content", "closed",
                  "object geometry/crop/wrap is content-specific unless a separate object rule requests change")

    # Body paragraph/run formatting: typed fields + canonical lossless bundle.
    if part == "word/document.xml" and ("/w:pPr" in xpath or "/w:rPr" in xpath):
        if "/w:sectPr" in xpath:
            return _d("section_format", "typed_common_bundle_plus_semantic_structure", "closed",
                      "common portable sectPr properties use a lossless bundle; type/columns/page numbering/references use semantic serializers")
        if any(x in xpath for x in ("/w:tbl/", "/w:tc/", "/w:tr/")):
            return _d("table_cell_format", "preserve_target_then_apply_explicit_table_rules", "closed",
                      "cell-local formatting follows target content; template-wide fonts/borders/width policies use explicit table rules")
        if any(x in xpath for x in ("/wps:", "/wpc:", "/w:txbxContent")):
            return _d("text_box_content", "preserve_target_content", "closed",
                      "template text-box contents and local formatting stay with the target object")
        if any(x in xpath for x in ("/w:pStyle", "/w:numPr", "/w:rStyle")):
            return _d("style_or_numbering_binding", "preserve_target_binding_and_semantic_rebind", "closed",
                      "raw IDs stay target-local; effective styles are materialized and numbering levels are resolved through target numId mappings")
        return _d("paragraph_run_format", "typed_plus_lossless_bundle", "closed",
                  "covered by FormatProperties + ppr/rpr canonical OOXML bundles")

    if part == "word/styles.xml":
        if "/w:pPr" in xpath or "/w:rPr" in xpath or "/w:docDefaults" in xpath:
            return _d("style_cascade_format", "resolve_and_materialize", "closed",
                      "basedOn/docDefaults are resolved; effective role format is materialized on target paragraphs")
        return _d("style_identity_ui", "preserve_target_or_rebind", "closed",
                  "style names/UI metadata are identities, not universal role formatting")

    if part == "word/numbering.xml":
        if any(x in low for x in ("nsid", "/w:tmpl", "abstractnumid", "/w:numid")):
            return _d("numbering_identity", "regenerate_and_remap", "closed",
                      "numbering IDs/templates are package-local")
        return _d("numbering_semantics", "semantic_role_level_bundle_or_preserve_target", "closed",
                  "homogeneous bound roles migrate numFmt/lvlText/start plus portable lvl bundle; ambiguous or unused definitions stay target-owned")

    if part in {"word/theme/theme1.xml", "word/fontTable.xml"}:
        return _d("theme_font_semantics", "preserve_target_theme_asset", "closed",
                  "theme/fontTable definitions are package assets, not role rules; current templates have no body-role theme-font reference to migrate")

    if part == "word/settings.xml":
        if any(x in low for x in ("proofstate", "stylepane", "docvars", "view", "zoom")):
            return _d("editor_ui_state", "preserve_target", "closed",
                      "editor/proofing UI state is not template formatting")
        if "/w:compat" in xpath:
            return _d("compatibility_mode", "preserve_target", "closed",
                      "compat flags belong to the target producer/version unless explicitly requested")
        return _d("document_behavior", "preserve_target_unless_explicit_structure_rule", "closed",
                  "settings remain target-owned; even/odd headers, field updates and other requested behavior use explicit verified actions")

    if part == "word/document.xml" and "/w:sectPr" in xpath:
        return _d("section_format", "typed_common_bundle_plus_semantic_structure", "closed",
                  "section geometry/common formatting and semantic columns/page/reference actions have separate owners")
    if part == "word/document.xml" and any(x in xpath for x in ("/w:tbl", "/w:tr", "/w:tc")):
        return _d("table_structure", "preserve_target_geometry_plus_common_table_bundle", "closed",
                  "target grid/merge/content geometry is preserved; portable common tblPr and explicit border/width/font rules are verified")

    if part.startswith("word/header") or part.startswith("word/footer"):
        if "/w:pPr" in xpath or "/w:rPr" in xpath:
            return _d("header_footer_format", "typed_plus_lossless_bundle", "closed",
                      "common header/footer pPr/rPr formats use the same typed+extras channel as body text without copying content or fields")
        return _d("header_footer_content", "preserve_or_rebuild_fields", "closed",
                  "header/footer text and PAGE fields are structural content, not body role samples")

    if part in {"word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"}:
        return _d("note_comment_structure", "preserve_target_content", "closed",
                  "notes/comments belong to target content; references remain package-local")

    if part in {"word/webSettings.xml", "docProps/core.xml", "docProps/app.xml",
                "docProps/custom.xml"}:
        return _d("producer_metadata", "preserve_target", "closed",
                  "metadata is inventoried but is not a formatting rule")

    if any(x in xpath for x in ("w:fldChar", "w:instrText", "w:fldSimple", "w:bookmark",
                                 "w:hyperlink", "w:sdt", "w:comment", "w:footnote", "w:endnote")):
        return _d("complex_content_structure", "preserve_target_content", "closed",
                  "field/bookmark/control/note structure is target content; dedicated repairs are explicit")

    if part == "word/document.xml":
        return _d("document_content", "preserve_target_content", "closed",
                  "ordinary body XML is content, not template formatting")

    return _d("package_extension", "preserve_target_or_owner_managed", "closed",
              "inventoried extension part; no implicit template-to-target copy")


def _d(domain: str, decision: str, coverage: str, reason: str) -> dict[str, str]:
    return {"domain": domain, "decision": decision, "coverage": coverage, "reason": reason}


_PIPELINE_CONTRACTS = {
    "paragraph_run_format": {
        "ir": "FormatProperties.ppr_extra_xml/rpr_extra_xml + typed fields",
        "spec": "decompiler semantic-role/cover rules",
        "planner": "violations grouped by stable paragraph target",
        "writer": "set_paragraph_format",
        "verifier": "verify_paragraph_format + second lint",
        "render": "four-template Word COM/PDF visual gate",
    },
    "style_cascade_format": {
        "ir": "StyleNode + basedOn/docDefaults resolved format",
        "spec": "decompiler style fallback and effective role materialization",
        "planner": "role/style violations become materialize-or-style actions",
        "writer": "create_or_update_style/set_paragraph_format",
        "verifier": "style/role lint + action verifier",
        "render": "four-template Word COM/PDF visual gate",
    },
    "section_format": {
        "ir": "SectionNode typed geometry + section_extra_xml",
        "spec": "section_all common bundle + structure semantic plan",
        "planner": "section_all violations grouped into structural actions",
        "writer": "set_page_margins_all/apply_section_columns/page/reference actions",
        "verifier": "verify_page_margins_all + structure verifiers + second lint",
        "render": "Word COM open plus PDF pagination/header/footer gate",
    },
    "table_structure": {
        "ir": "TableNode typed geometry + tblpr_extra_xml",
        "spec": "table_all common bundle + structural table parameters",
        "planner": "table_all violations grouped by portable property set",
        "writer": "set_table_alignment + explicit width/border/font actions",
        "verifier": "verify_table_alignment + table action verifiers + second lint",
        "render": "Word COM open plus PDF/PNG table visual gate",
    },
    "table_cell_format": {
        "ir": "target table-cell content/format inventory",
        "spec": "explicit table cell font/border policy only",
        "planner": "explicit table policy actions; unspecified cells excluded",
        "writer": "set_tbl_font/set_tbl_header_font/set_tbl_cell_valign",
        "verifier": "table action verifiers; unspecified cell-local format preserved",
        "render": "Word COM open plus PDF/PNG table visual gate",
    },
    "header_footer_format": {
        "ir": "DocumentIR.headers/footers FormatProperties",
        "spec": "header_all/footer_all common typed+extras rules",
        "planner": "header/footer scope violations become part-wide format actions",
        "writer": "format_headers/format_footers",
        "verifier": "verify_headers/verify_footers + second lint",
        "render": "Word COM open plus PDF header/footer text/field gate",
    },
    "numbering_semantics": {
        "ir": "ListDefinitionNode/ListLevelNode.level_extra_xml",
        "spec": "numbering_role homogeneous-level rules",
        "planner": "semantic role binding remapped to target abstractNumId/ilvl",
        "writer": "set_list_level_format against target abstractNumId",
        "verifier": "verify_list_level_format + numbering_role lint",
        "render": "Word COM open plus PDF numbering-text gate",
    },
    "style_or_numbering_binding": {
        "ir": "ParagraphNode style_id/list_binding",
        "spec": "semantic role rules; raw IDs excluded",
        "planner": "explicit semantic rebind only; otherwise preservation invariant",
        "writer": "preserve target binding; map_style/map_paragraph_to_list only when explicit",
        "verifier": "binding/action verifier or preservation invariant",
        "render": "Word COM open and affected-role visual gate",
    },
}


def pipeline_contract(domain: str, decision: str) -> dict[str, str]:
    contract = _PIPELINE_CONTRACTS.get(domain)
    if contract:
        return dict(contract)
    if decision.startswith("preserve"):
        return {
            "ir": "inventory_only_or_target_owner",
            "spec": "not_applicable_by_preservation_policy",
            "planner": "no_migration_action_by_policy",
            "writer": "preserve_target",
            "verifier": "content/package invariant",
            "render": "Word normal-open gate; visual comparison when layout-bearing",
        }
    if decision.startswith("regenerate"):
        return {
            "ir": "package inventory/relationship graph",
            "spec": "not_applicable_by_identity_policy",
            "planner": "package owner action only when a referenced part is rebuilt",
            "writer": "package owner regenerates/remaps",
            "verifier": "package invariant verifier",
            "render": "Word normal-open gate",
        }
    if decision == "intentionally_ignore":
        return {
            "ir": "inventory_only",
            "spec": "not_applicable_by_history_policy",
            "planner": "excluded_from_migration_plan",
            "writer": "do_not_copy",
            "verifier": "absence from migration plan",
            "render": "not_applicable; Word normal-open gate covers package validity",
        }
    return {
        "ir": "domain structural owner",
        "spec": "domain policy/explicit rule",
        "planner": "domain policy action or explicit no-action disposition",
        "writer": "domain action owner",
        "verifier": "domain verifier or preservation invariant",
        "render": "Word normal-open gate and domain visual gate when layout-bearing",
    }


def build_matrix(inventories: dict[str, dict[str, Any]]) -> dict[str, Any]:
    all_paths = sorted({p for inv in inventories.values() for p in inv["paths"]})
    rows = []
    coverage = Counter()
    decisions = Counter()
    domains = Counter()
    for key in all_paths:
        disp = disposition(key)
        coverage[disp["coverage"]] += 1
        decisions[disp["decision"]] += 1
        domains[disp["domain"]] += 1
        rows.append({
            "path": key,
            **disp,
            "pipeline": pipeline_contract(disp["domain"], disp["decision"]),
            "templates": {
                name: inv["paths"].get(key) for name, inv in inventories.items()
                if key in inv["paths"]
            },
        })
    return {
        "template_count": len(inventories),
        "union_path_count": len(rows),
        "unclassified_count": sum(1 for row in rows if not row.get("decision")),
        "unowned_pipeline_count": sum(
            1 for row in rows
            if any(not row.get("pipeline", {}).get(key)
                   for key in ("ir", "spec", "planner", "writer", "verifier", "render"))
        ),
        "coverage_counts": dict(coverage),
        "decision_counts": dict(decisions),
        "domain_counts": dict(domains),
        "rows": rows,
    }


def scan_and_build(paths: dict[str, str | Path]) -> tuple[dict[str, Any], dict[str, Any]]:
    inventories = {name: scan_package(path) for name, path in paths.items()}
    return inventories, build_matrix(inventories)


def write_report(paths: dict[str, str | Path], output_dir: str | Path) -> dict[str, Any]:
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    inventories, matrix = scan_and_build(paths)
    for name, inv in inventories.items():
        (out / f"{name}.json").write_text(
            json.dumps(inv, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    (out / "matrix.json").write_text(
        json.dumps(matrix, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    summary = {k: v for k, v in matrix.items() if k != "rows"}
    summary["templates"] = {
        name: {"file": inv["file"], "members": inv["member_count"], "paths": inv["path_count"]}
        for name, inv in inventories.items()
    }
    (out / "summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return summary
