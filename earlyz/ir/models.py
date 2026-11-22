"""早睡 (EarlyZ) 3.2 — Document IR 核心对象模型。

对应架构文档 §4。Document IR 是目标 DOCX 的只读语义模型：
- 每个可寻址对象都有稳定 node_id（不等于数组索引）；
- 格式对象保存 direct/style/basedOn/default/theme/grid 级联链；
- 所有模型带 schema_version，版本化演进。

本模块只定义数据模型，不解析 OOXML（解析在 analyzer.py）。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional

IR_SCHEMA_VERSION = "3.2"


# ────────────────────────────────────────────────────────────────
# 基础格式属性（可空字段 = 未声明；None 与 0 语义不同）
# ────────────────────────────────────────────────────────────────

@dataclass
class FormatProperties:
    """一组可比较的格式属性（twips / half-points / pt 与 OOXML 一致）。"""

    font_east_asia: Optional[str] = None
    font_ascii: Optional[str] = None
    font_size_pt: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[str] = None
    color: Optional[str] = None
    jc: Optional[str] = None              # left/center/both/right/distribute
    ind_first_line: Optional[int] = None  # twips
    ind_first_line_chars: Optional[int] = None  # 字符数 × 100
    ind_left: Optional[int] = None
    ind_right: Optional[int] = None
    ind_hanging: Optional[int] = None
    tab_pos: Optional[int] = None         # 首个左对齐制表位 twips（Abstract/分类号标签后）
    spacing_before: Optional[int] = None  # twips
    spacing_after: Optional[int] = None
    spacing_before_lines: Optional[float] = None  # 行单位（模板 beforeLines/100，2026-08-18）
    spacing_after_lines: Optional[float] = None    # 行单位（模板 afterLines/100）
    line: Optional[int] = None            # line 值（240 = 单倍）
    line_rule: Optional[str] = None       # auto/exact/atLeast
    line_spacing: Optional[float] = None  # 倍数行距（line_rule=auto 时 line/240，spec 惯例字段）
    keep_next: Optional[bool] = None
    keep_lines: Optional[bool] = None
    widow_control: Optional[bool] = None
    p_bdr_bottom: Optional[bool] = None   # 段落底边框（页眉下划线）
    page_break_before: Optional[bool] = None
    outline_lvl: Optional[int] = None
    snap_to_grid: Optional[bool] = None
    small_caps: Optional[bool] = None
    drop_cap: Optional[str] = None        # drop / margin（w:framePr/@w:dropCap）
    drop_cap_lines: Optional[int] = None  # 下沉行数
    # 已建模字段之外的可迁移 OOXML。身份/编号/分节属性不进入这里；它们
    # 由各自结构子系统处理。字符串采用规范化 XML，支持精确比较和写回。
    ppr_extra_xml: Optional[str] = None
    rpr_extra_xml: Optional[str] = None

    def matches(self, other: "FormatProperties") -> bool:
        """仅比较两侧都非 None 的属性；任何不一致返回 False。"""
        self_d = {k: v for k, v in self.__dict__.items() if v is not None}
        other_d = {k: v for k, v in other.__dict__.items() if v is not None}
        for k, v in self_d.items():
            if k in other_d and other_d[k] != v:
                return False
        return True

    def diff(self, other: "FormatProperties") -> dict[str, tuple[Any, Any]]:
        """返回 {属性: (期望, 实际)}，仅包含两侧都有且不同的属性。"""
        out: dict[str, tuple[Any, Any]] = {}
        self_d = {k: v for k, v in self.__dict__.items() if v is not None}
        other_d = {k: v for k, v in other.__dict__.items() if v is not None}
        for k, v in self_d.items():
            if k in other_d and other_d[k] != v:
                out[k] = (v, other_d[k])
        return out


@dataclass
class FormatProvenance:
    """某个属性的生效来源（供 explain_format 使用）。"""

    source: str            # direct_formatting/paragraph_style/style_chain/numbering/doc_defaults/theme/doc_grid
    property_name: str
    expected: Optional[Any] = None
    actual: Optional[Any] = None
    style_id: Optional[str] = None
    note: Optional[str] = None


@dataclass
class ResolvedFormat:
    """格式级联解析结果（§4.7）。"""

    direct: FormatProperties = field(default_factory=FormatProperties)
    style: FormatProperties = field(default_factory=FormatProperties)
    based_on: FormatProperties = field(default_factory=FormatProperties)
    numbering: FormatProperties = field(default_factory=FormatProperties)
    default: FormatProperties = field(default_factory=FormatProperties)
    theme: FormatProperties = field(default_factory=FormatProperties)
    effective: FormatProperties = field(default_factory=FormatProperties)
    provenance: list[FormatProvenance] = field(default_factory=list)


# ────────────────────────────────────────────────────────────────
# 定位与身份
# ────────────────────────────────────────────────────────────────

@dataclass
class SourceSpan:
    """对象在 part 中的字节区间或结构索引区间（两种均可）。"""

    start: int = -1
    end: int = -1
    index: Optional[int] = None  # 结构数组索引（仅诊断用，不作为稳定 ID）


@dataclass
class NodeSelector:
    """稳定选择器（§4.4）。node_id 优先；fallback 必须验证类型/文本/上下文。"""

    node_id: Optional[str] = None
    node_type: Optional[str] = None
    semantic_role: Optional[str] = None
    text_hash: Optional[str] = None
    section_id: Optional[str] = None
    occurrence: int = 0
    fallback: Optional[dict] = None  # {normalized_text, near_node_id, max_distance, ...}

    def to_dict(self) -> dict:
        d: dict[str, Any] = {}
        for k in ("node_id", "node_type", "semantic_role", "text_hash",
                  "section_id", "occurrence", "fallback"):
            v = getattr(self, k)
            if v is not None:
                d[k] = v
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "NodeSelector":
        return cls(**{k: v for k, v in d.items() if k in cls.__dataclass_fields__})


# ────────────────────────────────────────────────────────────────
# 通用节点
# ────────────────────────────────────────────────────────────────

@dataclass
class DocumentNode:
    """所有可寻址对象的基类（§4.2）。"""

    node_id: str
    node_type: str
    part_name: str
    structural_path: str = ""
    semantic_role: Optional[str] = None
    text_hash: Optional[str] = None
    content_hash: Optional[str] = None
    parent_id: Optional[str] = None
    children_ids: list[str] = field(default_factory=list)
    source_span: Optional[SourceSpan] = None
    direct_format: FormatProperties = field(default_factory=FormatProperties)
    resolved_format: Optional[ResolvedFormat] = None
    provenance: list[FormatProvenance] = field(default_factory=list)
    diagnostics: list[dict] = field(default_factory=list)
    capabilities: set[str] = field(default_factory=set)

    def to_dict(self) -> dict:
        d = {
            "node_id": self.node_id,
            "node_type": self.node_type,
            "part_name": self.part_name,
            "semantic_role": self.semantic_role,
            "text_hash": self.text_hash,
        }
        return {k: v for k, v in d.items() if v is not None}


@dataclass
class ParagraphNode(DocumentNode):
    node_type: str = "paragraph"
    part_name: str = "word/document.xml"
    text: str = ""
    style_id: Optional[str] = None
    outline_lvl: Optional[int] = None
    list_binding: Optional[dict] = None  # {num_id, ilvl}
    runs: list["RunNode"] = field(default_factory=list)
    is_heading: bool = False
    heading_level: Optional[int] = None
    pstyle_skipped: bool = False


@dataclass
class RunNode(DocumentNode):
    node_type: str = "run"
    part_name: str = "word/document.xml"
    text: str = ""
    rstyle_id: Optional[str] = None
    lang: Optional[str] = None
    is_field_member: bool = False
    has_tab: bool = False  # run 内含 w:tab（标签后分隔符，模板标签段自带）
    revision_state: Optional[str] = None  # insert/delete/format/none


@dataclass
class SectionNode(DocumentNode):
    node_type: str = "section"
    part_name: str = "word/document.xml"
    index: int = -1
    break_type: Optional[str] = None      # nextPage/continuous/oddPage/evenPage
    page_size: Optional[dict] = None      # {w,h,orient}
    margins: Optional[dict] = None        # {top,bottom,left,right,gutter,header,footer}
    columns: Optional[dict] = None
    pg_num_type: Optional[dict] = None    # {fmt,start}
    header_footers: list[dict] = field(default_factory=list)
    section_extra_xml: Optional[str] = None  # portable sectPr formatting bundle
    physical_start_page: Optional[int] = None
    governing_heading: Optional[str] = None


@dataclass
class StyleNode(DocumentNode):
    node_type: str = "style"
    part_name: str = "word/styles.xml"
    style_id: str = ""
    style_type: str = "paragraph"         # paragraph/character/table/numbering
    name: Optional[str] = None
    based_on: Optional[str] = None
    next: Optional[str] = None
    link: Optional[str] = None
    hidden: bool = False
    is_default: bool = False
    style_format: FormatProperties = field(default_factory=FormatProperties)
    inherits_from: list[str] = field(default_factory=list)


@dataclass
class ListLevelNode(DocumentNode):
    node_type: str = "list_level"
    part_name: str = "word/numbering.xml"
    level: int = 0
    num_fmt: Optional[str] = None
    lvl_text: Optional[str] = None
    start: Optional[int] = None
    indent: Optional[dict] = None
    style_link: Optional[str] = None
    level_extra_xml: Optional[str] = None  # portable lvl pPr/rPr/suffix/alignment bundle


@dataclass
class ListDefinitionNode(DocumentNode):
    node_type: str = "list_definition"
    part_name: str = "word/numbering.xml"
    num_id: Optional[int] = None
    abstract_num_id: Optional[int] = None
    levels: list[ListLevelNode] = field(default_factory=list)
    multi_level: bool = False
    restart_source: Optional[str] = None


@dataclass
class TableNode(DocumentNode):
    node_type: str = "table"
    part_name: str = "word/document.xml"
    index: int = -1
    row_count: int = 0
    col_count: int = 0
    grid_widths: list[int] = field(default_factory=list)
    layout: Optional[str] = None          # fixed/autofit
    alignment: Optional[str] = None
    tblpr_extra_xml: Optional[str] = None  # portable non-geometry tblPr bundle
    header_rows: list[int] = field(default_factory=list)
    merged_cells: list[dict] = field(default_factory=list)
    caption_node_id: Optional[str] = None
    overflow_risk: Optional[dict] = None  # {type, page, detail}


@dataclass
class FigureNode(DocumentNode):
    node_type: str = "figure"
    part_name: str = "word/document.xml"
    index: int = -1
    media_rel: Optional[str] = None
    docpr_id: Optional[str] = None        # wp:docPr @id（包级唯一，§7.1 repair_package_invariants）
    extent: Optional[dict] = None         # {cx, cy} EMU
    wrap: Optional[str] = None            # inline/float
    anchor: Optional[dict] = None
    caption_node_id: Optional[str] = None
    alt_text: Optional[str] = None
    bbox: Optional[dict] = None           # {x0,y0,x1,y1,page}
    ole_original: bool = False


@dataclass
class FieldNode(DocumentNode):
    node_type: str = "field"
    part_name: str = "word/document.xml"
    field_type: str = ""                  # PAGE/PAGEREF/TOC/REF/SEQ/STYLEREF/HYPERLINK/...
    instruction: Optional[str] = None
    result: Optional[str] = None
    complete: bool = True                 # begin/separate/end 三件套是否完整
    targets: list[str] = field(default_factory=list)


@dataclass
class CaptionNode(DocumentNode):
    node_type: str = "caption"
    part_name: str = "word/document.xml"
    caption_type: str = ""                # figure/table/equation
    label: Optional[str] = None           # 图/表/公式
    number: Optional[str] = None          # 3-1 / 2.3
    seq_field: Optional[str] = None
    cross_references: list[str] = field(default_factory=list)


# ────────────────────────────────────────────────────────────────
# IR 聚合根
# ────────────────────────────────────────────────────────────────

@dataclass
class DocumentIR:
    """整个 DOCX 的语义模型（§4.5）。"""

    schema_version: str = IR_SCHEMA_VERSION
    document_id: str = ""
    package: Optional[dict] = None         # PackageHealth
    metadata: Optional[dict] = None
    settings: Optional[dict] = None
    theme: Optional[dict] = None
    font_table: Optional[dict] = None
    styles: list[StyleNode] = field(default_factory=list)
    numbering: list[ListDefinitionNode] = field(default_factory=list)
    sections: list[SectionNode] = field(default_factory=list)
    blocks: list[DocumentNode] = field(default_factory=list)
    paragraphs: list[ParagraphNode] = field(default_factory=list)
    tables: list[TableNode] = field(default_factory=list)
    headers: dict = field(default_factory=dict)  # header part → [{text, fp}]（页眉格式规则）
    footers: dict = field(default_factory=dict)  # footer part → [{text, fp}]
    figures: list[FigureNode] = field(default_factory=list)
    fields: list[FieldNode] = field(default_factory=list)
    captions: list[CaptionNode] = field(default_factory=list)
    bookmarks: list[dict] = field(default_factory=list)
    hyperlinks: list[dict] = field(default_factory=list)
    notes: Optional[dict] = None
    revisions: Optional[dict] = None
    content_controls: list[dict] = field(default_factory=list)
    render_map: Optional[list[dict]] = None
    diagnostics: list[dict] = field(default_factory=list)
    capabilities: set[str] = field(default_factory=set)

    def node_by_id(self, node_id: str) -> Optional[DocumentNode]:
        for coll in (self.blocks, self.paragraphs, self.tables, self.figures,
                     self.fields, self.captions, self.sections):
            for n in coll:
                if n.node_id == node_id:
                    return n
        return None

    def nodes_by_type(self, node_type: str) -> list[DocumentNode]:
        out = []
        for coll in (self.blocks, self.paragraphs, self.tables, self.figures,
                     self.fields, self.captions, self.sections):
            out.extend(n for n in coll if n.node_type == node_type)
        return out

    def to_dict(self) -> dict:
        return {
            "schema_version": self.schema_version,
            "document_id": self.document_id,
            "sections": [s.index for s in self.sections],
            "paragraphs": len(self.paragraphs),
            "tables": len(self.tables),
            "figures": len(self.figures),
            "fields": len(self.fields),
            "styles": len(self.styles),
            "diagnostics": len(self.diagnostics),
        }
