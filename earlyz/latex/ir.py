"""L2 文档 IR：LaTeX 工程解析后的结构化文档（与输出格式无关；L7 出 docx、L8 出 tex 都从这里取）。

取值约定：
- 公式只存 LaTeX 源（已展开用户宏），转换交给 L4；
- 版式命令（字号/字体/间距/分页）不进 IR：场景 A 的版式由学校模板决定；
- 认不出的内容不丢：块级 → Raw（带原因），命令 → Document.unsupported 计数，行内 → Inline(t="raw")。
"""
from __future__ import annotations

from dataclasses import dataclass, field, fields, is_dataclass


@dataclass
class Inline:
    t: str            # text | math | ref | cite | footnote | link | code | br | image | counter | raw
    v: str = ""       # text 文本 / math LaTeX / ref 标签 / link URL / code 原文 / counter 计数器名
    style: tuple = ()                               # text：'b' 'i' 'tt' 'sc' 'u' 'sup' 'sub'
    data: dict = field(default_factory=dict)        # cite {keys,cmd,pre,post} / ref {cmd} / footnote {blocks}
    #                                                 link {text} / image {path,opts}


@dataclass
class Heading:
    level: int                    # part 0 · chapter 1 · section 2 · subsection 3 · subsubsection 4 · paragraph 5
    name: str                     # LaTeX 命令名（chapter/section…，cls 别名记原名）
    title: list
    numbered: bool = True
    short: list | None = None     # \section[短标题]{…}
    number: str = ""
    labels: list = field(default_factory=list)
    line: int = 0
    role: str = ""                # cls 章级环境的语义（acknowledgements / notation / achievements …）


@dataclass
class Para:
    inlines: list
    align: str = ""               # center | left | right | ""（两端对齐，模板定）
    noindent: bool = False
    cont: bool = False            # 紧跟行间公式、中间没有空行（TeX 不缩进，属于同一段的续行）
    quote: bool = False


@dataclass
class ListItem:
    blocks: list
    label: list | None = None
    labels: list = field(default_factory=list)
    number: str = ""


@dataclass
class ListBlock:
    style: str                    # itemize | enumerate | description
    items: list
    env: str = ""                 # 源环境名（cls 自定义列表如 achievements / bibliolist 记原名）


@dataclass
class Equation:
    latex: str
    env: str                      # equation / align / … / $$ / \[
    numbers: int = 0              # 占用的编号个数（align 多行各自编号）
    labels: list = field(default_factory=list)
    label_rows: list = field(default_factory=list)     # 与 labels 平行：标签所在行（多行环境逐行编号）
    tags: list = field(default_factory=list)          # [[行, 文本]]：\tag 自定编号（不占计数器）
    nonumber_rows: list = field(default_factory=list)  # \nonumber/\notag 所在行
    rows: int = 1
    group: int = 0                                     # subequations 组号（组内编号 = 组号 + a/b/c）
    number: list = field(default_factory=list)         # 每行的显示编号（"" = 不编号）
    line: int = 0


@dataclass
class Cell:
    inlines: list
    colspan: int = 1
    rowspan: int = 1
    align: str = ""


@dataclass
class Row:
    cells: list
    rule_above: str = ""          # hline | toprule | midrule | bottomrule
    rule_below: str = ""
    partial: list = field(default_factory=list)     # cmidrule/cline 列区间 [(a, b)]


@dataclass
class Grid:
    env: str
    colspec: str
    rows: list
    header_rows: int = 0          # longtable \endhead 之前的行（续页重复）


@dataclass
class Float:
    type: str                     # figure | table | algorithm | sub
    env: str
    caption: list = field(default_factory=list)
    caption_alt: list = field(default_factory=list)    # 双语题注的第二语言（\bicaption / Fig.\thefigure 行）
    caption_alt_prefix: str = ""
    short: list | None = None
    labels: list = field(default_factory=list)
    images: list = field(default_factory=list)          # [{path, opts}]
    grids: list = field(default_factory=list)
    subfloats: list = field(default_factory=list)
    notes: list = field(default_factory=list)           # tablenotes 条目（blocks）
    extra: list = field(default_factory=list)           # 浮动体里其余内容（占位框、tikz…）
    numbered: bool = True                               # \caption* / 无题注 → 不占编号
    caption_starred: bool = False                       # 当前题注来自 \caption*（后面再来 \caption 时让位）
    number: str = ""
    line: int = 0


@dataclass
class Theorem:
    env: str
    name: str                     # \newtheorem 给的显示名（定理/引理/Lemma）
    title: list = field(default_factory=list)
    blocks: list = field(default_factory=list)
    labels: list = field(default_factory=list)
    numbered: bool = True
    number: str = ""


@dataclass
class Proof:
    blocks: list
    title: list = field(default_factory=list)


@dataclass
class Code:
    text: str
    env: str
    lang: str = ""


@dataclass
class Abstract:
    lang: str                     # zh | en（按内容判定，不按命令名猜）
    env: str
    blocks: list


@dataclass
class Keywords:
    lang: str
    cmd: str
    items: list


@dataclass
class BibItem:
    key: str
    label: str
    inlines: list


@dataclass
class Bibliography:
    style: str                    # thebibliography | bibtex | biblatex
    items: list = field(default_factory=list)
    files: list = field(default_factory=list)
    bst: str = ""


@dataclass
class Marker:
    name: str                     # toc lof lot frontmatter mainmatter backmatter appendix pagebreak make:<cmd>


@dataclass
class Raw:
    latex: str
    reason: str
    env: str = ""


@dataclass
class Document:
    cls: str
    cls_options: list
    meta: dict
    blocks: list
    top_level: str = "chapter"            # chapter | section（article 类没有章）
    notes: list = field(default_factory=list)
    unsupported: dict = field(default_factory=dict)
    labels: dict = field(default_factory=dict)        # 标签 → {kind, number}（编号按 TeX 规则推算，可与 .aux 对账）
    nocite: list = field(default_factory=list)
    stats: dict = field(default_factory=dict)


def to_dict(obj):
    if is_dataclass(obj):
        d = {"kind": type(obj).__name__}
        for f in fields(obj):
            d[f.name] = to_dict(getattr(obj, f.name))
        return d
    if isinstance(obj, (list, tuple)):
        return [to_dict(x) for x in obj]
    if isinstance(obj, dict):
        return {k: to_dict(v) for k, v in obj.items()}
    return obj


def walk(blocks):
    """深度优先遍历所有块（含列表项、浮动体、定理、脚注里的块）。"""
    for b in blocks:
        yield b
        for name in ("blocks", "extra", "notes", "subfloats"):
            sub = getattr(b, name, None)
            if isinstance(sub, list):
                yield from walk(sub)
        if isinstance(b, ListBlock):
            for it in b.items:
                yield it
                yield from walk(it.blocks)
