"""版式保真判定（EZI-0005）：把「产物 vs 模板」逐项判成 PASS / FAIL / N-A。

纯只读：用 ``zipfile`` + ``lxml`` 直读 OOXML，不碰 Word COM，不改任何排版行为。
CLI 壳在 ``scripts/check_layout_fidelity.py``，这里只放可 import 的判定逻辑。

判定口径（七类，全部从 OOXML 取实测值）：

============  ==================================================================
标题          一级/二级/三级字号、标题中文字体、标题对齐。
              标题段落 = ``w:outlineLvl`` 经 ``w:pStyle``→``w:basedOn`` 链解析后等于
              level-1 的段落；字号/字体/对齐取段内首个有文字的 run 的有效属性。
              **取样口径（EZI-0012 A7）**：``outlineLvl`` 级别匹配 **且** 文本是
              编号标题（``第X章`` / ``1`` / ``1.1`` / ``一、`` / ``（一）``）才算样本；
              再排除表格内段落、后置节名（ABSTRACT/REFERENCES/APPENDIX…）、表单标签
              （冒号结尾短行）、封面拉丁大写行、**手写目录条目**（EZI-0041：文本形如
              ``1.1.\\t编写目的\\t4``，制表符分段 + 末段是页码，且不在真 TOC 域内）。
              **禁止**拿「全部 outlineLvl=0 段落的
              众数」当事实——交大中期 3 段里 2 段是校名行与表单字段，ACM 的众数被
              ABSTRACT/REFERENCES 带偏。排除后无样本 → ``N-A``。
              **适用性判据（EZI-0009）**：判定前先问源文档有没有这个角色——
              源文档该级编号标题样本数为 0 时三项一律 ``N-A``（产物没有该级段落是
              正确行为）；不给 ``source`` 且产物也没有样本同样 ``N-A``（不猜）。
              源文档有该级、产物没有 = 迁移丢失，一级判 ``FAIL``（二/三级沿用旧判法）。
              **期望值来源（EZI-0016）**：模板该级样本的**实测众数**优先（同一套取样
              函数，`_template_mode`），不再无条件信 ``spec.rules.headingN``——spec 的
              标题规则来自编译期的样式/簇取样，新疆模板三级 17/17 = center 却被 spec
              写成 both。两道护栏：模板样本 < ``TEMPLATE_MODE_MIN_SAMPLES`` 个或众数占比
              < ``TEMPLATE_MODE_MIN_SHARE`` = 没有主峰，退回 spec；模板/产物同级样本的
              **文本重合率** < ``HEADING_OVERLAP_MIN`` = 两侧比的不是同一批段落
              （模板「第X章」vs 产物条款项），同样退回 spec。note 逐项写明期望出处。
              **同级同文（EZI-0021 / EZI-0025）**：产物该级样本里混着「模板里根本不是
              标题」的段落时（重合率 < ``HEADING_OVERLAP_MIN``），**字号、字体、对齐三项
              一起**收窄到「与模板同级同文」的那批（``_comparable_heading_rows``），
              产物全体的众数不再当该级实测值；收窄后可比样本为空 → 三项 ``N-A``，绝不
              退回全量样本。期望侧同口径：模板**同文成对**的那批段落的该属性**实测众数**
              即期望（≥ ``HEADING_SAME_TEXT_MIN`` 段；1:1 对应不是统计推断，不受
              ``TEMPLATE_MODE_MIN_SAMPLES`` 限制）。**EZI-0039 起取众数即可，不再要求
              同文段 100% 一致**：最高频占比 ≥ ``TEMPLATE_MODE_MIN_SHARE``（0.60，
              与 ``spec/decompiler.py::_DOMINANCE_DEFAULT`` 同一口径）就采信——深圳招标
              一级标题模板同文 16.0pt×9 + 22.0pt×1（90%）曾经被判「不统一」退回 spec
              14pt，产物照模板排 16pt 反被判 FAIL。无单一主峰时字号取前二高频当双峰区间
              （地板同为 60%，同正文 EZI-0022），前二合计仍 <60% 才退回原护栏，并在
              note 写明「不统一（分布）／最高频／前二合计」的依据；字体/对齐是类别属性，
              无单一主峰直接退回原护栏（不造候选集合）。
正文          字号、中文字体、首行缩进（``w:ind/@w:firstLineChars``，缺则用
              ``@w:firstLine`` 按字号折算）、行距（**两侧同口径**：正文段落
              ``w:spacing/@w:line`` + ``@w:lineRule`` 的**设值**众数，排除表格/标题/
              目录条目/**表单标签**（EZI-0020：冒号结尾短行「学生姓名:」是字段名，不是
              正文，与标题取样共用同一条规则）；模板正文段落与默认段落样式都没设行距
              → ``N-A``，绝不凭空造期望值，EZI-0015）、对齐（``w:jc``，未设置按
              ``start`` 计）。正文取样还要排掉 **手写目录条目**及其列出的标题（EZI-0041：
              手写目录是这类源文档的标题索引，正文流里与条目同名的段落是标题，不是正文段）。
              正文**主字号**的期望值取模板正文段落的**实测众数**（同一 ``
              body_text_paragraphs`` 取样，排除标题/表格/题注/目录/空段），**不用**
              ``spec.rules.body.font_size_pt``——那是编译期从注解/样式推的值，本科模板
              正文实测 12pt×48/77 却被 spec 写成 14pt（EZI-0016）。主峰占比 <
              ``TEMPLATE_MODE_MIN_SHARE``（无主峰）时**不再退回 spec 注解**——交大中期
              16×7/14×6（50%）、开题 14×8/16×6（47%）就是模板自己混排，退回 spec 16pt
              等于用注解常量冒充实测；改判**双峰区间**：候选 = 前二高频，产物落在任一
              即 PASS，前二合计也 <60% 则 ``N-A``（EZI-0022）。只有样本不足（<
              ``TEMPLATE_MODE_MIN_SAMPLES``）才退回规范注解，并在 note 里显式写
              「期望来自规范注解（非模板实测）」。
页边距        ``w:sectPr/w:pgMar`` 的 top/right/bottom/left，容差 20 twips。
              **逐节配对后比**（EZI-0016 逐节、EZI-0027 改按角色）：模板与产物各自读出
              全部节；节数相同 → 第 i 节比第 i 节；**节数不同 → 按角色配对**（封面/声明/
              摘要/ABSTRACT/目录/正文/参考文献/致谢/附录，词表与引擎 ``section_geometry``
              同源），同一个产物节只比它角色对应的模板节——本科论文产物多一个「声明」节，
              按索引比会把产物 ABSTRACT(1417) 和模板 目录(1440) 对成一对，把全对的边距判
              FAIL。角色认不出来（只认出 ≤1 个模板节 / 读不出逐节段落）→ 退回逐节索引比
              ``min(模板, 产物)`` 节，note 写明用了哪种口径，**不静默跳过**。
              **禁止**两侧各取一个单值——本科模板 8 节里
              节0~2 = 1417/1701/1134/1701、节3~7 = 1440/1800/1440/1800，旧判法拿模板
              节0 比产物末节，把「产物正文节用了封面边距」这个真差异说成了「单值不符」。
目录域        ``w:fldSimple/@w:instr`` 或 ``w:instrText`` 里出现 ``TOC`` 指令。
              **迁移语义**：期望值来自**源文档**（``source``）——源有目录域则产物
              必须有（缺 → FAIL）；源没有则产物**也不许生成**（EZI-0038）：产物 0 个
              → PASS，产物有 TOC 域 → FAIL。不给 ``source`` 一律 ``N-A``，绝不按模板
              或常量猜期望。EZI-0044：旧口径把「源无 TOC」整支判 N-A（对错都不看），
              三对体检样本上把这一项做成了 100% 死指标；它本来就有一条可判的迁移
              约束（不得凭空生成目录域），不是「不可判」。
页眉页脚      **只按 PAGE 域判（EZI-0012 A1-A4）**：期望值只看模板同名部件里
              有没有 ``PAGE`` 域（``w:instrText``/``w:fldSimple``），**不看**
              ``headerReference``/``footerReference`` 存在性。
              模板页眉含 PAGE → 产物须有页眉且含 PAGE，否则 ``FAIL``；模板页眉不含
              PAGE（公文版头、合同编号这类纯文字）→ ``N-A``，按 2026-11-10 页眉红线
              （不写模板页眉文字、不清用户占位文字，只继承 PAGE 域链 + 段落格式），
              产物不该有这个页眉。页码域：模板或 spec 有 PAGE → 产物须有；两者均无
              → ``N-A``（不再有「产物须有页码域」这条常量判据）。页眉字号（EZI-0044）：
              基准 = **模板页眉 PAGE 段落的
              有效字号**（PAGE run 直接 ``w:sz`` → 段落标记 ``w:pPr/w:rPr/w:sz`` →
              段落/默认样式链），与「页眉有没有文字」无关（PAGE 是域不是文字，占位符
              ``\xa0`` 也不影响）；模板页眉没有 PAGE 域时 ``N-A``——那个字号是模板页眉
              **文字**的字号，按红线不迁，没有可继承的基准。产物没有页眉部件、或页眉
              没有 PAGE 域时 ``N-A``：缺的是产物侧 PAGE 链证据，已由「页码域(PAGE)」判，
              不重复计。**「页眉引用(headerReference)」已废弃（EZI-0044）**：11-10 红线
              禁止按 headerReference 存在性判 FAIL，而「模板页眉有 PAGE → 产物页眉须有
              PAGE」这条唯一有意义的语义与「页码域(PAGE)」「页眉字号」完全重叠 → 不再
              输出该项；页脚引用照旧（页脚没有「不迁文字」的红线）。
              页脚页码字号（EZI-0064 同构页眉口径）：模板页脚有 PAGE 但
              PAGE 段落样式链无字号证据、或产物无页脚/无 PAGE 时 ``N-A``。
题注绑定      图题（``^图\\s*[0-9一二三四五六七八九十]``）必须紧跟其**前**的图片段落，
              且该图片段落带 ``w:keepNext``；表题（``^表\\s*[0-9…]``）必须紧跟其**后**的
              表格，且表题段落带 ``w:keepNext``。方向由**对象位置**定（对象在题前=图、
              对象在题后=表）：``w:keepNext`` 只能绑「自己 + 后面内容」，图题在下所以
              绑的是图片段落，表题在上所以绑的是题注自己。
              题与对象不相邻（隔空段/他段/题注内嵌图对象文本框）时 ``N-A`` 并在 note
              写明原因——**不猜**。段文本只算**流内** ``w:t``（图形/文本框里的字不算
              段落文字，否则「图里带的题注」会被当成题注段）。
分节分栏      ``w:sectPr`` 数量：**模板单节时判产物必须 1 节**（EZI-0034 A5 反例：
              单节模板一律不切，多切 = 发明结构）；模板多节时 ``N-A``，note 给
              源/模板/产物 三个节数 + 缺哪类证据——产物节数 = 产物**属性变化点数**
              （cols/pgMar/pgSz/页码体系），而变化点由**源内容角色**决定、不由模板
              节数决定（EZI-0034），模板多节时没有「内容块→模板节」映射证据，拿模板
              节数当期望就是常量后门（EZI-0044）。``w:cols/@w:num``
              **逐节**读：两侧都取 body 末节（``body/w:sectPr``，缺失则最后一个段落里的
              sectPr）。旧判法拿模板 ``structure.body_cols``（「第一个非首节 ≥2 栏」）去比
              产物末节——两侧不是同一节，ACM 模板/产物 cols 序列都是 ``[1,3,1,2]``/
              ``[3,2]``（非 1 栏部分）却被判 ``3 vs 2`` FAIL（EZI-0015）。
============  ==================================================================

期望值来源（不自己发明）：``earlyz.spec.compilers.compile_format_spec`` 产出的
``spec.structure`` / ``spec.rules``，按 ``semantic_role`` 归并；模板侧没有规则的项
（页眉字号、页脚页码字号）取模板 docx 同名部件实测值——**两项都取 PAGE 段落的
样式链有效字号**（EZI-0044 页眉先行、EZI-0064 页脚同构接线：旧「页脚直写众数」
口径在 PAGE run 无直写时拿页脚**文字**字号错当页码字号，且样式链继承时整块 N-A），并在
``note`` 里写明出处；连模板也没有的项一律 ``N-A`` 并写明原因，绝不拿默认值凑 PASS。

**例外（EZI-0016）：三项以模板 docx 的实测值当期望，不信 spec 的推值**——正文主字号、
标题字号/字体/对齐、页边距。理由是 spec 的这三类值来自编译期的簇取样/注解（``min_cluster``
可以是 1 段、注解可以写「正文四号」），与产物侧的取样口径对不上：本科模板正文实测
12pt×48/77 而 spec 写 14pt、新疆模板三级 17/17 = center 而 spec 写 both、本科模板
8 节两套边距而旧判法只读首节。三道护栏保证不把模板的噪声当版式事实：样本量 <
``TEMPLATE_MODE_MIN_SAMPLES``、主峰占比 < ``TEMPLATE_MODE_MIN_SHARE``、标题文本重合率 <
``HEADING_OVERLAP_MIN`` 时一律退回 spec 并在 ``note`` 写明原因。

**例外之例外（EZI-0022）：正文主字号无主峰时不许退回 spec**。模板自己混排（最高频
< 60%）时「退回 spec 注解」= 又拿常量当期望，产物/注解撞上就假 PASS、错开就假 FAIL。
该支改判双峰区间（前二高频，落在任一即 PASS；前二合计 < 60% → N-A）；退回规范注解
只剩「样本不足（<5 段）」一种情形，且 ``note`` 必须写明「期望来自规范注解（非模板
实测）」。标题三项走 EZI-0016/EZI-0021/EZI-0025 口径：同文成对全一致 → 取模板同文
实测（不受样本下限约束）；否则按模板该级众数（两道护栏）→ 退回 spec。

分数口径（可复算）：``score = passed / (total - not_applicable)``，N-A 不进分母；
分母为 0 时 ``score = 0.0``。退出码 ``0``=无 FAIL、``1``=有 FAIL、``2``=跑不起来。
"""
from __future__ import annotations

import re
import zipfile
from collections import Counter
from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Sequence

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

PASS = "PASS"
FAIL = "FAIL"
NA = "N-A"

#: 页边距容差（twips）。沿用 2026-10-17 版判定，不放宽。
MARGIN_TOLERANCE_TWIPS = 20
#: 字号容差（pt）。沿用 2026-10-17 版判定，不放宽。
SIZE_TOLERANCE_PT = 0.6
#: 页眉/页脚字号容差（pt）。两侧都直接读 ``w:sz``（无单位换算），按近精确比。
HF_SIZE_EPS_PT = 0.1
#: 首行缩进容差（字符）。
INDENT_TOLERANCE_CHARS = 1.0
#: 模板实测众数当期望值的最小样本数（EZI-0016）。模板该级只有 3 个样本时，
#: 「3/3 一致」不代表版式——东台一级标题模板只有 3 个「第X章」，产物 34 个一级标题里
#: 24 个是对齐 both 的条款项；拿 3 个样本的众数去判 34 个段落，比的是两拨不相干的段落。
#: 例外（EZI-0025）：「与产物同文成对」的模板段落在该属性上**全一致**时不受此下限约束
#: ——那是 1:1 对应（「产物这一段该不该像模板那一段」），不是从 3 个样本推该级版式。
TEMPLATE_MODE_MIN_SAMPLES = 5
#: 「同文成对」当直接证据的最小段数（EZI-0025）。只有 1 段对得上说明该级只在个别段上
#: 有对应（或文本撞车），仍按 ``TEMPLATE_MODE_MIN_SAMPLES`` 的统计口径走，不搞单段特例。
HEADING_SAME_TEXT_MIN = 2
#: 模板实测众数当期望值的主峰占比下限（EZI-0016）。东台二级标题模板 16 个样本里
#: 仿宋 8 个（50%）= 模板自己就没统一，不能拿它当期望值判产物 FAIL。
TEMPLATE_MODE_MIN_SHARE = 0.6
#: 无主峰时「双峰区间」的候选个数（EZI-0022）。模板正文实测最高频 <
#: ``TEMPLATE_MODE_MIN_SHARE`` = 模板自己就是混排，此时**不许**退回 spec 注解值当期望
#: （那是常量后门），而是取前二高频当候选：产物落在任一候选即 PASS。前二合计仍须
#: ≥ ``TEMPLATE_MODE_MIN_SHARE``，否则分布过散（前二也代表不了模板）→ N-A 不判。
TEMPLATE_NO_PEAK_TOP_K = 2
#: 标题样本「文本重合率」下限（EZI-0016）：产物该级样本里至少这个比例要在模板同级
#: 样本里找到同文本对应，才算两侧比的是同一批标题（新疆模板二级是「第X章」，
#: 产物二级是条款项 → 重合 0/13，此时模板实测不能当期望）。
HEADING_OVERLAP_MIN = 0.5
#: 页边距比较的四个方向。
_MARGIN_KEYS = ("top", "right", "bottom", "left")

# ── 节角色词表（EZI-0027）────────────────────────────────────────────────
# 与引擎 ``earlyz/actions/section_geometry.py`` 的 ``_ROLE_HEADINGS`` **同源**；
# 判定层独立（只读 OOXML，不 import 会改文档的引擎模块），所以词表以常量复制。
# 两处词表如需增删角色，必须同步修改，否则「引擎按角色套几何」与「判定按角色比几何」
# 会各认一套角色，产物节数不等时又会错位。
_SECTION_ROLE_HEADINGS: dict[str, str] = {
    "声明": "declaration",
    "郑重声明": "declaration",
    "原创性声明": "declaration",
    "学术声明": "declaration",
    "学位论文原创性声明": "declaration",
    "摘要": "abstract",
    "中文摘要": "abstract",
    "abstract": "abstract_en",
    "英文摘要": "abstract_en",
    "目录": "toc",
    "参考文献": "references",
    "references": "references",
    "参考书目": "references",
    "致谢": "acknowledgements",
    "acknowledgements": "acknowledgements",
    "acknowledgments": "acknowledgements",
}
#: 附件表单区标题（EZI-0075）：「附件2」「附件 3 ：」——独立短行、不含制表符
#: （目录条目「附件2⇥45」不算）。与引擎 ``section_geometry._ATTACHMENT_RE``
#: 同口径；判定层缺这一档 → 产物附件区配对不到模板表单节（1680 真值）。
_SECTION_ATTACHMENT_RE = re.compile(r"^附\s*件\s*[0-9]{0,2}[：:]?$")
#: 目录页节长上限（EZI-0075）：超长节里的目录行是噪声不是节身份（同引擎
#: ``_TOC_PAGE_MAX_PARAS``；建设施工被吞节实测 1628 段）。
_TOC_SECTION_MAX_PARAS = 80
#: 角色 → note 里的人话名（配对表用；顺序即学位论文的前后顺序）。
_SECTION_ROLE_LABEL = {
    "cover": "封面",
    "declaration": "声明",
    "abstract": "摘要",
    "abstract_en": "ABSTRACT",
    "toc": "目录",
    "body": "正文",
    "references": "参考文献",
    "acknowledgements": "致谢",
    "appendix": "附录",
    "attachment": "附件",
}
#: 标题归一：去掉空白、全角空格、模板占位符（□）和省略号/点线（同引擎口径）。
_SECTION_PLACEHOLDER_RE = re.compile(r"[\s\u3000\u00a0□■…\u2026]+")
#: 「第 N 章」与「1 / 1.1 / 1．4．1」式章节标题（同引擎口径）。
_SECTION_CHAPTER_RE = re.compile(r"^第[0-9一二三四五六七八九十百]+章")
_SECTION_NUMBERED_RE = re.compile(r"^\d+(?:[.．]\d+){0,3}[、.．]?\s*\S")

#: 标题规则的键 → ``effective_props`` 的取样属性键（字号/字体两套命名不同）。
_HEAD_PROP_FIELD = {"font_size_pt": "sz", "font_east_asia": "east_asia", "jc": "jc"}
#: 上面三个键的中文人话名（note 里说「无可测字号/中文字体/对齐」）。
_HEAD_PROP_LABEL = {"font_size_pt": "字号", "font_east_asia": "中文字体", "jc": "对齐"}
#: 题注样式（正文主字号的实测样本要排掉题注段，EZI-0016）。
_CAPTION_STYLE = re.compile(r"^(Caption|题注|图注|表注|图表注|Figure|Table)", re.IGNORECASE)

CATEGORY_TITLES = "标题"
CATEGORY_BODY = "正文"
CATEGORY_MARGIN = "页边距"
CATEGORY_TOC = "目录域"
CATEGORY_HF = "页眉页脚"
CATEGORY_CAPTION = "题注绑定"
CATEGORY_SECTION = "分节分栏"

_TOC_STYLE = re.compile(r"^TOC", re.IGNORECASE)
_TOC_INSTR = re.compile(r"\bTOC\b")
_PAGE_INSTR = re.compile(r"\bPAGE\b")  # 不会命中 PAGEREF / NUMPAGES
_HEADER_PART = re.compile(r"^word/header\d*\.xml$")
_FOOTER_PART = re.compile(r"^word/footer\d*\.xml$")

#: 页眉红线的原文（2026-11-10）：页眉只继承 PAGE 域链 + 段落格式，不迁模板页眉文字。
_HEADER_REDLINE = ("2026-11-10 页眉红线：不写模板页眉文字、不清用户占位文字，"
                   "只继承 PAGE 域链 + 段落格式")

#: 编号标题的文本形态（EZI-0012 A7）：标题取样只认这些前缀。
#: 前三式是 issue 给的原式；``1.``/``（一）``/``（1）`` 是公文常见编号（同济校级发文的
#: 一级「1.正文」、二级「（一）」都靠它们），不补就会把真标题整批排除成 N-A。
#: ``\s`` 覆盖半角空格、全角空格（\\u3000）与 em space（\\u2003，ACM 用）。
_NUMBERED_HEADING_RES = (
    re.compile(r"^第\s*[一二三四五六七八九十百千0-9]+\s*[章节编]"),
    re.compile(r"^[0-9]+(?:\s*[.．]\s*[0-9]+)+\s*"),   # 1.1 / 1.1.2 / 1．1
    re.compile(r"^[0-9]+\s*[.．、]\s*"),                # 1 绪论（公文「1.正文」）
    re.compile(r"^[0-9]+\s+"),                          # 1 绪论（空格分隔）
    re.compile(r"^[一二三四五六七八九十百]+\s*[、.．]"),  # 一、
    re.compile(r"^[（(]\s*[一二三四五六七八九十百0-9]+\s*[）)]"),  # （一）/（1）
)
#: 后置/固定节名：无编号行，不该当编号标题样本（`APPENDIX A` 也在此列）。
_SECTION_NAME_RE = re.compile(
    r"^\s*(?:abstract|references?|bibliography|acknowledge?ments?|appendix(?:\s+[a-z])?"
    r"|index|contents|目\s*录|摘\s*要|参考文献|致\s*谢|附\s*录(?:\s*[a-z一二三四五六七八九十])?)"
    r"\s*[:：]?\s*$", re.IGNORECASE)
#: 表单标签：以冒号结尾的短行（`学院 （系）:`）——封面/表格里的字段名，不是标题。
_FORM_LABEL_RE = re.compile(r"[:：]\s*$")
_FORM_LABEL_MAX = 24

#: 手写目录条目（EZI-0041）：制表符分段、末段是页码——``1.1.\t编写目的\t4``。
#: 真 TOC 域的条目带 ``TOC1``/``TOC2`` 样式或 ``PAGEREF``，本来就进不了样本；这条规则
#: 管的是源文档自带、**没有任何域标记**的手写目录：``EZI-0038`` 不再生成目录后，它们
#: 既没有 TOC 样式也没有 PAGEREF，于是落进正文/标题样本池（可行研 7 段 16pt 条目把
#: 「正文主字号」的众数从 10.5 拖到 16.0，功能设计/安装手册的二/三级标题对齐因此变 FAIL）。
# EZI-0073：两段（标题\t…\t页码）与单段（标题\t页码）都算——重庆科技实锤
# 「1 正文格式说明\t1」单 tab 形态漏判混进 L1 标题样本，jc 众数被带偏；
# 真标题不以 tab+纯数字结尾，扩形态零误伤。
_MANUAL_TOC_ENTRY_RE = re.compile(r"^.+?\t+(?:.+?\t+)?\s*\d{1,4}\s*$")
#: 条目里标题前的编号（``1.1.`` / ``一、`` / ``（一）`` / ``第3章``）：比对正文标题时去掉。
_MANUAL_TOC_NUM_RE = re.compile(
    r"^\s*(?:第[0-9一二三四五六七八九十百千]+\s*[章节编]"
    r"|[0-9]+(?:\s*[.．]\s*[0-9]+)*\s*[.．、]?"
    r"|[一二三四五六七八九十百]+\s*[、.．]"
    r"|[（(]\s*[一二三四五六七八九十百0-9]+\s*[）)])\s*")

#: 题注：图 / 表 + 编号（阿拉伯或中文数字）。图题在图下方、表题在表上方。
_CAPTION_RE = re.compile(r"^(图|表)\s*[0-9一二三四五六七八九十]")
#: 图形的三种容器：DrawingML 图片、VML 图、嵌入对象（OLE/图表）。
_DRAWING_TAGS = ("drawing", "pict", "object")
#: 段内文本要排除的图形容器（图对象自带的文本框文字不算「段落文字」）。
_GRAPHIC_SCOPE = ("drawing", "pict", "object", "txbxContent")
#: 流内包装层：穿过它们取块级元素（内容控件里的段落仍然是流内段落）。
_FLOW_WRAPPERS = frozenset({"sdt", "sdtContent", "customXml", "smartTag",
                            "ins", "del", "moveTo", "moveFrom"})

_JC_ALIASES = {"left": "start", "right": "end", "start": "start", "end": "end"}


# ────────────────────────────── 报告模型 ──────────────────────────────


@dataclass(frozen=True)
class LayoutItem:
    """一条判定。``expected`` / ``actual`` 是可直接 json 序列化的原始值。"""

    name: str
    expected: Any
    actual: Any
    verdict: str
    category: str = ""
    note: str = ""

    def as_dict(self) -> dict[str, Any]:
        d = {
            "name": self.name,
            "expected": self.expected,
            "actual": self.actual,
            "verdict": self.verdict,
            "category": self.category,
        }
        if self.note:
            d["note"] = self.note
        return d


@dataclass
class LayoutReport:
    """一次判定的全部结果。``score`` 只由 verdict 决定，随时可复算。"""

    product: str
    template: str
    items: list[LayoutItem] = field(default_factory=list)
    #: 源文档（迁移语义的期望值来源）；没传就是 ``None``。不参与计分。
    source: str | None = None

    @property
    def passed(self) -> int:
        return sum(1 for i in self.items if i.verdict == PASS)

    @property
    def failed(self) -> int:
        return sum(1 for i in self.items if i.verdict == FAIL)

    @property
    def not_applicable(self) -> int:
        return sum(1 for i in self.items if i.verdict == NA)

    @property
    def total(self) -> int:
        return len(self.items)

    @property
    def judged(self) -> int:
        return self.total - self.not_applicable

    @property
    def score(self) -> float:
        return round(self.passed / self.judged, 4) if self.judged else 0.0

    @property
    def exit_code(self) -> int:
        return 1 if self.failed else 0

    def as_dict(self) -> dict[str, Any]:
        return {
            "product": self.product,
            "template": self.template,
            "source": self.source,
            "passed": self.passed,
            "failed": self.failed,
            "not_applicable": self.not_applicable,
            "total": self.total,
            "judged": self.judged,
            "score": self.score,
            "exit_code": self.exit_code,
            "items": [i.as_dict() for i in self.items],
        }


def score_items(items: Sequence[LayoutItem]) -> dict[str, Any]:
    """从原始判定行复算分数（不依赖 LayoutReport，便于外部对账）。"""
    rep = LayoutReport(product="", template="", items=list(items))
    return {
        "passed": rep.passed,
        "failed": rep.failed,
        "not_applicable": rep.not_applicable,
        "total": rep.total,
        "judged": rep.judged,
        "score": rep.score,
    }


def format_table(report: LayoutReport) -> str:
    """人类可读表格（老格式的超集：多了一列备注）。"""
    lines = [f"产物: {Path(report.product).name}", f"模板: {Path(report.template).name}"]
    if report.source:
        lines.append(f"源文档: {Path(report.source).name}")
    lines.append("")
    if not report.items:
        lines.append("  （没有可机器核对的项——模板是散文式规范）")
        return "\n".join(lines)
    width = max(len(i.name) for i in report.items)
    for it in report.items:
        lines.append(f"  [{it.verdict:<4}] {it.name:<{width}}  模板={_show(it.expected)}  产物={_show(it.actual)}")
        if it.note:
            lines.append(f"         └ {it.note}")
    lines.append(
        f"\n{report.passed}/{report.judged} 通过"
        f"（共 {report.total} 项，N-A {report.not_applicable} 项，"
        f"FAIL {report.failed} 项）  正确率={report.score:.2%}"
    )
    return "\n".join(lines)


def _show(v: Any) -> str:
    if v is None:
        return "None"
    if isinstance(v, dict):
        return "{" + ", ".join(f"{k}={v[k]}" for k in v) + "}"
    if isinstance(v, (tuple, list)):        # 双峰区间（EZI-0022）：16.0 或 14.0
        return " 或 ".join(_show(x) for x in v)
    return str(v)


# ────────────────────────────── OOXML 读取 ──────────────────────────────


def _load_docx(path: str | Path) -> dict[str, Any]:
    """把一个 docx 读进内存：document / styles / header 部件 / footer 部件。"""
    p = Path(path)
    if not p.is_file():
        raise FileNotFoundError(f"文件不存在：{p}")
    with zipfile.ZipFile(p) as z:
        names = set(z.namelist())
        if "word/document.xml" not in names:
            raise ValueError(f"不是合法的 docx（缺 word/document.xml）：{p}")
        doc = etree.fromstring(z.read("word/document.xml"))
        styles = etree.fromstring(z.read("word/styles.xml")) if "word/styles.xml" in names else None
        hf_parts: dict[str, Any] = {}
        for n in sorted(names):
            if _HEADER_PART.match(n) or _FOOTER_PART.match(n):
                hf_parts[n] = etree.fromstring(z.read(n))
    return {"path": p, "document": doc, "styles": styles, "hf_parts": hf_parts}


class _Styles:
    """样式链解析：docDefaults → pStyle→basedOn 逐级 → 段落/run 直接格式。"""

    def __init__(self, root: Any) -> None:
        self.by_id: dict[str, Any] = {}
        if root is not None:
            for s in root.iter(W + "style"):
                sid = s.get(W + "styleId")
                if sid:
                    self.by_id[sid] = s
        self.rpr_default = None
        self.ppr_default = None
        if root is not None:
            dd = root.find(W + "docDefaults")
            if dd is not None:
                rd = dd.find(W + "rPrDefault")
                if rd is not None:
                    self.rpr_default = rd.find(W + "rPr")
                pd = dd.find(W + "pPrDefault")
                if pd is not None:
                    self.ppr_default = pd.find(W + "pPr")
        self._chains: dict[str, list[Any]] = {}

    def chain(self, style_id: str | None) -> list[Any]:
        """从最基类到最派生（weakest → strongest）。带环保护。"""
        if not style_id:
            return []
        if style_id in self._chains:
            return self._chains[style_id]
        out: list[Any] = []
        seen: set[str] = set()
        sid: str | None = style_id
        while sid and sid in self.by_id and sid not in seen:
            seen.add(sid)
            st = self.by_id[sid]
            out.append(st)
            b = st.find(W + "basedOn")
            sid = b.get(W + "val") if b is not None else None
        out.reverse()
        self._chains[style_id] = out
        return out


def _pstyle(ppr: Any) -> str | None:
    if ppr is None:
        return None
    st = ppr.find(W + "pStyle")
    return st.get(W + "val") if st is not None else None


def _rpr_chain(p: Any, styles: _Styles, run: Any = None) -> list[Any]:
    out: list[Any] = []
    if styles.rpr_default is not None:
        out.append(styles.rpr_default)
    ppr = p.find(W + "pPr")
    for st in styles.chain(_pstyle(ppr)):
        r = st.find(W + "rPr")
        if r is not None:
            out.append(r)
    if ppr is not None:
        r = ppr.find(W + "rPr")
        if r is not None:
            out.append(r)
    if run is not None:
        r = run.find(W + "rPr")
        if r is not None:
            out.append(r)
    return out


def _ppr_chain(p: Any, styles: _Styles) -> list[Any]:
    out: list[Any] = []
    if styles.ppr_default is not None:
        out.append(styles.ppr_default)
    ppr = p.find(W + "pPr")
    for st in styles.chain(_pstyle(ppr)):
        x = st.find(W + "pPr")
        if x is not None:
            out.append(x)
    if ppr is not None:
        out.append(ppr)
    return out


def _val(chain: Iterable[Any], *tags: str, attr: str = "val") -> str | None:
    """按 tags 路径取属性（默认 ``@w:val``，rFonts 用 ``@w:eastAsia`` 等）。

    链上最后命中者胜（更派生/更直接的覆盖更弱的）。
    """
    v: str | None = None
    for el in chain:
        cur = el
        hit = True
        for t in tags:
            cur = cur.find(W + t)
            if cur is None:
                hit = False
                break
        if not hit:
            continue
        got = cur.get(W + attr)
        if got is not None:
            v = got
    return v


def _int(v: Any) -> int | None:
    try:
        return int(str(v).strip())
    except (TypeError, ValueError):
        return None


def _first_text_run(p: Any) -> Any:
    for r in p.iter(W + "r"):
        for t in r.findall(W + "t"):
            if (t.text or "").strip():
                return r
    return None


def _has_text(p: Any) -> bool:
    return any((t.text or "").strip() for t in p.iter(W + "t"))


def effective_props(p: Any, styles: _Styles) -> dict[str, Any]:
    """段落的「有效」格式：样式链 + 直接格式合并后的实测值。"""
    run = _first_text_run(p)
    rp = _rpr_chain(p, styles, run)
    pp = _ppr_chain(p, styles)
    sz_half = _int(_val(rp, "sz"))
    ppr = p.find(W + "pPr")
    spacing = None
    ind = None
    for el in pp:
        s = el.find(W + "spacing")
        if s is not None:
            spacing = s
        i = el.find(W + "ind")
        if i is not None:
            ind = i
    line = _int(spacing.get(W + "line")) if spacing is not None else None
    rule = spacing.get(W + "lineRule") if spacing is not None else None
    first_chars = _int(ind.get(W + "firstLineChars")) if ind is not None else None
    first_line = _int(ind.get(W + "firstLine")) if ind is not None else None
    return {
        "style_id": _pstyle(ppr),
        "sz": (sz_half / 2.0) if sz_half else None,
        "east_asia": _val(rp, "rFonts", attr="eastAsia"),
        "ascii": _val(rp, "rFonts", attr="ascii"),
        "jc": _val(pp, "jc"),
        "outline_lvl": _val(pp, "outlineLvl"),
        "first_line_chars": first_chars,
        "first_line": first_line,
        # 行距缺 lineRule 时 Word 按 auto 处理
        "line": line,
        "line_rule": (rule or "auto") if line is not None else None,
    }


def _inside_table(p: Any) -> bool:
    return any(a.tag == W + "tc" for a in p.iterancestors())


#: 封面边界缓存（EZI-0061）：path → find_cover_end(analyze_document(path))。
#: 每 doc 一次 IR 分析（~1s），evaluate 内 document_paragraphs 被调 3-4 次、
#: 基线跑批 73 对 ×2 侧，不缓存会重复分析同一文档数十次。
_COVER_END_CACHE: dict[str, int] = {}


def _cover_end_of(doc: dict[str, Any]) -> int:
    """本文档的封面区边界（lint 侧 `find_cover_end` 口径，两侧同源，禁第二套边界）。

    IR 解析失败退 0（不排封面）——判定宁可不排也不因 IR 异常整份崩。
    """
    p = doc.get("path")
    if not p:
        return 0
    key = str(p)
    if key not in _COVER_END_CACHE:
        try:
            from earlyz.ir.analyzer import analyze_document  # noqa: PLC0415 — 防环
            from earlyz.spec.decompiler import find_cover_end  # noqa: PLC0415

            _COVER_END_CACHE[key] = find_cover_end(analyze_document(key))
        except Exception:  # noqa: BLE001 — 封面排除是增强不是前提
            _COVER_END_CACHE[key] = 0
    return _COVER_END_CACHE[key]


def document_paragraphs(doc: dict[str, Any]) -> list[tuple[Any, dict[str, Any]]]:
    """正文（非表格内、**非文本框内、非封面区**）全部段落及其有效属性，一次算好复用。

    EZI-0061（2026-11-20）修复语义——EZI-0060 A1 全量实证（73 对：textbox 入池
    22 对 / 封面区入池 63 对 / 三键众数翻转 13 对，已证误判 3+）：

    * **文本框段排除**：``body.iter(w:p)`` 会下潜 ``w:txbxContent``——模板格式
      说明文（书脊说明/题目排版示范）混进判定正文池，英语笔译模板实测 465 段中
      306 段是文本框说明文，主字号期望 10.5 由污染池产出而模板真实正文 12.0；
      祖先含 ``w:txbxContent`` 的段不进池（与 spec 侧 ``analyze_document`` 口径
      对齐——IR 本就不收文本框段）；
    * **封面区段排除**：封面区（lint 侧 ``find_cover_end``，两侧同源）内的说明行
      （含「-435610-2345690存档论文…」锚号残行）不进池。非文本框段按文档序对齐
      IR 索引（analyze_document 不收文本框段，两侧索引同构）。

    A1 裁定（2026-11-20，13 翻转对离线重判全表 diff）：零翻红；6 个隐藏假 FAIL
    翻绿（英语笔译 3-2-1 主字号、矿大 1-1-2/2-1-2/上海交大 行距、立项申请表
    对齐、湖南师大主字号）+ N-A→PASS 样本面扩大 6 例。3-2-2 行距 FAIL 存活
    （说明书型模板——模板无真正文示范，属模板形态缺陷，排除法范围外，见
    EZI-0059/0061 账本）。
    """
    styles = _Styles(doc["styles"])
    body = doc["document"].find(W + "body")
    if body is None:
        return []
    cover_end = _cover_end_of(doc)
    out = []
    ir_idx = -1
    for p in body.iter(W + "p"):
        if _inside_table(p):
            continue
        if any(a.tag == f"{W}txbxContent" for a in p.iterancestors()):
            continue        # 文本框段（格式说明文）不是正文
        ir_idx += 1
        if ir_idx < cover_end:
            continue        # 封面区说明行不是正文
        out.append((p, effective_props(p, styles)))
    return out


def _is_form_label(text: str) -> bool:
    """表单标签（``学生姓名:`` / ``学院 （系）:``）：以冒号结尾的短行。

    封面/表格里的字段名不是正文段落——正文取样与标题取样共用这一条判定
    （``EZI-0012`` 的标题排除、``EZI-0020`` 的正文取样），不许各写一套。
    """
    return len(text) <= _FORM_LABEL_MAX and bool(_FORM_LABEL_RE.search(text))


def _toc_field_spans(body: Any) -> set[int]:
    """落在**复杂域**（``w:fldChar`` begin…end）TOC 区间的段落，返回 ``id(p)`` 集合。

    只用来回答「这段是不是真 TOC 域里的条目」——域内条目本来就带 ``TOC1``/``TOC2``
    样式（``body_paragraphs`` 与 ``heading_samples_by_level`` 已按样式排除），这里再
    守一道，避免把域内段落误判成「手写目录条目」而写错 note。
    """
    inside: set[int] = set()
    stack: list[dict[str, Any]] = []
    open_toc = 0
    cur: Any = None
    for el in body.iter():
        tag = el.tag
        if tag == W + "p":
            cur = el
            if open_toc:
                inside.add(id(cur))
            continue
        if tag == W + "fldChar":
            kind = el.get(W + "fldCharType")
            if kind == "begin":
                stack.append({"instr": [], "toc": False})
            elif kind == "end" and stack:
                if stack.pop()["toc"]:
                    open_toc -= 1
            continue
        if tag == W + "instrText" and stack:
            top = stack[-1]
            top["instr"].append(el.text or "")
            if not top["toc"] and _TOC_INSTR.search("".join(top["instr"])):
                top["toc"] = True
                open_toc += 1
                if cur is not None:
                    inside.add(id(cur))       # 指令所在段也是域内段
    return inside


def _in_toc_field(p: Any) -> bool:
    """段落是否在真目录域里：``w:fldSimple`` 祖先/自身域、同段 ``w:instrText``、跨段域区间。"""
    for anc in p.iterancestors():
        if anc.tag == W + "fldSimple" and _TOC_INSTR.search(anc.get(W + "instr") or ""):
            return True
    if any(_TOC_INSTR.search(fs.get(W + "instr") or "") for fs in p.iter(W + "fldSimple")):
        return True
    if any(_TOC_INSTR.search(t.text or "") for t in p.iter(W + "instrText")):
        return True
    body = next((a for a in p.iterancestors() if a.tag == W + "body"), None)
    return body is not None and id(p) in _toc_field_spans(body)


def _manual_toc_title(text: str) -> str | None:
    """手写目录条目的标题（去编号、去页码）：``1.1.\t编写目的\t4`` → ``编写目的``。"""
    parts = [s.strip() for s in text.split("\t")]
    if len(parts) < 2 or not parts[-1].isdigit():
        return None
    named = [s for s in parts[:-1] if s]
    if not named:
        return None
    title = _MANUAL_TOC_NUM_RE.sub("", named[-1]).strip()
    return title or None


def _is_manual_toc_entry(p: Any, text: str | None = None) -> bool:
    """手写目录条目：**文本以制表符分段且末段是页码**（``1.1.\\t编写目的\\t4``），
    且该段**不在真 TOC 域内**。

    正文取样（``body_paragraphs`` / ``body_text_paragraphs``）与标题取样
    （``heading_samples_by_level``）共用这一条，不许各写一套（EZI-0041）。
    """
    text = flow_text(p) if text is None else text
    if not text or not _MANUAL_TOC_ENTRY_RE.match(text):
        return False
    return not _in_toc_field(p)


def _manual_toc_titles(paras: Sequence[tuple[Any, dict[str, Any]]]) -> set[str]:
    """手写目录列出的标题文本集合（EZI-0041）。

    这类源文档的手写目录就是它们的**标题索引**：正文流里与条目同名的段落是标题，只是
    没带 ``w:outlineLvl``（``EZI-0038`` 之后引擎不再给它们标级别），于是被当成正文段
    参与字号/行距众数投票——可行研 26 段 16pt 标题把「正文主字号」的众数顶到 16.0。
    没有手写目录的文档返回空集，取样口径不变。
    """
    out: set[str] = set()
    for p, _props in paras:
        text = flow_text(p)
        if not _is_manual_toc_entry(p, text):
            continue
        title = _manual_toc_title(text)
        if title:
            out.add(title)
    return out


def body_paragraphs(paras: Sequence[tuple[Any, dict[str, Any]]], *,
                    exclude_form_labels: bool = True) -> list[dict[str, Any]]:
    """正文段落：有文字、不是标题层级、不是目录条目、不是表单标签。

    ``exclude_form_labels``（EZI-0020）默认开：交大中期的表单字段「学生姓名:」等
    6 段是 ``500/exact``，与 6 段真正文（``400/atLeast``）**打平**，``most_common``
    按文档序取到前者，行距项把已经排对（``400/atLeast``）的产物判成 FAIL——字段名
    不该进正文众数投票。

    正文四项（字号/字体/缩进/对齐）暂时传 ``False`` 走旧样本：主字号的主峰阈值（60%）
    是按含标签的旧样本校准的，去掉标签会让本科论文 12pt（39/68=57%）从单峰掉进
    无主峰分支——`EZI-0022` 改判双峰区间后至少不再退回 spec 14pt，但候选集合仍会从
    「12pt」变成「12pt/10.5pt」，判据强度变了。换取样口径属另一件事（见
    ``docs/issues/EZI-0020`` 决策记录）。

    ``EZI-0041``：再排掉**手写目录条目**（``1.1.\\t编写目的\\t4``）与**手写目录列出的
    标题**——这类源文档的手写目录就是标题索引，正文流里同名的段落是标题（只是没带
    ``outlineLvl``），不该进正文众数投票。
    """
    sel = []
    toc_titles = _manual_toc_titles(paras)
    from earlyz.ir.analyzer import is_format_instruction_text, _back_heading_role
    # EZI-0071：refs 区界状态机**先于一切跳过**看全文——References 标题自身
    # 是 Heading1（会被标题跳过跳掉），区界机看不到它就永不置位（3-2-1 实锤
    # 66 条 refs 全漏进池）。预判逐段 zone，按索引字典对齐（跳过链不消费）。
    in_refs = False
    _zone = {}
    for _zi, (_zp, _xp) in enumerate(paras):
        _text_raw = flow_text(_zp)
        _core_raw = re.sub(r"[\s　]+", "", _text_raw or "")
        _role_raw = _back_heading_role(_text_raw or "")
        if _role_raw == "refs_heading" or _core_raw in ("References", "REFERENCES"):
            in_refs = True
        elif _role_raw in ("bio_heading", "ack_heading", "appendix_heading"):
            in_refs = False
        _zone[_zi] = in_refs
    for _pi, (p, props) in enumerate(paras):
        if props["outline_lvl"] is not None:
            continue
        if _TOC_STYLE.match(props["style_id"] or ""):
            continue
        if not _has_text(p):
            continue
        text = flow_text(p)
        if _is_manual_toc_entry(p, text):
            continue
        if toc_titles and text in toc_titles:
            continue
        if exclude_form_labels and _is_form_label(text):
            continue
        # EZI-0071：格式说明段不进正文池——「字体：黑体；字号：三号；间距…」
        # 冒号/分号规格串（3-2-1 源 18 条实锤混进产物正文池，把缩进/字号众数
        # 带偏）。与引擎 strip_instruction_paragraphs、analyzer 内容 hash 豁免
        # **同一谓词**（earlyz.ir.analyzer.is_format_instruction_text，同源铁律）。
        if is_format_instruction_text(text):
            continue
        # EZI-0071：参考文献区不进正文池——refs 条目（[N] Author…，悬挂缩进、
        # 无 firstLineChars、TNR 字体）不是正文，量大时（英文论文 66 条实测）
        # 把正文缩进/字号众数带偏。区界状态机在预判段已对每段判好（refs 标题
        # 自身是 Heading1，必须先于标题跳过看到），此处按索引消费。
        if _zone.get(_pi):
            continue
        sel.append(props)
    return sel


def _is_cover_caps_line(text: str) -> bool:
    """封面拉丁大写行（``SHANGHAI JIAO TONG UNIVERSITY``）：全大写、无数字、非中文。

    带数字的不算——``1 INTRODUCTION`` 是编号标题，不该因为全大写被丢掉。
    """
    if not text.isascii() or any(c.isdigit() for c in text):
        return False
    letters = [c for c in text if c.isalpha()]
    return len(letters) >= 4 and all(c.isupper() for c in letters)


def _is_star_inline_note(text: str) -> bool:
    """星号内联格式说明行（EZI-0076，与引擎 ``new_actions._star_inline_format_note``
    同口径的判定层独立副本——判定层不 import 引擎模块，词表以常量复制）。

    行内含 ``*…*`` 记号且夹着格式术语（字体/字号/黑体/宋体/楷体/仿宋/加粗/
    居中/缩进/行距/标题 N），段长 8~120——真标题不带星号格式指令；目录条目
    已被前置的手写目录闸排除，不会误伤。
    """
    if not (8 <= len(text) <= 120) or "*" not in text:
        return False
    if not re.search(r"\*[^*]{2,40}\*", text):
        return False
    return bool(re.search(
        r"字体|字号|黑体|宋体|楷体|仿宋|加粗|居中|缩进|行距|标题[123一二三]|"
        r"Times|Roman|小[一二三四五]、|[一二三四]号", text))


def _heading_exclusion(text: str) -> str | None:
    """非标题段落的原因（也用于统计排除条数）；``None`` = 可当编号标题样本。

    顺序：后置节名 → 表单标签 → 封面拉丁大写行 → **编号前缀**。前三条是 issue
    点名的噪声形态，最后一条是正向要求（EZI-0012 A7：只认编号标题）。
    """
    if _SECTION_NAME_RE.match(text):
        return "后置节名"
    if _is_form_label(text):
        return "表单标签"
    if _is_cover_caps_line(text):
        return "封面大写行"
    if not any(r.match(text) for r in _NUMBERED_HEADING_RES):
        return "非编号行"
    return None


def heading_samples_by_level(
        doc: dict[str, Any]) -> tuple[dict[int, list[dict[str, Any]]], Counter]:
    """按 EZI-0012 A7 口径取一/二/三级标题样本，并统计各类排除条数。

    样本 = ``w:outlineLvl`` 级别匹配 **且** 文本是编号标题；再排除表格内段落、
    后置节名（ABSTRACT/REFERENCES/APPENDIX…）、表单标签（冒号结尾短行）、封面
    拉丁大写行、**手写目录条目**（EZI-0041：制表符分段 + 末段页码，且不在真 TOC 域内
    ——``1.1.\\t编写目的\\t4`` 是目录条目，不是二级标题）。**不再**拿「全部 outlineLvl=0 段落的众数」当事实：交大中期 3 段
    里只有 1 段是真标题（另两段是校名行 18pt 与表单字段 10.5pt），ACM 的众数被
    ``ABSTRACT``/``REFERENCES``（11pt）带偏——众数是噪声投票，不是版式事实。

    与 ``EZI-0009`` 的差别：那里适用性计数**包含**表格内段落（怕把 IEEE 的表格内
    真标题判成 N-A）；本 issue 的 A8 要求 IEEE 这 4 组标题项 N-A，改为源/产物同一套
    排除口径——宁可 N-A 也不拿表格、节名、表单标签当标题。
    """
    styles = _Styles(doc["styles"])
    body = doc["document"].find(W + "body")
    out: dict[int, list[dict[str, Any]]] = {1: [], 2: [], 3: []}
    excluded: Counter = Counter()
    if body is None:
        return out, excluded
    for p in body.iter(W + "p"):
        props = effective_props(p, styles)
        lvl = _int(props["outline_lvl"])
        if lvl is None or lvl > 2:
            continue
        text = flow_text(p)
        if not text:
            continue                       # 空段不是样本，也不算「排除」（没得判）
        if _inside_table(p):
            excluded["表格内"] += 1
            continue
        if _is_manual_toc_entry(p, text):
            excluded["手写目录条目"] += 1
            continue
        # EZI-0076：星号内联格式说明行（「1.1 二级标题 *字体为5号黑体*标题2」——
        # 模板作者把说明样例行挂真 outlineLvl，实为排版指令非内容标题）。引擎
        # 侧 0073 已剥除产物中的这类行；源/产物同一套排除口径（EZI-0012 A7
        # 纪律），源的池也必须排除——否则「源 1 个样本=说明行」会走迁移丢失
        # 分支悬起「待产品口径定」旗（计算机学报实锤，段 25 自证）。
        if _is_star_inline_note(text):
            excluded["星号内联说明行"] += 1
            continue
        why = _heading_exclusion(text)
        if why is not None:
            excluded[why] += 1
            continue
        out[lvl + 1].append({**props, "text": text[:40]})
    return out, excluded


def heading_role_counts(doc: dict[str, Any]) -> dict[int, int]:
    """源文档一/二/三级**编号标题样本**数（适用性判据，EZI-0012 A7 口径）。"""
    samples, _ = heading_samples_by_level(doc)
    return {lvl: len(v) for lvl, v in samples.items()}


def _excl_text(counter: Counter) -> str:
    """排除条数的人话（只列非零项，顺序固定，便于报告逐行对照）。"""
    order = ("表格内", "手写目录条目", "后置节名", "表单标签", "封面大写行", "非编号行")
    bits = [f"{k} {counter[k]}" for k in order if counter.get(k)]
    return "、".join(bits) if bits else "无"


def _mode(values: Iterable[Any]) -> tuple[Any, int, int]:
    """众数；(值, 命中数, 有效样本数)。空集返回 ``(None, 0, 0)``。"""
    vals = [v for v in values if v is not None]
    if not vals:
        return None, 0, 0
    value, n = Counter(vals).most_common(1)[0]
    return value, n, len(vals)


def _template_mode(samples: Sequence[Mapping[str, Any]], key: str, *,
                   jc: bool = False) -> dict[str, Any]:
    """模板侧同一取样口径的实测众数，并判断它够不够当期望值（EZI-0016）。

    ``key='jc'`` 时按未设置 = ``start`` 归一（与产物侧完全同口径）。``usable``
    要求「有效样本 ≥ ``TEMPLATE_MODE_MIN_SAMPLES`` 且众数占比 ≥ ``TEMPLATE_MODE_MIN_SHARE``」
    ——样本太少或模板自己没统一时，众数是噪声，不是版式事实。
    """
    vals = [_jc(p.get("jc")) if jc else p.get(key) for p in samples]
    value, n, tot = _mode(vals)
    share = (n / tot) if tot else 0.0
    return {"value": value, "n": n, "total": tot, "share": share,
            "usable": tot >= TEMPLATE_MODE_MIN_SAMPLES and share >= TEMPLATE_MODE_MIN_SHARE}


def _measured_txt(meas: Mapping[str, Any], unit: str = "") -> str:
    """实测众数的人话：``模板实测 12.0pt（48/77，62%）``。"""
    value = meas.get("value")
    shown = "无" if value is None else f"{value}{unit}"
    return (f"模板实测 {shown}（{meas.get('n', 0)}/{meas.get('total', 0)}，"
            f"{float(meas.get('share') or 0) * 100:.0f}%）")


def _sz_dist_txt(props: Sequence[Mapping[str, Any]], key: str = "sz") -> str:
    """字号分布人话（只列有值的段落，按命中数降序）：``12.0pt×48、18.0pt×10``。"""
    c = Counter(p.get(key) for p in props if p.get(key) is not None)
    if not c:
        return "无可测字号"
    return "、".join(f"{v}pt×{n}" for v, n in c.most_common())


def _no_peak_reason(meas: Mapping[str, Any], unit: str = "") -> str:
    """模板实测为什么不能当期望（样本不足 / 无主峰）的人话。EZI-0016。"""
    if not meas.get("total"):
        return "模板无可测样本"
    if meas["total"] < TEMPLATE_MODE_MIN_SAMPLES:
        return (f"{_measured_txt(meas, unit)}样本不足（须 ≥"
                f"{TEMPLATE_MODE_MIN_SAMPLES} 个）")
    return (f"{_measured_txt(meas, unit)}无主峰（占比须 ≥"
            f"{TEMPLATE_MODE_MIN_SHARE:.0%}）")


def _top_values(props: Sequence[Mapping[str, Any]], key: str, k: int,
                *, jc: bool = False) -> tuple[list[tuple[Any, int]], int]:
    """某一属性的前 ``k`` 高频（降序）与有效样本数：``([(16.0, 7), (14.0, 6)], 14)``。

    ``jc=True`` 时按未设置 = ``start`` 归一（与 ``_template_mode`` 同口径，否则 note 的
    「最高频 X%」和「前二合计 Y%」会各算一套分母）。
    """
    vals = [_jc(p.get(key)) if jc else p.get(key) for p in props]
    c = Counter(v for v in vals if v is not None)
    return c.most_common(k), sum(c.values())


def _body_size_expectation(tpl_props: Sequence[Mapping[str, Any]], spec_pt: Any
                           ) -> tuple[Any, tuple[float, ...] | None, str, str]:
    """正文主字号的期望值/出处（EZI-0016 实测优先；EZI-0022 无主峰给双峰区间）。

    返回 ``(expected, candidates, src, forced)``：``candidates`` 非空时按「产物落在
    任一候选的容差内即 PASS」判；``forced`` 为 ``NA`` 时该项**不判**（``expected``
    也留空），否则为 ``""``。三种情形：

    * **单峰**（众数占比 ≥ ``TEMPLATE_MODE_MIN_SHARE``）→ 候选 = 实测众数。
    * **无主峰**（样本够、最高频占比不够）→ 候选 = 前二高频；二者合计仍须 ≥
      ``TEMPLATE_MODE_MIN_SHARE``，否则分布过散、前二也代表不了模板 → ``N-A``。
      这一支**绝不**读 ``spec_pt``——交大中期 16×7/14×6（50%）、开题 14×8/16×6（47%）
      就是模板自己混排，退回 spec 16pt 等于拿注解常量冒充模板实测（EZI-0022）。
    * **样本不足**（< ``TEMPLATE_MODE_MIN_SAMPLES``）→ 实测不可用，退回 spec 注解值，
      但 note 必须显式标注「期望来自规范注解（非模板实测）」（EZI-0022 A4）。
    """
    meas = _template_mode(tpl_props, "sz")
    if meas["usable"]:
        val = float(meas["value"])
        return val, (val,), f"期望取{_measured_txt(meas, 'pt')}", ""
    if meas["total"] >= TEMPLATE_MODE_MIN_SAMPLES:
        top, total = _top_values(tpl_props, "sz", TEMPLATE_NO_PEAK_TOP_K)
        cand = tuple(float(v) for v, _ in top)
        cover = (sum(n for _, n in top) / total) if total else 0.0
        if len(cand) >= 2 and cover >= TEMPLATE_MODE_MIN_SHARE:
            picks = "、".join(f"{v}pt×{n}" for v, n in top)
            return (cand, cand,
                    f"{_no_peak_reason(meas, 'pt')}；前二高频 {picks}"
                    f"（合计 {cover:.0%}）→ 期望取双峰区间 "
                    f"{'/'.join(f'{v}pt' for v in cand)}"
                    f"（产物落在任一即 PASS，EZI-0022）", "")
        return (None, None,
                f"{_no_peak_reason(meas, 'pt')}；"
                f"{len(cand)} 个最高频合计 {cover:.0%} <"
                f"{TEMPLATE_MODE_MIN_SHARE:.0%}，分布过散、无可靠期望 → N-A"
                f"（不退回 spec 注解，EZI-0022）", NA)
    return (float(spec_pt) if spec_pt is not None else None, None,
            f"期望来自规范注解（非模板实测）{spec_pt}pt："
            f"{_no_peak_reason(meas, 'pt')}（EZI-0022 A4）", "")


def _size_ok_any(candidates: Iterable[float], act: Any,
                 tol: float = SIZE_TOLERANCE_PT) -> bool:
    """产物字号落在候选集合任一值的容差内（EZI-0022 双峰区间）。"""
    return act is not None and any(_size_ok(c, act, tol) for c in candidates)


def _norm_heading_text(text: str) -> str:
    """标题文本归一：去半角/全角空格与制表符（模板「第一章　投标文件的编写」与产物同）。"""
    return "".join(str(text).split())


def _heading_overlap(tpl_rows: Sequence[Mapping[str, Any]],
                     prod_rows: Sequence[Mapping[str, Any]]) -> tuple[int, float]:
    """产物该级样本里有多少条能在模板同级样本中找到同文本对应（归一化后）。"""
    if not prod_rows:
        return 0, 0.0
    tpl_texts = {_norm_heading_text(r["text"]) for r in tpl_rows}
    hit = sum(1 for r in prod_rows if _norm_heading_text(r["text"]) in tpl_texts)
    return hit, hit / len(prod_rows)


def _comparable_heading_rows(
        tpl_rows: Sequence[Mapping[str, Any]],
        prod_rows: Sequence[Mapping[str, Any]]
        ) -> tuple[list[Mapping[str, Any]], int, str, bool]:
    """标题**实测**样本收窄到「与模板同级同文」的那批（EZI-0021；EZI-0025 三项共用）。

    只在「两侧已判为不是同一批标题」（``_heading_overlap`` < ``HEADING_OVERLAP_MIN``）、
    模板该级有样本、且产物有同文对应时收窄；其余情况原样返回
    ``(prod_rows, 0, "", False)``，调用方拼出来的 note 与旧口径逐字相同。

    第 4 项 ``narrowed`` = 「本次按可比样本判」。它为 ``True`` 时调用方**不得**退回全量
    样本（EZI-0025 A3：可比样本为空 → ``N-A``）；产物一条同文对应都没有（重合 0）时
    不算收窄（``narrowed=False``，仍走全量 + spec）——否则新疆二级这类「两侧标题根本
    不是一批」的用例会退化成永远 N-A 的死指标（EZI-0021 决策记录 3）。

    为什么必须收窄（东台招标实测）：模板一级只有 3 个「第X章」（16pt 黑体 center），
    产物一级 34 段里 31 段是条款项/公告条目（10.5pt×14、16pt×11、12pt×9）。拿 34 段的
    众数当「一级标题字号」，测的是「哪类段落最多」，不是「章标题字号是否跟随模板」——
    真章标题 3/3 都是 16pt（=模板），却被判 FAIL。EZI-0016 已经认定「重合率 < 50% 时
    模板实测不能当期望」；同一个理由对实测侧同样成立：产物全体的众数也不能当该级实测值。
    两侧同口径 = 都只认同一批标题。

    EZI-0025：**字号、字体、对齐三项共用这一套样本**。EZI-0021 只接了字号，字体/对齐
    仍按全量样本取众数——EZI-0024 把真章标题改成模板的黑体之后，字体项的实测仍是那批
    条款项的众数（宋体），PASS 的理由是错的：尺子看不见引擎已经改对。
    """
    if not tpl_rows or not prod_rows:
        return list(prod_rows), 0, "", False
    hit, ratio = _heading_overlap(tpl_rows, prod_rows)
    if not hit:
        return list(prod_rows), 0, "", False
    if ratio >= HEADING_OVERLAP_MIN:
        # 不收窄（两侧基本是同一批标题）。note 只在该级样本**全部**同文时加一句「可比
        # 样本 N 段」——让人看清判的是哪批（东台一级 3/3 就是这一支，EZI-0025 A2 要求
        # note 写明可比样本数）；其余情况不加，note 与旧口径逐字相同。
        if hit == len(prod_rows):
            return list(prod_rows), 0, (
                f"；可比样本 {hit} 段（与模板同级同文，覆盖该级全部样本）"), False
        return list(prod_rows), 0, "", False
    tpl_texts = {_norm_heading_text(r["text"]) for r in tpl_rows}
    rows = [r for r in prod_rows if _norm_heading_text(r["text"]) in tpl_texts]
    if not rows:                    # 理论不可达（hit > 0 即至少 1 行同文）；兜底不静默退全量
        return [], len(prod_rows), (
            f"；可比样本 0 段（与模板同级同文）；该级共 {len(prod_rows)} 段全部无同文对应"
            f"（不退回全量样本，EZI-0025 A3）"), True
    dropped = len(prod_rows) - len(rows)
    return rows, dropped, (
        f"；可比样本 {len(rows)} 段（与模板同级同文）；该级共 {len(prod_rows)} 段，"
        f"未参与 {dropped} 段（模板同级无同文对应，属标题级别结构差异，见 EZI-0021）"), True


def _same_text_pairs(
        tpl_rows: Sequence[Mapping[str, Any]],
        prod_rows: Sequence[Mapping[str, Any]]
        ) -> list[tuple[Mapping[str, Any], Mapping[str, Any]]]:
    """产物段 ↔ 模板同级同文段成对（归一化文本相同；模板重名取首个）。EZI-0025。

    与 ``_comparable_heading_rows`` 同一条「同文」判据（``_norm_heading_text``）：那条
    收窄**实测样本**，这条取**模板那一侧的对应段**当期望出处——两侧比的必须是同一批段。
    """
    by_text: dict[str, Mapping[str, Any]] = {}
    for r in tpl_rows:
        by_text.setdefault(_norm_heading_text(r["text"]), r)
    out: list[tuple[Mapping[str, Any], Mapping[str, Any]]] = []
    for p in prod_rows:
        t = by_text.get(_norm_heading_text(p["text"]))
        if t is not None:
            out.append((p, t))
    return out


def _dist_txt(props: Sequence[Mapping[str, Any]], key: str, *,
              jc: bool = False, limit: int = 3) -> str:
    """某属性的取值分布人话：``仿宋_GB2312;仿宋×6、黑体×6、仿宋×2``（降序，最多 ``limit`` 种）。"""
    vals = [_jc(p.get(key)) if jc else p.get(key) for p in props]
    c = Counter(v for v in vals if v is not None)
    if not c:
        return "无可测值"
    shown = "、".join(f"{v}×{n}" for v, n in c.most_common(limit))
    return shown if len(c) <= limit else f"{shown} 等 {len(c)} 种"


def _heading_prop(key: str, exp: Any, act: Any, *, tot: int,
                  narrowed: bool, cmp_n: int,
                  candidates: Iterable[float] | None = None) -> tuple[str, str]:
    """标题一项的判定，并给出 EZI-0025 A3 的 N-A 理由（正常判时为空串）。

    收窄后该属性在可比样本里一个值都取不到（例如同文段全都没写 ``w:rFonts``）→ ``N-A``
    并写明「不退回全量样本」：退回全量等于又拿被排除的条款项去判章标题。

    ``candidates`` 非空 = 同文批次无单一主峰、改判双峰区间（EZI-0039；地板与正文主字号
    EZI-0022 同一条 ``TEMPLATE_MODE_MIN_SHARE``）：产物落在任一候选的容差内即 PASS。
    """
    if narrowed and not tot:
        return NA, (f"；可比样本 {cmp_n} 段均无可测{_HEAD_PROP_LABEL[key]}"
                    f"（不退回全量样本，EZI-0025 A3）")
    if key == "font_size_pt":
        ok = (_size_ok_any(candidates, act) if candidates
              else _size_ok(exp, act))
    else:
        ok = act is not None and exp is not None and act == exp
    return _verdict(ok), ""


def _same_text_expectation(key: str, field: str,
                           tpl_same: Sequence[Mapping[str, Any]],
                           *, unit: str = ""
                           ) -> tuple[Any, tuple[float, ...] | None, str]:
    """同文成对那批模板段落的期望值（EZI-0039：取众数，不再要求 100% 一致）。

    返回 ``(expected, candidates, note)``。``expected is None`` = 这批同文证据给不出期望
    （``note`` 写明为什么），调用方接着走「模板该级实测 / spec」原护栏。

    * **众数占比 ≥ ``TEMPLATE_MODE_MIN_SHARE``（0.60）即采信**——与
      ``earlyz/spec/decompiler.py::_DOMINANCE_DEFAULT`` 同一口径，禁另造阈值。深圳招标
      一级标题模板同文 10 段：16.0pt×9 + 第九章特例 22.0pt×1（90%）。旧口径要求
      「同文段 100% 一致」才算证据，于是判「不统一」→ 退回 spec 14.0pt；产物照模板排
      16pt 反被判 FAIL，引擎改成 14pt 反而 PASS——尺子把引擎往错的方向拉（EZI-0037）。
    * **无单一主峰 + 字号**：前二高频当双峰区间（``candidates``），地板仍是
      ``TEMPLATE_MODE_MIN_SHARE``，与正文主字号 EZI-0022 同一口径；前二合计 <60% =
      分布过散，给不出期望（调用方退回 spec 并写明依据，不静默取首个值）。
    * **无单一主峰 + 字体/对齐**：类别属性一半一半时两个值都「对」，造候选集合等于放弃
      判据（43 对基线里「模板同文 2 段 仿宋_GB2312×1、黑体×1」的项也会凭空变 FAIL），
      只写分布与前二占比，退回原护栏。
    """
    meas = _template_mode(tpl_same, field, jc=(key == "jc"))
    n, tot = int(meas["n"]), int(meas["total"])
    label = _HEAD_PROP_LABEL[key]
    dist = (_sz_dist_txt(tpl_same, field) if key == "font_size_pt"
            else _dist_txt(tpl_same, field, jc=(key == "jc")))
    if not tot:                         # 同文段一个可测值都没有（如全篇不写 rFonts）
        return None, None, f"模板同文 {len(tpl_same)} 段{label}不统一（{dist}）"
    share = float(meas["share"])
    if share >= TEMPLATE_MODE_MIN_SHARE:
        val = meas["value"]
        if key == "font_size_pt" and val is not None:
            val = float(val)
        if n == tot:                    # 全一致：EZI-0025 原文案（1:1 直接证据）不变
            return val, None, (f"期望取{_measured_txt(meas, unit)}"
                               f"（与产物同级同文成对 {len(tpl_same)} 段，"
                               f"直接证据不受样本下限约束，EZI-0025）")
        return val, None, (
            f"期望取{_measured_txt(meas, unit)}（与产物同级同文成对 "
            f"{len(tpl_same)} 段；模板同文 {len(tpl_same)} 段{label} {dist}，"
            f"最高频 {share:.0%} ≥{TEMPLATE_MODE_MIN_SHARE:.0%} 取众数，EZI-0039）")
    top, total = _top_values(tpl_same, field, TEMPLATE_NO_PEAK_TOP_K, jc=(key == "jc"))
    cover = (sum(k for _, k in top) / total) if total else 0.0
    picks = "、".join(f"{v}{unit}×{k}" for v, k in top)
    why = (f"模板同文 {len(tpl_same)} 段{label}不统一（{dist}）；"
           f"最高频 {share:.0%} <{TEMPLATE_MODE_MIN_SHARE:.0%} 无单一主峰，"
           f"前二高频 {picks}（合计 {cover:.0%}）")
    if key == "font_size_pt" and len(top) >= 2 and cover >= TEMPLATE_MODE_MIN_SHARE:
        cand = tuple(float(v) for v, _ in top if v is not None)
        return (cand, cand,
                f"{why} ≥{TEMPLATE_MODE_MIN_SHARE:.0%} → 期望取双峰区间 "
                f"{' / '.join(f'{v}{unit}' for v in cand)}"
                f"（产物落在任一即 PASS，地板同 EZI-0022）")
    return None, None, why


def _heading_expectation(rule: Mapping[str, Any], key: str,
                         tpl_rows: Sequence[Mapping[str, Any]],
                         prod_rows: Sequence[Mapping[str, Any]], *,
                         where: str, unit: str = ""
                         ) -> tuple[Any, tuple[float, ...] | None, str]:
    """标题某一项的期望值/候选/出处（EZI-0016 两道护栏；EZI-0025 同文成对优先）。

    三个来源按顺序：

    1. **同文成对（EZI-0025 / EZI-0039）**：模板该级里「与产物同文本」的那批段落。产物
       那一段与模板这一段是一一对应的同一批标题，不需要「≥ ``TEMPLATE_MODE_MIN_SAMPLES``
       个」的统计前提——东台一级只有 3 个「第X章」，该当期望的正是这 3 段的 3/3 黑体，
       而不是 spec 注解。EZI-0039 起**取众数**（≥60% 即采信，见
       ``_same_text_expectation``），不再要求同文段 100% 一致——「不统一」不是退回 spec
       的理由，只有「前二合计 <60%（分布过散）」才是。
    2. **模板该级实测众数（EZI-0016）**：样本量/主峰 + 文本重合率两道护栏都过才用。
    3. **退回 spec 规则**，并把「为什么没用模板实测」逐条写进 note。

    两道护栏：① 模板该级样本量/主峰不足（``_no_peak_reason``）；② 模板与产物的标题
    文本重合率 < ``HEADING_OVERLAP_MIN``——两侧比的不是同一批标题（模板「第X章」vs
    产物条款项）时，模板实测就是错位比较，宁可退回 spec 也不制造假 FAIL。

    第二项返回值是「双峰候选」（仅字号可能非空），第三项是 note。
    """
    field = _HEAD_PROP_FIELD[key]
    spec_val = rule.get(key)
    same = _same_text_pairs(tpl_rows, prod_rows)
    tpl_same = [t for _, t in same]
    same_why = ""
    if len(tpl_same) >= HEADING_SAME_TEXT_MIN:
        exp_same, cand_same, same_note = _same_text_expectation(
            key, field, tpl_same, unit=unit)
        if exp_same is not None:
            return exp_same, cand_same, same_note
        same_why = same_note
    elif same:
        same_why = f"模板同文仅 {len(same)} 段（须 ≥{HEADING_SAME_TEXT_MIN}）"
    meas = _template_mode(tpl_rows, field, jc=(key == "jc"))
    hit, overlap = _heading_overlap(tpl_rows, prod_rows)
    if meas["usable"] and overlap >= HEADING_OVERLAP_MIN:
        val = meas["value"]
        if key == "font_size_pt" and val is not None:
            val = float(val)
        return val, None, f"期望取{_measured_txt(meas, unit)}"
    why: list[str] = []
    if same_why:
        why.append(same_why)
    if not meas["total"]:
        why.append(f"模板无编号{where}样本")
    elif not meas["usable"]:
        why.append(_no_peak_reason(meas, unit))
    else:
        why.append(f"模板同级样本 {len(tpl_rows)} 个，与产物文本重合 {hit}/{len(prod_rows)}"
                   f"（{overlap * 100:.0f}%），两侧不是同一批标题（须 ≥"
                   f"{HEADING_OVERLAP_MIN:.0%}）")
    why.append("退回 spec 规则")
    return (float(spec_val) if key == "font_size_pt" and spec_val is not None else spec_val,
            None, "；".join(why))


def body_text_paragraphs(paras: Sequence[tuple[Any, dict[str, Any]]], *,
                         exclude_form_labels: bool = True) -> list[dict[str, Any]]:
    """正文实测样本（EZI-0016）：``body_paragraphs`` 再去掉题注段。

    ``body_paragraphs`` 已排掉标题层级、目录条目与空段（默认还有表单标签，EZI-0020）；
    这里再排掉 ``Caption*``/「题注/图注/表注」样式的题注段——题注是 10.5pt 小字，
    混进正文众数会把主字号带偏。两侧（模板 / 产物）必须用同一个函数，否则又是
    「一侧 spec、一侧实测」的老病。
    """
    return [p for p in body_paragraphs(paras, exclude_form_labels=exclude_form_labels)
            if not _CAPTION_STYLE.match(p["style_id"] or "")]


def _section_margins(doc: dict[str, Any]) -> list[dict[str, int | None]]:
    """逐节 ``w:pgMar``（按文档顺序；缺 pgMar 的节给空 dict）。EZI-0016。"""
    out: list[dict[str, int | None]] = []
    for sp in _sections(doc):
        mar = sp.find(W + "pgMar")
        out.append({k: _int(mar.get(W + k)) for k in _MARGIN_KEYS} if mar is not None else {})
    return out


def _margin_sig(mar: Mapping[str, Any]) -> tuple:
    return tuple(mar.get(k) for k in _MARGIN_KEYS)


def _margin_txt(sig: Sequence[Any]) -> str:
    return "/".join("—" if v is None else str(v) for v in sig)


def _margin_seq_txt(seq: Sequence[Mapping[str, Any]]) -> str:
    """节序列压成「值 ×重复数」的人话（节多时不刷屏）：``1417/1701/1134/1701 ×8；1440/… ×1``。"""
    runs: list[tuple[tuple, int]] = []
    for mar in seq:
        sig = _margin_sig(mar)
        if runs and runs[-1][0] == sig:
            runs[-1] = (sig, runs[-1][1] + 1)
        else:
            runs.append((sig, 1))
    return "；".join(_margin_txt(sig) + (f" ×{cnt}" if cnt > 1 else "") for sig, cnt in runs)


def _idx_ranges(idxs: Sequence[int]) -> str:
    """``[3,4,5,6,7]`` → ``3~7``；``[0,2]`` → ``0、2``。"""
    out: list[str] = []
    start: int | None = None
    for i, n in enumerate(idxs):
        if start is None:
            start = n
        if i + 1 == len(idxs) or idxs[i + 1] != n + 1:
            out.append(str(start) if start == n else f"{start}~{n}")
            start = None
    return "、".join(out)


def _margin_diff(tpl: Mapping[str, Any], prod: Mapping[str, Any]) -> bool:
    """单节四边是否超差（缺值也算超差：没有可比的值就是没复刻）。"""
    for k in _MARGIN_KEYS:
        a, b = tpl.get(k), prod.get(k)
        if a is None or b is None or abs(int(b) - int(a)) > MARGIN_TOLERANCE_TWIPS:
            return True
    return False


def _sect_props(doc: dict[str, Any]) -> dict[str, Any]:
    body = doc["document"].find(W + "body")
    sp = body.find(W + "sectPr") if body is not None else None
    if sp is None:
        return {}
    out: dict[str, Any] = {}
    mar = sp.find(W + "pgMar")
    if mar is not None:
        out["pg_mar"] = {k: _int(mar.get(W + k)) for k in
                         ("top", "right", "bottom", "left", "header", "footer", "gutter")}
    sz = sp.find(W + "pgSz")
    if sz is not None:
        out["pg_sz"] = {"w": _int(sz.get(W + "w")), "h": _int(sz.get(W + "h"))}
    cols = sp.find(W + "cols")
    if cols is not None:
        out["cols"] = {"num": cols.get(W + "num") or "1", "space": cols.get(W + "space")}
    return out


def _sections(doc: dict[str, Any]) -> list[Any]:
    return list(doc["document"].iter(W + "sectPr"))


def _section_texts(doc: dict[str, Any]) -> list[list[str]] | None:
    """按文档顺序返回每节的段落文本（与 ``_sections`` 同序）。

    分节符（段内 ``w:pPr/w:sectPr``）结束的段落算**上一节**的末段；body 级
    ``w:sectPr`` 结束最后一节。与引擎 ``section_geometry._section_blocks`` 同口径
    ——角色判定必须和「引擎按角色套几何」看同一批段落，否则两边认的节不是同一节。
    段落树与 sectPr 清单对不上（sectPr 不在 body 流里等）时返回 ``None``，
    调用方退回逐节索引口径。
    """
    body = doc["document"].find(W + "body")
    if body is None:
        return None
    out: list[list[str]] = []
    pending: list[str] = []
    for el in body.iter():
        if not isinstance(el.tag, str):
            continue
        if el.tag == W + "p":
            txt = "".join(t.text or "" for t in el.iter(W + "t")).strip()
            ppr = el.find(W + "pPr")
            sp = ppr.find(W + "sectPr") if ppr is not None else None
            if sp is not None:
                if txt:
                    pending.append(txt)
                out.append(pending)
                pending = []
            elif txt:
                pending.append(txt)
        elif el.tag == W + "sectPr":
            parent = el.getparent()
            if parent is not None and parent.tag == W + "body":
                out.append(pending)
                pending = []
    return out


def _norm_section_heading(text: str) -> str:
    return _SECTION_PLACEHOLDER_RE.sub("", text or "").strip().lower()


def _section_roles(texts: Sequence[str], *, first: bool) -> tuple[str, ...]:
    """一节的节角色（只认标签式短标题，精确匹配，不猜正文语义）。

    词表与引擎同源（见 ``_SECTION_ROLE_HEADINGS`` 的注释）。三条防误判：
    目录条目以页码结尾（「参考文献59」）不是标题；``toc`` 命中后只认目录本身
    （目录页会列出参考文献/致谢/附录条目）；第一节默认封面，但它已经有明确角色
    时不硬加封面（「无封面、直接从摘要开始」的文档不该映射到模板封面节）。
    """
    roles: set[str] = set()
    for raw in texts:
        t = _norm_section_heading(raw)
        if not t or len(t) > 40:
            continue
        role = _SECTION_ROLE_HEADINGS.get(t)
        if role:
            roles.add(role)
            continue
        if t.startswith("附录") and len(t) <= 8:
            roles.add("appendix")
            continue
        # EZI-0075：附件表单区标题（冒号形也是独立标题行）——引擎同口径。
        _raw_clean = (raw or "").strip()
        if (len(_raw_clean) <= 10 and "\t" not in _raw_clean
                and _SECTION_ATTACHMENT_RE.match(t)):
            roles.add("attachment")
            continue
        if re.search(r"\d$", t):          # 目录条目以页码结尾，不是标题
            continue
        if _SECTION_CHAPTER_RE.match(t) or _SECTION_NUMBERED_RE.match(t):
            roles.add("body")
    if "toc" in roles:
        if len(texts) <= _TOC_SECTION_MAX_PARAS:
            roles = {"toc"} | ({"cover"} & roles)
        else:
            # EZI-0075：超长节里的一行目录不是节身份（建设施工 1628 段合同正文
            # 大节含一行「目录」被吞成 toc 节 → 配对模板唯一目录节 1680 →
            # 沿用灌给附件节）。长节目录透明，与引擎 ``_detect_roles`` 同口径。
            roles.discard("toc")
    if first and (not roles or any("封面" in raw or "毕业设计" in raw for raw in texts)):
        roles.add("cover")
    return tuple(sorted(roles))


def _section_role_seq(doc: dict[str, Any]) -> list[tuple[str, ...]] | None:
    """逐节角色；段落结构与 ``sectPr`` 清单对不上时返回 ``None``（→ 退回索引口径）。"""
    texts = _section_texts(doc)
    n = len(_sections(doc))
    if texts is None or len(texts) != n:
        return None
    return [_section_roles(t, first=(i == 0)) for i, t in enumerate(texts)]


def _role_txt(roles: Sequence[str]) -> str:
    """角色集合的人话：``(封面)`` / ``(正文/目录)``；认不出返回空串。"""
    if not roles:
        return ""
    return "(" + "/".join(_SECTION_ROLE_LABEL.get(r, r) for r in roles) + ")"


def _pair_margin_sections(
    tpl_roles: Sequence[Sequence[str]], prod_roles: Sequence[Sequence[str]],
    tpl_margins: Sequence[Mapping[str, Any]] | None = None,
) -> tuple[list[tuple[int, int]], str, str]:
    """给每个产物节选一个模板节：``([(产物节, 模板节)], 口径, 说明)``。

    口径三档（A1/A3，EZI-0027）：

    * ``index`` —— 两侧节数相同：逐节按索引对齐（旧行为，不按角色改动等长文档）；
    * ``role`` —— 节数不同且角色能认出 ≥2 个模板节：按**角色**配对
      （封面/声明/摘要/ABSTRACT/目录/正文/参考文献/致谢/附录），不再机械比前 N 节
      ——本科论文产物多一个「声明」节，按索引比会把产物 ABSTRACT(1417) 与模板
      目录(1440) 对成一对，把全是正确的边距判成 FAIL；
    * ``index-fallback`` —— 节数不同但角色认不出来（只认出 ≤1 个模板节 /
      读不出逐节段落）：退回逐节索引，**说明里写明原因**，不静默跳过。

    角色配对规则与引擎 ``section_geometry._pair_sections`` 一致：映射只前进不后退，
    同角色按出现顺序消费；产物里认不出角色（或该角色模板已用尽）的节**沿用前一节**
    的模板节——正文被拆成多个子节时不至于掉回封面几何。
    """
    n_tpl, n_prod = len(tpl_roles), len(prod_roles)
    if n_tpl == n_prod:
        return [(i, i) for i in range(n_prod)], "index", "两侧节数相同"
    role_pos: dict[str, list[int]] = {}
    for ti, roles in enumerate(tpl_roles):
        for r in roles:
            role_pos.setdefault(r, []).append(ti)
    # EZI-0074：同角色多节混少数派变体（118 节合同实测：目录 1680 变体、
    # 横向附件页 1100/1580/280）——按序配对会把变体当期望。与引擎
    # ``section_geometry._role_map`` 同口径：消费前把**角色页边距众数**
    # 的节排到队首（组内稳定保序）；角色均匀时排序恒等、配对行为不变。
    if tpl_margins is not None:
        from collections import Counter as _Counter
        for r, idxs in role_pos.items():
            if len(idxs) < 2 or any(tpl_margins[ti] is None for ti in idxs):
                continue
            def _mkey(ti):
                m = tpl_margins[ti] or {}
                return tuple(sorted((str(k), str(v)) for k, v in m.items()))
            mode_key = _Counter(_mkey(ti) for ti in idxs).most_common(1)[0][0]
            role_pos[r] = sorted(idxs, key=lambda ti: 0 if _mkey(ti) == mode_key else 1)
    used = {r: 0 for r in role_pos}
    mapping: dict[int, int] = {}
    last_ti = 0
    for pi, roles in enumerate(prod_roles):
        best: int | None = None
        for r in roles:                   # roles 已排序 → 多角色命中取最靠前的模板节
            seq = role_pos.get(r) or []
            while used.get(r, 0) < len(seq) and seq[used[r]] < last_ti:
                used[r] += 1
            if used.get(r, 0) < len(seq):
                cand = seq[used[r]]
                if best is None or cand < best:
                    best = cand
        if best is None:
            # EZI-0075：单调前进耗尽但该角色模板节全在身后——合同/标书表单区
            # 装订在正文前、产物附件统一放尾部，两侧顺序天然不对称。回退取该
            # 角色队首（mode 重排后=众数节）；期望走候选集（同角色全部形态）。
            for r in roles:
                seq = role_pos.get(r) or []
                if seq:
                    best = seq[0]
                    break
        if best is not None:
            mapping[pi] = best
            for r in roles:
                seq = role_pos.get(r) or []
                if used.get(r, 0) < len(seq) and seq[used[r]] == best:
                    used[r] += 1
            last_ti = best
    matched = len(set(mapping.values()))
    if matched < 2:                       # 只认出封面一节 = 不足以信角色映射
        return ([(i, i) for i in range(min(n_tpl, n_prod))], "index-fallback",
                f"角色识别失败（只在 {matched} 个模板节上认出可比角色）")
    pairs: list[tuple[int, int]] = []
    cur = 0
    for pi in range(n_prod):
        if pi in mapping:
            cur = mapping[pi]
        pairs.append((pi, cur))
    paired = {ti for _, ti in pairs}
    extra = [f"节{ti}{_role_txt(tpl_roles[ti])}" for ti in range(n_tpl) if ti not in paired]
    note = ""
    if extra:
        note = f"模板{'、'.join(extra)}未被配对（产物无同角色节）"
    return pairs, "role", note


def _margin_pair_txt(pairs: Sequence[tuple[int, int]],
                     prod_roles: Sequence[Sequence[str]] | None,
                     tpl_roles: Sequence[Sequence[str]] | None,
                     limit: int = 12) -> str:
    """配对表人话：``节0(封面)→模板节0(封面)、节1(声明)→模板节0(封面)…``。"""
    items = []
    for pi, ti in pairs:
        items.append(f"节{pi}{_role_txt(prod_roles[pi]) if prod_roles else ''}"
                     f"→模板节{ti}{_role_txt(tpl_roles[ti]) if tpl_roles else ''}")
    if len(items) > limit:
        return "、".join(items[:limit]) + f"…（共 {len(items)} 节）"
    return "、".join(items)


def _seq_txt(seq: Sequence[str], limit: int = 12) -> str:
    """节序列的人话：超过 ``limit`` 节时保留头尾，中间用 ``…`` 省略。"""
    items = [str(x) for x in seq]
    if len(items) <= limit:
        return "[" + ",".join(items) + "]"
    head = items[: limit // 2]
    tail = items[-(limit - len(head)):]
    return "[" + ",".join(head) + ",…," + ",".join(tail) + "]"


def _section_cols_fact(doc: dict[str, Any]) -> dict[str, Any]:
    """逐节读 ``w:cols/@w:num``；正文分栏取**正文节**（EZI-0069）。

    为什么必须逐节：旧判法拿模板 ``structure.body_cols``（「第一个非首节 ≥2 栏」），
    去比产物 ``body/w:sectPr``（末节）。ACM 的 cols 序列两侧完全一样（``[1,3,1,2]``），
    索引 1 是 3 栏作者区、末节是 2 栏正文，于是判成「模板 3 产物 2」FAIL——
    比的是两个不同的节，不是版式差异。

    EZI-0015 A2 取末节治了「正文≠第一个多栏节」（ACM：末节=正文）。但期刊
    另一形态（自动化学报，20261121 基线两案实锤）：模板 ``[1,2,1]``——正文在
    **中段**双栏、末节是收稿信息**单栏**小节；产物 ``[1,2]`` 末节恰是正文 →
    「取末节」把模板尾部小节当正文比 → 假红 exp=1 act=2。

    EZI-0069/0070 正文节配对（机制层，无模板常量）：**全文档段量最大**的节为
    正文节（段量 = 该节非空文本段数，``_section_texts`` 与引擎
    ``section_geometry`` 同口径）；平票取**晚**节（正文通常在后）。0070 去
    0069 首版的「≥2 栏限定」——合同族实锤：双栏示例区（3 段）被限定条件
    误当正文，真正文是 240 段单栏大节。节-段树对不上（``_section_texts``
    返回 None）或全空文档 → 诚实退回取末节。缺 ``w:cols`` 按 1 栏计。
    """
    sects = _sections(doc)
    seq: list[str] = []
    for sp in sects:
        cols = sp.find(W + "cols")
        seq.append((cols.get(W + "num") or "1") if cols is not None else "1")
    base = {"seq": seq, "n": len(sects)}
    if not seq:
        return {**base, "num": None, "body_idx": None, "how": "无节"}
    texts = _section_texts(doc)
    if texts is not None and len(texts) == len(seq):
        counts = [sum(1 for x in ts if x) for ts in texts]
        if any(counts):
            # EZI-0070：正文节=**全文档段量最大**的节（去 EZI-0069 首版的
            # 「≥2 栏限定」）。合同族实锤：模板 [1,2,1,1,1] 第 2 节是 3 段的
            # 双栏「甲乙双方信息」示例区、真正文在 240 段的单栏末节——限定
            # ≥2 栏会把示例区当正文 → exp=2 act=1 假红。全文档段量最大对
            # 期刊（47 段正文双栏 vs 收稿 1 段）、ACM（末节正文）、合同
            # （单栏大节）三种形态都选得到真正文。平票取晚节（正文通常在后）。
            body_idx = max(range(len(seq)), key=lambda i: (counts[i], i))
            return {**base, "num": seq[body_idx], "body_idx": body_idx,
                    "how": f"正文节=第{body_idx + 1}节（全文档段量最大 "
                            f"{counts[body_idx]} 段）"}
    return {**base, "num": seq[-1], "body_idx": len(seq) - 1,
            "how": "取末节（段树不可配/全空文档）"}



def _line_dist_txt(dist: Mapping[tuple[Any, Any], int], total: int,
                   limit: int = 4) -> str:
    """行距分布：``240/exact×103（未设 115）`` 这种写法，按命中数降序。"""
    setd = {k: v for k, v in dist.items() if k[0] is not None}
    unset = sum(v for k, v in dist.items() if k[0] is None)
    parts = [f"{k[0]}/{k[1]}×{v}" for k, v in
             sorted(setd.items(), key=lambda kv: (-kv[1], str(kv[0])))[:limit]]
    if not parts:
        return f"未设 {unset}/{total}"
    if unset:
        parts.append(f"未设×{unset}")
    return "、".join(parts) + f"（正文 {total} 段）"


def _default_para_style_line(doc: dict[str, Any]) -> tuple[int | None, str | None]:
    """默认段落样式的 ``w:spacing/@w:line``（正文段落全未设行距时的兜底）。

    ``_ppr_chain`` 只跟显式 ``w:pStyle``，没写 pStyle 的段落不会继承 ``Normal``
    的 pPr。模板全靠默认样式定行距时，段落级实测会全空——这时取默认段落样式
    （``w:type="paragraph" w:default="1"``，缺则 ``Normal``）的样式链末次设值，
    仍然来自模板 styles.xml，不是凭空造值。
    """
    root = doc.get("styles")
    if root is None:
        return None, None
    styles = _Styles(root)
    target = None
    for s in root.iter(W + "style"):
        if s.get(W + "type") == "paragraph" and s.get(W + "default") in ("1", "true", "on"):
            target = s
            break
    if target is None:
        target = next((s for s in root.iter(W + "style")
                       if s.get(W + "styleId") == "Normal"), None)
    if target is None:
        return None, None
    chain: list[Any] = []
    if styles.ppr_default is not None:
        chain.append(styles.ppr_default)
    for st_el in styles.chain(target.get(W + "styleId")):
        ppr = st_el.find(W + "pPr")
        if ppr is not None:
            chain.append(ppr)
    line = rule = None
    for el in chain:
        sp = el.find(W + "spacing")
        if sp is not None and sp.get(W + "line") is not None:
            line = _int(sp.get(W + "line"))
            rule = sp.get(W + "lineRule") or "auto"
    return line, rule


def line_spacing_fact(doc: dict[str, Any]) -> dict[str, Any]:
    """正文行距实测（模板/产物**同一个函数**，EZI-0015 A3）。

    * 样本 = :func:`body_paragraphs`（排除表格内、标题层级、目录条目、空段、
      **表单标签**——EZI-0020：交大中期 6 段「学生姓名:」这类字段名是 ``500/exact``，
      与 6 段真正文 ``400/atLeast`` 打平后按文档序取到字段名，把已排对的产物判 FAIL）；
    * 众数只统计**设了** ``@w:line`` 的段落。旧实现把 ``None`` 也算进众数——
      IEEE 产物 218 段里 115 段没设、103 段是 ``240/exact``，众数成了 ``None``，
      于是「产物明明有 240/exact 105 段」被判成 actual=None FAIL；
    * 段落级全空 → 默认段落样式兜底；仍无 → ``line=None``（调用方判 ``N-A``，
      note 写「模板正文未设固定行距」），绝不拿默认值凑期望。
    """
    paras = body_paragraphs(document_paragraphs(doc))
    dist: Counter = Counter((p["line"], p["line_rule"]) for p in paras)
    setd = Counter({k: v for k, v in dist.items() if k[0] is not None})
    fact: dict[str, Any] = {"dist": dist, "total": len(paras),
                            "set": sum(setd.values()), "source": "正文段落"}
    if setd:
        (line, rule), n = setd.most_common(1)[0]
        fact.update({"line": line, "rule": rule, "n": n})
        return fact
    dline, drule = _default_para_style_line(doc)
    if dline is not None:
        fact.update({"line": dline, "rule": drule, "n": 0, "source": "默认段落样式"})
        return fact
    fact.update({"line": None, "rule": None, "n": 0, "source": "无"})
    return fact


def _line_value_txt(line: Any, rule: Any) -> str | None:
    """行距值的人话：``264/auto（1.10 倍行距）``；无值返回 ``None``。"""
    if line is None:
        return None
    mult = f"（{int(line) / 240:.2f} 倍行距）" if rule == "auto" else ""
    return f"{line}/{rule}{mult}"


def _field_instrs(root: Any) -> list[str]:
    """部件内所有域指令：``w:fldSimple/@w:instr`` + 每段的 ``w:instrText`` 拼接。

    同一段里多个域会被拼成一串，按 2+ 空格再切开，避免 TOC 指令把后面的
    PAGEREF 一起带进备注。
    """
    out: list[str] = []
    for fs in root.iter(W + "fldSimple"):
        instr = (fs.get(W + "instr") or "").strip()
        if instr:
            out.append(instr)
    for p in root.iter(W + "p"):
        parts = [(t.text or "") for t in p.iter(W + "instrText")]
        if not parts:
            continue
        joined = "".join(parts).strip()
        if not joined:
            continue
        out.extend(s.strip() for s in re.split(r"\s{2,}", joined) if s.strip())
    return out


def _toc_instrs(doc: dict[str, Any]) -> list[str]:
    """该 docx 主文档部件里命中的 TOC 域指令（``TOC \\o "1-3" \\h`` 这类）。"""
    return [i for i in _field_instrs(doc["document"]) if _TOC_INSTR.search(i)]


def _hf_stats(doc: dict[str, Any], kind: str) -> dict[str, Any]:
    """页眉/页脚部件统计：部件数、字号众数、是否含 PAGE 域、有没有文字。"""
    pat = _HEADER_PART if kind == "header" else _FOOTER_PART
    parts = {n: root for n, root in doc["hf_parts"].items() if pat.match(n)}
    sizes: list[float] = []
    has_page = False
    texts: list[str] = []
    for root in parts.values():
        for sz in root.iter(W + "sz"):
            v = _int(sz.get(W + "val"))
            if v:
                sizes.append(v / 2.0)
        if any(_PAGE_INSTR.search(i) for i in _field_instrs(root)):
            has_page = True
        txt = "".join(t.text or "" for t in root.iter(W + "t")).strip()
        if txt:
            texts.append(txt)
    size, n, _ = _mode(sizes)
    return {"parts": len(parts), "size": size, "size_n": n, "size_total": len(sizes),
            "has_page": has_page, "has_text": bool(texts),
            "text": " / ".join(texts)[:40]}


def _hf_refs(doc: dict[str, Any], kind: str) -> int:
    return sum(len(s.findall(W + kind + "Reference")) for s in _sections(doc))


def _page_runs(p: Any) -> list[Any]:
    """段落里承载 PAGE 域的 run（``fldSimple`` 取它包住的 run 与它自己的 rPr）。

    只看**域**，不看文字：``_PAGE_INSTR`` 带词边界，不会命中 PAGEREF/NUMPAGES。
    """
    runs: list[Any] = []
    for fs in p.iter(W + "fldSimple"):
        if _PAGE_INSTR.search(fs.get(W + "instr") or ""):
            # 先取域里的 run：``w:fldSimple`` 自己通常没有 rPr，先看它会把样式链
            # 默认值（Header 样式 11pt）当成 run 直接值，把域 run 的 21 半磅盖掉。
            runs.extend(fs.iter(W + "r"))
            runs.append(fs)
    for r in p.iter(W + "r"):
        if _PAGE_INSTR.search("".join(t.text or "" for t in r.iter(W + "instrText"))):
            runs.append(r)
    return runs


def _page_para_size(p: Any, styles: "_Styles") -> tuple[float | None, str]:
    """承载 PAGE 域段落的**有效**字号（pt）与证据出处（EZI-0044）。

    解析顺序（都走样式链，``docDefaults`` → ``pStyle``/``basedOn`` → 段落 ``rPr``
    → run ``rPr``，越靠后越强）：

    1. PAGE 域的 run 直接 ``w:sz``；
    2. 同段文字 run（复杂域的数字常在另一个 run 里，域代码 run 只有 ``instrText``）；
    3. 段落标记 ``w:pPr/w:rPr/w:sz`` / 段落样式 / 文档默认。

    这样「模板靠 Header 样式给页码 9pt」也能当证据；三条都拿不到就是 ``None``，
    调用方判 ``N-A``，绝不拿 Word 默认值凑期望。
    """
    for run in _page_runs(p):
        half = _int(_val(_rpr_chain(p, styles, run), "sz"))
        if half:
            return half / 2.0, "PAGE run"
    for r in p.iter(W + "r"):
        if any((t.text or "").strip() for t in r.findall(W + "t")):
            half = _int(_val(_rpr_chain(p, styles, r), "sz"))
            if half:
                return half / 2.0, "同段文字 run"
            break
    half = _int(_val(_rpr_chain(p, styles), "sz"))
    return (half / 2.0, "段落标记/样式") if half else (None, "")


def _hf_page_size(doc: dict[str, Any], kind: str) -> dict[str, Any]:
    """页眉/页脚里 **PAGE 段落**的有效字号（EZI-0044）。

    为什么不能用 :func:`_hf_stats` 的「全部件 ``w:sz`` 众数」当页眉期望：页眉主体是
    版头/编号**文字**（矿大模板 37 个页眉部件、北航 24 个，字号 10.5/14 混排），按
    11-10 红线那些文字不迁，它们的字号与产物没有任何继承关系。能继承的只有 PAGE
    域链 + 段落格式，所以基准必须锚在承载 PAGE 的那个段落上（页脚同理可用，但页脚
    页码字号沿用旧口径，见 ``页脚页码字号`` 的 note）。
    """
    pat = _HEADER_PART if kind == "header" else _FOOTER_PART
    styles = _Styles(doc["styles"])
    sizes: list[float] = []
    hows: Counter[str] = Counter()
    for name, root in doc["hf_parts"].items():
        if not pat.match(name):
            continue
        for p in root.iter(W + "p"):
            if not _page_runs(p):
                continue
            size, how = _page_para_size(p, styles)
            if size is not None:
                sizes.append(size)
                hows[how] += 1
    size, n, _total = _mode(sizes)
    return {"size": size, "size_n": n, "size_total": len(sizes),
            "how": "、".join(f"{k} {v}" for k, v in hows.most_common())}


# ───────────────────── 题注绑定（EZI-0011：图题绑图 / 表题绑表）─────────────────────


def flow_text(p: Any) -> str:
    """段落**自身**文本：只算流内 ``w:t``，图形/文本框里的字不算段落文字。

    为什么必须这样切：真实样本（``output/毕业设计.docx``）里有 30 张「图 + 题注
    文本框」的图对象——题注字写在 ``w:drawing`` 内嵌的 ``w:txbxContent`` 里。
    按 ``p.iter(w:t)`` 取文本会把这些图**当成题注段**（89 个「图题」里 60 个是这种
    图对象，另有 29 个是图形段本身），判出来的东西跟版面流里的题注没有关系。
    """
    out: list[str] = []
    for t in p.iter(W + "t"):
        anc = t.getparent()
        skip = False
        while anc is not None and anc is not p:
            if etree.QName(anc).localname in _GRAPHIC_SCOPE:
                skip = True
                break
            anc = anc.getparent()
        if not skip:
            out.append(t.text or "")
    return "".join(out).strip()


def has_graphic(el: Any) -> bool:
    """块内是否含图对象（DrawingML / VML / OLE 三种容器都算）。"""
    return el is not None and el.tag == W + "p" and any(
        el.find(".//" + W + t) is not None for t in _DRAWING_TAGS)


def has_keep_next(el: Any) -> bool:
    ppr = el.find(W + "pPr") if el is not None else None
    return ppr is not None and ppr.find(W + "keepNext") is not None


def flow_blocks(document: Any) -> list[Any]:
    """正文流内块级元素（``w:p`` / ``w:tbl``），穿过 sdt/customXml 等包装层。

    「相邻」按这个序列算：题注与对象之间夹了空段/他段就是不相邻（A2 要 N-A）。
    文本框里的段落不在此列（它们是图对象的一部分，不在版面流里）。
    """
    body = document.find(W + "body")
    if body is None:
        return []

    def walk(el: Any):
        for ch in el:
            name = etree.QName(ch).localname
            if name in _FLOW_WRAPPERS:
                yield from walk(ch)
            elif name in ("p", "tbl"):
                yield ch

    return list(walk(body))


def _inside_graphic(p: Any) -> bool:
    """段落是否被图形/文本框包住（不在版面流内）。"""
    anc = p.getparent()
    while anc is not None:
        if etree.QName(anc).localname in ("txbxContent", "drawing", "pict"):
            return True
        anc = anc.getparent()
    return False


def _na_reason(el: Any) -> str:
    """不相邻时的原因（写进 note，不猜、不含糊）。"""
    if el is None:
        return "无"
    name = etree.QName(el).localname
    if name == "tbl":
        return "表格"
    if name != "p":
        return name
    if has_graphic(el):
        return "图对象段"
    return "空段" if not flow_text(el) else "他段"


def caption_binding_stats(doc: dict[str, Any]) -> dict[str, Any]:
    """图题/表题与对象的相邻性 + ``w:keepNext`` 实测统计（纯只读，A1/A2 口径）。

    判定（严格按对象位置，不看题注文字猜测）：

    * **图题**（``^图`` 开头）：对象应在**题前**——前一个流内块是图对象段才算相邻；
      该图段带 ``w:keepNext`` → PASS，不带 → FAIL（题与图之间没有绑定）。
    * **表题**（``^表`` 开头）：对象应在**题后**——后一个流内块是 ``w:tbl`` 才算相邻；
      表题段带 ``w:keepNext`` → PASS，不带 → FAIL。
    * 不相邻（隔空段/他段/题前无块）→ N-A，原因记进 ``na_reasons``——**不猜**。
    * 题注字写在图形文本框里（不在版面流）→ 不参与判定，只计数（``embedded``），
      因为块相邻这件事对它没有定义。
    """
    blocks = flow_blocks(doc["document"])
    stats: dict[str, Any] = {
        kind: {"pass": 0, "fail": 0, "na": 0, "fails": [], "na_reasons": Counter()}
        for kind in ("figure", "table")
    }
    for i, el in enumerate(blocks):
        if el.tag != W + "p":
            continue
        m = _CAPTION_RE.match(flow_text(el))
        if m is None:
            continue
        kind = "figure" if m.group(1) == "图" else "table"
        st = stats[kind]
        if kind == "figure":
            prev = blocks[i - 1] if i else None
            if has_graphic(prev):
                if has_keep_next(prev):
                    st["pass"] += 1
                else:
                    st["fail"] += 1
                    st["fails"].append(flow_text(el)[:36])
            else:
                st["na"] += 1
                st["na_reasons"][_na_reason(prev)] += 1
        else:
            nxt = blocks[i + 1] if i + 1 < len(blocks) else None
            if nxt is not None and nxt.tag == W + "tbl":
                if has_keep_next(el):
                    st["pass"] += 1
                else:
                    st["fail"] += 1
                    st["fails"].append(flow_text(el)[:36])
            else:
                st["na"] += 1
                st["na_reasons"][_na_reason(nxt)] += 1
    embedded = {"figure": 0, "table": 0}
    body = doc["document"].find(W + "body")
    for p in (body.iter(W + "p") if body is not None else ()):
        if not _inside_graphic(p):
            continue
        m = _CAPTION_RE.match(flow_text(p))
        if m is not None:
            embedded["figure" if m.group(1) == "图" else "table"] += 1
    stats["embedded"] = embedded
    for kind in ("figure", "table"):
        st = stats[kind]
        st["total"] = st["pass"] + st["fail"] + st["na"]
    return stats


def _caption_verdict(st: dict[str, Any]) -> str:
    """聚合判定：有 FAIL 就 FAIL；有 PASS 且无 FAIL 就 PASS；全部不可判 → N-A。"""
    if st["fail"]:
        return FAIL
    if st["pass"]:
        return PASS
    return NA


# ────────────────────────────── 判定工具 ──────────────────────────────


def _verdict(ok: bool) -> str:
    return PASS if ok else FAIL


def _size_ok(exp: float, act: float | None, tol: float = SIZE_TOLERANCE_PT) -> bool:
    return act is not None and abs(float(act) - float(exp)) < tol


def _hf_size_actual(st: Mapping[str, Any]) -> Any:
    """页眉/页脚字号的实测写法：N-A 时 ``actual`` 也不许是 None（要能看出为什么没得比）。"""
    if st.get("parts") == 0:
        return "无该部件"
    if st.get("size") is None:
        return "无 w:sz"
    return st.get("size")


def _spec_has_page(structure: Any) -> bool:
    """spec 是否声明了 PAGE 域（EZI-0012 A3 的「spec 侧」判据）。

    ``structure.py`` 把模板页眉/页脚按段拆成 ``run_sig``，页码域落成 ``{PAGE}``
    token；``front.has_page`` 是封面节的同口径标记。两者都没有 = spec 没说要有
    页码域——此时不许按「产物须有页码域」这条常量判 FAIL。
    """
    if not isinstance(structure, Mapping):
        return False
    for part in structure.get("header_footer_parts") or []:
        if not isinstance(part, Mapping):
            continue
        for para in part.get("paragraphs") or []:
            if not isinstance(para, Mapping):
                continue
            for tok in para.get("run_sig") or []:
                if str(tok).strip("{} \t").upper() == "PAGE":
                    return True
    front = structure.get("front")
    return bool(isinstance(front, Mapping) and front.get("has_page"))


def _jc(value: str | None) -> str:
    """未设置 jc 时 Word 按左对齐（start）渲染；别名归一。"""
    if not value:
        return "start"
    return _JC_ALIASES.get(value, value)


def _indent_chars(props: dict[str, Any]) -> float:
    """首行缩进折算成「字符」：优先 firstLineChars/100，否则 firstLine/(字号*20)。

    两种属性都没有 = 没有首行缩进 = 0 字符（返回 0.0 而不是 None，保证每项都有实测值）。
    """
    if props.get("first_line_chars") is not None:
        return props["first_line_chars"] / 100.0
    if props.get("first_line") is not None and props.get("sz"):
        return props["first_line"] / (float(props["sz"]) * 20.0)
    return 0.0


def _rule_expected(raw: Any) -> dict[str, Any] | None:
    """``rule.expected`` → dict；**只有 Mapping 才算规则体**。

    真实模板里出现过 ``expected=True``（bool）的情况，``dict(True)`` 会抛
    ``TypeError: 'bool' object is not iterable``，整组打分直接崩（EZI-0009 A7）。
    这里只做类型收窄：不是 Mapping 就返回 ``None``，调用方跳过该规则——
    既不 ``dict(bool)``，也不拿默认值凑一条能判的规则。
    """
    if isinstance(raw, Mapping):
        return dict(raw)
    return None


def _spec_parts(spec: Any) -> tuple[dict[str, Any], dict[str, dict[str, Any]]]:
    """兼容对象式 spec 与纯 dict（单测可注入），按 semantic_role 归并 rules（首个合法胜）。

    ``expected`` 不是 Mapping 的规则直接跳过（见 :func:`_rule_expected`）；同一角色
    若先出现一条非法规则、后出现一条合法规则，用后者——跳过不等于把角色判死。
    """
    if isinstance(spec, Mapping):
        structure = spec.get("structure") or {}
        raw_rules = spec.get("rules") or []
        rules: dict[str, dict[str, Any]] = {}
        for r in raw_rules:
            if not isinstance(r, Mapping):
                continue
            sel = r.get("selector") or {}
            role = sel.get("semantic_role") if isinstance(sel, Mapping) else None
            expected = _rule_expected(r.get("expected"))
            if role and expected is not None and role not in rules:
                rules[role] = expected
        return structure, rules
    structure = getattr(spec, "structure", None) or {}
    rules = {}
    for r in getattr(spec, "rules", None) or []:
        sel = getattr(r, "selector", None) or {}
        role = sel.get("semantic_role") if isinstance(sel, Mapping) else None
        expected = _rule_expected(getattr(r, "expected", None))
        if role and expected is not None and role not in rules:
            rules[role] = expected
    return structure, rules


def _compile_spec(template: Path) -> Any:
    from earlyz.spec.compilers import compile_format_spec

    return compile_format_spec("uploaded_template_docx", str(template), "thesis").spec


# ────────────────────────────── 主判定 ──────────────────────────────


def evaluate_layout_fidelity(product: str | Path, template: str | Path, *,
                             spec: Any = None,
                             source: str | Path | None = None) -> LayoutReport:
    """对 ``product`` 逐项判定是否按 ``template`` 排。只读，无副作用。

    ``spec`` 可传现成的 format spec（对象或 dict）；不传则对 ``template``
    调 ``compile_format_spec`` 现场编译。

    ``source`` 是**迁移语义 + 适用性**的依据（EZI-0008 / EZI-0009 / EZI-0012）：

    * 目录域：源有 TOC 域 → 产物必须有，否则 FAIL；源无 → 产物必须 0 个
      （有 → FAIL，EZI-0038 不生成目录；0 个 → PASS，EZI-0044）；不传 → N-A。
    * 标题三项：只取「outlineLvl 级别匹配 + 编号标题文本」的样本（排除表格内/后置
      节名/表单标签/封面大写行，EZI-0012 A7）；源文档该级编号样本数为 0 → N-A
      （产物没有该级是正确行为）；源有 → 按字号/字体/对齐正常判；不传且产物也没有
      样本 → N-A（不猜）。
    * 页眉页脚：只按模板部件里的 PAGE 域判（EZI-0012 A1-A4），不看
      ``headerReference``/``footerReference`` 存在性——模板页眉只有版头/编号文字时
      按 11-10 红线不该迁，判 N-A。页眉字号锚模板页眉 PAGE 段落（EZI-0044）；
      页眉引用(headerReference) 已废弃，不再输出（语义与页码域/页眉字号重叠）。
    * 分节数：模板单节 → 判产物必须 1 节（EZI-0034 A5）；模板多节 → N-A 的信息项，
      note 给 源/模板/产物 三个节数 + 缺证据说明（EZI-0044）。

    其它判定项（正文、页边距、分栏）不读 ``source``。
    """
    prod_path, tpl_path = Path(product), Path(template)
    src_path = Path(source) if source is not None else None
    prod = _load_docx(prod_path)
    tpl = _load_docx(tpl_path)
    src = _load_docx(src_path) if src_path is not None else None
    if spec is None:
        spec = _compile_spec(tpl_path)
    structure, rules = _spec_parts(spec)

    paras = document_paragraphs(prod)
    # EZI-0020：行距取样要排掉表单标签（`line_spacing_fact` 默认排），但正文四项
    # （字号/字体/缩进/对齐）暂时**沿用含标签的旧样本**：主字号的主峰阈值（60%）是按
    # 旧样本校准的，样本去掉标签会让本科论文 12pt（39/68=57%）从单峰掉进无主峰分支，
    # 候选集合由「12pt」变成「12pt/10.5pt」——判据强度变了，属 EZI-0020 的取样决策，
    # 本 issue 不动。两侧同口径：产物 bprops 与模板 tpl_bprops 都传同一个开关，不搞
    # 一侧新一侧旧。
    bprops = body_paragraphs(paras, exclude_form_labels=False)

    report = LayoutReport(product=str(prod_path), template=str(tpl_path),
                          source=str(src_path) if src_path else None)

    def add(name: str, expected: Any, actual: Any, verdict: str,
            category: str, note: str = "") -> None:
        report.items.append(LayoutItem(name, expected, actual, verdict, category, note))

    # ── 标题：一级 / 二级 / 三级 ──────────────────────────────
    # A7（EZI-0012）：样本只认「outlineLvl 级别匹配 且 文本是编号标题」，并排除
    # 表格内/后置节名/表单标签/封面大写行。旧口径取「全部 outlineLvl=0 段落的众数」，
    # 交大中期 3 段里 2 段是校名行（18pt）与表单字段（10.5pt），ACM 的众数被
    # ABSTRACT/REFERENCES（11pt）带偏——把对的判成错的。
    # A1（EZI-0009）：判定前先问「源文档有没有这个角色」，源无编号标题 → N-A。
    #
    # EZI-0016：期望值不再无条件取 spec 规则，而是「模板同级样本实测众数优先」——
    # spec 的标题规则来自编译期的样式/簇取样（`_expected_from` 的 min_cluster 可以是
    # 1 个样本），新疆模板三级 17/17 都是 center，spec 却写成 both，于是产物按模板排
    # 反而被判 FAIL。两道护栏保证只在两侧可比时才用模板实测：
    #   ① 样本量/主峰：模板该级样本 < 5 个或众数占比 < 60% → 没有主峰，退回 spec；
    #   ② 文本重合率：产物该级样本至少一半能在模板同级找到同文本对应，否则比的是
    #      两拨不相干的段落（新疆模板二级是「第X章」，产物二级是条款项）。
    prod_heads, prod_head_excl = heading_samples_by_level(prod)
    tpl_heads, _ = heading_samples_by_level(tpl)
    if src is not None:
        src_heads, src_head_excl = heading_samples_by_level(src)
    else:
        src_heads, src_head_excl = {}, Counter()
    for level in (1, 2, 3):
        role = f"heading{level}"
        rule = rules.get(role) or {}
        hprops = prod_heads.get(level) or []
        where = f"{'一二三'[level - 1]}级标题"
        if not rule:
            add(f"{where}字号(pt)", None, None, NA, CATEGORY_TITLES,
                f"模板 spec 无 {role} 规则，无法判定")
            continue

        sz = ea = jc = None
        ea_n = ea_tot = jc_n = jc_tot = 0
        if hprops:
            # 字号的全样本众数只给「无该级样本」分支填 actual 用；实际字号判定在下面按
            # 可比样本重算（EZI-0021）。
            sz, _, _ = _mode(p["sz"] for p in hprops)
            ea, ea_n, ea_tot = _mode(p["east_asia"] for p in hprops)
            jc, jc_n, jc_tot = _mode(_jc(p["jc"]) for p in hprops)
        src_n = len(src_heads.get(level) or []) if src is not None else None
        src_txt = f"{src_n} 个编号{where}" if src_n is not None else "未提供源文档"
        missing = "无编号该级标题样本"

        tpl_rows = tpl_heads.get(level) or []

        # 适用性不成立（或无法确认）：三项一律 N-A，expected 留空（不出现常量期望）
        if src_n == 0 or (src_n is None and not hprops):
            note = (f"源文档无编号{where}，无法比较（编号样本 0；"
                    f"排除：{_excl_text(src_head_excl)}；EZI-0012 A7）"
                    if src_n == 0 else
                    "未提供源文档，且产物无编号该级标题（排除："
                    f"{_excl_text(prod_head_excl)}）——无法确认该级是否适用，不判")
            for name, key, actual in (
                (f"{where}字号(pt)", "font_size_pt", sz if hprops else missing),
                (f"{where}中文字体", "font_east_asia", ea if hprops else missing),
                (f"{where}对齐(jc)", "jc", jc if hprops else missing),
            ):
                if rule.get(key) is not None:
                    add(name, None, actual, NA, CATEGORY_TITLES, note)
            continue

        # 源有该级**真**编号标题、产物排除后没有 = 迁移丢失（EZI-0076 定案：
        # 说明行类噪声已被池排除前置过滤，走到这里的样本全是真标题——真内容
        # 丢了就该红，一级二级三级同口径，废除「二/三级不判失败待口径定」旗）。
        if not hprops:
            verdict = FAIL
            note = (f"源文档有编号{where}（{src_n} 个），产物排除后无样本 → 迁移丢失")
            for name, key, exp in (
                (f"{where}字号(pt)", "font_size_pt",
                 float(rule["font_size_pt"]) if rule.get("font_size_pt") is not None else None),
                (f"{where}中文字体", "font_east_asia", rule.get("font_east_asia")),
                (f"{where}对齐(jc)", "jc", _jc(rule["jc"]) if rule.get("jc") else None),
            ):
                if rule.get(key) is not None:
                    add(name, exp, missing, verdict, CATEGORY_TITLES, note)
            continue

        # 正常判：源有该级、产物也有编号样本
        # EZI-0021 + EZI-0025：字号、字体、对齐三项**共用同一套样本**——两侧不是同一批
        # 标题时（重合 < HEADING_OVERLAP_MIN），产物全体的众数同样不能当该级实测值，收窄
        # 到「与模板同级同文」的那批。不收窄时 cmp_txt 为空串，note 与旧口径逐字相同。
        # 收窄后可比样本为空 → 三项 N-A，**不退回全量样本**（A3）。
        cmp_rows, _cmp_dropped, cmp_txt, narrowed = _comparable_heading_rows(tpl_rows, hprops)
        if narrowed and not cmp_rows:
            note = (f"编号{where}样本 {len(hprops)} 个；{cmp_txt.lstrip('；')}"
                    f"；该级三项不判（EZI-0025 A3：不退回全量样本）")
            for name, key in ((f"{where}字号(pt)", "font_size_pt"),
                              (f"{where}中文字体", "font_east_asia"),
                              (f"{where}对齐(jc)", "jc")):
                if rule.get(key) is not None:
                    add(name, None, "无可比样本", NA, CATEGORY_TITLES, note)
            continue
        judge_rows = cmp_rows if narrowed else hprops
        c_sz, c_n, c_tot = _mode(p["sz"] for p in judge_rows)
        c_ea, c_ea_n, c_ea_tot = _mode(p["east_asia"] for p in judge_rows)
        c_jc, c_jc_n, c_jc_tot = _mode(_jc(p["jc"]) for p in judge_rows)
        first = judge_rows[0]["text"][:16]
        sample = (f"编号{where}样本 {len(hprops)} 个（首个「{first}」）；"
                  f"排除 {_excl_text(prod_head_excl)}")
        tail = f"；源 {src_txt}" if src_n is not None else ""
        if rule.get("font_size_pt") is not None:
            exp_sz, cand_sz, src_sz = _heading_expectation(
                rule, "font_size_pt", tpl_rows, hprops, where=where, unit="pt")
            verdict, na_why = _heading_prop("font_size_pt", exp_sz, c_sz, tot=c_tot,
                                            narrowed=narrowed, cmp_n=len(cmp_rows),
                                            candidates=cand_sz)
            add(f"{where}字号(pt)", None if verdict == NA else exp_sz, c_sz,
                verdict, CATEGORY_TITLES,
                f"{sample}{cmp_txt}，字号众数 {c_n}/{c_tot}；{src_sz}{na_why}{tail}")
        if rule.get("font_east_asia"):
            exp_ea, _cand_ea, src_ea = _heading_expectation(
                rule, "font_east_asia", tpl_rows, hprops, where=where)
            verdict, na_why = _heading_prop("font_east_asia", exp_ea, c_ea,
                                            tot=c_ea_tot,
                                            narrowed=narrowed, cmp_n=len(cmp_rows))
            add(f"{where}中文字体", None if verdict == NA else exp_ea, c_ea,
                verdict, CATEGORY_TITLES,
                f"{sample}{cmp_txt}，eastAsia 众数 {c_ea_n}/{c_ea_tot}；{src_ea}{na_why}{tail}")
        if rule.get("jc"):
            exp_jc, _cand_jc, src_jc = _heading_expectation(
                rule, "jc", tpl_rows, hprops, where=where)
            exp_jc = _jc(str(exp_jc)) if exp_jc is not None else None
            verdict, na_why = _heading_prop("jc", exp_jc, c_jc, tot=c_jc_tot,
                                            narrowed=narrowed, cmp_n=len(cmp_rows))
            add(f"{where}对齐(jc)", None if verdict == NA else exp_jc, c_jc,
                verdict, CATEGORY_TITLES,
                f"{sample}{cmp_txt}，jc 众数 {c_jc_n}/{c_jc_tot}；{src_jc}{na_why}{tail}")

    # ── 正文：字号 / 字体 / 缩进 / 行距 / 对齐 ────────────────
    # EZI-0071：产物正文样本不足（<5 段）→ 本块四项 N-A——排除链（说明段/refs
    # 区）正确净化后「真无正文」的产物（温州文献综述实锤：源=封面+题名+refs
    # 示例，池 0 段），act=None 判 FAIL 是拿噪声当判据；与模板侧「<5 段实测
    # 不可用」（EZI-0022）同一规矩对称。
    _act_pool_na = len(bprops) < 5
    _act_na_note = f"产物正文段落 {len(bprops)} 个（<5 实测不可用）→ N-A"
    body_rule = rules.get("body") or {}
    if body_rule.get("font_size_pt") is not None:
        # EZI-0016：期望取模板正文段落的实测众数（同一取样函数、同一排除口径），
        # 不再用 spec 注解推出来的值——本科模板正文 12pt×48/77（62%）被 spec 写成
        # 14pt（注解「正文四号」），产物照模板排 12pt 反而判 FAIL。
        # EZI-0022：无主峰时**不再**退回 spec 注解当期望（交大中期 50%、开题 47% 都
        # 退回了 spec 16pt，产物恰好 16pt 才 PASS——运气）。改判「双峰区间」：候选 =
        # 前二高频，产物落在任一即 PASS；前二合计也 <60% 则 N-A。只有样本不足
        # （<5 段，实测不可用）才退回规范注解，并在 note 里显式标注出处。
        tpl_bprops = body_text_paragraphs(document_paragraphs(tpl),
                                          exclude_form_labels=False)
        exp_bsz, cand_bsz, src_bsz, forced = _body_size_expectation(
            tpl_bprops, body_rule["font_size_pt"])
        sz, n, tot = _mode(p["sz"] for p in bprops)
        note = (f"正文段落 {len(bprops)} 个（排除表格/标题/题注/目录/空段/手写目录条目），字号众数 "
                f"{n}/{tot}；{src_bsz}；模板正文分布 "
                f"{_sz_dist_txt(tpl_bprops, 'sz')}")
        if _act_pool_na:
            add("正文主字号(pt)", None, sz, NA, CATEGORY_BODY, note + "；" + _act_na_note)
        elif forced == NA:
            add("正文主字号(pt)", None, sz, NA, CATEGORY_BODY, note)
        else:
            add("正文主字号(pt)", exp_bsz, sz,
                _verdict(_size_ok_any(cand_bsz or (exp_bsz,), sz)), CATEGORY_BODY, note)
    else:
        add("正文主字号(pt)", None, None, NA, CATEGORY_BODY, "模板 spec 无 body.font_size_pt")

    if body_rule.get("font_east_asia"):
        ea, n, tot = _mode(p["east_asia"] for p in bprops)
        if _act_pool_na:
            add("正文中文字体", None, ea, NA, CATEGORY_BODY, _act_na_note)
        else:
            add("正文中文字体", body_rule["font_east_asia"], ea,
                _verdict(ea is not None and ea == body_rule["font_east_asia"]), CATEGORY_BODY,
                f"eastAsia 众数 {n}/{tot}")

    if body_rule.get("ind_first_line_chars") is not None:
        exp_chars = float(body_rule["ind_first_line_chars"]) / 100.0
        ch, n, tot = _mode(round(_indent_chars(p), 2) for p in bprops)
        ok = ch is not None and abs(ch - exp_chars) <= INDENT_TOLERANCE_CHARS / 100.0 + 1e-9
        if _act_pool_na:
            add("正文首行缩进(字符)", None, ch, NA, CATEGORY_BODY, _act_na_note)
        else:
            add("正文首行缩进(字符)", exp_chars, ch, _verdict(ok), CATEGORY_BODY,
                f"firstLineChars/100 众数 {n}/{tot}；容差 ±{INDENT_TOLERANCE_CHARS / 100:.2f} 字符")

    # 行距（EZI-0015 A3）：两侧同口径——都用 ``line_spacing_fact`` 量正文本体段落
    # （排除表格/标题/TOC），众数只算设了 ``@w:line`` 的段落。旧实现拿
    # ``structure.rules.body.line`` 当期望、产物侧把 None 也算进众数：IEEE 模板正文
    # 段落 79 段 240/exact（经 PARA 样式继承），产物 103 段 240/exact，却被判
    # 「模板 240/exact vs 产物 None」FAIL；模板真的没设行距时又凭空给出期望值。
    tpl_line = line_spacing_fact(tpl)
    prod_line = line_spacing_fact(prod)
    tpl_line_txt = _line_value_txt(tpl_line["line"], tpl_line["rule"])
    prod_line_txt = _line_value_txt(prod_line["line"], prod_line["rule"])
    tpl_dist_txt = _line_dist_txt(tpl_line["dist"], tpl_line["total"])
    prod_dist_txt = _line_dist_txt(prod_line["dist"], prod_line["total"])
    if tpl_line["line"] is None:
        add("正文行距(line/lineRule)", None,
            prod_line_txt or f"未设（正文 {prod_line['total']} 段）", NA, CATEGORY_BODY,
            f"模板正文未设固定行距（模板正文 {tpl_line['total']} 段，w:spacing/@w:line "
            f"与默认段落样式均无设值）→ 无期望值，不判；产物分布：{prod_dist_txt}")
    elif prod_line["total"] < 5:
        # EZI-0071：产物正文样本不足 → N-A（与 bprops 四项同一规矩）
        add("正文行距(line/lineRule)", None,
            prod_line_txt or f"未设（正文 {prod_line['total']} 段）", NA, CATEGORY_BODY,
            f"产物正文段落 {prod_line['total']} 个（<5 实测不可用）→ N-A；产物分布：{prod_dist_txt}")
    else:
        ok = (prod_line["line"] == tpl_line["line"]
              and prod_line["rule"] == tpl_line["rule"])
        src_txt = ("（期望取自模板" + tpl_line["source"] +
                   (f"，{tpl_line['n']}/{tpl_line['set']} 段" if tpl_line["n"] else "") + "）")
        add("正文行距(line/lineRule)", f"{tpl_line_txt}{src_txt}",
            prod_line_txt or f"未设（正文 {prod_line['total']} 段）",
            _verdict(ok), CATEGORY_BODY,
            f"两侧同口径（正文段落，排除表格/标题/TOC/表单标签/手写目录条目；众数只算设值段落，line 单位 "
            f"1/20 pt）：模板 {tpl_dist_txt}；产物 {prod_dist_txt}")

    if body_rule.get("jc"):
        jc, n, tot = _mode(_jc(p["jc"]) for p in bprops)
        if _act_pool_na:
            add("正文对齐(jc)", None, jc, NA, CATEGORY_BODY, _act_na_note)
        else:
            add("正文对齐(jc)", _jc(body_rule["jc"]), jc,
                _verdict(jc is not None and jc == _jc(body_rule["jc"])), CATEGORY_BODY,
                f"jc 众数 {n}/{tot}；未设置按 start（左对齐）计")

    # ── 页边距（EZI-0016 逐节 / EZI-0027 节数不等时按角色配对，容差 20 twips）──
    # 旧判法：期望 = spec.structure.sections[0].pg_mar，实测 = 产物 body 末节。
    # 本科模板 8 节里节0~2 = 1417/1701/1134/1701（封面/摘要）、节3~7 = 1440/1800/1440/1800
    # （正文）—— 拿模板首节比产物末节，两侧根本不是同一节，报出来的
    # 「1417/1701/1134/1701 vs 1440/1800/1440/1800」既说不清差异在哪，也说不清哪边错。
    # EZI-0027：产物多一个「声明」节时，按索引比会把产物节3（ABSTRACT，1417）与模板
    # 节3（目录，1440）对成一对 → 边距全对也判 FAIL；节数不等时改按角色配对。
    tpl_margins = _section_margins(tpl)
    prod_margins = _section_margins(prod)
    if not tpl_margins or not any(tpl_margins):
        add("页边距(top/right/bottom/left)", None, None, NA, CATEGORY_MARGIN,
            "模板无 sectPr/pgMar——无期望值，不判")
    else:
        tpl_roles = _section_role_seq(tpl)
        prod_roles = _section_role_seq(prod)
        if tpl_roles is None or prod_roles is None:
            n_cmp = min(len(tpl_margins), len(prod_margins))
            pairs: list[tuple[int, int]] = [(i, i) for i in range(n_cmp)]
            mode, pair_note = "index-fallback", "角色识别失败（读不出逐节段落结构）"
        else:
            pairs, mode, pair_note = _pair_margin_sections(
                tpl_roles, prod_roles, tpl_margins=tpl_margins)
        # EZI-0074：模板同角色页边距**多态**（118 节合同实测：目录节里混
        # 1680 变体、正文节散布横向附件页）——拿任一节的值当唯一期望会把
        # 排对形态判成 FAIL。与 EZI-0022 双峰区间同一原则：该角色出现 >1
        # 种页边距形态时，产物落在**任一形态**即达标（配对节仍用于报告）。
        def _form(m):
            return tuple(sorted((k, m.get(k)) for k in ("top", "right", "bottom", "left")))
        _role_forms: dict[str, set] = {}
        for ti, roles in enumerate(tpl_roles):
            if tpl_margins[ti] is None:
                continue
            for r in roles:
                _role_forms.setdefault(r, set()).add(_form(tpl_margins[ti]))

        def _pair_ok(pi: int, ti: int) -> bool:
            pm = prod_margins[pi]
            tm = tpl_margins[ti]
            forms = set()
            for r in tpl_roles[ti]:
                forms |= _role_forms.get(r, set())
            if len(forms) > 1:
                pf = _form(pm)
                if pf is not None and all(v is not None for _k, v in pf):
                    return pf in forms
            return not _margin_diff(tm, pm)

        bad = [(pi, ti) for pi, ti in pairs
               if not _pair_ok(pi, ti)]
        exp_txt = f"{len(tpl_margins)} 节：{_margin_seq_txt(tpl_margins)}"
        act_txt = (f"{len(prod_margins)} 节：{_margin_seq_txt(prod_margins)}"
                   if prod_margins else "无 sectPr/pgMar")
        if mode == "role":
            bits = [f"两侧节数不同（{len(tpl_margins)} vs {len(prod_margins)}）→ "
                    f"按角色配对，容差 ±{MARGIN_TOLERANCE_TWIPS} twips",
                    "配对（产物节→模板节）："
                    + _margin_pair_txt(pairs, prod_roles, tpl_roles)]
            if pair_note:
                bits.append(pair_note)
        else:
            bits = [f"按索引逐节对齐比前 {len(pairs)} 节（模板 {len(tpl_margins)} 节 / 产物 "
                    f"{len(prod_margins)} 节），容差 ±{MARGIN_TOLERANCE_TWIPS} twips"]
        if bad:
            if mode == "role":
                shown = "；".join(
                    f"节{pi}{_role_txt(prod_roles[pi])}→模板节{ti}{_role_txt(tpl_roles[ti])} "
                    f"模板 {_margin_txt(_margin_sig(tpl_margins[ti]))} vs 产物 "
                    f"{_margin_txt(_margin_sig(prod_margins[pi]))}" for pi, ti in bad[:3])
            else:
                shown = "；".join(
                    f"节{pi} 模板 {_margin_txt(_margin_sig(tpl_margins[ti]))} vs 产物 "
                    f"{_margin_txt(_margin_sig(prod_margins[pi]))}" for pi, ti in bad[:3])
            bits.append(f"超差节 {_idx_ranges([pi for pi, _ in bad])}（{len(bad)} 节）：{shown}"
                        + (f"；其余 {len(bad) - 3} 节同类" if len(bad) > 3 else ""))
        if len(tpl_margins) != len(prod_margins) and mode != "role":
            fallback_txt = f"{pair_note} → 退回逐节索引口径；" if mode == "index-fallback" else ""
            bits.append(f"两侧节数不同（{len(tpl_margins)} vs {len(prod_margins)}）："
                        f"{fallback_txt}"
                        "多出的节没有同序号可比，节骨架本身是信息项"
                        "（见「分节数(sectPr)」），此处只比对齐得到的节")
        if not prod_margins:
            bits.append("产物没有任何 sectPr → 拿不到 pgMar，判 FAIL")
        add("页边距(top/right/bottom/left)", exp_txt, act_txt,
            _verdict(bool(prod_margins) and not bad), CATEGORY_MARGIN, "；".join(bits))

    # ── 目录域（迁移语义：源有则产物必有；源无则产物不许生成）──────────
    toc_hits = _toc_instrs(prod)
    prod_toc = f"存在（{len(toc_hits)} 处）" if toc_hits else "未找到"
    src_hits = _toc_instrs(src) if src is not None else None
    if src_hits is None:
        # 不给源文档就没有期望值——宁可 N-A 也不按模板/常量猜
        add("目录域(TOC)", None, prod_toc, NA, CATEGORY_TOC,
            "未提供源文档，无法判定（缺源文档证据：迁移语义的期望值只能来自源）")
    elif not src_hits:
        # EZI-0044：源无 TOC 不是「不可判」——「不生成目录」本身就是可判约束
        # （EZI-0038 修的就是源无 TOC 时仍生成目录）。产物 0 个 = 正确行为 = PASS；
        # 产物有 TOC 域 = 凭空生成，判 FAIL。旧口径整支 N-A，三对体检样本上把这一
        # 项做成了 100% 死指标，等于放过了「不该生成却生成」。
        add("目录域(TOC)", "0 个（源无目录域，不许生成）", prod_toc,
            _verdict(not toc_hits), CATEGORY_TOC,
            "迁移语义（EZI-0008/EZI-0038）：源文档无目录域 → 产物不得生成目录域；"
            + ("产物 0 个 → PASS" if not toc_hits
               else "产物生成了目录域 → FAIL（只迁移，不生成）"))
    else:
        add("目录域(TOC)", f"存在（源文档有 {len(src_hits)} 处，须迁移）", prod_toc,
            _verdict(bool(toc_hits)), CATEGORY_TOC,
            f"迁移语义：源文档有目录域（首条 {src_hits[0][:40]}），产物"
            + ("已迁移" if toc_hits else "缺目录域 → 必须 FAIL"))

    # ── 页眉页脚（EZI-0012 A1-A4：只按 PAGE 域判，不看 headerReference 存在性）──
    # 2026-11-10 页眉红线：不写模板页眉文字、不清用户占位文字，只继承 PAGE 域链 +
    # 段落格式。所以模板页眉里只有版头/编号文字时，产物**不该**有这个页眉——拿
    # 「模板有 headerReference」当期望，正好和红线对撞（公文/合同 5 例被判 FAIL）。
    #
    # EZI-0044：「页眉引用(headerReference)」**明确废弃**，不再输出该项。
    # 理由：① 红线禁止按 headerReference 存在性判 FAIL；② 唯一有意义的替代语义
    # （模板页眉有 PAGE → 产物页眉须有 PAGE）已被「页码域(PAGE)」与「页眉字号(pt)」
    # 完全覆盖，再留一项只能重复计（三对体检样本上它 100% N-A，就是死指标）。
    # 页脚引用保留：页脚没有「不迁文字」的红线，页脚部件本身就是页码的载体。
    n_sect = len(_sections(prod))
    ftr_refs = _hf_refs(prod, "footer")
    n_ftr_sect = sum(1 for s in _sections(prod) if s.findall(W + "footerReference"))
    prod_h, prod_f = _hf_stats(prod, "header"), _hf_stats(prod, "footer")
    tpl_h, tpl_f = _hf_stats(tpl, "header"), _hf_stats(tpl, "footer")

    name = "页脚引用(footerReference)"
    actual = f"{n_ftr_sect}/{n_sect} 节（{ftr_refs} 处）"
    if tpl_f["parts"] == 0:
        add(name, None, actual, NA, CATEGORY_HF,
            "模板无页脚部件 → 无期望值，不判页脚引用（EZI-0012 A2）")
    elif not tpl_f["has_page"]:
        shown = (f"仅文字：{tpl_f['text'][:24]}" if tpl_f["has_text"] else "页脚为空")
        add(name, None, actual, NA, CATEGORY_HF,
            f"模板页脚部件无 PAGE 域（{shown}）→ 页脚文字不迁，不判页脚引用"
            f"（EZI-0012 A2）")
    else:
        ok = prod_f["has_page"]
        add(name, "页脚含 PAGE 域", actual + ("，含 PAGE 域" if ok else "，无 PAGE 域"),
            _verdict(ok), CATEGORY_HF,
            "模板页脚部件含 PAGE 域 → 产物须有页脚且含 PAGE 域（EZI-0012 A2）")

    # A3：页码域。期望只来自「模板或 spec 有 PAGE」；两者均无 → N-A。
    # 旧口径在「模板与 spec 均无 PAGE」时仍写「产物须有页码域」判 FAIL（note 自己
    # 都写明了没有期望值），和 EZI-0008 的目录域是同一个病。
    spec_page = _spec_has_page(structure)
    tpl_slots = [k for k, s in (("页眉", tpl_h), ("页脚", tpl_f)) if s["has_page"]]
    prod_slots = [k for k, s in (("页眉", prod_h), ("页脚", prod_f)) if s["has_page"]]
    prod_slot_txt = "/".join(prod_slots) or "无"
    if tpl_slots or spec_page:
        bits = []
        if tpl_slots:
            bits.append(f"模板页码在{'/'.join(tpl_slots)}")
        if spec_page:
            bits.append("spec 声明 PAGE 域")
        add("页码域(PAGE)", "存在（" + "；".join(bits) + "）", prod_slot_txt,
            _verdict(bool(prod_slots)), CATEGORY_HF,
            "模板或 spec 有 PAGE 域 → 产物须有（EZI-0012 A3）；同侧缺失已由页脚引用/页眉字号判")
    else:
        add("页码域(PAGE)", None, prod_slot_txt, NA, CATEGORY_HF,
            "模板与 spec 均无 PAGE 域 → 不判（删掉「产物须有页码域」这条常量判据，"
            "EZI-0012 A3）")

    # A4（EZI-0044 改口径）：页眉字号只锚**模板页眉 PAGE 段落**的有效字号。
    # 旧口径拿模板页眉全部件 w:sz 众数当期望——页眉主体是版头/编号文字（矿大 37 个
    # 部件、北航 24 个），按 11-10 红线这些文字不迁，它们的字号与产物没有继承关系；
    # 又因为引擎只写 \xa0 占位符（无 w:t 文本）而连判三次 N-A，做成死指标。
    # 现在：模板页眉没有 PAGE 域 = 没有可继承的字号基准 → N-A（带 note 说清缺哪类
    # 证据）；两侧都有 PAGE 域才比，基准来自模板 PAGE 段落（run/段落标记/样式链）。
    tpl_h_page = _hf_page_size(tpl, "header")
    prod_h_page = _hf_page_size(prod, "header")
    hdr_size_note = (f"期望取模板页眉 PAGE 段落有效字号 {tpl_h_page['size_n']}/"
                     f"{tpl_h_page['size_total']}")
    if tpl_h["parts"] == 0:
        add("页眉字号(pt)", None, _hf_size_actual(prod_h), NA, CATEGORY_HF,
            "模板无页眉部件 → 无 PAGE 段落可继承，无期望值")
    elif not tpl_h["has_page"]:
        shown = (f"仅文字：{tpl_h['text'][:24]}" if tpl_h["has_text"] else "页眉为空")
        tpl_sz = f"（页眉 w:sz 众数 {tpl_h['size']}pt，属模板文字）" if tpl_h["size"] else ""
        add("页眉字号(pt)", None, _hf_size_actual(prod_h), NA, CATEGORY_HF,
            f"模板页眉无 PAGE 域（{shown}）{tpl_sz} → 按 {_HEADER_REDLINE}，"
            "页眉文字字号不迁，没有可继承的基准（缺「模板页眉 PAGE 段落」这类证据）")
    elif tpl_h_page["size"] is None:
        add("页眉字号(pt)", None, _hf_size_actual(prod_h), NA, CATEGORY_HF,
            "模板页眉有 PAGE 域，但 PAGE 段落及样式链都没有 w:sz "
            "（run / 段落标记 / 段落样式 / 文档默认均未设）→ 无字号证据，不判")
    elif prod_h["parts"] == 0:
        add("页眉字号(pt)", None, "无页眉部件", NA, CATEGORY_HF,
            f"产物无页眉部件 → 无 PAGE 段落可比（缺页眉 PAGE 链已由页码域(PAGE)判，"
            f"不重复计）；{hdr_size_note}")
    elif prod_h_page["size"] is None:
        add("页眉字号(pt)", None, "页眉无 PAGE 域", NA, CATEGORY_HF,
            f"产物页眉无 PAGE 域 → 无页码链、无字号可比（缺 PAGE 链由页码域判，"
            f"不重复计）；{hdr_size_note}")
    else:
        add("页眉字号(pt)", tpl_h_page["size"], prod_h_page["size"],
            _verdict(_size_ok(tpl_h_page["size"], prod_h_page["size"], HF_SIZE_EPS_PT)),
            CATEGORY_HF,
            f"{hdr_size_note}（证据出处：{tpl_h_page['how'] or '—'}；"
            f"产物 {prod_h_page['size_n']}/{prod_h_page['size_total']}，"
            f"出处：{prod_h_page['how'] or '—'}）；容差 ±{HF_SIZE_EPS_PT} pt；"
            "口径：只继承 PAGE 域链 + 段落格式，页眉文字字号不参与")

    # 页脚页码字号：模板页脚没有字号 / 产物没有页脚 / 产物页脚没有 PAGE 域时都
    # 无「页码字号」可比——缺页码域这件事已经由页脚引用与页码域各自判，不重复计。
    # EZI-0064：页脚页码字号锚点重构——期望改取**页脚 PAGE 段落的样式链有效字号**
    # （与页眉字号 EZI-0044 同构：run sz → 同段文字 run → 段落标记/样式/docDefaults）。
    # 旧口径拿「页脚全部件 w:sz 直写众数」当期望：模板页脚 PAGE run 无直写、
    # 字号走样式链继承时（矿大/北科/国科大等 16 案例）众数为 None → N-A，
    # 页码字号整块无担保（EZI-0063 盲区审计实锤）。旧口径还有反向病：页脚里
    # 混排的**文字**字号被当页码字号比——那是页脚文字的字号，与页码无继承关系。
    # 两侧同用 PAGE 段落有效口径，一处证据都没有就 N-A，不拿直写众数凑。
    tpl_f_page = _hf_page_size(tpl, "footer")
    prod_f_page = _hf_page_size(prod, "footer")
    ftr_size_note = (f"期望取模板页脚 PAGE 段落有效字号 {tpl_f_page['size_n']}/"
                     f"{tpl_f_page['size_total']}（出处：{tpl_f_page['how'] or '—'}）")
    if not tpl_f["has_page"]:
        add("页脚页码字号(pt)", None, _hf_size_actual(prod_f), NA, CATEGORY_HF,
            f"模板页脚无 PAGE 域 → 模板页脚字号 {tpl_f['size']}pt 是页脚文字的字号，"
            "不是页码字号，不判（EZI-0012 A3 同源）")
    elif tpl_f_page["size"] is None:
        add("页脚页码字号(pt)", None, _hf_size_actual(prod_f), NA, CATEGORY_HF,
            "模板页脚有 PAGE 域，但 PAGE 段落及样式链均无 w:sz"
            "（run / 同段文字 / 段落标记 / 样式 / 文档默认均未设）→ 无字号证据，不判"
            f"（旧口径会把页脚文字直写众数 {tpl_f['size']}pt 错当页码字号，EZI-0064 弃用）")
    elif prod_f["parts"] == 0:
        add("页脚页码字号(pt)", None, "无页脚部件", NA, CATEGORY_HF,
            "产物无页脚部件 → 不判页码字号；模板页脚字号只对模板自己的文字有效")
    elif not prod_f["has_page"]:
        add("页脚页码字号(pt)", None, "页脚无 PAGE 域", NA, CATEGORY_HF,
            "产物页脚无 PAGE 域 → 无页码字号可比（缺页码域由页脚引用/页码域判，不重复计）")
    elif prod_f_page["size"] is None:
        add("页脚页码字号(pt)", None, _hf_size_actual(prod_f), NA, CATEGORY_HF,
            "产物页脚 PAGE 段落及样式链均无 w:sz → 产物侧无字号证据，不判"
            f"；{ftr_size_note}")
    else:
        add("页脚页码字号(pt)", tpl_f_page["size"], prod_f_page["size"],
            _verdict(_size_ok(tpl_f_page["size"], prod_f_page["size"], HF_SIZE_EPS_PT)), CATEGORY_HF,
            f"{ftr_size_note}；产物 {prod_f_page['size_n']}/{prod_f_page['size_total']}"
            f"（出处：{prod_f_page['how'] or '—'}）；容差 ±{HF_SIZE_EPS_PT} pt；"
            "口径：页脚 PAGE 段落有效字号（样式链），页脚文字字号不参与（EZI-0064）")

    # ── 题注绑定（EZI-0011）：图题绑图段、表题绑表 ─────────────
    # 中文规范：图题在图**下方**、表题在表**上方**；``w:keepNext`` 是单向的（绑
    # 自己 + 后面的内容），所以方向必须按对象位置分——图题要绑的是**前面的图
    # 段落**，表题要绑的是**题注自己**。这里只判 OOXML 实测，不猜：
    # 题与对象不相邻一律 N-A（隔空段/他段/题注内嵌图对象文本框）。
    cap = caption_binding_stats(prod)
    for kind, label, expect in (
        ("figure", "图题绑定", "图段带 keepNext 且图紧邻题前"),
        ("table", "表题绑定", "表题带 keepNext 且表紧邻题后"),
    ):
        st = cap[kind]
        actual = (f"PASS {st['pass']} / FAIL {st['fail']} / N-A {st['na']}"
                  f"（{label[:1]}题 {st['total']} 个）")
        verdict = _caption_verdict(st)
        bits = ["口径：对象位置定方向（题前图 / 题后表）"]
        if st["fail"]:
            bits.append("FAIL 例=" + "；".join(st["fails"][:3]))
        if st["na"]:
            why = "、".join(f"{k} {v}" for k, v in sorted(st["na_reasons"].items()))
            bits.append(f"N-A 原因：{why}")
        if st["pass"] + st["fail"] == 0:
            emb = cap["embedded"][kind]
            bits.append(
                f"无可判定题注：流内{label[:1]}题为 0"
                + (f"，另有 {emb} 处题注字写在图对象文本框里（不在版面流，按 A2 不猜）"
                   if emb else ""))
        add(label, expect if verdict != NA else None, actual, verdict,
            CATEGORY_CAPTION, "；".join(bits))

    # ── 分节 / 分栏 ───────────────────────────────────────────
    # EZI-0044 定案：分节数只有一支有模板证据——**模板单节**。模板 1 节 = 模板自己
    # 没有任何属性变化点，产物多切一节就是无中生有（EZI-0034 A5「单节模板一律不切」），
    # 这一支判 PASS/FAIL。模板多节时判不了：产物节数 = 产物**属性变化点数**
    # （cols / pgMar / pgSz / 页码体系），而变化点由**源内容角色**决定，不由模板节数
    # 决定（EZI-0034：本科模板 8 节 → 产物 3 节、IEEE 模板 2 节 → 产物 2 节、
    # CJC 模板 8 节 → 产物 3 节）；模板节数还自带 1-twip 噪声与逐章空节
    # （北航模板 16 节几何变化点去噪后 4 个、产物 5 节），拿它当期望必然误判。
    # 缺的证据是「源内容块 → 模板节」的角色映射（在引擎 plan_sections 里，判定层
    # 不 import 会改文档的引擎模块），所以多节模板这一支结构性 N-A，note 写清缺什么。
    n_tpl_sect = len(_sections(tpl))
    src_sect_txt = str(len(_sections(src))) if src is not None else "?（未提供源文档）"
    sect_tail = f"源={src_sect_txt} 模板={n_tpl_sect} 产物={n_sect}"
    if n_tpl_sect == 1:
        add("分节数(sectPr)", 1, n_sect, _verdict(n_sect == 1), CATEGORY_SECTION,
            f"模板单节（无 cols/pgMar/pgSz/页码体系变化点）→ 产物必须 1 节"
            f"（EZI-0034 A5：单节模板一律不切，多切=发明结构）；{sect_tail}")
    else:
        add("分节数(sectPr)", None, n_sect, NA, CATEGORY_SECTION,
            f"信息项，不参与打分：{sect_tail}；产物节数 = 产物属性变化点数"
            "（cols/pgMar/pgSz/页码体系），变化点由源内容角色决定、不由模板节数决定"
            "（EZI-0034）→ 模板多节时缺「源内容块→模板节」映射证据，不判；"
            "模板单节时才判（产物必须 1 节）")

    # 正文分栏（EZI-0015 A2）：**逐节**对比，两侧都取 body 末节（最后一个 sectPr）。
    # 旧判法拿 ``structure.body_cols``（「第一个非首节 ≥2 栏」）比产物末节：ACM 模板
    # 与产物 cols 序列都是 [1,3,1,2]，索引 1=3 栏作者区、末节=2 栏正文，被判
    # 「模板 3 产物 2」FAIL——两侧比的不是同一节。逐节序列写进 note，节数不同时
    # 说明取舍（两侧一律取末节）。
    tpl_cols = _section_cols_fact(tpl)
    prod_cols = _section_cols_fact(prod)
    cols_note = (f"逐节 cols/@w:num：模板 {_seq_txt(tpl_cols['seq'])}"
                 f"（{tpl_cols['n']} 节） 产物 {_seq_txt(prod_cols['seq'])}"
                 f"（{prod_cols['n']} 节）；正文节配对（EZI-0069）：模板"
                 f" {tpl_cols['how']}、产物 {prod_cols['how']}")
    if tpl_cols["n"] and prod_cols["n"] and tpl_cols["n"] != prod_cols["n"]:
        cols_note += f"；节数不同（{tpl_cols['n']} vs {prod_cols['n']}）→ 取末节对比"
    if tpl_cols["num"] is None:
        add("正文分栏数", None, prod_cols["num"] or "1", NA, CATEGORY_SECTION,
            "模板没有任何 w:sectPr → 无期望值；" + cols_note)
    else:
        ok = str(tpl_cols["num"]) == str(prod_cols["num"])
        if ok and tpl_cols["seq"] != prod_cols["seq"]:
            cols_note += "；正文节一致（前序节不同不影响正文分栏）"
        add("正文分栏数", str(tpl_cols["num"]), prod_cols["num"] or "1",
            _verdict(ok), CATEGORY_SECTION, cols_note)

    return report
