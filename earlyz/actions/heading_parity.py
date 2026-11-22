"""同文标题对齐：目标里与模板同名标题段，直接照抄模板段落的对齐/加粗/字号。

为什么需要（2026-11-01 实测）：标题规则是**按级**取值（样本池 → 众数），而模板同级标题
常常本来就不一致——浙江政采 55 个标题段里公告区居中、格式区两端对齐；招投标文件（中国
仪器仪表学会）模板居中、产物两端。按级规则无论取哪个值都会错一半（heading 失败 33 份里
「对齐」17 项、「字号」23 项大多是这一类）。

模板对**同一段文字**已经给出了它自己的排法，照抄就是最忠实的做法。只处理「模板里有同名
标题段」的段落；模板没写的新内容仍走按级规则。

取值必须在 **XML 层**读（run 级优先）：模板段落常常 pPr 没写字号、run 上写 sz=24
（12pt），而样式链 BodyText 是 22（11pt）。只看段落/样式会抄到错值（实测：勘察合同
heading 仍判 11≠12）。作用属性：jc / bold / font_size_pt / **font_east_asia**
（前三项与评分器 heading 检查同口径）。

中文字体（EZI-0019，2026-11-16）：``w:rFonts`` 是四个互相独立的槽，本动作原先只搬
jc/bold/字号、**从不写 @w:eastAsia**，于是标题中文一个都没迁（东台招标三级标题：模板
28/28 仿宋，产物 10/10 宋体）。eastAsia 的取值**不按文本逐个抄**，而是走「模板同级
标题实测众数」——与评分器 ``_heading_expectation`` 完全同函数、同护栏
（``heading_samples_by_level`` + ``_template_mode``）。按文本抄会翻车：东台模板二级
标题自己就不统一（仿宋 8/16、黑体 6/16），抄到的那一处与评分器的期望值必然对不上；
同级无主峰时评分器退回 spec 规则，引擎也就必须**不下发**（A4），不许拿默认值猜。

级别进匹配键（EZI-0024，2026-11-16）：同文对齐原先**只按文本配对**，模板里同一段文字
出现在正文/格式段时，产物段哪怕是一级标题也照抄那段正文的 10.5/12pt——东台招标产物
34 个 ``Heading1`` 因此出现 3 种字号（16×11 / 12×9 / 10.5×14）。现在匹配键是
``(级别, 文本)``：产物段**有**级别就必须与模板同文段同级，否则一个属性都不抄（宁可
不抄也不跨级）；产物段**取不到**级别时保持旧行为（只按文本认人，它本来就不是标题段）。

字体证据的第二个例外（EZI-0024 A2）：EZI-0019 的「模板样本 < 5 不下发」按**数量**
判，把「3/3 全是黑体」也挡掉了——3 段全同值不是投票噪声，而是模板唯一的写法。
``TEMPLATE_FONT_CONSISTENT_MIN`` 起算，全一致即允许下发；样本不一致时护栏不变。

逐段证据（EZI-0028，2026-11-16）：模板该级**没有主峰**时（东台二级 8 仿宋 + 6 黑体，
主峰 8/16 = 50% < ``TEMPLATE_MODE_MIN_SHARE``），原先整级一个字体都不发，于是 6 个
「附件格式标题」一直保持产物的仿宋。众数没有主峰 ≠ 模板没写：这 6 段自己在 run 上
写着 ``w:eastAsia="黑体"``。现在该级没有可信众数时退到**逐段**证据——只抄「模板这一段
自己写的」值（run 级或样式链，docDefaults 不算），逐段各写各的，**不**把众数甩给全级
（「模板该级本就不统一时不得乱下发」）。有可信众数时仍走 EZI-0024 的护栏，一字未改。

同文成对采信（EZI-0037，2026-11-17）：判定层从 EZI-0025 起就把「与模板**同文同级**成对
≥ 2 段」当**第一来源**（``_heading_expectation``），绕过 ``TEMPLATE_MODE_MIN_SAMPLES``；
引擎这条链却还在用「有效样本 ≥ 5 + 重合率 ≥ 50%」，于是同一份证据两层结论相反——留出集
30 对里 43 对看不见的四项（一级/二级标题中文字体）就是这么来的：重庆科技大学一级模板只有
2 段（2/2 黑体）被样本下限挡掉、深圳招标一级同文 10 段（10/10 黑体）但同级 31 段重合 32%
被重合率挡掉，判定层判 FAIL、引擎一个字都不写。

采信判据是**众数**、不是「全一致」（EZI-0037 回工，与判定层 EZI-0039 同一口径）：成对
≥ ``HEADING_SAME_TEXT_MIN`` 段且最高频占比 ≥ ``_DOMINANCE_DEFAULT``（decompiler 里
0.6，本文件不另造阈值）。深圳招标一级同文 9×16 + 1×22 就是这一支：16pt 占 90% ≥ 60%。

采信之后字号/字体**逐段照抄模板这一段自己的值**，绝不拿该级众数整批覆盖：深圳招标第三章
模板本来就是 22pt（同文成对里的 1/10 少数），拿众数 16pt 整级下发会把它拖成 16pt——实测
32/32/32/32/32，比基线 32/32/44/32/32 更不像模板。**整级补齐**（判定层不收窄实测样本的
级别里，模板无同文对应的段也得跟上众数）因此只在同文成对**全部一致**时才做：重庆科技
大学一级 2/2 走这条，深圳 9×16+1×22 不走。没有模板证据时引擎**不写**字号——判定层此时
退回 spec 规则，照抄模板块自己的值只会离期望更远。
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

from lxml import etree

from earlyz.actions.new_actions import (
    W,
    _locate_paragraphs,
    _read_part,
    _read_styles_map,
    _rpr_insert_ordered,
    _write_part,
)
from earlyz.ir.selectors import normalize_text

_HEADING_ROLES = ("heading1", "heading2", "heading3")
_HEADING_LEVEL = {"heading1": 1, "heading2": 2, "heading3": 3}
# 西文/数字字体（项目字体铁律）：ascii/hAnsi 恒 Times New Roman，与中文字体分开写。
_LATIN_FONT = "Times New Roman"
# 产物段落自己的级别来自样式名（Word 内置 "Heading 1" / 中文界面 "标题 1"）。
_LEVEL_NAME_RE = re.compile(r"^(?:heading|标题)\s*([1-9])$", re.IGNORECASE)
#: 「样本少但全一致」也算证据时要求的最小样本数（EZI-0024 A2）。模板该级标题 3 段全写
#: 同一个中文字体时，众数不是投票噪声而是模板唯一写法（东台一级 3/3 黑体）；1~2 段
#: 没有独立重复，仍不算证据。
TEMPLATE_FONT_CONSISTENT_MIN = 3


def _text_of(p: etree._Element) -> str:
    return "".join(t.text or "" for t in p.iter(f"{W}t"))


def _para_props(p: etree._Element, style_map: dict | None = None,
                defaults: dict | None = None) -> dict:
    """与评分器同口径的有效值：run 直接格式 > 段落直接格式 > 样式链 > docDefaults。

    ⚠ 不能用「段落标记 rPr（pPr/rPr）」当兜底：勘察合同的模板标题段 run 上没写字号，
    12pt 写在样式链里，段落标记的 11pt 会把值抄错（实测 heading 仍判 11≠12）。

    返回里 ``east_asia_src`` 记中文字体的来源（``run`` / ``style`` / ``default``，
    EZI-0028）：逐段证据只采信前两者——``docDefaults`` 是全局默认值，拿它下发等于猜。
    """
    ppr = p.find(f"{W}pPr")
    jc = None
    if ppr is not None:
        el = ppr.find(f"{W}jc")
        if el is not None:
            jc = el.get(f"{W}val")
    bold = None
    size = None
    east_asia = None
    # eastAsia 的**来源**（EZI-0028）：run 直接格式 / 样式链 / docDefaults。前两者是模板
    # 对这一段自己的写法，可以当逐段证据；docDefaults 只是全局默认值，拿它下发等于猜。
    east_asia_src = None
    for r in p.iter(f"{W}r"):
        rpr = r.find(f"{W}rPr")
        if rpr is None:
            continue
        if bold is None:
            b = rpr.find(f"{W}b")
            if b is not None:
                bold = (b.get(f"{W}val") or "1") not in ("0", "false")
        if size is None:
            sz = rpr.find(f"{W}sz")
            if sz is not None and sz.get(f"{W}val"):
                try:
                    size = int(sz.get(f"{W}val")) / 2.0
                except ValueError:
                    pass
        if east_asia is None:
            # 中文字体（EZI-0024）：模板同文段自己的 eastAsia（run 级优先，缺失沿样式链）。
            # 用来做「逐段证据」：与众数不一致时说明这一段模板另有写法（东台封面
            # 「招 标 文 件」42pt 宋体，而一级众数是章标题的黑体），不许拿众数覆盖它。
            rf = rpr.find(f"{W}rFonts")
            if rf is not None and rf.get(f"{W}eastAsia"):
                east_asia = rf.get(f"{W}eastAsia")
                east_asia_src = "run"
        if bold is not None and size is not None and east_asia is not None:
            break
    sid = None
    if ppr is not None:
        pel = ppr.find(f"{W}pStyle")
        if pel is not None:
            sid = pel.get(f"{W}val")
    layer = (style_map or {}).get(sid) or {}
    if jc is None:
        jc = layer.get("jc")
    if bold is None and layer.get("bold") is not None:
        bold = bool(layer.get("bold"))
    if size is None:
        size = layer.get("font_size_pt")
    if size is None:
        size = (defaults or {}).get("font_size_pt")
    if east_asia is None:
        # 模板段没在 run 上写字体时沿样式链取（与评分器 `effective_props` 同源）：
        # 只认 run 级会把「样式里定字体」的模板标题全判成无证据，字体一个都不下发。
        east_asia = layer.get("font_east_asia")
        east_asia_src = "style" if east_asia else None
    if east_asia is None:
        east_asia = (defaults or {}).get("font_east_asia")
        east_asia_src = "default" if east_asia else None
    if jc is None:
        jc = (defaults or {}).get("jc")
    return {"jc": jc, "bold": bold, "size": size, "east_asia": east_asia,
            "east_asia_src": east_asia_src}


def _set_jc(p: etree._Element, jc: str | None) -> bool:
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        if not jc:
            return False
        ppr = etree.Element(f"{W}pPr")
        p.insert(0, ppr)
    cur = ppr.find(f"{W}jc")
    if jc is None:
        if cur is None:
            return False
        ppr.remove(cur)
        return True
    if cur is not None:
        if cur.get(f"{W}val") == jc:
            return False
        cur.set(f"{W}val", jc)
        return True
    el = etree.SubElement(ppr, f"{W}jc")
    el.set(f"{W}val", jc)
    return True


def _set_run_props(p: etree._Element, bold: bool | None, size_pt: float | None) -> bool:
    changed = False
    for r in p.iter(f"{W}r"):
        rpr = r.find(f"{W}rPr")
        if bold is not None:
            b = rpr.find(f"{W}b") if rpr is not None else None
            if bold and b is None:
                if rpr is None:
                    rpr = etree.Element(f"{W}rPr")
                    r.insert(0, rpr)
                etree.SubElement(rpr, f"{W}b")
                changed = True
            elif not bold and b is not None:
                rpr.remove(b)
                changed = True
        if size_pt:
            half = str(int(round(size_pt * 2)))
            if rpr is None:
                rpr = etree.Element(f"{W}rPr")
                r.insert(0, rpr)
            for tag in ("sz", "szCs"):
                el = rpr.find(f"{W}{tag}")
                if el is None:
                    el = etree.SubElement(rpr, f"{W}{tag}")
                    el.set(f"{W}val", half)
                    changed = True
                elif el.get(f"{W}val") != half:
                    el.set(f"{W}val", half)
                    changed = True
    return changed


def _set_east_asia(p: etree._Element, east_asia: str) -> bool:
    """给段落里带文字的 run 写 ``w:rFonts/@w:eastAsia``，ascii/hAnsi 钉在 TNR。

    为什么只写 eastAsia（EZI-0019）：``w:rFonts`` 的几个槽互相独立——``ascii`` /
    ``hAnsi`` 管西文与数字、``eastAsia`` 管中文。本动作此前只搬 jc/bold/字号，中文字体
    一个都没迁（东台招标三级标题：模板 28/28 仿宋，产物 10/10 宋体）。写入时**只覆盖
    eastAsia**，并把 ascii/hAnsi 校正到 Times New Roman（铁律：中文随模板、西文/数字
    TNR）——产物现状恰好两头都错：``ascii="仿宋_GB2312;仿宋" w:eastAsia="宋体"``，
    英文/数字跟着中文字体走。``cs``（复杂脚本）不动，与中英文分离无关。

    返回是否有改动（幂等：值已对就返回 False）。
    """
    changed = False
    for r in p.iter(f"{W}r"):
        if not any((t.text or "").strip() for t in r.iter(f"{W}t")):
            continue
        rpr = r.find(f"{W}rPr")
        if rpr is None:
            rpr = etree.Element(f"{W}rPr")
            r.insert(0, rpr)
        rf = rpr.find(f"{W}rFonts")
        if rf is None:
            rf = etree.Element(f"{W}rFonts")
            _rpr_insert_ordered(rpr, rf)          # rFonts 必须排在 b/sz 之前（RPR_SEQ）
            changed = True
        if rf.get(f"{W}eastAsia") != east_asia:
            rf.set(f"{W}eastAsia", east_asia)
            changed = True
        for attr in ("ascii", "hAnsi"):
            if rf.get(f"{W}{attr}") != _LATIN_FONT:
                rf.set(f"{W}{attr}", _LATIN_FONT)
                changed = True
    return changed


def _paragraph_level(p: etree._Element, style_map: dict | None) -> int | None:
    """产物段落**自己的**标题级别（1/2/3），取不到给 ``None``。

    只认两种已有事实：段落 ``pPr/w:outlineLvl``（EZI-0014 之后标题段恒有）或样式名
    （``Heading 1`` / ``标题 1``）。取不到级别就没有「同级模板众数」可对——宁可不下发
    也不拿别级的值凑（EZI-0019 A4）。第二遍（模板侧是标题、产物侧没被判成标题的段）
    靠它拿到级别。
    """
    ppr = p.find(f"{W}pPr")
    if ppr is not None:
        ol = ppr.find(f"{W}outlineLvl")
        if ol is not None:
            try:
                lvl = int(ol.get(f"{W}val") or "")
            except ValueError:
                lvl = -1
            if 0 <= lvl <= 2:
                return lvl + 1
    sid = None
    if ppr is not None:
        st = ppr.find(f"{W}pStyle")
        if st is not None:
            sid = st.get(f"{W}val")
    name = (style_map or {}).get(sid or "") or (sid or "")
    m = _LEVEL_NAME_RE.match(str(name).strip())
    return int(m.group(1)) if m else None


def _template_heading_east_asia(template_path: Path,
                               target_doc: dict | None = None) -> dict[int, dict]:
    """模板各级标题的**证据口径**（字体 EZI-0019 / 同文成对与字号 EZI-0037），取样复用 EZI-0016。

    ``heading_samples_by_level``（outlineLvl 级别 + 编号标题文本，排除表格内/后置节名/
    表单标签/封面大写行）+ ``_template_mode``（有效样本 ≥ ``TEMPLATE_MODE_MIN_SAMPLES``
    且众数占比 ≥ ``TEMPLATE_MODE_MIN_SHARE`` 才算 ``usable``）。**不许在此另写取样**：
    引擎与评分器各测各的，产物照模板排也会被判 FAIL（EZI-0016 的原病）。

    证据按判定层 ``_heading_expectation`` 的优先级排（EZI-0037 起与它同序）：

    1. ``same_text_*``：``_same_text_pairs``（**判定层同一个函数**）给出「与产物同文同级」
       的模板段；成对 ≥ ``HEADING_SAME_TEXT_MIN`` 段且该属性**取众数**、最高频占比
       ≥ ``_DOMINANCE_DEFAULT``（0.6，与 decompiler/判定层同一口径）时采信，**绕过**
       ``TEMPLATE_MODE_MIN_SAMPLES`` 与 ``HEADING_OVERLAP_MIN``。1:1 对应的同一批标题不需要
       统计前提（重庆科技大学一级 2/2 黑体、深圳招标一级 10/10 黑体都被「样本 < 5」挡过）。
       1 段同文仍不算证据（计算机学报投稿二级只有 1 段）。
    2. ``usable``：评分器口径（样本够 + 有主峰），保持不变；
    3. ``consistent``：样本数不够但**全都同值**（东台一级 3/3 黑体）——按数量一刀切会把
       「模板唯一写法」当噪声挡掉；``TEMPLATE_FONT_CONSISTENT_MIN`` 起算。这是字体独有的
       历史分支（判定层没有对应物），只对 ``font_east_asia`` 生效；
    4. ``mode``：该级众数，没有同文证据时的退路。

    ``trusted`` = 有任一条证据。``downlink``（整级下发）仍要求模板/产物同级标题文本重合率
    ≥ ``HEADING_OVERLAP_MIN``：那是判定层**不收窄实测样本**的条件
    （``_comparable_heading_rows``），两侧不是同一批标题时整级下发会改坏同级的别的段
    （东台一级 3/34 = 9%，见 EZI-0021 条目 4）。

    ``size_*``：字号用**同一份**同文成对证据（判定层字号/字体/对齐三项共用
    ``_same_text_pairs``）。与字体不同，判定层对字号**没有** ``consistent`` 分支——同文
    取不到众数就退回 spec，所以没有证据时引擎必须**不写**（``size_trusted`` 为假），照抄
    模板这一段自己的值只会离 spec 更远。

    ``size_downlink``（整级补齐字号）**只在同文成对全部一致时**才为真（``n == total``）：
    众数占比 60% 只够「采信这一批段各自的模板值」，不够「拿它刷掉整级」。深圳招标一级
    9×16 + 1×22（16pt 占 90%）时整级补齐会把模板自己写成 22pt 的第三章拖成 16pt。

    返回 ``{1: meas, 2: meas, 3: meas}``。``value/n/total/share/usable`` 恒为**该级众数**
    口径（评分器原样，``_no_peak_reason`` 与人话 note 都读它）；引擎要下发的值看
    ``east_asia``（字体）与 ``size_value``（字号）。``same_text_keys`` 是该级「与模板同文」
    的产物段文本键（归一化），供整级补齐跳过它们——**绝不用众数覆盖已与模板逐段一致的
    段**。无样本的级别 ``total=0``，由调用方写进 info 的 note（A4「模板无该级证据不下发」）。
    """
    from earlyz.spec.decompiler import _DOMINANCE_DEFAULT
    from earlyz.validation.layout_fidelity import (
        _HEAD_PROP_FIELD,
        _dist_txt,
        _heading_overlap,
        _load_docx,
        _norm_heading_text,
        _same_text_pairs,
        _template_mode,
        HEADING_OVERLAP_MIN,
        HEADING_SAME_TEXT_MIN,
        heading_samples_by_level,
    )

    ea_key = _HEAD_PROP_FIELD["font_east_asia"]
    sz_key = _HEAD_PROP_FIELD["font_size_pt"]
    samples, _excl = heading_samples_by_level(_load_docx(template_path))
    prod = heading_samples_by_level(target_doc)[0] if target_doc is not None else {}
    out: dict[int, dict] = {}
    for lvl, rows in samples.items():
        meas = dict(_template_mode(rows, ea_key))     # 众数口径（评分器原样保留）
        # 全一致：n == total (>0) 且样本数过「独立重复」下限。`_template_mode` 的
        # usable 仍按数量判（评分器口径），这里只给引擎加一条证据分支。
        consistent = bool(meas["total"]) and meas["n"] == meas["total"] and \
            meas["total"] >= TEMPLATE_FONT_CONSISTENT_MIN
        meas["consistent"] = consistent
        # 同文成对（EZI-0037，回工取众数）：与判定层 `_same_text_pairs` **同一函数、同一
        # 「同文」判据**（`_norm_heading_text` 去空白）。≥2 段且众数占比 ≥ 0.6 即采信
        # （判定层 EZI-0039 同一口径）；「全一致」另记，用来判能不能整级补齐。
        prod_rows = prod.get(lvl) or []
        same = _same_text_pairs(rows, prod_rows) if target_doc is not None else []
        same_rows = [t for _, t in same]
        same_ea = _template_mode(same_rows, ea_key)
        same_sz = _template_mode(same_rows, sz_key)
        enough = len(same) >= HEADING_SAME_TEXT_MIN
        font_ok = bool(enough and same_ea["total"]
                       and same_ea["share"] >= _DOMINANCE_DEFAULT)
        size_ok = bool(enough and same_sz["total"]
                       and same_sz["share"] >= _DOMINANCE_DEFAULT)
        font_uniform = bool(enough and same_ea["total"]
                            and same_ea["n"] == same_ea["total"])
        size_uniform = bool(enough and same_sz["total"]
                            and same_sz["n"] == same_sz["total"])
        meas["same_text"] = len(same)
        meas["same_text_keys"] = sorted({_norm_heading_text(p["text"]) for p, _ in same})
        meas["same_text_uniform"] = font_uniform
        meas["same_text_value"] = same_ea["value"]
        meas["same_text_share"] = round(float(same_ea["share"]), 3) if same_ea["total"] else 0.0
        meas["same_text_size_value"] = same_sz["value"]
        meas["same_text_size_share"] = (round(float(same_sz["share"]), 3)
                                       if same_sz["total"] else 0.0)
        # EZI-0062：同文批的 jc 众数——判定层 ``_same_text_expectation`` 对齐分支的
        # 引擎侧数据。0053 的 ``same_text_count``（同一文本在模板里的出现次数）与
        # 判定层 ``HEADING_SAME_TEXT_MIN``（跨文本聚合成对数）是假等价——唯一文本
        # 标题（深圳「第X章」×10 段各自 cnt=1）全被旧闸误禁。
        same_jc = _template_mode(same_rows, "jc")
        meas["same_text_jc_value"] = same_jc["value"]
        meas["same_text_jc_share"] = (round(float(same_jc["share"]), 3)
                                      if same_jc["total"] else 0.0)
        meas["same_text_jc_total"] = int(same_jc["total"])
        meas["east_asia"] = same_ea["value"] if font_ok else meas["value"]
        meas["trusted"] = bool(meas["usable"] or consistent or font_ok)
        # 字号（EZI-0037）：判定层没有 consistent 分支，只有「同文众数采信」与「该级众数
        # 可用且两侧同批」两种模板来源，其余退回 spec。
        size_mode = _template_mode(rows, sz_key)
        meas["size_mode_value"] = size_mode["value"]
        if size_ok:
            meas["size_value"], meas["size_source"] = same_sz["value"], "same_text"
        elif size_mode["usable"]:
            meas["size_value"], meas["size_source"] = size_mode["value"], "mode"
        else:
            meas["size_value"], meas["size_source"] = None, None
        # 「为什么没写字号」的人话（info["fonts"] 的 note 直接用）
        if size_ok:
            meas["size_reason"] = ""
        elif same:
            meas["size_reason"] = (
                (f"模板同文 {len(same)} 段字号无主峰（{_dist_txt(same_rows, sz_key)}，"
                 f"最高频 {float(same_sz['share']) * 100:.0f}% <"
                 f"{_DOMINANCE_DEFAULT:.0%}）")
                if enough else f"模板同文仅 {len(same)} 段（须 ≥{HEADING_SAME_TEXT_MIN}）")
        else:
            meas["size_reason"] = "模板该级无同文成对样本"
        if target_doc is None:
            meas.update({"overlap": None, "overlap_hit": None, "overlap_total": None,
                         "downlink": meas["trusted"],
                         "size_trusted": bool(size_ok or size_mode["usable"]),
                         "size_downlink": bool(size_uniform or size_mode["usable"])})
        else:
            hit, ratio = _heading_overlap(rows, prod_rows)
            in_batch = ratio >= HEADING_OVERLAP_MIN
            meas.update({
                "overlap": round(ratio, 3), "overlap_hit": hit,
                "overlap_total": len(prod_rows),
                "downlink": bool(meas["trusted"] and in_batch),
                # 逐段照抄模板字号只在「判定层会用模板实测」时成立：同文取到众数 → 逐段
                # 照抄模板这一段自己的值；众数可用且两侧同批 → 也是逐段照抄。其余退回 spec。
                "size_trusted": bool(size_ok or (size_mode["usable"] and in_batch)),
                # 整级补齐字号：只有同文成对**全部一致**（且判定层不收窄）才做——众数
                # 60% 只够采信各自的模板值，拿它刷整级会把少数派段拖平（深圳第三章 44）。
                "size_downlink": bool(size_uniform and in_batch),
            })
            if not meas["size_trusted"] and size_mode["usable"]:
                meas["size_reason"] += (f"；众数可用但两侧标题重合 {hit}/{len(prod_rows)}"
                                        f"（{ratio * 100:.0f}% <{HEADING_OVERLAP_MIN:.0%}）")
        out[int(lvl)] = meas
    return out


def _props_rank(d: dict) -> tuple:
    """模板同一段文字有多处时，挑「排成标题样子」的那一处：加粗/居中优先，其次字号大的。"""
    return (1 if (d.get("bold") or d.get("jc") == "center") else 0, d.get("size") or 0)


def _style_outline_levels(path: Path) -> dict[str, int]:
    """styleId → 样式链上的 ``outlineLvl`` + 1（无级别的样式不收）。

    Word 常把大纲级别写在**样式**里、段落上不写；``_paragraph_level`` 只认直接
    ``outlineLvl`` 和「Heading N / 标题 N」样式名，自定义样式名（「章标题」之类）会漏。
    评分器 ``effective_props`` 是沿样式链取 ``outline_lvl`` 的，引擎取级别必须同源，
    否则同文对齐会把「评分器认作一级」的模板段当成无级别（EZI-0024 A1）。
    """
    out: dict[str, int] = {}
    try:
        with zipfile.ZipFile(str(path)) as z:
            root = etree.fromstring(z.read("word/styles.xml"))
    except Exception:  # noqa: BLE001 — 没有/读不了 styles.xml 就按「无样式级」量
        return out
    raw: dict[str, tuple[int | None, str | None]] = {}
    for s in root.iter(f"{W}style"):
        if s.get(f"{W}type") != "paragraph":
            continue
        sid = s.get(f"{W}styleId")
        if not sid:
            continue
        lvl = None
        ol = s.find(f"{W}pPr/{W}outlineLvl")
        if ol is not None:
            try:
                v = int(ol.get(f"{W}val") or "")
                lvl = v + 1 if 0 <= v <= 2 else None
            except ValueError:
                lvl = None
        base = s.find(f"{W}basedOn")
        raw[sid] = (lvl, base.get(f"{W}val") if base is not None else None)

    def _resolve(sid: str | None, depth: int = 0, seen: set | None = None) -> int | None:
        seen = seen or set()
        if not sid or sid in seen or depth > 6 or sid not in raw:
            return None
        seen.add(sid)
        own, base = raw[sid]
        return own if own is not None else _resolve(base, depth + 1, seen)

    for sid in raw:
        lvl = _resolve(sid)
        if lvl is not None:
            out[sid] = lvl
    return out


def _template_heading_props(template_path: Path,
                            template_ir) -> dict[tuple[int | None, str], dict]:
    """模板里「被判定为标题」的段落 → ``(级别, 规范化文本)`` → run 级合成属性。

    标题证据 = IR 标题角色 ∪ 编号形态段（2026-11-07）：参考文献里的
    「一、中文著作」这类条款级标题，IR 常判成 body（它不是章标题），但它确实
    是模板排好的标题段；只按 IR 角色取证据，产物里同名段就永远拿不到模板的
    字号/对齐（华东政法 4 份、北科 6 份、数据提供合同 6 份 heading 全卡在这）。
    形态判据直接用 analyzer 的 `_looks_like_numbered_heading`（手标 126 段校准过，
    尾冒号拒、句子形态拒），不另造一套。

    **匹配键含级别（EZI-0024 A1）**：模板同文段是正文/格式段时级别是 ``None``，
    是二级标题时是 2——产物同文的一级标题段不许照抄它们（东台招标：34 个
    ``Heading1`` 被抄成 10.5/12/16 三种字号）。级别取模板段自己的 ``outlineLvl``
    / 样式名（``_paragraph_level``，与产物侧同一函数）。**不能拿 IR 角色兜底**：
    analyzer 的 ``_is_cn_numbered_heading`` 会把「一、项目基本情况」这类条款项也判成
    heading1（EZI-0021 条目 2），用它兜底等于把正文段升成一级，A1 立刻失守；评分器
    也只认 outlineLvl，取不到级别就按「无级别」处理才同源。
    同一段文字在模板里出现多次（目录条目 + 真标题）时仍按「排成标题样子」挑一处，
    只保留那一处的级别——这是旧口径的取舍，不额外扩一种匹配。
    """
    heading_texts = set()
    try:
        from earlyz.ir.analyzer import _looks_like_numbered_heading
    except Exception:  # noqa: BLE001
        _looks_like_numbered_heading = None  # type: ignore[assignment]
    for p in getattr(template_ir, "paragraphs", []) or []:
        role = getattr(p, "semantic_role", "") or ""
        raw = getattr(p, "text", "") or ""
        t = normalize_text(raw)
        if not t:
            continue
        if role in _HEADING_ROLES:
            heading_texts.add(t)
            continue
        if _looks_like_numbered_heading and _looks_like_numbered_heading(raw):
            heading_texts.add(t)
    cand: dict[str, list[tuple[int | None, dict]]] = {}
    if not heading_texts:
        return cand  # type: ignore[return-value]
    from earlyz.spec.decompiler import _style_effective_map

    style_map, defaults = _style_effective_map(template_path)
    # 取级别要用 styleId → **样式名**（`_paragraph_level` 匹配 "Heading 1"/"标题 1"）；
    # `_style_effective_map` 是 styleId → 有效格式，喂给它只会得到 None。
    style_names = _read_styles_map(template_path)
    style_levels = _style_outline_levels(template_path)
    # 没有 pStyle 的段落走「默认段落样式」——它常是 styleId="a"（name=Normal）这种
    # 非标准 id，`_style_effective_map` 的键里找不到，于是字号/对齐全丢
    # （华东政法参考文献「一、中文著作」：模板 12pt/两端，产物照抄成 18pt/居中）。
    try:
        with zipfile.ZipFile(str(template_path)) as z:
            _st = etree.fromstring(z.read("word/styles.xml"))
        for _s in _st.iter(f"{W}style"):
            if (_s.get(f"{W}type") == "paragraph"
                    and (_s.get(f"{W}default") or "") in ("1", "true")):
                _sid = _s.get(f"{W}styleId")
                if _sid and style_map.get(_sid):
                    style_map = dict(style_map)
                    style_map[None] = style_map[_sid]
                break
    except Exception:  # noqa: BLE001
        pass
    with zipfile.ZipFile(str(template_path)) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    for p in root.iter(f"{W}p"):
        # EZI-0053：表格内段落不是标题样本（判定层 heading_samples_by_level 同
        # 口径排除表格）。标书模板「7. 合同授予」在评标说明表格里有一份 jc=both
        # 的副本，_props_rank（加粗/对齐优先）把它排在正文真标题（jc 未设）前面，
        # 同文对齐把产物首发写好的 jc=left 改成 both → 判定层 FAIL。表格副本
        # 永不作为证据，全在表格里 = 不配对、产物保持首发值。
        if any(a.tag == f"{W}tc" for a in p.iterancestors()):
            continue
        t = normalize_text(_text_of(p))
        if t and t in heading_texts:
            lvl = _paragraph_level(p, style_names)
            if lvl is None:
                ps = p.find(f"{W}pPr/{W}pStyle")
                lvl = style_levels.get(ps.get(f"{W}val") if ps is not None else "")
            cand.setdefault(t, []).append((lvl, _para_props(p, style_map, defaults)))
    # 同一段文字在模板里常有两处：目录条目（小字、左对齐）与真正的标题段。
    # 取「排成标题样子」的那一处——加粗/居中优先，其次字号大的（评分器取证据同口径）。
    out: dict[tuple[int | None, str], dict] = {}
    for t, items in cand.items():
        lvl, spec = max(items, key=lambda it: _props_rank(it[1]))
        # EZI-0053：同文成对段数随 spec 带出——jc 镜像要与判定层同下限
        # （`HEADING_SAME_TEXT_MIN`），孤例不许单独定对齐。
        spec = dict(spec)
        spec["same_text_count"] = len(items)
        out[(lvl, t)] = spec
    return out


def align_headings_to_template(target_path: Path, template_path: Path,
                               template_ir=None,
                               spec_heading_jc: dict[int, str] | None = None) -> dict:
    """把「模板同名标题段」的对齐/加粗/字号/中文字体直写到目标段落上。非阻断。

    同文对齐的匹配键是 ``(级别, 文本)``（EZI-0024 A1）：产物段有级别时必须与模板
    同文段同级才抄——模板同文段常是正文/格式段（东台招标 31/34 个一级标题段），
    按纯文本抄会把 10.5/12pt 抄到 ``Heading1`` 上，同一级出现 3 种字号。产物段
    取不到级别时退回纯文本配对（它本来就不是标题段，抄来的也是同一段文字）。
    """
    target_path = Path(target_path)
    template_path = Path(template_path)
    info: dict = {"applied": 0, "roles": {}, "errors": []}
    try:
        from earlyz.ir.analyzer import analyze_document
        from earlyz.validation.layout_fidelity import HEADING_SAME_TEXT_MIN

        tir = template_ir if template_ir is not None else analyze_document(template_path)
        tmap = _template_heading_props(template_path, tir)
        if not tmap:
            return info
        # 同一段文字在模板里可能出现在不同级别（封面 / 目录 / 正文 / 标题）。按文本
        # 建一份「排得最像标题」的兜底表，只给产物**取不到级别**的段落用（旧行为）；
        # 产物段一旦有级别就必须同级命中——东台招标的条款项就是这样被抄成一级标题
        # 字号的（EZI-0024 A1：34 个 Heading1 出现 10.5/12/16 三种）。
        by_text: dict[str, dict] = {}
        for (_lvl, _t), _spec in tmap.items():
            _cur = by_text.get(_t)
            if _cur is None or _props_rank(_spec) > _props_rank(_cur):
                by_text[_t] = _spec

        def _parity_spec(text: str, level: int | None) -> dict | None:
            """模板同文段属性；产物段有级别就只认同级（EZI-0024 A1）。"""
            if level is not None:
                return tmap.get((level, text))
            return by_text.get(text)

        root = _read_part(target_path, "word/document.xml")
        styles_map = {}
        try:
            styles_map = _read_styles_map(target_path)
        except Exception:  # noqa: BLE001
            pass
        # EZI-0019：中文字体 = 模板同级标题实测众数（与评分器同函数、同护栏）。
        # 无样本/无主峰 → 不下发（A4）；样本少但全一致 → 下发（EZI-0024 A2）。
        # 该级与产物标题重合不足时只记「不整级下发」，逐段认同文同级（见 _apply_fonts）。
        font_stat: dict[str, dict] = {}
        # EZI-0062：jc 闸证据（聚合成对数 + 同文批 jc 众数），见两遍写入处的闸注释。
        level_jc_evidence: dict[int, dict] = {}
        # 同文段文本键（每级）：整级补齐必须跳过它们——「绝不用众数覆盖已与模板逐段一致
        # 的段」（EZI-0037 回工）。`parity_seen` 只覆盖 `_parity_spec` 命中的段；模板段没被
        # IR 认成标题时，它仍在同文成对证据里、却不在 `parity_seen` 里。
        same_text_guard: dict[int, set[str]] = {}
        try:
            from earlyz.validation.layout_fidelity import (
                HEADING_OVERLAP_MIN,
                TEMPLATE_MODE_MIN_SAMPLES,
                TEMPLATE_MODE_MIN_SHARE,
                _no_peak_reason,
                _norm_heading_text,
            )

            styles_root = None
            try:
                styles_root = _read_part(target_path, "word/styles.xml")
            except Exception:  # noqa: BLE001 — 没有 styles 部件就按「无样式」量
                pass
            meas_by_level = _template_heading_east_asia(
                template_path, {"document": root, "styles": styles_root})
            for lvl, meas in sorted(meas_by_level.items()):
                same_text_guard[int(lvl)] = set(meas.get("same_text_keys") or ())
                level_jc_evidence[int(lvl)] = {
                    "pairs": int(meas.get("same_text") or 0),
                    "jc": meas.get("same_text_jc_value"),
                    "jc_share": float(meas.get("same_text_jc_share") or 0.0),
                    "jc_total": int(meas.get("same_text_jc_total") or 0),
                }
                trusted = bool(meas.get("trusted"))
                # 整级下发（中文字体）：`downlink` = 有证据 + 重合率达标（判定层不收窄）
                applied = bool(meas.get("downlink") and meas.get("east_asia"))
                rec = {
                    "role": f"heading{lvl}",
                    "east_asia": meas.get("east_asia"),
                    "samples": int(meas.get("n") or 0),
                    "total": int(meas.get("total") or 0),
                    "share": round(float(meas.get("share") or 0.0), 3),
                    "consistent": bool(meas.get("consistent")),
                    "overlap": meas.get("overlap"),
                    # 同文成对段数（EZI-0037）：判定层 `_same_text_pairs` 的同一份证据
                    "same_text": int(meas.get("same_text") or 0),
                    "same_text_uniform": bool(meas.get("same_text_uniform")),
                    # 同文成对里字体/字号的众数占比（回工：≥0.6 才采信）
                    "same_text_share": float(meas.get("same_text_share") or 0.0),
                    "same_text_size_share": float(meas.get("same_text_size_share") or 0.0),
                    "trusted": trusted,
                    "applied": applied,
                    "paragraphs": 0,
                    # 该级没有可信众数时，按「模板同文且同级」的逐段证据写了多少段
                    # （EZI-0028：东台二级 8 仿宋 + 6 黑体，主峰 50% → 逐段各写各的）
                    "per_paragraph": 0,
                    # 字号（EZI-0037）：该写哪个值 / 有没有模板证据（没有就不写）
                    "size_value": meas.get("size_value"),
                    "size_source": meas.get("size_source"),
                    "size_trusted": bool(meas.get("size_trusted")),
                    "size_downlink": bool(meas.get("size_downlink")),
                    "size_paragraphs": 0,
                    # 整级补齐（同文段之外、该级模板无同文对应的段）改了几段
                    "level_wide": 0,
                }
                if not applied:
                    if not trusted:
                        rec["note"] = f"不下发：{_no_peak_reason(meas)}"
                    else:
                        rec["note"] = (
                            f"不整级下发：模板同级 {meas.get('n')} 段与产物同级标题文本"
                            f"重合 {meas.get('overlap_hit')}/{meas.get('overlap_total')}"
                            f"（{(meas.get('overlap') or 0) * 100:.0f}% <"
                            f" {HEADING_OVERLAP_MIN:.0%}），两侧不是同一批标题；"
                            "只写「模板同文且同级」的段落（EZI-0024）")
                if not rec["size_trusted"]:
                    rec["note"] = (rec["note"] + "；" if rec.get("note") else "") + (
                        f"字号不写：{meas.get('size_reason') or '无模板证据'}"
                        "（判定层此时退回 spec 规则，EZI-0037）")
                font_stat[f"heading{lvl}"] = rec
        except Exception as e:  # noqa: BLE001 — 非阻断：字体不下发，其余照旧
            info["errors"].append(f"heading_fonts: {type(e).__name__}: {e}")
        info["fonts"] = font_stat

        def _east_asia_for(level) -> str | None:
            """该级模板实测中文字体；该级众数没有证据（不是 trusted）给 None。"""
            rec = font_stat.get(f"heading{level}") if level else None
            return rec["east_asia"] if rec and rec["trusted"] else None

        # ⚠ 不能拿 id(p) 判「这段处理过没有」：lxml 的代理对象会被回收，同一条底层
        # 节点两次迭代可能给出不同 id（实测：6 段被数成 9 段）。存引用才稳。
        font_done: list[etree._Element] = []

        def _apply_fonts(p: etree._Element, level: int | None, spec: dict) -> bool:
            """按段落**自己的**级别下发中文字体（级别取不到就不下发）。

            级别用 `_paragraph_level`（outlineLvl > 样式名）而不是纯 role：评分器
            `heading_samples_by_level` 正是按 outlineLvl 分组，两侧口径必须同源。

            三级判据（EZI-0024）：① 该级众数可信（样本够/有主峰/**全一致**）才谈下发；
            ② 该级与产物标题重合不足时（东台一级 3/34 = 9%）不做**整级**下发，只给
            「模板同文且同级」的段落写；③ 模板同文段自己写的 ``eastAsia`` 必须等于该级
            众数——否则模板对这段另有排法（东台封面「招 标 文 件」42pt 宋体），不许覆盖。

            该级**没有可信众数**时（EZI-0028）：①~③ 全都不成立——没有众数可比，也就
            没有「整级下发」可言，此时退到**逐段证据**：模板对这一段自己写了字体
            （run 级或样式链，不含 docDefaults）就照抄。东台二级 8 仿宋 + 6 黑体
            （主峰 50% < 60%）原先把整级一个都不发，6 个「附件格式标题」因此一直保持
            产物的仿宋；逐段证据让它们各自拿到模板自己的黑体，而**不是**把众数发给全级。
            """
            if level is None:
                return False
            if any(x is p for x in font_done):
                return False
            rec = font_stat.get(f"heading{level}")
            if rec is None:
                return False
            ea = _east_asia_for(level)          # 该级可信众数（None = 没有可信众数）
            if ea:
                # 有可信众数：模板同文段自己另写字体时不覆盖（封面护栏，EZI-0024）
                if spec.get("east_asia") != ea:
                    return False
            else:
                # 无可信众数：只认「模板这一段自己写的」字体（run 级或样式链），
                # docDefaults 不算证据。样本不足时仍按 EZI-0019 A4 不下发——「该级本身
                # 没什么可测」和「模板有两套写法」是两件事：前者逐段抄会把孤例当规律，
                # 后者逐段各抄各的才是忠实还原（东台二级 16 段 8/6 属后者）。
                if int(rec.get("total") or 0) < TEMPLATE_MODE_MIN_SAMPLES:
                    return False
                if spec.get("east_asia_src") not in ("run", "style"):
                    return False
                ea = spec.get("east_asia")
                if not ea:
                    return False
                rec["per_paragraph"] = int(rec.get("per_paragraph") or 0) + 1
            font_done.append(p)
            rec["paragraphs"] += 1
            return _set_east_asia(p, ea)

        def _size_for(level, spec: dict) -> float | None:
            """该段该写的**字号**（``None`` = 不动）。EZI-0037 与判定层同口径。

            判定层 ``_heading_expectation('font_size_pt')`` 只有两种「模板来源」：同文成对
            取众数且最高频占比 ≥ ``_DOMINANCE_DEFAULT``（≥ ``HEADING_SAME_TEXT_MIN`` 段，
            绕过样本下限；EZI-0039 与回工同一口径），或该级众数可用且两侧标题重合
            ≥ ``HEADING_OVERLAP_MIN``；其余**退回 spec 规则**。所以：

            * 有模板证据 → 照抄**模板这一段自己**的字号（``spec['size']``）。**绝不**
              拿该级众数整批写：深圳招标一级同文 9×16 + 1×22（16pt 占 90%），第三章
              模板自己就是 22pt，拿众数整级写会把它拖成 16pt（产物 32/32/32/32/32，
              比基线 32/32/44/32/32 更不像模板）。
            * 没有证据 → **不写**：这条链拿不到 spec（pipeline 不在写作用域），唯一同口径
              的做法是让前面 spec 步骤留下的值留在产物里。照抄模板块自己的值只会离期望更远。
            * 产物段取不到级别（``level is None``）→ 保持旧行为，只按文本认人（EZI-0024）。
            """
            if level is None or spec is None:
                return spec["size"] if spec else None
            rec = font_stat.get(f"heading{level}")
            if rec is None:
                return spec["size"]
            return spec["size"] if rec.get("size_trusted") else None

        touched = 0
        # 同文命中的段（走 `_parity_spec` 的那批）：第三遍「整级补齐」必须跳过它们——
        # 它们受 EZI-0024 的封面护栏管辖（模板这一段自己另写字体时不许覆盖）。
        parity_seen: list[etree._Element] = []
        for role in _HEADING_ROLES:
            level_hint = _HEADING_LEVEL[role]
            for p in _locate_paragraphs(root, {"semantic_role": role}, styles_map):
                key = normalize_text(_text_of(p))
                level = _paragraph_level(p, styles_map) or level_hint
                spec = _parity_spec(key, level)
                if not spec:
                    continue
                parity_seen.append(p)
                ch = False
                # EZI-0062：jc 镜像闸改用**判定层的同变量**——0053 拿
                # ``same_text_count``（同一文本在模板里的出现次数）对齐判定层的
                # ``HEADING_SAME_TEXT_MIN``（跨文本聚合成对数）是假等价：唯一文本标题
                # （深圳「第X章」10 段各自 cnt=1）与「模板唯一样本即产物自己的格式
                # 示范」（计算机学报 L2/L3）全被旧闸误禁。现在三级：
                # ① 该级聚合成对 ≥ 下限且同文批 jc 众数占比 ≥
                #   ``TEMPLATE_MODE_MIN_SHARE`` → 写众数（判定层
                #   ``_same_text_expectation`` 同口径；深圳 10 对 center 直证）；
                # ② 成对不足 → 显式写 spec 规则的 jc（判定层同分支的兜底期望；
                #   计算机学报 L2/L3：6.4b 对自定义样式段没写到，靠这里补）；
                # ③ spec 也没有 → 保持现值（0053 行为）。
                # 开封 L3「7. 合同授予」孤例 both 走 ②（spec=left→start）不回归；
                # 加粗/字号/字体不受此闸（各有自己的证据链）。
                if level is not None:
                    _ev = level_jc_evidence.get(int(level)) or {}
                    if (_ev.get("pairs", 0) >= HEADING_SAME_TEXT_MIN
                            and _ev.get("jc_total")
                            and _ev.get("jc_share", 0.0) >= TEMPLATE_MODE_MIN_SHARE):
                        ch |= _set_jc(p, _ev["jc"] or "left")
                    elif (spec_heading_jc or {}).get(int(level)):
                        ch |= _set_jc(p, spec_heading_jc[int(level)])
                else:
                    # EZI-0062：by_text 纯文本兜底恢复**无条件照抄**（11-01 原契约
                    # 「模板对同一段文字已经给了排法」）——0053 把 same_text_count 计数
                    # 闸也压到了这条路上，唯一文本从此永不照抄（基线测试
                    # test_heading_parity_2026_11_01 红了三轮没人看见）。开封式
                    # 错位配对的真实形态走 level-matched 路径（②spec 兜底管住），
                    # 这条路的保护本来就靠表格排除 + _props_rank（EZI-0024/0053）。
                    ch |= _set_jc(p, spec["jc"] or "left")
                ch |= _set_run_props(p, spec["bold"], _size_for(level, spec))
                ch |= _apply_fonts(p, level, spec)
                if ch:
                    touched += 1
                    info["roles"][role] = info["roles"].get(role, 0) + 1
        # 第二遍：模板侧是标题、但产物侧**没被判成标题**的段（摘要 / ABSTRACT / 参考文献 /
        # 附录 …）。它们只按文本配对，角色不一致就永远轮不到第一遍（实测：北京大学博士
        # 「摘要」未加粗、Springer「References」字号 10≠12 都是这一类）。跳过目录条目
        # （pStyle 以 TOC 开头，或含 PAGEREF 域）与表格内段落。
        # 中文字体（EZI-0019）：这一遍没有 role，级别改从段落自己的 outlineLvl / 样式名
        # 取（`_paragraph_level`）；取不到级别就不下发——级别错了比不下发更糟。
        for p in root.iter(f"{W}p"):
            key = normalize_text(_text_of(p))
            level = _paragraph_level(p, styles_map)
            spec = _parity_spec(key, level)
            if not spec:
                continue
            pstyle = p.find(f"{W}pPr/{W}pStyle")
            sid = (pstyle.get(f"{W}val") if pstyle is not None else "") or ""
            if sid.upper().startswith("TOC"):
                continue
            if any("PAGEREF" in (x.text or "").upper() for x in p.iter(f"{W}instrText")):
                continue
            if p.getparent() is not None and any(
                    a.tag == f"{W}tc" for a in p.iterancestors()):
                continue
            parity_seen.append(p)
            ch = False
            # EZI-0062：与第一遍同闸（聚合成对 + 众数占比 → 众数；成对不足 → spec
            # 兜底；无级纯文本兜底路径保 0053 的 same_text_count 计数闸——东台/开封
            # 的错位配对保护不放松）。本遍不跳 parity_seen，标题段会再次路过——
            # 两遍闸必须同判，否则会把第一遍刚写的值改回另一套。
            if level is not None:
                _ev = level_jc_evidence.get(int(level)) or {}
                if (_ev.get("pairs", 0) >= HEADING_SAME_TEXT_MIN
                        and _ev.get("jc_total")
                        and _ev.get("jc_share", 0.0) >= TEMPLATE_MODE_MIN_SHARE):
                    ch |= _set_jc(p, _ev["jc"] or "left")
                elif (spec_heading_jc or {}).get(int(level)):
                    ch |= _set_jc(p, spec_heading_jc[int(level)])
            else:
                # EZI-0062：by_text 纯文本兜底无条件照抄（与第一遍同判；见第一遍闸注释）。
                ch |= _set_jc(p, spec["jc"] or "left")
            ch |= _set_run_props(p, spec["bold"], _size_for(level, spec))
            # 中文字体：级别只能从段落自己取（这一遍没有 role）；取不到就不写字体，
            # 别拿别的级别凑
            ch |= _apply_fonts(p, level, spec)
            if ch:
                touched += 1
                info["roles"]["text_only"] = info["roles"].get("text_only", 0) + 1
        # 第三遍（EZI-0037，回工后收窄）：**同文成对全部一致**时的整级补齐。判定层
        # 「不收窄实测样本」的级（两侧标题重合率 ≥ ``HEADING_OVERLAP_MIN``）里期望值是
        # **整级一个值**；模板没有同文对应的同级段（重庆科技大学一级的 2 段陈旧目录条目）
        # 不补就会把该级众数留在旧值上——判定层判 FAIL，引擎日志却写着「已下发」。
        #
        # 只在同文成对**全一致**（``size_uniform``/``applied``）时才补齐，**绝不**用众数
        # 覆盖已与模板逐段一致的段：
        #   * 同文段一律跳过（`parity_seen` + `same_text_guard`）——它们各自的模板值来自
        #     `_size_for`/`_apply_fonts`，深圳招标第三章 22pt 就是这样保住的；
        #   * 同文成对不统一（深圳 9×16 + 1×22）时整级一个值都写不得：众数 16pt 会把
        #     模板自己写成 22pt 的那一章拖平，产物从 32/32/44/32/32 退化成 32/32/32/32/32。
        # 表格内、目录样式段一律不碰。
        for p in root.iter(f"{W}p"):
            if any(x is p for x in parity_seen):
                continue
            if p.getparent() is not None and any(
                    a.tag == f"{W}tc" for a in p.iterancestors()):
                continue
            pstyle = p.find(f"{W}pPr/{W}pStyle")
            sid = (pstyle.get(f"{W}val") if pstyle is not None else "") or ""
            if sid.upper().startswith("TOC"):
                continue
            level = _paragraph_level(p, styles_map)
            rec = font_stat.get(f"heading{level}") if level else None
            if rec is None:
                continue
            if _norm_heading_text(_text_of(p)) in (same_text_guard.get(level) or set()):
                continue
            size_v = rec.get("size_value") if rec.get("size_downlink") else None
            ea_v = rec.get("east_asia") if rec.get("applied") else None
            if not size_v and not ea_v:
                continue
            ch = False
            if size_v:
                if _set_run_props(p, None, size_v):
                    ch = True
                    rec["size_paragraphs"] += 1
            if ea_v:
                ch |= _set_east_asia(p, ea_v)
            if ch:
                touched += 1
                rec["level_wide"] += 1
        # EZI-0028：该级没有可信众数时走的是「逐段证据」，note 要写明不是整级下发——
        # 否则 info 里只有「不下发」，而产物明明改了字体，读日志的人没法对齐。
        for _rec in font_stat.values():
            _n = int(_rec.get("per_paragraph") or 0)
            if _n and not _rec.get("applied"):
                _rec["note"] = (_rec.get("note") or "") + (
                    f"；但模板对这 {_n} 段自己写了字体（run/样式级），已按「模板同文且同级」"
                    f"逐段下发各自的模板值（EZI-0028）")
            # EZI-0037：整级补齐改了几段，必须能在 info 里看见（否则「整级下发」只体现
            # 在同文段的 paragraphs 上，读日志的人会以为没写）。回工后只在同文成对
            # **全部一致**时才会发生，note 里写清这条前提。
            _w = int(_rec.get("level_wide") or 0)
            if _w:
                _rec["note"] = (_rec.get("note") + "；" if _rec.get("note") else "") + (
                    f"该级另有 {_w} 段「模板无同文对应」，已按该级证据值补齐"
                    f"（同文成对全部一致才有整级值，EZI-0037 回工）")
        if touched:
            _write_part(target_path, "word/document.xml", root)
        info["applied"] = touched
    except Exception as e:  # noqa: BLE001 — 非阻断：失败只记录
        info["errors"].append(f"{type(e).__name__}: {e}")
    return info
