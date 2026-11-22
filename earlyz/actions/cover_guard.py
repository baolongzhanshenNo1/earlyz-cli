"""封面完整性兜底：交付前把未被重建的封面还原成源文档原样。

背景（2026-10-26 真实场景全量测试，199 份裸内容）：
正文规则（首行缩进、行距、标点归一、编号格式…）会连带改到封面，交付门
`cover polluted` 随即整篇拒绝交付。逐条给每个动作打补丁既漏又脆——这里
按"封面是模板/源文档托管区"的既有口径，在交付前统一还原：

- 只在 **本次没有重建封面** 时生效（重建时封面来自模板，本就该与源不同）；
- 封面范围与 report_fixes / new_actions 同一套判定：第一个承载 sectPr 的
  段落为止；单节文档退化为"第一个标题之前"；
- 段落数不一致（结构已变）时不动，宁可让交付门报错也不猜。
"""
from __future__ import annotations

import copy
import re
import shutil
import tempfile
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _cover_paragraphs(body) -> list:
    """返回封面区段落元素。

    优先用 new_actions._compute_cover_boundary（raw XML 版 find_cover_end，
    与 lint/writer 同一套锚点：章节序号/outlineLvl/语义标题词/固定锚点词）。
    2026-10-26 实测：劳动合同这类无分节符也无 heading 的文档，自写规则
    算不出封面区（返回 0 段），导致兜底空转。
    """
    paras = [c for c in body if c.tag == f"{W}p"]
    try:
        from earlyz.actions.new_actions import _compute_cover_boundary as _bound

        _n = int(_bound(body) or 0)
        if _n > 0:
            return paras[:_n]
    except Exception:  # noqa: BLE001
        pass
    end = None
    for i, p in enumerate(paras):
        ppr = p.find(f"{W}pPr")
        if ppr is not None and ppr.find(f"{W}sectPr") is not None:
            end = i
            break
    if end is not None:
        return paras[: end + 1]
    # 单节文档：第一个标题之前
    for i, p in enumerate(paras):
        ppr = p.find(f"{W}pPr")
        if ppr is None:
            continue
        st = ppr.find(f"{W}pStyle")
        ol = ppr.find(f"{W}outlineLvl")
        if (st is not None and "heading" in (st.get(f"{W}val") or "").lower()) or (
            ol is not None and (ol.get(f"{W}val") or "") in ("0", "1", "2")
        ):
            return paras[:i]
    return []


def restore_cover_from_source(target_path, source_path) -> dict:
    """把 target 的封面段落还原成 source 的原样；返回执行摘要。"""
    info = {"applied": False, "reason": "", "paragraphs": 0}
    tp, sp = Path(target_path), Path(source_path)
    if not tp.exists() or not sp.exists():
        info["reason"] = "缺文件"
        return info
    with zipfile.ZipFile(sp) as zs:
        s_root = etree.fromstring(zs.read("word/document.xml"))
    with zipfile.ZipFile(tp) as zt:
        t_root = etree.fromstring(zt.read("word/document.xml"))

    s_body, t_body = s_root.find(f"{W}body"), t_root.find(f"{W}body")
    if s_body is None or t_body is None:
        info["reason"] = "无 body"
        return info
    s_cov, t_cov = _cover_paragraphs(s_body), _cover_paragraphs(t_body)
    if not s_cov or not t_cov:
        info["reason"] = "封面区为空"
        return info

    # 书签对搬回时用**新 id**（w:bookmarkStart/@w:id 必须文档级唯一）：
    # 源段自带的书签可能占着目标段原来的 id，直接复用会重号。
    _bid = [0]
    for _el in t_body.iter(f"{W}bookmarkStart"):
        try:
            _bid[0] = max(_bid[0], int(_el.get(f"{W}id") or 0))
        except (TypeError, ValueError):
            pass

    # 不做「段落数必须相等」的硬判定：管道可能往封面区补段（页眉线/空段）。
    # 2026-10-26 JCM 实测：源 55 段 / 产物 57 段 → 守卫整段放弃 → 封面被
    # 正文缩进 879、两端对齐污染后仍整篇拒绝交付。
    # 改为按「文本指纹」配对：产物里能对上源封面段的，整段替换回原样；
    # 对不上的（新增段）原样保留。
    def _fp(p) -> str:
        return "".join(t.text or "" for t in p.iter(f"{W}t")).strip()

    # 归一化指纹：去掉空白与全部标点后再比对。产物里封面段的文本可能被
    # 标点归一改过（「—」→「——」等），精确指纹对不上会导致漏还原
    # （2026-10-26 JCM 实测：Abstract/Keywords 两段漏还原）。
    def _nfp(p) -> str:
        raw = _fp(p)
        return "".join(ch for ch in raw if ch.isalnum())

    s_by_fp: dict[str, list] = {}
    s_by_nfp: dict[str, list] = {}
    for sp in s_cov:
        key = _fp(sp)
        if key:
            s_by_fp.setdefault(key, []).append(sp)
        nkey = _nfp(sp)
        if nkey:
            s_by_nfp.setdefault(nkey, []).append(sp)

    changed = 0
    used: set[int] = set()
    for t_p in t_cov:
        key = _fp(t_p)
        if not key:
            continue
        s_p = None
        for pool, k in ((s_by_fp, key), (s_by_nfp, _nfp(t_p))):
            for i, c in enumerate(pool.get(k) or []):
                if id(c) not in used:
                    s_p = c
                    used.add(id(c))
                    break
            if s_p is not None:
                break
        if s_p is None:
            continue
        s_xml = etree.tostring(s_p)
        if etree.tostring(t_p) == s_xml:
            continue
        new_p = etree.fromstring(s_xml)
        # 结构不变量（2026-11-10 IEEE 英文期刊实锤）：分节符是**段落属性**
        # （w:pPr/w:sectPr），挂在封面区最后一段上。整段替换成源文档版本时
        # 源段没有 sectPr → 分节符被一并带走 → 题名区（1 栏）和正文（2 栏）
        # 合并成单节 → 题名被挤进左栏、页眉继承也错位。
        # 版式可以还原，结构不能丢：目标段带 sectPr 时原样挂回。
        _t_ppr = t_p.find(f"{W}pPr")
        _t_sp = _t_ppr.find(f"{W}sectPr") if _t_ppr is not None else None
        if _t_sp is None:
            # 兜底：极少数文档会有重复 pPr（历史转换产物），分节符可能在后面的那个
            _t_sp = next((sp for ppr in t_p.findall(f"{W}pPr")
                          for sp in ppr.findall(f"{W}sectPr")), None)
        if _t_sp is not None:
            _n_ppr = new_p.find(f"{W}pPr")
            if _n_ppr is None:
                _n_ppr = etree.Element(f"{W}pPr")
                new_p.insert(0, _n_ppr)
            if _n_ppr.find(f"{W}sectPr") is None:
                _n_ppr.append(copy.deepcopy(_t_sp))
        # 书签不变量（EZI-0038 实测）：整段替换会把目标段上的
        # w:bookmarkStart/w:bookmarkEnd 一起带走。目录条目的 PAGEREF 与
        # w:hyperlink/@w:anchor、题注交叉引用的 _RefF/_RefT 都指向这些书签——
        # 一旦丢失，Word 打开就是「错误！未找到引用源」（数据库设计说明书
        # 实测：封面边界把整篇当封面，17 个 _Toc* 书签被一次性换掉，
        # 17 个 PAGEREF 全悬空）。源段自带同名书签时以源为准（那是迁移过来的
        # 原书签）；源段没有的，把目标的整对搬回去。
        _t_bms = t_p.findall(f"{W}bookmarkStart")
        if _t_bms:
            _kept = {b.get(f"{W}name") or "" for b in new_p.findall(f"{W}bookmarkStart")}
            _lost = [b for b in _t_bms if (b.get(f"{W}name") or "") not in _kept]
            if _lost:
                _pos = new_p.find(f"{W}pPr")
                _at = list(new_p).index(_pos) + 1 if _pos is not None else 0
                _ends = []
                for _b in _lost:
                    _bid[0] += 1
                    _nb = copy.deepcopy(_b)
                    _nb.set(f"{W}id", str(_bid[0]))
                    new_p.insert(_at, _nb)
                    _at += 1
                    _ne = etree.Element(f"{W}bookmarkEnd")
                    _ne.set(f"{W}id", str(_bid[0]))
                    _ends.append(_ne)
                for _ne in _ends:
                    new_p.append(_ne)
        t_p.getparent().replace(t_p, new_p)
        changed += 1
    if not changed:
        info.update({"applied": True, "reason": "无需还原"})
        return info

    # 回写：整包重写（保持其它部件字节不变）
    with zipfile.ZipFile(tp) as zt:
        items = [(i, zt.read(i.filename)) for i in zt.infolist()]
    tmp = Path(tempfile.mkdtemp()) / tp.name
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zw:
        for item, data in items:
            if item.filename == "word/document.xml":
                data = etree.tostring(t_root, xml_declaration=True,
                                      encoding="UTF-8", standalone=True)
            zw.writestr(item, data)
    shutil.move(str(tmp), str(tp))
    info.update({"applied": True, "reason": "已还原", "paragraphs": changed})
    return info


def sanitize_heading_outline(target_path) -> dict:
    """交付前剥掉"句子形态段落"上的大纲级别。

    与交付门 _heading_shape_errors 同一判据：文本含「。」或「；」的段落不可能
    是标题。2026-10-26 真实场景：劳动合同条款「一、本合同文本供…；」被后续
    动作重新打上 outlineLvl（ensure_heading_outline 之后还有写入点），
    交付门判 heading shape 硬错误整篇拒绝。放在交付前统一收口。
    """
    info = {"applied": False, "stripped": 0}
    tp = Path(target_path)
    if not tp.exists():
        info["reason"] = "缺文件"
        return info
    with zipfile.ZipFile(tp) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    data = dict((i.filename, d) for i, d in items)
    if "word/document.xml" not in data:
        info["reason"] = "无 document.xml"
        return info
    root = etree.fromstring(data["word/document.xml"])
    stripped = 0
    for p in root.iter(f"{W}p"):
        text = "".join(t.text or "" for t in p.iter(f"{W}t"))
        if "。" not in text and "；" not in text:
            continue
        ppr = p.find(f"{W}pPr")
        if ppr is None:
            continue
        for tag in (f"{W}outlineLvl",):
            ol = ppr.find(tag)
            if ol is not None:
                ppr.remove(ol)
                stripped += 1
    if not stripped:
        info.update({"applied": True, "reason": "无需处理"})
        return info
    tmp = Path(tempfile.mkdtemp()) / tp.name
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zw:
        for item, blob in items:
            if item.filename == "word/document.xml":
                blob = etree.tostring(root, xml_declaration=True,
                                      encoding="UTF-8", standalone=True)
            zw.writestr(item, blob)
    shutil.move(str(tmp), str(tp))
    info.update({"applied": True, "stripped": stripped})
    return info


# ═══════════════════════════════════════════════════════════════════
# EZI-0014：引擎不得给非标题段落新增 outlineLvl
# ═══════════════════════════════════════════════════════════════════
#
# 由来（2026-11-16 交大中期检查报告实测，probe 追到写入点）：模板规范把
# outline_lvl 当「可移植类型属性」抄进**非标题角色**规则——affiliation（单位/
# 学院行）规则的 expected 里带 outline_lvl=0（decompiler `_PORTABLE_TYPED_KEYS`
# 含 outline_lvl，值来自模板样本段的直接格式）。`set_paragraph_format` 按 role
# 批量下发，封面表单字段「学院 （系）:」就被打上大纲级别 0；`outlineLvl` 正是
# Word 判定「进目录/进导航」的依据 → 表单字段作为一条一级条目进 TOC。
#
# 判据与 EZI-0012 的标题取样口径同源（`layout_fidelity.heading_samples_by_level`）：
#   · 保留：标题样式 / 编号标题形态（analyzer `_looks_like_numbered_heading`，
#     EZI-0012 取样认的是同一套编号判据）/ 语义章节词 / 后置节名（参考文献·
#     致谢·附录——引擎有意给它们大纲级别，F9 重建目录后不能消失）/ 摘掉
#     outlineLvl 后仍被判成 heading 的段（否则「有 outlineLvl 就算标题」自证）。
#   · 剥掉：表格内段落 / 表单标签（冒号结尾短行）/ 封面拉丁大写行 / 句子形态 /
#     没有任何标题证据的段。
# 只在「源文档本来没有」时动手：源文档自己的错标（交大封面 SHANGHAI… 校名行
# 带 outlineLvl=0）不归本动作管——那是 EZI-0012 判定层已排除的事。
# 例外（EZI-0023）：**章层级文档**（一级层级已被「第X章」占据）里，源文档标的
# 0 级也不可信——东台招标 14 段条款项 + 表格里的「60天」「1.6」全是源标的 0 级。
# 这时继承来的级别也要过一遍标题证据；非章层级文档维持上面的旧口径。

#: 允许写 outlineLvl 的角色（与 new_actions._outline_level_for_role 同一集合）
_OUTLINE_KEEP_ROLES = frozenset({
    "heading1", "heading2", "heading3", "heading4",
    "refs_heading", "ack_heading", "appendix_heading", "bio_heading",
})
#: 表单标签：冒号结尾的短行（「学院 （系）:」）。真标题另有标题样式/编号形态
#: 证据，会先被保留分支接住，不会因为这个上限被误剥。
_FORM_LABEL_MAX = 24
_FORM_LABEL_RE = re.compile(r"[:：]\s*$")


def _outline_text(p) -> str:
    """段内文字。"""
    return "".join(t.text or "" for t in p.iter(f"{W}t"))


def _outline_key(text: str) -> str:
    """配对指纹：只留字母/数字/汉字。

    产物会被标点归一（「学院 (系):」→「学院 （系）:」）与空白重排改字面，
    比对用指纹才能认出「这段源文档本来就有大纲级别」。
    """
    return "".join(ch for ch in (text or "") if ch.isalnum())


def _is_cover_caps_line(text: str) -> bool:
    """封面拉丁大写行（``SHANGHAI JIAO TONG UNIVERSITY``）。

    带数字的不算——``1 INTRODUCTION`` 是编号标题，不该因为全大写被剥。
    与 EZI-0012 ``layout_fidelity._is_cover_caps_line`` 同一判据。
    """
    if not text.isascii() or any(c.isdigit() for c in text):
        return False
    letters = [c for c in text if c.isalpha()]
    return len(letters) >= 4 and all(c.isupper() for c in letters)


def _para_font_size_pt(p) -> float | None:
    """段落有效字号（半磅 → pt）：取段落标记与各 run 上的最大值。"""
    sizes: list[float] = []
    mark = p.find(f"{W}pPr/{W}rPr")
    if mark is not None:
        sz = mark.find(f"{W}sz")
        if sz is not None and sz.get(f"{W}val"):
            try:
                sizes.append(int(sz.get(f"{W}val")) / 2.0)
            except ValueError:
                pass
    for r in p.iter(f"{W}r"):
        rpr = r.find(f"{W}rPr")
        sz = rpr.find(f"{W}sz") if rpr is not None else None
        if sz is not None and sz.get(f"{W}val"):
            try:
                sizes.append(int(sz.get(f"{W}val")) / 2.0)
            except ValueError:
                pass
    return max(sizes) if sizes else None


def _heading_evidence(p, text: str, styles_map: dict,
                      chapter_tier: bool | None = None,
                      level: int | None = None) -> tuple[bool, str]:
    """该段凭什么算标题；``(False, 原因)`` = 不是标题 → 可以剥大纲级别。

    ``chapter_tier``（EZI-0023）：该文档的一级层级是否已被「第X章」占据。章层级
    文档里 `outlineLvl=0` 只能落在章级文本上——「一、」「2.」开头的编号条目是
    公告条目/合同条款，不是一级标题。判据与 `classify_paragraph_role` 同源（同一份
    analyzer 规则，不在本文件另造一套）。1、2 级不受影响：小节层级照旧按形态判。
    """
    from earlyz.ir.analyzer import (
        _CHAPTER_TITLE_WORDS,
        _back_heading_role,
        _chapter_tier,
        _is_top_tier_text,
        _looks_like_numbered_heading,
        classify_paragraph_role,
    )

    if chapter_tier is None:
        chapter_tier = _chapter_tier(p)
    ppr = p.find(f"{W}pPr")
    if level is None:
        _ol = ppr.find(f"{W}outlineLvl") if ppr is not None else None
        try:
            level = int(_ol.get(f"{W}val")) if _ol is not None else None
        except (TypeError, ValueError):
            level = None
    ps = ppr.find(f"{W}pStyle") if ppr is not None else None
    sid = ((ps.get(f"{W}val") or "") if ps is not None else "")
    pstyle_name = styles_map.get(sid, "") or sid
    if "heading" in pstyle_name.lower() or re.match(r"^标题\s*\d", pstyle_name):
        return True, "标题样式"
    # 后置节名（REFERENCES / APPENDIX A / 摘 要 / 致 谢）：TOC 锚点，必须留
    if _back_heading_role(text) is not None:
        return True, "后置节名"
    if re.sub(r"[\s\u3000\u25a1]+", "", text or "") in _CHAPTER_TITLE_WORDS:
        return True, "章节词"
    in_table = any(a.tag == f"{W}tbl" for a in p.iterancestors())
    if in_table:
        return False, "表格内"
    # 摘掉 outlineLvl 再分类：classify_paragraph_role 把「段上有 outlineLvl」
    # 本身当 heading 证据（analyzer._heading_level）——原样判会自证，表单字段
    # 一旦被打上级别就永远是标题。
    bare = etree.fromstring(etree.tostring(ppr)) if ppr is not None else None
    if bare is not None:
        for el in bare.findall(f"{W}outlineLvl"):
            bare.remove(el)
    role = classify_paragraph_role(
        bare, pstyle_name, False, in_table,
        bare is not None and bare.find(f"{W}numPr") is not None, text,
        in_hyperlink=p.find(f"{W}hyperlink") is not None,
        chapter_tier=chapter_tier,
    )
    if role in _OUTLINE_KEEP_ROLES:
        # 章层级文档里 0 级必须由「章级角色/章级文本」支撑：`classify_paragraph_role`
        # 会把「6. 投标函」这类单级编号判成 heading2（章内条目），它的级别是 1 不是 0。
        if chapter_tier and level == 0 and role in ("heading2", "heading3", "heading4") \
                and not _is_top_tier_text(text):
            return False, f"0 级与角色不符:{role}"
        return True, f"角色:{role}"
    # 噪声形态先于「编号/短语形态」：`_looks_like_numbered_heading` 是形态闸门
    # （短短语无句读即过），封面大写行与表单标签都会从那里漏过去——EZI-0012 的
    # 取样口径同样是把这两类排在编号判据之前排除。
    if len((text or "").strip()) <= _FORM_LABEL_MAX and _FORM_LABEL_RE.search(text or ""):
        return False, "表单标签"
    if _is_cover_caps_line((text or "").strip()):
        return False, "封面大写行"
    if _looks_like_numbered_heading(text):
        # 章层级文档里，0 级只留给章级文本（「第一章 投标人须知」）；「2.招标内容
        # 清单」「一、付款方式」这类 0 级编号条目是源文档标错的条款项 → 不算证据。
        # 也**不**再让「大字号短行」兜底：东台 16pt 的附件标题与条款项同形同字号，
        # 字号在那里区分不出标题（非章层级文档照旧走字号兜底）。
        if chapter_tier and level == 0 and not _is_top_tier_text(text):
            return False, "章层级文档里的 0 级编号条目"
        return True, "编号形态"
    if "。" in (text or "") or "；" in (text or ""):
        return False, "句子形态"
    # 题名区大字号短行（22pt/二号题名——形态识别不出、样式也没挂的那类真标题）。
    # 放在噪声判据之后：表单字段/大写行先被剥，这条只兜「大字号、无句读、短」。
    size = _para_font_size_pt(p)
    if size is not None and size >= 15.0 and len((text or "").strip()) <= 40:
        return True, f"大字号短行:{size:g}pt"
    return False, "无标题证据"


def strip_invented_outline_levels(target_path, source_path) -> dict:
    """剥掉「源文档没有、也不是标题」的段落上的 outlineLvl（EZI-0014）。

    A1 落实：产物 outlineLvl 段落集合 ⊆ 源文档 outlineLvl 段落 ∪ 真标题集合。
    源文档自己有的大纲级别原样保留；引擎新增的只在有标题证据时保留——判据见模块
    注释（与 EZI-0012 取样口径同源）。幂等：已收口过再跑 stripped=0。

    EZI-0023（「不得照搬」）：上面那条「源文档自己标的就原样保留」在**章层级文档**
    里不成立——东台招标源文档 29 段正文 0 级里只有 3 段是「第X章」，其余是条款项
    （「2.采购代理机构信息」…）与表格数字（「60天」「1.6」「2.2」）。所以章层级
    文档里继承来的级别也要过一遍标题证据；非章层级文档（交大中期这类报告/表单）
    维持原口径，封面校名行等源标记照旧不动。
    """
    info: dict = {"applied": False, "stripped": 0,
                  "kept_inherited": 0, "kept_heading": 0, "details": []}
    tp, sp = Path(target_path), Path(source_path)
    if not tp.exists():
        info["reason"] = "缺文件"
        return info
    with zipfile.ZipFile(tp) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    data = dict((i.filename, d) for i, d in items)
    if "word/document.xml" not in data:
        info["reason"] = "无 document.xml"
        return info
    root = etree.fromstring(data["word/document.xml"])

    # 文档级事实（EZI-0023）：产物自己的一级层级是否被「第X章」占据。
    # 放在「逐段判」之前算一次（analyzer 侧按文档树缓存，不是每段扫全篇）。
    from earlyz.ir.analyzer import _chapter_tier_for_root

    chapter_tier = _chapter_tier_for_root(root)
    info["chapter_tier"] = chapter_tier

    # 源文档自己的大纲级别（多重集合：同一段文字可能出现多次）
    inherited: Counter = Counter()
    if sp.exists():
        try:
            with zipfile.ZipFile(sp) as zs:
                s_root = etree.fromstring(zs.read("word/document.xml"))
            for p in s_root.iter(f"{W}p"):
                ppr = p.find(f"{W}pPr")
                if ppr is None or ppr.find(f"{W}outlineLvl") is None:
                    continue
                key = _outline_key(_outline_text(p))
                if key:
                    inherited[key] += 1
        except Exception:  # noqa: BLE001 — 源读不了就只按标题证据判（保守剥）
            inherited = Counter()

    styles_map: dict = {}
    try:
        from earlyz.actions.new_actions import _read_styles_map

        styles_map = _read_styles_map(tp) or {}
    except Exception:  # noqa: BLE001
        styles_map = {}

    stripped = 0
    for p in root.iter(f"{W}p"):
        ppr = p.find(f"{W}pPr")
        if ppr is None:
            continue
        ol = ppr.find(f"{W}outlineLvl")
        if ol is None:
            continue
        text = _outline_text(p)
        key = _outline_key(text)
        if not key:
            continue  # 空段不承载标题语义（也就不进目录），不动
        is_heading, why = _heading_evidence(p, text, styles_map, chapter_tier)
        if inherited.get(key, 0) > 0:
            inherited[key] -= 1
            if is_heading or not chapter_tier:
                info["kept_inherited"] += 1
                continue
            # 章层级文档 + 源标记 + 无标题证据 → 照搬的错级别，落到下面剥掉
            why = f"继承的源级别（{why}）"
        elif is_heading:
            info["kept_heading"] += 1
            continue
        ppr.remove(ol)
        stripped += 1
        if len(info["details"]) < 20:
            info["details"].append({"text": text.strip()[:40],
                                    "val": ol.get(f"{W}val"), "why": why})

    if not stripped:
        info.update({"applied": True, "reason": "无需处理"})
        return info
    tmp = Path(tempfile.mkdtemp()) / tp.name
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zw:
        for item, blob in items:
            if item.filename == "word/document.xml":
                blob = etree.tostring(root, xml_declaration=True,
                                      encoding="UTF-8", standalone=True)
            zw.writestr(item, blob)
    shutil.move(str(tmp), str(tp))
    info.update({"applied": True, "stripped": stripped})
    return info



def snapshot_cover_ppr(target_path) -> list[bytes]:
    """快照封面段的 pPr（模板重建后立刻调用）。

    重建封面时不能整段还原成源文档（文案来自模板），但后续正文规则仍会
    往里写首行缩进/对齐（2026-10-26 实测 JCM：封面「Paper Title:」被套
    首行缩进 879）。这里锁住段落属性，交付前回填。
    """
    tp = Path(target_path)
    if not tp.exists():
        return []
    with zipfile.ZipFile(tp) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    out: list[bytes] = []
    for p in _cover_paragraphs(root.find(f"{W}body")):
        ppr = p.find(f"{W}pPr")
        out.append(etree.tostring(ppr) if ppr is not None else b"")
    return out


def restore_cover_ppr(target_path, snaps: list[bytes]) -> dict:
    """把封面段 pPr 回填成快照值；文案与 run 属性不动。"""
    info = {"applied": False, "restored": 0}
    tp = Path(target_path)
    if not tp.exists() or not snaps:
        info["reason"] = "无快照"
        return info
    with zipfile.ZipFile(tp) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    data = dict((i.filename, d) for i, d in items)
    if "word/document.xml" not in data:
        info["reason"] = "无 document.xml"
        return info
    root = etree.fromstring(data["word/document.xml"])
    paras = _cover_paragraphs(root.find(f"{W}body"))
    if len(paras) != len(snaps):
        info["reason"] = f"封面段数变化（{len(snaps)}→{len(paras)}）"
        return info
    changed = 0
    for p, snap in zip(paras, snaps):
        cur = p.find(f"{W}pPr")
        cur_xml = etree.tostring(cur) if cur is not None else b""
        if cur_xml == snap:
            continue
        if cur is not None:
            p.remove(cur)
        if snap:
            p.insert(0, etree.fromstring(snap))
        changed += 1
    if not changed:
        info.update({"applied": True, "reason": "无需回填"})
        return info
    tmp = Path(tempfile.mkdtemp()) / tp.name
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zw:
        for item, blob in items:
            if item.filename == "word/document.xml":
                blob = etree.tostring(root, xml_declaration=True,
                                      encoding="UTF-8", standalone=True)
            zw.writestr(item, blob)
    shutil.move(str(tmp), str(tp))
    info.update({"applied": True, "restored": changed})
    return info


def relax_cell_nowrap(target_path, min_chars: int = 20) -> dict:
    """放开长文本单元格的「禁止换行」，避免整行拉出页面。

    w:noWrap 告诉 Word 该单元格内容不换行。封面表单（如东南本科的
    「AI 工具使用情况说明表」）里常把填写示例写进 noWrap 单元格，
    一句话拉成 755pt → 渲染后判 image/table bbox overflow，整篇拒绝交付
    （2026-10-26 实测 p1 text_block x1=845.6 / 页宽 595.3）。
    短标签（≤ min_chars）保留 noWrap，避免破坏表单栅格。
    """
    info = {"applied": False, "relaxed": 0}
    tp = Path(target_path)
    if not tp.exists():
        info["reason"] = "缺文件"
        return info
    with zipfile.ZipFile(tp) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    data = dict((i.filename, d) for i, d in items)
    if "word/document.xml" not in data:
        info["reason"] = "无 document.xml"
        return info
    root = etree.fromstring(data["word/document.xml"])
    relaxed = 0
    for tc in root.iter(f"{W}tc"):
        tcpr = tc.find(f"{W}tcPr")
        if tcpr is None:
            continue
        nw = tcpr.find(f"{W}noWrap")
        if nw is None:
            continue
        text = "".join(t.text or "" for t in tc.iter(f"{W}t")).strip()
        if len(text) <= min_chars:
            continue
        tcpr.remove(nw)
        relaxed += 1
    if not relaxed:
        info.update({"applied": True, "reason": "无需处理"})
        return info
    tmp = Path(tempfile.mkdtemp()) / tp.name
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zw:
        for item, blob in items:
            if item.filename == "word/document.xml":
                blob = etree.tostring(root, xml_declaration=True,
                                      encoding="UTF-8", standalone=True)
            zw.writestr(item, blob)
    shutil.move(str(tmp), str(tp))
    info.update({"applied": True, "relaxed": relaxed})
    return info
