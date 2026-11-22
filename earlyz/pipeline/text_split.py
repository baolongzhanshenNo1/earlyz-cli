"""正文结构拆分归一（pre-analyze；2026-10-09 用户定案两项）。

与 latex 残留归一（`latex_scripts.normalize_latex_scripts_in_copy`）同一位
置：compile_document「0)」阶段、analyze **之前**、只作用于工作副本（用户
原件永不修改）——拆分后各段以自然段落进入分析/套模板/指纹链，格式、
角色、分页、目录全按拆后结构算。

两个确定性归一（均为保守判据，不满足即不动）：

1) `split_inline_enumerations`：行内序号列表化。段内「（1）…（2）…（3）…」
   形态、引导句以冒号收尾、编号自 1 连续升序（≥2 项）时，拆为「引导句 /
   各项 / 尾句」各自独立段落。用户 2026-10-09 定案：「顺序号就要按顺序排」
   ——每项末尾标点**一律原样**（「不要管每一项的末尾符号是不是句号」；
   「不确定符号对不对时先问用户」——本步从不改任何符号；仅删除编号与其
   后内容之间的空白，与用户验收图一致）。

2) `split_appendix_heading`：附录标题行内正文拆分。「附录 A 公式(3-5) 的
   NP-hard 性证明（详细证明置于此处，正文从略）」→「附录 A」标题段 + 其余
   内容正文段。用户 2026-10-09 定案：附录A是章标题，其余内容是属于正文
   （模板的附录标题即单独的「附录」）。

切片规格 = [(起始, 结束, 挖除区间列表或 None), ...]（字符级、相对段落文本，
挖除用于「编号后空白」这类不可达区间）。内容保持门：段落边界不进入内容
hash（analyzer._normalized_content_hash 口径），且本步在 analyze 之前作用
于工作副本——两层上内容保持检查不受影响。不改动的判定（跳过/歧义）如实
计数回报，供排查（不静默）。
"""

from __future__ import annotations

import copy
import re
import shutil
import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
_XMLSPACE = "{http://www.w3.org/XML/1998/namespace}space"

_ENUM_MARK_RE = re.compile(r"[（(](\d{1,2})[）)]")
_MARKER_TAIL_SPACE_RE = re.compile(r"^([（(]\d{1,2}[）)])[ \u3000\t]+")
_COLON_END_RE = re.compile(r"[：:]$")
_CJK_RE = re.compile(r"[\u4e00-\u9fff]")
_CHAPTER_RE = re.compile(r"^第\s*[0-9一二三四五六七八九十]{1,3}\s*章")
_DOT_LEADER_RE = re.compile(r"\.{4,}|…{2,}")
_APPENDIX_SPLIT_RE = re.compile(
    r"^[\t \u3000]*附录[\t \u3000]*[A-Za-z0-9一二三四五六七八九十]{1,4}")
_WS_CHARS = " \u3000\t"


# ── 段落手术基础设施 ────────────────────────────────────────────

def _segments(p):
    """段落的文本切片表 [(run, t_elem, text), ...]；含不支持内容返回 None。

    保守：任何非 pPr/r 子元素（hyperlink/sdt/bookmark/fldChar…）、run 内
    非 rPr/t 元素（tab/br/drawing/sym…）、单 run 多 t，一律放弃该段。
    """
    segs = []
    for c in p:
        tag = etree.QName(c).localname
        if tag == "pPr":
            continue
        if tag != "r":
            return None
        t_elems = []
        for rc in c:
            rtag = etree.QName(rc).localname
            if rtag == "rPr":
                continue
            if rtag != "t":
                return None
            t_elems.append(rc)
        if len(t_elems) > 1:
            return None
        if not t_elems:
            continue  # 空 run（仅 rPr）——不承载文本，忽略
        segs.append((c, t_elems[0], t_elems[0].text or ""))
    return segs or None


def _in_drops(pos, drops):
    for ds, de in drops:
        if ds <= pos < de:
            return True
    return False


def _slice_into_paragraphs(p, specs):
    """把段落按切片规格 [(a, b, drops), ...] 切为多个新段落元素。

    drops = 该区间内要删除的字符区间（如「编号后空白」）。part[0] 保留原
    段落属性（含 paraId）；part[1:] 去掉 w14:paraId/textId 与
    pageBreakBefore（拆分不引入分页）。
    """
    segs = _segments(p)
    if segs is None:
        return None
    total = sum(len(t) for _, _, t in segs)
    ppr = p.find(f"{W}pPr")
    parts = []
    for idx, spec in enumerate(specs):
        a, b = int(spec[0]), int(spec[1])
        drops = spec[2] if len(spec) > 2 and spec[2] else []
        a = max(0, min(a, total))
        b = max(0, min(b, total))
        if b <= a:
            return None
        new_p = etree.Element(p.tag)
        for k, v in p.attrib.items():
            if idx and k in (f"{_W14}paraId", f"{_W14}textId"):
                continue
            new_p.set(k, v)
        if ppr is not None:
            new_ppr = copy.deepcopy(ppr)
            if idx:
                pbb = new_ppr.find(f"{W}pageBreakBefore")
                if pbb is not None:
                    new_ppr.remove(pbb)
            new_p.append(new_ppr)
        pos = 0
        for (run, t_el, txt) in segs:
            seg_a = pos
            pos += len(txt)
            seg_b = pos
            if seg_b <= a or seg_a >= b:
                continue
            s0 = max(a, seg_a) - seg_a
            s1 = min(b, seg_b) - seg_a
            if drops:
                chars = []
                for i in range(s0, s1):
                    if _in_drops(seg_a + i, drops):
                        continue
                    chars.append(txt[i])
                piece = "".join(chars)
            else:
                piece = txt[s0:s1]
            if not piece:
                continue
            new_run = copy.deepcopy(run)
            kept_t = None
            for rc in list(new_run):
                if etree.QName(rc).localname == "t":
                    if kept_t is None:
                        kept_t = rc
                        rc.text = piece
                    else:
                        new_run.remove(rc)
            if kept_t is None:
                continue
            kept_t.set(_XMLSPACE, "preserve")
            new_p.append(new_run)
        parts.append(new_p)
    return parts


def _para_text_join(p):
    return "".join(t.text or "" for t in p.iter(f"{W}t"))


def _has_forbidden_descendant(p):
    """段内含域/书签/修订/嵌入对象等不可见结构 → 放弃。"""
    for el in p.iter():
        tag = etree.QName(el).localname
        if tag in ("fldChar", "instrText", "bookmarkStart", "bookmarkEnd",
                   "commentRangeStart", "commentRangeEnd", "ins", "del",
                   "drawing", "object", "pict", "sectPr"):
            return True
    return False


def _body_paragraphs(root):
    body = root.find(f"{W}body")
    if body is None:
        return []
    return [el for el in body if etree.QName(el).localname == "p"]


def _norm_head_key(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "").strip()


# ── 1) 行内序号列表化 ───────────────────────────────────────────

def _enum_zone_flags(body_ps):
    """逐段「允许列表化」标记：正文区（首见「第 X 章」之后、参考文献/
    致谢/附录类标题之前）内才允许；其余（前置件/目录/摘要/后部）全部跳过。"""
    allowed = [False] * len(body_ps)
    in_body = False
    for i, p in enumerate(body_ps):
        text = _para_text_join(p)
        stripped = text.strip()
        norm = _norm_head_key(text)
        if _CHAPTER_RE.match(stripped):
            in_body = True
            continue
        if norm and (norm.startswith(("参考文献", "致谢", "附录", "作者简介"))
                     or norm.lower() in ("references", "acknowledgment",
                                         "acknowledgements",
                                         "acknowledgement", "background")):
            in_body = False
            continue
        allowed[i] = in_body
    return allowed


def _enum_plan(p):
    """判定段落是否为可拆的行内序号枚举；返回 (specs, item_count) 或 None。"""
    if _has_forbidden_descendant(p):
        return None
    text = _para_text_join(p)
    if not text or "\t" in text:
        return None
    if _DOT_LEADER_RE.search(text):
        return None  # 目录式引导点
    if _segments(p) is None:
        return None
    markers = [(m.start(), m.end(), int(m.group(1)))
               for m in _ENUM_MARK_RE.finditer(text)]
    if len(markers) < 2:
        return None
    nums = [n for _, _, n in markers]
    if nums != list(range(1, len(nums) + 1)):
        return None
    lead = text[:markers[0][0]]
    if not _COLON_END_RE.search(lead.rstrip()):
        return None
    if len(lead.strip()) < 2 or not _CJK_RE.search(lead):
        return None

    total = len(text)
    specs = []
    lead_end = len(lead.rstrip())
    if lead_end <= 0:
        return None
    specs.append((0, lead_end, None))
    n_items = 0
    for i, (ms, me, _n) in enumerate(markers):
        nxt = markers[i + 1][0] if i + 1 < len(markers) else total
        if text[ms] not in "（(":
            return None
        end = nxt
        while end > ms and text[end - 1] in _WS_CHARS:
            end -= 1
        drop = None
        if me < end and text[me] in _WS_CHARS:
            ce = me
            while ce < end and text[ce] in _WS_CHARS:
                ce += 1
            drop = [(me, ce)]
        content_start = drop[0][1] if drop else me
        if end <= content_start or not text[content_start:end].strip():
            return None
        specs.append((ms, end, drop))
        n_items += 1
    # 尾句拆分：最后一项内首个「。」之后仍有非空白内容 → 独立成段
    lms, lend, ldrop = specs[-1]
    j = text.find("。", lms, lend)
    if j >= 0 and text[j + 1:total].strip():
        specs[-1] = (lms, j + 1, ldrop)
        ts = j + 1
        while ts < total and text[ts] in _WS_CHARS:
            ts += 1
        if ts < total:
            specs.append((ts, total, None))
    return specs, n_items


def _split_enumerations(root, stats):
    body_ps = _body_paragraphs(root)
    allowed = _enum_zone_flags(body_ps)
    n_splits = 0
    for i, p in enumerate(body_ps):
        if not allowed[i]:
            continue
        plan = _enum_plan(p)
        if plan is None:
            continue
        specs, n_items = plan
        parts = _slice_into_paragraphs(p, specs)
        if parts is None:
            stats["skipped"].append(_para_text_join(p)[:40])
            continue
        parent = p.getparent()
        for np_ in parts:
            p.addprevious(np_)
        parent.remove(p)
        n_splits += 1
        stats["items"] += n_items
        stats["previews"].append(_para_text_join(parts[0])[:36])
    stats["split"] = n_splits


# ── 2) 附录标题行内正文拆分 ─────────────────────────────────────

def _appendix_zone_flags(body_ps):
    """附录拆分允许区：参考文献/致谢类标题之后，或（无此类标题时）文档
    末尾 15 段内——防正文中以「附录 A …」开头的句子被误拆。"""
    n = len(body_ps)
    back_start = None
    for i, p in enumerate(body_ps):
        norm = _norm_head_key(_para_text_join(p))
        if not norm:
            continue
        low = norm.lower()
        if norm.startswith(("参考文献", "致谢")) \
                or low.startswith(("references", "acknowledg")):
            back_start = i
    flags = [False] * n
    for i in range(n):
        if back_start is not None:
            flags[i] = i > back_start
        else:
            flags[i] = i >= max(0, n - 15)
    return flags


def _appendix_plan(p):
    if _has_forbidden_descendant(p):
        return None
    text = _para_text_join(p)
    if not text:
        return None
    if _segments(p) is None:
        return None
    m = _APPENDIX_SPLIT_RE.match(text)
    if not m:
        return None
    rest = text[m.end():]
    if not rest.strip():
        return None
    first = rest.strip()[0]
    if first in "：:。.；;，,、":
        return None
    head_end = m.end()
    while head_end > 0 and text[head_end - 1] in _WS_CHARS:
        head_end -= 1
    rest_start = m.end()
    while rest_start < len(text) and text[rest_start] in _WS_CHARS:
        rest_start += 1
    if not text[rest_start:].strip() or not text[:rest_start].strip():
        return None
    return [(0, head_end, None), (rest_start, len(text), None)]


def _split_appendix(root, stats):
    body_ps = _body_paragraphs(root)
    flags = _appendix_zone_flags(body_ps)
    n_splits = 0
    for i, p in enumerate(body_ps):
        if not flags[i]:
            continue
        plan = _appendix_plan(p)
        if plan is None:
            continue
        parts = _slice_into_paragraphs(p, plan)
        if parts is None:
            stats["skipped"].append(_para_text_join(p)[:40])
            continue
        parent = p.getparent()
        for np_ in parts:
            p.addprevious(np_)
        parent.remove(p)
        n_splits += 1
        stats["previews"].append(
            f"{_para_text_join(parts[0])[:16]} | {_para_text_join(parts[1])[:24]}")
    stats["split"] = n_splits


# ── 主入口 ──────────────────────────────────────────────────────

def normalize_text_structure_in_docx(path) -> dict:
    """就地归一一个 .docx（返回值全部为计数，供汇报/排查）。"""
    path = Path(path)
    stats = {
        "changed": False,
        "enumerations": {"split": 0, "items": 0, "previews": [], "skipped": []},
        "appendix": {"split": 0, "previews": [], "skipped": []},
    }
    with zipfile.ZipFile(path) as z:
        try:
            data = z.read("word/document.xml")
        except KeyError:
            return stats
    try:
        root = etree.fromstring(data)
    except etree.XMLSyntaxError:
        return stats

    _split_enumerations(root, stats["enumerations"])
    _split_appendix(root, stats["appendix"])

    if not (stats["enumerations"]["split"] or stats["appendix"]["split"]):
        return stats
    stats["changed"] = True
    new_doc = etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                             standalone=True)
    tmp = path.with_suffix(".textsplit.tmp")
    with zipfile.ZipFile(path) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            payload = new_doc if item.filename == "word/document.xml" \
                else zin.read(item.filename)
            zout.writestr(item, payload)
    shutil.move(str(tmp), str(path))
    return stats


def normalize_text_structure_in_copy(src, workdir) -> tuple[Path, dict]:
    """工作副本上的归一（上游若已复制，本步就地处理）；无改动原样返回。"""
    src = Path(src)
    workdir = Path(workdir)
    workdir.mkdir(parents=True, exist_ok=True)
    copy_path = workdir / src.name
    if src.resolve() == copy_path.resolve():
        target = src  # 已是工作副本（latex 步已复制）→ 就地
    else:
        shutil.copy2(src, copy_path)
        target = copy_path
    stats = normalize_text_structure_in_docx(target)
    if not stats["changed"] and target != src:
        target.unlink(missing_ok=True)
        return src, stats
    return target, stats
