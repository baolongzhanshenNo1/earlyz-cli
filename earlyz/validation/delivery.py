"""早睡 (EarlyZ) 3.2 — Delivery Validator（§8.10 / §11，ADR-014）。

四层门禁：
1. Static Package Gate   — ZIP/ContentTypes/rels/重复 entry/缺失 part（复用 ir.package）
2. OOXML Invariant Gate  — docPr 唯一、字段完整、关键 ID 唯一
3. Semantic Format Gate  — 内容保持（content hash 对比）、FormatSpec 规则（可选）
4. Render Gate           — Word COM → PDF/分页/bbox（Windows）；无则标注验证等级

验证等级（§11.2）：
LEVEL_A_WORD_VERIFIED（真实 Word 通过）
LEVEL_B_FALLBACK_RENDERED
LEVEL_C_STATIC_ONLY
FAILED

只有 LEVEL_A 才能声明「已通过真实 Word 验收」。
"""

from __future__ import annotations

import hashlib
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

from earlyz.ir.analyzer import analyze_document
from earlyz.ir.package import inspect_package
from earlyz.validation.renderers import RendererUnavailableError, _table_pagination_pass
from earlyz.ir.complex_objects import (
    complex_object_gate_errors,
    complex_object_gate_warnings,
)

LEVEL_A_WORD_VERIFIED = "LEVEL_A_WORD_VERIFIED"
LEVEL_B_FALLBACK_RENDERED = "LEVEL_B_FALLBACK_RENDERED"
LEVEL_C_STATIC_ONLY = "LEVEL_C_STATIC_ONLY"
FAILED = "FAILED"

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


@dataclass
class DeliveryReport:
    input_path: str
    level: str = FAILED
    passed: bool = False
    hard_errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    content_preserved: Optional[bool] = None
    renderer_used: Optional[str] = None
    certificate: dict = field(default_factory=dict)

    @property
    def hard_error_count(self) -> int:
        return len(self.hard_errors)

    def to_dict(self) -> dict:
        return {
            "input_path": self.input_path,
            "validation_level": self.level,
            "passed": self.passed,
            "hard_errors": self.hard_errors,
            "warnings": self.warnings,
            "content_preserved": self.content_preserved,
            "renderer_used": self.renderer_used,
            "certificate": self.certificate,
        }


def _file_hash(path: str) -> str:
    return hashlib.sha1(Path(path).read_bytes()).hexdigest()


def _count_header_parts(path) -> int:
    """源文档自带的页眉部件数（0 = 文档原本没有页眉）。"""
    import re as _re
    import zipfile as _zip
    try:
        with _zip.ZipFile(str(path)) as z:
            return len([n for n in z.namelist() if _re.match(r"word/header\d+\.xml$", n)])
    except Exception:  # noqa: BLE001
        return 0


def _cover_residual_hints(ir, spec=None) -> list[str]:
    """封面残留模板样例文本检测（层 2.6，warning 级）。

    2026-08-16 起：spec.example_residuals（理解层提取的模板示例教学内容）
    存在时 → 规范化匹配目标文档段落，未清除的示例残留逐一列出；
    无 spec（独立调用）→ 回退旧启发式（封面区 ≥2 标题形状段互不包含）。
    """
    import re as _re

    from earlyz.spec.decompiler import (
        _DATE_RE, _CHAPTER_RE, _field_name, _cjk_ratio, find_cover_end,
        norm_text, _norm_residual,
    )
    if spec is not None and getattr(spec, "example_residuals", None):
        resid = set(spec.example_residuals)
        hits = []
        for p in ir.paragraphs:
            t = (p.text or "").strip()
            if not t:
                continue
            if _norm_residual(t) in resid:
                hits.append(t[:40])
        return hits
    cover_end = find_cover_end(ir)
    _front_fields = ("关键词", "关键字", "摘要", "abstract", "indexterms", "index terms",
                     "中图分类号", "文献标识码", "文章编号", "作者简介")
    titles: list[str] = []
    for i, p in enumerate(ir.paragraphs[:cover_end]):
        t = (p.text or "").strip()
        if not t:
            continue
        n = norm_text(t)
        # 期刊首页的「关键词：…」「摘要：…」也是短中文行，会被当成第二个标题 → 误报残留
        if any(n.startswith(k) for k in _front_fields):
            continue
        if (_cjk_ratio(t) >= 0.5 and 8 <= len(n) <= 30
                and not _field_name(t) and not _CHAPTER_RE.match(n)
                and not _DATE_RE.match(t)):
            titles.append(t)
    if len(titles) < 2:
        return []
    # 两行标题互不包含 → 疑似残留（真标题与副标题通常有包含/连续关系）
    hints = []
    for a in titles[:-1]:
        for b in titles[1:]:
            na, nb = norm_text(a), norm_text(b)
            if na not in nb and nb not in na and na != nb:
                hints.append(f"「{a[:24]}」与「{b[:24]}」")
                break
    return hints[:3]


def _check_duplicate_docpr(document_xml: str) -> list[str]:
    """w:drawing 的 docPr id 必须唯一（曾因重复 docPr 导致 Word 异常）。"""
    errors = []
    root = ET.fromstring(document_xml)
    seen: dict[str, str] = {}
    for docpr in root.iter(f"{WP}docPr"):
        did = docpr.get("id")
        if did is None:
            continue
        if did in seen:
            errors.append(f"duplicate docPr id={did}")
        seen[did] = ""
    return errors


def _read_style_names(path: str) -> dict[str, str]:
    """styleId → style name（角色判定需要 name：styleId=1 → "heading 1"）。"""
    import lxml.etree as LET   # 与 _cover_zone_snapshot 同款局部导入

    names: dict[str, str] = {}
    try:
        with zipfile.ZipFile(path) as z:
            root = LET.fromstring(z.read("word/styles.xml"))
    except (KeyError, OSError, ValueError):
        return names
    for s in root.findall(f"{W}style"):
        sid = s.get(f"{W}styleId")
        nm = s.find(f"{W}name")
        if sid and nm is not None:
            names[sid] = nm.get(f"{W}val", "") or sid
    return names


def _cover_zone_snapshot(path: str) -> tuple[int, list[dict], list[list[dict]]]:
    """封面区快照：逐段（文本 + 正文格式特征）与封面节页眉/页脚引用。

    边界判据与写动作的封面守卫同一套（`new_actions._compute_cover_boundary`，
    它内部又对齐 lint 的 `find_cover_end`）——单一真相，不另立一套。
    返回 (封面段数, 非空段快照, 封面节的 header/footer part 列表)。
    """
    from earlyz.actions.new_actions import _compute_cover_boundary
    import lxml.etree as LET   # 边界判据用 iterancestors（lxml 特性）

    with zipfile.ZipFile(path) as z:
        root = LET.fromstring(z.read("word/document.xml"))
        try:
            rels = LET.fromstring(z.read("word/_rels/document.xml.rels"))
        except KeyError:
            rels = None
        parts: dict = {}
        if rels is not None:
            for rel in rels:
                rid = rel.get("Id")
                tgt = (rel.get("Target") or "").lstrip("/")
                if not rid or not tgt:
                    continue
                name = tgt if tgt.startswith("word/") else f"word/{tgt}"
                try:
                    parts[rid] = z.read(name)
                except KeyError:
                    continue
        try:
            settings = LET.fromstring(z.read("word/settings.xml"))
        except (KeyError, LET.XMLSyntaxError):
            settings = None
    # B 轴按**有效引用**判（2026-09-29）：first 仅在 titlePg 下显示、even 仅在
    # settings 全局 evenAndOddHeaders 下显示；不显示的部件不算封面书眉。
    even_odd = _on_off(settings.find(f"{W}evenAndOddHeaders")
                       if settings is not None else None)

    body = root.find(f"{W}body")
    if body is None:
        return 0, [], []
    n_cover = _compute_cover_boundary(body)
    paras = [p for p in body.iter(f"{W}p")
             if not any(a.tag == f"{W}drawing" for a in p.iterancestors())]
    # 前部角色（期刊题名区判定用，2026-09-27）：与 analyzer 的相位机同源，禁第二套词表
    from earlyz.ir.analyzer import FrontRoleAssigner, classify_paragraph_role
    style_names = _read_style_names(path)
    assigner = FrontRoleAssigner()
    rows: list[dict] = []
    chrome: list[list[dict]] = []
    # 节未声明某类型引用时沿用前一节（OOXML 继承语义），按文档顺序累积
    inherited: dict[tuple[str, str], str] = {}
    for idx, p in enumerate(paras):
        if idx >= n_cover:
            break
        text = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        ppr = p.find(f"{W}pPr")
        sect = ppr.find(f"{W}sectPr") if ppr is not None else None
        if sect is not None:
            for tag in ("headerReference", "footerReference"):
                for r in sect.findall(f"{W}{tag}"):
                    rid = r.get(f"{R}id")
                    if rid:
                        inherited[(tag, r.get(f"{W}type") or "default")] = rid
            title_pg = _on_off(sect.find(f"{W}titlePg"))
            refs = []
            for (tag, rtype), rid in inherited.items():
                if (rtype == "first" and not title_pg) or (
                        rtype == "even" and not even_odd):
                    continue  # Word 不显示该部件
                refs.append({"kind": tag, "type": rtype, "xml": parts.get(rid)})
            if refs:
                chrome.append(refs)
        if not text:
            continue
        jc = ppr.find(f"{W}jc") if ppr is not None else None
        ind = ppr.find(f"{W}ind") if ppr is not None else None
        first = hanging = None
        if ind is not None:
            # 口径（2026-09-27）：w:hanging 是**悬挂**缩进，不是首行缩进——单位行/编号行
            # 的悬挂缩进是模板值，曾被误报成「被套上正文首行缩进 105」。
            first = (ind.get(f"{W}firstLine") or ind.get(f"{W}firstLineChars"))
            # 悬挂缩进单独成轴（2026-09-29 #19 漏报：学位论文封面「培养单位」被套上
            # 悬挂缩进，旧口径只看首行缩进 → 门放行）
            hanging = (ind.get(f"{W}hanging") or ind.get(f"{W}hangingChars"))
        pstyle = ppr.find(f"{W}pStyle") if ppr is not None else None
        sid = pstyle.get(f"{W}val", "") if pstyle is not None else ""
        sname = style_names.get(sid, "") or sid
        numpr = ppr is not None and ppr.find(f"{W}numPr") is not None
        role = classify_paragraph_role(ppr, sname, False, False, numpr, text)
        role = assigner.feed(text, sname, role)
        rows.append({
            "text": text,
            "jc": (jc.get(f"{W}val") if jc is not None else None),
            "first_line": _int_or_none(first),
            "hanging": _int_or_none(hanging),
            "role": role,
        })
    return n_cover, rows, chrome


def _int_or_none(v):
    try:
        return int(str(v).strip())
    except (TypeError, ValueError):
        return None


def _norm_ws(s: str) -> str:
    """文本比对口径：去掉全部空白（含全角空格/制表符），只看字符序列。"""
    return "".join((s or "").split()).replace("\u3000", "")


def _on_off(el) -> bool:
    """OOXML CT_OnOff：元素存在且 w:val 不是 false/0/off 即为真。"""
    if el is None:
        return False
    return (el.get(f"{W}val") or "true").lower() not in ("false", "0", "off")


def _chrome_is_dirty(refs: list[dict]) -> str:
    """封面节的页眉/页脚是否「写脏」：可见文本 或 PAGE 域 → 返回原因。"""
    for ref in refs:
        xml = ref.get("xml")
        if not xml:
            continue
        try:
            node = ET.fromstring(xml)
        except ET.ParseError:
            continue
        kind = "页眉" if ref["kind"] == "headerReference" else "页脚"
        body = ""
        for p in node.iter(f"{W}p"):
            body += "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        instr = " ".join(t.text or "" for t in node.iter(f"{W}instrText"))
        if "PAGE" in instr.upper() or "页码" in body:
            return f"{kind}含页码域"
        if body:
            return f"{kind}有文本「{body[:12]}」"
    return ""


def _zone_is_article_head(rows: list[dict]) -> bool:
    """封面区是否为「期刊题名区」（article head）——是则封面保全门 A/B 轴都不适用。

    判据委托 `analyzer.zone_is_article_head`（2026-09-29：交付门 / writer / lint 同一谓词，
    防写端豁免与门判据漂移）。只按**源稿**判定：源稿是无格式原稿，角色是稳定证据；
    产物的角色会被已写入的样式（如 article_title 规则带 outlineLvl）影响。
    """
    from earlyz.ir.analyzer import zone_is_article_head

    return zone_is_article_head(r.get("role") for r in rows)


_PUNCT_WIDTH = str.maketrans({
    "（": "(", "）": ")", "，": ",", "：": ":", "；": ";", "？": "?", "！": "!",
    "【": "[", "】": "]", "“": '"', "”": '"', "‘": "'", "’": "'", "、": ",",
    "．": ".", "　": " ",
})


def _norm_punct(text: str) -> str:
    """全半角/中英标点等价归一 —— 只用于判断「文本是否真的被改」。"""
    return "".join(str(text).split()).translate(_PUNCT_WIDTH)


def _cover_preservation_errors(source_path: str, product_path: str,
                              check_text: bool = True) -> list[str]:
    """封面保全门（2026-09-25 横向机制审计发现 1）。

    真机缺陷：封面守卫写好了、行为也修好了，但门禁对「封面被改」完全盲——
    把守卫关掉跑真实编译，封面「姓 名:」行被 `jc=None→center`、首行缩进
    `1680→None`，门禁仍报 `success=True / LEVEL_A_WORD_VERIFIED /
    hard_errors=[]`。这一层专看封面三件事（用户 2026-09-25 拍板 B）：

      A 轴 封面区非空段不得凭空长出一正文格式特征：首行缩进、两端对齐、悬挂缩进；
          文本不得被改（2026-09-29：标签等价口径把「中图分类号：」改写成模板标签）；
      B 轴 封面节（含封面性质节）**实际显示**的页眉/页脚不得有可见文本或页码域
          （first 仅 titlePg 生效、even 仅 evenAndOddHeaders 生效，未声明沿用前一节）。

    白名单（用户拍板）：空段计数差异（节尾清理）、段内 sectPr 与页边距/
    节属性（页面设置继续按模板封面节取值=现状）。
    """
    errors: list[str] = []
    src_n, src_rows, _ = _cover_zone_snapshot(source_path)
    out_n, out_rows, out_chrome = _cover_zone_snapshot(product_path)
    if src_n <= 0 or out_n <= 0:
        return errors  # 无封面区，本层不适用（期刊/纯正文文档）
    if _zone_is_article_head(src_rows):
        # 期刊题名区（2026-09-27）：题名/作者/单位/摘要/关键词的前部格式**由模板规则下发**
        # （题名 firstLine、单位悬挂缩进、摘要/关键词两端对齐、书眉含页码域），不是「封面被
        # 正文化」；08-20 口径即「期刊首页题名/作者/单位/摘要按模板前部取值，不当学位论文
        # 封面跳过」。A 轴（段落格式）与 B 轴（页眉页脚）都不适用。
        return errors
    # 2026-10-14 审计：段数变（合法的节尾/残留清理）时逐段格式比对会整体
    # 错位（用第 i+1 源段比第 i 产物段）→ 误报封面污染。三条格式轴与下方
    # 文本轴同口径：段数一致才逐段比（白名单：空段计数差异）。
    aligned = len(src_rows) == len(out_rows)
    for i, (a, b) in enumerate(zip(src_rows, out_rows)):
        if aligned and (b["first_line"] or 0) > 0 and (a["first_line"] or 0) <= 0:
            errors.append(
                f"cover polluted: 封面第 {i + 1} 个非空段「{b['text'][:12]}」"
                f"被套上正文首行缩进 {b['first_line']}（原稿无）"
            )
        if aligned and b["jc"] == "both" and a["jc"] != "both":
            errors.append(
                f"cover polluted: 封面第 {i + 1} 个非空段「{b['text'][:12]}」"
                f"被套上正文两端对齐（原稿 {a['jc']}）"
            )
        if aligned and (b.get("hanging") or 0) > 0 and (a.get("hanging") or 0) <= 0:
            errors.append(
                f"cover polluted: 封面第 {i + 1} 个非空段「{b['text'][:12]}」"
                f"被套上悬挂缩进 {b['hanging']}（原稿无）"
            )
        # 段数一致才逐段比文本（段数变 = 合法的节尾/残留清理，对不齐时不硬判）
        # check_text=False：封面已按模板重建（2026-10-22 定案的设计行为）——
        # 模板封面写「学院（系）:」、源文档写「学院(系):」，文本必然不同，
        # 拿源文档当基准比文字只会把设计行为判成污染。格式轴仍然生效。
        # 只差全半角/中英标点的行不算「文本被改」：封面槽位填充与模板封面重建
        # 本来就会把「学院(系):」写成「学院（系）:」，这是设计行为，不是污染。
        # 真篡改（增删字词）仍然会被下面的比较抓到。
        if (check_text and aligned
                and _norm_punct(a["text"]) != _norm_punct(b["text"])
                and _norm_ws(a["text"]) != _norm_ws(b["text"])):
            errors.append(
                f"cover polluted: 封面第 {i + 1} 个非空段文本被改"
                f"「{a['text'][:16]}」→「{b['text'][:16]}」"
            )
    for refs in out_chrome:
        why = _chrome_is_dirty(refs)
        if why:
            errors.append(f"cover polluted: 封面节页眉/页脚未清干净——{why}")
    return errors


_VML_NS = "urn:schemas-microsoft-com:vml"
_GEOM_ATTRS = ("style", "width", "height", "coordsize", "coordorigin")


def _object_signature(el) -> str:
    """w:object 的内容签名：剔除 VML 外框几何，只比对象本身。"""
    import copy as _copy  # noqa: PLC0415 — 与模块内其他 lxml 使用保持一致，函数内导入
    import lxml.etree as _LET  # noqa: PLC0415
    try:
        cp = _copy.deepcopy(el)
        for node in cp.iter():
            try:
                ns = _LET.QName(node).namespace
            except ValueError:
                continue
            if ns == _VML_NS:
                for k in _GEOM_ATTRS:
                    node.attrib.pop(k, None)
        return _LET.tostring(cp)
    except Exception:  # noqa: BLE001 — 取签名失败退回原始串
        return _LET.tostring(el)



def _equation_preservation_errors(source_path: str, product_path: str,
                                  warnings: list | None = None) -> list[str]:
    """层 3.5b：公式/上标保全门（2026-09-27）。

    背景：交付门原先只比对**文本**（content_preserved），数学与上标不在口径内。
    真机事故：模板 extras 回放把作者语义（`w:vertAlign`）当未建模属性清空 →
    真论文 139 处引用上标与作者角标 139→0，交付门全程绿灯；MathType 公式段
    被正文首行缩进右推，也没人管。本层把「公式/上标没被弄坏」变成可证明的事实。

    口径（源 vs 产物）：
      ① `word/embeddings/*` 部件字节一致（MathType OLE 载荷；无动作合法改动它）
      ② `w:object` 子树序列化一致（VML 形状 + OLE 引用）
      ③ `m:oMath` / `m:oMathPara` 子树序列化一致（OMML；公式动作只写段落属性，不进子树）
      ④ 上标/下标计数不得**减少**（作者语义；增加允许——LaTeX 归一动作会新增）
    """
    import hashlib as _hashlib
    import zipfile as _zipfile
    import lxml.etree as LET

    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"

    def _snapshot(path: str):
        with _zipfile.ZipFile(path) as z:
            # 只算真实部件：zip 目录条目（以 / 结尾）不是 OLE 载荷，
            # 重打包不写目录项，比较时会误报「OLE 部件丢失」（2026-10-19 南航模板实测）
            emb = {n: _hashlib.sha256(z.read(n)).hexdigest()
                   for n in z.namelist()
                   if n.startswith("word/embeddings/") and not n.endswith("/")}
            root = LET.fromstring(z.read("word/document.xml"))
        # w:object 的「内容」是 OLE payload / 图形数据；v:shape 的 style 尺寸只是外框。
        # 媒体缩放会改外框（2026-10-19 期刊模板封顶 439.2pt→252pt），
        # 若连外框一起比，会把合法缩放判成篡改。比较前剔除几何属性。
        objects = [_object_signature(el) for el in root.iter(f"{W}object")]
        maths = [LET.tostring(el) for el in root.iter(f"{M}oMathPara")]
        maths += [LET.tostring(el) for el in root.iter(f"{M}oMath")
                  if not any(True for _ in el.iterancestors(f"{M}oMathPara"))]
        sup = sub = 0
        for r in root.iter(f"{W}r"):
            va = r.find(f"{W}rPr/{W}vertAlign")
            if va is None:
                continue
            sup += va.get(f"{W}val") == "superscript"
            sub += va.get(f"{W}val") == "subscript"
        return emb, objects, maths, sup, sub

    errors: list[str] = []
    s_emb, s_obj, s_math, s_sup, s_sub = _snapshot(source_path)
    p_emb, p_obj, p_math, p_sup, p_sub = _snapshot(product_path)

    def _added_note(kind: str, n: int) -> None:
        # 2026-10-14 审计：公式「从文本/带 → 对象」是链内受控转换动作的产物
        # （§10.1 授权），新增不算违规、必须让用户看见（与「公式转换通报」
        # 同一哲学）；缺失/漂移仍为硬错误。
        msg = f"公式保全门：新增 {kind} {n} 个（公式转换动作产生，属受控变化）"
        if warnings is not None:
            warnings.append(msg)

    def _loss_add(src, prod):
        """列表快照的缺失/新增计数；元素不可哈希时退回长度差（保守）。"""
        try:
            from collections import Counter
            cs, cp = Counter(src), Counter(prod)
            return sum((cs - cp).values()), sum((cp - cs).values())
        except TypeError:
            if src == prod:
                return 0, 0
            if len(src) > len(prod):
                return len(src) - len(prod), 0
            if len(prod) > len(src):
                return 0, len(prod) - len(src)
            return 1, 0  # 同长但内容不同 → 视为漂移（保守判错）

    if s_emb != p_emb:
        missing = sorted(set(s_emb) - set(p_emb))[:3]
        added = sorted(set(p_emb) - set(s_emb))[:3]
        changed = sorted(n for n in set(s_emb) & set(p_emb) if s_emb[n] != p_emb[n])[:3]
        if missing or changed:
            errors.append("公式保全门：OLE 部件被改动"
                          f"（missing={missing} added={added} changed={changed}）")
        if added:
            _added_note("OLE 部件", len(set(p_emb) - set(s_emb)))
    if s_obj != p_obj:
        _miss, _add = _loss_add(s_obj, p_obj)
        if _miss:
            errors.append("公式保全门：w:object 子树不一致"
                          f"（源 {len(s_obj)} 个 → 产物 {len(p_obj)} 个）")
        if _add:
            _added_note("w:object 子树", _add)
    if s_math != p_math:
        _miss, _add = _loss_add(s_math, p_math)
        if _miss:
            errors.append("公式保全门：OMML 子树不一致"
                          f"（源 {len(s_math)} 个 → 产物 {len(p_math)} 个）")
        if _add:
            _added_note("OMML 子树", _add)
    if p_sup < s_sup:
        errors.append(f"公式保全门：上标被清（源 {s_sup} → 产物 {p_sup}）")
    if p_sub < s_sub:
        errors.append(f"公式保全门：下标被清（源 {s_sub} → 产物 {p_sub}）")
    return errors


def _heading_shape_errors(product_path: str, warnings: list | None = None) -> list[str]:
    """标题形态门（2026-09-26 用户「识别层得加强一下」）。

    真机缺陷：识别层把 30 段序号开头的正文句判成 heading1 → 模板标题取值下发
    （18pt 加粗居中 + outlineLvl + 段前分页）→ 全篇带大纲段从 31 涨到 52、
    导航窗格与目录层级失真；而所有既有门禁都报成功（行为层"一切正常"）。

    本层**不信角色、只看产物格式证据**（与识别判据彼此独立，互为交叉验证）：
      轴 A 段带 w:outlineLvl（进导航/进目录的那批）
      轴 B 段居中对齐且直系字号 ≥ 14pt（标题级视觉）
    命中任一轴者，其文本若含「。」或「；」即判**句子形态** → 硬错误。

    阈值刻意保守（真标题几乎不含句号/分号，含「，」「：」的标题合法故不判）：
    宁可漏报，不误杀长标题或期刊式标题。
    """
    errors: list[str] = []
    warnings = warnings if warnings is not None else []
    _hits: list[str] = []
    try:
        with zipfile.ZipFile(product_path) as z:
            body = ET.fromstring(z.read("word/document.xml")).find(f"{W}body")
    except (KeyError, ET.ParseError, zipfile.BadZipFile):
        return errors
    if body is None:
        return errors

    def _direct_size(p) -> int:
        sizes = []
        for r in p:
            if r.tag != f"{W}r":
                continue
            for sz in r.iter(f"{W}sz"):
                try:
                    sizes.append(int(sz.get(f"{W}val") or 0))
                except ValueError:
                    pass
        ppr = p.find(f"{W}pPr")
        if ppr is not None:
            rpr = ppr.find(f"{W}rPr")
            if rpr is not None:
                for sz in rpr.iter(f"{W}sz"):
                    try:
                        sizes.append(int(sz.get(f"{W}val") or 0))
                    except ValueError:
                        pass
        return max(sizes) if sizes else 0

    for i, p in enumerate(body.iter(f"{W}p")):
        text = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
        if not text or ("。" not in text and "；" not in text):
            continue
        # 目录条目（带 PAGEREF）不适用：条目本就带页码尾巴
        if any((x.text or "").strip().startswith("PAGEREF")
               for x in p.iter(f"{W}instrText")):
            continue
        ppr = p.find(f"{W}pPr")
        axis = ""
        n_char = len(text)
        if ppr is not None and ppr.find(f"{W}outlineLvl") is not None:
            # 短标题带句号属正常（「致谢。」「结论。」只有 2-3 字）。
            # 2026-10-25 全量实测：3 字「致谢。」被判句子形态 → 整篇拒绝交付。
            if n_char >= 15:
                axis = "带大纲级别（进导航/目录）"
        elif ppr is not None:
            jc = ppr.find(f"{W}jc")
            # 居中大字号：只对「长句 + 句末是。」成立才判——模板注解行
            # （「奇数页页眉：Times New Roman，五号）」）也居中大字号，但句末是右括号。
            if (jc is not None and (jc.get(f"{W}val") or "") == "center"
                    and _direct_size(p) >= 28 and n_char >= 30
                    and text.rstrip().endswith("。")):
                axis = f"居中 + {_direct_size(p) / 2:.0f}pt 标题级字号"
        if axis:
            _hits.append(
                f"第 {i + 1} 段被判为标题层级（{axis}）但文本是句子形态"
                f"——「{text[:24]}…」（共 {len(text)} 字）"
            )
    # 零星命中降级为提示（不进 hard_errors）：该门的目的是抓「识别层批量把正文句
    # 刷成标题」（30 段级），而不是拦单段。2026-10-26 真实场景实测：学位论文封面
    # 说明行、摘要里居中 14~16pt 的句子被单点命中 → 整篇拒绝交付。该门自身注释
    # 口径即「宁可漏报，不误杀」。≥3 段仍判硬错误（系统性误升的特征）。
    if len(_hits) >= 3:
        errors.extend(f"heading shape: {h}" for h in _hits)
    elif _hits:
        for _h in _hits:
            warnings.append(f"heading shape（已降级，不阻断）: {_h}")
    return errors


def _check_field_integrity(ir, warnings: list | None = None) -> list[str]:
    """域完整性（2026-10-14 审计降噪）：缺 end 的域分两档——

    ① 既缺 end 又无结果文本 → 硬错误（真残缺）；
    ② 仅缺 end 但有结果文本 → warning（Word 常仍可显示；解析口径曾整体
       误判过整篇，见 analyzer._FieldTracker 的事故注释）。域指令附在
       消息里便于定位。fldSimple 形态不在本 tracker 口径（已知差异）。
    """
    errors = []
    for f in ir.fields:
        if f.complete:
            continue
        result_text = (getattr(f, "result", "") or "").strip()
        instr = (getattr(f, "instruction", "") or "").strip()
        detail = f" instr={instr[:40]}" if instr else ""
        if result_text:
            msg = (f"broken field (unclosed fldChar begin, 结果仍在): "
                   f"{f.node_id}{detail}")
            if warnings is not None:
                warnings.append(msg)
            else:
                errors.append(msg)
        else:
            errors.append(
                f"broken field (unclosed fldChar begin, 无结果): {f.node_id}{detail}")
    return errors


def _check_render_map(ir) -> list[str]:
    """P0：渲染门禁的静态前置（完整 Word 渲染由外部 renderer 提供）。"""
    return []


def _render_retry_then_fallback(input_path: str, word_render_callback, first_err,
                                warnings: list, errors: list):
    """Word 渲染失败后的事后判定（2026-10-06 方案 B，用户拍板）。

    顺序：① 重试一次（瞬时故障大多一发即过）→ ② LibreOffice 兜底渲染复核 →
    ③ 双渲染器都失败才置硬错误（文档真实损坏的证据）。返回 render_info 或 None：
    重试成功=word（LEVEL_A）；LO 兜底成功=libreoffice（LEVEL_B）；
    无任何渲染器可用=降级（LEVEL_C，环境无法复核≠文档有问题）。
    """
    # ① 重试一次
    try:
        info = word_render_callback(input_path)
        warnings.append(
            f"Word 渲染首次失败（{type(first_err).__name__}），重试后通过——"
            "瞬时故障已留痕，可放心使用")
        if info.get("bbox_overflow"):
            errors.append(f"image/table bbox overflow: {len(info['bbox_overflow'])} items {_bbox_detail(info['bbox_overflow'])}")
        return info
    except Exception as retry_err:  # noqa: BLE001 — 再失败才走兜底
        second_err = retry_err

    # ② LibreOffice 兜底渲染复核
    try:
        from earlyz.validation.providers import render_with_provider
        r = render_with_provider("libreoffice", input_path)
    except Exception as lo_err:  # noqa: BLE001
        # 分水岭：「本机没有 LibreOffice」= 环境无法复核 → 降级交付；
        # 「LibreOffice 也渲染失败」= 双渲染器实证文档损坏 → 硬拒。
        from earlyz.validation.providers import _lo_available
        ok, why = _lo_available()
        if not ok:
            warnings.append(
                f"Word 渲染两次失败（{type(second_err).__name__}），本机也没有 "
                f"LibreOffice 可复核（{why}）——按环境问题降级交付（{LEVEL_C_STATIC_ONLY}）。"
                "不放心可重跑一次本次排版")
            return None
        # 复核组件缺失（如 pypdfium2/PDFium 未随环境安装）= 环境问题，不是文档证据
        # → 降级交付（2026-10-06 实测：托管 venv 缺 pypdfium2 时 Word/LO 两条路
        # 同点崩溃，曾被误判为「双渲染器失败=文档损坏」）
        for _e in (second_err, lo_err):
            if isinstance(_e, ImportError):
                warnings.append(
                    f"渲染复核组件缺失（{type(_e).__name__}: {_e}）——按环境问题"
                    f"降级交付（{LEVEL_C_STATIC_ONLY}）；修复组件后重跑可获得完整渲染复核")
                return None
        errors.append(
            f"render failed: word: {type(second_err).__name__}: {second_err}; "
            f"libreoffice: {type(lo_err).__name__}: {lo_err}")
        return None

    info = {"renderer": r.renderer, "pages": r.pages, "sections": r.sections,
            "bbox_overflow": r.bbox_overflow or [], "pdf_path": r.pdf_path}
    warnings.append(
        f"Word 渲染两次失败（{type(second_err).__name__}），已改用 LibreOffice "
        f"渲染复核通过（{r.pages} 页）——复核等级降档，文档本身无问题")
    if info["bbox_overflow"]:
        errors.append(f"image/table bbox overflow: {len(info['bbox_overflow'])} items {_bbox_detail(info['bbox_overflow'])}")
    return info


def validate_delivery(
    input_path: str | Path,
    source_path: Optional[str | Path] = None,
    spec=None,
    profile: str = "general",
    renderer: str = "word_preferred",
    word_render_callback=None,
    authorized_content_actions: Optional[list[str]] = None,
    capability_config: Optional[dict] = None,  # §20 P2：provider 选择配置（api 层注入，None → env/默认链）
    cover_rebuilt: bool = False,  # 封面已按模板重建 → 文字轴不适用（见 _cover_preservation_errors）
) -> DeliveryReport:
    """最终交付门禁（§8.10 工具契约）。

    word_render_callback: 可选回调 render(docx_path) -> dict
    （返回 {"pages": N, "sections": M, "bbox_overflow": [...], "renderer": "word"}），
    由上层（Windows + Word 环境）注入；无回调时只能到 LEVEL_C。
    """
    rep = DeliveryReport(input_path=str(input_path))
    errors: list[str] = []
    warnings: list[str] = []

    # ── 层 1：Static Package Gate ────────────────────────────────
    try:
        package_health = inspect_package(str(input_path))
    except ValueError as e:
        rep.hard_errors.append(f"package: {e}")
        rep.level = FAILED
        return rep
    errors.extend(package_health.errors)
    if package_health.duplicate_entries:
        errors.append(f"duplicate zip entries: {package_health.duplicate_entries[:3]}")

    # ── 层 2 + 计数：Analyzer / OOXML Invariant Gate ─────────────
    try:
        ir = analyze_document(str(input_path))
    except Exception as e:  # noqa: BLE001 — 解析失败即 hard error
        errors.append(f"analyze failed: {type(e).__name__}: {e}")
        rep.hard_errors = errors
        rep.level = FAILED
        return rep

    with zipfile.ZipFile(str(input_path)) as z:
        document_xml = z.read("word/document.xml").decode("utf-8", errors="replace")

    errors.extend(_check_duplicate_docpr(document_xml))
    errors.extend(_check_field_integrity(ir, warnings))
    errors.extend(_check_render_map(ir))

    # ── 层 2.5：复杂对象完整性（Phase 2：脚注/尾注引用=硬错误；修订/批注/书签=警告）──
    try:
        errors.extend(complex_object_gate_errors(str(input_path)))
        warnings.extend(complex_object_gate_warnings(str(input_path)))
    except Exception as e:  # noqa: BLE001 — 复杂对象分析失败不阻断（保守降级）
        warnings.append(f"complex object analysis skipped: {type(e).__name__}: {e}")

    # ── 层 2.6：封面残留检测（2026-08-15 真机回归）──
    # 用户从模板起稿常忘记删模板样例标题（如 Altera DDR…）→ 封面出现两行
    # 标题。工具只修格式不删内容（铁律），但必须告警提醒用户手动清理。
    try:
        residual = _cover_residual_hints(ir, spec)
        for hint in residual:
            warnings.append(f"模板示例残留未清除：{hint}")
    except Exception:  # noqa: BLE001 — 检测失败不阻断
        pass

    # ── 层 2.7：模板有页眉、源文档没有（2026-10-17）──
    # 提示由 format_compiler 按动作结果（parts_created）发出——这里拿到的
    # input_path 已是工作副本（可能刚被补出空白页眉），判不出「原本没有」。

    # ── 层 3：Semantic Format Gate（内容保持） ────────────────────
    src_tables: Optional[int] = None
    src_figures: Optional[int] = None
    if source_path is not None and Path(source_path).exists():
        # 2026-10-14 审计（fail-open 修复）：源解析失败与比对失败必须分开——
        # 源不可解析 = 无法证明内容保持 → 硬错误（此前整块 except 只记 warning、
        # 静默跳过，「守护」在最需要时失效、产物仍判 LEVEL A）；仅比对环节
        # 失败 → warning + content_preserved 保持 None（证书=unknown）+ 等级
        # 降档（见下方验证等级段）。
        src_ir = None
        try:
            src_ir = analyze_document(str(source_path))
        except Exception as e:  # noqa: BLE001
            errors.append(
                f"source analysis failed（无法证明内容保持）: "
                f"{type(e).__name__}: {e}")
        if src_ir is not None:
            try:
                src_hash = src_ir.metadata.get("content_hash")
                out_hash = ir.metadata.get("content_hash")
                rep.content_preserved = (src_hash == out_hash)
                if not rep.content_preserved:
                    # §10.1：内容变化例外必须由 ActionSpec 声明 + 用户授权。
                    # 声明了 content_change 的 action（如模板示例残留清理）执行后
                    # hash 变化是预期行为 → warning + 证书记录，不判 FAILED。
                    auth = authorized_content_actions or []
                    if auth:
                        warnings.append(
                            f"content changed by authorized actions: {','.join(auth)} "
                            f"（§10.1 例外：ActionSpec 声明 + 用户授权）"
                        )
                    else:
                        errors.append("content preservation FAILED: text hash changed")
                # 结构普查（2026-09-25，08-28 用户拍板「表格丢失必须报警而非
                # 静默」）：输入→产物表格/图片计数对比。表格丢失=内容丢失 →
                # 硬错误拦截交付；图片减少=警告（个别流程可能合法合并/内联）。
                # 机制层，与学校/模板无关。
                src_tables = len(src_ir.tables)
                src_figures = len(src_ir.figures)
                out_tables = len(ir.tables)
                if src_tables > out_tables:
                    errors.append(
                        f"structure loss: tables {src_tables}->{out_tables}"
                    )
                out_figs = len(ir.figures)
                if src_figures > out_figs:
                    warnings.append(
                        f"structure change: figures {src_figures}->{out_figs}"
                    )
            except Exception as e:  # noqa: BLE001
                warnings.append(
                    f"source comparison skipped: {type(e).__name__}: {e}")

    # ── 层 3.5：封面保全门（2026-09-25 横向机制审计发现 1）─────────────
    # 缺陷实证：把封面守卫关掉跑真实编译，封面「姓 名:」行被
    # jc=None→center、首行缩进 1680→None，门禁仍报 success=True /
    # LEVEL_A_WORD_VERIFIED / hard_errors=[] —— 行为修好了，感知没建。
    # 用户 2026-09-25 拍板 B：封面及封面性质页的段落格式不得被正文格式
    # 污染；页面设置（页边距/节属性）继续按模板封面节取值=现状；封面页
    # 页眉/页码被写脏必须纠正。详见 _cover_preservation_errors。
    if source_path is not None and Path(source_path).exists():
        try:
            errors.extend(
                _cover_preservation_errors(str(source_path), str(input_path),
                                           check_text=not cover_rebuilt))
        except Exception as e:  # noqa: BLE001 — 检测失败不阻断（保守降级）
            warnings.append(
                f"cover preservation check skipped: {type(e).__name__}: {e}")

    # ── 层 3.5b：公式/上标保全门（2026-09-27）───────────────────────
    # 缺陷实证：模板 extras 回放把作者语义（w:vertAlign）当未建模属性清空，
    # 真论文 139 处引用上标 + 作者角标 139→0、8 个 MathType 公式段被正文首行
    # 缩进右推，而交付门只比对文本、全程绿灯。数学与上标必须有独立口径。
    if source_path is not None and Path(source_path).exists():
        try:
            errors.extend(
                _equation_preservation_errors(str(source_path), str(input_path),
                                              warnings=warnings))
        except Exception as e:  # noqa: BLE001 — 检测失败不阻断（保守降级）
            warnings.append(
                f"equation preservation check skipped: {type(e).__name__}: {e}")

    # ── 层 3.6：标题形态门（2026-09-26 用户「识别层得加强一下」）────────
    # 常盲区：识别层把正文句判成标题 → 标题格式（放大 + 大纲 + 分页）下发，
    # 行为层与既有门禁全报成功。本层只看产物格式证据（大纲级别 / 居中大字号）
    # 与文本形态是否自相矛盾，与识别判据独立——识别层再松，出海前也会被拦。
    try:
        errors.extend(_heading_shape_errors(str(input_path), warnings))
    except Exception as e:  # noqa: BLE001 — 检测失败不阻断（保守降级）
        warnings.append(f"heading shape check skipped: {type(e).__name__}: {e}")

    # ── 层 4：Render Gate（P2 provider 化：回调优先，未注入则注册表 resolve+降级链）──
    render_info = None
    if renderer != "none":
        if word_render_callback is not None:
            try:
                render_info = word_render_callback(str(input_path))
                rep.renderer_used = render_info.get("renderer", "word")
                if render_info.get("bbox_overflow"):
                    errors.append(f"image/table bbox overflow: {len(render_info['bbox_overflow'])} items {_bbox_detail(render_info['bbox_overflow'])}")
            except RendererUnavailableError as e:  # Word 不可用≠文档问题 → warning 降级
                warnings.append(f"word renderer unavailable: {e}；静态门禁通过（{LEVEL_C_STATIC_ONLY}）")
            except Exception as e:  # noqa: BLE001
                # 2026-10-06 方案 B（用户拍板）：Word 渲染失败先重试一次（冷启动/
                # COM 忙/文件被用户占用都是环境瞬时故障——实证好文档也会撞上，见
                # tx_8a938def065b 事件）；重试仍失败再走 LibreOffice 兜底渲染复核；
                # **双渲染器都失败才硬拒**（那才是文档真实损坏的证据）。
                render_info = _render_retry_then_fallback(
                    str(input_path), word_render_callback, e, warnings, errors)
                if render_info is not None:
                    rep.renderer_used = render_info.get("renderer", "unknown")
        else:
            # provider 化路径（配置化而非 if-else）：resolve 后执行，失败按降级警告
            try:
                from .providers import resolve_provider, render_with_provider
                _prov = resolve_provider(renderer, config=capability_config)
                if _prov != "none":
                    _r = render_with_provider(_prov, str(input_path))
                    rep.renderer_used = _r.renderer
                    render_info = {
                        "renderer": _r.renderer, "pages": _r.pages,
                        "sections": _r.sections, "bbox_overflow": _r.bbox_overflow,
                        "pdf_path": _r.pdf_path,
                    }
                    if _r.bbox_overflow:
                        errors.append(f"image/table bbox overflow: {len(_r.bbox_overflow)} items {_bbox_detail(_r.bbox_overflow)}")
                else:
                    warnings.append(f"render provider unavailable；静态门禁通过（{LEVEL_C_STATIC_ONLY}）")
            except RendererUnavailableError:  # noqa: BLE001 — 同上语义：环境问题降级
                warnings.append(f"render provider unavailable；静态门禁通过（{LEVEL_C_STATIC_ONLY}）")
            except Exception as e:  # noqa: BLE001 — 渲染失败 = 硬错误（文档真实损坏）
                errors.append(f"render failed: {type(e).__name__}: {e}")

    # ── 层 4.7：目录页码真机回写（2026-09-26 用户定案）─────────────────
    # 页码唯一真源 = Word 排版：交付前读 _Toc 书签的实际页码写回目录条目缓存。
    # 目录条目文本长度不变 → 回写不影响分页，故可在渲染之后执行；无 Word 时降级。
    toc_pages_info: dict = {}
    try:
        from earlyz.validation.renderers import probe_bookmark_pages
        from earlyz.actions.new_actions import sync_toc_page_numbers as _sync_toc
        _pm = probe_bookmark_pages(str(input_path))
        # 优先用「显示页码」（wdActiveEndAdjustedPageNumber）：目录里要写的就是页脚
        # 上显示的那个数。物理页 − 1 的推算只对「正文从第 1 页起算」的模板成立，
        # 从封面连续编号的模板会整份偏小（2026-11-08：目录 1/9/10/11 vs 页脚 5/13/14/15）。
        _tb = _pm.get("toc_bookmarks_adj") or _pm.get("toc_bookmarks") or {}
        _adj_ok = bool(_pm.get("toc_bookmarks_adj"))
        if _tb:
            _sync = _sync_toc(str(input_path), {"page_map": _tb, "offset": 0 if _adj_ok else None})
            toc_pages_info = {
                "entries": _sync.get("entries"),
                "updated": _sync.get("updated"),
                "offset": _sync.get("offset"),
                "physical_pages": _pm.get("pages"),
                "source": "word-adjusted" if _adj_ok else "word",
            }
        else:
            warnings.append("目录页码回写跳过：未发现 _Toc 书签")
    except RendererUnavailableError as _exc:
        warnings.append(f"目录页码回写跳过（无 Word）：{_exc}")
    except Exception as _exc:  # noqa: BLE001 —— 不阻断交付
        warnings.append(f"目录页码回写异常：{type(_exc).__name__}: {_exc}")
    # ── 层 4.8：目录 F9 回环验收（2026-11-08 用户要求改口径）──────────────
    # 旧口径只判「有没有 TOC 域」，读的是我们自己写的缓存条目 → 「后置章节没大纲
    # 级别，F9 一按整条消失」「目录与正文同节，F9 后接排」在两套基线上全绿。
    # 新口径在 Word 里真更新一次域，再判条目完整 / 目录独立 / 页码一致。
    toc_after_update_info: dict = {}
    # 只在真用过 Word 渲染时做（没有 Word 的环境不必多开一次 COM，也避免测试变慢）
    if (render_info is not None and render_info.get("renderer") == "word"
            and (capability_config or {}).get("toc_after_update_verify", True)):
        try:
            from earlyz.validation.toc_after_update import verify as _verify_toc_after
            _ta = _verify_toc_after(Path(input_path))
            toc_after_update_info = _ta
            if not _ta.get("ok") and not _ta.get("skipped"):
                _ch = _ta.get("checks") or {}
                _miss = (_ch.get("entries_complete") or {}).get("missing") or []
                if _miss:
                    errors.append("目录 F9 后缺条目：" + "、".join(str(x) for x in _miss[:6]))
                _sep = _ch.get("toc_separate_from_body") or {}
                if _sep.get("ok") is False:
                    errors.append(
                        "目录与正文同页（更新域后接排）：目录末页 "
                        f"{_sep.get('toc_last_page')} == 正文首页 {_sep.get('first_head_page')}")
                _pn = _ch.get("page_numbers_match") or {}
                if _pn.get("ok") is False:
                    warnings.append(f"目录页码与页脚不一致 {_pn.get('count')} 条（更新域后）")
        except RendererUnavailableError as _exc:
            warnings.append(f"目录 F9 回环验收跳过（无 Word）：{_exc}")
        except Exception as _exc:  # noqa: BLE001 —— 验收器自身故障不阻断交付
            warnings.append(f"目录 F9 回环验收异常：{type(_exc).__name__}: {_exc}")

    # ── 层 4.9：标题编号完整性（2026-11-08 用户实锤「1. 1. 引言」）──────
    # 为什么单开一层：score_layout 里 numbering/numPr 出现 0 次 → 199 份基线对
    # 「样式引用悬空 numId → Word 默认列表多画一个号 + 720twips 缩进」完全免疫。
    # A 引用闭合（纯 XML，总是跑）；B 重复号 / C 异常缩进（Word 真机读数）。
    numbering_info: dict = {}
    try:
        from earlyz.validation.numbering_integrity import verify as _verify_numbering

        _num = _verify_numbering(Path(input_path))
        numbering_info = _num
        _refs = (_num.get("checks") or {}).get("styles_numbering_closed") or {}
        if _refs.get("ok") is False:
            errors.append("样式引用的编号定义缺失（numId "
                          + "、".join(str(x) for x in (_refs.get("missing") or [])[:4])
                          + "）——Word 会按默认列表补画编号并加缩进")
        _hd = (_num.get("checks") or {}).get("heading_numbers") or {}
        if _hd.get("ok") is False:
            _dup = _hd.get("duplicates") or []
            errors.append("标题出现重复编号 " + str(_hd.get("count")) + " 处"
                          + ("（如 " + str(_dup[0].get("text")) + " 前已画 " + str(_dup[0].get("list")) + "）"
                             if _dup else ""))
        if (_hd.get("indented_count") or 0) > 0:
            warnings.append(f"标题左缩进异常 {_hd.get('indented_count')} 处"
                            f"（最大 {max((x.get('left_cm') or 0) for x in (_hd.get('indented') or [{}]))}cm）")
    except Exception as _exc:  # noqa: BLE001 —— 验收器自身故障不阻断交付
        warnings.append(f"标题编号验收异常：{type(_exc).__name__}: {_exc}")

    # ── 层 4.6：表格分页复核（B 主，2026-09-26 定案）─────────────────────
    # A（估算 + 分类）已在编译链施工；此处用 Word 真机事实复核表单表是否真的
    # 整表同页，偏差则 force_keep 纠偏并复测。无 Word 时降级 warning，不降交付等级。
    table_pagination_info: dict = {}
    if render_info is not None and render_info.get("renderer") == "word":
        try:
            table_pagination_info = _table_pagination_pass(str(input_path), warnings)
        except Exception as e:  # noqa: BLE001
            warnings.append(f"table pagination pass crashed: {type(e).__name__}: {e}")

    rep.hard_errors = errors
    rep.warnings = warnings

    # ── 验证等级与证书 ───────────────────────────────────────────
    if errors:
        rep.level = FAILED
        rep.passed = False
    elif (source_path is not None and Path(source_path).exists()
          and rep.content_preserved is None):
        # 2026-10-14：源文件在、但内容保持无法证明（比对环节失败）→ 不宣称
        # 已复核：降为静态级别（不拦截交付，证书明示 content_preserved=unknown）。
        rep.level = LEVEL_C_STATIC_ONLY
        rep.passed = True
    elif render_info is not None:
        rep.level = LEVEL_A_WORD_VERIFIED if render_info.get("renderer") == "word" else LEVEL_B_FALLBACK_RENDERED
        rep.passed = True
    else:
        rep.level = LEVEL_C_STATIC_ONLY
        rep.passed = True

    rep.certificate = {
        "input_path": str(input_path),
        "input_hash": _file_hash(str(input_path)),
        "validation_level": rep.level,
        "profile": profile,
        "table_pagination": table_pagination_info,
        "toc_pages": toc_pages_info,
        "toc_after_update": toc_after_update_info,
        "numbering": numbering_info,
        "paragraphs": len(ir.paragraphs),
        "tables": len(ir.tables),
        "figures": len(ir.figures),
        "fields": len(ir.fields),
        "sections": len(ir.sections),
        "styles": len(ir.styles),
        "content_preserved": rep.content_preserved,
        "content_hash_scope": ("body 顶层段落 w:t；忽略全部空白/制表符；不含表格/"
                               "页眉页脚/脚注；TOC 与参考文献 [N] 前缀豁免"),
        "renderer_used": rep.renderer_used,
        "source_tables": src_tables,
        "source_figures": src_figures,
        "word_pages": render_info.get("pages") if render_info else None,
        "word_sections": render_info.get("sections") if render_info else None,
        "bbox_overflow": len(render_info.get("bbox_overflow", [])) if render_info else 0,
        "hard_error_count": rep.hard_error_count,
        "warnings": len(rep.warnings),
    }
    return rep


def _bbox_detail(items, limit: int = 5) -> str:
    """溢出项明细（页码/类型/bbox/页尺寸），避免只报数量无法定位。"""
    if not items:
        return ""
    out = []
    for it in list(items)[:limit]:
        try:
            out.append(f"p{it.get('page')} {it.get('kind')} {it.get('bbox')}/{it.get('page_size')}")
        except Exception:  # noqa: BLE001
            out.append(str(it)[:60])
    return " | ".join(out)
