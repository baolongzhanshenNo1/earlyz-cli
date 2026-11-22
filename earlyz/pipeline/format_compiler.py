"""T11 — Format Compiler 编排器（3.2 §8.2 核心工具调用链，ADR-009）。

compile_document 串起完整闭环：
analyze_document → compile_format_spec → lint_format → plan_format_repairs
→ normalize_plan_actions → TransactionExecutor → validate_delivery

- LLM 不参与规划（确定性，ADR-005）；
- 原文件永不修改（事务从副本开始，ADR-008）；
- 幂等：同一文档同一 spec 第二次执行有效 action 为 0；
- 输出 CompileResult（IR/spec/lint/plan/transaction/validation 全链路产物）。
"""

from __future__ import annotations

import json
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from lxml import etree

from earlyz.actions.catalog import build_catalog_registry
from earlyz.actions.new_actions import normalize_plan_actions
from earlyz.ir.analyzer import analyze_document
from earlyz.pipeline.latex_scripts import normalize_latex_scripts_in_copy
from earlyz.pipeline.text_split import normalize_text_structure_in_copy
from earlyz.ir.models import DocumentIR
from earlyz.lint.engine import LintResult, lint_format
from earlyz.planner.models import RepairAction, RepairPlan
from earlyz.planner.planner import PlanningPolicy, plan_format_repairs
from earlyz.spec.compilers import CompilationResult, compile_format_spec
from earlyz.spec.models import FormatSpec
from earlyz.transaction.executor import TransactionExecutor, TransactionResult
from earlyz.validation.delivery import DeliveryReport, validate_delivery


_ROLE_HEADINGS = {
    "原创性声明": "front", "郑重声明": "front", "学位论文原创性声明": "front",
    "摘要": "front", "Abstract": "front", "ABSTRACT": "front",
    "目录": "front", "参考文献": "back", "References": "back", "REFERENCES": "back",
    "致谢": "back", "附录": "back",
}
_CHAPTER_RE = re.compile(r"^第\s*[0-9一二三四五六七八九十]+\s*章")
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

#: 允许下发 outline_lvl 的角色（EZI-0014；与 new_actions._outline_level_for_role 同集合）
_OUTLINE_ROLES = frozenset({
    "heading1", "heading2", "heading3", "heading4",
    "refs_heading", "ack_heading", "appendix_heading", "bio_heading",
})


def _drop_non_heading_outline_levels(actions: list[dict]) -> dict:
    """从非标题角色的段落动作里删掉 ``outline_lvl`` 参数（EZI-0014）。

    由来（2026-11-16 交大中期实测）：模板规范把 outline_lvl 当「可移植类型
    属性」抄进非标题角色规则——affiliation（单位/学院行）规则的 expected 带
    ``outline_lvl=0``（spec/decompiler `_PORTABLE_TYPED_KEYS` 含 outline_lvl，
    值来自模板样本段的直接格式）。`set_paragraph_format` 按 role 批量下发后，
    封面表单字段「学院 （系）:」拿到大纲级别 0；后面的 TOC 生成/更新把它当
    一级条目收进目录。

    层级只允许由 heading*/后置章节角色写入（值由角色推导，不抄模板样例）。
    不设语义角色的段（node_id/text 定位的专项动作）不在此收口——无从判断角色
    时不猜，交给交付前 `cover_guard.strip_invented_outline_levels` 兜底。
    """
    dropped: list[str] = []
    kept: list[dict] = []
    for a in actions:
        if a.get("action") != "set_paragraph_format":
            kept.append(a)
            continue
        params = a.get("params") or {}
        props = params.get("properties") or {}
        if "outline_lvl" not in props:
            kept.append(a)
            continue
        role = str((params.get("selector") or {}).get("semantic_role") or "")
        if not role or role in _OUTLINE_ROLES:
            kept.append(a)
            continue
        props.pop("outline_lvl", None)
        dropped.append(role)
        if not props:
            continue  # 只剩 outline_lvl 的动作整条撤掉（空 properties 只会空跑一趟）
        kept.append(a)
    actions[:] = kept
    return {"dropped": len(dropped), "roles": sorted(set(dropped))}


def _template_body_line_fact(template_path: str | Path) -> dict | None:
    """模板正文行距实测（判定层 ``line_spacing_fact`` 同口径，EZI-0045）。

    为什么复用判定层函数而不是在引擎侧再写一套：正文行距项判的就是
    「模板/产物**同一个** ``body_paragraphs`` 取样 + **有效属性**（含 pStyle 链与
    默认段落样式继承）的众数」。引擎侧只统计段落**直接**设值（EZI-0017 旧口径）时，
    封面表单段（``中图分类号：``/``作者签名：`` 这类直接写 ``300/auto`` 的段落）会
    占据样本池，而真正的正文段从 ``Normal`` 继承的 1.5 倍行距根本进不了投票——
    矿大 3-3-1 模板正文段 80/146 是 ``360/auto``，引擎却抽出 ``300/auto``，产物被
    整篇刷成 1.25 倍。「期望只能来自模板正文段落」，所以取样与取值口径必须与判定层
    同源，两侧才不可能漂移。

    返回 ``{"line", "line_rule", "n", "set", "total", "source"}``；模板正文一个行距
    都没设时返回 ``None``（判定层会走默认样式/无值，这里不改 spec）。
    """
    from earlyz.validation.layout_fidelity import _load_docx, line_spacing_fact

    fact = line_spacing_fact(_load_docx(str(template_path)))
    if fact.get("line") is None:
        return None
    return {
        "line": int(fact["line"]),
        "line_rule": str(fact.get("rule") or "auto"),
        "n": int(fact.get("n") or 0),
        "set": int(fact.get("set") or 0),
        "total": int(fact.get("total") or 0),
        "source": fact.get("source"),
    }


def _align_body_line_with_template(spec: FormatSpec, template_path: str | Path) -> dict:
    """把 ``rule:tpl.body`` 的 line/line_rule 锚到模板正文段落（EZI-0045）。

    只改引擎侧的**期望值**，不改判定层、不按文档名分支、不引入常量。模板正文段落
    没设行距（判定层 N-A / 默认样式）时原样返回，保留 decompiler 的兜底链。
    """
    info: dict = {"applied": False}
    try:
        fact = _template_body_line_fact(template_path)
    except Exception as exc:  # noqa: BLE001 — 锚定失败退回 decompiler 取值，不阻断编译
        info["error"] = f"{type(exc).__name__}: {exc}"[:200]
        return info
    if fact is None:
        info["reason"] = "模板正文段落未设行距（判定层按默认样式/无值处理，不改 spec）"
        return info
    rule = next((r for r in spec.rules if r.rule_id == "rule:tpl.body"), None)
    if rule is None:
        info["reason"] = "spec 无 rule:tpl.body（模板没有正文样本规则）"
        return info
    before = f"{rule.expected.get('line')}/{rule.expected.get('line_rule')}"
    rule.expected["line"] = fact["line"]
    rule.expected["line_rule"] = fact["line_rule"]
    # `line_spacing` 是 lineRule=auto 时由 line/240 派生的便利字段；与 line 并存会
    # 因字典/动作顺序互相覆盖（exact/atLeast 尤其会被改回 auto）。
    rule.expected.pop("line_spacing", None)
    info.update(
        applied=True,
        before=before,
        after=f"{fact['line']}/{fact['line_rule']}",
        n=fact["n"], set=fact["set"], total=fact["total"], source=fact["source"],
    )
    try:
        if isinstance(spec.structure, dict):
            src = spec.structure.setdefault("body_line_source", {})
            if isinstance(src, dict):
                src["aligned"] = dict(info)
    except Exception:  # noqa: BLE001 — 溯源信息写不进不影响排版
        pass
    return info


def _template_body_size_fact(template_path: str | Path,
                             spec_pt: object = None) -> dict | None:
    """模板正文主字号实测（判定层 ``正文主字号(pt)`` 同口径，EZI-0047）。

    为什么必须复用判定层的取样/取值函数，而不是再写一套引擎侧众数：判定层判的是
    「模板/产物**同一个** ``body_text_paragraphs`` 取样（``body_paragraphs`` 去掉题注）
    + **有效属性**（含 pStyle 链与 Normal/docDefaults 继承）的众数」，双峰时用
    ``TEMPLATE_MODE_MIN_SHARE``（0.60）判主峰、不过就取前二当候选集合。引擎侧
    ``spec/decompiler._template_body_size`` 是另一套口径——只收「正文类角色 + 文本
    >10 字 + 封面区外」的段落，且只统计段落**自己**写的字号、覆盖率 <50% 就判定
    「字号是继承来的」而返回 None。两侧取样域差了 6~10 倍（人大本科：判定层 65 段 /
    引擎 10 段；湖南师大：192 段 / 7 段），于是期望值必然漂移，方向还双向：
    引擎返回 None 时走 Normal 兜底 → 产物 12.0/16.0（偏大）；引擎抽到少数派直接值
    → 产物 10.5/14.0（偏小）。「期望只能来自模板正文段落」，所以取样与取值口径必须
    与判定层同源，两侧才不可能漂移。

    返回 ``{"size", "candidates", "n", "total", "share", "source"}``；
    判定层这一支**不判**（样本不足 <5 段 → 退回 spec 注解 / 分布过散 → N-A）时返回
    ``None``，调用方保留 decompiler 的取值（两侧仍然同一个值）。
    """
    from earlyz.validation.layout_fidelity import (
        TEMPLATE_MODE_MIN_SAMPLES, _body_size_expectation, _load_docx,
        _template_mode, body_text_paragraphs, document_paragraphs,
    )

    props = body_text_paragraphs(document_paragraphs(_load_docx(str(template_path))),
                                 exclude_form_labels=False)
    meas = _template_mode(props, "sz")
    if int(meas.get("total") or 0) < TEMPLATE_MODE_MIN_SAMPLES:
        # 判定层同一支：样本不足 → 期望退回 spec 注解（= 引擎 decompiler 的值），
        # 引擎不改它才对得上；改了反而把「注释值 vs 实测值」的差造出来。
        return None
    _exp, cand, src, forced = _body_size_expectation(props, spec_pt)
    if forced or not cand:
        # 分布过散（前二合计 <60%）→ 判定层 N-A，不判；引擎保留 decompiler 取值。
        return None
    sizes = tuple(float(v) for v in cand)
    return {
        "size": sizes[0],            # 众数（双峰时即前二里的最高频，判定层任一候选均 PASS）
        "candidates": sizes,
        "n": int(meas.get("n") or 0),
        "total": int(meas.get("total") or 0),
        "share": float(meas.get("share") or 0.0),
        "source": src,
    }


def _pt_same(a: object, b: object, tol: float = 0.01) -> bool:
    """字号是否同一个值（pt 浮点直比会因 10.5/10.50 之类误判）。"""
    try:
        return a is not None and b is not None and abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def _align_body_size_with_template(spec: FormatSpec, template_path: str | Path) -> dict:
    """把 ``rule:tpl.body`` 的 font_size_pt 锚到模板正文段落（EZI-0047）。

    只改引擎侧的**期望值**，不改判定层、不按文档名分支、不在引擎里另造阈值：
    主峰/双峰/过散全部由判定层 ``_body_size_expectation`` 判，引擎只把它的候选集合里
    最高频的那个写进规则（判定层「落在任一候选即 PASS」）。

    ``appendix_body`` / ``bio_body`` 是 decompiler 从正文规则 deepcopy 出来的
    （「该区无独立内容样本，继承正文规则」），值仍等于改前正文值时一并跟随，
    否则同一篇的附录正文会被写成旧字号、把正文主字号众数拆散。
    """
    info: dict = {"applied": False}
    rule = next((r for r in spec.rules if r.rule_id == "rule:tpl.body"), None)
    if rule is None:
        info["reason"] = "spec 无 rule:tpl.body（模板没有正文样本规则）"
        return info
    before = rule.expected.get("font_size_pt")
    try:
        fact = _template_body_size_fact(template_path, before)
    except Exception as exc:  # noqa: BLE001 — 锚定失败退回 decompiler 取值，不阻断编译
        info["error"] = f"{type(exc).__name__}: {exc}"[:200]
        return info
    if fact is None:
        info["reason"] = ("模板正文样本不足/分布过散（判定层退回注解或 N-A），"
                          "不改 spec，保留 decompiler 取值")
        return info
    rule.expected["font_size_pt"] = fact["size"]
    propagated: list[str] = []
    for r in spec.rules:
        if r is rule:
            continue
        role = str((r.selector or {}).get("semantic_role") or "")
        if role in ("appendix_body", "bio_body") and _pt_same(
                r.expected.get("font_size_pt"), before):
            r.expected["font_size_pt"] = fact["size"]
            propagated.append(role)
    info.update(
        applied=True,
        before=before,
        after=fact["size"],
        candidates=list(fact["candidates"]),
        n=fact["n"], total=fact["total"], share=round(fact["share"], 3),
        source=fact["source"],
        propagated=sorted(propagated),
    )
    try:
        if isinstance(spec.structure, dict):
            src = spec.structure.setdefault("body_size_source", {})
            if isinstance(src, dict):
                src["aligned"] = dict(info)
    except Exception:  # noqa: BLE001 — 溯源信息写不进不影响排版
        pass
    return info


#: Word 的 `w:outlineLvl` 取值语义：0–8 = 大纲第 1–9 级（标题），9 = 「正文文本」
#: （显式声明不是标题）。过滤只保护 0–8；9 是正文，照旧按正文规则排。
_OUTLINE_LEVEL_MAX = 8


def _outline_level_of(p: object, lf_styles: object) -> int | None:
    """段落**有效**大纲级别（``docDefaults`` → pStyle 链 → 直接格式，判定层同一实现）。

    复用 ``earlyz.validation.layout_fidelity`` 的 ``_Styles`` / ``_ppr_chain`` /
    ``_val`` / ``_int``（EZI-0050）：判定层 ``heading_samples_by_level`` 就是按这条链
    取 ``outline_lvl`` 的，引擎侧另写一套必然与尺子错开（排除规则只许有一份实现）。
    lxml 与 stdlib ET 的这几处 API 同形，可直接复用。
    """
    from earlyz.validation.layout_fidelity import _int, _ppr_chain, _val

    try:
        return _int(_val(_ppr_chain(p, lf_styles), "outlineLvl"))
    except Exception:  # noqa: BLE001 — 取不到级别按「没有大纲级别」处理
        return None


def _body_targets_without_outline(doc_path: str | Path, selector: dict, *,
                                  heading_levels: set[int] | None = None) -> dict:
    """正文规则的作用集：按角色解析后**剔掉带大纲级别的段落**（EZI-0050）。

    为什么角色过滤不够：正文规则按 ``selector.semantic_role="body"`` 解析段落，而动作层
    的角色链里 ``FrontRoleAssigner`` 会把标题段判成 ``body``——引擎自己的编号标题判据
    ``_is_numeric_heading`` 认不出「1  引言」这种**编号后两个空格**的写法，于是
    「1  引言的说明」/「2  题目、作者和单位」/「3  摘要和关键词」被判成正文
    （航空学报实测：``classify_paragraph_role`` = heading1，``FrontRoleAssigner`` 之后
    = body）。判定层却按 ``outlineLvl`` + 编号形态把它们算进一级标题样本
    （``heading_samples_by_level``）；正文规则一盖章，标题值就被刷成正文值
    （重庆科技 6 项 PASS→FAIL；航空学报正文 11.0pt 与真标题 16.0pt 4:4 平票 →
    众数取到正文值，一级标题字号 PASS→FAIL）。

    ``heading_levels``（事务内**首发**与 6.4b **补齐**用）：只剔「引擎为该级配了标题
    动作」的段（1 表示 heading1…）——那些段归标题规则管，正文规则不许盖章；引擎没有
    该级标题动作时（计算机学报：二级/三级示例标题根本没进 lint），正文规则仍是唯一
    写入方，不能掐掉（否则那两级只剩 Normal 的 9.0pt，凭空多两个 FAIL）。
    ``None`` = 严格口径（交付前**终局重放** 5.991 用）：带大纲级别的段一个都不许重放。

    返回 ``{"captured", "kept", "dropped_outline"}``——捕获用动作层自己的指纹机制
    （``_capture_resolved_paragraphs``），与计划动作同一套定位。
    """
    from earlyz.actions import new_actions as _na
    from earlyz.validation.layout_fidelity import _Styles

    path = Path(doc_path)
    root = _na._read_part(path, "word/document.xml")
    styles_map = _na._read_styles_map(path)
    candidates = _na._locate_paragraphs(root, selector or {}, styles_map)
    style_root = None
    try:
        style_root = _na._read_part(path, "word/styles.xml")
    except Exception:  # noqa: BLE001 — 无 styles 部件时只剩直接格式
        style_root = None
    lf_styles = _Styles(style_root)
    kept: list = []
    dropped = 0
    for p in candidates:
        lvl = _outline_level_of(p, lf_styles)
        if lvl is not None and 0 <= lvl <= _OUTLINE_LEVEL_MAX:
            if heading_levels is not None and (lvl + 1) not in heading_levels:
                kept.append(p)      # 引擎没有该级标题规则：正文规则照旧
                continue
            dropped += 1
        else:
            kept.append(p)
    return {
        "captured": _na._capture_resolved_paragraphs(root, kept),
        "kept": len(kept),
        "dropped_outline": dropped,
    }


def _heading_levels_in_plan(actions: list[dict]) -> set[int]:
    """计划里配了标题动作的级别集合（``heading1`` → 1）。"""
    levels: set[int] = set()
    for a in actions:
        if str(a.get("action") or "") != "set_paragraph_format":
            continue
        role = str(((a.get("params") or {}).get("selector") or {})
                   .get("semantic_role") or "")
        m = re.fullmatch(r"heading(\d)", role)
        if m:
            levels.add(int(m.group(1)))
    return levels


def _filter_body_plan(actions: list[dict], doc_path: str | Path,
                      heading_levels: set[int] | None = None) -> dict:
    """把计划里**正文**动作的作用集收窄成「不属于标题规则的段」（EZI-0050）。

    事务内首发、6.4b 补齐与 5.991 终局重放必须同一条口径：否则首发把正文值（字号/字体/
    对齐）写到标题段上，终局重放又去修另一批段，两边各自把标题值刷回正文值。
    这里直接改动作参数：给 ``set_paragraph_format(role=body)`` 塞一份过滤后的
    ``_resolved_paragraphs`` 指纹集（与动作层「计划捕获」同一机制），动作层就不再按
    role 二次解析（那一步会把标题段圈回来）。
    """
    info: dict = {"applied": False}
    for a in actions:
        if str(a.get("action") or "") != "set_paragraph_format":
            continue
        params = a.get("params") or {}
        selector = params.get("selector") or {}
        if str(selector.get("semantic_role") or "") != "body":
            continue
        roster = _body_targets_without_outline(doc_path, selector,
                                               heading_levels=heading_levels)
        info = {"applied": True, **{k: roster[k] for k in ("kept", "dropped_outline")}}
        if roster["kept"] == 0:
            # 一个正文段都没有（整篇都是标题？）→ 保留原动作，交给动作层按 role 解析
            info["skipped"] = "过滤后无正文段，保留原动作"
            return info
        params["_resolved_paragraphs"] = roster["captured"]
        a["params"] = params
        return info
    return info


def _doc_section_count(path) -> int:
    try:
        with zipfile.ZipFile(path) as z:
            root = etree.fromstring(z.read("word/document.xml"))
        return len(list(root.iter(f"{{{W}}}sectPr")))
    except Exception:  # noqa: BLE001
        return 0


def _detect_role_boundaries(path) -> list[dict]:
    """按文档里的角色标题推断切节点（确定性，不猜内容）。

    只认精确标题（去空白后完全相等），因此不会命中目录里带页码的条目
    （「参考文献14」）。第一个「第 N 章」标题按正文章处理。
    """
    import zipfile as _zip
    from lxml import etree as _etree
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    try:
        with _zip.ZipFile(path) as z:
            root = _etree.fromstring(z.read("word/document.xml"))
    except Exception:  # noqa: BLE001
        return []
    body = root.find(f"{W}body")
    if body is None:
        return []
    seen: set[str] = set()
    out: list[dict] = []
    chapter_done = False
    for el in body.iter():
        if not isinstance(el.tag, str) or not el.tag.endswith("}p"):
            continue
        txt = re.sub(r"\s+", "", "".join(t.text or "" for t in el.iter(f"{W}t")))
        if not txt or len(txt) > 40:
            continue
        # 目录条目末尾带页码（「第1章引言6」「参考文献14」）——不是真正的标题
        if re.search(r"[0-9]$", txt):
            continue
        if not chapter_done and _CHAPTER_RE.match(txt):
            out.append({"heading": txt, "role": "body"})
            chapter_done = True
            seen.add(txt)
            continue
        role = _ROLE_HEADINGS.get(txt)
        if role and txt not in seen:
            seen.add(txt)
            out.append({"heading": txt, "role": role})
    return out


@dataclass
class CompileResult:
    document_id: str
    spec: Optional[FormatSpec] = None
    ir: Optional[DocumentIR] = None
    lint: Optional[LintResult] = None
    plan: Optional[RepairPlan] = None
    transaction: Optional[TransactionResult] = None
    validation: Optional[DeliveryReport] = None
    artifacts: dict = field(default_factory=dict)

    @property
    def passed(self) -> bool:
        return bool(self.validation and self.validation.passed)

    def to_dict(self) -> dict:
        return {
            "document_id": self.document_id,
            "spec": self.spec.to_dict() if self.spec else None,
            "lint_total": len(self.lint.violations) if self.lint else 0,
            "plan_actions": len(self.plan.actions) if self.plan else 0,
            "transaction_status": self.transaction.status if self.transaction else None,
            "validation_level": self.validation.level if self.validation else None,
            "passed": self.passed,
            "artifacts": self.artifacts,
        }


def _footer_page_number_props(hf_parts, template_size_pt) -> dict:
    """页脚页码的写入参数（EZI-0046）。

    **字号** ``font_size_pt`` 只认模板页脚实测（``template_size_pt`` =
    模板全部 footer part 里全部 ``w:sz`` 的众数，与判定层 ``_hf_stats`` 同一个
    统计量）。**故意不从 spec 的页脚段拿字号**：``hf_parts`` 那一条链是
    「第一个 part 的第一段首个 run → 样式链」，模板页脚首 run 常是只有
    ``rPr/sz=18`` 的空壳 run，或没有 sz 只靠 Footer 样式（18 半磅）→ 恒 9.0pt
    （8/9 例实测产物 9.0pt、模板 14.0/10.5 的根因）。``jc`` / 字体仍取模板
    PAGE 段的直接格式。

    ``template_size_pt is None``（模板页脚没有 w:sz）→ 不写 ``font_size_pt``：
    判定层此时判 N-A（无期望值），引擎不许拿 9/10.5/14 之类的常量顶替。
    """
    props: dict = {}
    for hp in (hf_parts or []):
        if (hp.get("kind") or "") != "footer":
            continue
        for pa in (hp.get("paragraphs") or []):
            if "PAGE" not in " ".join(str(x) for x in (pa.get("run_sig") or [])):
                continue
            for key in ("jc", "font_ascii", "font_eastasia"):
                if pa.get(key) is not None:
                    props[key] = pa[key]
            break
        if props:
            break
    if template_size_pt is not None:
        props["font_size_pt"] = float(template_size_pt)
    return props


def _template_footer_page_size_pt(spec_source) -> float | None:
    """模板页脚 ``w:sz`` 众数（磅）；取不到就是 ``None``（= 无实测证据）。"""
    if not spec_source or not str(spec_source).lower().endswith(".docx"):
        return None
    try:
        from earlyz.actions.new_actions import footer_page_size_pt
        return footer_page_size_pt(spec_source)
    except Exception:  # noqa: BLE001 — 读不到模板就不写字号，不兜底
        return None


def _template_front_mark_roles(spec_source_type: str, spec_source: str) -> list:
    """模板侧取值（2026-09-28 夜班）：作者行/单位行的编号角标在模板样本里是上标吗。

    只有 docx 模板源才可查（YAML / 内置规范没有样本段落）→ 返回空列表，
    调用方据此**不调度** `superscript_front_marks`（模板没写就不发明）。
    """
    try:
        if not spec_source:
            return []
        st = str(spec_source_type or "").lower()
        if "docx" not in st and not str(spec_source).lower().endswith(".docx"):
            return []
        from earlyz.actions.new_actions import front_mark_roles_in_document
        return front_mark_roles_in_document(spec_source)
    except Exception:  # noqa: BLE001 — 模板读不了就按「没这条规矩」处理（保守）
        return []



def _extract_cover_slots(source_path) -> dict:
    """从源文档封面抽槽位值（题目/英文题目/专业/姓名/学号/指导教师/日期）。

    2026-10-18：自动封面只传 template_path 不传 slots，模板样例标题与
    「指导教师：XXX」被整块带进产物。这里按源文档实际文字取值。
    """
    import zipfile as _zip
    from lxml import etree as _et

    W_ = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    try:
        doc = _et.fromstring(_zip.ZipFile(str(source_path)).read("word/document.xml"))
    except Exception:  # noqa: BLE001
        return {}
    paras = []
    for p in doc.iter(f"{W_}p"):
        t = "".join(x.text or "" for x in p.iter(f"{W_}t")).strip()
        if t:
            paras.append(t)
    head = paras[:30]
    out: dict = {}
    alias = {"专业名称": ("专业名称", "专业"), "学生姓名": ("学生姓名", "姓名"),
             "学生学号": ("学生学号", "学号"), "指导教师": ("指导教师", "导师姓名", "导师"),
             "日期": ("日期",)}
    for key, names in alias.items():
        for t in head:
            for nm in names:
                core = re.sub(r"[\s　]+", "", t)
                if key == "日期" and "提交" in core:
                    continue  # 提交日期不是封面日期
                if nm in core:
                    val = core.split(nm, 1)[1].strip("：:　 、,，")
                    val = re.sub(r"^(职称|及职称)[：:　 ]*", "", val).strip()
                    if val:
                        out[key] = val
                        break
            if key in out:
                break
    # 日期：时间只按源文档走（模板只提供样式）。优先取封面上的中文/数字日期行
    if "日期" not in out:
        for t in head:
            if re.search(r"^[〇○零一二三四五六七八九十]{2,4}年", t) or re.match(r"^\d{4}\s*年", t):
                out["日期"] = t
                break
    # 题目：跳过「毕业设计（论文）」后的第一行；英文题目：其后以括号/拉丁字母开头的一行
    body = [t for t in head if "毕业设计" not in t]
    for i, t in enumerate(body):
        if t.startswith("（") or re.match(r"^[A-Za-z]", t):
            continue
        if len(t) >= 6:
            out.setdefault("题目", t)
            for nxt in body[i + 1:i + 3]:
                if nxt.startswith("（") or re.match(r"^[A-Za-z]", nxt):
                    out.setdefault("英文题目", nxt.strip("（）()"))
            break
    return out


def _cover_slots_missing(source_path, labels: list[str]) -> list[str]:
    """封面槽位完整性：源文档里标签后面必须有非空取值（只用于决定是否自动重建）。"""
    import zipfile as _zip
    from lxml import etree as _et

    W_ = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    try:
        doc = _et.fromstring(_zip.ZipFile(str(source_path)).read("word/document.xml"))
    except Exception:  # noqa: BLE001
        return list(labels)
    texts = []
    for p in doc.iter(f"{W_}p"):
        t = "".join(x.text or "" for x in p.iter(f"{W_}t"))
        if t.strip():
            texts.append(re.sub(r"[\s　]+", "", t))
    missing = []
    for label in labels:
        core = re.sub(r"[\s　]+", "", str(label))
        # 同义标签：源文档常写「导师姓名、职称」「导师姓名及职称」，模板写「指导教师」
        _alias = {"指导教师": ("导师姓名", "导师"), "专业名称": ("专业",),
                  "学生姓名": ("姓名",), "学生学号": ("学号",), "日期": ("日期", "提交日期")}
        keys = [core] + [re.sub(r"[\s　]+", "", a) for a in _alias.get(core, ())]
        ok = False
        for t in texts[:80]:
            for key in keys:
                if key and key in t:
                    rest = t.split(key, 1)[1].strip("：:　 、,，")
                    if rest:
                        ok = True
                        break
            if ok:
                break
        if not ok:
            missing.append(str(label))
    return missing


def compile_document(
    input_path: str | Path,
    spec_source_type: str,
    spec_source: str,
    domain: str = "general",
    workdir: str | Path = ".",
    max_risk: str = "R2",
    preserve_content: bool = True,
    word_render_callback=None,
    print_mode: bool = False,
    libreoffice_path: Optional[str] = None,
    force_odd_from_section: int = 1,
    clean_example_residuals: bool = True,
    repair_section_headers: bool = True,
    format_captions: bool = True,
    format_table_cells: bool = True,
    normalize_footer_parts: bool = True,
    format_toc_entries: bool = True,
    progress_cb=None,
    capability_config: dict | None = None,  # §20 P2：renderer/provider 配置（api 层注入）
) -> CompileResult:
    """完整编译闭环（§8.2）。

    print_mode=True 时在事务产物上执行打印版最终化
    （奇数页强制 + 空白页 + 页码体系，IR 语义识别正文起始节）。
    repair_section_headers/format_captions/format_table_cells：源文档结构
    修复阶段（2026-08-16 真机回归接入）——节引用错位（正文绑目录页眉）、
    图表标题数字加粗、表内容 5 号。lint→planner 无 IR 级差异信号，
    故 compile 显式执行（纯格式操作，幂等）。
    """
    src = Path(input_path)
    work = Path(workdir)
    work.mkdir(parents=True, exist_ok=True)

    # 0) LaTeX 脚本残留归一（2026-09-27）：AI 粘贴内容常带 `$^{1)}$`（作者单位标注）、
    #    `$^{[1]}$`（文献引用）、`D_n` / `T_n^off`（正文伪公式记法）。必须在 analyze 之前——
    #    段指纹 `_stable_paragraph_fingerprint` 含规范化文本，晚做会让计划里的
    #    `_resolved_paragraphs` 命中集合错位（动作打空/打错段）。作用于工作副本，用户原件不动。
    src, _latex_stats = normalize_latex_scripts_in_copy(src, work)

    # 0.4) 正文结构拆分归一（2026-10-09 用户定案两项）：①行内序号列表化
    #      「（1）…（2）…（3）…」→ 引导句/各项/尾句独立成段（标点一律原样）；
    #      ②附录标题行内正文拆分「附录 A 描述…」→ 标题段 + 正文段。同 0)
    #      理由：必须在 analyze 之前作用于工作副本（指纹/角色/分页按拆后结构算）。
    src, _ts_stats = normalize_text_structure_in_copy(src, work)

    # 0.5) 字体名归一（2026-10-27）：替代字体（Noto/Liberation）与带分号的
    #      字体串（「宋体;SimSun」）都会被 Word 当不存在 → 回退默认字体。
    #      只改 rFonts 属性，不动文字；作用于工作副本。
    try:
        from earlyz.pipeline.fonts_norm import normalize_fonts_in_copy
        src, _font_stats = normalize_fonts_in_copy(src, work)
    except Exception:  # noqa: BLE001 — 归一是增强，失败不阻断编译
        _font_stats = {"changed": False}

    result = CompileResult(document_id=f"doc:{src.stem}")
    if _latex_stats.get("changed"):
        result.artifacts["latex_scripts"] = _latex_stats
    if _font_stats.get("changed"):
        result.artifacts["fonts_normalized"] = _font_stats
    if _ts_stats.get("changed"):
        result.artifacts["text_structure"] = _ts_stats

    # 1) Analyze
    result.ir = analyze_document(src)

    # 1.5) 源文档卫生：检测源稿自带的模板说明文字（只提示，不删——删文字属内容变更）
    try:
        from earlyz.validation.source_hygiene import detect_instruction_paragraphs
        hygiene = detect_instruction_paragraphs(result.ir)
        if hygiene.get("count"):
            result.artifacts["source_instructions"] = {
                "count": hygiene["count"],
                "samples": hygiene["samples"],
                "items": hygiene["items"][:50],
                "note": hygiene["note"],
            }
    except Exception:  # noqa: BLE001 — 检测失败不阻断编译
        hygiene = {}

    # 2) Compile spec（三源互斥由 compilers 保证）
    if progress_cb is not None:
        try:
            progress_cb({"type": "stage", "phase": "analyze", "summary": "解析文档结构与模板规范…"})
        except Exception:  # noqa: BLE001
            pass
    comp = compile_format_spec(spec_source_type, spec_source, domain)
    result.spec = comp.spec
    # 2.1) 正文行距期望锚定模板正文段落（EZI-0045）：decompiler 的正文行距取样只认
    #      段落**直接**设值且排掉封面/短行，与判定层的 `body_paragraphs` + 有效属性
    #      口径不一致（矿大 3-3-1：引擎 300/auto vs 判定层 360/auto）。这里按判定层
    #      同口径重算一次并写回 `rule:tpl.body`，产物才会被写成模板真正在用的行距。
    #      只改期望值：不改判定层、不按文档名分支、模板正文无行距时原样保留兜底链。
    if spec_source_type != "built_in_spec" and spec_source:
        try:
            _bl_align = _align_body_line_with_template(result.spec, spec_source)
            if _bl_align.get("applied"):
                result.artifacts["body_line_aligned"] = _bl_align
        except Exception as _bl_e:  # noqa: BLE001 — 锚定失败不阻断编译
            result.artifacts["body_line_aligned"] = {"applied": False,
                                                     "error": str(_bl_e)[:200]}
        # 2.2) 正文主字号期望锚定模板正文段落（EZI-0047）：与行距同一病根、同一个修法。
        #      decompiler 的 `_template_body_size` 只收「正文类角色 + >10 字 + 封面外」
        #      且要求过半段落**自己**写字号，与判定层的 `body_text_paragraphs` + 有效属性
        #      口径不一致（人大本科判定层 65 段/引擎 10 段；湖南师大 192 段/引擎 7 段），
        #      引擎取不到值就退回 Normal 12pt/16pt（产物偏大），取到少数派直接值就写小
        #      （产物偏小）。这里按判定层同口径重算，把它的候选集合里最高频的一个写回
        #      `rule:tpl.body`——判定层「落在任一候选即 PASS」，双峰模板也不会打架。
        try:
            _bs_align = _align_body_size_with_template(result.spec, spec_source)
            if _bs_align.get("applied"):
                result.artifacts["body_size_aligned"] = _bs_align
        except Exception as _bs_e:  # noqa: BLE001 — 锚定失败不阻断编译
            result.artifacts["body_size_aligned"] = {"applied": False,
                                                     "error": str(_bs_e)[:200]}

    # 模板规范自检（2026-10-29）：正文样本为空/过少、关键字段缺失都要在交付前
    # 变成告警——采样退化原来是静默的，产物看着正常但没按模板排。
    try:
        from earlyz.spec.health import spec_health_warnings
        _health = spec_health_warnings(result.spec)
        if _health:
            result.artifacts["spec_health"] = _health
    except Exception:  # noqa: BLE001
        _health = []
    if comp.unresolved:
        result.artifacts["spec_unresolved"] = comp.unresolved

    # 3) Lint
    result.lint = lint_format(result.ir, result.spec)

    # 4) Plan（确定性）
    if progress_cb is not None:
        try:
            progress_cb({"type": "stage", "phase": "plan",
                         "summary": f"生成排版方案（{len(result.lint.violations)} 处差异 → 动作）…"})
        except Exception:  # noqa: BLE001
            pass
    result.plan = plan_format_repairs(
        result.ir, result.spec, result.lint.violations,
        PlanningPolicy(max_risk=max_risk, preserve_content=preserve_content),
    )

    # 5) Apply（事务 + 参数归一化）——单一路径（2026-08-17 统一）：
    #    planner 段（lint 驱动）与 structural 段（compile 显式结构修复）
    #    全部走 TransactionExecutor：同一 journal、同一 verifier 门禁。
    reg = build_catalog_registry()
    normalized = normalize_plan_actions([a.to_dict() for a in result.plan.actions])
    # EZI-0014：非标题角色不得下发 outline_lvl（affiliation 规则带 outline_lvl=0 →
    # 封面表单字段「学院 （系）:」被标成一级大纲，进目录/导航）。必须在执行前改
    # 参数：之后的 TOC 生成步骤直接读 outlineLvl 取条目。
    _outline_drop = _drop_non_heading_outline_levels(normalized)
    if _outline_drop.get("dropped"):
        result.artifacts["outline_lvl_rejected"] = _outline_drop
    # EZI-0050：正文动作的作用集不得含带大纲级别的段（手写目录条目被 FrontRoleAssigner
    # 判成 body）。首发（这里）与 5.991 终局重放共用同一条过滤，见 `_filter_body_plan`。
    # 计划里配了标题动作的级别（正文动作对同级别的大纲段不盖章；只算一次，三处共用）
    _heading_levels = _heading_levels_in_plan(normalized)
    try:
        _body_scope = _filter_body_plan(normalized, src, _heading_levels)
        if _body_scope.get("applied"):
            result.artifacts["body_plan_scope"] = _body_scope
    except Exception as _bs_e:  # noqa: BLE001 — 过滤失败退回旧行为（按 role 解析）
        result.artifacts["body_plan_scope"] = {"applied": False, "error": str(_bs_e)[:200]}
    # 字体回退通报（2026-11-21）：Step 层在「目标字体确认没装」时已改用近似字体
    # （典型：公文标题的方正小标宋简体 → 黑体），这里把事实记进 artifacts，
    # 供末尾 warnings 让用户看见——字形换了必须说，不许静默。
    _font_subs = sorted({
        str(a.get("params", {}).get("font_substituted_from"))
        for a in normalized
        if a.get("action") == "insert_red_header"
        and a.get("params", {}).get("font_substituted_from")
    })
    if _font_subs:
        result.artifacts["font_substitutions"] = _font_subs

    from earlyz.planner.models import RepairAction
    tx_plan = RepairPlan(
        plan_id=result.plan.plan_id,
        actions=[RepairAction(**a) for a in normalized],
        invariants=result.plan.invariants,
        preconditions=result.plan.preconditions,
    )
    result.transaction = TransactionExecutor(reg, work).execute(
        tx_plan, src, progress_cb=progress_cb)

    # 5.5) 模板示例残留清理——独立内容清理阶段（§10.1 例外，用户授权）。
    #      不进入格式修复链（planner 不再生成该动作）；compile 显式执行。
    #      ActionSpec 声明 content_change → validate_delivery 授权放行。
    authorized_content_actions: list[str] = []
    if result.transaction.status == "applied" and clean_example_residuals:
        if result.spec.example_residuals:
            from earlyz.actions.new_actions import remove_example_residuals as _clean_residuals
            from earlyz.spec.decompiler import _norm_residual as _norm_res
            # 大面积删除护栏（2026-10-30）：源文档本身就是「规范/写作指南」时，
            # 通篇都是模板说明文字——自动清理会删掉几百上千字真实内容
            # （天津大学 -1634 字、自动化学报说明 -861 字实测）。
            # 命中段落 ≥5 且占全文 >10% → 只报告，不删除（删内容需用户确认）。
            _skip_mass = False
            try:
                import zipfile as _zf
                from lxml import etree as _et
                _W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
                with _zf.ZipFile(result.transaction.output_path) as _z:
                    _root = _et.fromstring(_z.read("word/document.xml"))
                _ps = list(_root.iter(f"{_W}p"))
                _res = {_norm_res(x) for x in result.spec.example_residuals}
                _hit = 0
                _hit_chars = 0
                _all_chars = 0
                for _p in _ps:
                    _txt = "".join(t.text or "" for t in _p.iter(f"{_W}t"))
                    _all_chars += len(_txt)
                    if _norm_res(_txt) in _res:
                        _hit += 1
                        _hit_chars += len(_txt)
                # 按**字数**占比判：几段长说明就能删掉上千字（天津大学 4 段 = 1634 字），
                # 按段落数占比会漏判。
                if _ps and _hit >= 2 and _all_chars and (_hit_chars / _all_chars) > 0.03:
                    _skip_mass = True
                    result.artifacts["residual_cleanup"] = {
                        "changed": False, "skipped": "mass_deletion_guard",
                        "matched": _hit, "paragraphs": len(_ps),
                        "matched_chars": _hit_chars, "total_chars": _all_chars,
                    }
            except Exception:  # noqa: BLE001 — 护栏失败不影响主流程
                _skip_mass = False
            if not _skip_mass:
                # 作者内容保护（2026-11-07）：把源文档的段落文字一起传给清理动作，
                # 与模板示例同形但源里也有的段落不删（软件编码说明书 13 段代码样例
                # 被当残留删掉 → content 判丢字）。
                _keep_texts: list[str] = []
                try:
                    with _zf.ZipFile(str(src)) as _zs:
                        _sroot = _et.fromstring(_zs.read("word/document.xml"))
                    _keep_texts = ["".join(t.text or "" for t in _p.iter(f"{_W}t"))
                                   for _p in _sroot.iter(f"{_W}p")]
                except Exception:  # noqa: BLE001 — 取不到源文本就退回旧行为
                    _keep_texts = []
                try:
                    clean_info = _clean_residuals(str(result.transaction.output_path),
                                                  {"residuals": result.spec.example_residuals,
                                                   "keep_texts": _keep_texts})
                    result.artifacts["residual_cleanup"] = clean_info
                    if clean_info.get("changed"):
                        authorized_content_actions.append("remove_example_residuals")
                        result.artifacts["residual_cleanup"]["note"] = (
                            "§10.1 内容变化例外：ActionSpec 声明 + 用户授权"
                            "（模板示例残留清理，2026-08-16 拍板）"
                        )
                except Exception as e:  # noqa: BLE001 — 清理失败不阻断格式交付
                    result.artifacts["residual_cleanup"] = {"changed": False,
                                                            "error": str(e)}

    # 5.55) 模板命名样式接线：把模板的「正文/标题 N/题注」样式克隆进产物，
    #       并给对应角色段落挂上 w:pStyle。此前全是直接格式，用户改样式无效。
    #       非阻断：失败只记录，不影响交付。
    if result.transaction.status == "applied":
        try:
            from earlyz.actions.style_actions import wire_template_styles as _wire_styles
            result.artifacts["named_styles"] = _wire_styles(
                Path(result.transaction.output_path), Path(spec_source), result.spec)
        except Exception as e:  # noqa: BLE001
            result.artifacts["named_styles"] = {"_errors": [str(e)]}

    # 5.56) titlePg 一致化：有 first 页眉/页脚引用就必须有 titlePg，
    #       没有就必须去掉（否则首页留白或该显示的没显示）。非阻断。
    if result.transaction.status == "applied":
        try:
            from earlyz.actions.header_footer_actions import ensure_title_pg_flag as _ensure_tp
            result.artifacts["title_pg"] = _ensure_tp(
                Path(result.transaction.output_path), {})
        except Exception as e:  # noqa: BLE001
            result.artifacts["title_pg"] = {"error": str(e)}

    # 5.6) 结构修复段——与 planner 段同一事务、同一 journal（路径统一，
    #      2026-08-17 用户彻查拍板「只要一个路径」）。修复失败降级为
    #      journal 失败记录（不回滚已应用的格式修复）。
    #      ⚠️ 方向债务：下列若干步仍含取值常量（TOC 色/表题加粗/罗马页码）。
    #      只减不增。新需求走 decompiler 取值，禁再抄学校补丁。
    #      见技能 earlyz-direction-guard。
    #      顺序铁律：footers 重建必须先于 section_refs（脏 part 新增引用
    #      触发 Word 损坏）；toc_entries 在 sdt 展开后才能命中 hyperlink 段。
    structural_failures: list[dict] = []
    if result.transaction.status == "applied":
        from earlyz.spec.structure import structural_params as _st_params
        st = _st_params(result.spec) if result.spec is not None else {}
        structural_steps: list[tuple[str, dict]] = []
        # EZI-0065 A4：内容动作开始前先剥源文稿里的模板说明段（整段【…】/*…*
        # 包裹；真机案例：用户源里粘了「【表说明采用黑体；表头文字采用中文】」
        # 被原样排进产物版面）。保守三闸 + 熔断，详见 action docstring。
        structural_steps.append(("strip_instruction_paragraphs", {}))
        # EZI-0046（页脚页码字号恒 9pt）：期望值 = **模板页脚部件 w:sz 众数**，
        # 与判定层 `layout_fidelity._hf_stats` 同一个统计量（引擎与尺子取样必须
        # 同源，EZI-0045 教训）。旧取值链是 spec 单段首 run → 样式链 → 模板
        # Footer 样式 18 半磅 → 恒 9.0pt；两条路径的常量兜底一并删掉。
        # 模板页脚没有任何 w:sz 时 `_ftr_size_pt=None`：判定层此时判 N-A（无期望
        # 值），引擎不写字号、不拿常量顶替。
        _ftr_size_pt = _template_footer_page_size_pt(spec_source)
        result.artifacts["footer_page_size"] = {
            "template": str(spec_source) if spec_source else "",
            "size_pt": _ftr_size_pt,
        }
        # 2026-11-16（EZI-0034）：按内容角色建节——产物节数 = 属性变化点数（栏数/边距/
        # 页码体系），逐节复制模板该节的 cols/pgMar/pgSz/**w:type**。必须在页眉页脚
        # 绑定之前跑：节骨架定了，页脚（页码）与各节页眉才能按最终结构绑。
        # 2026-11-17（回工）**骨架优先**：老口径按角色标题切（摘要/Abstract/目录/参考文献
        # 各一刀），计划按「属性变化点」切。判定层是按内容判角色再配对的，边界落在角色
        # 标题上才配得准——北航本科实测：骨架 5 节 PASS，改由计划重锚定成 5 节后参考文献
        # 节拿到 `1134` top → 页边距 PASS→FAIL。所以骨架切得不比计划少时用骨架，计划只
        # 在骨架切不出来（或无骨架可用）时建节。兜底不能常开：本科参考会切出 8 节
        # （声明/摘要/Abstract/目录/正文/参考文献/致谢 各一节），Abstract 落到 1440 组。
        _plan_ok = False
        try:
            from earlyz.actions.section_geometry import apply_section_plan as _apply_plan
            _cand = Path(result.transaction.output_path)
            _tpl_src = Path(spec_source) if spec_source else Path("")
            _bounds = _detect_role_boundaries(_cand)
            _dry = _apply_plan(_cand, _tpl_src, dry_run=True)
            _planned = int(_dry.get("sections_planned") or 0)
            if len(_bounds) >= 2 and (len(_bounds) + 1) <= _planned:
                result.artifacts["section_plan"] = dict(_dry, skipped="prefer-skeleton")
                result.artifacts["section_skeleton"] = {
                    "boundaries": [b["heading"] for b in _bounds],
                    "template_sections": len(st.get("sections") or []),
                    "reason": "骨架切得不少于计划（角色标题优先）",
                }
            else:
                _pl = _apply_plan(_cand, _tpl_src)
                result.artifacts["section_plan"] = _pl
                _plan_ok = bool(_pl.get("changed")) or (
                    not _pl.get("errors") and not _pl.get("skipped")
                    and int(_pl.get("sections_planned") or 0) >= 2)
        except Exception as _e:  # noqa: BLE001 — 建节失败不阻断格式交付
            result.artifacts["section_plan"] = {"error": str(_e)}
        if not _plan_ok:
            # 2026-10-18（Layer 1）：单节文档 + 多节模板 → 先建节骨架，再套页眉页脚。
            # 依据是文档里的角色标题（精确匹配，不猜内容）；找不到 ≥2 个切点就跳过。
            try:
                _cand = Path(result.transaction.output_path)
                _doc_sects = _doc_section_count(_cand)
                _tpl_sects = len(st.get("sections") or [])
                if _doc_sects <= 1 and _tpl_sects >= 3:
                    _bounds = _detect_role_boundaries(_cand)
                    if len(_bounds) >= 2:
                        structural_steps.append(("build_section_skeleton", {"boundaries": _bounds}))
                        result.artifacts["section_skeleton"] = {
                            "boundaries": [b["heading"] for b in _bounds],
                            "template_sections": _tpl_sects,
                        }
            except Exception as _e:  # noqa: BLE001 — 建节失败不阻断格式交付
                result.artifacts["section_skeleton"] = {"error": str(_e)}
        if normalize_footer_parts:
            structural_steps.append(("normalize_footer_parts", {}))
        hf_parts = st.get("header_footer_parts") or []
        if hf_parts:
            structural_steps.append(("apply_template_header_footer", {
                "template_parts": hf_parts,
                # 模板实际绑定的槽位（含空 part）；缺省时动作内按有内容 part 推断（旧口径）
                "template_slots": st.get("hf_bound_slots") or [],
                # 源文档完全没有页眉时，按模板补「只带格式的空白页眉」（2026-10-17）
                "create_missing": True,
            }))
        # 2026-10-18（Layer 1）：没有页脚就没有页码。分节完成后补页脚并绑到每一节，
        # 封面（多节时的第 0 节）不绑页码。
        structural_steps.append(("ensure_footer_parts", {
            "format": "{PAGE}", "alignment": "center", "skip_first": True,
            # 字号只能来自模板页脚实测（EZI-0046）；无证据就不传，避免 9pt 常量兜底
            **({"size_pt": _ftr_size_pt} if _ftr_size_pt is not None else {}),
        }))
        if repair_section_headers:
            structural_steps.append(("repair_section_references", {
                "header_format": st.get("header_format") or {},
                "sections": st.get("sections") or [],
                "front_matter_roman": st.get("front_matter_roman", False),
                "identity_headers": st.get("identity_headers") or [],
                "body_headers": st.get("body_headers") or {},
            }))
        toc_params = st.get("toc") or {}
        if format_toc_entries and toc_params:
            structural_steps.append(("format_toc_entries", {"toc": toc_params}))
            structural_steps.append(("ensure_toc_tab_stops", {"toc": toc_params}))
        # 脚注/尾注取值（2026-09-25 批3）：模板有脚注素材才跑（无素材=空参数，
        # 动作内部整体跳过，目标文档脚注保持原样）。
        footnote_params = st.get("footnote") or {}
        if footnote_params:
            structural_steps.append(("format_footnotes", {"footnote": footnote_params}))
        # 脚注编号机制（2026-10-17）：模板明写「每页…从 1 开始」才写 numRestart
        _fn_restart = st.get("footnote_restart")
        if _fn_restart:
            structural_steps.append(("set_footnote_numbering", {"restart": _fn_restart}))
        # 标题 outlineLvl（Word 导航/大纲识别；2026-08-17 彻查第三轮）
        structural_steps.append(("ensure_heading_outline", {}))
        # 前部角标上标（2026-09-28 夜班，用户报的缺陷）：作者行 `张 三1)李 四2)`
        # 的编号是普通文字，而模板作者行/单位行样本的编号 run 带
        # `vertAlign="superscript"`。vertAlign 是「run 级且行内选择性」属性，
        # 段落级 extras 回放搬不过来（且被显式排除，防整行变上标）→ 单列动作。
        # 取值出自模板：只有模板 docx 源且样本确实带上标时才调度。
        _front_roles = _template_front_mark_roles(spec_source_type, spec_source)
        if _front_roles:
            structural_steps.append(("superscript_front_marks", {"roles": _front_roles}))
        # 标题前空段清理（2026-08-18 用户彻查「1.3 前两个空白行」）
        # 必须在 ensure_heading_outline 之后（标题识别依赖 outlineLvl）
        structural_steps.append(("remove_blank_lines_before_headings", {}))
        # 目录从零生成（2026-10-17 用户拍板「论文没目录时从零生成」）：模板注释给了
        # 目录规范才有 toc_params；文档已有目录则动作 no-op。必须在标题 outlineLvl
        # 打完后（条目=正文标题）、update_toc_entries/ensure_toc_field 之前。
        if format_toc_entries and toc_params:
            structural_steps.append(("insert_toc_if_missing", {
                "enabled": True,
                "toc": toc_params,
                "toc_title": st.get("toc_title") or {},
            }))
            # 刚建出的条目补跑样式（幂等；已有目录时是第二次无害执行）
            structural_steps.append(("format_toc_entries", {"toc": toc_params}))
            structural_steps.append(("ensure_toc_tab_stops", {"toc": toc_params}))
        # MathType 章节断点（公式组件 W3，2026-10-01）：编号域链只读章号、全文无断点时，
        # 按章标题补隐藏断点（Word 才显示 (4.1) 而不是 (0.1)）。机制层：章号出自文档
        # 自己的章标题，编号格式沿用文档域链；文档已有断点则不动。必须在
        # ensure_heading_outline 之后（章标题识别依赖 outlineLvl）。无 MathType 编号时 no-op。
        structural_steps.append(("insert_equation_chapter_breaks", {}))
        if format_toc_entries and toc_params:
            # TOC 条目文本重算（2026-08-18 用户彻查「结论 vs 结论与展望」：
            # TOC 域缓存是源文档旧的——earlyz 自算替代 Word updateFields
            # 覆盖格式）。必须最后跑——format_toc_entries 先定型式；
            # ensure_heading_outline 先打 outlineLvl（TOC 序列依赖）。
            structural_steps.append(("update_toc_entries", {}))
            # TOC 域保活（2026-09-26 用户拍板「能自动更新」）——
            # 条目文本已由 update_toc_entries 算好，这里补回域外壳：
            # 打开显示我们算的条目（不弹窗，updateFields 由后面那步清），
            # 按 F9 可重建。必须最后跑（先算好文本再包域）。
            structural_steps.append(("ensure_toc_field", {}))
        # 前置章节（摘要/ABSTRACT/目录/致谢等）另起页——同节时补分页
        # 分页要求只从模板取（2026-10-06）：模板给了 section_break_words
        # （哪怕空集）就按模板执行；键缺失=无模板信息 → 动作侧旧口径兜底。
        _sbw = st.get("section_break_words")
        structural_steps.append(("ensure_section_page_breaks",
                                 {"section_break_words": _sbw} if _sbw is not None else {}))

        # 分节可能在上一步才生成 → 页眉引用要再补一次（动作幂等）
        if hf_parts:
            structural_steps.append(("apply_template_header_footer", {
                "template_parts": hf_parts,
                "template_slots": st.get("hf_bound_slots") or [],
                "create_missing": True,
            }))
        # settings.xml 移除 w:updateFields——不弹「更新域」弹窗
        # （2026-08-18 用户彻查；用户要求排版前提醒而非产物弹窗）
        structural_steps.append(("ensure_update_fields", {}))
        if format_captions:
            structural_steps.append(("format_caption_fonts", {
                "caption": st.get("caption") or {},
                "equation": st.get("equation") or {},
            }))
            structural_steps.append(("format_caption_numbers", {"caption": st.get("caption") or {}}))
            # 表格题注编号样式（2026-10-17）：模板明写罗马数字才转
            _tns = (st.get("caption") or {}).get("table_number_style")
            if _tns == "roman":
                structural_steps.append(("set_table_caption_numbering", {"style": "roman"}))
            # 表标题段 keepNext 防跨页（08-08 铁律②；catalog 有但从未接线
            # ——工具箱≠交付第四例，2026-08-17 彻查补齐）
            structural_steps.append(("normalize_table_caption_block", {}))
        # 图表题注交叉引用（2026-10-18：引擎此前只在 _next_steps 提示「可做」，从不做）
        structural_steps.append(("create_caption_cross_refs", {}))
        # 参考文献参数不存在时不调度：no-op handler 后接 verifier 会把
        # “模板无此要求”误报为结构失败。
        refs_params = st.get("refs") or {}
        if refs_params:
            structural_steps.append(("format_references", {"refs": refs_params}))
        if format_table_cells:
            structural_steps.append(("apply_section_columns", {
                "body_cols": st.get("body_cols") or {},
                "front_cols": st.get("front_cols") or {},
                "identity_headers": st.get("identity_headers") or [],
                "body_section_anchor": st.get("body_section_anchor") or "",
                "body_section_anchor_kind": st.get("body_section_anchor_kind") or "",
                "front_tokens": list(st.get("front_tokens") or []),
                "section_plan": list(st.get("section_plan") or []),
            }))
            # 分栏后再封一次表宽：分栏把每栏可用宽度砍半，之前按版心封顶的表会溢出（2026-10-19）
            structural_steps.append(("fit_tables_to_page", {}))
                # 表格封顶到节版心宽（2026-10-19 期刊模板：宽表超出页面 131pt）
            structural_steps.append(("fit_tables_to_page", {}))
            structural_steps.append(("normalize_table_widths", {"table_width": st.get("table_width") or {}}))
            structural_steps.append(("format_table_cells", {"table_cells": st.get("table_cells") or {}}))
            structural_steps.append(("apply_table_borders", {"table_borders": st.get("table_borders") or {}}))
            # 整表同页（2026-09-26 用户：表格分页是当前最明显短板）——必须在所有
            # 表格结构动作之后跑：全行 cantSplit + 能装下则 keepNext 全链锁同页。
            structural_steps.append(("keep_table_together", {}))
            # 表单表单元格水平+垂直居中（2026-09-26 用户定案：只动对齐，段落
            # 内容与正文一致）——必须在表格结构动作之后，避免居中前改列宽。
            structural_steps.append(("center_form_table_cells", {}))
        # C 通路（W8 Toggle TeX，2026-10-06）：`$…$` / `$$…$$` / `\[…\]` 文本 →
        # MathType 对象。确定性自动转（显式定界符即用户意图）；货币等非数学内容
        # 由 _looks_like_math 守卫跳过并如实通报。必须在公式段定型**之前**——
        # 转换出来的公式段要被定型步按公式段排（缩进归零/编号制表位）。
        # 公式：模板「公式尺寸」写明的标准字号透传给新生成对象（2026-10-17 接线）
        _eq_fs = (st.get("equation") or {}).get("font_size_pt")
        structural_steps.append(("toggle_tex_equations",
                                 {"font_size_pt": _eq_fs} if _eq_fs else {}))
        # 公式段定型（缩进归零 + 居中 + 编号制表位 + 超长缩字号）。
        # ⚠️ 顺序铁律（2026-09-28 真机定位，勿挪！）：必须在 `apply_section_columns`
        # **之后**。公式段的可用宽度与编号制表位是按「所属节的栏宽」算的；分栏步若排
        # 在后面，公式步拿的是**尚未定稿**的宽度 → 第一遍判「放得下」、第二遍才发现
        # 超宽要缩（真论文实测 pass1 shrunk=0 / pass2 shrunk=2，幂等被打破）。
        # 模板**有**公式样本时由 spec 规则处理（decompiler _ROLE_KEYS["equation_display"]），
        # 此处跳过 —— 取值一律以模板为准。
        _has_eq_rule = any(
            (getattr(r, "selector", None) or {}).get("semantic_role") == "equation_display"
            for r in (result.spec.rules if result.spec else []))
        if not _has_eq_rule:
            structural_steps.append(("normalize_equation_paragraphs", {}))
        else:
            # 2026-10-08 真机缺陷（用户报「公式没被处理」直查）：模板有公式样本时
            # 此步整体跳过 → 连**纯文本公式检测**（B 通路通报「识别到 N 条纯文本公式，
            # 回复『公式转对象』」）都被卡死，用户从未收到转换入口提示。检测与定型
            # 解耦：detect_only=只统计不改文档，通报照发；定型取值仍以模板为准。
            #
            # 2026-10-10 真机缺陷（用户报「公式编号不在最右、悬在行中」直查）：
            # detect_only 让编号落位彻底失管——spec 规则只表达 jc/缩进/字体等取值，
            # **表达不了「编号右贴版心缘」的制表位机制** → 编号停在源文档原处
            # （全角空格粘公式尾）。分权：取值类一律关（以模板为准：居中/字体/行距/
            # 缩字号），编号制表位照跑（通用机制、模板无关）；纯文本检测统计不变
            # （仍返回 equations_seen/plain_seen，B 通路通报照发）。
            structural_steps.append(("normalize_equation_paragraphs", {
                "center": False, "fonts": False, "line_spacing": False,
                "fit_width": False, "number_tab": True,
            }))
        drop_cap = {}
        if result.spec is not None:
            for _r in result.spec.rules:
                if (_r.selector or {}).get("semantic_role") == "drop_cap":
                    drop_cap = dict(_r.expected or {})
                    break
        if drop_cap.get("drop_cap"):
            structural_steps.append(("apply_drop_cap", {
                "properties": drop_cap,
                "layout": st.get("drop_cap_layout") or {},
            }))
        # 图片尺寸适应（2026-08-23 CJC 实锤：双栏 415pt 图被裁——按栏宽缩放）
        structural_steps.append(("fit_media_to_page", {}))
        # 图片段 keepLines/keepNext 防跨页+居中（08-08 铁律②；同上补接线）
        structural_steps.append(("normalize_figure_block", {"figure_width": st.get("figure_width") or {}}))
        # 半角标点低风险子集自动归一（2026-10-08 用户报「序号没处理」——(1)(2)(3)
        # 等中文语境数字括号此前不转全角；B2 口径已定「自动改 + 三保险」：对照表=
        # changes 账本随 structural_fixes 落结果、可撤=journal/备份、不静默=计数明细
        # 可见）。顺序：内容类动作全部完成后、改分页的动作之前（页码写回须反映最终
        # 文本）。保护区（参考文献/英文摘要/代码段）与英文语境一律不动。
        structural_steps.append(("normalize_cjk_punctuation", {}))
        # 目录 PAGEREF 真实页码写回（必须最后：前面所有动作都可能改分页）
        structural_steps.append(("refresh_page_fields", {}))
        # 目录收尾（2026-09-26 用户定案：目录放编译链最后一步 —— 先排完其他
        # 内容、条目与页码先正确、样式最后上；缩进交回 TOC1/TOC2 样式、
        # 页码右对齐版心右缘 + 点线前导符。样式由最后一步说了算，通用段落
        # 规则不再有权覆盖目录条目段。）
        structural_steps.append(("polish_toc_entries", {}))
        # 页眉引用收尾（2026-10-17）：正文分栏/公式分节等步骤会在中途新增
        # continuous 分节，早先绑好的 headerReference 不会跟到新节上 →
        # 动作幂等，放最后再补一次，保证每一节都有页眉（哪怕空格式页眉）。
        if hf_parts:
            structural_steps.append(("apply_template_header_footer", {
                "template_parts": hf_parts,
                "template_slots": st.get("hf_bound_slots") or [],
                "create_missing": True,
            }))
        # 页眉绑定之后再把共用部件拆成每节独立（否则「只改封面页眉」波及全文）
        structural_steps.append(("split_section_parts", {}))
        # 封面/声明节页眉必须是空的：模板绑定可能把「致谢」页眉绑到封面
        # （2026-10-18 真机 7 分 45 秒后 delivery 判 cover polluted 的直接原因）
        structural_steps.append(("strip_cover_header_text", {}))
        # 前置页（封面/声明/授权）去页眉页脚 + 页码按角色分段（2026-10-18 用户实测）
        structural_steps.append(("normalize_front_matter", {}))
        # 模板页眉里写着的角色字样（摘要/ABSTRACT/目录/参考文献/致谢/附录）落到对应节；
        # 只搬模板原文，含 XXXX 占位样例的跳过（2026-10-18：此前 14 个页眉全是空壳）
        _role_texts: dict[str, str] = {}
        for _hp in (hf_parts or []):
            if (_hp.get("kind") or "") != "header":
                continue
            _core = str(_hp.get("text_core") or "").strip()
            if _core in ("摘要", "ABSTRACT", "目录", "参考文献", "致谢", "附录"):
                _role_texts[_core] = _core
        # 2026-10-18 二次回滚：真实文档上「同一段落被连续分节切成多节」时，
        # 角色启发式仍会把「致谢」写到封面节（部件在各节间共享），交付门判
        # cover polluted。正确做法是在 build_section_skeleton 建节时把角色
        # 一并落到该节，而不是事后按首段猜。先不自动跑。
        if False and _role_texts:
            structural_steps.append(("apply_section_header_text", {"role_texts": _role_texts}))
        # 页码格式按模板取值（模板 footer PAGE 段 jc=center / 字体）。
        # EZI-0046：字号**不再**从 spec 单段（首 run → 样式链）拿——那条链在
        # 劳动合同/节水/重庆科技上恒返回 9.0pt（模板 Footer 样式 18 半磅），
        # 而判定层期望是模板页脚全部 w:sz 的众数（14.0 / 10.5）。字号改由
        # `_ftr_size_pt`（模板实测众数）注入，`format_footers` 写到每个 part
        # 的每个 run（含空壳 run）。模板无 w:sz 时不写、不兜底。
        _pn_props = _footer_page_number_props(hf_parts, _ftr_size_pt)
        if _pn_props:
            structural_steps.append(("format_footers", {"properties": _pn_props}))
        # 模板绑定过 even 槽位就必须开 evenAndOddHeaders，否则奇偶页眉永远不生效
        if any("even" in str(s) for s in (st.get("hf_bound_slots") or [])):
            structural_steps.append(("set_even_odd_headers_flag", {"enabled": True}))
        # 标题自动编号收口（2026-11-08 用户实锤「1. 1. 引言」「二三级 1.27cm」）：
        # ① 模板样式带 numPr，但模板的 numbering 定义不会随样式一起进产物 →
        #    悬空 numId 让 Word 按默认列表兜底（多画一个号 + 默认 720twips 缩进）；
        # ② 文本里已有手打序号的标题再压掉自动编号（内容不动）。
        # 顺序：先补定义（否则先压制会把「本该自动编号」的层级也一起关掉），
        # 再压制；两步都幂等。
        structural_steps.append(("reconcile_numbering_refs",
                                 {"template": str(spec_source) if spec_source else ""}))
        structural_steps.append(("suppress_duplicate_heading_numbers", {}))
        # EZI-0068：表内段落标题残留——heading 样式单元格段级 outlineLvl=9
        # 覆盖（suppress 家族只扫正文直挂段，表内段是其盲区；探针按
        # OutlineLevel 枚举 doc.Paragraphs 含表格内段）
        structural_steps.append(("clear_table_cell_outline", {}))
        # 目录与正文分节（2026-11-08 用户实锤「目录和正文在一页」）：**必须放在
        # 结构段最后**——前面的 normalize_front_matter / split_section_parts
        # 会重建前置区的节，早放会被覆盖成 continuous（实测产物里分节符变成
        # continuous，目录与正文仍同页）。分节符挂在目录末条后的空段上，
        # F9 更新目录冲不掉它（PBB 会被冲掉）。
        structural_steps.append(("ensure_toc_section_break", {}))
        # 5.6a) 题注抽流（EZI-0013，A1/A2）——**必须在题注字体/编号动作之前**：
        #   `format_caption_fonts` / `format_caption_numbers` / `renumber_*` 都只认
        #   「流内段落」，题注还锁在图对象文本框里时它们看不见（实测 30 条漏 30 条）。
        #   抽流后题注才是普通段落，编号/字号/加粗/图题绑定（5.65）才有对象可改。
        #   直调不走进事务：该动作尚未进 catalog registry（catalog.py 不在本 issue
        #   的 write_scope 内），与 5.65 的 `bind_caption_objects` 同一条直调路径。
        if format_captions:
            try:
                from earlyz.actions.captions import (
                    extract_caption_textboxes as _extract_caps,
                )
                _capx = _extract_caps(Path(result.transaction.output_path),
                                      {"apply": True})
                result.artifacts["caption_extraction"] = _capx
                if _capx.get("changed"):
                    # 文本 hash 必然变（题注从文本框搬进正文流 → 字符序列随段落
                    # 位置改变）。按 §10.1 显式授权：不声明就会被
                    # validate_delivery 判「content preservation FAILED」拦交付。
                    authorized_content_actions.append("extract_textbox_captions")
            except Exception as e:  # noqa: BLE001 — 非阻断（与其它结构修复一致）
                result.artifacts["caption_extraction"] = {"error": str(e)}
        structural_summary: dict = {}
        if structural_steps:
            post_plan = RepairPlan(
                plan_id=f"{result.plan.plan_id}#structural",
                actions=[RepairAction(action=a, params=p, risk="R1",
                                      source="结构修复段（compile 显式）")
                         for a, p in structural_steps],
            )
            from pathlib import Path as _P
            _tx = result.transaction
            post_result = TransactionExecutor(reg, work).execute_post_actions(
                post_plan,
                src,
                tx_id=_tx.transaction_id,
                candidate=_P(_tx.output_path),
                backup=_P(_tx.backup_path),
                journal_path=_P(_tx.journal_path),
            )
            result.transaction = post_result
            # 兼容字段（artifacts 结构保持，供报告/测试读取）
            for entry in post_result.action_results:
                info = entry.get("info") or {"changed": True}
                prev = structural_summary.get(entry["action"])
                if isinstance(prev, dict) and isinstance(info, dict):
                    # 同一动作可能跑多次（页眉补建）→ 计数字段取最大值，别被后一次 0 覆盖
                    merged = dict(prev)
                    for k in ("parts_created", "refs_bound", "parts_formatted"):
                        try:
                            merged[k] = max(int(prev.get(k) or 0), int(info.get(k) or 0))
                        except Exception:  # noqa: BLE001
                            merged[k] = info.get(k, prev.get(k))
                    structural_summary[entry["action"]] = merged
                else:
                    structural_summary[entry["action"]] = info
            journal_rows = json.loads(
                _P(_tx.journal_path).read_text(encoding="utf-8"))["actions"]
            # 2026-10-14（复现事故：授权静默丢失根因）：verifier 异议的动作只进
            # journal、不进 action_results（executor 的 continue 分支），但它的
            # info（含 changed 等计数）在 journal 里完好。授权/报告判断必须能
            # 看到失败动作的真实改动——否则「动作改了内容、复查器喊了一声」会
            # 静默变成「未授权内容变化」，两条硬错误齐发拦掉整次交付。
            for row in journal_rows:
                name = row.get("action")
                if (name and name not in structural_summary
                        and row.get("info") is not None):
                    structural_summary[name] = row["info"]
            journal_failed = [a for a in journal_rows
                              if a.get("status") == "failed"]
            if journal_failed:
                structural_failures = journal_failed
                structural_summary["failed_actions"] = [
                    {"action": a["action"], "error": a.get("error")}
                    for a in journal_failed]
            result.artifacts["structural_fixes"] = structural_summary

    # 5.65) 图题绑图 / 表题绑表（EZI-0011）——**必须落在结构段之后**：
    #   `normalize_table_caption_block` 在结构段里跑，它的正则把「图 / 表」写在
    #   一起（`^(表|Table|图|Figure|Fig)`），于是图题也被加上了 keepNext——方向
    #   是反的（keepNext 绑「自己 + 后面内容」，图题的后面是正文，不是图）。放在
    #   结构段之后才能把它改对：图题段去掉 keepNext、改绑前面的图片段落；表题
    #   段补齐 keepNext 并拆掉与表格之间的纯空段。顺序也保证在页码写回
    #   （refresh_page_fields）之后不再改动内容流——改动只是段落属性 + 拆纯
    #   空段，文字 hash 不受影响（分析器的 content_hash 不计空段）。
    #   cantSplit / tblHeader 一律不碰，动作结果里带修前/修后计数当证据（A4）。
    if result.transaction.status == "applied" and format_captions:
        try:
            from earlyz.actions.captions import bind_caption_objects as _bind_caps
            result.artifacts["caption_binding"] = _bind_caps(
                Path(result.transaction.output_path), {"apply": True})
        except Exception as e:  # noqa: BLE001 — 非阻断（与其它结构修复一致）
            result.artifacts["caption_binding"] = {"error": str(e)}

    # 5.7) 内容变化授权（2026-10-08 起；2026-10-14 审计扩展）：
    #      受控内容变化动作（半角→全角标点、$…$ 文本转 MathType 对象、
    #      模板示例残留清理等）属 §10.1 例外：ActionSpec 声明 content_change +
    #      本次运行授权 → validate_delivery 把「文本 hash 变化」降为
    #      warning + 证书记录；其余任何非预期文本变化照常拦截。
    #      授权判据 = 动作**确实改了内容**（info 计数真值），与它是否被
    #      verifier 判过失败无关（失败动作的改动同样留在产物里，见 5.6 段）。
    _sf = result.artifacts.get("structural_fixes") or {}
    from earlyz.actions.catalog import CONTENT_CHANGE_ACTIONS as _CCA
    _CHANGE_KEYS = ("changed", "converted", "added", "numbered", "renumbered",
                    "display_converted", "inline_converted")
    for _name in sorted(_CCA):
        _info = _sf.get(_name) or {}
        if any(_info.get(k) for k in _CHANGE_KEYS):
            authorized_content_actions.append(_name)
    _tt = _sf.get("toggle_tex_equations") or {}
    if (_tt.get("display_converted") or 0) + (_tt.get("inline_converted") or 0):
        authorized_content_actions.append("toggle_tex_equations")
    authorized_content_actions = list(dict.fromkeys(authorized_content_actions))

    # 6) Validate（四层门禁；内容保持对比原文件）
    # 5.88) 跨页续表标注（P3b）——默认关闭：label 由 spec 的 table_continued_label
    #       声明，未声明即 no-op（加字属内容变化，必须显式授权）。
    if result.transaction.status == "applied":
        try:
            from earlyz.actions.new_actions import mark_continued_tables as _mark_cont
            _label = ""
            try:
                _label = (getattr(result.spec, "table_continued_label", "") or "") if result.spec else ""
            except Exception:
                _label = ""
            result.artifacts["continued_tables"] = _mark_cont(
                Path(result.transaction.output_path), {"label": _label})
        except Exception as e:  # noqa: BLE001
            result.artifacts["continued_tables"] = {"error": str(e)}

    # 5.84) 目录标题（「目录」二字）补齐章节标题样式。
    # 2026-10-22 实测：源文档的「目录」段落只有 outlineLvl=0，无 pStyle、无字号 →
    # 渲染成 Normal 小字，与「摘要」「致谢」明显不一致。
    # 口径：复制同文档第一个一级标题的 pPr/rPr（取值来自文档自身，不硬编码字号）。
    if result.transaction.status == "applied":
        try:
            from earlyz.actions.new_actions import style_toc_title as _style_toc
            result.artifacts["toc_title_style"] = _style_toc(
                Path(result.transaction.output_path), {})
        except Exception as e:  # noqa: BLE001
            result.artifacts["toc_title_style"] = {"error": str(e)}

    # 5.86) 「第 N 节 / 第 N 条」引用（P3，默认 dry-run 只出清单）：
    #       合同类大量使用「第 N 条」，但条文编号跨文档不稳定，先只报告不改，
    #       待用户确认后再开 apply（单独开关，符合 P3「中风险、单独决策」口径）。
    if result.transaction.status == "applied":
        try:
            from earlyz.actions.cross_refs import link_section_references as _link_secs
            result.artifacts["section_refs"] = _link_secs(
                Path(result.transaction.output_path), {"apply": False})
        except Exception as e:  # noqa: BLE001
            result.artifacts["section_refs"] = {"error": str(e)}

    # 5.87) 数据表数值列对齐（P1，2026-10-21）：kind=data 且数值列 → 右对齐 + 数字不换行。
    #       纯格式改动；kind=form 的表完全不动。
    if result.transaction.status == "applied":
        try:
            from earlyz.actions.new_actions import align_data_table_columns as _align_tbl
            result.artifacts["data_table_align"] = _align_tbl(
                Path(result.transaction.output_path), {})
        except Exception as e:  # noqa: BLE001
            result.artifacts["data_table_align"] = {"error": str(e)}

    # 5.85) 章节号引用（P0，2026-10-21 用户拍板默认开启）：
    #       把正文「第 N 章」包成 REF 域，域结果保留原文字——不改字、显示不变，
    #       章号改动后 Word 里 F9 联动。找不到对应章标题的提及只报不改（不发明编号）。
    if result.transaction.status == "applied":
        try:
            from earlyz.actions.cross_refs import link_chapter_references as _link_chaps
            result.artifacts["chapter_refs"] = _link_chaps(
                Path(result.transaction.output_path), {"apply": True})
        except Exception as e:  # noqa: BLE001
            result.artifacts["chapter_refs"] = {"error": str(e)}

    # 5.9) 最终表格封顶：必须在分栏/结构动作全部落地之后再跑一次。
    # 2026-10-21 实测（建设工程监理合同等 4 份）：模板本身单栏，spec 却排出
    # 2 栏节（栏宽 195pt），表格声明宽 458pt 被塞进窄栏 → 渲染到页外，
    # image/table bbox overflow（本次真机看图确认表格右缘被裁）。
    # fit_tables_to_page 是栏感知的，但此前只在分栏之前跑，拿到的是整幅版心宽。
    if result.transaction.status == "applied":
        try:
            from earlyz.actions.new_actions import fit_tables_to_page as _fit_tables_final
            result.artifacts["final_table_fit"] = _fit_tables_final(
                Path(result.transaction.output_path), {})
        except Exception as e:  # noqa: BLE001
            result.artifacts["final_table_fit"] = {"error": str(e)}

    # 5.95) 跨页续表（P3b 真正生效版）：只有 spec 声明了 table_continued_label 才做。
    # 先渲染探测「每行落在第几页」→ 迭代断表 + 续页表头写「续表」→ 清孤儿表/空页。
    # 未声明即零成本跳过（不渲染、不改动）。
    # 续表拆分依赖 Word/LibreOffice 的分页探测；没有渲染器时探测结果不可信
    # （2026-10-28 全库基线实测：headless + 并发 → 17 行表被拆成 16 张单行表）。
    _renderer_available = bool(word_render_callback) or bool(libreoffice_path) or bool(print_mode)
    if result.transaction.status == "applied" and _renderer_available:
        _cont_label = ""
        try:
            _cont_label = (getattr(result.spec, "table_continued_label", "") or "") if result.spec else ""
        except Exception:
            _cont_label = ""
        if not _cont_label:
            # 未显式声明时按文档族兜底（用户 2026-10-21 拍板：只对学位论文/期刊开）
            try:
                from earlyz.validation.doc_family import continued_label_for as _cont_for
                _cont_label = _cont_for(spec_source)
            except Exception:
                _cont_label = ""
        if _cont_label:
            try:
                from earlyz.actions.new_actions import apply_continuation_splits as _split_cont
                _cs = _split_cont(Path(result.transaction.output_path), _cont_label)
                result.artifacts["continued_split"] = _cs
                # 授权判定（CONTENT_CHANGE_ACTIONS 扫描 structural_fixes）需要看到它
                _sfx = result.artifacts.setdefault("structural_fixes", {})
                _sfx["apply_continuation_splits"] = {
                    "changed": bool((_cs or {}).get("changed") or (_cs or {}).get("split_count")),
                    "split_count": (_cs or {}).get("split_count", 0),
                }
            except Exception as e:  # noqa: BLE001
                result.artifacts["continued_split"] = {"error": str(e)}

    # 续表插入「原表题（续）」属显式授权的加字行为；授权名单在上游已定稿，
    # 此处按实际结果补登记（否则内容保持门会判 text hash changed → FAILED）。
    _csi = result.artifacts.get("continued_split") or {}
    if isinstance(_csi, dict) and (_csi.get("changed") or (_csi.get("split_count") or _csi.get("splits"))):
        authorized_content_actions.append("apply_continuation_splits")
        authorized_content_actions = list(dict.fromkeys(authorized_content_actions))

    # 5.97) 包级清理（2026-11-07）：正文里已无 `w:object` 的 OLE 部件与关系是死重量。
    # 实测 42/199 份产物带孤儿 embeddings（大连民族 30 个 / 矿大 13 个），
    # 包体虚胖数 MB；关系指向的部件一旦缺失，Word 会弹「需要修复」。
    # 只在 embeddings 内做，保守判定「被引用」= 包内任何 XML 部件里出现过该关系 Id。
    if result.transaction.status == "applied":
        try:
            from earlyz.actions.package_cleanup import drop_orphan_embeddings as _drop_emb

            result.artifacts["package_cleanup"] = _drop_emb(Path(result.transaction.output_path))
        except Exception as e:  # noqa: BLE001 — 非阻断
            result.artifacts["package_cleanup"] = {"error": str(e)}

    if result.transaction.status == "applied":
        target_for_validation = result.transaction.output_path

        # 6.4) 模板页眉构件继承（带渲染门控）：模板页眉里的校徽/校名是版式的
        # 一部分，产物源自裸内容（无页眉）时要带过来。但「部件里带图」不等于
        # 「模板上看得见」——先渲染模板量页眉区像素，确认可见才继承。
        try:
            _tpl_path = Path(spec_source)
            if _tpl_path.suffix.lower() == ".docx" and _tpl_path.exists():
                from earlyz.actions.template_furniture import inherit_template_headers

                result.artifacts["template_furniture"] = inherit_template_headers(
                    target_for_validation, _tpl_path,
                    render_fn=word_render_callback, gate=True)
        except Exception as _e:  # noqa: BLE001
            result.artifacts["template_furniture"] = {"error": str(_e)}

        # 6.5) 打印版最终化（IR 语义识别正文起始节，复用验证过的 Word 分页循环）
        if print_mode:
            from earlyz.pipeline.print_layout import apply_print_layout

            print_out = work / f"{src.stem}-print.docx"
            apply_print_layout(
                target_for_validation,
                print_out,
                libreoffice_path or "",
                force_odd_from_section=force_odd_from_section,
                ir=result.ir,
            )
            target_for_validation = print_out
            result.artifacts["print_output"] = str(print_out)

        _cover_snap = None   # 重建后的封面 pPr 快照（供 6.48 回填）

        # 6.4) 封面自动重建（用户 2026-10-18 定案）：模板有封面区块且槽位取值
        # 齐全 → 直接跑；只有缺值（英文题目/指导教师等源文档没提供）才提示确认。
        try:
            if spec_source and str(spec_source).lower().endswith((".docx", ".doc")):
                from earlyz.actions import cover as _cover_mod

                _slots = _extract_cover_slots(src)
                _zres = _cover_mod.build_cover_from_template(
                    target_for_validation,
                    {"template_path": str(spec_source), "apply": False,
                     "zones": ["cover", "statement"], "slots": _slots})
                _zones = _zres.get("zones_found") or []
                _names = [str(z.get("name")) for z in _zones if isinstance(z, dict)]
                _labels = [s for z in _zones if isinstance(z, dict) for s in (z.get("slots") or [])]
                _missing = _cover_slots_missing(src, _labels) if _labels else []
                # 占位符值（XXX / XXXX）等同缺值：绝不能把模板样例写进用户封面
                _ph = [x for x in _labels if re.search(r"X{2,}", str(x))]
                # 2026-10-18 实测：自动重建后模板样例标题（"Altera DDR IPCore…"）
                # 没有被替换成源文档题目——缺「题目」槽位映射。宁可先不自动落盘，
                # 也不能交付带模板样例文字的封面。
                # 2026-10-22 定案：封面/声明是「产品托管区」——用模板版式**替换**
                # 源封面（force=True + 显式区段），不是叠加、也不是跳过。
                # 之前的两种错法都实测过：叠加 → 日期行/签名行重复；跳过 → 封面不生效。
                _map = {}
                try:
                    _root_b, _blocks = _cover_mod._target_blocks(target_for_validation)
                    _map = _cover_mod._find_target_map(_blocks)
                except Exception:
                    _map = {}
                if "cover" in _names and not _missing:
                    # 兜底：_find_target_map 认不出封面块时（P3 源文档实测），
                    # 用「文首→第一个前置标题之前」作为封面替换区段；
                    # 否则 cover.py 会静默退化成「插到文首不删旧封面」→ 内容重复。
                    if not _map.get("cover"):
                        _rng = _front_cover_range(Path(target_for_validation))
                        if _rng:
                            _map = dict(_map)
                            _map["cover"] = _rng
                    # 表单式封面（区段内含表格）不能整段替换：会把表格连同内容删掉。
                    _rng_chk = (_map or {}).get("cover")
                    if _rng_chk and _cover_range_has_table(Path(target_for_validation), _rng_chk):
                        result.artifacts["cover_auto"] = {
                            "applied": False,
                            "reason": "封面区段含表格（表单式封面），保留源文档原样不重建",
                            "range": list(_rng_chk)[:2],
                        }
                    else:
                        # 只重建封面；声明区（本人签名/日期）插一份就多一份，保持源文档原样。
                        _done = _cover_mod.build_cover_from_template(
                            target_for_validation,
                            {"template_path": str(spec_source), "apply": True,
                             "force": True, "replace": _map,
                             "zones": ["cover"], "slots": _slots})
                        result.artifacts["cover_auto"] = {"applied": True, "result": _done}
                        try:
                            from earlyz.actions.cover_guard import (
                                snapshot_cover_ppr as _snap_cover,
                            )

                            _cover_snap = _snap_cover(target_for_validation)
                        except Exception:  # noqa: BLE001
                            _cover_snap = None
                else:
                    result.artifacts["cover_auto"] = {
                        "applied": False, "missing_slots": _missing,
                        "zones": _names,
                        "reason": ("槽位缺值" if _missing else
                                   ("模板无封面区块" if "cover" not in _names else
                                    "封面题目槽位未映射，暂不自动重建（防模板样例文字外泄）")),
                    }
        except Exception as _e:  # noqa: BLE001
            result.artifacts["cover_auto"] = {"applied": False, "error": str(_e)[:200]}

        # 6.48) 封面完整性兜底（2026-10-26）：本次没有重建封面时，交付前把封面段落
        # 还原成源文档原样。正文规则（首行缩进/行距/标点归一/编号）不得污染封面——
        # 199 份真实场景全量测试中 cover polluted 是失败大头，逐条动作打补丁既漏又脆。
        try:
            from earlyz.actions.cover_guard import (
                restore_cover_from_source as _restore_cover,
                sanitize_heading_outline as _sanitize_headings,
            )

            result.artifacts["heading_sanitize"] = _sanitize_headings(target_for_validation)
            if _cover_snap:
                # 重建过封面：文案来自模板，不能整段还原成源文档，
                # 只把段落属性回填成重建刚完成时的样子（锁住缩进/对齐/行距）。
                from earlyz.actions.cover_guard import restore_cover_ppr as _restore_ppr

                result.artifacts["cover_restore"] = _restore_ppr(target_for_validation, _cover_snap)
            elif not (result.artifacts.get("cover_auto") or {}).get("applied"):
                result.artifacts["cover_restore"] = _restore_cover(target_for_validation, src)
            # 封面还原可能把源文档的 wp:docPr id 带回来，与产物里已重编号的
            # 图片/文本框撞车 → 交付门判 duplicate docPr id（2026-10-26 西工大
            # 英文版实测 id=11）。还原后立刻做一次包级不变量修复。
            try:
                from earlyz.actions.new_actions import repair_package_invariants as _repair_pkg

                result.artifacts["package_invariants"] = _repair_pkg(target_for_validation, {})
            except Exception as _rp:  # noqa: BLE001
                result.artifacts["package_invariants"] = {"error": str(_rp)[:200]}
            # 封面还原会把源文档封面里的布局表格一并带回来，宽度可能超过当前
            # 模板版心 → 渲染后判 image/table bbox overflow（2026-10-26 东南本科
            # 实测 p1 text_block 右边缘 845.6 / 页宽 595.3）。5.9 的表格封顶在
            # 还原之前跑，必须在这里再封一次。
            try:
                from earlyz.actions.new_actions import fit_tables_to_page as _fit_tables2

                result.artifacts["cover_table_fit"] = _fit_tables2(target_for_validation, {})
            except Exception as _ft:  # noqa: BLE001
                result.artifacts["cover_table_fit"] = {"error": str(_ft)[:200]}
            # 封面表单的 noWrap 单元格会把长句拉出页面（东南本科 AI 说明表实测），
            # 放开长文本单元格的禁止换行。
            try:
                from earlyz.actions.cover_guard import relax_cell_nowrap as _relax_nowrap

                result.artifacts["cell_nowrap"] = _relax_nowrap(target_for_validation)
            except Exception as _rn:  # noqa: BLE001
                result.artifacts["cell_nowrap"] = {"error": str(_rn)[:200]}
        except Exception as _cr:  # noqa: BLE001
            result.artifacts["cover_restore"] = {"applied": False, "error": str(_cr)[:200]}

        # 6.49) 大纲级别不得新增（EZI-0014，A1 硬不变量）：产物 outlineLvl 段落集合
        # 必须 ⊆ 源文档 outlineLvl 段落 ∪ 真标题集合。放在封面还原之后——还原会
        # 把源文档封面段（含原文自带的大纲级别）整段带回来，必须是「带回来之后」
        # 再看有没有新增。真标题（交大中期 22pt 题名）与源文档自己的标记原样保留。
        try:
            from earlyz.actions.cover_guard import (
                strip_invented_outline_levels as _strip_outline,
            )

            result.artifacts["outline_invented"] = _strip_outline(target_for_validation, src)
        except Exception as _ol_e:  # noqa: BLE001 — 非阻断守卫：失败只记录
            result.artifacts["outline_invented"] = {"applied": False, "error": str(_ol_e)[:200]}

        # 6.4b) 一致性自查 + 二次修复（2026-10-29）
        # 批量动作下发后总有一部分段落没拿到规则：locator 的段落集与 lint 的
        # 段落集不完全一致，或段落被后续结构步骤重写。这里对**产物**按同一
        # spec 再 lint/plan 一次，只补段落级格式动作（不动结构），再复检剩余。
        # 这正是「自查交付」：不等外部评分器发现，自己交出去前先查一遍。
        try:
            _tgt = target_for_validation
            _ir2 = analyze_document(_tgt)
            _lint2 = lint_format(_ir2, result.spec)
            _plan2 = plan_format_repairs(
                _ir2, result.spec, _lint2.violations,
                PlanningPolicy(max_risk="R2", preserve_content=True),
            )
            # 只补「格式属性」：对齐/制表位由公式段、题注、表格等专用步骤决定，
            # 二次修复重新判定会把它们覆盖（2026-10-29 公式编号段 jc 被刷成 center）。
            _SAFE_PROPS = {"font_size_pt", "font_east_asia", "font_ascii", "line",
                           "line_rule", "line_spacing", "spacing_before",
                           "spacing_after", "ind_first_line", "ind_first_line_chars",
                           "ind_left", "ind_right"}
            # 对齐（jc）默认不在二次修复范围内（正文的对齐由结构步骤决定）。
            # 但**前置区角色**没有任何结构步骤管对齐：模板要求题目/作者/单位居中，
            # 二次修复把它们重写成「只补字号字体」的动作，会把第一遍写好的
            # jc=center 一起丢掉（2026-11-08 用户实锤：期刊题目两端对齐、作者不居中）。
            _JC_SAFE_ROLES = {
                "article_title", "article_title_en", "authors", "authors_en",
                "affiliation", "affiliation_en", "abstract_title", "abstract_title_en",
                "toc_title", "keywords", "keywords_en", "abstract", "abstract_en",
            }
            # EZI-0053：**标题角色**的 jc 同样必须允许补写。「标题对齐由结构
            # 步骤决定」不成立：唯一写标题 jc 的通道就是首发的 set_paragraph_format
            # （同规则的字号/字体已在 _SAFE_PROPS 里随二次修复补齐）；heading_parity
            # 只镜像模板同文对，未配对的标题没有任何兜底写入点。实测（燕山示范文
            # L2）：6.48 封面还原把封面区边界外的示范标题段按文本指纹整段替换回
            # 源文档原样，首发写好的 jc=both 随 pPr 一起被抹；二次修复补了字号/
            # 行距/缩进、唯独不补 jc → 判定层读 start → FAIL。对齐与字体同源
            # （spec 规则 expected），允许补写只是与首发同契约。
            _JC_SAFE_ROLES |= {
                "heading1", "heading2", "heading3", "heading4",
                "refs_heading", "ack_heading", "appendix_heading", "bio_heading",
            }
            _para2 = []
            _body_scope2: dict = {}
            for _a in (_plan2.actions or []):
                if _a.action != "set_paragraph_format":
                    continue
                _sel = _a.params.get("selector") or {}
                if not _sel.get("semantic_role"):
                    continue
                _allowed = (_SAFE_PROPS | {"jc"}) if _sel.get("semantic_role") in _JC_SAFE_ROLES else _SAFE_PROPS
                _props = {k: v for k, v in (_a.params.get("properties") or {}).items()
                          if k in _allowed}
                if not _props:
                    continue
                _new_params = {"selector": _sel, "properties": _props}
                if str(_sel.get("semantic_role") or "") == "body":
                    # EZI-0050：正文补齐同样不得写到带大纲级别的段上（手写目录条目被
                    # FrontRoleAssigner 判成 body）。与首发 / 5.991 终局重放同一条过滤；
                    # 捕获就在这一步的产物上做，指纹与随后执行时逐段一致。
                    try:
                        _scope2 = _body_targets_without_outline(
                            Path(_tgt), _sel, heading_levels=_heading_levels)
                        _body_scope2 = {k: _scope2[k] for k in ("kept", "dropped_outline")}
                        if _scope2["kept"]:
                            _new_params["_resolved_paragraphs"] = _scope2["captured"]
                    except Exception as _bse:  # noqa: BLE001 — 过滤失败退回旧行为
                        _body_scope2 = {"error": str(_bse)[:200]}
                _para2.append(RepairAction(action=_a.action,
                                           params=_new_params,
                                           risk=_a.risk, reason=_a.reason, source=_a.source))
            _conf = {"checked": len(_lint2.violations), "planned": len(_para2)}
            if _body_scope2:
                _conf["body_scope"] = _body_scope2
            if _para2 and len(_para2) <= 40:
                _tx2 = TransactionExecutor(reg, work).execute_post_actions(
                    RepairPlan(plan_id=f"{(result.plan.plan_id if result.plan else 'plan')}#conformance",
                               actions=_para2),
                    src,
                    tx_id=result.transaction.transaction_id,
                    candidate=Path(_tgt),
                    backup=Path(result.transaction.backup_path),
                    journal_path=Path(result.transaction.journal_path),
                )
                result.transaction = _tx2
                target_for_validation = _tx2.output_path
                _lint3 = lint_format(analyze_document(target_for_validation), result.spec)
                _conf.update({"status": _tx2.status,
                              "remaining": len(_lint3.violations)})
            elif _para2:
                _conf["skipped"] = f"段落级动作 {len(_para2)} 条，超过单次补齐上限 40"
            result.artifacts["conformance"] = _conf
        except Exception as _cf:  # noqa: BLE001
            result.artifacts["conformance"] = {"error": str(_cf)[:200]}

        # 6.4c) 作者角标终局校验（2026-11-08）：
        # 现象：superscript_front_marks 早期把作者行 1) 2) 标成上标，随后又被
        # 前部一致性修复的段落重写抹掉（同一段落在两次写入之间被替换，
        # set_paragraph_format 内部拿到的快照已经是 0）」。
        # 处置：在渲染前最后再跑一次并**验证**——不是靠顺序侥幸，
        # 而是「上标存在」成为交付前的硬断言；幂等，已上标则不动。
        try:
            from earlyz.actions.new_actions import (
                superscript_front_marks as _sup2,
                verify_superscript_front_marks as _vsup2,
            )

            _roles2 = {"roles": ["authors", "affiliation", "authors_en", "affiliation_en"]}
            result.artifacts["front_marks_final"] = _sup2(target_for_validation, _roles2)
            result.artifacts["front_marks_ok"] = bool(_vsup2(target_for_validation, _roles2))
        except Exception as _sf:  # noqa: BLE001
            result.artifacts["front_marks_final"] = {"error": str(_sf)[:200]}

        # 6.5-pre) 渲染前结构门禁（2026-10-18）：结构/内容问题必须秒级暴露，
        # 不能等 Word 渲染完（真实大文档 7~8 分钟）才报封面污染这类硬错误。
        _pre_blocked = False
        try:
            from earlyz.validation.delivery import validate_delivery as _vd_pre

            _pre = _vd_pre(
                target_for_validation,
                source_path=src,
                spec=result.spec,
                renderer="none",
                authorized_content_actions=authorized_content_actions or None,
                capability_config=capability_config,
                cover_rebuilt=bool((result.artifacts.get("cover_auto") or {}).get("applied")),
            )
            if str(getattr(_pre, "level", "")).upper() == "FAILED":
                _pre_blocked = True
                result.validation = _pre
                result.artifacts["prerender_gate"] = {
                    "blocked": True,
                    "errors": list(getattr(_pre, "hard_errors", None) or [])[:6],
                }
                if progress_cb is not None:
                    try:
                        progress_cb({"type": "stage", "phase": "prerender_gate",
                                     "summary": "结构门禁未通过，已跳过渲染"})
                    except Exception:  # noqa: BLE001
                        pass
        except Exception as _e:  # noqa: BLE001
            result.artifacts["prerender_gate"] = {"blocked": False, "error": str(_e)[:200]}

        if progress_cb is not None and not _pre_blocked:
            try:
                progress_cb({"type": "stage", "phase": "render",
                             "summary": "渲染验证（Word/LibreOffice）…"})
            except Exception:  # noqa: BLE001
                pass
        # 6.59) 封面/扉页是两个**页面单元**（EZI-0042）。源稿靠封面高度自然溢出分页，
        #       产物把封面格式（字号/行距）改小后溢出消失 → 扉页并进封面页；扉页末尾
        #       那条 nextPage 分节符再强开一页 → 第 2 页纯空白（Word COM 实测：
        #       s1=[1..1] s2=[3..11]）。**必须放在封面还原（6.48）之后**：还原会把
        #       封面段落 pPr 从源文档整段回填，早期写进去的 pageBreakBefore 会被抹掉
        #       （实测：p8 的 pPr 变成只剩 `jc=center`）。只加页属性、不动节结构，
        #       也不删任何已有分页/空白页（源稿有意的偶数页留白保持原样）。
        try:
            from earlyz.actions.new_actions import (
                ensure_cover_page_unit_breaks as _cover_units,
            )

            result.artifacts["cover_page_units"] = _cover_units(
                target_for_validation,
                Path(spec_source) if spec_source else None,
            )
        except Exception as _cu:  # noqa: BLE001 — 分页失败不阻断格式交付
            result.artifacts["cover_page_units"] = {"error": str(_cu)[:200]}

        # 6.6) 收尾规范化（2026-09-28 用户判决「从源头解决，不留屎山」）：
        # 交付前统一按 schema 顺序重排 pPr/rPr/tblPr/trPr 子元素 + 修正 w:pPr
        # 位置 → 产物零违规。动作侧另有一道（new_actions._write_part，全仓唯一
        # 写盘函数），两层叠加：将来新增直写点也不会漏。顺序表真源 =
        # earlyz/tools/ooxml_order.py。
        from earlyz.tools.ooxml_order import canonicalize_docx

        # 动作会把模板的 ppr_extra_xml/rpr_extra_xml 原样写进产物，里面可能带
        # 「宋体;SimSun」这类脏字体名（Word 当不存在 → 回退默认字体）。
        try:
            from earlyz.pipeline.fonts_norm import normalize_fonts_in_place
            _font_post = normalize_fonts_in_place(target_for_validation)
            if _font_post.get("changed"):
                result.artifacts["fonts_normalized_post"] = _font_post
        except Exception:  # noqa: BLE001
            pass

        result.artifacts["canonicalized"] = canonicalize_docx(target_for_validation)

        result.validation = validate_delivery(
                target_for_validation,
                source_path=src,
                spec=result.spec,
                renderer="word_preferred" if word_render_callback else "none",
                word_render_callback=word_render_callback,
                authorized_content_actions=authorized_content_actions or None,
                capability_config=capability_config,
                cover_rebuilt=bool((result.artifacts.get("cover_auto") or {}).get("applied")),
            )
        _n_instr = int((result.artifacts.get("source_instructions") or {}).get("count") or 0)
        if _n_instr:
            result.validation.warnings.append(
                f"源文档里有 {_n_instr} 处模板说明文字（如「三号黑体居中，上下各空一行」）。"
                "按只改格式不动内容的原则已原样保留；诊断包里可见完整清单，确认后可清理。"
            )
        for _font_from in (result.artifacts.get("font_substitutions") or []):
            result.validation.warnings.append(
                f"本机没有安装「{_font_from}」，公文红头已用「黑体」近似：字号/位置不变，"
                "字形与装了该字体的机器略有差异。要完全一致请在本机安装该字体后重新排版。")
        if structural_failures:
            # 2026-10-14（复现事故 root）：失败按性质分流——
            #   · handler 异常（动作中途出错）→ 真损坏风险 → 硬错误 + FAILED；
            #   · verifier 异议（如「重扫仍有可转项」的口径之争）→ 动作改动
            #     已保留、内容/结构另有交付门普查 → 降为 warning + 证书明细，
            #     不再一票否决（一票否决曾把假失败放大为整篇无产物）。
            _soft, _hard = [], []
            for row in structural_failures:
                err = str(row.get("error") or "")
                (_soft if err.startswith("verifier ") else _hard).append(row)
            for row in _soft:
                result.validation.warnings.append(
                    f"{row.get('action', 'unknown action')} 未通过复检"
                    f"（更改已保留，已由交付门复核）：{row.get('error')}")
            if _soft:
                result.artifacts.setdefault("verifier_disagreements", []).extend(
                    {"action": r.get("action"), "error": r.get("error")}
                    for r in _soft)
            if _hard:
                result.validation.hard_errors.extend(
                    f"{row.get('action', 'unknown action')} failed: "
                    f"{row.get('error') or 'verification failed'}"
                    for row in _hard)
                result.validation.passed = False
                result.validation.level = "FAILED"

        # 模板示例残留清理失败 → 用户可见（2026-10-14 审计 D4：此前只躺在
        # artifacts 里，用户界面读的是 warnings/hard_errors）。
        _rc = result.artifacts.get("residual_cleanup") or {}
        if _rc.get("error"):
            result.validation.warnings.append(
                f"模板示例残留清理未执行（不阻断交付）：{_rc['error']}")

        # 公式超长自动缩字号 → **明确告诉用户**（用户 2026-09-28 拍板：「如果实在
        # 放不下只能缩小字体了，然后给用户说明情况」）。缩字号会动公式本体，属必须
        # 上报的动作：进 validation.warnings，用户在产品里能看到。
        _eq_info = structural_summary.get("normalize_equation_paragraphs") or {}
        _shrunk = _eq_info.get("shrink_details") or []
        if _shrunk:
            result.artifacts["equation_shrunk"] = _shrunk

        # 公式文本定型通报（B/C，2026-10-06）。两件事都必须让用户看见：
        # C：$…$ 文本转对象是内容级改写（用户拍板「公式是纯文本」缺陷的根治），
        #    转了几条、跳过几条都如实说；B：纯文本公式行只排了版没转对象，
        #    明确告知怎么转（E9：方案先给用户过目，确认后才替换）。
        _tt = structural_summary.get("toggle_tex_equations") or {}
        _tt_done = int(_tt.get("display_converted") or 0) + int(_tt.get("inline_converted") or 0)
        if _tt_done:
            result.validation.warnings.append(
                f"已把 {_tt_done} 条 $…$ / \\[…\\] 公式文本转换为可编辑公式对象"
                "（Word 里双击可编辑）")
        for _skip in (_tt.get("skipped") or []):
            result.validation.warnings.append(
                f"公式文本未转换：{_skip.get('text', '')} …（{_skip.get('skip_reason', '')}）")
        _plain = _eq_info.get("plain_seen") or 0
        if _plain:
            result.validation.warnings.append(
                f"识别到 {_plain} 条纯文本公式（内容是文字，还不是可编辑公式对象）。"
                "回复「公式转对象」，我会先给转换方案过目，你确认后再替换。")
        # 源文档没有页眉、模板有 → 已补空白页眉线，文字要用户补（2026-10-17 用户拍板）
        _hf = structural_summary.get("apply_template_header_footer") or {}
        if int(_hf.get("parts_created") or 0):
            result.validation.warnings.append(
                "模板有页眉（卷期/作者/刊名等信息），你的文档原本没有；"
                "已按模板补了空白页眉线（只含格式），页眉文字需要手工补，"
                "或者把内容告诉 EarlyZ 让 agent 帮你加。")
        for _row in _shrunk:
            _frm = (_row.get("from_half_pt") or 0) / 2
            _to = (_row.get("to_half_pt") or 0) / 2
            if _row.get("fits"):
                result.validation.warnings.append(
                    f"公式过长，已自动缩小字号：{_row.get('text', '')} … "
                    f"{_frm:g} 磅 → {_to:g} 磅（可用宽度 {_row.get('avail_twips')} twips，"
                    f"估算需要 {_row.get('need_twips')} twips）")
            else:
                result.validation.warnings.append(
                    f"公式过长，缩到下限仍可能放不下：{_row.get('text', '')} … "
                    f"{_frm:g} 磅（可用宽度 {_row.get('avail_twips')} twips，"
                    f"估算需要 {_row.get('need_twips')} twips）——建议缩短公式或改单栏")
    else:
        # 回滚必须带失败明细：此前只报「transaction rolled back」，
        # 用户和排查者都不知道是第几个动作、什么原因（2026-10-20）。
        _tx = result.transaction
        _detail = ""
        # 明细在 action_results 里：取第一个非 ok 的动作
        for _ar in (getattr(_tx, "action_results", None) or []):
            if str(_ar.get("status")) != "ok":
                _detail = (f"{_ar.get('action')}（#{_ar.get('index')}）："
                           f"{_ar.get('error') or _ar.get('skipped') or ''}")[:300]
                break
        for _attr in ("error", "message", "reason"):
            if _detail:
                break
            _v = getattr(_tx, _attr, None)
            if _v:
                _detail = str(_v)
                break
        if not _detail:
            for _attr in ("errors", "messages", "notes", "warnings"):
                _v = getattr(_tx, _attr, None)
                if _v:
                    _detail = "; ".join(str(x) for x in _v)[:400]
                    break
        result.validation = DeliveryReport(
            input_path=str(src), level="FAILED", passed=False,
            hard_errors=[f"transaction rolled back：{_detail}" if _detail else "transaction rolled back"],
        )

    # 5.99) 同文标题对齐（**必须最后跑，谁都别再改文档**）：模板对同一段文字自己怎么排
    #       （居中/两端/加粗/字号），就照抄。按级规则取的是众数，模板同级本来不一致时必然
    #       错一半（浙江政采 55 个标题段：公告区居中、格式区两端对齐）。实测放在 5.6 结构
    #       修复段之后仍会被后续步骤覆盖回旧值 → 挪到 artifacts 落地前最后一步。
    #       非阻断：失败只记录。
    if result.transaction.status == "applied" and result.transaction.output_path:
        # 5.98) 节几何对齐（2026-11-07；2026-11-16 EZI-0018 改按角色）：
        #       合同类模板 16–118 个节、节几何各不相同（封面 right=1300 / 正文
        #       right=1680），而 spec 只给一套页边距、被写进每一节；本科模板 8 节
        #       也是两套（封面/摘要 1417 组、目录/正文/参考文献/致谢/附录 1440 组）。
        #       节数相等时逐节按索引对位（判定器同口径）；节数不等时按角色
        #       （封面/声明/摘要/ABSTRACT/目录/正文/参考文献/致谢/附录）映射到模板
        #       同类节，禁止再按「模板首节/末节」二分。
        #       放这里：结构已定、分节已完，只搬 pgSz/pgMar 属性，不动分节结构。
        try:
            from earlyz.actions.section_geometry import align_section_geometry as _align_geo
            result.artifacts["section_geometry"] = _align_geo(
                Path(result.transaction.output_path), Path(spec_source))
        except Exception as e:  # noqa: BLE001
            result.artifacts["section_geometry"] = {"errors": [str(e)]}
        # 5.985) 按内容角色建节（EZI-0034，**结构定稿前的唯一一次**）：本步之前的
        #       `build_section_skeleton` / `apply_section_columns` /
        #       `ensure_toc_section_break` 各自按自己的判据插分节符（CJC 实测多出一个
        #       「Background」节、并把 continuous 改成 nextPage → 多空白页），所以结构
        #       到这里才定稿。**只增不删**（2026-11-17 回工）：只补缺失的切点、逐节复制
        #       模板该节的 cols/pgMar/pgSz/w:type；已有分节边界一条都不删，计划节数不多于
        #       现有节数时整份不动（删边界会让节内文本区间漂移，判定层按角色配对随之全变
        #       —— 北航/开封两项 PASS→FAIL 就是这个成因）。
        #       必须在 fit_tables_to_page（版心宽由节边距决定）之前，也必须在
        #       heading_parity（谁都别再改文档）之前。非阻断：失败只记录。
        try:
            from earlyz.actions.section_geometry import apply_section_plan as _apply_plan
            result.artifacts["section_plan_final"] = _apply_plan(
                Path(result.transaction.output_path), Path(spec_source))
        except Exception as e:  # noqa: BLE001
            result.artifacts["section_plan_final"] = {"errors": [str(e)]}
        # 5.99) 表格宽度收进版心（2026-11-07，必须在节几何之后——版心宽由节边距决定）：模板/源稿的表常按 A4 全宽（9060–9746）
        #       排版，目标版心只有 8306–9070 → 表格压到右边距外。只收显式宽度且超宽的
        #       表，按比例缩放 gridCol/tcW；type=auto 不动。
        try:
            from earlyz.actions.table_fit import fit_tables_to_page as _fit_tbl
            result.artifacts["table_fit"] = _fit_tbl(Path(result.transaction.output_path))
        except Exception as e:  # noqa: BLE001
            result.artifacts["table_fit"] = {"errors": [str(e)]}
        try:
            from earlyz.actions.heading_parity import align_headings_to_template as _align_h
            # EZI-0062：spec 规则的标题 jc 下发（判定层成对不足时的兜底期望）。
            # 只取通用规则 ``rule:tpl.headingN``——同名标题规则（.title.）作用域是
            # 单条标题，不做级别级兜底。
            _spec_jc: dict[int, str] = {}
            try:
                for _r in (result.spec.rules if result.spec else []):
                    _rid = str(getattr(_r, "rule_id", "") or "")
                    if _rid.startswith("rule:tpl.heading") and ".title." not in _rid:
                        try:
                            _lvl = int(_rid.rsplit("heading", 1)[1].split(".", 1)[0])
                        except ValueError:
                            continue
                        _v = (getattr(_r, "expected", None) or {}).get("jc")
                        if _v:
                            _spec_jc[_lvl] = str(_v)
            except Exception:  # noqa: BLE001
                _spec_jc = {}
            result.artifacts["heading_parity"] = _align_h(
                Path(result.transaction.output_path), Path(spec_source),
                spec_heading_jc=_spec_jc or None)
        except Exception as e:  # noqa: BLE001
            result.artifacts["heading_parity"] = {"errors": [str(e)]}
        # 5.991) 正文规则终局落地（EZI-0047）：5.99 的同文对齐有个纯文本兜底——产物段
        #        没有级别时按「同一段文字」抄模板段的字号。模板里的示例正文段
        #        （人大本科「[此处键入正文]」样例 12pt）会把**正文角色段落**的字号从
        #        正文规则的 10.5pt 改回样例的 12pt，正文主字号众数因此被顶回 12.0，
        #        判定层拿模板实测 10.5/20 去比必然 FAIL（正文行距的 300/auto 同理会
        #        被后置结构步骤改走）。正文段怎么排的唯一来源是 `rule:tpl.body`，所以
        #        同文对齐之后按正文规则把正文角色段落再写一次（没被改过的段落幂等空跑）。
        #        只重放**正文**这一条动作：标题/题注/参考文献的同文对齐结果不受影响。
        if (result.plan is not None and result.spec is not None
                and (result.artifacts.get("body_size_aligned") or {}).get("applied")):
            try:
                _body_action = next(
                    (a for a in result.plan.actions
                     if getattr(a, "action", "") == "set_paragraph_format"
                     and str(((getattr(a, "params", None) or {}).get("selector") or {})
                             .get("semantic_role") or "") == "body"),
                    None,
                )
                if _body_action is not None:
                    from earlyz.actions.new_actions import set_paragraph_format as _spf
                    # 计划里的动作带 `_resolved_paragraphs`（初次执行时按指纹捕获的段落集）；
                    # 结构步骤改过文档后这份捕获只剩 4 段还认得出。复位必须**按角色重新
                    # 解析**，否则大部分正文段根本没被写（人大实测 4/15）。
                    _reassert_params = {
                        k: v for k, v in (_body_action.params or {}).items()
                        if not str(k).startswith("_")
                    }
                    # EZI-0050：角色再解析会把带 outlineLvl 的段圈进来（手写目录条目被
                    # FrontRoleAssigner 判成 body），正文规则于是把 5.99 刚对齐好的标题值
                    # 刷回正文值——留出集 9 项 PASS→FAIL 的唯一成因。重放集先剔掉一切带
                    # 大纲级别的段（禁按字号白名单之类的补丁绕过角色过滤），再按指纹捕获
                    # 给动作层，与计划动作同一套定位机制。
                    _roster = _body_targets_without_outline(
                        Path(result.transaction.output_path),
                        _reassert_params.get("selector") or {})
                    if _roster["kept"] == 0:
                        result.artifacts["body_format_reassert"] = {
                            "ok": True, "changed": False,
                            "skipped": "重放集为空（按角色解析出的段落全带大纲级别）",
                            "kept": _roster["kept"],
                            "dropped_outline": _roster["dropped_outline"],
                        }
                    else:
                        _reassert_params["_resolved_paragraphs"] = _roster["captured"]
                        _reassert_out = _spf(
                            Path(result.transaction.output_path), _reassert_params)
                        if isinstance(_reassert_out, dict):
                            _reassert_out.setdefault("kept", _roster["kept"])
                            _reassert_out.setdefault("dropped_outline",
                                                     _roster["dropped_outline"])
                        result.artifacts["body_format_reassert"] = _reassert_out
            except Exception as e:  # noqa: BLE001 — 复位失败不阻断编译
                result.artifacts["body_format_reassert"] = {"ok": False, "error": str(e)[:200]}

    result.artifacts.update({
        "output_path": result.transaction.output_path,
        "journal_path": result.transaction.journal_path,
        "violations": len(result.lint.violations) if result.lint else 0,
        "plan_actions": len(result.plan.actions) if result.plan else 0,
    })
    # 示例文字大面积保留的告警（护栏触发时）
    try:
        _rc = result.artifacts.get("residual_cleanup") or {}
        if _rc.get("skipped") == "mass_deletion_guard" and result.validation is not None:
            result.validation.warnings.append(
                f"文档疑似规范/说明文件：{_rc.get('matched')} 段模板示例文字未自动删除"
                f"（占全文 {_rc.get('matched')}/{_rc.get('paragraphs')} 段），如需清理请确认"
            )
    except Exception:  # noqa: BLE001
        pass
    # 规范自检告警并入交付报告（warning 级，不阻断）
    try:
        for _w in (result.artifacts.get("spec_health") or []):
            if result.validation is not None and _w not in result.validation.warnings:
                result.validation.warnings.append(_w)
    except Exception:  # noqa: BLE001
        pass
    return result


def compile_document_idempotent_check(
    result: CompileResult, spec_source_type: str, spec_source: str
) -> bool:
    """幂等验收：对产物再编译一次，有效 action 数必须为 0。"""
    if not result.transaction or not result.transaction.output_path:
        return False
    out = result.transaction.output_path
    second = compile_document(
        out,
        spec_source_type=spec_source_type,
        spec_source=spec_source,
        domain=result.spec.domain if result.spec else "general",
        workdir=str(Path(out).parent / "idempotent"),
        max_risk="R2",
    )
    return len(second.plan.actions) == 0

def _cover_range_has_table(docx_path, rng):
    """封面区段里是否含表格。

    2026-10-25 实测（真实语料 199 份）：「开题报告 / 中期检查报告 / 材料集」这类
    文档的封面本身就是表单（表格），封面重建按段落区段整段替换会把表格连同内容
    一起删掉——交大中期检查 5 表→1 表、哈医大 6 表→5 表并触发 text hash changed。
    表单式封面一律保留原样，不做自动重建。
    """
    try:
        import zipfile
        from lxml import etree
        W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        root = etree.fromstring(zipfile.ZipFile(docx_path).read("word/document.xml"))
        body = root.find(f"{W}body")
        if body is None:
            return False
        start, end = int(rng[0]), int(rng[1])
        p_idx = -1
        for child in body:
            if child.tag == f"{W}p":
                p_idx += 1
                if p_idx > end:
                    break
            elif child.tag == f"{W}tbl":
                # 表格出现在封面区段内（含紧随最后一个封面段落之后）
                if p_idx >= start - 1 and p_idx <= end:
                    return True
        return False
    except Exception:
        return False


def _front_cover_range(docx_path):
    """计算源文档「封面区」段落区间 [start, end]（body 直属子元素下标，含端点）。

    以**分节结构**为准：封面 = 文档第一节，区间到第一个承载 sectPr 的段落为止。
    这样不会跨到声明/摘要/目录（按标题文字切分会把摘要一起圈进来，实测 [0,75]）。
    取不到（没有分节符）就返回 None，调用方退回 auto，不冒险删除。
    """
    try:
        import zipfile as _zf
        from lxml import etree as _et
        WNS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        with _zf.ZipFile(docx_path) as z:
            root = _et.fromstring(z.read("word/document.xml"))
        body = root.find(f"{WNS}body")
        if body is None:
            return None
        kids = list(body)
        end = None
        for i, el in enumerate(kids):
            if _et.QName(el).localname != "p":
                continue
            ppr = el.find(f"{WNS}pPr")
            if ppr is not None and ppr.find(f"{WNS}sectPr") is not None:
                end = i
                break
        if end is None:
            return None
        # 末尾的空段留在区间内（它们承载分节符与空行），前面的空段照常保留
        return [0, end]
    except Exception:  # noqa: BLE001
        return None
