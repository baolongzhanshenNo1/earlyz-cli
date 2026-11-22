"""语料化版式打分的聚合层（EZI-0007）。

打分本身在 ``earlyz.validation.layout_fidelity``（EZI-0005）：一次只判「一对」产物×模板，
返回 23 项 PASS/FAIL/N-A。这里补的是**统计那一层**：把 N 组判定行聚合成

* 逐项通过率（每个项名在多少组里出现、PASS/FAIL/N-A 各几条、通过率多少）；
* 失败聚类（同一项在哪些组失败、典型「期望值 vs 实测值」是什么）；
* Markdown 两张表 + 逐组结果表。

设计约束（决定了它长这样）：

* **纯函数、不碰 I/O 之外的东西**：不 import 排版引擎、不开 Word、不读 docx。
  输入是已判定好的 items（``LayoutItem`` 或等价 dict），所以单测能喂假数据把
  聚合口径钉死，而不用先跑一遍 4 分钟编译。
* **N-A 不进分母**：与 ``layout_fidelity`` 的分数口径一致，分母为 0 时算 0.0，
  不拿 0/0 当满分。
* **排序必须确定**：表是为了横比，次序每次都不同就没法看趋势。
  逐项表按「通过率升序 → 失败数降序 → 项名」；失败聚类按「失败组数降序 → 项名」。

CLI 壳在 ``scripts/corpus_layout_score.py``；本模块只放可 import 的纯逻辑。

元基准（EZI-0031）——**尺子自己也得有验收标准**：

* **分母健康度**（``metric_health``）：每项的 N-A 占比。超过 ``NA_RATIO_LIMIT``
  的项打 ``【疑似死指标】``——一半以上的组判不了，这项的通过率就没有代表性；
  在**全部组上恒为 N-A** 的项再打 ``【恒为 N-A】`` 并在报告顶部单独告警，
  不许静默占着「覆盖 27 项」的名分。
* **口径变更只读对照**（``compare_payloads``）：喂两份既有结果 JSON（旧口径跑出来的
  和现行口径跑出来的），输出逐项 verdict 翻转清单（项名 / 组 / 旧 / 新）+ 分母变动。
  只读：不重判、不编译、不开 Word。入口是本模块的 ``__main__``（``scripts/`` 下那个
  CLI 不在 EZI-0031 的 write_scope 内，所以对照入口放在这里而不是加 CLI 参数）::

      env -u PYTHONPATH python -m earlyz.validation.corpus_scoring \\
          --baseline-json old.json --current-json new.json --md compare.md
"""
from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

from .layout_fidelity import FAIL, NA, PASS

__all__ = [
    "CONSTANT_NA_MARK",
    "DEAD_METRIC_MARK",
    "NA_RATIO_LIMIT",
    "CaseResult",
    "ManifestCase",
    "ManifestError",
    "build_payload",
    "case_result",
    "compare_payloads",
    "error_summary",
    "load_manifest",
    "load_result_payload",
    "main",
    "metric_health",
    "render_comparison_markdown",
    "render_markdown",
    "step_5_95_headline",
    "summarize",
    "verdict_map",
]

#: manifest 里每组允许出现的键。多一个就报错——拼错的键（``expect_min_scor``）
#: 静默忽略等于让验收断言失效，这类错必须当场撞出来。
CASE_KEYS = ("name", "template", "doc", "expect_min_score")

#: 失败聚类里每组最多给几个「期望 vs 实测」样本。
MAX_EXAMPLES = 3

#: 「跑不起来」的错误摘要在报告里最多占多少字符——要能一眼看出错在哪，
#: 又不能把表格撑歪；完整 traceback 留在 JSON 的 ``error`` 字段里。
MAX_ERROR_CHARS = 120

#: N-A 占比上限（EZI-0031 A1）：某项在超过这个比例的组里都判不了，它的通过率
#: 就不代表产品，只代表「样本没覆盖到」。放宽到 50% 是刻意的——一半样本还能判
#: 就还有参考价值；低于一半就得打标，没人能再说「覆盖 27 项」。
NA_RATIO_LIMIT = 0.5

#: 报告里的两个标记。它们是报警的**字面锚点**：报告、issue 证据、评审都按这两个
#: 字符串找，改文案要连测试一起改（``tests/test_corpus_scoring.py``）。
DEAD_METRIC_MARK = "【疑似死指标】"
CONSTANT_NA_MARK = "【恒为 N-A】"

_VERDICT_ALIASES = {PASS: PASS, FAIL: FAIL, NA: NA, "NA": NA, "N/A": NA}


class ManifestError(ValueError):
    """manifest 结构不合法（缺键 / 键名拼错 / 文件不存在 / 重名）。"""


# ────────────────────────────── manifest ──────────────────────────────


@dataclass(frozen=True)
class ManifestCase:
    """一组用例：拿 ``template`` 当模板去排 ``doc``，再按 ``template`` 打分。"""

    name: str
    template: Path
    doc: Path
    template_rel: str = ""
    doc_rel: str = ""
    expect_min_score: float | None = None

    def as_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "template": self.template_rel or str(self.template),
            "doc": self.doc_rel or str(self.doc),
            "expect_min_score": self.expect_min_score,
        }

    @property
    def safe_name(self) -> str:
        """用例名 → 可当目录名用（编译工作目录按组名分子目录）。"""
        bad = '<>:"/\\|?*'
        out = "".join("_" if (ch in bad or ord(ch) < 32) else ch for ch in self.name).strip()
        return (out or "case")[:60]


def _require_str(raw: Mapping[str, Any], key: str, where: str) -> str:
    v = raw.get(key)
    if not isinstance(v, str) or not v.strip():
        raise ManifestError(f"{where}: `{key}` 必须是非空字符串，实际为 {v!r}")
    return v.strip()


def load_manifest(path: str | Path) -> list[ManifestCase]:
    """读 manifest 并解析成绝对路径的用例列表。

    约定（写进 ``docs/corpus/README.md``）：``template`` / ``doc`` 的相对路径
    **以 manifest 所在目录为基准**——这样清单跟着语料走，复制目录不会失效。
    """
    p = Path(path)
    if not p.is_file():
        raise ManifestError(f"manifest 不存在：{p}")
    try:
        payload = json.loads(p.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise ManifestError(f"manifest 不是合法 JSON：{p}（{exc}）") from exc
    if not isinstance(payload, dict) or "cases" not in payload:
        raise ManifestError(f"manifest 顶层必须是对象且含 `cases`：{p}")
    raw_cases = payload["cases"]
    if not isinstance(raw_cases, list) or not raw_cases:
        raise ManifestError(f"`cases` 必须是非空数组：{p}")

    base = p.resolve().parent
    out: list[ManifestCase] = []
    seen: set[str] = set()
    for idx, raw in enumerate(raw_cases):
        where = f"cases[{idx}]"
        if not isinstance(raw, dict):
            raise ManifestError(f"{where}: 必须是对象，实际为 {type(raw).__name__}")
        unexpected = [k for k in raw if k not in CASE_KEYS]
        if unexpected:
            raise ManifestError(
                f"{where}: 出现未定义键 {unexpected}；允许的键只有 {list(CASE_KEYS)}"
            )
        name = _require_str(raw, "name", where)
        if name in seen:
            raise ManifestError(f"{where}: 用例名重复 `{name}`（名字是汇总表的行标识，必须唯一）")
        seen.add(name)
        rel_tpl = _require_str(raw, "template", where)
        rel_doc = _require_str(raw, "doc", where)
        tpl = (base / rel_tpl).resolve()
        doc = (base / rel_doc).resolve()
        for label, f in (("template", tpl), ("doc", doc)):
            if not f.is_file():
                raise ManifestError(f"{where} `{name}`: {label} 不存在：{f}")
        exp = raw.get("expect_min_score")
        if exp is not None:
            if isinstance(exp, bool) or not isinstance(exp, (int, float)):
                raise ManifestError(f"{where} `{name}`: expect_min_score 必须是数字，实际为 {exp!r}")
            exp = float(exp)
            if not 0.0 <= exp <= 1.0:
                raise ManifestError(f"{where} `{name}`: expect_min_score 必须在 [0,1]，实际为 {exp}")
        out.append(ManifestCase(name=name, template=tpl, doc=doc,
                                template_rel=rel_tpl, doc_rel=rel_doc,
                                expect_min_score=exp))
    return out


# ────────────────────────────── 单组结果 ──────────────────────────────


def _item_dict(item: Any) -> dict[str, Any]:
    """把 ``LayoutItem`` / dict 归一成 dict，并校验 verdict 合法。"""
    if isinstance(item, Mapping):
        d = dict(item)
    else:
        d = {
            "name": getattr(item, "name", None),
            "expected": getattr(item, "expected", None),
            "actual": getattr(item, "actual", None),
            "verdict": getattr(item, "verdict", None),
            "category": getattr(item, "category", ""),
        }
        note = getattr(item, "note", "")
        if note:
            d["note"] = note
    name = d.get("name")
    if not isinstance(name, str) or not name:
        raise ValueError(f"item 缺 name：{d!r}")
    raw = d.get("verdict")
    verdict = _VERDICT_ALIASES.get(str(raw).strip().upper() if raw is not None else None)
    if verdict is None:
        raise ValueError(f"item `{name}` 的 verdict 非法：{raw!r}（只认 PASS/FAIL/N-A）")
    d["verdict"] = verdict
    d.setdefault("category", "")
    d.setdefault("expected", None)
    d.setdefault("actual", None)
    return d


@dataclass
class CaseResult:
    """一组「编译产物 vs 模板」的判定结果。

    ``error`` 非空表示这组没跑出判定（编译炸了 / 文件坏了）——此时
    ``exit_code=2``，不把「跑不起来」伪装成「0 项通过」。
    """

    name: str
    template: str
    doc: str
    items: list[dict[str, Any]] = field(default_factory=list)
    product: str = ""
    error: str = ""
    expect_min_score: float | None = None
    #: 单对 wall time（编译 + 打分），A4 要求逐对记录——5 分钟一对的事实得能被统计
    duration_s: float | None = None
    #: 其中编译占的部分（打分本身约 1s，用来确认瓶颈在编译不在判定）
    compile_seconds: float | None = None
    #: 这一对编译时的渲染口径（EZI-0036）：``word`` / ``lo`` / ``none``。
    #: 空串 = 老结果或没走到编译，报告里必须显示「未记录」，不许默认成有渲染器。
    renderer: str = ""
    #: 5.95 跨页续表这一步的真实结局（EZI-0036）：从 ``result.artifacts`` 读，
    #: ``ran`` / ``skipped:未声明table_continued_label`` / ``skipped:无渲染器``。
    step_5_95: str = ""

    def __post_init__(self) -> None:
        self.items = [_item_dict(i) for i in self.items]

    # ── 计数 ────────────────────────────────────────────────
    @property
    def passed(self) -> int:
        return sum(1 for i in self.items if i["verdict"] == PASS)

    @property
    def failed(self) -> int:
        return sum(1 for i in self.items if i["verdict"] == FAIL)

    @property
    def not_applicable(self) -> int:
        return sum(1 for i in self.items if i["verdict"] == NA)

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
        if self.error:
            return 2
        return 1 if self.failed else 0

    @property
    def meets_expectation(self) -> bool:
        """没设 ``expect_min_score`` 时，默认「没有 FAIL」即达标。"""
        if self.error:
            return False
        if self.expect_min_score is None:
            return self.failed == 0
        return self.score >= self.expect_min_score - 1e-9

    def as_dict(self) -> dict[str, Any]:
        d = {
            "name": self.name,
            "template": self.template,
            "doc": self.doc,
            "product": self.product,
            "passed": self.passed,
            "failed": self.failed,
            "not_applicable": self.not_applicable,
            "total": self.total,
            "judged": self.judged,
            "score": self.score,
            "exit_code": self.exit_code,
            "expect_min_score": self.expect_min_score,
            "meets_expectation": self.meets_expectation,
            "items": [dict(i) for i in self.items],
        }
        if self.duration_s is not None:
            d["duration_s"] = round(self.duration_s, 2)
        if self.compile_seconds is not None:
            d["compile_seconds"] = round(self.compile_seconds, 2)
        if self.error:
            d["error"] = self.error
        # EZI-0036：只在真有值时才写键——老结果 JSON 里没有这两个键，
        # 新结果缺值（编译没走到）也不该伪装成 ``none`` / ``ran``。
        if self.renderer:
            d["renderer"] = self.renderer
        if self.step_5_95:
            d["step_5_95"] = self.step_5_95
        return d


def case_result(case: ManifestCase, report: Any, *, product: str = "",
                compile_seconds: float | None = None,
                duration_s: float | None = None,
                renderer: str = "", step_5_95: str = "") -> CaseResult:
    """把 ``LayoutReport``（或任意有 ``items`` 的对象）包成一条汇总行。"""
    items = list(getattr(report, "items", report) or [])
    return CaseResult(
        name=case.name,
        template=case.template_rel or str(case.template),
        doc=case.doc_rel or str(case.doc),
        items=items,
        product=product,
        expect_min_score=case.expect_min_score,
        duration_s=duration_s if duration_s is not None else compile_seconds,
        compile_seconds=compile_seconds,
        renderer=renderer,
        step_5_95=step_5_95,
    )


# ────────────────────────────── 聚合 ──────────────────────────────


def _rate(passed: int, judged: int) -> float:
    return _ratio(passed, judged)


def _ratio(num: int, den: int) -> float:
    """占比，分母 0 时给 0.0（不拿 0/0 当满分，也不当 100%）。"""
    return round(num / den, 4) if den else 0.0


def metric_health(items: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    """分母健康度（EZI-0031 A1/A2）：每项的 N-A 占比 + 死指标/恒 N-A 标记。

    为什么要单列这一层：分数 = ``passed / (total - N-A)``，而 N-A 是**判定器自己说
    「这项我判不了」**。分母缩小时总分自动变好看（可判项 228→187，总分 88.2%→96.3%），
    所以必须有人检查分母是不是虚的——这个函数就是那个检查。

    * ``na_ratio``：该项 N-A 组数 / 该项出现组数；
    * ``suspect_dead``：``na_ratio > NA_RATIO_LIMIT``，报告里打 ``【疑似死指标】``；
    * ``constant_na``：``judged == 0``（出现了但一组都判不了），打 ``【恒为 N-A】``
      并由报告顶部单独告警——这类项通过率恒为 0.0，却常被当成「覆盖到了」。

    入参是 ``summarize()`` 的 ``items``（或任何含
    ``name/cases/judged/not_applicable`` 的行），所以老结果 JSON 也能现算健康度。
    """
    rows: list[dict[str, Any]] = []
    total = na_total = judged_total = 0
    for it in items:
        cases = int(it.get("cases") or 0)
        na = int(it.get("not_applicable") or 0)
        raw_judged = it.get("judged")
        judged = int(raw_judged) if raw_judged is not None else max(cases - na, 0)
        total += cases
        na_total += na
        judged_total += judged
        ratio = _ratio(na, cases)
        rows.append({
            "name": str(it.get("name") or ""),
            "category": str(it.get("category") or ""),
            "cases": cases,
            "not_applicable": na,
            "judged": judged,
            "na_ratio": ratio,
            "suspect_dead": cases > 0 and ratio > NA_RATIO_LIMIT,
            "constant_na": cases > 0 and judged == 0,
        })
    suspect = sorted((r for r in rows if r["suspect_dead"]),
                     key=lambda r: (-r["na_ratio"], r["name"]))
    constant = sorted((r for r in rows if r["constant_na"]), key=lambda r: r["name"])
    return {
        "total": total,
        "not_applicable": na_total,
        "judged": judged_total,
        "na_ratio": _ratio(na_total, total),
        "limit": NA_RATIO_LIMIT,
        "suspect_dead_metrics": suspect,
        "constant_na_items": constant,
    }


def summarize(results: Sequence[CaseResult]) -> dict[str, Any]:
    """逐项通过率 + 失败聚类 + 整体分数 + 分母健康度。只看 verdict，不重新判版式。"""
    results = list(results)
    passed = sum(r.passed for r in results)
    failed = sum(r.failed for r in results)
    na = sum(r.not_applicable for r in results)
    judged = sum(r.judged for r in results)
    total = sum(r.total for r in results)
    scored = [r.score for r in results if r.judged]
    macro = round(sum(scored) / len(scored), 4) if scored else 0.0

    # 项名 → 统计。同名项在不同组可能一个 FAIL 一个 PASS（不同模板期望值不同），
    # 所以逐项通过率按「组 × 项」为样本，不按组。
    stats: dict[str, dict[str, Any]] = {}
    for r in results:
        for it in r.items:
            s = stats.setdefault(it["name"], {
                "name": it["name"],
                "category": it.get("category") or "",
                "cases": 0, "passed": 0, "failed": 0, "na": 0,
                "failed_cases": [], "_examples": [],
            })
            s["cases"] += 1
            if it["verdict"] == PASS:
                s["passed"] += 1
            elif it["verdict"] == FAIL:
                s["failed"] += 1
                s["failed_cases"].append(r.name)
                s["_examples"].append({
                    "case": r.name,
                    "expected": it.get("expected"),
                    "actual": it.get("actual"),
                })
            else:
                s["na"] += 1

    items: list[dict[str, Any]] = []
    for s in stats.values():
        judged_i = s["cases"] - s["na"]
        na_ratio = _ratio(s["na"], s["cases"])
        items.append({
            "name": s["name"],
            "category": s["category"],
            "cases": s["cases"],
            "judged": judged_i,
            "passed": s["passed"],
            "failed": s["failed"],
            "not_applicable": s["na"],
            "na_ratio": na_ratio,
            "suspect_dead": s["cases"] > 0 and na_ratio > NA_RATIO_LIMIT,
            "constant_na": s["cases"] > 0 and judged_i == 0,
            "pass_rate": _rate(s["passed"], judged_i),
            "failed_cases": list(s["failed_cases"]),
        })
    items.sort(key=lambda r: (r["pass_rate"], -r["failed"], r["name"]))

    failures: list[dict[str, Any]] = []
    for s in stats.values():
        if not s["failed"]:
            continue
        seen: set[tuple[str, str]] = set()
        examples: list[dict[str, Any]] = []
        for ex in sorted(s["_examples"], key=lambda e: e["case"]):
            key = (_show(ex["expected"]), _show(ex["actual"]))
            if key in seen:
                continue
            seen.add(key)
            examples.append(ex)
            if len(examples) >= MAX_EXAMPLES:
                break
        failures.append({
            "name": s["name"],
            "category": s["category"],
            "failed_cases": s["failed"],
            "cases": list(s["failed_cases"]),
            "examples": examples,
        })
    failures.sort(key=lambda f: (-f["failed_cases"], f["name"]))

    durations = [r.duration_s for r in results if r.duration_s is not None]
    duration = {
        "n": len(durations),
        "total": round(sum(durations), 1),
        "mean": round(sum(durations) / len(durations), 1) if durations else 0.0,
        "max": round(max(durations), 1) if durations else 0.0,
        "min": round(min(durations), 1) if durations else 0.0,
    }

    # 分母健康度：与 render_markdown 走同一个函数，避免「JSON 的占比」和「报告里的
    # 占比」两套算法（这类漂移最后一定会有人拿着两份数对不上）。
    health = metric_health(items)

    return {
        "cases": len(results),
        "passed": passed,
        "failed": failed,
        "not_applicable": na,
        "total": total,
        "judged": judged,
        "score": _rate(passed, judged),
        "macro_score": macro,
        "items": items,
        "failures": failures,
        "errors": [{"name": r.name, "error": r.error} for r in results if r.error],
        "unmet_expectations": [r.name for r in results if not r.meets_expectation],
        "duration_s": duration,
        "na_ratio": health["na_ratio"],
        "na_ratio_limit": health["limit"],
        "suspect_dead_metrics": health["suspect_dead_metrics"],
        "constant_na_items": health["constant_na_items"],
    }


def overall_exit_code(results: Sequence[CaseResult]) -> int:
    """0=全部达标；1=有组未达标（含 FAIL）；2=有组跑不起来。"""
    results = list(results)
    if any(r.error for r in results):
        return 2
    return 0 if all(r.meets_expectation for r in results) else 1


def build_payload(manifest: str | Path, results: Sequence[CaseResult], *,
                  generated_at: str = "", workdir: str = "",
                  skip_render: bool = True, project_pairs: int | None = None,
                  renderer: str | None = None,
                  step_5_95: str | None = None) -> dict[str, Any]:
    """CLI 写盘的 JSON 结构：逐组结果 + summary（+ 可选的全量耗时推算）。

    ``renderer`` / ``step_5_95``（EZI-0036）是**本次跑批的口径标签**，只在调用方
    明确给出时才写进载荷——老调用点（以及历史结果 JSON）保持原样，报告层靠
    ``payload.get(...)`` 判「未记录」，不靠默认值假装有渲染器。
    """
    results = list(results)
    summary = summarize(results)
    if project_pairs and summary["duration_s"]["mean"]:
        mean = float(summary["duration_s"]["mean"])
        summary["projection"] = {
            "pairs": int(project_pairs),
            "mean_s_per_pair": mean,
            "hours": round(mean * int(project_pairs) / 3600.0, 1),
        }
    payload: dict[str, Any] = {
        "manifest": str(manifest),
        "generated_at": generated_at,
        "workdir": workdir,
        "env": {"EARLYZ_SKIP_WORD_RENDER": "1" if skip_render else "unset"},
        "exit_code": overall_exit_code(results),
        "cases": [r.as_dict() for r in results],
        "summary": summary,
    }
    if renderer is not None:
        payload["renderer"] = renderer
    if step_5_95 is not None:
        payload["step_5_95"] = step_5_95
    return payload


def step_5_95_headline(results: Sequence[CaseResult]) -> str:
    """把逐对的 5.95 结局压成一个报告头标签；混批则写 ``mixed（...计数...）``。

    为什么要压：报告头只留一行，但一批里可能有的组跑了 5.95、有的没声明
    ``table_continued_label``——把差别吞掉就等于伪造口径。逐组真值仍在
    ``cases[].step_5_95``。
    """
    states = [r.step_5_95 for r in results if r.step_5_95]
    if not states:
        return ""
    if len(set(states)) == 1:
        return states[0]
    counts = Counter(states)
    detail = "、".join(f"{k}={v}" for k, v in counts.most_common())
    return f"mixed（{detail}）"


# ────────────────────────────── Markdown ──────────────────────────────


def error_summary(err: str, *, limit: int = MAX_ERROR_CHARS) -> str:
    """错误文本 → 单行摘要（异常类型 + 首行），超长截断加 ``…``。

    为什么需要它：``CaseResult.error`` 常是一整段 traceback，直接塞进 Markdown 表格
    会把表撑歪；但只给 ``0/0`` 又等于静默失败。折中是「异常类型 + 首行」，最多
    ``limit`` 字符——够定位是哪一类错、错在什么消息上，细节仍看 JSON 的 ``error``。

    CLI 写 error 时按 ``f"{type(exc).__name__}: {exc}"`` 拼，所以首行天然以异常类型
    开头；这里再做一层归一（去空行 / 压空白 / 截断），让手写的 ``CaseResult`` 也不跑偏。
    """
    text = str(err or "")
    first = ""
    for line in text.splitlines():  # splitlines 已按行切分（含 CRLF），无需手工归一
        collapsed = " ".join(line.split())
        if collapsed:
            first = collapsed
            break
    if limit <= 0:
        return ""
    if len(first) <= limit:
        return first
    return first[: limit - 1].rstrip() + "…"


def _error_rows(cases: Sequence[Mapping[str, Any]]) -> list[tuple[str, str]]:
    """``[(组名, 错误摘要)]``，只保留跑不起来的组；空 error 的组不算。"""
    return [(str(c.get("name", "")), error_summary(str(c.get("error") or "")))
            for c in cases if c.get("error")]


def _show(v: Any) -> str:
    """值 → 单元格文本（与 layout_fidelity 的展示口径一致：dict 展开成 k=v）。"""
    if v is None:
        return "None"
    if isinstance(v, Mapping):
        return "{" + ", ".join(f"{k}={v[k]}" for k in v) + "}"
    return str(v)


def _row_na_ratio(it: Mapping[str, Any]) -> float:
    """一行的 N-A 占比。老结果 JSON 没有 ``na_ratio`` 字段时按 na/组数现算。"""
    raw = it.get("na_ratio")
    if raw is not None:
        return float(raw)
    return _ratio(int(it.get("not_applicable") or 0), int(it.get("cases") or 0))


def _cell(v: Any) -> str:
    """表格单元格：转义竖线、压掉换行（Markdown 表格一行就是一行）。"""
    text = _show(v).replace("|", "\\|")
    return text.replace("\r", " ").replace("\n", "<br>")


def _pct(rate: float) -> str:
    return f"{rate:.0%}"


def _delta_cell(old: Any, new: Any) -> str:
    """对照表里的「旧 → 新」单元格：两侧一样就只写一次，省得满屏箭头。"""
    before, after = _show(old), _show(new)
    if before == after:
        return _cell(old)
    return f"{_cell(old)} → {_cell(new)}"


def _caliber_lines(payload: Mapping[str, Any]) -> list[str]:
    """报告头的渲染口径两行（EZI-0036 A2）：``renderer=`` 与 ``step_5_95=``。

    老结果 JSON 没有这两个键——这时必须写「未记录」，而不是按默认值写成
    ``none`` / ``ran``：读数的人要能一眼看出「这份分数是 no-renderer 口径」，
    否则开渲染后的新数会被拿去和旧数并列比较（issue A4 明令禁止）。
    """
    renderer = str(payload.get("renderer") or "")
    step = str(payload.get("step_5_95") or "")
    if renderer:
        why = {
            "word": "Word COM 回调已注入，5.95 有渲染器可用",
            "lo": "仅 LibreOffice 可用（无 Word 回调），5.95 靠 LO 分页探测",
            "none": "无渲染器（未注入 word_render_callback、未定位 LibreOffice），5.95 整段跳过",
        }.get(renderer, "自定义口径")
        lines = [f"- renderer={renderer}（{why}）"]
    else:
        lines = ["- renderer=（未记录：老结果 JSON 无此字段，按 **no-renderer** 口径解读，"
                 "不得与开渲染结果并列比较）"]
    if step:
        note = {
            "ran": "5.95 跨页续表真的执行了",
            "skipped:无渲染器": "5.95 未执行——没有渲染器，分页探测不可信",
            "skipped:未声明table_continued_label": "5.95 未执行——spec 未声明 table_continued_label（且文档族无兜底）",
        }.get(step, "逐组见下表")
        lines.append(f"- step_5_95={step}（{note}）")
    else:
        lines.append("- step_5_95=（未记录：老结果 JSON 无此字段）")
    return lines


def render_markdown(payload: Mapping[str, Any]) -> str:
    """两张必需表（逐项通过率 / 失败聚类）+ 一张逐组结果表 + 分母健康度（EZI-0031）。"""
    summary = payload.get("summary") or {}
    cases = payload.get("cases") or []
    # 分母健康度从 summary["items"] 现算（老结果 JSON 没有 health 字段也能报警）
    items = summary.get("items") or []
    health = metric_health(items)
    # 「跑不起来的组」：error 非空 = 编译/打分阶段炸了，0/0 不是「全对」而是「没判」。
    # 摘要要同时进「逐组结果」和「失败聚类」下方的归因行，否则 exit 2 无从解释（EZI-0010）。
    error_rows = _error_rows(cases)
    lines: list[str] = []
    lines.append("# 语料化版式打分报告")
    lines.append("")
    lines.append(f"- manifest：`{payload.get('manifest', '')}`")
    if payload.get("generated_at"):
        lines.append(f"- 生成时间：{payload['generated_at']}")
    if payload.get("workdir"):
        lines.append(f"- 编译工作目录：`{payload['workdir']}`")
    env = (payload.get("env") or {}).get("EARLYZ_SKIP_WORD_RENDER")
    lines.append(f"- EARLYZ_SKIP_WORD_RENDER={env}（=1 时关掉渲染链；不是全部 COM，见 README 局限 2）")
    lines.extend(_caliber_lines(payload))
    lines.append("")
    lines.append(
        f"**总体：{summary.get('passed', 0)}/{summary.get('judged', 0)} 项通过"
        f"（{_pct(float(summary.get('score', 0.0)))}）；"
        f"{summary.get('cases', 0)} 组，FAIL {summary.get('failed', 0)} 项，"
        f"N-A {summary.get('not_applicable', 0)} 项；"
        f"组均正确率 {_pct(float(summary.get('macro_score', 0.0)))}**"
    )
    # ── A2：恒 N-A 的项必须在报告顶部被点名 ──
    # 这类项通过率恒为 0.0，却因为「名字在名单里」被算成覆盖到了；分母被抽掉一半
    # 而总分上涨（EZI-0031 背景 2）就是这么发生的。放在最顶上，翻不到就算没看见。
    constant = health["constant_na_items"]
    if constant:
        detail = "、".join(
            f"{_cell(r['name'])}（{r['cases']} 组全部 N-A，0 可判）" for r in constant
        )
        lines.append("")
        lines.append(
            f"> ⚠️ {CONSTANT_NA_MARK} {len(constant)} 项在其出现的全部组上都没有可判样本："
            f"{detail} —— 没有样本就没有通过率，这几项**不构成任何覆盖**，"
            "别拿它们充判定项名单"
        )
    if summary.get("errors"):
        names = "、".join(e["name"] for e in summary["errors"])
        lines.append("")
        lines.append(f"> ⚠️ 有 {len(summary['errors'])} 组没跑出判定（exit 2）：{names}")
    if summary.get("unmet_expectations"):
        lines.append("")
        lines.append(f"> ⚠️ 未达 `expect_min_score` 的组：{'、'.join(summary['unmet_expectations'])}")
    dur = summary.get("duration_s") or {}
    if dur.get("n"):
        lines.append("")
        lines.append(
            f"- 耗时：{dur['n']} 对合计 {dur['total']:.0f}s；"
            f"平均 {dur['mean']:.0f}s/对，最长 {dur['max']:.0f}s，最短 {dur['min']:.0f}s"
        )
    proj = summary.get("projection") or {}
    if proj:
        lines.append(
            f"- 按实测均速推算：{proj['pairs']} 对 ≈ "
            f"{proj['mean_s_per_pair']:.0f}s × {proj['pairs']} = **{proj['hours']:.1f} 小时**"
            "（一次跑完不现实，实际用 --limit / --only 抽样）"
        )
    lines.append("")

    # ── 分母健康度（EZI-0031 A1）：先看尺子有没有效，再看分 ──
    lines.append("## 分母健康度")
    lines.append("")
    if not items:
        lines.append("（无判定项）")
    else:
        lines.append(
            f"- 总 N-A 占比：{_pct(health['na_ratio'])}"
            f"（{health['not_applicable']}/{health['total']} 项次）；"
            f"可判样本 {health['judged']}/{health['total']}"
            " —— 通过率的分母就是这些可判样本"
        )
        lines.append(
            f"- N-A 占比上限：{_pct(health['limit'])}；超过即打 {DEAD_METRIC_MARK}"
            "（一半以上的组判不了，通过率就不代表产品）"
        )
        suspect = health["suspect_dead_metrics"]
        if not suspect:
            lines.append(
                f"- {DEAD_METRIC_MARK}：无（没有判定项的 N-A 占比超过 {_pct(health['limit'])}）"
            )
        else:
            lines.append(
                f"- {DEAD_METRIC_MARK}：{len(suspect)} 项"
                f"（N-A 占比 > {_pct(health['limit'])}，通过率不可信）"
            )
            lines.append("")
            lines.append("| 项名 | 类目 | 组数 | N-A | 可判 | N-A 占比 | 标记 |")
            lines.append("| --- | --- | ---: | ---: | ---: | ---: | --- |")
            for r in suspect:
                marks = DEAD_METRIC_MARK + (CONSTANT_NA_MARK if r["constant_na"] else "")
                lines.append(
                    f"| {_cell(r['name'])} | {_cell(r['category'])} | {r['cases']} "
                    f"| {r['not_applicable']} | {r['judged']} | {_pct(r['na_ratio'])} | {marks} |"
                )
            lines.append("")
            lines.append(
                "（每项的 N-A 占比同时列在下方「逐项通过率」表的同名一列；"
                "恒为 N-A 的项已在报告顶部单独告警）"
            )
    lines.append("")

    # ── 表 1：逐项通过率 ──
    lines.append("## 逐项通过率")
    lines.append("")
    if not items:
        lines.append("（无判定项）")
    else:
        lines.append("| 项名 | 类目 | 组数 | PASS | FAIL | N-A | N-A 占比 | 通过率 |")
        lines.append("| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |")
        for it in items:
            # 没有可判样本时不能写 0%——0/0 是「没判」，写成 0% 就成了「全错」
            rate = _pct(float(it["pass_rate"])) if it.get("judged") else "—"
            lines.append(
                f"| {_cell(it['name'])} | {_cell(it.get('category') or '')} | {it['cases']} "
                f"| {it['passed']} | {it['failed']} | {it['not_applicable']} "
                f"| {_pct(_row_na_ratio(it))} | {rate} |"
            )
    lines.append("")

    # ── 表 2：失败聚类 ──
    lines.append("## 失败聚类")
    lines.append("")
    failures = summary.get("failures") or []
    if not failures:
        lines.append("（没有失败项）")
    else:
        lines.append("| 项名 | 失败组数 | 典型期望值 vs 实测值 |")
        lines.append("| --- | ---: | --- |")
        for f in failures:
            bits = [
                f"{ex['case']}：模板 {_cell(ex['expected'])} vs 产物 {_cell(ex['actual'])}"
                for ex in f.get("examples") or []
            ]
            detail = "<br>".join(bits) or "（无实测值）"
            lines.append(f"| {_cell(f['name'])} | {f['failed_cases']} | {detail} |")
    lines.append("")

    # ── 跑不起来的组：exit 2 的归因（原来只有 0/0 一行，看不出为什么）──
    if error_rows:
        lines.append(
            f"**跑不起来的组：{len(error_rows)}**"
            "（编译或打分阶段抛异常，没有产出任何判定，退出码 2）"
        )
        lines.append("")
        for name, why in error_rows:
            lines.append(f"- `{name}`：{why}")
        lines.append("")

    # ── 表 3：逐组结果（形态对比就看这张）──
    lines.append("## 逐组结果")
    lines.append("")
    if not cases:
        lines.append("（无用例）")
    else:
        # 只有真有组跑不起来时才加「错误」列：正常报告一列都不许多，否则等于污染报告
        show_error = bool(error_rows)
        head = "| 组 | 文档 | 模板 | PASS/判定 | 正确率 | 耗时(s) | 退出码 |"
        sep = "| --- | --- | --- | ---: | ---: | ---: | ---: |"
        if show_error:
            head += " 错误 |"
            sep += " --- |"
        lines.append(head)
        lines.append(sep)
        for c in cases:
            note = "（失败）" if c.get("error") else ""
            secs = f"{c['duration_s']:.0f}" if c.get("duration_s") is not None else "—"
            row = (
                f"| {_cell(c['name'])}{note} | {_cell(c['doc'])} | {_cell(c['template'])} "
                f"| {c['passed']}/{c['judged']} | {_pct(float(c['score']))} | {secs} "
                f"| {c['exit_code']} |"
            )
            if show_error:
                # 错误摘要走 _cell：竖线转义、换行压平，保证表格不被 traceback 撑歪
                why = _cell(error_summary(str(c.get("error") or ""))) if c.get("error") else "—"
                row += f" {why} |"
            lines.append(row)
    lines.append("")
    return "\n".join(lines)


def total_failed_items(results: Iterable[CaseResult]) -> int:
    """便捷计数（CLI 汇总行用）。"""
    return sum(r.failed for r in results)


# ─────────────────── 口径变更的只读对照（EZI-0031 A3） ───────────────────
#
# 问题：判定口径一年改了 9 轮，每轮「翻了几项、翻在哪些组」都靠人肉 diff 报告。
# 这里把这件事机械化：两份既有结果 JSON → 逐项 verdict 翻转清单 + 分母变动。
#
# 为什么是「比两份结果」而不是「拿新逻辑重判旧产物」：
#   * 重判必须调 ``layout_fidelity``，本 issue 明确不动判定逻辑，也不该在这里起
#     第二条判定入口（会让「口径」有两个真相）；
#   * 结果 JSON 里本来就带 expected/actual/verdict，重判能给出的信息它都有；
#   * 只读 = 不编译、不开 Word、不碰产物，随时能跑，不会因为 5 分钟/对 变成摆设。
#
# 代价（写清楚）：必须真的用新口径跑过一遍同一批产物（换逻辑后重跑 harness），
# 对照才成立——它证明的是「两次跑之间判定变了多少」，不是「如果我改了会怎样」。


def _payload_cases(payload: Any, where: str) -> list[Mapping[str, Any]]:
    """校验一份结果 JSON 的骨架，返回 ``cases``。形状不对就报错，不静默当空。"""
    if not isinstance(payload, Mapping):
        raise ValueError(f"{where} 必须是结果 JSON 的对象，实际为 {type(payload).__name__}")
    cases = payload.get("cases")
    if not isinstance(cases, list) or not cases:
        raise ValueError(f"{where} 缺 `cases` 或为空（不是 corpus_layout_score.py 的产物？）")
    for idx, case in enumerate(cases):
        if not isinstance(case, Mapping):
            raise ValueError(f"{where} cases[{idx}] 必须是对象")
        if not isinstance(case.get("name"), str) or not case["name"]:
            raise ValueError(f"{where} cases[{idx}] 缺 `name`")
        if not isinstance(case.get("items"), list):
            raise ValueError(f"{where} cases[{idx}]（{case['name']}）缺 `items`")
    return cases


def load_result_payload(path: str | Path) -> dict[str, Any]:
    """读一份 ``corpus_layout_score.py`` 写出的结果 JSON（只读，不改文件）。"""
    p = Path(path)
    if not p.is_file():
        raise ValueError(f"结果 JSON 不存在：{p}")
    try:
        payload = json.loads(p.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"结果 JSON 不是合法 JSON：{p}（{exc}）") from exc
    _payload_cases(payload, str(p))
    return payload


def verdict_map(payload: Any, *, where: str = "结果") -> dict[tuple[str, str], dict[str, Any]]:
    """结果 JSON → ``{(组名, 项名): {verdict, expected, actual, category}}``。

    verdict 走 ``_item_dict`` 归一（大小写 / ``NA`` 别名），非法值当场报错；
    同一组里同名项出现两次也报错——否则对照会静默吃掉一条判定。
    """
    out: dict[tuple[str, str], dict[str, Any]] = {}
    for case in _payload_cases(payload, where):
        for raw in case["items"]:
            item = _item_dict(raw)
            key = (case["name"], item["name"])
            if key in out:
                raise ValueError(f"{where} `{case['name']}` 的 `{item['name']}` 重复出现")
            out[key] = {
                "verdict": item["verdict"],
                "expected": item.get("expected"),
                "actual": item.get("actual"),
                "category": item.get("category") or "",
            }
    return out


def _counts_from_map(vmap: Mapping[tuple[str, str], Mapping[str, Any]]) -> dict[str, Any]:
    passed = sum(1 for v in vmap.values() if v["verdict"] == PASS)
    failed = sum(1 for v in vmap.values() if v["verdict"] == FAIL)
    na = len(vmap) - passed - failed
    return {"total": len(vmap), "passed": passed, "failed": failed,
            "judged": passed + failed, "not_applicable": na,
            "score": _rate(passed, passed + failed)}


def _payload_meta(payload: Mapping[str, Any], label: str) -> dict[str, Any]:
    raw = payload.get("summary")
    summary: Mapping[str, Any] = raw if isinstance(raw, Mapping) else {}
    return {
        "label": label,
        "manifest": str(payload.get("manifest") or ""),
        "generated_at": str(payload.get("generated_at") or ""),
        "exit_code": payload.get("exit_code"),
        "cases": summary.get("cases"),
        "total": summary.get("total"),
        "passed": summary.get("passed"),
        "judged": summary.get("judged"),
        "not_applicable": summary.get("not_applicable"),
        "score": summary.get("score"),
        # 口径标签（EZI-0036）：对照报告必须显示两侧是不是同一种渲染口径
        "renderer": payload.get("renderer"),
        "step_5_95": payload.get("step_5_95"),
    }


def _flip_rank(change: str) -> int:
    """翻转的严重度：恶化 > 改善 > 覆盖变化（PASS/FAIL ↔ N-A）。只影响排序。"""
    if change == f"{PASS}→{FAIL}":
        return 0
    if change == f"{FAIL}→{PASS}":
        return 1
    return 2


def compare_payloads(baseline: Any, current: Any, *,
                     baseline_label: str = "", current_label: str = "") -> dict[str, Any]:
    """只读对照：两份结果 JSON → 逐项 verdict 翻转清单 + 分母变动（EZI-0031 A3）。

    ``baseline`` 是旧口径跑出来的结果，``current`` 是现行口径跑出来的结果；两者必须
    来自**同一批产物**（同一 manifest、同一组名），否则 ``only_in_*`` 会告诉你对不齐。

    返回字典（可直接 ``json.dumps``）::

        flips              逐项翻转 [{case, item, category, old, new, change,
                                     expected, old_actual, actual}]
        flip_counts        {"FAIL→PASS": 4, ...}
        improved/regressed 改善 = FAIL→PASS，恶化 = PASS→FAIL
        coverage_changed   PASS/FAIL 与 N-A 之间的翻转数（判得出来了 / 判不出来了）
        only_in_baseline   只在旧结果里出现的（组, 项）——项被删或改名
        only_in_current    只在现行结果里出现的
        denominator_shift  可判样本数变了的项（分母缩水/长大，就是总分漂移的根）
        score_delta/judged_delta/na_delta  整体位移
    """
    old = verdict_map(baseline, where=f"基线 {baseline_label}".strip())
    new = verdict_map(current, where=f"现行 {current_label}".strip())
    old_meta = _payload_meta(baseline, baseline_label)
    new_meta = _payload_meta(current, current_label)
    for meta, vmap, payload in ((old_meta, old, baseline), (new_meta, new, current)):
        if meta.get("judged") is None:
            meta.update(_counts_from_map(vmap))
        if meta.get("cases") is None:
            meta["cases"] = len(_payload_cases(payload, "结果"))

    common = [k for k in old if k in new]
    flips: list[dict[str, Any]] = []
    counts: Counter[str] = Counter()
    for key in common:
        before, after = old[key], new[key]
        if before["verdict"] == after["verdict"]:
            continue
        change = f"{before['verdict']}→{after['verdict']}"
        counts[change] += 1
        flips.append({
            "case": key[0],
            "item": key[1],
            "category": after["category"] or before["category"],
            "old": before["verdict"],
            "new": after["verdict"],
            "change": change,
            # 两侧的期望/实测都留：口径变更常常同时改「期望」和「实测」（EZI-0022 双峰区间），
            # 只留新侧会让人误以为旧判定是拿新期望判的。
            "old_expected": before.get("expected"),
            "expected": after.get("expected"),
            "old_actual": before.get("actual"),
            "actual": after.get("actual"),
        })
    flips.sort(key=lambda f: (_flip_rank(f["change"]), f["item"], f["case"]))

    only_old = sorted(({"case": k[0], "item": k[1], "verdict": old[k]["verdict"]}
                       for k in old if k not in new), key=lambda e: (e["case"], e["item"]))
    only_new = sorted(({"case": k[0], "item": k[1], "verdict": new[k]["verdict"]}
                       for k in new if k not in old), key=lambda e: (e["case"], e["item"]))

    def _by_item(vmap: Mapping[tuple[str, str], Mapping[str, Any]]) -> dict[str, dict[str, int]]:
        agg: dict[str, dict[str, int]] = {}
        for (_case, name), v in vmap.items():
            row = agg.setdefault(name, {"cases": 0, "judged": 0, "na": 0})
            row["cases"] += 1
            if v["verdict"] == NA:
                row["na"] += 1
            else:
                row["judged"] += 1
        return agg

    old_by_item, new_by_item = _by_item(old), _by_item(new)
    shift: list[dict[str, Any]] = []
    for name in sorted(set(old_by_item) | set(new_by_item)):
        before = old_by_item.get(name, {"cases": 0, "judged": 0, "na": 0})
        after = new_by_item.get(name, {"cases": 0, "judged": 0, "na": 0})
        if before["judged"] != after["judged"] or before["cases"] != after["cases"]:
            shift.append({
                "item": name,
                "old_cases": before["cases"], "new_cases": after["cases"],
                "old_judged": before["judged"], "new_judged": after["judged"],
                "old_na": before["na"], "new_na": after["na"],
            })

    judged_delta = na_delta = None
    if old_meta.get("judged") is not None and new_meta.get("judged") is not None:
        judged_delta = int(new_meta["judged"]) - int(old_meta["judged"])
    if old_meta.get("not_applicable") is not None and new_meta.get("not_applicable") is not None:
        na_delta = int(new_meta["not_applicable"]) - int(old_meta["not_applicable"])
    score_delta = None
    if old_meta.get("score") is not None and new_meta.get("score") is not None:
        score_delta = round(float(new_meta["score"]) - float(old_meta["score"]), 4)

    improved = counts.get(f"{FAIL}→{PASS}", 0)
    regressed = counts.get(f"{PASS}→{FAIL}", 0)
    return {
        "baseline": old_meta,
        "current": new_meta,
        "cases_compared": len({k[0] for k in common}),
        "verdicts_compared": len(common),
        "flips": flips,
        "flip_counts": dict(sorted(counts.items())),
        "improved": improved,
        "regressed": regressed,
        "coverage_changed": sum(counts.values()) - improved - regressed,
        "only_in_baseline": only_old,
        "only_in_current": only_new,
        "denominator_shift": shift,
        "score_delta": score_delta,
        "judged_delta": judged_delta,
        "na_delta": na_delta,
    }


def _fmt_meta(m: Mapping[str, Any]) -> str:
    """``基线：`x.json`（manifest …，生成 …）182/186（97.85%）``。"""
    bits = []
    if m.get("label"):
        bits.append(f"`{m['label']}`")
    if m.get("manifest"):
        bits.append(f"manifest `{m['manifest']}`")
    if m.get("generated_at"):
        bits.append(f"生成 {m['generated_at']}")
    if m.get("renderer"):
        bits.append(f"渲染 {m['renderer']}")
    elif "renderer" in m:
        bits.append("渲染 未记录（no-renderer 口径）")
    head = "（" + "，".join(bits) + "）" if bits else ""
    score = m.get("score")
    tail = (f"{m.get('passed')}/{m.get('judged')} 项通过（{_pct(float(score))}）"
            if score is not None else "（结果里没有 summary）")
    return f"{head}{tail}"


def render_comparison_markdown(cmp: Mapping[str, Any]) -> str:
    """对照结果 → Markdown（翻转清单 + 分母变动 + 对不齐的项）。"""
    b = cmp.get("baseline") or {}
    c = cmp.get("current") or {}
    flips = cmp.get("flips") or []
    lines = ["# 判定口径只读对照", ""]
    lines.append(f"- 基线（旧口径）：{_fmt_meta(b)}")
    lines.append(f"- 现行（新口径）：{_fmt_meta(c)}")
    lines.append(
        f"- 对齐样本：{cmp.get('cases_compared', 0)} 组 / {cmp.get('verdicts_compared', 0)} 项；"
        f"翻转 **{len(flips)}** 项（改善 {cmp.get('improved', 0)}，"
        f"恶化 {cmp.get('regressed', 0)}，覆盖变化 {cmp.get('coverage_changed', 0)}）；"
        f"仅基线有 {len(cmp.get('only_in_baseline') or [])}，"
        f"仅现行有 {len(cmp.get('only_in_current') or [])}"
    )
    if cmp.get("score_delta") is not None:
        lines.append(
            f"- 分数：{_pct(float(b.get('score') or 0.0))} → {_pct(float(c.get('score') or 0.0))}"
            f"（{float(cmp['score_delta']) * 100:+.2f}pt）"
        )
    if cmp.get("judged_delta") is not None:
        lines.append(
            f"- 可判样本：{b.get('judged')} → {c.get('judged')}"
            f"（{int(cmp['judged_delta']):+d}）；"
            f"N-A {b.get('not_applicable')} → {c.get('not_applicable')}"
            f"（{int(cmp.get('na_delta') or 0):+d}）"
        )
    lines.append("")

    lines.append("## verdict 翻转清单")
    lines.append("")
    if not flips:
        lines.append(
            f"（没有 verdict 翻转：{cmp.get('verdicts_compared', 0)} 项的判定完全一致）"
        )
    else:
        lines.append("| 项名 | 组 | 旧 verdict | 新 verdict | 旧期望 → 新期望 | 旧实测 → 新实测 |")
        lines.append("| --- | --- | --- | --- | --- | --- |")
        for f in flips:
            lines.append(
                f"| {_cell(f['item'])} | {_cell(f['case'])} | {_cell(f['old'])} "
                f"| {_cell(f['new'])} | {_delta_cell(f.get('old_expected'), f.get('expected'))} "
                f"| {_delta_cell(f.get('old_actual'), f.get('actual'))} |"
            )
    lines.append("")

    shift = cmp.get("denominator_shift") or []
    if shift:
        lines.append("## 分母变动（可判样本数变了的项）")
        lines.append("")
        lines.append("| 项名 | 旧组数 | 新组数 | 旧可判 | 新可判 | 旧 N-A | 新 N-A |")
        lines.append("| --- | ---: | ---: | ---: | ---: | ---: | ---: |")
        for s in shift:
            lines.append(
                f"| {_cell(s['item'])} | {s['old_cases']} | {s['new_cases']} "
                f"| {s['old_judged']} | {s['new_judged']} | {s['old_na']} | {s['new_na']} |"
            )
        lines.append("")

    for key, title in (("only_in_baseline", "只在基线里有（项被删/改名/该组没跑）"),
                       ("only_in_current", "只在现行里有（新增项/新增组）")):
        rows = cmp.get(key) or []
        if not rows:
            continue
        lines.append(f"## {title}")
        lines.append("")
        for row in rows:
            lines.append(f"- `{row['case']}` / {row['item']}：{row['verdict']}")
        lines.append("")
    return "\n".join(lines)


def main(argv: Sequence[str] | None = None) -> int:
    """``python -m earlyz.validation.corpus_scoring``：口径变更只读对照（A3）。

    退出码：``0``=对照完成（有没有翻转都是 0，这是 diff 工具不是闸门）、
    ``2``=输入读不了/不是本 harness 的产物。
    """
    ap = argparse.ArgumentParser(
        prog="python -m earlyz.validation.corpus_scoring",
        description=(
            "判定口径只读对照（EZI-0031 A3）：给两份既有结果 JSON，列逐项 verdict 翻转。"
            "只读——不重判、不编译、不开 Word。"
        ),
    )
    ap.add_argument("--baseline-json", required=True, type=Path, help="旧口径跑出的结果 JSON")
    ap.add_argument("--current-json", required=True, type=Path, help="新口径跑出的结果 JSON")
    ap.add_argument("--md", type=Path, help="把完整对照写成 Markdown")
    ap.add_argument("--json", dest="json_out", type=Path, help="把对照写成 JSON")
    ap.add_argument("-q", "--quiet", action="store_true", help="不打印摘要")
    args = ap.parse_args(argv)

    try:
        cmp = compare_payloads(
            load_result_payload(args.baseline_json),
            load_result_payload(args.current_json),
            baseline_label=str(args.baseline_json),
            current_label=str(args.current_json),
        )
        if args.json_out:
            _write(args.json_out, json.dumps(cmp, ensure_ascii=False, indent=2))
        if args.md:
            _write(args.md, render_comparison_markdown(cmp))
    except (ValueError, OSError) as exc:
        print(f"对照失败：{exc}", file=sys.stderr)
        return 2
    if not args.quiet:
        b, c = cmp["baseline"], cmp["current"]
        print(
            f"对照 {b.get('label') or '基线'} → {c.get('label') or '现行'}："
            f"{cmp['cases_compared']} 组 / {cmp['verdicts_compared']} 项，"
            f"翻转 {len(cmp['flips'])} 项（改善 {cmp['improved']}，恶化 {cmp['regressed']}，"
            f"覆盖变化 {cmp['coverage_changed']}）"
        )
        if cmp["score_delta"] is not None:
            print(f"分数 {_pct(float(b.get('score') or 0.0))} → "
                  f"{_pct(float(c.get('score') or 0.0))}（{cmp['score_delta'] * 100:+.2f}pt）；"
                  f"可判样本 {b.get('judged')} → {c.get('judged')}")
        for f in cmp["flips"]:
            print(f"  {f['old']}→{f['new']}  {f['item']}  [{f['case']}]")
        for row in cmp["only_in_baseline"]:
            print(f"  仅基线有  {row['item']}  [{row['case']}] {row['verdict']}")
        for row in cmp["only_in_current"]:
            print(f"  仅现行有  {row['item']}  [{row['case']}] {row['verdict']}")
        for s in cmp["denominator_shift"]:
            print(f"  分母变动  {s['item']}  可判 {s['old_judged']}→{s['new_judged']}")
    return 0


def _write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


if __name__ == "__main__":  # pragma: no cover - 入口薄壳，逻辑都在 main/compare_payloads
    raise SystemExit(main())
