"""格式检测报告修复动作（M1，2026-10-07 边界定稿）。

- normalize_cjk_punctuation：半角标点→全角「低风险子集」归一（三级守卫，
  B2 用户拍板：可自动改 + 三保险=对照表/可撤/不静默）。
  守卫：①保护区（参考文献/英文摘要/类代码段）②模式排除（ASCII 邻接 /
  小数与版本 / 扩展名 / O(N) 数学）③只改「中文语义」的括号与句读——含中文
  语境数字括号：公式引用（4.1）与内联序号 (1)(2)(3)→（1）（2）（3）
  （2026-10-08 用户报「序号没处理」补齐；英文语境数字括号一律不动）。
  实测（2026-10-07 送检论文 3.1 万字）：盲改错误率 ~49%；本子集 22 处全中、
  误伤 0（含 @AuditLog 代码行守卫）。每处变更全量入 changes 账本。
- set_footer_page_number_font：页脚页码段字号（检测报告「页面设置-页码
  问题-字号问题」的修复动作；取值=报告规范值）。无页码域时显式报错。

铁律：空参不静默（显式 error）；幂等（二次运行 changed=0）；不改保护区。
"""

from __future__ import annotations

import re
from collections import Counter
from pathlib import Path

from lxml import etree

from earlyz.actions.new_actions import (
    W,
    _part_names,
    _read_part,
    _rpr_insert_ordered,
    _set_run_properties,
    _write_part,
)
from earlyz.spec.decompiler import _CN_SIZE_PT

# ── 中文标点表 ─────────────────────────────────────────────────

_PUNCT_MAP = {",": "，", ".": "。", ";": "；", ":": "：", "?": "？", "!": "！"}
_FUNC_OPEN, _FUNC_CLOSE = "（", "）"
_HALF_OPEN, _HALF_CLOSE = "(", ")"
_FULLW_SET = set("，。；：？！（）【】《》、“”‘’…—·")


def _is_cjk(ch: str) -> bool:
    return ("\u4e00" <= ch <= "\u9fff") or ("\u3400" <= ch <= "\u4dbf")


# 中文语境标点（数字括号判定的前邻上下文——「包括：(1)」「方法；(2)」）。
# 2026-10-08 用户报「序号没处理」：此前仅 CJK 汉字前邻放行，(1)(2)(3) 落在
# 中文标点（：；，等）后一律进保守保留 → 序号括号永远不转全角。
_CJK_CTX_PUNCT = frozenset("，。、；：！？…—～·\u3000）（】」』》〉〕］｝")


def _digit_paren_cn_ctx(chars, k: int, close: int) -> bool:
    """数字括号是否中文语境。

    前邻 = 汉字或中文标点 → 是；段首/制表符前导 → 要求括号后（跳过空格）为汉字
    （防英文段落行首 "(1)" 误伤）；其余（ASCII 邻接）→ 否。
    """
    prev = chars[k - 1] if k > 0 else ""
    if _is_cjk(prev) or prev in _CJK_CTX_PUNCT:
        return True
    if prev in ("", "\t", "\n"):
        nxt = next((c for c in chars[close + 1:] if c not in (" ", "\t")), "")
        return _is_cjk(nxt)
    return False


def _norm_space(text: str) -> str:
    return text.replace(" ", "").replace("\u3000", "")


# 类代码段识别（三级守卫 Tier-1「代码保护区」）——样本实测：
# 附录 Java 代码段（@PutMapping/@PreAuthorize/@AuditLog/字段声明）必须整体跳过。
_CODE_SIGNS = re.compile(
    r"://|::|@\w+\s*\(|@Override"
    r"|\b(?:public|private|protected|import|package|return|void|class|interface|enum)\s"
    r"|\b(?:SELECT|INSERT|UPDATE|DELETE|CREATE|VARCHAR|DECIMAL|PRIMARY|FOREIGN)\b",
)


def _looks_like_code(text: str) -> bool:
    if _CODE_SIGNS.search(text):
        return True
    stripped = text.rstrip()
    if stripped.endswith((";", "{", "}")) and stripped.count(" ") >= 2:
        # 行尾分号/花括号且含缩进 → 代码行（中文正文极少此形态）
        ascii_ratio = sum(1 for c in stripped if c.isascii()) / max(len(stripped), 1)
        if ascii_ratio > 0.25:
            return True
    return False


def _para_stream(pel):
    """段落文本流 [(char, w:t元素, 元素内偏移) | (char, None, None)]。

    tab/br 以哨兵字符（\\t \\n）入流（元素 None，永不转换）——保证邻接判断
    不被制表符/换行骗过；w:instrText（域指令）不入流。
    """
    stream: list[tuple[str, object, int]] = []
    for el in pel.iter():
        tag = el.tag
        if tag == f"{W}t":
            txt = el.text or ""
            for k, ch in enumerate(txt):
                stream.append((ch, el, k))
        elif tag == f"{W}tab":
            stream.append(("\t", None, None))
        elif tag in (f"{W}br", f"{W}cr"):
            stream.append(("\n", None, None))
    return stream


def _para_text(pel) -> str:
    return "".join(t.text or "" for t in pel.iter(f"{W}t"))


def _is_heading_para(pel) -> bool:
    ppr = pel.find(f"{W}pPr")
    if ppr is None:
        return False
    if ppr.find(f"{W}outlineLvl") is not None:
        return True
    ps = ppr.find(f"{W}pStyle")
    if ps is not None:
        v = ps.get(f"{W}val") or ""
        if "Heading" in v or "标题" in v:
            return True
    return False


def _zone_map(paras) -> dict[int, str]:
    """保护区：参考文献 / 英文摘要。paras = [(pel, text)] → {idx: zone}。"""
    zones: dict[int, str] = {}
    n = len(paras)
    # 参考文献区（取最后一次出现的正文标题，避免命中目录条目）
    ref_start = None
    for i in range(n - 1, -1, -1):
        if _norm_space(paras[i][1]).strip() == "参考文献":
            ref_start = i
            break
    if ref_start is not None:
        j = ref_start + 1
        while j < n:
            t = _norm_space(paras[j][1]).strip()
            if t.startswith("附") and len(t) <= 8:
                break
            if j > ref_start + 1 and _is_heading_para(paras[j][0]):
                break
            zones[j] = "refs"
            j += 1
    # 封面（模板托管区）：文档第一节 = 到第一个承载 sectPr 的段落为止。
    # 2026-10-25 实测：封面槽位「学院(系):」被半角括号归一改成「学院（系）:」，
    # 触发 cover polluted 硬错误 → 整篇三次拒绝交付（用户明确说"不用管"也发不出去）。
    # 封面文字是模板填充区，不属于正文标点规范的作用范围。
    _cover_end = None
    for _i, (_pel, _t) in enumerate(paras):
        _ppr = _pel.find(f"{W}pPr")
        if _ppr is not None and _ppr.find(f"{W}sectPr") is not None:
            _cover_end = _i
            break
    if _cover_end is not None:
        for _i in range(0, _cover_end + 1):
            zones.setdefault(_i, "cover")
    else:
        # 单节文档（无分节符）没有「第一节」可用——改用结构判定：
        # 第一个标题之前的段落视为封面/前置区。2026-10-25 实测：
        # 「通用公文标准格式(…)-带注解版」单节，封面注解里的「顶格.」被改成
        # 「顶格。」→ cover polluted，整篇拒绝交付。
        _first_head = None
        for _i, (_pel, _t) in enumerate(paras):
            _ppr = _pel.find(f"{W}pPr")
            if _ppr is None:
                continue
            _st = _ppr.find(f"{W}pStyle")
            _ol = _ppr.find(f"{W}outlineLvl")
            _is_head = False
            if _st is not None and "heading" in (_st.get(f"{W}val") or "").lower():
                _is_head = True
            if _ol is not None and (_ol.get(f"{W}val") or "") in ("0", "1", "2"):
                _is_head = True
            if _is_head:
                _first_head = _i
                break
        # 前置标记也是边界：封面之后通常紧跟「摘要 / 目录 / ABSTRACT」。
        # 2026-10-25 实测：首标题恰好在第 0 段时 `if _first_head:` 为假（0 是假值），
        # 封面区等于没保护 → 「学院(系):」照旧被改成「学院（系）:」。
        _marker = None
        for _i, (_pel, _t) in enumerate(paras):
            _tt = "".join(str(_t).split()).upper()
            if _tt in ("摘要", "目录", "ABSTRACT", "中文摘要", "英文摘要"):
                _marker = _i
                break
        _bounds = [b for b in (_first_head if _first_head is not None else None, _marker)
                   if b is not None]
        _end = min(_bounds) if _bounds else None
        if _end is not None and _end > 0:
            for _i in range(0, _end):
                zones.setdefault(_i, "cover")
    # 英文摘要区：ABSTRACT … Key words
    en_start = None
    for i in range(n):
        ts = paras[i][1].strip()
        if ts.upper().startswith("ABSTRACT") and len(ts) < 60:
            en_start = i
            break
    if en_start is not None:
        j = en_start
        while j < n:
            ts = paras[j][1].strip().lower()
            if ts.startswith("key words") or ts.startswith("keywords"):
                zones[j] = "en_abs"
                break
            zones[j] = "en_abs"
            if j - en_start > 120:  # 防御：异常文档不吞全文
                break
            j += 1
    return zones


# ── 扫描（action / verifier / 工具预览 共用，只读）────────────────

def _scan_root(root, *, kept_limit: int = 12) -> dict:
    """只读扫描（同一棵 lxml 树内）——action 与独立扫描共用此实现。

    铁律：调用者若要在结果上执行转换（normalize_cjk_punctuation），必须把
    同一棵树的 root 传进来，否则 stream 里的元素引用属于另一棵树，写回无效
    （2026-10-07 首版实测 bug：二次运行不幂等 / verifier False）。
    """
    paras = [(pel, _para_text(pel)) for pel in root.iter(f"{W}p")]
    zones = _zone_map(paras)
    convertible: list[dict] = []
    kept: list[dict] = []
    zone_skips: Counter = Counter()
    for pi, (pel, text) in enumerate(paras):
        if not text.strip():
            continue
        zone = zones.get(pi)
        if zone:
            zone_skips[zone] += 1
            continue
        if _looks_like_code(text):
            zone_skips["code"] += 1
            continue
        stream = _para_stream(pel)
        chars = [c for c, _, _ in stream]
        k, n = 0, len(chars)
        while k < n:
            ch = chars[k]
            if ch == _HALF_OPEN:
                depth, close = 0, -1
                for kk in range(k, n):
                    if chars[kk] == _HALF_OPEN:
                        depth += 1
                    elif chars[kk] == _HALF_CLOSE:
                        depth -= 1
                        if depth == 0:
                            close = kk
                            break
                if close == -1:
                    kept.append({"para": pi, "sample": text[:60], "reason": "括号未闭合"})
                    k += 1
                    continue
                inner = "".join(chars[k + 1:close])
                if any(_is_cjk(c) for c in inner):
                    kind = "paren_cn"
                elif re.fullmatch(r"\d+(?:\.\d+)*", inner) and _digit_paren_cn_ctx(
                        chars, k, close):
                    # 数字括号（中文语境）：点分=公式/图表引用；纯数字=内联序号（(1)(2)(3)
                    # →（1）（2）（3），2026-10-08 用户报「序号没处理」）。
                    kind = "formula_ref" if "." in inner else "seq_no"
                else:
                    kept.append({
                        "para": pi,
                        "sample": "".join(chars[max(0, k - 20):min(n, close + 21)]),
                        "reason": "非中文内容/非公式引用（代码·数学·英文）",
                    })
                    k = close + 1
                    continue
                a, b = max(0, k - 18), min(n, close + 19)
                token_before = "".join(chars[a:k]) + _HALF_OPEN + inner + _HALF_CLOSE + "".join(chars[close + 1:b])
                token_after = "".join(chars[a:k]) + _FUNC_OPEN + inner + _FUNC_CLOSE + "".join(chars[close + 1:b])
                convertible.append({
                    "para": pi, "kind": kind, "i": k, "close": close,
                    "stream": stream, "before": token_before, "after": token_after,
                })
                k = close + 1
                continue
            if ch in _PUNCT_MAP:
                prev = chars[k - 1] if k > 0 else ""
                nxt = chars[k + 1] if k + 1 < n else ""
                if _is_cjk(prev) and _is_cjk(nxt):
                    a, b = max(0, k - 18), min(n, k + 19)
                    convertible.append({
                        "para": pi, "kind": "punct", "i": k, "close": -1,
                        "stream": stream,
                        "before": "".join(chars[a:k]) + ch + "".join(chars[k + 1:b]),
                        "after": "".join(chars[a:k]) + _PUNCT_MAP[ch] + "".join(chars[k + 1:b]),
                    })
            k += 1
    return {
        "convertible": convertible,
        "kept": kept[:kept_limit],
        "kept_total": len(kept),
        "zone_skips": dict(zone_skips),
    }


def scan_cjk_punctuation(path: str | Path, *, kept_limit: int = 12) -> dict:
    """只读扫描（独立读文件；预览/verifier 用——不要在结果上执行转换）。"""
    return _scan_root(_read_part(Path(path), "word/document.xml"),
                      kept_limit=kept_limit)


# ── Action 1：半角标点低风险子集归一 ─────────────────────────────

def _conversion_feasible(c: dict) -> bool:
    """本项能否安全转换：所需字符元素全部可寻址。

    括号类（paren_cn/formula_ref/seq_no）必须成对可改——只改半边会留下
    半全角混排（`（2）(3)`），宁可整体不转；与 verifier 同源判定。
    """
    stream = c["stream"]
    if c["kind"] in ("paren_cn", "formula_ref", "seq_no"):
        return all(stream[off][1] is not None for off in (c["i"], c["close"]))
    return stream[c["i"]][1] is not None


def normalize_cjk_punctuation(target_path: Path, params: dict) -> dict:
    """低风险子集：中文语义括号/公式引用括号/中文句读 → 全角。

    2026-10-14（复现 tx_6129fb0f5547 事故）：单遍转换与收敛核对口径不一致——
    `(2)(3)` 连写时「(3)」按转换前邻接受保守跳过，但「(2)」转全角后它的
    邻接判定变宽松，verifier 重扫即判「仍有可转项」→ 假失败 + 漏转。转换
    会改变后续项的上下文判定，必须循环到 fixpoint（单调收敛，2-3 遍即尽）。
    """
    path = Path(target_path)
    root = _read_part(path, "word/document.xml")
    changes: list[dict] = []
    kept_total = 0
    kept_samples: list = []
    zone_skips: dict = {}
    stuck_keys: set = set()
    for _pass in range(8):  # fixpoint；8 遍封顶
        scan = _scan_root(root)  # 同一棵树：stream 元素引用可直接改、直接写回
        kept_total = scan["kept_total"]
        kept_samples = scan["kept"]
        zone_skips = scan["zone_skips"]
        items = scan["convertible"]
        if not items:
            break
        progressed = False
        for c in items:
            key = (c["para"], c["kind"], c["before"])
            if key in stuck_keys:
                continue
            if not _conversion_feasible(c):
                # 元素不可寻址（跨 run 等）→ 本项无从下手；verifier 同源放行
                stuck_keys.add(key)
                continue
            stream = c["stream"]
            if c["kind"] in ("paren_cn", "formula_ref", "seq_no"):
                for off in (c["i"], c["close"]):
                    ch, el, rel = stream[off]
                    new = _FUNC_OPEN if ch == _HALF_OPEN else _FUNC_CLOSE
                    el.text = (el.text or "")[:rel] + new + (el.text or "")[rel + 1:]
            else:
                ch, el, rel = stream[c["i"]]
                el.text = (el.text or "")[:rel] + _PUNCT_MAP[ch] + (el.text or "")[rel + 1:]
            changes.append({
                "para": c["para"],
                "kind": c["kind"],
                "before": c["before"],
                "after": c["after"],
            })
            progressed = True
        if not progressed:
            break
    if not changes and not stuck_keys:
        return {"ok": True, "changed": 0, "converted": 0, "changes": [],
                "kept_total": kept_total, "zone_skips": zone_skips,
                "note": "未发现可安全归一的半角标点（保护区与保守项均未动）"}
    _write_part(path, "word/document.xml", root)
    out = {
        "ok": True,
        "changed": len(changes),
        "converted": len(changes),
        "changes": changes[:100],
        "changes_truncated": max(0, len(changes) - 100),
        "kept_total": kept_total,
        "kept_samples": kept_samples,
        "zone_skips": zone_skips,
    }
    if stuck_keys:
        out["stuck_unaddressable"] = len(stuck_keys)
    return out


def verify_normalize_cjk_punctuation(target_path: Path, params: dict = None) -> bool:
    """复查：全部可寻址的低风险项已归一（与动作同源的收敛判定）。

    2026-10-14：只放行「转换器本就无从下手」的残留——字符元素不可寻址的
    项动作无法转换，不得据此判失败（否则假失败经一票否决拦掉整次交付）。
    """
    try:
        scan = scan_cjk_punctuation(Path(target_path))
        for c in scan["convertible"]:
            if not _conversion_feasible(c):
                continue  # 不可寻址 → 动作侧同样跳过
            return False
        return True
    except Exception:  # noqa: BLE001
        return False


# ── Action 2：页脚页码字号 ─────────────────────────────────────

def _resolve_size(params: dict) -> tuple[int, str]:
    """params → (半磅值, 展示标签)。空参/未知字号 → ValueError（不静默）。"""
    cn = str(params.get("cn_size") or "").strip()
    pt_raw = params.get("size_pt")
    if cn:
        pt = _CN_SIZE_PT.get(cn)
        if pt is None:
            raise ValueError(f"未知中文字号: {cn!r}（支持：初号/小初/一号…小七 系列）")
        return int(round(pt * 2)), cn
    if pt_raw is not None:
        try:
            pt = float(pt_raw)
        except (TypeError, ValueError) as e:
            raise ValueError(f"size_pt 不是数字: {pt_raw!r}") from e
        if not (2.0 <= pt <= 72.0):
            raise ValueError(f"size_pt 超出合理范围(2-72磅): {pt}")
        return int(round(pt * 2)), f"{pt}磅"
    raise ValueError("必须提供 cn_size（如\"五号\"）或 size_pt（磅）——缺参显式拒绝")


def _footer_parts(path: Path) -> list[str]:
    return [n for n in _part_names(path)
            if re.fullmatch(r"word/footer\d*\.xml", n)]


def _para_has_page_field(pel) -> bool:
    instr = "".join(el.text or "" for el in pel.iter(f"{W}instrText"))
    return bool(re.search(r"\bPAGE\b", instr))


def footer_page_number_sizes(path: str | Path) -> list[dict]:
    """只读：各页脚 part 中页码段的字号现状（预览/证据用）。"""
    p = Path(path)
    out: list[dict] = []
    for part in _footer_parts(p):
        root = _read_part(p, part)
        sizes: list[int | None] = []
        paras_n = 0
        for pel in root.iter(f"{W}p"):
            if not _para_has_page_field(pel):
                continue
            paras_n += 1
            for r in pel.iter(f"{W}r"):
                sz = r.find(f"{W}rPr/{W}sz")
                try:
                    sizes.append(int(sz.get(f"{W}val")) if sz is not None else None)
                except (TypeError, ValueError):
                    sizes.append(None)
        if paras_n:
            out.append({"part": part, "paras": paras_n, "sizes": sizes})
    return out


def set_footer_page_number_font(target_path: Path, params: dict) -> dict:
    """页脚页码段字号修复（w:sz + w:szCs 写入页码所在段全部 run）。"""
    path = Path(target_path)
    half_points, label = _resolve_size(params)
    parts = _footer_parts(path)
    touched: list[dict] = []
    runs_total = 0
    for part in parts:
        root = _read_part(path, part)
        part_before: list[str] = []
        part_runs = 0
        hit = False
        for pel in root.iter(f"{W}p"):
            if not _para_has_page_field(pel):
                continue
            hit = True
            runs = list(pel.iter(f"{W}r"))
            if not runs:
                continue
            for r in runs:
                sz = r.find(f"{W}rPr/{W}sz")
                try:
                    part_before.append(sz.get(f"{W}val") if sz is not None else "—")
                except (TypeError, ValueError):
                    part_before.append("—")
            _set_run_properties(pel, {"font_size_pt": half_points / 2}, target_runs=runs)
            for r in runs:
                rpr = r.find(f"{W}rPr")
                if rpr is None:
                    rpr = etree.Element(f"{W}rPr")
                    r.insert(0, rpr)
                szcs = rpr.find(f"{W}szCs")
                if szcs is None:
                    szcs = etree.Element(f"{W}szCs")
                    _rpr_insert_ordered(rpr, szcs)
                szcs.set(f"{W}val", str(half_points))
            part_runs += len(runs)
        if hit:
            _write_part(path, part, root)
            touched.append({"part": part, "before": sorted(set(part_before)),
                            "runs": part_runs})
            runs_total += part_runs
    if not touched:
        raise ValueError("页脚里没有找到页码域（PAGE）——无法修复页码字号，请确认页脚含自动页码")
    return {
        "ok": True,
        "changed": len(touched),
        "half_points": half_points,
        "size": label,
        "parts": touched,
        "runs": runs_total,
    }


def verify_footer_page_number_font(target_path: Path, params: dict = None) -> bool:
    """复查：所有页码段 run 的 sz 均为目标值。"""
    try:
        half_points, _label = _resolve_size(params or {})
    except ValueError:
        return False
    try:
        p = Path(target_path)
        found = False
        for part in _footer_parts(p):
            root = _read_part(p, part)
            for pel in root.iter(f"{W}p"):
                if not _para_has_page_field(pel):
                    continue
                found = True
                for r in pel.iter(f"{W}r"):
                    sz = r.find(f"{W}rPr/{W}sz")
                    val = sz.get(f"{W}val") if sz is not None else None
                    if val is None or int(val) != half_points:
                        return False
        return found
    except Exception:  # noqa: BLE001
        return False
