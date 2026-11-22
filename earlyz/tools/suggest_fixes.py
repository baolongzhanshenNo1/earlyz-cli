"""§22.5 P2 建议层执行工具（2026-08-25 真机链路发现缺口后补全）。

场景：用户**对话描述**修正指令（「第1章与绪论之间应为一个空格」「目录中
章标题页码应加粗」）——LLM 把指令结构化为 registry 白名单 action 列表，
调本工具执行：

- R1/R2 → 立即执行（handler + verifier 门禁）；
- R3（内容修改/破坏性）→ 不执行，列入 pending_confirm（LLM 向用户展示，
  用户明确确认后 confirm_all=true 再调用执行——§9.4 授权模型对话确认）；
- 未知 action / 裸 para_index（无 text_snippet 防漂移）→ rejected；
- 执行前备份到 .earlyz_history（恢复语义，安全兜底）。

铁律：LLM 永不直接改 OOXML——本工具只调度 registry 权威 handler。
"""

from __future__ import annotations

import hashlib
import re
import shutil
import time
from pathlib import Path
from typing import Optional

from earlyz.actions.registry import RiskLevel


def _risk_of(registry, action_name: str) -> Optional[RiskLevel]:
    try:
        spec = registry.get(action_name)
        return spec.risk if spec else None
    except Exception:  # noqa: BLE001
        return None


def _jsonable(obj):
    """动作返回值 → JSON 可序列化结构（透传净化，2026-10-07）。

    透传前，这些返回值从不进入消息/SSE 的 json.dumps（严格模式、无 default
    兜底）；动作可能返回 Path/bytes 等类型，直接透传会在序列化处抛
    TypeError 导致任务中断——统一净化：bytes→长度占位、未知类型→str。
    """
    if obj is None or isinstance(obj, (str, int, float, bool)):
        return obj
    if isinstance(obj, dict):
        return {str(k): _jsonable(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_jsonable(v) for v in obj]
    if isinstance(obj, bytes):
        return f"<{len(obj)} bytes>"
    return str(obj)


# ── 建议层参数要点表（2026-08-26 Action 自查审计产出）──────────
# 审计（空参轰炸全白名单）：27 个 action 空参时 handler 直接键取 → KeyError
#（caller 喂回，浪费一轮）；静默无参动作（无参标记）为合法 noop；其余需参数
# action 静默 noop = 假成功。本表 = 校验（params 全空→rejected）+ hint（参数要点）
# 双用；以"（无参" 开头的条目为合法无参动作（不拦）。
ACTION_PARAM_HINTS: dict[str, str] = {
    "add_footer_reference": "section_index(节序号)",
    "add_page_number": "footer_path(页脚 part 路径)+font_ascii/font_size 可选",
    "apply_section_columns": "section_columns(分栏定义)",
    "apply_template_header_footer": "template_path(模板 docx 路径)",
    "clear_section_footers": "section_index(节序号)",
    "clear_section_headers": "section_index(节序号)",
    "clone_header_footer_from_spec": "kind(header/footer)+source_style_id",
    "configure_header_footer_variants": "variant(如 first_even_odd)",
    "create_section_headers": "section_headers(标题列表)",
    "ensure_section_page_breaks": "（无参，仅重排场景：节分页检查）",
    "fit_media_to_page": "（无参，图片页适配）",
    "format_caption_fonts": "（无参，图注字体规范化）",
    "format_caption_numbers": "（无参，图注编号规范化）",
    "format_footers": "properties(页脚格式属性)",
    "format_headers": "properties(页眉格式属性)",
    "format_paragraph": "para_index+text_snippet 定位段",
    "format_references": "（无参，参考文献域规范化）",
    "format_table_cells": "（无参，表格单元格规范化）",
    "insert_red_header": "org_name(红头单位名)",
    "insert_section_break": "heading_text(插入点标题文本)",
    "map_paragraph_to_list": "selector(段落定位)",
    "protect_image_para": "para_index+text_snippet",
    "repair_fields": "（无参，域结构+断链引用书签+书目编号修复）",
    "repair_section_references": "（无参，节引用修复）",
    "restart_list_numbering": "num_id+start_value",
    "set_body_font_ascii": "style_id+font_ascii",
    "set_cell_format": "valign 或 margins(至少一项)",
    "set_header_font": "header_path+font_ascii/font_size",
    "set_header_text": "header_path+text",
    "set_list_level_format": "abstract_num_id+ilvl+properties",
    "set_equation_from_latex": "selector 或 match_text 定位公式段 + latex(目标公式)",
    # 2026-10-07 真机事故：用户回「公式转对象」→ LLM 调 convert_text_equations 传空
    # params → 无表项守卫跳过 → 空转 changed=False → 转述为「无公式被匹配」。
    # 入表后：空参→rejected(missing_params)+hint；白名单行也带上参数要点。
    "convert_text_equations": "conversions 必填：[{match_text 公式原文, latex 目标LaTeX}…]（整段式/句内式自动判别：句内式原位转换、编号不动）；用户已确认 apply=true",
    # 2026-10-07：公式—编号点线引导（用户点名才加；style 必填防默认静默）
    "set_equation_leader": "style 必填：dot=点线（……）/ dash=虚线（----）/ none=去掉点线；match_text 可选（只处理指定公式）",
    # 2026-10-07 M1 报告修复动作（空参守卫 + 参数要点）
    "normalize_cjk_punctuation": "（无参，半角标点低风险子集归一；参考文献/英文摘要/类代码段不动）",
    "set_footer_page_number_font": "cn_size(如\"五号\") 或 size_pt(磅)——二者至少一个；缺参将被拒绝",
    # 2026-10-08 M2 题注/图目录（apply 必填：先预览对照清单，用户确认后落盘）
    "renumber_caption_numbers": "apply 必填：false=只出对照清单预览（不落盘），true=按清单落盘；kinds=figure/table 可选（默认全部）；style=dash/dot 可选",
    "generate_caption_toc": "apply 必填：false=只出预览，true=生成/重建图目录表目录；kinds=figure/table 可选（默认全部）",
    # 2026-10-08 M3 交叉引用（apply 必填：先预览待替换清单，用户确认后落盘）
    "generate_cross_references": "apply 必填：false=只出待替换清单预览（不落盘），true=落盘生成引用域；kinds=figure/table/equation 可选（默认全部）",
    # 2026-10-08 文献引用联动 / 章号引用联动 / 纯文本表格转换（apply 必填：先预览后落盘）
    "link_citation_references": "apply 必填：false=清单预览，true=落盘；只改单个 [n]，组合引用只报不改；superscript 默认 true（[n] 上标；要正文形态传 false）",
    "link_chapter_references": "apply 必填：false=出清单预览（不落盘），true=落盘；只改「第 N 章」提及，无对应标题只报不改",
    "convert_text_table": "先 apply=false 出候选行列预览；确认后 match(题注)+cells/columns 再 apply=true（未确认禁止落盘）",
    # 2026-10-08 M3 封面构建（template_path+apply 必填：先预览后落盘）
    "build_cover_from_template": "template_path 必填：模板绝对路径或上传文件名（包 id user:<sha12> 也可；解析不到自动换会话模板）；apply 必填：false=克隆预览，true=落盘；zones 默认 [\"cover\"]，可加 \"statement\"；slots={模板标签:值（题目/日期 特例；英文题目行用「英文题目」写第二行，先问用户）——取值照抄目标文档，取不到才问}",


    "superscript_front_marks": "（无参，作者行/单位行编号角标上标）",
    "set_page_border": "borders(边框定义)",
    "set_page_margins": "section_index+properties(top/bottom/left/right)",
    "set_page_number_type": "section_index+fmt(start 可选)",
    "set_paragraph_alignment": "para_index+text_snippet+val(center/left/right/both；val 或 alignment 别名均兼容)",
    "set_paragraph_border_shading": "borders 或 shading(至少一项)",
    "set_paragraph_format": "selector(段落定位)+properties",
    "set_paragraph_spacing": "para_index+text_snippet+spacing",
    "set_paragraph_tabs": "tabs(制表位定义)",
    "set_run_format": "properties(run 格式)",
    "set_table_alignment": "properties",
    "set_tbl_body_font": "tbl_index+font_ascii/font_eastAsia",
    "set_tbl_cell_valign": "tbl_index+valign",
    "set_tbl_col_widths": "tbl_index+widths",
    "set_tbl_font": "tbl_index+font_ascii/font_eastAsia",
    "set_tbl_header_font": "tbl_index+font_ascii/font_eastAsia",
    "set_tbl_header_repeat": "tbl_index+repeat(bool)",
    "set_toc_entry_styles": "（无参，目录条目样式对齐）",
    "convert_ole_to_image": "（无参，OLE 转图）",
    "clean_heading_text": "text_snippet(目标段前 20 字符，无空格) 或 para_index",
    "bold_toc_page_numbers": "（无参，全部 TOC1 章标题条目页码加粗）",
    "set_style": "style_id+properties(样式属性)",
    "insert_toc": "（无参，插入/重建目录域）",
}


def _validate_action(registry, action: dict) -> tuple[str, str | None, str | None]:
    """校验单个 action。返回 (issue, reason, hint)：issue 空=通过。

    hint 仅在 missing_params 时给出（拒绝原因里的参数要点，LLM 补参重试指引）。
    """
    name = (action or {}).get("action") or ""
    params = (action or {}).get("params") or {}
    if not name:
        return "rejected", "missing_action", None
    if _risk_of(registry, name) is None:
        return "rejected", "unknown_action", None
    # §4.4 兼容层：裸 para_index 拒绝；para_index+text_snippet（防漂移冻结）允许
    if "para_index" in params and not params.get("text_snippet"):
        return "rejected", "naked_para_index", None
    # 2026-08-26 Action 自查：需参数 action 传空 params → handler 静默 noop=
    # 假成功（clean_heading_text 真机翻版）——统一拦截为 rejected，LLM 补参重试
    hint = ACTION_PARAM_HINTS.get(name)
    if hint and not hint.startswith("（无参") and not params:
        return "rejected", "missing_params", hint
    return "", None, None


def _backup(target: Path) -> Path | None:
    """执行前备份（.earlyz_history 惯例，恢复语义）。"""
    try:
        hist = target.parent / ".earlyz_history"
        hist.mkdir(parents=True, exist_ok=True)
        stamp = time.strftime("%Y%m%d_%H%M%S")
        dst = hist / f"{_short_stem(target.stem)}_suggest_{stamp}{target.suffix}"
        shutil.copy2(target, dst)
        return dst
    except Exception:  # noqa: BLE001
        return None


def _short_stem(stem: str) -> str:
    """产物主名清洗：剥 uid 前缀链 + 历史时间戳链 + 限长（2026-10-08 两轮）。

    背景：上传与工作副本会在原名上叠加 (a) 8 位十六进制 uid 前缀（每次上传
    加一层）(b) 「-suggest-YYYYMMDD-HHMMSS」时间戳后缀（每轮加一段）——
    真机产物名一度 130+ 字符（用户投诉「文件名称太长了，能不能短一些」）。
    清洗后基线名稳定：剥全部 uid 前缀、剥全部历史戳（旧式 -suggest- 与
    新式 -MMDD-HHMMSS），截断到 40 字符；新戳由 _work_copy 挂单段。

    2026-10-08 起委托共享命名模块（earlyz.tools.naming）——与交付出口
    （agent_loop._deliver_product）同一实现，防两处口径漂移。
    """
    from earlyz.tools.naming import clean_stem
    return clean_stem(stem) or stem[:40]


def _work_copy(target: Path) -> tuple[Path, Path]:
    """创建执行工作副本（原文件永不触碰——Word/其他程序锁定免疫）。

    返回 (copy_path, backup_path)：copy_path=执行对象；backup_path=原文件备份。
    产物名 = 清洗基线 + 紧凑时间戳（-MMDD-HHMMSS）；同名冲突自动加序号。
    """
    stamp = time.strftime("%m%d-%H%M%S")
    base = _short_stem(target.stem)
    copy = target.parent / f"{base}-{stamp}{target.suffix}"
    k = 2
    while copy.exists():
        copy = target.parent / f"{base}-{stamp}-{k}{target.suffix}"
        k += 1
    shutil.copy2(target, copy)
    try:
        # copy2 会保留原文件的只读/属性——副本必须可写（执行对象）
        copy.chmod(0o644)
    except OSError:
        pass
    return copy, (_backup(target) or target)


def execute_suggested_fixes(target: str | Path, actions: list[dict], registry,
                            *, confirm_all: bool = True,
                            allow_empty: bool = False) -> dict:
    """建议层执行（§22.5 P2 对话确认链路）——**工作副本语义 + 指令即授权**。

    原文件永不修改：先把 target 复制为 `*-suggest-<ts>.docx` 工作副本，全部
    R1/R2/R3 动作在副本上执行；成功后返回 `result_path`（产物指针切换点，
    dispatch 层更新 ctx.working → 用户下载/预览指向新产物）。原文件被 Word
    占用时免疫 PermissionError（读 copy 不需要写锁）。

    **confirm_all 默认 True（2026-08-25 用户拍板：用户明确指令即授权）**——
    用户消息里的修正指令（「第1章与绪论之间应为一个空格」「目录页码应加粗」）
    已经是确认，不再二次询问；R3（内容修改/破坏性）随指令直接执行。
    调用方只有 LLM（SYSTEM_PROMPT 约束：仅用户明确提及的修正才列入 actions，
    未提及的变更型动作不执行、改为报告告知）。
    """
    target = Path(target)
    if not target.exists():
        return {"error": f"目标文件不存在: {target}"}
    if not allow_empty and not actions:
        return {"applied": [], "pending_confirm": [], "rejected": [],
                "error": "actions 为空"}
    applied: list[dict] = []
    pending: list[dict] = []
    rejected: list[dict] = []
    to_execute: list[tuple[dict, RiskLevel]] = []
    for a in actions or []:
        issue, reason, hint = _validate_action(registry, a)
        if issue == "rejected":
            entry: dict = {"action": a.get("action") or "?", "reason": reason}
            if hint:
                entry["hint"] = f"需要参数：{hint}"
            rejected.append(entry)
            continue
        risk = _risk_of(registry, a["action"])
        if risk is not None and risk.value > RiskLevel.R2.value:
            # R3：确认后才执行
            if confirm_all:
                to_execute.append((a, risk))
            else:
                pending.append({"action": a["action"], "params": a.get("params", {}),
                                "risk": risk.value,
                                "hint": "内容修改/破坏性动作——需要用户确认"})
            continue
        to_execute.append((a, risk))

    out: dict = {"applied": applied, "pending_confirm": pending, "rejected": rejected}
    if not to_execute:
        return out

    # 工作副本执行（原文件永不触碰）
    copy_path, backup = _work_copy(target)
    # changed 信号（2026-08-26 22:45 真机绕圈）：执行前后 document.xml hash 对比——
    # 「已居中」等幂等动作返回 changed=false（no-op）——LLM 收到后一句话答复，
    # 不再重复 apply/换工具/explain 确认（LLM 绕圈 5 次调用的根因）。
    def _doc_hash(p: Path) -> str:
        # 全包内容对比（2026-10-08）：只比 document.xml 会漏判「页眉页脚/样式/
        # 编号等非正文部件」的真实改动（changed=false 误报）——凡包内部件内容
        # 有变即 changed=true，交付层据此判断「这轮到底有没有产物可给用户」。
        try:
            import zipfile as _z
            h = hashlib.sha256()
            with _z.ZipFile(p) as z:
                for _name in sorted(z.namelist()):
                    h.update(_name.encode("utf-8"))
                    h.update(b"\x00")
                    h.update(z.read(_name))
            return h.hexdigest()
        except Exception:  # noqa: BLE001
            return ""
    _h_before = _doc_hash(copy_path)
    for a, _risk in to_execute:
        name = a["action"]
        params = a.get("params") or {}
        entry: dict = {"action": name, "status": "ok"}
        try:
            handler = registry.handler(name)
            if handler is None:
                entry.update({"status": "failed", "error": "no_handler"})
            else:
                result = handler(str(copy_path), params)
                # 2026-10-07：动作返回载荷必须透传（预览类动作的 zones_found/槽位
                # 当前值/对照清单/待替换清单全在返回值里）。旧实现丢弃返回值 →
                # LLM 只见 status=ok，封面/题注/交叉引用等「先预览」流程全部盲飞
                # （封面事故实锤）；返回 dict 内的 action 级 error 同时提升为
                # failed——「模板不存在」等不再被静默吞成 ok。
                if isinstance(result, dict):
                    entry["result"] = _jsonable(result)
                    if result.get("error"):
                        entry.update({"status": "failed",
                                      "error": str(result.get("error"))})
                verifier = registry.verifier(name)
                if entry["status"] == "ok" and verifier is not None \
                        and not verifier(str(copy_path), params):
                    entry.update({"status": "failed",
                                  "error": "verifier_false"})
        except Exception as e:  # noqa: BLE001
            entry.update({"status": "failed",
                          "error": f"{type(e).__name__}: {e}"})
        applied.append(entry)
    out["changed"] = _h_before != _doc_hash(copy_path)

    ok_any = any(entry["status"] == "ok" for entry in applied)
    # 2026-10-09 用户报「公式 6 成 1 败 → 到最后没有下载卡、没有产物」：
    # 批内部分失败时（7 条公式 1 条失败 → verifier 整批 False → ok_any=False），
    # 旧逻辑连已落盘的 6 条改动一起丢弃（不返回 result_path + 工作副本 unlink）
    # ——交付层据此判「无产物」，卡与文件全丢。修正：changed=True 即有真实改动，
    # 永远返回产物指针、保留副本（失败明细在 applied 里）；ok_any 仅用于
    # 「无改动且全失败」时清理。预览（apply=false）与幂等 no-op 均 changed=false，
    # 仍不曝光产物（2026-10-08 口径不变）。
    if out["changed"]:
        out["result_path"] = str(copy_path)
        out["backup_path"] = str(backup) if backup != target else None
        if not ok_any:
            out["partial_failed"] = sum(1 for e in applied if e["status"] != "ok")
    elif not ok_any:
        # 全部失败且包内无任何改动：清理工作副本（不污染产物目录）
        try:
            copy_path.unlink(missing_ok=True)
        except Exception:  # noqa: BLE001
            pass
    return out
