"""T10 — 41 Action 目录注册表 + 兼容适配层（3.2 §6 / §13.2）。

把现有 writer._ACTION_HANDLERS / fixer._VERIFIERS 全量注册进 ActionRegistry：
- handler 签名 (target_path: Path, params: dict) → 与 TransactionExecutor 兼容；
- verifier 签名 (target_path, params, get_xml, idx) → 包装为 (path, params) -> bool，
  存在 severity=error 的 FixerDiagnostic 即判定失败；
- 风险分级来自架构文档 §6.4；R3 不暴露给 Agent；
- 未命中风险表的 action 显式报告（不允许静默默认值掩盖缺口）。
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from earlyz.actions.registry import ActionRegistry, ActionSpec, RiskLevel
from earlyz.actions import new_actions
from earlyz.actions import equation_dsmt4
from earlyz.actions import equation_text
from earlyz.actions import style_actions
from earlyz.actions import list_actions
from earlyz.actions import header_footer_actions
from earlyz.actions import new_actions_suggest
from earlyz.actions import report_fixes
from earlyz.actions import captions
from earlyz.actions import numbering_fix
from earlyz.actions import cross_refs
from earlyz.actions import cover
from earlyz.actions import text_table
from earlyz.engine import fixer, writer

# ── 风险分级（架构文档 §6.4 定稿） ──────────────────────────────

R1_ACTIONS = {
    "set_style", "set_body_font_ascii", "set_paragraph_alignment",
    "set_paragraph_spacing", "format_paragraph", "set_code_block_style",
    "set_header_font", "set_page_margins", "set_tbl_font",
    "set_tbl_header_font", "set_tbl_body_font", "set_tbl_borders",
    "set_tbl_cell_valign", "set_tbl_col_widths", "set_tbl_header_repeat",
    "protect_image_para", "ensure_toc_tabs",
}

R2_ACTIONS = {
    "add_footer_reference", "add_page_number", "clear_section_footers",
    "clear_section_headers", "create_section_headers", "insert_red_header",
    "insert_section_break", "insert_toc", "normalize_image_paragraphs",
    "normalize_list_numbering", "normalize_tbl_grid", "remove_extra_sectpr",
    "set_header_text", "set_page_number_type", "set_sectpr_type",
    "set_toc_entry_styles", "set_tr_cant_split", "set_images_inline",
}

R3_ACTIONS = {
    "clean_heading_text", "convert_ole_to_image", "extract_textbox_captions",
    "split_merged_cells", "strip_style_numbering", "normalize_cjk_spacing",
}

# 3.2 新 P0 action（§7.1，Format Compiler 新链路执行端）
NEW_ACTIONS = {
    "set_doc_defaults": (RiskLevel.R2, new_actions.set_doc_defaults, new_actions.verify_doc_defaults, "styles"),
    "set_paragraph_format": (RiskLevel.R1, new_actions.set_paragraph_format, new_actions.verify_paragraph_format, "paragraph"),
    "set_page_margins_all": (RiskLevel.R1, new_actions.set_page_margins_all, new_actions.verify_page_margins_all, "section"),
    "set_table_alignment": (RiskLevel.R2, new_actions.set_table_alignment, new_actions.verify_table_alignment, "table"),
    "format_headers": (RiskLevel.R2, new_actions.format_headers, new_actions.verify_headers, "section"),
    "format_footers": (RiskLevel.R2, new_actions.format_footers, new_actions.verify_footers, "section"),
    "build_section_skeleton": (RiskLevel.R2, new_actions.build_section_skeleton,
                               new_actions.verify_section_skeleton, "section"),
    "ensure_toc_section_break": (RiskLevel.R2, new_actions.ensure_toc_section_break,
                                 new_actions.verify_toc_section_break, "section"),
    "ensure_footer_parts": (RiskLevel.R2, new_actions.ensure_footer_parts,
                            new_actions.verify_footer_refs, "section"),
    "split_section_parts": (RiskLevel.R2, new_actions.split_section_parts,
                            new_actions.verify_section_parts, "section"),
    "apply_section_header_text": (RiskLevel.R1, new_actions.apply_section_header_text,
                                  new_actions.verify_section_header_text, "section"),
    "strip_cover_header_text": (RiskLevel.R1, new_actions.strip_cover_header_text,
                                new_actions.verify_cover_header_clean, "section"),
    "normalize_front_matter": (RiskLevel.R2, new_actions.normalize_front_matter,
                               new_actions.verify_front_matter, "section"),
    "create_caption_cross_refs": (RiskLevel.R1, new_actions.create_caption_cross_refs,
                                  new_actions.verify_caption_cross_refs, "document"),
    "remove_example_residuals": (RiskLevel.R3, new_actions.remove_example_residuals,
                                 new_actions.verify_remove_residuals, "document"),
    "normalize_ooxml_order": (RiskLevel.R2, new_actions.normalize_ooxml_order,
                              new_actions.verify_ooxml_order, "package"),
    "repair_package_invariants": (RiskLevel.R2, new_actions.repair_package_invariants,
                                  new_actions.verify_package_invariants, "package"),
    "configure_page_numbering": (RiskLevel.R2, new_actions.configure_page_numbering,
                                 new_actions.verify_page_numbering, "section"),
    "restart_list_numbering": (RiskLevel.R2, new_actions.restart_list_numbering,
                               new_actions.verify_list_numbering, "list"),
    "set_paragraph_tabs": (RiskLevel.R1, new_actions.set_paragraph_tabs,
                           new_actions.verify_paragraph_tabs, "paragraph"),
    "set_paragraph_border_shading": (RiskLevel.R1, new_actions.set_paragraph_border_shading,
                                     new_actions.verify_paragraph_border_shading, "paragraph"),
    "set_run_format": (RiskLevel.R1, new_actions.set_run_format,
                       new_actions.verify_run_format, "run"),
    "configure_header_footer_variants": (RiskLevel.R2, new_actions.configure_header_footer_variants,
                                         new_actions.verify_header_footer_variants, "section"),
    "apply_template_header_footer": (RiskLevel.R2, new_actions.apply_template_header_footer,
                                     new_actions.verify_template_header_footer, "section"),
    "clone_header_footer_from_spec": (RiskLevel.R2, new_actions.clone_header_footer_from_spec,
                                      new_actions.verify_header_footer_cloned, "section"),
    "set_page_border": (RiskLevel.R2, new_actions.set_page_border,
                        new_actions.verify_page_border, "section"),
    "fit_table_to_page": (RiskLevel.R2, new_actions.fit_table_to_page,
                          new_actions.verify_table_fit, "table"),
    "set_row_layout": (RiskLevel.R2, new_actions.set_row_layout,
                       new_actions.verify_row_layout, "table"),
    "set_cell_format": (RiskLevel.R1, new_actions.set_cell_format,
                        new_actions.verify_cell_format, "table"),
    "merge_cells": (RiskLevel.R3, new_actions.merge_cells,
                    new_actions.verify_merged, "table"),
    "split_cells": (RiskLevel.R3, new_actions.split_cells,
                    new_actions.verify_split, "table"),
    "normalize_table_caption_block": (RiskLevel.R1, new_actions.normalize_table_caption_block,
                                      new_actions.verify_table_caption_block, "table"),
    "normalize_figure_block": (RiskLevel.R1, new_actions.normalize_figure_block,
                               new_actions.verify_figure_block, "image"),
    "fit_media_to_page": (RiskLevel.R2, new_actions.fit_media_to_page,
                          new_actions.verify_media_fit, "image"),
    "set_alt_text": (RiskLevel.R1, new_actions.set_alt_text,
                     new_actions.verify_alt_text, "image"),
    "repair_fields": (RiskLevel.R2, new_actions.repair_fields,
                      new_actions.verify_fields, "toc"),
    "repair_bookmarks": (RiskLevel.R2, new_actions.repair_bookmarks,
                         new_actions.verify_bookmarks, "toc"),
    "repair_section_references": (RiskLevel.R2, new_actions.repair_section_references,
                                  new_actions.verify_section_references, "section"),
    "format_caption_numbers": (RiskLevel.R1, new_actions.format_caption_numbers,
                               new_actions.verify_caption_numbers, "caption"),
    "format_table_cells": (RiskLevel.R1, new_actions.format_table_cells,
                           new_actions.verify_table_cells, "table"),
    "normalize_footer_parts": (RiskLevel.R1, new_actions.normalize_footer_parts,
                               new_actions.verify_footer_parts, "footer"),
    # 2026-08-17 六大投诉修复 action（18:41 真机产物取证）
    "format_toc_entries": (RiskLevel.R1, new_actions.format_toc_entries,
                           new_actions.verify_toc_entries, "toc"),
    "bold_toc_page_numbers": (RiskLevel.R1, new_actions_suggest.bold_toc_page_numbers,
                              new_actions_suggest.verify_toc_page_numbers, "toc"),
    "ensure_toc_tab_stops": (RiskLevel.R1, new_actions.ensure_toc_tab_stops,
                             new_actions.verify_toc_entries, "styles"),
    # 2026-08-18 用户彻查「结论 vs 结论与展望」：TOC 域缓存是源文档旧的，
    # earlyz 自算条目文本（替代 Word updateFields 覆盖格式）
    "update_toc_entries": (RiskLevel.R1, new_actions.update_toc_entries,
                           new_actions.verify_toc_entries_updated, "toc"),
    # 2026-08-18 用户拍板「参考文献补序号 + 注意模板段落样式」
    "format_references": (RiskLevel.R1, new_actions.format_references,
                          new_actions.verify_references, "refs"),
    # 2026-09-26 用户拍板「目录能自动更新（F9）这个是肯定要的」：条目算好后
    # 补回 TOC 域外壳（缓存=我们的条目，F9 才重建）；不弹窗交给 ensure_update_fields
    "ensure_toc_field": (RiskLevel.R1, new_actions.ensure_toc_field,
                         new_actions.verify_toc_field, "toc"),
    # 2026-08-18 用户彻查「打开就弹更新域」：TOC 域静态化（不弹窗不被重算）
    # 2026-09-26 停用：静态目录在 Word 里按 F9 无法重建，与「能自动更新」相冲。
    # 函数保留（回滚/对照用），但不再注册进工具箱，避免 LLM 选中它拆掉域。
    # 2026-08-18 用户彻查「标题前两个空白行」：标题前空段清理
    "remove_blank_lines_before_headings": (RiskLevel.R1, new_actions.remove_blank_lines_before_headings,
                                           new_actions.verify_no_blank_before_headings, "content"),
    "fit_tables_to_page": (RiskLevel.R2, new_actions.fit_tables_to_page,
                           new_actions.verify_tables_within_page, "table"),
    "normalize_table_widths": (RiskLevel.R1, new_actions.normalize_table_widths,
                               new_actions.verify_table_widths, "table"),
    "apply_section_columns": (RiskLevel.R1, new_actions.apply_section_columns,
                              new_actions.verify_section_columns, "section"),
    "format_caption_fonts": (RiskLevel.R1, new_actions.format_caption_fonts,
                             new_actions.verify_caption_fonts, "caption"),
    # 2026-08-17 彻查第三轮：标题 outlineLvl + 前置章节另起页
    "ensure_heading_outline": (RiskLevel.R1, new_actions.ensure_heading_outline,
                               new_actions.verify_heading_outline, "paragraph"),
    # 2026-11-08 用户实锤「1. 1. 引言」「二级标题 1.27cm 缩进」：样式带 numPr 但
    # numbering.xml 没有该 numId（悬空）→ Word 默认列表兜底多画一个号。
    "reconcile_numbering_refs": (RiskLevel.R1, numbering_fix.reconcile_numbering_refs,
                                 numbering_fix.verify_numbering_refs, "styles"),
    "suppress_duplicate_heading_numbers": (
        RiskLevel.R1, numbering_fix.suppress_duplicate_heading_numbers,
        numbering_fix.verify_no_duplicate_heading_numbers, "paragraph"),
    # EZI-0068：表内段落标题残留（heading 样式单元格被探针当标题+续编号链）
    "clear_table_cell_outline": (
        RiskLevel.R1, new_actions.clear_table_cell_outline,
        new_actions.verify_table_cell_outline, "paragraph"),
    "normalize_equation_paragraphs": (RiskLevel.R1, new_actions.normalize_equation_paragraphs,
                                      new_actions.verify_equation_paragraphs, "paragraph"),
    # 2026-09-28 夜班：前部（作者行/单位行）编号角标 → 上标（run 级，取值出自模板样本）
    "superscript_front_marks": (RiskLevel.R1, new_actions.superscript_front_marks,
                                new_actions.verify_superscript_front_marks, "paragraph"),
    # 2026-09-28 夜班：LaTeX → Word 原生公式（编辑公式能力；显式建议 + 采纳，不静默转换）
    "set_equation_from_latex": (RiskLevel.R2, new_actions.set_equation_from_latex,
                                new_actions.verify_equation_from_latex, "paragraph"),
    # 2026-10-01 深夜：LaTeX → DSMT4 OLE 对象（场景 A 学位论文，自研 MathType 组件，完全兼容）
    "set_equation_dsmt4": (RiskLevel.R2, equation_dsmt4.set_equation_dsmt4,
                          equation_dsmt4.verify_equation_dsmt4, "paragraph"),
    # 2026-10-06 公式文本定型（B/C 双通路，边界文档 W8 + E9 原则）：
    # C：$…$/$$…$$/\[…\] 文本 → MathType 对象（Toggle TeX 语义，自动 + 防货币误伤）
    "toggle_tex_equations": (RiskLevel.R2, equation_text.toggle_tex_equations,
                             equation_text.verify_toggle_tex_equations, "document"),
    # B：人工敲的纯文本公式 → 对象；apply=false 出提案（E9 先给用户确认），true 落盘
    "convert_text_equations": (RiskLevel.R2, equation_text.convert_text_equations,
                               equation_text.verify_convert_text_equations, "paragraph"),
    # 2026-10-07：公式—编号点线引导（模板沉默变体，用户点名才加；默认不加）
    "set_equation_leader": (RiskLevel.R1, equation_text.set_equation_leader,
                            equation_text.verify_equation_leader, "equation"),
    # 2026-09-28 用户拍板「公式编号的生成一定要做」：行间公式自动编号
    # （格式从文档已有编号推断、章编号感知；跳过已有编号 → 幂等）
    "generate_equation_numbers": (RiskLevel.R2, new_actions.generate_equation_numbers,
                                  new_actions.verify_equation_numbers, "equation"),
    # 2026-10-01 公式组件 W3（对标 MathType Insert Chapter/Section Break）：MathType 编号只读
    # 章号（SEQ MTChap \c）却无章节断点 → Word 显示 (0.x)。按章标题补隐藏断点，不动 OLE。
    "insert_equation_chapter_breaks": (RiskLevel.R1, new_actions.insert_equation_chapter_breaks,
                                       new_actions.verify_equation_chapter_breaks, "equation"),
    # 2026-09-28 用户拍板「要加」：行内公式段拆分（公式粘在正文句子里 → 拆成独立公式段）
    "split_inline_equations": (RiskLevel.R2, new_actions.split_inline_equations,
                               new_actions.verify_split_inline_equations, "equation"),
    # 2026-10-17 用户拍板：论文没目录时从零生成（模板有目录规范才启用；有目录 no-op）
    "insert_toc_if_missing": (RiskLevel.R2, new_actions.insert_toc_if_missing,
                              new_actions.verify_toc_generated, "toc"),
    "ensure_section_page_breaks": (RiskLevel.R1, new_actions.ensure_section_page_breaks,
                                   new_actions.verify_section_page_breaks, "section"),
    # 2026-09-25 批3 覆盖扩展：脚注/尾注取值按模板（模板无素材则整体跳过）
    "format_footnotes": (RiskLevel.R1, new_actions.format_footnotes,
                         new_actions.verify_footnotes, "package"),
    # 2026-10-17：脚注编号机制（CJC「每页的脚注从 1 开始编序号」）
    "set_footnote_numbering": (RiskLevel.R2, new_actions.set_footnote_numbering,
                               new_actions.verify_footnote_numbering, "section"),
    # 2026-10-17：表格题注编号样式（IEEE 要求罗马数字）
    "set_table_caption_numbering": (RiskLevel.R1,
                                    new_actions.set_table_caption_numbering,
                                    new_actions.verify_table_caption_numbering,
                                    "package"),
    "ensure_update_fields": (RiskLevel.R1, new_actions.ensure_update_fields,
                             new_actions.verify_update_fields, "settings"),
    # 2026-08-27 别名：LLM 按旧提示输出 update_fields（PROMPT 已统一，历史上下文仍可能带）——兼容
    "update_fields": (RiskLevel.R1, new_actions.ensure_update_fields,
                      new_actions.verify_update_fields, "settings"),
    # 2026-08-19 用户拍板：目录数字打开就精确且不弹窗
    "refresh_page_fields": (RiskLevel.R1, new_actions.refresh_page_fields,
                            new_actions.verify_page_fields_refreshed, "toc"),
    "apply_drop_cap": (RiskLevel.R1, new_actions.apply_drop_cap,
                       new_actions.verify_drop_cap, "paragraph"),
    "apply_table_borders": (RiskLevel.R1, new_actions.apply_table_borders,
                            new_actions.verify_table_borders, "table"),
    "keep_table_together": (RiskLevel.R2, new_actions.keep_table_together,
                            new_actions.verify_keep_table_together, "table"),
    "sync_toc_page_numbers": (RiskLevel.R2, new_actions.sync_toc_page_numbers,
                              new_actions.verify_toc_page_numbers, "toc"),
    "polish_toc_entries": (RiskLevel.R2, new_actions.polish_toc_entries,
                           new_actions.verify_toc_entries, "toc"),
    "center_form_table_cells": (RiskLevel.R2, new_actions.center_form_table_cells,
                                new_actions.verify_center_form_table_cells, "table"),
    # 2026-10-07 M1 格式检测报告修复动作（维普批注 docx 接入；边界见
    # docs/方案-格式检测报告接入与A档七项-2026-10-07.md）：
    # 标点=低风险子集自动改+三保险（B2 拍板）；页码字号=报告规范值
    "normalize_cjk_punctuation": (RiskLevel.R3, report_fixes.normalize_cjk_punctuation,
                                  report_fixes.verify_normalize_cjk_punctuation,
                                  "punctuation"),
    # EZI-0065 A4：源文稿里混入的模板说明段（整段【…】/*…* 包裹）不进产物
    "strip_instruction_paragraphs": (
        RiskLevel.R3, new_actions.strip_instruction_paragraphs,
        new_actions.verify_strip_instruction_paragraphs, "content"),
    "set_footer_page_number_font": (RiskLevel.R2, report_fixes.set_footer_page_number_font,
                                    report_fixes.verify_footer_page_number_font, "footer"),
    # 2026-10-08 M2：图/表目录 + 题注章号（A1/A4；Word COM 真机探针定案
    # 「字面编号 + 隐藏采集域」；边界见 docs/方案-格式检测报告接入与A档七项-2026-10-07.md）
    "renumber_caption_numbers": (RiskLevel.R2, captions.renumber_caption_numbers,
                                 captions.verify_caption_numbers_renumbered, "caption"),
    "generate_caption_toc": (RiskLevel.R2, captions.generate_caption_toc,
                             captions.verify_caption_toc, "toc"),
    # 2026-10-08 M3：交叉引用生成（A2；正文引用 → REF 域，字面编号+书签联动方案）
    "generate_cross_references": (RiskLevel.R2, cross_refs.generate_cross_references,
                                  cross_refs.verify_cross_references, "xref"),
    # 2026-10-08 M3：封面构建（A3；模板示例区块级克隆 + 槽位等宽回填）
    "build_cover_from_template": (RiskLevel.R2, cover.build_cover_from_template,
                                  cover.verify_cover_built, "cover"),
    # 2026-10-08：文献引用联动（正文 [n] ↔ 文末条目 → REF 域；提案→确认）
    "link_citation_references": (RiskLevel.R2, cross_refs.link_citation_references,
                                 cross_refs.verify_citation_links, "xref"),
    # 2026-10-08：章号引用联动（正文「第 X 章」提及 ↔ 章标题 → REF 域；提案→确认）
    "link_chapter_references": (RiskLevel.R2, cross_refs.link_chapter_references,
                                cross_refs.verify_chapter_links, "xref"),
    # 2026-10-08：纯文本表格转换（先给行列方案过目，用户确认后落盘）
    "convert_text_table": (RiskLevel.R2, text_table.convert_text_table,
                           text_table.verify_text_table, "table"),
}

# §10.1 内容变化例外清单——这些 action 经用户授权会修改内容（模板示例残留
# 清理，2026-08-16 用户拍板「模板是用来提取格式的！照搬模板内容干什么？」；
# 半角标点低风险子集归一，2026-10-08 用户拍板「自动改 + 三保险：对照表/
# 可撤/不静默」，见 report_fixes.normalize_cjk_punctuation）。
# validate_delivery 对声明了 content_change 的 action 放行（FAILED→warning+证书记录）。
# 2026-10-14 审计：把真实会写正文文字的动作纳入声明（机制消费方：format_compiler
# 运行期按 journal info 允许内容变化；见 5.7 段）——generate_equation_numbers（公式
# 编号落盘）、renumber_caption_numbers（题注编号改写）、format_references（补 [N] 序号）。
CONTENT_CHANGE_ACTIONS = {"remove_example_residuals", "normalize_cjk_punctuation",
                          "generate_equation_numbers", "renumber_caption_numbers",
                          "format_references",
                          # 2026-10-17：表格题注编号样式会改题注文本（表 1 ↔ 表 I）
                          "set_table_caption_numbering",
                          # 2026-10-22：跨页续表要在续页插一行「原表题（续）」，
                          # 文本必然变化——属显式授权的加字行为（用户要求续表）。
                          "apply_continuation_splits"}

# Phase 2 样式系统（§7.1 P0 清单）——确定性链路使用，暂不暴露给 Agent
STYLE_ACTIONS = {
    "create_or_update_style": (RiskLevel.R1, style_actions.create_or_update_style, style_actions.verify_style_created, "styles"),
    "clone_style_from_spec": (RiskLevel.R2, style_actions.clone_style_from_spec, style_actions.verify_style_cloned, "styles"),
    "clear_direct_formatting": (RiskLevel.R2, style_actions.clear_direct_formatting, style_actions.verify_no_direct_formatting, "paragraph"),
    "map_style": (RiskLevel.R2, style_actions.map_style, style_actions.verify_style_mapped, "paragraph"),
    "apply_paragraph_style": (RiskLevel.R1, style_actions.apply_paragraph_style, style_actions.verify_paragraph_style, "paragraph"),
    "ensure_title_pg_flag": (RiskLevel.R1, header_footer_actions.ensure_title_pg_flag, header_footer_actions.verify_title_pg_flag, "section"),
    "align_data_table_columns": (RiskLevel.R1, new_actions.align_data_table_columns, new_actions.verify_data_table_columns, "table"),
    "mark_continued_tables": (RiskLevel.R2, new_actions.mark_continued_tables, new_actions.verify_continued_tables, "table"),
}

# Phase 2 多级编号系统（合同/标书核心）——确定性链路专用
LIST_ACTIONS = {
    "set_list_level_format": (RiskLevel.R2, list_actions.set_list_level_format, list_actions.verify_list_level_format, "list"),
    "map_paragraph_to_list": (RiskLevel.R2, list_actions.map_paragraph_to_list, list_actions.verify_paragraph_list, "list"),
}

# Phase 2 复杂页眉页脚（first/even/default 三态）——论文前置部分核心
HEADER_FOOTER_ACTIONS = {
    "set_header_reference": (RiskLevel.R2, header_footer_actions.set_header_reference, header_footer_actions.verify_header_reference, "section"),
    "set_footer_reference": (RiskLevel.R2, header_footer_actions.set_footer_reference, header_footer_actions.verify_footer_reference, "section"),
    "set_title_page_flag": (RiskLevel.R1, header_footer_actions.set_title_page_flag, header_footer_actions.verify_title_page_flag, "section"),
    "set_even_odd_headers_flag": (RiskLevel.R1, header_footer_actions.set_even_odd_headers_flag, header_footer_actions.verify_even_odd_headers_flag, "section"),
}

# ── Agent 可见 action 显式清单（ADR-006）──────────────────────────
# 与旧版手写 TOOLS 的 22 个 action 完全一致（保持 Agent 能力不变量）。
# 注意：extract_textbox_captions / normalize_cjk_spacing / split_merged_cells
# 虽为 R3（破坏性），但旧 Agent 图文三轮流程依赖，必须保持暴露。
AGENT_EXPOSED = {
    "set_style", "set_body_font_ascii", "add_page_number", "set_images_inline",
    "set_paragraph_alignment", "set_paragraph_spacing", "format_paragraph",
    "protect_image_para", "set_tbl_header_repeat", "set_tr_cant_split",
    "set_tbl_font", "set_tbl_header_font", "set_tbl_body_font", "set_tbl_borders",
    "set_tbl_col_widths", "set_tbl_cell_valign", "split_merged_cells",
    "normalize_tbl_grid", "set_page_margins", "normalize_list_numbering",
    "normalize_cjk_spacing", "extract_textbox_captions",
    # Phase 2 暴露（安全部分；clear_direct_formatting 破坏性保留内部）
    "create_or_update_style", "clone_style_from_spec", "map_style",
    "set_list_level_format", "map_paragraph_to_list",
    "set_header_reference", "set_footer_reference",
    "set_title_page_flag", "set_even_odd_headers_flag",
    # 2026-10-06 公式转换通路修复：set_equation_*/generate_equation_numbers/
    # split_inline_equations 只注册未暴露（提示词引用了、schema 里却没有 →
    # 模型根本调不了「对话改公式」）；连同 B/C 新动作一并暴露。
    "set_equation_from_latex", "set_equation_dsmt4", "generate_equation_numbers",
    "split_inline_equations", "toggle_tex_equations", "convert_text_equations",
    "set_equation_leader",
    # 2026-10-07 M1 报告修复动作（对话修正指令可用；报告流走 fix_format_report_issues）
    "normalize_cjk_punctuation", "set_footer_page_number_font",
    # 2026-10-08 M2 题注/图目录表目录（对话指令：「按章重排题注编号」「生成图目录/表目录」）
    "renumber_caption_numbers", "generate_caption_toc",
    # 2026-10-08 M3 交叉引用（对话指令：「正文引用做成交叉引用」「引用编号联动」）
    "generate_cross_references",
    # 2026-10-08 M3 封面构建（对话指令：「用模板做封面」「封面排版乱重做」）
    "build_cover_from_template",
    # 2026-10-08 文献引用联动 / 章号引用联动 / 纯文本表格转换（对话指令）
    "link_citation_references", "link_chapter_references", "convert_text_table",
    # 2026-11-08 用户实锤「为什么改不了」：能清「段前分页」的只有 set_paragraph_format，
    # 但它只注册未暴露 → 模型只能用 format_paragraph（没有 page_break_before），
    # 于是瞎拼 {paragraph: 60, page_break_before: false} → 选择器为空 → 命中 174 段
    # → AMBIGUOUS → 回滚 → 用户看到「属性动作执行异常」。暴露 + 给 schema。
    "set_paragraph_format",
}

# ── Agent 中文说明（ADR-006：唯一真相源，agent_loop 的 TOOLS 描述由此生成）──
# 逐 action 的参数 schema（ADR-006 自动生成 apply_operations 的 schema）。
# 缺省时 spec.params_schema 为空对象 → 模型看不到任何参数名，只能瞎猜
# （2026-11-08 实锤：模型传 {paragraph: 60, page_break_before: false}，
#  action 只认 selector+properties → 选择器为空 → 命中 174 段 → AMBIGUOUS → 回滚，
#  用户看到「属性动作执行异常」，两次都改不了）。
_SELECTOR_SCHEMA = {
    "type": "object",
    "description": "段落定位器：semantic_role / node_id / text_hash / paragraph_index 至少给一个",
    "properties": {
        "semantic_role": {"type": "string", "description": "语义角色，如 body / heading1 / appendix_heading / refs_heading"},
        "occurrence": {"type": "integer", "description": "同角色里的第几个（0 基），缺省=全部"},
        "node_id": {"type": "string", "description": "稳定节点 ID（inspect_document 输出）"},
        "text_hash": {"type": "string", "description": "规范化文本 hash（inspect_document 输出）"},
        "paragraph_index": {"type": "integer", "description": "正文第 N 段（0 基，与 inspect_document 的段落编号同口径）"},
    },
}

PARAMS_SCHEMAS: dict[str, dict] = {
    "set_paragraph_format": {
        "type": "object",
        "properties": {
            "selector": _SELECTOR_SCHEMA,
            "properties": {
                "type": "object",
                "description": "要写入的段落属性（只写需要改的 key）",
                "properties": {
                    "jc": {"type": "string", "enum": ["left", "center", "right", "both"]},
                    "keep_next": {"type": "boolean"},
                    "keep_lines": {"type": "boolean"},
                    "page_break_before": {"type": "boolean", "description": "false = 清除段前分页"},
                    "line": {"type": "integer"},
                    "line_spacing": {"type": "number", "description": "倍数行距"},
                    "ind_first_line": {"type": "integer"},
                    "font_east_asia": {"type": "string"},
                    "font_ascii": {"type": "string"},
                    "font_size_pt": {"type": "number"},
                    "bold": {"type": "boolean"},
                },
            },
        },
        "required": ["selector"],
    },
}

ACTION_DESCRIPTIONS = {
    "set_paragraph_format": (
        "set_paragraph_format {selector:{paragraph_index | semantic_role(+occurrence) | "
        "node_id | text_hash}, properties:{jc?, bold?, keep_next?, keep_lines?, "
        "page_break_before?, line?, line_spacing?, ind_first_line?, font_east_asia?, "
        "font_ascii?, font_size_pt?}}：按选择器设置段落属性。"
        "「清除段前分页」= properties.page_break_before:false；"
        "只知道第几段就用 selector.paragraph_index（0 基，与 inspect_document 段落编号同口径）。"
    ),
    "set_style": "set_style {style_id, font_eastAsia?, font_ascii?, font_size_pt?, bold?, alignment?, before_lines?, after_lines?, line?, lineRule?, ind_first_line_chars?}：改某个样式的字体/段落格式（正文段落多时改默认样式一次生效）",
    "set_body_font_ascii": "set_body_font_ascii {style_id, font_ascii}：正文西文字体",
    "add_page_number": "add_page_number {footer_path, format}：加页码",
    "set_images_inline": "set_images_inline {}：将所有浮动图片转为嵌入型",
    "set_paragraph_alignment": "set_paragraph_alignment {para_index, val}：段落对齐（val=center/left/right/both）",
    "set_paragraph_spacing": "set_paragraph_spacing {para_index, before_twips?, after_twips?, line?, lineRule?}：定向修改段落段间距/行距（twips）",
    "format_paragraph": "format_paragraph {para_index, font_eastAsia?, font_ascii?, font_size_pt?, bold?, jc?, ind_first_line_chars?, line?, lineRule?, before_lines?, after_lines?, outline_lvl?}：段落级完整格式（标题/正文/关键词均可用；长款项可用 run_formats 拆分前缀与内容格式）",
    "protect_image_para": "protect_image_para {para_index, keepLines?, keepNext?, fixLineRule?}：保护段落不跨页/不被裁切",
    "set_tbl_header_repeat": "set_tbl_header_repeat {tbl_index}：表格表头跨页重复",
    "set_tr_cant_split": "set_tr_cant_split {tbl_index}：表格行禁止跨页截断",
    "set_tbl_font": "set_tbl_font {tbl_index, font_eastAsia?, font_ascii?, font_size_pt?, skip_header?}：表格字体",
    "set_tbl_header_font": "set_tbl_header_font {tbl_index, font_eastAsia?, font_ascii?, font_size_pt?, bold?}：表头行字体",
    "set_tbl_body_font": "set_tbl_body_font {tbl_index, font_eastAsia?, font_ascii?, font_size_pt?}：表体行字体",
    "set_tbl_borders": "set_tbl_borders {tbl_index, style}：表格边框（three_line/solid/none）",
    "set_tbl_col_widths": "set_tbl_col_widths {tbl_index, widths}：手动设置列宽，widths 为 twips 值列表（如 [2400,3600,1800]）或 \"auto\"",
    "set_tbl_cell_valign": "set_tbl_cell_valign {tbl_index, valign, skip_header?}：单元格垂直对齐，valign='top'/'center'/'bottom'",
    "split_merged_cells": "split_merged_cells {tbl_index}：检测合并单元格（gridSpan/vMerge），返回位置但不拆分（⚠️ 风险大仅警告）",
    "normalize_tbl_grid": "normalize_tbl_grid {tbl_index}：重建 tblGrid+清除 tcW，解决列宽不一致导致的边框不对齐",
    "set_page_margins": "set_page_margins {section_index, top?, bottom?, left?, right?, header?, footer?}：页面边距（twips）",
    "normalize_list_numbering": "normalize_list_numbering {}：统一列表编号层级，确保同级列表段落 ilvl 一致",
    "repair_fields": "repair_fields {}：引用修复——重建断链引用书签（「错误!未找到引用源」根因）+ 恢复参考文献列表编号（引用显示「0」根因）+ 补未闭合域；「引用不正确」时先跑此动作",
    "normalize_cjk_spacing": "normalize_cjk_spacing {}：清除中英混排多余空格",
    "extract_textbox_captions": "extract_textbox_captions {}：文本框内图标题提取为正文段落",
    "create_or_update_style": "create_or_update_style {style_id, name?, type?, properties:{font_east_asia?, font_ascii?, font_size_pt?, bold?, italic?, jc?, ...}}：创建或更新样式（缺失自动创建，同值幂等）",
    "clone_style_from_spec": "clone_style_from_spec {style_id, source_style_id, source_docx_path}：从规范模板克隆样式到本文档",
    "map_style": "map_style {from_style_id, to_style_id}：全文样式映射（所有 from 样式段落改为 to 样式）",
    "set_list_level_format": "set_list_level_format {abstract_num_id, ilvl, num_fmt?, lvl_text?, start?}：修改多级编号某级格式（如 1. 1.1 1.1.1）",
    "map_paragraph_to_list": "map_paragraph_to_list {selector:{semantic_role, occurrence}, num_id, ilvl}：把定位段落挂到指定编号层级",
    "set_header_reference": "set_header_reference {header_type: default/first/even, text?, page_number?, suffix?}：设置页眉（三态，first 自动启用首页不同）",
    "set_footer_reference": "set_footer_reference {footer_type: default/first/even, text?, page_number?, suffix?}：设置页脚（三态，常含页码域）",
    "set_title_page_flag": "set_title_page_flag {section_index?, enabled}：启用/关闭首页不同（titlePg）",
    "set_even_odd_headers_flag": "set_even_odd_headers_flag {enabled}：启用/关闭奇偶页眉脚不同",
    "set_equation_from_latex": "set_equation_from_latex {selector 或 match_text, latex}：把公式段改写为 Word 原生公式（LaTeX 语法，保留原编号）",
    "set_equation_dsmt4": "set_equation_dsmt4 {selector 或 match_text, latex}：把公式段改写为 MathType 公式对象（DSMT4，Word 里双击可编辑，保留原编号）",
    "generate_equation_numbers": "generate_equation_numbers {}：行间公式自动编号（格式从文档已有编号推断、章编号感知；跳过已有编号）",
    "split_inline_equations": "split_inline_equations {para_index}：把粘在正文句子里的公式拆成独立公式段（只搬 run 不改文字）",
    "toggle_tex_equations": "toggle_tex_equations {}：把 $…$ / $$…$$ / \\[…\\] 文本转成 MathType 公式对象（货币等非数学内容自动跳过并如实列出）",
    "convert_text_equations": "convert_text_equations {conversions: [{match_text, latex}], apply}：纯文本公式批量转公式对象（两种形态自动判别：整段式=段落就是公式、整段换对象；句内式=公式粘在正文句子里（如「…时延为 T_n=C_n/f_n。 (4)」）、只把公式跨度原位换对象、编号留在原位——A5）；apply=false 只出提案给用户过目（含命中形态 mode 与跨度预览），用户确认后 apply=true 落盘（E9：识别结果先给用户确认再替换）",
    "set_equation_leader": "set_equation_leader {style, match_text?}：公式与编号之间的连接样式——style=\"dot\"（点线 ……）/\"dash\"（虚线 ----）/\"none\"（去掉）；用户点名要加/去点线时才用（默认口径不加）；match_text 可选，只处理含该文本的公式段",
    "strip_instruction_paragraphs": "strip_instruction_paragraphs {}：源文稿里整段【…】或 *…* 包裹的模板说明段剥除（表格/页眉页脚/文本框/标题豁免；>20 段熔断防误删模板本体；只删不转规格——源不是规范权威）",
    "clear_table_cell_outline": "clear_table_cell_outline {}：表内段落标题残留清障——heading 样式单元格（样式带 outlineLvl/编号链，如 '10 MHz' 数值格）段级写 outlineLvl=9（正文级）覆盖样式级标题，Word 导航/探针不再视为标题；不动样式引用/字号/视觉/numPr；>300 段熔断",
    "normalize_cjk_punctuation": "normalize_cjk_punctuation {}：半角标点→全角中文标点的低风险子集归一（参考文献/英文摘要/类代码段保护区不动；英文语境/小数/公式引用等保守保留；每处变更入对照账本）",
    "set_footer_page_number_font": "set_footer_page_number_font {cn_size 或 size_pt}：设置页脚页码段字号（如 cn_size=\"五号\"；用于格式检测报告的「页面设置-页码问题-字号问题」修复）",
    "renumber_caption_numbers": "renumber_caption_numbers {kinds?, style?, apply}：题注编号按章重排（图 2-1/表 2-1 式；章切换处重新计数，如 图4.12→图4-1 风格归一或补章号）；apply=false 只出对照清单（重排前须先出清单），用户确认后 apply=true 落盘；章前题注/无编号题注不动并如实列出；style=dash(-)/dot(.) 可选（默认从现有编号多数派推断）",
    "generate_caption_toc": "generate_caption_toc {kinds?, heading_texts?, apply}：生成「图目录」「表目录」区（TOC \\c 域+条目缓存，Word 里可更新域；插入主目录之后、首章之前）；无题注时明确报「未发现题注」不生成空目录；已存在同 kind 目录区就地重建（幂等不堆叠）；apply=false 只出预览",
    "generate_cross_references": "generate_cross_references {kinds?, apply}：给题注/公式编号加书签，把正文里「图N-M/表N-M/式(N)」文字引用替换成可联动的 REF 域（Word 里 F9 可更新；编号重排后引用自动跟随）；重号歧义/找不到对应编号的引用只报不改；apply=false 只出待替换清单预览，用户确认后 apply=true 落盘；kinds=figure/table/equation 可选（默认全部）",
    "build_cover_from_template": "build_cover_from_template {template_path, apply, zones?, slots?, replace?}：用规范模板的封面/声明示例区块级克隆到目标文档（字体/logo 原样携带，槽位等宽回填：题目/专业名称/学生姓名/学生学号/指导教师/日期）；template_path=会话模板绝对路径（[模板已就绪] 通告里的值；上传文件名/模板包 id user:<sha12> 也可解析，解析不到会自动换用会话模板并在结果里说明）；zones=[\"cover\"] 或 [\"cover\",\"statement\"]；replace=auto（默认，自动定位目标旧封面块；定位不到不硬做）或 {\"cover\":[start,end]} 显式指定；apply=false 先出预览（含 zones_found 与槽位当前值），用户确认后 apply=true 落盘（本轮消息已明确授权＝已确认——预览通过就直接 apply=true，不再重复征求确认）；同值重跑幂等跳过",
    "link_citation_references": "link_citation_references {apply, superscript?}：把正文 [n] 文献引用做成可联动的交叉引用域（指向文末条目编号；条目编号改动后 Word 里 F9 联动更新）；只对已存在的条目编号生成，找不到条目/组合范围引用（[1,2]/[3-5]）只报不改；apply 必填：false=只出待替换清单预览（不落盘），用户确认后 true=落盘；superscript 可选（默认 true）：引用标记 [n] 整体做成上标（中文论文约定）；要保留正文形态传 false",
    "link_chapter_references": "link_chapter_references {apply}：把正文「第 X 章」提及做成可联动的交叉引用域（指向章标题；章号改动后 Word 里 F9 联动更新、可点击跳转）；只对已存在的「第 N 章」式章标题生成，找不到对应章标题的提及只报不改；目录/文本框/封面区/章标题自身不参与；apply 必填：false=只出待替换清单预览（不落盘），用户确认后 true=落盘",
    "convert_text_table": "convert_text_table {apply, match?, cells?, columns?, header?, template_path?}：把「表 N-M」题注下的纯文本数据块转成 Word 表格；apply=false 先出候选数据块与行列方案预览（未落盘）；apply=true 必须先与用户确认行列方案（match=题注 + cells 或 columns 二选一），禁止先斩后奏；表边框优先取模板首表（template_path），否则三线表默认",
}

_RISK_MAP: dict[str, RiskLevel] = {}
for _a in R1_ACTIONS:
    _RISK_MAP[_a] = RiskLevel.R1
for _a in R2_ACTIONS:
    _RISK_MAP[_a] = RiskLevel.R2
for _a in R3_ACTIONS:
    _RISK_MAP[_a] = RiskLevel.R3
for _a in NEW_ACTIONS:
    _RISK_MAP[_a] = NEW_ACTIONS[_a][0]
for _a in STYLE_ACTIONS:
    _RISK_MAP[_a] = STYLE_ACTIONS[_a][0]
for _a in LIST_ACTIONS:
    _RISK_MAP[_a] = LIST_ACTIONS[_a][0]
for _a in HEADER_FOOTER_ACTIONS:
    _RISK_MAP[_a] = HEADER_FOOTER_ACTIONS[_a][0]


# ── 适配器 ──────────────────────────────────────────────────────

def _wrap_handler(fn) -> callable:
    def handler(path, params: dict) -> dict:
        fn(Path(path), params)
        return {"ok": True}
    return handler


def _make_get_xml(docx_path: Path):
    import zipfile
    from lxml import etree

    cache: dict[str, object] = {}

    def get_xml(internal_path: str):
        if internal_path not in cache:
            with zipfile.ZipFile(docx_path, "r") as z:
                cache[internal_path] = etree.fromstring(z.read(internal_path))
        return cache[internal_path]

    return get_xml


def _wrap_verifier(fn) -> callable:
    """包装 (target_path, params, get_xml, idx) -> bool。

    fixer 的 verifier 返回 FixerDiagnostic 列表；存在 severity=error 即失败。
    """

    def verifier(path, params: dict) -> bool:
        try:
            diags = fn(Path(path), params, _make_get_xml(Path(path)), 0) or []
        except Exception:
            return False
        hard = [d for d in diags if getattr(d, "severity", "warning") == "error"]
        return not hard

    return verifier


# ── 目录注册 ────────────────────────────────────────────────────

def build_catalog_registry() -> ActionRegistry:
    """从 writer/fixer 实际注册表构建 3.2 ActionRegistry（唯一真相源）。"""
    reg = ActionRegistry()
    uncovered: list[str] = []

    for name, handler_fn in sorted(writer._ACTION_HANDLERS.items()):
        risk = _RISK_MAP.get(name)
        if risk is None:
            uncovered.append(name)
            continue
        verifier_fn = fixer._VERIFIERS.get(name)
        spec = ActionSpec(
            name=name,
            version=1,
            category=_category_of(name),
            scope="document",
            risk=risk,
            exposed_to_agent=name in AGENT_EXPOSED,
            idempotent=True,
            verifier=name if verifier_fn is not None else None,
            rollback_strategy="restore_part",
            content_change=name in CONTENT_CHANGE_ACTIONS,
            params_schema=PARAMS_SCHEMAS.get(name),
            description=ACTION_DESCRIPTIONS.get(name, ""),
        )
        reg.register(
            spec,
            handler=_wrap_handler(handler_fn),
            verifier=_wrap_verifier(verifier_fn) if verifier_fn is not None else None,
        )
    # 3.2 新 action 注册
    for name, (risk, handler_fn, verifier_fn, category) in sorted(NEW_ACTIONS.items()):
        spec = ActionSpec(
            name=name,
            version=1,
            category=category,
            scope="document",
            risk=risk,
            exposed_to_agent=name in AGENT_EXPOSED,
            idempotent=True,
            verifier=name,
            rollback_strategy="restore_part",
            content_change=name in CONTENT_CHANGE_ACTIONS,
            params_schema=PARAMS_SCHEMAS.get(name),
            description=ACTION_DESCRIPTIONS.get(name, ""),
        )
        reg.register(spec, handler=handler_fn, verifier=verifier_fn)
    # Phase 2 样式系统注册（确定性链路专用，不暴露给 Agent）
    for name, (risk, handler_fn, verifier_fn, category) in sorted(STYLE_ACTIONS.items()):
        spec = ActionSpec(
            name=name,
            version=1,
            category=category,
            scope="document",
            risk=risk,
            exposed_to_agent=name in AGENT_EXPOSED,
            idempotent=True,
            verifier=name,
            rollback_strategy="restore_part",
            description=ACTION_DESCRIPTIONS.get(name, ""),
        )
        reg.register(spec, handler=handler_fn, verifier=verifier_fn)
    # Phase 2 多级编号系统注册
    for name, (risk, handler_fn, verifier_fn, category) in sorted(LIST_ACTIONS.items()):
        spec = ActionSpec(
            name=name,
            version=1,
            category=category,
            scope="document",
            risk=risk,
            exposed_to_agent=name in AGENT_EXPOSED,
            idempotent=True,
            verifier=name,
            rollback_strategy="restore_part",
            description=ACTION_DESCRIPTIONS.get(name, ""),
        )
        reg.register(spec, handler=handler_fn, verifier=verifier_fn)
    # Phase 2 页眉页脚注册
    for name, (risk, handler_fn, verifier_fn, category) in sorted(HEADER_FOOTER_ACTIONS.items()):
        spec = ActionSpec(
            name=name,
            version=1,
            category=category,
            scope="document",
            risk=risk,
            exposed_to_agent=name in AGENT_EXPOSED,
            idempotent=True,
            verifier=name,
            rollback_strategy="restore_part",
            description=ACTION_DESCRIPTIONS.get(name, ""),
        )
        reg.register(spec, handler=handler_fn, verifier=verifier_fn)
    if uncovered:
        raise ValueError(
            f"actions without risk classification (update catalog): {uncovered}"
        )
    return reg


def _category_of(name: str) -> str:
    if name.startswith("set_tbl") or name.startswith("normalize_tbl"):
        return "table"
    if name.startswith("set_tbl_") and name in ("set_tbl_borders", "set_tbl_col_widths", "set_tbl_header_repeat"):
        return "table"
    if "header" in name or "footer" in name or "page_number" in name or "sectpr" in name:
        return "section"
    if "toc" in name:
        return "toc"
    if "image" in name or "ole" in name:
        return "image"
    if "list" in name or "numbering" in name:
        return "list"
    return "paragraph"


def catalog_consistency_report() -> dict:
    """Phase 1 验收指标：handler/verifier/schema 差集为 0。"""
    reg = build_catalog_registry()
    rep = reg.consistency_report()
    return {
        "total": len(reg.all()),
        "exposed": len(reg.exposed()),
        "consistent": rep.consistent,
        "missing_handler": rep.missing_handler,
        "missing_verifier": rep.missing_verifier,
    }


# ── Agent 工具描述自动生成（ADR-006 完整落地）────────────────────
# agent_loop 的 TOOLS 不再手写 action 清单；description/schema 全部
# 由 ActionRegistry（唯一真相源）生成，消除静态 Prompt 漂移。

_FIXED_RULES = [
    "图文排版标准流程：① set_images_inline → ② extract_textbox_captions → ③ set_paragraph_alignment + protect_image_para。",
    "务必先 inspect_document 拿到真实数据再调用。",
]


def build_apply_operations_description(reg: ActionRegistry | None = None) -> str:
    """从 registry 自动生成 apply_operations 的中文 description。

    每个 exposed action 一行（来自 ACTION_DESCRIPTIONS），缺失说明的
    action 自动生成占位行（防止静默遗漏）。
    """
    reg = reg or build_catalog_registry()
    lines = ["描述: 在当前文档上执行一组 OOXML 精细排版操作（无需规范模板）。常用 action："]
    for spec in reg.exposed():
        desc = spec.description or (
            f"- {spec.name} {{...}}：{spec.name}（参数说明待补充——联系开发补全 ACTION_DESCRIPTIONS）"
        )
        lines.append(f"- {desc}")
    lines.extend(_FIXED_RULES)
    return "\n".join(lines)


def build_apply_operations_schema(reg: ActionRegistry | None = None) -> dict:
    """从 registry 生成 apply_operations 的 parameters JSON Schema。

    保持旧格式兼容：{operations: [{action, params}]}。
    """
    reg = reg or build_catalog_registry()
    any_of = []
    for spec in reg.exposed():
        props = {"action": {"type": "string", "enum": [spec.name]}}
        params = spec.params_schema or {"type": "object", "properties": {}}
        props["params"] = params
        any_of.append({"type": "object", "properties": props, "required": ["action", "params"]})
    return {
        "type": "object",
        "properties": {
            "operations": {
                "type": "array",
                "description": "操作列表，每项 {action, params}",
                "items": {"anyOf": any_of},
            },
        },
        "required": ["operations"],
    }
