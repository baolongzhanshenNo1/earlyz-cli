# ADR-021 · TSD 模板语义反编译器（Template Semantic Decompiler）

> **状态：** Accepted / 已实现（2026-08-15）
> **对应 3.2 定稿：** §5.4（规则来源优先级）、§12/§17（Phase 1 工作包增补）
> **取代：** `spec/compilers.py::_compile_docx` 旧实现（只读 styles.xml）
> **代码：** `earlyz/spec/decompiler.py`（本目录）

## 0. 背景与根因（2026-08-15 全量审计）

真实高校模板（西安理工高科学院等）是**规范示例式**模板：规范不在
styles.xml（只有 9 个 header/footer 辅助样式），而在 document.xml
示例段落的**直接格式**里（黑体18pt居中章标题、宋体小四正文、逐段
封面布局），并伴随「（黑体小2加粗）」注释段。

旧 `_compile_docx` 只读 styles.xml → 产出 7 条样式级规则 →
741 段落几乎零命中 → lint 只报样式违规 → plan 全是
`create_or_update_style` → **排版能力为零**（用户看到"编译成功但
零修改"）。631 绿测试全过——因为验收指标只衡量"文档健康"，
不衡量"格式对齐"。

## 1. 设计原则

1. **单一真相源**：模板与目标文档用同一个 `analyze_document` 分类器
   （角色天然对齐，不造第二套分类）。
2. **确定性、零 LLM**：聚类 + 主导值 + 文本形状正则，全程可复现。
3. **克制原则**：宁可少排，不可错排——高门槛主导值（0.6/0.75）、
   位置伪影键永不产出、单样本噪声防护。
4. **属性键与下游严格对齐**：expected 键名 = lint `_actual_prop` =
   writer `_PPR_MAP/_RPR_MAP`（spec 惯例别名 `first_line_indent_chars`
   在 writer 侧归一化）。

## 2. 三类规则（+docDefaults 基线）

| 规则 | selector | 样本源 | 执行动作 |
|---|---|---|---|
| **语义角色规则** | `{semantic_role: role}` | 模板段落直接格式（按角色聚类，标题形状检测补分类）＋styles.xml 样式定义兜底 | `set_paragraph_format` 批量（幂等 no-op） |
| **节级页面几何规则** | `{scope: section_all}` | 模板各节 pgMar（全一致才出） | `set_page_margins_all`（单动作全节） |
| **封面锚点规则** | `{cover_anchor, match}` | 模板封面区段落（exact/prefix/date/title 四种匹配） | `set_paragraph_format`（锚点选择器） |
| docDefaults 基线 | `{scope: document_default}` | styles.xml docDefaults | `set_doc_defaults` |

### 属性键策略

- `_ROLE_KEYS`：每角色白名单（body 7 键 / heading 核心键 / caption…）。
- `_POSITIONAL_KEYS = {page_break_before}`：模板示例段的分页/锚定属性
  是**位置伪影**（示例章标题恰好落在分页点），不是角色规范——永不产出
  （否则摘要/目录标题被强加分页）。
- `_STRICT_KEYS`（行距/字体/段前后距）：矛盾高发，主导值门槛 0.75。

### 封面区边界（关键不变量）

封面区 = 首个**结构性锚点**（章节序号行 / heading 角色 / 固定锚点词
如"郑重声明"）之前的段落；**找不到锚点 = 无封面区（返回 0）**。

- `find_cover_end(ir)`（decompiler，lint 用）与
  `_compute_cover_boundary(body)`（new_actions，writer 用）语义对齐。
- lint 的 role 通道**跳过封面区**（封面段归 cover_anchor 规则管——
  防止 body 主导值把封面大标题刷成首行缩进/两端对齐）。
- title/date 匹配**只认封面区内段落**（防止正文 12 字中文行被误判为
  封面标题刷成黑体 26pt——2026-08-15 真实回归）。

## 3. 下游接线（本 ADR 引入的变更）

| 层 | 变更 |
|---|---|
| `spec/compilers.py` | `_compile_docx` → `decompile_template`；死代码清除 |
| `lint/engine.py` | 新增 `scope:section_all`（逐节 margins_diff 容差3twip）与 `cover_anchor` 两通道；role 通道加封面豁免 |
| `planner/planner.py` | role/cover/scope 三类规则按 rule_id 合并为批量动作；新增 `set_page_margins_all` 映射 |
| `actions/new_actions.py` | `_locate_paragraphs` 支持 cover_anchor 选择器；role 批量跳过封面区；新增 `set_page_margins_all` + verifier |
| `actions/catalog.py` | 注册 `set_page_margins_all`（R1，section）→ 54 action |

## 4. Phase 1 验收指标增补（堵"测试全绿但能力为零"）

**上传模板编译（uploaded_template_docx）验收**：

1. 真实示例式模板 → 规则 ≥ 4 类 semantic_role（body/heading1/
   heading2/caption）；
2. 真实模板 → 真实论文的 plan **必须含段落级动作 > 0**
   （`set_paragraph_format`）；
3. 产物段落规则命中率 ≥ 90%（lint 二轮违规为 0 = 幂等通过）；
4. `lint_total=0 且动作全 no-op` 时，报告必须显式说明
   「模板规范未命中任何段落」而非 passed=True。

实测（西安理工模板 + 马炳翔论文）：lint 645 违规（605 段落级 +
6 节级 + 34 封面）→ 11 动作 applied → 二轮幂等 → Word 渲染
PDF/PNG 视觉验证通过（封面居中层级清晰、章标题黑体18pt居中、
正文缩进2字符两端对齐、页眉页脚引用全保留）。

## 5. 与 §19 Template Pack / §20 Host 分层的关系

- TSD 是 §19 的**前置条件**：Template Pack 的「上传 docx → 提取规则
  → 存成 pack」入口依赖的正是本提取器；§19 资产化应在 TSD 验收后实施。
- §20/B 批与 TSD 正交，可并行。

## 6. 已知边界（后续增量）

- 封面**内容**（各字段值）仍属用户文档，TSD 只克隆布局属性。
- 页眉页脚变体规则（headerRefs/footerRefs 内容级）未提取——
  当前链路保留目标文档页眉页脚引用（23 refs 全保），不破坏。
- 图表目录/参考文献条目角色（caption 已支持，参考文献条目格式未单列）。
