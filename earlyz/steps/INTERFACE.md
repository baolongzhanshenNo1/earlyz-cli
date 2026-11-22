# 早睡 (EarlyZ) Pipeline Steps — 可组合的排版步骤

## 设计原则

- 每个 Step 是高内聚的纯函数：读文档 → 生成 actions
- Step 之间零耦合：不 import 彼此
- Writer 不管 Step 来源：只执行 fix_plan
- 加新领域 = 组合现有 Base Steps + 写专属 Steps

## 接口

```python
class PipelineStep:
    """一个排版步骤。返回 actions 列表，追加到 fix_plan。"""
    def run(self, target: Path, spec: dict) -> list[dict]: ...

class Reconciler:
    """按顺序执行 Steps，合并成 fix_plan。"""
    def __init__(self, steps: list[PipelineStep]): ...
    def reconcile(self, target, spec) -> list[dict]: ...
```

## Base Steps（所有领域共享）

```
                    ┌─────────────────────┐
                    │     PipelineStep     │
                    └──────────┬──────────┘
                               │
        ┌──────────────────────┼──────────────────────┐
        │                      │                      │
        ▼                      ▼                      ▼
┌───────────────┐    ┌───────────────┐    ┌───────────────┐
│  StyleStep    │    │ SectionBreak  │    │ PageNumber    │
│               │    │    Step       │    │    Step       │
│ 从 spec 读    │    │               │    │               │
│ H1/H2/Body    │    │ 所有 sectPr   │    │ 格式+对齐     │
│ 字体+字号     │    │ → nextPage    │    │ 从 spec 读取   │
└───────────────┘    └───────────────┘    └───────────────┘

┌───────────────┐    ┌───────────────┐    ┌───────────────┐
│ HeaderFont    │    │ HeaderText    │    │ FooterLink    │
│    Step       │    │    Step       │    │    Step       │
│               │    │               │    │               │
│ header XML    │    │ 每节独立 header│   │ sectPr 加      │
│ 字体设置      │    │ 文字标准化     │    │ footerReference│
└───────────────┘    └───────────────┘    └───────────────┘
```

## 论文专属 Steps（thesis/）

```
┌───────────────┐    ┌───────────────┐    ┌───────────────┐
│ FrontMatter   │    │ RomanArabic   │    │ OddPageEnforce│
│   ClearStep   │    │    Step       │    │    Step       │
│               │    │               │    │               │
│ 封面/声明      │    │ 前部→罗马     │    │ 偶数页起始      │
│ 清除页眉      │    │ 正文→阿拉伯    │    │ 前插空白页      │
└───────────────┘    └───────────────┘    └───────────────┘
```

## 公文专属 Steps（gov_doc/）

```
┌───────────────┐
│ CenteredPage  │    （公文不需要 FrontMatterClear/RomanArabic/OddPage）
│   NumStep     │
│               │
│ "— {PAGE} —" │
│ 居中对齐      │
└───────────────┘
```

## 技术文档专属 Steps（tech_doc/）

```
┌───────────────┐    ┌───────────────┐
│  CodeBlock    │    │   TocStep     │
│    Step       │    │               │
│               │    │ TOC 域插入    │
│ 代码块检测    │    │ 自动目录      │
│ 背景色设置    │    │               │
└───────────────┘    └───────────────┘
```

## 组合示例

```python
# 论文
thesis = Reconciler([
    StyleStep(),
    SectionBreakStep(),
    HeaderFontStep(),
    HeaderTextStep(h1_pattern=THESIS_H1),
    FooterLinkStep(),
    PageNumberStep(),
    FrontMatterClearStep(),
    RomanArabicStep(),
    # OddPageEnforce 在 Step 6 单独跑
])

# 公文
gov_doc = Reconciler([
    StyleStep(),
    SectionBreakStep(),
    HeaderFontStep(),
    HeaderTextStep(h1_pattern=GOV_H1, front_skip=set()),
    FooterLinkStep(),
    CenteredPageNumStep(),
])

# 技术文档
tech_doc = Reconciler([
    StyleStep(),
    SectionBreakStep(),
    HeaderFontStep(),
    HeaderTextStep(h1_pattern=TECH_H1, normalizer=noop),
    FooterLinkStep(),
    PageNumberStep(),
    CodeBlockStep(),
    TocStep(),
])
```

## 文件结构

```
earlyz/
├── engine/              # FormatEngine（不变）
├── tools/               # 用户入口
│   ├── thesis.py        # → thesis_reconciler.reconcile()
│   ├── gov_doc.py       # → gov_doc_reconciler.reconcile()
│   └── tech_doc.py      # → tech_doc_reconciler.reconcile()
└── steps/               # 可组合步骤 ✨
    ├── __init__.py
    ├── base.py           # PipelineStep + Reconciler + Base Steps
    ├── thesis.py         # 论文专属 Steps
    ├── gov_doc.py        # 公文专属 Steps
    └── tech_doc.py       # 技术文档专属 Steps
```
