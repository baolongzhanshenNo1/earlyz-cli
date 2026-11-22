# earlyz-cli

> **产品引擎面的开源版，供参考，不是产品本体。** EarlyZ 桌面版的核心排版引擎从主仓同步而来；评测语料、判定器基建与 agent 编排不在仓里。当作研究代码对待：可能有 bug，用前自行审阅调整。

earlyz-cli 是规范模板驱动的 Word 排版引擎，命令行形态。给它一份格式乱的文档和一份规范模板（都是 docx），它把两者的格式差异找出来修掉——只修格式，不动文字内容。三个命令：`format` 按规范排版，`analyze` 把规范模板提取成 YAML 规则，`enforce` 做奇偶页强制（需 LibreOffice）。所有操作在 OOXML ZIP 层直接读写，排版本身不依赖 Word 或 LibreOffice 运行时。

它的设计前提是**模板才是排版的规范权威**：源文档只提供排什么，不定义怎么排。字号线距页边距这些取值必须出自模板的 OOXML，不猜、不套学校常量；PDF 渲染走 PDFium（pypdfium2，Apache-2.0/BSD-3），无 AGPL 依赖。设计取舍与能力边界见 [docs/000-philosophy.md](docs/000-philosophy.md)。

## 接入

```bash
git clone https://github.com/baolongzhanshenNo1/earlyz-cli
pip install -e .
python examples/make_examples.py                 # 自制合成示例对（不含任何真实院校/期刊模板）
earlyz format examples/示例论文.docx --spec examples/示例模板.docx
python -m pytest -q tests/                        # 140 条引擎行为守卫（合成夹具）
```

环境：Python ≥ 3.11；LibreOffice 仅 `enforce` 命令需要。

## 引擎现状

排版正确率以双语料库可判项计：清洗集 579/579、留出集 388/388，双双 100%（2026-11）。开发走 issue 驱动工作流（[workflow-kit](https://github.com/baolongzhanshenNo1/workflow-kit)）：每张卡过五层门禁、证据绑定提交指纹、摆回即红。本仓引擎与桌面版主仓同源同步（`scripts/sync_cli_repo.py` 切面可复跑）。

与桌面版的关系：桌面版是闭源商业产品（[EarlyZ](https://baolongzhanshen.cn)），在此引擎之上另有公式（自研 MathType 对标组件）、LaTeX、事务回滚与 LLM Agent 闭环。

## License

[MIT](LICENSE)
