"""LaTeX 工程层（场景 A/B 共用入口）：工程解包、主文件识别、编码、cls 定位、缺失清单、前置字段；
L1 TeX 前端（tokens / macros / expand）+ L2 文档 IR（ir / parse_* / numbering / document，入口 parse_project）。

全部自研、stdlib only（不用 pandoc）。TeX 发行版只在用户本机已安装时调用（kpsewhich 查 cls），不打包。
"""
from earlyz.latex.document import parse_project
from earlyz.latex.project import LatexProject, MainCandidate, open_project, select_main

__all__ = ["LatexProject", "MainCandidate", "open_project", "parse_project", "select_main"]
