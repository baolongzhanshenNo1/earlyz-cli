"""EarlyZ 工具注册。"""

from __future__ import annotations
from earlyz.tools.base import EarlyZTool

_registry: dict[str, EarlyZTool] = {}

def register(tool: EarlyZTool) -> None:
    _registry[tool.tool_name] = tool

def get_tool(name: str) -> EarlyZTool | None:
    return _registry.get(name)

def list_tools() -> list[EarlyZTool]:
    return list(_registry.values())

# ── 注册 5 个内置工具 ──
from earlyz.tools.thesis import ThesisTool         # noqa: E402
from earlyz.tools.contract import ContractTool     # noqa: E402
from earlyz.tools.gov_doc import GovDocTool        # noqa: E402
from earlyz.tools.translate import TranslateTool   # noqa: E402
from earlyz.tools.tech_doc import TechDocTool      # noqa: E402

register(ThesisTool())
register(ContractTool())
register(GovDocTool())
register(TranslateTool())
register(TechDocTool())
