"""早睡 (EarlyZ) — enforce/ 包

奇偶页强制（按章节标题选择哪些节必须从奇数页开始）。

2026-11-15（EZI-0003）：此前实现只存在于仓库外的归档目录（office 侧那份副本），
``cli_new.py`` 的 ``enforce`` 子命令与 ``api/agent_loop.py`` 的 ``enforce_odd_pages_tool``
都靠 ``sys.path`` 注入去 import 它——换机器 clone 必然 ModuleNotFoundError。
本包把这段最小实现（config / pdf_reader / odd_pages，共约 260 行）迁进仓库，
PDF 读取改用仓库内现役 ``earlyz.validation.pdf_backend``（PDFium，非 AGPL），
PDF 转换改走 LibreOffice（安装包本来就不带归档，历史链路实际只走到这条兜底）。
"""

from .config import EnforceConfig
from .odd_pages import enforce_odd_pages
from .pdf_reader import get_section_start_pages

__all__ = [
    "EnforceConfig",
    "enforce_odd_pages",
    "get_section_start_pages",
]
