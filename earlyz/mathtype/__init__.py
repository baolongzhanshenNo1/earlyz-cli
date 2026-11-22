"""EarlyZ 公式引擎（无界面 MathType，自研）。

- mtef：MTEF v5 记录模型 + 二进制读写（Design Science 公开规格）
- cfb / ole：OLE 复合文档容器 + Equation.DSMT4 对象（\\x01CompObj / \\x01Ole / \\x03ObjInfo / Equation Native）
- latex_math / from_latex：LaTeX 数学 → MTEF（E6）
- render：MTEF → 矢量预览图（E3）
产物与 MathType 兼容：Word 显示预览图，双击仍可用 MathType 打开编辑。全部 stdlib，不依赖 MathType。
"""

from .from_latex import latex_to_mtef as latex_to_equation
from .mtef import parse, serialize, Equation
from .cfb import read_streams, write_streams
from .render import equation_to_wmf
from .ole import equation_to_ole

__all__ = [
    "latex_to_equation",
    "parse",
    "serialize",
    "Equation",
    "read_streams",
    "write_streams",
    "equation_to_wmf",
    "equation_to_ole",
]
