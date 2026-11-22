"""OLE 对象写出器：MTEF Equation → Equation.DSMT4 OLE 对象（CFB 复合文档）。

Word 中的 MathType 公式是 OLE 嵌入对象，结构：
- CLSID: {0002CE03-0000-0000-C000-000000000046}
- 4 个流：
  1. \\x01CompObj（28 字节）：版本标记 + 长度前缀字符串（"Equation.DSMT4" 等）
  2. \\x01Ole（20 字节）：OLE 版本标记
  3. \\x03ObjInfo（6 字节）：对象标志
  4. Equation Native（变长）：28 字节 MathType 头 + MTEF v5 二进制

产物：CFB 二进制（512 字节扇区，单层存储，所有流 < 4096 字节走 mini stream）。
"""
from __future__ import annotations

import struct
from .mtef import Equation, write_equation
from .render import equation_to_wmf
from .cfb import write_streams

# CLSID {0002CE03-0000-0000-C000-000000000046}
DSMT4_CLSID = bytes.fromhex("03CE0200 00000000 C0000000 00000046".replace(" ", ""))


def equation_to_ole(eq: Equation, preview_wmf: bytes | None = None) -> bytes:
    """MTEF Equation → Equation.DSMT4 OLE 对象（CFB 二进制）。
    
    Args:
        eq: MTEF Equation 对象
        preview_wmf: WMF 预览图（可选）。不传则自动生成。
    
    Returns:
        CFB 二进制（可直接嵌入 Word 的 <w:object>）
    """
    if preview_wmf is None:
        preview_wmf = equation_to_wmf(eq)
    
    # 构造 4 个流（按 Word 的顺序）
    streams = [
        ('\x01CompObj', _build_compobj()),
        ('\x01Ole', _build_ole()),
        ('\x03ObjInfo', _build_objinfo()),
        ('Equation Native', _build_native(eq)),
    ]
    
    # 用 cfb 模块写出
    return write_streams(streams, DSMT4_CLSID)


def _build_compobj() -> bytes:
    """\\x01CompObj 流：OLE 复合文档的组件信息（Word 靠它找公式服务器）。

    逐字节对齐真实 MathType 对象（105 字节，电工技术学报模板实测）：
      u32 0xFFFE0001           版本
      u32 0x00000A03           flags
      u32 0xFFFFFFFF           reserved
      CLSID 16 字节
      u32 + ANSI 串（含结尾 NUL）：user type "MathType 6.0 Equation"
      u32 + ANSI 串（含结尾 NUL）：clipboard format "MathType EF"
      u32 + ANSI 串（含结尾 NUL）：progID "Equation.DSMT4"
      u32 0xB271F439           标记
      u32 ×3 = 0

    早先写成 UTF-16 字符串 + 字节序反了的版本号，且缺 clipboard format，
    Word 找不到公式服务器 → 打开文档闪退。
    """
    def _ansi(s: str) -> bytes:
        raw = s.encode("ascii") + b"\x00"
        return struct.pack("<I", len(raw)) + raw

    out = struct.pack("<III", 0xFFFE0001, 0x00000A03, 0xFFFFFFFF)
    out += DSMT4_CLSID
    out += _ansi("MathType 6.0 Equation")
    out += _ansi("MathType EF")
    out += _ansi("Equation.DSMT4")
    out += bytes.fromhex("f439b271")   # 按真实字节序原样写，不能当 u32 打包
    out += b"\x00" * 12
    assert len(out) == 105, len(out)
    return out


def _build_ole() -> bytes:
    """\\x01Ole 流：20 字节。真实对象为 01 00 00 02 08 00 00 00 + 12 字节 0。"""
    return bytes.fromhex("0100000208000000") + b"\x00" * 12


def _build_objinfo() -> bytes:
    """\\x03ObjInfo 流：6 字节。真实对象为 00 00 03 00 01 00（嵌入 + 已初始化）。"""
    return bytes.fromhex("000003000100")


def _build_native(eq: Equation) -> bytes:
    """Equation Native 流：28 字节 MathType 头 + MTEF v5。

    头格式以真实 MathType 对象实测为准（电工技术学报模板 43 个对象逐个核对）：
      +0  cbHdr     (2) = 28
      +2  version   (2) = 0
      +4  cf        (2) = 2          ← 固定值，不是 0
      +6  reserved  (2)
      +8  cbObject  (4) = MTEF 字节数 ← **在偏移 8，不是 6**
      +12 reserved  (16)

    早先写成 `HHHI`（cbObject 落在偏移 6、cf=0、version=2），Word/EQNEDT32 读到
    的长度字段错位，打开含公式的文档时会直接闪退。
    """
    mtef_bin = write_equation(eq)
    header = struct.pack('<HHHHI', 28, 0, 2, 0, len(mtef_bin)) + b'\x00' * 16
    assert len(header) == 28
    return header + mtef_bin
