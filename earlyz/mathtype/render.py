"""MTEF → WMF 矢量预览图（Windows Metafile，16 位格式）。

Word 嵌入的 DSMT4 对象需要预览图：v:shape 里的 imagedata 指向 WMF 流。WMF 是简单的
绘图命令序列（MoveTo, LineTo, TextOut…），我们用 stdlib 构造最小可显示的占位图：

1. **简化策略**：不做真正的排版和渲染，生成一个带公式文本的占位矩形
   - 遍历 MTEF 树，提取可见字符（CHAR 记录的 fontpos 或 mtcode）
   - 用 TextOut 绘制提取的文本
   - 用 Rectangle 画边框
   - 尺寸固定（如 1000×400 逻辑单位）

2. **WMF 结构**（16 位格式，[MS-WMF]）：
   - Placeable header（22 字节）：key=0x9AC6CDD7，bbox，inch（DPI）
   - 标准 WMF header：type=1（内存），header_size=9，version=0x0300，size，num_objects，max_record
   - META_SETMAPMODE, META_SETWINDOWORG, META_SETWINDOWEXT
   - META_CREATEPENINDIRECT, META_CREATEBRUSHINDIRECT, META_SELECTOBJECT
   - META_TEXTOUT / META_RECTANGLE
   - META_DELETEOBJECT（清理笔刷）
   - META_EOF

3. **未来增强**（E3 优先级）：
   - 真正的数学排版（分数、根号、上下标的几何布局）
   - 字体映射（Symbol, Times New Roman, MT Extra）
   - 精确的边界框计算

当前实现：占位图，显示 "[Math: ...]" 加提取的字符。
"""
from __future__ import annotations

import struct
from .mtef import Equation, Line, Char, Tmpl, Pile, Matrix


def equation_to_wmf(eq: Equation) -> bytes:
    """MTEF Equation → WMF 二进制。
    
    返回带 Placeable header 的完整 WMF（22 字节 + 标准 WMF）。
    """
    # 提取可见文本
    text = _extract_text(eq)
    if len(text) > 60:
        text = text[:57] + '...'
    display = f"[Math: {text}]" if text else "[Math]"
    
    # 固定尺寸（逻辑单位）
    width, height = 2000, 600
    
    # 构造 WMF 记录
    records = []
    
    # SetMapMode(MM_ANISOTROPIC=8)
    records.append(_meta_record(0x0103, struct.pack('<H', 8)))
    
    # SetWindowOrg(0, 0)
    records.append(_meta_record(0x020B, struct.pack('<hh', 0, 0)))
    
    # SetWindowExt(width, height)
    records.append(_meta_record(0x020C, struct.pack('<hh', height, width)))
    
    # CreatePenIndirect(PS_SOLID=0, width=1, color=0x000000)
    records.append(_meta_record(0x02FA, struct.pack('<HhI', 0, 1, 0x000000)))
    
    # CreateBrushIndirect(BS_NULL=1, color=0, hatch=0)
    records.append(_meta_record(0x02FC, struct.pack('<HHH', 1, 0, 0)))
    
    # SelectObject(pen, index=0)
    records.append(_meta_record(0x012D, struct.pack('<H', 0)))
    
    # SelectObject(brush, index=1)
    records.append(_meta_record(0x012D, struct.pack('<H', 1)))
    
    # Rectangle(left=10, top=10, right=width-10, bottom=height-10)
    records.append(_meta_record(0x041B, struct.pack('<hhhh', height - 10, width - 10, 10, 10)))
    
    # TextOut(x=50, y=height/2, text)
    # META_TEXTOUT: size, y, x, count, text(偶数长度), x, y（重复）
    text_bytes = display.encode('latin-1', errors='replace')
    text_len = len(text_bytes)
    if text_len % 2 == 1:
        text_bytes += b'\x00'
    x, y = 50, height // 2
    rec_data = struct.pack('<H', text_len) + text_bytes + struct.pack('<hh', x, y)
    records.append(_meta_record(0x0521, rec_data))
    
    # DeleteObject(pen)
    records.append(_meta_record(0x01F0, struct.pack('<H', 0)))
    
    # DeleteObject(brush)
    records.append(_meta_record(0x01F0, struct.pack('<H', 1)))
    
    # EOF
    records.append(_meta_record(0x0000, b''))
    
    # 计算总大小（16 位 word 为单位）
    total_size = 9  # header size
    for rec in records:
        total_size += len(rec) // 2
    
    # 标准 WMF header（18 字节；2026-10-06 修复：此前 14 字节漏 <HHHIHH 的后半）
    wmf_header = struct.pack(
        '<HHHIHIH',
        1,              # type: memory metafile
        9,              # header size (words)
        0x0300,         # version: Windows 3.0
        total_size,     # file size (words)
        2,              # num_objects
        50,             # max_record (words, 估算)
        0,              # no parameters
    )
    
    wmf_body = wmf_header + b''.join(records)
    
    # Placeable header（22 字节）
    # bbox 单位是 twips（1/1440 inch），假设 96 DPI
    inch = 1440
    left, top, right, bottom = 0, 0, width * 15, height * 15
    checksum = 0x9AC6CDD7 ^ left ^ top ^ right ^ bottom ^ inch
    placeable = struct.pack(
        '<IHhhhhHIH',
        0x9AC6CDD7,     # key (DWORD)
        0,              # handle (WORD)
        left, top, right, bottom,  # bbox (4×SHORT)
        inch,           # units per inch (WORD)
        0,              # reserved (DWORD)
        checksum & 0xFFFF  # checksum (WORD)
    )
    
    return placeable + wmf_body


def bitmap_to_wmf(png_bytes: bytes) -> tuple[bytes, int, int]:
    """PNG 位图 → 含 StretchDIB 的 Placeable WMF（2026-10-06）。

    比手写矢量记录可靠得多：只有 SetMapMode/Window/Viewport + 一条 META_STRETCHDIB
    + EOF，像素直接内嵌 DIB。Word/.NET GDI+ 都能渲染（旧实现记录畸形：笔刷/文字
    记录参数错位 → 红叉）。
    """
    import io

    from PIL import Image

    img = Image.open(io.BytesIO(png_bytes)).convert("RGB")
    w, h = img.size
    px = img.load()
    row = ((w * 3 + 3) // 4) * 4           # DIB 行 4 字节对齐
    bits = bytearray()
    for y in range(h - 1, -1, -1):         # DIB 自下而上
        for x in range(w):
            r, g, b = px[x, y]
            bits += bytes((b, g, r))       # BGR
        bits += b"\x00" * (row - w * 3)
    bih = struct.pack("<IiiHHIIiiII", 40, w, h, 1, 24, 0, len(bits), 0, 0, 0, 0)
    dib = bih + bytes(bits)

    records = []
    records.append(_meta_record(0x0103, struct.pack("<H", 8)))       # SETMAPMODE ANISOTROPIC
    records.append(_meta_record(0x020B, struct.pack("<hh", 0, 0)))   # SETWINDOWORG
    records.append(_meta_record(0x020C, struct.pack("<hh", h, w)))   # SETWINDOWEXT (y, x)
    records.append(_meta_record(0x0209, struct.pack("<hh", 0, 0)))   # SETVIEWPORTORG
    records.append(_meta_record(0x020A, struct.pack("<hh", h, w)))   # SETVIEWPORTEXT (y, x)
    # META_STRETCHDIB：ROP + Usage + SrcH/SrcW/SrcY/SrcX + DstH/DstW/DstY/DstX + DIB
    records.append(_meta_record(
        0x0F43,
        struct.pack("<IHhhhhhhhh", 0x00CC0020, 0, h, w, 0, 0, h, w, 0, 0) + dib))
    records.append(_meta_record(0x0000, b""))                        # EOF

    total_size = 9
    max_rec = 0
    for rec in records:
        _words = len(rec) // 2
        total_size += _words
        max_rec = max(max_rec, _words)
    # ⚠️ METAHEADER = 18 字节：mtType, mtHeaderSize, mtVersion, mtSize(DWORD),
    # mtNoObjects, mtMaxRecord(DWORD), mtNoParameters —— 2026-10-06 修复：旧代码
    # 少了后 4 字节（14 字节头），解析器按 18 读 → 整体错位 → Word 红叉/坏图。
    wmf_header = struct.pack("<HHHIHIH", 1, 9, 0x0300, total_size, 0, max_rec, 0)
    body = wmf_header + b"".join(records)

    inch = 1440
    left, top, right, bottom = 0, 0, w * 15, h * 15
    checksum = 0x9AC6CDD7 ^ left ^ top ^ right ^ bottom ^ inch
    placeable = struct.pack("<IHhhhhHIH", 0x9AC6CDD7, 0, left, top, right, bottom,
                            inch, 0, checksum & 0xFFFF)
    return placeable + body, w, h


def equation_preview_wmf(latex: str, fontsize: float = 10.5,
                         max_width_pt: float | None = None) -> tuple[bytes, float, float]:
    """LaTeX → (预览 WMF, 宽 pt, 高 pt)（matplotlib 渲染 + StretchDIB 封装）。"""
    from .preview import latex_to_png

    png, w_pt, h_pt = latex_to_png(latex, fontsize=fontsize,
                                   max_width_pt=max_width_pt)
    wmf, _w, _h = bitmap_to_wmf(png)
    return wmf, w_pt, h_pt


def _meta_record(func: int, data: bytes) -> bytes:
    """构造单条 META_* 记录。
    
    格式：size(32bit, words), function(16bit), params...
    """
    size = (len(data) + 6) // 2  # 6 = sizeof(size) + sizeof(function)
    return struct.pack('<IH', size, func) + data


def _extract_text(eq: Equation, limit: int = 100) -> str:
    """从 MTEF 树提取可见文本（用于占位图）。"""
    out = []
    
    def walk(obj, depth: int = 0):
        if depth > 20 or len(out) >= limit:
            return
        
        if isinstance(obj, Char):
            if obj.fontpos is not None:
                ch = chr(obj.fontpos) if 32 <= obj.fontpos < 127 else '?'
                out.append(ch)
            elif obj.mtcode is not None:
                # mtcode 是 Symbol 字体的码位，简化为 '?'
                out.append('?')
        
        elif isinstance(obj, Line):
            for item in obj.items:
                walk(item, depth + 1)
            if not obj.null:
                out.append(' ')
        
        elif isinstance(obj, Tmpl):
            # 简化：只提取子项文本，不管模板类型
            for item in obj.items:
                walk(item, depth + 1)
        
        elif isinstance(obj, Pile):
            for line in obj.lines:
                walk(line, depth + 1)
                out.append(' ')
        
        elif isinstance(obj, Matrix):
            for cell in obj.cells:
                walk(cell, depth + 1)
                out.append(' ')
    
    for rec in eq.records:
        walk(rec)
    
    text = ''.join(out).strip()
    # 移除多余空格
    import re
    text = re.sub(r'\s+', ' ', text)
    return text
