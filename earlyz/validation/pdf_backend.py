"""PDF 读取 / 渲染后端：PDFium（pypdfium2，Apache-2.0 / BSD-3-Clause）。

2026-09-30 定案替换 PyMuPDF（AGPL）。实测（nightB/cmp_pdfium，8 份 Word/TeX PDF）：
7–8 ms/页、文字层与 MuPDF 一致度 ≥0.998、图片数一致、老 CJK 字形墨量比 1.00–1.04。

约束：
- PDFium **非线程安全**（pypdfium2 文档：任意两个 pdfium 调用不得并发，哪怕是不同
  文档）。本模块所有 pdfium 调用都在进程级 ``_LOCK`` 内完成，并**显式 close** 全部
  句柄——不把关闭留给 GC finalizer（finalizer 可能在别的线程里触发 pdfium 调用）。
- 坐标统一换成 PyMuPDF 口径：原点在页面（CropBox）左上角、y 向下、单位 pt，调用方的
  越界判定 / y 区间判定原样沿用。旋转页不做换算（Word / TeX 产物不出旋转页）。
- 不依赖 Pillow / numpy：PNG 用标准库 zlib 编码。
"""

from __future__ import annotations

import struct
import threading
import zlib
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Optional

_LOCK = threading.RLock()

# 表单 XObject 嵌套上限（防恶意/损坏 PDF 无限递归；正常文档 ≤3 层）
_MAX_FORM_DEPTH = 8


@dataclass
class PageInfo:
    """单页读取结果（坐标：左上原点、y 向下、pt）。"""

    index: int
    width: float
    height: float
    text: str = ""
    # 文本行：(x0, y0, x1, y1, text)；同一基线上的 PDFium 文本矩形合并成一行
    lines: list[tuple[float, float, float, float, str]] = field(default_factory=list)
    # 图片放置框：(x0, y0, x1, y1)，含表单 XObject 内嵌图片（已按矩阵链换到页面坐标）
    images: list[tuple[float, float, float, float]] = field(default_factory=list)


def _pdfium():
    import pypdfium2 as pdfium  # noqa: PLC0415 — 延迟导入，非渲染环境不强依赖
    import pypdfium2.raw as raw  # noqa: PLC0415

    return pdfium, raw


def _norm_text(s: str) -> str:
    return (s or "").replace("\r\n", "\n").replace("\r", "\n").replace("\x00", "")


def page_count(path: str | Path) -> int:
    pdfium, _ = _pdfium()
    with _LOCK:
        pdf = pdfium.PdfDocument(str(path))
        try:
            return len(pdf)
        finally:
            pdf.close()


def _page_frame(page) -> tuple[float, float, float, float]:
    """(left, top, width, height)：把 PDF 用户坐标换成左上原点所需的偏移。"""
    w, h = page.get_size()
    # 必须用 get_bbox()：PDFium 的 get_cropbox()/get_mediabox() 不继承页面树父节点，
    # 盒子写在 /Pages 上的 PDF（TeX 产物常见）会回退成 ANSI A 612×792 → 全部坐标
    # 偏 49.9pt（2026-09-30 thuthesis/xaut 实测）。get_bbox() 走继承且已与 MediaBox 求交。
    try:
        l, b, r, t = page.get_bbox()
    except Exception:  # noqa: BLE001 — 个别损坏 PDF 无 box，退回 (0,0)-(w,h)
        l, b, r, t = 0.0, 0.0, w, h
    return float(l), float(t), float(w), float(h)


def _to_top_left(frame, l, b, r, t) -> tuple[float, float, float, float]:
    left, top, _, _ = frame
    return (l - left, top - t, r - left, top - b)


def _merge_lines(rects: list[tuple[float, float, float, float, str]]):
    """PDFium 文本矩形 → 行：竖向重叠 ≥ 较矮者一半视为同一行，行内按 x 排序拼接。"""
    rects = sorted(rects, key=lambda r: (r[1], r[0]))
    lines: list[list] = []
    for x0, y0, x1, y1, s in rects:
        placed = False
        for ln in lines:
            ov = min(y1, ln[3]) - max(y0, ln[1])
            if ov >= 0.5 * min(y1 - y0, ln[3] - ln[1]) and ov > 0:
                ln[0], ln[1] = min(ln[0], x0), min(ln[1], y0)
                ln[2], ln[3] = max(ln[2], x1), max(ln[3], y1)
                ln[4].append((x0, s))
                placed = True
                break
        if not placed:
            lines.append([x0, y0, x1, y1, [(x0, s)]])
    out = []
    for x0, y0, x1, y1, parts in lines:
        text = "".join(s for _, s in sorted(parts, key=lambda p: p[0]))
        out.append((x0, y0, x1, y1, _norm_text(text)))
    out.sort(key=lambda r: (r[1], r[0]))
    return out


def _image_boxes(page, raw, frame) -> list[tuple[float, float, float, float]]:
    """页面图片框（含表单 XObject 内图片）。

    表单内子对象的 bounds 在表单坐标系里，需依次乘上各层表单对象的矩阵才到页面坐标。
    """
    out: list[tuple[float, float, float, float]] = []

    def walk(form, chain, depth):
        for obj in page.get_objects(max_depth=1, form=form):
            if obj.type == raw.FPDF_PAGEOBJ_IMAGE:
                try:
                    l, b, r, t = obj.get_bounds()
                except Exception:  # noqa: BLE001 — 取不到框的对象跳过
                    continue
                pts = [(l, b), (l, t), (r, b), (r, t)]
                for m in chain:  # 由内向外
                    pts = [m.on_point(x, y) for x, y in pts]
                xs, ys = [p[0] for p in pts], [p[1] for p in pts]
                out.append(_to_top_left(frame, min(xs), min(ys), max(xs), max(ys)))
            elif obj.type == raw.FPDF_PAGEOBJ_FORM and depth < _MAX_FORM_DEPTH:
                walk(obj, [obj.get_matrix()] + chain, depth + 1)

    walk(None, [], 0)
    return out


def read_pages(path: str | Path, *, text: bool = True, lines: bool = False,
               images: bool = False, pages: Optional[Iterable[int]] = None) -> list[PageInfo]:
    """按需读取每页文字 / 文本行 / 图片框。pages 为 0 基页号（None=全部）。"""
    pdfium, raw = _pdfium()
    result: list[PageInfo] = []
    with _LOCK:
        pdf = pdfium.PdfDocument(str(path))
        try:
            n = len(pdf)
            idx = range(n) if pages is None else [i for i in pages if 0 <= i < n]
            for i in idx:
                page = pdf[i]
                tp = None
                try:
                    frame = _page_frame(page)
                    info = PageInfo(index=i, width=frame[2], height=frame[3])
                    if text or lines:
                        tp = page.get_textpage()
                    if text:
                        info.text = _norm_text(tp.get_text_range())
                    if lines:
                        rects = []
                        for k in range(tp.count_rects()):
                            l, b, r, t = tp.get_rect(k)
                            s = tp.get_text_bounded(l, b, r, t)
                            if s and s.strip():
                                rects.append((*_to_top_left(frame, l, b, r, t), s))
                        info.lines = _merge_lines(rects)
                    if images:
                        info.images = _image_boxes(page, raw, frame)
                    result.append(info)
                finally:
                    if tp is not None:
                        tp.close()
                    page.close()
        finally:
            pdf.close()
    return result


def render_rgb(path: str | Path, index: int, scale: float) -> tuple[int, int, bytes]:
    """渲染单页为紧凑 RGB（无行填充）：返回 (width, height, data)。"""
    got = list(_render_iter(path, [index], scale))  # 耗尽生成器 → 文档句柄在本线程关闭
    if not got:
        raise IndexError(f"page {index} out of range: {path}")
    return got[0][1:]


def _render_iter(path, indices, scale):
    """逐页渲染：每页只在锁内完成 pdfium 调用与拷贝，编码 PNG 等纯 Python 工作放锁外。"""
    pdfium, _ = _pdfium()
    with _LOCK:
        pdf = pdfium.PdfDocument(str(path))
        n = len(pdf)
    try:
        for i in indices:
            if not 0 <= i < n:
                continue
            with _LOCK:
                page = pdf[i]
                try:
                    bm = page.render(scale=scale, rev_byteorder=True)  # 白底 RGB
                    try:
                        w, h, stride, nc = bm.width, bm.height, bm.stride, bm.n_channels
                        buf = bytes(bm.buffer)
                    finally:
                        bm.close()
                finally:
                    page.close()
            row = w * nc
            if nc != 3:  # 理论上 rev_byteorder+BGR → RGB 3 通道；防御：丢弃多余通道
                data = b"".join(
                    bytes(v for j, v in enumerate(buf[y * stride:y * stride + row]) if j % nc < 3)
                    for y in range(h))
            elif stride == row:
                data = buf[: row * h]
            else:
                data = b"".join(buf[y * stride:y * stride + row] for y in range(h))
            yield i, w, h, data
    finally:
        with _LOCK:
            pdf.close()


def write_png(dest: str | Path, width: int, height: int, rgb: bytes, level: int = 6) -> None:
    """标准库 PNG 编码（8-bit RGB，filter 0）。"""
    row = width * 3
    raw_rows = b"".join(b"\x00" + rgb[y * row:(y + 1) * row] for y in range(height))

    def chunk(tag: bytes, payload: bytes) -> bytes:
        return (struct.pack(">I", len(payload)) + tag + payload
                + struct.pack(">I", zlib.crc32(tag + payload) & 0xFFFFFFFF))

    png = (b"\x89PNG\r\n\x1a\n"
           + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
           + chunk(b"IDAT", zlib.compress(raw_rows, level))
           + chunk(b"IEND", b""))
    Path(dest).write_bytes(png)


def render_png_pages(path: str | Path, out_dir: str | Path, *, dpi: float = 110,
                     pages: Optional[Iterable[int]] = None,
                     name: str = "page_{n}.png") -> list[Path]:
    """渲染为 PNG；name 中 {n} 为 1 基页号。返回写出的文件列表（按页序）。"""
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    idx = list(range(page_count(path))) if pages is None else list(pages)
    written: list[Path] = []
    for i, w, h, data in _render_iter(path, idx, dpi / 72.0):
        dest = out / name.format(n=i + 1)
        write_png(dest, w, h, data)
        written.append(dest)
    return written


def compose_contact_sheet(tiles: list[tuple[int, int, bytes]], dest: str | Path,
                          gap: int = 24) -> None:
    """把若干 RGB 图横向拼成一张白底联系表（无文字标注）。"""
    if not tiles:
        return
    height = max(h for _, h, _ in tiles) + gap * 2
    width = sum(w for w, _, _ in tiles) + gap * (len(tiles) + 1)
    white = b"\xff"
    rows = []
    for y in range(height):
        parts = [white * (gap * 3)]
        for w, h, data in tiles:
            ty = y - gap
            if 0 <= ty < h:
                parts.append(data[ty * w * 3:(ty + 1) * w * 3])
            else:
                parts.append(white * (w * 3))
            parts.append(white * (gap * 3))
        rows.append(b"".join(parts))
    write_png(dest, width, height, b"".join(rows))
