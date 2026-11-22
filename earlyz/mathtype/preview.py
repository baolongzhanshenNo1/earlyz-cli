"""LaTeX → 公式预览位图（2026-10-06：替换 [Math:…] 占位实现）。

背景：Word 里 DSMT4 对象靠 v:shape/imagedata 的缓存预览图显示，双击才进编辑器。
旧 render.equation_to_wmf 是占位框（"[Math: …]" 矩形）且 WMF 记录畸形——Word 只能
画红叉（.NET GDI+ 也拒载）。本模块用 matplotlib mathtext（Agg）真渲染公式为透明
底 PNG，再由 render.bitmap_to_wmf 封装进 StretchDIB WMF 作为预览。

口径：渲染失败如实抛异常，不静默降级（用户宁可看到报错也不要红叉）。
"""
from __future__ import annotations

import io


def latex_to_png(latex: str, fontsize: float = 10.5, dpi: int = 300,
                 max_width_pt: float | None = None) -> tuple[bytes, float, float]:
    """LaTeX → (PNG 字节, 宽 pt, 高 pt)。透明底、紧致裁剪、过宽等比缩。

    字体口径（2026-10-07 真机「字体不对」）：mathtext 用 **STIX**（Times 系衬线
    数学字体，与论文正文西文观感一致）；默认的 DejaVu Sans（无衬线）一眼假。
    字号默认 10.5 磅＝CJC 期刊模板说明条 14）「公式尺寸：标准 10.5 磅」。
    max_width_pt：镜像编译侧可用宽度（4191 twips≈209pt，留余量取 205），
    超宽等比缩小而不是撑破版心。
    """
    import matplotlib
    matplotlib.use("Agg", force=True)
    from matplotlib import mathtext
    from matplotlib.font_manager import FontProperties

    if max_width_pt is None:
        max_width_pt = 175.0
    s = latex if latex.strip().startswith("$") else f"${latex}$"
    buf = io.BytesIO()
    with matplotlib.rc_context({"mathtext.fontset": "stix"}):
        mathtext.math_to_image(s, buf, prop=FontProperties(size=fontsize),
                               dpi=dpi, format="png")
    buf.seek(0)

    from PIL import Image
    img = Image.open(io.BytesIO(buf.getvalue())).convert("RGBA")
    bbox = img.getbbox()
    if bbox:
        img = img.crop(bbox)
    w_pt = img.width * 72.0 / dpi
    h_pt = img.height * 72.0 / dpi
    if w_pt > max_width_pt and w_pt > 0:
        k = max_width_pt / w_pt
        img = img.resize((max(1, int(img.width * k)), max(1, int(img.height * k))),
                         Image.LANCZOS)
        w_pt, h_pt = max_width_pt, h_pt * k
    out = io.BytesIO()
    img.save(out, format="PNG")
    return out.getvalue(), w_pt, h_pt
