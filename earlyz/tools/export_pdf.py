"""A6 · PDF 导出（M1，2026-10-07）。

产物 = 与 Word 视图一致的 PDF（优先 Word COM 导出；无 Word 时 LibreOffice
兜底——兜底不改变版式验证结论，仅作导出通道）。写入 OUTPUT 后经
/api/download/<name> 下载。
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

from earlyz.validation import proc_track as _proc


def export_pdf(docx_path: str | Path, out_dir: str | Path) -> dict:
    """docx → PDF。返回 {success, pdf, name, renderer, pages?}；失败 raise。"""
    src = Path(docx_path)
    if not src.exists():
        raise FileNotFoundError(f"文档不存在: {src}")
    out_dir = Path(out_dir)
    errors: list[str] = []
    pdf: Path | None = None
    renderer = ""
    pages = None

    # ① Word COM（版式与 Word 视图一致，首选）
    try:
        from earlyz.validation.renderers import render_with_word, word_render_available
        if word_render_available():
            info = render_with_word(src)
            cand = Path(info.get("pdf_path") or "")
            if cand.exists():
                pdf = cand
                renderer = "word"
                pages = info.get("pages")
            else:
                errors.append("Word 导出未产生 PDF")
        else:
            errors.append("Word 不可用")
    except Exception as e:  # noqa: BLE001
        errors.append(f"Word 导出失败: {type(e).__name__}: {e}")

    # ② LibreOffice 兜底
    if pdf is None:
        try:
            from earlyz.validation.providers import _find_soffice
            soffice = _find_soffice()
            if not soffice:
                errors.append("LibreOffice 不可用（soffice 未安装）")
            else:
                with tempfile.TemporaryDirectory(prefix="earlyz-pdf-") as td:
                    _proc.run(
                        [soffice, "--headless", "--convert-to", "pdf",
                         "--outdir", td, str(src.resolve())],
                        check=True, timeout=300, capture_output=True,
                    )
                    cand = next(Path(td).glob("*.pdf"), None)
                    if cand is not None:
                        out_dir.mkdir(parents=True, exist_ok=True)
                        staged = out_dir / cand.name  # {stem}.pdf
                        shutil.copy2(cand, staged)   # 临时目录消亡前迁出
                        pdf = staged
                        renderer = "libreoffice"
        except Exception as e:  # noqa: BLE001
            errors.append(f"LibreOffice 导出失败: {type(e).__name__}: {e}")

    if pdf is None:
        raise RuntimeError("PDF 导出失败（Word 与 LibreOffice 均不可用）: " + "；".join(errors))

    out = out_dir / f"{src.stem}.pdf"
    if pdf.resolve() != out.resolve():
        out_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(pdf, out)
    return {"success": True, "pdf": str(out), "name": out.name,
            "renderer": renderer, "pages": pages}
