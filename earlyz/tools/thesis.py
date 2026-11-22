"""论文排版工具 — 唯一路径 compile_document（3.2 新链路）。"""

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Optional

from earlyz.tools.base import EarlyZTool, ToolResult


class ThesisTool(EarlyZTool):
    """高校论文排版：委托 format_compiler，不再走 EarlyZAgent / Steps。"""

    @property
    def tool_name(self) -> str:
        return "thesis"

    @property
    def description(self) -> str:
        return "高校论文排版：按上传模板反编译规则整篇编译"

    def run(
        self,
        target_docx: str | Path,
        spec_docx: Optional[str | Path] = None,
        output_docx: Optional[str | Path] = None,
        print_mode: bool = False,
        libreoffice_path: Optional[str] = None,
        template_name: Optional[str] = None,  # 保留签名，忽略 YAML 旧模板
    ) -> ToolResult:
        spec = Path(spec_docx) if spec_docx else None
        target = Path(target_docx)
        if spec is None or not spec.exists():
            return ToolResult(
                success=False, tool_name=self.tool_name,
                logs=["整篇论文排版需要规范模板（uploaded_template_docx）。"],
            )
        if not target.exists():
            return ToolResult(
                success=False, tool_name=self.tool_name,
                logs=[f"目标文档不存在: {target}"],
            )
        out = Path(output_docx) if output_docx else target.parent / f"{target.stem}_formatted.docx"
        work = out.parent / f"_compile_{target.stem}"
        work.mkdir(parents=True, exist_ok=True)

        from earlyz.pipeline.format_compiler import compile_document

        res = compile_document(
            str(target),
            "uploaded_template_docx",
            str(spec),
            domain="thesis",
            workdir=str(work),
            print_mode=bool(print_mode),
            libreoffice_path=libreoffice_path,
        )
        produced = res.transaction.output_path if res.transaction else None
        if produced and Path(produced).exists():
            shutil.copy2(produced, out)
        return ToolResult(
            success=bool(res.passed and produced),
            tool_name=self.tool_name,
            output_path=out if out.exists() else None,
            logs=[f"compile {res.transaction.status if res.transaction else 'none'} passed={res.passed}"],
            metadata={
                "fixed_count": len(res.plan.actions) if res.plan else 0,
                "remaining": 0 if res.passed else 1,
                "transaction_status": res.transaction.status if res.transaction else None,
            },
        )
