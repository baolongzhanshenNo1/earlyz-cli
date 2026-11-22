# -*- coding: utf-8 -*-
"""EZI-0076 源样本星号说明行池排除 + 二/三级迁移丢失升级 FAIL（待产品口径定旗机械化解）。

口径拍板（2026-11-22 对话）：
- 源唯一「编号二级标题」= 模板说明样例行（计算机学报段 25 自证
  「1.1  二级标题 *字体为5号黑体*标题2」，带真 outlineLvl）→ 属噪声，
  源/产物同一套排除口径（EZI-0012 A7 纪律）→ 池排除 → 走标准 N-A；
- 走到「源有真样本、产物排除后无」= 真内容迁移丢失 → 二/三级与一级
  同口径判 FAIL，废除「不判失败待口径定」旗。
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest
from docx import Document
from docx.shared import Twips

sys.path.insert(0, str(Path(__file__).parent))
from test_layout_fidelity_score import build_docx  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# ---------- 谓词 ----------

def test_star_inline_note_predicate_positive():
    from earlyz.validation.layout_fidelity import _is_star_inline_note
    for s in ("1.1  二级标题 *字体为5号黑体*标题2",
              "一级标题*字体为4号黑体*标题1",
              "1.1.1  三级标题 *字体为5号宋体*标题3",
              "正文文本*宋体小4号、行距20磅*示例如下"):
        assert _is_star_inline_note(s), s


def test_star_inline_note_predicate_negative():
    from earlyz.validation.layout_fidelity import _is_star_inline_note
    for s in ("1.1 实验设计与结果分析",
              "2.1 数据获取与预处理方法",
              "3 系统架构与模块设计",
              "摘要",
              "基于深度学习的文档版式迁移方法研究",
              "1.1.2 实验结果" * 20):        # 真标题可以很长，但不含星号格式指令
        assert not _is_star_inline_note(s), s


# ---------- 池排除 ----------

def _mk_pool_docx(path: Path, rows: list[tuple[str, int]]) -> Path:
    """rows: (text, outlineLvl 0/1/2)。最小 doc：只有带 outlineLvl 的段落。"""
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Twips(1417)
    for text, lvl in rows:
        p = doc.add_paragraph(text)
        ppr = p._p.get_or_add_pPr()
        ol = ppr.makeelement(f"{{{W}}}outlineLvl", {f"{{{W}}}val": str(lvl)})
        ppr.append(ol)
    doc.save(str(path))
    return path


def test_heading_pool_excludes_star_inline_note(tmp_path):
    from earlyz.validation.layout_fidelity import _load_docx, heading_samples_by_level
    fx = _mk_pool_docx(tmp_path / "pool.docx", [
        ("1.1  二级标题 *字体为5号黑体*标题2", 1),   # 说明行
        ("1.1 实验设计与结果分析", 1),                # 真标题
        ("一级标题*字体为4号黑体*标题1", 0),          # 说明行（L1）
        ("1 引言", 0),                                 # 真标题（L1）
    ])
    samples, excluded = heading_samples_by_level(_load_docx(fx))
    assert [s["text"] for s in samples[2]] == ["1.1 实验设计与结果分析"]
    assert [s["text"] for s in samples[1]] == ["1 引言"]
    assert excluded["星号内联说明行"] == 2, excluded


# ---------- 迁移丢失：二/三级升级 FAIL ----------

def _drop_level2(path: Path) -> Path:
    """删掉文档里全部二级标题段（build_docx 的 L2 走 Heading 2 样式链，
    直写 pPr/outlineLvl 删不到——按已知文本「1.1 研究背景」删）。"""
    doc = Document(str(path))
    for p in list(doc.paragraphs):
        if (p.text or "").strip().startswith("1.1 研究背景"):
            p._p.getparent().remove(p._p)
    out = path.with_name("prod_no_l2.docx")
    doc.save(str(out))
    return out


def test_level2_migration_loss_now_fails(tmp_path):
    """源有真二级标题、产物没有 → FAIL（升级实锚；旧口径 N-A+悬旗废除）。"""
    from earlyz.validation.layout_fidelity import evaluate_layout_fidelity
    tpl = build_docx(tmp_path / "tpl.docx")
    src = build_docx(tmp_path / "src.docx")
    prod = _drop_level2(build_docx(tmp_path / "prod.docx"))
    rep = evaluate_layout_fidelity(prod, tpl, source=src)
    by_name = {i.name: i for i in rep.items}
    assert "二级标题字号(pt)" in by_name, sorted(by_name)
    for name in ("二级标题字号(pt)", "二级标题中文字体", "二级标题对齐(jc)"):
        it = by_name.get(name)
        if it is None:
            continue                     # 规则缺 jc 时该项不生成——合法
        assert it.verdict == "FAIL", (name, it.verdict, it.note)
        assert "迁移丢失" in it.note, it.note
        assert "待产品口径定" not in it.note, it.note


def test_star_inline_source_takes_na_path_not_flag(tmp_path):
    """源唯一 L2=说明行 → 池排除 → 走「源无编号」N-A（旗文本彻底消失）。"""
    from earlyz.validation.layout_fidelity import evaluate_layout_fidelity
    tpl = build_docx(tmp_path / "tpl.docx")
    src = _mk_pool_docx(tmp_path / "src.docx", [
        ("1 引言", 0),
        ("1.1  二级标题 *字体为5号黑体*标题2", 1),
    ])
    prod = _mk_pool_docx(tmp_path / "prod.docx", [("1 引言", 0)])
    rep = evaluate_layout_fidelity(prod, tpl, source=src)
    it = next((i for i in rep.items if i.name == "二级标题字号(pt)"), None)
    if it is not None:
        assert it.verdict != "FAIL", it.note
        assert "待产品口径定" not in (it.note or ""), it.note


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-q"]))
