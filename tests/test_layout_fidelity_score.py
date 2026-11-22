"""EZI-0005 版式正确率打分器单测：不依赖 Word COM，现场造 docx 断言 PASS/FAIL。

造三份最小 docx（python-docx，1 节，含页眉 + 页脚 PAGE 域）：

* ``template.docx``  —— 期望值来源，交给生产编译器 ``compile_format_spec`` 推规则；
* ``product_ok.docx`` —— 与模板同排版，应全 PASS、exit 0、score 1.0；
* ``product_bad.docx`` —— 故意错字号/字体/缩进/行距/对齐/页边距且没有目录域，
  对应项必须 FAIL、exit 1、score < 1。

目录域（``目录域(TOC)``）按 **EZI-0008 迁移语义**判：期望值只看源文档，不看模板常量。
另外造两份源文档（有/无 TOC 域），覆盖「无源 / 源无 TOC / 源有且产物有 / 源有产物无」。
**EZI-0044**：源无 TOC 时产物必须 0 个（0 个 → PASS，凭空生成 → FAIL）——旧口径整支
N-A，三对体检样本上把这一项做成了 100% 死指标。

标题三项按 **EZI-0009 适用性判据**判：源文档该级段落数为 0 → N-A（产物没有该级是
正确行为），源有才按字号/字体/对齐判 PASS/FAIL。**EZI-0044**：分节数不再是「恒 N-A 的
信息项」——模板单节 → 判产物必须 1 节（EZI-0034 A5），模板多节 → 结构性 N-A（产物
节数=属性变化点数，由源内容角色决定，模板节数不是期望）。另造「无任何标题层级」与
「3 节」「多节模板」几份 docx 覆盖这几条。

EZI-0012 追加两条口径：**页眉/页脚/页码只按 PAGE 域判**（模板部件里无 PAGE 就是
N-A，不按 ``headerReference`` 存在性判，见 2026-11-10 页眉红线）；**标题取样只认
「outlineLvl 级别匹配 + 编号标题文本」**，排除表格内段落、后置节名、表单标签、
封面拉丁大写行，排除后无样本 → N-A（但不许把「源有编号标题、产物没有」也放过）。
**EZI-0044**：页眉引用(headerReference) 明确废弃（不再输出该项）；页眉字号改锚
「模板页眉 PAGE 段落有效字号」，与页眉有没有文字无关。

EZI-0020 追加一条口径：**正文行距取样排掉表单标签**（冒号结尾短行，「学生姓名:」
这类字段名不是正文段落，与标题取样共用同一条规则）——交大中期模板 6 段标签
``500/exact`` 与 6 段真正文 ``400/atLeast`` 打平，``most_common`` 按文档序取到标签，
把已经排对（``400/atLeast``）的产物判成 FAIL。

同时覆盖：``rule.expected`` 不是 Mapping（bool）时不许把整组打分打崩（A7）、
``--json`` 结构（A3）、CLI 退出码 0/1/2、``--source``、N-A 不算失败、
分数可复算、只读不改文件。
"""
from __future__ import annotations

import base64
import hashlib
import importlib.util
import io
import json
import sys
from pathlib import Path

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Twips

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from earlyz.validation.layout_fidelity import (  # noqa: E402
    FAIL,
    NA,
    PASS,
    evaluate_layout_fidelity,
    format_table,
    score_items,
)

SCRIPT_PATH = ROOT / "scripts" / "check_layout_fidelity.py"
#: EZI-0077：开源 CLI 面不带 scripts/ 跑器——跑器缺位时跳过 CLI 壳测试
#: （主仓永远在，主仓永不跳）。引擎行为测试不受影响。
pytest.importorskip = pytest.importorskip  # noqa: 保 typing
CLI_RUNNER_PRESENT = SCRIPT_PATH.is_file()


# ────────────────────────────── 造 docx ──────────────────────────────


def _set_run(run, east_asia: str, size_pt: float) -> None:
    """把 run 的字体名/字号写成显式直接格式（eastAsia + ascii + hAnsi）。"""
    run.font.size = Pt(size_pt)
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(attr), east_asia)


def _add_field(paragraph, instr: str, size_pt: float | None = None, text: str = "1") -> None:
    """往段落里插一个 ``w:fldSimple``（最简域形态）。"""
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr)
    run = OxmlElement("w:r")
    if size_pt is not None:
        rpr = OxmlElement("w:rPr")
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(round(size_pt * 2))))
        rpr.append(sz)
        run.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    run.append(t)
    fld.append(run)
    paragraph._p.append(fld)


def build_docx(path: Path, *, body_size: float = 12.0, body_font: str = "宋体",
               indent_chars: int = 200, line_spacing: float = 1.5,
               body_align=WD_ALIGN_PARAGRAPH.JUSTIFY, h1_size: float = 16.0,
               h1_font: str = "宋体", h1_align=WD_ALIGN_PARAGRAPH.CENTER,
               margins: tuple[int, int, int, int] = (1417, 1134, 850, 1587),
               with_toc: bool = True, header_size: float = 10.5,
               footer_size: float = 10.5, with_headings: bool = True,
               sections: int = 1,
               header_text: str | None = "某某大学本科毕业设计（论文）",
               header_page: bool = False, with_header: bool = True,
               footer_text: str | None = None, footer_page: bool = True,
               with_footer: bool = True) -> Path:
    """造一份最小论文式 docx：``sections`` 节 + 一级/二级标题 + 3 段正文 + 页眉页脚。

    ``with_headings=False`` 造「裸内容」型文档（一个标题层级都没有——合同/中介合同
    那类源文档），用来验证 EZI-0009 的适用性判据。

    页眉/页脚的四种形态（EZI-0012 A1-A5）：

    * 默认：页眉是**纯文字版头**（无 PAGE）、页脚有 PAGE 域；
    * ``header_page=True`` / ``footer_page=False``：换 PAGE 域的有无；
    * ``header_text=None``：页眉部件空着（交大中期的真实产物形态）；
    * ``with_header=False`` / ``with_footer=False``：连部件都不建。
    """
    doc = Document()
    sec = doc.sections[0]
    top, right, bottom, left = margins
    sec.top_margin, sec.right_margin = Twips(top), Twips(right)
    sec.bottom_margin, sec.left_margin = Twips(bottom), Twips(left)
    sec.header_distance, sec.footer_distance = Twips(1134), Twips(567)

    if with_headings:
        h1 = doc.add_heading("", level=1)
        h1.paragraph_format.alignment = h1_align
        _set_run(h1.add_run("1 绪论"), h1_font, h1_size)

        h2 = doc.add_heading("", level=2)
        _set_run(h2.add_run("1.1 研究背景"), "黑体", 14.0)

    for i in range(6):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = line_spacing
        p.paragraph_format.alignment = body_align
        if indent_chars:
            ppr = p._element.get_or_add_pPr()
            ind = OxmlElement("w:ind")
            ind.set(qn("w:firstLineChars"), str(indent_chars))
            ind.set(qn("w:firstLine"), str(int(body_size * 20 * indent_chars / 100)))
            ppr.append(ind)
        _set_run(p.add_run(f"第 {i + 1} 段正文，用于验证字号、缩进、行距与对齐的判定。"),
                 body_font, body_size)

    if with_toc:
        toc_p = doc.add_paragraph()
        _add_field(toc_p, ' TOC \\o "1-2" \\h \\z \\u ', text="")

    if with_header:
        hp = sec.header.paragraphs[0]
        if header_text:
            _set_run(hp.add_run(header_text), "宋体", header_size)
        if header_page:
            _add_field(hp, " PAGE   \\* MERGEFORMAT ", size_pt=header_size, text="1")
    if with_footer:
        fp = sec.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if footer_text:
            _set_run(fp.add_run(footer_text), "宋体", footer_size)
        if footer_page:
            _add_field(fp, " PAGE   \\* MERGEFORMAT ", size_pt=footer_size, text="1")

    for _ in range(sections - 1):        # 多节：验证「分节数」信息项的三个数各自独立
        doc.add_section()

    doc.save(str(path))
    return path


@pytest.fixture(scope="module")
def docs(tmp_path_factory):
    d = tmp_path_factory.mktemp("layout_fidelity")
    return {
        "template": build_docx(d / "template.docx"),
        "ok": build_docx(d / "product_ok.docx"),
        "bad": build_docx(
            d / "product_bad.docx",
            body_size=10.5, body_font="楷体", indent_chars=0, line_spacing=1.0,
            body_align=WD_ALIGN_PARAGRAPH.LEFT, h1_size=18.0, h1_font="黑体",
            h1_align=WD_ALIGN_PARAGRAPH.LEFT, margins=(1417, 1701, 1134, 1701),
            with_toc=False,
        ),
        # 迁移语义的期望值来源：源有 TOC / 源无 TOC
        "source_toc": build_docx(d / "source_with_toc.docx"),
        "source_no_toc": build_docx(d / "source_without_toc.docx", with_toc=False),
        # EZI-0009 适用性判据：裸内容型源（一个标题层级都没有）+ 多节源（3 节）
        "source_no_heads": build_docx(d / "source_without_headings.docx",
                                      with_headings=False, with_toc=False),
        "source_3sect": build_docx(d / "source_three_sections.docx", sections=3),
    }


@pytest.fixture(scope="module")
def reports(docs):
    return {
        "ok": evaluate_layout_fidelity(docs["ok"], docs["template"]),
        "bad": evaluate_layout_fidelity(docs["bad"], docs["template"]),
    }


def _load_cli():
    """直接 import 脚本（不走子进程），拿到 CLI 壳的 main()。"""
    if not CLI_RUNNER_PRESENT:
        pytest.skip("scripts/check_layout_fidelity.py 不在本仓（开源 CLI 面）")
    spec = importlib.util.spec_from_file_location("check_layout_fidelity_cli", SCRIPT_PATH)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    return mod


# ────────────────────────────── A1/A2：覆盖与判定 ──────────────────────────────


def test_compliant_product_all_pass(reports):
    rep = reports["ok"]
    assert rep.total >= 12, f"判定项不足：{rep.total}"
    bad = [(i.name, i.expected, i.actual) for i in rep.items if i.verdict not in (PASS, NA)]
    assert bad == [], f"合规产物出现 FAIL 项：{bad}"
    assert rep.failed == 0
    assert rep.exit_code == 0
    assert rep.score == 1.0


def test_six_categories_covered(reports):
    """七类（EZI-0011 加了「题注绑定」这一类）——类目名固定，免得新项偷偷落空类。"""
    cats = {i.category for i in reports["ok"].items}
    assert cats == {"标题", "正文", "页边距", "目录域", "页眉页脚", "题注绑定", "分节分栏"}, cats


def test_every_item_has_name_expected_actual_verdict(reports):
    for rep in reports.values():
        for it in rep.items:
            assert it.name and it.category
            assert it.verdict in (PASS, FAIL, NA)
            assert it.expected is not None or it.verdict == NA, it.name
            assert it.actual is not None, f"{it.name} 给不出实测值"


def test_broken_product_fails_the_right_items(reports):
    rep = reports["bad"]
    by_name = {i.name: i for i in rep.items}
    expected_fails = [
        "一级标题字号(pt)",       # 18pt vs 16pt
        "一级标题中文字体",        # 黑体 vs 宋体
        "一级标题对齐(jc)",       # start vs center
        "正文主字号(pt)",         # 10.5 vs 12
        "正文中文字体",           # 楷体 vs 宋体
        "正文首行缩进(字符)",      # 0 vs 2
        "正文行距(line/lineRule)",  # 240/auto vs 360/auto
        "正文对齐(jc)",           # start vs both
        "页边距(top/right/bottom/left)",
        # 目录域不在这一列：EZI-0008 起按迁移语义判（期望值读源文档，不是常量）。
        # 各情形的专项断言见下面「A4：目录域迁移语义」小节。
    ]
    missing = [n for n in expected_fails if n not in by_name]
    assert missing == [], f"缺少判定项：{missing}"
    not_failed = [(n, by_name[n].actual) for n in expected_fails if by_name[n].verdict != FAIL]
    assert not_failed == [], f"该 FAIL 却未 FAIL：{not_failed}"
    assert rep.failed >= len(expected_fails)
    assert rep.score < 1.0
    assert rep.exit_code == 1
    # 未动的项仍应 PASS，证明判定不是「一律 FAIL」
    for name in ("页码域(PAGE)", "正文分栏数"):
        assert by_name[name].verdict == PASS, name
    # EZI-0044：页眉引用(headerReference) **已废弃**——红线禁止按存在性判 FAIL，
    # 唯一有意义的替代语义（模板页眉有 PAGE → 产物页眉须有 PAGE）已被「页码域(PAGE)」
    # 与「页眉字号(pt)」覆盖，再留一项只能重复计。此项不许再出现在报告里。
    assert "页眉引用(headerReference)" not in by_name, \
        "页眉引用已废弃（EZI-0044），不许悄悄回来"
    # 页眉字号：模板页眉是纯文字版头（无 PAGE 域）→ 没有可继承的基准，N-A 且 note 引红线
    hdr = by_name["页眉字号(pt)"]
    assert hdr.verdict == NA and "2026-11-10" in hdr.note, (hdr.verdict, hdr.note)


def test_score_is_recomputable(reports):
    for rep in reports.values():
        rec = score_items(rep.items)
        assert rec["score"] == rep.score
        assert rec["passed"] == rep.passed
        assert rec["total"] == rep.total
        assert rec["judged"] == rep.total - rep.not_applicable
        assert rec["score"] == pytest.approx(
            rep.passed / rep.judged if rep.judged else 0.0, abs=1e-4)


def test_na_is_not_a_failure(reports):
    rep = reports["ok"]
    na = [i for i in rep.items if i.verdict == NA]
    assert na, "合规样本应至少有一项 N-A（三级标题在模板里无规则/产物无该级段落）"
    assert rep.exit_code == 0  # N-A 不算失败
    assert all(i.verdict != FAIL for i in na)


# ────────────────────────────── A4：目录域迁移语义 ──────────────────────────────


def _toc_item(rep):
    return next(i for i in rep.items if i.name == "目录域(TOC)")


def test_toc_without_source_is_na(docs):
    """不给源文档 → N-A，note 写明原因，且不影响 exit code（禁止猜期望）。"""
    rep = evaluate_layout_fidelity(docs["ok"], docs["template"])
    it = _toc_item(rep)
    assert it.verdict == NA, (it.expected, it.actual, it.note)
    assert it.expected is None
    assert "未提供源文档" in it.note
    assert rep.exit_code == 0
    assert rep.source is None


def test_toc_source_without_toc_and_product_without_toc_passes(docs):
    """EZI-0044：源无目录域 + 产物 0 个 = **PASS**（「不生成」是可判约束）。

    旧口径把这一支判 N-A（「源没有就不判」），等于对「不该生成却生成」不设防；
    EZI-0038 修的正是「源无 TOC 域仍生成目录」。现在 0 个 → PASS，有 → FAIL。
    """
    rep = evaluate_layout_fidelity(docs["bad"], docs["template"], source=docs["source_no_toc"])
    it = _toc_item(rep)
    assert it.verdict == PASS, (it.expected, it.actual, it.note)
    assert it.expected == "0 个（源无目录域，不许生成）", it.expected
    assert "源文档无目录域" in it.note and "PASS" in it.note, it.note
    assert it.actual == "未找到"
    assert rep.source == str(docs["source_no_toc"])


def test_toc_source_without_toc_but_product_has_toc_fails(docs):
    """EZI-0044 反例：源无目录域、产物凭空生成 TOC 域 → FAIL（不是恒 PASS）。"""
    rep = evaluate_layout_fidelity(docs["ok"], docs["template"], source=docs["source_no_toc"])
    it = _toc_item(rep)
    assert it.verdict == FAIL, (it.expected, it.actual, it.note)
    assert it.actual.startswith("存在"), it.actual
    assert "不得生成" in it.note and "只迁移，不生成" in it.note, it.note
    assert rep.exit_code == 1


def test_toc_source_with_toc_and_product_with_toc_passes(docs):
    """源有目录域 + 产物也有 → PASS。"""
    rep = evaluate_layout_fidelity(docs["ok"], docs["template"], source=docs["source_toc"])
    it = _toc_item(rep)
    assert it.verdict == PASS, (it.expected, it.actual, it.note)
    assert it.expected is not None and "源文档有" in str(it.expected)
    assert it.actual.startswith("存在")


def test_toc_source_with_toc_but_product_lacks_toc_fails(docs):
    """源有目录域但产物没有 → FAIL（迁移丢失），exit code 必须为 1。"""
    rep = evaluate_layout_fidelity(docs["bad"], docs["template"], source=docs["source_toc"])
    it = _toc_item(rep)
    assert it.verdict == FAIL, (it.expected, it.actual, it.note)
    assert it.actual == "未找到"
    assert rep.failed >= 1
    assert rep.exit_code == 1


# ───────────────── A5（EZI-0009）：标题适用性判据 / 分节数信息项 ─────────────────


def _by_name(rep, name):
    return next(i for i in rep.items if i.name == name)


_HEADING_ITEMS_L1 = ("一级标题字号(pt)", "一级标题中文字体", "一级标题对齐(jc)")


def test_heading_items_na_when_source_has_no_such_level(docs):
    """A5-1：源文档该级段落数为 0 → 三项 N-A（产物有没有该级都不判对错）。

    两种产物都覆盖：

    * 产物也没有该级段落——合同/中介合同的真实形态（源与模板都 0 个一级标题，
      旧口径却报「一级标题字号 FAIL 模板 16.0 vs 产物 无该级段落」）；
    * 产物自己带了该级段落（源没有 = 不是迁移来的），仍不判字号/字体/对齐。
    """
    for product in (docs["source_no_heads"], docs["bad"]):
        rep = evaluate_layout_fidelity(product, docs["template"], source=docs["source_no_heads"])
        for name in _HEADING_ITEMS_L1:
            it = _by_name(rep, name)
            assert it.verdict == NA, (product.name, name, it.verdict, it.note)
            assert it.expected is None, f"{name} 判 N-A 时不该出现期望值：{it.expected}"
            assert "源文档无编号" in it.note and "无法比较" in it.note, it.note
        assert _by_name(rep, "二级标题字号(pt)").verdict == NA
        assert not [i for i in rep.items
                    if i.name.startswith("一级标题") and i.verdict != NA]


def test_heading_items_pass_when_source_and_product_both_have_level(docs):
    """A5-2：源有该级、产物也有且符合模板 → 三项 PASS，note 写明源的级数。"""
    rep = evaluate_layout_fidelity(docs["ok"], docs["template"], source=docs["ok"])
    for name in _HEADING_ITEMS_L1:
        it = _by_name(rep, name)
        assert it.verdict == PASS, (name, it.expected, it.actual, it.note)
        assert "源 1 个" in it.note, it.note


def test_heading_items_fail_when_source_and_product_both_have_level_but_differ(docs):
    """A5-3：源有该级、产物也有但排版不符 → 照旧 FAIL（真实缺陷不许被 N-A 吃掉）。"""
    rep = evaluate_layout_fidelity(docs["bad"], docs["template"], source=docs["source_toc"])
    for name in _HEADING_ITEMS_L1:
        it = _by_name(rep, name)
        assert it.verdict == FAIL, (name, it.expected, it.actual, it.note)
        assert it.expected is not None
    assert rep.exit_code == 1


def test_heading_level1_fails_when_source_has_it_but_product_drops_it(docs):
    """源有一级标题、产物没有 = 迁移丢失，仍判 FAIL（新口径不是「一律放过」）。"""
    rep = evaluate_layout_fidelity(docs["source_no_heads"], docs["template"],
                                   source=docs["source_toc"])
    it = _by_name(rep, "一级标题字号(pt)")
    assert it.verdict == FAIL, (it.expected, it.actual, it.note)
    assert it.actual == "无编号该级标题样本"
    assert "迁移丢失" in it.note


def test_section_count_single_section_template_judges_product(docs):
    """EZI-0044：模板单节（没有任何属性变化点）→ 产物必须 1 节，PASS/FAIL 都判。

    旧口径「恒 N-A 信息项」把这条硬约束也一起吞了（EZI-0034 A5 反例要求
    「模板单节（公文/合同）产物节数不得增加」）。
    """
    rep = evaluate_layout_fidelity(docs["ok"], docs["template"], source=docs["source_3sect"])
    it = _by_name(rep, "分节数(sectPr)")
    assert it.verdict == PASS, (it.expected, it.actual, it.note)
    assert it.expected == 1 and it.actual == 1
    assert "模板单节" in it.note and "EZI-0034" in it.note, it.note
    assert ("源=3" in it.note and "模板=1" in it.note and "产物=1" in it.note), it.note

    rep2 = evaluate_layout_fidelity(docs["source_3sect"], docs["template"],
                                   source=docs["source_3sect"])
    it2 = _by_name(rep2, "分节数(sectPr)")
    assert it2.verdict == FAIL, (it2.expected, it2.actual, it2.note)
    assert it2.expected == 1 and it2.actual == 3
    assert "发明结构" in it2.note, it2.note


def test_section_count_multi_section_template_is_structural_na(tmp_path, docs):
    """EZI-0044：模板多节 → 结构性 N-A，note 说清「产物节数=属性变化点数」与缺的证据。"""
    tpl3 = build_docx(tmp_path / "tpl_3sect.docx", sections=3)
    rep = evaluate_layout_fidelity(docs["ok"], tpl3, source=docs["source_3sect"])
    it = _by_name(rep, "分节数(sectPr)")
    assert it.verdict == NA and it.expected is None, (it.verdict, it.expected, it.note)
    assert ("源=3" in it.note and "模板=3" in it.note and "产物=1" in it.note), it.note
    assert "属性变化点" in it.note and "映射证据" in it.note, it.note
    assert "信息项" in it.note, it.note

    it2 = _by_name(evaluate_layout_fidelity(docs["ok"], tpl3), "分节数(sectPr)")
    assert it2.verdict == NA and "源=?（未提供源文档）" in it2.note, it2.note


# ─────────── A5/A7（EZI-0012）：页眉页脚页码只按 PAGE 域判 + 标题取样口径 ───────────


def _hf_report(product, template):
    """只判页眉页脚：给空 spec（标题/正文一律 N-A），不依赖模板编译。"""
    return evaluate_layout_fidelity(
        product, template, spec={"structure": {"sections": []}, "rules": []})


def _named(rep, name):
    return next(i for i in rep.items if i.name == name)


def test_header_reference_item_is_deprecated(tmp_path):
    """EZI-0044：页眉引用(headerReference) **废弃**——任何形态都不再输出该项。

    红线禁止按 headerReference 存在性判 FAIL；唯一有意义的替代语义（模板页眉有 PAGE
    → 产物页眉须有 PAGE）已由「页码域(PAGE)」+「页眉字号(pt)」覆盖。这里同时锁住
    「废弃 ≠ 把缺陷也放走」：模板页眉有 PAGE、产物完全没有页眉时，页码域必须 FAIL。
    """
    text_only = build_docx(tmp_path / "tpl_text_only.docx")           # 模板页眉纯文字
    hdr_page = build_docx(tmp_path / "tpl_hdr_page.docx", header_text=None, header_page=True)
    no_hf = build_docx(tmp_path / "prod_no_hf.docx", with_header=False, with_footer=False)
    hdr_page_prod = build_docx(tmp_path / "prod_hdr_page.docx", header_text=None,
                               header_page=True)
    for prod, tpl in ((text_only, text_only), (hdr_page_prod, hdr_page),
                      (no_hf, hdr_page), (no_hf, text_only)):
        rep = _hf_report(prod, tpl)
        names = {i.name for i in rep.items}
        assert "页眉引用(headerReference)" not in names, names
        assert "页脚引用(footerReference)" in names        # 页脚照旧（无「不迁文字」红线）
    # 真缺陷仍有归属：模板页眉有 PAGE、产物连页眉页脚都没有 → 页码域 FAIL
    rep = _hf_report(no_hf, hdr_page)
    assert _named(rep, "页码域(PAGE)").verdict == FAIL, _named(rep, "页码域(PAGE)").note


def test_header_size_na_when_template_header_has_no_page(docs):
    """EZI-0044：模板页眉无 PAGE 域 → 那个字号是模板**文字**的字号，不迁 → N-A。

    三对体检样本全是这个形态（矿大 37 个页眉部件、北航 24 个，页码在页脚），
    这也是旧口径 100% N-A 的真实原因之一：不是取样写坏了，而是基准选错了
    （拿模板页眉文字字号去比产物 \xa0 占位符）。
    """
    it = _named(evaluate_layout_fidelity(docs["ok"], docs["template"]), "页眉字号(pt)")
    assert it.verdict == NA and it.expected is None, (it.verdict, it.expected, it.note)
    assert "模板页眉无 PAGE 域" in it.note and "2026-11-10" in it.note, it.note
    assert "缺「模板页眉 PAGE 段落」" in it.note, it.note


def test_header_size_passes_when_both_sides_page_sizes_match(tmp_path):
    """EZI-0044：模板页眉有 PAGE（10.5pt）→ 产物页眉 PAGE 也 10.5pt → PASS。

    与「页眉有没有文字」无关：两边都不写文字（``header_text=None``），只有 PAGE 域。
    """
    tpl = build_docx(tmp_path / "tpl_hdr_page.docx", header_text=None, header_page=True,
                     header_size=10.5)
    prod = build_docx(tmp_path / "prod_hdr_page.docx", header_text=None, header_page=True,
                      header_size=10.5)
    it = _named(_hf_report(prod, tpl), "页眉字号(pt)")
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == 10.5 and it.actual == 10.5
    assert "PAGE 段落有效字号" in it.note and "PAGE run" in it.note, it.note


def test_header_size_fails_when_page_size_differs(tmp_path):
    """EZI-0044 反例：两侧都有 PAGE、字号不符 → FAIL（新口径不是「一律 N-A/PASS」）。"""
    tpl = build_docx(tmp_path / "tpl_hdr_page.docx", header_text=None, header_page=True,
                     header_size=10.5)
    prod = build_docx(tmp_path / "prod_hdr_big.docx", header_text=None, header_page=True,
                      header_size=14.0)
    it = _named(_hf_report(prod, tpl), "页眉字号(pt)")
    assert it.verdict == FAIL, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == 10.5 and it.actual == 14.0


def _strip_sizes(path: Path) -> Path:
    """删掉 styles.xml 与 header 部件里的全部 ``w:sz``（造「没有字号证据」的样本）。

    判定层解析 PAGE 段落字号时走样式链（run → 段落标记 → 段落样式 → docDefaults），
    python-docx 默认模板的 ``Header`` 样式带 11pt，所以「模板 PAGE 段落无字号证据」
    只能靠把 styles.xml 里的 sz 一起删干净来造。
    """
    import zipfile

    from lxml import etree

    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    for name in list(parts):
        if name != "word/styles.xml" and not name.startswith("word/header"):
            continue
        root = etree.fromstring(parts[name])
        for sz in list(root.iter(qn("w:sz"))):
            sz.getparent().remove(sz)
        parts[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                                     standalone=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in parts.items():
            z.writestr(n, data)
    return path


def test_header_size_na_when_template_page_paragraph_has_no_size(tmp_path):
    """模板页眉有 PAGE 但 run/段落标记/样式链都没有 w:sz → 无字号证据，N-A。"""
    tpl = build_docx(tmp_path / "tpl_hdr_page_nosz.docx", header_text=None,
                     header_page=True)
    _strip_sizes(tpl)                       # 连 Header 样式与 docDefaults 的 sz 一起删
    prod = build_docx(tmp_path / "prod_hdr_page.docx", header_text=None, header_page=True,
                      header_size=10.5)
    it = _named(_hf_report(prod, tpl), "页眉字号(pt)")
    assert it.verdict == NA and it.expected is None, (it.verdict, it.expected, it.note)
    assert "没有 w:sz" in it.note and "无字号证据" in it.note, it.note


def test_header_size_na_when_product_has_no_header_part(tmp_path):
    """EZI-0044：模板页眉有 PAGE、产物没有页眉部件 → N-A（缺 PAGE 链由页码域判）。"""
    tpl = build_docx(tmp_path / "tpl_hdr_page.docx", header_text=None, header_page=True)
    prod = build_docx(tmp_path / "prod_no_header.docx", with_header=False)
    it = _named(_hf_report(prod, tpl), "页眉字号(pt)")
    assert it.verdict == NA and it.expected is None, (it.verdict, it.expected, it.note)
    assert it.actual == "无页眉部件", it.actual
    assert "不重复计" in it.note and "PAGE 段落有效字号" in it.note, it.note


def test_header_size_na_when_product_header_has_no_page(tmp_path):
    """EZI-0044：产物页眉存在但没有 PAGE 域 → 无页码链，无字号可比 → N-A。"""
    tpl = build_docx(tmp_path / "tpl_hdr_page.docx", header_text=None, header_page=True)
    prod = build_docx(tmp_path / "prod_hdr_empty.docx", header_text=None)
    it = _named(_hf_report(prod, tpl), "页眉字号(pt)")
    assert it.verdict == NA and it.expected is None, (it.verdict, it.expected, it.note)
    assert it.actual == "页眉无 PAGE 域", it.actual
    assert "不重复计" in it.note, it.note


def test_page_field_na_when_no_page_anywhere(tmp_path):
    """A5-④：模板/spec/产物均无 PAGE 域 → 页码域 N-A（删掉常量判据）。"""
    tpl = build_docx(tmp_path / "tpl_no_page.docx", header_text="版头文字",
                     footer_text="编号 001", footer_page=False)
    prod = build_docx(tmp_path / "prod_no_page.docx", header_text="版头文字",
                      footer_text="编号 001", footer_page=False)
    rep = _hf_report(prod, tpl)
    it = _named(rep, "页码域(PAGE)")
    assert it.verdict == NA and it.expected is None, (it.verdict, it.expected, it.note)
    assert "均无 PAGE 域" in it.note, it.note
    assert _named(rep, "页脚引用(footerReference)").verdict == NA    # 模板页脚也是纯文字
    assert rep.failed == 0


def test_page_field_fails_when_template_has_page_and_product_lacks_it(tmp_path):
    """A3 反例：模板页脚含 PAGE、产物无页脚 → 页码域 FAIL（模板真有的期望）。"""
    tpl = build_docx(tmp_path / "tpl_ftr_page.docx", header_text=None)
    prod = build_docx(tmp_path / "prod_no_footer.docx", with_footer=False)
    rep = _hf_report(prod, tpl)
    page = _named(rep, "页码域(PAGE)")
    assert page.verdict == FAIL, (page.verdict, page.note)
    assert "模板页码在页脚" in str(page.expected)
    assert _named(rep, "页脚引用(footerReference)").verdict == FAIL


def test_footer_page_size_na_when_template_footer_has_no_page(tmp_path):
    """模板页脚只有文字（无 PAGE）→ 那个字号是页脚文字的，不是页码字号 → N-A。"""
    tpl = build_docx(tmp_path / "tpl_ftr_text.docx", header_text=None,
                     footer_text="编号 001", footer_page=False)
    prod = build_docx(tmp_path / "prod_ftr_page.docx", header_text=None)
    it = _named(_hf_report(prod, tpl), "页脚页码字号(pt)")
    assert it.verdict == NA and it.expected is None, (it.verdict, it.expected, it.note)
    assert "不是页码字号" in it.note, it.note


# ── A7：标题取样只认「outlineLvl 级别匹配 + 编号标题文本」──

_HEADING1_SPEC = {"structure": {"sections": []},
                  "rules": [{"selector": {"semantic_role": "heading1"},
                             "expected": {"font_size_pt": 16.0}}]}


def _build_heading_docx(path: Path, blocks):
    """``blocks`` = [(文本, 字号, 对齐)]，全部用 Heading 1（outlineLvl=0）。"""
    doc = Document()
    for text, size, align in blocks:
        p = doc.add_paragraph(style="Heading 1")
        p.paragraph_format.alignment = align
        _set_run(p.add_run(text), "宋体", size)
    doc.save(str(path))
    return path


def _one_heading_item(path):
    return _named(evaluate_layout_fidelity(path, path, spec=_HEADING1_SPEC, source=path),
                  "一级标题字号(pt)")


def test_heading_ignores_section_names(tmp_path):
    """ACM 形态：ABSTRACT/REFERENCES 也是 outlineLvl=0，取样只认编号行。"""
    p = _build_heading_docx(tmp_path / "acm_like.docx", [
        ("ABSTRACT", 11.0, WD_ALIGN_PARAGRAPH.LEFT),
        ("1  Introduction", 16.0, WD_ALIGN_PARAGRAPH.JUSTIFY),
        ("REFERENCES", 11.0, WD_ALIGN_PARAGRAPH.LEFT),
    ])
    it = _one_heading_item(p)
    assert it.verdict == PASS, (it.verdict, it.actual, it.note)
    assert it.actual == 16.0
    assert "后置节名 2" in it.note, it.note


def test_heading_ignores_form_label_and_cover_caps_line(tmp_path):
    """交大中期形态：校名大写行 18pt、表单字段 10.5pt 不许拉低真标题（16pt）。"""
    p = _build_heading_docx(tmp_path / "sjtu_like.docx", [
        ("SHANGHAI JIAO TONG UNIVERSITY", 18.0, WD_ALIGN_PARAGRAPH.CENTER),
        ("1 毕业设计（论文）中期检查报告", 16.0, WD_ALIGN_PARAGRAPH.JUSTIFY),
        ("学院 （系）:", 10.5, None),
    ])
    it = _one_heading_item(p)
    assert it.verdict == PASS and it.actual == 16.0, (it.verdict, it.actual, it.note)
    assert "封面大写行 1" in it.note and "表单标签 1" in it.note, it.note


def test_heading_ignores_paragraph_inside_table(tmp_path):
    """IEEE 形态：真标题装在表格里（双栏布局），按 A7 不算样本。"""
    doc = Document()
    tbl = doc.add_table(rows=1, cols=1)
    cell_p = tbl.cell(0, 0).paragraphs[0]
    cell_p.style = doc.styles["Heading 1"]
    _set_run(cell_p.add_run("1 表格里的标题"), "宋体", 10.0)
    path = tmp_path / "ieee_like.docx"
    doc.save(str(path))

    it = _one_heading_item(path)
    assert it.verdict == NA, (it.verdict, it.actual, it.note)
    assert "表格内 1" in it.note and "无法比较" in it.note, it.note


def test_heading_na_when_only_noise_paragraphs(tmp_path):
    """排除后无样本 → N-A（note 写「无编号一级标题，无法比较」+ 各类排除条数）。"""
    p = _build_heading_docx(tmp_path / "noise_only.docx", [
        ("ABSTRACT", 11.0, None),
        ("学院 （系）:", 10.5, None),
        ("SHANGHAI JIAO TONG UNIVERSITY", 18.0, None),
        ("本科生毕业设计（论文）中期检查报告", 22.0, None),
    ])
    it = _one_heading_item(p)
    assert it.verdict == NA and it.expected is None
    assert "无编号一级标题，无法比较" in it.note, it.note
    for bit in ("后置节名 1", "表单标签 1", "封面大写行 1", "非编号行 1"):
        assert bit in it.note, (bit, it.note)


def test_heading_level1_fails_when_source_has_numbered_heading_but_product_only_noise(tmp_path):
    """A7 不许把标题项搞成永远 N-A：源有编号标题、产物只有噪声 → 仍判 FAIL。"""
    src = _build_heading_docx(tmp_path / "src_numbered.docx", [("1  绪论", 16.0, None)])
    prod = _build_heading_docx(tmp_path / "prod_noise.docx", [("ABSTRACT", 12.0, None)])
    rep = evaluate_layout_fidelity(prod, prod, spec=_HEADING1_SPEC, source=src)
    it = _named(rep, "一级标题字号(pt)")
    assert it.verdict == FAIL, (it.verdict, it.actual, it.note)
    assert "迁移丢失" in it.note, it.note


# ────── EZI-0041：手写目录条目不许当正文段 / 二级标题（EZI-0038 合并后的假 FAIL）──────


def _add_tabbed_run(paragraph, text: str, size_pt: float) -> None:
    """把**字面制表符**写进 ``w:t``（python-docx 的 ``Run.text`` 会把 ``\\t`` 变 ``w:tab``）。

    真实样本（可行研/功能设计/安装手册）里手写目录条目就是 ``<w:t>1.1.\\t编写目的\\t4</w:t>``，
    取样函数只拼 ``w:t``，所以这里必须造同形态的 XML。
    """
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(round(size_pt * 2))))
    rpr.append(sz)
    run.append(rpr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    run.append(t)
    paragraph._p.append(run)


_HEADING2_SPEC = {"structure": {"sections": []},
                  "rules": [{"selector": {"semantic_role": "heading2"},
                             "expected": {"font_size_pt": 16.0, "jc": "both"}}]}


def _build_manual_toc_headings(path: Path, entries, real=()) -> Path:
    """造「手写目录条目（Heading 2，制表符 + 页码）+ 真二级标题」的 docx。"""
    doc = Document()
    for text, size in entries:
        p = doc.add_paragraph(style="Heading 2")
        _add_tabbed_run(p, text, size)
    for text, size, align in real:
        p = doc.add_paragraph(style="Heading 2")
        p.paragraph_format.alignment = align
        _set_run(p.add_run(text), "宋体", size)
    doc.save(str(path))
    return path


def test_heading_ignores_manual_toc_entries(tmp_path):
    """EZI-0041 A1/A3：``1.1.\\t编写目的\\t4``（制表符 + 页码）是手写目录条目，不是二级标题。

    EZI-0038 不再生成 TOC 域后，这些条目既没有 ``TOC2`` 样式也没有 ``PAGEREF``，于是
    落进二级标题样本池：功能设计/安装手册 14、15 条 jc=start 的条目（16pt）把「二级标题
    对齐」从 N-A 判成 FAIL（模板无编号二级标题样本，退回 spec both）。排掉条目后只剩真
    标题（16pt/both）→ 三项都按真标题判。
    """
    p = _build_manual_toc_headings(
        tmp_path / "manual_toc.docx",
        entries=[("1.1.\t编写目的\t4", 16.0), ("1.2.\t文档范围\t5", 16.0)],
        real=[("1.1 编写目的", 16.0, WD_ALIGN_PARAGRAPH.JUSTIFY)])
    rep = evaluate_layout_fidelity(p, p, spec=_HEADING2_SPEC, source=p)
    for name in ("二级标题字号(pt)", "二级标题对齐(jc)"):
        it = _named(rep, name)
        assert it.verdict == PASS, (name, it.verdict, it.actual, it.note)
        assert "手写目录条目 2" in it.note, it.note
    assert _named(rep, "二级标题对齐(jc)").actual == "both"


def test_heading_na_when_only_manual_toc_entries(tmp_path):
    """EZI-0041 A3：样本全是手写目录条目 → N-A（不是 FAIL，也不再拿条目算字号/对齐）。"""
    p = _build_manual_toc_headings(
        tmp_path / "only_toc.docx", entries=[("2.2.1.\t准备工作\t4", 16.0)])
    it = _named(evaluate_layout_fidelity(p, p, spec=_HEADING2_SPEC), "二级标题对齐(jc)")
    assert it.verdict == NA, (it.verdict, it.actual, it.note)
    assert it.expected is None
    assert "手写目录条目 1" in it.note, it.note


def test_real_numbered_heading_without_page_number_is_still_a_sample(tmp_path):
    """EZI-0041 A4 反例：``1.1 编写目的``（空格分隔、无页码）不是目录条目。

    它必须仍进样本：真·对齐错误（left vs spec both）照判 FAIL，不许被「排目录条目」
    这条规则顺手放过。
    """
    p = _build_manual_toc_headings(
        tmp_path / "real_bad.docx",
        entries=[("1.2.\t文档范围\t5", 16.0)],
        real=[("1.1 编写目的", 16.0, WD_ALIGN_PARAGRAPH.LEFT)])
    it = _named(evaluate_layout_fidelity(p, p, spec=_HEADING2_SPEC), "二级标题对齐(jc)")
    assert it.verdict == FAIL, (it.verdict, it.actual, it.note)
    assert it.actual == "start", it.actual
    assert "编号二级标题样本 1 个" in it.note, it.note


def test_body_size_ignores_manual_toc_entries_and_listed_headings(tmp_path):
    """EZI-0041 A2：正文取样排掉手写目录条目**及其列出的标题**。

    可行研形态：手写目录 2 条 + 正文流里同名的 2 段标题（16pt，没有 ``outlineLvl``）
    + 3 段真正文（10.5pt）。不排的话 16pt 有 4 票压过 3 票，产物照模板排 10.5 却被判
    FAIL（实测：57 段里 26 段 16pt 把众数顶到 16.0）。
    """
    doc = Document()
    for text in ("1.1.\t编写目的\t4", "1.2.\t文档范围\t4"):
        p = doc.add_paragraph()
        _add_tabbed_run(p, text, 16.0)
    for text in ("编写目的", "文档范围"):
        p = doc.add_paragraph()
        _set_run(p.add_run(text), "宋体", 16.0)
    for text in ("说明文档的编写目的", "说明文档的主要内容", "略",
                 "补充正文段一", "补充正文段二"):
        p = doc.add_paragraph()
        _set_run(p.add_run(text), "宋体", 10.5)
    path = tmp_path / "body_manual_toc.docx"
    doc.save(str(path))

    spec = {"structure": {"sections": []},
            "rules": [{"selector": {"semantic_role": "body"},
                       "expected": {"font_size_pt": 10.5}}]}
    it = _named(evaluate_layout_fidelity(path, path, spec=spec), "正文主字号(pt)")
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.actual == 10.5, it.actual
    assert "正文段落 5 个" in it.note, it.note


def test_manual_toc_entry_guard_skips_real_toc_field(tmp_path):
    """EZI-0041 A1 守卫：真 TOC 域内的段落不算「手写目录条目」（域内条目另有 TOC 样式口径）。"""
    from earlyz.validation import layout_fidelity as lf

    doc = Document()
    inside = doc.add_paragraph()
    _add_field(inside, r'TOC \o "1-2" \h', text="1.1.\t编写目的\t4")
    plain = doc.add_paragraph()
    _add_tabbed_run(plain, "1.1.\t编写目的\t4", 10.5)
    path = tmp_path / "toc_field.docx"
    doc.save(str(path))

    paras = lf.document_paragraphs(lf._load_docx(path))
    flags = [lf._is_manual_toc_entry(p) for p, _props in paras]
    assert flags == [False, True], flags


# ───────────────── A7：rule.expected 不是 Mapping 不许把整组打分打崩 ─────────────────


def test_non_mapping_rule_expected_is_skipped_not_crashing(docs):
    """A7：``rule.expected=True``（bool）时 ``dict(True)`` 抛 TypeError，整组打分崩。

    修法口径：只有 ``Mapping`` 才收成规则；bool/None/其它**整条跳过**该规则
    （不 ``dict(bool)``，也不塞默认值），同一角色的合法规则仍然生效。
    """
    class Rule:
        def __init__(self, role, expected):
            self.selector = {"semantic_role": role}
            self.expected = expected

    class Spec:
        structure = {"sections": []}
        rules = [Rule("heading1", True), Rule("body", {"font_size_pt": 12.0})]

    rep = evaluate_layout_fidelity(docs["ok"], docs["template"], spec=Spec())
    it = _by_name(rep, "一级标题字号(pt)")
    assert it.verdict == NA and "无 heading1 规则" in it.note, (it.verdict, it.note)
    assert _by_name(rep, "正文主字号(pt)").verdict == PASS     # 合法规则照常判
    assert rep.failed == 0

    class Spec2:                     # 同一角色：先非法后合法 → 用合法那条
        structure = {"sections": []}
        rules = [Rule("heading1", True), Rule("heading1", {"font_size_pt": 16.0})]

    rep2 = evaluate_layout_fidelity(docs["ok"], docs["template"], spec=Spec2())
    assert _by_name(rep2, "一级标题字号(pt)").verdict == PASS

    # dict 形态的 spec 走另一个分支，同样不许 dict(bool)
    rep3 = evaluate_layout_fidelity(docs["ok"], docs["template"], spec={
        "structure": {"sections": []},
        "rules": [{"selector": {"semantic_role": "heading1"}, "expected": True},
                  {"selector": {"semantic_role": "body"}, "expected": {"font_size_pt": 12.0}}],
    })
    it3 = _by_name(rep3, "一级标题字号(pt)")
    assert it3.verdict == NA and "无 heading1 规则" in it3.note, (it3.verdict, it3.note)
    assert _by_name(rep3, "正文主字号(pt)").verdict == PASS


# ────────────────────────────── A3：JSON 与退出码 ──────────────────────────────


def test_json_payload_shape(docs, capsys):
    cli = _load_cli()
    code = cli.main(["check_layout_fidelity.py", str(docs["ok"]), str(docs["template"]), "--json"])
    out = capsys.readouterr().out
    payload = json.loads(out)
    assert code == 0 and payload["exit_code"] == 0
    for key in ("passed", "total", "score", "items"):
        assert key in payload, key
    assert payload["passed"] == payload["total"] - payload["not_applicable"]
    assert payload["score"] == pytest.approx(
        payload["passed"] / (payload["total"] - payload["not_applicable"]), abs=1e-4)
    assert len(payload["items"]) == payload["total"]
    for item in payload["items"]:
        assert {"name", "expected", "actual", "verdict"} <= set(item)
        assert item["verdict"] in (PASS, FAIL, NA)


def test_cli_exit_codes(docs, capsys):
    cli = _load_cli()
    assert cli.main(["x", str(docs["ok"]), str(docs["template"])]) == 0
    assert cli.main(["x", str(docs["bad"]), str(docs["template"])]) == 1
    assert cli.main(["x", str(docs["ok"])]) == 2          # 少参数 = 用法错误
    assert cli.main(["x", "不存在.docx", str(docs["template"])]) == 2
    capsys.readouterr()


def test_cli_source_flag_drives_toc_verdict(docs, capsys):
    """``--source X`` 与 ``--source=X`` 等价；不给就跟「无源」一样 N-A。"""
    cli = _load_cli()
    base = ["x", str(docs["bad"]), str(docs["template"])]

    assert cli.main(base) == 1                       # 其它项仍在 FAIL
    out = capsys.readouterr().out
    assert "[N-A ] 目录域(TOC)" in out
    assert "未提供源文档" in out

    assert cli.main(base + ["--source", str(docs["source_toc"])]) == 1
    out = capsys.readouterr().out
    assert "[FAIL] 目录域(TOC)" in out
    assert "缺目录域" in out

    assert cli.main(base + [f"--source={docs['source_no_toc']}"]) == 1
    out = capsys.readouterr().out
    assert "[PASS] 目录域(TOC)" in out          # EZI-0044：源无 TOC + 产物 0 个 = PASS
    assert "源文档无目录域" in out

    assert cli.main(base + ["--source"]) == 2        # 缺路径 = 用法错误
    capsys.readouterr()


def test_json_payload_carries_source_and_toc_verdict(docs, capsys):
    cli = _load_cli()
    code = cli.main(["x", str(docs["bad"]), str(docs["template"]),
                     "--source", str(docs["source_toc"]), "--json"])
    payload = json.loads(capsys.readouterr().out)
    assert code == 1 and payload["exit_code"] == 1
    assert payload["source"] == str(docs["source_toc"])
    toc = next(i for i in payload["items"] if i["name"] == "目录域(TOC)")
    assert toc["verdict"] == FAIL


def test_table_output_lists_every_item(reports, docs, capsys):
    cli = _load_cli()
    assert cli.main(["x", str(docs["ok"]), str(docs["template"])]) == 0
    out = capsys.readouterr().out
    for it in reports["ok"].items:
        assert it.name in out
    assert "正确率" in out


# ────────────────────────────── A5：只读 ──────────────────────────────


def test_evaluate_is_read_only(docs):
    prod = docs["bad"]
    before = hashlib.sha256(prod.read_bytes()).hexdigest(), prod.stat().st_mtime_ns
    evaluate_layout_fidelity(prod, docs["template"])
    after = hashlib.sha256(prod.read_bytes()).hexdigest(), prod.stat().st_mtime_ns
    assert before == after


def test_format_table_smoke(reports):
    text = format_table(reports["bad"])
    assert "[FAIL]" in text and "正确率" in text


# ───────────── A1/A2（EZI-0011）：图题绑图 / 表题绑表（方向按对象位置）─────────────

#: 1×1 PNG（只为造一个真实 `w:drawing`，内容不重要）
_PNG_1PX = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFAAH/"
    "q842iQAAAABJRU5ErkJggg=="
)


def _set_keep_next(p, on: bool = True) -> None:
    p.paragraph_format.keep_with_next = bool(on)


def _add_picture_paragraph(doc):
    """加一个只含图片的段落（= 「图片段落」，引擎要在这里加 keepNext）。"""
    p = doc.add_paragraph()
    p.add_run().add_picture(io.BytesIO(base64.b64decode(_PNG_1PX)), width=Pt(20))
    return p


def _caption_report(path: Path):
    """只判题注绑定：给空 spec（其它项一律 N-A），不依赖模板编译。"""
    return evaluate_layout_fidelity(
        path, path, spec={"structure": {"sections": []}, "rules": []})


def _build_caption_docx(path: Path, *, figure_keep_next: bool = True,
                        caption_keep_next: bool = False,
                        table_caption_keep_next: bool = True,
                        blank_before_table: bool = False,
                        table_before_caption: bool = False,
                        figure: bool = True, table: bool = True):
    """造一份只含「图 + 图题」「表题 + 表」的最小文档，覆盖绑定/不绑定/不相邻。

    返回 ``Document``（已经落盘到 ``path``），方便调用方继续加段落。
    """
    doc = Document()
    doc.add_paragraph("正文一段，用来验证题注之外的段落不受影响。")
    if figure:
        fig = _add_picture_paragraph(doc)
        if figure_keep_next:
            _set_keep_next(fig)
        cap = doc.add_paragraph("图3.1 学生用例图")
        if caption_keep_next:
            _set_keep_next(cap)
    if table:
        if not table_before_caption:
            tcap = doc.add_paragraph("表4.1 用户表")
            if table_caption_keep_next:
                _set_keep_next(tcap)
            if blank_before_table:
                doc.add_paragraph("")
            doc.add_table(rows=2, cols=2)
        else:
            doc.add_table(rows=2, cols=2)
            tcap = doc.add_paragraph("表4.1 用户表")
            if table_caption_keep_next:
                _set_keep_next(tcap)
    doc.save(str(path))
    return doc


def _caption_item(rep, name):
    return next(i for i in rep.items if i.name == name)


def test_figure_caption_passes_when_image_paragraph_has_keep_next(tmp_path):
    """A1：图题紧凑跟在前面的图片段落、且图段带 keepNext → PASS。"""
    path = tmp_path / "fig_ok.docx"
    _build_caption_docx(path, figure_keep_next=True)
    it = _caption_item(_caption_report(path), "图题绑定")
    assert it.verdict == PASS, (it.verdict, it.actual, it.note)
    assert it.actual.startswith("PASS 1 / FAIL 0"), it.actual
    assert it.expected is not None        # 有期望值才算「判过」，不是 N-A 蒙混
    assert it.category == "题注绑定"


def test_figure_caption_fails_when_image_paragraph_lacks_keep_next(tmp_path):
    """A1 反例：图段没有 keepNext（图与图题之间没有绑定）→ FAIL。"""
    path = tmp_path / "fig_bad.docx"
    _build_caption_docx(path, figure_keep_next=False)
    it = _caption_item(_caption_report(path), "图题绑定")
    assert it.verdict == FAIL, (it.verdict, it.actual, it.note)
    assert "FAIL 1" in it.actual
    assert "图3.1" in it.note          # FAIL 例要点名，便于定位


def test_figure_caption_keep_next_on_caption_itself_is_not_a_pass(tmp_path):
    """方向错反例：只给**图题段**加 keepNext（旧引擎的做法）→ 仍然 FAIL。

    keepNext 绑的是「自己 + 后面内容」；图题后面是正文，跟图没关系。
    """
    path = tmp_path / "fig_wrong_dir.docx"
    _build_caption_docx(path, figure_keep_next=False, caption_keep_next=True)
    it = _caption_item(_caption_report(path), "图题绑定")
    assert it.verdict == FAIL, (it.verdict, it.actual, it.note)


def test_table_caption_passes_when_bound_and_adjacent(tmp_path):
    """A1：表题紧跟表格且表题段带 keepNext → PASS。"""
    path = tmp_path / "tab_ok.docx"
    _build_caption_docx(path, table_caption_keep_next=True)
    it = _caption_item(_caption_report(path), "表题绑定")
    assert it.verdict == PASS, (it.verdict, it.actual, it.note)
    assert it.actual.startswith("PASS 1 / FAIL 0"), it.actual


def test_table_caption_fails_when_missing_keep_next(tmp_path):
    """A1 反例：表题段没有 keepNext → FAIL（题与表会被分页拆开）。"""
    path = tmp_path / "tab_bad.docx"
    _build_caption_docx(path, table_caption_keep_next=False)
    it = _caption_item(_caption_report(path), "表题绑定")
    assert it.verdict == FAIL, (it.verdict, it.actual, it.note)
    assert "表4.1" in it.note


def test_table_caption_na_when_blank_paragraph_between(tmp_path):
    """A2：题与表之间隔了空段 → N-A（不猜），note 写明原因。"""
    path = tmp_path / "tab_blank.docx"
    _build_caption_docx(path, table_caption_keep_next=False, blank_before_table=True)
    it = _caption_item(_caption_report(path), "表题绑定")
    assert it.verdict == NA, (it.verdict, it.actual, it.note)
    assert it.expected is None, "判 N-A 时不该出现期望值"
    assert "空段" in it.note and "N-A" in it.actual


def test_table_caption_above_table_is_na_not_guessed(tmp_path):
    """A2：表格在题**前**（对象位置不对）→ N-A，别把「题在表下方」判成绑表。"""
    path = tmp_path / "tab_above.docx"
    _build_caption_docx(path, table_before_caption=True, table_caption_keep_next=True)
    it = _caption_item(_caption_report(path), "表题绑定")
    assert it.verdict == NA, (it.verdict, it.actual, it.note)
    assert it.expected is None
    assert "N-A 原因" in it.note and "无可判定题注" in it.note


def test_figure_caption_na_when_no_adjacent_image(tmp_path):
    """A2：图题前面不是图片段落 → N-A（隔正文/无图），不许判 FAIL 也不许判 PASS。"""
    doc = _build_caption_docx(tmp_path / "fig_na.docx", figure=False, table=False)
    doc.add_paragraph("正文里提到图2.1，但这里不是题注。")
    doc.add_paragraph("图2.1 只有题注没有图")      # 前段是正文 → 无法判定
    doc.save(str(tmp_path / "fig_na.docx"))
    it = _caption_item(_caption_report(tmp_path / "fig_na.docx"), "图题绑定")
    assert it.verdict == NA, (it.verdict, it.actual, it.note)
    assert it.expected is None
    assert "N-A" in it.actual and "他段" in it.note


def test_mixed_binding_fails_aggregate(tmp_path):
    """聚合口径：一好一坏 → 整体 FAIL（不许拿多数 PASS 盖住真实缺陷）。"""
    doc = _build_caption_docx(tmp_path / "fig_mixed.docx", figure_keep_next=True, table=False)
    _add_picture_paragraph(doc)                   # 第二张图没有 keepNext
    doc.add_paragraph("图3.2 第二张用例图")
    doc.save(str(tmp_path / "fig_mixed.docx"))
    it = _caption_item(_caption_report(tmp_path / "fig_mixed.docx"), "图题绑定")
    assert it.verdict == FAIL, (it.verdict, it.actual, it.note)
    assert "PASS 1 / FAIL 1" in it.actual


def test_caption_text_inside_graphic_textbox_is_not_judged(tmp_path):
    """题注字写在图形文本框里（``w:txbxContent``）→ 不算流内题注，N-A 说明清楚。

    真实样本（``output/毕业设计.docx``）30 张图的题注就是这样：按 ``p.iter(w:t)``
    取字会把图对象当题注段，方向判反。这里用最小 VML 文本框复现该结构。
    """
    doc = Document()
    p = doc.add_paragraph()
    pict = OxmlElement("w:pict")
    txbx = OxmlElement("w:txbxContent")
    inner = OxmlElement("w:p")
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "图5.1 藏在文本框里的题注"
    r.append(t)
    inner.append(r)
    txbx.append(inner)
    pict.append(txbx)
    p._p.append(pict)
    doc.save(str(tmp_path / "fig_txbx.docx"))

    it = _caption_item(_caption_report(tmp_path / "fig_txbx.docx"), "图题绑定")
    assert it.verdict == NA, (it.verdict, it.actual, it.note)
    assert "流内图题为 0" in it.note
    assert "1 处题注字写在图对象文本框里" in it.note, it.note


# ───────── A4（EZI-0011）：引擎动作 bind_caption_objects（放在本文件，因 write_scope）─────────
#
# 说明：本 issue 的 write_scope 只允许改这一个 tests 文件，所以引擎动作的回归也放这里
# ——它守的正是上面两项判定的口径（判定说 FAIL，动作必须能把 FAIL 修成 PASS）。


def _bind(path, **params):
    from earlyz.actions.captions import bind_caption_objects
    return bind_caption_objects(path, {"apply": True, **params})


def test_action_binds_figure_paragraph_and_unbinds_caption(tmp_path):
    """图题方向修正：keepNext 长在**图片段落**上，图题段自己的 keepNext 要去掉。"""
    path = tmp_path / "fix_fig.docx"
    _build_caption_docx(path, figure_keep_next=False, caption_keep_next=True, table=False)
    before = _caption_item(_caption_report(path), "图题绑定")
    assert before.verdict == FAIL          # 方向错：题注自己有 keepNext、图段没有

    res = _bind(path)
    assert res["figures_bound"] == 1
    assert res["keep_next_added"] == 1
    assert res["caption_keepnext_removed"] == 1
    after = _caption_item(_caption_report(path), "图题绑定")
    assert after.verdict == PASS, (after.verdict, after.actual, after.note)

    from earlyz.actions.captions import _flow_blocks, _p_has
    from lxml import etree
    import zipfile
    root = etree.fromstring(zipfile.ZipFile(path).read("word/document.xml"))
    blocks = _flow_blocks(root.find(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body"))
    fig = next(b for b in blocks if b.find(
        ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing") is not None)
    cap = next(b for b in blocks if "图3.1" in "".join(
        t.text or "" for t in b.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")))
    assert _p_has(fig, "keepNext") and _p_has(fig, "keepLines")
    assert not _p_has(cap, "keepNext")


def test_action_binds_table_caption_and_drops_blank_gap(tmp_path):
    """表题：补 keepNext，并拆掉题与表之间的纯空段（空段会断 keepNext 链）。"""
    path = tmp_path / "fix_tab.docx"
    _build_caption_docx(path, table_caption_keep_next=False, blank_before_table=True,
                        figure=False)
    assert _caption_item(_caption_report(path), "表题绑定").verdict == NA   # 隔空段 → 不判

    res = _bind(path)
    assert res["tables_bound"] == 1
    assert res["blank_paras_removed"] == 1
    after = _caption_item(_caption_report(path), "表题绑定")
    assert after.verdict == PASS, (after.verdict, after.actual, after.note)


def test_action_keeps_cant_split_and_tbl_header_and_is_idempotent(tmp_path):
    """A4：cantSplit / tblHeader 一个都不许少；动作幂等（二次运行 changed=False）。"""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    path = tmp_path / "fix_keep_flags.docx"
    doc = _build_caption_docx(path, table_caption_keep_next=False, figure=False)
    tbl = doc.tables[0]
    for row in tbl.rows:                    # 模拟引擎早先写好的行属性
        trpr = row._tr.get_or_add_trPr()
        trpr.append(OxmlElement("w:cantSplit"))
    tbl.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
    doc.save(str(path))

    res = _bind(path)
    assert res["table_flags_preserved"] is True, (res["table_flags_before"], res["table_flags_after"])
    assert res["table_flags_before"]["cant_split"] == 2
    assert res["table_flags_before"]["tbl_header"] == 1
    assert res["table_flags_after"] == res["table_flags_before"]

    again = _bind(path)
    assert again["changed"] is False
    assert again["table_flags_preserved"] is True
    assert again["keep_next_added"] == 0 and again["blank_paras_removed"] == 0


def test_action_noop_on_document_without_inflow_captions(tmp_path):
    """没有流内题注（题注字全在图对象文本框里）→ 不动盘，如实报 0。"""
    doc = Document()
    p = doc.add_paragraph()
    pict = OxmlElement("w:pict")
    txbx = OxmlElement("w:txbxContent")
    inner = OxmlElement("w:p")
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "图7.1 文本框里的题注"
    r.append(t)
    inner.append(r)
    txbx.append(inner)
    pict.append(txbx)
    p._p.append(pict)
    path = tmp_path / "noop.docx"
    doc.save(str(path))

    before = path.read_bytes()
    res = _bind(path)
    assert res["changed"] is False
    assert res["figures"] == 0 and res["tables"] == 0
    assert res["embedded_figure_captions"] == 1
    assert path.read_bytes() == before          # 幂等且没改盘


# ───────── A2/A3/A5（EZI-0015）：正文分栏逐节取末节 + 行距两侧同口径 ─────────


def _set_section_cols(doc, index: int, num: int) -> None:
    """给第 ``index`` 节写 ``w:cols/@w:num``（python-docx 不暴露分栏属性）。"""
    sect_pr = doc.sections[index]._sectPr
    cols = sect_pr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        grid = sect_pr.find(qn("w:docGrid"))
        if grid is not None:
            grid.addprevious(cols)      # schema 顺序：w:cols 在 w:docGrid 之前
        else:
            sect_pr.append(cols)
    cols.set(qn("w:num"), str(num))


def _build_cols_docx(path: Path, cols_seq: list[int]) -> Path:
    """造 ``len(cols_seq)`` 节、逐节写 cols/@w:num 的 docx（正文段落落在末节前）。"""
    doc = Document()
    for i, num in enumerate(cols_seq):
        if i:
            doc.add_section()
        _set_section_cols(doc, i, num)
    for j in range(3):
        doc.add_paragraph(f"第 {j + 1} 段正文，用来验证分栏判定的取样。")
    doc.save(str(path))
    return path


def _strip_line_attrs(docx_path: Path, out_path: Path,
                      *, doc_defaults_only: bool = False) -> Path:
    """重写 docx，删掉 ``w:spacing/@w:line``（造「模板未设固定行距」样本）。

    ``doc_defaults_only=True`` 只删 docDefaults 里的行距——段落没写 ``w:pStyle`` 时
    ``_ppr_chain`` 只拿到 docDefaults，这样「未设行距」的段落实测才会是 None。
    """
    import zipfile

    from lxml import etree

    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    with zipfile.ZipFile(docx_path) as src:
        items = [(info, src.read(info.filename)) for info in src.infolist()]
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as dst:
        for info, data in items:
            if info.filename in ("word/document.xml", "word/styles.xml"):
                root = etree.fromstring(data)
                scopes = ([root.find(W + "docDefaults")] if doc_defaults_only
                          else [root])
                for scope in scopes:
                    if scope is None:
                        continue
                    for sp in scope.iter(W + "spacing"):
                        for attr in ("line", "lineRule"):
                            if sp.get(W + attr) is not None:
                                del sp.attrib[W + attr]
                data = etree.tostring(root, xml_declaration=True,
                                      encoding="UTF-8", standalone=True)
            dst.writestr(info, data)
    return out_path


def test_body_cols_compares_last_section_not_first_multicol_section(tmp_path):
    """A2 回归（ACM 形态）：两侧 cols 序列完全一样 → 必须 PASS。

    旧判法拿模板 ``structure.body_cols``（「第一个非首节 ≥2 栏」= 索引 1 的 3 栏
    作者区）去比产物 ``body/w:sectPr``（末节 2 栏），比的是两个不同的节，判成
    「模板 3 产物 2」FAIL。这里复刻 ACM 的 ``[1,3,1,2]``：两侧一致，必须 PASS。
    """
    tpl = _build_cols_docx(tmp_path / "tpl_cols.docx", [1, 3, 1, 2])
    prod = _build_cols_docx(tmp_path / "prod_cols.docx", [1, 3, 1, 2])
    it = _named(evaluate_layout_fidelity(prod, tpl), "正文分栏数")
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert (it.expected, it.actual) == ("2", "2"), (it.expected, it.actual)
    assert "[1,3,1,2]" in it.note, it.note
    # EZI-0069 措辞迁移：取末节 → 正文节配对（ACM 形态下正文节=第4节，
    # 与旧末节口径同值 2=2，行为不变）
    assert "正文节配对" in it.note, it.note


def test_body_cols_sequence_order_differs_but_body_section_matches(tmp_path):
    """A5：多节且 cols 顺序不同，但正文末节一致 → PASS（note 说明取末节）。

    模板 ``[1,3,1,2]``，产物 ``[1,1,3,2]``：3 栏节的位置不同（节骨架由「分节数」
    信息项承载），正文末节都是 2 栏 → 分栏判 PASS。
    """
    tpl = _build_cols_docx(tmp_path / "tpl_cols_order.docx", [1, 3, 1, 2])
    prod = _build_cols_docx(tmp_path / "prod_cols_order.docx", [1, 1, 3, 2])
    it = _named(evaluate_layout_fidelity(prod, tpl), "正文分栏数")
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert (it.expected, it.actual) == ("2", "2")
    assert "前序节不同不影响正文分栏" in it.note, it.note


def test_body_cols_fails_when_body_section_really_differs(tmp_path):
    """反向断言：末节真的不同（模板 2 栏、产物 1 栏）→ FAIL，不许被新口径放过。"""
    tpl = _build_cols_docx(tmp_path / "tpl_cols_2.docx", [1, 3, 2])
    prod = _build_cols_docx(tmp_path / "prod_cols_1.docx", [1, 3, 1])
    it = _named(evaluate_layout_fidelity(prod, tpl), "正文分栏数")
    assert it.verdict == FAIL, (it.verdict, it.expected, it.actual, it.note)
    # EZI-0070：全文档段量最大（平票取晚）——均匀夹具下两侧正文节均取末节
    # → act 回 1，真不同仍 FAIL 的反向语义不变
    assert (it.expected, it.actual) == ("2", "1")


def test_line_spacing_na_when_template_sets_no_line(tmp_path):
    """A3/A5：模板正文没有任何 w:line（含默认段落样式）→ N-A，不凭空造期望值。"""
    tpl = _strip_line_attrs(build_docx(tmp_path / "tpl_noline.docx"),
                            tmp_path / "tpl_noline_stripped.docx")
    prod = build_docx(tmp_path / "prod_line.docx")
    it = _named(evaluate_layout_fidelity(prod, tpl), "正文行距(line/lineRule)")
    assert it.verdict == NA, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected is None, it.expected
    assert it.actual is not None, "N-A 也要给产物实测值"
    assert "模板正文未设固定行距" in it.note, it.note


def test_line_spacing_mode_ignores_unset_paragraphs(tmp_path):
    """A3 回归（IEEE 形态）：多数正文段落没写行距，产物仍应判 PASS。

    旧实现把 None 也算进众数：IEEE 产物 115/218 段未设、103 段 240/exact →
    众数成了 None，actual=None，判 FAIL。这里造「3 段设 240/auto + 5 段未设」，
    众数只算设值段落 → 240/auto，与模板一致。
    """
    tpl = _strip_line_attrs(build_docx(tmp_path / "tpl_240.docx", line_spacing=1.0),
                            tmp_path / "tpl_240_stripped.docx",
                            doc_defaults_only=True)
    prod_path = build_docx(tmp_path / "prod_mixed.docx", line_spacing=1.0)
    doc = Document(str(prod_path))
    for j in range(5):                     # 多数段落不写行距（继承 docDefaults）
        doc.add_paragraph(f"补充段落 {j + 1}，没有直接写行距。")
    doc.save(str(prod_path))
    prod = _strip_line_attrs(prod_path, tmp_path / "prod_mixed_stripped.docx",
                             doc_defaults_only=True)

    it = _named(evaluate_layout_fidelity(prod, tpl), "正文行距(line/lineRule)")
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.actual.startswith("240/auto"), it.actual
    assert "未设×" in it.note and "两侧同口径" in it.note, it.note


# ────── A1/A2/A4（EZI-0020）：正文取样排掉表单标签（交大中期行距误判）──────


def _set_para_line(paragraph, line: int, rule: str) -> None:
    """给段落写 ``w:spacing/@w:line`` + ``@w:lineRule``（python-docx 只暴露倍数）。"""
    ppr = paragraph._p.get_or_add_pPr()
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    spacing.set(qn("w:line"), str(line))
    spacing.set(qn("w:lineRule"), rule)


_FORM_LABELS = ("学生姓名:", "学生学号:", "学院 （系）:")
_FORM_BODIES = tuple(f"第 {i} 段正文，用于验证正文行距取样排掉表单标签。" for i in (1, 2, 3, 4, 5))


def _build_form_docx(path: Path, *, labels=_FORM_LABELS, label_line=(500, "exact"),
                     label_size: float = 16.0, bodies=_FORM_BODIES,
                     body_line=(400, "atLeast"), body_size: float = 16.0) -> Path:
    """按「先表单标签、后真正文」的文档序造 docx（交大中期封面-正文形态）。

    标签与正文各写自己的行距/字号：标签段是字段名（短行 + 冒号），正文段是长段落。
    """
    doc = Document()
    for text in labels:
        p = doc.add_paragraph()
        _set_run(p.add_run(text), "宋体", label_size)
        _set_para_line(p, *label_line)
    for text in bodies:
        p = doc.add_paragraph()
        _set_run(p.add_run(text), "宋体", body_size)
        _set_para_line(p, *body_line)
    doc.save(str(path))
    return path


def test_line_spacing_mode_ignores_form_labels(tmp_path):
    """A1/A2 回归（交大中期形态）：表单标签与真正文打平后，只按真正文判。

    旧口径把「学生姓名:」这类字段名也算正文：模板 500/exact×3（标签）与
    400/atLeast×3（真正文）打平，``most_common`` 按文档序取到 500/exact；产物只剩
    2 段标签（EZI-0017 引擎修完后真正文已是 400/atLeast×3），产物众数 400/atLeast
    → 判 FAIL。排掉标签后模板 400/atLeast×3 = 产物 → PASS。
    """
    tpl = _build_form_docx(tmp_path / "tpl_form.docx")
    prod = _build_form_docx(tmp_path / "prod_form.docx", labels=_FORM_LABELS[:2])
    it = _named(evaluate_layout_fidelity(prod, tpl), "正文行距(line/lineRule)")
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected.startswith("400/atLeast"), it.expected
    assert it.actual.startswith("400/atLeast"), it.actual
    assert "表单标签" in it.note, it.note
    assert "500/exact" not in it.note, it.note     # 字段名不进样本分布
    assert "（正文 5 段）" in it.note, it.note      # 真正文一段不少


def test_line_spacing_keeps_long_colon_ended_paragraph(tmp_path):
    """A4 反向：以冒号结尾的**长**段仍是正文——「短行」才是表单标签的判据。

    只认「≤ ``_FORM_LABEL_MAX`` 字符 + 冒号结尾」：长段落即使收尾是冒号也不许排掉，
    短标签即使没有冒号也不会被这条规则误伤（另一条正向样本在下面）。
    """
    long_colon = ("请每位学生根据学校及院（系）检查的要求认真进行自查，"
                  "并及时发现课题研究过程中存在的问题：")
    assert len(long_colon) > 24
    bodies = (long_colon, "第 1 段正文。", "第 2 段正文。", "第 3 段正文。", "第 4 段正文。")
    tpl = _build_form_docx(tmp_path / "tpl_long.docx", labels=("学生姓名:",),
                           label_line=(300, "exact"), bodies=bodies,
                           body_line=(240, "auto"))
    prod = _build_form_docx(tmp_path / "prod_long.docx", labels=("学生姓名:",),
                            label_line=(300, "exact"), bodies=bodies,
                            body_line=(240, "auto"))
    it = _named(evaluate_layout_fidelity(prod, tpl), "正文行距(line/lineRule)")
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected.startswith("240/auto"), it.expected
    assert "300/exact" not in it.note, it.note      # 短标签（300/exact）被排掉
    assert "（正文 5 段）" in it.note, it.note       # 长冒号段留着


def test_body_font_size_sample_still_includes_form_labels_for_now(tmp_path):
    """口径边界（EZI-0020 决策记录）：正文四项暂用旧样本，标签仍计入正文段落数。

    行距取样排标签；正文主字号的主峰阈值（60%）是按含标签的旧样本校准的，同一处去掉
    标签会让本科论文 12pt（39/68=57%）从单峰掉进无主峰分支（EZI-0022 后候选集合由
    「12pt」变「12pt/10.5pt」）——那是另一件事，本 issue 不动。哪天把字号也改成排标签，
    这个断言要跟着改（说明决策变了），不许悄悄变。
    """
    spec = {"structure": {"sections": []},
            "rules": [{"selector": {"semantic_role": "body"},
                       "expected": {"font_size_pt": 12.0}}]}
    tpl = _build_form_docx(tmp_path / "tpl_sz.docx", label_size=14.0, body_size=12.0,
                           label_line=(360, "auto"), body_line=(360, "auto"))
    prod = _build_form_docx(tmp_path / "prod_sz.docx", label_size=14.0, body_size=12.0,
                            label_line=(360, "auto"), body_line=(360, "auto"))
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=spec), "正文主字号(pt)")
    assert "正文段落 8 个" in it.note, it.note      # 3 标签 + 5 正文（旧样本口径）


# ────── A1/A2（EZI-0016）：页边距逐节比 / 正文主字号与标题期望取模板实测 ──────

_MARGIN_COVER = (1417, 1701, 1134, 1701)      # 封面/摘要那套（本科模板节0~2）
_MARGIN_BODY = (1440, 1800, 1440, 1800)       # 正文那套（本科模板节3~7）


def _set_section_margins(doc, index: int, margins) -> None:
    sec = doc.sections[index]
    top, right, bottom, left = margins
    sec.top_margin, sec.right_margin = Twips(top), Twips(right)
    sec.bottom_margin, sec.left_margin = Twips(bottom), Twips(left)


def _build_margin_docx(path: Path, margins_seq, *, with_pgmar: bool = True) -> Path:
    """造 ``len(margins_seq)`` 节、逐节写 ``w:pgMar`` 的 docx（各节边距可不同）。"""
    doc = Document()
    for i, margins in enumerate(margins_seq):
        if i:
            doc.add_section()
        _set_section_margins(doc, i, margins)
    for j in range(3):
        doc.add_paragraph(f"第 {j + 1} 段正文，用来验证逐节页边距比对。")
    if not with_pgmar:                       # 造「模板没有 pgMar」→ 该项 N-A
        for sec in doc.sections:
            mar = sec._sectPr.find(qn("w:pgMar"))
            if mar is not None:
                sec._sectPr.remove(mar)
    doc.save(str(path))
    return path


def _margin_item(product, template):
    return _named(evaluate_layout_fidelity(product, template), "页边距(top/right/bottom/left)")


def test_margin_compares_per_section_by_index(tmp_path):
    """A1 回归（本科形态）：模板两套边距（封面 ×3 + 正文 ×5），产物逐节复刻 → PASS。

    旧判法拿模板首节比产物末节：模板 1417/1701/1134/1701 vs 产物 1440/1800/1440/1800，
    报出来的是「两个不同节」的差；逐节对齐后两侧一致 → 必须 PASS。
    """
    tpl = _build_margin_docx(tmp_path / "tpl_margin.docx",
                             [_MARGIN_COVER] * 3 + [_MARGIN_BODY] * 5)
    prod = _build_margin_docx(tmp_path / "prod_margin.docx",
                              [_MARGIN_COVER] * 3 + [_MARGIN_BODY] * 5)
    it = _margin_item(prod, tpl)
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected.startswith("8 节："), it.expected
    assert "1417/1701/1134/1701 ×3" in it.expected, it.expected
    assert "1440/1800/1440/1800 ×5" in it.expected, it.expected
    assert "按索引逐节对齐" in it.note, it.note


def test_margin_fails_when_body_sections_use_cover_margins(tmp_path):
    """A1 反向（本科真形态）：产物 9 节里前 8 节都用封面边距 → 逐节比出节3~7 超差。

    旧判法只看「模板首节 vs 产物末节」，差异被藏起来；逐节对齐后产物节3~7（1417 组）
    vs 模板节3~7（1440 组）必须判 FAIL，并在 note 点名超差的节号与两侧取值。
    """
    tpl = _build_margin_docx(tmp_path / "tpl_margin8.docx",
                             [_MARGIN_COVER] * 3 + [_MARGIN_BODY] * 5)
    prod = _build_margin_docx(tmp_path / "prod_margin9.docx",
                              [_MARGIN_COVER] * 8 + [_MARGIN_BODY])
    it = _margin_item(prod, tpl)
    assert it.verdict == FAIL, (it.verdict, it.expected, it.actual, it.note)
    assert "9 节：" in it.actual and "×8" in it.actual, it.actual
    assert "超差节 3~7（5 节）" in it.note, it.note
    assert "节3 模板 1440/1800/1440/1800 vs 产物 1417/1701/1134/1701" in it.note, it.note
    assert "两侧节数不同（8 vs 9）" in it.note, it.note


def test_margin_detects_swapped_sections_that_single_value_would_miss(tmp_path):
    """A1 核心：两节的边距整套对调 → 单值比法（模板首节 vs 产物末节）会判 PASS。

    模板 [封面, 正文]，产物 [正文, 封面]：单值取「模板节0 = 1417 组」比「产物末节 =
    1417 组」→ 假 PASS；逐节对齐比出节0/节1 都错 → FAIL。
    """
    tpl = _build_margin_docx(tmp_path / "tpl_swap.docx", [_MARGIN_COVER, _MARGIN_BODY])
    prod = _build_margin_docx(tmp_path / "prod_swap.docx", [_MARGIN_BODY, _MARGIN_COVER])
    it = _margin_item(prod, tpl)
    assert it.verdict == FAIL, (it.verdict, it.expected, it.actual, it.note)
    assert "超差节 0~1（2 节）" in it.note, it.note


def test_margin_passes_when_only_section_count_differs(tmp_path):
    """A1：节数不同但能对齐的节都一致 → PASS（节骨架是信息项，不按节数判失败）。"""
    tpl = _build_margin_docx(tmp_path / "tpl_margin2.docx", [_MARGIN_COVER, _MARGIN_BODY])
    prod = _build_margin_docx(tmp_path / "prod_margin3.docx",
                              [_MARGIN_COVER, _MARGIN_BODY, _MARGIN_BODY])
    it = _margin_item(prod, tpl)
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert "两侧节数不同（2 vs 3）" in it.note, it.note


def test_margin_na_when_template_has_no_pgmar(tmp_path):
    """A1：模板没有 pgMar → N-A（无期望值，不猜）。"""
    tpl = _build_margin_docx(tmp_path / "tpl_nomar.docx", [_MARGIN_COVER], with_pgmar=False)
    prod = _build_margin_docx(tmp_path / "prod_margin.docx", [_MARGIN_COVER])
    it = _margin_item(prod, tpl)
    assert it.verdict == NA and it.expected is None, (it.verdict, it.expected, it.note)
    assert "模板无 sectPr/pgMar" in it.note, it.note


# ────── A1-A4（EZI-0027）：节数不等时按角色配对，不再机械比前 N 节 ──────
#
# 本科论文真形态：模板 8 节（封面/摘要/ABSTRACT 同一套封面边距 + 目录/正文/
# 参考文献/致谢/附录 同一套正文边距），产物 9 节（多一个「声明」节）。按索引比
# 前 8 节 → 产物节3 ABSTRACT(1417) 撞模板节3 目录(1440) → 边距全对也 FAIL。

_TPL_ROLE_SECTIONS = [
    ("封面示例：毕业设计（论文）", _MARGIN_COVER),
    ("摘  要", _MARGIN_COVER),
    ("ABSTRACT", _MARGIN_COVER),
    ("目  录", _MARGIN_BODY),
    ("1 绪论", _MARGIN_BODY),
    ("参考文献", _MARGIN_BODY),
    ("致谢", _MARGIN_BODY),
    ("附录", _MARGIN_BODY),
]


def _build_role_margin_docx(path: Path, sections, *, with_pgmar: bool = True) -> Path:
    """造「每节自带标题段」的 docx：``sections = [(标题文本, 边距), ...]``。

    角色判定只认节内标题段，而 ``_build_margin_docx`` 的正文段落全落在末节
    （其它节没有文字 → 认不出角色）。这里每节先写标题段再插分节符，让两侧都能
    按标题判出角色。
    """
    doc = Document()
    for i, (heading, _) in enumerate(sections):
        if i:
            doc.add_section()
        doc.add_paragraph(heading)
    for i, (_, margins) in enumerate(sections):
        _set_section_margins(doc, i, margins)
    if not with_pgmar:
        for sec in doc.sections:
            mar = sec._sectPr.find(qn("w:pgMar"))
            if mar is not None:
                sec._sectPr.remove(mar)
    doc.save(str(path))
    return path


def test_margin_pairs_by_role_when_product_has_extra_declaration_section(tmp_path):
    """A1/A2（EZI-0027）：产物多一个「声明」节 → 按角色配对，不能按索引错位判 FAIL。

    模板 8 节 / 产物 9 节，两侧边距全部按角色正确（封面组 1417、正文组 1440）。
    按索引比前 8 节会把产物 ABSTRACT(1417) 和模板 目录(1440) 对成一对 → 假 FAIL；
    按角色配对后每个产物节都只比同角色模板节 → PASS，且 note 必须给出配对表。
    """
    tpl = _build_role_margin_docx(tmp_path / "tpl_role8.docx", _TPL_ROLE_SECTIONS)
    prod_sections = [
        ("毕业设计（论文）", _MARGIN_COVER),
        ("郑 重 声 明", _MARGIN_COVER),          # 模板没有独立声明节（声明在封面节内）
        ("摘  要", _MARGIN_COVER),
        ("ABSTRACT", _MARGIN_COVER),
        ("目  录", _MARGIN_BODY),
        ("1 绪论", _MARGIN_BODY),
        ("参考文献", _MARGIN_BODY),
        ("致谢", _MARGIN_BODY),
        ("附录", _MARGIN_BODY),
    ]
    prod = _build_role_margin_docx(tmp_path / "prod_role9.docx", prod_sections)
    it = _margin_item(prod, tpl)
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected.startswith("8 节："), it.expected
    assert it.actual.startswith("9 节："), it.actual
    assert "两侧节数不同（8 vs 9）→ 按角色配对" in it.note, it.note
    assert "节3(ABSTRACT)→模板节2(ABSTRACT)" in it.note, it.note
    assert "节4(目录)→模板节3(目录)" in it.note, it.note
    assert "节5(正文)→模板节4(正文)" in it.note, it.note
    assert "节8(附录)→模板节7(附录)" in it.note, it.note


def test_margin_role_pairing_still_fails_when_body_uses_cover_margins(tmp_path):
    """A4 反例：角色配对后，正文节真用封面边距（1417/1701）仍必须 FAIL。

    其余节全对，只把产物「正文」节写成封面那套：按角色它应比模板正文节
    （1440/1800/1440/1800），超差必须点名「节5(正文)→模板节4(正文)」。
    这条守住角色配对不变成「降级放行」。
    """
    tpl = _build_role_margin_docx(tmp_path / "tpl_role8_bad.docx", _TPL_ROLE_SECTIONS)
    prod_sections = [
        ("毕业设计（论文）", _MARGIN_COVER),
        ("郑 重 声 明", _MARGIN_COVER),
        ("摘  要", _MARGIN_COVER),
        ("ABSTRACT", _MARGIN_COVER),
        ("目  录", _MARGIN_BODY),
        ("1 绪论", _MARGIN_COVER),               # ← 正文节真用了封面边距
        ("参考文献", _MARGIN_BODY),
        ("致谢", _MARGIN_BODY),
        ("附录", _MARGIN_BODY),
    ]
    prod = _build_role_margin_docx(tmp_path / "prod_role9_bad.docx", prod_sections)
    it = _margin_item(prod, tpl)
    assert it.verdict == FAIL, (it.verdict, it.expected, it.actual, it.note)
    assert "按角色配对" in it.note, it.note
    assert "超差节 5（1 节）" in it.note, it.note
    assert ("节5(正文)→模板节4(正文) 模板 1440/1800/1440/1800 "
            "vs 产物 1417/1701/1134/1701") in it.note, it.note


def test_margin_role_pairing_falls_back_to_index_and_says_so(tmp_path):
    """A3：认不出角色 → 退回逐节索引，note 写明口径，不静默跳过。

    两侧都没有标题段（`_build_margin_docx` 的段落全在末节）→ 角色只剩「首节=封面」
    （不足 2 个可比模板节）→ 必须退回按索引比前 8 节，并在 note 点名「角色识别失败
    … 退回逐节索引口径」。节数不同、前 8 节边距一致 → PASS。
    """
    tpl = _build_margin_docx(tmp_path / "tpl_norole8.docx",
                             [_MARGIN_COVER] * 3 + [_MARGIN_BODY] * 5)
    prod = _build_margin_docx(tmp_path / "prod_norole9.docx",
                              [_MARGIN_COVER] * 3 + [_MARGIN_BODY] * 5 + [_MARGIN_COVER])
    it = _margin_item(prod, tpl)
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert "按索引逐节对齐比前 8 节" in it.note, it.note
    assert "角色识别失败" in it.note and "退回逐节索引口径" in it.note, it.note
    assert "两侧节数不同（8 vs 9）" in it.note, it.note


_BODY_ONLY_SPEC = {"structure": {"sections": []},
                   "rules": [{"selector": {"semantic_role": "body"},
                              "expected": {"font_size_pt": 14.0}}]}


def _build_body_docx(path: Path, sizes) -> Path:
    """造只含正文段落的 docx：``sizes`` 每项一段，字号写直接格式。"""
    doc = Document()
    for i, size in enumerate(sizes):
        p = doc.add_paragraph()
        _set_run(p.add_run(f"第 {i + 1} 段正文，用来验证正文主字号的期望来源。"), "宋体", size)
    doc.save(str(path))
    return path


def _body_size_item(product, template):
    return _named(evaluate_layout_fidelity(product, template, spec=_BODY_ONLY_SPEC),
                  "正文主字号(pt)")


def test_body_size_expectation_comes_from_template_mode_not_spec(tmp_path):
    """A2（本科形态）：模板正文实测众数 12pt、spec 注解写 14pt → 期望取 12pt。

    产物照模板排 12pt 就不该被判 FAIL。note 必须写明「期望取模板实测 12.0pt」，
    不许把 spec 注解值当事实。
    """
    tpl = _build_body_docx(tmp_path / "tpl_body12.docx", [12.0] * 5 + [14.0])
    prod = _build_body_docx(tmp_path / "prod_body12.docx", [12.0] * 5)
    it = _body_size_item(prod, tpl)
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == 12.0, it.expected
    assert "期望取模板实测 12.0pt（5/6，83%）" in it.note, it.note
    assert "模板正文分布 12.0pt×5、14.0pt×1" in it.note, it.note


def test_body_size_falls_back_to_spec_when_template_samples_too_few(tmp_path):
    """A4（EZI-0022）：模板正文只有 2 段（<5）→ 退回规范注解，note 必须写明出处。

    样本太少的众数是噪声（不能当版式事实），但退回的 spec 值**不是模板实测**——note
    必须显式写「期望来自规范注解（非模板实测）」，不许与模板实测混为一谈。
    """
    tpl = _build_body_docx(tmp_path / "tpl_body2.docx", [12.0, 12.0])
    prod = _build_body_docx(tmp_path / "prod_body12.docx", [12.0] * 5)
    it = _body_size_item(prod, tpl)
    assert it.verdict == FAIL, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == 14.0, it.expected
    assert "样本不足" in it.note, it.note
    assert "期望来自规范注解（非模板实测）14.0pt" in it.note, it.note
    assert "期望取模板实测" not in it.note, it.note


def test_body_size_no_main_peak_uses_top_two_candidates(tmp_path):
    """A2(b)（交大中期形态）：16×7/14×6/18×1（50%，无主峰）→ 双峰区间 16/14。

    期望取模板实测**前二高频**（不是 spec 的 14pt，也不是退回的 spec 值），产物 16pt
    落在第一候选 → PASS。note 要给完整分布与前二频次（EZI-0022）。
    """
    tpl = _build_body_docx(tmp_path / "tpl_jd_mid.docx",
                           [16.0] * 7 + [14.0] * 6 + [18.0])
    prod = _build_body_docx(tmp_path / "prod_jd_mid.docx", [16.0] * 9 + [14.0] * 3)
    it = _body_size_item(prod, tpl)
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == (16.0, 14.0), it.expected
    assert "无主峰" in it.note and "退回 spec" not in it.note, it.note
    assert "前二高频 16.0pt×7、14.0pt×6（合计 93%）" in it.note, it.note
    assert "双峰区间 16.0pt/14.0pt" in it.note, it.note
    assert "模板正文分布 16.0pt×7、14.0pt×6、18.0pt×1" in it.note, it.note


def test_body_size_no_main_peak_passes_on_second_candidate(tmp_path):
    """A2(b)（交大开题形态）：14×8/16×6/18×2/15×1（47%）→ 产物 16pt 是第二候选 → PASS。

    这一条是 A1 的反面：旧口径「无主峰退回 spec」时产物撞上的恰好是 spec 16pt，分不清
    「照模板排」还是「撞常量」；双峰区间只认模板实测的两个候选，产物落在**次高频**仍
    PASS，依据是实测 16.0×6/17 而不是注解。
    """
    tpl = _build_body_docx(tmp_path / "tpl_jd_kt.docx",
                           [14.0] * 8 + [16.0] * 6 + [18.0] * 2 + [15.0])
    prod = _build_body_docx(tmp_path / "prod_jd_kt.docx", [16.0] * 10 + [14.0] * 5)
    it = _body_size_item(prod, tpl)
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert set(it.expected) == {14.0, 16.0}, it.expected
    assert "前二高频 14.0pt×8、16.0pt×6（合计 82%）" in it.note, it.note
    assert "双峰区间 14.0pt/16.0pt" in it.note, it.note
    assert "模板正文分布 14.0pt×8、16.0pt×6、18.0pt×2、15.0pt×1" in it.note, it.note


def test_body_size_no_main_peak_fails_when_product_matches_neither(tmp_path):
    """A1/A2(b)：产物 12pt 既不是 16×7 也不是 14×6 → FAIL，不许拿 spec 12pt 冒充期望。

    旧口径下这条是**假 PASS**：无主峰 → 退回 spec（本例 spec 写 12pt）→ 产物 12pt
    判 PASS，而模板正文里根本没有 12pt。这就是 A1 要杀的「注解常量」后门。
    """
    spec = {"structure": {"sections": []},
            "rules": [{"selector": {"semantic_role": "body"},
                       "expected": {"font_size_pt": 12.0}}]}
    tpl = _build_body_docx(tmp_path / "tpl_jd_mid.docx",
                           [16.0] * 7 + [14.0] * 6 + [18.0])
    prod = _build_body_docx(tmp_path / "prod_body12.docx", [12.0] * 9)
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=spec), "正文主字号(pt)")
    assert it.verdict == FAIL, (it.verdict, it.expected, it.actual, it.note)
    assert set(it.expected) == {16.0, 14.0}, it.expected
    assert "双峰区间" in it.note, it.note


def test_body_size_tie_both_candidates_pass(tmp_path):
    """A5 双峰：12×3/14×3（并列 50%）→ 12 与 14 都是候选，产物任一即 PASS。"""
    tpl = _build_body_docx(tmp_path / "tpl_tie.docx", [12.0] * 3 + [14.0] * 3)
    for size, tag in ((12.0, "p12"), (14.0, "p14")):
        prod = _build_body_docx(tmp_path / f"prod_{tag}.docx", [size] * 5)
        it = _body_size_item(prod, tpl)
        assert it.verdict == PASS, (size, it.verdict, it.expected, it.note)
        assert set(it.expected) == {12.0, 14.0}, it.expected
        assert "双峰区间 12.0pt/14.0pt" in it.note, it.note


def test_body_size_na_when_top_two_cover_too_little(tmp_path):
    """A5 无主峰且分布过散：前二合计 <60% → N-A（不判，也不退回 spec 注解）。

    模板 16×3/14×2/12×2/10.5×2/18×1：最高频 30%、前二合计 50%——前二也代表不了模板，
    任何期望都是编的。旧口径在无主峰时一律退回 spec（14pt）把产物判 FAIL / PASS；
    新口径 N-A，expected 留空（无期望值，不猜）。
    """
    tpl = _build_body_docx(tmp_path / "tpl_scatter.docx",
                           [16.0] * 3 + [14.0] * 2 + [12.0] * 2 + [10.5] * 2 + [18.0])
    prod = _build_body_docx(tmp_path / "prod_body12.docx", [12.0] * 5)
    it = _body_size_item(prod, tpl)
    assert it.verdict == NA, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected is None, it.expected
    assert "无主峰" in it.note and "分布过散" in it.note, it.note
    assert "不退回 spec 注解" in it.note, it.note


_H3_SPEC = {"structure": {"sections": []},
            "rules": [{"selector": {"semantic_role": "heading3"},
                       "expected": {"font_size_pt": 12.0,
                                    "font_east_asia": "宋体",
                                    "jc": "both"}}]}


def _build_level_docx(path: Path, blocks, *, level: int = 3) -> Path:
    """``blocks`` = [(文本, 字号, 中文字体, 对齐)]，全部用 Heading ``level``（带 outlineLvl）。"""
    doc = Document()
    for text, size, font, align in blocks:
        p = doc.add_paragraph(style=f"Heading {level}")
        if align is not None:
            p.paragraph_format.alignment = align
        _set_run(p.add_run(text), font, size)
    doc.save(str(path))
    return path


_H3_TPL = [(f"1.{i} 采购条款{i}", 12.0, "仿宋_GB2312", None) for i in range(1, 7)]


def _h3_item(product, template, name="三级标题对齐(jc)"):
    return _named(evaluate_layout_fidelity(product, template, spec=_H3_SPEC), name)


def test_heading_jc_expectation_comes_from_template_not_spec(tmp_path):
    """A2（新疆形态）：模板三级 6/6 都是左对齐（未设 jc），spec 写 both → 期望取 start。

    产物同样左对齐 → PASS；旧口径拿 spec 的 both 当期望，把按模板排的产物判 FAIL。
    """
    tpl = _build_level_docx(tmp_path / "tpl_h3_start.docx", _H3_TPL)
    prod = _build_level_docx(tmp_path / "prod_h3_start.docx", _H3_TPL)
    it = _h3_item(prod, tpl)
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == "start", it.expected
    assert "期望取模板实测 start（6/6，100%）" in it.note, it.note


def test_heading_font_expectation_comes_from_template_not_spec(tmp_path):
    """A2（东台形态）：同一批三级标题模板用仿宋、产物用宋体 → 期望取模板 仿宋 → FAIL。

    口径改对后不许把真差异放过：两侧标题文本完全对得上（重合 6/6）、字号一致，
    只有中文字体不同，就是真缺陷。
    """
    tpl = _build_level_docx(tmp_path / "tpl_h3_fangsong.docx", _H3_TPL)
    prod = _build_level_docx(tmp_path / "prod_h3_song.docx",
                             [(t, s, "宋体", a) for t, s, _, a in _H3_TPL])
    it = _h3_item(prod, tpl, "三级标题中文字体")
    assert it.verdict == FAIL, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == "仿宋_GB2312", it.expected
    assert "期望取模板实测 仿宋_GB2312（6/6，100%）" in it.note, it.note


def test_heading_falls_back_to_spec_when_populations_do_not_overlap(tmp_path):
    """A2 护栏（新疆二级形态）：模板是「第X章」、产物是条款项，文本重合 0 → 退回 spec。

    两侧比的不是同一批标题时，拿模板实测去判产物就是错位比较（EZI-0015 的老病），
    必须退回 spec 而不是制造一个假 FAIL。
    """
    tpl = _build_level_docx(tmp_path / "tpl_h3_chapter.docx",
                            [(f"第{i}章 投标人须知", 12.0, "宋体", WD_ALIGN_PARAGRAPH.CENTER)
                             for i in range(1, 7)])
    prod = _build_level_docx(tmp_path / "prod_h3_clause.docx",
                             [(f"1.{i} 采购条款{i}", 12.0, "宋体", WD_ALIGN_PARAGRAPH.JUSTIFY)
                              for i in range(1, 7)])
    it = _h3_item(prod, tpl)
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == "both", it.expected        # spec 值
    assert "两侧不是同一批标题" in it.note and "重合 0/6" in it.note, it.note


def _ezi25_tpl_three(tmp_path, name="tpl_h1_three.docx"):
    """模板一级 = 3 个「第X章」16pt 黑体居中（东台骨架）。"""
    return _build_level_docx(
        tmp_path / name,
        [(f"第{i}章 投标人须知", 16.0, "黑体", WD_ALIGN_PARAGRAPH.CENTER) for i in range(1, 4)],
        level=1)


def _ezi25_prod_three_plus_five(tmp_path, name="prod_h1_many.docx",
                                ch_size=10.5, ch_font="宋体", ch_align=WD_ALIGN_PARAGRAPH.LEFT):
    """产物一级 = 3 个同文章标题 + 5 个无同文的条款项（东台 work17 形态）。"""
    return _build_level_docx(
        tmp_path / name,
        [(f"第{i}章 投标人须知", ch_size, ch_font, ch_align) for i in range(1, 4)]
        + [(f"{i} 获取招标文件", 10.5, "宋体", WD_ALIGN_PARAGRAPH.LEFT) for i in range(1, 6)],
        level=1)


_EZI25_SPEC_SIZE = {"structure": {"sections": []},
                    "rules": [{"selector": {"semantic_role": "heading1"},
                               "expected": {"font_size_pt": 16.0}}]}


def test_heading_expectation_uses_same_text_pairs_below_sample_floor(tmp_path):
    """EZI-0025：模板只有 3 段，但与产物**同文成对**且全一致 → 期望取模板实测。

    EZI-0016 的样本下限（≥ ``TEMPLATE_MODE_MIN_SAMPLES``）防的是「拿 3 个样本的众数推
    该级版式」；这里比的是一一对应的同一批段落（产物那 3 段就是模板那 3 段），不是统计
    推断。东台一级正是这个形态：spec 写宋体、模板 3/3 黑体——旧口径退回 spec 之后，
    尺子既看不见模板是什么，也看不见引擎把真章标题改成了黑体。
    """
    tpl = _ezi25_tpl_three(tmp_path)
    prod = _ezi25_prod_three_plus_five(tmp_path)
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=_EZI25_SPEC_SIZE),
                "一级标题字号(pt)")
    assert it.verdict == FAIL, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == 16.0, it.expected
    assert "同文成对 3 段" in it.note, it.note


def test_heading_expectation_keeps_sample_floor_when_nothing_is_paired(tmp_path):
    """EZI-0025 只把「同文成对」当直接证据，没有取消 EZI-0016 的样本下限。

    同样的 3 段模板，产物一级全是无同文的条款项（重合 0）→ 3 个模板样本依旧是噪声，
    仍退回 spec 并在 note 写明「样本不足」+「两侧不是同一批标题」。
    """
    tpl = _ezi25_tpl_three(tmp_path)
    prod = _build_level_docx(
        tmp_path / "prod_only_clauses.docx",
        [(f"{i} 获取招标文件", 10.5, "宋体", WD_ALIGN_PARAGRAPH.LEFT) for i in range(1, 9)],
        level=1)
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=_EZI25_SPEC_SIZE),
                "一级标题字号(pt)")
    assert it.expected == 16.0, it.expected
    assert "样本不足" in it.note, it.note          # 无同文 → 3 样本仍是噪声


# ── EZI-0021：一级标题字号只判「与模板同级同文」的那批 ──

_EZI21_CHAPTERS = ("第一章 投标人须知", "第二章  投标文件格式", "第三章  合同文件")
_EZI21_CLAUSES = ("2.招标内容清单", "1.供货时间及地点", "2.质量及验收方法",
                  "三、付款方式", "五、违约责任")
_EZI21_SPEC_SIZE = {"structure": {"sections": []},
                    "rules": [{"selector": {"semantic_role": "heading1"},
                               "expected": {"font_size_pt": 16.0}}]}
_EZI21_SPEC_SIZE_JC = {"structure": {"sections": []},
                       "rules": [{"selector": {"semantic_role": "heading1"},
                                  "expected": {"font_size_pt": 16.0, "jc": "both"}}]}


def _ezi21_tpl(tmp_path, name="tpl.docx"):
    """模板一级 = 3 个「第X章」，16pt 居中（东台招标实测形态）。"""
    return _build_level_docx(
        tmp_path / name,
        [(t, 16.0, "黑体", WD_ALIGN_PARAGRAPH.CENTER) for t in _EZI21_CHAPTERS], level=1)


def _ezi21_prod(tmp_path, ch_size=16.0, clause_size=10.5, name="prod.docx",
                clause_align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    """产物一级 = 3 个章标题 + 5 个条款项（条款项被标成 outlineLvl=0）。"""
    return _build_level_docx(
        tmp_path / name,
        [(t, ch_size, "宋体", WD_ALIGN_PARAGRAPH.CENTER) for t in _EZI21_CHAPTERS]
        + [(t, clause_size, "宋体", clause_align) for t in _EZI21_CLAUSES], level=1)


def test_heading_size_judges_only_comparable_samples(tmp_path):
    """EZI-0021（东台形态）：模板一级 3 章 16pt，产物一级 8 段混排（16×3 + 10.5×5）。

    旧口径拿 8 段的众数（10.5）当实测 → 把 3/3 都排对的章标题判 FAIL（真样本上的
    「一级标题字号 10.5 ≠ 16」就是这么来的）。真章标题与模板同级同文 → PASS，
    note 必须写明被排除的条款项条数（A2）。
    """
    tpl, prod = _ezi21_tpl(tmp_path), _ezi21_prod(tmp_path)
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=_EZI21_SPEC_SIZE, source=prod),
                "一级标题字号(pt)")
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.actual == 16.0, it.actual
    assert "可比样本 3 段" in it.note and "未参与 5 段" in it.note, it.note


def test_heading_size_still_fails_when_real_chapters_have_wrong_size(tmp_path):
    """A1 保留判 FAIL 的能力：章标题 3 段 12pt、条款项 5 段 16pt。

    旧口径（8 段众数 16）会 PASS——那才是真的放过；收窄后 3/3 = 12 ≠ 16 → FAIL。
    """
    tpl = _ezi21_tpl(tmp_path)
    prod = _ezi21_prod(tmp_path, ch_size=12.0, clause_size=16.0)
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=_EZI21_SPEC_SIZE, source=prod),
                "一级标题字号(pt)")
    assert it.verdict == FAIL, (it.verdict, it.expected, it.actual, it.note)
    assert it.actual == 12.0, it.actual
    assert "可比样本 3 段" in it.note, it.note


def test_heading_size_keeps_full_sample_when_populations_mostly_overlap(tmp_path):
    """重合 ≥ HEADING_OVERLAP_MIN 不收窄：模板里没有的新标题仍参与该级实测。"""
    tpl = _build_level_docx(
        tmp_path / "tpl6.docx",
        [(f"第{i}章 正文章节", 16.0, "黑体", WD_ALIGN_PARAGRAPH.CENTER) for i in range(1, 7)],
        level=1)
    prod = _build_level_docx(
        tmp_path / "prod8.docx",
        [(f"第{i}章 正文章节", 16.0, "宋体", WD_ALIGN_PARAGRAPH.CENTER) for i in range(1, 7)]
        + [("7 附录A", 12.0, "宋体", WD_ALIGN_PARAGRAPH.CENTER),
           ("8 附录B", 12.0, "宋体", WD_ALIGN_PARAGRAPH.CENTER)], level=1)
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=_EZI21_SPEC_SIZE, source=prod),
                "一级标题字号(pt)")
    assert it.verdict == PASS and it.actual == 16.0, (it.verdict, it.actual, it.note)
    assert "可比样本" not in it.note, it.note          # 6/8 = 75% ≥ 50% → 全样本


def test_heading_size_keeps_full_sample_when_template_has_no_headings(tmp_path):
    """模板该级无编号样本（本科论文形态）：不受收窄影响，仍按 spec 判全样本。"""
    tpl = _build_level_docx(tmp_path / "tpl_none.docx",
                            [("ABSTRACT", 16.0, "黑体", WD_ALIGN_PARAGRAPH.CENTER)], level=1)
    prod = _build_level_docx(tmp_path / "prod3.docx",
                             [(t, 16.0, "宋体", WD_ALIGN_PARAGRAPH.CENTER)
                              for t in _EZI21_CHAPTERS], level=1)
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=_EZI21_SPEC_SIZE, source=prod),
                "一级标题字号(pt)")
    assert it.verdict == PASS and it.actual == 16.0, (it.verdict, it.actual, it.note)
    assert "可比样本" not in it.note, it.note


def test_heading_size_keeps_full_sample_when_nothing_is_comparable(tmp_path):
    """新疆二级形态：模板 6 个「第X章」、产物 6 个条款项，重合 0 → 无可比样本。

    不 N-A 也不收窄（收窄会把该项变成永远 N-A 的死指标）：沿用全样本 + spec，
    note 保留「两侧不是同一批标题」。
    """
    tpl = _build_level_docx(
        tmp_path / "tpl_chapters.docx",
        [(f"第{i}章 招标内容", 12.0, "宋体", WD_ALIGN_PARAGRAPH.CENTER) for i in range(1, 7)],
        level=1)
    prod = _build_level_docx(
        tmp_path / "prod_clauses.docx",
        [(f"{i}.{i} 采购条款{i}", 12.0, "宋体", WD_ALIGN_PARAGRAPH.JUSTIFY) for i in range(1, 7)],
        level=1)
    spec = {"structure": {"sections": []},
            "rules": [{"selector": {"semantic_role": "heading1"},
                       "expected": {"font_size_pt": 12.0}}]}
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=spec, source=prod), "一级标题字号(pt)")
    assert it.verdict == PASS and it.actual == 12.0, (it.verdict, it.actual, it.note)
    assert "可比样本" not in it.note, it.note
    assert "两侧不是同一批标题" in it.note and "重合 0/6" in it.note, it.note


def test_heading_jc_follows_the_same_narrowing_as_size(tmp_path):
    """EZI-0025：收窄不再只作用于字号——对齐项也只认「与模板同级同文」的那 3 段。

    EZI-0021 当年留过「只收窄字号」的口子，理由是期望侧模板只有 3 样本会退回 spec 的
    both（模板 3/3 实为 center）→ 造出假 FAIL。EZI-0025 把期望侧一并改成「模板同文成对」
    之后这条假 FAIL 的成因没了，口子必须关掉：实测 center（3/3）对模板同文 center（3/3）
    → PASS。旧口径拿 8 段众数 both 去比（条款项主导），同样 PASS 但理由全错。
    """
    tpl = _ezi21_tpl(tmp_path)
    prod = _ezi21_prod(tmp_path)
    rep = evaluate_layout_fidelity(prod, tpl, spec=_EZI21_SPEC_SIZE_JC, source=prod)
    sz = _named(rep, "一级标题字号(pt)")
    jc = _named(rep, "一级标题对齐(jc)")
    assert "可比样本 3 段" in sz.note, sz.note
    assert jc.verdict == PASS and jc.actual == "center", (jc.verdict, jc.actual, jc.note)
    assert jc.expected == "center", jc.expected
    assert "可比样本 3 段" in jc.note and "未参与 5 段" in jc.note, jc.note


# ── EZI-0025：字体与对齐也按「同级同文」判（EZI-0021 只接了字号）─────────────

#: spec 故意写成与模板不同的值（宋体/both）：期望该来自模板实测，不是这份注解。
_EZI25_SPEC = {"structure": {"sections": []},
               "rules": [{"selector": {"semantic_role": "heading1"},
                          "expected": {"font_size_pt": 16.0,
                                       "font_east_asia": "宋体",
                                       "jc": "both"}}]}


def _ezi25_tpl(tmp_path, name="tpl_ezi25.docx"):
    """模板一级 = 3 个「第X章」，16pt 黑体居中（东台骨架）。"""
    return _build_level_docx(
        tmp_path / name,
        [(t, 16.0, "黑体", WD_ALIGN_PARAGRAPH.CENTER) for t in _EZI21_CHAPTERS], level=1)


def _ezi25_prod(tmp_path, name="prod_ezi25.docx", ch_font="黑体",
                ch_align=WD_ALIGN_PARAGRAPH.CENTER, with_clauses=True):
    """产物一级 = 3 个章标题 + 5 个条款项（条款项被标成 outlineLvl=0）。"""
    rows = [(t, 16.0, ch_font, ch_align) for t in _EZI21_CHAPTERS]
    if with_clauses:
        rows += [(t, 10.5, "宋体", WD_ALIGN_PARAGRAPH.JUSTIFY) for t in _EZI21_CLAUSES]
    return _build_level_docx(tmp_path / name, rows, level=1)


def test_heading_font_judged_on_comparable_samples_not_spec(tmp_path):
    """A2（东台形态）：spec 写宋体、模板 3/3 黑体、产物真章标题 3/3 黑体 → PASS。

    期望必须取模板同文的 3 段黑体（note 写明「可比样本 3 段」「同文成对 3 段」），
    而不是 spec 的宋体——旧口径在 31 段条款项混进一级的产物上 PASS，但理由是错的：
    它既没看见模板是什么，也没看见引擎（EZI-0024）已经把真章标题改对。
    """
    tpl, prod = _ezi25_tpl(tmp_path), _ezi25_prod(tmp_path)
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=_EZI25_SPEC, source=prod),
                "一级标题中文字体")
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == "黑体", it.expected
    assert it.actual == "黑体", it.actual
    assert "可比样本 3 段" in it.note and "同文成对 3 段" in it.note, it.note


def test_heading_jc_judged_on_comparable_samples_not_spec(tmp_path):
    """A1/A2（东台形态）：spec 写 both、模板 3/3 居中 → 期望取模板 center，产物 center → PASS。"""
    tpl, prod = _ezi25_tpl(tmp_path), _ezi25_prod(tmp_path)
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=_EZI25_SPEC, source=prod),
                "一级标题对齐(jc)")
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == "center" and it.actual == "center", (it.expected, it.actual)
    assert "同文成对 3 段" in it.note, it.note


def test_heading_font_counterexample_real_chapters_wrong_font_fails(tmp_path):
    """A4 反例：真章标题字体被改错（模板黑体 → 产物宋体）→ 必须 FAIL。

    不许因为「可比样本只有 3 段、模板样本不足」就放过：那 3 段是唯一能判真假的样本。
    """
    tpl = _ezi25_tpl(tmp_path)
    prod = _ezi25_prod(tmp_path, ch_font="宋体")
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=_EZI25_SPEC, source=prod),
                "一级标题中文字体")
    assert it.verdict == FAIL, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == "黑体" and it.actual == "宋体", (it.expected, it.actual)


def test_heading_jc_counterexample_real_chapters_wrong_align_fails(tmp_path):
    """A4 反例：真章标题对齐被改错（模板 center → 产物 both）→ 必须 FAIL。"""
    tpl = _ezi25_tpl(tmp_path)
    prod = _ezi25_prod(tmp_path, ch_align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=_EZI25_SPEC, source=prod),
                "一级标题对齐(jc)")
    assert it.verdict == FAIL, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == "center" and it.actual == "both", (it.expected, it.actual)


def test_heading_font_and_jc_narrow_like_size(tmp_path):
    """A1：字体/对齐与字号共用一套样本——5 段条款项（宋体/both）不再参与这两项的实测。

    收窄前字体众数 8 段里 5 段是宋体、对齐众数 5/8 是 both（两项都会被条款项带偏）。
    """
    tpl, prod = _ezi25_tpl(tmp_path), _ezi25_prod(tmp_path)
    rep = evaluate_layout_fidelity(prod, tpl, spec=_EZI25_SPEC, source=prod)
    for name, actual, expected in (("一级标题中文字体", "黑体", "黑体"),
                                   ("一级标题对齐(jc)", "center", "center")):
        it = _named(rep, name)
        assert it.verdict == PASS, (name, it.verdict, it.expected, it.actual, it.note)
        assert it.actual == actual and it.expected == expected, (name, it.actual, it.expected)
        assert "可比样本 3 段" in it.note and "未参与 5 段" in it.note, it.note


def test_heading_font_falls_back_to_spec_when_template_same_text_is_mixed(tmp_path):
    """A3 侧边界：模板同文 3 段黑体 + 3 段仿宋 = 模板自己就不统一 → 不拿众数当期望。

    note 必须写明「模板同文 6 段中文字体不统一（…）」，再按旧护栏退回 spec。判 6/14 段
    各用各的字体时，众数（50%）不是版式事实——东台二级正是这个形态（仿宋 6 / 黑体 6 /
    仿宋 2）。
    """
    tpl = _build_level_docx(
        tmp_path / "tpl_mixed.docx",
        [(t, 16.0, "黑体", WD_ALIGN_PARAGRAPH.CENTER) for t in _EZI21_CHAPTERS]
        + [(f"{i}.格式附件{i}", 15.0, "仿宋", WD_ALIGN_PARAGRAPH.CENTER) for i in range(1, 4)],
        level=1)
    prod = _build_level_docx(
        tmp_path / "prod_mixed.docx",
        [(t, 16.0, "黑体", WD_ALIGN_PARAGRAPH.CENTER) for t in _EZI21_CHAPTERS]
        + [(f"{i}.格式附件{i}", 15.0, "仿宋", WD_ALIGN_PARAGRAPH.CENTER) for i in range(1, 4)]
        + [(t, 10.5, "仿宋", WD_ALIGN_PARAGRAPH.JUSTIFY) for t in _EZI21_CLAUSES], level=1)
    spec = {"structure": {"sections": []},
            "rules": [{"selector": {"semantic_role": "heading1"},
                       "expected": {"font_east_asia": "仿宋"}}]}
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=spec, source=prod),
                "一级标题中文字体")
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == "仿宋", it.expected                  # 退回 spec，不是模板众数
    assert "模板同文 6 段中文字体不统一" in it.note, it.note
    assert "退回 spec 规则" in it.note, it.note


def test_heading_font_na_when_comparable_samples_have_no_font(tmp_path):
    """A3：收窄后可比样本一个字体都测不到 → N-A，**不退回全量样本**的宋体条款项。

    章标题不写 ``w:rFonts``（eastAsia 无值）而条款项写了宋体：全量样本能测到「宋体」，
    但那是 5 段条款项的字体，拿去判 3 段章标题就是错位比较。
    """
    from docx import Document

    tpl = _ezi25_tpl(tmp_path)
    prod = tmp_path / "prod_no_font.docx"
    doc = Document()
    for t in _EZI21_CHAPTERS:
        doc.add_paragraph(style="Heading 1").add_run(t)      # 不写 rFonts → eastAsia 无值
    for t in _EZI21_CLAUSES:
        _set_run(doc.add_paragraph(style="Heading 1").add_run(t), "宋体", 10.5)
    doc.save(str(prod))
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=_EZI25_SPEC, source=prod),
                "一级标题中文字体")
    assert it.verdict == NA, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected is None, it.expected
    assert "可比样本 3 段均无可测中文字体" in it.note, it.note
    assert "不退回全量样本" in it.note, it.note


def test_heading_font_keeps_full_sample_when_nothing_is_comparable(tmp_path):
    """A4 边界：两侧一条同文都对不上（重合 0）→ 字体仍按全量 + spec 判，不 N-A。

    EZI-0021 决策记录 3 的理由（否则成了永远不判的死指标）对字体同样成立；note 必须
    写明「两侧不是同一批标题」。
    """
    tpl = _build_level_docx(
        tmp_path / "tpl_chapters.docx",
        [(f"第{i}章 招标内容", 12.0, "宋体", WD_ALIGN_PARAGRAPH.CENTER) for i in range(1, 7)],
        level=1)
    prod = _build_level_docx(
        tmp_path / "prod_clauses.docx",
        [(f"{i}.{i} 采购条款{i}", 12.0, "宋体", WD_ALIGN_PARAGRAPH.JUSTIFY)
         for i in range(1, 7)], level=1)
    spec = {"structure": {"sections": []},
            "rules": [{"selector": {"semantic_role": "heading1"},
                       "expected": {"font_east_asia": "宋体"}}]}
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=spec, source=prod),
                "一级标题中文字体")
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == "宋体" and it.actual == "宋体", (it.expected, it.actual)
    assert "两侧不是同一批标题" in it.note and "重合 0/6" in it.note, it.note


# ── EZI-0039：同文成对取众数（≥60%），不得因「不统一」退回 spec ──
# 深圳招标形态：模板一级 31 段，其中 10 段与产物同文；这 10 段里 9 段 16pt + 第九章
# 22pt（90%）。旧口径要求「同文段 100% 一致」→ 判「不统一」→ 退回 spec 14pt，
# 产物照模板排 16pt 反被判 FAIL；引擎改成 spec 值 14pt 反而 PASS（尺子拉错方向）。

_EZI39_SPEC_SIZE = {"structure": {"sections": []},
                    "rules": [{"selector": {"semantic_role": "heading1"},
                               "expected": {"font_size_pt": 14.0}}]}
_EZI39_CHAPTERS = [f"第{i}章 招标文件条款{i}" for i in range(1, 11)]


def _ezi39_shenzhen_pair(tmp_path, prod_size=16.0):
    """模板一级 31 段（10 段与产物同文，其中第九章 22pt）+ 产物一级 31 段。"""
    tpl_blocks = [(t, 16.0, "黑体", WD_ALIGN_PARAGRAPH.CENTER) for t in _EZI39_CHAPTERS]
    tpl_blocks[9] = (_EZI39_CHAPTERS[9], 22.0, "黑体", WD_ALIGN_PARAGRAPH.CENTER)
    tpl_blocks += [(f"第{i}章 附则{i}", 16.0, "黑体", WD_ALIGN_PARAGRAPH.CENTER)
                   for i in range(11, 32)]
    tpl = _build_level_docx(tmp_path / "tpl_shenzhen.docx", tpl_blocks, level=1)
    prod_blocks = [(t, prod_size, "黑体", WD_ALIGN_PARAGRAPH.CENTER) for t in _EZI39_CHAPTERS]
    prod_blocks += [(f"{i}.{i} 投标人资格条款{i}", 10.5, "宋体", WD_ALIGN_PARAGRAPH.JUSTIFY)
                    for i in range(1, 22)]
    prod = _build_level_docx(tmp_path / "prod_shenzhen.docx", prod_blocks, level=1)
    return tpl, prod


def test_heading_size_same_text_takes_mode_instead_of_spec(tmp_path):
    """A1/A2（深圳形态）：同文 10 段 16.0×9 + 22.0×1（90%）→ 期望取众数 16.0，不退回 spec。

    可比样本 10/31 = 32% < ``HEADING_OVERLAP_MIN`` → 收窄到那 10 段；同文段「不统一」
    不是退回 spec 的理由——最高频 90% ≥ 60%（= ``_DOMINANCE_DEFAULT``）即采信。
    """
    tpl, prod = _ezi39_shenzhen_pair(tmp_path)
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=_EZI39_SPEC_SIZE, source=prod),
                "一级标题字号(pt)")
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == 16.0, it.expected
    assert it.actual == 16.0, it.actual
    assert "16.0pt×9" in it.note and "22.0pt×1" in it.note, it.note
    assert "取众数" in it.note and "EZI-0039" in it.note, it.note
    assert "退回 spec" not in it.note, it.note


def test_heading_size_spec_value_is_not_a_free_pass(tmp_path):
    """A2 反例：产物改成 spec 的 14pt（旧口径下的「标准答案」）→ 必须 FAIL。

    旧口径把 spec 14.0 当期望，引擎写成 14pt 就 PASS（产物变错反而过关）；按模板同文
    众数 16.0 判，14pt 就是真差异。
    """
    tpl, prod = _ezi39_shenzhen_pair(tmp_path, prod_size=14.0)
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=_EZI39_SPEC_SIZE, source=prod),
                "一级标题字号(pt)")
    assert it.verdict == FAIL, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == 16.0 and it.actual == 14.0, (it.expected, it.actual)


def test_heading_size_same_text_no_single_peak_uses_top_two_candidates(tmp_path):
    """A4 上半：同文 16.0×4 / 14.0×4 / 12.0×2（最高频 40%）→ 前二合计 80% ≥60% → 双峰区间。

    地板与正文主字号 EZI-0022 同一条 ``TEMPLATE_MODE_MIN_SHARE``：产物落在 14.0 或
    16.0 任一即 PASS，note 必须写明「双峰区间」与前二频次。
    """
    blocks = ([(t, 16.0, "黑体", WD_ALIGN_PARAGRAPH.CENTER) for t in _EZI39_CHAPTERS[:4]]
              + [(t, 14.0, "黑体", WD_ALIGN_PARAGRAPH.CENTER) for t in _EZI39_CHAPTERS[4:8]]
              + [(t, 12.0, "黑体", WD_ALIGN_PARAGRAPH.CENTER) for t in _EZI39_CHAPTERS[8:]])
    tpl = _build_level_docx(tmp_path / "tpl_bimodal.docx", blocks, level=1)
    prod = _build_level_docx(
        tmp_path / "prod_bimodal.docx",
        [(t, 14.0, "黑体", WD_ALIGN_PARAGRAPH.CENTER) for t in _EZI39_CHAPTERS], level=1)
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=_EZI39_SPEC_SIZE, source=prod),
                "一级标题字号(pt)")
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == (16.0, 14.0), it.expected
    assert "双峰区间" in it.note and "EZI-0022" in it.note, it.note


def test_heading_size_same_text_scattered_falls_back_with_reason(tmp_path):
    """A4 下半：同文 5 段 5 个不同字号（前二合计 40% <60%）→ 退回 spec 并写明依据。

    分布过散时前二也代表不了模板：期望取 spec 18.0（不是静默取首个 16.0），note 里
    「最高频」「前二高频」「退回 spec 规则」三件套必须齐全。
    """
    sizes = (16.0, 15.0, 14.0, 12.0, 10.5)
    tpl = _build_level_docx(
        tmp_path / "tpl_scattered.docx",
        [(t, s, "黑体", WD_ALIGN_PARAGRAPH.CENTER)
         for t, s in zip(_EZI39_CHAPTERS[:5], sizes)], level=1)
    prod = _build_level_docx(
        tmp_path / "prod_scattered.docx",
        [(t, 16.0, "黑体", WD_ALIGN_PARAGRAPH.CENTER) for t in _EZI39_CHAPTERS[:5]], level=1)
    spec = {"structure": {"sections": []},
            "rules": [{"selector": {"semantic_role": "heading1"},
                       "expected": {"font_size_pt": 18.0}}]}
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=spec, source=prod),
                "一级标题字号(pt)")
    assert it.expected == 18.0, it.expected
    assert it.verdict == FAIL, (it.verdict, it.expected, it.actual, it.note)
    assert "最高频 20%" in it.note, it.note
    assert "前二高频" in it.note and "合计 40%" in it.note, it.note
    assert "退回 spec 规则" in it.note, it.note


def test_heading_size_same_text_two_pairs_still_accepted(tmp_path):
    """A3（重庆形态）：同文 2/2 全一致 → 期望仍 16.0，不因本轮改众数而漂移。"""
    tpl = _build_level_docx(
        tmp_path / "tpl_two.docx",
        [(t, 16.0, "黑体", WD_ALIGN_PARAGRAPH.CENTER) for t in _EZI39_CHAPTERS[:2]], level=1)
    prod = _build_level_docx(
        tmp_path / "prod_two.docx",
        [(t, 16.0, "黑体", WD_ALIGN_PARAGRAPH.CENTER) for t in _EZI39_CHAPTERS[:2]], level=1)
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=_EZI39_SPEC_SIZE, source=prod),
                "一级标题字号(pt)")
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == 16.0, it.expected
    assert "同文成对 2 段" in it.note, it.note


def test_heading_font_same_text_takes_mode_when_dominant(tmp_path):
    """字体同口径：同文 5 段 黑体×4 + 仿宋×1（80%）→ 期望取众数 黑体。

    字体与字号共用 ``_same_text_expectation``；旧的「100% 一致」门槛会把这一批判成
    「模板自己就不统一」再退回 spec。产物与模板同文段一致 → PASS。
    """
    tpl = _build_level_docx(
        tmp_path / "tpl_font_mode.docx",
        [(t, 16.0, "黑体", WD_ALIGN_PARAGRAPH.CENTER) for t in _EZI39_CHAPTERS[:4]]
        + [(_EZI39_CHAPTERS[4], 16.0, "仿宋", WD_ALIGN_PARAGRAPH.CENTER)], level=1)
    prod = _build_level_docx(
        tmp_path / "prod_font_mode.docx",
        [(t, 16.0, "黑体", WD_ALIGN_PARAGRAPH.CENTER) for t in _EZI39_CHAPTERS[:5]], level=1)
    spec = {"structure": {"sections": []},
            "rules": [{"selector": {"semantic_role": "heading1"},
                       "expected": {"font_east_asia": "仿宋"}}]}
    it = _named(evaluate_layout_fidelity(prod, tpl, spec=spec, source=prod),
                "一级标题中文字体")
    assert it.verdict == PASS, (it.verdict, it.expected, it.actual, it.note)
    assert it.expected == "黑体", it.expected                     # 众数，不是 spec 仿宋
    assert "取众数" in it.note and "黑体×4" in it.note, it.note

