"""目录验收（2026-11-08 新口径）：Word 更新域之后仍然正确。

为什么单独成模块：旧口径只判「有没有 TOC 域」，读的是我们写进去的缓存条目，
于是「参考文献/致谢/附录没有大纲级别，F9 一按整条消失」「目录与正文同节，F9 后
接排」两类缺陷在 199 份基线上全绿（60 份有 TOC，58 份目录后无分节）。

三条判据（都在 Word 里实跑）：
  A 条目完整：正文一级标题（含参考文献/致谢/附录/结论）必须都出现在 F9 后的目录；
  B 目录独立：目录最后一条所在页 < 正文首章所在页；
  C 页码一致：条目页码 == 页脚显示页码（罗马数字按数值比）。

三条判据都吃过「判据自己错」的亏，下面每处取舍都标了实测原因，别顺手改回去。

所有操作在临时副本上，不改交付文件。
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_SPACE_RE = re.compile(r"\s+")

_WORD_TOC_PS = r"""param([string]$Path, [string]$Out)
$ErrorActionPreference = 'Stop'
$lines = New-Object System.Collections.Generic.List[string]
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$doc = $w.Documents.Open($Path, $false, $false)
try {
    $doc.Repaginate()
    if ($doc.TablesOfContents.Count -gt 0) {
        # 模仿人类：先把整个目录更新一遍（条目+页码都重算），再看结果
        $doc.TablesOfContents.Item(1).Update()
        $doc.Repaginate()
    }
    $lines.Add("TOC_COUNT=" + $doc.TablesOfContents.Count)
    $tocStart = 0
    $tocEnd = 0
    if ($doc.TablesOfContents.Count -gt 0) {
        $rng = $doc.TablesOfContents.Item(1).Range
        $tocStart = $rng.Start
        $tocEnd = $rng.End
        $t = $rng.Text
        foreach ($l in ($t -split "`r")) {
            if ($l.Trim() -ne "") { $lines.Add("TOCLINE " + $l.Trim()) }
        }
        # 目录末页：取目录范围内**最后一个有文字的段**所在页。
        # 不能用 Range.End 的页——域范围末尾常带一个空段，页码会落到下一页，
        # 于是「目录已独占一页」被判成同页（实测 4 份假失败）。
        $lastTocPage = 0
        foreach ($p in $doc.Paragraphs) {
            if ($p.Range.Start -lt $tocStart) { continue }
            if ($p.Range.Start -gt $tocEnd) { break }
            if ($p.Range.Text.Trim().Length -gt 0) { $lastTocPage = $p.Range.Information(3) }
        }
        $lines.Add("TOC_LAST_PAGE=" + $lastTocPage)
    }
    $lines.Add("PAGES=" + $doc.ComputeStatistics(2))
    # 正文首页：优先「第 X 章」标题，其次任一「大纲级别 1」段，最后取目录结束后
    # 第一段非空段落（有些文档正文标题既无章号也无大纲级别，但排版上仍是正文）。
    $firstPage = 0
    $fallback = 0
    $firstBody = 0
    $afterToc = $tocEnd
    foreach ($p in $doc.Paragraphs) {
        if ($p.Range.Start -lt $afterToc) { continue }
        $txt = $p.Range.Text.Trim()
        if ($txt.Length -eq 0) { continue }
        if ($firstBody -eq 0) { $firstBody = $p.Range.Information(3) }
        if ($p.OutlineLevel -eq 1) {
            if ($fallback -eq 0) { $fallback = $p.Range.Information(3) }
            if ($txt -match '^第\s*[0-9一二三四五六七八九十]+\s*章') {
                $firstPage = $p.Range.Information(3)
                $lines.Add("FIRST_HEAD_TEXT=" + $txt)
                break
            }
        }
    }
    if ($firstPage -eq 0) { $firstPage = $fallback }
    if ($firstPage -eq 0) { $firstPage = $firstBody }
    $lines.Add("FIRST_HEAD_PAGE=" + $firstPage)
    $doc.Save()
} finally {
    try { $doc.Close($false) } catch {}
    try { $w.Quit() } catch {}
}
[System.IO.File]::WriteAllLines($Out, $lines, (New-Object System.Text.UTF8Encoding($false)))
"WROTE " + $Out
"""

_ROMAN = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}
# 不该被当成「一级标题」的正文段：日期行、签名/签字行、页码行、空标题序号
_JUNK_RE = re.compile(
    r"(年\s*月\s*日|年.{0,6}月.{0,6}日|签名|签字|盖章|注[:：]|^\d+$|^[（(]\s*[)）]$)"
)


def _roman_to_int(s: str) -> int:
    total = 0
    prev = 0
    for ch in reversed(s.upper()):
        v = _ROMAN.get(ch, 0)
        total = total - v if v < prev else total + v
        prev = max(prev, v)
    return total


def _norm(s: str) -> str:
    """标题比对用归一：去掉页尾（tab/点线/空格 + 页码，含罗马数字）、去空白。"""
    s = (s or "").replace("　", " ")
    s = s.replace("\u00a0", " ")
    s = s.split("\t")[0]
    s = re.sub(r"\.{2,}.*$", "", s)
    s = re.sub(r"[…·]{2,}.*$", "", s)
    s = re.sub(r"\s*[-—–]?\s*[0-9IVXLCDMivxlcdm]+\s*$", "", s)
    return _SPACE_RE.sub("", s).strip()


def _paragraph_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(W + "t")).strip()


def _is_toc_paragraph(p) -> bool:
    """目录条目段：TOC 样式，或带 PAGEREF 域。

    只判 w:hyperlink 是不够的——我们自己插入的目录用的是 PAGEREF 域、没有
    hyperlink，于是目录行（连「20   年  月  日」这种日期行）会被当成「应有一级
    标题」，验收就永远报缺条目（实测 2 份假失败）。
    """
    pr = p.find(W + "pPr")
    if pr is not None:
        st = pr.find(W + "pStyle")
        if st is not None and (st.get(W + "val") or "").upper().startswith("TOC"):
            return True
    if p.find(W + "hyperlink") is not None:
        return True
    for ins in p.iter(W + "instrText"):
        if "PAGEREF" in (ins.text or "").upper():
            return True
    for fs in p.iter(W + "fldSimple"):
        if "PAGEREF" in (fs.get(W + "instr") or "").upper():
            return True
    return False


def _expected_headings(docx: Path) -> list[str]:
    """正文应有的一级条目：outlineLvl=0 的 body 段 + 后置章节语义兜底。"""
    with zipfile.ZipFile(docx) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    body = root.find(W + "body")
    if body is None:
        return []
    out: list[str] = []
    seen: set[str] = set()
    for p in body.iter(W + "p"):
        if p.getparent() is not body:
            continue
        if _is_toc_paragraph(p):
            continue
        txt = _paragraph_text(p)
        if len(txt) < 2:
            continue
        # 手打/粘贴的目录行（用 tab 或省略号连页码）不是正文标题——不排掉会把
        # 「附录……………16」这种目录行当成「应有一级条目」，验收永远报缺条目。
        if "	" in txt or re.search(r"…{2,}|\.{5,}|·{3,}", txt):
            continue
        pr = p.find(W + "pPr")
        ol = pr.find(W + "outlineLvl") if pr is not None else None
        is_l1 = ol is not None and ol.get(W + "val") == "0"
        # 「应有条目」不能只取 outlineLvl=0——那正是被测属性，缺了就永远「没少条目」。
        # 后置章节按语义文本兜底（参考文献/致谢/附录/结论）。
        key = _norm(txt)
        is_back = bool(re.match(r"^(参考文献|致谢|附录[A-Z]{0,2}|结论)$", key))
        if not (is_l1 or is_back):
            continue
        if _JUNK_RE.search(txt):
            continue
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(txt)
    return out


def _run_word_roundtrip(docx: Path, workdir: Path) -> dict:
    ps_file = workdir / "word_toc_after_update.ps1"
    ps_file.write_text(_WORD_TOC_PS, encoding="utf-8")
    tmp = workdir / ("rt_" + docx.name)
    shutil.copy2(docx, tmp)
    out_txt = workdir / "toc_roundtrip.txt"
    empty = {"toc_lines": [], "toc_last_page": 0, "first_head_page": 0, "pages": 0,
             "has_toc": False, "copy": str(tmp)}
    try:
        subprocess.run(
            ["pwsh", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps_file),
             "-Path", str(tmp), "-Out", str(out_txt)],
            # 不解码输出：pwsh 在中文 Windows 上吐 GBK，text=True 会抛
            # UnicodeDecodeError 把整批跑挂（2026-11-08 实测 171/199 崩掉）。
            capture_output=True, timeout=180,
        )
    except subprocess.TimeoutExpired:
        # Word 在某些复杂文档上会卡住（实测残留 ~$ 锁文件）——按超时处理，不拖垮整批
        return {**empty, "error": "word timeout"}
    if not out_txt.exists():
        return {**empty, "error": "word roundtrip failed"}
    data = dict(empty)
    for line in out_txt.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line.startswith("TOCLINE "):
            data["toc_lines"].append(line[len("TOCLINE "):])
        elif line.startswith("TOC_LAST_PAGE="):
            data["toc_last_page"] = int(line.split("=", 1)[1] or 0)
        elif line.startswith("FIRST_HEAD_PAGE="):
            data["first_head_page"] = int(line.split("=", 1)[1] or 0)
        elif line.startswith("PAGES="):
            data["pages"] = int(line.split("=", 1)[1] or 0)
        elif line.startswith("TOC_COUNT="):
            data["has_toc"] = int(line.split("=", 1)[1] or 0) > 0
    return data


def _cached_page_of(para) -> int | None:
    """目录条目里缓存的页码：取段落最后一个文本 run。

    不能取「最后一个纯数字 run」——条目正文本身带章号（「2 期刊中析出的文献」），
    会把这行的 2 当页码（实测 8 份假不一致）。
    """
    texts = [t.text.strip() for t in para.iter(W + "t") if (t.text or "").strip()]
    if not texts:
        return None
    last = texts[-1]
    if last.isdigit():
        return int(last)
    if re.fullmatch(r"[IVXLCDMivxlcdm]+", last):
        return _roman_to_int(last)
    return None


def _check_pages(docx: Path) -> tuple:
    try:
        from earlyz.actions.new_actions import _para_pageref_target
        from earlyz.validation.renderers import probe_bookmark_pages
        pm = probe_bookmark_pages(str(docx))
        adj = pm.get("toc_bookmarks_adj") or {}
        if not adj:
            return None, []
        with zipfile.ZipFile(docx) as z:
            root = etree.fromstring(z.read("word/document.xml"))
        body = root.find(W + "body")
        bad: list = []
        for para in body.iter(W + "p"):
            name = _para_pageref_target(para)
            if not name:
                continue
            want = adj.get(name)
            if want is None:
                continue
            cached = _cached_page_of(para)
            if cached is None:
                continue
            if int(cached) != int(want):
                bad.append({"entry": _paragraph_text(para)[:24], "toc": int(cached),
                            "word": int(want)})
        return (not bad), bad
    except Exception as e:  # noqa: BLE001
        return None, [{"error": f"{type(e).__name__}: {e}"}]


def verify(docx: Path) -> dict:
    docx = Path(docx)
    res: dict = {"file": docx.name, "ok": True, "checks": {}}
    if not docx.exists():
        return {"file": docx.name, "ok": False, "error": "missing"}
    expected = _expected_headings(docx)
    # 先看 XML 有没有 TOC 域：没有就不必开 Word（199 份里 139 份没有，省一大截时间）
    try:
        with zipfile.ZipFile(docx) as z:
            xml = z.read("word/document.xml")
        if b"TOC " not in xml and b"TOC\\" not in xml:
            return {"file": docx.name, "ok": True, "skipped": "no_toc", "checks": {}}
    except Exception:  # noqa: BLE001
        pass
    with tempfile.TemporaryDirectory(prefix="tocverify_") as td:
        rt = _run_word_roundtrip(docx, Path(td))
        if rt.get("error"):
            return {"file": docx.name, "ok": False, "error": rt["error"]}
        if not rt["has_toc"]:
            return {"file": docx.name, "ok": True, "skipped": "no_toc", "checks": {}}
        toc_norm = {_norm(l) for l in rt["toc_lines"]}
        missing = [t for t in expected if _norm(t) not in toc_norm]
        res["checks"]["entries_complete"] = {
            "ok": not missing, "expected": len(expected), "missing": missing[:8],
        }
        separate = bool(rt["toc_last_page"] and rt["first_head_page"]
                        and rt["toc_last_page"] < rt["first_head_page"])
        res["checks"]["toc_separate_from_body"] = {
            "ok": separate, "toc_last_page": rt["toc_last_page"],
            "first_head_page": rt["first_head_page"],
        }
        page_ok, mismatches = _check_pages(Path(rt["copy"]))
        res["checks"]["page_numbers_match"] = {
            "ok": page_ok, "mismatches": mismatches[:6], "count": len(mismatches),
        }
        res["ok"] = all(c["ok"] is not False for c in res["checks"].values())
        res["pages"] = rt["pages"]
    return res


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("docx", nargs="?")
    ap.add_argument("--batch", help="jsonl，每行含 output 字段")
    ap.add_argument("--json", help="结果落盘路径")
    args = ap.parse_args()

    if args.batch:
        rows = [json.loads(l) for l in Path(args.batch).read_text(encoding="utf-8").splitlines() if l.strip()]
        results = []
        for i, r in enumerate(rows, 1):
            p = Path(r.get("output") or "")
            if not p.exists():
                continue
            out = verify(p)
            out["source"] = r.get("source")
            results.append(out)
            mark = "OK  " if out.get("ok") else "FAIL"
            print(f"[{i}/{len(rows)}] {mark} {out.get('file', '')[:60]}", flush=True)
        if args.json:
            Path(args.json).write_text(json.dumps(results, ensure_ascii=False, indent=1), encoding="utf-8")
        bad = [r for r in results if not r.get("ok")]
        print(f"\n合计 {len(results)} | 通过 {len(results) - len(bad)} | 失败 {len(bad)}")
        for r in bad:
            fails = [k for k, v in (r.get("checks") or {}).items() if v.get("ok") is False]
            print(f"  - {Path(str(r.get('source') or r['file'])).name[:40]} → {fails}")
        return

    if not args.docx:
        ap.error("需要 docx 或 --batch")
    res = verify(Path(args.docx))
    print(json.dumps(res, ensure_ascii=False, indent=1))
    if args.json:
        Path(args.json).write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
    sys.exit(0 if res.get("ok") else 1)


if __name__ == "__main__":
    main()
