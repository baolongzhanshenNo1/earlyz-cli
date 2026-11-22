"""标题编号完整性验收（2026-11-08 新判据）。

为什么单独成模块：`scripts/score_layout.py` 里 `numbering` / `numPr` 出现 0 次，
199 份基线因此对「1. 1. 引言」「二三级 1.27cm 缩进」完全免疫——判据里根本没有
编号这一维。这里补三条，全部以「Word 打开后看到的样子」为准：

  A 引用闭合：styles.xml 引用的 numId 必须在 numbering.xml 解析得到
    （悬空 numId 会让 Word 退回默认列表：多画一个号 + 默认 720twips 缩进）；
  B 不重复：标题段落既被画出自动编号、文本里又自带序号 → 就是「1. 1.」；
  C 无异常缩进：一级/二级/三级标题左缩进不得凭空出现（阈值 1.0cm）。

B/C 需要 Word 真机读数（ListFormat.ListString / LeftIndent），单份约 3~5 秒。
"""
from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
# 序号前缀：三种形态都算（2026-11-08 实测漏判）
#   "2.1 基于优化" / "1.1编写目的"（多级，可无空格） / "1.  范围"（数字+空格） / "3．方法"
# 反例必须不匹配："2026年工作总结"（数字后直接跟汉字且无分隔）
_NUM_PREFIX_RE = re.compile(r"^\s*(?:\d+(?:\s*[.．]\s*\d+)+|\d+\s*[.．、]|\d+\s)\s*")

_WORD_PS = r"""param([string]$Path, [string]$Out)
$ErrorActionPreference = 'Stop'
$lines = New-Object System.Collections.Generic.List[string]
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$doc = $w.Documents.Open($Path, $false, $true)
try {
    $doc.Repaginate()
    foreach ($p in $doc.Paragraphs) {
        if ($p.OutlineLevel -ge 1 -and $p.OutlineLevel -le 3) {
            $txt = $p.Range.Text.Trim()
            if ($txt.Length -eq 0) { continue }
            if ($txt.Length -gt 40) { $txt = $txt.Substring(0, 40) }
            $ls = ""
            try { $ls = $p.Range.ListFormat.ListString } catch {}
            $li = [math]::Round($p.LeftIndent / 28.3465, 2)
            $lines.Add("H`t$($p.OutlineLevel)`tlist=$ls`tleft=$li`t$txt")
        }
    }
} finally {
    try { $doc.Close($false) } catch {}
    try { $w.Quit() } catch {}
}
[System.IO.File]::WriteAllLines($Out, $lines, (New-Object System.Text.UTF8Encoding($false)))
"WROTE " + $Out
"""


def _styles_numbering_refs(docx: Path) -> set[str]:
    with zipfile.ZipFile(docx) as z:
        if "word/styles.xml" not in z.namelist():
            return set()
        root = etree.fromstring(z.read("word/styles.xml"))
    out: set[str] = set()
    for np in root.iter(f"{W}numPr"):
        nid = np.find(f"{W}numId")
        if nid is not None:
            val = nid.get(f"{W}val")
            if val and val != "0":
                out.add(val)
    return out


def _numbering_ids(docx: Path) -> set[str]:
    with zipfile.ZipFile(docx) as z:
        if "word/numbering.xml" not in z.namelist():
            return set()
        root = etree.fromstring(z.read("word/numbering.xml"))
    return {n.get(f"{W}numId") for n in root.findall(f"{W}num")}


def check_refs(docx: Path) -> dict:
    refs = _styles_numbering_refs(docx)
    have = _numbering_ids(docx)
    return {"ok": not (refs - have), "referenced": sorted(refs), "missing": sorted(refs - have)}


def _word_headings(docx: Path, workdir: Path) -> list[dict]:
    ps_file = workdir / "heading_probe.ps1"
    ps_file.write_text(_WORD_PS, encoding="utf-8")
    tmp = workdir / ("hp_" + docx.name)
    shutil.copy2(docx, tmp)
    out_txt = workdir / "heads.txt"
    try:
        subprocess.run(["pwsh", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                        str(ps_file), "-Path", str(tmp), "-Out", str(out_txt)],
                       capture_output=True, timeout=180)   # 不解码：pwsh 中文输出是 GBK
    except subprocess.TimeoutExpired:
        return [{"error": "word timeout"}]
    if not out_txt.exists():
        return [{"error": "word probe failed"}]
    rows = []
    for line in out_txt.read_text(encoding="utf-8").splitlines():
        parts = line.split("\t")
        if len(parts) < 5 or not parts[0].startswith("H"):
            continue
        liststr = left = ""
        for seg in parts[2:4]:
            if seg.startswith("list="):
                liststr = seg[5:]
            elif seg.startswith("left="):
                left = seg[5:]
        rows.append({"level": parts[1], "list": liststr, "left_cm": left,
                     "text": parts[4] if len(parts) > 4 else ""})
    return rows


def verify(docx: Path) -> dict:
    docx = Path(docx)
    res: dict = {"file": docx.name, "ok": True, "checks": {}}
    if not docx.exists():
        return {"file": docx.name, "ok": False, "error": "missing"}
    res["checks"]["styles_numbering_closed"] = check_refs(docx)
    with tempfile.TemporaryDirectory(prefix="numverify_") as td:
        rows = _word_headings(docx, Path(td))
    if any(r.get("error") for r in rows):
        res["checks"]["heading_numbers"] = {"ok": None, "error": rows[0]["error"]}
        res["ok"] = res["checks"]["styles_numbering_closed"]["ok"] is not False
        return res
    dup: list[dict] = []
    indented: list[dict] = []
    for r in rows:
        text = r.get("text", "")
        if r.get("list") and _NUM_PREFIX_RE.match(text):
            dup.append({"text": text[:26], "list": r["list"], "level": r["level"]})
        try:
            left = float(r.get("left_cm") or 0)
        except ValueError:
            left = 0.0
        if left >= 1.0:
            indented.append({"text": text[:26], "left_cm": left, "level": r["level"]})
    res["checks"]["heading_numbers"] = {
        "ok": not dup, "duplicates": dup[:8], "count": len(dup),
        "indented": indented[:8], "indented_count": len(indented),
        "headings": len(rows),
    }
    res["ok"] = all(c.get("ok") is not False for c in res["checks"].values())
    return res
