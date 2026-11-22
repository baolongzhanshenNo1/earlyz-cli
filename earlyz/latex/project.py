"""LaTeX 工程入口：.zip / 文件夹 / 单个 .tex → LatexProject（主文件候选）→ select_main 分析。

规则来自 nightB/probe_mainfile.py 在 12 个真实工程上的验证（边界文档 §2.5）：
  R0 `% !TEX root = x.tex` 魔法注释直接定位
  R1 候选 = 去注释后同时含 \\documentclass 与 \\begin{document}
  R2 排除被其它 .tex \\input/\\include/\\subfile 引用的
  R3 test*/testfiles/doc/docs 目录降权；\\documentclass{ltxdoc}（宏包文档）排除
  R4 排序：顶层 > 文件名 main/thesis/paper/*-example > 引用章节多 > 体积大
唯一 → 自动选定；多个 → 列出候选，首位预选，由用户确认（不猜）。
"""
from __future__ import annotations

import glob
import os
import re
import shutil
import subprocess
import zipfile
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath

from earlyz.latex.front_fields import BUILTIN, FrontReport, extract_front

# zip 安全上限（防压缩炸弹 / 路径穿越）
MAX_FILES = 5000
MAX_TOTAL_BYTES = 512 * 1024 * 1024
MAX_RATIO = 200
MAX_CANDIDATES = 8

_NAME_PRI = re.compile(r"^(main|thesis|paper|manuscript|.*-example|.*_example)$", re.I)
_SKIP_DIRS = re.compile(r"^(test|tests|testfiles|doc|docs|\.git|build|out|_.*)$", re.I)
_REF_RE = re.compile(r"\\(?:input|include|subfile|InputIfFileExists)\s*\{([^}]+)\}")
_CLS_RE = re.compile(r"\\documentclass\s*(?:\[[^\]]*\])?\s*\{([^}]+)\}")
# 正文照常解析、无需 cls 源码的标准文档类
_STANDARD_CLS = {"article", "report", "book", "letter", "ctexart", "ctexrep", "ctexbook",
                 "ctexbeamer", "beamer", "amsart", "amsbook", "memoir", "scrartcl", "scrreprt",
                 "scrbook", "standalone"}
_GFX_EXT = ("", ".pdf", ".png", ".jpg", ".jpeg", ".eps", ".ps")


class ProjectError(ValueError):
    """面向用户的错误（中文文案，直接展示）。"""


def read_text(p: Path) -> tuple[str, str]:
    """→ (文本, 实际编码)。UTF-8 严格解码失败再试 GB18030（GBK 超集），都不行按 latin-1 兜底。"""
    b = p.read_bytes()
    if b.startswith(b"\xef\xbb\xbf"):
        return b[3:].decode("utf-8", errors="replace"), "utf-8"
    for enc in ("utf-8", "gb18030"):
        try:
            return b.decode(enc), enc
        except UnicodeDecodeError:
            pass
    return b.decode("latin-1"), "latin-1"


def strip_comments(t: str) -> str:
    return re.sub(r"(?m)(?<!\\)%.*$", "", t)


def declared_encoding(t: str) -> str | None:
    """源码声明的编码（去注释后）：inputenc 选项 / CJKutf8 / CJK 环境的 {GBK}{…}。"""
    t = strip_comments(t)
    if re.search(r"\\usepackage\s*(?:\[[^\]]*\])?\s*\{[^}]*\bCJKutf8\b", t):
        return "utf-8"
    m = re.search(r"\\usepackage\s*\[([^\]]+)\]\s*\{inputenc\}", t)
    if m:
        opt = m.group(1).strip().lower()
        return "utf-8" if opt in ("utf8", "utf8x") else opt
    if re.search(r"\\begin\{CJK\*?\}\s*(?:\[[^\]]*\])?\s*\{GBK\}", t):
        return "gbk"
    if re.search(r"\\begin\{CJK\*?\}\s*(?:\[[^\]]*\])?\s*\{UTF8\}", t):
        return "utf-8"
    return None


def _enc_family(e: str | None) -> str | None:
    if e is None:
        return None
    e = e.lower()
    return "gbk" if e in ("gbk", "gb2312", "gb18030", "cp936") else e


@dataclass
class MainCandidate:
    file: str                     # 相对工程根，正斜杠
    cls: str
    rule: str                     # R0 | R1-R4
    depth: int = 0
    name_pri: bool = False
    n_refs: int = 0
    size: int = 0
    skip_dir: bool = False


@dataclass
class LatexProject:
    root: Path
    kind: str                                         # zip | dir | tex
    tex_files: list[str] = field(default_factory=list)
    candidates: list[MainCandidate] = field(default_factory=list)
    raw_candidates: int = 0

    @property
    def unique(self) -> bool:
        return len(self.candidates) == 1

    def to_question(self) -> dict | None:
        """多候选 → dsh ask_user_question 同构的问题（推荐项排第一、label 带「（推荐）」）。"""
        if len(self.candidates) <= 1:
            return None
        opts = []
        for i, c in enumerate(self.candidates):
            desc = f"文档类 {c.cls}" if c.cls else "未声明文档类"
            if c.n_refs:
                desc += f"，引用 {c.n_refs} 个子文件"
            opts.append({"label": c.file + ("（推荐）" if i == 0 else ""), "description": desc})
        return {"id": "latex_main", "header": "主文件",
                "question": f"这个工程里有 {len(self.candidates)} 篇独立文档，要排哪一篇？",
                "options": opts, "multi_select": False}


def _safe_extract(zp: Path, dest: Path) -> None:
    """解压并拒绝路径穿越、绝对路径、链接、超量文件、压缩炸弹。"""
    try:
        zf = zipfile.ZipFile(zp)
    except zipfile.BadZipFile as e:
        raise ProjectError("压缩包已损坏或不是 zip 格式") from e
    with zf:
        infos = zf.infolist()
        if len(infos) > MAX_FILES:
            raise ProjectError(f"压缩包文件数超过 {MAX_FILES} 个，已拒绝")
        total = sum(i.file_size for i in infos)
        packed = max(1, sum(i.compress_size for i in infos))
        if total > MAX_TOTAL_BYTES or total / packed > MAX_RATIO:
            raise ProjectError("压缩包解压后体积异常（超过 512MB 或压缩比过高），已拒绝")
        root = dest.resolve()
        for info in infos:
            name = info.filename.replace("\\", "/")
            # 非 UTF-8 标记的中文文件名：zipfile 按 cp437 解出，还原成 GBK（Windows 压缩软件默认）
            if not (info.flag_bits & 0x800):
                try:
                    name = name.encode("cp437").decode("gb18030")
                except (UnicodeEncodeError, UnicodeDecodeError):
                    pass
            pp = PurePosixPath(name)
            if pp.is_absolute() or ".." in pp.parts or re.match(r"^[A-Za-z]:", name):
                raise ProjectError("压缩包里有指向工程外的路径，已拒绝")
            if (info.external_attr >> 16) & 0o170000 == 0o120000:
                raise ProjectError("压缩包里有符号链接，已拒绝")
            target = (root / Path(*pp.parts)).resolve()
            if root not in target.parents and target != root:
                raise ProjectError("压缩包里有指向工程外的路径，已拒绝")
            if info.is_dir():
                target.mkdir(parents=True, exist_ok=True)
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info) as src, open(target, "wb") as out:
                shutil.copyfileobj(src, out, 1 << 20)


def _single_top_dir(root: Path) -> Path:
    """Overleaf/GitHub 导出常见「zip 内只有一层目录」→ 以该目录为工程根。"""
    kids = [k for k in root.iterdir() if not k.name.startswith(("__MACOSX", "."))]
    if len(kids) == 1 and kids[0].is_dir():
        return kids[0]
    return root


def open_project(src: str | Path, workdir: str | Path) -> LatexProject:
    """src = .zip / 文件夹 / .tex。zip 解到 workdir/<stem>/；文件夹与 .tex 原地只读，不复制。"""
    src = Path(src)
    if not src.exists():
        raise ProjectError("找不到上传的工程")
    if src.is_dir():
        root, kind = src, "dir"
    elif src.suffix.lower() == ".zip":
        dest = Path(workdir) / (src.stem[:40] or "project")
        if dest.exists():
            shutil.rmtree(dest)
        dest.mkdir(parents=True)
        _safe_extract(src, dest)
        root, kind = _single_top_dir(dest), "zip"
    elif src.suffix.lower() == ".tex":
        root, kind = src.parent, "tex"
    else:
        raise ProjectError("只支持 .zip 工程包、文件夹或 .tex 文件")
    proj = LatexProject(root=root, kind=kind)
    _detect(proj, only=src if kind == "tex" else None)
    if not proj.candidates:
        raise ProjectError("工程里没有找到主文件（需要同时含 \\documentclass 和 \\begin{document} 的 .tex）")
    return proj


def _detect(proj: LatexProject, only: Path | None) -> None:
    root = proj.root
    texs = sorted(p for p in root.rglob("*.tex") if p.is_file() and not p.name.startswith("_ez"))
    proj.tex_files = [p.relative_to(root).as_posix() for p in texs]
    if only is not None:
        texs_scan = [only]
    else:
        texs_scan = texs
    text = {p: read_text(p)[0] for p in texs}
    if only is not None and only not in text:
        text[only] = read_text(only)[0]
    # R0：魔法注释
    for p in texs_scan:
        m = re.search(r"(?mi)^\s*%\s*!\s*TEX\s+root\s*=\s*(\S+)", text[p])
        if m:
            tgt = (p.parent / m.group(1)).resolve()
            if tgt.is_file():
                t = text.get(tgt) or read_text(tgt)[0]
                cm = _CLS_RE.search(strip_comments(t))
                proj.candidates = [MainCandidate(file=_rel(root, tgt), cls=cm.group(1).strip() if cm else "",
                                                 rule="R0")]
                proj.raw_candidates = 1
                return
    referenced: set[Path] = set()
    for p, t in text.items():
        tb = strip_comments(t)
        # 只数「文档」里的引用：AAS 的 rb_template.tex 只有 \RequirePackage{latexrelease}+\input{template.tex}，
        # 是包装壳，不能因此把真正的 template.tex 排除掉（2026-09-30 实测误选 example/）
        if not (re.search(r"\\documentclass", tb) or re.search(r"\\begin\{document\}", tb)):
            continue
        for ref in _REF_RE.findall(tb):
            for name in ref.split(","):
                q = p.parent / name.strip()
                referenced.add((q if q.suffix else q.with_suffix(".tex")).resolve())
                # \input 路径相对主文件目录（TeX 语义），也按工程根再解析一次
                q2 = root / name.strip()
                referenced.add((q2 if q2.suffix else q2.with_suffix(".tex")).resolve())
    cands = []
    for p in texs_scan:
        body = strip_comments(text[p])
        if not (re.search(r"\\documentclass", body) and re.search(r"\\begin\{document\}", body)):
            continue
        cm = _CLS_RE.search(body)
        rel = p.relative_to(root)
        cands.append(MainCandidate(
            file=rel.as_posix(), cls=cm.group(1).strip() if cm else "", rule="R1-R4",
            depth=len(rel.parts) - 1, name_pri=bool(_NAME_PRI.match(p.stem)),
            n_refs=len(_REF_RE.findall(body)), size=len(text[p]),
            skip_dir=any(_SKIP_DIRS.match(x) for x in rel.parts[:-1]),
        ))
    proj.raw_candidates = len(cands)
    live = [c for c in cands if (root / c.file).resolve() not in referenced and c.cls != "ltxdoc"]
    top = [c for c in live if not c.skip_dir] or live
    top.sort(key=lambda c: (c.depth, not c.name_pri, -c.n_refs, -c.size))
    proj.candidates = top[:MAX_CANDIDATES]


def _rel(root: Path, p: Path) -> str:
    try:
        return p.resolve().relative_to(root.resolve()).as_posix()
    except ValueError:
        return p.name


# ── 选定主文件后的分析 ───────────────────────────────────────────────


@dataclass
class ClsInfo:
    name: str
    where: str            # project | tex | standard | builtin_map | unknown
    path: str = ""        # project 内相对路径 / kpsewhich 绝对路径


@dataclass
class MainAnalysis:
    main: str
    cls: ClsInfo
    encoding: str                         # 实际编码
    declared_encoding: str | None         # 源码声明
    encoding_mismatch: bool
    included: list[str] = field(default_factory=list)       # 展开到的子文件（相对根）
    missing: list[dict] = field(default_factory=list)       # {kind, name, from}
    front: FrontReport | None = None
    tex_available: bool = False
    notes: list[str] = field(default_factory=list)          # 面向用户的说明（中文）

    def to_dict(self) -> dict:
        return {
            "main": self.main, "cls": vars(self.cls), "encoding": self.encoding,
            "declared_encoding": self.declared_encoding, "encoding_mismatch": self.encoding_mismatch,
            "included": self.included, "missing": self.missing,
            "front": vars(self.front) if self.front else None,
            "tex_available": self.tex_available, "notes": self.notes,
        }


def find_tex_tool(name: str = "kpsewhich") -> str | None:
    """用户本机 TeX（TeX Live / MiKTeX）：PATH 优先，其次常见安装位置。找不到返回 None（不打包 TeX）。"""
    exe = shutil.which(name)
    if exe:
        return exe
    pats = [r"C:\texlive\*\bin\windows", r"C:\texlive\*\bin\win64", r"C:\texlive\*\bin\win32",
            r"D:\texlive\*\bin\windows", r"E:\texlive\*\bin\windows",
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\MiKTeX\miktex\bin\x64"),
            r"C:\Program Files\MiKTeX\miktex\bin\x64"]
    for pat in pats:
        for d in sorted(glob.glob(pat), reverse=True):
            cand = Path(d) / f"{name}.exe"
            if cand.is_file():
                return str(cand)
    return None


def _kpsewhich(fname: str) -> str | None:
    exe = find_tex_tool("kpsewhich")
    if not exe:
        return None
    try:
        r = subprocess.run([exe, fname], capture_output=True, text=True, timeout=20,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except (OSError, subprocess.TimeoutExpired):
        return None
    out = (r.stdout or "").strip().splitlines()
    return out[0].strip() if r.returncode == 0 and out else None


def locate_cls(proj_root: Path, main_dir: Path, cls: str) -> ClsInfo:
    """cls 定位顺序：工程内（主文件目录 → 工程任意位置）→ 标准类 → 本机 TeX → 内置映射 → 未知。"""
    if not cls:
        return ClsInfo(name="", where="unknown")
    fname = cls + ".cls"
    for d in (main_dir, proj_root):
        p = d / fname
        if p.is_file():
            return ClsInfo(cls, "project", _rel(proj_root, p))
    hits = [p for p in proj_root.rglob(fname) if p.is_file()]
    if hits:
        return ClsInfo(cls, "project", _rel(proj_root, hits[0]))
    if cls in _STANDARD_CLS:
        return ClsInfo(cls, "standard")
    hit = _kpsewhich(fname)
    if hit:
        return ClsInfo(cls, "tex", hit)
    if cls in BUILTIN:
        return ClsInfo(cls, "builtin_map")
    return ClsInfo(cls, "unknown")


def _inside(p: Path, root: Path) -> bool:
    try:
        p.resolve().relative_to(root.resolve())
        return True
    except (ValueError, OSError):
        return False


def _resolve_ref(base: Path, root: Path, name: str, exts: tuple[str, ...]) -> Path | None:
    """解析 \\input/\\include/bib 引用；只认工程根内的文件（绝对路径、../ 越界一律视为缺失——
    工程来自用户上传的 zip/文件夹，不能借 \\input{C:/...} 读本机任意文件；与 document.py 口径一致）。"""
    name = name.strip().strip('"')
    for d in (base, root):
        for ext in exts:
            p = d / (name + ext if ext and not name.lower().endswith(ext) else name)
            if p.is_file() and _inside(p, root):
                return p
    return None


def _expand(p: Path, root: Path, seen: set[Path], included: list[str], missing: list[dict],
            depth: int = 0) -> str:
    """展开 \\input/\\include/\\subfile（深度 ≤ 12，防循环）；找不到的子文件记入 missing。"""
    t, _ = read_text(p)
    t = strip_comments(t)
    if depth > 12:
        return t
    base = p.parent

    def sub(m: re.Match) -> str:
        out = []
        for name in m.group(1).split(","):
            q = _resolve_ref(base, root, name, (".tex", ""))
            if q is None:
                missing.append({"kind": "tex", "name": name.strip(), "from": _rel(root, p)})
                continue
            rq = q.resolve()
            if rq in seen:
                continue
            seen.add(rq)
            included.append(_rel(root, q))
            out.append(_expand(q, root, seen, included, missing, depth + 1))
        return "\n".join(out)

    return _REF_RE.sub(sub, t)


def _graphics_paths(text: str) -> list[str]:
    m = re.search(r"\\graphicspath\s*\{((?:\{[^}]*\})+)\}", text)
    return re.findall(r"\{([^}]*)\}", m.group(1)) if m else []


def _missing_assets(text: str, main_dir: Path, root: Path, main_rel: str) -> list[dict]:
    miss = []
    gdirs = [main_dir] + [main_dir / g for g in _graphics_paths(text)] + [root]
    for m in re.finditer(r"\\(?:includegraphics|psfig|epsfig)\*?\s*(?:\[[^\]]*\])?\s*\{([^}]+)\}", text):
        arg = m.group(1).strip()
        if "file=" in arg or "figure=" in arg:           # \psfig{file=x.eps,width=…}
            fm = re.search(r"(?:file|figure)\s*=\s*([^,}\s]+)", arg)
            arg = fm.group(1) if fm else arg
        if "#" in arg or "\\" in arg:                      # 宏参数/命令拼路径，静态查不了
            continue
        if not any((d / (arg + e)).is_file() for d in gdirs for e in _GFX_EXT):
            miss.append({"kind": "image", "name": arg, "from": main_rel})
    for m in re.finditer(r"\\(?:bibliography|addbibresource)\s*(?:\[[^\]]*\])?\s*\{([^}]+)\}", text):
        for name in m.group(1).split(","):
            name = name.strip()
            if name and _resolve_ref(main_dir, root, name, (".bib", "")) is None:
                miss.append({"kind": "bib", "name": name, "from": main_rel})
    # \usepackage 不核对：宏包归 TeX 发行版，工程里没有不代表缺（场景 B 编译时由 TeX 报）
    # 去重，保持顺序
    seen, out = set(), []
    for x in miss:
        k = (x["kind"], x["name"])
        if k not in seen:
            seen.add(k)
            out.append(x)
    return out


_WHERE_NOTE = {
    "project": "文档类 {cls} 在工程里（{path}），按工程自带的定义解析。",
    "standard": "文档类 {cls} 是标准文档类，直接解析。",
    "tex": "文档类 {cls} 不在工程里，已在本机 TeX 发行版找到（{path}）。",
    "builtin_map": ("工程里没有 {cls}.cls，本机也没有安装 TeX，已按内置的 {cls} 字段对照表读取题目、"
                    "作者、摘要等封面信息；正文的章节、公式、图表、参考文献照常解析。"),
    "unknown": ("工程里没有 {cls}.cls，本机也没有安装 TeX，内置对照表里也没有这个文档类。"
                "正文的章节、公式、图表、参考文献照常解析；封面信息只能按通用写法"
                "（\\title、\\author、abstract 环境）读取，读不到的字段列在下面，请核对。"),
}


def select_main(proj: LatexProject, main: str | None = None) -> MainAnalysis:
    """main=None 时取首位候选（唯一候选或用户已确认）；main 必须是候选之一（防任意路径）。"""
    files = {c.file: c for c in proj.candidates}
    if main is None:
        cand = proj.candidates[0]
    else:
        main = re.sub(r"（推荐）$", "", main.strip())
        if main not in files:
            raise ProjectError("选中的主文件不在候选列表里")
        cand = files[main]
    root = proj.root
    mp = root / cand.file
    raw, enc = read_text(mp)
    decl = declared_encoding(raw)
    included: list[str] = []
    missing: list[dict] = []
    full = _expand(mp, root, {mp.resolve()}, included, missing)
    cls = locate_cls(root, mp.parent, cand.cls)
    ana = MainAnalysis(
        main=cand.file, cls=cls, encoding=enc, declared_encoding=decl,
        encoding_mismatch=bool(decl) and _enc_family(decl) != _enc_family(enc) and enc != "latin-1",
        included=included, tex_available=find_tex_tool("kpsewhich") is not None,
    )
    ana.missing = missing + _missing_assets(full, mp.parent, root, cand.file)
    ana.front = extract_front(full, cls.name)
    if cls.name:
        ana.notes.append(_WHERE_NOTE[cls.where].format(cls=cls.name, path=cls.path))
    if ana.encoding_mismatch:
        ana.notes.append(f"主文件实际是 {enc.upper()} 编码，但源码声明的是 {decl.upper()}，已按实际编码读取。")
    if cls.where in ("builtin_map", "unknown") and ana.front and ana.front.unresolved:
        ana.notes.append("封面区有这些命令没能识别，内容可能没读到：" + "、".join(ana.front.unresolved[:12]))
    if ana.front and ana.front.missing and cls.where in ("builtin_map", "unknown"):
        ana.notes.append("没读到的封面字段：" + "、".join(ana.front.missing))
    if ana.missing:
        kinds = {"tex": "子文件", "image": "图片", "bib": "参考文献库"}
        by = {}
        for x in ana.missing:
            by.setdefault(kinds.get(x["kind"], x["kind"]), []).append(x["name"])
        ana.notes.append("工程里缺这些文件：" + "；".join(f"{k} {len(v)} 个（{'、'.join(v[:5])}{'…' if len(v) > 5 else ''}）"
                                                    for k, v in by.items()))
    return ana

