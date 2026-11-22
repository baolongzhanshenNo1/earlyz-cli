"""前置字段（题目/作者/摘要/关键词…）：内置映射 + 抽取。

用途：工程里找不到 cls、本机也没有 TeX 时，按 cls 名走内置映射读封面信息；
识别不了的前置命令列给用户（不静默丢）。映射逐条来自素材实测（2026-09-30，
D:\\Desktop\\期刊模板\\Latex），不是凭印象写的：
- thuthesis / ustcthesis / fduthesis / sjtuthesis：\\xxxsetup{key = value} 键值式
  （sjtu 是 info = { zh / title = … }，fdu 是 info = { title = … }）
- xautthesis：\\ctitle \\cauthor … 命令式 + cabstract/eabstract 环境
- aas（中文）：\\cntitle \\cnauthor \\cnabstract …；aase/aasen（英文）：\\title \\abstract \\keyword
- cas-sc / cas-dc（Elsevier）：\\title \\author \\affiliation + abstract/highlights/keywords 环境
- CjC（计算机学报）：**没有字段命令**，前置信息是 CJK 环境 + \\zihao 手排，只能按
  「摘  要」「关键词」「中图法分类号」这类版面标签定位 → kind="typeset"
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

# 标准字段名 → 显示名（面向用户的中文，不出现命令名）
FIELD_LABELS = {
    "title": "中文题目", "title_en": "英文题目", "author": "作者", "author_en": "英文作者",
    "supervisor": "导师", "supervisor_en": "英文导师名", "department": "院系",
    "discipline": "学科专业", "major": "专业", "degree": "学位", "student_id": "学号",
    "affiliation": "作者单位", "affiliation_en": "英文作者单位",
    "abstract": "中文摘要", "abstract_en": "英文摘要",
    "keywords": "中文关键词", "keywords_en": "英文关键词",
    "clc": "中图分类号", "udc": "UDC", "doi": "DOI", "date": "日期", "funding": "基金/脚注",
    "citation": "中文引用格式", "citation_en": "英文引用格式",
}


@dataclass
class ClsMap:
    kind: str                                   # setup | command | typeset
    setup: str = ""                             # \xxxsetup 命令名
    keys: dict[str, str] = field(default_factory=dict)       # 原键/命令 → 标准字段
    envs: dict[str, str] = field(default_factory=dict)       # 环境名 → 标准字段
    labels: dict[str, str] = field(default_factory=dict)     # typeset：版面标签 → 标准字段


_STD_CMDS = {"title": "title", "author": "author", "date": "date", "thanks": "funding"}
_STD_ENVS = {"abstract": "abstract"}

BUILTIN: dict[str, ClsMap] = {
    "thuthesis": ClsMap("setup", "thusetup", {
        "title": "title", "title*": "title_en", "author": "author", "author*": "author_en",
        "supervisor": "supervisor", "supervisor*": "supervisor_en",
        "associate-supervisor": "supervisor", "department": "department",
        "discipline": "discipline", "discipline*": "discipline",
        "degree-category": "degree", "degree-category*": "degree",
        "keywords": "keywords", "keywords*": "keywords_en"},
        {"abstract": "abstract", "abstract*": "abstract_en"}),
    "ustcthesis": ClsMap("setup", "ustcsetup", {
        "title": "title", "title*": "title_en", "author": "author", "department": "department",
        "discipline": "discipline", "research-direction": "discipline", "supervisor": "supervisor",
        "student-id": "student_id", "clc": "clc", "udc": "udc", "date": "date",
        "keywords": "keywords", "keywords*": "keywords_en"},
        {"abstract": "abstract", "abstract*": "abstract_en"}),
    "fduthesis": ClsMap("setup", "fdusetup", {
        "title": "title", "title*": "title_en", "author": "author", "supervisor": "supervisor",
        "major": "major", "degree": "degree", "department": "department",
        "student-id": "student_id", "keywords": "keywords", "keywords*": "keywords_en"},
        {"abstract": "abstract", "abstract*": "abstract_en"}),
    "sjtuthesis": ClsMap("setup", "sjtusetup", {
        "zh/title": "title", "en/title": "title_en", "zh/author": "author", "en/author": "author_en",
        "zh/supervisor": "supervisor", "en/supervisor": "supervisor_en", "id": "student_id",
        "zh/department": "department", "en/department": "department",
        "zh/major": "major", "en/major": "major", "zh/degree": "degree", "en/degree": "degree",
        "zh/keywords": "keywords", "en/keywords": "keywords_en"},
        {"abstract": "abstract"}),
    "xautthesis": ClsMap("command", "", {
        "ctitle": "title", "etitle": "title_en", "cauthor": "author", "csupervisor": "supervisor",
        "cdepartment": "department", "cdiscipline": "discipline", "cmajor": "major",
        "cdate": "date", "studentid": "student_id", "classification": "clc",
        "ckeywords": "keywords", "ekeywords": "keywords_en"},
        {"cabstract": "abstract", "eabstract": "abstract_en"}),
    "aas": ClsMap("command", "", {
        "cntitle": "title", "entitle": "title_en", "cnauthor": "author", "enauthor": "author_en",
        # cnaddress/enaddress 在 aas.cls 里是「引用本文」格式（\bf Citation），不是作者单位；
        # 单位写在 \thanks 里（2026-09-30 读 aas.cls:155 实测）
        "cnaddress": "citation", "enaddress": "citation_en", "cnabstract": "abstract",
        "enabstract": "abstract_en", "cnkeyword": "keywords", "enkeyword": "keywords_en",
        "doi": "doi", "thanks": "funding"}),
    "aase": ClsMap("command", "", {
        "title": "title_en", "author": "author_en", "abstract": "abstract_en",
        "keyword": "keywords_en", "doi": "doi", "thanks": "funding"}),
    "cas-sc": ClsMap("command", "", {
        "title": "title_en", "author": "author_en", "affiliation": "affiliation_en",
        "tnotetext": "funding"},
        {"abstract": "abstract_en", "keywords": "keywords_en", "highlights": "abstract_en"}),
    "CjC": ClsMap("typeset", labels={
        "摘要": "abstract", "关键词": "keywords", "中图法分类号": "clc", "DOI号": "doi",
        "Abstract": "abstract_en", "Keywords": "keywords_en"}),
}
BUILTIN["aasen"] = BUILTIN["aase"]
BUILTIN["cas-dc"] = BUILTIN["cas-sc"]

# 前置区里常见、但不承载字段内容的排版命令（出现不算「识别不了」）；取自素材前置区实测
_LAYOUT_CMDS = {
    "maketitle", "begin", "end", "vspace", "vskip", "hspace", "quad", "qquad", "centering",
    "pagestyle", "thispagestyle", "setcounter", "setlength", "zihao", "selectfont", "textbf",
    "textsl", "textit", "rm", "sl", "hei", "song", "fs", "kai", "noindent", "par", "hline",
    "linespread", "multicolumn", "rule", "footnote", "mbox", "makeatletter", "makeatother",
    "newcommand", "renewcommand", "newenvironment", "let", "def", "include", "input",
    "frontmatter", "mainmatter", "tableofcontents", "listoffigures", "listoftables",
    "includegraphics", "cortext", "ead", "credit", "shorttitle", "shortauthors", "nonumnote",
    "fntext", "printSq", "cit", "hfill", "ldots", "scriptscriptstyle", "sep", "nocite",
    "copyrightpage", "makedeclaration", "makeatother", "hyphenpenalty", "captionsetup",
    # 通用标记命令（不承载封面字段）
    "label", "ref", "eqref", "cite", "item", "text", "texttt", "textrm", "textsf", "emph", "url",
    "href", "footnotemark", "footnotetext", "boldsymbol", "mathbf", "mathrm", "textasciitilde",
    "newcolumntype", "tnotemark", "cormark", "fnmark", "tnoteref", "fnref", "small", "large",
}


def _group(t: str, i: int) -> tuple[str, int]:
    """t[i] == '{'：返回配平内容与结束位置；越界时返回剩余文本。"""
    depth, j = 0, i
    while j < len(t):
        c = t[j]
        if c == "\\":
            j += 2
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return t[i + 1:j], j + 1
        j += 1
    return t[i + 1:], len(t)


def _skip_ws_opt(t: str, i: int) -> int:
    """跳过空白和一个可选参数 [..]。"""
    while i < len(t) and t[i] in " \t\r\n":
        i += 1
    if i < len(t) and t[i] == "[":
        depth = 0
        while i < len(t):
            if t[i] == "[":
                depth += 1
            elif t[i] == "]":
                depth -= 1
                if depth == 0:
                    i += 1
                    break
            i += 1
        while i < len(t) and t[i] in " \t\r\n":
            i += 1
    return i


def _kv_pairs(body: str, prefix: str = "") -> list[tuple[str, str]]:
    """key = value 列表（顶层逗号切分）；value 为 {…} 且内含 = 时递归展开（sjtu/fdu 的 info = {…}）。"""
    out, depth, cur = [], 0, []
    for c in body + ",":
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
        if c == "," and depth == 0:
            s = "".join(cur).strip()
            cur = []
            m = re.match(r"([A-Za-z][\w*\-]*(?:\s*/\s*[A-Za-z][\w*\-]*)?)\s*=\s*(.*)$", s, re.S)
            if not m:
                continue
            key = re.sub(r"\s*/\s*", "/", m.group(1))
            val = m.group(2).strip()
            if val.startswith("{") and val.endswith("}"):
                inner = val[1:-1]
                if re.search(r"(?m)^\s*[A-Za-z][\w*\-]*(?:\s*/\s*\w+)?\s*=", inner):
                    out += _kv_pairs(inner)
                    continue
                val = inner
            out.append((prefix + key, val.strip()))
        else:
            cur.append(c)
    return out


def _drop_groups(v: str, names: tuple[str, ...]) -> str:
    """去掉 \\name{…} 整组（配平）：AAS 的 \\cntitle{题目 \\thanks{收稿日期…}} 只要题目。"""
    out, i = [], 0
    pat = re.compile(r"\\(?:" + "|".join(names) + r")\s*(?=\{)")
    while True:
        m = pat.search(v, i)
        if not m:
            out.append(v[i:])
            return "".join(out)
        out.append(v[i:m.start()])
        _, i = _group(v, m.end())


def _clean(v: str, limit: int = 60) -> str:
    v = _drop_groups(v, ("thanks", "tnoteref", "fnref", "cortext", "footnote"))
    v = re.sub(r"\\[A-Za-z]+\*?", " ", v)
    v = re.sub(r"[{}$^_~]|\\.", " ", v)
    v = re.sub(r"\s+", " ", v).strip()
    return v[:limit]


@dataclass
class FrontReport:
    source: str                         # builtin:<cls> | generic
    fields: dict[str, str] = field(default_factory=dict)      # 标准字段 → 取到的值（截断）
    unresolved: list[str] = field(default_factory=list)       # 前置区里认不出的命令（去重）
    missing: list[str] = field(default_factory=list)          # 映射里有、文档里没取到的字段（显示名）


def extract_front(text: str, cls: str, limit: int = 60) -> FrontReport:
    """text = 主文件（已展开 \\input、已去注释）。前置区 = 导言区 + \\begin{document} 到首个章节。"""
    cmap = BUILTIN.get(cls)
    rep = FrontReport(source=f"builtin:{cls}" if cmap else "generic")
    bi = text.find("\\begin{document}")
    pre = text[:bi] if bi >= 0 else text
    rest = text[bi:] if bi >= 0 else ""
    m = re.search(r"\\(?:chapter|section|mainmatter)\b", rest)
    front_body = rest[: m.start()] if m else rest[:8000]
    region = pre + "\n" + front_body

    keys = dict(_STD_CMDS)
    envs = dict(_STD_ENVS)
    if cmap:
        envs.update(cmap.envs)
        if cmap.kind == "command":
            keys.update(cmap.keys)
    known_cmds = set(keys)

    def put(fid: str, val: str) -> None:
        val = _clean(val, limit)
        if val and fid not in rep.fields:
            rep.fields[fid] = val

    # 键值式 \xxxsetup{…}（导言区和 thusetup.tex 这类被 \input 的配置文件都已展开进 text）
    if cmap and cmap.kind == "setup":
        known_cmds.add(cmap.setup)
        # 全文找：thu 的 keywords 写在 data/abstract.tex 的 \thusetup 里（\frontmatter 之后）
        for sm in re.finditer(r"\\" + re.escape(cmap.setup) + r"\s*(?=\{)", text):
            body, _ = _group(text, sm.end())
            for k, v in _kv_pairs(body):
                if k in cmap.keys:
                    put(cmap.keys[k], v)
        # 摘要在正文前部之外的单独文件里（thu data/abstract.tex）：全文找环境
        region_env = text
    else:
        region_env = region
    # 命令式
    for cm in re.finditer(r"\\([A-Za-z]+\*?)", region):
        name = cm.group(1)
        if name in keys:
            j = _skip_ws_opt(region, cm.end())
            if j < len(region) and region[j] == "{":
                val, _ = _group(region, j)
                put(keys[name], val)
    # 环境式
    for env, fid in envs.items():
        em = re.search(r"\\begin\{" + re.escape(env) + r"\}(?:\[[^\]]*\])?(.*?)\\end\{" + re.escape(env) + r"\}",
                       region_env, re.S)
        if em:
            put(fid, em.group(1))
    # sjtu：abstract[zh] / abstract[en] 用可选参数区分语言
    if cls == "sjtuthesis":
        for lang, fid in (("zh", "abstract"), ("en", "abstract_en")):
            em = re.search(r"\\begin\{abstract\}\[" + lang + r"\](.*?)\\end\{abstract\}", text, re.S)
            if em:
                rep.fields[fid] = _clean(em.group(1), limit)
    # 版面标签式（CjC）：标签后面的正文。先去掉字体切换环境的「\begin{CJK*}{GBK}{song}」
    # 整组参数（否则值读成「CJK* GBK song」），再去命令和花括号。
    if cmap and cmap.kind == "typeset":
        plain = re.sub(r"\\(?:begin|end)\{CJK\*?\}(?:\{[^}]*\}){0,2}", " ", front_body)
        # 只吃命令自己的参数：\zihao{5-}{…} 后面那组是被排版的内容，不能一起删
        plain = re.sub(r"\\(?:zihao|vspace|vskip|hspace)\s*\{[^}]*\}", " ", plain)
        plain = re.sub(r"\\setlength\s*\{[^}]*\}\s*\{[^}]*\}", " ", plain)
        plain = re.sub(r"\\[A-Za-z]+\*?|[{}]|\\.", " ", plain)
        plain = re.sub(r"\s+", " ", plain)
        for label, fid in cmap.labels.items():
            pat = r"\s*".join(map(re.escape, label))
            # 值截到下一个版面标签 / 表格残留词为止
            lm = re.search(pat + r"[\s:：]*(.{1,120}?)(?=\s(?:tabular|摘\s*要|关键词|中图法分类号|DOI号|Key\s*words|Abstract)\b|$)",
                           plain)
            if lm:
                put(fid, lm.group(1))

    # 识别不了的：正文前部出现、带参数、既不是字段也不是排版命令、也不是文档自己定义的宏
    defined = set(re.findall(
        r"\\(?:newcommand|renewcommand|providecommand|DeclareMathOperator|DeclareRobustCommand|"
        r"NewDocumentCommand|DeclareDocumentCommand)\*?\s*\{?\s*\\([A-Za-z]+)", text))
    defined |= set(re.findall(r"\\(?:def|let|gdef|edef)\s*\\([A-Za-z]+)", text))
    # 数学内容（$…$、\(…\)、\[…\]）里的命令不是封面字段（fdu 符号表里的 \bra \ket）
    scan = re.sub(r"\$\$.*?\$\$|\$[^$]*\$|\\\(.*?\\\)|\\\[.*?\\\]", " ", front_body, flags=re.S)
    unknown = []
    for cm in re.finditer(r"\\([A-Za-z]+)\*?\s*(?=[\[{])", scan):
        name = cm.group(1)
        if name in known_cmds or name in _LAYOUT_CMDS or name in defined or name in unknown:
            continue
        unknown.append(name)
    rep.unresolved = unknown
    expected = set(cmap.keys.values()) | set(cmap.envs.values()) | set(cmap.labels.values()) if cmap else \
        {"title", "author", "abstract"}
    rep.missing = [FIELD_LABELS.get(f, f) for f in sorted(expected - set(rep.fields))]
    return rep
