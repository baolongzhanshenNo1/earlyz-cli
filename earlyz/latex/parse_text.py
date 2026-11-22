"""L2 解析用的常量表与行内缓冲：文本符号、重音、样式命令、无内容的版式命令。

版式命令（字号/字体/间距/页眉标记）只跳过参数、不进 IR：场景 A 的版式由学校模板决定，
场景 B 原文照搬。表里每一项都是「跳过几个可选参数 + 几个必选参数」，不猜内容。
"""
from __future__ import annotations

import re
import unicodedata

from .ir import Inline

TEXT_SYMBOLS = {
    "%": "%", "&": "&", "#": "#", "$": "$", "_": "_", "{": "{", "}": "}", " ": " ", "\\ ": " ",
    "ldots": "…", "dots": "…", "textellipsis": "…", "LaTeX": "LaTeX", "TeX": "TeX",
    "LaTeXe": "LaTeX2ε", "XeTeX": "XeTeX", "XeLaTeX": "XeLaTeX", "BibTeX": "BibTeX",
    "pdfLaTeX": "pdfLaTeX", "LuaLaTeX": "LuaLaTeX", "textbackslash": "\\", "textasciitilde": "~",
    "textasciicircum": "^", "textbullet": "•", "S": "§", "P": "¶", "copyright": "©",
    "textregistered": "®", "texttrademark": "™", "textdegree": "°", "textendash": "–",
    "textemdash": "—", "textquoteleft": "‘", "textquoteright": "’", "textquotedblleft": "“",
    "textquotedblright": "”", "ss": "ß", "o": "ø", "O": "Ø", "ae": "æ", "AE": "Æ", "oe": "œ",
    "OE": "Œ", "aa": "å", "AA": "Å", "l": "ł", "L": "Ł", "i": "ı", "j": "ȷ", "pounds": "£",
    "euro": "€", "textyen": "¥", "dag": "†", "ddag": "‡", "textbar": "|", "textless": "<",
    "textgreater": ">", "slash": "/", "textunderscore": "_", "quad": "\u2003",
    "qquad": "\u2003\u2003", "enskip": "\u2002", "enspace": "\u2002", ",": "\u2009",
    "thinspace": "\u2009", ";": "\u2005", ":": "\u2005", ">": "\u2005", "!": "", "/": "",
    "-": "", "@": "", "nobreakspace": "\u00a0", "textperiodcentered": "·", "checkmark": "✓",
    "textcelsius": "℃", "degree": "°", "textpm": "±", "texttimes": "×", "textdiv": "÷",
    "textmu": "µ", "textonehalf": "½", "textperthousand": "‰", "textsection": "§",
    "backslash": "\\",
}
ACCENTS = {"'": "\u0301", "`": "\u0300", "^": "\u0302", '"': "\u0308", "~": "\u0303",
           "=": "\u0304", ".": "\u0307", "u": "\u0306", "v": "\u030c", "H": "\u030b",
           "c": "\u0327", "k": "\u0328", "r": "\u030a", "b": "\u0331", "d": "\u0323"}

STYLE_CMDS = {  # 命令 → 样式（None = 透明，只保留内容）
    "textbf": "b", "textit": "i", "emph": "i", "textsl": "i", "texttt": "tt", "textsc": "sc",
    "underline": "u", "uline": "u", "textsuperscript": "sup", "textsubscript": "sub",
    "textrm": None, "textsf": None, "textup": None, "textmd": None, "textnormal": None,
    "mbox": None, "text": None, "fbox": None, "textsuperscript*": "sup", "hl": None,
    "tnote": "sup",                      # threeparttable 表注标记
    "hologo": None, "strong": "b",       # fontspec \\strong = 加粗强调                      # \\hologo{XeTeX}：logo 名原样当文字
}
STYLE_OPT_CMDS = {"makebox": 2, "framebox": 2, "parbox": 1, "colorbox": 0, "raisebox": 2}
STYLE_DECLS = {"bfseries": "b", "bf": "b", "itshape": "i", "it": "i", "em": "i", "slshape": "i",
               "sl": "i", "ttfamily": "tt", "tt": "tt", "scshape": "sc", "sc": "sc"}
STYLE_RESETS = {"normalfont": None, "rm": None, "rmfamily": None, "sffamily": None, "sf": None,
                "upshape": "i", "mdseries": "b"}

IGNORE_ARGS = {  # 版式命令：(可选参数个数, 必选参数个数)；* 形式先吃掉星号
    "vspace": (0, 1), "hspace": (0, 1), "fontsize": (0, 2), "setlength": (0, 2),
    "addtolength": (0, 2), "setcounter": (0, 2), "addtocounter": (0, 2), "stepcounter": (0, 1),
    "refstepcounter": (0, 1), "pagestyle": (0, 1), "thispagestyle": (0, 1),
    "pagenumbering": (0, 1), "linespread": (0, 1), "color": (1, 1), "CJKfamily": (0, 1),
    "zihao": (0, 1), "fontseries": (0, 1), "fontshape": (0, 1), "fontfamily": (0, 1),
    "usefont": (0, 4), "captionsetup": (1, 1), "graphicspath": (0, 1), "hypersetup": (0, 1),
    "markboth": (0, 2), "markright": (0, 1), "addcontentsline": (0, 3), "enlargethispage": (0, 1),
    "setcitestyle": (0, 1), "rule": (1, 2), "urlstyle": (0, 1), "DeclareGraphicsExtensions": (0, 1),
    "pagecolor": (1, 1), "titleformat": (1, 5), "titlespacing": (0, 4), "ctexset": (0, 1),
    "setmainfont": (1, 1), "setCJKmainfont": (1, 1), "newcounter": (1, 1), "counterwithin": (0, 2),
    "counterwithout": (0, 2), "numberwithin": (0, 2), "addvspace": (0, 1), "nocite": (0, 1),
    "bibliographystyle": (0, 1), "addbibresource": (1, 1), "usepackage": (1, 1),
    "RequirePackage": (1, 1), "selectlanguage": (0, 1), "AtBeginDocument": (0, 1),
    "AtEndDocument": (0, 1), "tnotemark": (1, 0), "fnmark": (1, 0), "cormark": (1, 0),
    "CJKtilde": (0, 0), "vphantom": (0, 1), "hphantom": (0, 1), "phantom": (0, 1),
    "includeonly": (0, 1), "bibpunct": (1, 6), "newcolumntype": (1, 2), "allowdisplaybreaks": (1, 0),
    "tikzset": (0, 1), "usetikzlibrary": (0, 1), "pgfplotsset": (0, 1), "SetAlgoLined": (0, 0),
    "bio": (1, 1),                       # Elsevier cas：\\bio[选项]{照片} 作者简介开始（正文照留，\\endbio 结束）
}
NOOP = frozenset({
    "relax", "protect", "indent", "small", "footnotesize", "scriptsize", "tiny", "large", "Large",
    "LARGE", "huge", "Huge", "normalsize", "selectfont", "hfill", "vfill", "hfil", "vfil", "null",
    "leavevmode", "nopagebreak", "pagebreak", "nolinebreak", "smallskip", "medskip", "bigskip",
    "break", "allowbreak", "sloppy", "fussy", "onecolumn", "twocolumn", "strut", "unskip",
    "ignorespaces", "phantomsection", "boldmath", "unboldmath", "frenchspacing", "nonfrenchspacing",
    "raggedbottom", "flushbottom", "par@", "centering@", "endinput", "makeatletter", "makeatother",
    "displaystyle", "textstyle", "normalcolor", "ttfamily@", "sffamily@", "noalign", "hss", "vss",
    "zihaowu@", "maketitle@", "balance", "FloatBarrier", "noindent@", "printcredits",
    "printnomenclature", "WriteBookmarks", "today@", "arraybackslash", "xspace", "insertTableNotes",
    "hrule", "vrule", "nobreak", "newline@", "endbio",
    # 中文字体切换是声明（不带参数）：ctex 的 \\heiti、xaut/CjC cls 的 \\songti \\hei \\song …
    "songti", "heiti", "kaishu", "fangsong", "lishu", "youyuan", "hei", "song", "kai", "fs", "li",
    "CJKfamily@", "arialfont", "timesfont",
})
# 寄存器赋值（\\tabcolsep=2pt、\\hyphenpenalty 10000）：吃掉可选的 = 与数值/单位，不留进正文
DIMEN_CMDS = frozenset({"vskip", "hskip", "kern", "mkern", "hglue", "vglue", "mskip", "lineskip",
                        "baselineskip", "parskip", "parindent", "mathindent", "clubpenalty",
                        "widowpenalty", "hyphenpenalty", "tolerance", "emergencystretch", "tabcolsep",
                        "arrayrulewidth", "columnsep", "doublerulesep", "LTpre", "LTpost",
                        "abovedisplayskip", "belowdisplayskip", "abovecaptionskip", "belowcaptionskip",
                        "jot", "textfloatsep", "intextsep", "floatsep", "exhyphenpenalty"})
DUMMY_TEXT = frozenset({"lipsum", "zhlipsum", "kant", "blindtext", "Blindtext", "lorem"})
# 透明环境：跳过若干必选参数后，内容按普通内容解析（行内、块级共用）
TRANSPARENT_ENVS = {"CJK": 2, "CJK*": 2, "multicols": 1, "multicols*": 1, "threeparttable": 0,
                    "ThreePartTable": 0, "landscape": 0, "small": 0, "footnotesize": 0, "spacing": 1,
                    "singlespace": 0, "onehalfspace": 0, "adjustwidth": 2, "document": 0, "frame": 0,
                    "linenumbers": 0, "sloppypar": 0, "strip": 0, "graphicalabstract": 0}
CITE_CMDS = frozenset({
    "cite", "citep", "citet", "citealt", "citealp", "citeauthor", "citeyear", "citeyearpar",
    "parencite", "textcite", "autocite", "footcite", "supercite", "upcite", "citenum", "Cite",
    "Citep", "Citet", "Parencite", "Textcite", "Autocite", "cites", "parencites", "smartcite",
})
REF_CMDS = frozenset({"ref", "eqref", "autoref", "cref", "Cref", "pageref", "nameref", "vref",
                      "cpageref", "Autoref", "subref", "labelcref"})
COUNTER_FMT = frozenset({"arabic", "roman", "Roman", "alph", "Alph", "chinese", "zhnum", "fnsymbol",
                         "CJKnumber", "zhdig"})
SECTION_LEVEL = {"part": 0, "chapter": 1, "section": 2, "subsection": 3, "subsubsection": 4,
                 "paragraph": 5, "subparagraph": 6}

_CJK = r"\u2e80-\u9fff\uf900-\ufaff\u3000-\u303f\uff00-\uffef"
_CJK_SPACE = re.compile(r"(?<=[" + _CJK + r"]) (?=[" + _CJK + r"])")
_LIG = (("---", "—"), ("--", "–"), ("``", "“"), ("''", "”"), ("`", "‘"), ("'", "’"))


def is_cjk_text(s: str) -> bool:
    cjk = len(re.findall("[\u4e00-\u9fff]", s))
    letters = len(re.findall(r"[A-Za-z]", s))
    return cjk > 0 and cjk * 3 >= letters


def accent(ch: str, mark: str) -> str:
    return unicodedata.normalize("NFC", (ch or " ")[0] + mark + (ch or " ")[1:])


class InlineBuf:
    def __init__(self) -> None:
        self.items: list[Inline] = []

    def text(self, s: str, style: tuple) -> None:
        if not s:
            return
        if s == " " and (not self.items or (self.items[-1].t == "text" and self.items[-1].v.endswith(" "))):
            return
        last = self.items[-1] if self.items else None
        if last is not None and last.t == "text" and last.style == style:
            last.v += s
        else:
            self.items.append(Inline("text", s, style))

    def add(self, inl: Inline) -> None:
        self.items.append(inl)

    def has_content(self) -> bool:
        return any(i.t != "text" or i.v.strip() for i in self.items)

    def take(self) -> list[Inline]:
        items, self.items = self.items, []
        return finalize(items)


def finalize(items: list[Inline]) -> list[Inline]:
    """TeX 连字（-- --- `` ''）、CJK 字间换行不产生空格（xeCJK 行为）、段首尾空白去掉。"""
    out = []
    for it in items:
        if it.t == "text":
            v = it.v
            if "tt" not in it.style:
                for a, b in _LIG:
                    v = v.replace(a, b)
            v = _CJK_SPACE.sub("", v)
            if not v:
                continue
            it.v = v
        out.append(it)
    while out and out[0].t == "text" and not out[0].v.strip():
        out.pop(0)
    while out and out[-1].t == "text" and not out[-1].v.strip():
        out.pop()
    if out and out[0].t == "text":
        out[0].v = out[0].v.lstrip(" ")
    if out and out[-1].t == "text":
        out[-1].v = out[-1].v.rstrip(" ")
    return out


def inline_text(items) -> str:
    """行内 → 纯文本（标题比对、语言判断、报告用）。"""
    parts = []
    for it in items or ():
        if it.t == "text":
            parts.append(it.v)
        elif it.t == "math":
            parts.append(it.v)
        elif it.t == "footnote":
            continue
        elif it.t in ("link",):
            parts.append(it.data.get("text") or it.v)
        elif it.t == "code":
            parts.append(it.v)
    return "".join(parts).strip()
