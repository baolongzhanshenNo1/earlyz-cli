"""MTEF v5 读写（MathType 4.0+ 公式二进制格式；Design Science 公开规格「MathType MTEF v.5」）。

模型是「记录树」：结构记录（LINE/CHAR/TMPL/PILE/MATRIX）带对象列表，其余记录（SIZE/FONT_STYLE_DEF/
COLOR/定义记录…）按出现位置原样留在列表里——读进来再写出去逐字节一致（真论文 8 个对象是回归 oracle）。
"""
from __future__ import annotations

from dataclasses import dataclass, field

# 记录类型
END, LINE, CHAR, TMPL, PILE, MATRIX, EMBELL, RULER = 0, 1, 2, 3, 4, 5, 6, 7
FONT_STYLE_DEF, SIZE, FULL, SUB, SUB2, SYM, SUBSYM = 8, 9, 10, 11, 12, 13, 14
COLOR, COLOR_DEF, FONT_DEF, EQN_PREFS, ENCODING_DEF = 15, 16, 17, 18, 19

# 选项位
OPT_NUDGE = 0x08
OPT_CHAR_EMBELL, OPT_CHAR_FUNC_START, OPT_CHAR_ENC_CHAR_8 = 0x01, 0x02, 0x04
OPT_CHAR_ENC_CHAR_16, OPT_CHAR_ENC_NO_MTCODE = 0x10, 0x20
OPT_LINE_NULL, OPT_LP_RULER, OPT_LINE_LSPACE = 0x01, 0x02, 0x04
COLOR_CMYK, COLOR_NAME = 0x01, 0x04

# 字样（typeface）
FN_TEXT, FN_FUNCTION, FN_VARIABLE, FN_LCGREEK, FN_UCGREEK, FN_SYMBOL = 1, 2, 3, 4, 5, 6
FN_VECTOR, FN_NUMBER, FN_USER1, FN_USER2, FN_MTEXTRA, FN_TEXT_FE = 7, 8, 9, 10, 11, 12
FN_EXPAND, FN_MARKER, FN_SPACE = 22, 23, 24

# 模板选择子（Template selectors）
TM_ANGLE, TM_PAREN, TM_BRACE, TM_BRACK, TM_BAR, TM_DBAR, TM_FLOOR, TM_CEILING = range(8)
TM_OBRACK, TM_INTERVAL, TM_ROOT, TM_FRACT, TM_UBAR, TM_OBAR, TM_ARROW, TM_INTEG = range(8, 16)
TM_SUM, TM_PROD, TM_COPROD, TM_UNION, TM_INTER, TM_INTOP, TM_SUMOP, TM_LIM = range(16, 24)
TM_HBRACE, TM_HBRACK, TM_LDIV, TM_SUB, TM_SUP, TM_SUBSUP, TM_DIRAC, TM_VEC = range(24, 32)
TM_TILDE, TM_HAT, TM_ARC, TM_JSTATUS, TM_STRIKE, TM_BOX = range(32, 38)


class MtefError(ValueError):
    pass


@dataclass
class Rec:
    """非结构记录（SIZE / FULL… / FONT_STYLE_DEF / COLOR / 定义记录 / RULER / FUTURE）：原字节保留。"""
    type: int
    raw: bytes                    # 不含类型字节
    info: dict = field(default_factory=dict)


@dataclass
class Embell:
    kind: int
    options: int = 0
    nudge: bytes = b""


@dataclass
class Char:
    typeface: int                 # 正 = 字样（FN_*），负 = 显式字体（FONT_DEF 下标取负）
    mtcode: int | None = None
    fontpos: int | None = None
    options: int = 0
    nudge: bytes = b""
    embells: list = field(default_factory=list)


@dataclass
class Line:
    items: list = field(default_factory=list)
    options: int = 0
    nudge: bytes = b""
    lspace: int | None = None
    ruler: Rec | None = None

    @property
    def null(self) -> bool:
        return bool(self.options & OPT_LINE_NULL)


@dataclass
class Tmpl:
    selector: int
    variation: int
    tmpl_options: int = 0
    items: list = field(default_factory=list)   # 子对象（LINE / CHAR / 尺寸记录…）
    options: int = 0
    nudge: bytes = b""


@dataclass
class Pile:
    halign: int
    valign: int
    lines: list = field(default_factory=list)
    options: int = 0
    nudge: bytes = b""
    ruler: Rec | None = None


@dataclass
class Matrix:
    valign: int
    h_just: int
    v_just: int
    rows: int
    cols: int
    row_parts: bytes
    col_parts: bytes
    cells: list = field(default_factory=list)
    options: int = 0
    nudge: bytes = b""


@dataclass
class Equation:
    version: int = 5
    platform: int = 1
    product: int = 0
    pversion: int = 7
    psubversion: int = 4
    app_key: str = "DSMT7"
    inline: bool = False
    eqn_options: int = 0
    records: list = field(default_factory=list)  # 顶层：定义记录 + 初始尺寸 + LINE/PILE
    tail: bytes = b""                            # 结尾 END 之后的字节（MathType 不写；保留以便逐字节往返）

    def body(self):
        """顶层结构对象（LINE / PILE）。"""
        for r in self.records:
            if isinstance(r, (Line, Pile)):
                return r
        return None


# ── 读 ────────────────────────────────────────────────────────────────────────

class _Reader:
    def __init__(self, b: bytes, pos: int = 0):
        self.b, self.p = b, pos

    def u8(self) -> int:
        if self.p >= len(self.b):
            raise MtefError("unexpected end of MTEF")
        v = self.b[self.p]
        self.p += 1
        return v

    def u16(self) -> int:
        v = self.u8()
        return v | (self.u8() << 8)

    def uint(self) -> int:
        v = self.u8()
        return self.u16() if v == 255 else v

    def sint(self) -> int:
        v = self.u8()
        return self.u16() - 32768 if v == 255 else v - 128

    def cstr(self) -> str:
        e = self.b.index(b"\0", self.p)
        s = self.b[self.p:e].decode("latin-1")
        self.p = e + 1
        return s

    def take(self, start: int) -> bytes:
        return self.b[start:self.p]

    def nudge(self, opts: int) -> bytes:
        if not opts & OPT_NUDGE:
            return b""
        s = self.p
        dx, dy = self.u8(), self.u8()
        if dx == 128 and dy == 128:
            self.p += 4
        return self.take(s)

    def nibbles_dims(self) -> None:
        """维度数组：计数字节 + nibble 流（高 4 位在前）。每个维度 = 单位 nibble（0..4）+ 值 nibble… + 0xF；
        奇数个 nibble 时补一个 0 凑整字节。"""
        n = self.u8()
        expect_unit = True
        while n:
            byte = self.u8()
            for v in (byte >> 4, byte & 0xF):
                if n == 0:
                    break                     # 末尾补齐 nibble
                if expect_unit:
                    expect_unit = False
                elif v == 0xF:
                    n -= 1
                    expect_unit = True


def _ruler(r: _Reader) -> Rec:
    s = r.p
    n = r.u8()
    r.p += 3 * n
    return Rec(RULER, r.take(s))


def _object_list(r: _Reader) -> list:
    out = []
    while True:
        t = r.u8()
        if t == END:
            return out
        out.append(_record(r, t))


def _record(r: _Reader, t: int):
    if t == LINE:
        o = r.u8()
        ln = Line(options=o, nudge=r.nudge(o))
        if o & OPT_LINE_LSPACE:
            ln.lspace = r.u16()
        if o & OPT_LP_RULER:
            if r.u8() != RULER:
                raise MtefError("LINE ruler expected")
            ln.ruler = _ruler(r)
        if not o & OPT_LINE_NULL:
            ln.items = _object_list(r)
        return ln
    if t == CHAR:
        o = r.u8()
        c = Char(typeface=0, options=o, nudge=r.nudge(o))
        c.typeface = r.sint()
        if not o & OPT_CHAR_ENC_NO_MTCODE:
            c.mtcode = r.u16()
        if o & OPT_CHAR_ENC_CHAR_8:
            c.fontpos = r.u8()
        elif o & OPT_CHAR_ENC_CHAR_16:
            c.fontpos = r.u16()
        if o & OPT_CHAR_EMBELL:
            for e in _object_list(r):
                if not isinstance(e, Embell):
                    raise MtefError("embellishment list holds non-EMBELL record")
                c.embells.append(e)
        return c
    if t == TMPL:
        o = r.u8()
        tm = Tmpl(selector=0, variation=0, options=o, nudge=r.nudge(o))
        tm.selector = r.u8()
        v = r.u8()
        if v & 0x80:
            v = (v & 0x7F) | (r.u8() << 8)
        tm.variation = v
        tm.tmpl_options = r.u8()
        tm.items = _object_list(r)
        return tm
    if t == PILE:
        o = r.u8()
        pl = Pile(halign=0, valign=0, options=o, nudge=r.nudge(o))
        pl.halign, pl.valign = r.u8(), r.u8()
        if o & OPT_LP_RULER:
            if r.u8() != RULER:
                raise MtefError("PILE ruler expected")
            pl.ruler = _ruler(r)
        pl.lines = _object_list(r)
        return pl
    if t == MATRIX:
        o = r.u8()
        m = Matrix(0, 0, 0, 0, 0, b"", b"", options=o, nudge=r.nudge(o))
        m.valign, m.h_just, m.v_just, m.rows, m.cols = (r.u8() for _ in range(5))
        nr, nc = (2 * (m.rows + 1) + 7) // 8, (2 * (m.cols + 1) + 7) // 8
        m.row_parts = bytes(r.u8() for _ in range(nr))
        m.col_parts = bytes(r.u8() for _ in range(nc))
        m.cells = _object_list(r)
        return m
    if t == EMBELL:
        o = r.u8()
        nd = r.nudge(o)
        return Embell(kind=r.u8(), options=o, nudge=nd)
    s = r.p
    if t == RULER:
        return _ruler(r)
    if t == FONT_STYLE_DEF:
        idx = r.uint()
        r.u8()
        return Rec(t, r.take(s), {"font": idx})
    if t == SIZE:
        k = r.u8()
        if k == 101:
            r.u16()
        elif k == 100:
            r.u8()
            r.u16()
        else:
            r.u8()
        return Rec(t, r.take(s))
    if FULL <= t <= SUBSYM:
        return Rec(t, b"")
    if t == COLOR:
        r.uint()
        return Rec(t, r.take(s))
    if t == COLOR_DEF:
        o = r.u8()
        r.p += 2 * (4 if o & COLOR_CMYK else 3)
        if o & COLOR_NAME:
            r.cstr()
        return Rec(t, r.take(s))
    if t == FONT_DEF:
        enc = r.uint()
        name = r.cstr()
        return Rec(t, r.take(s), {"enc": enc, "name": name})
    if t == EQN_PREFS:
        r.u8()
        r.nibbles_dims()
        r.nibbles_dims()
        n = r.u8()
        for _ in range(n):
            if r.uint():
                r.u8()
        return Rec(t, r.take(s))
    if t == ENCODING_DEF:
        name = r.cstr()
        return Rec(t, r.take(s), {"name": name})
    if t >= 100:
        n = r.uint()
        r.p += n
        return Rec(t, r.take(s))
    raise MtefError(f"unknown MTEF record type {t} at {r.p - 1}")


def parse(b: bytes) -> Equation:
    """MTEF 字节（Equation Native 去掉 28 字节头之后的部分）→ Equation。"""
    if len(b) < 6 or b[0] != 5:
        raise MtefError(f"not MTEF v5 (version byte {b[0] if b else None})")
    r = _Reader(b)
    eq = Equation(version=r.u8(), platform=r.u8(), product=r.u8(), pversion=r.u8(),
                  psubversion=r.u8())
    eq.app_key = r.cstr()
    eq.eqn_options = r.u8()
    eq.inline = bool(eq.eqn_options & 1)
    while r.p < len(b):
        t = r.u8()
        if t == END:
            break
        eq.records.append(_record(r, t))
    eq.tail = b[r.p:]
    return eq


def read_equation(b: bytes) -> Equation:
    """MTEF 二进制 → Equation（parse 的公开接口别名）。"""
    return parse(b)


# ── 写 ────────────────────────────────────────────────────────────────────────

def _uint(v: int) -> bytes:
    return bytes([v]) if v < 255 else bytes([255, v & 0xFF, v >> 8 & 0xFF])


def _sint(v: int) -> bytes:
    if -128 <= v < 127:
        return bytes([v + 128])
    v += 32768
    return bytes([255, v & 0xFF, v >> 8 & 0xFF])


def _u16(v: int) -> bytes:
    return bytes([v & 0xFF, v >> 8 & 0xFF])


def _emit_list(items, out: bytearray) -> None:
    for it in items:
        _emit(it, out)
    out.append(END)


def _emit(x, out: bytearray) -> None:
    if isinstance(x, Line):
        o = x.options
        out += bytes([LINE, o]) + x.nudge
        if o & OPT_LINE_LSPACE:
            out += _u16(x.lspace or 0)
        if o & OPT_LP_RULER and x.ruler is not None:
            out += bytes([RULER]) + x.ruler.raw
        if not o & OPT_LINE_NULL:
            _emit_list(x.items, out)
    elif isinstance(x, Char):
        o = x.options
        out += bytes([CHAR, o]) + x.nudge + _sint(x.typeface)
        if not o & OPT_CHAR_ENC_NO_MTCODE:
            out += _u16(x.mtcode or 0)
        if o & OPT_CHAR_ENC_CHAR_8:
            out.append(x.fontpos & 0xFF)
        elif o & OPT_CHAR_ENC_CHAR_16:
            out += _u16(x.fontpos or 0)
        if o & OPT_CHAR_EMBELL:
            _emit_list(x.embells, out)
    elif isinstance(x, Tmpl):
        out += bytes([TMPL, x.options]) + x.nudge + bytes([x.selector])
        v = x.variation
        out += bytes([v]) if v < 0x80 else bytes([(v & 0x7F) | 0x80, v >> 8])
        out.append(x.tmpl_options)
        _emit_list(x.items, out)
    elif isinstance(x, Pile):
        out += bytes([PILE, x.options]) + x.nudge + bytes([x.halign, x.valign])
        if x.options & OPT_LP_RULER and x.ruler is not None:
            out += bytes([RULER]) + x.ruler.raw
        _emit_list(x.lines, out)
    elif isinstance(x, Matrix):
        out += bytes([MATRIX, x.options]) + x.nudge
        out += bytes([x.valign, x.h_just, x.v_just, x.rows, x.cols]) + x.row_parts + x.col_parts
        _emit_list(x.cells, out)
    elif isinstance(x, Embell):
        out += bytes([EMBELL, x.options]) + x.nudge + bytes([x.kind])
    elif isinstance(x, Rec):
        out.append(x.type)
        out += x.raw
    else:
        raise MtefError(f"cannot serialize {type(x).__name__}")


def serialize(eq: Equation) -> bytes:
    out = bytearray([eq.version, eq.platform, eq.product, eq.pversion, eq.psubversion])
    out += eq.app_key.encode("latin-1") + b"\0"
    out.append((eq.eqn_options & ~1) | (1 if eq.inline else 0))
    for r in eq.records:
        _emit(r, out)
    out.append(END)
    out += eq.tail
    return bytes(out)


def write_equation(eq: Equation) -> bytes:
    """Equation → MTEF 二进制（serialize 的公开接口别名）。"""
    return serialize(eq)


# ── 构造辅助（LaTeX → MTEF、Format Equations 用）──────────────────────────────

def size_rec(kind: int) -> Rec:
    """FULL / SUB / SUB2 / SYM / SUBSYM 简写尺寸记录。"""
    return Rec(kind, b"")


def encoding_def(name: str) -> Rec:
    return Rec(ENCODING_DEF, name.encode("latin-1") + b"\0", {"name": name})


def font_def(enc_index: int, name: str) -> Rec:
    return Rec(FONT_DEF, _uint(enc_index) + name.encode("latin-1") + b"\0", {"enc": enc_index, "name": name})


def walk(items):
    """深度优先遍历结构对象。"""
    for it in items:
        yield it
        if isinstance(it, Line):
            yield from walk(it.items)
        elif isinstance(it, Tmpl):
            yield from walk(it.items)
        elif isinstance(it, Pile):
            yield from walk(it.lines)
        elif isinstance(it, Matrix):
            yield from walk(it.cells)
