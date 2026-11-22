"""LaTeX 数学 → MTEF（场景 A：学位论文公式，E6 优先级）。

支持范围：
- 基础运算符：+, -, *, /, =, <, >, ≤, ≥, ≠, ≈, ±, ∓, ×, ÷, ·
- 上下标：^, _
- 分数：\frac{}{}, \dfrac{}{}, \tfrac{}{}
- 根号：\sqrt{}, \sqrt[n]{}
- 求和积分：\sum, \int, \prod, \lim, \oint
- 括号：(), [], {}, \left \right, \big \Big \bigg \Bigg
- 希腊字母：\alpha, \beta, \gamma, \delta, \theta, \lambda, \sigma, \omega, \Gamma, \Delta, \Theta, \Lambda, \Sigma, \Omega 等
- 重音符号：\hat{}, \bar{}, \tilde{}, \vec{}, \dot{}, \ddot{}
- 字体：\mathbf{}, \mathrm{}, \mathit{}, \mathcal{}, \mathbb{}, \text{}
- 矩阵：\begin{matrix} ... \end{matrix}, pmatrix, bmatrix, vmatrix, cases
- 定界符：\langle, \rangle, \lfloor, \rfloor, \lceil, \rceil
- 特殊符号：\infty, \partial, \nabla, \prime, \cdots, \ldots, \dots

不支持（需要时返回错误）：
- 高级宏包（amsmath 的 align/gather/multline 等多行环境 → 需要在 L2 层拆成多个公式）
- 自定义命令（\newcommand, \def）
- TikZ、chemfig 等绘图
- \boxed, \fbox 等装饰框
- 复杂的数组对齐（array 环境带列对齐参数 → 简化为居中）

设计：
- 递归下降解析器，直接构造 MTEF 记录树
- 优先级：^ _ > \frac > 其他命令 > 空格分隔的项
- 括号自动匹配（\left \right）
- 数学空格（\,, \;, \quad）映射到 MTEF nudge
"""
from __future__ import annotations

from .mtef import (
    Equation, Line, Char, Tmpl, Pile, Rec, Embell,
    CHAR, SIZE, FULL, FN_TEXT, FN_VARIABLE, FN_NUMBER, FN_FUNCTION, FN_SYMBOL,
    FN_LCGREEK, FN_UCGREEK, FN_MTEXTRA,
    TM_FRACT, TM_ROOT, TM_ANGLE, TM_PAREN, TM_BRACK, TM_BRACE, TM_BAR,
    TM_SUB, TM_SUP, TM_SUBSUP, TM_SUM, TM_INTEG, TM_PROD, TM_LIM,
    TM_HAT, TM_TILDE, TM_VEC, TM_BAR as TM_OVERBAR, TM_UBAR,
    OPT_CHAR_ENC_CHAR_8, OPT_CHAR_EMBELL,
)


class LatexParseError(ValueError):
    pass


# ── 词法 ────────────────────────────────────────────────────────────────────

class _Tok:
    def __init__(self, kind: str, val: str = "", pos: int = 0):
        self.kind, self.val, self.pos = kind, val, pos
    def __repr__(self):
        return f"{self.kind}({self.val!r})" if self.val else self.kind


def _tokenize(src: str) -> list[_Tok]:
    """LaTeX → token 流。"""
    out, i = [], 0
    while i < len(src):
        c = src[i]
        if c.isspace():
            # 数学模式下空格不重要，除非显式 \, \; \quad
            i += 1
            continue
        if c == '\\':
            j = i + 1
            if j < len(src) and not src[j].isalpha():
                # \{, \}, \,, \;, \!, \quad 等单字符命令
                out.append(_Tok('CMD', src[j], i))
                i = j + 1
            else:
                # \alpha, \frac, \begin 等
                while j < len(src) and src[j].isalpha():
                    j += 1
                out.append(_Tok('CMD', src[i+1:j], i))
                i = j
        elif c in '{}':
            out.append(_Tok('BRACE', c, i))
            i += 1
        elif c in '^_':
            out.append(_Tok('SCRIPT', c, i))
            i += 1
        elif c in '+-*/=<>()[]|':
            out.append(_Tok('OP', c, i))
            i += 1
        elif c == '&':
            out.append(_Tok('ALIGN', '&', i))
            i += 1
        elif c.isdigit() or c == '.':
            j = i
            while j < len(src) and (src[j].isdigit() or src[j] == '.'):
                j += 1
            out.append(_Tok('NUM', src[i:j], i))
            i = j
        elif c.isalpha():
            j = i
            while j < len(src) and src[j].isalpha():
                j += 1
            out.append(_Tok('VAR', src[i:j], i))
            i = j
        else:
            # 其他 Unicode 字符（如直接输入的希腊字母）
            out.append(_Tok('CHAR', c, i))
            i += 1
    return out


# ── 解析 ────────────────────────────────────────────────────────────────────

class _Parser:
    def __init__(self, toks: list[_Tok]):
        self.toks, self.i = toks, 0

    def peek(self) -> _Tok | None:
        return self.toks[self.i] if self.i < len(self.toks) else None

    def eat(self, kind: str | None = None, val: str | None = None) -> _Tok:
        t = self.peek()
        if t is None:
            raise LatexParseError(f"unexpected end, expected {kind or val}")
        if kind and t.kind != kind:
            raise LatexParseError(f"expected {kind}, got {t.kind} at pos {t.pos}")
        if val and t.val != val:
            raise LatexParseError(f"expected {val!r}, got {t.val!r} at pos {t.pos}")
        self.i += 1
        return t

    def parse(self) -> Equation:
        """顶层入口：解析整个公式。"""
        records = [
            Rec(SIZE, b'\x00\x02'),  # SIZE FULL = 默认尺寸
        ]
        line = self.parse_line()
        records.append(line)
        return Equation(records=records)

    def parse_line(self) -> Line:
        """一行（顶层或子公式）：优先级最低，空格分隔的多项。"""
        items = []
        while True:
            t = self.peek()
            if t is None or t.kind == 'BRACE' and t.val == '}':
                break
            if t.kind == 'ALIGN':  # 矩阵/对齐环境的列分隔符
                break
            expr = self.parse_expr()
            # 展平：_text_word 返回 Line，需要拆开
            if isinstance(expr, Line) and all(isinstance(x, Char) for x in expr.items):
                items.extend(expr.items)
            else:
                items.append(expr)
        return Line(items=items)

    def parse_expr(self):
        """表达式：处理 ^ _ 上下标（优先级最高）。"""
        base = self.parse_term()
        # 检查上下标
        sup_line, sub_line = None, None
        while True:
            t = self.peek()
            if t and t.kind == 'SCRIPT':
                if t.val == '^':
                    self.eat()
                    sup_line = self.parse_script_arg()
                elif t.val == '_':
                    self.eat()
                    sub_line = self.parse_script_arg()
            else:
                break
        # 构造模板
        if sup_line and sub_line:
            return Tmpl(TM_SUBSUP, 0, items=[base, sub_line, sup_line])
        elif sup_line:
            return Tmpl(TM_SUP, 0, items=[base, sup_line])
        elif sub_line:
            return Tmpl(TM_SUB, 0, items=[base, sub_line])
        else:
            return base

    def parse_script_arg(self):
        """上下标参数：{ } 或单个 token。"""
        t = self.peek()
        if t and t.kind == 'BRACE' and t.val == '{':
            self.eat()
            line = self.parse_line()
            self.eat('BRACE', '}')
            return line
        else:
            # 单字符上标，如 x^2
            item = self.parse_atom()
            return Line(items=[item])

    def parse_term(self):
        """项：\frac, \sqrt, 命令, 原子。"""
        t = self.peek()
        if t is None:
            raise LatexParseError("unexpected end in term")
        
        if t.kind == 'CMD':
            cmd = t.val
            self.eat()
            
            # 分数
            if cmd in ('frac', 'dfrac', 'tfrac'):
                num = self.parse_arg()
                den = self.parse_arg()
                return Tmpl(TM_FRACT, 0, items=[num, den])
            
            # 根号
            elif cmd == 'sqrt':
                # 可选参数 [n]
                idx_line = None
                if self.peek() and self.peek().kind == 'OP' and self.peek().val == '[':
                    self.eat()
                    idx_line = self.parse_line()
                    if self.peek() and self.peek().kind == 'OP' and self.peek().val == ']':
                        self.eat()
                radicand = self.parse_arg()
                if idx_line:
                    return Tmpl(TM_ROOT, 1, items=[radicand, idx_line])  # variation 1 = n 次根
                else:
                    return Tmpl(TM_ROOT, 0, items=[radicand])
            
            # 大运算符：\sum, \int, \prod, \lim
            elif cmd == 'sum':
                return Char(FN_SYMBOL, mtcode=0xE5)  # Σ
            elif cmd == 'int':
                return Char(FN_SYMBOL, mtcode=0xF2)  # ∫
            elif cmd == 'prod':
                return Char(FN_SYMBOL, mtcode=0xD5)  # Π
            elif cmd == 'lim':
                return self._text_word('lim', FN_FUNCTION)
            
            # 希腊字母（小写）
            elif cmd in _GREEK_LOWER:
                return Char(FN_LCGREEK, mtcode=_GREEK_LOWER[cmd])
            # 希腊字母（大写）
            elif cmd in _GREEK_UPPER:
                return Char(FN_UCGREEK, mtcode=_GREEK_UPPER[cmd])
            
            # 重音符号
            elif cmd == 'hat':
                arg = self.parse_arg()
                return Tmpl(TM_HAT, 0, items=[arg])
            elif cmd == 'bar':
                arg = self.parse_arg()
                return Tmpl(TM_OVERBAR, 0, items=[arg])
            elif cmd == 'tilde':
                arg = self.parse_arg()
                return Tmpl(TM_TILDE, 0, items=[arg])
            elif cmd == 'vec':
                arg = self.parse_arg()
                return Tmpl(TM_VEC, 0, items=[arg])
            
            # 括号：\left \right
            elif cmd == 'left':
                return self.parse_delim_group()
            
            # 字体
            elif cmd in ('mathbf', 'textbf'):
                arg = self.parse_arg()
                # 简化：暂时不改字体，返回原内容
                return arg
            elif cmd in ('mathrm', 'textrm', 'text'):
                arg = self.parse_arg()
                # text 内容用 FN_TEXT
                if isinstance(arg, Line):
                    for item in arg.items:
                        if isinstance(item, Char):
                            item.typeface = FN_TEXT
                return arg
            
            # 矩阵环境
            elif cmd == 'begin':
                env_name = self.parse_arg()
                env_str = self._line_to_text(env_name) if isinstance(env_name, Line) else ""
                if env_str in ('matrix', 'pmatrix', 'bmatrix', 'vmatrix', 'cases'):
                    return self.parse_matrix(env_str)
                raise LatexParseError(f"unsupported environment: {env_str}")
            
            # 特殊符号
            elif cmd == 'infty':
                return Char(FN_SYMBOL, mtcode=0xA5)  # ∞
            elif cmd == 'partial':
                return Char(FN_SYMBOL, mtcode=0xB6)  # ∂
            elif cmd == 'cdot':
                return Char(FN_SYMBOL, mtcode=0xD7)  # ·
            elif cmd == 'times':
                return Char(FN_SYMBOL, mtcode=0xB4)  # ×
            elif cmd == 'div':
                return Char(FN_SYMBOL, mtcode=0xB8)  # ÷
            elif cmd == 'pm':
                return Char(FN_SYMBOL, mtcode=0xB1)  # ±
            elif cmd == 'leq':
                return Char(FN_SYMBOL, mtcode=0xA3)  # ≤
            elif cmd == 'geq':
                return Char(FN_SYMBOL, mtcode=0xB3)  # ≥
            elif cmd == 'neq':
                return Char(FN_SYMBOL, mtcode=0xB9)  # ≠
            elif cmd == 'approx':
                return Char(FN_SYMBOL, mtcode=0xBB)  # ≈
            
            # 空格（忽略或映射到 nudge）
            elif cmd in (',', ';', '!', 'quad', 'qquad'):
                # 暂时忽略
                return Line(items=[])
            
            # 未知命令
            else:
                # 尝试作为函数名
                return self._text_word(cmd, FN_FUNCTION)
        
        else:
            return self.parse_atom()

    def parse_atom(self):
        """原子：数字、变量、括号、{ }。"""
        t = self.peek()
        if t is None:
            raise LatexParseError("unexpected end in atom")
        
        if t.kind == 'NUM':
            self.eat()
            return self._text_word(t.val, FN_NUMBER)
        
        elif t.kind == 'VAR':
            self.eat()
            # 单字母变量用 FN_VARIABLE，多字母视为函数名
            if len(t.val) == 1:
                return Char(FN_VARIABLE, fontpos=ord(t.val))
            else:
                return self._text_word(t.val, FN_FUNCTION)
        
        elif t.kind == 'OP':
            self.eat()
            # 运算符映射
            ch = t.val
            if ch in '+-*/=':
                return Char(FN_TEXT, fontpos=ord(ch))
            elif ch == '<':
                return Char(FN_SYMBOL, mtcode=0x3C)
            elif ch == '>':
                return Char(FN_SYMBOL, mtcode=0x3E)
            elif ch == '(':
                return Char(FN_TEXT, fontpos=ord('('))
            elif ch == ')':
                return Char(FN_TEXT, fontpos=ord(')'))
            elif ch == '[':
                return Char(FN_TEXT, fontpos=ord('['))
            elif ch == ']':
                return Char(FN_TEXT, fontpos=ord(']'))
            elif ch == '|':
                return Char(FN_TEXT, fontpos=ord('|'))
            else:
                raise LatexParseError(f"unsupported operator: {ch}")
        
        elif t.kind == 'BRACE' and t.val == '{':
            self.eat()
            line = self.parse_line()
            self.eat('BRACE', '}')
            return line
        
        elif t.kind == 'CHAR':
            self.eat()
            return Char(FN_TEXT, fontpos=ord(t.val))
        
        else:
            raise LatexParseError(f"unexpected token {t.kind} at pos {t.pos}")

    def parse_arg(self):
        """解析 {} 包裹的参数。"""
        self.eat('BRACE', '{')
        line = self.parse_line()
        self.eat('BRACE', '}')
        return line

    def parse_delim_group(self):
        r"""\\left ... \\right 定界符组。"""
        # 吃掉 \left 后的定界符
        left_t = self.peek()
        if left_t and left_t.kind == 'OP':
            left_delim = left_t.val
            self.eat()
        elif left_t and left_t.kind == 'CMD':
            left_delim = '\\' + left_t.val
            self.eat()
        else:
            raise LatexParseError("expected delimiter after \\left")
        
        # 解析中间内容（递归调用 parse_expr，不是 parse_line）
        items = []
        while True:
            t = self.peek()
            if t is None:
                raise LatexParseError("unexpected end, expected \\right")
            # 检查 \right
            if t.kind == 'CMD' and t.val == 'right':
                break
            items.append(self.parse_expr())
        
        content = Line(items=items)
        
        # 吃掉 \right
        self.eat('CMD', 'right')
        
        right_t = self.peek()
        if right_t and right_t.kind == 'OP':
            right_delim = right_t.val
            self.eat()
        elif right_t and right_t.kind == 'CMD':
            right_delim = '\\' + right_t.val
            self.eat()
        else:
            raise LatexParseError("expected delimiter after \\right")
        
        # 映射定界符到 MTEF 模板
        sel = _delim_to_selector(left_delim, right_delim)
        return Tmpl(sel, 0, items=[content])

    def parse_matrix(self, env_name: str):
        """解析矩阵环境。"""
        # 简化实现：只支持简单的行列结构
        rows = []
        current_row = []
        
        while True:
            t = self.peek()
            if t is None:
                raise LatexParseError("unexpected end in matrix")
            
            # 检查 \end{env_name}
            if t.kind == 'CMD' and t.val == 'end':
                self.eat()
                self.parse_arg()  # 吃掉 {env_name}
                break
            
            # 行分隔符 \\
            if t.kind == 'CMD' and t.val == '\\':
                self.eat()
                if current_row:
                    rows.append(current_row)
                    current_row = []
                continue
            
            # 列分隔符 &
            if t.kind == 'ALIGN':
                self.eat()
                continue
            
            # 单元格内容
            cell = self.parse_expr()
            current_row.append(cell)
        
        if current_row:
            rows.append(current_row)
        
        # 构造 MTEF Matrix 或 Pile（简化为 Pile）
        lines = []
        for row in rows:
            line = Line(items=row)
            lines.append(line)
        
        return Pile(halign=1, valign=0, lines=lines)

    def _text_word(self, word: str, typeface: int) -> Line:
        """多字符单词 → Line of Chars。"""
        chars = [Char(typeface, fontpos=ord(c)) for c in word]
        return Line(items=chars)

    def _line_to_text(self, line: Line) -> str:
        """简单提取 Line 中的文本（用于环境名等）。"""
        out = []
        for item in line.items:
            if isinstance(item, Char) and item.fontpos:
                out.append(chr(item.fontpos))
        return ''.join(out)


# ── 希腊字母映射 ──────────────────────────────────────────────────────────

_GREEK_LOWER = {
    'alpha': 0x61, 'beta': 0x62, 'gamma': 0x67, 'delta': 0x64,
    'epsilon': 0x65, 'zeta': 0x7A, 'eta': 0x68, 'theta': 0x71,
    'iota': 0x69, 'kappa': 0x6B, 'lambda': 0x6C, 'mu': 0x6D,
    'nu': 0x6E, 'xi': 0x78, 'pi': 0x70, 'rho': 0x72,
    'sigma': 0x73, 'tau': 0x74, 'upsilon': 0x75, 'phi': 0x66,
    'chi': 0x63, 'psi': 0x79, 'omega': 0x77,
    'varepsilon': 0x65, 'vartheta': 0x4A, 'varpi': 0x76, 'varrho': 0x72,
    'varsigma': 0x56, 'varphi': 0x6A,
}

_GREEK_UPPER = {
    'Gamma': 0x47, 'Delta': 0x44, 'Theta': 0x51, 'Lambda': 0x4C,
    'Xi': 0x58, 'Pi': 0x50, 'Sigma': 0x53, 'Upsilon': 0x55,
    'Phi': 0x46, 'Psi': 0x59, 'Omega': 0x57,
}


def _delim_to_selector(left: str, right: str) -> int:
    """定界符对 → MTEF 模板选择子。"""
    pair = (left, right)
    if pair == ('(', ')'):
        return TM_PAREN
    elif pair == ('[', ']'):
        return TM_BRACK
    elif pair == ('{', '}') or pair == ('\\{', '\\}'):
        return TM_BRACE
    elif pair == ('|', '|'):
        return TM_BAR
    elif pair == ('\\langle', '\\rangle'):
        return TM_ANGLE
    else:
        # 默认圆括号
        return TM_PAREN


# ── 公开接口 ──────────────────────────────────────────────────────────────

def latex_to_mtef(latex_src: str) -> Equation:
    """LaTeX 数学表达式 → MTEF Equation。
    
    Args:
        latex_src: LaTeX 数学源码（不含 $ $ 或 \\[ \\] 定界符）
    
    Returns:
        Equation 对象（可序列化为 MTEF 二进制）
    
    Raises:
        LatexParseError: 解析失败
    """
    toks = _tokenize(latex_src)
    parser = _Parser(toks)
    return parser.parse()
