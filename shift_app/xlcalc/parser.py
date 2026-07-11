"""Excel数式のトークナイザ+パーサ。

AST はタプルで表現する:
  ('num', float) / ('str', s) / ('bool', b) / ('err', code)
  ('ref', sheet|None, col:int, row:int)
  ('range', sheet|None, c1, r1, c2, r2)
  ('call', NAME, [args])
  ('bin', op, left, right) / ('un', op, x) / ('pct', x)
"""
import re

_ERR_RE = re.compile(r'#(REF!|DIV/0!|VALUE!|NAME\?|N/A|NULL!|NUM!|SPILL!)')
_NUM_RE = re.compile(r'\d+(?:\.\d+)?(?:[eE][+-]?\d+)?')
_CELL_RE = re.compile(r'^\$?([A-Za-z]{1,3})\$?(\d+)$')
_COL_RE = re.compile(r'^\$?[A-Za-z]{1,3}$')
# ident の終端になる文字
_IDENT_END = set('()+-*/^&=<>,:%! \t\n"\'')

COL_A = ord('A')


def col_to_num(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - COL_A + 1)
    return n


def num_to_col(n: int) -> str:
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(COL_A + r) + s
    return s


def tokenize(src: str):
    toks = []
    i, n = 0, len(src)
    while i < n:
        ch = src[i]
        if ch in ' \t\n':
            i += 1
            continue
        if ch == '"':
            j = i + 1
            buf = []
            while j < n:
                if src[j] == '"':
                    if j + 1 < n and src[j + 1] == '"':
                        buf.append('"')
                        j += 2
                        continue
                    break
                buf.append(src[j])
                j += 1
            toks.append(('STR', ''.join(buf)))
            i = j + 1
            continue
        if ch == "'":
            j = i + 1
            buf = []
            while j < n:
                if src[j] == "'":
                    if j + 1 < n and src[j + 1] == "'":
                        buf.append("'")
                        j += 2
                        continue
                    break
                buf.append(src[j])
                j += 1
            # クオート済みシート名は直後に '!' が来る
            if j + 1 < n and src[j + 1] == '!':
                toks.append(('SHEET', ''.join(buf)))
                i = j + 2
                continue
            raise ValueError(f'quoted name without !: {src!r}')
        if ch == '#':
            m = _ERR_RE.match(src, i)
            if m:
                toks.append(('ERR', m.group(0)))
                i = m.end()
                continue
            raise ValueError(f'bad error literal in {src!r} at {i}')
        if ch.isdigit() or (ch == '.' and i + 1 < n and src[i + 1].isdigit()):
            m = _NUM_RE.match(src, i)
            toks.append(('NUM', float(m.group(0))))
            i = m.end()
            continue
        if ch in '<>':
            if i + 1 < n and src[i + 1] in '=>':
                toks.append(('OP', src[i:i + 2]))
                i += 2
            else:
                toks.append(('OP', ch))
                i += 1
            continue
        if ch in '()+-*/^&=,:%':
            toks.append(('OP', ch))
            i += 1
            continue
        # ident (関数名/セル参照/シート名/TRUE/FALSE)
        j = i
        while j < n and src[j] not in _IDENT_END:
            j += 1
        word = src[i:j]
        if not word:
            raise ValueError(f'unexpected char {ch!r} in {src!r}')
        if j < n and src[j] == '!':
            toks.append(('SHEET', word))
            i = j + 1
            continue
        up = word.upper()
        if up == 'TRUE':
            toks.append(('BOOL', True))
        elif up == 'FALSE':
            toks.append(('BOOL', False))
        else:
            m = _CELL_RE.match(word)
            if m and col_to_num(m.group(1)) <= 16384:
                toks.append(('CELL', (col_to_num(m.group(1)), int(m.group(2)))))
            else:
                toks.append(('NAME', word))
        i = j
    return toks


class Parser:
    def __init__(self, toks):
        self.toks = toks
        self.i = 0

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def next(self):
        t = self.peek()
        self.i += 1
        return t

    def expect(self, kind, val=None):
        k, v = self.next()
        if k != kind or (val is not None and v != val):
            raise ValueError(f'expected {kind} {val}, got {k} {v}')
        return v

    # 優先順位: 比較 < 連結& < 加減 < 乗除 < 累乗 < 単項 < %
    def parse(self):
        node = self.expr_cmp()
        if self.peek()[0] is not None:
            raise ValueError(f'trailing tokens: {self.toks[self.i:]}')
        return node

    def expr_cmp(self):
        node = self.expr_concat()
        while self.peek() == ('OP', '=') or (self.peek()[0] == 'OP' and self.peek()[1] in ('<', '>', '<=', '>=', '<>')):
            op = self.next()[1]
            node = ('bin', op, node, self.expr_concat())
        return node

    def expr_concat(self):
        node = self.expr_add()
        while self.peek() == ('OP', '&'):
            self.next()
            node = ('bin', '&', node, self.expr_add())
        return node

    def expr_add(self):
        node = self.expr_mul()
        while self.peek()[0] == 'OP' and self.peek()[1] in '+-':
            op = self.next()[1]
            node = ('bin', op, node, self.expr_mul())
        return node

    def expr_mul(self):
        node = self.expr_pow()
        while self.peek()[0] == 'OP' and self.peek()[1] in '*/':
            op = self.next()[1]
            node = ('bin', op, node, self.expr_pow())
        return node

    def expr_pow(self):
        node = self.expr_unary()
        while self.peek() == ('OP', '^'):
            self.next()
            node = ('bin', '^', node, self.expr_unary())
        return node

    def expr_unary(self):
        k, v = self.peek()
        if k == 'OP' and v in '+-':
            self.next()
            return ('un', v, self.expr_unary())
        return self.expr_postfix()

    def expr_postfix(self):
        node = self.primary()
        while self.peek() == ('OP', '%'):
            self.next()
            node = ('pct', node)
        return node

    def primary(self):
        k, v = self.next()
        if k == 'NUM':
            return ('num', v)
        if k == 'STR':
            return ('str', v)
        if k == 'BOOL':
            return ('bool', v)
        if k == 'ERR':
            return ('err', v)
        if k == 'OP' and v == '(':
            node = self.expr_cmp()
            self.expect('OP', ')')
            return node
        if k == 'SHEET':
            sheet = v
            k2, v2 = self.next()
            if k2 == 'ERR':  # 例: 入力シート!#REF!
                return ('err', v2)
            if k2 == 'NAME' and _COL_RE.match(v2):
                return self._col_range(sheet, v2)
            if k2 != 'CELL':
                raise ValueError(f'expected cell after sheet, got {k2} {v2}')
            return self._maybe_range(sheet, v2)
        if k == 'CELL':
            return self._maybe_range(None, v)
        if k == 'NAME' and _COL_RE.match(v) and self.peek() == ('OP', ':'):
            return self._col_range(None, v)
        if k == 'NAME':
            if self.peek() == ('OP', '('):
                self.next()
                args = []
                if self.peek() != ('OP', ')'):
                    args.append(self.expr_cmp())
                    while self.peek() == ('OP', ','):
                        self.next()
                        if self.peek() in (('OP', ','), ('OP', ')')):
                            args.append(('blank',))
                            continue
                        args.append(self.expr_cmp())
                self.expect('OP', ')')
                name = v.upper()
                if name.startswith('_XLFN.'):
                    name = name[6:]
                return ('call', name, args)
            raise ValueError(f'unknown name {v!r}')
        raise ValueError(f'unexpected token {k} {v}')

    def _col_range(self, sheet, col1):
        """列全体参照: C:C / $E:$BH / シート!$B:$B"""
        self.expect('OP', ':')
        k, v = self.next()
        if k == 'SHEET':
            k, v = self.next()
        if k != 'NAME' or not _COL_RE.match(v):
            raise ValueError(f'bad column range end: {k} {v}')
        c1 = col_to_num(col1.replace('$', ''))
        c2 = col_to_num(v.replace('$', ''))
        return ('crange', sheet, min(c1, c2), max(c1, c2))

    def _maybe_range(self, sheet, cell):
        if self.peek() == ('OP', ':'):
            self.next()
            k2, v2 = self.next()
            if k2 == 'SHEET':
                k2, v2 = self.next()
            if k2 != 'CELL':
                raise ValueError('bad range end')
            c1, r1 = cell
            c2, r2 = v2
            return ('range', sheet, min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2))
        return ('ref', sheet, cell[0], cell[1])


_parse_cache = {}


def parse_formula(src: str):
    """'=' 始まりの数式文字列を AST に変換(文字列単位でキャッシュ)。"""
    node = _parse_cache.get(src)
    if node is None:
        body = src[1:] if src.startswith('=') else src
        node = Parser(tokenize(body)).parse()
        _parse_cache[src] = node
    return node
