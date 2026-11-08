"""Bhayankar extractor, DAX module: Tableau calculation -> draft DAX.

Parses a Tableau formula into a tree and emits DAX with a kind (measure / calculated column /
visual calculation), a confidence (high / medium / low) and review notes. Drafts are a starting
point for the engineer: every LOD, table calc and unknown function is marked for review.
Python 3 standard library only.
"""
import re

KEYWORDS = {'IF', 'THEN', 'ELSEIF', 'ELSE', 'END', 'CASE', 'WHEN', 'AND', 'OR', 'NOT', 'FIXED',
            'INCLUDE', 'EXCLUDE', 'TRUE', 'FALSE', 'NULL', 'IN'}
TOKEN = re.compile(r'''
    (?P<ws>\s+)
  | (?P<field>\[(?:[^\]]|\]\])*\](?:\.\[(?:[^\]]|\]\])*\])?)
  | (?P<str>'(?:[^']|'')*'|"(?:[^"]|"")*")
  | (?P<date>\#[^#]*\#)
  | (?P<num>\d+\.?\d*(?:[eE][-+]?\d+)?|\.\d+)
  | (?P<ident>[A-Za-z_][A-Za-z_0-9]*)
  | (?P<op><=|>=|<>|!=|==|&&|\|\||[-+*/%^=<>(),{}:])
''', re.X)

AGGS = {'SUM': ('SUM', 'SUMX'), 'AVG': ('AVERAGE', 'AVERAGEX'), 'MIN': ('MIN', 'MINX'),
        'MAX': ('MAX', 'MAXX'), 'COUNT': ('COUNT', 'COUNTX'), 'MEDIAN': ('MEDIAN', 'MEDIANX'),
        'STDEV': ('STDEV.S', 'STDEVX.S'), 'STDEVP': ('STDEV.P', 'STDEVX.P'),
        'VAR': ('VAR.S', 'VARX.S'), 'VARP': ('VAR.P', 'VARX.P')}
def lookup_zero(n):
    """('fn', 'LOOKUP', [expr, ('num', '0')]): Tableau LOOKUP with offset 0 is the value itself."""
    return n[0] == 'fn' and n[1] == 'LOOKUP' and len(n[2]) == 2 and n[2][1][0] == 'num' and float(n[2][1][1]) == 0


SAME = {'ABS', 'ROUND', 'SQRT', 'EXP', 'LN', 'POWER', 'SIGN', 'LEFT', 'RIGHT', 'LEN', 'UPPER', 'LOWER',
        'TRIM', 'YEAR', 'MONTH', 'DAY', 'QUARTER', 'TODAY', 'NOW', 'HOUR', 'MINUTE', 'SECOND', 'PI',
        'SIN', 'COS', 'TAN', 'ASIN', 'ACOS', 'ATAN', 'DEGREES', 'RADIANS', 'ISOWEEKNUM'}
TABLE_CALCS = {'RUNNING_SUM': 'RUNNINGSUM', 'RUNNING_AVG': 'MOVINGAVERAGE', 'RUNNING_COUNT': 'RUNNINGSUM',
               'RUNNING_MAX': 'RUNNINGSUM', 'RUNNING_MIN': 'RUNNINGSUM', 'WINDOW_SUM': 'WINDOW', 'WINDOW_AVG': 'WINDOW',
               'WINDOW_MAX': 'WINDOW', 'WINDOW_MIN': 'WINDOW', 'WINDOW_COUNT': 'WINDOW', 'WINDOW_MEDIAN': 'WINDOW',
               'LOOKUP': 'OFFSET', 'PREVIOUS_VALUE': 'PREVIOUS', 'INDEX': 'ROWNUMBER', 'FIRST': 'FIRST',
               'LAST': 'LAST', 'SIZE': 'ROWS', 'RANK': 'RANK', 'RANK_DENSE': 'RANK', 'RANK_UNIQUE': 'ROWNUMBER',
               'RANK_MODIFIED': 'RANK', 'RANK_PERCENTILE': 'RANK', 'TOTAL': 'COLLAPSEALL'}
DATE_UNITS = {'year': 'YEAR', 'quarter': 'QUARTER', 'month': 'MONTH', 'week': 'WEEK', 'day': 'DAY',
              'hour': 'HOUR', 'minute': 'MINUTE', 'second': 'SECOND'}


class ParseError(Exception):
    pass


COMMENT = re.compile(r"""('(?:[^']|'')*'|"(?:[^"]|"")*")|//[^\n]*""")


def tokenize(f):
    f = COMMENT.sub(lambda m: m.group(1) or '', f.replace('\r', ''))  # Tableau // comments, not inside strings
    out, pos = [], 0
    while pos < len(f):
        m = TOKEN.match(f, pos)
        if not m:
            raise ParseError(f'unexpected character {f[pos]!r} at {pos}')
        pos = m.end()
        kind = m.lastgroup
        if kind == 'ws':
            continue
        val = m.group(kind)
        if kind == 'ident' and val.upper() in KEYWORDS:
            kind, val = 'kw', val.upper()
        out.append((kind, val))
    out.append(('eof', None))
    return out


class Parser:
    """Precedence (low -> high): OR, AND, NOT, comparison/IN, + -, * / %, ^, unary -, primary."""

    def __init__(self, f):
        self.t = tokenize(f)
        self.i = 0

    def peek(self, k=0):
        return self.t[self.i + k]

    def take(self, kind=None, val=None):
        tk = self.t[self.i]
        if (kind and tk[0] != kind) or (val and tk[1] != val):
            raise ParseError(f'expected {val or kind}, got {tk[1]!r}')
        self.i += 1
        return tk

    def at(self, kind, val=None):
        tk = self.t[self.i]
        return tk[0] == kind and (val is None or tk[1] == val)

    def parse(self):
        e = self.expr()
        if not self.at('eof'):
            raise ParseError(f'unexpected {self.peek()[1]!r}')
        return e

    def expr(self):
        return self.or_()

    def or_(self):
        e = self.and_()
        while self.at('kw', 'OR') or self.at('op', '||'):
            self.i += 1
            e = ('bin', '||', e, self.and_())
        return e

    def and_(self):
        e = self.not_()
        while self.at('kw', 'AND') or self.at('op', '&&'):
            self.i += 1
            e = ('bin', '&&', e, self.not_())
        return e

    def not_(self):
        if self.at('kw', 'NOT'):
            self.i += 1
            return ('not', self.not_())
        return self.cmp()

    def cmp(self):
        e = self.add()
        while True:
            if self.at('op') and self.peek()[1] in ('=', '==', '!=', '<>', '<', '<=', '>', '>='):
                op = self.take()[1]
                op = {'==': '=', '!=': '<>'}.get(op, op)
                e = ('bin', op, e, self.add())
            elif self.at('kw', 'IN'):
                self.i += 1
                self.take('op', '(')
                vals = [self.expr()]
                while self.at('op', ','):
                    self.i += 1
                    vals.append(self.expr())
                self.take('op', ')')
                e = ('in', e, vals)
            else:
                return e

    def add(self):
        e = self.mul()
        while self.at('op') and self.peek()[1] in ('+', '-'):
            op = self.take()[1]
            e = ('bin', op, e, self.mul())
        return e

    def mul(self):
        e = self.pow_()
        while self.at('op') and self.peek()[1] in ('*', '/', '%'):
            op = self.take()[1]
            e = ('bin', op, e, self.pow_())
        return e

    def pow_(self):
        e = self.unary()
        if self.at('op', '^'):
            self.i += 1
            e = ('bin', '^', e, self.pow_())
        return e

    def unary(self):
        if self.at('op', '-'):
            self.i += 1
            return ('neg', self.unary())
        if self.at('op', '+'):
            self.i += 1
            return self.unary()
        return self.primary()

    def primary(self):
        kind, val = self.peek()
        if kind == 'num':
            self.i += 1
            return ('num', val)
        if kind == 'str':
            self.i += 1
            q = val[0]
            return ('str', val[1:-1].replace(q + q, q))
        if kind == 'date':
            self.i += 1
            return ('date', val.strip('#'))
        if kind == 'field':
            self.i += 1
            return ('field', val)
        if kind == 'kw' and val in ('TRUE', 'FALSE'):
            self.i += 1
            return ('bool', val)
        if kind == 'kw' and val == 'NULL':
            self.i += 1
            return ('null',)
        if kind == 'kw' and val == 'IF':
            return self.if_()
        if kind == 'kw' and val == 'CASE':
            return self.case()
        if kind == 'op' and val == '(':
            self.i += 1
            e = self.expr()
            self.take('op', ')')
            return e
        if kind == 'op' and val == '{':
            return self.lod()
        if kind == 'ident':
            self.i += 1
            name = val.upper()
            args = []
            self.take('op', '(')
            if not self.at('op', ')'):
                args.append(self.expr())
                while self.at('op', ','):
                    self.i += 1
                    args.append(self.expr())
            self.take('op', ')')
            return ('fn', name, args)
        raise ParseError(f'unexpected {val!r}')

    def if_(self):
        self.take('kw', 'IF')
        branches = []
        c = self.expr()
        self.take('kw', 'THEN')
        branches.append((c, self.expr()))
        other = None
        while True:
            if self.at('kw', 'ELSEIF'):
                self.i += 1
                c = self.expr()
                self.take('kw', 'THEN')
                branches.append((c, self.expr()))
            elif self.at('kw', 'ELSE'):
                self.i += 1
                other = self.expr()
            else:
                break
        self.take('kw', 'END')
        return ('if', branches, other)

    def case(self):
        self.take('kw', 'CASE')
        subject = self.expr()
        whens, other = [], None
        while self.at('kw', 'WHEN'):
            self.i += 1
            w = [self.expr()]
            while self.at('op', ','):  # WHEN a, b THEN
                self.i += 1
                w.append(self.expr())
            self.take('kw', 'THEN')
            whens.append((w, self.expr()))
        if self.at('kw', 'ELSE'):
            self.i += 1
            other = self.expr()
        self.take('kw', 'END')
        return ('case', subject, whens, other)

    def lod(self):
        self.take('op', '{')
        kind = 'FIXED'
        dims = []
        if self.at('kw') and self.peek()[1] in ('FIXED', 'INCLUDE', 'EXCLUDE'):
            kind = self.take()[1]
            if not self.at('op', ':'):
                dims.append(self.expr())
                while self.at('op', ','):
                    self.i += 1
                    dims.append(self.expr())
            self.take('op', ':')
        e = self.expr()
        self.take('op', '}')
        return ('lod', kind, dims, e)


def walk(n):
    yield n
    if not isinstance(n, tuple):
        return
    for x in n[1:]:
        if isinstance(x, tuple):
            yield from walk(x)
        elif isinstance(x, list):
            for y in x:
                if isinstance(y, tuple) and y and isinstance(y[0], str):
                    yield from walk(y)
                elif isinstance(y, tuple):  # (cond, then) pairs or (whens, then)
                    for z in y:
                        if isinstance(z, list):
                            for q in z:
                                yield from walk(q)
                        elif isinstance(z, tuple):
                            yield from walk(z)


def walk_outer(n):
    """walk() that does not enter LOD expressions: what the formula does at its own level (a FIXED LOD is one value per
    row there, whatever aggregation it holds inside)."""
    yield n
    if isinstance(n, tuple) and n and n[0] == 'lod':
        return
    for x in walk(n):
        if x is n:
            continue
        if isinstance(x, tuple) and x and x[0] == 'lod':
            yield x
            continue
        if not any(x is y or _inside(x, y) for y in _lods(n)):
            yield x


def _lods(n):
    return [x for x in walk(n) if isinstance(x, tuple) and x and x[0] == 'lod' and x is not n]


def _inside(x, lod):
    return any(x is y for y in walk(lod) if y is not lod)


def strip_brackets(s):
    """'(x + y)' -> 'x + y' when the outer brackets enclose the whole DAX text ('(a) / (b)' is left alone). Strings,
    quoted table names and [column] names may hold brackets of their own."""
    if not (s.startswith('(') and s.endswith(')')):
        return s
    depth, i, close = 0, 0, {'"': '"', "'": "'", '[': ']'}
    while i < len(s):
        ch = s[i]
        if ch in close:
            j = s.find(close[ch], i + 1)
            while j != -1 and ch != '[' and s[j + 1:j + 2] == close[ch]:  # "" and '' escapes
                j = s.find(close[ch], j + 2)
            if j == -1:
                return s
            i = j + 1
            continue
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
            if depth == 0 and i < len(s) - 1:
                return s
        i += 1
    return s[1:-1]


def q(name):
    return "'" + name.replace("'", "''") + "'"


def col(table, name):
    return f"{q(table)}[{name.replace(']', ']]')}]"


class Translator:
    """resolve(field_token) -> dict(kind='column'|'measure'|'rowcalc'|'param'|'unknown',
    table=, name=, datatype=) supplied by the extractor."""

    def __init__(self, resolve):
        self.resolve = resolve

    def translate(self, formula, role=None, default_agg=None, name=None, window=None):
        """window: the DAX columns a table calculation runs over on one sheet (Tableau "compute using", e.g.
        ["'Date'[Year Month]"]; [] = each mark on its own). With it, RANK / WINDOW_* / RUNNING_* / TOTAL / INDEX /
        SIZE / LOOKUP become measures; without it (or for a table calc it cannot write) they stay placeholders."""
        self.notes, self.conf = [], 'high'
        self.window, self._tc_fallback = window, False
        if not formula or not formula.strip():
            return {'dax': 'BLANK()', 'kind': 'measure', 'confidence': 'low', 'notes': ['empty formula']}
        try:
            tree = Parser(formula).parse()
        except ParseError as e:
            return {'dax': None, 'kind': 'unknown', 'confidence': 'low', 'notes': [f'parse error: {e}; translate by hand']}
        nodes = list(walk(tree))
        outer = list(walk_outer(tree))  # an aggregation inside {FIXED ...} does not make the formula a measure
        param_like = lambda r: r['kind'] == 'param' or (r['kind'] == 'measure' and r.get('param_only'))
        has_agg = any(n[0] == 'fn' and (n[1] in AGGS or n[1] in ('COUNTD', 'ATTR')) for n in outer) or \
            any(n[0] == 'field' and self.resolve(n[1])['kind'] == 'measure' and not self.resolve(n[1]).get('param_only')
                for n in outer) or \
            any(n[0] == 'lod' and n[1] != 'FIXED' for n in outer if n is not tree)  # INCLUDE/EXCLUDE: view level
        has_tc = any(n[0] == 'fn' and n[1] in TABLE_CALCS and not lookup_zero(n) for n in nodes)
        has_lod = any(n[0] == 'lod' for n in nodes)
        fields = [n for n in nodes if n[0] == 'field' and self.resolve(n[1])['kind'] in ('column', 'rowcalc')]
        uses_param = any(n[0] == 'field' and param_like(self.resolve(n[1])) for n in nodes)  # incl. parameter-only calcs
        # a row-level formula over calcs that are measures only by the way Tableau shows them (a row calc the view
        # aggregates, a FIXED LOD): one value per row, so a calculated column that inlines their row-level DAX
        refs = [self.resolve(n[1]) for n in nodes if n[0] == 'field']
        meas_refs = [r for r in refs if r['kind'] == 'measure' and not r.get('param_only')]
        self.rowify = bool(meas_refs) and all(r.get('row_dax') for r in meas_refs) and not uses_param and not has_tc \
            and not has_lod and not any(n[0] == 'fn' and (n[1] in AGGS or n[1] in ('COUNTD', 'ATTR')) for n in nodes)
        if has_tc and window is None:
            kind = 'visual calculation'
        elif has_tc:
            kind = 'measure'  # written over the sheet's addressing columns (falls back below when it cannot be)
        elif self.rowify:
            kind = 'calculated column'
            self.note('row-level formula over FIXED / row calcs: calculated column with their row-level DAX inlined', 'medium')
        elif has_agg or (has_lod and tree[0] == 'lod') or not fields:
            kind = 'measure'
        else:
            kind = 'calculated column'
        self.measure_ctx = kind != 'calculated column'
        self.bare = []  # columns referenced outside any aggregation while in measure context
        # CASE/IF on a parameter whose branches return columns = a dynamic field selector -> field parameter
        branches = [t for _, t in tree[2]] if tree[0] == 'case' else [t for _, t in tree[1]] if tree[0] == 'if' else []
        subject_param = (tree[0] == 'case' and tree[1][0] == 'field' and self.resolve(tree[1][1])['kind'] == 'param') or \
            (tree[0] == 'if' and uses_param)
        kinds = [self.resolve(b[1])['kind'] if b[0] == 'field' else None for b in branches]
        # every branch a true/false row field (CASE [Param] WHEN "Last Week" THEN [Last Week] ...): the parameter picks
        # which rows pass a sheet filter, not which field a visual shows. A row formula (SWITCH over the parameter
        # value) that the build folds into the measures of the sheets it filters
        flags = all(b[0] == 'field' and (self.resolve(b[1]).get('datatype') or '') == 'boolean' and
                    self.resolve(b[1])['kind'] in ('column', 'rowcalc') for b in branches)
        if subject_param and branches and not flags and all(k in ('column', 'rowcalc', 'measure') for k in kinds) and \
                any(k in ('column', 'rowcalc') for k in kinds):  # field parameters may hold columns and measures
            names = [self.resolve(b[1])['name'] for b in branches]
            return {'dax': None, 'kind': 'field parameter', 'confidence': 'medium', 'fields': names,
                    'notes': [f'dynamic field selector over {", ".join(names)}: build a field parameter '
                              '(field-parameters.md), not SWITCH (bare columns are invalid in a measure)']}
        sel = self.measure_slice(tree) if self.measure_ctx else None
        dax = sel or self.emit(tree, row=not self.measure_ctx)
        row_dax = dax if kind == 'calculated column' else self.fixed_row_dax(tree)
        if self.bare:
            self.note('mixes row-level columns (' + ', '.join(sorted(set(self.bare))) + ') with measures/aggregates: '
                      'decide the grain (measure with an iterator, or a measure filter); not valid DAX as written', 'low')
        if kind == 'calculated column' and uses_param:
            self.note('reads a parameter: a calculated column is computed at refresh and cannot react to slicers; '
                      'use it as a measure filter or move the logic into measures', 'medium')
        if kind == 'calculated column' and role == 'measure':
            # row-level calc that Tableau aggregates in the view (default aggregation)
            agg = (default_agg or 'Sum').upper()
            agg = {'USER': 'SUM', 'COUNTD': 'COUNTD'}.get(agg, agg)
            table = self.main_table(tree)
            if agg in AGGS and table:
                dax = f'{AGGS[agg][1]}({q(table)}, {dax})'
                kind = 'measure'
                self.note(f'row-level calc aggregated in the view by {agg}: written as an iterator measure', 'medium')
        if kind == 'calculated column':
            self.notes.append('row-level: prefer a Power Query column (folds to the source); DAX column shown')
        if re.search(r'(^|[(,\s])1\s*[-+]\s*DIVIDE\(', dax or ''):
            self.note('dax-guidelines: never 1 - (x/y); rewrite with VAR and one DIVIDE', 'medium')
        if has_tc and self._tc_fallback:
            kind = 'visual calculation'
        out = {'dax': dax, 'kind': kind, 'confidence': self.conf, 'notes': self.notes}
        if row_dax and not uses_param:
            out['row_dax'] = row_dax  # the per-row value, for row-level calcs that reference this one
        return out

    def fixed_row_dax(self, tree):
        """{FIXED dims: agg} over one table's plain columns -> its value on each row of that table (calc column)."""
        if tree[0] != 'lod' or tree[1] != 'FIXED' or not all(self.plain_column(d) for d in tree[2]):
            return None
        table = self.main_table(tree[3])
        if not table or any(self.resolve(d[1]).get('table') != table for d in tree[2]):
            return None
        saved, self.measure_ctx = self.measure_ctx, True
        inner = self.emit(tree[3], row=False)
        self.measure_ctx = saved
        if not tree[2]:  # {FIXED : agg} / {agg}: one value for the whole table (e.g. latest date), the same on every row
            return f'CALCULATE({inner}, REMOVEFILTERS({q(table)}))'
        return f'CALCULATE({inner}, ALLEXCEPT({q(table)}, {", ".join(self.emit(d, True) for d in tree[2])}))'

    def note(self, msg, conf):
        if msg not in self.notes:
            self.notes.append(msg)
        order = {'high': 0, 'medium': 1, 'low': 2}
        if order[conf] > order[self.conf]:
            self.conf = conf

    def main_table(self, n):
        for x in walk(n):
            if x[0] == 'field':
                r = self.resolve(x[1])
                if r['kind'] in ('column', 'rowcalc') and r.get('table'):
                    return r['table']
        return None

    def is_string(self, n):
        if n[0] == 'str':
            return True
        if n[0] == 'field':
            return (self.resolve(n[1]).get('datatype') or '') == 'string'
        if n[0] == 'fn' and n[1] in ('STR', 'LEFT', 'RIGHT', 'MID', 'UPPER', 'LOWER', 'TRIM', 'LTRIM', 'RTRIM',
                                     'REPLACE', 'ATTR', 'DATENAME', 'SPACE'):
            return n[1] != 'ATTR' or any(self.is_string(a) for a in n[2])
        if n[0] == 'bin' and n[1] == '+':
            return self.is_string(n[2]) or self.is_string(n[3])
        if n[0] in ('if',):
            return any(self.is_string(t) for _, t in n[1])
        if n[0] == 'case':
            return any(self.is_string(t) for _, t in n[2])
        return False

    def non_numeric(self, n):
        """True when an expression gives text, a date or true/false (MIN/MAX/ATTR of such a field, a string formula)."""
        if self.is_string(n):
            return True
        if n[0] == 'fn' and n[1] in ('MIN', 'MAX', 'ATTR') and len(n[2]) == 1:
            return self.non_numeric(n[2][0])
        if n[0] == 'field':
            return (self.resolve(n[1]).get('datatype') or '') in ('string', 'date', 'datetime', 'boolean')
        return n[0] in ('date', 'str', 'bool')

    def plain_column(self, n):
        return n[0] == 'field' and self.resolve(n[1])['kind'] == 'column'

    # ---------- emitter ----------
    # DAX precedence of what emit() writes for a node (higher binds tighter; 10 = a function call or a bracketed form).
    # The parse tree has no bracket nodes, so a sub-expression that binds looser than its parent must be bracketed again:
    # Tableau ([a] + [b]) * [c] is not DAX 'T'[a] + 'T'[b] * 'T'[c].
    PREC = {'=': 4, '<>': 4, '<': 4, '<=': 4, '>': 4, '>=': 4, '&': 5, '+': 6, '-': 6, '*': 7, '/': 7}

    def prec(self, n):
        if n[0] == 'bin':
            op = n[1]
            if op == '+' and (self.is_string(n[2]) or self.is_string(n[3])):
                return self.PREC['&']
            return self.PREC.get(op, 10)  # && || are written bracketed, % ^ as MOD() / POWER()
        return {'in': 4, 'not': 3, 'neg': 8}.get(n[0], 10)

    def bracket(self, parent, a, ea, b, eb):
        """(ea, eb) bracketed where the operand binds looser than the parent operator (or as loose, on the right of a
        non-associative one: a - (b - c), a / (b * c), comparisons)."""
        pp = self.prec(parent)
        if pp >= 10:
            return ea, eb
        if self.prec(a) < pp:
            ea = f'({ea})'
        pb = self.prec(b)
        if pb < pp or (pb == pp and parent[1] in ('-', '/', '=', '<>', '<', '<=', '>', '>=')):
            eb = f'({eb})'
        return ea, eb

    def emit(self, n, row):
        k = n[0]
        if k == 'num':
            return n[1]
        if k == 'str':
            return '"' + n[1].replace('"', '""') + '"'
        if k == 'bool':
            return n[1] + '()'
        if k == 'null':
            return 'BLANK()'
        if k == 'date':
            m = re.match(r'(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?', n[1])
            if not m:
                self.note(f'date literal #{n[1]}# kept as text', 'medium')
                return f'DATEVALUE("{n[1]}")'
            y, mo, d, h, mi, s = m.groups()
            out = f'DATE({int(y)}, {int(mo)}, {int(d)})'
            if h:
                out += f' + TIME({int(h)}, {int(mi)}, {int(s or 0)})'
            return out
        if k == 'field':
            r = self.resolve(n[1])
            if r['kind'] == 'param':
                return f'[{r["name"]} Value]'
            if r['kind'] == 'measure':
                if getattr(self, 'rowify', False) and r.get('row_dax'):
                    return f'({r["row_dax"]})'
                return f'[{r["name"]}]'
            if r['kind'] in ('column', 'rowcalc'):
                if not r.get('table'):
                    self.note(f'source table of [{r["name"]}] unknown: replace <Table>', 'medium')
                if self.measure_ctx and not row:
                    self.bare.append(r['name'])
                return col(r.get('table') or '<Table>', r['name'])
            self.note(f'unresolved field {n[1]}', 'low')
            return n[1]
        if k == 'neg':
            inner = self.emit(n[1], row)
            return '-' + (f'({inner})' if self.prec(n[1]) < 10 else inner)  # -([a] + [b]) is not -'T'[a] + 'T'[b]
        if k == 'not':
            return f'NOT ({self.emit(n[1], row)})'
        if k == 'in':
            return f'{self.emit(n[1], row)} IN {{{", ".join(self.emit(v, row) for v in n[2])}}}'
        if k == 'bin':
            op, a, b = n[1], n[2], n[3]
            # dax-guidelines: never 1 - x/y or x/y - 1: compute the full expression with VAR + one DIVIDE
            one = lambda x: x[0] == 'num' and float(x[1]) == 1
            ratio = lambda x: x[0] == 'bin' and x[1] == '/'
            if op in ('-', '+') and self.measure_ctx and not row and ((one(a) and ratio(b)) or (ratio(a) and one(b))):
                r_ = b if ratio(b) else a
                num, den = self.emit(r_[2], row), self.emit(r_[3], row)
                if one(a):
                    body = f'_den {op} _num' if op == '+' else '_den - _num'
                else:
                    body = f'_num {op} _den'
                return f'(VAR _num = {num} VAR _den = {den} RETURN DIVIDE({body}, _den))'
            ea, eb = self.bracket(n, a, self.emit(a, row), b, self.emit(b, row))
            if op in ('=', '<>') and (a[0] == 'bool') != (b[0] == 'bool'):
                # Tableau: NULL = TRUE/FALSE is NULL (IF takes ELSE). DAX: BLANK() = FALSE() is TRUE, so a blank
                # (no related row, NULL text) would count as FALSE. == is strict: BLANK() == FALSE() is FALSE
                if op == '=':
                    return f'{ea} == {eb}'
                return f'(NOT ISBLANK({eb if a[0] == "bool" else ea}) && {ea} <> {eb})'
            if op in ('<', '<=', '>', '>=', '=', '<>') and (row or not self.measure_ctx) and (a[0] == 'num') != (b[0] == 'num') \
                    and (b if a[0] == 'num' else a)[0] not in ('str', 'bool', 'null', 'date'):
                # Tableau: NULL compared with a number is NULL (IF takes ELSE); DAX: BLANK() <= 40 is TRUE
                self._vn = getattr(self, '_vn', 0) + 1
                v = f'_v{self._vn}'
                cmp_ = f'{v} {op} {eb}' if b[0] == 'num' else f'{ea} {op} {v}'
                return f'(VAR {v} = {eb if a[0] == "num" else ea} RETURN NOT ISBLANK({v}) && {cmp_})'
            if op == '+' and (self.is_string(a) or self.is_string(b)):
                return f'{ea} & {eb}'
            if op == '%':
                return f'MOD({ea}, {eb})'
            if op == '^':
                return f'POWER({ea}, {eb})'
            if op == '/':
                # Tableau gives NULL for x / 0; DAX "/" gives infinity (row level too): DIVIDE, unless the
                # denominator is a non-zero number
                if not (b[0] == 'num' and float(b[1]) != 0):
                    return f'DIVIDE({strip_brackets(ea)}, {strip_brackets(eb)})'
                return f'{ea} / {eb}'  # operands bracketed by precedence above
            return f'({ea} {op} {eb})' if op in ('&&', '||') else f'{ea} {op} {eb}'
        if k == 'if':
            out = self.emit(n[2], row) if n[2] is not None else None
            for c, t in reversed(n[1]):
                out = f'IF({self.emit(c, row)}, {self.emit(t, row)}' + (f', {out})' if out is not None else ')')
            return out
        if k == 'case':
            parts = [self.emit(n[1], row)]
            for ws, t in n[2]:
                et = self.emit(t, row)
                for w in ws:
                    parts += [self.emit(w, row), et]
            if n[3] is not None:
                parts.append(self.emit(n[3], row))
            return 'SWITCH(' + ', '.join(parts) + ')'
        if k == 'lod':
            return self.lod(n)
        if k == 'fn':
            return self.fn(n[1], n[2], row)
        self.note(f'unhandled node {k}', 'low')
        return '/* ? */'

    def lod(self, n):
        kind, dims, e = n[1], n[2], n[3]
        if kind == 'FIXED' and not self.measure_ctx:  # inside a calculated column: the value on this row
            row = self.fixed_row_dax(n)
            if row:
                return row
        inner = self.emit(e, row=False)
        dcols = [self.emit(d, row=True) for d in dims]
        table = self.main_table(e) or (self.main_table(dims[0]) if dims else None) or '<Table>'
        if kind == 'FIXED':
            self.note('FIXED LOD: value per grain ignoring other filters on this table (Tableau keeps only context '
                      'filters); filters from related dimension tables still apply in Power BI. Check which slicers must '
                      'affect it. View aggregation assumed SUM', 'medium')
            if not dims:
                return f'CALCULATE({inner}, REMOVEFILTERS())'
            # per grain value: keep only the grain key(s), remove every other filter on the table
            # (dax-guidelines: VAR with _ prefix, REMOVEFILTERS over ALL, column predicates as filters)
            keys = [f'_key{i + 1}' for i in range(len(dcols))]
            plain = [self.plain_column(d) for d in dims]
            if all(plain):
                rows = f'VALUES({dcols[0]})' if len(dcols) == 1 else f'SUMMARIZE({q(table)}, {", ".join(dcols)})'
                grain, preds = dcols, [f'{c} = {k}' for k, c in zip(keys, dcols)]
            else:
                # an expression in the grain (e.g. DATEDIFF(...) <= 12): SUMMARIZE groups by columns only, so add it
                # as a column first and filter it back with FILTER over the table
                self.note('FIXED LOD grain holds an expression: grouped through ADDCOLUMNS', 'medium')
                ext = ', '.join(f'"@d{i + 1}", {c}' for i, (c, p) in enumerate(zip(dcols, plain)) if not p)
                grain = [c if p else f'[@d{i + 1}]' for i, (c, p) in enumerate(zip(dcols, plain))]
                rows = f'SUMMARIZE(ADDCOLUMNS({q(table)}, {ext}), {", ".join(grain)})'
                preds = [f'{c} = {k}' if p else f'FILTER(ALL({q(table)}), {c} = {k})'
                         for k, c, p in zip(keys, dcols, plain)]
            var = ' '.join(f'VAR {k} = {c}' for k, c in zip(keys, grain))
            # the view adds the per-grain values up (SUM), but a text or date value cannot be summed: SUMX over text
            # errors and over dates gives a meaningless number. Those take the MIN / MAX of the grain values instead
            outer = 'SUMX'
            if self.non_numeric(e):
                outer = 'MINX' if e[0] == 'fn' and e[1] == 'MIN' else 'MAXX'
                self.note(f'FIXED LOD of a text / date value: {outer} over the grain (a sum is meaningless)', 'medium')
            return f'{outer}({rows}, {var} RETURN CALCULATE({inner}, REMOVEFILTERS({q(table)}), {", ".join(preds)}))'
        if kind == 'INCLUDE':
            self.note('INCLUDE LOD: outer aggregation assumed SUM over the finer grain; check the view aggregation', 'medium')
            grain = ', '.join(dcols)
            return f'SUMX(SUMMARIZE({q(table)}, {grain}), CALCULATE({inner}))'
        self.note('EXCLUDE LOD: removes the listed dimensions from the filter context', 'medium')
        return f'CALCULATE({inner}, ' + ', '.join(f'REMOVEFILTERS({c})' for c in dcols) + ')'

    def agg(self, name, a, row):
        simple, iterator = AGGS[name]
        if a[0] == 'field' and self.resolve(a[1])['kind'] == 'measure':
            # SUM([calc]) where the calc became a measure: it already aggregates (DAX SUM takes a column only)
            self.note(f'{name}() around a calc that is a measure: the measure is used as is', 'medium')
            return self.emit(a, row)
        if self.plain_column(a):
            return f'{simple}({self.emit(a, True)})'
        table = self.main_table(a)
        if not table:
            if name in ('AVG', 'MIN', 'MAX', 'MEDIAN'):
                # AVG(0), MIN(1), ... of a constant (a reference-line helper): the average/min/max of a constant is the
                # constant; DAX AVERAGE/MIN/MAX of a literal is not valid (they need a column)
                self.note(f'{name}() of a constant: the constant itself (DAX {simple} needs a column)', 'low')
                return self.emit(a, row)
            return f'{simple}({self.emit(a, row)})'
        return f'{iterator}({q(table)}, {self.emit(a, True)})'

    WEEKDAYS = ('sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday')

    def week_start(self, a):
        """Tableau start_of_week literal -> days after Sunday (0..6); None when not given or not a literal."""
        if a is None or a[0] != 'str' or a[1].lower() not in self.WEEKDAYS:
            return None
        return self.WEEKDAYS.index(a[1].lower())

    def date_unit(self, a):
        if a[0] != 'str':
            self.note('date part is not a literal', 'low')
            return 'DAY'
        return DATE_UNITS.get(a[1].lower(), a[1].upper())

    # Date table columns whose text does not sort in time order: order by their number column
    ORDER_KEY = {"'Date'[Month]": "'Date'[Month Number]", "'Date'[Year Quarter]": "'Date'[Year Quarter Number]",
                 "'Date'[Day of Week]": "'Date'[Day of Week Number]", "'Date'[Month Day]": "'Date'[Month Day Number]"}
    WINDOW_X = {'SUM': 'SUMX', 'AVG': 'AVERAGEX', 'MAX': 'MAXX', 'MIN': 'MINX', 'COUNT': 'COUNTX', 'MEDIAN': 'MEDIANX'}

    def window_fn(self, name, args):
        """A Tableau table calculation as a DAX measure over self.window (the addressing columns of the sheet; the
        visual's other columns partition it through the filter context, as Tableau's partitioning dimensions do).
        None when this function / argument shape is not handled (the caller keeps the placeholder)."""
        cols = self.window
        rel = f'ALLSELECTED({", ".join(cols)})' if cols else None
        E = lambda x: self.emit(x, False)  # noqa: E731
        lit = lambda a: a[0] == 'num' or (a[0] == 'neg' and a[1][0] == 'num')  # noqa: E731
        if name in ('RANK', 'RANK_DENSE', 'RANK_UNIQUE', 'RANK_MODIFIED') and 1 <= len(args) <= 2:
            if len(args) == 2 and args[1][0] != 'str':
                return None
            e = E(args[0])
            order = 'ASC' if len(args) == 2 and args[1][1].lower() == 'asc' else 'DESC'
            if not rel:
                return f'IF(ISBLANK({e}), BLANK(), 1)'
            if name in ('RANK_UNIQUE', 'RANK_MODIFIED'):
                self.note(f'{name}: tied marks share a rank in RANKX (Tableau numbers ties differently)', 'medium')
            ties = 'DENSE' if name == 'RANK_DENSE' else 'SKIP'
            return f'IF(ISBLANK({e}), BLANK(), RANKX({rel}, CALCULATE({e}), , {order}, {ties}))'
        if name == 'INDEX' and not args:
            if not cols:
                return '1'
            keys = [self.ORDER_KEY.get(c, c) for c in cols]
            rel_k = f'ALLSELECTED({", ".join(dict.fromkeys(cols + keys))})'
            one = ' && '.join(f'HASONEVALUE({c})' for c in cols)
            return f'IF({one}, ROWNUMBER({rel_k}, ORDERBY({", ".join(f"{k}, ASC" for k in keys)})))'
        if name == 'SIZE' and not args:
            return f'COUNTROWS({rel})' if rel else '1'
        if name == 'TOTAL' and len(args) == 1:
            e = E(args[0])
            return f'CALCULATE({e}, {rel})' if rel else e
        if name.startswith('WINDOW_') and name[7:] in self.WINDOW_X and len(args) == 1:
            e = E(args[0])
            return f'{self.WINDOW_X[name[7:]]}({rel}, CALCULATE({e}))' if rel else e
        if name.startswith('RUNNING_') and name[8:] in self.WINDOW_X and len(args) == 1 and len(cols) <= 1:
            e = E(args[0])
            if not cols:
                return e
            c = cols[0]
            k = self.ORDER_KEY.get(c, c)
            rows = f'ALLSELECTED({c}, {k})' if k != c else f'ALLSELECTED({c})'
            return f'{self.WINDOW_X[name[8:]]}(FILTER({rows}, {k} <= MAX({k})), CALCULATE({e}))'
        if name == 'LOOKUP' and len(args) == 2 and lit(args[1]) and len(cols) == 1:
            n = int(float(args[1][1] if args[1][0] == 'num' else '-' + args[1][1][1]))
            c = cols[0]
            k = self.ORDER_KEY.get(c, c)
            rows = f'ALLSELECTED({c}, {k})' if k != c else f'ALLSELECTED({c})'
            return f'IF(HASONEVALUE({c}), CALCULATE({E(args[0])}, OFFSET({n}, {rows}, ORDERBY({k}, ASC))))'
        return None

    def date_part(self, u, d):
        """DATEPART unit (as date_unit() returns it) of the DAX date d. Units with no DAX function of their name
        (dayofyear, iso-*) are computed; an unknown unit is marked, never written as a function that does not exist."""
        if u in ('YEAR', 'QUARTER', 'MONTH', 'DAY', 'HOUR', 'MINUTE', 'SECOND'):
            return f'{u}({d})'
        other = {'WEEK': f'WEEKNUM({d})', 'WEEKDAY': f'WEEKDAY({d})',
                 'DAYOFYEAR': f'(DATEDIFF(DATE(YEAR({d}), 1, 1), {d}, DAY) + 1)',
                 'ISO-WEEK': f'WEEKNUM({d}, 21)', 'ISO-WEEKDAY': f'WEEKDAY({d}, 2)',
                 'ISO-YEAR': f'YEAR({d} - WEEKDAY({d}, 2) + 4)'}  # ISO year: the year of the week's Thursday
        if u in other:
            return other[u]
        self.note(f'date part {u.lower()} not mapped', 'low')
        return f'/* TODO date part {u.lower()} */ BLANK()'

    def measure_slice(self, tree):
        """CASE [column] WHEN v1, v2 THEN [measure] END (one branch, no ELSE) in a measure: the measure for the rows
        whose column holds those values, i.e. CALCULATE([measure], KEEPFILTERS(column IN {v1, v2})). As written it
        mixes a row-level column with a measure, which DAX rejects."""
        if tree[0] != 'case' or tree[3] is not None or len(tree[2]) != 1 or tree[1][0] != 'field':
            return None
        subj = self.resolve(tree[1][1])
        vals, then = tree[2][0]
        if subj['kind'] != 'column' or then[0] != 'field' or self.resolve(then[1])['kind'] != 'measure' or \
                not all(v[0] in ('num', 'str', 'bool') for v in vals):
            return None
        col_dax = self.emit(tree[1], row=True)
        lits = ', '.join(self.emit(v, row=True) for v in vals)
        self.note(f'CASE on [{subj["name"]}] returning a measure: the measure filtered to those values', 'medium')
        return f'CALCULATE({self.emit(then, row=False)}, KEEPFILTERS({col_dax} IN {{{lits}}}))'

    def fn(self, name, args, row):
        E = lambda x: self.emit(x, row)
        if name in AGGS and len(args) == 1:
            return self.agg(name, args[0], row)
        if name in ('MIN', 'MAX') and len(args) == 2:
            return f'{name}({E(args[0])}, {E(args[1])})'
        if name == 'COUNTD':
            a = args[0]
            # Tableau COUNTD ignores NULL; DISTINCTCOUNT and DISTINCT() keep BLANK as one more value (+1 whenever a row is
            # blank, and in COUNTD(IF c THEN x END) whenever any row fails c)
            if self.plain_column(a):
                return f'DISTINCTCOUNTNOBLANK({self.emit(a, True)})'
            t = self.main_table(a) or '<Table>'
            return f'COUNTROWS(FILTER(DISTINCT(SELECTCOLUMNS({q(t)}, "@v", {self.emit(a, True)})), NOT ISBLANK([@v])))'
        if name == '__SET_TOPN' and len(args) == 3 and self.plain_column(args[0]):
            # Tableau TOP/BOTTOM n set ordered by its own field (e.g. the last 4 weeks): IN/OUT per row
            c = self.emit(args[0], True)
            n = args[1][1] if args[1][0] == 'num' else 1
            order = 'ASC' if args[2][0] == 'str' and args[2][1].upper() == 'ASC' else 'DESC'
            self.note('Tableau top-n set -> IN TOPN over the distinct values (computed at refresh: the set does not '
                      'follow other filters, as in Tableau without context filters)', 'medium')
            return f'({c} IN SELECTCOLUMNS(TOPN({n}, ALL({c}), {c}, {order}), "@m", {c}))'
        if name == 'ATTR':
            if self.plain_column(args[0]) or (args[0][0] == 'field' and self.resolve(args[0][1])['kind'] == 'rowcalc'):
                self.note('ATTR -> SELECTEDVALUE (blank when several values, Tableau shows *)', 'medium')
                return f'SELECTEDVALUE({self.emit(args[0], True)})'
            # SELECTEDVALUE takes a column only: ATTR of a measure or an expression is the expression itself
            self.note('ATTR of a measure/expression -> the expression (Tableau shows * when it varies)', 'medium')
            return f'({self.emit(args[0], row)})'
        if name == 'ROUND' and len(args) == 1:  # Tableau ROUND(x): decimals optional; DAX ROUND needs them
            return f'ROUND({E(args[0])}, 0)'
        if name in SAME:
            return f'{name}({", ".join(E(a) for a in args)})'
        if name == 'IIF':
            return f'IF({E(args[0])}, {E(args[1])}, {E(args[2])})'
        if name == 'IFNULL':
            return f'COALESCE({E(args[0])}, {E(args[1])})'
        if name == 'ZN':
            return f'COALESCE({E(args[0])}, 0)'
        if name == 'ISNULL':
            return f'ISBLANK({E(args[0])})'
        if name in ('CEILING', 'FLOOR'):
            return f'{name}({E(args[0])}, {E(args[1]) if len(args) > 1 else 1})'
        if name == 'LOG':
            return f'LOG({E(args[0])}, {E(args[1]) if len(args) > 1 else 10})'
        if name == 'DIV':
            return f'QUOTIENT({E(args[0])}, {E(args[1])})'
        if name == 'SQUARE':
            return f'POWER({E(args[0])}, 2)'
        if name == 'MID':
            s = E(args[0])
            return f'MID({s}, {E(args[1])}, {E(args[2]) if len(args) > 2 else f"LEN({s})"})'
        if name in ('LTRIM', 'RTRIM'):
            self.note(f'{name} -> TRIM (DAX TRIM also removes inner double spaces)', 'medium')
            return f'TRIM({E(args[0])})'
        if name == 'REPLACE':
            return f'SUBSTITUTE({E(args[0])}, {E(args[1])}, {E(args[2])})'
        if name == 'CONTAINS':
            return f'CONTAINSSTRING({E(args[0])}, {E(args[1])})'
        if name == 'STARTSWITH':
            return f'(LEFT({E(args[0])}, LEN({E(args[1])})) = {E(args[1])})'
        if name == 'ENDSWITH':
            return f'(RIGHT({E(args[0])}, LEN({E(args[1])})) = {E(args[1])})'
        if name == 'FIND':  # Tableau FIND(string, substring) -> DAX FIND(find_text, within_text)
            start = E(args[2]) if len(args) > 2 else '1'
            return f'FIND({E(args[1])}, {E(args[0])}, {start}, 0)'
        if name == 'SPACE':
            return f'REPT(" ", {E(args[0])})'
        if name == 'STR':
            return f'CONVERT({E(args[0])}, STRING)'
        if name == 'INT':
            return f'TRUNC({E(args[0])})'
        if name == 'FLOAT':
            return f'CONVERT({E(args[0])}, DOUBLE)'
        if name == 'DATE':
            a = args[0]
            if self.is_string(a):
                return f'DATEVALUE({E(a)})'
            x = E(a)
            return f'DATE(YEAR({x}), MONTH({x}), DAY({x}))'
        if name == 'DATETIME':
            return f'CONVERT({E(args[0])}, DATETIME)'
        if name == 'MAKEDATE':
            return f'DATE({E(args[0])}, {E(args[1])}, {E(args[2])})'
        if name == 'MAKEDATETIME':
            return f'({E(args[0])} + {E(args[1])})'
        if name == 'WEEK':
            return f'WEEKNUM({E(args[0])})'
        if name == 'DATEPART':
            return self.date_part(self.date_unit(args[0]), E(args[1]))
        if name == 'DATENAME':
            u, d = self.date_unit(args[0]), E(args[1])
            return {'MONTH': f'FORMAT({d}, "MMMM")', 'WEEKDAY': f'FORMAT({d}, "dddd")'}.get(
                u, f'FORMAT({self.date_part(u, d)}, "0")')
        if name == 'DATETRUNC':
            u, d = self.date_unit(args[0]), E(args[1])
            if u == 'WEEK':
                k = self.week_start(args[2] if len(args) > 2 else None)
                if k is None:
                    self.note('DATETRUNC week: week start assumed Sunday', 'medium')
                    k = 0
                return f'({d} - WEEKDAY({d} - {k}, 1) + 1)' if k else f'({d} - WEEKDAY({d}, 1) + 1)'
            return {'YEAR': f'DATE(YEAR({d}), 1, 1)', 'QUARTER': f'DATE(YEAR({d}), (QUARTER({d}) - 1) * 3 + 1, 1)',
                    'MONTH': f'DATE(YEAR({d}), MONTH({d}), 1)'}.get(u, f'DATE(YEAR({d}), MONTH({d}), DAY({d}))')
        if name == 'DATEADD':
            u, n, d = self.date_unit(args[0]), E(args[1]), E(args[2])
            mult = {'YEAR': 12, 'QUARTER': 3, 'MONTH': 1}
            if u in mult:
                self.note('DATEADD -> EDATE drops the time part; fine for date comparisons', 'medium')
                return f'EDATE({d}, {n})' if mult[u] == 1 else f'EDATE({d}, ({n}) * {mult[u]})'
            per_day = {'DAY': None, 'WEEK': None, 'HOUR': 24, 'MINUTE': 1440, 'SECOND': 86400}
            if u not in per_day:
                self.note(f'DATEADD unit {u.lower()} not mapped', 'low')
                return f'/* TODO DATEADD {u.lower()} */ BLANK()'
            return f'({d} + ({n}))' if u == 'DAY' else f'({d} + ({n}) * 7)' if u == 'WEEK' else f'({d} + ({n}) / {per_day[u]})'
        if name == 'DATEDIFF':
            u = self.date_unit(args[0])
            # Tableau gives NULL when a date is NULL; DAX counts from 12/30/1899 (a huge number, so "<= 40" flips).
            # Row level only: guard the date columns once, not again in a nested DATEDIFF over the same columns.
            guard = getattr(self, '_dd_guard', set())
            cols = [] if (self.measure_ctx and not row) else list(dict.fromkeys(
                self.emit(x, True) for a in args[1:3] for x in walk(a) if self.plain_column(x)))
            new = [c for c in cols if c not in guard]
            self._dd_guard = guard | set(new)
            a, b = E(args[1]), E(args[2])
            self._dd_guard = guard
            out = f'DATEDIFF({a}, {b}, {u})'
            if u == 'WEEK':
                # DAX counts Sunday week boundaries; Tableau's start_of_week argument moves them. Shifting both dates
                # back k days makes a k-days-after-Sunday boundary land on Sunday ('monday' -> 1).
                k = self.week_start(args[3] if len(args) > 3 else None)
                if k is None:
                    self.note('DATEDIFF week: Tableau uses the data source week start (Sunday assumed)', 'medium')
                elif k:
                    out = f'DATEDIFF({a} - {k}, {b} - {k}, WEEK)'
            return f'IF({" || ".join(f"ISBLANK({c})" for c in new)}, BLANK(), {out})' if new else out
        if name in ('USERNAME', 'FULLNAME', 'USERDOMAIN', 'ISMEMBEROF', 'ISUSERNAME', 'ISFULLNAME'):
            # USERNAME() is a Tableau login; USERPRINCIPALNAME() is an Entra sign-in (usually an email): they only match
            # through a user table. Compared with a column, the value changes per viewer: that is RLS design, not DAX.
            self.note(f'{name}: security function. Do not use this DAX as is: map UPN -> login in a user table and '
                      'implement the logic as an RLS role (per-viewer rows if the result is used as a column)', 'low')
            return 'USERPRINCIPALNAME()' if not args else f'(USERPRINCIPALNAME() = {E(args[0])})'
        if name == 'LOOKUP' and lookup_zero(('fn', name, args)):
            return E(args[0])  # LOOKUP(x, 0): the mark's own value (offset 0), no table calculation
        if name in TABLE_CALCS and getattr(self, 'window', None) is not None:
            out = self.window_fn(name, args)
            if out is not None:
                self.note(f'table calculation {name} over {", ".join(self.window) or "each mark on its own"} '
                          '(Tableau "compute using" of the sheet); check it restarts where Tableau does', 'medium')
                return out
            self._tc_fallback = True
        if name in TABLE_CALCS:
            self.note(f'table calculation {name}: depends on "compute using" (partition/addressing). '
                      f'Rebuild as a visual calculation ({TABLE_CALCS[name]}) or RANKX/WINDOW measure', 'low')
            inner = ', '.join(E(a) for a in args)
            if name.startswith('RANK'):
                return f'RANKX(ALLSELECTED(<Dimension>), {inner})'
            if name == 'TOTAL':
                return f'CALCULATE({inner}, ALLSELECTED())'
            return f'/* visual calc */ {TABLE_CALCS[name]}({inner})'
        if name in ('REGEXP_MATCH', 'REGEXP_EXTRACT', 'REGEXP_EXTRACT_NTH', 'REGEXP_REPLACE', 'SPLIT',
                    'RAWSQL_STR', 'RAWSQL_INT', 'RAWSQL_REAL', 'RAWSQL_BOOL', 'RAWSQL_DATE', 'SCRIPT_REAL',
                    'SCRIPT_STR', 'SCRIPT_INT', 'SCRIPT_BOOL', 'MAKEPOINT', 'MAKELINE', 'HEXBINX', 'HEXBINY'):
            self.note(f'{name} has no DAX equivalent: do it in Power Query or the source', 'low')
            return f'/* TODO {name} */ BLANK()'
        self.note(f'function {name} not mapped', 'low')
        return f'/* TODO */ {name}({", ".join(E(a) for a in args)})'


def format_string(tableau_fmt, datatype=None, name=''):
    """Tableau number format ('n#,##0;-#,##0', 'c"$"#,##0', 'p0.0%', '*+0.0%;-0.0%') -> Power BI formatString.
    Falls back by data type (modeling-guidelines: every visible measure needs a formatString)."""
    if datatype == 'string' or datatype == 'boolean':
        return None  # text measures take no number format
    f = (tableau_fmt or '').strip()
    if f and f[0] in 'ncp*' and len(f) > 1:
        f = f[1:]
    if f in ('S', 's', 'D', 'd'):
        return 'Short Date'
    if f in ('L', 'l'):
        return 'Long Date'
    if f and any(ch in f for ch in '#0'):
        pos = f.split(';')[0]
        return pos if f.count(';') == 0 or f.split(';')[1] == '-' + pos else f
    if '%' in name or (datatype == 'real' and re.search(r'\b(ratio|pct|percent|rate)\b', name, re.I)):
        return '0.0%'
    return {'integer': '#,##0', 'real': '#,##0.00', 'date': 'Short Date', 'datetime': 'General Date'}.get(datatype, '#,##0')


ABBREV = re.compile(r'\b([A-Z]{2,4})\b')


def suggest_measure_name(name, dax):
    """Naming-conventions: business names, '#' prefix for counts, no abbreviations (flagged, never guessed)."""
    out = re.sub(r'\s+', ' ', name).strip()
    if re.match(r'^(DISTINCTCOUNTNOBLANK|DISTINCTCOUNT|COUNTROWS|COUNT)\(', dax or '') and not out.startswith('#'):
        out = '# ' + re.sub(r'\s*\b(count|cnt|countd|#)\b\s*', ' ', out, flags=re.I).strip()
    abbrev = sorted({a for a in ABBREV.findall(out) if a not in ('YTD', 'MTD', 'QTD', 'ID', 'KPI')})
    return out, abbrev


def group_column(table, source_col, groups, default):
    """Tableau group (categorical bin) -> SWITCH calculated column."""
    parts = []
    for label, vals in groups.items():
        vs = ', '.join('"' + v.strip('"').replace('"', '""') + '"' for v in vals)
        parts.append(f'{col(table, source_col)} IN {{{vs}}}, "{label.strip(chr(34))}"')
    other = f'"{default.strip(chr(34))}"' if default else col(table, source_col)
    return 'SWITCH(TRUE(),\n    ' + ',\n    '.join(parts) + f',\n    {other})'


def param_table_name(name):
    """Disconnected table of a Tableau parameter: '<name> Parameter', but a parameter already called
    'DC Parameter' keeps its name (not 'DC Parameter Parameter')."""
    return name if re.search(r'\bparameter\s*$', name or '', re.I) else f'{name} Parameter'


def parameter_pattern(p):
    """Tableau parameter -> Power BI pattern + DAX (table + value measure)."""
    name, dt, dom = p['name'], p.get('datatype'), p.get('domain')
    val = (p.get('value') or '').strip()
    lit = lambda v: (f'DATE({int(v[1:5])}, {int(v[6:8])}, {int(v[9:11])})' if v.startswith('#') and len(v) >= 11
                     else v.replace('"', '"') if v.startswith('"') else v)
    default = lit(val) if val else 'BLANK()'
    tname = param_table_name(name)
    if dt in ('date', 'datetime'):
        lo = lit(p['min']) if p.get('min') else 'DATE(2015, 1, 1)'
        hi = lit(p['max']) if p.get('max') else 'DATE(YEAR(TODAY()) + 1, 12, 31)'
        return {'pattern': 'disconnected date table + single-date slicer (or one Between slicer for a Start/End pair)',
                'table_dax': f"{q(tname)} = CALENDAR({lo}, {hi})",
                'measure_dax': f"[{name} Value] = SELECTEDVALUE({col(tname, 'Date')}, {default})"}
    if p.get('members'):
        rows = ', '.join('{' + (m[0] if not m[0].startswith('"') else m[0]) + (f', "{m[1]}"' if m[1] else '') + '}'
                         for m in p['members'])
        return {'pattern': 'disconnected list table + single-select slicer (field parameter if the values pick measures/fields)',
                'table_dax': f"{q(tname)} = DATATABLE(\"Value\", {'STRING' if dt == 'string' else 'INTEGER' if dt == 'integer' else 'DOUBLE'}"
                             + (', "Label", STRING' if any(m[1] for m in p['members']) else '') + f', {{{rows}}})',
                'measure_dax': f"[{name} Value] = SELECTEDVALUE({col(tname, 'Value')}, {default})"}
    if dt in ('integer', 'real') and dom == 'range':
        return {'pattern': 'numeric range parameter (Modeling > New parameter > Numeric range)',
                'table_dax': f"{q(tname)} = GENERATESERIES({p.get('min', 0)}, {p.get('max', 100)}, {p.get('granularity', 1)})",
                'measure_dax': f"[{name} Value] = SELECTEDVALUE({col(tname, 'Value')}, {default})"}
    return {'pattern': 'free-entry parameter: Power BI has no free text input; use a list slicer or a numeric range',
            'table_dax': None, 'measure_dax': f'[{name} Value] = {default}'}
