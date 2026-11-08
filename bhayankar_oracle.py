"""Bhayankar tester, oracle module: recomputes what a Tableau worksheet shows from source data files.

Independent of Power BI: it evaluates the Tableau formulas themselves (parsed with bhayankar_dax.Parser)
over CSV/XLSX copies of the source tables, applying data source, context, dimension, Top N and measure
filters in Tableau's order, FIXED/INCLUDE/EXCLUDE LODs, groups, blending on linking fields and the
parameter defaults. What it cannot reproduce (table calcs other than percent of total, sets, RLS
functions without a test user, custom SQL without an extract) is reported as not testable, never guessed.
Python 3 standard library only.
"""
import csv
import datetime as dt
import math
import os
import re
import zipfile
import xml.etree.ElementTree as ET

from bhayankar_dax import Parser, ParseError
from bhayankar_extractor import parse_instance, U

QREF = re.compile(r'\[([^\[\]]+)\]\.\[([^\[\]]+)\]')
AGG_FN = {'SUM', 'AVG', 'COUNT', 'COUNTD', 'ATTR', 'MEDIAN', 'STDEV', 'STDEVP', 'VAR', 'VARP', 'PERCENTILE'}
TABLE_FN = re.compile(r'^(RUNNING_\w+|WINDOW_\w+|RANK\w*|INDEX|LOOKUP|PREVIOUS_VALUE|TOTAL|FIRST|LAST|SIZE|SCRIPT_\w+|RAWSQL\w*)$')
USER_FN = {'USERNAME', 'FULLNAME', 'ISMEMBEROF', 'USERDOMAIN', 'ISUSERNAME', 'ISFULLNAME', 'USERATTRIBUTE', 'USERATTRIBUTEINCLUDES'}
DERIV = {'none': 'none', 'sum': 'sum', 'avg': 'avg', 'count': 'cnt', 'countd': 'ctd', 'min': 'min', 'max': 'max',
         'attribute': 'attr', 'user': 'usr', 'year': 'yr', 'quarter': 'qr', 'month': 'mn', 'day': 'dy', 'week': 'wk',
         'weekday': 'wd', 'year-trunc': 'tyr', 'quarter-trunc': 'tqr', 'month-trunc': 'tmn', 'week-trunc': 'twk',
         'day-trunc': 'tdy', 'median': 'med', 'stdev': 'stdev', 'var': 'var', 'hour': 'hr', 'minute': 'mi',
         'mdy': 'mdy', 'my': 'my', 'md': 'md'}
AGG_CODE = {'sum': 'SUM', 'avg': 'AVG', 'cnt': 'COUNT', 'ctd': 'COUNTD', 'cntd': 'COUNTD', 'min': 'MIN', 'max': 'MAX',
            'attr': 'ATTR', 'med': 'MEDIAN', 'stdev': 'STDEV', 'var': 'VAR'}
DATE_CODE = {'yr', 'qr', 'mn', 'dy', 'wk', 'wd', 'tyr', 'tqr', 'tmn', 'twk', 'tdy', 'hr', 'mi', 'mdy', 'my', 'md'}
QUICK = {'pcto': 'PERCENT_OF_TOTAL', 'pcdf': 'PERCENT_DIFFERENCE', 'diff': 'DIFFERENCE', 'cum': 'RUNNING_TOTAL',
         'rank': 'RANK', 'rank_dense': 'RANK', 'rank_modified': 'RANK', 'rank_unique': 'RANK', 'pctrank': 'PERCENTILE',
         'mavg': 'MOVING_AVG', 'ytd': 'YTD_TOTAL', 'ytdgr': 'YTD_GROWTH', 'cgr': 'COMPOUND_GROWTH', 'win': 'WINDOW'}
MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October',
          'November', 'December']
DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
DATE_FORMATS = ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M:%S.%f', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%dT%H:%M:%S.%f',
                '%Y-%m-%d %H:%M', '%Y-%m-%d', '%m/%d/%Y %H:%M:%S', '%m/%d/%Y %I:%M:%S %p', '%m/%d/%Y %H:%M',
                '%m/%d/%Y %I:%M %p', '%m/%d/%Y', '%d-%b-%Y', '%d-%b-%y', '%Y/%m/%d', '%d.%m.%Y')


class Unsupported(Exception):
    """Raised for logic the oracle does not reproduce; the message says what and why."""


class Star:
    """Tableau's '*' from ATTR() over several values."""
    def __repr__(self):
        return '*'


STAR = Star()


def norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


# ---------- reading data files ----------
def read_table(path):
    """CSV / TSV (any encoding Tableau or Excel writes) or the first sheet of an .xlsx -> (headers, rows)."""
    if path.lower().endswith(('.xlsx', '.xlsm')):
        return read_xlsx(path)
    with open(path, 'rb') as f:
        raw = f.read()
    for enc in (('utf-16',) if raw[:2] in (b'\xff\xfe', b'\xfe\xff') else ('utf-8-sig', 'cp1252', 'latin-1')):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    head = text[:4096]
    delim = max(('\t', ',', ';', '|'), key=lambda d: head.split('\n')[0].count(d))
    rows = list(csv.reader(text.splitlines(), delimiter=delim))
    rows = [r for r in rows if any(c.strip() for c in r)]
    if not rows:
        return [], []
    return [h.strip() for h in rows[0]], rows[1:]


def read_xlsx(path, sheet=None):
    ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si', ns):
                shared.append(''.join(t.text or '' for t in si.iter('{%s}t' % ns['m'])))
        sheets = sorted(n for n in z.namelist() if re.match(r'xl/worksheets/sheet\d+\.xml$', n))
        if not sheets:
            return [], []
        root = ET.fromstring(z.read(sheet or sheets[0]))
    out = []
    for r in root.iter('{%s}row' % ns['m']):
        cells = {}
        for c in r.findall('m:c', ns):
            ref = re.match(r'([A-Z]+)', c.get('r', 'A')).group(1)
            idx = 0
            for ch in ref:
                idx = idx * 26 + ord(ch) - 64
            v = c.find('m:v', ns)
            t = c.get('t')
            if t == 's' and v is not None:
                val = shared[int(v.text)]
            elif t == 'inlineStr':
                val = ''.join(x.text or '' for x in c.iter('{%s}t' % ns['m']))
            else:
                val = v.text if v is not None else ''
            cells[idx - 1] = val
        if cells:
            out.append([cells.get(i, '') for i in range(max(cells) + 1)])
    out = [r for r in out if any(str(c).strip() for c in r)]
    if not out:
        return [], []
    width = max(len(r) for r in out)
    out = [r + [''] * (width - len(r)) for r in out]
    return [str(h).strip() for h in out[0]], out[1:]


def parse_date(s):
    if isinstance(s, (dt.date, dt.datetime)):
        return s
    s = str(s).strip()
    if not s:
        return None
    for fmt in DATE_FORMATS:
        try:
            return dt.datetime.strptime(s, fmt)
        except ValueError:
            continue
    if re.match(r'^\d{5}(\.\d+)?$', s):  # Excel serial date
        return dt.datetime(1899, 12, 30) + dt.timedelta(days=float(s))
    return None


def tableau_literal(v):
    """Tableau member / parameter literal: '"Y"' -> 'Y', '#2026-01-01#' -> date, '2025' -> 2025, '%null%' -> None."""
    if v is None:
        return None
    s = str(v).strip()
    if s in ('%null%', 'Null', 'NULL'):
        return None
    if len(s) >= 2 and s[0] == s[-1] and s[0] in '"\'':
        return s[1:-1].replace(s[0] * 2, s[0])
    if s.startswith('#') and s.endswith('#'):
        d = parse_date(s.strip('#'))
        return d.date() if d and d.time() == dt.time() and len(s) <= 12 else d
    if s.lower() in ('true', 'false'):
        return s.lower() == 'true'
    try:
        f = float(s)
        return int(f) if f.is_integer() and '.' not in s and 'e' not in s.lower() else f
    except ValueError:
        return s


def as_dt(v):
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, dt.date):
        return dt.datetime(v.year, v.month, v.day)
    return v


def comparable(a, b):
    """Make two values comparable the way Tableau does (date vs datetime, bool vs number, number vs numeric text)."""
    if isinstance(a, (dt.date, dt.datetime)) or isinstance(b, (dt.date, dt.datetime)):
        a = as_dt(parse_date(a) if isinstance(a, str) else a)
        b = as_dt(parse_date(b) if isinstance(b, str) else b)
    elif isinstance(a, bool) and isinstance(b, (int, float)) and not isinstance(b, bool):
        a = int(a)
    elif isinstance(b, bool) and isinstance(a, (int, float)) and not isinstance(a, bool):
        b = int(b)
    elif isinstance(a, str) and isinstance(b, (int, float)) and not isinstance(b, bool):
        try:
            a = float(a)
        except ValueError:
            b = str(b)
    elif isinstance(b, str) and isinstance(a, (int, float)) and not isinstance(a, bool):
        try:
            b = float(b)
        except ValueError:
            a = str(a)
    return a, b


def to_str(v):
    if v is None:
        return None
    if isinstance(v, bool):
        return 'true' if v else 'false'
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, dt.datetime):
        return v.strftime('%Y-%m-%d %H:%M:%S') if v.time() != dt.time() else v.strftime('%Y-%m-%d')
    if isinstance(v, dt.date):
        return v.isoformat()
    return str(v)


def round_half_away(x, n=0):
    m = 10 ** n
    return math.floor(abs(x) * m + 0.5) / m * (1 if x >= 0 else -1)


# ---------- dates ----------
def trunc(unit, d):
    d = as_dt(d)
    if d is None:
        return None
    unit = unit.lower()
    if unit == 'year':
        return dt.datetime(d.year, 1, 1)
    if unit == 'quarter':
        return dt.datetime(d.year, 3 * ((d.month - 1) // 3) + 1, 1)
    if unit == 'month':
        return dt.datetime(d.year, d.month, 1)
    if unit == 'week':
        x = d - dt.timedelta(days=(d.weekday() + 1) % 7)  # weeks start on Sunday (Tableau default)
        return dt.datetime(x.year, x.month, x.day)
    if unit == 'day':
        return dt.datetime(d.year, d.month, d.day)
    if unit == 'hour':
        return d.replace(minute=0, second=0, microsecond=0)
    if unit == 'minute':
        return d.replace(second=0, microsecond=0)
    return d


def add_months(d, n):
    y, m = divmod(d.month - 1 + n, 12)
    y += d.year
    m += 1
    last = [31, 29 if y % 4 == 0 and (y % 100 or y % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]
    return d.replace(year=y, month=m, day=min(d.day, last))


def date_add(unit, n, d):
    d = as_dt(d)
    if d is None or n is None:
        return None
    n = int(n)
    unit = unit.lower()
    if unit == 'year':
        return add_months(d, 12 * n)
    if unit == 'quarter':
        return add_months(d, 3 * n)
    if unit == 'month':
        return add_months(d, n)
    if unit == 'week':
        return d + dt.timedelta(weeks=n)
    if unit in ('day', 'weekday', 'dayofyear'):
        return d + dt.timedelta(days=n)
    if unit == 'hour':
        return d + dt.timedelta(hours=n)
    if unit == 'minute':
        return d + dt.timedelta(minutes=n)
    return d + dt.timedelta(seconds=n)


def date_part(unit, d):
    d = as_dt(d)
    if d is None:
        return None
    unit = unit.lower()
    if unit == 'year':
        return d.year
    if unit == 'quarter':
        return (d.month - 1) // 3 + 1
    if unit == 'month':
        return d.month
    if unit == 'day':
        return d.day
    if unit == 'weekday':
        return (d.weekday() + 1) % 7 + 1  # Sunday = 1
    if unit == 'week':
        jan1 = dt.datetime(d.year, 1, 1)
        start = jan1 - dt.timedelta(days=(jan1.weekday() + 1) % 7)
        return (d - start).days // 7 + 1
    if unit == 'dayofyear':
        return d.timetuple().tm_yday
    if unit == 'hour':
        return d.hour
    if unit == 'minute':
        return d.minute
    if unit == 'second':
        return d.second
    if unit == 'iso-year':
        return d.isocalendar()[0]
    if unit == 'iso-week':
        return d.isocalendar()[1]
    raise Unsupported(f'DATEPART unit {unit}')


def date_diff(unit, a, b):
    a, b = as_dt(a), as_dt(b)
    if a is None or b is None:
        return None
    unit = unit.lower()
    if unit == 'year':
        return b.year - a.year
    if unit == 'quarter':
        return (b.year - a.year) * 4 + ((b.month - 1) // 3 - (a.month - 1) // 3)
    if unit == 'month':
        return (b.year - a.year) * 12 + b.month - a.month
    if unit == 'week':
        return (trunc('week', b) - trunc('week', a)).days // 7
    if unit in ('day', 'weekday', 'dayofyear'):
        return (trunc('day', b) - trunc('day', a)).days
    if unit == 'hour':
        return int((trunc('hour', b) - trunc('hour', a)).total_seconds() // 3600)
    if unit == 'minute':
        return int((trunc('minute', b) - trunc('minute', a)).total_seconds() // 60)
    return int((b - a).total_seconds())


def derive_date(code, v):
    """Shelf date derivation ('yr', 'tmn', ...) applied to a value."""
    if v is None:
        return None
    if not isinstance(v, (dt.date, dt.datetime)):
        v = parse_date(v)
        if v is None:
            return None
    return {'yr': lambda: date_part('year', v), 'qr': lambda: date_part('quarter', v), 'mn': lambda: date_part('month', v),
            'dy': lambda: date_part('day', v), 'wk': lambda: date_part('week', v), 'wd': lambda: date_part('weekday', v),
            'hr': lambda: date_part('hour', v), 'mi': lambda: date_part('minute', v),
            'tyr': lambda: trunc('year', v), 'tqr': lambda: trunc('quarter', v), 'tmn': lambda: trunc('month', v),
            'twk': lambda: trunc('week', v), 'tdy': lambda: trunc('day', v), 'mdy': lambda: trunc('day', v),
            'my': lambda: trunc('month', v), 'md': lambda: (v.month, v.day)}[code]()


# ---------- scopes ----------
class RowScope:
    __slots__ = ('row',)

    def __init__(self, row):
        self.row = row


class GroupScope:
    __slots__ = ('rows',)

    def __init__(self, rows):
        self.rows = rows


class Ctx:
    """Per-worksheet evaluation context (what LODs need to know)."""
    def __init__(self, lod_rows=None, view_dims=None, view_rows=None):
        self.lod_rows = lod_rows or []
        self.view_dims = view_dims or []
        self.view_rows = view_rows or []
        self.cache = {}


class Oracle:
    def __init__(self, wb, data_dirs=(), today=None, username=None, keep_empty=(), table_files=None):
        self.wb = wb
        self.today = today or dt.date.today()
        self.username = username
        self.keep_empty = {norm(x) for x in keep_empty}
        self.files = {}
        for d in data_dirs:
            for dp, _, fs in os.walk(d):
                for fn in fs:
                    if fn.lower().endswith(('.csv', '.txt', '.tsv', '.xlsx')) and not fn.startswith('~$'):
                        self.files.setdefault(norm(os.path.splitext(fn)[0]), os.path.join(dp, fn))
        self.table_files = {norm(k): v for k, v in (table_files or {}).items()}
        self._ast, self._agg, self._rows, self._tables, self._access = {}, {}, {}, {}, {}
        self.notes = []
        self.loaded = {}  # file -> row count, for the report

    # ----- data -----
    def find_file(self, *names):
        for n in names:
            k = norm(n)
            if not k:
                continue
            if k in self.table_files and os.path.exists(self.table_files[k]):
                return self.table_files[k]
            if k in self.files:
                return self.files[k]
        return None

    def load_file(self, path):
        if path not in self._tables:
            headers, rows = read_table(path)
            idx = {}
            for i, h in enumerate(headers):
                idx.setdefault(h.upper(), i)
                idx.setdefault(norm(h), i)
            self._tables[path] = (headers, rows, idx)
            self.loaded[path] = len(rows)
        return self._tables[path]

    @staticmethod
    def leaves(rel):
        if not rel:
            return []
        if 'join' in rel:
            return [x for r in rel['inputs'] for x in Oracle.leaves(r)]
        if 'union' in rel:
            return [x for r in rel['union'] for x in Oracle.leaves(r)]
        return [rel]

    def leaf_rows(self, leaf, ds_caption):
        tname = (leaf.get('table') or '').replace('[', '').replace(']', '')
        last = tname.split('.')[-1] if tname else ''
        path = self.find_file(leaf.get('alias'), last, tname, ds_caption if leaf.get('custom_sql') else None)
        if not path:
            what = 'custom SQL "' + (leaf.get('alias') or '') + '"' if leaf.get('custom_sql') else (tname or leaf.get('alias'))
            raise Unsupported(f'no data file for {what}: put {last or leaf.get("alias")}.csv in the data folder')
        headers, rows, _ = self.load_file(path)
        alias = (leaf.get('alias') or last).upper()
        hu = [h.upper() for h in headers]
        return [{**{(alias, h): r[i] if i < len(r) else '' for i, h in enumerate(hu)},
                 **{h: r[i] if i < len(r) else '' for i, h in enumerate(hu)}} for r in rows]

    def join(self, rel, ds_caption):
        if 'join' not in rel:
            if 'union' in rel:
                return [r for x in rel['union'] for r in self.join(x, ds_caption)]
            return self.leaf_rows(rel, ds_caption)
        left = self.join(rel['inputs'][0], ds_caption)
        for right_rel, on in zip(rel['inputs'][1:], rel['on'] or [None]):
            right = self.join(right_rel, ds_caption)
            pairs = re.findall(r'\[([^\]]+)\]\.\[([^\]]+)\]\s*=\s*\[([^\]]+)\]\.\[([^\]]+)\]', on or '')
            if not pairs:
                raise Unsupported(f'join condition not understood: {on}')
            lk = [(a.upper(), b.upper()) for a, b, _, _ in pairs]
            rk = [(c.upper(), d.upper()) for _, _, c, d in pairs]
            if right and lk[0] in right[0] and not (left and lk[0] in left[0]):
                lk, rk = rk, lk
            index = {}
            for r in right:
                index.setdefault(tuple(r.get(k, '') for k in rk), []).append(r)
            kind = (rel.get('join') or 'inner').lower()
            out, hit = [], set()
            for l in left:
                m = index.get(tuple(l.get(k, '') for k in lk), [])
                for r in m:
                    hit.add(id(r))
                    out.append({**r, **l})
                if not m and kind in ('left', 'full'):
                    out.append(dict(l))
            if kind in ('right', 'full'):
                out.extend(dict(r) for r in right if id(r) not in hit)
            left = out
        return left

    def ds_rows(self, ds):
        """Source rows of one Tableau data source (raw strings), joined as the data source defines."""
        if ds in self._rows:
            if isinstance(self._rows[ds], Exception):
                raise self._rows[ds]
            return self._rows[ds]
        try:
            info = self.wb.ds[ds]
            conn = info['connection'] or {}
            conn = conn.get('tds') or conn
            flat = self.find_file(info['caption'], info['caption'].split(' (')[0], ds)
            if flat:
                headers, rows, _ = self.load_file(flat)
                hu = [h.upper() for h in headers]
                out = [{h: r[i] if i < len(r) else '' for i, h in enumerate(hu)} for r in rows]
            elif conn.get('logical') and len(conn['logical'].get('objects', [])) > 1:
                objs = conn['logical']['objects']
                phys = conn.get('physical_by_object') or {}
                out = self.join(phys[objs[0]], info['caption'])
                done = {objs[0]}
                for rl in conn['logical'].get('relationships', []):
                    a, b = (rl['between'] + [None, None])[:2]
                    new = b if a in done else a if b in done else None
                    if not new or new in done or new not in phys:
                        continue
                    out = self._join_rows(out, self.join(phys[new], info['caption']), rl['on'] or '')
                    done.add(new)
                self.notes.append(f'{info["caption"]}: Tableau relationships evaluated as left joins from {objs[0]}; '
                                  'measures from the many side can differ from Tableau')
            elif conn.get('relations'):
                out = self.join(conn['relations'], info['caption'])
            else:
                raise Unsupported(f'{info["caption"]}: no table definition; put a flattened "{info["caption"]}.csv" in the data folder')
        except Unsupported as e:
            self._rows[ds] = e
            raise
        self._rows[ds] = out
        return out

    def _join_rows(self, left, right, on):
        pairs = re.findall(r'\[([^\]]+?)(?: \([^)]*\))?\]', on)
        if len(pairs) != 2:
            raise Unsupported(f'relationship not understood: {on}')
        a, b = pairs[0].upper(), pairs[1].upper()
        index = {}
        for r in right:
            index.setdefault(r.get(b, r.get(a, '')), []).append(r)
        out = []
        for l in left:
            m = index.get(l.get(a, l.get(b, '')), [])
            out.extend({**r, **l} for r in m) if m else out.append(dict(l))
        return out

    # ----- columns -----
    def accessor(self, key):
        """(ds, col) source column -> function(raw row) -> typed value."""
        if key in self._access:
            return self._access[key]
        c = self.wb.cols[key]
        rem = c.get('remote')
        cands = []
        if rem:
            tbl = (rem[1] or (rem[0] or '').strip('[]') or '').upper()
            if rem[2]:
                cands += [(tbl, rem[2].upper()), rem[2].upper()]
        for n in (key[1], c.get('caption')):
            if n:
                cands += [n.upper()]
        dtype = c.get('datatype')
        keep = norm(rem[2] if rem and rem[2] else key[1]) in self.keep_empty

        def pick(row):
            for k in cands:
                if k in row:
                    return k
            nk = {norm(x if isinstance(x, str) else x[1]): x for x in row}
            for k in cands:
                kk = norm(k if isinstance(k, str) else k[1])
                if kk in nk:
                    return nk[kk]
            return None
        state = {'k': None, 'fmt': None}

        def get(row):
            if state['k'] is None:
                state['k'] = pick(row) or ''
                if not state['k']:
                    raise Unsupported(f'column {c.get("caption") or key[1]} ({rem[2] if rem else "?"}) not in the data file')
            v = row.get(state['k'], '')
            if v is None or (v == '' and not keep):
                return None
            if dtype in ('integer', 'real'):
                try:
                    x = float(str(v).replace(',', ''))
                except ValueError:
                    return None
                return int(x) if dtype == 'integer' and x.is_integer() else x
            if dtype in ('date', 'datetime'):
                d = parse_date(v)
                return (d.date() if dtype == 'date' and isinstance(d, dt.datetime) else d)
            if dtype == 'boolean':
                s = str(v).strip().lower()
                return True if s in ('true', '1', 't', 'yes', 'y') else False if s in ('false', '0', 'f', 'no', 'n') else None
            return v
        self._access[key] = get
        return get

    # ----- formulas -----
    def ast(self, key):
        if key not in self._ast:
            c = self.wb.cols[key]
            if c['bins']:
                self._ast[key] = ('bins', key)
            elif c['formula'] is None:
                self._ast[key] = ('src', key)
            else:
                try:
                    tree = Parser(c['formula']).parse()
                except ParseError as e:
                    raise Unsupported(f'formula of {c["caption"] or key[1]} not parsed: {e}')
                self._ast[key] = self.resolve(tree, key[0])
        return self._ast[key]

    def resolve(self, n, ds):
        if not isinstance(n, tuple):
            return n
        if n[0] == 'field':
            k = self.wb.lookup(ds, n[1])
            return ('key', k) if k else ('unknown', n[1])
        if n[0] == 'if':
            return ('if', [(self.resolve(c, ds), self.resolve(t, ds)) for c, t in n[1]], self.resolve(n[2], ds))
        if n[0] == 'case':
            return ('case', self.resolve(n[1], ds), [([self.resolve(w, ds) for w in ws], self.resolve(t, ds)) for ws, t in n[2]],
                    self.resolve(n[3], ds))
        if n[0] == 'in':
            return ('in', self.resolve(n[1], ds), [self.resolve(v, ds) for v in n[2]])
        if n[0] == 'fn':
            return ('fn', n[1], [self.resolve(a, ds) for a in n[2]])
        if n[0] == 'lod':
            return ('lod', n[1], [self.resolve(d, ds) for d in n[2]], self.resolve(n[3], ds))
        return tuple(self.resolve(x, ds) if isinstance(x, tuple) else x for x in n)

    def is_agg(self, key, seen=None):
        """True when a calc aggregates (needs a group), False when it is row level."""
        if key in self._agg:
            return self._agg[key]
        seen = seen or set()
        if key in seen:
            return False
        seen.add(key)
        res = self._node_agg(self.ast(key), seen)
        self._agg[key] = res
        return res

    def _node_agg(self, n, seen):
        if not isinstance(n, tuple):
            return False
        t = n[0]
        if t == 'lod':
            return False
        if t == 'fn' and (n[1] in AGG_FN or (n[1] in ('MIN', 'MAX') and len(n[2]) == 1)):
            return True
        if t == 'fn' and TABLE_FN.match(n[1]):
            return True
        if t == 'key':
            k = n[1]
            if k[0] == 'Parameters' or self.wb.cols[k]['formula'] is None:
                return False
            return self.is_agg(k, seen)
        for x in n[1:]:
            if isinstance(x, tuple) and self._node_agg(x, seen):
                return True
            if isinstance(x, list):
                for y in x:
                    if isinstance(y, tuple) and y and isinstance(y[0], str) and self._node_agg(y, seen):
                        return True
                    if isinstance(y, tuple) and y and not isinstance(y[0], str):
                        for z in y:
                            zz = z if isinstance(z, list) else [z]
                            if any(isinstance(q, tuple) and self._node_agg(q, seen) for q in zz):
                                return True
        return False

    def param(self, key):
        c = self.wb.cols[key]
        return tableau_literal((c.get('param') or {}).get('value') or c.get('formula'))

    def field_value(self, key, scope, ctx):
        if key[0] == 'Parameters':
            return self.param(key)
        c = self.wb.cols[key]
        if isinstance(scope, GroupScope):
            if c['formula'] is not None and self.is_agg(key):
                return self.ev(self.ast(key), scope, ctx)
            return self.attr([self.field_value(key, RowScope(r), ctx) for r in scope.rows])
        row = scope.row
        if key in row:
            return row[key]
        per_sheet = c['formula'] is not None and self.has_lod(key)  # LOD results depend on the sheet's filters
        if per_sheet and (id(row), key) in ctx.cache:
            return ctx.cache[(id(row), key)]
        if c['bins']:
            src = self.field_value((key[0], c['bins']['source']), scope, ctx)
            if c['bins'].get('kind') == 'numeric-bin':
                size = tableau_literal(c['bins'].get('size')) or 1
                v = None if src is None else math.floor(src / size) * size
            else:
                v = None
                for label, vals in c['bins']['groups'].items():
                    if to_str(src) in [to_str(tableau_literal(x)) for x in vals]:
                        v = tableau_literal(label)
                        break
                else:
                    d = c['bins'].get('default')
                    v = tableau_literal(d) if d else src
        elif c['formula'] is None:
            if key[1] in ('Number of Records', '__tableau_internal_object_id__'):
                v = 1
            else:
                v = self.accessor(key)(row)
        else:
            if self.is_agg(key):
                raise Unsupported(f'aggregate calc {c["caption"] or key[1]} used at row level')
            v = self.ev(self.ast(key), scope, ctx)
        if per_sheet:
            ctx.cache[(id(row), key)] = v
        else:
            row[key] = v
        return v

    @staticmethod
    def attr(vals):
        vs = {to_str(v) for v in vals if v is not None}
        if not vs:
            return None
        if len(vs) == 1:
            return next(v for v in vals if v is not None)
        return STAR

    def aggregate(self, fn, vals, extra=None):
        if fn == 'ATTR':
            return self.attr(vals)
        if fn == 'COUNT':
            return sum(1 for v in vals if v is not None)
        if fn == 'COUNTD':
            return len({to_str(v) for v in vals if v is not None})
        nums = [v for v in vals if v is not None and not isinstance(v, Star)]
        if fn in ('MIN', 'MAX'):
            if not nums:
                return None
            try:
                return min(nums) if fn == 'MIN' else max(nums)
            except TypeError:
                return (min if fn == 'MIN' else max)(nums, key=lambda x: str(x))
        nums = [float(v) if isinstance(v, bool) else v for v in nums]
        nums = [v for v in nums if isinstance(v, (int, float))]
        if not nums:
            return None
        if fn == 'SUM':
            s = sum(nums)
            return s
        if fn == 'AVG':
            return sum(nums) / len(nums)
        if fn == 'MEDIAN':
            s = sorted(nums)
            m = len(s) // 2
            return s[m] if len(s) % 2 else (s[m - 1] + s[m]) / 2
        if fn in ('STDEV', 'VAR', 'STDEVP', 'VARP'):
            n = len(nums)
            if n < (1 if fn.endswith('P') else 2):
                return None
            mu = sum(nums) / n
            var = sum((x - mu) ** 2 for x in nums) / (n if fn.endswith('P') else n - 1)
            return math.sqrt(var) if fn.startswith('STDEV') else var
        if fn == 'PERCENTILE':
            s = sorted(nums)
            p = float(extra or 0)
            k = (len(s) - 1) * p
            f = math.floor(k)
            return s[f] if f == k else s[f] + (s[min(f + 1, len(s) - 1)] - s[f]) * (k - f)
        raise Unsupported(f'aggregation {fn}')

    def lod(self, n, scope, ctx):
        _, kind, dims, expr = n
        if isinstance(scope, GroupScope):
            return self.attr([self.lod(n, RowScope(r), ctx) for r in scope.rows])
        ck = ('lod', id(n))
        if ck not in ctx.cache:
            if kind == 'FIXED':
                rows = ctx.lod_rows
                keyf = lambda r: tuple(to_str(self.ev(d, RowScope(r), ctx)) for d in dims)
            else:
                if ctx.view_dims is None:
                    raise Unsupported(f'{kind} LOD outside a view')
                extra = [d for d in dims]
                if kind == 'INCLUDE':
                    keyf = lambda r: tuple(to_str(f(r)) for f in ctx.view_dims) + \
                        tuple(to_str(self.ev(d, RowScope(r), ctx)) for d in extra)
                else:
                    excl = {d[1] for d in dims if isinstance(d, tuple) and d[0] == 'key'}
                    keep = [f for f in ctx.view_dims if getattr(f, 'key', None) not in excl]
                    keyf = lambda r, keep=keep: tuple(to_str(f(r)) for f in keep)
                rows = ctx.view_rows
            groups = {}
            for r in rows:
                groups.setdefault(keyf(r), []).append(r)
            ctx.cache[ck] = ({k: self.ev(expr, GroupScope(g), ctx) for k, g in groups.items()}, keyf)
        table, keyf = ctx.cache[ck]
        return table.get(keyf(scope.row))

    def ev(self, n, scope, ctx):
        t = n[0]
        if t == 'num':
            v = float(n[1])
            return int(v) if v.is_integer() and '.' not in n[1] else v
        if t == 'str':
            return n[1]
        if t == 'bool':
            return n[1] == 'TRUE'
        if t == 'null':
            return None
        if t == 'date':
            d = parse_date(n[1])
            return d.date() if d and d.time() == dt.time() and len(n[1]) <= 10 else d
        if t == 'key':
            return self.field_value(n[1], scope, ctx)
        if t in ('src', 'bins'):
            return self.field_value(n[1], scope, ctx)
        if t == 'unknown':
            raise Unsupported(f'unknown field {n[1]}')
        if t == 'lod':
            return self.lod(n, scope, ctx)
        if t == 'neg':
            v = self.ev(n[1], scope, ctx)
            return None if v is None else -v
        if t == 'not':
            v = self.ev(n[1], scope, ctx)
            return None if v is None or isinstance(v, Star) else not v
        if t == 'if':
            for c, then in n[1]:
                cv = self.ev(c, scope, ctx)
                if cv is True or (cv not in (None, False) and not isinstance(cv, Star) and cv != 0):
                    return self.ev(then, scope, ctx)
            return self.ev(n[2], scope, ctx) if n[2] is not None else None
        if t == 'case':
            subj = self.ev(n[1], scope, ctx)
            if subj is not None and not isinstance(subj, Star):
                for whens, then in n[2]:
                    for w in whens:
                        a, b = comparable(subj, self.ev(w, scope, ctx))
                        if a == b:
                            return self.ev(then, scope, ctx)
            return self.ev(n[3], scope, ctx) if n[3] is not None else None
        if t == 'in':
            v = self.ev(n[1], scope, ctx)
            if v is None:
                return None
            for x in n[2]:
                a, b = comparable(v, self.ev(x, scope, ctx))
                if a == b:
                    return True
            return False
        if t == 'bin':
            return self.binop(n[1], n[2], n[3], scope, ctx)
        if t == 'fn':
            return self.fn(n[1], n[2], scope, ctx)
        raise Unsupported(f'expression {t}')

    def binop(self, op, an, bn, scope, ctx):
        if op in ('&&', '||'):
            a = self.ev(an, scope, ctx)
            if op == '&&' and a is False:
                return False
            if op == '||' and a is True:
                return True
            b = self.ev(bn, scope, ctx)
            vals = [a, b]
            if op == '&&':
                return False if False in vals else None if any(v is None or isinstance(v, Star) for v in vals) else bool(a and b)
            return True if True in vals else None if any(v is None or isinstance(v, Star) for v in vals) else bool(a or b)
        a, b = self.ev(an, scope, ctx), self.ev(bn, scope, ctx)
        if a is None or b is None or isinstance(a, Star) or isinstance(b, Star):
            return None
        if op in ('=', '<>', '<', '<=', '>', '>='):
            a, b = comparable(a, b)
            try:
                return {'=': a == b, '<>': a != b, '<': a < b, '<=': a <= b, '>': a > b, '>=': a >= b}[op]
            except TypeError:
                a, b = str(a), str(b)
                return {'=': a == b, '<>': a != b, '<': a < b, '<=': a <= b, '>': a > b, '>=': a >= b}[op]
        if op == '+':
            if isinstance(a, str) or isinstance(b, str):
                return str(to_str(a)) + str(to_str(b))
            if isinstance(a, (dt.date, dt.datetime)):
                return as_dt(a) + dt.timedelta(days=float(b))
            return a + b
        if op == '-':
            if isinstance(a, (dt.date, dt.datetime)) and isinstance(b, (dt.date, dt.datetime)):
                return (as_dt(a) - as_dt(b)).total_seconds() / 86400
            if isinstance(a, (dt.date, dt.datetime)):
                return as_dt(a) - dt.timedelta(days=float(b))
            return a - b
        if op == '*':
            return a * b
        if op == '/':
            return None if b == 0 else a / b
        if op == '%':
            return None if b == 0 else math.fmod(a, b)
        if op == '^':
            return a ** b
        raise Unsupported(f'operator {op}')

    def fn(self, name, args, scope, ctx):
        if name in USER_FN:
            if self.username is None:
                raise Unsupported(f'{name}() depends on the signed-in user: set data.username in test_map.json to test as one user')
            if name in ('USERNAME', 'FULLNAME', 'USERDOMAIN'):
                return self.username
            if name in ('ISUSERNAME', 'ISFULLNAME'):
                v = self.ev(args[0], scope, ctx)
                return None if v is None else str(v).lower() == str(self.username).lower()
            raise Unsupported(f'{name}() is not reproduced')
        if TABLE_FN.match(name):
            raise Unsupported(f'table calculation {name}() is not reproduced')
        if name in AGG_FN or (name in ('MIN', 'MAX') and len(args) == 1):
            rows = scope.rows if isinstance(scope, GroupScope) else [scope.row]
            vals = [self.ev(args[0], RowScope(r), ctx) for r in rows]
            return self.aggregate(name, vals, self.ev(args[1], scope, ctx) if name == 'PERCENTILE' and len(args) > 1 else None)
        if name == 'IIF':
            c = self.ev(args[0], scope, ctx)
            if c is None and len(args) > 3:
                return self.ev(args[3], scope, ctx)
            return self.ev(args[1] if c else args[2], scope, ctx) if c is not None else None
        if name in ('IFNULL', 'ZN'):
            v = self.ev(args[0], scope, ctx)
            if v is None:
                return 0 if name == 'ZN' else self.ev(args[1], scope, ctx)
            return v
        if name == 'ISNULL':
            return self.ev(args[0], scope, ctx) is None
        a = [self.ev(x, scope, ctx) for x in args]
        if name == 'TODAY':
            return self.today
        if name == 'NOW':
            return dt.datetime.combine(self.today, dt.datetime.now().time())
        if name in ('MIN', 'MAX'):
            if any(v is None for v in a):
                return None
            x, y = comparable(a[0], a[1])
            return (min if name == 'MIN' else max)(x, y)
        if name in ('DATEPART', 'DATENAME', 'DATETRUNC', 'DATEADD', 'DATEDIFF'):
            unit = str(a[0]).lower()
            if name == 'DATEPART':
                return date_part(unit, a[1])
            if name == 'DATENAME':
                if a[1] is None:
                    return None
                d = as_dt(a[1])
                return MONTHS[d.month - 1] if unit == 'month' else DAYS[d.weekday()] if unit == 'weekday' \
                    else f'Q{date_part("quarter", d)}' if unit == 'quarter' else str(date_part(unit, d))
            if name == 'DATETRUNC':
                return trunc(unit, a[1])
            if name == 'DATEADD':
                return date_add(unit, a[1], a[2])
            return date_diff(unit, a[1], a[2])
        if any(v is None for v in a) and name not in ('MAKEDATE',):
            return None
        s = lambda v: to_str(v)
        f = {
            'ABS': lambda: abs(a[0]), 'ROUND': lambda: round_half_away(a[0], int(a[1]) if len(a) > 1 else 0),
            'CEILING': lambda: math.ceil(a[0]), 'FLOOR': lambda: math.floor(a[0]), 'SQRT': lambda: math.sqrt(a[0]),
            'POWER': lambda: a[0] ** a[1], 'EXP': lambda: math.exp(a[0]), 'LN': lambda: math.log(a[0]),
            'LOG': lambda: math.log(a[0], a[1] if len(a) > 1 else 10), 'SIGN': lambda: (a[0] > 0) - (a[0] < 0),
            'INT': lambda: int(float(a[0])) if not isinstance(a[0], str) else int(float(a[0])),
            'FLOAT': lambda: float(a[0]), 'DIV': lambda: int(a[0] / a[1]) if a[1] else None,
            'STR': lambda: s(a[0]), 'LEN': lambda: len(s(a[0])), 'UPPER': lambda: s(a[0]).upper(),
            'LOWER': lambda: s(a[0]).lower(), 'TRIM': lambda: s(a[0]).strip(), 'LTRIM': lambda: s(a[0]).lstrip(),
            'RTRIM': lambda: s(a[0]).rstrip(), 'LEFT': lambda: s(a[0])[:int(a[1])],
            'RIGHT': lambda: s(a[0])[-int(a[1]):] if int(a[1]) else '',
            'MID': lambda: s(a[0])[int(a[1]) - 1:(int(a[1]) - 1 + int(a[2])) if len(a) > 2 else None],
            'CONTAINS': lambda: s(a[1]) in s(a[0]), 'STARTSWITH': lambda: s(a[0]).startswith(s(a[1])),
            'ENDSWITH': lambda: s(a[0]).endswith(s(a[1])),
            'FIND': lambda: s(a[0]).find(s(a[1]), int(a[2]) - 1 if len(a) > 2 else 0) + 1,
            'REPLACE': lambda: s(a[0]).replace(s(a[1]), s(a[2])),
            'SPLIT': lambda: self._split(s(a[0]), s(a[1]), int(a[2])), 'SPACE': lambda: ' ' * int(a[0]),
            'ASCII': lambda: ord(s(a[0])[0]) if s(a[0]) else None, 'CHAR': lambda: chr(int(a[0])),
            'YEAR': lambda: date_part('year', a[0]), 'MONTH': lambda: date_part('month', a[0]),
            'DAY': lambda: date_part('day', a[0]), 'QUARTER': lambda: date_part('quarter', a[0]),
            'WEEK': lambda: date_part('week', a[0]), 'ISOWEEK': lambda: date_part('iso-week', a[0]),
            'DATE': lambda: (lambda d: d.date() if isinstance(d, dt.datetime) else d)(parse_date(a[0]) if isinstance(a[0], str) else as_dt(a[0])),
            'DATETIME': lambda: parse_date(a[0]) if isinstance(a[0], str) else as_dt(a[0]),
            'MAKEDATE': lambda: dt.date(int(a[0]), int(a[1]), int(a[2])),
            'MAKEDATETIME': lambda: dt.datetime.combine(as_dt(a[0]).date(), as_dt(a[1]).time()),
            'ISDATE': lambda: parse_date(a[0]) is not None,
            'PI': lambda: math.pi,
        }.get(name)
        if f is None:
            raise Unsupported(f'function {name}() is not reproduced')
        try:
            return f()
        except (ValueError, TypeError, ZeroDivisionError, IndexError, OverflowError):
            return None

    @staticmethod
    def _split(s, d, n):
        parts = s.split(d)
        try:
            return parts[n - 1] if n > 0 else parts[n]
        except IndexError:
            return ''

    # ----- worksheets -----
    def has_lod(self, key, seen=None):
        """Row-level calcs that contain an LOD depend on the sheet's filters, so they are cached per sheet."""
        seen = seen or set()
        if key in seen or key[0] == 'Parameters' or self.wb.cols[key]['formula'] is None:
            return False
        seen.add(key)
        stack = [self.ast(key)]
        while stack:
            n = stack.pop()
            if isinstance(n, tuple) and n:
                if n[0] == 'lod':
                    return True
                if n[0] == 'key' and self.has_lod(n[1], seen):
                    return True
                stack.extend(x for x in n[1:] if isinstance(x, (tuple, list)))
            elif isinstance(n, list):
                stack.extend(x for x in n if isinstance(x, (tuple, list)))
        return False

    def base_rows(self, ds):
        """Data source rows after the data source filters (cached)."""
        ck = ('base', ds)
        if ck not in self._rows:
            rows = self.ds_rows(ds)
            ctx = Ctx(rows)
            for el in self.wb.ds[ds]['filters']:
                flt = self.parse_filter(el, ds)
                if flt and flt['scope'] == 'row':
                    rows = [r for r in rows if flt['test'](r, ctx)]
            self._rows[ck] = rows
        return self._rows[ck]

    def field(self, ds, inner):
        return Field(self, ds, inner)

    def parse_filter(self, el, default_ds):
        """<filter> element -> {field, scope: row|group|topn|skip, test, text, context} or None."""
        m = QREF.match(el.get('column') or '')
        if not m:
            return None
        ds, inner = m.groups()
        if inner in (':Measure Names', 'Multiple Values', ':Measure Values'):
            return None
        fld = self.field(ds, inner)
        out = {'field': fld, 'context': el.get('context') == 'true', 'scope': 'row', 'ds': ds,
               'text': f'{fld.label}'}
        cls = el.get('class')
        if cls == 'categorical':
            g = el.find('./groupfilter')
            if g is None:
                return None
            if any(x.get('function') in ('end', 'filter') for x in g.iter('groupfilter')):
                out['scope'] = 'topn'
                out['gf'] = g
                out['text'] += ' (Top N / condition)'
                return out
            if g.get('function') == 'level-members' or (g.get(U + 'ui-enumeration') == 'all' and g.get('function') != 'except'):
                return None
            neg = g.get(U + 'ui-enumeration') == 'exclusive' and g.get('function') != 'except'
            test_v = lambda v, g=g, neg=neg: (not self.gf_match(g, v)) if neg else self.gf_match(g, v)
            out['text'] += ' ' + ('NOT ' if neg else '') + self.gf_text(g)
        elif cls == 'quantitative':
            lo = tableau_literal(el.findtext('min')) if el.findtext('min') is not None else None
            hi = tableau_literal(el.findtext('max')) if el.findtext('max') is not None else None
            inc = el.get('included-values') or 'in-range'

            def test_v(v, lo=lo, hi=hi, inc=inc):
                if v is None or isinstance(v, Star):
                    return inc in ('in-range-or-null', 'null', 'all')
                if inc == 'null':
                    return False
                if inc in ('non-null', 'all'):
                    return True
                a, b = comparable(v, lo) if lo is not None else (v, None)
                c, d = comparable(v, hi) if hi is not None else (v, None)
                return (b is None or a >= b) and (d is None or c <= d)
            out['text'] += f' RANGE {to_str(lo) or "-inf"}..{to_str(hi) or "+inf"}' + (f' ({inc})' if inc != 'in-range' else '')
        elif cls == 'relative-date':
            unit = (el.get('period-type-v2') or el.get('period-type') or 'day').lower()
            first, last = int(el.get('first-period') or 0), int(el.get('last-period') or 0)
            anchor = parse_date(el.get('anchor-date') or '') or as_dt(self.today)
            start = date_add(unit, first, trunc(unit, anchor))
            end = date_add(unit, last + 1, trunc(unit, anchor))
            inc_null = el.get('include-null') == 'true'

            def test_v(v, start=start, end=end, inc_null=inc_null):
                if v is None:
                    return inc_null
                v = as_dt(v)
                return start <= v < end
            out['text'] += f' RELATIVE {unit} {first}..{last} (from {to_str(start)} to before {to_str(end)})'
        else:
            out['scope'] = 'skip'
            out['text'] += f' ({cls} filter not reproduced)'
            return out
        if fld.is_measure_filter():
            out['scope'] = 'group'
            out['test'] = lambda rows, ctx, test_v=test_v, fld=fld: test_v(fld.measure_value(rows, ctx))
        else:
            out['test'] = lambda r, ctx, test_v=test_v, fld=fld: test_v(fld.value(r, ctx))
        return out

    def gf_match(self, g, v):
        f = g.get('function')
        if f == 'member':
            lit = tableau_literal(g.get('member'))
            if v is None or lit is None:
                return v is None and lit is None
            a, b = comparable(v, lit)
            return a == b or to_str(a) == to_str(b)
        if f == 'level-members':
            return True
        kids = g.findall('./groupfilter')
        if f == 'union':
            return any(self.gf_match(k, v) for k in kids)
        if f == 'intersection':
            return all(self.gf_match(k, v) for k in kids)
        if f == 'except':
            return bool(kids) and self.gf_match(kids[0], v) and not any(self.gf_match(k, v) for k in kids[1:])
        if f == 'range':
            if v is None:
                return False
            lo, hi = tableau_literal(g.get('from')), tableau_literal(g.get('to'))
            a, b = comparable(v, lo) if lo is not None else (v, None)
            c, d = comparable(v, hi) if hi is not None else (v, None)
            try:
                return (b is None or a >= b) and (d is None or c <= d)
            except TypeError:
                return False
        raise Unsupported(f'filter function {f}')

    def gf_text(self, g):
        f = g.get('function')
        if f == 'member':
            return to_str(tableau_literal(g.get('member'))) if tableau_literal(g.get('member')) is not None else 'NULL'
        kids = [self.gf_text(k) for k in g.findall('./groupfilter')]
        if f == 'union':
            return 'IN(' + ', '.join(kids) + ')'
        if f == 'except':
            return f'{kids[0]} EXCEPT {", ".join(kids[1:])}' if kids else 'EXCEPT'
        if f == 'level-members':
            return 'ALL'
        if f == 'range':
            return f'RANGE({to_str(tableau_literal(g.get("from")))}..{to_str(tableau_literal(g.get("to")))})'
        return f'{f}(' + ', '.join(kids) + ')'

    def topn_members(self, flt, rows, ctx):
        """Members kept by a Top N / condition filter, computed on the context-filtered rows."""
        g = flt['gf']
        fld = flt['field']
        groups = {}
        for r in rows:
            groups.setdefault(to_str(fld.value(r, ctx)), []).append(r)
        keep = set(groups)
        for e in g.iter('groupfilter'):
            fn = e.get('function')
            if fn == 'filter':
                cond = self.resolve(Parser(e.get('expression') or 'TRUE').parse(), fld.ds)
                keep &= {k for k, rs in groups.items() if self.ev(cond, GroupScope(rs), ctx) is True}
            elif fn == 'end':
                order = e.find('./groupfilter')
                expr = (order.get('expression') if order is not None and order.get('expression') else e.get('expression'))
                if not expr:
                    raise Unsupported('Top N filter without a ranking expression')
                node = self.resolve(Parser(expr).parse(), fld.ds)
                desc = (e.get('end') or 'top') == 'top'
                scored = [(self.ev(node, GroupScope(rs), ctx), k) for k, rs in groups.items()]
                scored = [x for x in scored if x[0] is not None and not isinstance(x[0], Star)]
                scored.sort(key=lambda x: x[0], reverse=desc)
                keep &= {k for _, k in scored[:int(e.get('count') or 10)]}
        return keep

    def run_sheet(self, ws_el, group_limit=None):
        """Recompute one worksheet. Returns {status, reason, dims, measures, rows, total, filters, notes, row_count}."""
        name = ws_el.get('name')
        res = {'sheet': name, 'status': 'ok', 'reason': None, 'dims': [], 'measures': [], 'rows': [], 'total': {},
               'filters': [], 'notes': [], 'row_count': None, 'datasource': None, 'confidence': 'high',
               'measure_errors': {}}
        try:
            self._run(ws_el, res)
        except Unsupported as e:
            res['status'] = 'skip'
            res['reason'] = str(e)
        except (RecursionError, ParseError) as e:
            res['status'] = 'skip'
            res['reason'] = f'not evaluated: {e}'
        if res['status'] == 'ok' and res['measure_errors'] and len(res['measure_errors']) == len(res['measures']):
            res['status'] = 'skip'
            res['reason'] = '; '.join(sorted(set(res['measure_errors'].values())))
        elif res['status'] == 'ok' and res['measure_errors']:
            res['status'] = 'partial'
        return res

    def _run(self, ws_el, res):
        tbl = ws_el.find('./table')
        view = tbl.find('./view')
        texts = [tbl.findtext('./rows') or '', tbl.findtext('./cols') or '']
        enc_cols = []
        for p in tbl.findall('./panes/pane'):
            for e in list(p.find('./encodings') if p.find('./encodings') is not None else []):
                enc_cols.append((e.tag, e.get('column') or ''))
        refs, seen = [], set()
        for t in texts + [c for _, c in enc_cols]:
            t = re.sub(r'\[[^\[\]]+\]\.\[__tableau_internal_object_id__\]\.\[', '[__tableau_internal_object_id__].[', t)
            for m in QREF.finditer(t):
                if m.group(0) not in seen:
                    seen.add(m.group(0))
                    refs.append(m.groups())
        fields = [self.field(ds, inner) for ds, inner in refs if ds != 'Parameters']
        mv = [f for f in fields if f.special == 'Measure Values']
        fields = [f for f in fields if not f.special]
        if mv:
            for flt in view.findall('./filter'):
                if (flt.get('column') or '').endswith('[:Measure Names]'):
                    for g in flt.iter('groupfilter'):
                        mm = QREF.search((g.get('member') or '').strip('"'))
                        if g.get('function') == 'member' and mm and mm.group(0) not in seen:
                            seen.add(mm.group(0))
                            fields.append(self.field(*mm.groups()))
        dims = [f for f in fields if not f.measure]
        meas = [f for f in fields if f.measure]
        if not dims and not meas:
            raise Unsupported('no fields on the view (text or image sheet)')
        ds_list = [d.get('name') for d in view.findall('./datasources/datasource') if d.get('name') != 'Parameters']
        primary = next((f.ds for f in dims if f.ds in self.wb.ds), None) or \
            next((f.ds for f in meas if f.ds in self.wb.ds), None) or (ds_list[0] if ds_list else None)
        if primary is None or primary not in self.wb.ds:
            raise Unsupported('no data source on the view')
        res['datasource'] = self.wb.ds[primary]['caption']
        res['dims'] = [f.label for f in dims]
        res['measures'] = [f.label for f in meas]
        sec_dims = [f for f in dims if f.ds != primary]
        if sec_dims:
            raise Unsupported('dimensions from a blended secondary data source: ' + ', '.join(f.label for f in sec_dims))
        rows = self.base_rows(primary)
        filters = [x for x in (self.parse_filter(el, primary) for el in view.findall('./filter')) if x]
        for flt in filters:
            if flt['scope'] == 'skip':
                res['notes'].append('not applied: ' + flt['text'])
        prim = [f for f in filters if f['ds'] == primary]
        ctx = Ctx(rows)
        for flt in prim:
            if flt['context'] and flt['scope'] == 'row':
                rows = [r for r in rows if flt['test'](r, ctx)]
                res['filters'].append('context: ' + flt['text'])
        ctx.lod_rows = rows
        keep_sets = []
        for flt in prim:
            if flt['scope'] == 'topn':
                keep_sets.append((flt['field'], self.topn_members(flt, rows, ctx)))
                res['filters'].append(flt['text'])
        for fld, keep in keep_sets:
            rows = [r for r in rows if to_str(fld.value(r, ctx)) in keep]
        for flt in prim:
            if not flt['context'] and flt['scope'] == 'row':
                rows = [r for r in rows if flt['test'](r, ctx)]
                res['filters'].append(flt['text'])
        dim_fns = []
        for f in dims:
            fn = (lambda r, f=f: f.value(r, ctx))
            fn.key = f.key
            dim_fns.append(fn)
        ctx.view_dims = dim_fns
        ctx.view_rows = rows
        res['row_count'] = len(rows)
        groups = {}
        for r in rows:
            groups.setdefault(tuple(to_str(fn(r)) for fn in dim_fns), []).append(r)
        gfilters = [f for f in prim if f['scope'] == 'group']
        if gfilters:
            for f in gfilters:
                res['filters'].append('measure filter: ' + f['text'])
            groups = {k: g for k, g in groups.items() if all(f['test'](g, ctx) for f in gfilters)}
        secondary = {}
        for f in meas:
            if f.ds != primary:
                secondary.setdefault(f.ds, []).append(f)
        if secondary:
            res['confidence'] = 'medium'
            res['notes'].append('blended: ' + ', '.join(self.wb.ds[d]['caption'] if d in self.wb.ds else d for d in secondary)
                                + ' aggregated on the linking fields, then joined to ' + res['datasource'])
        blend = {d: self.blend(d, fs, dims, filters, res) for d, fs in secondary.items()}
        out_rows = []
        for k, g in sorted(groups.items(), key=lambda kv: tuple('' if x is None else x for x in kv[0])):
            vals = {}
            for f in meas:
                vals[f.label] = self._measure(f, g, ctx, res, blend, k, dims)
            out_rows.append((k, vals))
        for f in meas:
            if f.quick == ['pcto']:
                tot = sum(v[f.label] for _, v in out_rows if isinstance(v[f.label], (int, float)))
                for _, v in out_rows:
                    if isinstance(v[f.label], (int, float)):
                        v[f.label] = v[f.label] / tot if tot else None
        res['rows'] = out_rows
        total = {}
        for f in meas:
            if f.quick == ['pcto']:
                total[f.label] = 1.0 if out_rows else None
            else:
                total[f.label] = self._measure(f, rows, ctx, res, blend, None, dims)
        res['total'] = total

    def _measure(self, f, rows, ctx, res, blend, key, dims):
        if f.label in res['measure_errors']:
            return None
        try:
            if f.ds in blend:
                table, link_pos = blend[f.ds]
                lk = tuple(key[i] for i in link_pos) if key is not None else None
                if key is None:
                    return table.get(('__total__',), {}).get(f.label)
                return table.get(lk, {}).get(f.label)
            if f.quick and f.quick != ['pcto']:
                raise Unsupported(f'quick table calculation {QUICK.get(f.quick[0], f.quick[0])} is not reproduced')
            return f.measure_value(rows, ctx)
        except Unsupported as e:
            res['measure_errors'][f.label] = str(e)
            return None

    def blend(self, sec, fields, dims, filters, res):
        """Secondary data source measures, aggregated per linking field (same caption in both sources)."""
        wbc = self.wb.cols
        link_fields, link_pos = [], []
        for i, d in enumerate(dims):
            cap = (wbc.get(d.key) or {}).get('caption') or d.key[1]
            hit = next((n for (ds, n), c in wbc.items() if ds == sec and ((c.get('caption') or n) == cap or n == d.key[1])), None)
            if hit:
                inner = d.inner.replace(d.key[1], hit) if d.inner != d.key[1] else hit
                link_fields.append(self.field(sec, inner))
                link_pos.append(i)
        rows = self.base_rows(sec)
        ctx = Ctx(rows)
        for flt in filters:
            if flt['ds'] == sec and flt['scope'] == 'row':
                rows = [r for r in rows if flt['test'](r, ctx)]
        ctx.lod_rows = ctx.view_rows = rows
        groups = {}
        for r in rows:
            groups.setdefault(tuple(to_str(lf.value(r, ctx)) for lf in link_fields), []).append(r)
        table = {k: {f.label: f.measure_value(g, ctx) for f in fields} for k, g in groups.items()}
        table[('__total__',)] = {f.label: f.measure_value(rows, ctx) for f in fields}
        if not link_fields:
            res['notes'].append(f'{self.wb.ds[sec]["caption"]}: no linking field in the view, its measures are source totals')
            table[()] = table[('__total__',)]
        return table, link_pos


class Field:
    """One field on a worksheet shelf, e.g. [ds].[yr:Date Sent:ok] or [ds].[sum:Hours:qk]."""
    def __init__(self, oracle, ds, inner):
        self.o = oracle
        wb = oracle.wb
        self.ds, self.inner = ds, inner
        self.label = wb.resolve_ref(ds, inner)
        info = wb.field_info(ds, inner) if ds != '__tableau_internal_object_id__' else {'measure': True}
        self.measure = info['measure']
        self.special = info.get('special')
        self.rows_count = ds == '__tableau_internal_object_id__' or self.label.startswith('COUNT(rows')
        pre, col = [], inner
        if (ds, inner) not in wb.cols:
            pi = parse_instance(inner)
            hit = wb.inst.get((ds, inner))
            if pi:
                pre, col = [p.lower() for p in pi[0]], pi[1]
            elif hit:
                col, pre = hit[0], [DERIV.get((hit[1] or 'none').lower(), (hit[1] or 'none').lower())]
        self.key = (ds, col)
        self.quick = [p for p in pre if p in QUICK]
        self.agg = next((AGG_CODE[p] for p in pre if p in AGG_CODE), None)
        self.date = next((p for p in pre if p in DATE_CODE), None)
        self.user = 'usr' in pre
        if not self.rows_count and self.key not in wb.cols and not self.special:
            self.missing = True
        else:
            self.missing = False

    def is_measure_filter(self):
        if self.rows_count or self.agg or self.quick:
            return True
        if self.key in self.o.wb.cols and self.o.wb.cols[self.key]['formula'] is not None:
            return self.o.is_agg(self.key)
        return False

    def value(self, row, ctx):
        if self.missing:
            raise Unsupported(f'field {self.label} is not defined in the workbook')
        v = self.o.field_value(self.key, RowScope(row), ctx)
        return derive_date(self.date, v) if self.date else v

    def measure_value(self, rows, ctx):
        o = self.o
        if self.rows_count:
            return len(rows)
        if self.missing:
            raise Unsupported(f'field {self.label} is not defined in the workbook')
        c = o.wb.cols[self.key]
        if c['formula'] is not None and o.is_agg(self.key):
            return o.ev(o.ast(self.key), GroupScope(rows), ctx)
        fn = self.agg or {'sum': 'SUM', 'avg': 'AVG', 'count': 'COUNT', 'countd': 'COUNTD', 'min': 'MIN', 'max': 'MAX',
                          'median': 'MEDIAN', 'attr': 'ATTR'}.get((c.get('aggregation') or 'Sum').lower(), 'SUM')
        return o.aggregate(fn, [o.field_value(self.key, RowScope(r), ctx) for r in rows])
