#!/usr/bin/env python3
r"""Bhayankar tester: checks a Power BI rebuild (PBIP) against the Tableau workbook it replaces.

Reads the workbook the same way the Bhayankar extractor does (never the raw XML in an AI context),
reads the PBIP (TMDL model + PBIR report) as files, matches Tableau objects to Power BI objects, and
writes Unit, Functional and Data test results to Excel and CSV. Static: Power BI does not need to run.

Usage (from the toolkit folder):
  python bhayankar_test.py                    # scan test_input\ and write results to test_output\  (the normal way)
  python bhayankar_test.py <workbook.twb|.twbx> <project.pbip|project folder>
         [--data <folder>] [--tds <file>] [--tableau-exports <folder>] [--pbi-exports <folder>]
         [--map <test_map.json>] [--today YYYY-MM-DD] [--username <login>] [--out <folder>]

test_input\<Name>\ (one folder per report; the TWB and PBIP share the name):
  <Name>.twb / .twbx                  the Tableau workbook
  <Name>.pbip + .SemanticModel + .Report   the Power BI project (or any folder holding them)
  *.tds / *.tdsx                      published data sources (optional, matched like the extractor)
  data\                               source tables as <TABLE>.csv / .xlsx (optional: enables the data tests;
                                      the PBIP's own sample-data folder parameter is used when this is missing)
  tableau_exports\<Sheet>.csv|xlsx    Tableau crosstab / data exports per sheet (optional)
  pbi_exports\<Sheet or visual>.csv|xlsx   Power BI "Export data" per visual (optional)
  test_map.json                       corrections to the automatic matching, plus data options (optional)
test_output\<Name>\: test_report.xlsx (Summary, Issues, Unit, Functional, Data, Mapping, Expected Values,
  DAX Queries), the same tables as CSV, expected_values\<sheet>.csv, test_map_suggested.json, test_summary.md.
Modules (keep together): bhayankar_extractor.py, bhayankar_dax.py, bhayankar_report.py, bhayankar_pbip.py,
bhayankar_oracle.py, bhayankar_xlsx.py, bhayankar_validate.py. Python 3 standard library only.
"""
import argparse
import csv
import datetime as dt
import hashlib
import json
import os
import re
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import bhayankar_extractor as bx  # noqa: E402
from bhayankar_report import Rebuild  # noqa: E402
import bhayankar_pbip as bp  # noqa: E402
import bhayankar_oracle as bo  # noqa: E402
import bhayankar_toolkit as bt  # noqa: E402
import bhayankar_validate as bv  # noqa: E402
from bhayankar_xlsx import write_xlsx  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = bt.test_input_dir()
OUTPUT_DIR = bt.test_output_dir()
COLUMNS = ['ID', 'Category', 'Area', 'Test', 'Tableau object', 'Power BI object', 'Expected', 'Actual', 'Status',
           'Severity', 'Details']
SYN = {'flg': 'flag', 'lvl': 'level', 'emp': 'employee', 'num': 'number', 'nbr': 'number', 'qty': 'quantity',
       'amt': 'amount', 'dt': 'date', 'yr': 'year', 'mth': 'month', 'mnth': 'month', 'pct': 'percent',
       'cnt': 'count', 'avg': 'average', 'tot': 'total', 'hrs': 'hour', 'hr': 'hour', 'desc': 'description',
       'mgr': 'manager', 'dept': 'department', 'officier': 'officer', 'qtr': 'quarter', 'no': 'number'}
STOP = {'the', 'of', 'by', 'a', 'an', 'and', 'true', 'is', 'for', 'to', 'in', 'on', 'vs', 'v', 'j'}
DATE_FN = {'YEAR': ['year'], 'QUARTER': ['quarter'], 'MONTH': ['month'], 'DAY': ['day', 'date'], 'WEEK': ['week'],
           'WEEKDAY': ['weekday', 'day'], 'TRUNC_YEAR': ['year'], 'TRUNC_QUARTER': ['quarter'],
           'TRUNC_MONTH': ['month'], 'TRUNC_WEEK': ['week'], 'TRUNC_DAY': ['date', 'day'], 'MONTH_YEAR': ['month'],
           'DATE': ['date'], 'MONTH_DAY': ['day'], 'HOUR': ['hour'], 'MINUTE': ['minute']}
AGG_LABEL = {'SUM', 'AVG', 'COUNT', 'COUNTD', 'MIN', 'MAX', 'ATTR', 'MEDIAN', 'STDEV', 'VAR'}
DAX_AGG = {'SUM': {'SUM', 'SUMX'}, 'AVG': {'AVERAGE', 'AVERAGEX'}, 'COUNTD': {'DISTINCTCOUNT', 'DISTINCTCOUNTNOBLANK'},
           'COUNT': {'COUNT', 'COUNTA', 'COUNTROWS', 'COUNTX', 'COUNTAX', 'COUNTBLANK'}, 'MIN': {'MIN', 'MINX'},
           'MAX': {'MAX', 'MAXX'}, 'MEDIAN': {'MEDIAN', 'MEDIANX'}}
FAMILY = {'bar': {'clusteredBarChart', 'stackedBarChart', 'hundredPercentStackedBarChart', 'clusteredColumnChart',
                  'stackedColumnChart', 'hundredPercentStackedColumnChart', 'barChart', 'columnChart', 'waterfallChart'},
          'line': {'lineChart', 'areaChart', 'stackedAreaChart', 'hundredPercentStackedAreaChart', 'lineClusteredColumnComboChart',
                   'lineStackedColumnComboChart', 'ribbonChart'},
          'pie': {'pieChart', 'donutChart', 'treemap'},
          'table': {'tableEx', 'pivotTable', 'table', 'matrix'},
          'card': {'cardVisual', 'card', 'multiRowCard', 'kpi', 'gauge'},
          'map': {'azureMap', 'map', 'filledMap', 'shapeMap'},
          'scatter': {'scatterChart'},
          'text': {'textbox'}}
MARK_FAMILY = {'Bar': 'bar', 'GanttBar': 'bar', 'Line': 'line', 'Area': 'line', 'Pie': 'pie', 'Text': 'table',
               'Square': 'table', 'Circle': 'scatter', 'Shape': 'scatter', 'Map': 'map', 'Multipolygon': 'map'}
NON_DATA = {'textbox', 'image', 'shape', 'basicShape', 'actionButton', 'pageNavigator', 'bookmarkNavigator', 'group', None}
SLICERS = {'slicer', 'listSlicer', 'advancedSlicerVisual', 'textSlicer'}
DAX_REF = re.compile(r"(?:'((?:[^']|'')+)'|(?<![\w\]'])([A-Za-z_][A-Za-z0-9_]*))?\[((?:[^\]]|\]\])+)\]")
DAX_STR = re.compile(r'"(?:[^"]|"")*"')
SEV = {'FAIL': 0, 'WARN': 1, 'MANUAL': 2, 'SKIP': 3, 'INFO': 4, 'PASS': 5}


# ---------- names ----------
def norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def tokens(s):
    s = str(s or '').replace('%', ' percent ').replace('#', ' count ')
    s = re.sub(r'([a-z])([A-Z])', r'\1 \2', s)
    out = set()
    for w in re.split(r'[^A-Za-z0-9]+', s.lower()):
        if not w:
            continue
        w = SYN.get(w, w)
        if len(w) > 3 and w.endswith('s') and not w.endswith('ss'):
            w = w[:-1]
        if w not in STOP:
            out.add(w)
    return out


def sim(a, b):
    if norm(a) and norm(a) == norm(b):
        return 1.0
    A, B = tokens(a), tokens(b)
    if not A or not B:
        return 0.0
    i = len(A & B)
    if not i:
        return 0.0
    return round((i / len(A | B) + i / min(len(A), len(B))) / 2 * 0.95, 3)


def oid(t, n):
    """DAX reference: a ']' in a name is written ']]', a "'" in a table name "''" (measure [1-[Ratio]] -> [1-[Ratio]]])."""
    n = str(n).replace(']', ']]')
    return f"'{str(t).replace(chr(39), chr(39) * 2)}'[{n}]" if t else f'[{n}]'


def strip_dax(e):
    e = DAX_STR.sub('""', e or '')
    e = re.sub(r'//[^\n]*|--[^\n]*', '', e)
    return re.sub(r'/\*.*?\*/', '', e, flags=re.S)


def dax_refs(e):
    return [((m.group(1) or m.group(2) or '').replace("''", "'") or None, m.group(3).replace(']]', ']'))
            for m in DAX_REF.finditer(strip_dax(e))]


def label_parts(label):
    """'SUM(Email Sent)' -> ('SUM', 'Email Sent'); 'PERCENT_OF_TOTAL(SUM(X))' -> ('PERCENT_OF_TOTAL', 'SUM(X)')."""
    m = re.match(r'^([A-Z_]+)\((.*)\)$', (label or '').strip())
    return (m.group(1), m.group(2)) if m else (None, (label or '').strip())


def fmt(v):
    if v is None:
        return ''
    if isinstance(v, float):
        return f'{v:.6g}' if abs(v) < 1e15 else str(v)
    if isinstance(v, (set, list, tuple)):
        return ', '.join(fmt(x) for x in v)
    return str(v)


# ---------- export comparison ----------
MONTH_NUM = {m.lower(): i + 1 for i, m in enumerate(bo.MONTHS)}
MONTH_NUM.update({m[:3].lower(): i + 1 for i, m in enumerate(bo.MONTHS)})


def key_norm(v):
    if v is None:
        return '(null)'
    s = str(v).strip()
    if s.lower() in ('', 'null', '(blank)', '%null%', 'none', '(null)'):
        return '(null)'
    if s.lower() in MONTH_NUM:
        return str(MONTH_NUM[s.lower()])
    m = re.match(r'^q([1-4])$', s, re.I)
    if m:
        return m.group(1)
    for f in ('%B %Y', '%b %Y', '%Y-%m', '%b-%y', '%B, %Y'):
        try:
            return dt.datetime.strptime(s, f).strftime('%Y-%m-01')
        except ValueError:
            pass
    d = bo.parse_date(s) if re.match(r'^\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}', s) else None
    if d:
        return d.strftime('%Y-%m-%d') if d.time() == dt.time() else d.strftime('%Y-%m-%d %H:%M:%S')
    try:
        f = float(s.replace(',', ''))
        return str(int(f)) if f.is_integer() else repr(round(f, 9))
    except ValueError:
        pass
    if s.lower() in ('true', 'false'):
        return s.lower()
    return re.sub(r'\s+', ' ', s.lower())


def parse_num(s):
    """Displayed number -> (value, decimals shown, was percent) or None."""
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return float(s), None, False
    t = str(s).strip()
    if t.lower() in ('', 'null', '(blank)', '-', 'nan'):
        return None
    pct = t.endswith('%')
    neg = t.startswith('(') and t.endswith(')')
    core = re.sub(r'[^0-9.\-eE]', '', t.replace(',', ''))
    try:
        v = float(core)
    except ValueError:
        return None
    dec = len(core.split('.')[1]) if '.' in core and 'e' not in core.lower() else 0
    return (-v if neg else v), dec, pct


def num_match(expected, shown):
    if expected is None or isinstance(expected, bo.Star):
        return shown is None
    if shown is None:
        return False
    if isinstance(expected, bool):
        expected = float(expected)
    if not isinstance(expected, (int, float)):
        return key_norm(expected) == key_norm(shown[0])
    v, dec, pct = shown
    e = expected * 100 if pct else expected
    if dec is not None and dec < 9 and str(shown[0]) != repr(v):
        if abs(bo.round_half_away(e, dec) - v) <= 10 ** -(dec + 3) or abs(e - v) <= 0.5 * 10 ** -dec + 1e-9:
            return True
    return abs(e - v) <= max(1e-6, abs(e) * 1e-6)


# ---------- the tester ----------
class Tester:
    def __init__(self, case):
        self.case = case
        self.rows = []
        self.notes = []
        self.mapping = []
        self.expected = []
        self.dax_queries = []
        self.dax_parts = {}  # sheet -> {groups, measures, notes, filters}: lets bhayankar_verify run the query in Desktop
        self.overrides = case.get('test_map') or {}
        self.t0 = time.time()

    def add(self, cat, area, test, tab='', pbi='', expected='', actual='', status='INFO', severity='', details=''):
        if not severity:
            severity = {'FAIL': 'High', 'WARN': 'Medium', 'MANUAL': 'Medium'}.get(status, 'Low')
        self.rows.append({'Category': cat, 'Area': area, 'Test': test, 'Tableau object': fmt(tab),
                          'Power BI object': fmt(pbi), 'Expected': fmt(expected), 'Actual': fmt(actual),
                          'Status': status, 'Severity': severity, 'Details': fmt(details)})

    # ----- loading -----
    def load(self):
        c = self.case
        root, assets = bx.load_root(c['twb'])
        troots = [(bx.load_root(t, ('.tds',))[0], t) for t in c.get('tds', [])]
        self.root = root
        self.wb = bx.Workbook(root, assets, troots, c.get('table_map') or {}, strict_tds=c.get('strict_tds', True))
        self.res = self.wb.build()
        self.res['workbook']['tds_merged'] = self.wb.tds_used
        self.spec = Rebuild(self.res).build()
        self.proj = bp.load_project(c['pbip'])
        self.model = self.proj['model'] or {'tables': {}, 'relationships': [], 'roles': [], 'expressions': {}, 'measures': {}}
        self.report = self.proj['report'] or {'pages': [], 'filters': []}
        for e in self.proj['errors']:
            self.add('Unit', 'PBIP', 'Project files readable', '', c['pbip'], 'model and report readable', e, 'FAIL',
                     details='fix the path or the project; tests that need it are skipped')
        self.ws_el = {w.get('name'): w for w in root.findall('./worksheets/worksheet')}
        self.ws = {w['name']: w for w in self.res['worksheets']}
        self.calc_by_name, self.calc_versions = {}, {}
        for cl in self.res['calculations']:
            self.calc_by_name.setdefault(cl['name'], cl)
            self.calc_versions.setdefault(cl['name'], []).append(cl)  # one per data source that defines the name
        self.calc_keys = {}
        for (ds, n), col in self.wb.cols.items():
            if col['formula'] is not None and ds != 'Parameters':
                self.calc_keys.setdefault(col['caption'] or n, (ds, n))
        self.index_model()

    def index_model(self):
        m = self.model
        self.all_cols = {}   # (table, col) -> column
        self.all_meas = {}   # (table, measure) -> measure
        for t in m['tables'].values():
            for c in t['columns'].values():
                self.all_cols[(t['name'], c['name'])] = c
            for ms in t['measures'].values():
                self.all_meas[(t['name'], ms['name'])] = ms
        self.meas_by_name = {}
        for (t, n), ms in self.all_meas.items():
            self.meas_by_name.setdefault(n, (t, n))
        self.col_names = {}
        for (t, n) in self.all_cols:
            self.col_names.setdefault(n, []).append((t, n))
        self._closure = {}
        used = set()
        for p in self.report['pages']:
            for v in p['visuals']:
                for r in v['all_refs']:
                    used.add((r['table'], r['name']))
            for f in p['filters']:
                if f['field']:
                    used.add((f['field']['table'], f['field']['name']))
        self.report_refs = used

    def closure(self, kind, key, seen=None):
        """Transitive DAX of a measure / calculated column -> {'cols', 'roots', 'fns', 'text', 'measures'}."""
        ck = (kind, key)
        if ck in self._closure:
            return self._closure[ck]
        seen = seen or set()
        out = {'cols': set(), 'roots': set(), 'fns': set(), 'text': '', 'measures': set()}
        if ck in seen:
            return out
        seen.add(ck)
        if kind == 'measure':
            expr = (self.all_meas.get(key) or {}).get('expression') or ''
        else:
            col = self.all_cols.get(key) or {}
            if col.get('kind') != 'calculated':
                out['cols'].add(key)
                root = col.get('source_root') or key[1]
                out['roots'].add(str(root).upper())
                t = self.model['tables'].get(key[0]) or {}
                out['text'] = t.get('m', '') if root in t.get('m_added', set()) else ''
                self._closure[ck] = out
                return out
            expr = col.get('expression') or ''
        out['text'] = expr
        out['fns'] |= {f.upper() for f in re.findall(r'\b([A-Za-z][A-Za-z0-9.]*)\s*\(', strip_dax(expr))}
        home = key[0]
        for t, n in dax_refs(expr):
            if t and (t, n) in self.all_meas:
                sub = self.closure('measure', (t, n), seen)
            elif t and (t, n) in self.all_cols:
                sub = self.closure('column', (t, n), seen)
            elif not t and n in self.meas_by_name:
                sub = self.closure('measure', self.meas_by_name[n], seen)
                out['measures'].add(self.meas_by_name[n])
            elif not t and (home, n) in self.all_cols:
                sub = self.closure('column', (home, n), seen)
            else:
                continue
            if t and (t, n) in self.all_meas:
                out['measures'].add((t, n))
            for k in ('cols', 'roots', 'fns', 'measures'):
                out[k] |= sub[k]
            out['text'] += '\n' + sub['text']
        self._closure[ck] = out
        return out

    # ----- matching -----
    def match(self):
        self.match_tables()
        self.match_columns()
        self.match_calcs()
        self.match_pages()
        self.match_visuals()

    def src_tables(self):
        names = set()
        for c in self.res['columns']:
            if c['table']:
                names.add(c['table'])
        for d in self.wb.ds.values():
            conn = (d['connection'] or {})
            conn = conn.get('tds') or conn
            for leaf in bo.Oracle.leaves(conn.get('relations')):
                t = (leaf.get('table') or '').replace('[', '').replace(']', '').split('.')[-1]
                if t:
                    names.add(t)
            for leaf in (conn.get('physical_by_object') or {}).values():
                t = (leaf.get('table') or '').replace('[', '').replace(']', '').split('.')[-1] if leaf else ''
                if t:
                    names.add(t)
        return sorted(names)

    def match_tables(self):
        ov = self.overrides.get('tables') or {}
        tmap = self.case.get('table_map') or {}
        self.table_match = {}
        for s in self.src_tables():
            if s in ov:
                self.table_match[s] = [ov[s]] if ov[s] else []
                self.mapping.append(['Table', s, ov[s] or '(not migrated)', 1.0, 'test_map.json'])
                continue
            cands = []
            for t in self.model['tables'].values():
                lits = [x.upper() for x in t.get('literals', [])]
                su = s.upper()
                sc = 1.0 if su in lits else 0.85 if any(re.search(r'(?<![A-Z0-9_])' + re.escape(su) + r'(?![A-Z0-9_])', x)
                                                          for x in lits) else 0
                sc = max(sc, sim(s, t['name']), 1.0 if tmap.get(s) == t['name'] else 0)
                if sc >= 0.6:
                    cands.append((sc, t['name']))
            cands.sort(key=lambda x: -x[0])
            self.table_match[s] = [n for _, n in cands]
            self.mapping.append(['Table', s, ', '.join(n for _, n in cands[:3]) or '(none)',
                                 cands[0][0] if cands else 0, 'M source / name'])

    def match_columns(self):
        ov = self.overrides.get('columns') or {}
        self.col_match = {}
        for c in self.res['columns']:
            name = c['name']
            if name in ov:
                t, n = self.parse_obj(ov[name])
                self.col_match.setdefault(name, []).extend([(t, n)] if n else [])
                continue
            src = (c['source_column'] or '').upper()
            pref = set(self.table_match.get(c['table'] or '', []))
            cands = []
            for (t, n), col in self.all_cols.items():
                sr = str(col.get('source_root') or '').upper()
                s = 0.0
                if src and sr == src:
                    s = 0.9
                elif src and norm(sr) == norm(src):
                    s = 0.85
                ns = sim(name, n)
                if ns >= 0.8:
                    s = max(s, 0.85 * ns)
                if s and pref:
                    s = s + 0.1 if t in pref else s - 0.15
                if s >= 0.6:
                    cands.append((round(s, 3), t, n))
            cands.sort(key=lambda x: -x[0])
            lst = self.col_match.setdefault(name, [])
            for s, t, n in cands:
                if (t, n) not in lst:
                    lst.append((t, n))
            self.mapping.append(['Column', f'{c["table"]}.{c["source_column"]} ({name})',
                                 ', '.join(oid(t, n) for _, t, n in cands[:3]) or '(none)', cands[0][0] if cands else 0,
                                 'source column / name'])

    def parse_obj(self, s):
        """'Table[Name]' / "'My Table'[Name]" / '[Name]' / 'Name' -> (table, name)."""
        if not s:
            return None, None
        m = re.match(r"^\s*(?:'((?:[^']|'')+)'|([^\[]+))?\[(.+)\]\s*$", s)
        if m:
            return ((m.group(1) or m.group(2) or '').strip() or None), m.group(3)
        return None, s.strip()

    def resolve_obj(self, t, n):
        """Override target -> ('measure'|'column', (table, name)) or None."""
        if t and (t, n) in self.all_meas:
            return 'measure', (t, n)
        if t and (t, n) in self.all_cols:
            return 'column', (t, n)
        if n in self.meas_by_name:
            return 'measure', self.meas_by_name[n]
        if n in self.col_names:
            return 'column', self.col_names[n][0]
        return None

    def calc_fingerprint(self, name):
        key = self.calc_keys.get(name)
        roots, fns, lits, seen = set(), set(), set(), set()
        stack = [key] if key else []
        while stack:
            k = stack.pop()
            if k in seen or k not in self.wb.cols:
                continue
            seen.add(k)
            c = self.wb.cols[k]
            if c['formula'] is None:
                if k[0] != 'Parameters' and c.get('remote') and c['remote'][2]:
                    roots.add(c['remote'][2].upper())
                elif k[0] != 'Parameters':
                    roots.add(norm(c['caption'] or k[1]).upper())
                continue
            f = c['formula']
            fns |= {x.upper() for x in re.findall(r'\b([A-Za-z_]+)\s*\(', f)}
            lits |= {x[1:-1] for x in re.findall(r'"[^"]*"|\'[^\']*\'', f) if len(x) > 2}
            lits |= {x for x in re.findall(r'(?<![\w.\[])(\d+\.?\d*)(?![\w\]])', re.sub(r'\[[^\]]*\]', '', f))
                     if x not in ('0', '1', '0.0', '1.0')}
            stack.extend(self.wb.formula_refs(k[0], f))
            if c['bins']:
                stack.append((k[0], c['bins']['source']))
        return roots, fns, lits

    def desc_mentions(self, desc, name):
        if not desc or not name:
            return False
        d = desc.lower()
        n = re.escape(name.strip().lower())
        return bool(re.search(r'tableau[^.;]{0,40}?[\[\"\']?' + n + r'(?![a-z0-9])', d) or
                    re.search(r'\[' + n + r'\]', d))

    def match_calcs(self):
        ov = self.overrides.get('calcs') or {}
        self.calc_match = {}
        cands_all = [('measure', k, ms['name'], ms.get('description')) for k, ms in self.all_meas.items()] + \
                    [('column', k, c['name'], c.get('description')) for k, c in self.all_cols.items()]
        scored = []
        for cl in self.res['calculations']:
            name = cl['name']
            if name in ov:
                t, n = self.parse_obj(ov[name]) if ov[name] else (None, None)
                hit = self.resolve_obj(t, n) if n else None
                self.calc_match[name] = (hit[0], hit[1], 1.0, 'test_map.json') if hit else (None, None, 1.0, 'test_map.json: not migrated' if not ov[name] else f'test_map.json: {ov[name]} not found')
                continue
            roots, fns, _ = self.calc_fingerprint(name)
            row_level = cl['pbi_kind'] == 'calculated column' or cl['role'] == 'dimension'
            for kind, key, pname, desc in cands_all:
                s_name = sim(name, pname)
                s_desc = self.desc_mentions(desc, name)
                if not s_desc and s_name < 0.3:
                    continue
                clo = self.closure(kind, key)
                s_fp = 0.0
                if roots and clo['roots']:
                    j = len(roots & clo['roots']) / len(roots | clo['roots'])
                    tab_fns = {f for f in fns if f in DAX_AGG}
                    dax_fns = clo['fns']
                    agg_ok = all(DAX_AGG[f] & dax_fns for f in tab_fns) if tab_fns else True
                    s_fp = 0.7 * j + 0.3 * agg_ok
                if s_name == 1.0:
                    score = 1.0
                elif s_desc:
                    score = 0.75 + 0.25 * max(s_name, s_fp)
                else:
                    score = 0.6 * s_name + 0.4 * s_fp if roots else 0.9 * s_name
                if row_level and kind == 'measure':
                    score *= 0.85
                elif not row_level and kind == 'column' and cl['role'] == 'measure' and \
                        self.wb.cols.get(self.calc_keys.get(name), {}).get('formula') and 'SUM(' in (cl['formula'] or '').upper():
                    score *= 0.85
                if score >= 0.45:
                    method = 'same name' if s_name == 1.0 else 'description mentions it' if s_desc else 'name + logic'
                    scored.append((round(score, 3), name, kind, key, method))
        scored.sort(key=lambda x: -x[0])
        taken = set()
        for score, name, kind, key, method in scored:
            if name in self.calc_match:
                continue
            if (kind, key) in taken and score < 0.9:
                continue
            self.calc_match[name] = (kind, key, score, method)
            taken.add((kind, key))
        for score, name, kind, key, method in scored:  # second pass: same Power BI object may serve two calcs
            if name not in self.calc_match and score >= 0.6:
                self.calc_match[name] = (kind, key, score * 0.95, method + ' (shared)')
        for cl in self.res['calculations']:
            m = self.calc_match.get(cl['name'])
            self.mapping.append(['Calculation', cl['name'], oid(*m[1]) if m and m[1] else '(none)',
                                 m[2] if m else 0, m[3] if m else 'no candidate'])

    def match_pages(self):
        ov = self.overrides.get('pages') or {}
        self.page_match = {}
        pages = {p['name']: p for p in self.report['pages']}
        scored = []
        for d in self.res['dashboards']:
            if d['name'] in ov:
                self.page_match[d['name']] = pages.get(ov[d['name']]) if ov[d['name']] else None
                continue
            for p in self.report['pages']:
                s = sim(d['name'], p['name'])
                if s >= 0.55:
                    scored.append((s, d['name'], p['id']))
        scored.sort(key=lambda x: -x[0])
        used = set()
        by_id = {p['id']: p for p in self.report['pages']}
        for s, d, pid in scored:
            if d in self.page_match or pid in used:
                continue
            self.page_match[d] = by_id[pid]
            used.add(pid)
        for d in self.res['dashboards']:
            p = self.page_match.get(d['name'])
            self.mapping.append(['Page', d['name'], p['name'] if p else '(none)',
                                 sim(d['name'], p['name']) if p else 0, 'test_map.json' if d['name'] in ov else 'name'])

    # ----- Tableau field label -> Power BI objects -----
    def candidates(self, label, depth=0):
        """Every Power BI (table, name) that can stand for a Tableau field label."""
        out = []
        if not label or depth > 3:
            return out
        if label.startswith('Param:'):
            p = label[6:]
            for (t, n) in list(self.all_cols) + list(self.all_meas):
                if sim(p, n) >= 0.7 or sim(p, t) >= 0.7:
                    out.append((t, n))
            return out
        fn, inner = label_parts(label)
        if fn in DATE_FN:
            return self.date_candidates(fn, inner)
        if fn in ('PERCENT_OF_TOTAL', 'RUNNING_TOTAL', 'DIFFERENCE', 'PERCENT_DIFFERENCE', 'RANK', 'WINDOW',
                  'MOVING_AVG', 'YTD_TOTAL'):
            out = self.candidates(inner, depth + 1)
            for (t, n) in self.all_meas:
                if sim(label.replace('PERCENT_OF_TOTAL', '% of total'), n) >= 0.6:
                    out.append((t, n))
            return out
        if fn == 'COUNT' and inner.startswith('rows of'):
            for k, ms in self.all_meas.items():
                if 'COUNTROWS' in (ms['expression'] or '').upper():
                    out.append(k)
            return out
        base = inner if fn in AGG_LABEL else label
        # a model object with exactly this name first: the name matching below ignores symbols, so "c.Profit YTD -"
        # would otherwise stand for "c.Profit | YTD" and "Delta -" for "Delta +"
        for name in dict.fromkeys((label, base)):  # the whole label too: the build names constants 'MIN(0.2)'
            low = name.strip().lower()
            out.extend(k for k in list(self.all_meas) + list(self.all_cols) if str(k[1]).strip().lower() == low)
        if base in self.calc_match and self.calc_match[base][1]:
            out.append(self.calc_match[base][1])
        if len(self.calc_versions.get(base, ())) > 1:
            # a name several data sources define: the build makes one measure per data source ("<calc> (<source>)"),
            # each described as the Tableau calc; any of them stands for the label
            tag = f'Tableau calc "{base}":'
            out.extend(k for k, ms in self.all_meas.items() if (ms.get('description') or '').startswith(tag))
            out.extend(k for k in self.all_cols if k[1] == base)  # a row calc keeps its name on each source's table
        cols = list(self.col_match.get(base, []))
        out.extend(cols)
        if fn in AGG_LABEL and fn != 'ATTR':
            fam = DAX_AGG.get(fn, set())
            for k, ms in self.all_meas.items():
                e = strip_dax(ms['expression']).strip()
                m = re.match(r'^([A-Z.]+)\s*\(\s*(.*)\)$', e, re.I | re.S)
                if not m or m.group(1).upper() not in fam:
                    continue
                refs = dax_refs(m.group(2))
                if any((t or k[0], n) in cols for t, n in refs):
                    out.append(k)
        if not out:
            for (t, n) in list(self.all_meas) + list(self.all_cols):
                if sim(base, n) >= 0.85:
                    out.append((t, n))
        return list(dict.fromkeys(out))

    def date_candidates(self, fn, inner):
        words = DATE_FN[fn]
        base = self.candidates(inner, 1)
        out = list(base)
        date_tables = set()
        # the build relates Date through a hidden date-only key '<column> (Date)'
        keys = base + [(t, f'{c} (Date)') for t, c in base if (t, f'{c} (Date)') in self.all_cols]
        for t, c in keys:
            for r in self.model['relationships']:
                if r['from_table'] == t and r['from_column'] == c:
                    date_tables.add(r['to_table'])
                if r['to_table'] == t and r['to_column'] == c:
                    date_tables.add(r['from_table'])
        for t, c in base:
            for (tt, n) in self.all_cols:
                tk = tokens(n)
                if tt in date_tables and any(w in tk for w in words):
                    out.append((tt, n))
                elif tt == t and any(w in tk for w in words) and tokens(inner) & tk:
                    out.append((tt, n))
        return list(dict.fromkeys(out))

    def visual_fields(self, v):
        out = {(f['table'], f['name']) for f in v['fields']} | {(f['table'], f['name']) for f in v['all_refs']}
        return out | self.variant_bases(out)

    def variant_bases(self, keys):
        """A sheet-filtered measure variant ('WO Count (Flag DC Assigned)', made when one measure is filtered
        differently on different sheets) stands for its base measure: same logic plus the sheet's filters."""
        names = {n: k for k, n in ((k, k[1]) for k in self.all_meas)}
        out = set()
        for k in keys:
            ms = self.all_meas.get(k)
            desc = (ms or {}).get('description') or ''
            tc = re.match(r'Tableau calc "(.+?)": ', desc)
            if tc and tc.group(1) != k[1] and tc.group(1) in names:
                # the same Tableau calc under another model name: a table calc over another sheet's addressing
                # ("<calc> (along X)") or the copy of another data source ("<calc> (<data source>)")
                out.add(names[tc.group(1)])
                continue
            if not ms or 'filtered variant' not in desc:
                continue
            n = k[1]
            for i in [i for i in range(len(n)) if n.startswith(' (', i)][::-1]:
                if n[:i] in names:
                    out.add(names[n[:i]])
                    break
        return out

    def to_model_refs(self, dax):
        """Draft DAX of a Tableau calc names columns by their Tableau caption; the model may keep the source
        name ('...'[Date Application Review] -> 'VW_X'[DATE_APPLICATION_REVIEW]). Rewrite each column reference
        to the model column it was matched to, so the draft can be found inside a measure the build folded it into."""
        def sub(m):
            name = m.group(3).replace(']]', ']')
            hit = next(iter(self.col_match.get(name) or []), None)
            return f"'{hit[0]}'[{hit[1]}]" if hit else m.group(0)
        return DAX_REF.sub(sub, dax or '')

    @staticmethod
    def dax_shape(s):
        """DAX without spaces, with parameter value references and hoisted VARs as one token: a calc's DAX and the
        same condition folded into a measure (VAR _p1 = [Start Date Value] ... _p1) compare equal."""
        s = re.sub(r"\[[^\]]+ Value\]|\b_p\d+\b", '§', s or '')
        return re.sub(r'\s+', '', s).lower()

    def sheet_labels(self, ws):
        """Tableau labels on a worksheet -> (dims, measures, tooltip-only)."""
        dims, meas, tips = [], [], []
        for f in ws['shelf']['rows'] + ws['shelf']['cols']:
            if f.get('special') == 'Measure Values':
                meas.extend(ws.get('measure_values') or [])
            elif not f.get('special') and not f.get('parameter'):
                (meas if f['measure'] else dims).append(f['label'])
        for p in ws['panes']:
            for e in p.get('enc', []):
                if e.get('special') == 'Measure Values':
                    meas.extend(ws.get('measure_values') or [])
                elif e.get('special') or e.get('parameter'):
                    continue
                elif e['channel'] == 'tooltip':
                    tips.append(e['label'])
                else:
                    (meas if e['measure'] else dims).append(e['label'])
        def real(lab):
            """Measure Values come as '[SUM(X)]'; constants ('" "', MIN(0), a calc without fields) and deleted fields
            are layout scaffolding in Tableau and show nothing: the rebuild leaves them out, so does this check."""
            lab = lab[1:-1] if lab.startswith('[') and lab.endswith(']') else lab
            if lab in ('Measure Names', "''", '""') or lab.startswith('(deleted field') or '(deleted field' in lab or \
                    re.search(r'\b(?:Latitude|Longitude|Geometry) \(generated\)', lab):  # Tableau geocoding: no source column
                return None
            fn, inner = label_parts(lab)
            if re.fullmatch(r'-?[\d.]+|"[^"]*"|\'[^\']*\'', inner.strip()):
                return None
            calc = self.calc_by_name.get(inner)
            if calc and '[' not in (calc.get('formula') or '['):
                return None
            return lab
        dd = lambda xs: list(dict.fromkeys(x for x in map(real, xs) if x))
        dims, meas, tips = dd(dims), dd(meas), dd(tips)
        return dims, meas, [t for t in tips if t not in dims and t not in meas]

    def match_visuals(self):
        ov = self.overrides.get('visuals') or {}
        self.vis_match = {}  # (dashboard, sheet) -> visual
        for sp in self.spec['pages']:
            page = self.page_match.get(sp['name'])
            if not page:
                continue
            pw, ph = (page.get('width') or 1280), (page.get('height') or 720)
            tw, th = (sp.get('width') or pw), (sp.get('height') or ph)
            cands = [v for v in page['visuals'] if v['type'] not in SLICERS and v['type'] not in ('image', 'actionButton',
                                                                                                   'shape', 'basicShape', 'group',
                                                                                                   'pageNavigator')]
            scored = []
            by_id = {v['id']: v for v in page['visuals']}
            for ev in sp['visuals']:
                sheet = ev['sheet']
                k = f'{sp["name"]} / {sheet}'
                if k in ov:
                    v = next((x for x in page['visuals'] if ov[k] and (x['id'] == ov[k] or x['title'] == ov[k])), None)
                    self.vis_match[(sp['name'], sheet)] = v
                    continue
                # bhayankar_build names each visual sha1("<page>:<id>:<sheet>"): an exact match needs no title or field guess
                # (an untitled title card, or two sheets whose names share a prefix, used to be missed or swapped)
                exact = by_id.get(hashlib.sha1(f'{sp["name"]}:{ev.get("id")}:{sheet}'.encode('utf-8')).hexdigest()[:20])
                if exact:
                    self.vis_match[(sp['name'], sheet)] = dict(exact, _score=1.0)
                    exact['_score'] = 1.0
                    continue
                ws = self.ws.get(sheet) or {}
                dims, meas, tips = self.sheet_labels(ws) if ws else ([], [], [])
                want = set()
                for lb in dims + meas:
                    want |= set(self.candidates(lb))
                titles = [sheet, (ev.get('title') or {}).get('text') or '', ws.get('title') or '']
                pos = ev.get('position') or {}
                for v in cands:
                    vt = (v['title'] or v.get('text') or '')
                    s_title = max(sim(t, vt) for t in titles if t) if vt else 0
                    vf = self.visual_fields(v)
                    s_fields = len(want & vf) / len(want) if want else 0
                    if pos and v['w']:
                        cx = (pos['x'] + pos['width'] / 2) / tw - (v['x'] + v['w'] / 2) / pw
                        cy = (pos['y'] + pos['height'] / 2) / th - (v['y'] + v['h'] / 2) / ph
                        s_pos = max(0.0, 1 - (cx * cx + cy * cy) ** 0.5 * 2)
                    else:
                        s_pos = 0
                    score = 0.4 * s_title + 0.4 * s_fields + 0.2 * s_pos
                    if score >= 0.3:
                        scored.append((round(score, 3), sheet, v['id'], v))
            scored.sort(key=lambda x: -x[0])
            used = {v['id'] for (pn, _), v in self.vis_match.items() if pn == sp['name'] and v}  # exact matches are taken
            for score, sheet, vid, v in scored:
                if (sp['name'], sheet) in self.vis_match or vid in used:
                    continue
                self.vis_match[(sp['name'], sheet)] = v
                v['_score'] = score
                used.add(vid)
            for score, sheet, vid, v in scored:  # one Power BI visual can replace several Tableau sheets
                if (sp['name'], sheet) not in self.vis_match and score >= 0.5:
                    self.vis_match[(sp['name'], sheet)] = dict(v, _score=score, _shared=True)
            for ev in sp['visuals']:
                v = self.vis_match.get((sp['name'], ev['sheet']))
                self.mapping.append(['Visual', f'{sp["name"]} / {ev["sheet"]}',
                                     f'{page["name"]} / {v["type"]} "{v["title"] or v.get("text") or v["id"]}"' if v else '(none)',
                                     v.get('_score', 1.0) if v else 0, 'title + fields + position'])

    # ----- unit tests -----
    def unit(self):
        self.u_columns()
        self.u_calcs()
        self.u_dax()
        self.u_params()
        self.u_relationships()
        self.u_rls()
        self.u_date()

    def u_columns(self):
        tmap = {'string': {'string'}, 'integer': {'int64', 'double', 'decimal'}, 'real': {'double', 'decimal', 'int64'},
                'date': {'datetime', 'date'}, 'datetime': {'datetime', 'date'}, 'boolean': {'boolean'}}
        for c in self.res['columns']:
            hits = self.col_match.get(c['name']) or []
            src = f'{c["table"]}.{c["source_column"]}' if c['source_column'] else c['name']
            if not hits and c['name'] in ('Latitude (generated)', 'Longitude (generated)', 'Number of Records', 'Measure Names',
                                          'Measure Values'):
                self.add('Unit', 'COL', 'Source column used by Tableau exists in the model', c['name'], '', 'not a source column',
                         'made by Tableau', 'SKIP', details='Tableau generates this field (map coordinates, row count, '
                                                            'measure names/values); Power BI maps use the location columns')
                continue
            if not hits:
                calc = next((cl for cl in self.calc_by_name.values() if cl['name'] == c['name']), None)
                self.add('Unit', 'COL', 'Source column used by Tableau exists in the model', f'{c["name"]} ({src})', '',
                         'a model column with this source column', 'not found', 'FAIL' if not calc else 'WARN',
                         details='add the column, or map it in test_map.json "columns" if it was renamed or merged')
                continue
            t, n = hits[0]
            col = self.all_cols[(t, n)]
            ok_types = tmap.get(c['datatype'] or '', set())
            pdt = (col.get('dataType') or '').lower()
            self.add('Unit', 'COL', 'Source column used by Tableau exists in the model', f'{c["name"]} ({src})', oid(t, n),
                     c['source_column'] or c['name'], col.get('source_root') or n, 'PASS',
                     details=f'{len(hits)} model column(s) carry it' if len(hits) > 1 else '')
            if ok_types and pdt and pdt not in ok_types:
                self.add('Unit', 'TYPE', 'Data type compatible', c['name'], oid(t, n), c['datatype'], pdt, 'WARN',
                         details='numbers, dates and booleans should keep their type so totals and slicers behave the same')

    def u_calcs(self):
        fmts = {}
        for w in self.res['worksheets']:
            for k, v in (w.get('number_formats') or {}).items():
                fmts.setdefault(k.strip('[]').strip(), v)
        # row calcs that read a parameter (directly or through another calc) are not columns: bhayankar_build inlines
        # their formula into the measures / measure filters that use them, so there is nothing of that name to find
        by_name = {c['name']: c for c in self.res['calculations']}
        param_row = {c['name'] for c in self.res['calculations'] if c.get('pbi_kind') == 'calculated column'
                     and any('cannot react to slicers' in n for n in c.get('dax_notes') or [])}
        grew = True
        while grew:
            grew = False
            for c in self.res['calculations']:
                if c['name'] not in param_row and c.get('pbi_kind') == 'calculated column' and \
                        {r for r in re.findall(r'\[([^\]]+)\]', c.get('formula') or '') if r in by_name} & param_row:
                    param_row.add(c['name'])
                    grew = True
        for cl in self.res['calculations']:
            name = cl['name']
            m = self.calc_match.get(name)
            used_vis = [u for u in cl['used_in'] if not u.startswith(('via:', 'dashboard:', 'action:', 'datasource-filter'))]
            if (not m or not m[1]) and name in param_row:
                self.add('Unit', 'CALC', 'Tableau calculation has a Power BI equivalent', name, '', cl['pbi_kind'],
                         'inlined into measures', 'PASS', details='reads a parameter: a column cannot react to a slicer, so its '
                                                                    'formula is inside the measures that use it')
                continue
            if (not m or not m[1]) and cl.get('placeholder'):
                self.add('Unit', 'CALC', 'Tableau calculation has a Power BI equivalent', name, '', cl['pbi_kind'], 'not needed',
                         'SKIP', details=f'constant placeholder ({cl["formula"]}) used only for layout in: '
                                         f'{", ".join(cl["used_in"][:6])}; nothing to rebuild')
                continue
            if not m or not m[1]:
                reason = m[3] if m else 'no measure or column with a similar name, description or logic'
                status = 'SKIP' if m and 'not migrated' in m[3] else 'FAIL' if used_vis else 'WARN'
                self.add('Unit', 'CALC', 'Tableau calculation has a Power BI equivalent', name, '', cl['pbi_kind'], 'not found',
                         status, details=f'{reason}; used in: {", ".join(cl["used_in"][:6])}. Formula: {cl["formula"][:200]}')
                continue
            kind, key, score, method = m
            status = 'PASS' if score >= 0.7 else 'WARN'
            self.add('Unit', 'CALC', 'Tableau calculation has a Power BI equivalent', name, f'{kind} {oid(*key)}', cl['pbi_kind'],
                     kind, status, details=f'match {score:.2f} by {method}' + ('' if status == 'PASS' else
                                                                                  '; confirm, or fix it in test_map.json "calcs"'))
            jg = cl.get('joined_grain')
            if jg:
                self.add('Unit', 'GRAIN', 'Total counts the same rows as Tableau', name, oid(*key),
                         f'{jg["level"]} over rows of {", ".join(jg["tables"])}', 'cannot be checked from files', 'WARN',
                         details='Tableau computes this over joined rows (a FIXED value or a cross-table row value repeats '
                                 'once per detail row); compare the Power BI number with Tableau for one filter selection')
            clo = self.closure(kind, key)
            roots, fns, lits = self.calc_fingerprint(name)
            tab_aggs = {f for f in fns if f in DAX_AGG}
            # COUNTD(IF ... THEN [x] END) translates to COUNTROWS(FILTER(DISTINCT(...))): a distinct count as well
            distinct_rows = 'COUNTROWS' in clo['fns'] and clo['fns'] & {'DISTINCT', 'VALUES', 'SUMMARIZE', 'SUMMARIZECOLUMNS'}
            missing = [f for f in tab_aggs if not (DAX_AGG[f] & clo['fns']) and not (f == 'COUNTD' and distinct_rows)]
            if kind == 'measure' and missing:
                if 'COUNTD' in missing:
                    self.add('Unit', 'AGG', 'Aggregation matches Tableau', name, oid(*key), 'DISTINCTCOUNT (Tableau COUNTD)',
                             ', '.join(sorted(clo['fns'] & set().union(*DAX_AGG.values()))) or 'none', 'WARN',
                             details='Tableau counts distinct values; a SUM/COUNT gives a bigger number when values repeat')
                else:
                    self.add('Unit', 'AGG', 'Aggregation matches Tableau', name, oid(*key), ', '.join(sorted(tab_aggs)),
                             ', '.join(sorted(clo['fns'] & set().union(*DAX_AGG.values()))) or 'none', 'INFO',
                             details='different aggregation function; fine if the logic was moved into columns, check the numbers')
            elif kind == 'measure':
                self.add('Unit', 'AGG', 'Aggregation matches Tableau', name, oid(*key), ', '.join(sorted(tab_aggs)) or '(row level)',
                         ', '.join(sorted(clo['fns'] & set().union(*DAX_AGG.values()))) or '-', 'PASS')
            text = clo['text'] + '\n' + '\n'.join(t.get('m', '') for t in self.model['tables'].values()
                                                  if key[0] == t['name'])
            nums = {float(x) for x in re.findall(r'(?<![\w.])\d+(?:\.\d+)?', text)}
            low = text.lower()
            date_fns = re.search(r'\b(DATEDIFF|DATEADD|EDATE|EOMONTH|DATE|YEAR|MONTH|DAY|QUARTER|WEEKNUM|WEEKDAY)\s*\(', text, re.I)

            def kept(x):
                """DAX writes Tableau literals its own way: #2026-01-01# -> DATE(2026, 1, 1) (01 is 1), 'month' as a
                date unit -> MONTH / EDATE, a week start 'monday' -> a WEEKDAY shift."""
                if x in text or x.lower() in low:
                    return True
                if re.fullmatch(r'\d+(\.\d+)?', x):
                    return float(x) in nums
                if x.lower() in ('day', 'week', 'month', 'quarter', 'year', 'hour', 'minute', 'second'):
                    return bool(date_fns)
                if x.lower() in ('sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'):
                    return 'weekday' in low or 'week' in low
                return False
            missing_lits = [x for x in lits if not kept(x)]
            if lits:
                self.add('Unit', 'LIT', 'Hard-coded values of the Tableau formula are kept', name, oid(*key),
                         ', '.join(sorted(lits))[:300], 'all found' if not missing_lits else 'missing: ' + ', '.join(sorted(missing_lits))[:300],
                         'PASS' if not missing_lits else 'WARN',
                         details='' if not missing_lits else 'a filter value or constant from Tableau is not in the DAX (or the M of its table); '
                                                             'check it was not dropped or changed')
            tf = fmts.get(name) or (self.wb.cols.get(self.calc_keys.get(name)) or {}).get('default_format')
            obj = self.all_meas.get(key) if kind == 'measure' else self.all_cols.get(key)
            pf = (obj or {}).get('formatString') or ''
            if tf and kind == 'measure':
                t_pct = '%' in tf
                p_pct = '%' in pf
                self.add('Unit', 'FMT', 'Number format matches Tableau', name, oid(*key), tf, pf or '(none)',
                         'PASS' if t_pct == p_pct and pf else 'WARN',
                         details='' if t_pct == p_pct and pf else 'percent vs number format differs, or no formatString')

    def u_dax(self):
        cols_any = {n for (_, n) in self.all_cols}
        for (t, n), ms in sorted(self.all_meas.items()):
            e = ms['expression'] or ''
            allowed = {x[1:-1].replace('""', '"') for x in DAX_STR.findall(e)}
            vars_ = {v.lower() for v in re.findall(r'\bVAR\s+([A-Za-z_]\w*)', e, re.I)}
            broken = []
            for rt, rn in dax_refs(e):
                if rt:
                    if rt.lower() in vars_ or rt.startswith('_'):
                        continue
                    if rt not in self.model['tables']:
                        broken.append(oid(rt, rn) + ' (no such table)')
                    elif (rt, rn) not in self.all_cols and (rt, rn) not in self.all_meas and rn not in allowed:
                        broken.append(oid(rt, rn))
                elif rn not in self.meas_by_name and rn not in cols_any and rn not in allowed:
                    broken.append(f'[{rn}]')
            if broken:
                self.add('Unit', 'DAX', 'Measure references resolve', '', oid(t, n), 'every reference exists',
                         ', '.join(broken[:6]), 'FAIL', details='the measure errors in Power BI; fix the reference or the name')
            clean = strip_dax(e)
            # names may hold a "/" of their own ([Avg $/ Str], 'T'[Include/Exclude]): look at the code only
            code = re.sub(r"'(?:[^']|'')*'|\[(?:[^\]]|\]\])*\]", 'n', clean)
            if re.search(r'(?<![/*])/(?![/*])\s*(?![\d\s.]+(?:\)|$|,|\s))', code):
                self.add('Unit', 'DAX', 'Division uses DIVIDE', '', oid(t, n), 'DIVIDE(x, y)', '"/" operator', 'WARN', 'Low',
                         details='"/" errors on a zero denominator where Tableau returns NULL; use DIVIDE (semantic-model rule)')
            if re.search(r'\bIFERROR\s*\(|\bISERROR\s*\(', clean, re.I):
                self.add('Unit', 'DAX', 'No IFERROR / ISERROR', '', oid(t, n), 'DIVIDE or explicit checks', 'IFERROR/ISERROR', 'WARN', 'Low')
            tag = re.search(r'Tableau calc "(.+?)": ', ms.get('description') or '')
            calc_type = next((cv.get('datatype') for cv in self.calc_versions.get(tag.group(1), [])), None) if tag else None
            if not ms['hidden'] and not ms.get('formatString') and not ms.get('dynamic_format') and \
                    calc_type not in ('string', 'boolean') and \
                    not re.match(r'^\s*(SELECTEDVALUE|CONCATENATE|FORMAT|"|IF\s*\(.*")', clean, re.I | re.S):
                # a text or true/false result (an arrow, a label, a show/hide flag) takes no number format
                self.add('Unit', 'DAX', 'Visible measure has a formatString', '', oid(t, n), 'formatString', '(none)', 'WARN', 'Low')
        # a BLANK() placeholder (not translated) and every measure that reads one: Power BI shows a number, Tableau another
        stubs = {k for k, ms in self.all_meas.items() if re.sub(r'\s+', '', ms['expression'] or '') == 'BLANK()'}
        for k in sorted(self.all_meas):
            if k in stubs:
                if k in self.report_refs:
                    self.add('Unit', 'DAX', 'Measure is translated (not a BLANK() placeholder)', '', oid(*k), 'a DAX translation',
                             'BLANK() placeholder', 'FAIL', details='the report shows nothing here; see its description for the draft DAX')
                continue
            dep = sorted(self.closure('measure', k)['measures'] & stubs)
            if dep:
                self.add('Unit', 'DAX', 'Measure does not read a BLANK() placeholder', '', oid(*k), 'no placeholder in its dependencies',
                         'reads ' + ', '.join(oid(*d) for d in dep[:6]), 'FAIL',
                         details='BLANK() + x = x: the number looks plausible but leaves out the placeholder part; fix the placeholder')
        reached = set()
        for k in self.all_meas:
            if k in self.report_refs:
                reached |= self.closure('measure', k)['measures']
        unused = [oid(t, n) for (t, n) in self.all_meas if (t, n) not in self.report_refs and (t, n) not in reached]
        if unused:
            self.add('Unit', 'DAX', 'Measures used by the report', '', f'{len(unused)} measure(s)', 'used on a page or by another measure',
                     'not used: ' + ', '.join(unused[:25]) + (' …' if len(unused) > 25 else ''), 'INFO',
                     details='helpers for later pages are fine; otherwise remove or hide them')

    def u_params(self):
        all_names = [(t, n) for (t, n) in self.all_cols] + [(t, n) for (t, n) in self.all_meas]
        exprs = list(self.model['expressions'])
        for p in self.res['parameters']:
            hits = [oid(t, n) for t, n in all_names if sim(p['name'], n) >= 0.7] + \
                   [oid(None, t) for t in self.model['tables'] if sim(p['name'], t) >= 0.7] + [e for e in exprs if sim(p['name'], e) >= 0.7]
            self.add('Unit', 'PARAM', 'Tableau parameter has a Power BI equivalent', p['name'], ', '.join(hits[:3]),
                     f'default {p.get("value")}', 'found' if hits else 'not found', 'PASS' if hits else 'WARN',
                     details=(p.get('powerbi') or {}).get('pattern') or '' if not hits else
                     'check the default selection on the slicer equals the Tableau default ' + str(p.get('value')))

    def u_relationships(self):
        for r in self.model['relationships']:
            label = f"{oid(r['from_table'], r['from_column'])} -> {oid(r['to_table'], r['to_column'])}"
            a = self.all_cols.get((r['from_table'], r['from_column']))
            b = self.all_cols.get((r['to_table'], r['to_column']))
            if not a or not b:
                self.add('Unit', 'REL', 'Relationship columns exist', '', label, 'both columns exist', 'missing column', 'FAIL')
                continue
            if (a.get('dataType') or '').lower() != (b.get('dataType') or '').lower():
                self.add('Unit', 'REL', 'Relationship key types match', '', label, a.get('dataType'), b.get('dataType'), 'WARN',
                         details='different key types never match rows')
            if r['cross'] == 'bothDirections':
                self.add('Unit', 'REL', 'Single-direction filtering', '', label, 'oneDirection', 'bothDirections', 'WARN', 'Low',
                         details='bidirectional filters can change totals and slow the model; keep only if needed (e.g. RLS)')
            if r['from_card'] == 'many' and r['to_card'] == 'many':
                self.add('Unit', 'REL', 'No many-to-many relationship', '', label, 'many-to-one', 'many-to-many', 'WARN')
        for d in self.res['datasources']:
            conn = d['connection'].get('tds', d['connection'])
            joins = []

            def walk(rel):
                if rel and 'join' in rel:
                    for on in rel['on']:
                        joins.append((rel['join'], on))
                    for x in rel['inputs']:
                        walk(x)
            walk(conn.get('relations'))
            for rl in (conn.get('logical') or {}).get('relationships', []):
                joins.append(('relationship', rl['on']))
            for kind, on in joins:
                pairs = re.findall(r'\[([^\]]+)\]\.\[([^\]]+)\]', on or '')
                tabs = [self.table_match.get(t, []) for t, _ in pairs]
                pt = [t[0] for t in tabs if t]
                if len(pt) == 2 and pt[0] == pt[1]:
                    status, actual = 'PASS', f'tables merged into {pt[0]}'
                elif len(pt) == 2 and any({r['from_table'], r['to_table']} == set(pt) for r in self.model['relationships']):
                    status, actual = 'PASS', f'relationship {pt[0]} - {pt[1]}'
                else:
                    status, actual = 'WARN', 'no relationship between the matched tables'
                self.add('Unit', 'REL', 'Tableau join / relationship exists in the model', f'{d["name"]}: {kind} ON {on}',
                         ', '.join(pt), 'relationship or merged table', actual, status)

    def u_rls(self):
        flags = {f['code'] for f in self.res['flags']}
        tab_rls = bool(flags & {'RLS', 'PER_USER_COLUMN', 'HARDCODED_USERS'})
        roles = self.model['roles']
        user_fn = [r for r in roles if any(re.search(r'USERPRINCIPALNAME|USERNAME|CUSTOMDATA', e, re.I) for e in r['filters'].values())]
        if tab_rls:
            where = ', '.join(sorted({f['where'] for f in self.res['flags'] if f['code'] in ('RLS', 'PER_USER_COLUMN')}))[:300]
            self.add('Unit', 'RLS', 'User-based security is rebuilt as RLS roles', where, ', '.join(r['name'] for r in roles),
                     'a role filtering on USERPRINCIPALNAME()', f'{len(user_fn)} dynamic role(s)' if user_fn else 'none',
                     'PASS' if user_fn else 'FAIL', details='Tableau filters rows by the signed-in user; without a role every viewer sees all rows'
                     if not user_fn else 'test with "View as" in Desktop / the service for a real user')
            if user_fn:
                self.add('Unit', 'RLS', 'Role behaviour checked with real users', '', ', '.join(r['name'] for r in user_fn),
                         'same rows as Tableau for 2-3 users', 'not testable from files', 'MANUAL',
                         details='use Modeling > View as > Other user, compare with Tableau as that user')
        elif roles:
            self.add('Unit', 'RLS', 'Roles without a Tableau counterpart', '', ', '.join(r['name'] for r in roles),
                     'none in the TWB', f'{len(roles)} role(s)', 'INFO', details='security may come from the published data source or Tableau Server')

    def u_date(self):
        uses_parts = any(re.search(r'\b(YEAR|QUARTER|MONTH|WEEK|TRUNC_\w+)\(', f['label'])
                         for w in self.res['worksheets'] for f in w['shelf']['rows'] + w['shelf']['cols'])
        dts = [t['name'] for t in self.model['tables'].values() if t['is_date_table']]
        if uses_parts:
            related = [t for t in dts if any(t in (r['from_table'], r['to_table']) for r in self.model['relationships'])]
            self.add('Unit', 'DATE', 'Date table for date parts', 'YEAR/QUARTER/MONTH on shelves', ', '.join(dts),
                     'a related date table', ', '.join(related) or 'none related' if dts else 'none',
                     'PASS' if related else 'WARN', details='' if related else 'date parts need a date table (marked, related) per the modeling rules')

    # ----- functional tests -----
    def functional(self):
        self.f_validate()
        self.f_refs()
        self.f_pages()
        self.f_visuals()
        self.f_slicers()
        self.f_sheet_filters()
        self.f_visual_model_paths()
        self.f_actions()
        self.f_tooltips()

    def f_validate(self):
        rp = self.proj.get('report_path')
        if not rp or self.case.get('no_validate'):
            return
        try:  # offline: the PBIR rules of powerbi-report-author, schemas bundled in resources\ (no Node.js, no internet)
            data = bv.validate_report(rp)
        except Exception as e:  # never stop the other checks
            self.add('Functional', 'VAL', 'PBIR validates', '', rp, '0 errors', f'validator not run: {e}', 'SKIP')
            return
        errs, warns = data.get('errorCount', 0), data.get('warningCount', 0)
        self.add('Functional', 'VAL', 'PBIR validates (offline validator)', '', os.path.basename(rp), '0 errors',
                 f'{errs} error(s), {warns} warning(s)', 'PASS' if not errs and not warns else 'FAIL' if errs else 'WARN')
        diags = data.get('diagnostics') or data.get('errors') or []
        if isinstance(diags, dict):  # newer CLI: {code: {severity, items: [{message, file}]}}
            diags = [{**it, 'severity': g.get('severity', 'error'), 'code': code}
                     for code, g in diags.items() for it in (g.get('items') or [g])]
        for d in diags[:40]:
            self.add('Functional', 'VAL', 'Validator finding', '', d.get('path') or d.get('file') or '',
                     d.get('code', ''), d.get('message') or json.dumps(d)[:300],
                     'FAIL' if d.get('severity', 'error') == 'error' else 'WARN')

    def f_refs(self):
        tables = self.model['tables']
        if not tables:
            return
        broken = {}
        for p in self.report['pages']:
            items = [(f'{p["name"]} / {v["type"]} "{v["title"] or v["id"]}"', v['all_refs']) for v in p['visuals']]
            items.append((f'{p["name"]} / page filters', [f['field'] for f in p['filters'] if f['field']]))
            for where, refs in items:
                for r in refs:
                    t, n = r['table'], r['name']
                    ok = t in tables and ((t, n) in self.all_cols or (t, n) in self.all_meas or
                                          n in tables[t]['hierarchies'] or r['kind'] == 'Hierarchy' and n in tables[t]['hierarchies'])
                    if not ok:
                        broken.setdefault(where, []).append(oid(t, n))
        for f in self.report['filters']:
            if f['field'] and (f['field']['table'], f['field']['name']) not in self.all_cols and \
                    (f['field']['table'], f['field']['name']) not in self.all_meas:
                broken.setdefault('report filters', []).append(oid(f['field']['table'], f['field']['name']))
        for where, b in broken.items():
            self.add('Functional', 'REF', 'Report fields exist in the model', '', where, 'every field resolves',
                     ', '.join(sorted(set(b))[:8]), 'FAIL', details='the visual shows "Something\'s wrong with one or more fields"')
        if not broken:
            n = sum(len(v['all_refs']) for p in self.report['pages'] for v in p['visuals'])
            self.add('Functional', 'REF', 'Report fields exist in the model', '', f'{n} field references', 'every field resolves',
                     'all resolve', 'PASS')

    def f_pages(self):
        # dashboard tabs only: a hidden worksheet often shares its dashboard's name
        hidden_tabs = {t['name'] for t in self.res['tabs'] if t.get('hidden') and t.get('class', 'dashboard') == 'dashboard'}
        for d in self.res['dashboards']:
            p = self.page_match.get(d['name'])
            if not p:
                self.add('Functional', 'PAGE', 'Dashboard has a report page', d['name'], '', 'a page', 'not found', 'FAIL',
                         details='add the page, or map it in test_map.json "pages" (null = intentionally dropped)')
                continue
            self.add('Functional', 'PAGE', 'Dashboard has a report page', d['name'], p['name'], d['name'], p['name'], 'PASS',
                     details=f'{p["type"]} page' if p['type'] else '')
            if d['name'] in hidden_tabs and not p['hidden'] and p['type'] != 'Drillthrough':
                self.add('Functional', 'PAGE', 'Hidden dashboard is a hidden page', d['name'], p['name'], 'hidden', 'visible', 'WARN', 'Low')
            data_vis = [v for v in p['visuals'] if v['type'] not in NON_DATA and v['type'] not in SLICERS]
            own_actions = [a for a in self.res['actions'] if a.get('kind') in ('FILTER', 'HIGHLIGHT') and
                           (a.get('source') or '').split(' / ')[0].strip() == d['name'] and (a.get('target') or d['name']) == d['name']]
            if len(data_vis) > 1 and not own_actions and not p['interactions']:
                self.add('Functional', 'INTER', 'Visual interactions match Tableau', d['name'], p['name'],
                         'no cross-filtering between visuals (no Tableau filter action on this dashboard)',
                         'Power BI default: every visual filters the others', 'WARN', 'Low',
                         details='set Format > Edit interactions to None where Tableau sheets were independent')
        matched = {id(p) for p in self.page_match.values() if p}
        for p in self.report['pages']:
            if id(p) not in matched:
                self.add('Functional', 'PAGE', 'Extra report page', '', p['name'], '', 'no Tableau dashboard of that name', 'INFO',
                         details=f'{p["type"] or "normal"} page{" (hidden)" if p["hidden"] else ""}')

    def family(self, vtype):
        return next((f for f, s in FAMILY.items() if vtype in s), vtype)

    def f_visuals(self):
        for sp in self.spec['pages']:
            page = self.page_match.get(sp['name'])
            if not page:
                continue
            for ev in sp['visuals']:
                sheet = ev['sheet']
                ws = self.ws.get(sheet) or {}
                v = self.vis_match.get((sp['name'], sheet))
                dims, meas, tips = self.sheet_labels(ws) if ws else ([], [], [])
                if not v:
                    text_only = not dims and not meas
                    role = ws.get('role')
                    if role in ('help', 'decoration', 'text', 'button'):  # set by the extractor: not a data visual
                        self.add('Functional', 'VIS', 'Sheet has a visual on the page', f'{sp["name"]} / {sheet}', page['name'],
                                 ev.get('visualType') or 'a visual', 'not found', 'WARN' if role == 'button' else 'INFO',
                                 details={'help': 'help sheet ("?" icon + tooltip): a textbox/icon with a tooltip page is enough',
                                          'decoration': 'decoration/spacer sheet: a shape or nothing',
                                          'text': 'text/title sheet: may be a textbox or the page title',
                                          'button': 'link/navigation shape: check the page has a button doing the same '
                                                    '(see the ACTION rows)'}[role])
                        continue
                    self.add('Functional', 'VIS', 'Sheet has a visual on the page', f'{sp["name"]} / {sheet}', page['name'],
                             ev.get('visualType') or 'a visual', 'not found', 'WARN' if text_only else 'FAIL',
                             details='text/title sheet: may be a textbox or the page title' if text_only else
                             'add the visual, or map it in test_map.json "visuals" ("Dashboard / Sheet": "visual title or id")')
                    continue
                vname = f'{page["name"]} / {v["type"]} "{v["title"] or v.get("text") or v["id"]}"'
                self.add('Functional', 'VIS', 'Sheet has a visual on the page', f'{sp["name"]} / {sheet}', vname, sheet,
                         v['title'] or v['id'], 'PASS' if v.get('_score', 1) >= 0.45 else 'WARN',
                         details=f'match {v.get("_score", 1):.2f}' + (' (one visual shared by several sheets)' if v.get('_shared') else ''))
                exp_t = ev.get('visualType')
                mark = (ws.get('panes') or [{}])[0].get('mark') if ws else None
                exp_f = self.family(exp_t) if exp_t else MARK_FAMILY.get(mark)
                act_f = self.family(v['type'])
                if exp_f:
                    self.add('Functional', 'VTYPE', 'Visual type matches the Tableau mark', f'{sheet} ({mark})', vname,
                             exp_t or exp_f, v['type'], 'PASS' if exp_f == act_f else 'WARN', 'Low' if exp_f != act_f else '',
                             details='' if exp_f == act_f else 'different chart family; fine if it was a design decision')
                vf = self.visual_fields(v)
                miss_d = [lb for lb in dims if not (set(self.candidates(lb)) & vf)]
                miss_m = [lb for lb in meas if not (set(self.candidates(lb)) & vf)]
                if dims or meas:
                    st = 'PASS' if not miss_d and not miss_m else 'WARN'
                    self.add('Functional', 'FIELDS', 'Visual shows the same fields', sheet, vname,
                             'dims: ' + ', '.join(dims) + ' | measures: ' + ', '.join(meas),
                             'all present' if st == 'PASS' else 'missing: ' + ', '.join(miss_d + miss_m), st,
                             details='' if st == 'PASS' else 'a missing dimension changes the grain; a missing measure changes what is shown. '
                                                             'If the name differs, map it in test_map.json')
                miss_t = [lb for lb in tips if not (set(self.candidates(lb)) & vf)]
                if tips:
                    self.add('Functional', 'TIPS', 'Tooltip fields kept', sheet, vname, ', '.join(tips),
                             'all present' if not miss_t else 'missing: ' + ', '.join(miss_t), 'PASS' if not miss_t else 'INFO')
                t_title = (ev.get('title') or {}).get('text') or ws.get('title')
                if t_title and (ev.get('title') or {}).get('show', True):
                    dyn = '<' in (ws.get('title') or '') and '>' in (ws.get('title') or '')
                    s = sim(t_title, v['title'] or '')
                    ok = s >= 0.5 or (dyn and (v['title'] or '').startswith('['))
                    self.add('Functional', 'TITLE', 'Visual title matches', sheet, vname, t_title, v['title'] or '(no title)',
                             'PASS' if ok else 'WARN', 'Low',
                             details='' if ok else ('Tableau title is dynamic: bind the title to a measure' if dyn else 'title differs'))
                if ws.get('reference_lines'):
                    self.add('Functional', 'REFLINE', 'Reference lines kept', sheet, vname,
                             ', '.join(r['value'] or r['formula'] or '' for r in ws['reference_lines']),
                             'present' if v.get('ref_lines') else 'not found', 'PASS' if v.get('ref_lines') else 'WARN', 'Low')
                if ws.get('sorts') and v['type'] not in ('cardVisual', 'textbox'):
                    # a manual (hand-ordered) Tableau sort is a sort-by column on the model column
                    manual = [m.group(1) for s_ in ws['sorts'] for m in [re.match(r'^\[(.+?)\] manual', s_)] if m]
                    sorted_cols = [n for t, n in self.visual_fields(v) if (self.all_cols.get((t, n)) or {}).get('sortByColumn')
                                   and any(set(self.candidates(lb)) & {(t, n)} for lb in manual)]
                    ok = bool(v.get('sort')) or bool(manual and sorted_cols)
                    act = ', '.join(f'{n} {d}' for n, d in v['sort']) if v.get('sort') else \
                        (f'sort-by column on {", ".join(sorted_cols)}' if sorted_cols else 'default sort')
                    self.add('Functional', 'SORT', 'Sort order set', sheet, vname, '; '.join(ws['sorts'])[:200],
                             act, 'PASS' if ok else 'MANUAL', 'Low',
                             details='' if ok else 'check the order on screen; manual Tableau sorts need a sort column')

    def page_filters(self, page):
        out = list(page['filters']) + list(self.report['filters'])
        for v in page['visuals']:
            out.extend(v['filters'])
        return out

    def f_slicers(self):
        for sp in self.spec['pages']:
            page = self.page_match.get(sp['name'])
            if not page:
                continue
            slicers = [v for v in page['visuals'] if v['type'] in SLICERS]
            for s in sp['slicers']:
                label = s['field']['label']
                cands = set(self.candidates(label))
                if s.get('parameter') and not cands:
                    cands = set(self.candidates('Param:' + s['parameter']))
                # the slicer on exactly this field first: a group slicer and one on its base column can share a page
                hit = next((v for v in slicers if any(f['name'] == s['field'].get('name') for f in v['fields'])
                            and cands & self.visual_fields(v)), None) or \
                    next((v for v in slicers if cands & self.visual_fields(v)), None)
                if not hit:
                    hit = next((v for v in slicers if any(sim(label.replace('Param:', ''), f['name']) >= 0.75 or
                                                          sim(label.replace('Param:', ''), v.get('slicer_header') or '') >= 0.75
                                                          for f in v['fields'])), None)
                if not hit:
                    pane = [f for f in page['filters'] + self.report['filters'] if f['field'] and
                            (f['field']['table'], f['field']['name']) in cands and not f['hidden']]
                    self.add('Functional', 'SLICER', 'Tableau quick filter / parameter control has a slicer', f'{sp["name"]} / {label}',
                             'filter pane' if pane else '', s.get('mode') or 'slicer', 'in the filter pane' if pane else 'not found',
                             'WARN' if pane else 'FAIL', details='' if pane else 'add a slicer, or map the field in test_map.json')
                    continue
                hf = hit['fields'][0] if hit['fields'] else {}
                self.add('Functional', 'SLICER', 'Tableau quick filter / parameter control has a slicer', f'{sp["name"]} / {label}',
                         f'{page["name"]} / slicer {oid(hf.get("table"), hf.get("name"))}', s.get('mode'), hit.get('slicer_mode') or hit['type'], 'PASS')
                exp = (s.get('default') or {}).get('values')
                act = hit.get('slicer_selection') or next((f['values'] for f in hit['filters'] if f['values']), None)
                if exp is not None or act:
                    e = sorted(key_norm(bo.tableau_literal(x)) for x in (exp or []))
                    a = sorted(key_norm(x) for x in (act or []))
                    inv = (s.get('default') or {}).get('inverted')
                    self.add('Functional', 'DEFAULT', 'Slicer default selection matches Tableau', f'{sp["name"]} / {label}',
                             oid(hf.get('table'), hf.get('name')), ('NOT ' if inv else '') + (', '.join(e) or 'all'),
                             ', '.join(a) or 'all', 'PASS' if e == a else 'WARN',
                             details='' if e == a else 'the page opens with different numbers than Tableau; set the slicer selection '
                                                       '(or a bookmark) to the Tableau default')

    def f_sheet_filters(self):
        exposed = set()
        for d in self.res['dashboards']:
            for it in d['items']:
                if it['type'] in ('filter', 'paramctrl') and it.get('field'):
                    exposed.add((it.get('name'), it['field']))
        for sp in self.spec['pages']:
            page = self.page_match.get(sp['name'])
            if not page:
                continue
            pf = self.page_filters(page)
            for ev in sp['visuals']:
                ws = self.ws.get(ev['sheet'])
                if not ws:
                    continue
                v = self.vis_match.get((sp['name'], ev['sheet']))
                vtext = ''
                if v:
                    for t, n in self.visual_fields(v):
                        if (t, n) in self.all_meas:
                            vtext += '\n' + self.closure('measure', (t, n))['text']
                for f in ws['filters']:
                    rule = f['rule']
                    if rule.startswith('ALL') or (ev['sheet'], f['field']) in exposed:
                        continue
                    vals = [x.strip().strip('"') for x in re.findall(r'"[^"]*"|\b\d{4}\b|\btrue\b|\bfalse\b', rule.split('[values')[0])]
                    cands = set(self.candidates(f['field']))
                    calc = self.calc_by_name.get(f['field'])
                    if calc:
                        roots, _, _ = self.calc_fingerprint(calc['name'])
                        for (t, n), col in self.all_cols.items():
                            if str(col.get('source_root') or '').upper() in roots:
                                cands.add((t, n))
                    flt_hit = [x for x in pf if x['field'] and (x['field']['table'], x['field']['name']) in cands]
                    exact = [x for x in flt_hit if not vals or {key_norm(a) for a in vals} <= {key_norm(b) for b in x['values']}
                             or x.get('condition')]
                    tables_m = '\n'.join(self.model['tables'][t].get('m', '') for t, _ in cands if t in self.model['tables'])
                    in_dax = vals and all(('"' + x + '"') in vtext or x in vtext for x in vals)
                    if calc and not in_dax:
                        # a filter on a parameter-driven row calc is folded into the measures as that calc's DAX (the
                        # version of the sheet's own data source when several define the name)
                        shape = self.dax_shape(vtext)
                        in_dax = any(self.dax_shape(self.to_model_refs(cv['dax'])) in shape and
                                     all(f'={x.lower()}()' in shape for x in vals if x.lower() in ('true', 'false'))
                                     for cv in self.calc_versions.get(f['field'], [calc]) if len(cv.get('dax') or '') >= 12)
                    in_m = vals and all(('"' + x + '"') in tables_m for x in vals)
                    # a quick filter on the dashboard that reaches this sheet (Tableau saves its selection into every
                    # sheet it applies to): the page slicer must open with that selection
                    sl_hit = [x for x in page['visuals'] if x['type'] in SLICERS and cands & self.visual_fields(x)]
                    sl_sel = next((x.get('slicer_selection') or next((g['values'] for g in x['filters'] if g['values']), None)
                                   for x in sl_hit), None)
                    if exact:
                        st, act = 'PASS', 'visual/page/report filter'
                    elif in_dax:
                        st, act = 'PASS', 'folded into the visual\'s measures'
                    elif in_m:
                        st, act = 'PASS', 'applied in Power Query'
                    elif sl_hit and vals and sl_sel and {key_norm(a) for a in vals} == {key_norm(b) for b in sl_sel}:
                        st, act = 'PASS', 'slicer on the page opens with the Tableau selection'
                    elif sl_hit:
                        st, act = 'WARN', 'controlled by a slicer on the page, which opens with ' + \
                            (', '.join(sl_sel)[:80] if sl_sel else 'all values') + ' (see the DEFAULT rows)'
                    elif flt_hit:
                        st, act = 'WARN', 'filter on an equivalent field: ' + ', '.join(oid(x['field']['table'], x['field']['name']) for x in flt_hit[:2])
                    else:
                        st, act = ('MANUAL' if calc or not vals else 'WARN'), 'not found'
                    self.add('Functional', 'FILTER', 'Fixed sheet filter is applied', f'{ev["sheet"]}: {f["field"]}' + (' (context)' if f.get('context') else ''),
                             f'{page["name"]}' + (f' / "{v["title"] or v["id"]}"' if v else ''), rule[:200], act, st,
                             'Low' if st == 'PASS' else 'Medium',
                             details='' if st == 'PASS' else 'without this filter the visual shows more rows than Tableau; '
                                                             'add a visual/page filter or fold it into the measure, then re-run')

    def filters_reach(self, src, dst):
        """True when a filter on table src reaches table dst: same table, or a path along relationships in their
        filter direction (one side -> many side; both ways for bothDirections)."""
        if src == dst:
            return True
        nxt = {}
        for r in self.model['relationships']:
            if r.get('active') is False:
                continue
            nxt.setdefault(r['to_table'], set()).add(r['from_table'])
            if r.get('cross') == 'bothDirections':
                nxt.setdefault(r['from_table'], set()).add(r['to_table'])
        seen, todo = {src}, [src]
        while todo:
            for q in nxt.get(todo.pop(), ()):
                if q not in seen:
                    seen.add(q)
                    todo.append(q)
        return dst in seen

    def f_visual_model_paths(self):
        """A visual grouping a measure by a column of another table needs a relationship path from the column's
        table to the tables the measure aggregates. Without it Desktop shows "Can't determine relationships", or
        (wrong direction) every row shows the same total. Tableau's published sources are one joined row set, so
        this never shows there."""
        for sp in self.spec['pages']:
            for ev in sp['visuals']:
                v = self.vis_match.get((sp['name'], ev['sheet']))
                if not v or v['type'] in SLICERS:
                    continue
                fields = [(f['table'], f['name']) for f in v['fields']]
                col_tables = {t for t, n in fields if (t, n) in self.all_cols and t not in self.disconnected_tables()}
                meas_tables = set()
                for k in fields:
                    if k in self.all_meas:
                        meas_tables |= {t for t, _ in self.closure('measure', k)['cols']} - self.disconnected_tables()
                if not col_tables or not meas_tables:
                    continue
                bad = sorted(f'{c} -> {m}' for c in col_tables for m in meas_tables if not self.filters_reach(c, m))
                self.add('Functional', 'REL', 'Visual columns filter its measures', f'{sp["name"]} / {ev["sheet"]}',
                         f'{v["type"]} "{v["title"] or v["id"]}"', 'a relationship path from each column table to the measure tables',
                         'all reached' if not bad else 'no path: ' + ', '.join(bad)[:200], 'PASS' if not bad else 'WARN',
                         details='' if not bad else 'relate the tables (one side -> many side), or base the measure on the '
                                                    "column's table (e.g. count the key on the many side)")

    def disconnected_tables(self):
        """Tables that are disconnected by design: parameter, slice and report-period tables (no relationships)."""
        if not hasattr(self, '_disc'):
            related = {r['from_table'] for r in self.model['relationships']} | {r['to_table'] for r in self.model['relationships']}
            # no relationship and no Power Query source: a calculated parameter / slice / period table
            self._disc = {t for t, tb in self.model['tables'].items() if t not in related and
                          (not (tb.get('m') or '').strip() or re.search(r'Parameter|Report Period|Slices$', t))}
            # parameter-driven dimensions: their value reaches the measures through IN VALUES(), not a relationship
            self._disc |= {pd['name'] for pd in (self.spec.get('model') or {}).get('parameter_dimensions', [])}
        return self._disc

    def f_actions(self):
        pages_by_name = {p['name']: p for p in self.report['pages']}
        for a in self.res['actions']:
            kind = a.get('kind')
            src_db = (a.get('source') or '').split(' / ')[0].strip()
            tgt = (a.get('target') or '').strip()
            tp = self.page_match.get(tgt) if tgt else None
            if kind == 'FILTER' and tgt and tgt != src_db:
                if not tp:
                    self.add('Functional', 'ACTION', 'Filter action becomes drillthrough', a['name'], '', f'drillthrough to {tgt}',
                             'target page not found', 'FAIL')
                    continue
                dt_fields = [f['field'] for f in tp['filters'] if f['field'] and (f.get('how') == 'Drillthrough')]
                if tp['type'] == 'Drillthrough' or dt_fields:
                    spec_p = next((p for p in self.spec['pages'] if p['name'] == tgt), {})
                    want = [f['label'] for f in (spec_p.get('drillthrough') or {}).get('fields', [])]
                    have = {(f['table'], f['name']) for f in dt_fields}
                    got = [lb for lb in want if set(self.candidates(lb)) & have]
                    self.add('Functional', 'ACTION', 'Filter action becomes drillthrough', a['name'], f'{tp["name"]} (drillthrough)',
                             f'drillthrough to {tgt}' + (f' on {", ".join(want[:6])}' if want else ''),
                             'drillthrough on ' + ', '.join(oid(t, n) for t, n in have), 'PASS',
                             details=f'{len(got)} of {len(want)} Tableau action fields are drillthrough fields' if want else '')
                else:
                    nav = [v for p in self.report['pages'] for v in p['visuals'] if v.get('nav') and v['nav'].get('page') == tp['id']]
                    spec_src = next((p for p in self.spec['pages'] if any(nb['target'] == tgt for nb in p.get('nav_buttons', []))), None)
                    if nav and spec_src:  # the source sheets show no dimension: the Tableau click passes no value, only opens the page
                        self.add('Functional', 'ACTION', 'Filter action becomes drillthrough', a['name'], tp['name'],
                                 'navigation (source sheet has no dimension to pass)', 'navigation button', 'PASS')
                        continue
                    self.add('Functional', 'ACTION', 'Filter action becomes drillthrough', a['name'], tp['name'],
                             'drillthrough page', 'navigation button only' if nav else 'not a drillthrough page', 'WARN' if nav else 'FAIL',
                             details='a click in Tableau filters the target; Power BI needs drillthrough fields on the target page')
            elif kind == 'URL':
                url = a.get('url') or ''
                base = re.split(r'[<?]', url)[0]
                web_cols = [oid(t, n) for (t, n), c in self.all_cols.items() if (c.get('dataCategory') or '').lower() == 'weburl']
                in_model = base and any(base in (ms['expression'] or '') for ms in self.all_meas.values()) or \
                    base and any(base in (c.get('expression') or '') or base in self.model['tables'][t].get('m', '')
                                 for (t, n), c in self.all_cols.items())
                buttons = [v for p in self.report['pages'] for v in p['visuals'] if v.get('nav') and v['nav'].get('url')]
                ok = bool(web_cols or in_model or buttons)
                self.add('Functional', 'ACTION', 'URL action has a link', a['name'], ', '.join(web_cols[:3]) or ('URL in the model' if in_model else ''),
                         url[:150], 'link found' if ok else 'not found', 'PASS' if ok else 'FAIL',
                         details='' if ok else 'add a Web URL column (Data category = Web URL) built from the same fields')
            elif kind in ('NAVIGATE', 'GOTO') and tgt:
                nav = [v for p in self.report['pages'] for v in p['visuals'] if v.get('nav') and tp and v['nav'].get('page') == tp['id']] \
                    or [v for p in self.report['pages'] for v in p['visuals'] if v['type'] == 'pageNavigator']
                self.add('Functional', 'ACTION', 'Navigation action has a button', a['name'], tp['name'] if tp else '',
                         f'button to {tgt}', f'{len(nav)} button(s)' if nav else 'not found', 'PASS' if nav else 'WARN')
            elif kind in ('PARAMETER', 'SET'):
                self.add('Functional', 'ACTION', f'{kind.title()} action rebuilt', a['name'], '', 'equivalent behaviour',
                         'cannot be checked from files', 'MANUAL', details='no direct Power BI equivalent; check the redesign by clicking through')
            elif kind == 'FILTER':
                self.add('Functional', 'ACTION', 'Same-dashboard filter action', a['name'], src_db, 'cross-filter between visuals',
                         'Power BI cross-filters by default', 'INFO')
        back = [p['name'] for p in self.report['pages'] if p['type'] == 'Drillthrough' and
                not any(v.get('nav') and (v['nav'].get('type') == 'Back') for v in p['visuals'])]
        for b in back:
            self.add('Functional', 'ACTION', 'Drillthrough page has a back button', '', b, 'back button', 'none', 'WARN', 'Low')

    def f_tooltips(self):
        need = [f['where'] for f in self.res['flags'] if f['code'] == 'VIZ_IN_TOOLTIP']
        tips = [p['name'] for p in self.report['pages'] if p['type'] == 'Tooltip']
        for w in need:
            self.add('Functional', 'TOOLTIP', 'Viz-in-tooltip has a tooltip page', w, ', '.join(tips), 'a Tooltip page',
                     f'{len(tips)} tooltip page(s)' if tips else 'none', 'PASS' if tips else 'WARN')

    # ----- data tests -----
    def data(self):
        dirs = self.case.get('data_dirs') or []
        opts = self.overrides.get('data') or {}
        if self.case.get('no_data'):
            return
        if not dirs:
            self.add('Data', 'SRC', 'Source data available', '', '', 'CSV/XLSX copies of the source tables', 'no data folder',
                     'SKIP', details='put <TABLE>.csv files in a data\\ folder next to the workbook (or pass --data) to get expected numbers')
            return
        today = self.case.get('today') or opts.get('today')
        today = dt.date.fromisoformat(today) if isinstance(today, str) else today
        self.oracle = bo.Oracle(self.wb, dirs, today=today, username=self.case.get('username') or opts.get('username'),
                                keep_empty=opts.get('keep_empty_strings') or [], table_files=opts.get('table_files'))
        o = self.oracle
        for d in self.wb.ds:
            if d == 'Parameters' or self.wb.ds[d]['is_parameters']:
                continue
            cap = self.wb.ds[d]['caption']
            try:
                n = len(o.ds_rows(d))
                self.add('Data', 'SRC', 'Source data loaded', cap, '', 'rows', n, 'PASS' if n else 'WARN',
                         details='; '.join(f'{os.path.basename(p)}: {c} rows' for p, c in o.loaded.items())[:400])
            except bo.Unsupported as e:
                self.add('Data', 'SRC', 'Source data loaded', cap, '', 'data file', str(e), 'SKIP')
        on_dash = {ev['sheet'] for sp in self.spec['pages'] for ev in sp['visuals']}
        exports_t = self.export_files(self.case.get('tableau_exports'))
        exports_p = self.export_files(self.case.get('pbi_exports'))
        for name, el in self.ws_el.items():
            if on_dash and name not in on_dash:
                continue
            t0 = time.time()
            r = o.run_sheet(el)
            secs = time.time() - t0
            if r['status'] == 'skip':
                self.add('Data', 'ORACLE', 'Expected values computed from source data', name, '', 'values', 'not computed', 'SKIP',
                         details=r['reason'])
                continue
            self.add('Data', 'ORACLE', 'Expected values computed from source data', name, '', f'{len(r["measures"])} measure(s)',
                     f'{len(r["rows"])} row(s) from {r["row_count"]} source rows ({secs:.1f}s)',
                     'PASS' if r['status'] == 'ok' else 'WARN', 'Low',
                     details='; '.join((['filters: ' + ' AND '.join(r['filters'])] if r['filters'] else []) +
                                       r['notes'] + [f'{k}: {v}' for k, v in r['measure_errors'].items()])[:900])
            for n in r['notes']:
                self.notes.append(f'{name}: {n}')
            self.record_expected(r)
            dax = self.dax_query(name, r)
            self.dax_queries.append((name, dax))
            vis = next((v for (d, s), v in self.vis_match.items() if s == name and v), None)
            texp = self.find_export(exports_t, [name])
            pexp = self.find_export(exports_p, [name] + ([vis['title']] if vis and vis['title'] else []))
            for m in r['measures']:
                if m in r['measure_errors']:
                    continue
                exp_v = r['total'].get(m)
                pbi = self.candidates(m)
                pbi_txt = oid(*pbi[0]) if pbi else '(no match)'
                status, actual, details = 'MANUAL', '', 'run the DAX in the "DAX Queries" sheet (or read the visual total) and compare'
                if pexp or texp:
                    status, actual, details = None, [], []
                    for label, ex in (('Power BI export', pexp), ('Tableau export', texp)):
                        if not ex:
                            continue
                        text = isinstance(exp_v, str)  # a label calc ("$69.7K"): compare the text, not a number read from it
                        tot = self.export_total(ex, r, m, raw=text)
                        if tot is None:
                            details.append(f'{label}: no total for this measure')
                            continue
                        if text:
                            ok = key_norm(exp_v) == key_norm(tot)
                            actual.append(f'{label}: {tot}')
                        else:
                            ok = num_match(exp_v, tot)
                            actual.append(f'{label}: {tot[0]:g}{"%" if tot[2] else ""}')
                        status = 'FAIL' if not ok else (status or 'PASS')
                    status = status or 'MANUAL'
                    actual, details = '; '.join(actual), '; '.join(details)
                self.add('Data', 'TOTAL', 'Total matches the Tableau logic', f'{name}: {m}', pbi_txt, exp_v, actual, status,
                         details=details)
            for label, ex in (('Power BI export', pexp), ('Tableau export', texp)):
                if ex:
                    self.compare_export(name, r, ex, label)

    def record_expected(self, r):
        for k, vals in r['rows'][:20000]:
            for m, v in vals.items():
                self.expected.append([r['sheet'], ' | '.join(f'{d}={x if x is not None else "(null)"}' for d, x in zip(r['dims'], k)),
                                      m, v if not isinstance(v, bo.Star) else '*'])
        for m, v in r['total'].items():
            self.expected.append([r['sheet'], '(total)', m, v if not isinstance(v, bo.Star) else '*'])
        out = os.path.join(self.case['out'], 'expected_values')
        os.makedirs(out, exist_ok=True)
        fn = re.sub(r'[^\w\- ]', '_', r['sheet'])[:80] + '.csv'
        with open(os.path.join(out, fn), 'w', newline='', encoding='utf-8-sig') as f:
            w = csv.writer(f)
            w.writerow(r['dims'] + r['measures'])
            for k, vals in r['rows']:
                w.writerow(list(k) + [fmt(vals.get(m)) if not isinstance(vals.get(m), bo.Star) else '*' for m in r['measures']])
            w.writerow(['Grand Total'] + [''] * (len(r['dims']) - 1) + [fmt(r['total'].get(m)) for m in r['measures']]
                       if r['dims'] else [fmt(r['total'].get(m)) for m in r['measures']])

    def dax_query(self, sheet, r):
        fn = re.sub(r'[^\w\- ]', '_', sheet)[:80]
        lines = [f'// {sheet}: expected values in expected_values\\{fn}.csv',
                 f'// Tableau filters applied: {" AND ".join(r["filters"]) or "none"}']
        groups, measures, notes = [], [], []
        for d in r['dims']:
            c = self.candidates(d)
            col = next((x for x in c if x in self.all_cols), None)
            if col:
                groups.append(oid(*col))
            else:
                notes.append(f'// no Power BI column for {d}')
        for m in r['measures']:
            c = self.candidates(m)
            ms = next((x for x in c if x in self.all_meas), None)
            fn, inner = label_parts(m)
            col = next((x for x in c if x in self.all_cols), None)
            if ms:
                measures.append(f'"{m}", {oid(None, ms[1])}')
            elif col and fn in DAX_AGG:
                dfn = {'AVG': 'AVERAGE', 'COUNTD': 'DISTINCTCOUNTNOBLANK'}.get(fn, fn)
                measures.append(f'"{m}", {dfn}({oid(*col)})')
            else:
                notes.append(f'// no Power BI measure for {m}')
        # a field twice on one sheet (Profit as a header and as a value, two MIN(0.0) spacers): DAX accepts each group
        # column and each result name once ("specified more than once" / "same name with an existing column")
        groups = list(dict.fromkeys(groups))
        seen, uniq = set(), []
        for x in measures:
            name = x.split('", ', 1)[0].lower()
            if name not in seen:
                seen.add(name)
                uniq.append(x)
        measures = uniq
        self.dax_parts[sheet] = {'groups': list(groups), 'measures': list(measures), 'notes': list(notes),
                                 'filters': list(r['filters']), 'dims': list(r['dims'])}
        body = ',\n    '.join(groups + measures)
        lines += notes + ['EVALUATE', f'SUMMARIZECOLUMNS(\n    {body}\n)' if body else '{ BLANK() }']
        if groups:
            lines.append('ORDER BY ' + ', '.join(groups))
        return '\n'.join(lines)

    def export_files(self, folder):
        out = {}
        if folder and os.path.isdir(folder):
            for fn in os.listdir(folder):
                if fn.lower().endswith(('.csv', '.txt', '.tsv', '.xlsx')) and not fn.startswith('~$'):
                    out[os.path.splitext(fn)[0]] = os.path.join(folder, fn)
        return out

    def find_export(self, files, names):
        best, path = 0, None
        for stem, p in files.items():
            for n in names:
                s = 1.0 if norm(stem) == norm(n) else sim(stem, n)
                if s > best:
                    best, path = s, p
        if best < 0.8:
            return None
        try:
            headers, rows = bo.read_table(path)
        except Exception as e:  # unreadable export: report and continue
            self.notes.append(f'{os.path.basename(path)}: not read ({e})')
            return None
        return {'path': path, 'headers': headers, 'rows': rows}

    def aliases(self, label):
        fn, inner = label_parts(label)
        out = {label, inner}
        if fn in DATE_FN:
            word = {'TRUNC_MONTH': 'Month', 'MONTH_YEAR': 'Month', 'TRUNC_YEAR': 'Year', 'TRUNC_QUARTER': 'Quarter',
                    'TRUNC_DAY': 'Day', 'TRUNC_WEEK': 'Week'}.get(fn, fn.title())
            out |= {f'{word} of {inner}', f'{inner} ({word})', word, f'{inner} {word}', f'{word} {inner}'}
        if fn in AGG_LABEL:
            out |= {f'Sum of {inner}', f'Avg. {inner}', f'Average of {inner}', f'Count of {inner}', f'Distinct count of {inner}',
                    f'Count Distinct of {inner}', f'CNTD({inner})', f'Min. {inner}', f'Max. {inner}', f'Median of {inner}',
                    f'Total {inner}', f'{fn}({inner})'}
        if fn == 'PERCENT_OF_TOTAL':
            f2, i2 = label_parts(inner)
            out |= {f'% of Total {i2}', f'% of Total Sum of {i2}', f'{i2} (% of total)', f'% of total {i2}', f'%GT {i2}'}
        for k, v in (self.res.get('measure_aliases') or {}).items():
            if k.strip('[]') == label:
                out.add(v)
        for t, n in self.candidates(label):
            out.add(n)
        for (_, s), v in self.vis_match.items():
            if v:
                for f in v['fields']:
                    if (f['table'], f['name']) in set(self.candidates(label)) and f.get('display'):
                        out.add(f['display'])
        return {norm(x) for x in out if x}

    def column_for(self, headers, label, taken):
        # the exact name first: norm() drops symbols, so "Delta -" and "Delta +" (or "% Sales" and "Sales") are one key
        exact = str(label).strip().lower()
        for i, h in enumerate(headers):
            if i not in taken and str(h).strip().lower() == exact:
                return i
        al = self.aliases(label)
        for i, h in enumerate(headers):
            if i not in taken and norm(h) in al:
                return i
        digits = re.findall(r'\d+(?:\.\d+)?', str(label))
        for i, h in enumerate(headers):
            if i not in taken and max((sim(h, a) for a in [label, label_parts(label)[1]]), default=0) >= 0.85 \
                    and re.findall(r'\d+(?:\.\d+)?', str(h)) == digits:  # "min(0.2)" is not "min(2.0)"
                return i
        return None

    def export_total(self, ex, r, m, raw=False):
        """The export's total of measure m: parse_num(cell), or the cell text itself (raw=True: a text measure such as
        a "$69.7K" label, which parse_num would read as the number 69.7)."""
        if raw:
            conv = lambda v: None if v is None else str(v).strip()  # noqa: E731
        else:
            conv = parse_num
        taken = set()
        dcols = [self.column_for(ex['headers'], d, taken) for d in r['dims']]
        mi = self.column_for(ex['headers'], m, set(x for x in dcols if x is not None))
        if mi is None:
            if len(r['measures']) == 1:
                nums = [i for i in range(len(ex['headers'])) if i not in dcols and
                        any(parse_num(row[i]) for row in ex['rows'][:20] if i < len(row))]
                mi = nums[-1] if nums else None
            if mi is None:
                return None
        for row in ex['rows']:
            if any(str(c).strip().lower() in ('grand total', 'total') for c in row[:max(1, len(r['dims']))]):
                return conv(row[mi]) if mi < len(row) else None
        if not r['dims'] and ex['rows']:
            return conv(ex['rows'][0][mi])
        return None

    def compare_export(self, sheet, r, ex, label):
        headers, rows = ex['headers'], ex['rows']
        taken = set()
        dcols = []
        for d in r['dims']:
            i = self.column_for(headers, d, taken)
            dcols.append(i)
            if i is not None:
                taken.add(i)
        mcols = {}
        for m in r['measures']:
            i = self.column_for(headers, m, taken)
            if i is not None:
                mcols[m] = i
                taken.add(i)
        fname = os.path.basename(ex['path'])
        if not mcols:
            self.add('Data', 'EXPORT', f'{label} matches expected values', sheet, fname, ', '.join(r['measures']),
                     'no measure column recognised: ' + ', '.join(headers)[:200], 'SKIP',
                     details='rename the export columns to the Tableau field names, or check the export is the right sheet')
            return
        dims_known = [i for i in dcols if i is not None]
        if len(dims_known) < len(r['dims']):
            missing = [d for d, i in zip(r['dims'], dcols) if i is None]
            self.notes.append(f'{sheet} / {fname}: dimension column(s) not found ({", ".join(missing)}); compared at the coarser grain')
        # ratios do not add up, so rows are compared only at the sheet's own grain
        exp = {tuple(key_norm(x) for x in k): vals for k, vals in r['rows']} if len(dims_known) == len(r['dims']) else {}
        last = [''] * len(headers)
        ok = bad = missing = 0
        seen = set()
        for row in rows:
            row = list(row) + [''] * (len(headers) - len(row))
            if any(str(row[i]).strip().lower() in ('grand total', 'total') for i in dims_known):
                continue
            for i in dims_known:  # crosstabs leave repeated headers blank
                if str(row[i]).strip() == '' and last[i]:
                    row[i] = last[i]
                last[i] = row[i]
            key = tuple(key_norm(row[i]) for i in dcols if i is not None)
            if not exp:
                break
            vals = exp.get(key)
            if vals is None:
                missing += 1
                if missing <= 15:
                    self.add('Data', 'EXPORT', f'{label} row exists in expected values', sheet, fname, '(row not expected)',
                             ' | '.join(str(row[i]) for i in dims_known), 'FAIL', 'Medium',
                             details='the report shows a group the Tableau logic does not produce (filter or grain difference)')
                continue
            seen.add(key)
            for m, i in mcols.items():
                shown = parse_num(row[i])
                if num_match(vals.get(m), shown):
                    ok += 1
                else:
                    bad += 1
                    if bad <= 40:
                        self.add('Data', 'EXPORT', f'{label} value matches expected', f'{sheet}: {m}', fname,
                                 fmt(vals.get(m)), str(row[i]), 'FAIL', details=' | '.join(f'{d}={row[c]}' for d, c in zip(r['dims'], dcols) if c is not None))
        not_shown = len([k for k in exp if k not in seen]) if exp else 0
        if not exp:
            self.add('Data', 'EXPORT', f'{label} matches expected values', sheet, fname, 'row-level comparison',
                     'dimension columns not recognised', 'SKIP', details='; '.join(headers)[:300])
            return
        total = ok + bad
        self.add('Data', 'EXPORT', f'{label} matches expected values', sheet, fname, f'{len(exp)} row(s) x {len(mcols)} measure(s)',
                 f'{ok} of {total} values match; {missing} unexpected row(s); {not_shown} expected row(s) not in the export',
                 'PASS' if total and not bad and not missing and not not_shown else 'FAIL' if bad or missing else 'WARN',
                 details='compared measures: ' + ', '.join(mcols))

    # ----- output -----
    def finish(self):
        counters = {}
        self.rows.sort(key=lambda r: ({'Unit': 0, 'Functional': 1, 'Data': 2}[r['Category']]))
        for r in self.rows:
            k = (r['Category'][0], r['Area'])
            counters[k] = counters.get(k, 0) + 1
            r['ID'] = f'{k[0]}-{k[1]}-{counters[k]:03d}'

    def suggested_map(self):
        def obj(k):
            return oid(*k) if k else None
        return {
            'tables': {s: (v[0] if v else None) for s, v in self.table_match.items()},
            'columns': {c['name']: obj((self.col_match.get(c['name']) or [None])[0]) for c in self.res['columns']},
            'calcs': {n: obj(m[1]) for n, m in self.calc_match.items()} |
                     {cl['name']: None for cl in self.res['calculations'] if cl['name'] not in self.calc_match},
            'pages': {d['name']: (self.page_match[d['name']]['name'] if self.page_match.get(d['name']) else None)
                      for d in self.res['dashboards']},
            'visuals': {f'{d} / {s}': (v['title'] or v['id']) if v else None for (d, s), v in self.vis_match.items()},
            'data': {'username': None, 'today': None, 'keep_empty_strings': [], 'table_files': {}},
            '_how_to_use': 'copy to test_map.json next to the workbook, correct wrong matches, set null for objects '
                           'intentionally not migrated, and re-run',
        }

    def summary_rows(self):
        cats = ['Unit', 'Functional', 'Data']
        stats = ['PASS', 'FAIL', 'WARN', 'MANUAL', 'SKIP', 'INFO']
        table = [[c] + [sum(1 for r in self.rows if r['Category'] == c and r['Status'] == s) for s in stats] for c in cats]
        table.append(['Total'] + [sum(t[i + 1] for t in table) for i in range(len(stats))])
        fails_high = sum(1 for r in self.rows if r['Status'] == 'FAIL' and r['Severity'] == 'High')
        fails = sum(1 for r in self.rows if r['Status'] == 'FAIL')
        warns = sum(1 for r in self.rows if r['Status'] == 'WARN')
        verdict = 'FAIL' if fails_high else 'FAIL (minor)' if fails else 'PASS with warnings' if warns else 'PASS'
        return table, stats, verdict

    def write(self):
        out = self.case['out']
        os.makedirs(out, exist_ok=True)
        self.finish()
        table, stats, verdict = self.summary_rows()
        w = self.res['workbook']
        info = [
            ['Verdict', verdict],
            ['Tableau workbook', self.case['twb']],
            ['Power BI project', self.case['pbip']],
            ['Semantic model', f'{self.proj.get("model_path")} ({self.model.get("format", "-")})'],
            ['Report', f'{self.proj.get("report_path")} ({self.report.get("format", "-")})'],
            ['Published data sources merged', ', '.join(w.get('tds_merged') or []) or 'none'],
            ['Data folders', '; '.join(self.case.get('data_dirs') or []) or 'none (data tests skipped)'],
            ['Tableau exports', self.case.get('tableau_exports') or 'none'],
            ['Power BI exports', self.case.get('pbi_exports') or 'none'],
            ['Test map', self.case.get('test_map_path') or 'none (automatic matching only)'],
            ['Tableau scope', f'{w["dashboards"]} dashboards, {w["worksheets"]} sheets, {w["calcs_used"]} calcs used, '
                              f'{w["parameters_used"]} parameters, {w["actions"]} actions'],
            ['Power BI scope', f'{len(self.report["pages"])} pages, {sum(len(p["visuals"]) for p in self.report["pages"])} visuals, '
                               f'{len(self.model["tables"])} tables, {len(self.all_meas)} measures, {len(self.model["roles"])} roles'],
            ['Run', f'{dt.datetime.now():%Y-%m-%d %H:%M} in {time.time() - self.t0:.1f}s'],
            ['', ''],
            ['How to read', 'FAIL = must fix; WARN = check; MANUAL = needs a person (Desktop, screenshots, users); '
                            'SKIP = not testable with the inputs given; INFO = for awareness'],
            ['Matching', 'Tableau and Power BI objects are matched by name, description, logic and position. The Mapping sheet '
                         'shows every match; correct wrong ones in test_map.json (template: test_map_suggested.json)'],
        ] + [['Note', n] for n in sorted(set(self.notes))[:60]]
        summary = [['Category'] + stats] + table
        cols = COLUMNS
        as_list = lambda rs: [[r[c] for c in cols] for r in rs]
        issues = sorted([r for r in self.rows if r['Status'] in ('FAIL', 'WARN', 'MANUAL')],
                        key=lambda r: (SEV[r['Status']], {'High': 0, 'Medium': 1, 'Low': 2}.get(r['Severity'], 3), r['ID']))
        by_cat = {c: [r for r in self.rows if r['Category'] == c] for c in ('Unit', 'Functional', 'Data')}
        sheets = [('Summary', ['Item', 'Value'], info + [['', '']] + [[r[0], ' '.join(f'{s}={v}' for s, v in zip(stats, r[1:]))]
                                                                     for r in table], [30, 120]),
                  ('Counts', summary[0], summary[1:]),
                  ('Issues', cols, as_list(issues)),
                  ('Unit Tests', cols, as_list(by_cat['Unit'])),
                  ('Functional Tests', cols, as_list(by_cat['Functional'])),
                  ('Data Tests', cols, as_list(by_cat['Data'])),
                  ('Mapping', ['Type', 'Tableau', 'Power BI', 'Score', 'Method'], self.mapping),
                  ('Expected Values', ['Sheet', 'Group', 'Measure', 'Expected'], self.expected[:60000]),
                  ('DAX Queries', ['Sheet', 'DAX query'], [[s, q] for s, q in self.dax_queries], [30, 120])]
        path = os.path.join(out, 'test_report.xlsx')
        try:
            write_xlsx(path, sheets)
        except PermissionError:
            path = os.path.join(out, f'test_report_{dt.datetime.now():%H%M%S}.xlsx')
            write_xlsx(path, sheets)
            self.notes.append('test_report.xlsx was open in Excel; wrote ' + os.path.basename(path))
        for name, hdr, rows in [('unit_tests.csv', cols, as_list(by_cat['Unit'])),
                                ('functional_tests.csv', cols, as_list(by_cat['Functional'])),
                                ('data_tests.csv', cols, as_list(by_cat['Data'])),
                                ('issues.csv', cols, as_list(issues)),
                                ('mapping.csv', ['Type', 'Tableau', 'Power BI', 'Score', 'Method'], self.mapping)]:
            try:
                with open(os.path.join(out, name), 'w', newline='', encoding='utf-8-sig') as f:
                    wr = csv.writer(f)
                    wr.writerow(hdr)
                    wr.writerows(rows)
            except PermissionError:
                self.notes.append(f'{name} is open in another program; not rewritten')
        with open(os.path.join(out, 'dax_queries.dax'), 'w', encoding='utf-8') as f:
            f.write('\n\n'.join(q for _, q in self.dax_queries))
        with open(os.path.join(out, 'test_map_suggested.json'), 'w', encoding='utf-8') as f:
            json.dump(self.suggested_map(), f, indent=1, ensure_ascii=False)
        md = [f'# Bhayankar test: {self.case["name"]}', '', f'**Verdict: {verdict}**', '',
              '| Category | ' + ' | '.join(stats) + ' |', '|---|' + '---|' * len(stats)]
        md += ['| ' + ' | '.join(str(x) for x in r) + ' |' for r in table]
        md += ['', '## Top issues', '']
        md += [f'- **{r["Status"]}** {r["ID"]} {r["Test"]}: {r["Tableau object"]} → {r["Power BI object"]} '
               f'(expected {r["Expected"][:80]}, got {r["Actual"][:80]})' for r in issues[:40]]
        with open(os.path.join(out, 'test_summary.md'), 'w', encoding='utf-8') as f:
            f.write('\n'.join(md) + '\n')
        return {'workbook': os.path.basename(self.case['twb']), 'pbip': os.path.basename(self.case['pbip']), 'verdict': verdict,
                **{f'{c.lower()}_{s.lower()}': next(t for t in table if t[0] == c)[stats.index(s) + 1]
                   for c in ('Unit', 'Functional', 'Data') for s in ('PASS', 'FAIL', 'WARN')},
                'report': path}

    def run(self):
        self.load()
        self.match()
        if self.model['tables']:
            self.unit()
        if self.report['pages']:
            self.functional()
        self.data()
        return self.write()


# ---------- inputs ----------
def find_pbip(folder, name):
    best = None
    for dp, dirs, fs in walk(folder):
        for fn in fs:
            if fn.lower().endswith('.pbip'):
                s = 1.0 if norm(os.path.splitext(fn)[0]) == norm(name) else sim(os.path.splitext(fn)[0], name)
                if s >= 0.9 and (not best or s > best[0]):
                    best = (s, os.path.join(dp, fn))
    return best[1] if best else None


def walk(base, depth=3):
    """os.walk limited to `depth` levels below base, skipping PBIP internals and outputs."""
    base = os.path.abspath(base)
    for dp, dirs, fs in os.walk(base):
        dirs[:] = [d for d in dirs if not d.lower().endswith(('.semanticmodel', '.report')) and not d.startswith('.')
                   and d.lower() not in ('test_output', 'output', '__pycache__')]
        if dp[len(base):].count(os.sep) >= depth:
            dirs[:] = []
        yield dp, dirs, fs


def data_dirs_for(base, name, proj_path, wb_tables):
    """Data folder candidates: data\\, <Name>_Data\\, any sub-folder with CSVs named like source tables,
    then text parameters of the PBIP that point at an existing folder of CSVs (sample-data folders)."""
    out = []
    wanted = {norm(t) for t in wb_tables}
    for dp, dirs, fs in walk(base):
        dirs[:] = [d for d in dirs if d.lower() not in ('tableau_exports', 'pbi_exports', 'powerbi_exports')]
        stems = {norm(os.path.splitext(f)[0]) for f in fs if f.lower().endswith(('.csv', '.xlsx', '.txt'))}
        tail = os.path.basename(dp).lower()
        if stems and (stems & wanted or tail in ('data', 'csv', norm(name) + 'data') or norm(tail) == norm(name + ' data')):
            out.append(dp)
    if not out and proj_path:
        try:
            proj = bp.load_project(proj_path)
            for e in (proj['model'] or {}).get('expressions', {}).values():
                v = e.get('value')
                if isinstance(v, str) and os.path.isdir(v) and any(f.lower().endswith('.csv') for f in os.listdir(v)):
                    out.append(v)
        except Exception:  # a broken project is reported by the tests themselves
            pass
    return list(dict.fromkeys(out))


def load_json(path):
    if path and os.path.exists(path):
        with open(path, encoding='utf-8-sig') as f:
            return json.load(f)
    return None


def first_dir(base, names):
    for n in names:
        p = os.path.join(base, n)
        if os.path.isdir(p):
            return p
    return None


def build_case(twb, pbip, base, out_root, args, explicit=False):
    name = os.path.splitext(os.path.basename(twb))[0]
    root_tables = []
    try:
        r, _ = bx.load_root(twb)
        for t in r.iter('relation'):
            if t.get('table'):
                root_tables.append(t.get('table').replace('[', '').replace(']', '').split('.')[-1])
            if t.get('name'):
                root_tables.append(t.get('name'))
    except (Exception, SystemExit):
        pass
    map_path = args.map or next((p for p in (os.path.join(base, f'{name}.test_map.json'), os.path.join(base, 'test_map.json'))
                                 if os.path.exists(p)), None)
    tmap_path = next((p for p in (os.path.join(base, 'table_map.json'), os.path.join(INPUT_DIR, 'table_map.json'))
                      if os.path.exists(p)), None)
    tds = list(args.tds or [])
    if not tds:
        tds = [os.path.join(dp, f) for dp, _, fs in walk(base, 2) for f in fs if f.lower().endswith(('.tds', '.tdsx'))]
    data = list(args.data or []) or (data_dirs_for(base, name, pbip, root_tables) if base else [])
    return {'name': name, 'twb': os.path.abspath(twb), 'pbip': os.path.abspath(pbip) if pbip else None,
            'tds': tds, 'strict_tds': not (explicit and args.tds), 'data_dirs': data,
            'tableau_exports': args.tableau_exports or (first_dir(base, ['tableau_exports', 'Tableau Exports', 'tableau exports']) if base else None),
            'pbi_exports': args.pbi_exports or (first_dir(base, ['pbi_exports', 'powerbi_exports', 'Power BI Exports', 'pbi exports']) if base else None),
            'test_map': load_json(map_path) or {}, 'test_map_path': map_path, 'table_map': load_json(tmap_path) or {},
            'today': args.today, 'username': args.username, 'no_data': args.no_data, 'no_validate': args.no_validate,
            'out': os.path.join(out_root, name)}


def run_case(case):
    print(f'  testing {case["name"]} ...', flush=True)
    t = Tester(case)
    row = t.run()
    print(f'  {row["verdict"]:<18} {case["name"]}: unit {row["unit_pass"]} pass / {row["unit_fail"]} fail / {row["unit_warn"]} warn, '
          f'functional {row["functional_pass"]}/{row["functional_fail"]}/{row["functional_warn"]}, '
          f'data {row["data_pass"]}/{row["data_fail"]}/{row["data_warn"]} -> {row["report"]}')
    return row


def main():
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('twb', nargs='?', help=f'workbook, or a folder to scan (default: {INPUT_DIR})')
    ap.add_argument('pbip', nargs='?', help='.pbip file or project folder (when a workbook is given)')
    ap.add_argument('--data', action='append', help='folder of source tables as CSV/XLSX (repeatable)')
    ap.add_argument('--tds', action='append', help='published data source .tds/.tdsx (repeatable)')
    ap.add_argument('--tableau-exports', help='folder of Tableau sheet exports (CSV/XLSX named after the sheet)')
    ap.add_argument('--pbi-exports', help='folder of Power BI "Export data" files (named after the sheet or visual title)')
    ap.add_argument('--map', help='test_map.json with matching corrections and data options')
    ap.add_argument('--today', help='date used for TODAY() and relative date filters (YYYY-MM-DD; default: today)')
    ap.add_argument('--username', help='login used for USERNAME() in Tableau calcs (data tests only)')
    ap.add_argument('--out', help=f'output folder (default: {OUTPUT_DIR})')
    ap.add_argument('--no-data', action='store_true', help='skip the data tests')
    ap.add_argument('--no-validate', action='store_true', help='skip the PBIR validation (offline validator)')
    a = ap.parse_args()
    out = a.out or OUTPUT_DIR
    target = a.twb or INPUT_DIR
    rows = []
    if os.path.isdir(target):
        os.makedirs(target, exist_ok=True)
        print(f'Bhayankar tester: scanning {target}')
        twbs = [os.path.join(dp, f) for dp, dirs, fs in os.walk(target) for f in sorted(fs)
                if f.lower().endswith(('.twb', '.twbx')) and not f.startswith(('~$', '.'))]
        if not twbs:
            print(f'No .twb/.twbx found. Put each report in {target}\\<Name>\\ (workbook + PBIP with the same name) and run again.')
            return
        for twb in twbs:
            name = os.path.splitext(os.path.basename(twb))[0]
            base = os.path.dirname(twb)
            pbip = find_pbip(base, name) or find_pbip(target, name)
            if not pbip:
                print(f'  ERR {name}: no {name}.pbip found next to the workbook (the PBIP must have the same name)')
                rows.append({'workbook': os.path.basename(twb), 'verdict': 'ERROR: no PBIP with the same name'})
                continue
            case = build_case(twb, pbip, base if os.path.abspath(base) != os.path.abspath(target) else target, out, a)
            try:
                rows.append(run_case(case))
            except (Exception, SystemExit) as e:  # keep the batch going
                import traceback
                traceback.print_exc()
                rows.append({'workbook': os.path.basename(twb), 'verdict': f'ERROR {e}'})
        os.makedirs(out, exist_ok=True)
        keys = ['workbook', 'pbip', 'verdict'] + [f'{c}_{s}' for c in ('unit', 'functional', 'data') for s in ('pass', 'fail', 'warn')] + ['report']
        with open(os.path.join(out, 'test_inventory.csv'), 'w', newline='', encoding='utf-8-sig') as f:
            wr = csv.DictWriter(f, keys, extrasaction='ignore')
            wr.writeheader()
            wr.writerows(rows)
        print(f'{len(rows)} report(s) tested -> {out} (summary: test_inventory.csv)')
        return
    if not a.pbip:
        ap.error('give the PBIP too: python bhayankar_test.py <workbook> <project.pbip>')
    run_case(build_case(target, a.pbip, os.path.dirname(os.path.abspath(target)), out, a, explicit=True))


if __name__ == '__main__':
    main()
