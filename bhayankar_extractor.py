#!/usr/bin/env python3
r"""Bhayankar extractor: token-lean Tableau workbook extractor for the Power BI rebuild.

Reads a .twb/.twbx (optionally plus the published data source .tds/.tdsx) and
writes a compact digest that contains only what changes a number, a filter
result, a relationship, navigation, a visual's type/encoding, or security.
Everything else (thumbnails, metadata-records for unused columns, window state,
zone styling, repeated datasource-dependencies, orphan calcs) is dropped.

Usage (from the toolkit folder):
  python bhayankar_extractor.py              # scan input\ and write results to output\  (the normal way)
  python bhayankar_extractor.py <workbook.twb|.twbx> [--tds <published.tds|.tdsx>]   # one workbook -> output\<name>
  python bhayankar_extractor.py <workbook> --show "<sheet|dashboard|calc|parameter|action name>"
  python bhayankar_extractor.py --refresh-catalog  # maintenance only: re-pull visual roles (needs the Node.js CLI)

input\ : .twb/.twbx workbooks, the published data sources they use as .tds/.tdsx (matched by data source
         name), and optionally table_map.json {"SOURCE_TABLE": "Power BI table"}
output\: one folder per workbook, plus inventory.csv (one row per workbook):
  bhayankar_digest.md       read this, not the TWB: logic + Power BI model + rebuild plan per page
  bhayankar_digest_NN_*.md  per-dashboard files when the digest is large (split mode)
  bhayankar_digest.json     extracted Tableau logic, structured
  bhayankar_rebuild.json    Power BI rebuild spec (model, pages, visuals, slicers, drillthrough) for generators
  bhayankar_measures.dax    draft measures as a DAX query view script (DEFINE MEASURE ... EVALUATE)
Modules (keep together): bhayankar_dax.py (Tableau -> DAX), bhayankar_report.py (rebuild spec),
bhayankar_visual_catalog.json (roles snapshot from powerbi-report-author).

Python 3 standard library only.
"""
import argparse
import csv
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bhayankar_dax import Translator, group_column, parameter_pattern  # noqa: E402
from bhayankar_report import Rebuild, render_model, render_page, dax_script  # noqa: E402

U = '{http://www.tableausoftware.com/xml/user}'
QREF = re.compile(r'\[([^\[\]]+)\]\.\[([^\[\]]+)\]')
ANYREF = re.compile(r'(?:\[([^\[\]]+)\]\.)?\[([^\[\]]+)\]')
INST = re.compile(r'^(\w*):(.+?)(?::(\w+))?$')

AGG = {'sum': 'SUM', 'avg': 'AVG', 'cnt': 'COUNT', 'ctd': 'COUNTD', 'min': 'MIN', 'max': 'MAX',
       'attr': 'ATTR', 'med': 'MEDIAN', 'stdev': 'STDEV', 'var': 'VAR', 'cntd': 'COUNTD'}
DATEPART = {'yr': 'YEAR', 'qr': 'QUARTER', 'mn': 'MONTH', 'wk': 'WEEK', 'dy': 'DAY', 'hr': 'HOUR',
            'mi': 'MINUTE', 'wd': 'WEEKDAY', 'tyr': 'TRUNC_YEAR', 'tqr': 'TRUNC_QUARTER',
            'tmn': 'TRUNC_MONTH', 'twk': 'TRUNC_WEEK', 'tdy': 'TRUNC_DAY', 'md': 'MONTH_DAY',
            'my': 'MONTH_YEAR', 'mdy': 'DATE'}
QUICK_TC = {'pcto': 'PERCENT_OF_TOTAL', 'pcdf': 'PERCENT_DIFFERENCE', 'diff': 'DIFFERENCE', 'cum': 'RUNNING_TOTAL',
            'win': 'WINDOW', 'rank': 'RANK', 'rank_dense': 'RANK_DENSE', 'rank_modified': 'RANK_MODIFIED',
            'rank_unique': 'RANK_UNIQUE', 'pctrank': 'PERCENTILE', 'mavg': 'MOVING_AVG', 'ytd': 'YTD_TOTAL',
            'ytdgr': 'YTD_GROWTH', 'cgr': 'COMPOUND_GROWTH'}
FIELD_TYPES = {'nk', 'ok', 'qk'}
# date level of a blend link (Tableau links a date at the level used in the view, or the one set in Edit Blend Relationships)
BLEND_LEVEL = {'tyr': 'year', 'yr': 'year', 'tqr': 'quarter', 'qr': 'quarter', 'tmn': 'month', 'mn': 'month', 'my': 'month',
               'twk': 'week', 'wk': 'week', 'tdy': 'day', 'dy': 'day', 'mdy': 'day'}
PER_USER = re.compile(r"USERNAME\(\)\s*!?=\s*\[|\]\s*!?=\s*USERNAME\(\)", re.I)
USER_LITERAL = re.compile(r"USERNAME\(\)\s*!?=\s*['\"]([^'\"]+)['\"]", re.I)
OBJECT_ID = re.compile(r'\[[^\[\]]+\]\.\[__tableau_internal_object_id__\]\.\[')
ROWCOUNT = re.compile(r'^(.+?)(?: \([^)]*\))?_[0-9A-F]{32}$')


def parse_instance(inner):
    """'pcto:sum:Sales:qk:2' -> (['pcto', 'sum'], 'Sales'). Returns (prefixes, name) or None for plain names."""
    parts = inner.split(':')
    if len(parts) < 2:
        return None
    while len(parts) > 1 and (parts[-1].isdigit() or parts[-1] in FIELD_TYPES or parts[-1].startswith('vt')):
        parts.pop()
    pre = []
    known = set(AGG) | set(DATEPART) | set(QUICK_TC) | {'none', 'usr', 'io', ''}
    while len(parts) > 1 and parts[0] in known:
        pre.append(parts.pop(0))
    return (pre, ':'.join(parts)) if pre else None


TABLE_CALC = re.compile(r'\b(RUNNING_\w+|WINDOW_\w+|RANK\w*|INDEX|LOOKUP|PREVIOUS_VALUE|TOTAL|FIRST|LAST|SIZE)\s*\(', re.I)
RLS_FUNCS = re.compile(r'\b(USERNAME|FULLNAME|ISMEMBEROF|USERDOMAIN|ISUSERNAME|ISFULLNAME|USERATTRIBUTE\w*)\s*\(', re.I)
RISKY_FUNCS = re.compile(r'\b(RAWSQL\w*|SCRIPT_\w+|REGEXP_\w+|MODEL_\w+|SPLIT|MAKEPOINT|MAKELINE|HEXBIN\w)\s*\(', re.I)
DATE_TYPES = {'date', 'datetime'}
GRAIN_AGGS = re.compile(r'\b(SUM|AVG|COUNT|MEDIAN|STDEVP?|VARP?|PERCENTILE)\s*\(', re.I)  # COUNTD/MIN/MAX ignore duplicates

# XML the extractor reads, or skips on purpose (styling, caches, ids). Anything else is reported as UNHANDLED_XML,
# so a Tableau feature this tool does not know about shows up in the digest instead of being dropped silently.
KNOWN_XML = {
    'zone': {'layout-basic', 'layout-flow', None, 'title', 'filter', 'empty', 'color', 'size', 'shape', 'map', 'bitmap',
             'paramctrl', 'text', 'dashboard-object', 'button', 'navigation', 'web', 'flipboard', 'flipboard-nav'},
    'groupfilter': {'member', 'range', 'level-members', 'union', 'intersection', 'crossjoin', 'except', 'filter', 'end',
                    'order', 'empty-level'},
    'filter': {'categorical', 'quantitative', 'relative-date'},
    'mark': {'Automatic', 'Bar', 'Line', 'Area', 'Square', 'Circle', 'Shape', 'Text', 'Map', 'Pie', 'GanttBar',
             'Polygon', 'Multipolygon', 'Density', 'VizExtension'},
    'actions': {'action', 'edit-parameter-action', 'nav-action', 'datasources', 'datasource-dependencies'},
    'table': {'view', 'style', 'panes', 'rows', 'cols', 'subtotals', 'show-full-range', 'join-lod-include-overrides',
              'tooltip-style', 'mark-labels', 'forecast-specification'},
    'view': {'datasources', 'datasource-dependencies', 'filter', 'slices', 'aggregation', 'manual-sort', 'computed-sort',
             'natural-sort', 'sort', 'hide-sort-controls', 'shelf-sorts', 'mapsources', 'single-value-per-nest-shelf-sorts'},
    'pane': {'view', 'mark', 'mark-sizing', 'encodings', 'reference-line', 'customized-tooltip', 'customized-label',
             'style', 'trendline'},
    'datasource': {'aliases', 'column', 'connection', 'column-instance', 'group', 'extract', 'layout', 'style',
                   'semantic-values', 'datasource-dependencies', 'object-graph', 'drill-paths', 'date-options',
                   'field-sort-info', 'folder', 'default-sorts', 'repository-location', 'overridable-settings',
                   'folders-common', 'data-source-tree', 'filter', 'calculations', 'metadata-records'},
    'table-calc': {'Rows', 'Columns', 'Field', 'CellInPane', 'ColumnInPane', 'RowInPane', 'PaneCol', 'PaneRow', 'Table',
                   'Pane', 'Cell', None},
}


FCP = re.compile(r'_\.fcp\.\w+\.(true|false)\.\.\.')


def normalize_fcp(root):
    """Tableau 2020.2+ writes feature-flagged names such as '_.fcp.ObjectModelEncapsulateLegacy.true...relation'.
    Keep the 'true' variant under its plain name and drop the 'false' legacy fallback."""
    for parent in root.iter():
        for ch in list(parent):
            if ch.tag.startswith('_.fcp.'):
                if '.false...' in ch.tag:
                    parent.remove(ch)
                    continue
                ch.tag = FCP.sub('', ch.tag)
        for k in [k for k in parent.attrib if k.startswith('_.fcp.')]:
            v = parent.attrib.pop(k)
            if '.false...' not in k:
                parent.attrib[FCP.sub('', k)] = v
    return root


RECOVERED = {}  # path -> description, for files parsed in recovery mode


def parse_recover(data, label):
    """Parse XML; if it is truncated or corrupt, keep every element read before the error.
    Unfinished elements are kept with the children they already have."""
    try:
        return ET.fromstring(data)
    except ET.ParseError as e:
        err = e
    p = ET.XMLPullParser(events=('start',))
    root = None
    try:
        for i in range(0, len(data), 1 << 20):
            p.feed(data[i:i + (1 << 20)])
            for _, el in p.read_events():
                if root is None:
                    root = el
    except ET.ParseError:
        for _, el in p.read_events():
            if root is None:
                root = el
    if root is None:
        raise SystemExit(f'{label}: XML parse error ({err}) before any content; re-download the file')
    RECOVERED[label] = str(err)
    return root


def load_root(path, want=('.twb',)):
    """Parse a TWB/TDS, unzipping .twbx/.tdsx. Returns (root, packaged file list)."""
    if zipfile.is_zipfile(path):
        with zipfile.ZipFile(path) as z:
            inner = [n for n in z.namelist() if n.lower().endswith(want)]
            if not inner:
                raise SystemExit(f'No {want} inside {path}')
            assets = [(n, z.getinfo(n).file_size) for n in z.namelist() if n not in inner]
            return normalize_fcp(parse_recover(z.read(inner[0]), path)), assets
    with open(path, 'rb') as f:
        return normalize_fcp(parse_recover(f.read(), path)), []


def ua(e, k):
    return e.get(U + k)


def clean_formula(f):
    if f is None:
        return None
    lines = [re.sub(r'//.*$', '', ln) for ln in f.replace('\r', '').split('\n')]
    return re.sub(r'\s+', ' ', ' '.join(lines)).strip()


def strip_brackets(n):
    return n[1:-1] if n and n.startswith('[') and n.endswith(']') else n


AGG_IN_EXPR = re.compile(r'\b(SUM|AVG|COUNT|COUNTD|MIN|MAX|ATTR|MEDIAN)\s*\(', re.I)


def set_formula(g):
    """Tableau set (<group> with a <groupfilter>) -> an equivalent row-level Tableau formula (TRUE = IN), or None when
    the rule is not row-level (a condition on an aggregate per member, ranges): those stay a set flag for a person.
    TOP/BOTTOM n ordered by the set's own field uses __SET_TOPN, which bhayankar_dax turns into IN TOPN(...)."""
    def level(e):  # '[none:StartDt:qk]' -> '[StartDt]'
        lv = e.get('level')
        for k in e.iter('groupfilter'):
            lv = lv or k.get('level')
        if not lv:
            return None
        p = parse_instance(strip_brackets(lv))
        return f'[{p[1] if p else strip_brackets(lv)}]'

    def member(m):
        return 'NULL' if m in ('%null%', None) else m

    e = g.find('./groupfilter')
    if e is None:
        return None
    f, lv = e.get('function'), level(e)
    if not lv:
        return None
    if f == 'member':
        m = member(e.get('member'))
        return f'ISNULL({lv})' if m == 'NULL' else f'{lv} = {m}'
    if f == 'union' and all(k.get('function') == 'member' for k in e.findall('./groupfilter')):
        ms = [member(k.get('member')) for k in e.findall('./groupfilter')]
        return ' OR '.join(f'ISNULL({lv})' if m == 'NULL' else f'{lv} = {m}' for m in ms)
    if f == 'level-members':
        return 'TRUE'
    if f == 'except':
        kids = e.findall('./groupfilter')
        if len(kids) == 2 and kids[0].get('function') == 'level-members' and kids[1].get('function') == 'member':
            m = member(kids[1].get('member'))
            return f'NOT ISNULL({lv})' if m == 'NULL' else f'{lv} <> {m}'
        return None
    if f == 'filter':
        x = e.get('expression') or ''
        if AGG_IN_EXPR.search(re.sub(r'\{[^{}]*\}', '', x)):  # SUM([Sales]) > 100 per member: not a row-level test
            return None
        return f'({x})'
    if f == 'end':
        order = e.find('./groupfilter')
        if order is None or order.get('function') != 'order':
            return None
        ox = (order.get('expression') or '').replace(' ', '')
        if ox.upper() not in (f'MAX({lv})'.upper(), f'MIN({lv})'.upper()):
            return None  # ordered by another measure: a TOPN over a measure, left to a person
        desc = (order.get('direction') or 'DESC').upper() == 'DESC'
        if (e.get('end') or 'top') == 'bottom':
            desc = not desc
        return f"__SET_TOPN({lv}, {int(e.get('count') or 1)}, '{'DESC' if desc else 'ASC'}')"
    return None


JOIN_DUP = re.compile(r'^(.+?) \((?:Custom SQL Query|[^()]+)\)(?: #\d+)?$')
KEY_WORD = re.compile(r'\b(ID|FWO|WONUM|NUM|NUMBER|KEY|CODE)\b', re.I)
KEY_TYPE_WORDS = {'id', 'num', 'number', 'key', 'code', 'empl', 'dept', 'mgr', 'name', 'custom', 'sql', 'query'}


def _abbrev_of(tok, word):
    """'wo' abbreviates 'workorder' (same first letter, letters in order); 'pm' is 'pm'."""
    if tok == word:
        return True
    if not tok or not word or tok[0] != word[0] or len(tok) > len(word):
        return False
    it = iter(word)
    return all(ch in it for ch in tok)


def guess_joins(ds_el):
    """Published data source without its .tds: the TWB still lists every column of every logical table of the
    source (metadata-record <family>). When Tableau joins two tables that both hold a column of the same name, it
    renames the second copy '<name> (<table>) #n'; such a renamed column next to an unrenamed one of the same name in
    another table is the join key. Returns [{'many', 'one', 'many_col', 'one_col', 'physical_many', 'basis'}];
    'one' is the table the key belongs to (its name matches the key: WO Number -> WORK_ORDERS), else the table that
    kept the plain name, else the smaller table. Always a guess: the build marks each one DECISION NEEDED."""
    fam = {}
    for mr in ds_el.iter('metadata-record'):
        f, rn = mr.findtext('family'), mr.findtext('remote-name')
        if mr.get('class') == 'column' and f and rn:
            fam.setdefault(f, []).append(rn)
    norm = lambda s: re.sub(r'[\s_\-]+', '', s.lower())
    by_key = {}
    for f, cols in fam.items():
        for rn in cols:
            m = JOIN_DUP.match(rn)
            base = m.group(1) if m else rn
            if not KEY_WORD.search(base):
                continue
            by_key.setdefault(norm(base), []).append({'family': f, 'col': rn, 'base': base, 'renamed': bool(m)})
    out = []
    for key, hits in by_key.items():
        fams = {h['family'] for h in hits}
        if len(fams) < 2 or not any(h['renamed'] for h in hits):
            continue  # same name in two tables but never renamed by a join: not evidence of a join
        hits = list({h['family']: h for h in sorted(hits, key=lambda h: h['renamed'])}.values())  # one per table
        words = [w for w in re.findall(r'[a-z0-9]+', hits[0]['base'].lower()) if w not in KEY_TYPE_WORDS]

        def owns(h):
            ftoks = re.findall(r'[a-z0-9]+', h['family'].lower())
            return any(_abbrev_of(w, t) for w in words for t in ftoks)
        named = [h for h in hits if owns(h)]
        plain = [h for h in hits if not h['renamed']]
        if len(named) == 1:
            one, basis = named[0], f'the key "{hits[0]["base"]}" is named after {named[0]["family"]}'
        elif len(plain) == 1:
            one, basis = plain[0], f'{plain[0]["family"]} holds "{plain[0]["col"]}" under its own name; the others are join copies'
        else:
            one = min(hits, key=lambda h: len(fam[h['family']]))
            basis = f'{one["family"]} is the smallest table holding the key (weak guess)'
        for h in hits:
            if h is not one:
                out.append({'many': h['family'], 'one': one['family'], 'many_col': h['col'], 'one_col': one['col'],
                            'physical_many': h['base'], 'physical_one': one['base'], 'basis': basis})
    return out


class Workbook:
    def __init__(self, root, assets=(), tds_roots=(), table_map=None, strict_tds=False):
        self.table_map = table_map or {}
        self._dax = {}
        self.root = root
        self.assets = list(assets)
        self.ds = {}            # ds name -> info dict
        self.cols = {}          # (ds, col) -> column info
        self.inst = {}          # (ds, instance) -> (col, derivation, type)
        self.flags = []         # (code, where, detail)
        self.unhandled = {}     # (element, value) -> places, for UNHANDLED_XML
        self._read_datasources()
        self.tds_used = []
        for t in tds_roots:  # a root, or (root, path) so the file name can be matched too
            troot, tpath = t if isinstance(t, tuple) else (t, None)
            if self._merge_tds(troot, tpath, strict_tds) and tpath:
                self.tds_used.append(os.path.basename(tpath))
        self._index_instances()
        self.used = self._used_columns()
        self.blend_maps = self._blend_maps()

    # ---------- catalog ----------
    def _read_datasources(self):
        for d in self.root.findall('./datasources/datasource'):
            name = d.get('name')
            info = {'name': name, 'caption': d.get('caption') or name,
                    'is_parameters': name == 'Parameters', 'connection': None,
                    'filters': d.findall('./filter'), 'sets': {}, 'el': d,
                    'extract': d.find('./extract') is not None and d.find('./extract').get('enabled') != 'false'}
            conn = d.find('./connection')
            if conn is not None:
                info['connection'] = self._connection(conn, d)
            remote = {}
            for mr in d.iter('metadata-record'):
                if mr.get('class') not in ('column', 'measure'):
                    continue
                ln = mr.findtext('local-name')
                if ln:
                    remote[strip_brackets(ln)] = (mr.findtext('parent-name'), mr.findtext('family'),
                                                  mr.findtext('remote-name'), mr.findtext('local-type'))
            for mp in d.findall('./connection/cols/map'):  # field -> [table].[column]
                k, v = strip_brackets(mp.get('key')), mp.get('value') or ''
                if k not in remote and '].[' in v:
                    t, col = v.strip('[]').split('].[', 1)
                    remote[k] = (None, t, col, None)
            info['remote'] = remote
            if info['connection'] and info['connection'].get('published'):
                info['connection']['guessed_joins'] = guess_joins(d)
            for c in d.findall('./column'):
                self._add_column(name, c, remote)
            for g in d.findall('./group'):
                gname = strip_brackets(g.get('name'))
                if gname.startswith('Action (') or g.get(U + 'auto-column'):
                    continue  # hidden sets Tableau generates behind filter actions
                info['sets'][gname] = g
                formula = set_formula(g)
                if formula and (name, gname) not in self.cols:
                    # a set used in a calc ("IF [Last Week] THEN ...") is IN/OUT per row: a boolean calc the translator
                    # handles like any other (calc column, FIXED LOD, parameter inlining)
                    self.cols[(name, gname)] = {
                        'ds': name, 'name': gname, 'caption': g.get('caption'), 'datatype': 'boolean', 'role': 'dimension',
                        'type': 'nominal', 'formula': formula, 'calc_class': 'tableau', 'default_format': None,
                        'remote': None, 'aliases': {}, 'param': None, 'bins': None, 'hidden': False, 'aggregation': None,
                        'geo_role': None, 'from_set': True}
            self.ds[name] = info
            # datasource-level colour maps (Tableau saves field colours here)
            info['palettes'] = {}
            for enc in d.findall('./style/style-rule/encoding'):
                if enc.get('attr') == 'color':
                    info['palettes'][enc.get('field')] = self._palette(enc)
            # drill paths (Power BI hierarchies), display folders, default sort orders
            info['hierarchies'] = [(dp.get('name'), [strip_brackets(f.text or '') for f in dp.findall('./field')])
                                   for dp in d.findall('./drill-paths/drill-path')]
            info['folders'] = {strip_brackets(it.get('name') or ''): fo.get('name')
                               for fo in d.iter('folder') for it in fo.findall('./folder-item')}
            info['default_sorts'] = {}
            for ms in d.findall('./default-sorts/manual-sort'):
                pi = parse_instance(strip_brackets(ms.get('column') or ''))
                col = pi[1] if pi else strip_brackets(ms.get('column') or '')
                info['default_sorts'][col] = [(b.text or '').strip('"') for b in ms.findall('./dictionary/bucket')]
            for ch in d:
                self._known('datasource', ch.tag, info['caption'])

    def _add_column(self, ds, c, remote):
        n = strip_brackets(c.get('name'))
        calc = c.find('./calculation')
        info = {'ds': ds, 'name': n, 'caption': c.get('caption'), 'datatype': c.get('datatype'),
                'role': c.get('role'), 'type': c.get('type'), 'formula': None, 'calc_class': None,
                'default_format': c.get('default-format'), 'remote': remote.get(n),
                'aliases': {a.get('key'): a.get('value') for a in c.findall('./aliases/alias')},
                'param': None, 'bins': None, 'hidden': c.get('hidden') == 'true', 'aggregation': c.get('aggregation'),
                'geo_role': c.get('semantic-role')}
        if calc is not None:
            info['calc_class'] = calc.get('class')
            info['formula'] = calc.get('formula')
            if calc.get('class') == 'categorical-bin':
                info['bins'] = {'source': strip_brackets(calc.get('column')),
                                'default': calc.get('default'),
                                'groups': {b.get('value'): [v.text for v in b.findall('value')]
                                           for b in calc.findall('bin')}}
            elif calc.get('class') == 'bin':
                info['bins'] = {'source': strip_brackets(calc.get('column')), 'size': calc.get('size-parameter') or calc.get('size'),
                                'kind': 'numeric-bin'}
        if c.get('param-domain-type'):
            dom = {'domain': c.get('param-domain-type'), 'value': c.get('value')}
            r = c.find('./range')
            if r is not None:
                dom.update({k: r.get(k) for k in ('min', 'max', 'granularity', 'period-type-v2') if r.get(k)})
            mem = c.findall('./members/member')
            if mem:
                dom['members'] = [(m.get('value'), m.get('alias')) for m in mem]
            if c.get('source-field'):  # list "when the workbook opens" filled from a field's values
                dom['source_field_ref'] = c.get('source-field')
            info['param'] = dom
        if remote.get(n) is None and c.find('./calculation') is None and remote:
            # published DS columns are often renamed: match on caption
            info['remote'] = remote.get(info['caption'] or '')
        self.cols[(ds, n)] = info

    def _connection(self, conn, ds_el):
        cls = conn.get('class')
        out = {'class': cls}
        if cls == 'sqlproxy':
            out.update({'published': True, 'server': conn.get('server'), 'dbname': conn.get('dbname'),
                        'friendly': conn.get('server-ds-friendly-name')})
            rl = ds_el.find('./repository-location')
            if rl is not None:
                out['path'] = rl.get('path'), rl.get('id')
            return out
        out['published'] = False
        # what a Power Query connector needs per source type (Snowflake warehouse, Excel / text file, Oracle service)
        extra = ('warehouse', 'filename', 'directory', 'service', 'port')

        def info(c, name=None):
            d = {'class': c.get('class'), 'server': c.get('server'), 'dbname': c.get('dbname'), 'schema': c.get('schema')}
            if name is not None:
                d = {'name': name, **d}
            d.update({k: c.get(k) for k in extra if c.get(k)})
            return d
        out['named'] = [info(nc.find('connection'), nc.get('name'))
                        for nc in conn.findall('./named-connections/named-connection')
                        if nc.find('connection') is not None]
        if not out['named']:
            out['named'] = [info(conn)]
        rel = conn.find('./relation')
        out['relations'] = self._relation(rel) if rel is not None else None
        # Tableau 2020.2+ logical layer (relationships)
        og = ds_el.find('./object-graph')
        if og is not None:
            objs = {o.get('id'): o.get('caption') for o in og.findall('./objects/object')}
            fmap = self._field_map(ds_el)
            rels = []
            for r in og.findall('./relationships/relationship'):
                ex = r.find('./expression')
                ends = r.findall('./first-end-point') + r.findall('./second-end-point')
                ep = [objs.get(x.get('object-id'), x.get('object-id')) for x in ends]
                rel = {'between': ep, 'on': self._join_expr(ex), 'keys': self._logical_keys(ex, fmap)}
                # Performance options: cardinality "One" = unique-key on that end; "All records match" = referential
                # integrity. Default (nothing saved) is many-to-many, some records match.
                rel['unique'] = [any(k.endswith('unique-key') and v == 'true' for k, v in x.attrib.items()) for x in ends]
                # the table behind each end (an object holding a single table; a joined object has no one name)
                rel['end_tables'] = []
                for x in ends:
                    o = og.find(f"./objects/object[@id='{x.get('object-id')}']")
                    pr = o.find('./properties/relation') if o is not None else None
                    nm = pr.get('name') if pr is not None and pr.get('type') == 'table' else None
                    rel['end_tables'].append(self.table_map.get(nm, nm) if nm else None)
                ri = [x.get('referential-integrity') for x in ends if x.get('referential-integrity')]
                if ri:
                    rel['integrity'] = ri[0]
                rels.append(rel)
            out['logical'] = {'objects': list(objs.values()), 'relationships': rels}
            out['physical_by_object'] = {o.get('caption'): self._relation(o.find('./properties/relation'))
                                         for o in og.findall('./objects/object')
                                         if o.find('./properties/relation') is not None}
        return out

    def _field_map(self, ds_el):
        """Local field name -> (table, column) from a data source (TWB or .tds): <cols><map key='[Region (People)]'
        value='[People].[Region]'/>, else the metadata record's parent; calculated fields -> ('calc', caption, formula)."""
        out = {}
        for mr in ds_el.iter('metadata-record'):
            ln = strip_brackets(mr.findtext('local-name') or '')
            if mr.get('class') == 'column' and ln and mr.findtext('remote-name'):
                out[ln] = (strip_brackets(mr.findtext('parent-name') or '') or mr.findtext('family'), mr.findtext('remote-name'))
        for mp in ds_el.findall('./connection/cols/map'):
            v = mp.get('value') or ''
            if '].[' in v:
                out[strip_brackets(mp.get('key'))] = tuple(v.strip('[]').split('].[', 1))
        for c in ds_el.findall('./column'):
            calc = c.find('./calculation')
            if calc is not None and calc.get('formula') is not None:
                out[strip_brackets(c.get('name'))] = ('calc', c.get('caption') or strip_brackets(c.get('name')), calc.get('formula'))
        return out

    def _logical_keys(self, ex, fmap):
        """Relationship clause -> [{'op', 'left', 'right'}], each side {'field', 'table', 'column'} or a calc
        {'field', 'calc': formula}; [] when the clause has another shape. Tableau names the keys by local field name
        ([Region] = [Region (People)]), not by table."""
        if ex is None:
            return []
        kids = ex.findall('./expression')
        if ex.get('op') == 'AND':
            out = []
            for k in kids:
                sub = self._logical_keys(k, fmap)
                if not sub:
                    return []
                out += sub
            return out
        if len(kids) != 2 or any(k.findall('./expression') for k in kids):
            return []

        def side(k):
            f = strip_brackets(k.get('op') or '')
            hit = fmap.get(f)
            if hit and hit[0] == 'calc':
                return {'field': hit[1], 'calc': hit[2]}
            if hit:
                return {'field': f, 'table': self.table_map.get(hit[0], hit[0]), 'column': hit[1]}
            return {'field': f}
        return [{'op': ex.get('op'), 'left': side(kids[0]), 'right': side(kids[1])}]

    def _join_expr(self, ex):
        if ex is None:
            return None
        kids = ex.findall('./expression')
        if not kids:
            return ex.get('op')
        op = ex.get('op')
        parts = [self._join_expr(k) for k in kids]
        return f' {op} '.join(p for p in parts if p) if op not in ('=',) or len(parts) != 2 else f'{parts[0]} = {parts[1]}'

    def _relation(self, r):
        if r is None:
            return None
        t = r.get('type')
        if t == 'table':
            return {'table': r.get('table'), 'alias': r.get('name'), 'connection': r.get('connection')}
        if t == 'text':
            return {'custom_sql': clean_formula(r.text), 'alias': r.get('name')}
        if t == 'join':
            on = [self._join_expr(cl.find('./expression')) for cl in r.findall('./clause')]
            return {'join': r.get('join'), 'on': on, 'inputs': [self._relation(x) for x in r.findall('./relation')]}
        if t == 'union':
            return {'union': [self._relation(x) for x in r.findall('./relation')]}
        return {'type': t, 'name': r.get('name'), 'table': r.get('table')}

    def _merge_tds(self, troot, tds_path=None, strict=False):
        """Attach a published data source's real connection/joins/SQL to the matching sqlproxy DS.
        Matches on the data source name (formatted-name/caption or file name) against the published source's
        friendly name or server name. strict=True (input-folder mode) never falls back to an unnamed match."""
        conn = troot.find('./connection')
        if conn is None:
            return False
        info = self._connection(conn, troot)
        norm = lambda s: re.sub(r'[^a-z0-9]', '', (s or '').split('|')[0].lower())
        keys = {norm(troot.get('formatted-name')), norm(troot.get('caption'))}
        if tds_path:
            keys.add(norm(os.path.splitext(os.path.basename(tds_path))[0]))
        keys.discard('')
        targets = [d for d in self.ds.values() if d['connection'] and d['connection'].get('published')]
        match = [d for d in targets if keys & {norm(d['connection'].get('friendly')), norm(d['connection'].get('dbname')),
                                                norm(d['caption'])}]
        if not match and not strict and len(targets) == 1:
            match = targets  # a single published source and a .tds named explicitly on the command line
        for d in match[:1]:
            d['connection']['tds'] = info
            remote = {}
            for mp in troot.findall('./connection/cols/map'):  # authoritative field -> [table].[db column]
                v = mp.get('value') or ''
                if '].[' in v:
                    t, colname = v.strip('[]').split('].[', 1)
                    remote[strip_brackets(mp.get('key'))] = (None, t, colname, None)
            for mr in troot.iter('metadata-record'):
                ln = strip_brackets(mr.findtext('local-name') or '')
                if mr.get('class') in ('column', 'measure') and ln and ln not in remote:
                    remote[ln] = (mr.findtext('parent-name'), strip_brackets(mr.findtext('parent-name') or '') or mr.findtext('family'),
                                  mr.findtext('remote-name'), mr.findtext('local-type'))
            n_hit = 0
            for (ds, n), c in self.cols.items():
                if ds == d['name']:
                    hit = remote.get(n) or remote.get(c['caption'] or '')
                    if hit:  # the .tds knows the database table/column; the TWB only knew published names
                        c['remote'] = hit
                        n_hit += 1
            d['connection']['tds_columns_matched'] = n_hit
            # calcs defined in the published source but absent from the TWB copy
            for tc in troot.findall('./column'):
                k = (d['name'], strip_brackets(tc.get('name')))
                if k not in self.cols:
                    self._add_column(d['name'], tc, remote)
            d['connection']['tds_file'] = os.path.basename(tds_path) if tds_path else None
        return bool(match)

    def _known(self, element, value, where):
        """Remember XML this extractor does not read (see KNOWN_XML)."""
        if value not in KNOWN_XML[element] and not (isinstance(value, str) and value.startswith('_.fcp.')):
            places = self.unhandled.setdefault((element, value), [])
            if where not in places:
                places.append(where)

    def _zone_visibility(self):
        """Dynamic zone visibility (Tableau 2022.3+): {(dashboard uuid, zone id): field label}.
        The datagraph links a field node's value-output-guid to a zone node's visibility-input-guid by an edge."""
        g = self.root.find('./datagraph/graph')
        if g is None:
            return {}
        outputs = {}
        for n in g.iter():
            for k, v in n.attrib.items():
                if k.endswith('output-guid'):
                    outputs[v] = n
        feeds = {e.get('to'): e.get('from') for e in g.iter('edge')}
        out = {}
        for n in g.iter('dashboard-zone-visibility-node'):
            src = outputs.get(feeds.get(n.get('visibility-input-guid')))
            ref = src.get('fieldname') if src is not None else None
            out[(n.get('dashboard-identifier'), n.get('zone-id'))] = self.resolve_text(ref).strip('[]') if ref else '?'
        return out

    def _palette(self, enc):
        # Measure Names buckets are field references ("[ds].[sum:X:qk]"): show the readable label instead
        out = {}
        for m in enc.findall('./map'):
            for b in m.findall('./bucket'):
                key = self.resolve_text((b.text or '').strip('"'))
                if 'deleted field' in key:  # colour kept for a field that no longer exists
                    continue
                out[strip_brackets(key) if QREF.search(b.text or '') else key] = m.get('to')
        return out

    def _index_instances(self):
        for deps in self.root.iter('datasource-dependencies'):
            ds = deps.get('datasource')
            for c in deps.findall('./column'):
                if (ds, strip_brackets(c.get('name'))) not in self.cols and ds in self.ds:
                    self._add_column(ds, c, self.ds[ds].get('remote', {}))  # defined only inside a worksheet
            for ci in deps.findall('./column-instance'):
                self.inst[(ds, strip_brackets(ci.get('name')))] = (strip_brackets(ci.get('column')),
                                                                   ci.get('derivation'), ci.get('type'))

    # ---------- naming ----------
    def disp(self, ds, col):
        c = self.cols.get((ds, col))
        name = (c['caption'] if c and c['caption'] else col)
        if ds == 'Parameters':
            return f'Param:{name}'
        return name

    def resolve_ref(self, ds, inner):
        """'[ds].[deriv:Name:type]' parts -> readable label."""
        if inner in (':Measure Names',):
            return 'Measure Names'
        if inner in ('Multiple Values', ':Measure Values'):
            return 'Measure Values'
        if (ds, inner) in self.cols:
            return self.disp(ds, inner)
        parsed = parse_instance(inner)
        hit = self.inst.get((ds, inner))
        if parsed:
            pre, col = parsed
        elif hit:
            pre, col = [hit[1] or ''], hit[0]
        else:
            return inner
        if ds == '__tableau_internal_object_id__' or ((ds, col) not in self.cols and ROWCOUNT.match(col)):
            m = ROWCOUNT.match(col)
            return f'COUNT(rows of {m.group(1) if m else col})'
        if col == 'AdhocCluster':
            return 'CLUSTER (Tableau analytics)'
        if (ds, col) not in self.cols and col.startswith('Calculation_'):
            return f'(deleted field {col})'
        label = self.disp(ds, col)
        for p in reversed(pre):  # innermost derivation first: pcto:sum:X -> PERCENT_OF_TOTAL(SUM(X))
            p = p.lower()
            if p in AGG:
                label = f'{AGG[p]}({label})'
            elif p in DATEPART:
                label = f'{DATEPART[p]}({label})'
            elif p in QUICK_TC:
                label = f'{QUICK_TC[p]}({label})'
            elif p == 'io':
                label = f'IN/OUT({label})'
        return label

    def resolve_text(self, s):
        if not s:
            return s
        s = OBJECT_ID.sub('[__tableau_internal_object_id__].[', s)
        return QREF.sub(lambda m: f'[{self.resolve_ref(m.group(1), m.group(2))}]', s)

    def resolve_formula(self, ds, f):
        if not f:
            return f
        def rep(m):
            d, n = m.group(1) or ds, m.group(2)
            if (d, n) in self.cols:
                return f'[{self.disp(d, n)}]'
            return m.group(0)
        return ANYREF.sub(rep, clean_formula(f))

    def formula_refs(self, ds, f):
        out = set()
        for m in ANYREF.finditer(f or ''):
            d, n = m.group(1) or ds, m.group(2)
            if (d, n) in self.cols:
                out.add((d, n))
            else:  # formulas may reference by caption
                for (cd, cn), c in self.cols.items():
                    if cd == d and c['caption'] == n:
                        out.add((cd, cn))
        return out

    # ---------- DAX ----------
    def model_table(self, c):
        """Power BI table for a source column: source table (Tableau metadata family) through --table-map."""
        rem = c.get('remote')
        src = (rem[1] or strip_brackets(rem[0] or '')) if rem else None
        if not src and self.ds.get(c['ds'], {}).get('connection', {}) and \
                not (self.ds[c['ds']]['connection'] or {}).get('published'):
            src = None
        return self.table_map.get(src, src) if src else None

    def lookup(self, ds, tok):
        """Formula token '[x]' or '[ds].[x]' -> catalog key."""
        m = re.match(r'^\[((?:[^\]]|\]\])+)\](?:\.\[((?:[^\]]|\]\])+)\])?$', tok)
        if not m:
            return None
        d, n = (m.group(1), m.group(2)) if m.group(2) else (ds, m.group(1))
        d, n = d.replace(']]', ']'), n.replace(']]', ']')
        if (d, n) in self.cols:
            return (d, n)
        for (cd, cn), c in self.cols.items():
            if cd == d and c['caption'] == n:
                return (cd, cn)
        return None

    def dax_resolver(self, ds):
        def resolve(tok):
            key = self.lookup(ds, tok)
            if not key:
                return {'kind': 'unknown', 'name': tok}
            c = self.cols[key]
            name = c['caption'] or key[1]
            if key[0] == 'Parameters':
                return {'kind': 'param', 'name': name}
            if c['formula'] is not None or c['bins']:
                tr = self.calc_translation(key)
                return {'kind': 'rowcalc' if tr['kind'] in ('calculated column', 'field parameter') else 'measure', 'name': name,
                        'table': self.calc_table(key), 'datatype': c['datatype'], 'row_dax': tr.get('row_dax'),
                        'param_only': self.param_only(key)}
            return {'kind': 'column', 'name': name, 'table': self.model_table(c), 'datatype': c['datatype']}
        return resolve

    def _ref_pos(self, formula, key):
        """Where a field first appears in a formula (by name or caption); fields not found sort last."""
        hits = [formula.find(f'[{n}]') for n in (key[1], self.cols[key]['caption']) if n]
        return min([h for h in hits if h >= 0] or [len(formula)])

    def param_only(self, key, seen=None):
        """True for a calc that reads parameters only (e.g. DATEDIFF('day', [Min Date], [Max Date])): one value for the
        whole report, so a row-level formula may read it like a parameter."""
        seen = seen or set()
        if key in seen:
            return False
        seen.add(key)
        c = self.cols.get(key) or {}
        f = c.get('formula')
        if f is None or c.get('bins') or AGG_IN_EXPR.search(f) or TABLE_CALC.search(f):
            return False
        refs = self.formula_refs(key[0], f)
        return bool(refs) and all(r[0] == 'Parameters' or self.param_only(r, seen) for r in refs)

    def calc_table(self, key, seen=None):
        seen = seen or set()
        if key in seen:
            return None
        seen.add(key)
        c = self.cols[key]
        if c['bins']:
            src = self.cols.get((key[0], c['bins']['source']))
            return self.model_table(src) if src else None
        # first field in formula order wins (a set's order changes between runs, which made the home table random)
        for r in sorted(self.formula_refs(key[0], c['formula']), key=lambda r: (self._ref_pos(c['formula'], r), r)):
            rc = self.cols[r]
            if r[0] == 'Parameters':
                continue
            t = self.calc_table(r, seen) if rc['formula'] is not None else self.model_table(rc)
            if t:
                return t
        return None

    def calc_translation(self, key):
        if key in self._dax:
            return self._dax[key]
        self._dax[key] = {'dax': None, 'kind': 'measure', 'confidence': 'low', 'notes': ['circular reference']}
        c = self.cols[key]
        tr = Translator(self.dax_resolver(key[0]))
        self._dax[key] = tr.translate(c['formula'], self.view_role(key, c['role']), c.get('aggregation'), c['caption'] or key[1])
        return self._dax[key]

    # Tableau column-instance derivations that aggregate (sum:, avg:, usr: ...); every other one (none:, mn:, yr: ...)
    # puts the field on the view as a dimension
    AGG_DERIVATIONS = {'Sum', 'Avg', 'Min', 'Max', 'Count', 'CountD', 'Median', 'Attr', 'Stdev', 'StdevP', 'Var',
                       'VarP', 'User', 'Percentile', 'Collect'}

    def view_role(self, key, declared):
        """The role a calc really has in the views. A row-level calc whose Tableau role is "measure" but that a sheet
        places as a dimension (derivation None, e.g. a business-day count shown as a column of a detail table) must
        be a calculated column: as an iterator measure it cannot be a table column or a row header."""
        if declared != 'measure':
            return declared
        ders = {der or 'None' for (ds, _i), (col, der, _t) in self.inst.items() if (ds, col) == key}
        return 'dimension' if ders - self.AGG_DERIVATIONS else declared

    # ---------- usage ----------
    def _refs_in(self, el, skip_deps=True):
        out = set()
        stack = [el]
        while stack:
            e = stack.pop()
            if skip_deps and e.tag == 'datasource-dependencies':
                continue
            vals = list(e.attrib.values()) + ([e.text] if e.text else [])
            for v in vals:
                for m in QREF.finditer(v):
                    ds, inner = m.group(1), m.group(2)
                    col = self.inst.get((ds, inner), (None,))[0]
                    if not col:
                        pi = parse_instance(inner) if (ds, inner) not in self.cols else None
                        col = pi[1] if pi else inner
                    if (ds, col) in self.cols:
                        out.add((ds, col))
                    elif e.tag == 'run' and ds in self.ds and col:
                        out.add((ds, col))  # title / text field: _used_columns resolves case or builds it from metadata
            stack.extend(list(e))
        return out

    def _used_columns(self):
        used = {}
        lower = {}
        for (ds, cn) in self.cols:
            lower.setdefault((ds, cn.lower()), (ds, cn))

        def add(key, where):
            # Tableau field names are case-insensitive: a title may say [Task COMP by] for the column [Task COMP By]
            if key not in self.cols and key[1] and (key[0], key[1].lower()) in lower:
                key = lower[(key[0], key[1].lower())]
            if key not in self.cols and key[0] in self.ds and key[1]:
                # a source column referenced only in a sheet title (no datasource-dependencies entry): build it
                # from the data source metadata, as the shelves' columns are
                remote = self.ds[key[0]].get('remote', {})
                hit = next((k for k in remote if k.lower() == key[1].lower()), None)
                if hit:
                    el = ET.Element('column', {'name': f'[{hit}]', 'datatype': remote[hit][3] or 'string',
                                               'role': 'dimension', 'type': 'nominal'})
                    self._add_column(key[0], el, remote)
                    key = (key[0], hit)
                    lower.setdefault((key[0], hit.lower()), key)
            used.setdefault(key, set()).add(where)
        for ws in self.root.findall('./worksheets/worksheet'):
            w = ws.get('name')
            for deps in ws.iter('datasource-dependencies'):
                ds = deps.get('datasource')
                for c in deps.findall('./column'):
                    add((ds, strip_brackets(c.get('name'))), w)
            for k in self._refs_in(ws):
                add(k, w)
            for f in ws.iter('filter'):  # a Top N 'order by' calc is saved bare ([Calculation_77], no data source)
                fm = QREF.match(f.get('column', ''))
                for g in f.iter('groupfilter'):
                    om = re.fullmatch(r'\[([^\[\]]+)\]', (g.get('expression') or '').strip()) if g.get('function') == 'order' else None
                    if fm and om and (fm.group(1), om.group(1)) in self.cols:
                        add((fm.group(1), om.group(1)), w)
        for db in self.root.findall('./dashboards/dashboard'):
            for k in self._refs_in(db):
                add(k, 'dashboard:' + db.get('name').strip())
        acts = self.root.find('./actions')
        if acts is not None:
            for a in acts.findall('./action'):
                exp = (a.find('./link').get('expression') if a.find('./link') is not None else '') or ''
                for n in re.findall(r'<\[([^\]]+)\]>', exp):
                    for (ds, cn), c in self.cols.items():
                        if cn == n or c['caption'] == n:
                            add((ds, cn), 'action:' + (a.get('caption') or a.get('name')))
        for d in self.ds.values():
            for f in d['filters']:
                # '[ds].[Order Date]' or '[ds].[none:Order Date:nk]' (a split on '].[' left 'Order Date]': never matched)
                m = QREF.match(f.get('column', ''))
                inner = m.group(2) if m else strip_brackets(f.get('column', ''))
                pi = parse_instance(inner)
                k = (d['name'], pi[1] if pi else inner)
                if k in self.cols:
                    add(k, 'datasource-filter')
        # transitive closure through calculations
        frontier = list(used)
        while frontier:
            k = frontier.pop()
            c = self.cols.get(k)
            if not c:
                continue
            src = self.formula_refs(k[0], c['formula'])
            if c['bins'] and c['bins'].get('source'):
                src.add((k[0], c['bins']['source']))
            for s in src:
                if s not in used:
                    used[s] = set()
                    frontier.append(s)
                used[s].add('via:' + self.disp(*k))
        return used

    # ---------- filters ----------
    def gf(self, e):
        f = e.get('function')
        lvl = e.get('level')
        if f == 'member':
            return e.get('member')
        if f == 'range':
            return f'RANGE({e.get("from")}..{e.get("to")})'
        if f == 'level-members':
            return 'ALL'
        if f in ('union', 'intersection', 'crossjoin'):
            parts = [self.gf(k) for k in e.findall('./groupfilter')]
            if f == 'union' and all(k.get('function') == 'member' for k in e.findall('./groupfilter')):
                return 'IN(' + ', '.join(parts) + ')'
            return f' {f.upper()} '.join(parts)
        if f == 'except':
            kids = [self.gf(k) for k in e.findall('./groupfilter')]
            return f'{kids[0]} EXCEPT {" , ".join(kids[1:])}' if kids else 'EXCEPT'
        if f == 'filter':
            return f'WHERE {self.resolve_text(e.get("expression"))}'
        if f == 'end':
            inner = e.find('./groupfilter')
            return f'{e.get("end", "top").upper()} {e.get("count")} BY {self.resolve_text(e.get("expression") or "")} ' + \
                   (f'OF {self.gf(inner)}' if inner is not None else '')
        if f == 'order':
            inner = e.find('./groupfilter')
            return f'ORDER BY {self.order_expr(e.get("expression"))} {e.get("direction", "")} ' + (self.gf(inner) if inner is not None else '')
        return f'{f}({lvl})'

    def order_expr(self, expr):
        """Top N 'order by' field: '[ds].[sum:Sales:qk]', or a bare '[Calculation_123]' of the filter's own data source
        (Tableau saves the Top N 'by field' calc without its source) -> its readable label."""
        expr = expr or ''
        if QREF.search(expr):
            return self.resolve_text(expr)
        m = re.fullmatch(r'\[([^\[\]]+)\]', expr.strip())
        ds = getattr(self, '_gf_ds', None)
        if m and ds:
            n = m.group(1)
            if (ds, n) in self.cols:
                return f'[{self.disp(ds, n)}]'
            return f'[{self.resolve_ref(ds, n)}]'
        return expr

    def describe_filter(self, f):
        col = f.get('column', '')
        m = QREF.match(col)
        self._gf_ds = m.group(1) if m else None  # Top N 'order by' calcs are saved without their data source
        label = self.resolve_ref(m.group(1), m.group(2)) if m else col
        cls = f.get('class')
        if cls == 'categorical':
            g = f.find('./groupfilter')
            if g is None:
                return label, 'categorical'
            enum = ua(g, 'ui-enumeration')
            txt = self.gf(g)
            if enum == 'exclusive':
                txt = 'NOT ' + txt
            dom = ua(g, 'ui-domain')
            if dom:
                txt += f'  [values:{dom}]'
            return label, txt
        if cls == 'quantitative':
            mn, mx = f.findtext('min'), f.findtext('max')
            inc = f.get('included-values')
            return label, f'RANGE {mn or "-inf"}..{mx or "+inf"}' + (f' ({inc})' if inc and inc != 'in-range' else '')
        if cls == 'relative-date':
            # every attribute except the column / class, so newer ones (period-type-v2, include-future) are kept
            skip = ('class', 'column', 'filter-group')
            return label, 'RELATIVE ' + ' '.join(f'{k}={v}' for k, v in f.attrib.items() if k not in skip and '}' not in k)
        return label, cls

    # ---------- worksheets ----------
    def field_info(self, ds, inner):
        """Structured view of one shelf/encoding field: label, measure vs dimension, continuous, date."""
        label = self.resolve_ref(ds, inner)
        if label in ('Measure Names', 'Measure Values'):
            return {'label': label, 'measure': label == 'Measure Values', 'continuous': False, 'date': False,
                    'special': label}
        tm = re.search(r':(nk|ok|qk)(?::\d+)?$', inner)
        pi = parse_instance(inner)
        pre = [p.lower() for p in (pi[0] if pi else [])]
        colname = pi[1] if pi else inner
        c = self.cols.get((ds, colname)) or {}
        if ds == '__tableau_internal_object_id__' or label.startswith('COUNT(rows'):
            measure = True
        elif any(p in AGG or p in QUICK_TC for p in pre):
            measure = True
        elif 'none' in pre or any(p in DATEPART for p in pre):
            measure = False
        else:
            measure = c.get('role') == 'measure'
        continuous = (tm.group(1) == 'qk') if tm else c.get('type') == 'quantitative'
        info = {'label': label, 'measure': measure, 'continuous': continuous,
                'date': c.get('datatype') in DATE_TYPES or any(p in DATEPART for p in pre)}
        order = (getattr(self, '_ws_tc', None) or {}).get((ds, inner))
        if order:  # quick table calc "compute using": Rows = Table (across), Columns = Table (down), ...
            info['tc_ordering'] = order
        more = (getattr(self, '_ws_tc_more', None) or {}).get((ds, inner))
        if more:  # Field = compute using the listed dimensions (in this order)
            info.update({'tc_' + k: v for k, v in more.items()})
        if ds == 'Parameters':
            info['parameter'] = True
        elif ds in self.ds:  # the field's own data source: two sources may both have a [Region] (blending)
            info['ds'] = self.ds[ds]['caption']
        return info

    def shelf(self, text):
        if not text:
            return []
        text = OBJECT_ID.sub('[__tableau_internal_object_id__].[', text)
        return [self.field_info(m.group(1), m.group(2)) for m in QREF.finditer(text)]

    def worksheet(self, ws):
        name = ws.get('name')
        tbl = ws.find('./table')
        view = tbl.find('./view')
        self._ws_tc = {}  # this sheet's quick table calcs: (ds, instance) -> ordering-type
        self._ws_tc_more = {}  # (ds, instance) -> {'type', 'fields'}: "compute using" specific dimensions
        for deps in ws.iter('datasource-dependencies'):
            for ci in deps.findall('./column-instance'):
                tc = ci.find('./table-calc')
                if tc is None:
                    continue
                key = (deps.get('datasource'), strip_brackets(ci.get('name')))
                self._known('table-calc', tc.get('ordering-type'), name)
                if tc.get('ordering-type'):
                    self._ws_tc[key] = tc.get('ordering-type')
                more = {'type': tc.get('type'),
                        'fields': [self.resolve_text(o.get('field')).strip('[]') for o in tc.findall('./order')]}
                if more['type'] or more['fields']:
                    self._ws_tc_more[key] = {k: v for k, v in more.items() if v}
        for ch in tbl:
            self._known('table', ch.tag, name)
        for ch in view:
            self._known('view', ch.tag, name)
        for p in tbl.findall('./panes/pane'):
            for ch in p:
                self._known('pane', ch.tag, name)
            if p.find('./mark') is not None:
                self._known('mark', p.find('./mark').get('class'), name)
        for f in view.findall('./filter'):
            self._known('filter', f.get('class'), name)
            for g in f.iter('groupfilter'):
                self._known('groupfilter', g.get('function'), name)
        out = {'name': name, 'title': self._text(ws.find('./layout-options/title/formatted-text'))}
        out['datasources'] = [self.ds[d.get('name')]['caption'] if d.get('name') in self.ds else d.get('name')
                              for d in view.findall('./datasources/datasource')]
        out['rows'] = self.resolve_text(tbl.findtext('./rows'))
        out['cols'] = self.resolve_text(tbl.findtext('./cols'))
        out['shelf'] = {'rows': self.shelf(tbl.findtext('./rows')), 'cols': self.shelf(tbl.findtext('./cols')),
                        'dual': any(e.get('synchronized') == 'true' for e in tbl.iter('encoding')) or None}
        if (tbl.find('./rows') is not None and tbl.find('./rows').get('total') == 'true') or \
           (tbl.find('./cols') is not None and tbl.find('./cols').get('total') == 'true'):
            out['grand_totals'] = True
        panes = []
        for p in tbl.findall('./panes/pane'):
            mark = p.find('./mark')
            enc, enc_s = [], []
            encs = p.find('./encodings')
            for e in (list(encs) if encs is not None else []):
                col = OBJECT_ID.sub('[__tableau_internal_object_id__].[', e.get('column') or '')
                m = QREF.match(col)
                item = f'{e.tag}={self.resolve_ref(m.group(1), m.group(2)) if m else col}'
                if item not in enc:
                    enc.append(item)
                    if m:
                        enc_s.append({'channel': e.tag, **self.field_info(m.group(1), m.group(2))})
            pd = {'mark': mark.get('class') if mark is not None else None, 'encodings': enc, 'enc': enc_s}
            for sr in p.findall('./style/style-rule'):
                for fm in sr.findall('./format'):
                    if fm.get('attr') == 'mark-color':
                        pd['mark_color'] = fm.get('value')
            if p.get('y-axis-name') or p.get('x-axis-name'):
                pd['axis'] = self.resolve_text(p.get('y-axis-name') or p.get('x-axis-name'))
            lbl = p.find('./customized-label/formatted-text')
            if lbl is not None:
                pd['label'] = self._text(lbl)
            tip = p.find('./customized-tooltip/formatted-text')
            if tip is not None:
                pd['tooltip'] = self._text(tip)  # full text in JSON; the digest shows the first 160 characters
                if '<Sheet name=' in ET.tostring(tip, encoding='unicode') or 'tooltip-viz' in ET.tostring(tip, encoding='unicode'):
                    pd['viz_in_tooltip'] = True
            for sr in p.findall('./style/style-rule'):
                for fm in sr.findall('./format'):
                    if fm.get('attr') == 'mark-labels-show' and fm.get('value') == 'true':
                        pd['labels_shown'] = True
            tl = p.find('./trendline')
            if tl is not None and tl.get('enabled') != 'false':
                pd['trendline'] = tl.get('fit') or 'linear'
            panes.append(pd)
        # drop empty duplicate panes
        out['panes'] = [p for i, p in enumerate(panes) if p['encodings'] or p.get('mark') not in (None, 'Automatic') or i == 0]
        filters = []
        for f in view.findall('./filter'):
            label, txt = self.describe_filter(f)
            if label == 'Measure Names':
                out['measure_values'] = [self.resolve_text(g.get('member', '').strip('"').replace('\\', ''))
                                         for g in f.iter('groupfilter') if g.get('function') == 'member']
                continue
            if txt.startswith('ALL'):
                txt = 'ALL (interactive)'
            item = {'field': label, 'rule': txt}
            fm = QREF.match(f.get('column', ''))
            if fm and fm.group(1) in self.ds and fm.group(1) != 'Parameters':
                item['ds'] = self.ds[fm.group(1)]['caption']
            if f.get('context') == 'true':
                item['context'] = True
            filters.append(item)
        out['filters'] = filters
        sorts = []
        for s in view.findall('./sort') + view.findall('./computed-sort') + view.findall('./natural-sort'):
            sorts.append(f'{self.resolve_text(s.get("column"))} {s.get("direction", "")} by {self.resolve_text(s.get("using") or s.tag)}')
        for s in view.findall('./manual-sort'):
            if s.get('column', '').endswith('[:Measure Names]'):
                # display order of the measure columns: replaces the filter-order list
                out['measure_values'] = [self.resolve_text((b.text or '').strip('"').replace('\\', ''))
                                         for b in s.findall('./dictionary/bucket')]
                continue
            members = [self.resolve_text((b.text or '').strip('"')) for b in s.findall('./dictionary/bucket')]
            sorts.append(f'{self.resolve_text(s.get("column"))} manual: ' + ', '.join(members)[:200])
            out.setdefault('manual_sorts', {})[self.resolve_text(s.get('column')).strip('[]')] = members
        for s in view.findall('./shelf-sorts/shelf-sort-v2'):  # Tableau 2020.3+ sort format
            by = s.get('measure-to-sort-by')
            sorts.append(f'{self.resolve_text(s.get("dimension-to-sort"))} {s.get("direction", "")}'
                         + (f' by {self.resolve_text(by)}' if by else ''))
        if sorts:
            out['sorts'] = sorts
        subt = [self.resolve_text(c.text or '').strip('[]') for c in tbl.findall('./subtotals/column')]
        if subt:
            out['subtotals'] = subt
        if tbl.find('./forecast-specification') is not None:
            self.flags.append(('FORECAST', name, 'Tableau forecast: Power BI line chart Analytics > Forecast '
                               '(different model; numbers will not match)'))
        trends = sorted({p['trendline'] for p in panes if p.get('trendline')})
        if trends:
            self.flags.append(('TREND_LINE', name, f'{", ".join(trends)} trend line: Analytics pane > Trend line '
                               '(linear only; other fits need a DAX line)'))
        refs = []
        for r in tbl.iter('reference-line'):
            refs.append({'formula': r.get('formula'), 'value': self.resolve_text(r.get('value-column')),
                         'axis': self.resolve_text(r.get('axis-column')), 'scope': r.get('scope'),
                         'label': r.get('label')})
            if r.get('value-column') != r.get('axis-column'):
                self.flags.append(('REFLINE_CROSS_AXIS', name, f'{refs[-1]["value"]} drawn on {refs[-1]["axis"]} axis'))
        if refs:
            out['reference_lines'] = refs
        fmts = {}
        palettes = {}
        for sr in tbl.findall('./style/style-rule'):
            for fm in sr.findall('./format'):
                if fm.get('attr') == 'text-format' and fm.get('field'):
                    fmts[self.resolve_text(fm.get('field'))] = fm.get('value')
            for enc in sr.findall('./encoding'):
                if enc.get('attr') == 'color' and enc.find('./map') is not None:
                    palettes[self.resolve_text(enc.get('field'))] = self._palette(enc)
        if fmts:
            out['number_formats'] = fmts
        if palettes:
            out['colors'] = palettes
        # flags
        raw = ET.tostring(tbl, encoding='unicode')
        qtc = sorted({self.resolve_ref(m.group(1), m.group(2)) for m in QREF.finditer(raw)
                      if (parse_instance(m.group(2)) or ([], ''))[0] and
                      any(p in QUICK_TC for p in parse_instance(m.group(2))[0])})
        if qtc:
            self.flags.append(('QUICK_TABLE_CALC', name, ', '.join(qtc) + ': check "Compute using" with --show; '
                               'rebuild as DAX (window/visual calcs)'))
        if 'AdhocCluster' in raw:
            self.flags.append(('CLUSTER', name, 'Tableau clustering: no DAX equivalent'))
        real_ds = [d.get('name') for d in view.findall('./datasources/datasource') if d.get('name') != 'Parameters']
        if len(real_ds) > 1:  # data source captions, not the internal federated.xxx ids
            out['blend'] = self.blend(view, real_ds)
            links = '; '.join(f'{s["datasource"]} on ' + (', '.join(
                lk['primary_field'] + (f' = {lk["secondary_field"]}' if lk['secondary_field'] != lk['primary_field'] else '')
                + (f' ({lk["level"]})' if lk['level'] else '') for lk in s['links']) or 'NO ACTIVE LINK')
                for s in out['blend']['secondaries'])
            self.flags.append(('BLENDING', name, ', '.join(self.ds[d]['caption'] if d in self.ds else d for d in real_ds)
                               + f' | primary {out["blend"]["primary"]}; links: {links}'))
            for s in out['blend']['secondaries']:
                if not s['links']:
                    self.flags.append(('BLEND_NO_LINK', name, f'{s["datasource"]} has no active link to {out["blend"]["primary"]}: '
                                       'Tableau shows its grand total on every mark; Power BI does the same (no relationship)'))
                for lk in s['links']:
                    if lk['calc']:
                        self.flags.append(('BLEND_LINK_CALC', name, f'{lk["primary_field"]} = {lk["secondary_field"]}: a calculated '
                                           'field links the sources; the build cannot make a shared dimension from it (relate by hand)'))
                    elif lk['datatype'] != lk['secondary_datatype'] and not {lk['datatype'], lk['secondary_datatype']} <= DATE_TYPES:
                        self.flags.append(('BLEND_LINK_TYPE', name, f'{lk["primary_field"]} ({lk["datatype"]}) = '
                                           f'{lk["secondary_field"]} ({lk["secondary_datatype"]}): different data types'))
        shelf = (out['rows'] or '') + (out['cols'] or '')
        if any(e.get('synchronized') == 'true' for e in tbl.iter('encoding')):
            self.flags.append(('DUAL_AXIS', name, 'synchronized dual axis (often a trick for labels/ref lines); usually one axis in Power BI'))
        elif '+' in shelf:
            self.flags.append(('MULTI_MEASURE_AXES', name, 'side-by-side measure axes ("+"): separate visuals or small multiples in Power BI'))
        for f in view.findall('./filter'):
            if f.get('class') == 'categorical' and any(g.get('function') == 'range' for g in f.iter('groupfilter')):
                self.flags.append(('STRING_RANGE_FILTER', name, f'{self.describe_filter(f)[0]}: alphabetical range keeps any value '
                                   'sorting inside it; confirm the intended list'))
            if f.get('context') == 'true':
                self.flags.append(('CONTEXT_FILTER', name, self.describe_filter(f)[0]))
            if any(g.get('function') == 'end' for g in f.iter('groupfilter')):
                self.flags.append(('TOP_N_FILTER', name, self.describe_filter(f)[0]))
        if 'Measure Values' in shelf or any('Measure Values' in e for p in out['panes'] for e in p['encodings']):
            self.flags.append(('MEASURE_VALUES', name, 'one visual column per measure: ' + ', '.join(out.get('measure_values', []))))
        if any(p.get('viz_in_tooltip') for p in out['panes']):
            self.flags.append(('VIZ_IN_TOOLTIP', name, 'needs a report-page tooltip'))
        return out

    # ---------- data blending ----------
    def _blend_maps(self):
        """Edit Blend Relationships (custom links), saved at workbook level:
        {(primary ds, secondary ds): [(primary instance, secondary instance)]}. Real XML (Retail Performance Dashboard):
        <datasource-relationships><datasource-relationship source='ds1' target='ds2'><column-mapping>
        <map key='[ds1].[none:Region:nk]' value='[ds2].[none:Sales Territory:nk]' /> ..."""
        out = {}
        for rel in self.root.findall('./datasource-relationships/datasource-relationship'):
            for mp in rel.findall('./column-mapping/map'):
                k, v = QREF.match(mp.get('key') or ''), QREF.match(mp.get('value') or '')
                if not (k and v) or k.group(1) == v.group(1):
                    continue  # a mapping of a source onto itself (Retail saves some) links nothing
                pair = (k.group(1), v.group(1))
                if (k.group(2), v.group(2)) not in out.setdefault(pair, []):
                    out[pair].append((k.group(2), v.group(2)))
        return out

    def _link_end(self, ds, inner):
        """'tmn:Order Date:ok' -> ('Order Date', 'month'); 'none:Region:nk' -> ('Region', None)."""
        pi = parse_instance(inner)
        hit = self.inst.get((ds, inner))
        col = pi[1] if pi else (hit[0] if hit else inner)
        pre = [p.lower() for p in (pi[0] if pi else [])]
        return col, next((BLEND_LEVEL[p] for p in pre if p in BLEND_LEVEL), None)

    def blend(self, view, real_ds):
        """A blended sheet: the primary data source (first one used in the view) and, per secondary, the ACTIVE links.
        Tableau links on the custom mappings of Edit Blend Relationships when the pair has any, else automatically on
        the dimensions of the same name; a link is active only while its primary field is used in the view (shelves,
        marks, filters: the sheet's datasource-dependencies). A date link works at the level the view uses (or the
        level of the custom mapping). Secondary values are aggregated to the linked fields: the Power BI equivalent is
        a shared dimension per link, related to both tables (the build makes it)."""
        cap = lambda d: self.ds[d]['caption'] if d in self.ds else d
        primary = real_ds[0]
        used = {}  # ds -> {column: level it is used at}
        for deps in view.findall('./datasource-dependencies'):
            u = used.setdefault(deps.get('datasource'), {})
            for ci in deps.findall('./column-instance'):
                col, lev = self._link_end(deps.get('datasource'), strip_brackets(ci.get('name')))
                if col not in u or (lev and not u[col]):
                    u[col] = lev
        out = {'primary': cap(primary), 'secondaries': []}
        for sec in real_ds[1:]:
            pairs = [(self._link_end(primary, a), self._link_end(sec, b), True) for a, b in self.blend_maps.get((primary, sec), [])]
            pairs += [(self._link_end(primary, b), self._link_end(sec, a), True) for a, b in self.blend_maps.get((sec, primary), [])]
            if not pairs:  # automatic links: dimensions with the same name in both sources
                pcols = {self.disp(primary, n).lower(): n for (d, n), c in self.cols.items()
                         if d == primary and c.get('role') == 'dimension' and not n.startswith(':') and not c.get('param')}
                pairs = [((pcols[self.disp(sec, n).lower()], None), (n, None), False) for (d, n), c in self.cols.items()
                         if d == sec and c.get('role') == 'dimension' and not n.startswith(':') and not c.get('param')
                         and self.disp(sec, n).lower() in pcols]
            links = []
            for (pcol, plev), (scol, slev), custom in pairs:
                pc, sc = self.cols.get((primary, pcol)), self.cols.get((sec, scol))
                if not pc or not sc or pcol not in used.get(primary, {}):
                    continue  # the primary field is not in the view: a potential (inactive) link
                date = pc.get('datatype') in DATE_TYPES
                level = (plev or slev or used[primary].get(pcol) or 'day') if date else None
                # a link column the workbook's metadata does not place (published source without its .tds, column only
                # declared in the sheet): the one table every other used column of that source comes from
                one = lambda d, c: self.model_table(c) or (next(iter(self._ds_tables(d))) if len(self._ds_tables(d)) == 1
                                                           and c.get('formula') is None else None)
                links.append({'primary_field': self.disp(primary, pcol), 'secondary_field': self.disp(sec, scol),
                              'primary_datasource': cap(primary), 'secondary_datasource': cap(sec),
                              'primary_table': one(primary, pc), 'secondary_table': one(sec, sc),
                              'primary_column': (pc.get('remote') or (None,) * 3)[2] or pcol,  # source (database) names
                              'secondary_column': (sc.get('remote') or (None,) * 3)[2] or scol,
                              'datatype': pc.get('datatype'), 'secondary_datatype': sc.get('datatype'), 'level': level,
                              'custom': custom, 'calc': pc.get('formula') is not None or sc.get('formula') is not None})
            out['secondaries'].append({'datasource': cap(sec), 'links': links})
        return out

    def _text(self, ft):
        if ft is None:
            return None
        s = ''.join((r.text or '') for r in ft.findall('./run'))
        s = s.replace('Æ', ' ').replace('\n', ' ').strip()
        return re.sub(r'\s+', ' ', self.resolve_text(s))

    # ---------- dashboards ----------
    def dashboard(self, db):
        name = db.get('name').strip()
        size = db.find('./size')
        out = {'name': name, 'size': None, 'zones': [], 'items': []}
        W, H = 1000, 800  # Tableau's default dashboard size
        if size is not None:
            out['size'] = f'{size.get("sizing-mode", "")} {size.get("minwidth") or size.get("maxwidth")}x{size.get("minheight") or size.get("maxheight")}'
            W = int(size.get('minwidth') or size.get('maxwidth') or W)
            H = int(size.get('minheight') or size.get('maxheight') or H)
            out['sizing'] = size.get('sizing-mode') or 'automatic'
        else:
            out['sizing'] = 'automatic'
        out['px'] = [W, H]
        sid = db.find('./simple-id')
        uuid = sid.get('uuid') if sid is not None else None
        dzv = getattr(self, '_dzv', None) or {}
        if db.get('type') == 'storyboard':
            out['story'] = [{'caption': sp.get('caption'), 'sheet': sp.get('captured-sheet')}
                            for sp in db.iter('story-point')]
            self.flags.append(('STORY', name, f'{len(out["story"])} story points: one Power BI page (or bookmark) per point '
                               '+ page navigator'))
        def walk(z, depth):
            t = z.get('type-v2') or z.get('type')
            self._known('zone', t, name)
            shown_when = dzv.get((uuid, z.get('id')))
            if shown_when:
                self.flags.append(('DYNAMIC_ZONE_VISIBILITY', name, 'zones shown/hidden by a field or parameter value: '
                                   'bookmarks + buttons, or a measure that blanks the visual'))
            if z.get('hidden-by-user') == 'true':
                self.flags.append(('HIDDEN_ZONE', name, 'zone hidden by default (show/hide button): bookmark + button'))
            geo = ''
            box = None
            if all(z.get(k) for k in ('x', 'y', 'w', 'h')):
                geo = '@{:.0f},{:.0f} {:.0f}x{:.0f}%'.format(*(int(z.get(k)) / 1000 for k in ('x', 'y', 'w', 'h')))
                box = [round(int(z.get('x')) * W / 100000), round(int(z.get('y')) * H / 100000),
                       round(int(z.get('w')) * W / 100000), round(int(z.get('h')) * H / 100000)]
            line = None
            param = z.get('param') or ''
            pm = QREF.match(param)
            field = self.resolve_ref(pm.group(1), pm.group(2)) if pm else param
            if t not in ('layout-basic', 'layout-flow', 'empty'):
                out['items'].append({'type': t or 'sheet', 'name': z.get('name'), 'field': field if param else None,
                                     'mode': z.get('mode'), 'box': box, 'show_title': z.get('show-title') != 'false',
                                     'values': z.get('values'),
                                     'text': self._text(z.find('./formatted-text')) if t == 'text' else None})
                if shown_when:
                    out['items'][-1]['shown_when'] = shown_when
                if z.get('hidden-by-user') == 'true':
                    out['items'][-1]['hidden'] = True
            if t in ('layout-basic', 'layout-flow'):
                kids = z.findall('./zone')
                if kids:
                    tag = {'horz': 'ROW', 'vert': 'COLUMN'}.get(param, 'BOX')
                    if shown_when:
                        geo += f' (SHOWN WHEN [{shown_when}])'
                    if z.get('hidden-by-user') == 'true':
                        geo += ' (HIDDEN until shown by a show/hide button)'
                    if len(kids) > 1 or depth == 0 or geo.endswith(')'):
                        out['zones'].append('  ' * depth + f'{tag} {geo}')
                        depth += 1
                for k in kids:
                    walk(k, depth)
                return
            if t is None and z.get('name'):
                line = f'SHEET "{z.get("name")}" {geo}' + (' (title hidden)' if z.get('show-title') == 'false' else '')
            elif t == 'filter':
                line = f'FILTER {field} [{z.get("mode")}] for "{z.get("name")}"'
            elif t == 'paramctrl':
                line = f'PARAM {field} [{z.get("mode")}]'
            elif t in ('color', 'size', 'shape', 'map'):
                line = f'LEGEND {t} {field} for "{z.get("name")}"'
            elif t == 'bitmap':
                line = f'IMAGE {param} {geo}'
            elif t == 'text':
                line = f'TEXT "{(self._text(z.find("./formatted-text")) or "")[:160]}" {geo}'
            elif t == 'title':
                line = 'DASHBOARD TITLE'
            elif t == 'web':
                line = f'WEB {z.get("url") or param}'
            elif t in ('dashboard-object', 'button', 'navigation'):
                btn = z.find('./button')
                line = f'BUTTON {ET.tostring(btn, encoding="unicode")[:160] if btn is not None else param}'
            elif t == 'empty':
                return
            elif t == 'flipboard':
                line = 'STORY POINTS: ' + ' | '.join(f'{sp.get("caption")} -> "{sp.get("captured-sheet")}"'
                                                    for sp in z.iter('story-point'))
            elif t == 'flipboard-nav':
                line = 'STORY NAVIGATOR'
            else:
                line = f'{t} {param} {geo}'.strip()
            if shown_when:
                line += f' (SHOWN WHEN [{shown_when}])'
            if z.get('hidden-by-user') == 'true':
                line += ' (HIDDEN until shown by a show/hide button)'
            out['zones'].append('  ' * depth + line)
        for z in db.findall('./zones/zone'):
            walk(z, 0)
        return out

    # ---------- actions ----------
    def actions(self):
        acts = []
        node = self.root.find('./actions')
        if node is None:
            return acts
        for a in node:
            self._known('actions', a.tag, a.get('caption') or a.get('name') or 'actions')
            if a.tag not in ('action', 'edit-parameter-action', 'nav-action'):
                continue
            src = a.find('./source')
            act = a.find('./activation')
            cmd = a.find('./command')
            link = a.find('./link')
            item = {'name': a.get('caption') or a.get('name'),
                    'source': ((src.get('dashboard') or '').strip() if src is not None else '') +
                              (f' / {src.get("worksheet")}' if src is not None and src.get('worksheet') else ''),
                    'trigger': act.get('type') if act is not None else 'menu'}
            if src is not None and src.findall('./exclude-sheet'):
                item['excludes'] = [x.get('name') for x in src.findall('./exclude-sheet')]
            params = {p.get('name'): p.get('value') for p in a.findall('./params/param')}
            if a.tag == 'edit-parameter-action':
                # click a mark -> set a parameter to the clicked value (Tableau 2019.2+)
                fm = QREF.match(params.get('source-field') or '')
                pm = QREF.match(params.get('target-parameter') or '')
                item.update({'kind': 'PARAMETER',
                             'source_field': self.resolve_ref(fm.group(1), fm.group(2)) if fm else params.get('source-field'),
                             'parameter': self.resolve_ref(pm.group(1), pm.group(2)) if pm else params.get('target-parameter'),
                             'aggregation': (a.find('./agg-type').get('type') if a.find('./agg-type') is not None else None)})
                co = a.find('./clear-option')
                if co is not None:
                    item['on_clear'] = co.get('type') + (f' {co.get("value")}' if co.get('value') else '')
                item['powerbi'] = ('slicer or cross-filter on the source field when the parameter only filters; '
                                   'when it drives a calc, a disconnected table + SELECTEDVALUE measure (set by a slicer)')
                self.flags.append(('PARAMETER_ACTION', item['name'], f'click on "{item["source"].strip(" /")}" sets [{item["parameter"]}] '
                                   f'from [{item["source_field"]}]: {item["powerbi"]}'))
                acts.append(item)
                continue
            if a.tag == 'nav-action':  # "Go to Sheet" button / navigation object
                item.update({'kind': 'NAVIGATE', 'target': (params.get('sheet') or '').strip()})
                self.flags.append(('NAVIGATION_ACTION', item['name'], f'→ {item["target"]}: page navigation button'))
                acts.append(item)
                continue
            if link is not None and (link.get('expression') or '').startswith('tsl:'):
                # internal sheet link = navigation (optionally filtered): tsl:Sheet?[ds].[field]~s0=<[ds].[field]>
                from urllib.parse import unquote
                exp = unquote(link.get('expression')[4:])
                target, _, flt = exp.partition('?')
                item['kind'] = 'NAVIGATE'
                item['target'] = target
                fields = sorted({self.resolve_ref(m.group(1), m.group(2)) for m in QREF.finditer(flt)})
                if fields:
                    item['passes'] = fields
                self.flags.append(('NAVIGATION_ACTION', item['name'], f'→ {target}: page navigation (button/bookmark or drillthrough)'))
            elif link is not None:
                item['kind'] = 'URL'
                # '<[federated.x].[Order ID]>' -> '<[Order ID]>', '<[Parameters].[Parameter 2]>' -> '<[Param:Divested?]>'
                item['url'] = self.resolve_text(link.get('expression'))
                self.flags.append(('URL_ACTION', item['name'], 'field placeholders: ' + ', '.join(re.findall(r'<\[([^\]]+)\]>', item['url'] or ''))))
            elif cmd is not None:
                kind = {'tsc:tsl-filter': 'FILTER', 'tsc:brush': 'HIGHLIGHT', 'tsc:goto-sheet': 'GOTO',
                        'tsc:edit-parameter-action': 'PARAMETER', 'tsc:change-set-action': 'SET'}.get(cmd.get('command'), cmd.get('command'))
                item['kind'] = kind
                params = {p.get('name'): p.get('value') for p in cmd.findall('./param')}
                if 'target' in params:
                    item['target'] = (params.pop('target') or '').strip()
                item.update({k: self.resolve_text(v) for k, v in params.items() if v})
                if kind == 'FILTER':
                    self.flags.append(('FILTER_ACTION', item['name'], f'→ {item.get("target")}: drillthrough/cross-filter candidate'))
                if kind in ('PARAMETER', 'SET'):
                    self.flags.append((kind + '_ACTION', item['name'], 'no direct Power BI equivalent'))
            acts.append(item)
        return acts

    # ---------- assembly ----------
    def build(self):
        r = self.root
        self._dzv = self._zone_visibility()
        wss = [self.worksheet(w) for w in r.findall('./worksheets/worksheet')]
        dbs = [self.dashboard(d) for d in r.findall('./dashboards/dashboard')]
        acts = self.actions()
        windows = [{'class': w.get('class'), 'name': (w.get('name') or '').strip(), 'hidden': w.get('hidden') == 'true'}
                   for w in r.findall('./windows/window')]
        res = {'workbook': {}, 'datasources': [], 'parameters': [], 'calculations': [], 'groups': [], 'sets': [],
               'columns': [], 'worksheets': wss, 'dashboards': dbs, 'actions': acts, 'tabs': windows,
               'orphans': {}, 'flags': [], 'tier': {}}
        n_calc = n_calc_used = 0
        orphan_calcs, orphan_params = [], []
        for d in self.ds.values():
            if d['is_parameters']:
                for (ds, n), c in self.cols.items():
                    if ds != d['name']:
                        continue
                    if (ds, n) in self.used:
                        prm = {'name': c['caption'] or n, 'datatype': c['datatype'], **(c['param'] or {}),
                               'used_by': sorted(self._where(ds, n))}
                        sf = re.match(r'^\[([^\]]+)\]\.\[(.+)\]$', prm.pop('source_field_ref', '') or '')
                        if sf:
                            prm['source_field'] = self.resolve_ref(sf.group(1), sf.group(2))
                            prm['source_datasource'] = (self.ds.get(sf.group(1)) or {}).get('caption')
                        prm['powerbi'] = parameter_pattern(prm)
                        res['parameters'].append(prm)
                    else:
                        orphan_params.append(c['caption'] or n)
                continue
            conn = d['connection'] or {}
            dsrec = {'name': d['caption'], 'connection': conn, 'extract': d['extract'],
                     'filters': [dict(zip(('field', 'rule'), self.describe_filter(f))) for f in d['filters']]}
            if conn.get('published') and not conn.get('tds'):
                self.flags.append(('PUBLISHED_DS', d['caption'],
                                   'joins, custom SQL and source type are NOT in the TWB; get the .tds/.tdsx '
                                   '(Tableau Server: Data Sources > Download, or REST/tabcmd) and re-run with --tds'))
            if d['extract']:
                self.flags.append(('EXTRACT', d['caption'], 'workbook uses an extract; confirm the live source for Power BI'))
            lg = conn.get('tds', conn).get('logical') or {}
            for rl in lg.get('relationships', []):
                keys = re.findall(r'\[([^\]]+?)(?: \([^)]*\))?\]', rl['on'] or '')
                if len(keys) == 2 and keys[0].lower() != keys[1].lower():
                    self.flags.append(('KEY_NAME_MISMATCH', d['caption'], f'{" ↔ ".join(x or "?" for x in rl["between"])} '
                                       f'ON {rl["on"]}: keys have different names; confirm the relationship'))
            rels = json.dumps(conn.get('tds', conn))
            if 'custom_sql' in rels:
                self.flags.append(('CUSTOM_SQL', d['caption'], 'move to a view or Power Query native query'))
            if '"join":' in rels:
                self.flags.append(('PHYSICAL_JOIN', d['caption'], 'tables are joined before Tableau aggregates: SUM/AVG/COUNT '
                                   'of a column from the "one" side counts it once per joined row. Power BI relationships '
                                   'count it once; check those totals (or keep the join in Power Query)'))
            res['datasources'].append(dsrec)
            all_cols = [(k, c) for k, c in self.cols.items() if k[0] == d['name']]
            for (ds, n), c in all_cols:
                if n.startswith(':') or n in ('Multiple Values', 'Number of Records') and not c['formula']:
                    if n == ':Measure Names' and c['aliases']:  # renamed measure headers
                        res.setdefault('measure_aliases', {}).update(
                            {self.resolve_text(k.strip('"').replace('\\', '')): v for k, v in c['aliases'].items()})
                    continue  # Tableau-generated fields, not source columns
                is_calc = c['formula'] is not None or c['bins'] is not None
                used = (ds, n) in self.used
                if is_calc:
                    n_calc += 1
                if not used:
                    if is_calc:
                        orphan_calcs.append(c['caption'] or n)
                    continue
                where = sorted(self._where(ds, n))
                if c['bins'] and c['bins'].get('groups'):
                    src = self.cols.get((ds, c['bins']['source'])) or {}
                    gt = self.model_table(src) if src else None
                    res['groups'].append({'name': self.disp(ds, n), 'source': self.disp(ds, c['bins']['source']),
                                          'default': c['bins']['default'], 'groups': c['bins']['groups'], 'used_in': where,
                                          'table': gt, 'dax': group_column(gt or '<Table>', self.disp(ds, c['bins']['source']),
                                                                           c['bins']['groups'], c['bins']['default'])})
                    continue
                if is_calc:
                    n_calc_used += 1
                    f = self.resolve_formula(ds, c['formula'])
                    flags = self._calc_flags(ds, n, c, f)
                    tr = self.calc_translation((ds, n))
                    calc = {'name': self.disp(ds, n), 'tableau_id': n if c['caption'] else None,
                            'datasource': d['caption'],
                            'datatype': c['datatype'], 'role': c['role'], 'formula': f,
                            'used_in': where, 'flags': flags, 'pbi_kind': tr['kind'],
                            'table': self.calc_table((ds, n)) if tr['kind'] == 'calculated column' else None,
                            'dax': tr['dax'], 'confidence': tr['confidence'], 'dax_notes': tr['notes'], 'dax_fields': tr.get('fields')}
                    if c['formula'] is not None and not self.formula_refs(ds, c['formula']):
                        calc['constant'] = True  # no field: a literal used for layout (spacer, help icon colour, ...)
                        if not calc['name'].strip(' "\''):
                            calc['placeholder'] = True
                    if d['folders'].get(n):
                        calc['folder'] = d['folders'][n]
                    grain = self._joined_grain(ds, n, c, tr)
                    if grain:
                        calc['flags'] = flags + ['JOINED_GRAIN']
                        calc['joined_grain'] = grain
                    res['calculations'].append(calc)
                    continue
                rem = c['remote']
                col = {'name': self.disp(ds, n), 'datatype': c['datatype'], 'role': c['role'], 'datasource': d['caption'],
                       'table': self.model_table(c), 'source_column': rem[2] if rem else None,
                       'source': f'{rem[1] or strip_brackets(rem[0] or "")}.{rem[2]}' if rem else None,
                       'aliases': c['aliases'] or None, 'format': c['default_format'],
                       'color_map': d['palettes'].get(f'[{n}]') or d['palettes'].get(f'[{d["name"]}].[{n}]')}
                for k, v in (('folder', d['folders'].get(n)), ('geo_role', c.get('geo_role')),
                             ('sort_order', d['default_sorts'].get(n))):
                    if v:
                        col[k] = v
                res['columns'].append(col)
            for sname, g in d['sets'].items():
                if (d['name'], sname) in self.used or any(f'[{sname}]' in ET.tostring(w, encoding='unicode')
                                                          for w in r.findall('./worksheets/worksheet')):
                    res['sets'].append({'name': g.get('caption') or sname,
                                        'rule': self.gf(g.find('./groupfilter')) if g.find('./groupfilter') is not None else None})
                    self.flags.append(('SET', g.get('caption') or sname, 'model as a calculated column/table or a filter'))
        # data-source colour palettes keyed by readable field label, e.g. 'YEAR(WO APPR Date)' -> {'2025': '#1f77b4'}
        res['palettes'] = {}
        for d in self.ds.values():
            for key, pal in d.get('palettes', {}).items():
                m = re.match(r'^\[([^\[\]]+)\]$', key or '')
                if m and pal:
                    res['palettes'][self.resolve_ref(d['name'], m.group(1))] = pal
        # same name + formula (worksheet-local copies) -> one entry; same formula under different names -> flag
        merged_calcs, by_formula = {}, {}
        for c in res['calculations']:
            # the DAX names the table: the same formula in two data sources (Retail "Last Week" on Store/Category and on
            # Chain/Class) is two columns on two tables, and a calc of each source must find its own
            key = (c['name'], c['formula'], c.get('dax'))
            if key in merged_calcs:
                merged_calcs[key]['used_in'] = sorted(set(merged_calcs[key]['used_in']) | set(c['used_in']))
                n_calc_used -= 1
            else:
                merged_calcs[key] = c
                by_formula.setdefault(re.sub(r'\s+', '', (c['formula'] or '').lower()), []).append(c['name'])
        res['calculations'] = list(merged_calcs.values())
        # same name, different formulas, all used: say which data source / sheets use which version
        by_name = {}
        for c in res['calculations']:
            by_name.setdefault(c['name'], []).append(c)
        for name, versions in by_name.items():
            if len(versions) > 1:
                for c in versions:
                    c['flags'] = c['flags'] + ['NAME_COLLISION']
                self.flags.append(('NAME_COLLISION', name, f'{len(versions)} different formulas share this name: ' + '; '.join(
                    f'[{c["datasource"]}] used in {", ".join(c["used_in"]) or "-"}' for c in versions)))
        for f, names in by_formula.items():
            if len(set(names)) > 1 and len(f) > 12:
                self.flags.append(('DUPLICATE_LOGIC', ', '.join(sorted(set(names))), 'same formula: build one measure'))
        used_names = {c['name'] for c in res['calculations']}  # unused copies of a used calc are not orphans
        res['orphans'] = {'calculations': sorted(set(orphan_calcs) - used_names), 'parameters': sorted(orphan_params),
                          'unused_source_columns': sum(1 for k, c in self.cols.items()
                                                       if k not in self.used and c['formula'] is None and not self.ds[k[0]]['is_parameters'])}
        res['workbook'] = {'worksheets': len(wss), 'dashboards': len(dbs), 'actions': len(acts),
                           'datasources': len(res['datasources']), 'calcs_total': n_calc, 'calcs_used': n_calc_used,
                           'parameters_used': len(res['parameters']), 'source_columns_used': len(res['columns']),
                           'assets': [a for a in self.assets if not a[0].lower().endswith(('.hyper', '.tde'))]}
        imgs = sorted({z.split('IMAGE ')[1].split(' @')[0] for d in dbs for z in d['zones'] if 'IMAGE ' in z})
        if imgs:
            res['workbook']['images'] = imgs
        res['hierarchies'] = self._hierarchies()
        # every active blend link once (the build makes one shared dimension per linked field pair)
        blends = {}
        for ws in wss:
            for s in (ws.get('blend') or {}).get('secondaries', []):
                for lk in s['links']:
                    key = (lk['primary_table'], lk['primary_field'], lk['secondary_table'], lk['secondary_field'], lk['level'])
                    blends.setdefault(key, {**lk, 'sheets': []})['sheets'].append(ws['name'])
        if blends:
            res['blends'] = list(blends.values())
        constants = {c['name'] for c in res['calculations'] if c.get('constant')}
        link_sheets = {(a.get('source') or '').split(' / ')[-1].strip() for a in acts if a.get('kind') in ('URL', 'NAVIGATE', 'GOTO')}
        for ws in wss:
            ws['role'] = sheet_role(ws, constants, link_sheets)
        for (el, val), places in self.unhandled.items():
            self.flags.append(('UNHANDLED_XML', ', '.join(places[:6]) + (f' +{len(places) - 6} more' if len(places) > 6 else ''),
                               f'<{el}> "{val}" is not read by the extractor: check it by hand (--show "<name>")'))
        merged = {}
        for code, where, detail in self.flags:  # one row per finding, listing every place it occurs
            wl = merged.setdefault((code, detail), [])
            if where not in wl:
                wl.append(where)
        res['flags'] = [{'code': c, 'where': ', '.join(w), 'detail': d} for (c, d), w in merged.items()]
        res['tier'] = self._tier(res)
        return res

    def _where(self, ds, n):
        return {w for w in self.used.get((ds, n), set())}

    def _hierarchies(self):
        """Tableau drill paths that are used -> Power BI hierarchies (one table each)."""
        out = []
        for d in self.ds.values():
            for hname, fields in d.get('hierarchies', []):
                keys = [(d['name'], f) for f in fields if (d['name'], f) in self.cols]
                if not any(k in self.used for k in keys):
                    continue
                tables = {self.calc_table(k) if self.cols[k]['formula'] is not None else self.model_table(self.cols[k])
                          for k in keys}
                h = {'name': hname, 'datasource': d['caption'], 'levels': [self.disp(*k) for k in keys],
                     'table': next(iter(tables)) if len(tables) == 1 else None}
                if len(tables) > 1:
                    h['note'] = 'levels come from several tables: a Power BI hierarchy needs them in one table'
                out.append(h)
        return out

    def _row_tables(self, key, seen=None):
        """Source tables a field reaches row by row (through row-level calcs, not through measures)."""
        seen = set() if seen is None else seen
        if key in seen or key[0] == 'Parameters' or key not in self.cols:
            return set()
        seen.add(key)
        c = self.cols[key]
        if c['bins']:
            return self._row_tables((key[0], c['bins']['source']), seen)
        if c['formula'] is None:
            t = self.model_table(c)
            return {t} if t else set()
        if self.calc_translation(key)['kind'] not in ('calculated column', 'field parameter'):
            return set()
        out = set()
        for r in self.formula_refs(key[0], c['formula']):
            out |= self._row_tables(r, seen)
        return out

    def _joined_grain(self, ds, n, c, tr):
        """A row-level value that mixes columns of two tables is computed once per JOINED row in Tableau
        (one per row of the more detailed table). Aggregating it on the other table in Power BI gives another number."""
        if tr['kind'] == 'calculated column':
            ts = self._row_tables((ds, n))
            if len(ts) >= 2:
                return {'level': 'row', 'tables': sorted(ts)}
            return None
        f = clean_formula(c['formula'] or '')
        for m in GRAIN_AGGS.finditer(f):
            depth, i = 0, m.end() - 1
            for i in range(m.end() - 1, len(f)):
                depth += {'(': 1, ')': -1}.get(f[i], 0)
                if depth == 0:
                    break
            ts, lods = set(), []
            for r in self.formula_refs(ds, f[m.end():i]):
                ts |= self._row_tables(r)
                if re.search(r'\{\s*FIXED\b', self.cols[r]['formula'] or '', re.I):
                    lods.append(self.disp(*r))
            agg = m.group(1).upper()
            if len(ts) >= 2:
                self.flags.append(('JOINED_GRAIN', self.disp(ds, n),
                                   f'{agg}() over a row-level value mixing {", ".join(sorted(ts))}: Tableau computes it once '
                                   f'per joined row; in Power BI iterate the more detailed table ({agg}X over it) '
                                   'and compare the numbers'))
                return {'level': agg, 'tables': sorted(ts)}
            src = self._ds_tables(ds)
            if lods and len(src) >= 2:
                # a FIXED value is repeated on every row of the data source: in a multi-table source SUM/AVG can count
                # it once per joined row, while a Power BI measure on the FIXED table counts it once per key
                self.flags.append(('JOINED_GRAIN', self.disp(ds, n),
                                   f'{agg}() over FIXED LOD {", ".join(lods)} in a source with several tables '
                                   f'({", ".join(sorted(src))}): Tableau may repeat the FIXED value on every joined row; '
                                   'compare this total with Tableau before sign-off'))
                return {'level': agg, 'tables': sorted(src), 'fixed': lods}
        return None

    def _ds_tables(self, ds):
        """Source tables behind the used columns of one data source."""
        cache = self.__dict__.setdefault('_ds_tables_cache', {})
        if ds not in cache:
            cache[ds] = {t for k in self.used if k[0] == ds and self.cols.get(k) and self.cols[k]['formula'] is None
                         for t in [self.model_table(self.cols[k])] if t}
        return cache[ds]

    def _calc_flags(self, ds, n, c, f):
        fl = []
        txt = f or ''
        name = self.disp(ds, n)
        for kind in ('FIXED', 'INCLUDE', 'EXCLUDE'):
            if re.search(r'\{\s*' + kind + r'\b', txt, re.I):
                fl.append('LOD_' + kind)
        if TABLE_CALC.search(txt):
            fl.append('TABLE_CALC')
        if RLS_FUNCS.search(txt):
            fl.append('RLS')
            # USERNAME() compared with a column: the value changes per signed-in user (Tableau "per-user column").
            # Power BI columns cannot do that: needs a user table + RLS role, and usually one row copy per viewer.
            if PER_USER.search(txt):
                fl.append('PER_USER_COLUMN')
                self.flags.append(('PER_USER_COLUMN', name, 'value depends on the signed-in user: do not translate '
                                   'USERNAME() 1:1; design RLS (user table mapping UPN -> login) and per-viewer rows'))
            ids = sorted(set(USER_LITERAL.findall(txt)))
            if ids:
                self.flags.append(('HARDCODED_USERS', name, f'logins {", ".join(ids)} hard-coded: move them to the '
                                   'user / security table (e.g. an admin flag) instead of measures'))
        m = RISKY_FUNCS.search(txt)
        if m:
            fl.append('UNSUPPORTED_FN:' + m.group(1).upper())
        if 'ATTR(' in txt.upper():
            fl.append('ATTR')
        refs = self.formula_refs(ds, c['formula'])
        p_dates = [r for r in refs if r[0] == 'Parameters' and self.cols[r]['datatype'] == 'date']
        dt_cols = [r for r in refs if r[0] != 'Parameters' and self.cols[r]['datatype'] == 'datetime']
        if p_dates and dt_cols and re.search(r'[<>]=?', txt):
            fl.append('DATE_PARAM_VS_DATETIME')
            self.flags.append(('DATE_PARAM_VS_DATETIME', name, 'date parameter compared to a datetime field: '
                               '"<= End Date" excludes the end day after midnight; confirm inclusive/exclusive'))
        nested = [r for r in refs if self.cols[r]['formula']]
        if len(nested) >= 3:
            fl.append(f'NESTED({len(nested)})')
        for code in fl:
            if code.startswith(('LOD_', 'TABLE_CALC', 'RLS', 'UNSUPPORTED')):
                self.flags.append((code.split(':')[0], name, (f or '')[:120]))
        return fl

    def _tier(self, res):
        lvl = {'Low': 0, 'Medium': 1, 'High': 2}
        names = ['Low', 'Medium', 'High']
        codes = {f['code'] for f in res['flags']}
        hard = sum(1 for c in res['calculations'] if any(x.startswith(('LOD_', 'TABLE_CALC', 'UNSUPPORTED')) for x in c['flags']))
        nc = len(res['calculations'])
        f = {}
        f['calculated fields'] = (2 if nc >= 16 or hard >= 3 else 1 if nc >= 6 or hard else 0, f'{nc} used, {hard} hard')
        vis = max([sum(1 for z in d['zones'] if 'SHEET ' in z) for d in res['dashboards']] or [0])
        f['visuals per page'] = (2 if vis >= 13 else 1 if vis >= 7 else 0, f'max {vis}')
        pages = len(res['dashboards']) or len(res['worksheets'])
        f['pages'] = (2 if pages >= 6 else 1 if pages >= 3 else 0, str(pages))
        nds = len(res['datasources'])
        f['data sources'] = (2 if nds >= 4 or 'BLENDING' in codes else 1 if nds >= 2 else 0, str(nds))
        f['RLS'] = (1 if 'RLS' in codes else 0, 'user functions found' if 'RLS' in codes else 'none in TWB')
        inter = 2 if codes & {'PARAMETER_ACTION', 'SET_ACTION', 'SET'} or len(res['parameters']) > 2 else \
            1 if 'FILTER_ACTION' in codes or res['parameters'] else 0
        f['interactivity'] = (inter, ', '.join(sorted(codes & {'FILTER_ACTION', 'URL_ACTION', 'PARAMETER_ACTION', 'SET_ACTION'})) +
                              f'; {len(res["parameters"])} params')
        f['tooltips'] = (2 if 'VIZ_IN_TOOLTIP' in codes else 0, 'viz-in-tooltip' if 'VIZ_IN_TOOLTIP' in codes else 'default/custom')
        top = max(v[0] for v in f.values())
        return {'suggested': names[top], 'factors': {k: f'{names[v[0]]} ({v[1]})' for k, v in f.items()},
                'note': 'highest factor wins (complexity-tiering.md); a human confirms the tier'}


def sheet_role(ws, constants=(), link_sheets=()):
    """What a worksheet is for: data (a real visual), button (link/navigation shape), help ("?" icon with a tooltip),
    text (title only) or decoration (spacer/logo shape). Same order of checks as Rebuild.infer_visual."""
    def real(fs):
        return [f for f in fs if f['label'] not in constants]
    enc = real([e for p in ws['panes'] for e in p.get('enc', [])])
    if real(ws['shelf']['rows']) or real(ws['shelf']['cols']) or ws.get('measure_values'):
        return 'data'
    if any(e['channel'] in ('text', 'label') for e in enc):
        return 'data'  # card / KPI
    if ws['name'] in link_sheets:
        return 'button'
    if any(p.get('tooltip') for p in ws['panes']):
        return 'help'
    if any(e['channel'] == 'color' for e in enc):
        return 'data'  # status indicator: shape coloured by a field
    return 'text' if ws.get('title') and not enc else 'decoration'


def extract_tables(path, export_dir=None):
    """Tables, columns and row counts of the .hyper extracts packaged in a .twbx. Needs the optional tableauhyperapi
    package; without it (or when its engine cannot start) each extract is listed with the reason.
    export_dir: also write each table to <export_dir>/<table>.csv (real rows, for a sample build)."""
    if not zipfile.is_zipfile(path):
        return []
    with zipfile.ZipFile(path) as z:
        hypers = [n for n in z.namelist() if n.lower().endswith('.hyper')]
        if not hypers:
            return []
        try:
            from tableauhyperapi import HyperProcess, Telemetry, Connection
        except ImportError:
            return [{'file': n, 'status': 'not read: install the optional package (pip install tableauhyperapi)'}
                    for n in hypers]
        for _ in range(4):  # the Hyper engine sometimes fails to start on Windows; a retry usually works
            out = _read_hypers(z, hypers, HyperProcess, Telemetry, Connection, export_dir)
            if not any('Failed to start' in e.get('status', '') for e in out):
                break
        return out


def _read_hypers(z, hypers, HyperProcess, Telemetry, Connection, export_dir):
    import shutil
    import tempfile
    tmp = tempfile.mkdtemp(prefix='bhayankar_hyper_')
    out = []
    try:
        with HyperProcess(Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU, parameters={'log_dir': tmp}) as hp:
            for n in hypers:
                stem = re.sub(r'[^A-Za-z0-9_.-]+', '_', os.path.splitext(os.path.basename(n))[0]).strip('_')
                with Connection(hp.endpoint, z.extract(n, tmp)) as c:
                    for s in c.catalog.get_schema_names():
                        for t in c.catalog.get_table_names(s):
                            cols = [(str(x.name).strip('"'), str(x.type)) for x in c.catalog.get_table_definition(t).columns]
                            info = {'file': n, 'table': str(t), 'rows': c.execute_scalar_query(f'SELECT COUNT(*) FROM {t}'),
                                    'columns': cols}
                            if export_dir:  # one CSV per extract table, named <extract file>__<table>.csv
                                os.makedirs(export_dir, exist_ok=True)
                                fn = stem + '__' + re.sub(r'[^A-Za-z0-9_.-]+', '_', str(t).replace('"', '')).strip('_') + '.csv'
                                with open(os.path.join(export_dir, fn), 'w', newline='', encoding='utf-8') as f:
                                    wr = csv.writer(f)
                                    wr.writerow([cn for cn, _ in cols])
                                    with c.execute_query(f'SELECT * FROM {t}') as rows:
                                        for row in rows:
                                            wr.writerow(['' if v is None else str(v) for v in row])
                                info['csv'] = fn
                            out.append(info)
    except Exception as e:  # engine blocked (antivirus, old package) or a damaged extract
        reason = (str(e).strip().splitlines() or [type(e).__name__])[0][:160]
        out = [{'file': n, 'status': f'not read: {reason}'} for n in hypers]
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return out


# ---------- rendering ----------
SPLIT_BYTES = 40000  # above this the digest is split: common file + one file per dashboard


def _filter_txt(f):
    return f'{f["field"]} = {f["rule"]}' + (' (context)' if f.get('context') else '')


def common_filters(worksheets, min_sheets=3):
    """Non-interactive filters repeated on >= min_sheets sheets get an ID (F1, F2...) and are printed once."""
    count = {}
    for ws in worksheets:
        for f in ws['filters']:
            if not f['rule'].startswith('ALL'):
                count[_filter_txt(f)] = count.get(_filter_txt(f), 0) + 1
    many = [t for t, n in count.items() if n >= min(min_sheets, max(2, len(worksheets)))]
    return {t: f'F{i + 1}' for i, t in enumerate(many)}


def _used(names, scope=None, cap=6):
    """'used in' list: limited to the page's sheets (and calc dependencies) when scope is given, then capped."""
    if scope is not None:
        names = [n for n in names if n in scope or n.startswith('via:')]
    shown = ', '.join(names[:cap])
    return shown + (f' +{len(names) - cap} more' if len(names) > cap else '')


def render_calcs(calcs, groups, sets, scope=None):
    L = []
    for c in calcs:
        fl = f' **[{", ".join(c["flags"])}]**' if c['flags'] else ''
        pbi = f' → {c.get("pbi_kind")} [{c.get("confidence")}]' if c.get('pbi_kind') else ''
        src = f' from [{c["datasource"]}]' if 'NAME_COLLISION' in c['flags'] and c.get('datasource') else ''
        L.append(f'- **{c["name"]}**{src} ({c["datatype"]}/{c["role"]}){fl}{pbi}: `{c["formula"]}`  \n  used in: {_used(c["used_in"], scope)}')
    for g in groups:
        L.append(f'- **{g["name"]}** = GROUP of [{g["source"]}], default {g["default"]}: ' +
                 '; '.join(f'{k}={",".join(v)}' for k, v in g['groups'].items()) + f'  \n  used in: {_used(g["used_in"], scope)}')
    for s in sets:
        L.append(f'- **{s["name"]}** = SET {s["rule"]}')
    return L


def render_worksheets(worksheets, fids):
    L = []
    for ws in worksheets:
        role = ws.get('role')
        L.append(f'### {ws["name"]}' + (f'  ({role} sheet: no data visual)' if role and role != 'data' else ''))
        if ws['title']:
            L.append(f'- title: "{ws["title"]}"')
        L.append(f'- rows: `{ws["rows"] or ""}`  \n- cols: `{ws["cols"] or ""}`' + ('  \n- grand totals on' if ws.get('grand_totals') else ''))
        lines = []  # collapse identical consecutive panes (dual-axis copies)
        tips = list(dict.fromkeys(p['tooltip'][:160] for p in ws['panes'] if p.get('tooltip')))
        for p in ws['panes']:
            extra = f'; label: {p["label"]}' if p.get('label') else ''
            extra += f'; {p["trendline"]} trend line' if p.get('trendline') else ''
            ln = (f'mark {p["mark"]}; {", ".join(p["encodings"]) or "no encodings"}'
                  + ('; labels shown' if p.get('labels_shown') else '') + extra)
            if lines and lines[-1][1] == ln:
                lines[-1][0] += 1
            else:
                lines.append([1, ln])
        for n, ln in lines:
            L.append(f'- {"pane" if n == 1 else f"{n} panes"}: {ln}')
        for t in tips:
            L.append(f'- custom tooltip: "{t}"')
        if ws.get('measure_values'):
            L.append(f'- measure values: {", ".join(ws["measure_values"])}')
        interactive = [f['field'] for f in ws['filters'] if f['rule'].startswith('ALL')]
        ids = [fids[_filter_txt(f)] for f in ws['filters'] if _filter_txt(f) in fids]
        if ids:
            L.append(f'- common filters: {", ".join(ids)}')
        for f in ws['filters']:
            if not f['rule'].startswith('ALL') and _filter_txt(f) not in fids:
                L.append(f'- filter: {_filter_txt(f)}')
        if interactive:
            L.append(f'- interactive filters (no saved selection): {", ".join(interactive)}')
        for s in ws.get('sorts', []):
            L.append(f'- sort: {s}')
        if ws.get('subtotals'):
            L.append(f'- subtotals by: {", ".join(ws["subtotals"])}')
        tcs = {f['label']: f for f in ws['shelf']['rows'] + ws['shelf']['cols'] + [e for p in ws['panes'] for e in p.get('enc', [])]
               if f.get('tc_ordering') or f.get('tc_fields')}
        for lb, f in tcs.items():
            L.append(f'- table calc {lb}: compute using ' + (', '.join(f['tc_fields']) if f.get('tc_fields') else f.get('tc_ordering', '?'))
                     + (f' ({f["tc_type"]})' if f.get('tc_type') else ''))
        for r in ws.get('reference_lines', []):
            L.append(f'- ref line: {r["formula"]} of {r["value"]} on axis {r["axis"]} ({r["scope"]}), label "{r["label"]}"')
        if ws.get('number_formats'):
            L.append('- formats: ' + '; '.join(f'{k} {v}' for k, v in ws['number_formats'].items()))
        if ws.get('colors'):
            L.append('- colors: ' + json.dumps(ws['colors'], ensure_ascii=False))
    return L


def rel_txt(r):
    """Physical relation tree -> 'A LEFT JOIN B ON x = y' / custom SQL."""
    if not r:
        return '?'
    if 'table' in r and 'join' not in r:
        return r['table'] or r.get('alias') or '?'
    if 'custom_sql' in r:
        return f'CUSTOM SQL "{r.get("alias")}": {r["custom_sql"][:400]}'
    if 'join' in r:
        ins = [rel_txt(x) for x in r['inputs']]
        return f'({ins[0]} {r["join"].upper()} JOIN {" , ".join(ins[1:])} ON {" AND ".join(o for o in r["on"] if o)})'
    if 'union' in r:
        return 'UNION(' + ', '.join(rel_txt(x) for x in r['union']) + ')'
    return json.dumps(r)


def conn_txt(c):
    named = ', '.join(' '.join(x for x in (n.get('class'), n.get('server') and f'server={n["server"]}',
                                           n.get('dbname') and f'db={n["dbname"]}', n.get('schema') and f'schema={n["schema"]}') if x)
                      for n in c.get('named') or [])
    out = [f'source: {named or c.get("class")}']
    lg = c.get('logical')
    if lg and (lg['relationships'] or len(lg['objects']) > 1):
        out.append('logical tables: ' + '; '.join(f'{o}: {rel_txt(c["physical_by_object"].get(o))}' for o in lg['objects']))
        out += [f'relationship: {" ↔ ".join(x or "?" for x in r["between"])} ON {r["on"]}' for r in lg['relationships']]
    else:
        out.append('tables: ' + rel_txt(c.get('relations')))
    return out


def render_common(res, src_name, fids, split_index=None):
    L = []
    w = res['workbook']
    L.append(f'# TWB digest: {src_name}')
    L.append('Generated by bhayankar_extractor.py. Read this instead of the TWB. For raw XML of one item: '
             '`python bhayankar_extractor.py <twb> --show "<name>"`.\n')
    L.append(f'**Counts:** {w["worksheets"]} worksheets, {w["dashboards"]} dashboards, {w["datasources"]} data sources, '
             f'{w["calcs_used"]}/{w["calcs_total"]} calcs used, {w["parameters_used"]} parameters used, '
             f'{w["source_columns_used"]} source columns used, {w["actions"]} actions.  ')
    L.append(f'**Suggested tier:** {res["tier"]["suggested"]} — ' + '; '.join(f'{k}: {v}' for k, v in res['tier']['factors'].items()))
    if w.get('images'):
        L.append(f'**Images:** {", ".join(w["images"])}')
    if split_index:
        L.append('\n**Split digest.** Worksheets and calculations are in one file per dashboard (each file is self-contained):')
        L.extend(f'- `{fn}`: {desc}' for fn, desc in split_index)
    L.append('\n## Flags (seed for open questions)')
    L.append('| Code | Where | Detail |\n|---|---|---|')
    for f in res['flags']:
        L.append(f'| {f["code"]} | {f["where"]} | {(f["detail"] or "").replace("|", "/")} |')
    L.append('\n## Data sources')
    for d in res['datasources']:
        c = d['connection'] or {}
        if c.get('published'):
            L.append(f'- **{d["name"]}**: PUBLISHED (`sqlproxy`) server `{c.get("server")}`, name `{c.get("dbname")}`'
                     + (f', path `{c["path"][0]}`' if c.get('path') else ''))
            if c.get('tds'):
                L.extend(f'  - (.tds) {x}' for x in conn_txt(c['tds']))
        else:
            L.append(f'- **{d["name"]}**')
            L.extend(f'  - {x}' for x in conn_txt(c))
        if d['extract']:
            L.append('  - extract enabled')
        for f in d['filters']:
            L.append(f'  - datasource filter: {f["field"]} {f["rule"]}')
    tables = sorted({c['source'].split('.')[0] for c in res['columns'] if c['source']})
    if tables:
        L.append(f'- Tables behind used columns (from metadata): {", ".join(tables)}')
    L.append('\n## Source columns used')
    L.append('| Field | Type | Role | Source table.field | Notes |\n|---|---|---|---|---|')
    for c in sorted(res['columns'], key=lambda x: x['name']):
        notes = []
        if c['aliases']:
            notes.append('aliases ' + json.dumps(c['aliases'], ensure_ascii=False))
        if c['format']:
            notes.append('fmt ' + c['format'])
        if c['color_map']:
            notes.append('colors ' + json.dumps(c['color_map'], ensure_ascii=False))
        if c.get('geo_role'):
            notes.append('geo ' + c['geo_role'])
        if c.get('sort_order'):
            notes.append('sort ' + ', '.join(c['sort_order'])[:120])
        if c.get('folder'):
            notes.append('folder ' + c['folder'])
        L.append(f'| {c["name"]} | {c["datatype"]} | {c["role"]} | {c["source"] or "?"} | {"; ".join(notes)} |')
    if res.get('hierarchies'):
        L.append('\n## Hierarchies (Tableau drill paths)')
        L.extend(f'- **{h["name"]}** ({h["table"] or "several tables"}): {" > ".join(h["levels"])}'
                 + (f' — {h["note"]}' if h.get('note') else '') for h in res['hierarchies'])
    if res.get('extracts'):
        L.append('\n## Packaged extracts (.hyper)')
        for e in res['extracts']:
            if 'table' in e:
                L.append(f'- `{e["file"]}` {e["table"]}: {e["rows"]:,} rows, {len(e["columns"])} columns'
                         + (f' → extract_data\\{e["csv"]}' if e.get('csv') else ''))
            else:
                L.append(f'- `{e["file"]}`: {e["status"]}')
    if res['parameters']:
        L.append('\n## Parameters (used)')
        for p in res['parameters']:
            dom = {k: v for k, v in p.items() if k not in ('name', 'datatype', 'used_by', 'domain', 'value')}
            L.append(f'- **{p["name"]}** ({p["datatype"]}, {p.get("domain")}) default `{p.get("value")}` '
                     f'{json.dumps(dom, ensure_ascii=False) if dom else ""} — used by: {_used(p["used_by"])}')
    if res.get('measure_aliases'):
        L.append('\n## Measure header aliases (display names for measures in Measure Names/Values)')
        L.extend(f'- {k} → "{v}"' for k, v in res['measure_aliases'].items())
    if fids:
        L.append('\n## Common filters (referenced by ID in worksheets)')
        L.extend(f'- **{i}**: {t}' for t, i in fids.items())
    L.append('\n## Dashboards')
    for d in res['dashboards']:
        if split_index:  # layouts live in the page files
            L.append(f'- {d["name"]} ({d["size"]}): {sum(1 for z in d["zones"] if "SHEET " in z)} sheets, '
                     f'{sum(1 for z in d["zones"] if "FILTER " in z or "PARAM " in z)} filter/parameter controls')
            continue
        L.append(f'### {d["name"]}  ({d["size"]})')
        L.append('```')
        L.extend(d['zones'])
        L.append('```')
    vis = [t for t in res['tabs'] if not t['hidden']]
    if res['tabs']:
        L.append('**Tabs visible:** ' + ', '.join(f'{t["name"]} ({t["class"]})' for t in vis) +
                 ('; hidden: ' + ', '.join(f'{t["name"]} ({t["class"]})' for t in res['tabs'] if t['hidden'])
                  if len(vis) != len(res['tabs']) else ''))
    if res['actions']:
        L.append('\n## Actions')
        for a in res['actions']:
            rest = {k: v for k, v in a.items() if k not in ('name', 'kind', 'source', 'trigger')}
            L.append(f'- **{a["name"]}** {a.get("kind")} on {a["trigger"]} from "{a["source"]}" ' + json.dumps(rest, ensure_ascii=False))
    o = res['orphans']
    L.append('\n## Not migrated (orphans)')
    L.append(f'- {len(o["calculations"])} unused calcs: {", ".join(o["calculations"])}')
    if o['parameters']:
        L.append(f'- {len(o["parameters"])} unused parameters: {", ".join(o["parameters"])}')
    L.append(f'- {o["unused_source_columns"]} unused source columns (not listed)')
    return L


def calcs_for(res, sheet_names):
    """Calcs/groups needed by these sheets, including the calcs they depend on (via:...)."""
    names = set(sheet_names)
    items = res['calculations'] + res['groups']
    need = set()
    changed = True
    while changed:
        changed = False
        for c in items:
            if c['name'] in need:
                continue
            if any(u in names or (u.startswith('via:') and u[4:] in need) for u in c['used_in']):
                need.add(c['name'])
                changed = True
    return ([c for c in res['calculations'] if c['name'] in need],
            [g for g in res['groups'] if g['name'] in need])


def render_md(res, src_name, in_bytes, split=None, spec=None):
    """Returns {filename: markdown}. Splits into per-dashboard files when large (or split=True)."""
    fids = common_filters(res['worksheets'])
    pages = {p['name']: p for p in (spec or {}).get('pages', [])}
    rebuild_hdr = ['\n## Power BI rebuild plan (pages)',
                   'Visual types and roles checked against the powerbi-report-author catalog; positions in px on the '
                   'Tableau canvas. [high]/[medium]/[low] = confidence. Full detail: `bhayankar_rebuild.json`.']
    full = render_common(res, src_name, fids)
    full += ['\n## Calculations (used only)'] + render_calcs(res['calculations'], res['groups'], res['sets'])
    full += ['\n## Worksheets'] + render_worksheets(res['worksheets'], fids)
    if spec:
        full += [''] + render_model(spec['model']) + rebuild_hdr
        for pg in spec['pages']:
            full += render_page(pg)
    md = '\n'.join(full) + '\n'
    if split is False or (split is None and len(md.encode('utf-8')) <= SPLIT_BYTES):
        md += f'\n_Size: TWB {in_bytes:,} bytes → digest {len(md.encode("utf-8")):,} bytes (~{len(md) // 4:,} tokens)._\n'
        return {'bhayankar_digest.md': md}
    files, index = {}, []
    by_name = {ws['name']: ws for ws in res['worksheets']}
    placed = set()
    for i, d in enumerate(res['dashboards'], 1):
        sheets = [m.group(1) for z in d['zones'] for m in [re.search(r'SHEET "(.+?)"', z)] if m and m.group(1) in by_name]
        sheets = list(dict.fromkeys(sheets))
        placed.update(sheets)
        calcs, groups = calcs_for(res, sheets + ['dashboard:' + d['name']])
        fn = f'bhayankar_digest_{i:02d}_{re.sub(r"[^A-Za-z0-9]+", "_", d["name"]).strip("_")[:50]}.md'
        body = [f'# Dashboard: {d["name"]}  ({d["size"]})',
                f'Part of the `{src_name}` digest; data sources, columns, parameters, common filters (F#) and flags are in `bhayankar_digest.md`.',
                '```'] + d['zones'] + ['```', '\n## Calculations needed'] + render_calcs(calcs, groups, [], set(sheets)) + \
               ['\n## Worksheets'] + render_worksheets([by_name[s] for s in sheets], fids)
        if d['name'] in pages:
            body += rebuild_hdr + render_page(pages.pop(d['name']))
        files[fn] = '\n'.join(body) + '\n'
        index.append((fn, f'{len(sheets)} sheets, {len(calcs) + len(groups)} calcs, ~{len(files[fn]) // 4:,} tokens'))
    rest = [ws for ws in res['worksheets'] if ws['name'] not in placed]
    if rest:
        calcs, groups = calcs_for(res, [ws['name'] for ws in rest])
        fn = 'bhayankar_digest_99_sheets_not_on_a_dashboard.md'
        files[fn] = '\n'.join(['# Worksheets not placed on any dashboard',
                               'Visible tabs or helper sheets; check the Tabs line in `bhayankar_digest.md` before migrating.',
                               '\n## Calculations needed'] + render_calcs(calcs, groups, [], {ws['name'] for ws in rest}) +
                              ['\n## Worksheets'] + render_worksheets(rest, fids) +
                              (rebuild_hdr + [ln for pg in pages.values() for ln in render_page(pg)] if pages else [])) + '\n'
        index.append((fn, f'{len(rest)} sheets'))
    common = render_common(res, src_name, fids, index)
    if res['sets']:
        common += ['\n## Sets'] + render_calcs([], [], res['sets'])
    if spec:
        common += [''] + render_model(spec['model'])
    files['bhayankar_digest.md'] = '\n'.join(common) + '\n'
    total = sum(len(v.encode('utf-8')) for v in files.values())
    files['bhayankar_digest.md'] += (f'\n_Size: TWB {in_bytes:,} bytes → {len(files)} digest files, {total:,} bytes in all '
                               f'(~{total // 4:,} tokens); a single page needs this file + its page file._\n')
    return files


def show(root, name):
    """Print the pruned XML of one artifact (worksheet, dashboard, calc, parameter, action)."""
    def prune(e):
        for ch in list(e):
            if ch.tag in ('datasource-dependencies', 'style', 'zone-style', 'simple-id', 'repository-location',
                          'thumbnail', 'metadata-records', 'layout-options'):
                e.remove(ch)
            else:
                prune(ch)
    hits = []
    for e in root.iter():
        if e.tag in ('worksheet', 'dashboard', 'action', 'window') and (e.get('name') or '').strip() == name.strip() \
                or e.tag == 'action' and e.get('caption') == name:
            hits.append(e)
        elif e.tag == 'column' and name in (e.get('caption'), strip_brackets(e.get('name') or '')) and e.find('calculation') is not None:
            hits.append(e)
            break
    for h in hits[:1] + [x for x in hits[1:] if x.tag != 'window']:
        import copy
        c = copy.deepcopy(h)
        prune(c)
        print(ET.tostring(c, encoding='unicode'))
    if not hits:
        print(f'No artifact named "{name}"')


def process(path, tds, out_dir, split=None, table_map=None, strict_tds=False, extract_data=False):
    root, assets = load_root(path)
    troots = [(load_root(t, ('.tds',))[0], t) for t in tds]
    wb = Workbook(root, assets, troots, table_map, strict_tds)
    res = wb.build()
    res['workbook']['tds_merged'] = wb.tds_used
    ext = extract_tables(path, os.path.join(out_dir, 'extract_data') if extract_data else None)
    if ext:
        res['extracts'] = ext
        read = [e for e in ext if 'table' in e]
        if read:
            res['flags'].append({'code': 'EXTRACT_DATA', 'where': os.path.basename(path), 'detail':
                                 f'{len(read)} extract table(s), {sum(e["rows"] for e in read):,} rows: '
                                 + ('written to extract_data\\ (real rows for a sample build)' if extract_data else
                                    're-run with --extract-data to write them as CSV for a sample build')})
        else:
            res['flags'].append({'code': 'EXTRACT_DATA', 'where': os.path.basename(path),
                                 'detail': f'{len(ext)} packaged .hyper extract(s) {ext[0]["status"]}'})
    if path in RECOVERED:
        res['flags'].insert(0, {'code': 'TRUNCATED_FILE', 'where': os.path.basename(path),
                                'detail': f'{RECOVERED[path]}: digest built from the readable part only; '
                                          're-download the workbook and re-run before building'})
    os.makedirs(out_dir, exist_ok=True)
    for old in os.listdir(out_dir):  # stale page files from an earlier split run
        if old.startswith('bhayankar_digest_') and old.endswith('.md'):
            os.remove(os.path.join(out_dir, old))
    spec = Rebuild(res).build()
    files = render_md(res, os.path.basename(path), os.path.getsize(path), split, spec)
    with open(os.path.join(out_dir, 'bhayankar_rebuild.json'), 'w', encoding='utf-8') as f:
        json.dump(spec, f, indent=1, ensure_ascii=False, default=str)
    with open(os.path.join(out_dir, 'bhayankar_measures.dax'), 'w', encoding='utf-8') as f:
        f.write(dax_script(spec))
    for fn, text in files.items():
        with open(os.path.join(out_dir, fn), 'w', encoding='utf-8') as f:
            f.write(text)
    with open(os.path.join(out_dir, 'bhayankar_digest.json'), 'w', encoding='utf-8') as f:
        json.dump(res, f, indent=1, ensure_ascii=False, default=str)
    return res, files


HERE = os.path.dirname(os.path.abspath(__file__))
import bhayankar_toolkit as bt  # noqa: E402
INPUT_DIR = bt.input_dir()     # drop .twb/.twbx (+ .tds/.tdsx, table_map.json) here (or set BHAYANKAR_INPUT)
OUTPUT_DIR = bt.output_dir()   # one sub-folder per workbook + inventory.csv (or set BHAYANKAR_OUTPUT)


def run_folder(folder, out, split=None, table_map=None, extract_data=False):
    """Process every workbook under `folder`; .tds/.tdsx files in the same folder are matched to workbooks by name."""
    wbs, tds = [], []
    for dp, dirs, fs in os.walk(folder):
        dirs[:] = [d for d in dirs if os.path.abspath(os.path.join(dp, d)) != os.path.abspath(out)
                   and d not in ('_digests', 'Extraction')]
        for fn in sorted(fs):
            p = os.path.join(dp, fn)
            if fn.startswith(('~$', '.')):
                continue
            if fn.lower().endswith(('.twb', '.twbx')):
                wbs.append(p)
            elif fn.lower().endswith(('.tds', '.tdsx')):
                tds.append(p)
    rows, used = [], set()
    for p in wbs:
        fn = os.path.basename(p)
        try:
            res, files = process(p, tds, os.path.join(out, os.path.splitext(fn)[0]), split, table_map, strict_tds=True,
                                 extract_data=extract_data)
            w = res['workbook']
            used.update(w.get('tds_merged') or [])
            low = sum(1 for c in res['calculations'] if c.get('confidence') == 'low')
            decisions = sum(t.count('DECISION NEEDED') for t in files.values())
            rows.append({'workbook': fn, 'twb_bytes': os.path.getsize(p),
                         'digest_bytes': sum(len(v.encode()) for v in files.values()),
                         'largest_file_bytes': max(len(v.encode()) for v in files.values()),
                         'worksheets': w['worksheets'], 'dashboards': w['dashboards'], 'datasources': w['datasources'],
                         'calcs_used': w['calcs_used'], 'calcs_total': w['calcs_total'], 'low_confidence_calcs': low,
                         'decisions_needed': decisions, 'tds_merged': ' '.join(w.get('tds_merged') or []),
                         'tier': res['tier']['suggested'], 'flags': ' '.join(sorted({f['code'] for f in res['flags']}))})
            print(f'  ok  {fn}: {w["dashboards"]} dashboards, {w["calcs_used"]} calcs, tier {res["tier"]["suggested"]}'
                  + (f', .tds: {", ".join(w["tds_merged"])}' if w.get('tds_merged') else '')
                  + (f', {decisions} decision(s) needed' if decisions else ''))
        except (Exception, SystemExit) as e:  # keep the batch going
            rows.append({'workbook': fn, 'flags': f'ERROR {e}'})
            print(f'  ERR {fn}: {e}')
    for t in tds:
        if os.path.basename(t) not in used:
            print(f'  note: {os.path.basename(t)} did not match any workbook\'s published data source '
                  '(match is by data source name; rename the file to the published source name)')
    os.makedirs(out, exist_ok=True)
    keys = ['workbook', 'twb_bytes', 'digest_bytes', 'largest_file_bytes', 'worksheets', 'dashboards', 'datasources',
            'calcs_used', 'calcs_total', 'low_confidence_calcs', 'decisions_needed', 'tds_merged', 'tier', 'flags']
    with open(os.path.join(out, 'inventory.csv'), 'w', newline='', encoding='utf-8') as f:
        wr = csv.DictWriter(f, keys)
        wr.writeheader()
        wr.writerows(rows)
    return rows


def main():
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('path', nargs='?', help=f'workbook or folder (default: {INPUT_DIR})')
    ap.add_argument('--refresh-catalog', action='store_true',
                    help='re-snapshot visual roles from the powerbi-report-author CLI and exit')
    ap.add_argument('--tds', action='append', default=[], help='published data source .tds/.tdsx (repeatable)')
    ap.add_argument('--out', help=f'output folder (default: {OUTPUT_DIR})')
    ap.add_argument('--show', help='print pruned XML of one named artifact and exit')
    ap.add_argument('--split', dest='split', action='store_true', default=None, help='always split per dashboard')
    ap.add_argument('--no-split', dest='split', action='store_false', help='never split')
    ap.add_argument('--table-map', help='JSON file {"SOURCE_TABLE": "Power BI table"} used in DAX drafts '
                                        '(default: input/table_map.json when present)')
    ap.add_argument('--extract-data', action='store_true',
                    help='write the rows of packaged .hyper extracts to output\\<workbook>\\extract_data\\*.csv '
                         '(needs the optional tableauhyperapi package)')
    a = ap.parse_args()
    if a.refresh_catalog:
        from bhayankar_report import refresh_catalog
        print(refresh_catalog())
        return
    path = a.path or INPUT_DIR
    if not os.path.exists(path):  # a workbook name, as every other Bhayankar command takes: input\<name>.twb(x)
        hit = next((os.path.join(INPUT_DIR, path + ext) for ext in ('.twbx', '.twb')
                    if os.path.exists(os.path.join(INPUT_DIR, path + ext))), None)
        if not hit:
            sys.exit(f'{path}: not found. Give a .twb/.twbx path, or the name of a workbook in {INPUT_DIR}')
        path = hit
    out = a.out or OUTPUT_DIR
    tmap_file = a.table_map or (os.path.join(INPUT_DIR, 'table_map.json')
                                if os.path.exists(os.path.join(INPUT_DIR, 'table_map.json')) else None)
    tmap = {}
    if tmap_file:
        with open(tmap_file, encoding='utf-8') as f:
            tmap = json.load(f)
    if a.show:
        show(load_root(path)[0], a.show)
        return
    if os.path.isdir(path):
        os.makedirs(path, exist_ok=True)
        print(f'{bt.banner()}: extractor scanning {path}' + (f' (table map: {os.path.basename(tmap_file)})' if tmap_file else ''))
        rows = run_folder(path, out, a.split, tmap, a.extract_data)
        if not rows:
            print(f'No .twb/.twbx found. Put workbooks (and their published .tds/.tdsx) in {path} and run again.')
            return
        print(f'{len(rows)} workbook(s) → {out} (summary: inventory.csv)')
        return
    if not os.path.exists(path):
        ap.error(f'not found: {path}')
    dest = os.path.join(out, os.path.splitext(os.path.basename(path))[0])
    res, files = process(path, a.tds, dest, a.split, tmap, extract_data=a.extract_data)
    total = sum(len(v.encode()) for v in files.values())
    print(f'Wrote {len(files)} digest file(s) + bhayankar_digest.json to {dest}: {total:,} bytes (~{total // 4:,} tokens) '
          f'from {os.path.getsize(path):,} bytes; {len(res["flags"])} flags; tier {res["tier"]["suggested"]}')


if __name__ == '__main__':
    main()
