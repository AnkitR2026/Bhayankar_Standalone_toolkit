"""Bhayankar extractor, report module: digest -> Power BI rebuild spec.

Turns the extracted Tableau logic into what the Power BI build needs:
  model  : tables/columns, measures (incl. explicit measures for Tableau's implicit aggregations),
           calculated columns, date table, relationships, parameters, field-parameter candidates
  pages  : visuals (type + roles checked against the powerbi-report-author catalog), positions,
           titles, sorts, labels, colours, reference lines, filters, slicers with defaults,
           sync groups, drillthrough, tooltip pages, interactions, static items
Rules follow the Fabric skills: semantic-model-authoring (naming, explicit measures, formatString,
measures on their fact tables, single-direction relationships, Power Query date table) and
powerbi-report-cli (modern visual types, CLI roles, slicer types). Everything uncertain is a note.
Python 3 standard library only.
"""
import json
import os
import re

from bhayankar_dax import Translator, format_string, suggest_measure_name, q, col, param_table_name

HERE = os.path.dirname(os.path.abspath(__file__))
AGG_FN = {'SUM': ('Total', 'SUM'), 'AVG': ('Average', 'AVERAGE'), 'MIN': ('Min', 'MIN'), 'MAX': ('Max', 'MAX'),
          'COUNT': ('#', 'COUNT'), 'COUNTD': ('#', 'DISTINCTCOUNTNOBLANK'), 'MEDIAN': ('Median', 'MEDIAN'),
          'ATTR': (None, 'SELECTEDVALUE'), 'STDEV': ('Std Dev', 'STDEV.S'), 'VAR': ('Variance', 'VAR.S')}
NUMERIC_ONLY_AGG = {'SUM', 'AVG', 'MEDIAN', 'STDEV', 'STDEVP', 'VAR', 'VARP'}  # DAX cannot do these over text/dates
DATE_COL = {'YEAR': 'Year', 'QUARTER': 'Quarter', 'MONTH': 'Month', 'WEEK': 'Week', 'DAY': 'Day', 'WEEKDAY': 'Day of Week',
            'TRUNC_YEAR': 'Year', 'TRUNC_QUARTER': 'Year Quarter', 'TRUNC_MONTH': 'Year Month', 'TRUNC_WEEK': 'Week Start',
            'TRUNC_DAY': 'Date', 'DATE': 'Date', 'MONTH_YEAR': 'Year Month', 'MONTH_DAY': 'Month Day', 'HOUR': 'Hour'}
QTC_NAME = {'PERCENT_OF_TOTAL': '% of Total', 'PERCENT_DIFFERENCE': '% Difference', 'DIFFERENCE': 'Difference',
            'RUNNING_TOTAL': 'Running Total', 'RANK': 'Rank', 'MOVING_AVG': 'Moving Average', 'WINDOW': 'Window',
            'PERCENTILE': 'Percentile', 'YTD_TOTAL': 'YTD', 'YTD_GROWTH': 'YTD Growth', 'COMPOUND_GROWTH': 'CAGR'}
SLICER_MODE = {  # Tableau quick-filter mode -> (visualType, data.mode, selection) per powerbi-report-cli slicers.md
    'checkdropdown': ('slicer', 'Dropdown', 'multi'), 'dropdown': ('slicer', 'Dropdown', 'single'),
    'checklist': ('listSlicer', None, 'multi'), 'radiolist': ('listSlicer', None, 'single'),
    'slider': ('slicer', 'Between', 'range'), 'typein': ('textSlicer', None, 'search'),
    'pattern': ('textSlicer', None, 'search'), 'wildcard': ('textSlicer', None, 'search'),
    'datetime': ('slicer', 'Between', 'range'), 'relative-date': ('slicer', 'Relative', 'relative'),
    'compact': ('slicer', 'Dropdown', 'single'), 'list': ('listSlicer', None, 'single'),
    'type_in': ('textSlicer', None, 'search'), 'radiolist-h': ('advancedSlicerVisual', None, 'single')}


def load_catalog():
    p = os.path.join(HERE, 'bhayankar_visual_catalog.json')
    try:
        with open(p, encoding='utf-8') as f:
            return json.load(f)
    except OSError:
        return None


def refresh_catalog():
    """Re-snapshot visual types + roles from the installed powerbi-report-author CLI (source of truth for PBIR)."""
    import subprocess

    def cli(*a):
        out = subprocess.run(['powerbi-report-author', *a], capture_output=True, text=True, shell=os.name == 'nt',
                             encoding='utf-8')
        if out.returncode:
            raise SystemExit(f'powerbi-report-author {" ".join(a)} failed: {out.stderr.strip()[:300]}')
        return json.loads(out.stdout)['data']
    lst = cli('catalog', 'list')
    ver = subprocess.run('powerbi-report-author --version', capture_output=True, text=True, shell=True).stdout.strip()
    cat = {'source': 'powerbi-report-author catalog (snapshot)', 'cli_version': ver,
           'deprecated': {d['visualType']: d for d in lst['deprecated']},
           'legacy_do_not_create': {'card': 'cardVisual', 'multiRowCard': 'cardVisual', 'table': 'tableEx',
                                    'matrix': 'pivotTable', 'map': 'azureMap', 'filledMap': 'azureMap'},
           'visuals': {}}
    for t in lst['visualTypes']:
        d = cli('catalog', 'describe', t)
        cat['visuals'][t] = {'roles': {k: v['kind'] for k, v in (d.get('roles') or {}).items()},
                             'required': d.get('requiredRoles'), 'max': d.get('maxPerRole')}
    with open(os.path.join(HERE, 'bhayankar_visual_catalog.json'), 'w', encoding='utf-8') as f:
        json.dump(cat, f, indent=1)
    return f'{len(cat["visuals"])} visual types from powerbi-report-author {ver}'


def dq(s):
    """DAX string literal."""
    return '"' + s.replace('"', '""') + '"'


# Tableau relative-date units -> PBIR TimeUnit (semanticQuery schema: 0 Day, 1 Week, 2 Month, 3 Year; no Quarter)
REL_UNITS = {'day': 0, 'week': 1, 'month': 2, 'year': 3}


def dax_literal(v, datatype=None):
    """Tableau filter member -> DAX literal: numbers stay numbers for a numeric field, #dates# become dt"..." values,
    true/false become TRUE()/FALSE(), anything else a string."""
    v = str(v)
    if datatype in ('integer', 'real') and re.fullmatch(r'-?\d+(\.\d+)?', v):
        return v
    if datatype == 'boolean' and v.lower() in ('true', 'false'):
        return v.upper() + '()'
    m = re.fullmatch(r'#(\d{4}-\d{2}-\d{2})(?:[ T](\d{2}:\d{2}:\d{2}))?#', v)
    if m:
        return f'dt"{m.group(1)}' + (f'T{m.group(2)}"' if m.group(2) else '"')
    return dq(v)


def pretty_table(src):
    """Source table name -> business-style suggestion (naming-conventions). Always a suggestion: use --table-map."""
    if not src:
        return None
    s = re.sub(r'^\[?(dbo|edw|stg)\]?\.', '', src.strip('[]'), flags=re.I)
    s = re.sub(r'^(DIM|FACT|STG|F|D|VW|V|TBL)_', '', s, flags=re.I)
    s = re.sub(r'_(EDW|DW|V|VW|VIEW|TBL|TABLE|STG)$', '', s, flags=re.I)
    s = re.sub(r'[#$].*$', '', s)
    words = re.split(r'[_\s]+', s)
    return ' '.join(w if (w.isupper() and len(w) <= 2) else w.capitalize() for w in words if w)


class SheetCalcs(dict):
    """name -> digest calc. A name that several data sources define (each with its own formula) resolves to the
    one of the data source the current sheet uses (Rebuild._ws); otherwise, and outside a sheet, the last one."""

    def __init__(self, calcs, owner):
        super().__init__((c['name'], c) for c in calcs)
        self.by_name, self.owner = {}, owner
        for c in calcs:
            self.by_name.setdefault(c['name'], []).append(c)

    def pick(self, name):
        lst = self.by_name.get(name)
        if lst and len(lst) > 1:
            for ds in (getattr(self.owner, '_ws', None) or {}).get('datasources') or []:
                hit = next((c for c in lst if c.get('datasource') == ds), None)
                if hit:
                    return hit
        return dict.get(self, name)

    def get(self, name, default=None):
        hit = self.pick(name)
        return default if hit is None else hit

    def __getitem__(self, name):
        hit = self.pick(name)
        if hit is None:
            raise KeyError(name)
        return hit


class Rebuild:
    _ds = None     # data source of the field being resolved (field_of)
    cols_ds = {}   # (data source, column name) -> column (set per instance in __init__)

    def __init__(self, res, catalog=None):
        self.res = res
        self.cat = catalog if catalog is not None else load_catalog()
        self.notes = []
        self.measures = {}        # name -> measure dict (explicit + translated)
        self.date_roles = []      # source date columns used with date parts
        self.url_columns = []     # Power Query columns built from Tableau URL actions
        self.param_dims = {}      # parameter-driven calcs used as dimensions -> disconnected value tables
        self.cols = {c['name']: c for c in res['columns']}
        # the same column name in two data sources (blending: Orders[Region] and Targets[Region]): the field's own
        # data source picks the table (self.cols keeps only the last one)
        self.cols_ds = {(c.get('datasource'), c['name']): c for c in res['columns']}
        self._ds = None
        self.calcs = SheetCalcs(res['calculations'], self)
        self.groups = {g['name']: g for g in res['groups']}
        self.params = {p['name']: p for p in res['parameters']}
        self.formats = {}
        for ws in res['worksheets']:
            for k, v in (ws.get('number_formats') or {}).items():
                self.formats.setdefault(k.strip('[]'), v)
        self.split_same_named_calcs(res['calculations'])
        for c in res['calculations']:
            if '[' not in (c['formula'] or '') and not c['name'].strip(' "'):
                continue  # constant placeholder calc (e.g. a help icon's colour field): nothing to rebuild
            if c['pbi_kind'] in ('measure', 'visual calculation'):
                self.add_measure(c.get('model_name') or c['name'], c['dax'], c.get('table') or self.home_of(c), c['datatype'],
                                 f'Tableau calc "{c["name"]}": {c["formula"]}', c['confidence'], c['dax_notes'],
                                 folder='Calculations', tableau=c['name'])

    def split_same_named_calcs(self, calcs):
        """The same calc name in two data sources with different formulas (Retail "$ Contr." on Store/Category and on
        Chain/Class): one measure per data source. The first keeps the name, the others get "<name> (<data source>)",
        and the measures of that data source that call it are pointed at their own copy. Before, the first formula
        served every sheet, so the other data source's sheets showed numbers of the wrong table."""
        by = {}
        for c in calcs:
            by.setdefault(c['name'], []).append(c)
        renamed = {}  # (data source, name) -> model name
        while True:  # until no change: a calc whose text is the same everywhere ([Share] * 2) differs once [Share] split
            new = {}
            for name, lst in by.items():
                if len(lst) < 2 or len({c.get('dax') or '' for c in lst}) < 2:
                    continue
                for c in lst[1:]:
                    if c['pbi_kind'] in ('measure', 'visual calculation') and c.get('datasource') and not c.get('model_name'):
                        c['model_name'] = f'{name} ({c["datasource"]})'
                        new[(c['datasource'], name)] = c['model_name']
            if not new:
                break
            renamed.update(new)
            for c in calcs:
                for (ds, name), mn in new.items():
                    if c.get('datasource') == ds and c.get('dax'):
                        c['dax'] = re.sub(r"(?<!['\]\w])\[" + re.escape(name) + r"\]", lambda _m: f'[{mn}]', c['dax'])
        if not renamed:
            return
        self.notes.append('same calc name in several data sources: ' + ', '.join(sorted(renamed.values())))

    # ---------- model objects ----------
    def home_of(self, c):
        m = re.search(r"'((?:[^']|'')+)'\[", c.get('dax') or '')
        return m.group(1).replace("''", "'") if m else None

    def add_measure(self, name, dax, table, datatype, desc, conf='high', notes=(), folder='Base', tableau=None, fmt_key=None):
        if name in self.measures:
            return self.measures[name]
        # fmt_key: the field as the sheet names it (PERCENT_OF_TOTAL(Task Count) for "Task Count (% of Total)")
        fmt = format_string((self.formats.get(fmt_key) if fmt_key else None) or self.formats.get(name) or
                            self.formats.get(f'SUM({name})') or self.formats.get(f'[{name}]'), datatype, name)
        new, abbrev = suggest_measure_name(name, dax)
        m = {'name': name, 'suggested_name': new if new != name else None, 'abbreviations': abbrev or None,
             'table': table, 'dax': dax, 'formatString': fmt, 'displayFolder': f'Measures\\{folder}',
             'description': desc, 'confidence': conf, 'notes': list(notes), 'tableau': tableau}
        if not table:
            m['notes'].append('home table unknown: place it on its fact table (modeling-guidelines: no single measures table)')
        self.measures[name] = m
        return m

    def primary_ds(self):
        """Data source the current sheet is built on: the first one its view uses (Tableau's blend primary)."""
        ws = getattr(self, '_ws', None) or {}
        return (ws.get('blend') or {}).get('primary') or next((d for d in ws.get('datasources') or [] if d != 'Parameters'), None)

    def col_for(self, name):
        """Source column by label: the field's own data source, else the sheet's primary, else any source."""
        for ds in (self._ds, self.primary_ds()):
            if ds and (ds, name) in self.cols_ds:
                return self.cols_ds[(ds, name)]
        return self.cols.get(name)

    def field_of(self, label, ds, measure_hint=None, **kw):
        """field() for a shelf / filter field of a known data source."""
        prev, self._ds = self._ds, ds
        try:
            return self.field(label, measure_hint, **kw)
        finally:
            self._ds = prev

    def column_ref(self, name):
        if name in self.cols:
            c = self.col_for(name)
            return {'kind': 'Column', 'table': c['table'], 'name': name}
        if name in self.groups:
            return {'kind': 'Column', 'table': self.groups[name].get('table'), 'name': name}
        if name in self.calcs:
            c = self.calcs[name]
            if c['pbi_kind'] == 'field parameter':
                return {'kind': 'Column', 'table': name, 'name': name, 'field_parameter': True}
            if c['pbi_kind'] == 'calculated column':
                return self.param_dimension(name) or {'kind': 'Column', 'table': c.get('table'), 'name': name}
            mn = c.get('model_name') or name
            return {'kind': 'Measure', 'table': self.measures.get(mn, {}).get('table'), 'name': mn}
        return None

    # ---------- table calculations written over the sheet's addressing ----------
    def resolver(self):
        """Field token of a digest formula ([Caption], [Param:Name]) -> what bhayankar_dax.Translator needs, from the
        digest (the extractor's own resolver works on the workbook XML, which the report does not read)."""
        def res(tok):
            m = re.match(r'^\[(?:[^\]]*\]\.\[)?((?:[^\]]|\]\])+)\]$', tok)
            n = (m.group(1) if m else tok.strip('[]')).replace(']]', ']')
            if n.startswith('Param:'):
                return {'kind': 'param', 'name': n[6:]}
            c = self.calcs.get(n)
            if c:
                kind = 'rowcalc' if c['pbi_kind'] in ('calculated column', 'field parameter') else 'measure'
                return {'kind': kind, 'name': n, 'table': c.get('table') or self.home_of(c), 'datatype': c.get('datatype')}
            c = self.cols.get(n) or self.groups.get(n)
            if c:
                return {'kind': 'column', 'name': n, 'table': c.get('table'), 'datatype': c.get('datatype') or 'string'}
            return {'kind': 'unknown', 'name': tok}
        return res

    def tc_addressing(self, tc, tc_fields):
        """[(table, column)] a table calc of the current sheet runs over, or None when it cannot be told. Tableau's
        default (no "compute using" saved) is Table (across): the Columns shelf."""
        if tc == 'Field' and tc_fields:
            out = []
            for lab in tc_fields:
                fr = self.field(lab, False)
                if fr.get('kind') != 'Column' or not fr.get('table'):
                    return None
                if (fr['table'], fr['name']) not in out:
                    out.append((fr['table'], fr['name']))
            return out
        if tc in (None, 'Rows', 'Columns'):
            return self.shelf_addressing(tc or 'Rows')
        if tc == 'Table':
            a, b = self.shelf_addressing('Rows'), self.shelf_addressing('Columns')
            return None if a is None or b is None else a + [x for x in b if x not in a]
        if tc in ('Cell', 'CellInPane'):
            return []
        return None  # pane-relative orderings: not written

    def table_calc(self, c, tc, tc_fields):
        """A Tableau table-calculation calc used on the current sheet -> a measure over that sheet's addressing columns.
        The first use gives the calc's own measure its DAX; a sheet that addresses other columns gets a variant
        "<calc> (along <columns>)". None when the addressing or a function in it cannot be written."""
        addr = self.tc_addressing(tc, tc_fields)
        if addr is None or any(t == '<Table>' for t, _ in addr):
            return None
        tr = Translator(self.resolver()).translate(c['formula'], c.get('role'), None, c['name'], window=[col(t, n) for t, n in addr])
        dax = tr.get('dax')
        if tr['kind'] != 'measure' or not dax or '<Table>' in dax or '<Dimension>' in dax or '/* visual calc' in dax:
            return None
        along = ', '.join(n for _, n in addr) or 'each mark'
        desc = f'Tableau calc "{c["name"]}": {c["formula"]} | table calculation over {along} (Tableau compute using)'
        base = c.get('model_name') or c['name']
        nm, cur = base, self.measures.get(base)
        if cur and cur['dax'] != dax and cur.get('tc_along') is not None:
            nm = f'{base} (along {along})'
            cur = self.measures.get(nm)
        if cur and cur.get('tc_along') is None:  # the calc's placeholder measure: it gets the DAX
            cur.update({'dax': dax, 'confidence': tr['confidence'], 'notes': tr['notes'], 'description': desc,
                        'displayFolder': 'Measures\\Table Calcs', 'tc_along': along})
            return cur
        if cur:
            return cur
        m = self.add_measure(nm, dax, (self.measures.get(base) or {}).get('table') or c.get('table') or self.home_of({'dax': dax}),
                             c.get('datatype'), desc, tr['confidence'], tr['notes'], folder='Table Calcs', tableau=c['name'])
        m['tc_along'] = along
        return m

    def shelf_addressing(self, ordering):
        """Dimensions a Table (across) / Table (down) calc runs along: the Columns shelf for Rows ordering (across),
        the Rows shelf for Columns ordering (down). None when the ordering is anything else."""
        ws = getattr(self, '_ws', None)
        if not ws or ordering not in ('Rows', 'Columns'):
            return None
        out = []
        for f in ws['shelf']['cols' if ordering == 'Rows' else 'rows']:
            if f['measure'] or f.get('special') or self.is_constant(f['label']):
                continue
            fr = self.field(f['label'], False)
            if fr['kind'] == 'Column' and fr.get('table') and (fr['table'], fr['name']) not in out:
                out.append((fr['table'], fr['name']))
        return out

    def pct_addressing(self, ordering):
        """Columns a Tableau PERCENT_OF_TOTAL adds up over on the current sheet, or None when the ordering is not
        Table (across) / Table (down). Rows = Table (across): the Rows shelf partitions, every other dimension
        (Columns shelf, marks card) is addressed. Columns = Table (down): the Columns shelf partitions."""
        ws = getattr(self, '_ws', None)
        if not ws or ordering not in ('Rows', 'Columns'):
            return None
        part = 'rows' if ordering == 'Rows' else 'cols'
        other = ws['shelf']['cols' if part == 'rows' else 'rows'] + [e for p in ws['panes'] for e in p.get('enc', [])]
        out = []
        for f in other:
            if f['measure'] or f.get('special') or self.is_constant(f['label']):
                continue
            fr = self.field(f['label'], False)
            if fr['kind'] == 'Column' and fr.get('table') and (fr['table'], fr['name']) not in out:
                out.append((fr['table'], fr['name']))
        return out

    def field(self, label, measure_hint=None, tc=None, tc_fields=None, shelf=False):
        """Digest label -> {'kind': Column|Measure, 'table', 'name', 'label'} creating explicit measures as needed.
        shelf: the field sits on the current sheet's shelves / marks card (tc, tc_fields: its Tableau "compute using"),
        so a table calculation can be written over the sheet's addressing columns."""
        lab = label.strip('[]')
        out = {'label': lab}
        if shelf:
            ma = re.match(r'^(?:AGG|SUM|AVG|MIN|MAX)\((.+)\)$', lab)
            c = self.calcs.get(ma.group(1) if ma and ma.group(1) in self.calcs else lab)
            if c and c.get('pbi_kind') == 'visual calculation':
                tm = self.table_calc(c, tc, tc_fields)
                if tm:
                    return {**out, 'kind': 'Measure', 'table': tm['table'], 'name': tm['name']}
        if lab.startswith('Param:'):
            p = lab[6:]
            pat = (self.params.get(p) or {}).get('powerbi') or {}
            if measure_hint:
                return {**out, 'kind': 'Measure', 'table': None, 'name': f'{p} Value'}
            return {**out, 'kind': 'Column', 'table': param_table_name(p),
                    'name': 'Date' if 'CALENDAR' in (pat.get('table_dax') or '') else 'Value'}
        m = re.match(r'^COUNT\(rows of (.+)\)$', lab)
        if m:
            t = m.group(1)
            mm = self.add_measure(f'# {pretty_table(t)} Rows', f'COUNTROWS({q(t)})', t, 'integer',
                                  f'Tableau row count of {t}', folder='Counts')
            return {**out, 'kind': 'Measure', 'table': t, 'name': mm['name']}
        m = re.match(r'^([A-Z_]+)\((.+)\)$', lab)
        if m and m.group(1) in QTC_NAME:
            inner = self.field(m.group(2), True)
            nm = f'{inner["name"]} ({QTC_NAME[m.group(1)]})'
            addr = self.pct_addressing(tc) if m.group(1) == 'PERCENT_OF_TOTAL' and tc else None
            if addr is not None:  # Tableau "compute using" known: remove only the addressed dimensions
                total = f'ALLSELECTED({", ".join(col(t, c) for t, c in addr)})' if addr else None
                dax = f'DIVIDE([{inner["name"]}], CALCULATE([{inner["name"]}], {total}))' if total else \
                    f'DIVIDE([{inner["name"]}], [{inner["name"]}])'
                if nm in self.measures and self.measures[nm]['dax'] != dax:
                    nm = f'{nm[:-1]}, {" and ".join(c for _, c in addr) or "each mark"})'
                mm = self.add_measure(nm, dax, inner.get('table'), 'real',
                                      f'Tableau quick table calc PERCENT_OF_TOTAL of {inner["name"]}, compute using '
                                      f'Table ({"across" if tc == "Rows" else "down"}): total over '
                                      f'{", ".join(c for _, c in addr) or "the mark itself"}', 'high', [], folder='Table Calcs',
                                      fmt_key=lab)
                return {**out, 'kind': 'Measure', 'table': mm['table'], 'name': nm}
            if m.group(1) in ('RUNNING_TOTAL', 'RUNNING_SUM'):
                # Table (across) runs along the Columns shelf, Table (down) along Rows; marks-card dimensions
                # (a colour series) partition it, which the series filter context does by itself in DAX
                addr = self.shelf_addressing(tc) if tc else None
                if addr and len(addr) == 1:
                    # running sum along the one addressed column (Tableau "compute using" Table across / down)
                    t_, c_ = addr[0]
                    dax = (f'VAR _cur = MAX({col(t_, c_)})\nRETURN CALCULATE([{inner["name"]}], '
                           f'FILTER(ALLSELECTED({col(t_, c_)}), {col(t_, c_)} <= _cur))')
                    if nm in self.measures and self.measures[nm]['dax'] != dax:  # same calc along another column elsewhere
                        nm = f'{nm[:-1]}, along {c_})'
                    mm = self.add_measure(nm, dax, inner.get('table'), 'real',
                                          f'Tableau quick table calc RUNNING_TOTAL of {inner["name"]} along {c_}', 'medium',
                                          [f'running total along {t_}[{c_}] (Tableau compute using); check it restarts where '
                                           'Tableau does (partition by the other dimensions)'], folder='Table Calcs', fmt_key=lab)
                    return {**out, 'kind': 'Measure', 'table': mm['table'], 'name': nm}
            if m.group(1) == 'PERCENT_OF_TOTAL':  # share of the visual's total: the usual Tableau "Table (across)" case
                mm = self.add_measure(nm, f'DIVIDE([{inner["name"]}], CALCULATE([{inner["name"]}], ALLSELECTED()))',
                                      inner.get('table'), 'real', f'Tableau quick table calc PERCENT_OF_TOTAL of {inner["name"]}',
                                      'medium', ['percent of the visual total (ALLSELECTED); check Tableau "compute using" (--show)'],
                                      folder='Table Calcs', fmt_key=lab)
                return {**out, 'kind': 'Measure', 'table': mm['table'], 'name': nm}
            mm = self.add_measure(nm, f'/* visual calculation */ {m.group(1)}([{inner["name"]}])', inner.get('table'),
                                  'real', f'Tableau quick table calc {m.group(1)} of {inner["name"]}', 'low',
                                  ['quick table calc: rebuild as a visual calculation; check Tableau "compute using" (--show)'],
                                  folder='Table Calcs', fmt_key=lab)
            return {**out, 'kind': 'Measure', 'table': mm['table'], 'name': nm, 'visual_calc': True}
        if m and m.group(1) in DATE_COL:
            src = m.group(2)
            if src not in self.date_roles:
                self.date_roles.append(src)
            return {**out, 'kind': 'Column', 'table': 'Date', 'name': DATE_COL[m.group(1)], 'date_of': src}
        if m and m.group(1) in AGG_FN:
            fn, base = m.group(1), m.group(2)
            ref = self.column_ref(base)
            if ref and ref['kind'] == 'Measure':
                return {**out, **ref}  # SUM(<calc measure>): the translated measure already aggregates
            if ref:
                prefix, dfn = AGG_FN[fn]
                nm = f'{prefix} {base}' if prefix and prefix != '#' else (f'# {base}' if prefix == '#' else base)
                if fn == 'ATTR':
                    return {**out, **ref}
                dtype = (self.cols.get(base) or self.calcs.get(base) or {}).get('datatype')
                if fn in NUMERIC_ONLY_AGG and dtype in ('string', 'boolean', 'date', 'datetime'):
                    # SUM/AVG of a text, true/false or date field cannot be a DAX measure (the engine rejects
                    # SUM over text); a dimension in a measure slot is Tableau's ATTR: use the field itself
                    self.notes.append(f'{fn}([{base}]) on a {dtype} field: used as the field (Tableau ATTR), not as '
                                      f'a {fn} measure')
                    return {**out, **ref}
                dax = f'{dfn}({col(ref["table"] or "<Table>", base)})'
                if nm in self.measures and self.measures[nm]['dax'] != dax and ref['table']:
                    nm = f'{nm} ({ref["table"]})'  # SUM([Sales]) of a second data source (blending): its own measure
                mm = self.add_measure(nm, dax, ref['table'],
                                      'integer' if prefix == '#' else (self.cols.get(base) or {}).get('datatype'),
                                      f'Explicit measure for Tableau {fn}([{base}])',
                                      folder='Counts' if prefix == '#' else 'Base')
                if base in self.cols:
                    self.col_for(base)['hide'] = True  # modeling-guidelines: hide the aggregated base column
                return {**out, 'kind': 'Measure', 'table': mm['table'], 'name': nm}
        ref = self.column_ref(lab)
        if ref:
            if ref['kind'] == 'Column' and measure_hint:
                # continuous measure field without visible aggregation: Tableau default SUM
                return self.field(f'SUM({lab})', True)
            return {**out, **ref}
        self.notes.append(f'field "{lab}" not resolved to a model object')
        return {**out, 'kind': 'Unknown', 'table': None, 'name': lab}

    # ---------- data blending ----------
    def blend_attrs(self, fields, blend):
        """Fields of a secondary source that are dimensions (not its link fields) -> ATTR headers. Tableau aggregates
        every secondary field to the primary's marks: a dimension shows its one value, or "*" when there are several.
        In Power BI a column of another fact table cannot group the primary's measure (no relationship between them:
        every manager would show every region's total), so it becomes an ATTR measure."""
        links = {(s['datasource'], lk['secondary_field']) for s in blend['secondaries'] for lk in s['links']}
        out = []
        for f in fields:
            ds = f.get('ds')
            m = re.match(r'^ATTR\((.+)\)$', f['label'].strip('[]'))
            base = m.group(1) if m else f['label'].strip('[]')
            if ds and ds != blend['primary'] and (m or not f['measure']) and not f.get('special') \
                    and not f.get('date') and (ds, base) not in links:
                ref = self.field_of(base, ds, False)
                if ref.get('kind') == 'Column' and ref.get('table'):
                    f = {**f, 'label': f'ATTR({base})', 'measure': True, 'continuous': False, 'blend_attr': base}
            out.append(f)
        return out

    def attr_measure(self, f):
        """Tableau ATTR() of a blended secondary dimension: its value when the mark has one, "*" when it has several."""
        base = f['blend_attr']
        ref = self.field_of(base, f.get('ds'), False)
        if ref.get('kind') != 'Column' or not ref.get('table'):
            return ref
        c = col(ref['table'], ref['name'])
        dtype = (self.cols_ds.get((f.get('ds'), base)) or self.col_for(base) or {}).get('datatype')
        dax = f'IF(DISTINCTCOUNTNOBLANK({c}) > 1, "*", SELECTEDVALUE({c}))' if dtype in (None, 'string') else \
            f'VAR _v = SELECTEDVALUE({c})\nRETURN IF(DISTINCTCOUNTNOBLANK({c}) > 1, "*", IF(NOT ISBLANK(_v), _v & ""))'
        nm = f'{base} (ATTR)'
        if nm in self.measures and self.measures[nm]['dax'] != dax:
            nm = f'{base} (ATTR, {ref["table"]})'
        mm = self.add_measure(nm, dax, ref['table'], 'string',
                              f'Tableau ATTR([{base}]) of the blended data source {f.get("ds")}: its value per mark, "*" when '
                              'there are several (Tableau shows the asterisk too)', folder='Blend')
        return {'label': base, 'kind': 'Measure', 'table': mm['table'], 'name': mm['name']}

    def blend_filter(self, ws):
        """Tableau blending is a left join: the primary source decides which marks exist; secondary values only fill
        them. A Power BI visual over a shared dimension shows every member any of its measures has, so members only the
        secondary has would appear. Visual filter: rows of the primary's link table exist ("# <table> Rows" not blank)."""
        b = ws.get('blend') or {}
        tables = [lk['primary_table'] for s in b.get('secondaries', []) for lk in s['links'] if lk.get('primary_table')
                  and not lk.get('calc')]
        if not tables:
            return None
        fr = self.field(f'COUNT(rows of {tables[0]})', True)
        return {'field': fr, 'type': 'NotBlank', 'blend_primary': True,
                'tableau': f'blend: only the marks of the primary data source {b["primary"]} (Tableau left join)'}

    # ---------- visuals ----------
    def is_constant(self, label):
        """Constant-valued fields Tableau authors use as layout scaffolding (MIN(0.4), "SALES", Dummy)."""
        lab = label.strip('[]')
        m = re.match(r'^[A-Z_]+\((.+)\)$', lab, re.I)
        inner = m.group(1) if m else lab
        if re.match(r'^-?[\d.]+$', inner) or re.match(r'^"[^"]*"$', inner):
            return True
        c = self.calcs.get(inner)
        return bool(c and '[' not in (c['formula'] or '[') )

    def infer_visual(self, ws):
        self._ws = ws  # sheet context for table calcs ("compute using" depends on the shelves)
        panes = ws['panes']
        marks = [p['mark'] for p in panes if p.get('mark')]
        mark = next((mk for mk in marks if mk != 'Automatic'), marks[0] if marks else 'Automatic')
        keep = lambda fl: [f for f in fl if not self.is_constant(f['label'])]
        rows, cols = keep(ws['shelf']['rows']), keep(ws['shelf']['cols'])
        enc = keep([e for p in panes for e in p.get('enc', [])])
        if ws.get('blend'):  # dimensions of a secondary source are aggregated (ATTR): values, not groups
            rows, cols, enc = (self.blend_attrs(fl, ws['blend']) for fl in (rows, cols, enc))
        chan = lambda c: [e for e in enc if e['channel'] == c]
        mv = ws.get('measure_values') or []
        has_mv = any(f.get('special') == 'Measure Values' for f in rows + cols + enc)
        mvs = lambda: [{'label': m, 'measure': True, 'continuous': True} for m in mv]
        axis_m = lambda fl: [f for f in fl if f['measure'] and f.get('continuous') and not f.get('special')] + \
            (mvs() if any(f.get('special') == 'Measure Values' for f in fl) else [])
        mrows, mcols = axis_m(rows), axis_m(cols)
        dims = lambda fl: [f for f in fl if not f['measure'] and not f.get('special')]
        drows, dcols = dims(rows), dims(cols)
        headers = [f for f in rows + cols if f['measure'] and not f.get('continuous') and not f.get('special')]
        notes, conf, roles, extra = [], 'high', {}, {}
        if headers:
            notes.append('discrete aggregated field(s) on a shelf used as header labels: '
                         + ', '.join(h['label'] for h in headers) + ' (show as a column/label, not an axis)')
        text_enc = chan('text') + chan('label')
        color = chan('color')
        detail = chan('lod') + chan('detail')
        # a secondary source's dimension on Rows/Columns (ATTR header): a chart shows it in the tooltip
        tips = chan('tooltip') + [h for h in headers if h.get('blend_attr')]

        def add(role, flist, hint=None):
            rk = self.roles_of(vt).get(role, 'GroupingOrMeasure') if vt else 'GroupingOrMeasure'
            for f in flist:
                lab = f['label'].strip('[]')
                if lab == 'Measure Values':
                    add(role, mvs(), True)
                    continue
                if lab in ('Measure Names',) or lab.startswith('(deleted field') or self.is_constant(lab) or \
                        re.search(r'\b(?:Latitude|Longitude) \(generated\)', lab):
                    # Tableau's geocoded axes have no source column: a Power BI map geocodes the location field itself
                    continue
                if f.get('blend_attr'):  # secondary-source dimension: Tableau ATTR, one value per mark or "*"
                    fr = self.attr_measure(f)
                else:
                    fr = self.field_of(lab, f.get('ds'), hint if hint is not None else f['measure'], tc=f.get('tc_ordering'),
                                       tc_fields=f.get('tc_fields'), shelf=True)
                if rk == 'Measure' and fr['kind'] == 'Column':
                    # dimension in a measure-only role (tooltip/card): aggregation projection, as Tableau ATTR()
                    fr = {**fr, 'kind': 'Aggregation', 'function': 'Max' if fr.get('date_of') or
                          (self.cols.get(fr['name']) or {}).get('datatype') in ('date', 'datetime') else 'Min'}
                if rk == 'Grouping' and fr['kind'] == 'Measure':
                    if role == 'Series':
                        extra['color_measure'] = fr['name']
                        notes.append(f'colour from aggregate calc [{fr["name"]}]: field-driven dataPoint.fill (fx) '
                                     'bound to that measure, not a legend')
                    elif vt in ('pivotTable', 'tableEx') and role == 'Rows':
                        # a table shows an aggregate "dimension" per row as a value column: same text, same rows
                        if fr not in roles.setdefault('Values', []):
                            roles['Values'].append(fr)
                        notes.append(f'[{fr["name"]}] (aggregate calc on Tableau Rows) shown as a value column')
                    else:
                        extra.setdefault('unplaced', []).append(fr['name'])
                        notes.append(f'[{fr["name"]}] is an aggregate or parameter-driven dimension: Power BI cannot '
                                     'group by a measure; use a field parameter or a calculated column')
                    continue
                if fr not in roles.setdefault(role, []):
                    roles[role].append(fr)

        def done(vtype, c=None, n=None):
            out = {'visualType': vtype, 'roles': {k: v for k, v in roles.items() if v}, 'confidence': c or conf,
                   'notes': notes + (n or []), **extra}
            tt = [t for t in tips if not (t.get('blend_attr') and vtype in ('pivotTable', 'tableEx'))]  # tables: a value column
            if tt and vtype and 'Tooltips' in self.roles_of(vtype) and vtype not in ('help-icon', 'custom'):
                before = len(roles.get('Tooltips', []))
                add('Tooltips', tt, True)
                out['roles'] = {k: v for k, v in roles.items() if v}
            return out

        vt = None
        if mark == 'VizExtension':
            return done('custom', 'low', ['Tableau viz extension: needs an approved Power BI custom visual'])
        if not rows and not cols and not mv:
            if text_enc:
                vt = 'cardVisual'
                add('Data', text_enc, True)
                return done(vt)
            if any(p.get('tooltip') for p in panes):  # detail fields here only feed the tooltip text
                return {'visualType': 'help-icon', 'pattern': 'textbox/shape "?" + report-page tooltip (Tooltip page type)',
                        'roles': {}, 'confidence': 'medium', 'notes': ['Tableau help/info sheet: no data visual'],
                        'tooltip_text': next(p['tooltip'] for p in panes if p.get('tooltip'))}
            if color and mark in ('Circle', 'Shape', 'Square'):
                return {'visualType': 'shape', 'roles': {}, 'confidence': 'medium', 'color_by': color[0]['label'],
                        'notes': ['status indicator: shape with conditional fill (or a cardVisual with an icon)']}
            return {'visualType': 'shape' if not enc else None, 'roles': {}, 'confidence': 'medium' if not enc else 'low',
                    'notes': ['spacer/decoration sheet: shape or nothing' if not enc else 'sheet with no shelves: decide by hand']}
        if (mark in ('Text', 'Automatic') and not mrows and not mcols) and (text_enc or has_mv or (drows or dcols)):
            # discrete aggregates on Rows/Columns (a FIXED day count shown as a header) are values of a table row
            if drows and dcols and not any(f.get('special') == 'Measure Names' for f in cols):
                vt = 'pivotTable'
                add('Rows', drows)
                add('Columns', dcols)
                add('Values', headers + [e for e in text_enc if e['measure']] + (mvs() if has_mv else []), True)
                return done(vt)
            measures_txt = headers + [e for e in text_enc if e['measure']] + (mvs() if has_mv else [])
            if not measures_txt:
                vt = 'tableEx'  # dimensions only: a plain table
                add('Values', drows + dcols + [e for e in text_enc if not e['measure']], False)
                return done(vt)
            vt = 'pivotTable'  # powerbi-report-cli pitfall: tableEx with dimensions + measures can render no rows
            # Tableau Columns-shelf dimensions run across (Met | Not Met in one row): matrix Columns, not Rows
            add('Rows', drows)
            add('Columns', dcols)
            add('Values', measures_txt, True)
            return done(vt, n=['Tableau text table: pivotTable with stepped layout off and subtotals off reads like a flat table'])
        if mark in ('Circle', 'Shape', 'Square', 'Automatic') and mrows and mcols:
            vt = 'scatterChart'
            add('X', mcols[:1], True)
            add('Y', mrows[:1], True)
            cat = [e for e in detail if not e['measure']] + drows + dcols
            add('Category', cat[:1])
            if len(cat) > 1:
                notes.append('several detail fields: Power BI scatter takes one Category; combine them in a column or '
                             'move one to Legend')
            add('Series', [c for c in color if not c['measure']][:1] or [c for c in color])
            add('Size', chan('size'), True)
            return done(vt, 'medium' if len(cat) > 1 else None)
        if mark in ('Line', 'Polygon') and (mrows and mcols or chan('path')) and not (drows or dcols):
            return done('custom', 'low', ['chart drawn with Tableau path/trigonometry (radial, sankey, dumbbell...): '
                                          'needs a custom visual or a redesign'])
        if mark == 'Square' and drows and dcols and [c for c in color if c['measure']]:
            vt = 'pivotTable'
            add('Rows', drows)
            add('Columns', dcols)
            add('Values', [c for c in color if c['measure']], True)
            return done(vt, n=['heat map: background colour conditional formatting (FillRule gradient) on the value'])
        if mark == 'Pie':
            vt = 'pieChart'
            # "Multiple Values" is how some workbooks name Measure Names on the colour shelf
            pie_cat = [c for c in color if not c['measure'] and c.get('special') != 'Measure Names'
                       and c['label'].strip('[]') not in ('Measure Names', 'Multiple Values', ':Measure Names')]
            if not pie_cat:  # slices are Measure Names: Power BI pies need a Category column
                add('Y', [e for e in chan('wedge-size') + chan('size') if e['measure']] or mrows or mcols, True)
                out = done(vt, 'medium', ['pie of several measures (Measure Names): the build adds a slice table (one row per '
                                          'measure) and a SWITCH measure, so the pie gets a Category'])
                out['measure_slices'] = True
                return out
            add('Category', pie_cat)
            add('Y', [e for e in chan('wedge-size') + chan('size') if e['measure']] or mrows or mcols, True)
            return done(vt)
        if mark == 'Multipolygon' or any(re.search(r'\b(?:Latitude|Longitude) \(generated\)', f['label']) for f in
                                         ws['shelf']['rows'] + ws['shelf']['cols']):  # also AVG(Latitude (generated))
            vt = 'azureMap'
            add('Category', [e for e in detail if not e['measure']])
            if not roles.get('Category'):
                # the location is a parameter-chosen calc (IF [Param] = 'Region' THEN [Region] ... ELSE [State] END),
                # which cannot be a map column: the first geographic column it reads, so the map is still drawn
                geo = next((r for e in detail for r in re.findall(r'\[([^\]]+)\]', (self.calcs.get(e['label']) or {}).get('formula') or '')
                            if (self.cols.get(r) or {}).get('geo_role')), None)
                if geo:
                    add('Category', [{'label': geo, 'measure': False}])
                    notes.append(f'map location [{geo}]: the Tableau location field is chosen by a parameter; '
                                 'add a field parameter for the other levels')
            add('Size', chan('size'), True)
            add('Series', color)
            return done(vt, 'medium', ['map: set data categories (State, City, ...) on the location columns (map.md)'])
        if not (mrows or mcols) and mark in ('Shape', 'Circle', 'Square') and (drows or dcols):
            # a grid of icons/dots on discrete headers (detail table with status shapes, calendar of days): a matrix of
            # the values; the shapes become icon conditional formatting
            vals = [e for e in text_enc + color + chan('size') + chan('shape') if e['measure']] + (mvs() if has_mv else [])
            if not vals:
                vt = 'tableEx'
                add('Values', drows + dcols + [e for e in text_enc + color + chan('shape') if not e['measure']], False)
                return done(vt, 'medium', [f'{mark} marks on discrete headers only: a table of those fields; the shapes '
                                           'become icon conditional formatting'])
            vt = 'pivotTable'
            add('Rows', drows)
            add('Columns', dcols)
            add('Values', vals, True)
            return done(vt, 'medium', [f'{mark} marks on discrete headers (icon grid / calendar): a matrix of the values; '
                                       'add icon conditional formatting for the shapes'])
        if not (mrows or mcols):
            return done(None, 'low', [f'no mapping for mark {mark} (no continuous measure on an axis); decide by hand'])
        horizontal = bool(mcols) and not mrows
        meas = mcols if horizontal else mrows
        cats = drows if horizontal else dcols
        other = dcols if horizontal else drows  # dimensions on the measure side -> small multiples
        if not cats and detail:
            cats = [e for e in detail if not e['measure']][:1]
        if mark == 'GanttBar':
            if not (re.search('waterfall', ws['name'], re.I) or chan('size')):
                return done('custom', 'low', ['Gantt bar: needs an approved custom visual (visual-mapping.md)'])
            vt = 'waterfallChart'
            add('Category', cats)
            add('Y', meas[:1], True)
            return done(vt, 'medium', ['Tableau waterfall (Gantt bars + running total): Power BI waterfall computes the steps'])
        if not cats:
            vt = 'cardVisual'
            add('Data', meas, True)
            return done(vt, 'medium', ['single-value chart (e.g. bullet bar): card or a bar with a target line; '
                                       'Power BI bar charts need a category'])
        marks_set = {p.get('mark') for p in panes if p.get('mark') not in (None, 'Automatic')}
        distinct = list(dict.fromkeys(f['label'] for f in meas))
        dual = ws['shelf'].get('dual')
        if dual and len(distinct) == 1:
            notes.append('dual axis of the same measure (label/ref-line trick): one axis in Power BI')
        combo = len(marks_set) > 1 and {'Bar', 'Line'} <= marks_set and len(distinct) > 1
        line = mark == 'Line' or (mark == 'Automatic' and any(f['date'] and f['continuous'] for f in cats))
        inner = cats[-1]
        col_dims = [c for c in color if not c['measure'] and c.get('special') != 'Measure Names']
        same = [c for c in col_dims if c['label'] == inner['label']]
        series = [c for c in col_dims if c['label'] != inner['label']]
        if combo:
            vt = 'lineClusteredColumnComboChart'
        elif line:
            vt = 'lineChart'
        elif mark == 'Area':
            vt = 'stackedAreaChart' if series else 'areaChart'
        elif same and len(cats) > 1:
            vt = 'clusteredBarChart' if horizontal else 'clusteredColumnChart'
        elif series:
            vt = 'barChart' if horizontal else 'columnChart'  # Tableau stacks a colour dimension by default
        else:
            vt = 'clusteredBarChart' if horizontal else 'clusteredColumnChart'
        if same and len(cats) > 1:
            add('Category', cats[:-1])  # nested inner dimension that is also the colour -> legend
            add('Series', same)
        else:
            add('Category', cats)
            add('Series', series[:1] + [c for c in color if c['measure']][:1])
            if same:
                notes.append(f'bars coloured by their own category ({same[0]["label"]}): per-category dataPoint.fill')
        if combo:
            add('Y', [f for f in meas if f['label'] == distinct[0]], True)
            add('Y2', [f for f in meas if f['label'] != distinct[0]], True)
        else:
            add('Y', meas, True)
        if other:
            add('Rows', other)
            notes.append('dimensions on the measure shelf: Power BI small multiples (Rows role)')
        if len(distinct) > 1 and not combo and not dual and not has_mv:
            extra['split'] = {'measures': distinct, 'direction': 'horizontal' if horizontal else 'vertical'}
            extra['_tips'] = [t for t in tips if t['measure']]
            notes.append(f'{len(distinct)} measures on separate Tableau axes -> {len(distinct)} visuals side by side')
            return {'visualType': vt, 'roles': {k: v for k, v in roles.items() if v}, 'confidence': conf,
                    'notes': notes, **extra}
        return done(vt)

    def roles_of(self, vt):
        return ((self.cat or {}).get('visuals', {}).get(vt) or {}).get('roles', {}) if self.cat else \
            {'Tooltips': 'Measure'}

    def check(self, v):
        """Validate type + roles against the powerbi-report-author catalog snapshot."""
        vt = v['visualType']
        errs = []
        if not self.cat or not vt or vt in ('help-icon', 'custom') or v.get('measure_slices'):
            return errs  # measure-slice pies get their Category at build time
        if vt in self.cat.get('legacy_do_not_create', {}):
            errs.append(f'legacy type {vt}: use {self.cat["legacy_do_not_create"][vt]}')
        spec = self.cat['visuals'].get(vt)
        if not spec:
            return [f'unknown visual type {vt}']
        for role, fields in v['roles'].items():
            kind = spec['roles'].get(role)
            if not kind:
                errs.append(f'role {role} not in {vt}')
                continue
            for f in fields:
                if kind == 'Grouping' and f['kind'] in ('Measure', 'Aggregation'):
                    errs.append(f'{role} takes columns, got measure {f["name"]}')
                if kind == 'Measure' and f['kind'] == 'Column':  # Aggregation projections are fine
                    errs.append(f'{role} takes measures, got column {f["name"]}')
            mx = (spec.get('max') or {}).get(role)
            if mx and len(fields) > mx:
                errs.append(f'{role} allows {mx} field(s), got {len(fields)}')
        for r in spec.get('required') or []:
            if not v['roles'].get(r):
                errs.append(f'required role {r} empty')
        return errs

    def title(self, ws, sheet_name):
        t = ws.get('title')
        if not t:
            return {'text': sheet_name}
        t = t.replace('<Sheet Name>', sheet_name)
        toks = re.findall(r'<\[([^\]]+)\]>', t)
        if not toks:
            return {'text': t}
        parts, pos = [], 0
        for m in re.finditer(r'<\[([^\]]+)\]>', t):
            if m.start() > pos:
                parts.append('"' + t[pos:m.start()].replace('"', '""') + '"')
            lab = m.group(1)
            if lab.startswith('Param:'):
                p = self.params.get(lab[6:]) or {}
                fmt = '"M/d/yyyy"' if p.get('datatype') in ('date', 'datetime') else '"General"'
                parts.append(f'FORMAT([{lab[6:]} Value], {fmt})')
            elif lab.startswith('(deleted field') or '(deleted field' in lab:
                pass  # Tableau shows nothing for a field deleted from the data source
            else:
                fr = self.field(lab, None)
                if fr.get('kind') != 'Measure' and not fr.get('table'):
                    # titles keep the caption as typed ("Task COMP by"); the column may differ in case ("COMP By")
                    hit = next((k for k in self.cols if k.lower() == lab.lower()), None)
                    if hit:
                        fr = self.field(hit, None)
                if fr.get('kind') != 'Measure' and not fr.get('table'):
                    self.notes.append(f'title of "{sheet_name}": field [{lab}] not found in the model; left out of the title')
                    pos = m.end()
                    continue
                parts.append(f'[{fr["name"]}]' if fr['kind'] == 'Measure' else f'SELECTEDVALUE({col(fr["table"], fr["name"])})')
            pos = m.end()
        if pos < len(t):
            parts.append('"' + t[pos:].replace('"', '""') + '"')
        if not any(not p.startswith('"') for p in parts):  # nothing dynamic left: a plain text title
            return {'text': re.sub(r'<\[[^\]]+\]>', '', t).strip() or sheet_name}
        name = f'{sheet_name} Title'
        self.add_measure(name, ' & '.join(parts), None, 'string', f'Dynamic title of Tableau sheet "{sheet_name}"',
                         folder='Titles')
        return {'text': t, 'measure': name, 'note': 'bind title text to the measure (fx)'}

    def filter_spec(self, f):
        rule = f['rule']
        fr = self.field_of(f['field'], f.get('ds'), None)
        out = {'field': fr, 'tableau': rule}
        core = re.sub(r'\s+\[values:\w+\]$', '', rule)
        m = re.match(r'^(NOT )?IN\((.*)\)$', core)
        single = re.match(r'^(NOT )?"((?:[^"]|"")*)"$', core)
        rng = re.match(r'^RANGE\("((?:[^"]|"")*)"\.\."((?:[^"]|"")*)"\)$', core)
        lits = self.calc_literals(f['field'])
        # every member except the listed ones (ui-enumeration "exclusive" adds the NOT; it is UI only). Several
        # members come wrapped: ALL EXCEPT IN("a", "b")
        exc = re.match(r'^(?:NOT )?ALL EXCEPT (?:IN\((.+)\)|(.+))$', core)
        if exc:
            exc = re.match(r'^(.*)$', exc.group(1) or exc.group(2))
        qnull = re.match(r'^RANGE -inf\.\.\+inf(?: \((non-null|null|all)\))?$', core)
        # Tableau writes booleans and numbers without quotes: `true`, `NOT false`, `750`
        bare = re.match(r'^(NOT )?(true|false|-?\d+(?:\.\d+)?)$', core)
        # quantitative range: dates (#..#) or numbers, open ends as -inf / +inf
        qend = r'(#[^#]+#|-?\d+(?:\.\d+)?|[-+]inf)'
        qrng = re.match(rf'^RANGE {qend}\.\.{qend}(?: \((non-null|null|all)\))?$', core)
        # discrete range of whole numbers, e.g. MONTH(...) RANGE(1..12)
        drng = re.match(r'^RANGE\((-?\d+)\.\.(-?\d+)\)$', core)
        if exc and exc.group(1).strip() == '%null%' or (qnull and qnull.group(1) == 'non-null'):
            out.update({'type': 'NotBlank'})  # Tableau "exclude Null" / quantitative "non-null values"
        elif qnull and qnull.group(1) == 'null':
            out.update({'type': 'Blank'})
        elif qnull:
            out.update({'type': 'All'})
        elif exc and '%null%' not in exc.group(1):
            out.update({'type': 'Categorical', 'values': self.members(exc.group(1)), 'inverted': True})
        elif qrng:
            lo, hi = (None if x in ('-inf', '+inf') else x for x in qrng.groups()[:2])
            out.update({'type': 'Range', 'lo': lo, 'hi': hi, 'nonnull': qrng.group(3) == 'non-null'})
        elif drng and 0 <= int(drng.group(2)) - int(drng.group(1)) <= 400:
            out.update({'type': 'Categorical', 'inverted': False,
                        'values': [str(i) for i in range(int(drng.group(1)), int(drng.group(2)) + 1)]})
        elif bare:
            out.update({'type': 'Categorical', 'values': [bare.group(2)], 'inverted': bool(bare.group(1))})
        elif rule.startswith('RELATIVE'):
            attrs = dict(re.findall(r'([\w-]+)=(\S+)', rule))
            unit = attrs.get('period-type-v2') or attrs.get('period-type') or ''
            try:
                first, last = int(attrs.get('first-period', 'x')), int(attrs.get('last-period', 'x'))
            except ValueError:
                first = last = None
            if unit in REL_UNITS and first is not None and first <= last:
                out.update({'type': 'RelativeDate', 'unit': unit, 'first': first, 'last': last,
                            'include_future': attrs.get('include-future', 'true') == 'true',
                            'include_null': attrs.get('include-null') == 'true'})
            else:
                out.update({'type': 'RelativeDate', 'note': f'relative date rule not understood: {rule}'})
        elif rng and lits:
            # alphabetical range over a calc with a fixed set of outputs: list the outputs inside the range
            lo, hi = (x.replace('""', '"').lower() for x in rng.groups())
            out.update({'type': 'Categorical', 'values': [v for v in lits if lo <= v.lower() <= hi], 'inverted': False,
                        'note': f'alphabetical range {rng.group(1)}..{rng.group(2)} over the calc\'s outputs {lits}'})
        elif re.search(r'\b(UNION|INTERSECTION|EXCEPT|CROSSJOIN)\b|RANGE\(', core):
            # mixed rule (members + alphabetical ranges): never collapse it to the first member
            out.update({'type': 'Manual', 'members': re.findall(r'(?<![.(])"((?:[^"]|"")*)"(?!\.\.)', core),
                        'ranges': re.findall(r'RANGE\("([^"]*)"\.\."([^"]*)"\)', core),
                        'note': 'DECISION NEEDED: Tableau rule mixes members and alphabetical ranges; list the intended '
                                'values (or invert: exclude the unwanted ones) after checking the real value list'})
            return out
        elif m:
            out.update({'type': 'Categorical', 'values': self.members(m.group(2)), 'inverted': bool(m.group(1))})
        elif single:
            out.update({'type': 'Categorical', 'values': [single.group(2)], 'inverted': bool(single.group(1))})
        elif rule.startswith('RANGE ') or rule.startswith('RELATIVE'):
            out.update({'type': 'Advanced' if rule.startswith('RANGE') else 'RelativeDate'})
        elif rule.startswith(('TOP', 'BOTTOM')):
            out.update({'type': 'TopN', 'note': 'Top N filter: by the measure in the rule'})
        else:
            out.update({'type': 'Manual', 'note': 'translate this filter by hand (union/range/except)'})
        c = self.calcs.get(f['field'])
        if out.get('type') == 'Manual':
            return out  # a rule that was not understood is never folded into measures (it became `= "?"`): hand-translate
        if c and c['pbi_kind'] == 'calculated column' and 'Param:' in (c['formula'] or ''):
            # A calculated column is computed at refresh and cannot see slicers. A visual-level measure filter is no
            # substitute: it only hides groups whose totals still include the excluded rows. Tableau's dimension
            # filter removes rows, so the condition is applied inside the visual's measures.
            tbl = c.get('table') or '<Table>'
            val = out.get('values', ['?'])[0] if out.get('values') else '?'
            vals = out.get('values') or []
            lit = lambda v: dax_literal(v, c.get('datatype'))  # noqa: E731
            if out.get('type') in ('RelativeDate', 'TopN', 'Advanced'):
                out['note'] = (out.get('note') + '; ' if out.get('note') else '') + \
                    'filter on a parameter-driven row calc: fold it into the measures by hand'
                return out
            if out.get('type') == 'Range':
                parts = ([f'{c["dax"]} >= {lit(out["lo"])}'] if out['lo'] is not None else []) + \
                        ([f'{c["dax"]} <= {lit(out["hi"])}'] if out['hi'] is not None else [])
                val = f'{out["lo"] or "-inf"}..{out["hi"] or "+inf"}'
                cond = ' && '.join(parts) or f'NOT ISBLANK({c["dax"]})'
            elif vals and all(v in ('true', 'false') for v in vals) and len(vals) == 1:  # boolean calc = True / False
                # == is strict: BLANK() = FALSE() is TRUE in DAX, so "= False" would keep the NULL rows Tableau drops
                val, cond = vals[0], f'({c["dax"]}) == {vals[0].upper()}()'
                cond = f'NOT ({cond})' if out.get('inverted') else cond
            elif out.get('type') == 'NotBlank':
                val, cond = '(not blank)', f'NOT ISBLANK({c["dax"]})'
            elif len(vals) > 1 or out.get('inverted'):
                val = ', '.join(vals)
                cond = f'{c["dax"]} IN {{{", ".join(lit(v) for v in vals)}}}'
                cond = f'NOT ({cond})' if out.get('inverted') else cond
            else:
                cond = f'{c["dax"]} = {lit(val)}'
            out.update({'type': 'RowCondition', 'table': tbl, 'calc': c['name'], 'value': val,
                        'condition': cond,
                        'note': f'row-level calc reads a parameter: fold into this visual\'s measures as '
                                f'CALCULATE(<measure>, KEEPFILTERS(FILTER({q(tbl)}, <condition>)))'})
        return out

    @staticmethod
    def members(text):
        """'"a", "b"' or '225, 226' (numbers are unquoted in Tableau rules) -> ['a', 'b'] / ['225', '226']."""
        out = []
        for qd, bare in re.findall(r'"((?:[^"]|"")*)"|([^,\s][^,]*)', text):
            v = qd.replace('""', '"') if qd or not bare else bare.strip()
            if v:
                out.append(v)
        return out

    def calc_literals(self, name):
        """Every value a calc can return when each branch returns a string literal (IF/CASE ... THEN "x"), else None."""
        c = self.calcs.get(name)
        f = (c or {}).get('formula') or ''
        if not re.match(r'^\s*(IF|CASE)\b', f, re.I):
            return None
        heads = re.findall(r'\b(?:THEN|ELSE)\b', f, re.I)
        lits = re.findall(r'''\b(?:THEN|ELSE)\s+(?:"((?:[^"]|"")*)"|'((?:[^']|'')*)')''', f, re.I)
        if not heads or len(heads) != len(lits):
            return None
        return list(dict.fromkeys(a.replace('""', '"') if a else b.replace("''", "'") for a, b in lits))

    def param_dimension(self, name):
        """A calc that reads a parameter, used as a dimension: a calculated column cannot follow the slicer, so the
        build gives it a disconnected table of its possible values and filters the visual's measures with
        <calc DAX> IN VALUES(<that table>). None when its values are not a fixed list."""
        c = self.calcs.get(name)
        if not c or c['pbi_kind'] != 'calculated column' or 'Param:' not in (c['formula'] or '') or not c.get('dax'):
            return None
        lits = self.calc_literals(name)
        if not lits:
            return None
        if name not in self.param_dims:
            order = next((ws['manual_sorts'][name] for ws in self.res['worksheets'] if name in (ws.get('manual_sorts') or {})), [])
            vals = [v for v in order if v in lits] + [v for v in lits if v not in order]
            self.param_dims[name] = {'name': name, 'values': vals, 'host': c.get('table'), 'dax': c['dax']}
        return {'kind': 'Column', 'table': name, 'name': name, 'param_dim': True}

    def param_dim_condition(self, fr):
        pd = self.param_dims[fr['name']]
        return {'type': 'RowCondition', 'field': fr, 'table': pd['host'] or '<Table>', 'calc': fr['name'],
                'value': '(each row of the visual)',
                'condition': f"{pd['dax']} IN VALUES({q(fr['name'])}[{fr['name']}])",
                'tableau_calc_text': f'[{fr["name"]}] = the {fr["name"]} value of the row (parameter-driven dimension)'}

    def add_param_dim_filters(self, v, extra=()):
        """Visuals grouping by a parameter-driven dimension (or drilled through on one): fold the row condition into
        their measures. A visual with no measure cannot use the disconnected column: it is dropped, with a note."""
        fields = [f for fs in v['roles'].values() for f in fs if f.get('param_dim')] + list(extra)
        if not fields or not v.get('roles'):
            return
        has_measure = any(f['kind'] == 'Measure' for fs in v['roles'].values() for f in fs)
        if not has_measure:
            for r in list(v['roles']):
                v['roles'][r] = [f for f in v['roles'][r] if not f.get('param_dim')]
            v['notes'] = v.get('notes', []) + [f'[{f["name"]}] reads a parameter and this visual has no measure to filter: '
                                               'column left out (add a measure or a visual-level measure filter by hand)'
                                               for f in fields]
            return
        have = {x.get('calc') for x in v.get('visual_filters', []) if x.get('type') == 'RowCondition'
                and 'IN VALUES(' in x.get('condition', '')}
        for f in fields:
            if f['name'] not in have:
                v.setdefault('visual_filters', []).append(self.param_dim_condition(f))
                have.add(f['name'])

    # ---------- pages ----------
    def build(self):
        res = self.res
        by_ws = {ws['name']: ws for ws in res['worksheets']}
        pages = []
        visible = {t['name'] for t in res['tabs'] if not t['hidden']}
        tooltip_sheets = set()
        # A Tableau quick filter stays interactive as a (synced) slicer, so on its own dashboard it is neither a
        # page nor a visual filter: the saved selection would be locked in. Slicers reach only pages that hold
        # them (and drillthrough targets, as hidden synced copies, see sync()); on any other dashboard the sheet's
        # saved filter is fixed and must be applied, or the visual shows more rows than Tableau.
        db_slicers = {d['name']: {it['field'] for it in d['items'] if it['type'] == 'filter'} for d in res['dashboards']}
        drill_from = {}
        for a in res.get('actions', []):
            src_db = (a.get('source') or '').split(' / ')[0].strip()
            if a.get('kind') == 'FILTER' and a.get('target') and a['target'] != src_db and src_db in db_slicers:
                drill_from.setdefault(a['target'].strip(), set()).add(src_db)
        for d in res['dashboards']:
            W, H = d['px']
            page = {'name': d['name'], 'width': W, 'height': H, 'tableau_sizing': d.get('sizing'),
                    'visible': d['name'] in visible or not res['tabs'], 'visuals': [], 'slicers': [],
                    'static': [], 'page_filters': [], 'notes': []}
            if d.get('sizing') != 'fixed':
                page['notes'].append(f'Tableau sizing "{d.get("sizing")}": page set to {W}x{H}; confirm the canvas size')
            sheets = [it for it in d['items'] if it['type'] == 'sheet' and it['name'] in by_ws]
            sheet_names = [s['name'] for s in sheets]
            # filters every sheet on the page shares (and that are not slicers) -> page filters
            slicer_fields = db_slicers[d['name']].union(*(db_slicers[s] for s in drill_from.get(d['name'], ())))
            common = None
            for s in sheet_names:
                fl = [json.dumps(f, sort_keys=True) for f in by_ws[s]['filters'] if not f['rule'].startswith('ALL')]
                common = set(fl) if common is None else common & set(fl)
            common = common or set()
            measure_filters = []  # Power BI allows measure filters only at visual level
            for f in sorted(common):
                fd = json.loads(f)
                if fd['field'] not in slicer_fields:
                    fs = self.filter_spec(fd)
                    (measure_filters if fs.get('type') in ('Measure', 'RowCondition') or (fs.get('field') or {}).get('kind') == 'Measure'
                     else page['page_filters']).append(fs)
            for i, it in enumerate(d['items']):
                box = it['box']
                pos = {'x': box[0], 'y': box[1], 'width': box[2], 'height': box[3]} if box else None
                if it['type'] == 'sheet' and it['name'] in by_ws:
                    ws = by_ws[it['name']]
                    v = self.infer_visual(ws)
                    v.update({'id': f'v{i:02d}', 'sheet': ws['name'], 'position': pos,
                              'title': {**self.title(ws, ws['name']), 'show': it['show_title']}})
                    if any(p.get('labels_shown') for p in ws['panes']) and \
                            v['visualType'] not in ('pivotTable', 'tableEx', 'cardVisual', 'help-icon', None):
                        v['data_labels'] = True
                    if ws.get('sorts'):
                        v['sort'] = ws['sorts']
                        sb = self.sort_by(ws, v)
                        if sb:
                            v['sort_by'] = sb
                    if ws.get('reference_lines'):
                        v['reference_lines'] = []
                        for r in ws['reference_lines']:
                            val = self.field(r['value'], True)
                            same = r['value'] == r['axis']
                            v['reference_lines'].append({
                                'type': f'{r["formula"]} line' if same else 'measure-driven constant line',
                                'measure': val['name'], 'label': r['label'],
                                'note': None if same else f'value of {val["name"]} drawn on the {r["axis"]} axis: '
                                                          'constant line bound to a measure (fx)'})
                    cols_ = {}
                    for k, pal in (ws.get('colors') or {}).items():
                        cols_[k] = pal
                    colour_fields = [e['label'] for p_ in ws['panes'] for e in p_.get('enc', []) if e['channel'] == 'color']
                    for lab in colour_fields + [fr['label'] for r in ('Series', 'Category') for fr in v['roles'].get(r, [])]:
                        pal = (self.res.get('palettes') or {}).get(lab) or (self.cols.get(lab) or {}).get('color_map')
                        if pal:
                            cols_[lab] = pal
                    if cols_ and v['visualType'] not in ('help-icon', None):
                        v['colors'] = {'map': cols_, 'how': 'dataPoint.fill per series/category with Literal hex (color-strategy.md)'}
                    vf = [self.filter_spec(f) for f in ws['filters'] if not f['rule'].startswith('ALL')
                          and json.dumps(f, sort_keys=True) not in common and f['field'] not in slicer_fields]
                    vf += measure_filters
                    bf = self.blend_filter(ws) if v.get('visualType') not in ('help-icon', 'shape', None) else None
                    if bf:
                        vf.append(bf)
                    if vf:
                        v['visual_filters'] = vf
                    self.add_param_dim_filters(v)
                    if any(p.get('viz_in_tooltip') for p in ws['panes']):
                        v['tooltip'] = 'report-page tooltip (viz in tooltip)'
                    for vv in self.split_visual(v):
                        vv['catalog_errors'] = self.check(vv)
                        if vv['catalog_errors']:
                            vv['confidence'] = 'low'
                        page['visuals'].append(vv)
                elif it['type'] == 'filter':
                    vt, mode, sel = SLICER_MODE.get(it['mode'] or '', ('slicer', 'Dropdown', 'multi'))
                    fr = self.field(it['field'], False)
                    target = by_ws.get(it['name']) or {}
                    saved = next((f for f in target.get('filters', []) if f['field'] == it['field']), None)
                    default = None
                    if saved and not saved['rule'].startswith('ALL'):
                        fs = self.filter_spec(saved)
                        default = {'values': fs.get('values') or (fs.get('members', []) + [f'{a}..{b}' for a, b in fs.get('ranges', [])]),
                                   'inverted': fs.get('inverted'), 'tableau': saved['rule'], 'note': fs.get('note')}
                    page['slicers'].append({'id': f's{i:02d}', 'visualType': vt, 'mode': mode, 'selection': sel,
                                            'field': fr, 'position': pos, 'default': default,
                                            'applies_to_sheet': it['name'], 'sync_group': fr['name'],
                                            'tableau_field': it['field'],
                                            'scope_sheets': self.filter_scope(it['field'], sheet_names, by_ws)})
                elif it['type'] == 'paramctrl':
                    lab = it['field'] or ''
                    p = self.params.get(lab.replace('Param:', '')) or {}
                    page['slicers'].append({'id': f's{i:02d}', 'visualType': 'slicer',
                                            'mode': 'Between' if p.get('datatype') in ('date', 'datetime') else 'Dropdown',
                                            'selection': 'single', 'field': self.field(lab, False), 'position': pos,
                                            'default': {'values': [p.get('value')]}, 'parameter': p.get('name'),
                                            'powerbi': p.get('powerbi')})
                elif it['type'] in ('bitmap', 'text', 'title', 'web', 'button', 'navigation', 'dashboard-object'):
                    kind = {'bitmap': 'image', 'text': 'textbox', 'title': 'textbox', 'web': 'web content (flag)'}.get(
                        it['type'], 'actionButton')
                    text = it['text'] or (d['name'] if it['type'] == 'title' else None)
                    if text:  # a dashboard text object's <Sheet Name> is the dashboard's own name
                        text = text.replace('<Sheet Name>', d['name'].strip())
                    page['static'].append({'type': kind, 'position': pos, 'text': text,
                                           'image': it['field'] if it['type'] == 'bitmap' else None})
                elif it['type'] in ('color', 'size', 'shape'):
                    for v in page['visuals']:
                        if v['sheet'] == it['name']:
                            v['legend'] = {'show': True, 'tableau_zone': pos}
            self.date_param_pair(page)
            page['interactions'] = 'Tableau sheets do not filter each other unless an action says so: set the other ' \
                                   'visuals\' interactions to None except the pairs listed under actions'
            pages.append(page)
            for v in page['visuals']:
                if v.get('visualType') == 'help-icon':
                    tooltip_sheets.add(v['sheet'])
        # visible worksheet tabs not placed on any dashboard become their own page (a workbook saved without its tab
        # list counts every data sheet as visible, as the dashboards above do)
        placed = {v['sheet'] for p in pages for v in p['visuals']}
        for ws in res['worksheets']:
            shown = ws['name'] in visible or (not res['tabs'] and ws.get('role', 'data') == 'data')
            if shown and ws['name'] not in placed:
                v = self.infer_visual(ws)
                v.update({'id': 'v00', 'sheet': ws['name'], 'position': {'x': 0, 'y': 0, 'width': 1280, 'height': 720},
                          'title': self.title(ws, ws['name'])})
                # the sheet's own filters, as for a sheet on a dashboard (a worksheet tab used to lose them)
                vf = [self.filter_spec(f) for f in ws['filters'] if not f['rule'].startswith('ALL')]
                bf = self.blend_filter(ws) if v.get('visualType') not in ('help-icon', 'shape', None) else None
                vf += [bf] if bf else []
                if vf:
                    v['visual_filters'] = vf
                self.add_param_dim_filters(v)
                v['catalog_errors'] = self.check(v)
                pages.append({'name': ws['name'], 'width': 1280, 'height': 720, 'visible': True, 'visuals': [v],
                              'slicers': [], 'static': [], 'page_filters': [], 'notes': ['worksheet tab, not a dashboard']})
        # Tableau tab order (the first visible tab opens the workbook); the dashboards' XML order is not the tab order
        tab_pos = {t['name'].strip(): k for k, t in enumerate(res['tabs'])}
        pages.sort(key=lambda p: (not p.get('visible', True), tab_pos.get(p['name'].strip(), len(tab_pos))))
        self.actions(pages, by_ws)
        for p in pages:  # drillthrough on a parameter-driven dimension: the target page's measures follow it
            pdims = [f for f in (p.get('drillthrough') or {}).get('fields', []) if f.get('param_dim')]
            for v in p['visuals'] if pdims else []:
                self.add_param_dim_filters(v, pdims)
        self.sync(pages, by_ws)
        return {'model': self.model(), 'pages': pages, 'tooltip_pages': sorted(tooltip_sheets),
                'notes': sorted(set(self.notes)),
                'rules': 'Fabric skills: semantic-model-authoring + powerbi-report-cli (catalog '
                         f'{(self.cat or {}).get("cli_version", "not loaded")})'}

    def split_visual(self, v):
        """Tableau side-by-side measure panes -> one visual per measure, sharing the sheet's box."""
        sp = v.pop('split', None)
        tips = v.pop('_tips', [])
        if not sp:
            return [v]
        out, n = [], len(sp['measures'])
        pos = v.get('position') or {}
        ys = v['roles'].get('Y', [])
        for k, mlabel in enumerate(sp['measures']):
            vv = json.loads(json.dumps(v))
            vv['id'] = f'{v["id"]}{chr(97 + k)}'
            vv['roles']['Y'] = [f for f in ys if f['label'] == mlabel or f['name'] == mlabel.strip('[]')] or ys[k:k + 1]
            others, seen = [], {f['name'] for f in vv['roles']['Y']}
            for t in tips:
                o = self.attr_measure(t) if t.get('blend_attr') else self.field_of(t['label'], t.get('ds'), True)
                if o['name'] not in seen:
                    seen.add(o['name'])
                    others.append(o)
            vv['roles'].pop('Tooltips', None)
            if others and 'Tooltips' in self.roles_of(vv['visualType']):
                vv['roles']['Tooltips'] = others
            if pos:
                if sp['direction'] == 'horizontal':
                    w = pos['width'] // n
                    vv['position'] = {**pos, 'x': pos['x'] + k * w, 'width': w}
                else:
                    h = pos['height'] // n
                    vv['position'] = {**pos, 'y': pos['y'] + k * h, 'height': h}
            if k:
                vv['title'] = {'text': vv['roles']['Y'][0]['name'] if vv['roles']['Y'] else '', 'show': True}
            out.append(vv)
        return out

    def date_param_pair(self, page):
        dps = [s for s in page['slicers'] if s.get('parameter') and s['mode'] == 'Between']
        if len(dps) == 2:
            a, b = dps
            page['slicers'] = [s for s in page['slicers'] if s not in dps]
            merged = {'id': a['id'], 'visualType': 'slicer', 'mode': 'Between', 'selection': 'range',
                      'field': {'kind': 'Column', 'table': 'Report Period', 'name': 'Date', 'label': 'Report Period'},
                      'position': a['position'], 'default': {'from': a['default']['values'][0], 'to': b['default']['values'][0]},
                      'parameters': [a['parameter'], b['parameter']],
                      'note': f'Tableau "{a["parameter"]}"/"{b["parameter"]}" date parameters -> one Between slicer on a '
                              'disconnected date table; [Start/End Value] measures read MIN/MAX of the selection '
                              '(After/Before slicers render poorly)'}
            page['slicers'].insert(0, merged)

    def actions(self, pages, by_ws):
        names = {p['name'] for p in pages}
        for a in self.res['actions']:
            src = (a.get('source') or '').split(' / ')[0].strip()
            tgt = (a.get('target') or '').strip()
            page = next((p for p in pages if p['name'] == src), None)
            if page is None and a.get('kind') in ('PARAMETER', 'SET', 'NAVIGATE', 'FILTER'):  # sheet-level action: its sheet's page
                sheet = (a.get('source') or '').split(' / ')[-1].strip()
                page = next((p for p in pages if any(v['sheet'] == sheet for v in p['visuals'])), None)
            if page is not None:
                src = page['name']
            if a.get('kind') == 'FILTER' and tgt in names and tgt != src and page is not None:
                one = (a.get('source') or '').split(' / ', 1)[1].strip() if ' / ' in (a.get('source') or '') else None
                sheets = [v['sheet'] for v in (page or {}).get('visuals', []) if v['sheet'] not in (a.get('excludes') or [])
                          and (one is None or v['sheet'] == one)]
                fields = []
                for s in sheets:
                    for v in page['visuals']:
                        if v['sheet'] == s:
                            for role in ('Category', 'Series', 'Rows', 'Columns', 'Legend'):
                                for fr in v['roles'].get(role, []):
                                    if fr['kind'] == 'Column' and fr not in fields:
                                        fields.append(fr)
                            # a card shows its dimension as MIN/MAX/ATTR of a text/boolean/date field (Tableau "all
                            # fields" passes that dimension): the column itself is the drillthrough field
                            for fr in (x for xs in v['roles'].values() for x in xs):
                                if fr['kind'] == 'Aggregation' and fr.get('function') in ('Min', 'Max'):
                                    info = self.cols.get(fr['name']) or self.calcs.get(fr['name']) or {}
                                    if info.get('datatype') in ('string', 'boolean', 'date', 'datetime'):
                                        col_ref = {'kind': 'Column', 'table': fr['table'], 'name': fr['name'],
                                                   'label': fr['name']}
                                        if col_ref not in fields:
                                            fields.append(col_ref)
                tp = next(p for p in pages if p['name'] == tgt)
                if not fields:
                    # the source sheets show no dimension (KPI cards, totals): a Tableau click passes no value, so the action
                    # only opens the target page -> a navigation button on each source visual, and a back button
                    for s in sheets:
                        nb = {'sheet': s, 'target': tgt, 'action': a['name']}
                        if nb not in page.setdefault('nav_buttons', []):
                            page['nav_buttons'].append(nb)
                    tp.setdefault('nav_back', [])
                    if src not in tp['nav_back']:
                        tp['nav_back'].append(src)
                    continue
                dt = tp.get('drillthrough')
                if dt:  # several actions target this page: one drillthrough on the union of their fields
                    dt['fields'] += [f for f in fields if f not in dt['fields']]
                    dt['tableau_action'] += ', ' + a['name']
                    if src and src not in dt['from_page'].split(', '):
                        dt['from_page'] = ', '.join(x for x in (dt['from_page'], src) if x)
                else:
                    tp['drillthrough'] = {'from_page': src, 'fields': fields, 'tableau_action': a['name'],
                                          'trigger': a.get('trigger'),
                                          'note': 'Tableau click-to-filter -> Power BI drillthrough (right-click); keep a back button'}
            elif a.get('kind') == 'FILTER' and page is not None:
                page.setdefault('cross_filter', []).append({'action': a['name'], 'note': 'keep default filter interaction '
                                                            'between these visuals'})
            elif a.get('kind') == 'URL':
                for p in pages:
                    for v in p['visuals']:
                        if v['sheet'] == (a.get('source') or '').split(' / ')[-1]:
                            m = re.findall(r'<\[([^\]]+)\]>', a.get('url') or '')
                            # parameters cannot be in a column (computed at refresh): the build puts their default value
                            refs = [r for r in (self.field(x, False) for x in m if not x.startswith('Param:')) if r.get('table')]
                            url_col = f'{a["name"]} URL'
                            v['url_action'] = {'url': a['url'], 'fields': m, 'column': url_col,
                                               'table': refs[0]['table'] if refs else None,
                                               'how': 'Web URL column (dataCategory WebUrl) built in Power Query/DAX; '
                                                      'table webURL conditional formatting'}
                            uc = {'table': refs[0]['table'] if refs else None, 'name': url_col,
                                  'template': a['url'], 'fields': [r['name'] for r in refs]}
                            if not any(x['table'] == uc['table'] and x['name'] == url_col for x in self.url_columns):
                                self.url_columns.append(uc)  # one column per action, however many sheets use it
                            if v['visualType'] == 'pivotTable' and not v['roles'].get('Columns'):
                                # pivotTable takes web URLs on measure values only; the Tableau link sits on a
                                # dimension column, so a flat text table with a URL action becomes tableEx
                                v['visualType'] = 'tableEx'
                                v['roles'] = {'Values': v['roles'].get('Rows', []) + v['roles'].get('Values', [])}
                                v['url_action']['link_on'] = v['roles']['Values'][0]['name'] if v['roles']['Values'] else None
                                v['notes'] = [n for n in v['notes'] if 'pivotTable with stepped' not in n] + [
                                    'tableEx because the URL action needs a web URL on a dimension column (pivotTable '
                                    'supports it on values only); check rows render (powerbi-report-cli tableEx pitfall)']
                                v['catalog_errors'] = self.check(v)
            elif a.get('kind') == 'NAVIGATE' and page is not None:
                page['static'].append({'type': 'actionButton', 'action': 'PageNavigation', 'target': a.get('target'),
                                       'note': f'Tableau navigation action "{a["name"]}"'})
            elif a.get('kind') == 'PARAMETER' and a.get('parameter') and page is not None:
                page['notes'].append(f'PARAMETER action "{a["name"]}": a click sets [{a["parameter"]}] from '
                                     f'[{a.get("source_field")}]; {a.get("powerbi")}')
            elif a.get('kind') in ('PARAMETER', 'SET', 'HIGHLIGHT') and page is not None:
                page['notes'].append(f'{a["kind"]} action "{a["name"]}": no direct equivalent; review')

    def sync(self, pages, by_ws=None):
        by_ws = by_ws or {}
        for p in pages:
            dt = p.get('drillthrough')
            if not dt:
                continue
            for src_name in dt['from_page'].split(', '):
                src = next((x for x in pages if x['name'] == src_name), None)
                for s in (src or {}).get('slicers', []):
                    if s['field']['name'] in {x['field']['name'] for x in p['slicers']}:
                        continue
                    scope = None
                    if s.get('scope_sheets') is not None:  # a quick filter reaches only the sheets that carry it
                        scope = self.filter_scope(s['tableau_field'], [v['sheet'] for v in p['visuals']], by_ws)
                        if not scope:
                            continue
                    p['slicers'].append({**s, 'id': s['id'] + 'h', 'hidden': True, 'scope_sheets': scope,
                                         'note': 'hidden synced copy so the source page selection applies here'})

    def filter_scope(self, field, sheets, by_ws):
        """Sheets a Tableau quick filter applies to: the ones holding a filter on that field (Tableau writes the
        filter into every worksheet it is applied to). None when it is a data-source filter (applies to all)."""
        if not field or any(field in json.dumps(d.get('filters') or []) for d in self.res.get('datasources', [])):
            return None
        return [s for s in sheets if any(f['field'] == field for f in (by_ws.get(s) or {}).get('filters', []))]

    def sort_by(self, ws, v):
        """Tableau '[Dim] DESC by [Measure]' (computed sort) on the chart's category -> {'measure', 'direction'};
        '[Dim] ASC by natural-sort' -> {'category': True, 'direction'}. None when no sort is on the category."""
        cats = [f['label'] for f in v['roles'].get('Category', [])]
        for srt in ws.get('sorts') or []:
            m = re.match(r'^\[(.+?)\]\s+(ASC|DESC)\s+by\s+(.+)$', srt)
            if not m or not cats or m.group(1) != cats[0]:
                continue
            direction = 'Descending' if m.group(2) == 'DESC' else 'Ascending'
            by = m.group(3).strip()
            if by.startswith('['):
                fr = self.field(by.strip('[]'), True)
                if fr['kind'] == 'Measure':
                    return {'measure': fr['name'], 'direction': direction}
            elif 'natural' in by:
                return {'category': True, 'direction': direction}
        return None

    # ---------- model ----------
    def model(self):
        tables = {}
        for c in self.res['columns']:
            t = c['table'] or '<unknown table>'
            tables.setdefault(t, {'name': t, 'suggested_name': pretty_table(c['table']), 'columns': [],
                                  'calculated_columns': [], 'measures': []})
            if c.get('source') and not tables[t].get('source_table'):
                tables[t]['source_table'] = c['source'].split('.')[0]  # physical relation behind the model table
            tables[t]['columns'].append({'name': c['name'], 'source_column': c.get('source_column'),
                                         'datatype': c['datatype'], 'hidden': bool(c.get('hide')) or None,
                                         'summarizeBy': 'none' if c['datatype'] in ('integer',) and re.search(r'\b(id|year|number|num|key|code)\b', c['name'], re.I) else None})
        for c in self.res['calculations']:
            if c['pbi_kind'] == 'calculated column':
                t = c.get('table') or '<unknown table>'
                tables.setdefault(t, {'name': t, 'suggested_name': pretty_table(t), 'columns': [], 'calculated_columns': [], 'measures': []})
                tables[t]['calculated_columns'].append({'name': c['name'], 'dax': c['dax'], 'prefer': 'Power Query column',
                                                       'confidence': c['confidence'], 'notes': c['dax_notes']})
        for g in self.res['groups']:
            t = g.get('table') or '<unknown table>'
            tables.setdefault(t, {'name': t, 'suggested_name': pretty_table(t), 'columns': [], 'calculated_columns': [], 'measures': []})
            tables[t]['calculated_columns'].append({'name': g['name'], 'dax': g['dax'], 'prefer': 'Power Query conditional column',
                                                   'confidence': 'high', 'notes': ['Tableau group']})
        for m in self.measures.values():
            t = m['table'] or '<home table>'
            tables.setdefault(t, {'name': t, 'suggested_name': pretty_table(t) if m['table'] else None, 'columns': [],
                                  'calculated_columns': [], 'measures': []})
            tables[t]['measures'].append(m)
        rels = []
        for ds in self.res['datasources']:
            c = ds['connection'] or {}
            src = c.get('tds', c)
            for r in (src.get('logical') or {}).get('relationships', []):
                keys = (r['on'] or '').count('=')
                rel = {'from': r['between'][0] if r['between'] else None, 'to': r['between'][1] if len(r['between']) > 1 else None,
                       'on': r['on'], 'crossFilter': 'OneDirection',
                       'note': 'composite key: build one key column in Power Query' if keys > 1 else None}
                rel.update(self.logical_rel(r))
                rels.append(rel)
            self._joins(src.get('relations'), rels)
        published = [ds for ds in self.res['datasources']
                     if (ds['connection'] or {}).get('published') and not (ds['connection'] or {}).get('tds')]
        if published and not rels:
            used = set(tables) | {t.get('source_table') for t in tables.values()}  # model names and source (family) names
            for ds in published:
                for g in ds['connection'].get('guessed_joins') or []:
                    if g['many'] not in used or g['one'] not in used:
                        continue  # a table the workbook does not use is not built: no relationship to it
                    rels.append({'from': g['many'], 'to': g['one'], 'crossFilter': 'OneDirection', 'guessed': True,
                                 'on': f'[{g["many"]}].[{g["physical_many"]}] = [{g["one"]}].[{g["physical_one"]}]',
                                 'note': f'GUESSED from Tableau join-column naming ({g["many_col"]} / {g["one_col"]}): '
                                         f'{g["basis"]}. Confirm with the published .tds'})
            if rels:
                self.notes.append(f'relationships GUESSED ({len(rels)}): the source is a published data source without its '
                                  '.tds; the joins come from Tableau\'s renamed join columns. Confirm them with the .tds')
            else:
                self.notes.append('relationships UNKNOWN: the source is a published data source; run with --tds to get the '
                                  'joins, then model them as single-direction relationships (star schema)')
        model = {'tables': list(tables.values()), 'relationships': rels, 'url_columns': self.url_columns,
                 'parameters': [{'name': p['name'], **(p.get('powerbi') or {})} for p in self.res['parameters']],
                 'parameter_dimensions': list(self.param_dims.values()), 'notes': []}
        # Tableau data source filters (every sheet of the source sees only these rows) -> bhayankar_build.datasource_filters
        dsf = []
        for ds in self.res['datasources']:
            for f in ds.get('filters') or []:
                if not f.get('field') or (f.get('rule') or '').startswith('ALL'):
                    continue
                prev, self._ws = getattr(self, '_ws', None), {'datasources': [ds['name']]}
                try:
                    fs = self.filter_spec({**f, 'ds': ds['name']})
                finally:
                    self._ws = prev
                dsf.append({**fs, 'datasource': ds['name']})
        if dsf:
            model['datasource_filters'] = dsf
        if self.res.get('blends'):
            # Tableau blend links -> one shared dimension per linked field (bhayankar_build.blend_dimensions)
            model['blend_links'] = self.res['blends']
        if self.date_roles:
            model['date_table'] = {'build': 'Power Query date table (modeling-guidelines), mark as date table, disable auto date/time',
                                   'columns': sorted({v for v in DATE_COL.values()} - {'Hour'}),
                                   'active_relationship_to': self.date_roles[0],
                                   'other_date_columns': self.date_roles[1:] or None}
            if len(self.date_roles) > 1:
                model['notes'].append('several dates use date parts: one active relationship, others inactive + USERELATIONSHIP '
                                      'measures, or a second date table')
        fp = []
        for p in self.res['parameters']:
            users = [c for c in self.res['calculations'] if f'[Param:{p["name"]}]' in (c['formula'] or '')
                     and re.match(r'^\s*CASE\s+\[Param:', c['formula'] or '', re.I)]
            if users and p.get('members'):
                fp.append({'parameter': p['name'], 'calcs': [c['name'] for c in users],
                           'note': 'CASE on a parameter that picks measures/fields: build a field parameter, not SWITCH '
                                   '(field-parameters.md)'})
        for c in self.res['calculations']:
            if c['pbi_kind'] == 'field parameter':
                fp.append({'parameter': c['name'], 'calcs': [c['name']], 'fields': c.get('dax_fields'),
                           'note': 'Tableau CASE on a parameter returning columns: field parameter over these columns '
                                   '(table_operations CreateFieldParameter via the modeling MCP)'})
        if fp:
            model['field_parameter_candidates'] = fp
        abbrev = sorted({a for m in self.measures.values() for a in (m['abbreviations'] or [])})
        if abbrev:
            model['notes'].append(f'abbreviations in names: {", ".join(abbrev)}: expand with the business glossary '
                                  '(naming-conventions)')
        return model

    def logical_rel(self, r):
        """A Tableau relationship (logical layer) with resolved keys -> table-qualified 'on' the build reads,
        '[ORDERS].[Region] = [PEOPLE].[Region] AND ...', plus the one side Tableau states (Performance options:
        cardinality One) and, when Power BI cannot express it, 'unsupported' (the build logs DECISION NEEDED).
        A key on a calculated field is kept as {'calc': caption, 'table': its model table} when the calc is a
        calculated column (Power BI relates on calculated columns); otherwise it is unsupported."""
        keys = r.get('keys') or []
        if not keys:
            return {}
        out = {'tableau_on': r['on'], 'integrity': r.get('integrity')}
        ends = []
        for k in keys:
            if k['op'] != '=':
                return {**out, 'unsupported': f'operator {k["op"]} ({r["on"]}): Power BI relationships match equal keys '
                                              'only; use a Power Query merge or a DAX band lookup'}
            pair = []
            for s in (k['left'], k['right']):
                if s.get('calc') is not None:
                    c = self.calcs.get(s['field'])
                    if not c or c.get('pbi_kind') != 'calculated column' or not c.get('table'):
                        return {**out, 'unsupported': f'key [{s["field"]}] is a calculated field ({s["calc"]}) that is not a '
                                                      'row-level column: add it as a Power Query column on its table'}
                    pair.append((c['table'], s['field'], True))
                elif s.get('table'):
                    pair.append((s['table'], s['column'], False))
                else:
                    return {**out, 'unsupported': f'key [{s["field"]}] not found in the data source columns'}
            ends.append(pair)
        if len({(a[0], b[0]) for a, b in ends}) != 1:
            return {**out, 'unsupported': f'the key pairs of {r["on"]} join different tables'}
        (ta, _, _), (tb, _, _) = ends[0]
        out.update({'from': ta, 'to': tb, 'on': ' AND '.join(f'[{a[0]}].[{a[1]}] = [{b[0]}].[{b[1]}]' for a, b in ends),
                    'calc_keys': [x[1] for a, b in ends for x in (a, b) if x[2]] or None})
        uniq, et = r.get('unique') or [], r.get('end_tables') or []
        if len(uniq) == 2 and any(uniq):
            # which end is the left key's table: by name when the ends are single tables, else Tableau's own order
            # (first end = the left field, as in the saved workbooks)
            ua, ub = (uniq if et[:1] != [tb] and et[1:2] != [ta] else uniq[::-1])
            out['one_side'] = 'both' if ua and ub else ('a' if ua else 'b')
        if len(ends) > 1:
            out['note'] = f'composite key ({len(ends)} fields): the build adds one key column per table in Power Query'
        return out

    def _joins(self, r, rels):
        if not r or 'join' not in r:
            return
        ins = r.get('inputs') or []
        names = [x.get('alias') or x.get('table') for x in ins if x and 'table' in x]
        for on in r.get('on') or []:
            if not on:
                continue
            tabs = re.findall(r'\[([^\]]+)\]\.\[', on)
            rels.append({'from': tabs[0] if tabs else (names[0] if names else None),
                         'to': tabs[1] if len(tabs) > 1 else (names[1] if len(names) > 1 else None),
                         'on': on, 'tableau_join': r['join'], 'crossFilter': 'OneDirection',
                         'note': 'Tableau join flattened the tables: in Power BI keep them separate (star schema) '
                                 'and relate them'})
        for x in ins:
            self._joins(x, rels)


# ---------- rendering ----------
def fref(f):
    if not f:
        return '?'
    if f['kind'] == 'Measure':
        return f'[{f["name"]}]'
    if f['kind'] == 'Aggregation':
        return f"{f['function']}('{f['table'] or '?'}'[{f['name']}])"
    return f"'{f['table'] or '?'}'[{f['name']}]"


def render_page(p):
    L = [f'### Page "{p["name"]}" {p["width"]}x{p["height"]}' + ('' if p.get('visible', True) else ' (hidden)')]
    L += [f'- note: {n}' for n in p.get('notes', [])]
    if p.get('drillthrough'):
        dt = p['drillthrough']
        L.append(f'- drillthrough from "{dt["from_page"]}" on {", ".join(fref(f) for f in dt["fields"]) or "?"} ({dt["note"]})')
    for f in p.get('page_filters', []):
        L.append(f'- page filter: {fref(f["field"])} {f.get("type")} {f.get("values") or f.get("measure") or ""}'
                 + (' inverted' if f.get('inverted') else '') + (f' [{f["note"]}]' if f.get('note') else ''))
    for s in p['slicers']:
        pos = s.get('position') or {}
        d = s.get('default') or {}
        dflt = (f'{d.get("from")}..{d.get("to")}' if 'from' in d else
                (('NOT ' if d.get('inverted') else '') + ', '.join(str(x) for x in d.get('values') or [])) if d else 'none')
        L.append(f'- slicer {s["visualType"]}' + (f'/{s["mode"]}' if s.get('mode') else '') + f' {s["selection"]} on {fref(s["field"])}'
                 f' @{pos.get("x")},{pos.get("y")} {pos.get("width")}x{pos.get("height")} default: {dflt}'
                 + (' (hidden, synced)' if s.get('hidden') else '') + (f' [{s["note"]}]' if s.get('note') else '')
                 + (f' [{d["note"]}]' if d.get('note') else ''))
    for v in p['visuals']:
        pos = v.get('position') or {}
        roles = '; '.join(f'{r}: {", ".join(fref(f) for f in fs)}' for r, fs in v['roles'].items() if fs)
        L.append(f'- **{v["id"]} "{v["sheet"]}" → {v["visualType"]}** @{pos.get("x")},{pos.get("y")} '
                 f'{pos.get("width")}x{pos.get("height")} [{v["confidence"]}] {roles}')
        t = v.get('title') or {}
        if t.get('measure'):
            L.append(f'  - title: measure [{t["measure"]}] ("{t["text"]}")')
        elif t.get('text') and t.get('show', True):
            L.append(f'  - title: "{t["text"]}"')
        if v.get('data_labels'):
            L.append('  - data labels on')
        for r in v.get('reference_lines', []):
            L.append(f'  - {r["type"]}: [{r["measure"]}] label "{r["label"]}"' + (f' ({r["note"]})' if r.get('note') else ''))
        for s in v.get('sort', []):
            L.append(f'  - sort: {s}')
        if v.get('colors'):
            L.append(f'  - colors: {json.dumps(v["colors"]["map"], ensure_ascii=False)}')
        for f in v.get('visual_filters', []):
            if f.get('type') == 'RowCondition':
                L.append(f'  - row condition (fold into the measures): [{f["calc"]}] = "{f["value"]}" -> `{f["condition"]}`')
                continue
            L.append(f'  - visual filter: {fref(f["field"])} {f.get("type")} {f.get("values") or f.get("measure") or ""}'
                     + (' inverted' if f.get('inverted') else '') + (f' [{f["note"]}]' if f.get('note') else ''))
        if v.get('url_action'):
            L.append(f'  - link: {v["url_action"]["how"]}; fields {v["url_action"]["fields"]}')
        if v.get('tooltip'):
            L.append(f'  - tooltip: {v["tooltip"]}')
        for n in v.get('notes', []):
            L.append(f'  - note: {n}')
        for e in v.get('catalog_errors', []):
            L.append(f'  - CATALOG CHECK FAILED: {e}')
    for s in p['static']:
        pos = s.get('position') or {}
        L.append(f'- {s["type"]}' + (f' "{(s.get("text") or "")[:80]}"' if s.get('text') else '')
                 + (f' {s["image"]}' if s.get('image') else '') + (f' → {s["target"]}' if s.get('target') else '')
                 + (f' @{pos.get("x")},{pos.get("y")} {pos.get("width")}x{pos.get("height")}' if pos else ''))
    if p.get('cross_filter'):
        L.append(f'- interactions: keep filtering for {", ".join(c["action"] for c in p["cross_filter"])}; others None')
    else:
        L.append('- interactions: set all visual-to-visual interactions to None (Tableau had no filter action on this page)')
    return L


def render_model(m):
    L = ['## Power BI model']
    for t in m['tables']:
        sug = f' (suggested name "{t["suggested_name"]}")' if t.get('suggested_name') and t['suggested_name'] != t['name'] else ''
        L.append(f'### Table {t["name"]}{sug}')
        if t['columns']:
            L.append('- columns: ' + ', '.join(f'{c["name"]}' + (f' ←{c["source_column"]}' if c.get('source_column') and c['source_column'] != c['name'] else '')
                                               + (' (hidden)' if c.get('hidden') else '') + (' (no summarize)' if c.get('summarizeBy') else '')
                                               for c in t['columns']))
        for c in t['calculated_columns']:
            L.append(f'- calc column **{c["name"]}** [{c["confidence"]}, prefer {c["prefer"]}]: `{c["dax"]}`'
                     + ''.join(f'\n  - {n}' for n in c['notes'] if 'prefer a Power Query' not in n))
        for x in t['measures']:
            nm = x['name'] + (f' → suggest "{x["suggested_name"]}"' if x.get('suggested_name') else '')
            L.append(f'- measure **{nm}** [{x["confidence"]}] fmt `{x["formatString"]}` folder `{x["displayFolder"]}`: `{x["dax"]}`'
                     + ''.join(f'\n  - {n}' for n in x['notes']))
    for r in m['relationships']:
        L.append(f'- relationship {r["from"]} → {r["to"]} ON {r["on"]} ({r["crossFilter"]})' + (f' [{r["note"]}]' if r.get('note') else ''))
    if m.get('date_table'):
        dt = m['date_table']
        L.append(f'- date table: {dt["build"]}; active relationship to [{dt["active_relationship_to"]}]'
                 + (f'; other dates {dt["other_date_columns"]}' if dt.get('other_date_columns') else ''))
    for p in m['parameters']:
        L.append(f'- parameter **{p["name"]}**: {p.get("pattern")}' + (f'\n  - table: `{p["table_dax"]}`' if p.get('table_dax') else '')
                 + f'\n  - measure: `{p.get("measure_dax")}`')
    for f in m.get('field_parameter_candidates', []):
        L.append(f'- field parameter candidate: {f["parameter"]} used by {", ".join(f["calcs"])} ({f["note"]})')
    L += [f'- note: {n}' for n in m.get('notes', [])]
    return L


def dax_script(spec, measure_host_fallback=None):
    """DAX query view script: DEFINE every measure, ready to test and 'update model'."""
    if measure_host_fallback is None:  # measures without a home table go to the table hosting most measures
        homes = [m['table'] for t in spec['model']['tables'] for m in t['measures'] if m.get('table')]
        measure_host_fallback = max(set(homes), key=homes.count) if homes else '<Fact table>'
    L = ['// Bhayankar extractor: draft measures for DAX query view. Review every [medium]/[low] item first.',
         '// Measures are hosted on their fact table (semantic-model-authoring: no single measures table).', 'DEFINE']
    for t in spec['model']['tables']:
        for m in t['measures']:
            if not m['dax'] or m['dax'].startswith('/*'):
                L.append(f'    // TODO {m["name"]}: {m["dax"]}')
                continue
            host = m['table'] or measure_host_fallback
            moved = '' if m['table'] else f' (no home table found: placed on {host}, check)'
            L.append(f'    // [{m["confidence"]}] {m["description"][:150]}{moved}')
            L.append(f"    MEASURE {q(host)}[{m['name']}] = {m['dax']}")
    names = [m['name'] for t in spec['model']['tables'] for m in t['measures'] if m['dax'] and not m['dax'].startswith('/*')]
    L.append('EVALUATE')  # ROW() tolerates mixed number/text results, a table constructor does not
    L.append('    ROW(' + ', '.join(f'"{n}", [{n}]' for n in names[:20]) + ')' if names else '    { BLANK() }')
    return '\n'.join(L) + '\n'
