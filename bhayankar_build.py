#!/usr/bin/env python3
r"""Bhayankar build: rebuild spec -> Power BI project (TMDL semantic model + PBIR report).

Standalone: when output\<wb>\ has no rebuild spec, or input\ changed since it was made, the extractor runs first.

Reads, per workbook:
  input\<wb>.twb/.twbx (+ .tds)   extracted on demand into the two files below
  output\<wb>\bhayankar_rebuild.json, output\<wb>\bhayankar_digest.json   (from bhayankar_extractor.py)
  input\sources.json     where each source relation loads from (sample CSV folder or SQL endpoint)
  input\decisions.json   answers to DECISION NEEDED items, measure renames
  input\table_map.json   source table -> Power BI table (already applied by the extractor)
Writes output\<wb>\pbip\<wb>.pbip (+ .SemanticModel, .Report), then validates it offline with
bhayankar_validate.py (the PBIR rules of `powerbi-report-author validate`, with the schemas bundled in resources\).

Usage (from the toolkit folder):
  python bhayankar_build.py                  # every workbook in input\ (and any already extracted in output\)
  python bhayankar_build.py "<workbook name>"

Rules: semantic-model-authoring (TMDL guidelines, measures on fact tables, formatString, Power Query date
table, single-direction relationships) and powerbi-report-cli (roles from the CLI catalog, modern visuals,
validate after writing). Python 3 standard library only: no Node.js, no internet. Optional second opinion: set
BHAYANKAR_CLI_CHECK=1 to also run `powerbi-report-author validate` when that CLI is installed.
"""
import csv
import hashlib
import copy
import json
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import bhayankar_scaffold as bs  # noqa: E402
import bhayankar_toolkit as bt  # noqa: E402
import bhayankar_validate as bv  # noqa: E402
import bhayankar_modelcheck as bmc  # noqa: E402
INPUT_DIR = bt.input_dir()
OUTPUT_DIR = bt.output_dir()
VISUAL_SCHEMA = 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.9.0/schema.json'
PAGE_SCHEMA = 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json'
TMDL_TYPE = {'string': 'string', 'integer': 'int64', 'real': 'double', 'datetime': 'dateTime', 'date': 'dateTime',
             'boolean': 'boolean'}
M_TYPE = {'string': 'type text', 'integer': 'Int64.Type', 'real': 'type number', 'date': 'type date',
          'boolean': 'type logical'}
DARK, GREY = '#333333', '#605E5C'
from bhayankar_dax import param_table_name  # noqa: E402
from bhayankar_report import load_catalog, REL_UNITS  # noqa: E402

CATALOG = load_catalog() or {}


# Date table columns beyond Year/Quarter/Month/Year Month/Day (bhayankar_report.DATE_COL binds Tableau date parts to them)
DATE_EXTRA = [('Week Start', 'dateTime'), ('Week', 'int64'), ('Year Quarter', 'string'), ('Year Quarter Number', 'int64'),
              ('Day of Week', 'string'), ('Day of Week Number', 'int64'), ('Month Day', 'string'), ('Month Day Number', 'int64')]


def hid(key):
    return hashlib.sha1(key.encode('utf-8')).hexdigest()[:20]  # stable ids (random ids break Desktop reload)


def table_file(name):
    """File name for a table's .tmdl: the table keeps its own name inside the file, but the characters Windows does not
    allow in a file name must not be in the file name (a colon creates a hidden stream and an empty file, so the
    whole table was lost)."""
    return re.sub(r'[<>:"/\\|?*]', '_', name).rstrip(' .') or 'table'


def fid(key):
    return 'Filter' + hashlib.sha1(('filter:' + key).encode('utf-8')).hexdigest()[:24]


def tq(name):
    """TMDL object name: always single-quoted, ' escaped."""
    return "'" + name.replace("'", "''") + "'"


def tmdl_file(name):
    """File name for a table's .tmdl: characters Windows forbids (\\ / : * ? " < > |, control characters, a trailing dot
    or space) are %-encoded as Power BI Desktop does, so "Fame/Shame Index" or "Divested?" can be written. The table name
    inside the file is unchanged; Desktop reads it from there."""
    out = re.sub(r'[\\/:*?"<>|%\x00-\x1f]', lambda m: '%{:02X}'.format(ord(m.group())), name)
    out = re.sub(r'[. ]$', lambda m: '%{:02X}'.format(ord(m.group())), out)
    return out + '.tmdl'


def mstr(s):
    return '"' + s.replace('"', '""') + '"'


def mid(name):
    """M identifier: plain when it is a regular identifier, otherwise #\"quoted\" (spaces, hyphens, digits first...)."""
    return name if re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]*', name) else '#' + mstr(name)


def mfield(name):
    """M field access, e.g. [WO_ID] or [#\"Dept Num-WO Owner Now\"]."""
    return '[' + mid(name) + ']'


def load_json(path, default=None):
    if not os.path.exists(path):
        return default
    with open(path, encoding='utf-8') as f:
        return json.load(f)


# ---------- PBIR expression helpers ----------
def lit(v):
    return {'expr': {'Literal': {'Value': v}}}


def sq(s):
    return lit("'" + str(s).replace("'", "''") + "'")


def boolean(b):
    return lit('true' if b else 'false')


def num(n):
    return lit(f'{n}D')


def colour(hexv):
    return {'solid': {'color': sq(hexv)}}


def col_ref(entity, prop):
    return {'Column': {'Expression': {'SourceRef': {'Entity': entity}}, 'Property': prop}}


def meas_ref(entity, prop):
    return {'Measure': {'Expression': {'SourceRef': {'Entity': entity}}, 'Property': prop}}


def proj(field, display=None, active=False):
    inner = field.get('Column') or field.get('Measure') or field.get('Aggregation', {}).get('Expression', {}).get('Column')
    entity = inner['Expression']['SourceRef']['Entity']
    p = {'field': field, 'queryRef': f'{entity}.{inner["Property"]}', 'nativeQueryRef': inner['Property']}
    if 'Aggregation' in field:
        p['queryRef'] = f'Min({entity}.{inner["Property"]})'
    if display:
        p['displayName'] = display
    if active:
        p['active'] = True
    return p


def padding(v):
    return [{'properties': {'top': num(v), 'bottom': num(v), 'left': num(v), 'right': num(v)}}]


def call_args(dax, fn):
    """Top-level arguments of the first fn( ... ) call in a DAX expression (commas inside strings, quoted names or
    nested calls do not split), or None when there is no such call."""
    mm = re.search(r'\b' + fn + r'\(', dax)
    if not mm:
        return None
    depth, cur, args, quote = 0, '', [], None
    for ch in dax[mm.end():]:
        if quote:
            quote = None if ch == quote else quote
        elif ch in '"\'':
            quote = ch
        elif ch == '(':
            depth += 1
        elif ch == ')':
            if depth == 0:
                return args + [cur.strip()]
            depth -= 1
        elif ch == ',' and depth == 0:
            args.append(cur.strip())
            cur = ''
            continue
        cur += ch
    return None


def key_owner(col, table):
    """True when a key column is named after its table: EVENT_ID in Events, WO Number in WORK_ORDERS ('wo'
    abbreviates 'workorder'). Words such as ID / NUM / NUMBER / KEY / CODE do not count."""
    skip = {'id', 'num', 'number', 'key', 'code', 'no', 'nbr'}
    col = re.sub(r'([a-z0-9])([A-Z])', r'\1 \2', col)  # ProductKey -> Product Key
    words = [w for w in re.findall(r'[a-z0-9]+', col.lower()) if w not in skip]
    table = re.sub(r'([a-z0-9])([A-Z])', r'\1 \2', table)  # DimProduct -> Dim Product
    toks = [re.sub(r's$', '', re.sub(r'ies$', 'y', t)) for t in re.findall(r'[a-z0-9]+', table.lower())]  # Events -> event, Categories -> category

    def abbrev(w, t):
        it = iter(t)
        return w == t or (w and t and w[0] == t[0] and len(w) <= len(t) and all(ch in it for ch in w))
    return any(abbrev(re.sub(r's$', '', w), t) for w in words for t in toks)


def tableau_date(v):
    """'#2026-01-01#' -> (2026, 1, 1)"""
    m = re.match(r'#?(\d{4})-(\d{2})-(\d{2})', (v or '').strip('"'))
    return tuple(int(x) for x in m.groups()) if m else None


class Build:
    def __init__(self, wb_name):
        self.wb = wb_name
        self.wb_out = os.path.join(OUTPUT_DIR, wb_name)
        self.spec = load_json(os.path.join(self.wb_out, 'bhayankar_rebuild.json'))
        self.digest = load_json(os.path.join(self.wb_out, 'bhayankar_digest.json'))
        self.sources = load_json(os.path.join(INPUT_DIR, 'sources.json'), None)
        if self.sources and self.sources.get('sample_folder') and not os.path.isabs(self.sources['sample_folder']):
            # relative to input\, so a shared folder works on any PC
            self.sources['sample_folder'] = os.path.normpath(os.path.join(INPUT_DIR, self.sources['sample_folder']))
        self.sample_note = None
        if self.sources is None:  # no sources.json: use input\<workbook>_Data when it holds the extracts, else warehouse
            self.sources = {'mode': 'warehouse', 'tables': {}}
            for cand in (wb_name + '_Data', wb_name.replace(' ', '_') + '_Data'):
                folder = os.path.join(INPUT_DIR, cand)
                if os.path.isdir(folder) and any(f.lower().endswith('.csv') for f in os.listdir(folder)):
                    self.sources = {'mode': 'sample', 'sample_folder': folder, 'tables': {}}
                    self.sample_note = f'no input\\sources.json: loading the sample CSVs in {folder} (UseSampleData = true)'
                    break
        self.decisions = (load_json(os.path.join(INPUT_DIR, 'decisions.json'), {}) or {}).get(wb_name, {})
        self.renames = dict(self.decisions.get('measure_names', {}))  # a copy: the build adds its own (name clashes)
        self.log = [self.sample_note] if self.sample_note else []
        self.pbip_dir = self.decisions.get('pbip_folder') or os.path.join(self.wb_out, 'pbip')
        self.item = self.item_name()
        self.model_dir = os.path.join(self.pbip_dir, f'{self.item}.SemanticModel')
        self.report_dir = os.path.join(self.pbip_dir, f'{self.item}.Report')

    MAX_PATH = 259  # Windows MAX_PATH (260) minus the terminating NUL; Desktop cannot open a longer file path
    REPORT_TAIL = len('.Report\\definition\\pages\\') + 20 + len('\\visuals\\') + 20 + len('\\visual.json')

    def item_name(self):
        """Name of the .Report / .SemanticModel folders. The .pbip keeps the workbook name; the folders get a shorter
        name when <folder>\\<name>.Report\\definition\\pages\\<id>\\visuals\\<id>\\visual.json would pass MAX_PATH
        ("path too long" in Desktop). When even a short name does not fit, the project moves to a short folder."""
        longest_table = max([len(t['name']) for t in (self.spec or {}).get('model', {}).get('tables', [])] + [40])
        tail = max(self.REPORT_TAIL, len('.SemanticModel\\definition\\tables\\') + longest_table + len('.tmdl'))
        # the full path Windows uses: a short 8.3 folder name (JOHNDO~1) is expanded (John.Doe); measuring the short
        # form let a 261-character file through, which Desktop then could not open
        budget = self.MAX_PATH - len(os.path.realpath(self.pbip_dir)) - 1 - tail
        if len(self.wb) <= budget:
            return self.wb
        if budget < 8:
            self.pbip_dir = self.fallback_folder()
            self.log.append(f'PATH TOO LONG for output\\{self.wb}\\pbip: project written to {self.pbip_dir} instead '
                            '(or set decisions.json "pbip_folder")')
            budget = self.MAX_PATH - len(os.path.realpath(self.pbip_dir)) - 1 - tail
            if len(self.wb) <= budget:
                return self.wb
        words, short = self.wb.split(), ''
        for w in words:
            if len((short + ' ' + w).strip()) > budget:
                break
            short = (short + ' ' + w).strip()
        short = short or self.wb[:max(budget, 8)].strip()
        self.log.append(f'folders named "{short}.Report" / "{short}.SemanticModel" (the full name passes the Windows '
                        f'{self.MAX_PATH + 1}-character path limit); open {self.wb}.pbip as usual')
        return short

    MARKER = '.bhayankar_build'  # written into a fallback folder: only a folder holding it may be rebuilt

    def fallback_folder(self):
        """C:\\PBI\\<initials> for a project whose output\\ path is too long. A folder of that name that this tool did
        not make (e.g. a hand-fixed copy of the report) is never touched: the next free <initials>_2, _3, ... is used."""
        root = os.path.join(os.environ.get('SystemDrive', 'C:') + os.sep, 'PBI')
        base = re.sub(r'[^A-Za-z0-9]', '', ''.join(w[0] for w in self.wb.split()))[:12] or 'Report'
        for i in range(1, 100):
            cand = os.path.join(root, base if i == 1 else f'{base}_{i}')
            marker = os.path.join(cand, self.MARKER)
            owner = None
            if os.path.exists(marker):
                with open(marker, encoding='utf-8') as f:
                    owner = f.read().strip()
            if not os.path.exists(cand) or owner == self.wb:
                os.makedirs(cand, exist_ok=True)
                with open(marker, 'w', encoding='utf-8') as f:
                    f.write(self.wb)
                return cand
        raise SystemExit(f'no free folder under {root} for {self.wb}: set decisions.json "pbip_folder"')

    # =====================================================================================
    # semantic model
    # =====================================================================================
    def tds_relations(self):
        """Physical relations from the merged .tds: alias -> {'table': [db].[t]} or {'custom_sql': ...}."""
        out = {}

        def walk(r):
            if not r:
                return
            if 'join' in r:
                for x in r['inputs']:
                    walk(x)
            elif r.get('alias'):
                out[r['alias']] = r
        for ds in self.digest['datasources']:
            c = ds['connection'] or {}
            walk((c.get('tds') or c).get('relations'))
        return out

    def sample_rows(self, relation):
        """Rows of a relation's sample CSV after the sources.json transforms (for key uniqueness checks)."""
        if self.sources.get('mode') != 'sample':
            return None
        # sources.json entry, else the CSV the partition reads by default (<relation>.csv in the sample folder)
        cfg = self.sources.get('tables', {}).get(relation) or {'file': f'{relation}.csv'}
        path = os.path.join(self.sources['sample_folder'], cfg['file'])
        if not os.path.exists(path):
            return None
        with open(path, encoding='utf-8-sig', newline='') as f:
            rows = list(csv.DictReader(f))
        for k, v in (cfg.get('where') or {}).items():
            rows = [r for r in rows if r.get(k) == v]
        if cfg.get('columns'):  # renamed copies; the other CSV columns stay (keys need not be mapped)
            rows = [{**r, **{alias: r.get(src) for alias, src in cfg['columns'].items()}} for r in rows]
        return rows

    def build_model(self):
        spec_model = self.spec['model']
        rels_spec = spec_model['relationships']
        tmap = load_json(os.path.join(INPUT_DIR, 'table_map.json'), {}) or {}
        to_model = lambda src: tmap.get(src, src)
        relations = self.tds_relations()
        tables = {}
        self.moved_tables, self.moved_cols = {}, {}
        self.param_rowcalcs = []  # (spec table, calc column) not created because they read a parameter
        for t in spec_model['tables']:
            if t['name'].startswith('<'):
                continue
            src = t.get('source_table') or next((s for s, m in tmap.items() if m == t['name']), t['name'])
            same = next((n for n in tables if n.lower() == t['name'].lower()), None)
            if same:
                # Power BI names are case-insensitive (and so are Windows file names): "PM OWNER" and "PM Owner" from two
                # published sources are one table. Two tables would overwrite each other's .tmdl and break visuals.
                host = tables[same]
                for c in t['columns']:
                    hit = next((x for x in host['columns'] if x['name'].lower() == c['name'].lower()), None)
                    if hit:
                        self.moved_cols[(t['name'], c['name'])] = (same, hit['name'])
                    else:
                        host['columns'].append(dict(c))
                self.moved_tables[t['name']] = same
                self.log.append(f'table {t["name"]} merged into {same}: names differ only in case (one table in Power BI)')
                continue
            tables[t['name']] = {'name': t['name'], 'source': src, 'columns': [dict(c) for c in t['columns']],
                                 'steps': [], 'calc_columns': [], 'measures': []}
        to_model = lambda src, _tm=to_model: self.moved_tables.get(_tm(src), _tm(src))
        # relationship keys (loaded as hidden columns when Tableau did not expose them)
        self.relationships = []
        self.m2m = set()           # (many table, other table): many-to-many relationships (no one side)
        self.calc_key_rels = []    # relationships on calculated columns, checked once the calc columns exist
        for r in rels_spec:
            if r.get('unsupported'):
                self.log.append(f'DECISION NEEDED: Tableau relationship {r.get("from")} - {r.get("to")} not made: {r["unsupported"]}')
                continue
            pairs = re.findall(r'\[([^\]]+)\]\.\[([^\]]+)\]\s*=\s*\[([^\]]+)\]\.\[([^\]]+)\]', r.get('on') or '')
            if not pairs:
                self.log.append(f'relationship not parsed: {r.get("on")}')
                continue
            a_src, a_col, b_src, b_col = pairs[0]
            a, b = to_model(a_src), to_model(b_src)
            if a not in tables or b not in tables:
                self.log.append(f'relationship skipped (table not in model): {r["on"]}')
                continue
            if a == b:
                continue
            if not r.get('guessed') and (len(pairs) > 1 or r.get('one_side') or r.get('calc_keys') or r.get('tableau_on')):
                self.logical_relationship(r, pairs, tables, to_model)
                continue
            if r.get('guessed'):
                # no .tds: a join guessed from Tableau's column naming. Keep the model a tree (Power BI rejects a second
                # path between two tables) and say so in the log; sample data, when present, still decides the one side
                if self.connected(a, b):
                    self.log.append(f'guessed join {a}[{a_col}] -> {b}[{b_col}] not made: the tables are already connected '
                                    '(a second path would be ambiguous)')
                    continue
                if self.one_side(a_src, a_col, b_src, b_col) == 'a':
                    a, b, a_col, b_col, a_src, b_src = b, a, b_col, a_col, b_src, a_src
                ka, kb = self.key_column(tables[a], a_col), self.key_column(tables[b], b_col)
                self.relationships.append((a, ka, b, kb))
                self.log.append(f'DECISION NEEDED: relationship {a}[{ka}] -> {b}[{kb}] GUESSED (no .tds). {r.get("note") or ""}')
                continue
            ka, kb = self.key_column(tables[a], a_col), self.key_column(tables[b], b_col)
            one = self.one_side(a_src, a_col, b_src, b_col)
            if one is None:  # no sample data: the table the key is named after is the one side (EVENT_ID -> Events)
                own_a, own_b = key_owner(a_col, a), key_owner(b_col, b)
                one = 'a' if own_a and not own_b else 'b' if own_b and not own_a else None
                if one:
                    self.log.append(f'DECISION NEEDED: one side of {a}.{ka} = {b}.{kb} taken as {a if one == "a" else b} '
                                    f'(the key is named after it); sample data or the .tds would prove it')
            if one == 'b':
                self.relationships.append((a, ka, b, kb))
            elif one == 'a':
                self.relationships.append((b, kb, a, ka))
            else:
                self.relationships.append((a, ka, b, kb))
                self.log.append(f'cardinality of {a}.{ka} = {b}.{kb} not proven by sample data: assumed {b} is the one side')
        # tables with no relationship filter nothing: every FILTER('<table>', ...) folded into a measure is a no-op
        self.merge_same_sample(tables)
        if not self.relationships:
            self.infer_relationships(tables)
        self.flag_orphans(tables)
        self.apply_remaps(spec_model)
        # date table: relate the active date column through a date-only key
        dt = spec_model.get('date_table')
        self.date_key = None
        if dt:
            dcol = dt['active_relationship_to']
            hosts = [t for t in tables.values() if any(c['name'] == dcol for c in t['columns'])]
            if len(hosts) > 1 and self.relationships:  # facts related to each other: a second Date path would be ambiguous
                self.log.append(f'Date relates to {hosts[0]["name"]} only: {", ".join(h["name"] for h in hosts[1:])} also have '
                                f'[{dcol}] but reach Date through their relationships')
                hosts = hosts[:1]
            for host in hosts:  # blended sources (no fact-to-fact relationships): Date filters each of them
                c = next(c for c in host['columns'] if c['name'] == dcol)
                key = f'{dcol} (Date)'
                host['steps'].append((f'Added {key[:40]}', f'Table.AddColumn({{prev}}, {mstr(key)}, each if {mfield(c["source_column"])} = null '
                                                          f'then null else Date.From({mfield(c["source_column"])}), type nullable date)'))
                host['columns'].append({'name': key, 'source_column': key, 'datatype': 'date', 'hidden': True})
                self.relationships.append((host['name'], key, 'Date', 'Date'))
                self.date_key = self.date_key or (host['name'], key)
        # URL columns (Tableau URL actions)
        for u in spec_model.get('url_columns', []):
            host = tables.get(u['table'])
            if not host or any(c['name'] == u['name'] for c in host['columns']):
                continue  # several sheets running the same URL action: one column (a second one breaks the TMDL load)
            parts, pos, tpl = [], 0, u['template']
            for mm in re.finditer(r'<\[([^\]]+)\]>', tpl):
                if mm.start() > pos:
                    parts.append(mstr(tpl[pos:mm.start()]))
                pos = mm.end()
                if mm.group(1).startswith('Param:'):  # a column cannot follow a parameter: its default value
                    pname = mm.group(1)[6:]
                    prm = next((x for x in self.digest.get('parameters', []) if x.get('name') == pname), {})
                    val = str(prm.get('value') if prm.get('value') is not None else prm.get('default') or '').strip('"#')
                    parts.append(mstr(val))
                    self.log.append(f'DECISION NEEDED: URL action "{u["name"]}" reads parameter [{pname}]: the link uses its '
                                    f'default "{val}"; for a link that follows the parameter, make the URL a measure')
                    continue
                fc = next((c for c in host['columns'] if c['name'] == mm.group(1)), None)
                parts.append(f'Text.From({mfield(fc["source_column"] if fc else mm.group(1))})')
            if pos < len(tpl):
                parts.append(mstr(tpl[pos:]))
            fields = [next((c['source_column'] for c in host['columns'] if c['name'] == f), f) for f in u['fields']]
            guard = ' or '.join(f'{mfield(f)} = null' for f in fields) or 'false'
            host['steps'].append((f'Added {u["name"][:40]}', f'Table.AddColumn({{prev}}, {mstr(u["name"])}, each if {guard} then null '
                                                             f'else {" & ".join(parts)}, type nullable text)'))
            host['columns'].append({'name': u['name'], 'source_column': u['name'], 'datatype': 'string', 'hidden': True,
                                    'dataCategory': 'WebUrl'})
        # groups -> Power Query conditional columns; other row-level calcs -> DAX calculated columns
        for t in spec_model['tables']:
            for cc in t['calculated_columns']:
                host = tables.get(self.moved_tables.get(t['name'], t['name']))
                if not host:
                    continue
                g = next((g for g in self.digest['groups'] if g['name'] == cc['name']), None)
                if g:
                    src = next((c for c in host['columns'] if c['name'] == g['source']), None)
                    if not src:
                        self.log.append(f'group {g["name"]}: source column {g["source"]} not in {host["name"]}')
                        continue
                    pairs = ', '.join(f'{mid(v.strip(chr(34)))} = {mstr(k.strip(chr(34)))}' for k, vs in g['groups'].items() for v in vs)
                    default = mstr((g.get('default') or '"Other"').strip('"'))
                    host['steps'].append((f'Added {g["name"][:40]}',
                                          f'Table.AddColumn({{prev}}, {mstr(g["name"])}, each if {mfield(src["source_column"])} = null then null '
                                          f'else Record.FieldOrDefault([{pairs}], Text.From({mfield(src["source_column"])}), {default}), type nullable text)'))
                    host['columns'].append({'name': g['name'], 'source_column': g['name'], 'datatype': 'string'})
                    continue
                calc = next((c for c in self.digest['calculations'] if c['name'] == cc['name']), None)
                if calc and any('cannot react to slicers' in n for n in calc.get('dax_notes', [])):
                    self.log.append(f'calc column "{cc["name"]}" reads a parameter: not created; its filter is folded into the measures')
                    self.param_rowcalcs.append((t['name'], cc))  # measures that read it get its formula inlined (load_guard)
                    continue
                host['calc_columns'].append({'name': cc['name'], 'dax': cc['dax'],
                                             'datatype': (calc or {}).get('datatype') or 'string'})
        for t, k in getattr(self, 'calc_key_rels', []):  # a relationship on a calc column the build did not create
            if not any(c['name'] == k for c in tables[t]['calc_columns']):
                self.relationships = [x for x in self.relationships if not ((x[0], x[1]) == (t, k) or (x[2], x[3]) == (t, k))]
                self.log.append(f'DECISION NEEDED: relationship on calculated field {t}[{k}] dropped: the calc is not a '
                                'calculated column in this model; add it as a Power Query column')
        self.manual_sort_columns(tables)
        if dt and not self.date_key:
            # the active date is a Tableau calculated field (e.g. "Application Recorded Date" = IF ... THEN [Received]
            # ELSE [Paid]): Power Query cannot see it, so the date-only key is a DAX column on its host. Without it the
            # Date table relates to nothing and every Date slicer / axis filters nothing.
            dcol = dt['active_relationship_to']
            host = next((t for t in tables.values() if any(c['name'] == dcol for c in t['calc_columns'])), None)
            if host:
                key = f'{dcol} (Date)'
                ref = f'{tq(host["name"])}[{dcol}]'
                host['calc_columns'].append({'name': key, 'datatype': 'date', 'hidden': True,
                                             'dax': f'IF(ISBLANK({ref}), BLANK(), DATE(YEAR({ref}), MONTH({ref}), DAY({ref})))'})
                self.relationships.append((host['name'], key, 'Date', 'Date'))
                self.date_key = (host['name'], key)
                self.log.append(f'Date relates to {host["name"]}[{key}]: a DAX date key, because [{dcol}] is a calculated column')
            else:
                self.log.append(f'DECISION NEEDED: no table holds the date column [{dcol}]: the Date table relates to nothing, '
                                'so Date slicers and axes filter nothing; relate it by hand')
        self.tables = tables
        self.relations = relations
        renamed = self.keep_source_names(tables)
        if self.place_calc_columns(tables) or renamed:
            self.apply_remaps(spec_model)
        new = self.unique_column_names(tables)
        if new:  # only the new renames: the earlier ones are already applied (applying them twice could chain)
            saved = self.moved_tables, self.moved_cols
            self.moved_tables, self.moved_cols = {}, new
            self.apply_remaps(spec_model)
            self.moved_tables, self.moved_cols = saved[0], {**saved[1], **new}
        self.blend_dimensions(tables)
        self.build_measures()
        self.load_guard()
        self.filter_directions()
        self.datasource_filters()
        self.value_aliases()

    def value_aliases(self):
        """Tableau value aliases (a member shown under another name: "Anna Andreadi" as "Devi A.") -> a Power Query
        display column '<column> (alias)' beside the source column. Visual groupings (axis, legend, rows / columns)
        show it; filters, slicers and DAX keep the source column, whose raw values Tableau's own filters and formulas
        compare against. It sorts like the source column when each alias stands for one value."""
        amap = {}
        for c in self.digest.get('columns', []):
            if c.get('aliases') and c.get('table'):
                amap[self.remap(c['table'], c['name'])] = c['aliases']
        if not amap:
            return
        remap = {}
        for (tn, cn), aliases in amap.items():
            t = self.tables.get(tn)
            col = next((c for c in (t or {}).get('columns', []) if cn in (c['name'], c.get('caption'))), None)
            if col is None or not col.get('source_column') or col.get('derived'):
                continue
            pairs = {k.strip('"').replace('""', '"'): v for k, v in aliases.items()}
            name = f'{col["name"]} (alias)'
            if any(c['name'].lower() == name.lower() for c in t['columns'] + t['calc_columns']):
                continue
            src = col['source_column']
            rec = ', '.join(f'{mid(k)} = {mstr(v)}' for k, v in pairs.items())
            t['steps'].append((f'Added {name[:40]}', f'Table.AddColumn({{prev}}, {mstr(name)}, each if {mfield(src)} = null then null '
                                                      f'else Record.FieldOrDefault([{rec}], Text.From({mfield(src)}), Text.From({mfield(src)})), type nullable text)'))
            new = {'name': name, 'source_column': name, 'datatype': 'string', 'derived': True}
            if len(set(pairs.values())) == len(pairs) and not set(pairs.values()) & (set(pairs) - set(pairs.values())):
                new['sort_by'] = col.get('sort_by') or col['name']  # one alias per value: sorts like the source column
            t['columns'].append(new)
            if tn in getattr(self, 'model_cols', {}):  # the report guard's list of written columns (made earlier)
                self.model_cols[tn].add(name)
            remap[(tn, col['name'])] = (tn, name)
            self.log.append(f'{tn}[{col["name"]}]: {len(pairs)} Tableau value alias(es) -> display column [{name}] on visual '
                            'axes / legends / rows; filters and slicers keep the source values')
        grouping = ('Category', 'Series', 'Rows', 'Columns', 'Legend', 'Axis', 'Values', 'Details')
        for p in self.spec['pages']:
            for v in p.get('visuals', []):
                for role, fl in (v.get('roles') or {}).items():
                    if role not in grouping:
                        continue
                    for f in fl:
                        hit = remap.get((f.get('table'), f.get('name'))) if f.get('kind') == 'Column' else None
                        if hit:
                            f['table'], f['name'] = hit

    @staticmethod
    def m_literal(v, dtype):
        """Tableau filter value ('East', 750, #2021-09-01#, true) -> Power Query literal of the column's type."""
        s = str(v).strip()
        m = re.fullmatch(r'#?(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}):(\d{2}))?#?', s)
        if m and dtype in ('date', 'datetime'):
            y, mo, d, h, mi, se = (int(x) if x else 0 for x in m.groups())
            return f'#date({y}, {mo}, {d})' if dtype == 'date' else f'#datetime({y}, {mo}, {d}, {h}, {mi}, {se})'
        if dtype in ('integer', 'real') and re.fullmatch(r'-?\d+(\.\d+)?', s):
            return s
        if dtype == 'boolean' and s.lower() in ('true', 'false'):
            return s.lower()
        return mstr(s.strip('"'))

    def datasource_filters(self):
        """Tableau data source filters: rows the whole data source never shows. A source column -> Power Query row
        filter on its table (the rows never load). When that table filters others (the one side of a relationship) or
        the field is not a source column (calc column, relative date), the filter also goes on every page, so related
        tables follow it as Tableau's joined rows did. Aggregate (measure) filters: DECISION NEEDED."""
        ones = {b for _, _, b, _ in self.relationships}
        for f in self.spec['model'].get('datasource_filters') or []:
            fr, kind = f.get('field') or {}, f.get('type')
            where = f'data source filter [{fr.get("name")}] ({f.get("datasource")}: {f.get("tableau")})'
            if kind in ('All', None):
                continue
            if fr.get('kind') != 'Column' or not fr.get('table'):
                self.log.append(f'DECISION NEEDED: {where}: not a column (aggregate or unresolved): apply it by hand')
                continue
            tn, cn = self.remap(fr['table'], fr['name'])  # merged tables / kept source names
            fr = {**fr, 'table': tn, 'name': cn}
            f = {**f, 'field': fr}
            t = self.tables.get(tn)
            col = next((c for c in (t or {}).get('columns', []) if c['name'] == cn), None)
            cond = None
            if col is not None and col.get('source_column') and not col.get('derived'):
                x, dt = mfield(col['source_column']), col.get('datatype')
                if kind == 'Categorical' and f.get('values'):
                    cond = f'List.Contains({{{", ".join(self.m_literal(v, dt) for v in f["values"])}}}, {x})'
                    cond = f'not {cond}' if f.get('inverted') else cond
                elif kind == 'Range' and (f.get('lo') is not None or f.get('hi') is not None):
                    parts = ([f'{x} >= {self.m_literal(f["lo"], dt)}'] if f.get('lo') is not None else []) + \
                            ([f'{x} <= {self.m_literal(f["hi"], dt)}'] if f.get('hi') is not None else [])
                    cond = ' and '.join([f'{x} <> null'] + parts)
                elif kind == 'NotBlank':
                    cond = f'{x} <> null'
                elif kind == 'Blank':
                    cond = f'{x} = null'
            if cond:
                t['steps'].append((f'Data source filter {col["source_column"][:30]}', f'Table.SelectRows({{prev}}, each {cond})'))
                self.log.append(f'{where}: Power Query row filter on {t["name"]}')
            if not cond or t['name'] in ones:
                for p in self.spec['pages']:
                    p.setdefault('page_filters', []).append({k: v for k, v in f.items() if k != 'datasource'})
                self.log.append(f'{where}: filter on every page' + (f' ({t["name"]} filters related tables)' if cond else ''))

    # ---------- data blending ----------
    def blend_column(self, table, name, source, dtype):
        """Model column of a blend link: by name (after renames), else by source column; added hidden when no visual
        used it (a secondary's link field is usually not on a shelf)."""
        c = next((c for c in table['columns'] if c['name'] == name), None) or \
            next((c for c in table['columns'] if c.get('source_column') == source), None)
        if c is None and source:
            c = {'name': source, 'source_column': source, 'datatype': dtype or 'string', 'hidden': True}
            table['columns'].append(c)
        return c

    def blend_dimensions(self, tables):
        """Tableau data blending -> Power BI shared dimensions. Tableau aggregates the secondary source to the linked
        fields of the primary; without a relationship Power BI shows the secondary's grand total on every mark. For each
        active link (spec model "blend_links", from the extractor):
          - a date link (Order Date = Month Start at month level) relates the secondary's date to the Date table, so a
            Date column (Year Month, Year ...) filters both sources at the level the visual shows;
          - any other link becomes a shared dimension table (Power Query: the distinct keys of every linked column),
            related one-to-many to each blended table. Links of the same field across sheets share one dimension.
        Visuals, slicers and filters on a linked column are pointed at the shared dimension (so they filter both
        sources), the linked fact columns are hidden. Not made when the tables are already related (a second path
        would be ambiguous) or the link is a calculated field / mixes data types: DECISION NEEDED in the log."""
        links = self.spec['model'].get('blend_links') or []
        if not links:
            return
        has_date = bool(self.spec['model'].get('date_table'))
        root, node_col, node_name, order = {}, {}, {}, []  # union-find over (table, column) link ends
        self.blend_linked, self.blend_date_warned, calc_links = set(), set(), []

        def find(x):
            while root.get(x, x) != x:
                x = root[x]
            return x
        for lk in links:
            where = f'blend link {lk["primary_field"]} = {lk["secondary_field"]} ({", ".join(lk.get("sheets", []))})'
            if lk.get('calc'):
                if lk['primary_field'] not in calc_links:
                    calc_links.append(lk['primary_field'])
                continue
            ends = []
            for side in ('primary', 'secondary'):
                t, c = self.remap(lk[f'{side}_table'] or '', lk[f'{side}_field'])
                col = self.blend_column(tables[t], c, lk.get(f'{side}_column'), lk.get('datatype') if side == 'primary'
                                        else lk.get('secondary_datatype')) if t in tables else None
                if not col:
                    break
                ends.append((t, col))
            if len(ends) < 2:
                self.log.append(f'DECISION NEEDED: {where}: a linked table or column is not in the model; relate by hand')
                continue
            (ta, ca), (tb, cb) = ends
            if ta == tb:
                continue  # both sources read the same table: it is one table in Power BI already
            dates = {ca['datatype'], cb['datatype']} <= {'date', 'datetime'}
            if ca['datatype'] != cb['datatype'] and not dates:
                self.log.append(f'DECISION NEEDED: {where} links {ca["datatype"]} to {cb["datatype"]}: no relationship '
                                '(Power BI keys need one data type); convert one of them in Power Query')
                continue
            if dates and has_date:
                for t, c in ends:
                    self.blend_date_link(tables[t], c, where, lk.get('level'))
                continue
            a, b = (ta, ca['name']), (tb, cb['name'])
            for x, c in ((a, ca), (b, cb)):
                if x not in node_col:
                    node_col[x] = c
                    order.append(x)
            node_name.setdefault(a, lk['primary_field'])  # the dimension is named after the primary's field
            ra, rb = find(a), find(b)
            if ra != rb:
                root[rb] = ra
        if calc_links:
            self.log.append(f'DECISION NEEDED: blend links on calculated fields ({", ".join(calc_links)}): no shared dimension '
                            'made (Power Query cannot see a DAX calc). A filter on such a field reaches only its own source; '
                            'where the secondary must follow it, add the calc as a Power Query column on both tables and '
                            'relate them through a shared dimension')
        groups = {}
        for x in order:  # first-seen order: the first primary end names the dimension
            groups.setdefault(find(x), []).append(x)
        taken = {n.lower() for n in tables} | {'date'} | {(p.get('name') or '').lower() for p in self.spec['model'].get('parameters', [])}
        self.blend_remap = {}
        for ends in groups.values():
            fact_tables = sorted({t for t, _ in ends})
            # a new dimension filters each linked table and what those filter in turn (one side -> many side). Two
            # linked tables that reach a common table (one already filters the other) would give it two filter paths:
            # Power BI calls that ambiguous. A common one side (Date over both) is fine: filters do not flow up.
            down = {t: self.downstream(t) for t in fact_tables}
            clash = [(x, y) for i, x in enumerate(fact_tables) for y in fact_tables[i + 1:] if down[x] & down[y]]
            if clash:
                self.log.append(f'DECISION NEEDED: blend link on {", ".join(f"{t}[{c}]" for t, c in ends)}: '
                                f'{clash[0][0]} and {clash[0][1]} both filter {", ".join(sorted(down[clash[0][0]] & down[clash[0][1]]))}; '
                                'a shared dimension would make two filter paths (ambiguous); no dimension made')
                continue
            field = next((node_name[x] for x in ends if x in node_name), ends[0][1])
            dname = field if field.lower() not in taken else f'{field} (blend)'
            k = 2
            while dname.lower() in taken:
                dname, k = f'{field} (blend {k})', k + 1
            taken.add(dname.lower())
            dcol = field
            first = node_col[ends[0]]
            text = first['datatype'] in ('string', None)
            steps, parts = [], []
            for t, cn in ends:
                c = node_col[(t, cn)]
                step = f'{t} {cn} keys'
                expr = f'Table.SelectColumns({mid(t)}, {{{mstr(c["source_column"])}}})'
                if c['source_column'] != dcol:
                    expr = f'Table.RenameColumns({expr}, {{{{{mstr(c["source_column"])}, {mstr(dcol)}}}}})'
                steps.append((step, expr))
                parts.append(f'#{mstr(step)}')
            steps.append(('Combined', 'Table.Combine({' + ', '.join(parts) + '})'))
            if text:
                # Power BI keys ignore case and trailing spaces: "East" and "east " are one key on the one side
                steps.append(('Trimmed', f'Table.TransformColumns(#"Combined", {{{{{mstr(dcol)}, each if _ = null then null '
                                         f'else Text.TrimEnd(Text.From(_)), type nullable text}}}})'))
            steps.append(('Removed blanks', f'Table.SelectRows(#{mstr(steps[-1][0])}, each {mfield(dcol)} <> null)'))
            steps.append(('Kept one row per key', f'Table.Distinct(#"Removed blanks", {{{mstr(dcol)}'
                                                  + (', Comparer.OrdinalIgnoreCase' if text else '') + '})'))
            lines = ['let'] + [f'    #{mstr(s)} = {e}' + (',' if i < len(steps) - 1 else '') for i, (s, e) in enumerate(steps)] \
                + ['in', f'    #{mstr(steps[-1][0])}']
            tables[dname] = {'name': dname, 'source': None, 'blend_dim': True, 'm_lines': lines, 'steps': [], 'calc_columns': [],
                             'measures': [], 'columns': [{'name': dcol, 'source_column': dcol, 'datatype': first['datatype'],
                                                          'summarizeBy': 'none'}]}
            for t, cn in ends:
                c = node_col[(t, cn)]
                c['hidden'] = True  # report authors use the shared dimension
                self.relationships.append((t, c['name'], dname, dcol))
                self.blend_remap[(t, c['name'])] = (dname, dcol)
                self.blend_linked.add(t)
            self.log.append(f'blend: shared dimension {dname}[{dcol}] (distinct keys of {", ".join(f"{t}[{c}]" for t, c in ends)}) '
                            f'related to {", ".join(fact_tables)}; visuals on the linked column use it')
        if self.blend_remap:  # report only: DAX keeps the fact columns (row context needs them)
            saved = self.moved_tables, self.moved_cols
            self.moved_tables, self.moved_cols = {}, self.blend_remap
            self.remap_spec(self.spec['pages'])
            self.moved_tables, self.moved_cols = saved
        # blended tables are related now (through the shared dimension or Date): their orphan warnings no longer apply
        self.log = [x for x in self.log if not (isinstance(x, str) and x.startswith('WARNING ORPHAN TABLE:')
                                                and x.split(':', 1)[1].strip().split(' ')[0] in self.blend_linked)]

    def downstream(self, x):
        """x and every table x filters through single-direction relationships (one side -> many side)."""
        seen, todo = {x}, [x]
        while todo:
            n = todo.pop()
            for a, _, b, _ in self.relationships:
                if b == n and a not in seen:
                    seen.add(a)
                    todo.append(a)
        return seen

    def blend_date_link(self, table, c, where, level):
        """A date blend link: the table's date relates to the Date table (date-only key), as the primary's date does."""
        # the date step names the key after the Tableau caption ("Start Dt (Date)"); keep_source_names may since have
        # renamed the column itself to the source name (StartDt), so both spellings are the same key
        keys = {f'{c["name"]} (Date)', f'{c.get("caption") or c["name"]} (Date)'}
        key = f'{c.get("caption") or c["name"]} (Date)'
        cur = [(a, ka) for a, ka, b, _ in self.relationships if a == table['name'] and b == 'Date']
        if cur:
            if cur[0][1] in keys:
                self.blend_linked.add(table['name'])
            elif (table['name'], c['name']) not in self.blend_date_warned:
                self.blend_date_warned.add((table['name'], c['name']))
                self.log.append(f'DECISION NEEDED: {where}: {table["name"]} already relates to Date through [{cur[0][1]}], not '
                                f'[{c["name"]}]; the {level or "date"} link needs a second (inactive) Date relationship')
            return
        if not any(x['name'] == key for x in table['columns']):
            src = c.get('source_column') or c['name']
            table['steps'].append((f'Added {key[:40]}', f'Table.AddColumn({{prev}}, {mstr(key)}, each if {mfield(src)} = null '
                                                         f'then null else Date.From({mfield(src)}), type nullable date)'))
            table['columns'].append({'name': key, 'source_column': key, 'datatype': 'date', 'hidden': True})
        self.relationships.append((table['name'], key, 'Date', 'Date'))
        self.blend_linked.add(table['name'])
        self.log.append(f'{where}: {table["name"]}[{c["name"]}] related to Date (the {level or "date"} link: Date columns filter '
                        'both sources' + (', e.g. Year Month' if level == 'month' else '') + ')')

    def manual_sort_columns(self, tables):
        """Tableau manual sort (a hand-ordered member list) -> a hidden '<column> Sort' column (Power Query: the
        member's position in the list, unknown members last) and sortByColumn on the column. Text columns read by
        Power Query only (source and group columns): a DAX sort column derived from the column it sorts can make a
        circular dependency. When two sheets order a column differently, the first order wins (one sort per column)."""
        done = set()
        for ws in self.digest.get('worksheets', []):
            for label, order in (ws.get('manual_sorts') or {}).items():
                vals = [v for v in order if v and not v.startswith('%')]
                if len(vals) < 2:
                    continue
                for t in tables.values():
                    cc = next((c for c in t['calc_columns'] if c['name'] == label and c.get('dax') and 'BLANK()' != c['dax']
                               and (c.get('datatype') or 'string') == 'string'), None)
                    if cc and (t['name'], label) not in done:
                        # a calculated column: the sort column evaluates the same formula (it reads the base columns, not
                        # the sorted column, so there is no circular dependency)
                        done.add((t['name'], label))
                        sname = f'{label} Sort'
                        if not any(x['name'] == sname for x in t['columns'] + t['calc_columns']):
                            cases = ', '.join(f'{mstr(v)}, {i + 1}' for i, v in enumerate(vals))
                            t['calc_columns'].append({'name': sname, 'datatype': 'integer', 'hidden': True,
                                                      'dax': f'SWITCH({cc["dax"]}, {cases}, {len(vals) + 1})'})
                            cc['sort_by'], cc['sort_order'] = sname, vals
                            self.log.append(f'{t["name"]}[{label}] sorted by hand in Tableau ("{ws["name"]}"): sort column {sname}')
                        continue
                    c = next((c for c in t['columns'] if label in (c['name'], c.get('caption'))), None)
                    if not c or c.get('derived') or c.get('key') or (c.get('datatype') or 'string') != 'string':
                        continue
                    if (t['name'], c['name']) in done:
                        if c.get('sort_order') != vals:
                            self.log.append(f'{t["name"]}[{c["name"]}]: sheet "{ws["name"]}" orders it differently by hand; '
                                            'the first manual order is used (Power BI has one sort per column)')
                        continue
                    done.add((t['name'], c['name']))
                    sname = f'{c["name"]} Sort'
                    if any(x['name'] == sname for x in t['columns'] + t['calc_columns']):
                        continue
                    lst = '{' + ', '.join(mstr(v) for v in vals) + '}'
                    t['steps'].append((f'Added {sname[:40]}', f'Table.AddColumn({{prev}}, {mstr(sname)}, each let p = '
                                       f'List.PositionOf({lst}, {mfield(c["source_column"])}) in if p < 0 then {len(vals) + 1} '
                                       f'else p + 1, Int64.Type)'))
                    t['columns'].append({'name': sname, 'source_column': sname, 'datatype': 'integer', 'hidden': True,
                                         'derived': True, 'summarizeBy': 'none'})
                    c['sort_by'], c['sort_order'] = sname, vals
                    self.log.append(f'{t["name"]}[{c["name"]}] sorted by hand in Tableau ("{ws["name"]}"): sort column {sname}')

    def reads_rows(self, name, seen=None):
        """True when a measure (or a measure it calls) names a model table: 'T'[col], COUNTROWS('T'), FILTER('T', ...).
        A parameter value measure or a text constant names none."""
        seen = seen if seen is not None else set()
        if name in seen or name not in self.measures:
            return False
        seen.add(name)
        dax = self.measures[name]['dax'] or ''
        if {t.replace("''", "'") for t in re.findall(r"'((?:[^']|'')+)'", dax)} & set(self.tables):
            return True
        return any(self.reads_rows(r, seen) for r in re.findall(r"(?<!['\]\w])\[([^\]]+)\]", dax))

    def measure_tables(self, name, seen=None):
        """Tables whose columns a measure aggregates (through the measures it calls)."""
        seen = seen if seen is not None else set()
        if name in seen or name not in self.measures:
            return set()
        seen.add(name)
        dax = self.measures[name]['dax'] or ''
        out = {t.replace("''", "'") for t in re.findall(r"'((?:[^']|'')+)'\[", dax)}
        for ref in re.findall(r"(?<!')\[([^\]]+)\]", re.sub(r"'(?:[^']|'')+'\[[^\]]+\]", '', dax)):
            out |= self.measure_tables(ref, seen)
        return out & set(self.tables)

    def filter_directions(self):
        """Tableau's published sources are one joined row set: a sheet can count work orders per task status. In
        Power BI a single-direction relationship filters only from the one side (WORK_ORDERS) to the many side
        (TASKS), so the same visual shows one total on every row. For each relationship a visual needs the other
        way, filter both directions, only while the model has no loop (both directions on a tree is never ambiguous)."""
        self.both = set()
        rels = list(self.relationships)
        root = {}

        def find(x):
            while root.get(x, x) != x:
                x = root[x]
            return x
        for a, _, b, _ in rels:
            ra, rb = find(a), find(b)
            if ra == rb:  # a loop (or two relationships between the same tables): leave every direction alone
                return
            root[ra] = rb
        up = {}  # many -> [(one, relationship)]
        for r in rels:
            up.setdefault(r[0], []).append((r[2], r))

        def path_up(src, dst):
            """Relationships from src (many side) up to dst (one side), or None."""
            todo, prev = [src], {src: None}
            while todo:
                n = todo.pop()
                if n == dst:
                    out = []
                    while prev[n]:
                        n, r = prev[n]
                        out.append(r)
                    return out
                for one, r in up.get(n, ()):
                    if one not in prev:
                        prev[one] = (n, r)
                        todo.append(one)
            return None
        for p in self.spec['pages']:
            for v in p['visuals']:
                fs = [f for fl in v.get('roles', {}).values() for f in fl]
                cols = {f['table'] for f in fs if f.get('kind') == 'Column' and f.get('table') in self.tables}
                meas = set().union(*[self.measure_tables(self.measure_for(v, f['name'])) for f in fs if f.get('kind') == 'Measure'] or [set()])
                for c in cols:
                    for m in meas - {c}:
                        need = path_up(c, m)  # the column sits on the many side of the measure's table
                        for r in need or ():
                            if r not in self.both:
                                self.both.add(r)
                                self.log.append(f'DECISION NEEDED: relationship {r[0]}[{r[1]}] -> {r[2]}[{r[3]}] filters both '
                                                f'directions: "{v.get("sheet")}" groups a measure of {m} by a column of {c} '
                                                '(Tableau\'s joined rows). Check the totals, or count the key on the many side instead')

    # ---------- load guard: Desktop refuses the whole PBIP on some DAX errors, so never write them ----------
    SECURITY_FN = re.compile(r'\b(USERNAME|USERPRINCIPALNAME|CUSTOMDATA|USERCULTURE)\s*\(\s*\)')

    def preview_login(self):
        """Login that stands in for USERNAME() in calculated columns (decisions.json "preview_login", else the
        first hard-coded admin login of the workbook)."""
        if self.decisions.get('preview_login'):
            return self.decisions['preview_login']
        for f in self.digest.get('flags', []):
            if f['code'] == 'HARDCODED_USERS':
                m = re.search(r'logins ([^\s,:]+)', f['detail'])
                if m:
                    return m.group(1)
        return ''

    def load_guard(self):
        cols = {t['name']: {c['name'] for c in t['columns']} | {c['name'] for c in t['calc_columns']}
                for t in self.tables.values()}
        cols['Date'] = {'Date', 'Year', 'Quarter', 'Month Number', 'Month', 'Year Month', 'Day'} | {c for c, _ in DATE_EXTRA}
        if self.report_period:
            cols[self.report_period] = {'Date'}
        for tname in self.slice_tables:
            cols[tname] = {'Slice', 'Order'}
        for tname in getattr(self, 'pdim_tables', {}):
            cols[tname] = {tname, 'Order'}
        for tname, (_, col, _) in self.param_tables.items():
            cols[tname] = {col}
        written = set(cols)
        self.model_cols = cols
        fact_tables = [t for t in self.tables]
        for m in self.measures.values():
            if m['table'] not in written:  # a measure on a table that is never written is silently lost
                self.log.append(f'measure "{m["name"]}": home table {m["table"]} is not in the model; moved to {fact_tables[0]}')
                m['table'] = fact_tables[0]
        login = self.preview_login()
        # 1. security functions are not allowed in calculated columns (PFE_XL_DYNAMIC_SECURITY_FUNCTION_NOT_ALLOWED)
        for t in self.tables.values():
            for c in t['calc_columns']:
                if self.SECURITY_FN.search(c['dax']):
                    c['dax'] = self.SECURITY_FN.sub('"' + login.replace('"', '""') + '"', c['dax'])
                    self.log.append(f'DECISION NEEDED: calc column {t["name"]}[{c["name"]}] is per viewer (USERNAME()); '
                                    f'built for preview login "{login}" (decisions.json "preview_login"). Production needs RLS / a user table')

        # 2. references that do not resolve, and visual-calculation-only syntax, break the model load
        def problems(dax, home, is_measure):
            out = []
            if '/* visual calc' in dax or '<Table>' in dax or '<Dimension>' in dax:
                out.append('needs a visual calculation or a placeholder was left in')
            for tbl, col in re.findall(r"'((?:[^']|'')+)'\[([^\]]+)\]", dax):
                tbl = tbl.replace("''", "'")
                if tbl not in cols:
                    out.append(f"table '{tbl}' not in the model")
                elif col not in cols[tbl] and col not in self.measures:
                    out.append(f"'{tbl}'[{col}] not in the model")
            for ref in re.findall(r"(?<!['\]\w])\[([^@\]][^\]]*)\]", dax):  # [@x] = ADDCOLUMNS extension column
                if ref not in self.measures and not (not is_measure and ref in cols.get(home, ())):
                    out.append(f'[{ref}] is not a measure')
            if re.search(r"SELECTEDVALUE\(\s*(?:\[|\()", dax):
                out.append('SELECTEDVALUE of a measure or expression (needs a column)')
            for args in summarize_args(dax):  # SUMMARIZE(table, groupBy columns..., "name", expr...)
                for a in args[1:]:
                    if a.startswith('"'):
                        break
                    if not re.fullmatch(r"'(?:[^']|'')+'\[[^\]]+\]|\[@[^\]]+\]", a):
                        out.append(f'SUMMARIZE groups by an expression ({a[:60]}); it takes columns only')
            return sorted(set(out))

        def summarize_args(dax):
            for mm in re.finditer(r'\bSUMMARIZE\(', dax):
                depth, cur, args, i, quote = 0, '', [], mm.end(), None
                while i < len(dax):
                    ch = dax[i]
                    if quote:
                        quote = None if ch == quote else quote
                    elif ch in '"\'':
                        quote = ch
                    elif ch == '(':
                        depth += 1
                    elif ch == ')':
                        if depth == 0:
                            break
                        depth -= 1
                    elif ch == ',' and depth == 0:
                        args.append(cur.strip())
                        cur = ''
                        i += 1
                        continue
                    cur += ch
                    i += 1
                yield args + [cur.strip()]
        # a row-level calc that reads a parameter is not a column (a column cannot react to a slicer): measures that read
        # it inside an iterator get its formula instead, which is evaluated per row like Tableau does
        inline = {}
        for tname, cc in getattr(self, 'param_rowcalcs', []):
            host = self.moved_tables.get(tname, tname)
            if cc.get('dax') and cc['name'] not in cols.get(host, ()):
                inline[(host, cc['name'])] = cc['dax']
        if inline:
            pat = re.compile(r"'((?:[^']|'')+)'\[([^\]]+)\]")
            key = lambda mm: (mm.group(1).replace("''", "'"), mm.group(2))
            # a calc column that reads such a calc reads the parameter too: not a column either, inlined the same way
            changed = True
            while changed:
                changed = False
                for t in self.tables.values():
                    for c in list(t['calc_columns']):
                        if any(key(mm) in inline for mm in pat.finditer(c['dax'])):
                            dax = c['dax']
                            for _ in range(4):
                                dax = pat.sub(lambda mm: f'({inline[key(mm)]})' if key(mm) in inline else mm.group(0), dax)
                            inline[(t['name'], c['name'])] = dax
                            t['calc_columns'].remove(c)
                            cols.get(t['name'], set()).discard(c['name'])
                            self.log.append(f'calc column "{c["name"]}" reads a parameter through another calc: not created; '
                                            'its formula is inlined into the measures that use it')
                            changed = True
            # functions that take a column only get the expression form over the calc's table
            col_only = re.compile(r"\b(SELECTEDVALUE|DISTINCTCOUNTNOBLANK|DISTINCTCOUNT)\(\s*'((?:[^']|'')+)'\[([^\]]+)\]\s*\)")

            def expr_form(mm):
                fn, t, c = mm.group(1), mm.group(2).replace("''", "'"), mm.group(3)
                if (t, c) not in inline:
                    return mm.group(0)
                rows = f'DISTINCT(SELECTCOLUMNS({tq(t)}, "@x", {inline[(t, c)]}))'
                if fn == 'SELECTEDVALUE':  # ATTR(): the one value, blank when several
                    return f'(VAR _one = {rows} RETURN IF(COUNTROWS(_one) = 1, MAXX(_one, [@x])))'
                return f'COUNTROWS(FILTER({rows}, NOT ISBLANK([@x])))'
            for m in self.measures.values():
                before = m['dax']
                m['dax'] = col_only.sub(expr_form, m['dax'])
                for _ in range(4):  # a parameter calc can read another one
                    m['dax'] = pat.sub(lambda mm: f'({inline[(mm.group(1).replace(chr(39) * 2, chr(39)), mm.group(2))]})'
                                       if (mm.group(1).replace(chr(39) * 2, chr(39)), mm.group(2)) in inline else mm.group(0), m['dax'])
                if m['dax'] != before:
                    m.setdefault('notes', []).append('parameter-reading row calc inlined (per row, reacts to the slicer)')
        bad = {}
        for m in self.measures.values():
            p = problems(m['dax'], m['table'], True)
            if any('not valid DAX as written' in n for n in m.get('notes', [])):
                p = sorted(set(p) | {'mixes row-level columns with measures (translator: not valid DAX as written)'})
            todo = re.findall(r'/\* TODO ?([^*]*)\*/', m['dax'])
            if todo:  # a Tableau function with no DAX equivalent was written as BLANK(): the number looks real but is not
                p = sorted(set(p) | {'not translated: ' + ', '.join(sorted({x.strip() or 'function' for x in todo}))})
            if p:
                bad[m['name']] = p
        # a measure that reads itself (directly or round a loop) is a circular dependency: Desktop rejects it
        deps = {n: {r for r in re.findall(r"(?<!['\]\w])\[([^\]]+)\]", m['dax']) if r in self.measures}
                for n, m in self.measures.items()}
        for start in deps:
            path, todo_ = {start: None}, [start]
            while todo_:
                x = todo_.pop()
                for y in deps.get(x, ()):
                    if y == start:
                        chain, z = [start], x
                        while z != start:
                            chain.insert(1, z)
                            z = path[z]
                        bad.setdefault(start, []).append('circular reference: ' + ' -> '.join(f'[{c}]' for c in chain + [start]))
                        todo_ = []
                        break
                    if y not in path:
                        path[y] = x
                        todo_.append(y)
        for name, p in bad.items():
            m = self.measures[name]
            m['description'] = (f'NOT TRANSLATED ({"; ".join(p)}). Draft DAX: {" ".join(m["dax"].split())} | ' + m['description'])
            m['dax'] = 'BLANK()'
            m['displayFolder'] = 'Measures\\_Needs fix'
            self.log.append(f'measure "{name}" -> BLANK() (folder _Needs fix): {"; ".join(p)}')
        # measures that read a BLANK() stub look fine but are wrong: BLANK() + x = x, so a sum silently loses a term
        stubbed, changed = set(bad), True
        uses = {}
        while changed:
            changed = False
            for m in self.measures.values():
                if m['name'] in stubbed:
                    continue
                hit = [r for r in re.findall(r"(?<!['\]\w])\[([^\]]+)\]", m['dax']) if r in stubbed]
                if hit:
                    stubbed.add(m['name'])
                    uses[m['name']] = sorted(set(hit))
                    changed = True
        for name, hit in uses.items():
            m = self.measures[name]
            m['confidence'] = 'low'
            m['displayFolder'] = 'Measures\\_Needs fix'
            m['description'] = f'WRONG UNTIL FIXED: reads not-translated measure(s) {", ".join(hit)} (BLANK()). | ' + m['description']
            self.log.append(f'DECISION NEEDED: measure "{name}" reads not-translated {", ".join("[" + h + "]" for h in hit)} '
                            '(BLANK()): its numbers are wrong until those are fixed (folder _Needs fix)')
        self.stub_dependents = uses
        for t in self.tables.values():
            for c in t['calc_columns']:
                p = problems(c['dax'], t['name'], False)
                if p:
                    self.log.append(f'calc column {t["name"]}[{c["name"]}] -> BLANK(): {"; ".join(p)}. Draft DAX: {c["dax"]}')
                    c['dax'] = 'BLANK()'
                    continue
                todo = re.findall(r'/\* TODO ?([^*]*)\*/', c['dax'])
                if todo:  # the rest of the formula still works (e.g. a masking IF); say which part is missing
                    self.log.append(f'DECISION NEEDED: calc column {t["name"]}[{c["name"]}]: '
                                    f'{", ".join(sorted({x.strip() or "a function" for x in todo}))} has no DAX equivalent '
                                    'and is BLANK() inside it; do that part in Power Query or the source')

    # ---------- tables without relationships ----------
    def csv_header(self, relation):
        cfg = self.sources.get('tables', {}).get(relation) or {}
        path = os.path.join(self.sources.get('sample_folder') or '', cfg.get('file') or f'{relation}.csv')
        if self.sources.get('mode') != 'sample' or not os.path.exists(path):
            return []
        with open(path, encoding='utf-8-sig', newline='') as f:
            return next(csv.reader(f), [])

    def csv_name(self, t, c):
        """CSV header a model column is read from in sample mode."""
        cfg = self.sources.get('tables', {}).get(t['source']) or {}
        return next((a for a, b in self.sample_renames(t, cfg) + t.get('extra_renames', []) if b == c['source_column']),
                    c['source_column'])

    def related(self, name):
        return any(name in (a, b) for a, _, b, _ in self.relationships)

    def connected(self, x, y):
        """True when the relationships made so far already join tables x and y (in any direction)."""
        seen, todo = {x}, [x]
        while todo:
            n = todo.pop()
            for a, _, b, _ in self.relationships:
                for p, q in ((a, b), (b, a)):
                    if p == n and q not in seen:
                        seen.add(q)
                        todo.append(q)
        return y in seen

    def merge_same_sample(self, tables):
        """Sample mode: tables read from the same CSV (same where) that no relationship connects are one source row:
        their columns move onto the widest of them. Without the published .tds Tableau's join is unknown, and a table
        of one column with no key would make every filter on it silently do nothing."""
        if self.sources.get('mode') != 'sample':
            return
        groups = {}
        for t in tables.values():
            cfg = self.sources.get('tables', {}).get(t['source'])
            if cfg and cfg.get('file'):
                groups.setdefault((cfg['file'].lower(), json.dumps(cfg.get('where') or {}, sort_keys=True)), []).append(t)
        for members in groups.values():
            if len(members) < 2:
                continue
            members.sort(key=lambda t: (self.related(t['name']), len(t['columns'])), reverse=True)
            host = members[0]
            for t in members[1:]:
                if self.related(t['name']):
                    continue
                by_csv = {self.csv_name(host, c): c for c in host['columns']}
                for c in t['columns']:
                    src = self.csv_name(t, c)
                    if src in by_csv:  # the host already reads that CSV column: use its column
                        self.moved_cols[(t['name'], c['name'])] = (host['name'], by_csv[src]['name'])
                        continue
                    if any(x['name'] == c['name'] for x in host['columns']):
                        self.log.append(f'{t["name"]}[{c["name"]}] not merged into {host["name"]}: a column of that name exists')
                        continue
                    host['columns'].append(dict(c))
                    if src != c['source_column']:
                        host.setdefault('extra_renames', []).append((src, c['source_column']))
                    by_csv[src] = c
                    self.moved_cols[(t['name'], c['name'])] = (host['name'], c['name'])
                host['steps'] += t['steps']
                self.moved_tables[t['name']] = host['name']
                del tables[t['name']]
                self.log.append(f'DECISION NEEDED: table {t["name"]} has no relationship and reads the same sample CSV as '
                                f'{host["name"]}: its columns were put on {host["name"]} (same row). In the warehouse it is a '
                                'separate table: add the published .tds to input\\ so the join is known')

    def infer_relationships(self, tables):
        """No .tds: Tableau names a joined column that clashes '<name> (<TABLE>)', so a sample CSV header such as
        'WO Number (TASKS)' in TASKS and 'WO Number' in another table is the join key. Kept only when the
        sample data proves the other side unique."""
        heads = {n: self.csv_header(t['source']) for n, t in tables.items()}
        for a, ta in tables.items():
            for h in heads[a]:
                m = re.fullmatch(r'(.+) \((.+)\)', h)
                if not m or m.group(2) not in (ta['source'], a):
                    continue
                for b, tb in tables.items():
                    if b == a or m.group(1) not in heads[b] or self.related(a) and self.related(b):
                        continue
                    if self.one_side(ta['source'], h, tb['source'], m.group(1)) != 'b':
                        self.log.append(f'join {a}[{h}] = {b}[{m.group(1)}] suggested by the column names, but {b}[{m.group(1)}] '
                                        'is not unique in the sample: no relationship made')
                        continue
                    ka, kb = self.key_column(ta, h), self.key_column(tb, m.group(1))
                    self.relationships.append((a, ka, b, kb))
                    self.log.append(f'DECISION NEEDED: relationship {a}[{ka}] -> {b}[{kb}] inferred from Tableau join column '
                                    f'naming (no .tds); confirm with the published data source')

    def flag_orphans(self, tables):
        if len(tables) < 2:
            return
        for n in tables:
            if not self.related(n):
                self.log.append(f'WARNING ORPHAN TABLE: {n} has no relationship to any other table, so a filter or slicer on '
                                'it, and every measure filter folded as FILTER(' + tq(n) + ', ...), changes nothing in the '
                                'other tables. Add the published .tds to input\\ (joins) or relate it by hand')

    def remap(self, table, name):
        """(table, column) after tables were merged: where the column lives in the written model."""
        return self.moved_cols.get((table, name)) or (self.moved_tables.get(table, table), name)

    def apply_remaps(self, spec_model):
        """After tables were merged or calc columns moved: rewrite the spec and DAX that still name the old place."""
        if not (self.moved_tables or self.moved_cols):
            return
        self.remap_spec(self.spec['pages'])
        for t in spec_model['tables']:
            for m in t['measures']:
                m['table'] = self.moved_tables.get(m['table'], m['table'])
                m['dax'] = self.remap_dax(m['dax'])
            for cc in t['calculated_columns']:
                cc['dax'] = self.remap_dax(cc['dax'])
        for u in spec_model.get('url_columns', []):
            u['table'] = self.moved_tables.get(u['table'], u['table'])
        for pd in spec_model.get('parameter_dimensions', []):
            pd['host'] = self.moved_tables.get(pd['host'], pd['host'])
            pd['dax'] = self.remap_dax(pd['dax'])
        for c in self.digest['calculations']:
            c['dax'] = self.remap_dax(c.get('dax'))
        for t in getattr(self, 'tables', {}).values():
            for cc in t['calc_columns']:
                cc['dax'] = self.remap_dax(cc['dax'])

    def keep_source_names(self, tables):
        """Model columns are named as the source names them (warehouse column, or the extract header the developer
        gave), never by the Tableau caption: the tables are repointed to the warehouse later, so names must equal what
        it holds. A Tableau caption that differs survives as the visual header (projection displayName: the spec
        field keeps its label). Returns True when a column was renamed."""
        measure_names = {m['name'] for t in self.spec['model']['tables'] for m in t['measures']}
        done = False
        for t in tables.values():
            names = {c['name'] for c in t['columns']} | {c['name'] for c in t['calc_columns']}
            for c in t['columns']:
                old, new = c['name'], c.get('source_column')
                if not new or new == old or c.get('key') or new in names or old in measure_names:
                    continue
                c['name'], c['caption'] = new, old  # sources.json may still name the column by its caption
                names.discard(old)
                names.add(new)
                for k, v in list(self.moved_cols.items()):  # columns merged in earlier now live under the new name
                    if v == (t['name'], old):
                        self.moved_cols[k] = (t['name'], new)
                self.moved_cols[(t['name'], old)] = (t['name'], new)
                tn = t['name']
                self.relationships = [(a, new if (a, ka) == (tn, old) else ka, b, new if (b, kb) == (tn, old) else kb)
                                      for a, ka, b, kb in self.relationships]
                self.log.append(f'{t["name"]}[{new}]: kept the source column name (Tableau caption "{old}" is the visual header only)')
                done = True
        return done

    def unique_column_names(self, tables):
        """Power BI object names are case-insensitive: a Tableau calc "EVENT" = UPPER([Event]) next to the source column
        Event, or two calcs "POINT OF SALE" / "Point of Sale", are one name twice and Desktop refuses the model. The
        source column keeps its name; the later calc gets " (calc)". Names that differ only in case are told apart in
        the DAX and the report spec (exact-case references), so they are remapped. An exact duplicate cannot be told
        apart there: a calc that only repeats the column (Region = [Region]) is dropped, any other one is renamed
        without remapping (its visuals keep the source column). Returns {(table, old): (table, new)} to remap."""
        remap = {}
        for t in tables.values():
            seen = {}
            for c in t['columns'] + t['calc_columns']:
                low = c['name'].lower()
                if low not in seen:
                    seen[low] = c
                    continue
                first, old = seen[low], c['name']
                if old == first['name'] and c in t['calc_columns'] and \
                        re.fullmatch(re.escape(tq(t['name'])) + r'\[' + re.escape(old) + r'\]', (c.get('dax') or '').strip()):
                    t['calc_columns'].remove(c)
                    self.log.append(f'calc column {t["name"]}[{old}] = the column of the same name: not created twice')
                    continue
                base = f'{old} (calc)' if c in t['calc_columns'] else f'{old} (2)'
                new, k = base, 2
                while new.lower() in seen:
                    new, k = f'{base[:-1]} {k})', k + 1
                c['name'] = new
                seen[new.lower()] = c
                for x in t['columns'] + t['calc_columns']:
                    if x.get('sort_by') == old and x is not first:
                        x['sort_by'] = new
                if old != first['name']:
                    remap[(t['name'], old)] = (t['name'], new)
                    self.relationships = [(a, new if (a, ka) == (t['name'], old) else ka, b,
                                           new if (b, kb) == (t['name'], old) else kb) for a, ka, b, kb in self.relationships]
                    self.log.append(f'DECISION NEEDED: {t["name"]}[{old}] renamed [{new}]: Power BI names are not case-'
                                    f'sensitive, so it clashes with [{first["name"]}]')
                else:
                    self.log.append(f'DECISION NEEDED: {t["name"]}[{old}] is written twice (a Tableau calc named like a '
                                    f'column); the second is [{new}] and visuals use the first: check which one they need')
        return remap

    def place_calc_columns(self, tables):
        """A row-level calc reading columns of several tables lives on the many side of their relationship and reads
        the one side through RELATED (a calc column cannot see another table's row otherwise)."""
        many_to_one = {(a, b) for a, _, b, _ in self.relationships}
        moved = False
        for t in list(tables.values()):
            for c in list(t['calc_columns']):
                refs = {x.replace("''", "'") for x in re.findall(r"'((?:[^']|'')+)'\[", c['dax'])} - {'Date'}
                if len(refs) < 2 and refs <= {t['name']}:
                    continue
                host = next((h for h in [t['name']] + sorted(refs) if h in tables and
                             all(x == h or (h, x) in many_to_one for x in refs)), None)
                whole = [x for x in refs if x != host and re.search(re.escape(tq(x)) + r'(?!\[)', c['dax'])]
                if not host or whole:
                    self.log.append(f'calc column {t["name"]}[{c["name"]}] -> BLANK(): reads tables {", ".join(sorted(refs))} '
                                    + ('that no single many-side table reaches' if not host else
                                       f'and iterates {", ".join(whole)} from another table') + f'; write it by hand. Draft DAX: {c["dax"]}')
                    c['dax'] = 'BLANK()'
                    continue
                c['dax'] = re.sub(r"(?<!RELATED\()('(?:[^']|'')+')\[([^\]]+)\]",
                                  lambda m: m.group(0) if m.group(1) in (tq(host), "'Date'") else f'RELATED({m.group(0)})', c['dax'])
                if host != t['name']:
                    t['calc_columns'].remove(c)
                    tables[host]['calc_columns'].append(c)
                    self.moved_cols[(t['name'], c['name'])] = (host, c['name'])
                    moved = True
                    self.log.append(f'calc column [{c["name"]}] placed on {host} (many side), {t["name"]} read through RELATED')
        return moved

    def remap_spec(self, node):
        """Point every field of the report spec (roles, slicers, filters, drillthrough, URL actions, row conditions)
        at the merged table, so no visual keeps a binding to a table that is not written."""
        if isinstance(node, list):
            for x in node:
                self.remap_spec(x)
            return
        if not isinstance(node, dict):
            return
        if node.get('table') and node.get('name') and node.get('kind'):
            node['table'], node['name'] = self.remap(node['table'], node['name'])
        elif node.get('table') and node.get('column'):  # url_action
            node['table'], node['column'] = self.remap(node['table'], node['column'])
        elif node.get('type') == 'RowCondition':
            node['table'] = self.moved_tables.get(node['table'], node['table'])
            node['condition'] = self.remap_dax(node['condition'])
        for v in node.values():
            if isinstance(v, (dict, list)):
                self.remap_spec(v)

    def remap_dax(self, dax):
        if not dax or not (self.moved_cols or self.moved_tables):
            return dax

        def rep(m):
            t, c = m.group(1).replace("''", "'"), m.group(2)
            nt, nc = self.remap(t, c)
            return f'{tq(nt)}[{nc}]'
        dax = re.sub(r"'((?:[^']|'')+)'\[([^\]]+)\]", rep, dax)
        for old, new in self.moved_tables.items():  # whole-table references: FILTER('<old>', ...), COUNTROWS(...)
            dax = dax.replace(tq(old), tq(new))
        return dax

    def key_column(self, table, physical):
        c = next((c for c in table['columns'] if c.get('source_column') == physical), None)
        if c:
            return c['name']
        table['columns'].append({'name': physical, 'source_column': physical, 'datatype': 'string', 'hidden': True, 'key': True})
        return physical

    def sample_unique(self, src, colname):
        """True / False when the sample CSV shows the key unique or not; None without sample data."""
        rows = self.sample_rows(src)
        if not rows or colname not in rows[0]:
            return None
        vals = [r[colname] for r in rows if r.get(colname)]
        return len(vals) == len(set(vals))

    def composite_key(self, table, names):
        """One key column for a Tableau relationship on several fields (Power BI relates on one column): Power Query
        text '<a>|<b>' on the table, null when any part is null (Tableau never matches null keys)."""
        cols = [next(c for c in table['columns'] if c['name'] == n) for n in names]
        key = ' | '.join(names) + ' (key)'
        if not any(c['name'] == key for c in table['columns']):
            src = [c['source_column'] for c in cols]
            table['steps'].append((f'Added {key[:40]}', f'Table.AddColumn({{prev}}, {mstr(key)}, each if '
                                   + ' or '.join(f'{mfield(s)} = null' for s in src) + ' then null else Text.Combine({'
                                   + ', '.join(f'Text.From({mfield(s)})' for s in src) + '}, "|"), type nullable text)'))
            table['columns'].append({'name': key, 'source_column': key, 'datatype': 'string', 'hidden': True, 'derived': True})
        return key

    def logical_relationship(self, r, pairs, tables, to_model):
        """A Tableau relationship (logical layer, 2020.2+) -> a Power BI relationship. Keys: one column, a composite
        key column (several field pairs), or a calculated column. One side, in order: Tableau's Performance options
        (cardinality One), the sample data, the key named after its table (EVENT_ID -> EVENTS). With none of them the
        relationship is many-to-many: Tableau's own default, and unlike a guessed one side it never drops rows (a
        guessed one side keeps one row per key). Sample data showing neither side unique -> many-to-many too."""
        a_src, _, b_src, _ = pairs[0]
        a, b = to_model(a_src), to_model(b_src)
        calc = set(r.get('calc_keys') or [])
        label = f'Tableau relationship {a} - {b} ({r.get("tableau_on") or r.get("on")})'
        if calc and len(pairs) > 1:
            self.log.append(f'DECISION NEEDED: {label} not made: a composite key with a calculated field (Power Query cannot '
                            'read a DAX column); add the calc as a Power Query column')
            return
        if self.connected(a, b):
            self.log.append(f'DECISION NEEDED: {label} not made: {a} and {b} are already related (a second path is ambiguous)')
            return

        def key(t, cols):
            names = [c if c in calc else self.key_column(tables[t], c) for c in cols]
            return names[0] if len(names) == 1 else self.composite_key(tables[t], names)
        ka, kb = key(a, [p[1] for p in pairs]), key(b, [p[3] for p in pairs])
        # a key column no visual used is added as text; a key on a calculated (or typed) column has a type: one type
        calc_types = {c['name']: c.get('datatype') for c in self.digest.get('calculations', [])}
        ends = []
        for t, k in ((a, ka), (b, kb)):
            col = next((c for c in tables[t]['columns'] if c['name'] == k), None)
            ends.append((col, calc_types.get(k) if k in calc else (col or {}).get('datatype')))
        for (col, _), (_, other) in ((ends[0], ends[1]), (ends[1], ends[0])):
            if col is not None and col.get('key') and other and other != 'string' and not col.get('derived'):
                col['key'], col['datatype'] = False, other
        one, why = None, None
        if r.get('one_side') in ('a', 'b', 'both'):
            one, why = ('b' if r['one_side'] == 'both' else r['one_side']), 'Tableau Performance options (cardinality One)'
        elif len(pairs) == 1 and not calc:
            ua, ub = self.sample_unique(a_src, pairs[0][1]), self.sample_unique(b_src, pairs[0][3])
            if ua is False and ub is False:
                one, why = 'many', 'the sample data has repeated keys on both sides'
            elif ub or ua:
                one, why = ('b' if ub else 'a'), 'the sample data shows the key unique'
        if one is None and len(pairs) == 1:
            own_a, own_b = key_owner(ka, a), key_owner(kb, b)
            if own_a != own_b:
                one, why = ('a' if own_a else 'b'), 'the key is named after that table; sample data or Tableau cardinality would prove it'
        if one is None:
            one, why = 'many', 'Tableau default (many-to-many): no cardinality set, no sample data'
        if one == 'a':
            a, ka, b, kb = b, kb, a, ka
        self.relationships.append((a, ka, b, kb))
        if one == 'many':
            self.m2m.add((a, b))
            self.log.append(f'DECISION NEEDED: {label}: many-to-many {a}[{ka}] - {b}[{kb}] ({why}); if {b}[{kb}] is unique, make it '
                            f'one-to-many (faster, allows RELATED)')
        else:
            self.log.append(f'{label}: {a}[{ka}] -> {b}[{kb}], one side {b} ({why})'
                            + (f'; Tableau referential integrity "{r["integrity"]}"' if r.get('integrity') else ''))
        for t, k in ((a, ka), (b, kb)):
            if k in calc:
                self.calc_key_rels.append((t, k))

    def one_side(self, a_src, a_col, b_src, b_col):
        ua, ub = self.sample_unique(a_src, a_col), self.sample_unique(b_src, b_col)
        if ub and not ua:
            return 'b'
        if ua and not ub:
            return 'a'
        if ua and ub:
            return 'b'
        return None

    # ---------- measures ----------
    def rename(self, name):
        return self.renames.get(name, name)

    def fix_refs(self, dax):
        for old, new in self.renames.items():
            dax = re.sub(r"(?<!')\[" + re.escape(old) + r"\]", '[' + new + ']', dax)
        return dax

    def build_measures(self):
        spec_model = self.spec['model']
        fact_tables = [t for t in self.tables if t not in ('Date',)]
        self.measures = {}

        def add(name, dax, table, fmt=None, desc='', folder='Measures', hidden=False, notes=()):
            self.measures[name] = {'name': name, 'dax': dax, 'table': table, 'formatString': fmt, 'description': desc,
                                   'displayFolder': folder, 'hidden': hidden, 'notes': list(notes)}
        # parameters: a Start/End date pair shown as one Between slicer -> 'Report Period' + value measures
        pages = self.spec['pages']
        merged = next((s for p in pages for s in p['slicers'] if s.get('parameters')), None)
        self.report_period = None
        params = {p['name']: p for p in self.digest['parameters']}
        if merged:
            self.report_period = 'Report Period'
            for pname, fn in zip(merged['parameters'], ('MIN', 'MAX')):
                d = tableau_date((params.get(pname) or {}).get('value'))
                fallback = f'DATE({d[0]}, {d[1]}, {d[2]})' if d else 'BLANK()'
                add(f'{pname} Value', f"IF(ISFILTERED('Report Period'[Date]), {fn}('Report Period'[Date]), {fallback})",
                    'Report Period', 'm/d/yyyy', f'Tableau parameter "{pname}": {fn} of the Report Period slicer selection, '
                    f'default {params.get(pname, {}).get("value")}', 'Measures\\Parameters', True)
        # every other parameter -> a disconnected slicer table (when it has a value list) + a hidden value measure
        self.param_tables = {}
        for p in spec_model.get('parameters', []):
            if not p.get('measure_dax') or (merged and p['name'] in merged['parameters']):
                continue
            nm, dax = p['measure_dax'].split(' = ', 1)
            tbl = self.param_table(p, params.get(p['name']) or {})
            if tbl:
                tname, expr, col, dtype = tbl
                self.param_tables[tname] = (expr, col, dtype)
                sv = call_args(dax, 'SELECTEDVALUE')  # a default like "Demenko,Natalya S" holds a comma
                default = dax if not sv else (sv[1] if len(sv) > 1 else 'BLANK()')
                dax = f'SELECTEDVALUE({tq(tname)}[{col}], {default})'
                home, desc = tname, f'Tableau parameter "{p["name"]}": the {tname} slicer selection, default {default}'
            else:
                home, desc = self.report_period or fact_tables[0], f'Tableau parameter "{p["name"]}" (free entry: fixed to its default)'
            add(nm.strip('[]'), dax, home, None, desc, 'Measures\\Parameters', True)
        # parameter-driven calcs used as dimensions -> disconnected value tables (the visuals' measures are filtered
        # by <calc> IN VALUES(<table>) through the row-condition folding below)
        self.pdim_tables = {pd['name']: pd['values'] for pd in spec_model.get('parameter_dimensions', [])
                            if pd.get('values') and pd.get('host') in self.tables}
        for pd in spec_model.get('parameter_dimensions', []):
            if pd['name'] not in self.pdim_tables:
                self.log.append(f'parameter-driven dimension [{pd["name"]}] not built (host table {pd.get("host")} not in the model)')
        self.unique_measure_names(spec_model, fact_tables)
        # translated + explicit measures from the spec
        for t in spec_model['tables']:
            for m in t['measures']:
                if m['name'].endswith(' Filter') and 'COUNTROWS(FILTER' in (m['dax'] or ''):
                    continue
                table = m['table'] if m['table'] in self.tables else None
                if not table:
                    table = self.report_period if self.report_period and 'Value]' in (m['dax'] or '') else fact_tables[0]
                add(self.rename(m['name']), self.fix_refs(m['dax'] or 'BLANK()'), table, m['formatString'],
                    m['description'], m['displayFolder'], m['displayFolder'].endswith('Titles'), m.get('notes') or ())
        # fold row conditions (Tableau dimension filters on parameter-driven row calcs) into the measures they filter
        usage = {}
        for p in pages:
            for v in p['visuals']:
                conds = tuple(sorted(json.dumps(f, sort_keys=True) for f in v.get('visual_filters', []) if f.get('type') == 'RowCondition'))
                names = [f['name'] for fs in v['roles'].values() for f in fs if f['kind'] == 'Measure']
                names += [r['measure'] for r in v.get('reference_lines', [])]
                for n in names:
                    usage.setdefault(self.rename(n), set()).add(conds)
        self.variants = {}
        for name, condsets in usage.items():
            m = self.measures.get(name)
            if not m or not self.reads_rows(name):
                continue  # a parameter value or a constant reads no rows: a row filter changes nothing in it
            for conds in condsets:
                if not conds:
                    continue
                folded = self.fold(m['dax'], [json.loads(c) for c in conds])
                if len(condsets) == 1:
                    m['dax'] = folded
                    m['description'] += ' | Tableau sheet filter folded in: ' + '; '.join(json.loads(c)['tableau_calc_text']
                                                                                          if 'tableau_calc_text' in json.loads(c) else
                                                                                          f'[{json.loads(c)["calc"]}] = "{json.loads(c)["value"]}"'
                                                                                          for c in conds)
                else:
                    base = vn = f'{name} ({", ".join(dict.fromkeys(json.loads(c)["calc"] for c in conds))})'
                    k = 2
                    while vn in self.measures and self.measures[vn]['dax'] != folded:  # same calcs, other values
                        vn, k = f'{base[:-1]} {k})', k + 1
                    add(vn, folded, m['table'], m['formatString'], m['description'] + ' | filtered variant', m['displayFolder'])
                    self.variants[(name, conds)] = vn
        # a sheet filter on a parameter-driven row calc, on a visual with no measure to fold it into (a table of names,
        # a chart of a column): a hidden row count under the same condition, and the visual keeps the rows where it is
        # not blank. Before, such a table showed every row (Putney "Medal Haul 1" listed every player, not Player 1)
        for p in pages:
            for v in p['visuals']:
                conds = [f for f in v.get('visual_filters', []) if f.get('type') == 'RowCondition']
                fields = [f for fs in v['roles'].values() for f in fs]
                if not conds or any(f['kind'] == 'Measure' for f in fields) or not any(f['kind'] == 'Column' for f in fields) \
                        or v.get('visualType') in (None, 'cardVisual', 'card', 'multiRowCard', 'textbox', 'shape', 'image'):
                    continue
                host = {c.get('table') for c in conds}
                if len(host) != 1 or next(iter(host)) not in self.tables:
                    continue
                t = next(iter(host))
                nm = f'{v["sheet"]} Rows'
                if nm not in self.measures:
                    add(nm, self.fold(f'COUNTROWS({tq(t)})', conds), t, '#,##0',
                        f'Rows of Tableau sheet "{v["sheet"]}" that pass its filter(s) ' +
                        '; '.join(c.get('tableau_calc_text') or f'[{c.get("calc")}] = "{c.get("value")}"' for c in conds) +
                        ': the visual keeps the rows where this is not blank', 'Measures\\Sheet Filters', True)
                v['visual_filters'] = v.get('visual_filters', []) + [
                    {'type': 'NotBlank', 'field': {'kind': 'Measure', 'table': t, 'name': nm, 'label': nm}}]
                self.log.append(f'{p["name"]} / {v["sheet"]}: sheet filter on a parameter-driven calc applied through '
                                f'the hidden measure [{nm}] (visual filter "is not blank")')
        # pies of several measures (Tableau Measure Names slices): a slice table + SWITCH measure give them a Category
        self.slice_tables = {}
        for p in pages:
            for v in p['visuals']:
                if not v.get('measure_slices'):
                    continue
                ys = [f for f in v['roles'].get('Y', []) if f['kind'] == 'Measure']
                if not ys:
                    continue
                tname = f'{v["sheet"]} Slices'
                aliases = self.digest.get('measure_aliases') or {}  # Tableau headers renamed on the Measure Names card
                labels = [aliases.get(f.get('label')) or aliases.get(f'[{f.get("label")}]') or f.get('label') or f['name'] for f in ys]
                vname = f'{v["sheet"]} Slice Value'
                cases = ', '.join(f'{mstr(lab)}, [{self.measure_for(v, f["name"])}]' for lab, f in zip(labels, ys))
                fmt = (self.measures.get(self.measure_for(v, ys[0]['name'])) or {}).get('formatString') or '#,##0'
                add(vname, f"SWITCH(SELECTEDVALUE({tq(tname)}[Slice]), {cases})", tname, fmt,
                    f'Value of each slice of the Tableau pie "{v["sheet"]}" (one slice per measure)', 'Measures', False)
                self.slice_tables[tname] = labels
                v['roles'] = {'Category': [{'kind': 'Column', 'table': tname, 'name': 'Slice', 'label': 'Measure Names'}],
                              'Y': [{'kind': 'Measure', 'table': tname, 'name': vname, 'label': vname}]}
                self.log.append(f'{p["name"]} / {v["sheet"]}: pie of {len(ys)} measures -> slice table {tname} + [{vname}]')
        # reference lines -> measure-driven constant lines (Tableau "average per table" over the marks)
        self.refline_measures = {}
        for p in pages:
            for v in p['visuals']:
                for r in v.get('reference_lines', []):
                    cats = [f for role in ('Category', 'Series') for f in v['roles'].get(role, []) if f['kind'] == 'Column']
                    base = self.measure_for(v, r['measure'])
                    nm = f'Average {base} per {" and ".join(c["name"] for c in cats)}' if cats else f'Average {base}'
                    grid = 'CROSSJOIN(' + ', '.join(f"ALLSELECTED('{c['table']}'[{c['name']}])" for c in cats) + ')' if len(cats) > 1 else \
                        (f"ALLSELECTED('{cats[0]['table']}'[{cats[0]['name']}])" if cats else None)
                    dax = (f'AVERAGEX(FILTER({grid}, NOT ISBLANK([{base}])), [{base}])' if grid else f'[{base}]')
                    fmt = (self.measures.get(base) or {}).get('formatString') or '#,##0'
                    add(nm, dax, self.measures.get(base, {}).get('table') or fact_tables[0], fmt,
                        f'Tableau reference line "{r["label"]}": average of [{base}] over the visual\'s marks', 'Measures\\Reference Lines', True)
                    self.refline_measures[(v['id'], r['measure'])] = nm
        # dimensions in measure-only roles (Tableau ATTR in tooltips/cards): the model discourages implicit measures,
        # so an Aggregation projection fails with "Error fetching data"; bind an explicit SELECTEDVALUE measure instead
        for p in pages:
            for v in p['visuals']:
                for fs in v['roles'].values():
                    for f in fs:
                        if f['kind'] != 'Aggregation' or not f.get('table'):
                            continue
                        nm = f'{f["name"]} (ATTR)'
                        if nm not in self.measures:
                            add(nm, f"SELECTEDVALUE({tq(f['table'])}[{f['name']}])", f['table'], None,
                                f'Tableau ATTR([{f["name"]}]): the single value in context, blank when several (Tableau shows *)',
                                'Measures\\Attributes', True)
                        f.update({'kind': 'Measure', 'name': nm})
        # help icon for Tableau help sheets
        if any(v['visualType'] == 'help-icon' for p in pages for v in p['visuals']):
            add('Help Icon', '"?"', self.report_period or fact_tables[0], None, 'Text of the "?" help icon (Tableau help sheet)',
                'Measures\\Info', True)

    def unique_measure_names(self, spec_model, fact_tables):
        """A measure name must be unique in the whole model and differ from the columns of its own table, ignoring case
        ("Total Ty_units_Ship" on one table and "Total TY_UNITS_Ship" on another are one name to Power BI, and Desktop
        refuses the model). The later one gets its table name added, through the measure_names renames, so every
        reference (exact case) and visual binding follows. A rename the developer set in decisions.json is kept."""
        exact = set(self.measures)  # the same name twice is one measure (add() keeps one): not a clash
        used = {n.lower() for n in exact}
        cols = {name: {c['name'].lower() for c in t['columns'] + t['calc_columns']} for name, t in self.tables.items()}
        for t in spec_model['tables']:
            for m in t['measures']:
                if m['name'].endswith(' Filter') and 'COUNTROWS(FILTER' in (m['dax'] or ''):
                    continue
                name = self.rename(m['name'])
                table = m['table'] if m['table'] in self.tables else fact_tables[0]
                clash = name not in exact and name.lower() in used
                if not clash and name.lower() not in cols.get(table, ()) or m['name'] in self.decisions.get('measure_names', {}):
                    exact.add(name)
                    used.add(name.lower())
                    continue
                base = new = f'{name} ({table})'
                k = 2
                while new.lower() in used or new.lower() in cols.get(table, ()):
                    new, k = f'{base[:-1]} {k})', k + 1
                self.renames[m['name']] = new
                exact.add(new)
                used.add(new.lower())
                self.log.append(f'DECISION NEEDED: measure [{name}] renamed [{new}]: Power BI names are not case-sensitive, '
                                'so it clashes with another measure or a column of its table (decisions.json '
                                '"measure_names" to choose another name)')

    def param_table(self, p, dp):
        """(table, DAX table expression, value column, TMDL type) for a Tableau parameter with a value list, or None."""
        tname = param_table_name(p['name'])
        rhs = (p.get('table_dax') or '').split(' = ', 1)[-1] if p.get('table_dax') else None
        if rhs and rhs.startswith('CALENDAR'):
            return tname, rhs, 'Date', 'dateTime'
        if rhs:
            dtype = 'string' if '"Value", STRING' in rhs or (dp.get('datatype') == 'string') else \
                'int64' if dp.get('datatype') == 'integer' else 'double'
            return tname, rhs, 'Value', dtype
        if dp.get('source_field'):  # list filled from a field's values when the workbook opens
            calc = next((c for c in self.digest['calculations'] if c['name'] == dp['source_field'] and
                         c.get('datasource') in (None, dp.get('source_datasource')) and c.get('dax')), None)
            col = next((c for c in self.digest['columns'] if c['name'] == dp['source_field']), None)
            if calc:
                m = re.search(r"'((?:[^']|'')+)'\[", calc['dax'])
                table, expr = (m.group(1).replace("''", "'") if m else None), calc['dax']
            elif col and col.get('table'):
                table, cname = self.remap(col['table'], col['name'])
                expr = f"{tq(table)}[{cname}]"
            else:
                return None
            if table not in self.tables:
                return None
            dtype = TMDL_TYPE.get(dp.get('datatype') or 'string', 'string')
            return tname, f'DISTINCT(SELECTCOLUMNS({tq(table)}, "Value", {expr}))', 'Value', dtype
        return None

    def param_tmdl(self, name, expr, col, dtype):
        L = [f'/// Disconnected slicer table of a Tableau parameter; [{col}] feeds the parameter value measure.',
             f'table {tq(name)}', '']
        for m in self.measures.values():
            if m['table'] == name:
                L += self.measure_tmdl(m)
        L += [f'\tcolumn {tq(col)}', f'\t\tdataType: {dtype}', '\t\tsummarizeBy: none', f'\t\tsourceColumn: [{col}]', '',
              f'\tpartition {tq(name)} = calculated', '\t\tmode: import', f'\t\tsource = {expr}', '']
        return '\n'.join(L)

    def fold(self, dax, conds):
        hoisted, body = {}, []
        for c in conds:
            cond = c['condition']
            for ref in sorted(set(re.findall(r"(?<!')\[([^\]]+ Value)\]", cond))):
                hoisted.setdefault(ref, f'_p{len(hoisted) + 1}')
                cond = cond.replace(f'[{ref}]', hoisted[ref])
            body.append(f"KEEPFILTERS(FILTER({tq(c['table'])}, {cond}))")
        lines = [f'VAR {v} = [{r}]' for r, v in hoisted.items()]
        lines.append(f'RETURN CALCULATE({dax}, {", ".join(body)})')
        return '\n'.join(lines)

    def measure_for(self, v, name):
        name = self.rename(name)
        conds = tuple(sorted(json.dumps(f, sort_keys=True) for f in v.get('visual_filters', []) if f.get('type') == 'RowCondition'))
        return self.variants.get((name, conds), name)

    # ---------- TMDL writers ----------
    @staticmethod
    def sample_renames(t, cfg):
        """sources.json "columns" {name: CSV header} -> [(CSV header, source column)] for one table. The name may be the
        column's source column or its Power BI name; either way the sample rows come out under the source column, the
        name the warehouse branch, sourceColumn and every later M step use."""
        out = []
        for alias, src in (cfg.get('columns') or {}).items():
            c = next((c for c in t['columns'] if c.get('source_column') == alias), None) or \
                next((c for c in t['columns'] if alias in (c['name'], c.get('caption'))), None)  # caption: keep_source_names
            target = c['source_column'] if c and c.get('source_column') else alias
            if src != target and (src, target) not in out:
                out.append((src, target))
        return out

    # ---------- source connectors ----------
    def source_connections(self):
        """relation alias -> (relation, Tableau connection) for every table a workbook (or its .tds) reads: plain and
        joined relations and the tables behind each logical object (relationships)."""
        if getattr(self, '_src_conns', None) is not None:
            return self._src_conns
        out = {}
        for ds in (self.digest or {}).get('datasources', []):
            c = ds.get('connection') or {}
            src = c.get('tds') or c
            named = src.get('named') or []

            def walk(r):
                if not r:
                    return
                for x in r.get('inputs') or r.get('union') or []:
                    walk(x)
                if r.get('alias') and r.get('table'):
                    conn = next((n for n in named if n.get('name') and n.get('name') == r.get('connection')), None)
                    if conn is None and len(named) == 1:
                        conn = named[0]
                    out.setdefault(r['alias'], (r, conn))
                    # a text file's relation is 'rates.csv' / [rates#csv]; the model table is 'rates'
                    out.setdefault(re.sub(r'\.(csv|txt|tsv)$', '', r['alias'], flags=re.I), (r, conn))
            walk(src.get('relations'))
            for r in (src.get('physical_by_object') or {}).values():
                walk(r)
        self._src_conns = out
        return out

    def connector_mode(self):
        """'warehouse': every table from the target warehouse (WarehouseServer / WarehouseDatabase, SQL endpoint), as
        before; 'source': each table through the connector of the source Tableau read it from (Snowflake, Excel...).
        sources.json "connector" decides; without it: warehouse when a warehouse server is set, else source."""
        m = (self.sources or {}).get('connector')
        if m in ('warehouse', 'source'):
            return m
        return 'warehouse' if ((self.sources or {}).get('warehouse') or {}).get('server') else 'source'

    def source_expr(self, t, sql=None):
        """Power Query that reads table t (or runs its custom SQL) from its Tableau source with the matching connector,
        or None (published source without .tds, unknown type, warehouse mode): the warehouse SourceTable then."""
        if self.connector_mode() != 'source':
            return None
        rel, conn = self.source_connections().get(t['source'], (None, None))
        if not conn or not rel:
            return None
        cls = (conn.get('class') or '').lower()
        parts = [p for p in re.findall(r'\[((?:[^\]]|\]\])+)\]', rel.get('table') or '')] or [rel.get('table') or t['source']]
        name = parts[-1].replace(']]', ']')
        schema = parts[-2] if len(parts) > 1 else conn.get('schema')
        db = parts[-3] if len(parts) > 2 else conn.get('dbname')
        srv = mstr(conn.get('server') or '')
        nav = lambda keys: ''.join('{[' + ', '.join(f'{k} = {mstr(v)}' for k, v in kv) + ']}[Data]' for kv in keys)
        if cls == 'sqlserver':
            root = f'Sql.Database({srv}, {mstr(db or "")})'
            item = nav([[('Schema', schema or 'dbo'), ('Item', name)]])
        elif cls == 'snowflake':
            # Name only (no Kind): the same lookup finds a table or a view
            root = f'Snowflake.Databases({srv}, {mstr(conn.get("warehouse") or "")})' + nav([[('Name', db or '')]])
            item = nav([[('Name', schema or 'PUBLIC')], [('Name', name)]])
        elif cls in ('postgres', 'postgresql'):
            root = f'PostgreSQL.Database({srv}, {mstr(db or "")})'
            item = nav([[('Schema', schema or 'public'), ('Item', name)]])
        elif cls == 'oracle':
            s = conn.get('server') or ''
            root = f'Oracle.Database({mstr(s + ("/" + conn["service"] if conn.get("service") else ""))})'
            item = nav([[('Schema', schema or ''), ('Item', name)]])
        elif cls == 'teradata':
            root = f'Teradata.Database({srv})'
            item = nav([[('Schema', schema or db or ''), ('Item', name)]])
        elif cls == 'mysql':
            root = f'MySQL.Database({srv}, {mstr(db or "")})'
            item = nav([[('Schema', db or ''), ('Item', name)]])
        elif cls == 'redshift':
            root = f'AmazonRedshift.Database({srv}, {mstr(db or "")})'
            item = nav([[('Name', schema or 'public')], [('Name', name)]])
        elif cls == 'excel-direct' and not sql:
            sheet = name[:-1] if name.endswith('$') else name
            kind = ', Kind = "Sheet"' if name.endswith('$') else ''
            self.log.append(f'{t["name"]}: reads Excel {conn.get("filename") or "(file not recorded)"} sheet "{sheet}" (Tableau source)')
            return (f'Table.PromoteHeaders(Excel.Workbook(File.Contents({mstr(conn.get("filename") or "")}), null, true)'
                    f'{{[Item = {mstr(sheet)}{kind}]}}[Data], [PromoteAllScalars = true])')
        elif cls == 'textscan' and not sql:
            d = conn.get('directory') or ''
            path = (d.rstrip('/\\') + ('/' if '/' in d else '\\') if d else '') + (conn.get('filename') or name.replace('#', '.'))
            return (f'Table.PromoteHeaders(Csv.Document(File.Contents({mstr(path)}), [Delimiter = ",", Encoding = 65001, '
                    'QuoteStyle = QuoteStyle.Csv]), [PromoteAllScalars = true])')
        else:
            if cls not in ('sqlproxy', ''):
                self.log.append(f'DECISION NEEDED: {t["name"]}: Tableau source type "{cls}" has no connector mapping; it reads '
                                'the warehouse (SourceTable). Set the Power Query source by hand')
            return None
        self.log.append(f'{t["name"]}: reads {cls} {conn.get("server") or ""} {".".join(x for x in (db, schema, name) if x)} '
                        '(Tableau source connector; credentials are asked on the first refresh)')
        if sql:
            return f'Value.NativeQuery({root}, {mstr(sql)}, null, [EnableFolding = true])'
        return root + item

    def m_partition(self, t):
        if t.get('m_lines'):  # a table this build makes from other tables (blend shared dimension)
            return t['m_lines']
        rel = t['source']
        cfg = self.sources.get('tables', {}).get(rel, {})
        tds_rel = self.relations.get(rel, {})
        sql = tds_rel.get('custom_sql')
        phys = []
        for c in t['columns']:
            if c.get('source_column') and c['source_column'] not in phys and not c['name'].endswith(' (Date)') \
                    and not c.get('dataCategory') and not c.get('derived') \
                    and c['name'] not in [g['name'] for g in self.digest['groups']]:
                phys.append(c['source_column'])
        steps = []
        sample_file = cfg.get('file') or f'{rel}.csv'
        # sources.json sample transforms (where / columns / text_replace) apply to every table, not only custom SQL
        sample_steps = []
        for k, v in (cfg.get('where') or {}).items():
            sample_steps.append(f'Table.SelectRows({{prev}}, each {mfield(k)} = {mstr(v)})')
        renames = self.sample_renames(t, cfg) + t.get('extra_renames', [])  # extra: columns of merged tables
        if renames:
            # rename on the whole table (the used-columns step below selects); a CSV column already holding a target
            # name is dropped first so the rename cannot collide
            clash = [b for _, b in renames if b not in {a for a, _ in renames}]
            base = '{prev}' if not clash else \
                'Table.RemoveColumns({prev}, {' + ', '.join(mstr(b) for b in clash) + '}, MissingField.Ignore)'
            sample_steps.append(f'Table.RenameColumns({base}, {{' + ', '.join('{' + mstr(a) + ', ' + mstr(b) + '}' for a, b in renames) + '})')
        for colname, (a, b) in (cfg.get('text_replace') or {}).items():
            sample_steps.append(f'Table.TransformColumns({{prev}}, {{{{{mstr(colname)}, each if _ = null then null else Text.Replace(_, {mstr(a)}, {mstr(b)}), type nullable text}}}})')
        sample_expr = f'SampleTable({mstr(sample_file)})'
        for s in sample_steps:
            sample_expr = s.replace('{prev}', sample_expr)
        own = self.source_expr(t, sql)  # the Tableau source's own connector ('source' mode), else None
        if sql:
            native = own or f'Value.NativeQuery(Sql.Database(WarehouseServer, WarehouseDatabase), {mstr(sql)}, null, [EnableFolding = true])'
            steps.append(('Read source rows', f'if UseSampleData then {sample_expr} else {native}'))
            if not own:
                self.log.append(f'{t["name"]}: custom SQL from the .tds runs as a native query in warehouse mode; check the SQL '
                                f'dialect (Tableau source may not be T-SQL): {sql[:90]}')
        elif own:
            steps.append(('Read source rows', f'if UseSampleData then {sample_expr} else {own}'))
        else:
            src_table = (tds_rel.get('table') or rel).split('.')[-1].strip('[]')
            if sample_steps:
                steps.append(('Read source rows', f'if UseSampleData then {sample_expr} else SourceTable({mstr(src_table)}, {mstr(sample_file)})'))
            else:
                steps.append(('Read source rows', f'SourceTable({mstr(src_table)}, {mstr(sample_file)})'))
        steps.append(('Selected used columns', 'Table.SelectColumns({prev}, {' + ', '.join(mstr(p) for p in phys) + '})'))
        typed = []
        stamps = []
        for c in t['columns']:
            if c.get('source_column') not in phys:
                continue
            dt = c['datatype']
            if dt == 'datetime':
                stamps.append(c['source_column'])
            elif c.get('key'):
                typed.append('{' + mstr(c['source_column']) + ', type text}')
            elif dt in M_TYPE:
                typed.append('{' + mstr(c['source_column']) + ', ' + M_TYPE[dt] + '}')
        if typed:
            steps.append(('Changed column types', 'Table.TransformColumnTypes({prev}, {' + ', '.join(typed) + '}, "en-US")'))
        if stamps:
            steps.append(('Converted timestamps', 'Table.TransformColumns({prev}, {' + ', '.join('{' + mstr(s) + ', ToDateTime, type nullable datetime}' for s in stamps) + '})'))
        # one side of a relationship: one row per key
        late = []  # one-row-per-key on a column Power Query adds in t['steps'] (composite key): after those steps
        for frm, fc, to, tc in self.relationships:
            if to == t['name'] and (frm, to) not in getattr(self, 'm2m', ()):  # many-to-many: every row stays
                kcol = next((c for c in t['columns'] if c['name'] == tc), None)
                if kcol is None:
                    continue  # a DAX calculated column: Power Query cannot see it
                kc = kcol['source_column']
                step = (f'Kept one row per {kc[:30]}', f'Table.Distinct(Table.SelectRows({{prev}}, each {mfield(kc)} <> null), {{{mstr(kc)}}})')
                (late if kcol.get('derived') else steps).append(step)
        steps += t['steps'] + late
        lines, prev = ['let'], None
        for i, (label, expr) in enumerate(steps):
            nm = f'#"{label}"'
            e = expr.replace('{prev}', prev or '')
            lines.append(f'    {nm} = {e}' + (',' if i < len(steps) - 1 else ''))
            prev = nm
        lines += ['in', f'    {prev}']
        return lines

    def write_model(self):
        d = os.path.join(self.model_dir, 'definition')
        os.makedirs(os.path.join(d, 'tables'), exist_ok=True)
        for dp, _, fs in os.walk(d):  # replace files in place (OneDrive can hold folder handles briefly)
            for fn in fs:
                fp = os.path.join(dp, fn)
                # also the empty files an earlier build left when a table name held a ':' (no .tmdl extension)
                if fn.endswith('.tmdl') or (os.path.basename(dp) == 'tables' and os.path.getsize(fp) == 0):
                    os.remove(fp)
        with open(os.path.join(self.model_dir, 'definition.pbism'), 'w', encoding='utf-8') as f:
            json.dump({'$schema': 'https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json',
                       'version': '4.2', 'settings': {'qnaEnabled': True}}, f, indent=4)
        def w(name, text):
            with open(os.path.join(d, name), 'w', encoding='utf-8') as fh:
                fh.write(text)
        w('database.tmdl', f'database {tq(self.wb)}\n\tcompatibilityLevel: 1702\n\tcompatibilityMode: powerBI\n\tlanguage: 1033\n')
        names = list(self.tables) + ['Date'] + ([self.report_period] if self.report_period else []) + list(self.slice_tables) \
            + list(self.param_tables) + list(self.pdim_tables)
        w('model.tmdl', 'model Model\n\tculture: en-US\n\tdefaultPowerBIDataSourceVersion: powerBI_V3\n\tdiscourageImplicitMeasures\n'
                        '\tsourceQueryCulture: en-US\n\nannotation __PBI_TimeIntelligenceEnabled = 0\n\n'
                        + ''.join(f'ref table {tq(n)}\n' for n in names))
        w('expressions.tmdl', self.expressions())
        rel_lines = []
        for r in self.relationships:
            frm, fc, to, tc = r
            rel_lines += [f'relationship {tq(frm + " to " + to)}', f'\tfromColumn: {tq(frm)}.{tq(fc)}', f'\ttoColumn: {tq(to)}.{tq(tc)}']
            if (frm, to) in getattr(self, 'm2m', ()):
                rel_lines.append('\ttoCardinality: many')  # fromCardinality is many by default
            if r in getattr(self, 'both', ()):
                rel_lines.append('\tcrossFilteringBehavior: bothDirections')
            rel_lines.append('')
        w('relationships.tmdl', '\n'.join(rel_lines))
        written = set()

        def wt(tname, text):  # one file per table; names that differ only in case would overwrite each other on Windows
            fn = tmdl_file(tname)
            if fn.lower() in written:
                raise ValueError(f'two model tables named "{tname}" (names differ only in case): give one of them another name')
            written.add(fn.lower())
            w(os.path.join('tables', fn), text)
        for t in self.tables.values():
            wt(t['name'], self.table_tmdl(t))
        wt('Date', self.date_tmdl('Date', True))
        if self.report_period:
            wt(self.report_period, self.date_tmdl(self.report_period, False))
        for tname, labels in self.slice_tables.items():
            wt(tname, self.slice_tmdl(tname, labels))
        for tname, (expr, col, dtype) in self.param_tables.items():
            wt(tname, self.param_tmdl(tname, expr, col, dtype))
        for tname, values in self.pdim_tables.items():
            wt(tname, self.pdim_tmdl(tname, values))

    def pdim_tmdl(self, name, values):
        """Calculated table with the values of a parameter-driven Tableau calc, in Tableau's sort order."""
        rows = ', '.join('{' + mstr(v) + f', {i}' + '}' for i, v in enumerate(values, 1))
        L = [f'/// Values of the Tableau calc [{name}] (reads a parameter, so it cannot be a calculated column). Visuals '
             f'group by this column; their measures keep the rows where the calc equals the value.', f'table {tq(name)}', '']
        for m in self.measures.values():
            if m['table'] == name:
                L += self.measure_tmdl(m)
        L += [f'\tcolumn {tq(name)}', '\t\tdataType: string', '\t\tsummarizeBy: none', f'\t\tsourceColumn: [{name}]',
              '\t\tsortByColumn: Order', '',
              '\tcolumn Order', '\t\tdataType: int64', '\t\tisHidden', '\t\tsummarizeBy: none', '\t\tsourceColumn: [Order]', '',
              f'\tpartition {tq(name)} = calculated', '\t\tmode: import',
              f'\t\tsource = DATATABLE({mstr(name)}, STRING, "Order", INTEGER, {{{rows}}})', '']
        return '\n'.join(L)

    def slice_tmdl(self, name, labels):
        """Calculated table with one row per pie slice (measure name), sorted in Tableau's order."""
        rows = ', '.join('{' + mstr(lab) + f', {i}' + '}' for i, lab in enumerate(labels, 1))
        L = [f'/// Slices of a Tableau pie built from several measures (Measure Names); [Slice] is the pie Category.',
             f'table {tq(name)}', '']
        for m in self.measures.values():
            if m['table'] == name:
                L += self.measure_tmdl(m)
        L += ['\tcolumn Slice', '\t\tdataType: string', '\t\tsummarizeBy: none', '\t\tsourceColumn: [Slice]', '\t\tsortByColumn: Order', '',
              '\tcolumn Order', '\t\tdataType: int64', '\t\tisHidden', '\t\tsummarizeBy: none', '\t\tsourceColumn: [Order]', '',
              f'\tpartition {tq(name)} = calculated', '\t\tmode: import',
              f'\t\tsource = DATATABLE("Slice", STRING, "Order", INTEGER, {{{rows}}})', '']
        return '\n'.join(L)

    def expressions(self):
        wh = self.sources.get('warehouse', {})
        sample = self.sources.get('mode') == 'sample'
        folder = self.sources.get('sample_folder', '')
        e = []
        e.append(f'/// SQL endpoint of the warehouse (used when UseSampleData is false).\nexpression WarehouseServer = {mstr(wh.get("server", "TBD"))} meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]\n')
        e.append(f'/// Warehouse (database) name.\nexpression WarehouseDatabase = {mstr(wh.get("database", "TBD"))} meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]\n')
        e.append(f'/// Warehouse schema of the source tables.\nexpression WarehouseSchema = {mstr(wh.get("schema", "dbo"))} meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]\n')
        e.append(f'/// TRUE = load the sample CSVs from SampleDataFolder; FALSE = load from the warehouse.\nexpression UseSampleData = {"true" if sample else "false"} meta [IsParameterQuery=true, Type="Logical", IsParameterQueryRequired=true]\n')
        e.append(f'/// Folder with the sample CSV extracts (read only).\nexpression SampleDataFolder = {mstr(folder)} meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]\n')
        e.append('/// Reads one sample CSV with headers; empty cells become null.\nexpression SampleTable =\n\t\t(fileName as text) as table =>\n\t\t    let\n'
                 '\t\t        Source = Csv.Document(File.Contents(SampleDataFolder & "\\" & fileName), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),\n'
                 '\t\t        Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),\n'
                 '\t\t        BlanksAsNull = Table.TransformColumns(Promoted, List.Transform(Table.ColumnNames(Promoted), (c) => {c, each if _ = "" then null else _}))\n'
                 '\t\t    in\n\t\t        BlanksAsNull\n')
        e.append('/// Returns a source table: the sample CSV when UseSampleData is true, otherwise WarehouseSchema.tableName.\nexpression SourceTable =\n'
                 '\t\t(tableName as text, sampleFile as text) as table =>\n\t\t    if UseSampleData then SampleTable(sampleFile)\n'
                 '\t\t    else Sql.Database(WarehouseServer, WarehouseDatabase){[Schema = WarehouseSchema, Item = tableName]}[Data]\n')
        e.append('/// Converts a warehouse datetime or an ISO-8601 text timestamp (e.g. 2021-11-22T19:13:33Z) to a datetime without zone.\nexpression ToDateTime =\n'
                 '\t\t(value as any) as nullable datetime =>\n\t\t    if value = null then null\n\t\t    else if value is datetime then value\n'
                 '\t\t    else if value is datetimezone then DateTimeZone.RemoveZone(value)\n\t\t    else if value is date then DateTime.From(value)\n'
                 '\t\t    else DateTimeZone.RemoveZone(DateTimeZone.FromText(value))\n')
        return '\n'.join(e)

    def measure_tmdl(self, m):
        out = []
        if m['description']:
            out.append(f'\t/// {m["description"][:400]}')
        dax = m['dax']
        if '\n' in dax:
            out.append(f'\tmeasure {tq(m["name"])} = ```')
            out += [f'\t\t\t{ln}' for ln in dax.split('\n')]
            out.append('\t\t\t```')
        else:
            out.append(f'\tmeasure {tq(m["name"])} = {dax}')
        if m['formatString']:
            out.append(f'\t\tformatString: {m["formatString"]}')
        if m['displayFolder']:
            out.append(f'\t\tdisplayFolder: {m["displayFolder"]}')
        if m['hidden']:
            out.append('\t\tisHidden')
        out.append('')
        return out

    def table_tmdl(self, t):
        L = [f'table {tq(t["name"])}', '']
        for m in self.measures.values():
            if m['table'] == t['name']:
                L += self.measure_tmdl(m)
        for c in t['columns']:
            L.append(f'\tcolumn {tq(c["name"])}')
            L.append(f'\t\tdataType: {TMDL_TYPE.get(c["datatype"], "string") if not c.get("key") else "string"}')
            if c['datatype'] in ('date', 'datetime'):
                L.append('\t\tformatString: m/d/yyyy')
            if c.get('hidden'):
                L.append('\t\tisHidden')
            if c.get('dataCategory'):
                L.append(f'\t\tdataCategory: {c["dataCategory"]}')
            if c['datatype'] in ('integer', 'real') or c.get('summarizeBy') == 'none':
                L.append('\t\tsummarizeBy: none')
            if c.get('sort_by'):
                L.append(f'\t\tsortByColumn: {tq(c["sort_by"])}')
            L.append(f'\t\tsourceColumn: {c["source_column"]}')
            L.append('')
        names = {c['name'] for c in t['calc_columns']}
        for c in t['calc_columns']:
            L.append(f'\tcolumn {tq(c["name"])} = {c["dax"]}')
            L.append(f'\t\tdataType: {TMDL_TYPE.get(c["datatype"], "string")}')
            if c.get('hidden'):
                L.append('\t\tisHidden')
            if c.get('sort_by') in names:  # a sort column moved elsewhere (or blanked) is never referenced
                L.append(f'\t\tsortByColumn: {tq(c["sort_by"])}')
            L.append('')
        L.append(f'\tpartition {tq(t["name"])} = m')
        L.append('\t\tmode: import')
        L.append('\t\tsource =')
        L += [f'\t\t\t\t{ln}' for ln in self.m_partition(t)]
        L.append('')
        return '\n'.join(L)

    def date_tmdl(self, name, full):
        years = [tableau_date(p.get('min'))[0] for p in self.digest['parameters'] if tableau_date(p.get('min'))]
        start = min(years) if years else 2015
        L = [f'table {tq(name)}']
        if full:
            L.append('\tdataCategory: Time')
        L.append('')
        for m in self.measures.values():
            if m['table'] == name:
                L += self.measure_tmdl(m)
        cols = [('Date', 'dateTime', 'Date', True)]
        if full:  # every column bhayankar_report.DATE_COL can bind a Tableau date part to (except Hour: the table is daily)
            cols += [('Year', 'int64', 'Year', False), ('Quarter', 'string', 'Quarter', False), ('Month Number', 'int64', 'Month Number', False),
                     ('Month', 'string', 'Month', False), ('Year Month', 'string', 'Year Month', False), ('Day', 'int64', 'Day', False)]
            cols += [(c, t, c, False) for c, t in DATE_EXTRA]
        sort_by = {'Month': 'Month Number', 'Year Quarter': 'Year Quarter Number', 'Day of Week': 'Day of Week Number',
                   'Month Day': 'Month Day Number'}
        for cname, dtype, src, is_key in cols:
            L.append(f'\tcolumn {tq(cname)}')
            L.append(f'\t\tdataType: {dtype}')
            if dtype == 'dateTime':
                L.append('\t\tformatString: m/d/yyyy')
            if is_key and full:
                L.append('\t\tisKey')
            if dtype == 'int64':
                L.append('\t\tsummarizeBy: none')
            if cname in sort_by:
                L.append(f'\t\tsortByColumn: {tq(sort_by[cname])}')
            if cname in sort_by.values():
                L.append('\t\tisHidden')
            L.append(f'\t\tsourceColumn: {src}')
            L.append('')
        m = ['let', f'    #"Listed dates" = List.Dates(#date({start}, 1, 1), Duration.Days(#date(Date.Year(DateTime.LocalNow()) + 1, 12, 31) - #date({start}, 1, 1)) + 1, #duration(1, 0, 0, 0)),',
             '    #"Converted to table" = Table.FromList(#"Listed dates", Splitter.SplitByNothing(), {"Date"}),',
             '    #"Changed date type" = Table.TransformColumnTypes(#"Converted to table", {{"Date", type date}})']
        last = '#"Changed date type"'
        if full:
            adds = [('Added year', 'Year', 'Date.Year([Date])', 'Int64.Type'),
                    ('Added quarter', 'Quarter', '"Q" & Text.From(Date.QuarterOfYear([Date]))', 'type text'),
                    ('Added month number', 'Month Number', 'Date.Month([Date])', 'Int64.Type'),
                    ('Added month name', 'Month', 'Date.ToText([Date], "MMMM", "en-US")', 'type text'),
                    ('Added year month', 'Year Month', 'Date.ToText([Date], "yyyy-MM")', 'type text'),
                    ('Added day', 'Day', 'Date.Day([Date])', 'Int64.Type'),
                    # Tableau's default week starts on Sunday (WEEK, TRUNC_WEEK)
                    ('Added week start', 'Week Start', 'Date.StartOfWeek([Date], Day.Sunday)', 'type date'),
                    ('Added week', 'Week', 'Date.WeekOfYear([Date], Day.Sunday)', 'Int64.Type'),
                    ('Added year quarter', 'Year Quarter', 'Text.From(Date.Year([Date])) & " Q" & Text.From(Date.QuarterOfYear([Date]))', 'type text'),
                    ('Added year quarter number', 'Year Quarter Number', 'Date.Year([Date]) * 10 + Date.QuarterOfYear([Date])', 'Int64.Type'),
                    ('Added day of week', 'Day of Week', 'Date.ToText([Date], "dddd", "en-US")', 'type text'),
                    ('Added day of week number', 'Day of Week Number', 'Date.DayOfWeek([Date], Day.Sunday) + 1', 'Int64.Type'),
                    ('Added month day', 'Month Day', 'Date.ToText([Date], "MMMM d", "en-US")', 'type text'),
                    ('Added month day number', 'Month Day Number', 'Date.Month([Date]) * 100 + Date.Day([Date])', 'Int64.Type')]
            m[-1] += ','
            for i, (label, cname, expr, typ) in enumerate(adds):
                m.append(f'    #"{label}" = Table.AddColumn({last}, "{cname}", each {expr}, {typ})' + (',' if i < len(adds) - 1 else ''))
                last = f'#"{label}"'
        m += ['in', f'    {last}']
        L += [f'\tpartition {tq(name)} = m', '\t\tmode: import', '\t\tsource =']
        L += [f'\t\t\t\t{ln}' for ln in m]
        L.append('')
        return '\n'.join(L)

    # =====================================================================================
    # report
    # =====================================================================================
    def field_expr(self, f, v=None):
        if f['kind'] == 'Measure':
            name = self.measure_for(v, f['name']) if v else self.rename(f['name'])
            return meas_ref(self.measures.get(name, {}).get('table') or f['table'], name)
        if f['kind'] == 'Aggregation':
            return {'Aggregation': {'Expression': col_ref(f['table'], f['name']), 'Function': 3}}
        return col_ref(f['table'], f['name'])

    def scaffold(self):
        long_prefix = '\\\\?\\' if os.name == 'nt' else ''  # rmtree of a too-long earlier build needs the \\?\ form
        shutil.rmtree(long_prefix + os.path.abspath(self.report_dir), ignore_errors=True)
        if self.item != self.wb:  # earlier builds with the full name
            for ext in ('.Report', '.SemanticModel'):
                old = long_prefix + os.path.abspath(os.path.join(self.pbip_dir, self.wb + ext))
                shutil.rmtree(old, ignore_errors=True)
                if os.path.exists(old):
                    self.log.append(f'could not remove the earlier build folder {self.wb}{ext} (open in Power BI Desktop or '
                                    f'held by OneDrive); it is not used by {self.wb}.pbip: delete it when it is closed')
        for fn in os.listdir(self.pbip_dir) if os.path.isdir(self.pbip_dir) else []:
            if fn.endswith('.pbip'):
                os.remove(os.path.join(self.pbip_dir, fn))
        # offline scaffold (no Node.js): the .pbip keeps the workbook name and points at the (maybe short) folders
        bs.scaffold(self.pbip_dir, self.item, f'../{self.item}.SemanticModel', pbip_name=self.wb)
        return os.path.join(self.report_dir, 'definition', 'pages')

    def write_report(self):
        pages_dir = self.scaffold()
        defdir = os.path.join(self.report_dir, 'definition')
        report = load_json(os.path.join(defdir, 'report.json'))
        self.images = {}
        order, written = [], []
        tooltip_pages = []
        for p in self.spec['pages']:
            page = {'name': hid('page:' + p['name']), 'displayName': p['name'], 'displayOption': 'FitToPage',
                    'width': p['width'], 'height': p['height']}
            if not p.get('visible', True):
                page['visibility'] = 'HiddenInViewMode'
            visuals = []
            data_visuals = []
            self.growth = []
            slicer_visuals, scoped = [], []
            for s in p['slicers']:
                x = self.safe(p, f'slicer {s["field"].get("label")}', self.slicer, p, s)
                if x:
                    slicer_visuals.append(x)
                    if s.get('scope_sheets'):
                        scoped.append((x['name'], set(s['scope_sheets'])))
            sheet_of = {}
            for v in p['visuals']:
                vis = self.safe(p, f'visual {v.get("sheet")}', self.visual, p, v, tooltip_pages)
                if vis:
                    visuals.append(vis)
                    if v['visualType'] not in ('help-icon', 'shape', None, 'custom'):
                        data_visuals.append(vis['name'])
                        sheet_of[vis['name']] = v['sheet']
            for st in p['static']:
                vis = self.safe(p, f'static {st.get("type")}', self.static, p, st, report)
                if vis:
                    visuals.append(vis)
            for nb in p.get('nav_buttons', []):  # Tableau filter action from a sheet with no dimension: opens the page
                host = next((x for x in visuals if sheet_of.get(x['name']) == nb['sheet']), None)
                hp = host['position'] if host else {'x': p['width'] - 120, 'y': 8, 'width': 0, 'height': 48}
                visuals.append({'name': hid(f'{p["name"]}:nav:{nb["sheet"]}:{nb["target"]}'),
                                'position': {'x': max(0, hp['x'] + hp['width'] - 96), 'y': max(0, hp['y'] + hp['height'] - 32),
                                             'width': 92, 'height': 28},
                                'visual': {'visualType': 'actionButton',
                                           'objects': {'icon': [{'properties': {'shapeType': sq('blank')}, 'selector': {'id': 'default'}}],
                                                       'text': [{'properties': {'show': boolean(True), 'text': sq('Details ›')},
                                                                 'selector': {'id': 'default'}}]},
                                           'visualContainerObjects': {'visualLink': [{'properties': {
                                               'show': boolean(True), 'type': sq('PageNavigation'),
                                               'navigationSection': sq(hid('page:' + nb['target'])),
                                               'tooltip': sq(f'Open {nb["target"]} (Tableau action {nb["action"]})')}}]}}})
            if p.get('nav_back') and not p.get('drillthrough'):
                visuals.append({'name': hid(f'{p["name"]}:back'), 'position': {'x': p['width'] - 80, 'y': 8, 'width': 64, 'height': 40},
                                'visual': {'visualType': 'actionButton',
                                           'objects': {'icon': [{'properties': {'shapeType': sq('back')}, 'selector': {'id': 'default'}}]},
                                           'visualContainerObjects': {'visualLink': [{'properties': {'show': boolean(True), 'type': sq('Back'),
                                                                                                     'tooltip': sq(f'Back to {", ".join(p["nav_back"])}')}}]}}})
            if self.growth:
                # taller slicers push the content below them down (and shorten it) so nothing overlaps
                edge, delta = min(g[0] for g in self.growth), max(g[1] for g in self.growth)
                for vis in visuals:
                    ps = vis['position']
                    if ps['y'] >= edge - 1:
                        ps['y'] += delta
                        ps['height'] = max(ps['height'] - delta, 40)
            # multi-measure cards are split last, once their final box is known (the growth step above moves and
            # shrinks whole visuals; applied to the parts one by one it would leave gaps and overflow)
            split = []
            for vis in visuals:
                parts = self.split_card(vis)
                if len(parts) > 1 and vis['name'] in sheet_of:
                    data_visuals.remove(vis['name'])
                    for part in parts:
                        data_visuals.append(part['name'])
                        sheet_of[part['name']] = sheet_of[vis['name']]
                split.extend(parts)
            visuals = split
            visuals += slicer_visuals
            if p.get('drillthrough') and not p['drillthrough']['fields']:
                # e.g. a Tableau action from "all sheets" passing all fields: nothing to bind, so no drillthrough page
                self.log.append(f'{p["name"]}: drillthrough target without fields (action {p["drillthrough"]["tableau_action"]}); '
                                'pick the drillthrough fields by hand')
            elif p.get('drillthrough'):
                dt = p['drillthrough']
                # several Tableau fields can land on one model field (YEAR() of different dates -> Date[Year]):
                # one drillthrough filter per model field, or the names repeat (PBIR_FILTER_NAME_DUPLICATE_LOCAL)
                uniq = {}
                for f in dt['fields']:
                    uniq.setdefault(json.dumps(self.field_expr(f), sort_keys=True), f)
                dt = {**dt, 'fields': list(uniq.values())}
                page['filterConfig'] = {'filters': [{'name': fid(f'{p["name"]}:drill:{f["name"]}'), 'field': self.field_expr(f),
                                                     'type': 'Categorical', 'howCreated': 'Drillthrough'} for f in dt['fields']]}
                page['pageBinding'] = {'name': 'Pod', 'type': 'Drillthrough', 'parameters': [
                    {'name': 'Param_' + fid(f'{p["name"]}:drill:{f["name"]}'), 'boundFilter': fid(f'{p["name"]}:drill:{f["name"]}'),
                     'fieldExpr': self.field_expr(f)} for f in dt['fields']]}
                visuals.append({'name': hid(f'{p["name"]}:back'), 'position': {'x': p['width'] - 80, 'y': 8, 'width': 64, 'height': 40},
                                'visual': {'visualType': 'actionButton',
                                           'objects': {'icon': [{'properties': {'shapeType': sq('back')}, 'selector': {'id': 'default'}}]},
                                           'visualContainerObjects': {'visualLink': [{'properties': {'show': boolean(True), 'type': sq('Back'),
                                                                                                     'tooltip': sq(f'Back to {dt["from_page"]}')}}]}}})
            seen = {}
            for pf in p.get('page_filters', []):
                key = f'{p["name"]}:pagefilter:{pf["field"]["name"]}'
                seen[key] = seen.get(key, 0) + 1
                flt = self.categorical_filter(pf, key if seen[key] == 1 else f'{key}:{seen[key]}')
                if flt:
                    page.setdefault('filterConfig', {}).setdefault('filters', []).append(flt)
            if not p.get('cross_filter') and len(data_visuals) > 1:
                page['visualInteractions'] = [{'source': a, 'target': b, 'type': 'NoFilter'}
                                              for a in data_visuals for b in data_visuals if a != b]
            for sname, scope in scoped:  # Tableau quick filter "apply to selected worksheets"
                outside = [b for b in data_visuals if sheet_of.get(b) not in scope]
                if outside:
                    page.setdefault('visualInteractions', []).extend({'source': sname, 'target': b, 'type': 'NoFilter'} for b in outside)
            self.write_page(pages_dir, page, visuals)
            order.append(page['name'])
        for tp in tooltip_pages:
            self.write_page(pages_dir, tp['page'], tp['visuals'])
            order.append(tp['page']['name'])
        if not order:  # a report needs a page: say why it is empty instead of failing the build
            note = 'This Tableau workbook has no dashboards and no visible worksheets to rebuild.'
            page = {'name': hid('page:empty'), 'displayName': 'Page 1', 'displayOption': 'FitToPage', 'width': 1280, 'height': 720}
            text = self.static(page, {'type': 'textbox', 'text': note,
                                      'position': {'x': 40, 'y': 300, 'width': 1200, 'height': 80}}, report)
            self.write_page(pages_dir, page, [text])
            order.append(page['name'])
            self.log.append('no dashboards or visible worksheets: wrote one empty page with a note')
        if self.images:
            rr = os.path.join(self.report_dir, 'StaticResources', 'RegisteredResources')
            os.makedirs(rr, exist_ok=True)
            items = []
            for src, nm in self.images.items():
                shutil.copyfile(src, os.path.join(rr, nm))
                items.append({'name': nm, 'path': nm, 'type': 'Image'})
            report.setdefault('resourcePackages', [])
            report['resourcePackages'] = [rp for rp in report['resourcePackages'] if rp.get('name') != 'RegisteredResources']
            report['resourcePackages'].append({'name': 'RegisteredResources', 'type': 'RegisteredResources', 'items': items})
            with open(os.path.join(defdir, 'report.json'), 'w', encoding='utf-8') as f:
                json.dump(report, f, indent=2)
        order = list(dict.fromkeys(order))  # a page id listed twice fails validation (PBIR_PAGE_ORDER_DUPLICATE)
        pj = load_json(os.path.join(pages_dir, 'pages.json'))
        pj['pageOrder'] = order
        pj['activePageName'] = order[0]
        with open(os.path.join(pages_dir, 'pages.json'), 'w', encoding='utf-8') as f:
            json.dump(pj, f, indent=2)

    def unresolved_ref(self, node):
        """First Column/Measure reference anywhere inside `node` that does not resolve in the model, as 'Table.Name'."""
        if isinstance(node, dict):
            for kind in ('Column', 'Measure'):
                v = node.get(kind)
                if isinstance(v, dict) and 'Property' in v:
                    if not self.field_ok({kind: v}):
                        return f'{v.get("Expression", {}).get("SourceRef", {}).get("Entity")}.{v["Property"]}'
            for v in node.values():
                r = self.unresolved_ref(v)
                if r:
                    return r
        elif isinstance(node, list):
            for v in node:
                r = self.unresolved_ref(v)
                if r:
                    return r
        return None

    def field_ok(self, f):
        """True when a PBIR field reference resolves in the model that was written."""
        for kind in ('Column', 'Measure'):
            if kind in f:
                ent = f[kind].get('Expression', {}).get('SourceRef', {}).get('Entity')
                prop = f[kind].get('Property')
                if kind == 'Column':
                    return prop in self.model_cols.get(ent, ())
                return prop in self.measures and self.measures[prop]['table'] == ent
        if 'Aggregation' in f:
            return self.field_ok(f['Aggregation'].get('Expression', {}))
        return True

    def ref_ok(self, qref):
        """'Table.Field' / 'Min(Table.Field)' (queryRef, metadata selector) resolves in the written model."""
        inner = re.sub(r'^\w+\((.*)\)$', r'\1', qref or '')
        for i in [k for k, ch in enumerate(inner) if ch == '.']:
            ent, prop = inner[:i], inner[i + 1:]
            if prop in self.model_cols.get(ent, ()) or (prop in self.measures and self.measures[prop]['table'] == ent):
                return True
        return False

    def filter_ok(self, x):
        ents = {s.get('Entity') for s in (x.get('filter') or {}).get('From', [])}
        return self.field_ok(x.get('field', {})) and all(e in self.model_cols for e in ents)

    def report_guard(self, page, visuals):
        """Drop report bindings that do not resolve in the model (a missing table, a measure used as a column
        filter, ...): Desktop shows "Something's wrong with one or more filters" / "Error fetching data" for them.
        Checks every place a visual names a field: projections, sorts, filters (field and From entity), and
        formatting selectors (metadata / webURL)."""
        where = page['displayName']
        flts = page.get('filterConfig', {}).get('filters', [])
        keep = [x for x in flts if self.filter_ok(x)]
        for x in flts:
            if x not in keep:
                self.log.append(f'{where}: page filter {x.get("name")} dropped (field not in the model)')
        if flts:
            page['filterConfig']['filters'] = keep
            if page.get('pageBinding'):
                names = {x['name'] for x in keep}
                page['pageBinding']['parameters'] = [b for b in page['pageBinding']['parameters'] if b['boundFilter'] in names]
        out = []
        for v in visuals:
            vis = v.get('visual', {})
            qs = vis.get('query', {}).get('queryState')
            dropped = []
            if qs:
                for role in list(qs):
                    projs = qs[role]['projections']
                    good = [p for p in projs if self.field_ok(p['field'])]
                    dropped += [p['queryRef'] for p in projs if p not in good]
                    if good:
                        qs[role]['projections'] = good
                    else:
                        del qs[role]
                srt = vis['query'].get('sortDefinition', {}).get('sort')
                if srt:
                    vis['query']['sortDefinition']['sort'] = [s for s in srt if self.field_ok(s['field'])]
                    if not vis['query']['sortDefinition']['sort']:
                        del vis['query']['sortDefinition']
            vf = v.get('filterConfig', {}).get('filters', [])
            if vf:
                good = [x for x in vf if self.filter_ok(x)]
                dropped += [x.get('name') for x in vf if x not in good]
                v['filterConfig']['filters'] = good
            for oname, entries in list((vis.get('objects') or {}).items()):
                kept = []
                for e in entries:
                    meta = (e.get('selector') or {}).get('metadata')
                    url = (e.get('properties') or {}).get('webURL', {}).get('expr')
                    scoped = [d.get('scopeId', {}).get('Comparison', {}).get('Left') for d in (e.get('selector') or {}).get('data') or []]
                    bad_scope = next((s for s in scoped if s and not self.field_ok(s)), None)  # per-value colour on a missing column
                    if (meta and not self.ref_ok(meta)) or (url and not self.field_ok(url)) or bad_scope:
                        what = meta or ('webURL' if url and not self.field_ok(url) else
                                        (bad_scope.get('Column') or {}).get('Property', 'a data point'))
                        dropped.append(f'{oname} formatting on {what}')
                        continue
                    bad = self.unresolved_ref(e.get('selector'))  # e.g. data-point colours keyed on a column that is not in the model
                    if bad:
                        dropped.append(f'{oname} formatting keyed on {bad}')
                        continue
                    kept.append(e)
                if kept:
                    vis['objects'][oname] = kept
                else:
                    del vis['objects'][oname]
            if qs is not None and not qs:
                self.log.append(f'{where}: {vis.get("visualType")} {v["name"]} removed: none of its fields are in the model ({", ".join(dropped)})')
                continue
            need = ((bv.metadata().get('visuals') or {}).get(vis.get('visualType')) or {}).get('requiredRoles') or []
            lost = [r for r in need if qs is not None and r not in qs]
            if lost:  # a visual without a required role is invalid PBIR: required_roles_ok() below makes it a card of its
                # values (DECISION NEEDED) or leaves it out; say which fields it lost
                self.log.append(f'{where}: {vis.get("visualType")} {v["name"]}: required role(s) {", ".join(lost)} lost their '
                                f'fields ({", ".join(dropped) or "none bound"})')
                dropped = []
            if dropped:
                self.log.append(f'{where}: {vis.get("visualType")} {v["name"]}: dropped bindings not in the model: {", ".join(dropped)}')
            if qs is not None and not self.required_roles_ok(where, v):
                continue
            out.append(v)
        return out

    def required_roles_ok(self, where, v):
        """A chart missing a required role fails validation (PBIR_ROLE_REQUIRED_MISSING), e.g. a line chart whose axis is a
        parameter-driven Tableau dimension. Charts that still have values become a card of those values (DECISION NEEDED);
        the rest are removed. Returns False when the visual is removed."""
        vis = v['visual']
        vt, qs = vis.get('visualType'), vis['query']['queryState']
        need = ((CATALOG.get('visuals') or {}).get(vt) or {}).get('required') or []
        missing = [r for r in need if not qs.get(r)]
        if not missing:
            return True
        values = [p for role in ('Y', 'Values', 'Data', 'Size', 'X') for p in (qs.get(role) or {}).get('projections', [])
                  if 'Measure' in p['field'] or 'Aggregation' in p['field']]
        if not values:
            self.log.append(f'{where}: {vt} {v["name"]} removed: required role(s) {", ".join(missing)} empty and no values to show')
            return False
        vis['visualType'] = 'cardVisual'
        vis['query'] = {'queryState': {'Data': {'projections': values}}}
        vis.pop('objects', None)
        self.log.append(f'DECISION NEEDED: {where}: {vt} {v["name"]} had no {", ".join(missing)} (e.g. the Tableau axis is a '
                        'parameter-driven field): built as a card of its values; rebuild the axis with a field parameter')
        return True

    def safe(self, p, what, fn, *args):
        """Generate one report item; an unexpected error skips that item (logged) instead of stopping the build."""
        try:
            return fn(*args)
        except Exception as e:  # noqa: BLE001
            self.log.append(f'{p["name"]}: {what} NOT GENERATED (build error {type(e).__name__}: {e}); add it by hand '
                            'and report the error')
            print(f'  warning: {p["name"]}: {what} not generated ({type(e).__name__}: {e})')
            return None

    def write_page(self, pages_dir, page, visuals):
        visuals = self.report_guard(page, visuals)
        if page.get('visualInteractions'):  # a visual the guard removed must not stay in "Edit interactions"
            ids = {v['name'] for v in visuals}
            page['visualInteractions'] = [x for x in page['visualInteractions'] if x['source'] in ids and x['target'] in ids]
            if not page['visualInteractions']:
                del page['visualInteractions']
        d = os.path.join(pages_dir, page['name'])
        os.makedirs(os.path.join(d, 'visuals'), exist_ok=True)
        with open(os.path.join(d, 'page.json'), 'w', encoding='utf-8') as f:
            json.dump({'$schema': PAGE_SCHEMA, **page}, f, indent=2)
        for i, v in enumerate(visuals):
            ps = v['position']
            if page.get('height') and ps['y'] + ps['height'] > page['height']:  # pushed down by a taller slicer
                ps['y'] = max(0, page['height'] - ps['height'])
            ps['x'], ps['y'] = max(0, ps['x']), max(0, ps['y'])  # Tableau zones at -1 px after rounding (QFP)
            v['position']['z'] = (i + 1) * 1000
            v['position']['tabOrder'] = (i + 1) * 1000
            vd = os.path.join(d, 'visuals', v['name'])
            os.makedirs(vd, exist_ok=True)
            with open(os.path.join(vd, 'visual.json'), 'w', encoding='utf-8') as f:
                json.dump({'$schema': VISUAL_SCHEMA, **v}, f, indent=2)
        # folders left behind by an earlier build (a sync client can keep them after their files are deleted)
        keep = {v['name'] for v in visuals}
        vroot = os.path.join(d, 'visuals')
        for stale in os.listdir(vroot):
            sp = os.path.join(vroot, stale)
            if stale not in keep and os.path.isdir(sp):
                shutil.rmtree(sp, ignore_errors=True)

    def title_vco(self, v):
        t = v.get('title') or {}
        props = {'show': boolean(t.get('show', True)), 'fontSize': num(12), 'fontColor': colour(DARK)}
        if t.get('measure'):
            mname = t['measure']
            props['text'] = {'expr': meas_ref(self.measures.get(mname, {}).get('table') or 'Report Period', mname)}
        elif t.get('text'):
            props['text'] = sq(t['text'])
        return [{'properties': props}]

    def colour_entries(self, v):
        out = []
        cmap = (v.get('colors') or {}).get('map') or {}
        targets = [f for role in ('Series', 'Category') for f in v['roles'].get(role, []) if f['kind'] == 'Column']
        for f in targets:
            pal = cmap.get(f['label'])
            if not pal:
                continue
            numeric = f['table'] == 'Date' and f['name'] in ('Year', 'Month Number', 'Day')
            for val, hexv in sorted(pal.items()):
                right = f'{val}L' if numeric and val.isdigit() else "'" + val.replace("'", "''") + "'"
                out.append({'properties': {'fill': colour(hexv)}, 'selector': {'data': [{'scopeId': {'Comparison': {
                    'ComparisonKind': 0, 'Left': col_ref(f['table'], f['name']), 'Right': {'Literal': {'Value': right}}}}}]}})
            break
        return out

    def visual(self, p, v, tooltip_pages):
        pos = dict(v.get('position') or {'x': 0, 'y': 0, 'width': 300, 'height': 200})
        vt = v['visualType']
        name = hid(f'{p["name"]}:{v["id"]}:{v["sheet"]}')
        if vt == 'help-icon':
            tp_name = hid('page:' + v['sheet'] + ' Tooltip')
            text = v.get('tooltip_text') or ''
            refs = re.findall(r'<\[([^\]]+)\]>', text)
            clean = re.sub(r'\s*<\[[^\]]+\]>\s*', ' ', text).strip()
            tvis = [{'name': hid(f'{tp_name}:text'), 'position': {'x': 12, 'y': 12, 'width': 396, 'height': 236},
                     'visual': {'visualType': 'textbox', 'objects': {'general': [{'properties': {'paragraphs': [
                         {'textRuns': [{'value': clean, 'textStyle': {'fontFamily': 'Segoe UI', 'fontSize': '9pt', 'color': DARK}}]}]}}]},
                                'visualContainerObjects': {'background': [{'properties': {'show': boolean(False)}}]}}}]
            for i, r in enumerate(refs):
                mname = self.rename(r)
                if mname in self.measures:
                    tvis.append({'name': hid(f'{tp_name}:card:{r}'), 'position': {'x': 12, 'y': 252 + 88 * i, 'width': 396, 'height': 84},
                                 'visual': {'visualType': 'cardVisual', 'query': {'queryState': {'Data': {'projections': [
                                     proj(meas_ref(self.measures[mname]['table'], mname))]}}},
                                            'objects': {'value': [{'properties': {'fontSize': num(14)}, 'selector': {'id': 'default'}}]}}})
            if any(tp['page']['name'] == tp_name for tp in tooltip_pages):
                tooltip_pages[:] = [tp for tp in tooltip_pages if tp['page']['name'] != tp_name]  # same sheet on two dashboards: one tooltip page
            tooltip_pages.append({'page': {'name': tp_name, 'displayName': f'{v["sheet"]} Tooltip', 'type': 'Tooltip',
                                           'displayOption': 'ActualSize', 'width': 420, 'height': 260 + 88 * len(refs),
                                           'visibility': 'HiddenInViewMode', 'pageBinding': {'name': 'Pod', 'type': 'Tooltip', 'parameters': []}},
                                  'visuals': tvis})
            return {'name': name, 'position': pos, 'visual': {
                'visualType': 'cardVisual',
                'query': {'queryState': {'Data': {'projections': [proj(meas_ref(self.measures['Help Icon']['table'], 'Help Icon'))]}}},
                'objects': {'value': [{'properties': {'fontSize': num(24), 'bold': boolean(True), 'fontColor': colour(GREY)}, 'selector': {'id': 'default'}}],
                            'label': [{'properties': {'show': boolean(False)}, 'selector': {'id': 'default'}}],
                            'outline': [{'properties': {'show': boolean(False)}, 'selector': {'id': 'default'}}]},
                'visualContainerObjects': {'background': [{'properties': {'show': boolean(False)}}],
                                           'visualTooltip': [{'properties': {'show': boolean(True), 'type': sq('Canvas'), 'section': sq(tp_name)}}]}}}
        title_measure = (v.get('title') or {}).get('measure')
        if vt is None and title_measure in self.measures:
            # title-only Tableau sheet (no shelves, text built from fields): a card showing the dynamic title measure
            self.log.append(f'{p["name"]} / {v["sheet"]}: title-only sheet -> card bound to [{title_measure}]')
            return {'name': name, 'position': pos, 'visual': {
                'visualType': 'cardVisual',
                'query': {'queryState': {'Data': {'projections': [proj(meas_ref(self.measures[title_measure]['table'], title_measure))]}}},
                'objects': {'value': [{'properties': {'fontSize': num(14), 'bold': boolean(True), 'fontColor': colour(DARK)}, 'selector': {'id': 'default'}}],
                            'label': [{'properties': {'show': boolean(False)}, 'selector': {'id': 'default'}}],
                            'outline': [{'properties': {'show': boolean(False)}, 'selector': {'id': 'default'}}]},
                'visualContainerObjects': {'background': [{'properties': {'show': boolean(False)}}],
                                           'title': [{'properties': {'show': boolean(False)}}]}}}
        if vt in (None, 'custom', 'shape'):
            self.log.append(f'{p["name"]} / {v["sheet"]}: {vt or "unmapped"} visual not generated ({"; ".join(v.get("notes", []))[:120]})')
            return None
        qs = {}
        for role, fields in v['roles'].items():
            projs = []
            for f in fields:
                display = None if f['kind'] == 'Measure' or f.get('date_of') or f['label'] == f['name'] else f['label']
                projs.append(proj(self.field_expr(f, v), display, active=role in ('Category', 'Rows')))
            if projs:
                qs[role] = {'projections': projs}
        query = {'queryState': qs}
        objects, vco = {}, {'title': self.title_vco(v), 'padding': padding(2)}
        if vt in ('clusteredBarChart', 'clusteredColumnChart', 'barChart', 'columnChart', 'lineChart', 'areaChart', 'stackedAreaChart'):
            cat = (v['roles'].get('Category') or [None])[0]
            if cat and cat['kind'] == 'Column':
                sb = v.get('sort_by') or {}
                by = next((f for f in v['roles'].get('Y', []) if f['name'] == sb.get('measure')), None) if sb.get('measure') else None
                if by:  # Tableau "[Dim] DESC by [Measure]": sort the bars by the plotted measure
                    query['sortDefinition'] = {'sort': [{'field': self.field_expr(by, v), 'direction': sb['direction']}], 'isDefaultSort': False}
                else:
                    if sb.get('measure'):
                        self.log.append(f'{p["name"]} / {v["sheet"]}: Tableau sorts by [{sb["measure"]}], which is not on the visual; '
                                        f'sorted by {cat["name"]} instead (add the sort by hand)')
                    query['sortDefinition'] = {'sort': [{'field': self.field_expr(cat), 'direction': sb.get('direction', 'Ascending')
                                                         if sb.get('category') else 'Ascending'}], 'isDefaultSort': False}
                if cat['table'] == 'Date' and cat['name'] in ('Year', 'Month Number', 'Day'):
                    objects['categoryAxis'] = [{'properties': {'axisType': sq('Categorical')}}]  # numeric years run bottom-up otherwise
            objects['labels'] = [{'properties': {'show': boolean(bool(v.get('data_labels')))}}]
            if v['roles'].get('Series'):
                objects['legend'] = [{'properties': {'show': boolean(True), 'position': sq('Top')}}]
            fills = self.colour_entries(v)
            if fills:
                objects['dataPoint'] = fills
            refs = []
            for i, r in enumerate(v.get('reference_lines', [])):
                mname = self.refline_measures.get((v['id'], r['measure']))
                if not mname:
                    continue
                label = re.sub(r':?\s*<Value>', '', r['label'] or '').strip() or mname
                refs.append({'properties': {'show': boolean(True), 'displayName': sq(label),
                                            'value': {'expr': meas_ref(self.measures[mname]['table'], mname)},
                                            'style': sq('dotted'), 'lineColor': colour('#000000' if i == 0 else '#F28E2B'),
                                            'dataLabelShow': boolean(True), 'dataLabelText': sq('ValueAndName')},
                             'selector': {'id': str(i)}})
            if refs:
                objects['y1AxisReferenceLine'] = refs
        if vt in ('tableEx', 'pivotTable'):
            cols_all = [f for r in ('Values', 'Rows') for f in v['roles'].get(r, [])]
            if vt == 'tableEx' and len(cols_all) >= 7 and pos.get('width'):
                # many columns: fixed widths that fit the Tableau box (long text columns get more room, and wrap)
                weight = lambda f: 3 if re.search(r'descr|comment|note', f['name'], re.I) else 1.5 if re.search(r'name|title', f['name'], re.I) else 1
                total = sum(weight(f) for f in cols_all)
                usable = pos['width'] - 30
                objects['columnHeaders'] = [{'properties': {'autoSizeColumnWidth': boolean(False), 'customColumnWidth': boolean(True),
                                                            'wordWrap': boolean(True), 'bold': boolean(True), 'fontSize': num(9)}}]
                objects['columnWidth'] = [{'properties': {'value': num(round(usable * weight(f) / total))},
                                           'selector': {'metadata': proj(self.field_expr(f, v))['queryRef']}} for f in cols_all]
            else:
                objects['columnHeaders'] = [{'properties': {'columnAdjustment': sq('growToFit'), 'autoSizeColumnWidth': boolean(True),
                                                            'bold': boolean(True), 'fontSize': num(9)}}]
            objects['values'] = [{'properties': {'fontSize': num(9)}}]
            if vt == 'tableEx':  # pivotTable has no total.totals (PBIR_FORMATTING_PROP_UNKNOWN); it uses subTotals below
                objects['total'] = [{'properties': {'totals': boolean(False)}}]
            vco['stylePreset'] = [{'properties': {'name': sq('None')}}]
            ua = v.get('url_action')
            link_field = next((f for f in v['roles'].get('Values', []) if ua and f['name'] == ua.get('link_on')), None) or \
                (v['roles'].get('Values') or [None])[0] if ua and vt == 'tableEx' and ua.get('link_on') else None
            if link_field:
                objects['values'].append({'properties': {'webURL': {'expr': {'Aggregation': {'Expression': col_ref(ua['table'], ua['column']), 'Function': 3}}}},
                                          'selector': {'data': [{'dataViewWildcard': {'matchingOption': 1}}],
                                                       'metadata': f'{link_field["table"]}.{link_field["name"]}'}})
            if vt == 'pivotTable':
                objects['rowHeaders'] = [{'properties': {'stepped': boolean(False), 'showExpandCollapseButtons': boolean(False)}}]
                objects['subTotals'] = [{'properties': {'rowSubtotals': boolean(False), 'columnSubtotals': boolean(False)}}]
        seen = {}
        for vf in v.get('visual_filters', []):
            key = f'{p["name"]}:{v["id"]}:filter:{vf["field"]["name"]}'
            seen[key] = seen.get(key, 0) + 1  # two filters on one field (sheet + shared): names must stay unique
            flt = self.categorical_filter(vf, key if seen[key] == 1 else f'{key}:{seen[key]}')
            if flt:
                query.setdefault('_filters', []).append(flt)
        filters = query.pop('_filters', None)
        out = {'name': name, 'position': pos, 'visual': {'visualType': vt, 'query': query, 'objects': objects,
                                                         'visualContainerObjects': vco}}
        if filters:
            out['filterConfig'] = {'filters': filters}
        return out

    def split_card(self, vis):
        """A card with several measures is split into one card per measure, stacked inside the original box.
        Power BI lays several measures of one card side by side and clips each to a few characters (its layout
        settings did not change that), so each measure gets its own full-width card instead."""
        v = vis.get('visual') or {}
        data = ((v.get('query') or {}).get('queryState') or {}).get('Data') or {}
        projs = data.get('projections') or []
        if v.get('visualType') != 'cardVisual' or len(projs) < 2 or len(v['query']['queryState']) != 1:
            return [vis]
        n = len(projs)
        pos = vis['position']
        inset = 6  # keep clear of the header text above and the element below the card
        if pos['height'] - 2 * inset < 24 * n:  # a short box: no inset, and never grow past the original box
            inset = 0
        top, total = pos['y'] + inset, pos['height'] - 2 * inset
        h = max(total // n, 12)
        out = []
        for i, pr in enumerate(projs):
            part = copy.deepcopy(vis)
            part['name'] = hid(f'{vis["name"]}:measure:{i}')
            part['annotations'] = [{'name': 'bhayankarCardGroup', 'value': vis['name']}]
            part['visual']['query']['queryState']['Data']['projections'] = [pr]
            # the default card font is sized for a tall card and is cut off in a short one: size the text to the height
            fs = max(7, min(16, int(h * 0.28)))
            part['visual'].setdefault('objects', {}).update({
                'value': [{'properties': {'fontSize': num(fs), 'bold': boolean(True)}, 'selector': {'id': 'default'}}],
                'label': [{'properties': {'fontSize': num(8)}, 'selector': {'id': 'default'}}]})
            part['visual'].setdefault('visualContainerObjects', {})['padding'] = padding(0)
            if i:  # the title stays on the first card only
                part['visual']['visualContainerObjects']['title'] = [{'properties': {'show': boolean(False)}}]
            part['position'].update({'y': top + i * h, 'height': h, 'z': pos.get('z', 0) + i, 'tabOrder': pos.get('tabOrder', 0) + i})
            if i and part.get('filterConfig'):  # filter names must be unique across the report
                for flt in part['filterConfig'].get('filters', []):
                    flt['name'] = fid(f'{part["name"]}:{flt["name"]}')
            out.append(part)
        self.log.append(f'{vis["name"]}: card with {n} measures split into {n} stacked cards')
        return out

    def sweep_empty_visual_dirs(self):
        """Remove visual folders without a visual.json. An earlier build's folders can outlive their files for a
        moment while a sync client or Power BI Desktop still holds them, so removal is retried."""
        for _ in range(5):
            left = 0
            for dp, dns, fns in os.walk(os.path.join(self.report_dir, 'definition', 'pages')):
                if os.path.basename(os.path.dirname(dp)) == 'visuals' and 'visual.json' not in fns:
                    shutil.rmtree(dp, ignore_errors=True)
                    left += os.path.isdir(dp)
            if not left:
                return
            time.sleep(1)
        self.log.append('some empty visual folders from an earlier build could not be removed (held by another program): delete them when it is closed')

    def col_type(self, table, name):
        if table == 'Date':
            return 'integer' if name in ('Year', 'Month Number', 'Day') else 'date' if name == 'Date' else 'string'
        t = self.tables.get(table) or {}
        c = next((c for c in t.get('columns', []) + t.get('calc_columns', []) if c['name'] == name), None)
        return (c or {}).get('datatype') or 'string'

    def pbir_literal(self, value, dtype):
        """Filter value as a PBIR literal: 225L for whole numbers, 1.5D for decimals, 'text' otherwise."""
        if value.strip() == '%null%':
            return 'null'  # Tableau's Null member: a text literal here made Power BI reject the whole visual query
        if dtype == 'integer' and re.fullmatch(r'-?\d+', value):
            return f'{value}L'
        if dtype == 'real' and re.fullmatch(r'-?\d+(\.\d+)?', value):
            return f'{value}D'
        if dtype == 'boolean' and value.lower() in ('true', 'false'):
            return value.lower()
        return "'" + value.replace("'", "''") + "'"

    def range_literal(self, v, dtype):
        """Tableau range end (#2021-09-01 15:21:59#, 12, 1.5) -> PBIR literal, or None when it cannot be typed."""
        m = re.fullmatch(r'#(\d{4}-\d{2}-\d{2})(?:[ T](\d{2}:\d{2}:\d{2}))?#', v)
        if m:
            return f"datetime'{m.group(1)}T{m.group(2) or '00:00:00'}'"
        if re.fullmatch(r'-?\d+', v):
            return f'{v}L' if dtype not in ('real', 'double') else f'{v}D'
        if re.fullmatch(r'-?\d+\.\d+', v):
            return f'{v}D'
        return None

    @staticmethod
    def relative_bounds(f):
        """Tableau relative date (first/last period of a unit, today = 0) -> PBIR Between bounds. DateSpan widens a
        bound to its whole period, so -3..0 years is 1 Jan three years ago .. 31 Dec this year (Tableau's
        "last 3 years" with include-future). Without include-future the upper bound stops at today."""
        unit = REL_UNITS[f['unit']]

        def at(n):
            return {'Now': {}} if n == 0 else {'DateAdd': {'Expression': {'Now': {}}, 'Amount': n, 'TimeUnit': unit}}
        lower = {'DateSpan': {'Expression': at(f['first']), 'TimeUnit': unit}}
        if f['last'] == 0 and not f.get('include_future'):
            upper = {'DateSpan': {'Expression': {'Now': {}}, 'TimeUnit': 0}}
        else:
            upper = {'DateSpan': {'Expression': at(f['last']), 'TimeUnit': unit}}
        return lower, upper

    def categorical_filter(self, f, key):
        """One Tableau sheet / page filter -> PBIR filter: Categorical (members), Advanced (blank, range, or a
        comparison on a measure) or RelativeDate. None (logged) when it cannot be expressed."""
        kind = f.get('type')
        if kind == 'All':
            return None  # Tableau range over every value, nulls included: no filter
        alias = 'f'
        table, name = f['field'].get('table'), f['field']['name']
        is_measure = f['field'].get('kind') == 'Measure'
        if is_measure:
            name = self.rename(name)
            table = (self.measures.get(name) or {}).get('table') or table
        src = {('Measure' if is_measure else 'Column'): {'Expression': {'SourceRef': {'Source': alias}}, 'Property': name}}
        dtype = 'real' if is_measure else self.col_type(table, name)
        ftype, cond, why = 'Advanced', None, None
        if kind in ('NotBlank', 'Blank'):  # Tableau "exclude Null" / "non-null values": Advanced filter "is not blank"
            cond = {'Comparison': {'ComparisonKind': 0, 'Left': src, 'Right': {'Literal': {'Value': 'null'}}}}
            cond = {'Not': {'Expression': cond}} if kind == 'NotBlank' else cond
        elif kind == 'Range':
            lo = self.range_literal(f['lo'], dtype) if f.get('lo') is not None else None
            hi = self.range_literal(f['hi'], dtype) if f.get('hi') is not None else None
            if (f.get('lo') is not None and lo is None) or (f.get('hi') is not None and hi is None):
                why = 'range end is neither a date nor a number'
            else:
                parts = ([{'Comparison': {'ComparisonKind': 2, 'Left': src, 'Right': {'Literal': {'Value': lo}}}}] if lo else []) + \
                        ([{'Comparison': {'ComparisonKind': 4, 'Left': src, 'Right': {'Literal': {'Value': hi}}}}] if hi else [])
                if not parts:
                    return None
                cond = parts[0] if len(parts) == 1 else {'And': {'Left': parts[0], 'Right': parts[1]}}
        elif kind == 'RelativeDate' and f.get('unit') in REL_UNITS and not is_measure:
            lower, upper = self.relative_bounds(f)
            cond = {'Between': {'Expression': src, 'LowerBound': lower, 'UpperBound': upper}}
            ftype = 'RelativeDate'
        elif kind == 'Categorical' and f.get('values') and is_measure:
            # Tableau filter on an aggregate calc (one member): a visual-level measure comparison
            if len(f['values']) == 1 and re.fullmatch(r'-?\d+(\.\d+)?', f['values'][0]):
                cond = {'Comparison': {'ComparisonKind': 0, 'Left': src, 'Right': {'Literal': {'Value': f'{f["values"][0]}D'}}}}
                cond = {'Not': {'Expression': cond}} if f.get('inverted') else cond
            else:
                why = 'member filter on a measure (only one number can be compared)'
        elif kind == 'Categorical' and f.get('values'):
            ftype = 'Categorical'
            cond = {'In': {'Expressions': [src], 'Values': [[{'Literal': {'Value': self.pbir_literal(x, dtype)}}] for x in f['values']]}}
            if f.get('inverted'):
                cond = {'Not': {'Expression': cond}}
        elif kind != 'RowCondition':
            why = f.get('note') or kind
        if cond is None:
            if kind != 'RowCondition':
                self.log.append(f'filter on {name} ({kind}) not generated: {f.get("tableau")}' + (f' [{why}]' if why else ''))
            return None
        field = (meas_ref if is_measure else col_ref)(table, name)
        return {'name': fid(key), 'field': field, 'type': ftype,
                'filter': {'Version': 2, 'From': [{'Name': alias, 'Entity': table, 'Type': 0}], 'Where': [{'Condition': cond}]}}

    def slicer(self, p, s):
        pos = dict(s.get('position') or {'x': 0, 'y': 0, 'width': 160, 'height': 64})
        if s['visualType'] == 'slicer' and (s['mode'] or 'Dropdown') in ('Dropdown', 'Between', 'Before', 'After'):
            # slicers.md sizing: header 28 + selector 32 + padding 8/8 = 76, snapped to 80
            if pos['height'] < 80 and not s.get('hidden'):
                self.growth.append((pos['y'] + pos['height'], 80 - pos['height']))
            pos['height'] = max(pos['height'], 80)
        if s['mode'] == 'Between':
            pos['width'] = max(pos['width'], 216)
        f = s['field']
        ptab = param_table_name(s.get('parameter') or '')
        if s.get('parameter') and ptab in self.param_tables:  # parameter slicer: bind to its disconnected table
            f = {**f, 'kind': 'Column', 'table': ptab, 'name': self.param_tables[ptab][1]}
        if not f.get('table') or f.get('kind') not in ('Column', None):
            # e.g. a field from a published data source with no .tds in input\, or a free-entry parameter
            self.log.append(f'{p["name"]}: slicer "{f.get("label") or f.get("name")}" not generated: its field has no table in the model '
                            '(add the published .tds to input\\, or add the slicer by hand)')
            return None
        data = {'mode': sq(s['mode'] or 'Dropdown')}
        objects = {'data': [{'properties': data}],
                   'header': [{'properties': {'show': boolean(True), 'text': sq(re.sub(r'^Param:', '', f['label'])), 'textSize': num(9),
                                              'fontColor': colour(DARK)}}],
                   'items': [{'properties': {'textSize': num(9)}}]}
        alias = 's'
        src = {'Column': {'Expression': {'SourceRef': {'Source': alias}}, 'Property': f['name']}}
        cond = None
        decision = (self.decisions.get('slicer_defaults') or {}).get(f['label']) or (self.decisions.get('slicer_defaults') or {}).get(f['name'])
        d = s.get('default') or {}
        if decision and decision.get('exclude'):
            data['isInvertedSelectionMode'] = boolean(True)
            cond = {'Not': {'Expression': {'In': {'Expressions': [src], 'Values': [[{'Literal': {'Value': "'" + x + "'"}}] for x in decision['exclude']]}}}}
        elif decision and decision.get('include'):
            cond = {'In': {'Expressions': [src], 'Values': [[{'Literal': {'Value': "'" + x + "'"}}] for x in decision['include']]}}
        elif s['mode'] == 'Between' and d.get('from'):
            a, b = tableau_date(d['from']), tableau_date(d['to'])
            if a and b:
                sa, sb = f"datetime'{a[0]:04d}-{a[1]:02d}-{a[2]:02d}T00:00:00'", f"datetime'{b[0]:04d}-{b[1]:02d}-{b[2]:02d}T00:00:00'"
                data['startDate'] = lit(sa)
                data['endDate'] = lit(sb)
                cond = {'And': {'Left': {'Comparison': {'ComparisonKind': 2, 'Left': src, 'Right': {'Literal': {'Value': sa}}}},
                                'Right': {'Comparison': {'ComparisonKind': 4, 'Left': src, 'Right': {'Literal': {'Value': sb}}}}}}
        elif d.get('values') and not d.get('note'):
            # Tableau keeps string parameter values quoted ('"Current Year"'); the slicer table holds them bare
            vals = [x[1:-1].replace('""', '"') if len(x) > 1 and x[0] == x[-1] == '"' else x
                    for x in d['values'] if x and '..' not in x]
            if vals:
                cond = {'In': {'Expressions': [src], 'Values': [[{'Literal': {'Value': "'" + x.replace("'", "''") + "'"}}] for x in vals]}}
                if d.get('inverted'):
                    cond = {'Not': {'Expression': cond}}
        elif d.get('note') and not decision:
            self.log.append(f'{p["name"]}: slicer {f["label"]} has an unanswered decision; no default selection written')
        if cond:
            objects['general'] = [{'properties': {'filter': {'filter': {'Version': 2, 'From': [{'Name': alias, 'Entity': f['table'], 'Type': 0}],
                                                                        'Where': [{'Condition': cond}]}}}}]
        vis = {'name': hid(f'{p["name"]}:slicer:{s["id"]}:{f["name"]}'), 'position': pos,
               'visual': {'visualType': s['visualType'] if s['visualType'] in ('slicer', 'listSlicer', 'textSlicer', 'advancedSlicerVisual') else 'slicer',
                          'query': {'queryState': {'Values': {'projections': [proj(col_ref(f['table'], f['name']))]}}},
                          'objects': objects,
                          'syncGroup': {'groupName': re.sub(r'[^A-Za-z0-9]', '', f['table'] + f['name'])[:60], 'fieldChanges': True, 'filterChanges': True},
                          'visualContainerObjects': {'padding': padding(0)}}}
        if s['visualType'] != 'slicer':
            vis['visual']['objects'].pop('data', None)
        if s.get('hidden'):
            vis['isHidden'] = True
        return vis

    def static(self, p, st, report):
        pos = dict(st.get('position') or {'x': 0, 'y': 0, 'width': 200, 'height': 40})
        if st['type'] == 'image' and st.get('image'):
            src = self.find_image(st['image'])
            if not src:
                self.log.append(f'image {st["image"]} not found next to the workbook in input\\ (copy the Image folder there)')
                return None
            nm = re.sub(r'[^A-Za-z0-9.]', '', os.path.basename(src))
            self.images[src] = nm
            return {'name': hid(f'{p["name"]}:image:{st["image"]}'), 'position': pos, 'visual': {
                'visualType': 'image',
                'objects': {'general': [{'properties': {'imageUrl': {'expr': {'ResourcePackageItem': {'PackageName': 'RegisteredResources', 'PackageType': 1, 'ItemName': nm}}}}}]},
                'visualContainerObjects': {'padding': padding(0)}}}
        if st['type'] == 'textbox' and st.get('text'):
            return {'name': hid(f'{p["name"]}:text:{st["text"][:40]}'), 'position': pos, 'visual': {
                'visualType': 'textbox',
                'objects': {'general': [{'properties': {'paragraphs': [{'textRuns': [{'value': st['text'], 'textStyle': {
                    'fontFamily': 'Segoe UI', 'fontSize': '18pt', 'color': DARK}}], 'horizontalTextAlignment': 'center'}]}}]},
                'visualContainerObjects': {'background': [{'properties': {'show': boolean(False)}}], 'padding': padding(0)}}}
        self.log.append(f'{p["name"]}: static {st["type"]} not generated')
        return None

    def find_image(self, rel):
        twb = next((os.path.join(dp, fn) for dp, _, fs in os.walk(INPUT_DIR) for fn in fs
                    if os.path.splitext(fn)[0] == self.wb and fn.lower().endswith(('.twb', '.twbx'))), None)
        if not twb:
            return None
        if zipfile.is_zipfile(twb):
            with zipfile.ZipFile(twb) as z:
                n = next((n for n in z.namelist() if n.replace('\\', '/').endswith(rel.replace('\\', '/'))), None)
                if n:
                    dest = os.path.join(self.wb_out, '_assets', os.path.basename(n))
                    os.makedirs(os.path.dirname(dest), exist_ok=True)
                    with open(dest, 'wb') as f:
                        f.write(z.read(n))
                    return dest
            return None
        cand = os.path.join(os.path.dirname(twb), rel.replace('/', os.sep))
        return cand if os.path.exists(cand) else None

    # =====================================================================================
    def run(self):
        os.makedirs(self.pbip_dir, exist_ok=True)
        self.build_model()
        self.write_model()
        self.write_report()
        pbip = os.path.join(self.pbip_dir, self.wb + '.pbip')
        self.sweep_empty_visual_dirs()
        res = bv.validate_report(self.report_dir, pbip_path=pbip)
        # the PBIR validator does not read the model: what makes Desktop refuse it (name clashes, unknown
        # references or functions, circular measures, broken relationships) is checked here
        self.model_check = bmc.check(self.pbip_dir, self.report_dir)
        self.data_request()
        second = self.cli_check() if os.environ.get('BHAYANKAR_CLI_CHECK') else None
        with open(os.path.join(self.wb_out, 'bhayankar_build_log.md'), 'w', encoding='utf-8') as f:
            f.write(f'# Build log: {self.wb}\n\nBuilt by {bt.banner()}.\n\nProject: `{pbip}`\n\n'
                    f'## Validation ({bv.VALIDATOR})\n\n```json\n{json.dumps(res, indent=1)[:6000]}\n```\n\n'
                    f'## Model check (bhayankar_modelcheck: the semantic model, offline)\n\n{bmc.render(self.model_check)}\n\n'
                    + (f'## Second opinion (powerbi-report-author)\n\n```json\n{json.dumps(second, indent=1)[:3000]}\n```\n\n'
                       if second else '')
                    + '## Notes\n\n' + '\n'.join(f'- {x}' for x in self.log) + '\n')
        return res

    def source_columns(self, t):
        """(model column, CSV header) of every column a table reads from its source, in model order."""
        groups = {g['name'] for g in self.digest.get('groups', [])}
        out = []
        if t.get('m_lines'):  # built from other tables (blend shared dimension): reads no source of its own
            return out
        for c in t['columns']:
            if not c.get('source_column') or c.get('derived') or c.get('dataCategory') or c['name'].endswith(' (Date)') \
                    or c['name'] in groups:
                continue
            h = self.csv_name(t, c)
            if h not in [x for _, x in out]:
                out.append((c['name'], h))
        return out

    def data_request(self):
        """output\\<wb>\\DATA_REQUEST.md: exactly which file and columns each table needs (sample CSV name, CSV headers,
        warehouse table), checked against the sample folder when there is one; plus header-only CSV templates in
        data_request\\ to hand to whoever extracts the data. Headers only: the toolkit never makes up rows."""
        folder = self.sources.get('sample_folder') if self.sources.get('mode') == 'sample' else None
        tdir = os.path.join(self.wb_out, 'data_request')
        os.makedirs(tdir, exist_ok=True)
        L = [f'# Data request: {self.wb}', '',
             'What the Power BI model reads. Give the data team this file and the header-only templates in `data_request\\`.',
             '', '- **Sample mode:** one CSV per table in `input\\' + f'{self.wb}_Data\\` (or the `sources.json` `sample_folder`), '
             'UTF-8, comma-separated, first row = these headers exactly.',
             '- **Warehouse mode:** the same columns from the warehouse table (`WarehouseSchema`.`table`).', '']
        if folder:
            L.append(f'Checked against the sample folder `{folder}`.' if os.path.isdir(folder) else
                     f'**The sample folder `{folder}` does not exist.**')
            L.append('')
        problems = 0
        for t in self.tables.values():
            cols = self.source_columns(t)
            if not cols:
                continue
            cfg = self.sources.get('tables', {}).get(t['source']) or {}
            fname = cfg.get('file') or f'{t["source"]}.csv'
            have = None
            if folder and os.path.isdir(folder):
                path = os.path.join(folder, fname)
                have = set(self.csv_header(t['source'])) if os.path.exists(path) else False
            status = '' if have is None else (' **missing file**' if have is False else '')
            L += [f'## {t["name"]}', '', f'- CSV file: `{fname}`{status}', f'- Warehouse table: `{t["source"]}`',
                  f'- Columns ({len(cols)}):', '']
            check = have is not None
            L += ['| CSV header | Model column |' + (' In the sample |' if check else ''),
                  '|---|---|' + ('---|' if check else '')]
            for name, h in cols:
                ok = '' if not have else ('yes' if h in have else '**missing**')
                problems += ok == '**missing**'
                L.append(f'| `{h}` | {name if name != h else ""} |' + (f' {ok or "(no file)"} |' if check else ''))
            if have:
                extra = sorted(have - {h for _, h in cols})
                if extra:
                    L += ['', f'Extra columns in the CSV (not used, fine): {", ".join(f"`{x}`" for x in extra[:30])}']
            problems += have is False
            L.append('')
            with open(os.path.join(tdir, fname), 'w', encoding='utf-8', newline='') as f:
                csv.writer(f).writerow([h for _, h in cols])
        if folder:
            L.insert(6, f'**{problems} problem(s) found** (missing files or headers): fix them before Refresh.\n'
                        if problems else 'Every file and header is present.\n')
            if problems:
                self.log.append(f'DATA REQUEST: {problems} missing sample file(s)/header(s): see DATA_REQUEST.md')
        with open(os.path.join(self.wb_out, 'DATA_REQUEST.md'), 'w', encoding='utf-8') as f:
            f.write('\n'.join(L) + '\n')

    def cli_check(self):
        """Optional (BHAYANKAR_CLI_CHECK=1): the Node.js CLI's own verdict, for comparison. Never needed to build."""
        try:
            r = subprocess.run(['powerbi-report-author', 'validate', self.report_dir], capture_output=True, text=True,
                               shell=os.name == 'nt', encoding='utf-8', timeout=300)
            return json.loads(r.stdout[r.stdout.find('{'):])['data']
        except (OSError, ValueError, KeyError, subprocess.TimeoutExpired) as e:
            return {'result': 'not run', 'reason': str(e)[:300] or 'powerbi-report-author is not installed'}


def ensure_extracted(names):
    """Standalone build: extract a workbook from input\\ first when its rebuild spec is missing or older than the
    workbook, a .tds or table_map.json. Returns the workbook names that have a spec to build."""
    import bhayankar_extractor as bx
    wbs, tds = {}, []
    for dp, dirs, fs in os.walk(INPUT_DIR) if os.path.isdir(INPUT_DIR) else []:
        for fn in fs:
            if fn.startswith(('~$', '.')):
                continue
            if fn.lower().endswith(('.twb', '.twbx')):
                wbs.setdefault(os.path.splitext(fn)[0], os.path.join(dp, fn))
            elif fn.lower().endswith(('.tds', '.tdsx')):
                tds.append(os.path.join(dp, fn))
    tmap_file = os.path.join(INPUT_DIR, 'table_map.json')
    tmap = load_json(tmap_file, {}) or {}
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    if not names:  # every workbook in input\ plus any already extracted
        names = sorted(set(wbs) | {d for d in os.listdir(OUTPUT_DIR)
                                   if os.path.exists(os.path.join(OUTPUT_DIR, d, 'bhayankar_rebuild.json'))})
    ready = []
    for n in names:
        spec = os.path.join(OUTPUT_DIR, n, 'bhayankar_rebuild.json')
        src = wbs.get(n)
        if src:
            newest = max(os.path.getmtime(p) for p in [src, *tds, *([tmap_file] if os.path.exists(tmap_file) else [])])
            if not os.path.exists(spec) or os.path.getmtime(spec) < newest:
                print(f'{n}: {"no extract yet" if not os.path.exists(spec) else "input changed since the last extract"}; '
                      f'running the extractor first')
                try:
                    bx.process(src, tds, os.path.join(OUTPUT_DIR, n), None, tmap, strict_tds=True)
                except (Exception, SystemExit) as e:  # keep the batch going
                    print(f'{n}: extract failed: {e}')
                    continue
        if os.path.exists(spec):
            ready.append(n)
        else:
            print(f'{n}: not found. Put "{n}.twb" (or .twbx) in {INPUT_DIR} and run again.')
    return ready


def main():
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    print(f'{bt.banner()}: build (offline)')
    names = ensure_extracted(sys.argv[1:])
    if not names:
        print(f'Nothing to build: put a Tableau workbook (.twb/.twbx) in {INPUT_DIR} and run again.')
    for n in names:
        b = Build(n)
        res = b.run()
        where = f'pbip\\{n}.pbip' if os.path.normcase(b.pbip_dir) == os.path.normcase(os.path.join(b.wb_out, 'pbip')) \
            else os.path.join(b.pbip_dir, f'{n}.pbip')  # moved to a short folder (path length) or decisions.json pbip_folder
        print(f'{n}: {where} written; validate: {res.get("result")} '
              f'({res.get("errorCount", "?")} errors, {res.get("warningCount", "?")} warnings); {bmc.summary(b.model_check)}; '
              f'notes: {len(b.log)} (bhayankar_build_log.md)')


if __name__ == '__main__':
    main()
