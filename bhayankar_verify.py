#!/usr/bin/env python3
r"""Bhayankar verify: checks a built project inside Power BI Desktop's own engine. Offline, nothing to install.

Power BI Desktop runs a local Analysis Services engine for every open file. This tool finds the one that holds your
project and, through the query library Desktop ships with (Microsoft.PowerBI.AdomdClient.dll, used from Windows
PowerShell), checks:
  1. Model health   every measure, column and table partition loads without an error (INFO.MEASURES/COLUMNS/...)
  2. Data loaded    rows per table (0 rows everywhere means: click Refresh in Desktop, then run again)
  3. Measures       every measure evaluates; the value and the time it took
  4. Numbers        per Tableau sheet, the tester's query runs in the engine and is compared with what the Tableau
                    logic gives on the same source data (bhayankar_oracle). Sheets whose Tableau filters live in
                    Power BI slicers or visual filters are listed, not compared (the query does not apply them).
Writes output\<workbook>\verify\: VERIFY_REPORT.md (read this), verify_results.json, engine_exports\<sheet>.csv and
test_report.xlsx (the tester's report with the number comparisons filled in).

Usage:
  1. Open output\<workbook>\pbip\<workbook>.pbip in Power BI Desktop and click Refresh.
  2. python bhayankar_verify.py "<workbook>"          (Run_Verify.bat)
  python bhayankar_verify.py --list                    models open in Power BI Desktop (port, tables)
  python bhayankar_verify.py "<workbook>" --open       open the project in Desktop first and wait for it to load
  python bhayankar_verify.py "<workbook>" --port N     use this engine (see --list)
  python bhayankar_verify.py "<workbook>" --no-numbers model checks only
Read only: it never changes the model or any file of the project (--refresh only loads data into the open Desktop session). Python 3 standard library + Windows PowerShell.
"""
import argparse
import base64
import csv
import glob
import json
import os
import re
import socket
import subprocess
import sys
import tempfile
import time
import types

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import bhayankar_toolkit as bt  # noqa: E402

INPUT_DIR = bt.input_dir()
OUTPUT_DIR = bt.output_dir()
WORKSPACE_ROOTS = [
    os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Microsoft', 'Power BI Desktop', 'AnalysisServicesWorkspaces'),
    os.path.join(os.environ.get('USERPROFILE', ''), 'Microsoft', 'Power BI Desktop Store App', 'AnalysisServicesWorkspaces'),
]
DESKTOP_BINS = [
    os.path.join(os.environ.get('ProgramFiles', r'C:\Program Files'), 'Microsoft Power BI Desktop', 'bin'),
    os.path.join(os.environ.get('ProgramFiles(x86)', r'C:\Program Files (x86)'), 'Microsoft Power BI Desktop', 'bin'),
]
ADOMD_DLL = 'Microsoft.PowerBI.AdomdClient.dll'
SYSTEM_TABLE = re.compile(r'^(DateTableTemplate_|LocalDateTable_)')
# TOM ObjectState: 1 Ready, 3 NoData and 4 CalculationNeeded (nothing refreshed yet: not an error), 5 SemanticError,
# 6 EvaluationError, 7 DependencyError, 8 Incomplete, 9 SyntaxError
OK_STATES = {1, 3, 4}

PS_SCRIPT = r'''
$ErrorActionPreference = 'Stop'
[void][Reflection.Assembly]::LoadFrom($DllPath)
Add-Type -AssemblyName System.Web.Extensions
Add-Type -AssemblyName System.Web
$ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$ser.MaxJsonLength = [int]::MaxValue
$req = $ser.DeserializeObject([IO.File]::ReadAllText($RequestFile, [Text.Encoding]::UTF8))
$inv = [Globalization.CultureInfo]::InvariantCulture
# JSON writer: values that PowerShell adds to .NET lists are stored wrapped, which JavaScriptSerializer cannot write;
# PowerShell's own -is tests see through the wrapper
function Write-Json($v, [Text.StringBuilder]$sb) {
    if ($null -eq $v -or $v -is [DBNull]) { [void]$sb.Append('null') }
    elseif ($v -is [string]) { [void]$sb.Append([Web.HttpUtility]::JavaScriptStringEncode($v, $true)) }
    elseif ($v -is [bool]) { [void]$sb.Append($(if ($v) { 'true' } else { 'false' })) }
    elseif ($v -is [datetime]) { [void]$sb.Append('"' + $v.ToString('yyyy-MM-ddTHH:mm:ss', $inv) + '"') }
    elseif ($v -is [double] -or $v -is [single] -or $v -is [decimal]) {
        $d = [double]$v
        if ([double]::IsNaN($d) -or [double]::IsInfinity($d)) { [void]$sb.Append('null') } else { [void]$sb.Append($d.ToString('R', $inv)) }
    }
    elseif ($v -is [int] -or $v -is [long] -or $v -is [int16] -or $v -is [byte] -or $v -is [uint32] -or $v -is [uint64]) {
        [void]$sb.Append(([decimal]$v).ToString($inv))
    }
    elseif ($v -is [Collections.IDictionary]) {
        [void]$sb.Append('{')
        $first = $true
        foreach ($k in $v.Keys) {
            if (-not $first) { [void]$sb.Append(',') }
            $first = $false
            [void]$sb.Append([Web.HttpUtility]::JavaScriptStringEncode([string]$k, $true))
            [void]$sb.Append(':')
            Write-Json $v[$k] $sb
        }
        [void]$sb.Append('}')
    }
    elseif ($v -is [Collections.IEnumerable]) {
        [void]$sb.Append('[')
        $first = $true
        foreach ($x in $v) {
            if (-not $first) { [void]$sb.Append(',') }
            $first = $false
            Write-Json $x $sb
        }
        [void]$sb.Append(']')
    }
    else { [void]$sb.Append([Web.HttpUtility]::JavaScriptStringEncode([string]$v, $true)) }
}
$payload = New-Object 'System.Collections.Generic.Dictionary[string,object]'
$results = New-Object 'System.Collections.Generic.List[object]'
$conn = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdConnection ('Data Source=localhost:' + $EnginePort)
try {
    $conn.Open()
    $payload['database'] = $conn.Database
    if ($req.ContainsKey('tmsl') -and $req['tmsl']) {
        $c = $conn.CreateCommand()
        $c.CommandText = $req['tmsl']
        $c.CommandTimeout = 3600
        try { [void]$c.ExecuteNonQuery(); $payload['tmsl'] = 'ok' } catch { $payload['tmsl'] = $_.Exception.Message }
    }
    foreach ($q in $req['queries']) {
        $r = New-Object 'System.Collections.Generic.Dictionary[string,object]'
        $r['id'] = $q['id']
        $sw = [Diagnostics.Stopwatch]::StartNew()
        try {
            $cmd = $conn.CreateCommand()
            $cmd.CommandText = $q['dax']
            $cmd.CommandTimeout = [int]$q['timeout']
            $rd = $cmd.ExecuteReader()
            try {
                $cols = New-Object 'System.Collections.Generic.List[object]'
                for ($i = 0; $i -lt $rd.FieldCount; $i++) { $cols.Add($rd.GetName($i)) }
                $rows = New-Object 'System.Collections.Generic.List[object]'
                $max = [int]$q['max_rows']
                $n = 0
                while ($rd.Read()) {
                    $n++
                    if ($rows.Count -lt $max) {
                        $row = New-Object 'System.Collections.Generic.List[object]'
                        for ($i = 0; $i -lt $rd.FieldCount; $i++) {
                            $v = $rd.GetValue($i)
                            if ($v -is [DBNull]) { $v = $null }
                            elseif ($v -is [DateTime]) { $v = $v.ToString('yyyy-MM-ddTHH:mm:ss') }
                            elseif ($v -is [decimal] -or $v -is [single]) { $v = [double]$v }
                            $row.Add($v)
                        }
                        $rows.Add($row)
                    }
                }
                $r['columns'] = $cols
                $r['rows'] = $rows
                $r['row_count'] = $n
                $r['ok'] = $true
            } finally { $rd.Close() }
        } catch {
            $r['ok'] = $false
            $r['error'] = $_.Exception.Message
        }
        $r['ms'] = $sw.ElapsedMilliseconds
        $results.Add($r)
    }
} catch {
    $payload['connect_error'] = $_.Exception.Message
} finally {
    if ($conn.State -eq 'Open') { $conn.Close() }
}
$payload['results'] = $results
$sb = New-Object Text.StringBuilder
Write-Json $payload $sb
[IO.File]::WriteAllText($ResultFile, $sb.ToString(), (New-Object Text.UTF8Encoding $false))
'''


# =====================================================================================
# Desktop engines
# =====================================================================================
def read_port(path):
    with open(path, 'rb') as f:
        raw = f.read()
    text = raw.decode('utf-16-le', 'ignore') if b'\x00' in raw else raw.decode('ascii', 'ignore')
    digits = re.sub(r'\D', '', text)
    return int(digits) if digits else None


def port_open(port):
    for host in ('127.0.0.1', '::1'):
        try:
            with socket.create_connection((host, port), timeout=0.5):
                return True
        except OSError:
            continue
    return False


def engines():
    """[{port, workspace, started}] for every Power BI Desktop engine that is running, newest first."""
    out = []
    for root in WORKSPACE_ROOTS:
        for pf in glob.glob(os.path.join(root, '*', 'Data', 'msmdsrv.port.txt')):
            try:
                port = read_port(pf)
            except OSError:
                continue
            if port and port_open(port):
                out.append({'port': port, 'workspace': os.path.dirname(os.path.dirname(pf)), 'started': os.path.getmtime(pf)})
    return sorted(out, key=lambda e: -e['started'])


def adomd_dll():
    cands = [os.environ.get('BHAYANKAR_ADOMD')] + [os.path.join(b, ADOMD_DLL) for b in DESKTOP_BINS]
    for c in cands:
        if c and os.path.isfile(c):
            return c
    return None


def desktop_exe():
    for b in DESKTOP_BINS:
        p = os.path.join(b, 'PBIDesktop.exe')
        if os.path.isfile(p):
            return p
    return None


class EngineError(Exception):
    pass


def run_queries(port, queries, dll=None, tmsl=None, work=None):
    """Run DAX queries in the Desktop engine on `port`. queries: [{id, dax, timeout?, max_rows?}].
    Returns {database, results: {id: {ok, columns, rows, row_count, error, ms}}, tmsl}. Read only unless tmsl is given."""
    dll = dll or adomd_dll()
    if not dll:
        raise EngineError(f'{ADOMD_DLL} not found. It ships with Power BI Desktop in {DESKTOP_BINS[0]}; '
                          'set BHAYANKAR_ADOMD to its full path if Desktop is installed elsewhere.')
    work = work or tempfile.mkdtemp(prefix='bhy_engine_')
    os.makedirs(work, exist_ok=True)
    req_p, out_p, ps_p = (os.path.join(work, n) for n in ('engine_request.json', 'engine_result.json', 'engine_query.ps1'))
    req = {'queries': [{'id': q['id'], 'dax': q['dax'], 'timeout': q.get('timeout', 300), 'max_rows': q.get('max_rows', 100000)}
                       for q in queries]}
    if tmsl:
        req['tmsl'] = tmsl
    with open(req_p, 'w', encoding='utf-8') as f:
        json.dump(req, f)
    if os.path.exists(out_p):
        os.remove(out_p)
    with open(ps_p, 'w', encoding='utf-8-sig') as f:  # BOM: Windows PowerShell 5.1 reads the file as UTF-8
        # PowerShell variable names ignore case: keep the parameters distinct from the script's own variables
        f.write('param([string]$DllPath, [int]$EnginePort, [string]$RequestFile, [string]$ResultFile)\n' + PS_SCRIPT)
    ps = ['powershell', '-NoProfile', '-NonInteractive']
    p = subprocess.run(ps + ['-ExecutionPolicy', 'Bypass', '-File', ps_p, '-DllPath', dll, '-EnginePort', str(port),
                             '-RequestFile', req_p, '-ResultFile', out_p],
                       capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=7200)
    if not os.path.exists(out_p) and re.search(r'execution polic|not digitally signed|scripts is disabled', p.stderr or '', re.I):
        # a policy that blocks script files does not apply to an encoded command
        def lit(s):
            return "'" + str(s).replace("'", "''") + "'"
        body = (f'$DllPath = {lit(dll)}; $EnginePort = {int(port)}; $RequestFile = {lit(req_p)}; '
                f'$ResultFile = {lit(out_p)}\n' + PS_SCRIPT)
        p = subprocess.run(ps + ['-EncodedCommand', base64.b64encode(body.encode('utf-16-le')).decode('ascii')],
                           capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=7200)
    if not os.path.exists(out_p):
        raise EngineError('the PowerShell query bridge failed: ' + ((p.stderr or p.stdout or '').strip()[:600] or 'no output'))
    with open(out_p, encoding='utf-8-sig') as f:
        res = json.load(f)
    if res.get('connect_error'):
        raise EngineError(f'could not connect to the engine on port {port}: {res["connect_error"][:300]}')
    res['results'] = {r['id']: r for r in res.get('results') or []}
    return res


def cols(result):
    """Result columns without brackets/table prefix: '[Name]' / 'T[Name]' -> 'Name'."""
    return [re.sub(r'^.*\[(.*)\]$', r'\1', c) for c in result.get('columns') or []]


def records(result):
    names = cols(result)
    return [dict(zip(names, row)) for row in result.get('rows') or []]


def dq_table(t):
    return "'" + t.replace("'", "''") + "'"


def dq_name(n):
    return '[' + n.replace(']', ']]') + ']'


def model_tables(pbip):
    import bhayankar_pbip as bp
    proj = bp.load_project(pbip)
    if not proj.get('model'):
        raise SystemExit(f'cannot read the semantic model of {pbip}: ' + '; '.join(proj.get('errors') or ['not found']))
    return proj, set(proj['model']['tables'])


def find_engine(tables, port=None, work=None):
    """The open Desktop engine whose model has the project's tables: (engine, its table names, score)."""
    cands = [e for e in engines() if not port or e['port'] == port]
    if port and not cands:
        cands = [{'port': port, 'workspace': None, 'started': 0}]
    best = None
    for e in cands:
        try:
            res = run_queries(e['port'], [{'id': 't', 'dax': 'EVALUATE SELECTCOLUMNS(INFO.TABLES(), "Name", [Name])'}],
                              work=work)
        except EngineError:
            continue
        r = res['results'].get('t') or {}
        have = {row[0] for row in r.get('rows') or [] if row and not SYSTEM_TABLE.match(str(row[0]))}
        score = len(have & tables) / max(len(tables), 1)
        if not best or score > best[2]:
            best = ({**e, 'database': res.get('database')}, have, score)
    return best


# =====================================================================================
# checks
# =====================================================================================
def health(port, work):
    """Model objects with an error, and rows per table."""
    q = [{'id': k, 'dax': f'EVALUATE INFO.{k.upper()}()'} for k in ('tables', 'measures', 'columns', 'partitions')]
    res = run_queries(port, q, work=work)['results']
    for k in ('tables', 'measures', 'columns', 'partitions'):
        if not res[k].get('ok'):
            raise EngineError(f'INFO.{k.upper()}() failed: {res[k].get("error", "")[:300]} '
                              '(needs a Power BI Desktop from 2024 or later)')
    tables = {r['ID']: r['Name'] for r in records(res['tables'])}
    problems = []

    def check(kind, rs, name_key):
        for r in rs:
            err = (r.get('ErrorMessage') or '').strip()
            state = r.get('State')
            if err or (state is not None and state not in OK_STATES):
                problems.append({'kind': kind, 'table': tables.get(r.get('TableID'), ''), 'name': r.get(name_key) or r.get('Name') or '',
                                 'state': state, 'error': err})
    check('measure', records(res['measures']), 'Name')
    check('column', [r for r in records(res['columns']) if r.get('Type') != 3], 'ExplicitName')  # 3 = RowNumber
    check('partition', records(res['partitions']), 'Name')
    names = sorted(n for n in tables.values() if not SYSTEM_TABLE.match(n))
    counts = {}
    if names:
        q = [{'id': f'rows:{n}', 'dax': f'EVALUATE ROW("Rows", COUNTROWS({dq_table(n)}))'} for n in names]
        rr = run_queries(port, q, work=work)['results']
        for n in names:
            r = rr[f'rows:{n}']
            counts[n] = (r['rows'][0][0] or 0) if r.get('ok') and r.get('rows') else None
    measures = [(tables.get(r.get('TableID'), ''), r.get('Name')) for r in records(res['measures'])]
    return problems, counts, measures


def evaluate_measures(port, measures, work):
    q = [{'id': f'm:{i}', 'dax': f'EVALUATE ROW("Value", {dq_table(t)}{dq_name(m)})', 'timeout': 300}
         for i, (t, m) in enumerate(measures)]
    rr = run_queries(port, q, work=work)['results'] if q else {}
    out = []
    for i, (t, m) in enumerate(measures):
        r = rr.get(f'm:{i}') or {}
        out.append({'table': t, 'measure': m, 'ok': bool(r.get('ok')), 'ms': r.get('ms'),
                    'value': (r.get('rows') or [[None]])[0][0] if r.get('ok') else None, 'error': r.get('error')})
    return out


def find_workbook(name):
    for dp, _, fs in os.walk(INPUT_DIR) if os.path.isdir(INPUT_DIR) else []:
        for fn in fs:
            if os.path.splitext(fn)[0] == name and fn.lower().endswith(('.twb', '.twbx')):
                return os.path.join(dp, fn)
    return None


def sample_folder(name):
    src = None
    p = os.path.join(INPUT_DIR, 'sources.json')
    if os.path.exists(p):
        with open(p, encoding='utf-8') as f:
            src = json.load(f)
    if src and src.get('mode') == 'sample' and src.get('sample_folder'):
        folder = src['sample_folder']
        return folder if os.path.isabs(folder) else os.path.normpath(os.path.join(INPUT_DIR, folder))
    for cand in (name + '_Data', name.replace(' ', '_') + '_Data'):
        folder = os.path.join(INPUT_DIR, cand)
        if os.path.isdir(folder):
            return folder
    return None


def safe_name(s):
    return re.sub(r'[^\w\- ]', '_', s)[:80]


def table_files(folder):
    """{table: csv path} from input/sources.json "tables": {"T": {"file": "x.csv"}} (relative to the sample folder)."""
    p = os.path.join(INPUT_DIR, 'sources.json')
    if not os.path.exists(p):
        return {}
    with open(p, encoding='utf-8') as f:
        src = json.load(f)
    return {t: os.path.join(folder, cfg['file']) for t, cfg in (src.get('tables') or {}).items()
            if isinstance(cfg, dict) and cfg.get('file')}


def numbers(name, pbip, port, out_dir, work, log):
    """Tester pass 1 (oracle + queries), engine run, tester pass 2 with the engine results as Power BI exports."""
    import bhayankar_test as btst
    twb = find_workbook(name)
    data = sample_folder(name)
    if not twb:
        return {'status': 'skip', 'reason': f'no {name}.twb/.twbx in {INPUT_DIR}'}
    if not data or not os.path.isdir(data):
        return {'status': 'skip', 'reason': 'no sample data folder (sources.json sample_folder or input\\<workbook>_Data): '
                                            'the Tableau logic needs the source tables as CSV to compute expected numbers'}
    args = types.SimpleNamespace(map=None, tds=None, data=[data], tableau_exports=None, pbi_exports=None, today=None,
                                 username=None, no_data=False, no_validate=True)
    pass1 = os.path.join(work, 'pass1')
    case = btst.build_case(twb, pbip, INPUT_DIR, pass1, args)
    files = table_files(data)
    if files:  # sources.json maps a table to a differently named CSV (e.g. the --extract-data files)
        case['test_map'] = {**(case.get('test_map') or {}), 'data': {**((case.get('test_map') or {}).get('data') or {}),
                                                                      'table_files': files}}
    t1 = btst.Tester(case)
    t1.run()
    exports = os.path.join(out_dir, 'engine_exports')
    if os.path.isdir(exports):
        for fn in os.listdir(exports):
            os.remove(os.path.join(exports, fn))
    os.makedirs(exports, exist_ok=True)
    queries, compared, skipped = [], [], []
    for sheet, p in sorted(t1.dax_parts.items()):
        if not p['measures']:
            skipped.append({'sheet': sheet, 'why': 'no Power BI measure matched the Tableau measures'})
            continue
        if p['filters']:
            skipped.append({'sheet': sheet, 'why': 'Tableau filters (' + ' AND '.join(p['filters'])[:200] + ') are slicers or '
                                                  'visual filters in Power BI; the query does not apply them'})
            continue
        queries.append({'id': f'total:{sheet}', 'dax': 'EVALUATE ROW(' + ', '.join(p['measures']) + ')'})
        if p['groups']:
            queries.append({'id': f'rows:{sheet}', 'dax': 'EVALUATE SUMMARIZECOLUMNS(' + ', '.join(p['groups'] + p['measures']) + ')'})
    rr = run_queries(port, queries, work=work)['results'] if queries else {}
    for sheet, p in sorted(t1.dax_parts.items()):
        tot = rr.get(f'total:{sheet}')
        if not tot:
            continue
        if not tot.get('ok'):
            skipped.append({'sheet': sheet, 'why': 'query error: ' + (tot.get('error') or '')[:300]})
            continue
        heads = [re.sub(r"^'(.*)'\[(.*)\]$", r'\2', g) for g in p['groups']]
        labels = [re.match(r'^"((?:[^"]|"")*)"', m).group(1).replace('""', '"') for m in p['measures']]
        rows = []
        det = rr.get(f'rows:{sheet}')
        if det is not None:
            if not det.get('ok'):
                skipped.append({'sheet': sheet, 'why': 'query error: ' + (det.get('error') or '')[:300]})
                continue
            rows = det.get('rows') or []
        total_row = (tot.get('rows') or [[None] * len(labels)])[0]
        with open(os.path.join(exports, safe_name(sheet) + '.csv'), 'w', newline='', encoding='utf-8-sig') as f:
            w = csv.writer(f)
            w.writerow(heads + labels)
            for row in rows:
                w.writerow(['' if v is None else repr(v) if isinstance(v, float) else v for v in row])
            if heads:
                w.writerow(['Grand Total'] + [''] * (len(heads) - 1) + ['' if v is None else v for v in total_row])
            else:
                w.writerow(['' if v is None else v for v in total_row])
        compared.append(sheet)
    case2 = btst.build_case(twb, pbip, INPUT_DIR, out_dir,
                            types.SimpleNamespace(**{**vars(args), 'pbi_exports': exports, 'no_validate': False}))
    case2['out'] = out_dir  # write test_report.xlsx straight into verify\
    if files:
        case2['test_map'] = case['test_map']
    t2 = btst.Tester(case2)
    row = t2.run()
    data_rows = [r for r in t2.rows if r['Category'] == 'Data' and r['Area'] in ('TOTAL', 'EXPORT')]
    return {'status': 'ok', 'compared': compared, 'skipped': skipped, 'report': row.get('report'),
            'pass': sum(1 for r in data_rows if r['Status'] == 'PASS'),
            'fail': sum(1 for r in data_rows if r['Status'] == 'FAIL'),
            'manual': sum(1 for r in data_rows if r['Status'] in ('MANUAL', 'SKIP')),
            'failures': [{'test': r['Test'], 'object': r['Tableau object'], 'expected': r['Expected'], 'actual': r['Actual'],
                          'details': r['Details']} for r in data_rows if r['Status'] == 'FAIL'][:40]}


# =====================================================================================
# report
# =====================================================================================
def verdict(st):
    if st.get('error'):
        return 'NOT VERIFIED'
    if st['problems']:
        return 'MODEL ERRORS'
    empty = [t for t, n in st['counts'].items() if not n]
    if st['counts'] and len(empty) * 2 > len(st['counts']):
        return 'NOT REFRESHED'
    if any(not m['ok'] for m in st['measures']):
        return 'MEASURE ERRORS'
    n = st.get('numbers') or {}
    if n.get('status') == 'ok' and n.get('fail'):
        return 'NUMBERS DIFFER'
    return 'ENGINE CHECKS PASSED'


def render(st):
    L = [f'# Verify: {st["workbook"]}', '', f'**Verdict: {st["verdict"]}**', '',
         f'{bt.banner()}, run at {st["time"]}. Read only: nothing in the model or the project was changed.', '']
    if st.get('error'):
        L += ['## Could not verify', '', st['error'], '', 'Open the project in Power BI Desktop (`--open` does it), wait until it '
              'has loaded, click Refresh, then run Run_Verify.bat again.']
        return '\n'.join(L) + '\n'
    e = st['engine']
    L += [f'Project: `{st["project"]}`', f'Engine: Power BI Desktop on port {e["port"]} (model tables matched: '
          f'{e["match"]:.0%})', '']
    L += ['## 1. Model health', '']
    if st['problems']:
        L += ['| Kind | Table | Name | State | Error |', '|---|---|---|---|---|']
        L += [f'| {p["kind"]} | {p["table"]} | {p["name"]} | {p["state"]} | {p["error"][:200]} |' for p in st['problems']]
    else:
        L.append('No measure, column or partition reports an error.')
    L += ['', '## 2. Data loaded', '', '| Table | Rows |', '|---|---|']
    L += [f'| {t} | {"error" if n is None else f"{n:,}"} |' for t, n in st['counts'].items()]
    empty = [t for t, n in st['counts'].items() if not n]
    if st['counts'] and len(empty) * 2 > len(st['counts']):
        L += ['', f'{len(empty)} of {len(st["counts"])} tables are empty: click **Refresh** in Power BI Desktop (or run with '
              '`--refresh`), then run Run_Verify.bat again.']
    bad = [m for m in st['measures'] if not m['ok']]
    slow = sorted([m for m in st['measures'] if m['ok'] and (m['ms'] or 0) >= 2000], key=lambda m: -m['ms'])
    L += ['', '## 3. Measures', '', f'{len(st["measures"]) - len(bad)} of {len(st["measures"])} measures evaluate.']
    if bad:
        L += ['', '| Table | Measure | Error |', '|---|---|---|'] + [f'| {m["table"]} | {m["measure"]} | {(m["error"] or "")[:200]} |' for m in bad]
    if slow:
        L += ['', 'Slow (2 s or more on the whole table): ' + ', '.join(f'{m["measure"]} ({m["ms"] / 1000:.1f} s)' for m in slow[:10])]
    n = st.get('numbers') or {}
    L += ['', '## 4. Numbers vs Tableau', '']
    if n.get('status') != 'ok':
        L.append(f'Not compared: {n.get("reason", "skipped (--no-numbers)")}')
    else:
        L += [f'{n["pass"]} match, **{n["fail"]} differ**, {n["manual"]} not compared automatically. '
              f'Sheets compared: {len(n["compared"])}. Details: `{os.path.basename(n["report"] or "test_report.xlsx")}` '
              '(Data sheet).', '']
        if n['failures']:
            L += ['| Check | Sheet: field | Tableau logic | Power BI | Details |', '|---|---|---|---|---|']
            L += [f'| {f["test"]} | {f["object"]} | {f["expected"]} | {f["actual"]} | {str(f["details"])[:160]} |' for f in n['failures']]
        if n['skipped']:
            L += ['', 'Not compared automatically (check these by hand with the matching slicer selections):', '']
            L += [f'- **{s["sheet"]}**: {s["why"]}' for s in n['skipped']]
    L += ['', '## Next', '', '- A model error or a measure error: fix it in the generator or by hand, rebuild, verify again.',
          '- Numbers that differ: compare the two values with the Tableau screenshot; the expected values per sheet are in '
          '`verify\\expected_values\\`.', '- Everything passed: work through docs\\QA_CHECKLIST.md before hand-off.']
    return '\n'.join(L) + '\n'


def project_for(name):
    cand = os.path.join(OUTPUT_DIR, name, 'pbip', name + '.pbip')
    if os.path.exists(cand):
        return cand
    log = os.path.join(OUTPUT_DIR, name, 'bhayankar_build_log.md')
    if os.path.exists(log):
        with open(log, encoding='utf-8') as f:
            m = re.search(r'Project: `([^`]+)`', f.read())
        if m and os.path.exists(m.group(1)):
            return m.group(1)
    return None


def open_in_desktop(pbip, tables, wait_minutes=10):
    exe = desktop_exe()
    if not exe:
        raise EngineError('PBIDesktop.exe not found; open the project in Power BI Desktop by hand, then run again')
    before = {e['port'] for e in engines()}
    # started through PowerShell Start-Process: a detached, normal desktop launch (a direct child of this process did
    # not come up when the tool ran under Git Bash)
    def lit(s):
        return "'" + s.replace("'", "''") + "'"
    subprocess.run(['powershell', '-NoProfile', '-NonInteractive', '-Command',
                    f'Start-Process -FilePath {lit(exe)} -ArgumentList {lit(chr(34) + pbip + chr(34))}'],
                   capture_output=True, timeout=120)
    deadline = time.time() + wait_minutes * 60
    while time.time() < deadline:
        time.sleep(5)
        for e in engines():
            if e['port'] in before:
                continue
            found = find_engine(tables, port=e['port'])
            if found and found[2] >= 0.8:
                return found
    raise EngineError(f'Power BI Desktop did not finish loading the project within {wait_minutes} minutes')


def verify(name, port=None, open_first=False, with_numbers=True, refresh=False):
    out_dir = os.path.join(OUTPUT_DIR, name, 'verify')
    os.makedirs(out_dir, exist_ok=True)
    work = os.path.join(out_dir, '_work')
    st = {'workbook': name, 'time': time.strftime('%Y-%m-%d %H:%M'), 'problems': [], 'counts': {}, 'measures': []}
    try:
        pbip = project_for(name)
        if not pbip:
            raise EngineError(f'no built project for "{name}" in {OUTPUT_DIR} (run Run_Build.bat first)')
        st['project'] = pbip
        _, tables = model_tables(pbip)
        found = open_in_desktop(pbip, tables) if open_first else find_engine(tables, port, work=work)
        if not found or found[2] < 0.8:
            raise EngineError('the project is not open in Power BI Desktop (no open model has its tables). Open '
                              f'{pbip} in Desktop and click Refresh, or run with --open.')
        eng, _, score = found
        st['engine'] = {'port': eng['port'], 'workspace': eng.get('workspace'), 'database': eng.get('database'), 'match': score}
        if refresh:
            tmsl = json.dumps({'refresh': {'type': 'full', 'objects': [{'database': eng.get('database')}]}})
            st['refresh'] = run_queries(eng['port'], [], tmsl=tmsl, work=work).get('tmsl')
        st['problems'], st['counts'], measures = health(eng['port'], work)
        st['measures'] = evaluate_measures(eng['port'], measures, work)
        if with_numbers and any(st['counts'].values()):
            st['numbers'] = numbers(name, pbip, eng['port'], out_dir, work, None)
        elif with_numbers:
            st['numbers'] = {'status': 'skip', 'reason': 'the model has no rows yet (click Refresh in Desktop)'}
    except (EngineError, SystemExit) as e:
        st['error'] = str(e)
    st['verdict'] = verdict(st)
    with open(os.path.join(out_dir, 'verify_results.json'), 'w', encoding='utf-8') as f:
        json.dump(st, f, indent=1, ensure_ascii=False, default=str)
    with open(os.path.join(out_dir, 'VERIFY_REPORT.md'), 'w', encoding='utf-8') as f:
        f.write(render(st))
    return st


def main():
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('workbook', nargs='?', help='workbook name (as in output\\<workbook>\\)')
    ap.add_argument('--list', action='store_true', help='list the models open in Power BI Desktop')
    ap.add_argument('--port', type=int, help='engine port (from --list)')
    ap.add_argument('--open', action='store_true', help='open the project in Power BI Desktop first and wait for it')
    ap.add_argument('--refresh', action='store_true', help='ask the engine for a full refresh first (as the Refresh button)')
    ap.add_argument('--no-numbers', action='store_true', help='model checks only, no comparison with the Tableau logic')
    a = ap.parse_args()
    print(f'{bt.banner()}: verify in Power BI Desktop (read only)')
    if a.list:
        es = engines()
        if not es:
            print('No Power BI Desktop file is open.')
        for e in es:
            try:
                r = run_queries(e['port'], [{'id': 't', 'dax': 'EVALUATE SELECTCOLUMNS(INFO.TABLES(), "Name", [Name])'}])
                names = [row[0] for row in r['results']['t'].get('rows') or [] if not SYSTEM_TABLE.match(str(row[0]))]
                print(f'  port {e["port"]}: {len(names)} tables: {", ".join(names[:8])}{" ..." if len(names) > 8 else ""}')
            except EngineError as x:
                print(f'  port {e["port"]}: {x}')
        return
    if not a.workbook:
        ap.error('give the workbook name, e.g. python bhayankar_verify.py "My Report"')
    st = verify(a.workbook, a.port, a.open, not a.no_numbers, a.refresh)
    n = st.get('numbers') or {}
    print(f'{a.workbook}: {st["verdict"]}' + (f' ({st["error"]})' if st.get('error') else ''))
    if not st.get('error'):
        print(f'   model errors: {len(st["problems"])}; measures: {sum(m["ok"] for m in st["measures"])}/{len(st["measures"])} '
              f'evaluate; rows: {sum(v or 0 for v in st["counts"].values()):,}' +
              (f'; numbers: {n["pass"]} match / {n["fail"]} differ / {n["manual"]} not compared' if n.get('status') == 'ok' else ''))
    print(f'   report: output\\{a.workbook}\\verify\\VERIFY_REPORT.md')
    sys.exit(0 if st['verdict'] == 'ENGINE CHECKS PASSED' else 1)


if __name__ == '__main__':
    main()
