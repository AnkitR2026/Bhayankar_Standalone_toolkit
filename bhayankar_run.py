#!/usr/bin/env python3
r"""Bhayankar workflow runner: one command from a Tableau workbook to a tested Power BI project.

Runs the stages of the migration workflow in order and stops where a person must decide:

  1. extract   bhayankar_extractor  -> output\<wb>\bhayankar_digest.md and the rebuild spec
  2. gate      reads the digest flags: blockers stop the run, checks are listed for the developer
  3. build     bhayankar_build      -> output\<wb>\pbip (only when there are no blockers)
  4. test      bhayankar_test       -> output\<wb>\test (structure checks; data checks need Tableau exports)
  5. status    output\<wb>\WORKFLOW_STATUS.md (read this) and workflow_status.json

Usage (from the bhayankar folder):
  python bhayankar_run.py scan                       which workbooks in input\ are new, stale or done
  python bhayankar_run.py run                        run every workbook in input\ that is new or changed
  python bhayankar_run.py run "<workbook name>"      run one workbook (name without extension)
  python bhayankar_run.py run ... --no-build         stop after extract + gate
  python bhayankar_run.py run ... --no-test          stop after build
  python bhayankar_run.py run ... --force            run again even when nothing changed
  python bhayankar_run.py status "<workbook name>"   print its WORKFLOW_STATUS.md

Approvals live in input\decisions.json under the workbook name (read by the gate, ignored by the build):
  "workflow": {"rls_approved_by": "<Architect>", "accept_guessed_joins": true, "tier_confirmed": "Medium"}

Python 3 standard library only.
"""
import argparse
import hashlib
import io
import json
import os
import re
import subprocess
import sys
import time
from contextlib import redirect_stdout

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import bhayankar_toolkit as bt  # noqa: E402
INPUT_DIR = bt.input_dir()
OUTPUT_DIR = bt.output_dir()

STATUS_JSON = 'workflow_status.json'
STATUS_MD = 'WORKFLOW_STATUS.md'

# Flags that mean "compare this number with Tableau before sign-off" (they do not stop the build)
COMPARE_FLAGS = {
    'PUBLISHED_DS_SAMPLE': 'joins are inferred from sample data, not read from the .tds',
    'JOINED_GRAIN': 'a total may count rows differently from Tableau',
    'PHYSICAL_JOIN': 'tables are joined before Tableau aggregates',
    'UNHANDLED_XML': 'a Tableau feature this tool does not read',
    'NAME_COLLISION': 'the same calc name has different formulas',
    'DUPLICATE_LOGIC': 'the same formula under several names',
    'DATE_PARAM_VS_DATETIME': 'a date parameter is compared with a datetime field',
    'FORECAST': 'Tableau forecast numbers will not match Power BI',
    'CLUSTER': 'Tableau clustering has no DAX equivalent',
    'BLENDING': 'blended sources: built as shared dimensions over the blend links; compare totals per linked field',
    'BLEND_NO_LINK': 'a secondary source with no active link shows its grand total on every mark',
    'BLEND_LINK_CALC': 'a calculated field links blended sources: relate them by hand',
    'BLEND_LINK_TYPE': 'blend link fields of different data types: no relationship made',
    'CUSTOM_SQL': 'move custom SQL to a view or Power Query',
}
# Flags that need a design decision after the build (listed, never blocking)
DESIGN_FLAGS = {
    'PARAMETER_ACTION': 'a click sets a parameter: slicer, cross-filter or disconnected table',
    'DYNAMIC_ZONE_VISIBILITY': 'zones shown by a field or parameter: bookmarks + buttons (not built yet)',
    'HIDDEN_ZONE': 'show/hide zones: bookmarks + buttons (not built yet)',
    'STORY': 'Tableau story: one page or bookmark per point',
    'TREND_LINE': 'analytics-pane trend line',
    'QUICK_TABLE_CALC': 'table calculation: check "compute using"',
    'TABLE_CALC': 'table calculation: visual calculation or hand-written measure',
    'EXTRACT': 'the workbook uses an extract: confirm the live source',
    'EXTRACT_DATA': 'the .twbx packages extract data',
    'URL_ACTION': 'URL action needs a Web URL column',
    'NAVIGATION_ACTION': 'navigation action becomes a page button',
    'FILTER_ACTION': 'filter action becomes drillthrough or cross-filter',
}


def load_json(path, default=None):
    try:
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    except (OSError, ValueError):
        return default


def read_text(path):
    with open(path, encoding='utf-8') as f:
        return f.read()


def sha(paths):
    h = hashlib.sha256()
    for p in sorted(paths):
        with open(p, 'rb') as f:
            h.update(os.path.basename(p).encode() + f.read())
    return h.hexdigest()[:16]


def find_inputs():
    """({name: workbook path}, [tds paths], table_map) found under input\\."""
    wbs, tds = {}, []
    for dp, _dirs, fs in os.walk(INPUT_DIR) if os.path.isdir(INPUT_DIR) else []:
        for fn in sorted(fs):
            if fn.startswith(('~$', '.')):
                continue
            p = os.path.join(dp, fn)
            if fn.lower().endswith(('.twb', '.twbx')):
                wbs.setdefault(os.path.splitext(fn)[0], p)
            elif fn.lower().endswith(('.tds', '.tdsx')):
                tds.append(p)
    return wbs, tds, load_json(os.path.join(INPUT_DIR, 'table_map.json'), {}) or {}


def out_dir(name):
    return os.path.join(OUTPUT_DIR, name)


def read_status(name):
    return load_json(os.path.join(out_dir(name), STATUS_JSON))


def fingerprint(name):
    wbs, tds, _ = find_inputs()
    files = [wbs[name], *tds] if name in wbs else []
    extra = [p for p in (os.path.join(INPUT_DIR, 'sources.json'), os.path.join(INPUT_DIR, 'decisions.json'),
                         os.path.join(INPUT_DIR, 'table_map.json')) if os.path.exists(p)]
    return sha(files + extra) if files else None


def scan():
    """[{name, file, state}] state: new | changed | done | blocked | failed."""
    wbs, _tds, _ = find_inputs()
    rows = []
    for name, path in sorted(wbs.items()):
        st = read_status(name)
        if not st:
            state = 'new'
        elif st.get('fingerprint') != fingerprint(name):
            state = 'changed'
        else:
            state = st.get('state', 'done')
        rows.append({'name': name, 'file': os.path.relpath(path, HERE), 'state': state,
                     'verdict': (st or {}).get('verdict')})
    return rows


# ---------- stages ----------
def stage_extract(name, path, tds, tmap):
    import bhayankar_extractor as bx
    buf = io.StringIO()
    t0 = time.time()
    try:
        with redirect_stdout(buf):
            res, files = bx.process(path, tds, out_dir(name), None, tmap, strict_tds=True)
    except (Exception, SystemExit) as e:  # a damaged workbook must not stop the other workbooks
        return None, {'ok': False, 'error': str(e)[:300]}
    return res, {'ok': True, 'seconds': round(time.time() - t0, 1), 'dashboards': res['workbook']['dashboards'],
                 'worksheets': res['workbook']['worksheets'], 'calcs_used': res['workbook']['calcs_used'],
                 'files': sorted(files)}


def decision_lines(name):
    lines = []
    d = out_dir(name)
    if os.path.isdir(d):
        for fn in sorted(os.listdir(d)):
            if fn.endswith(('.md',)) and fn != STATUS_MD:
                with open(os.path.join(d, fn), encoding='utf-8') as f:
                    for ln in f:
                        if 'DECISION NEEDED' in ln:
                            ln = ln.strip().lstrip('-*| ').strip()
                            if ln not in lines:
                                lines.append(ln[:240])
    return lines


def sample_mode(name):
    """True when this workbook builds on sample CSVs: an input/<workbook>_Data folder with CSVs, or a sources.json in
    sample mode whose sample folder is named after the workbook."""
    for cand in (name + '_Data', name.replace(' ', '_') + '_Data'):
        folder = os.path.join(INPUT_DIR, cand)
        if os.path.isdir(folder) and any(f.lower().endswith('.csv') for f in os.listdir(folder)):
            return True
    src = load_json(os.path.join(INPUT_DIR, 'sources.json'), {}) or {}
    return src.get('mode') == 'sample' and os.path.basename(str(src.get('sample_folder', ''))).startswith(name)


def gate(name, res, approvals):
    """(blockers, compare, design, people) from the digest. Blockers stop the build; the rest are listed."""
    flags = {}
    for f in res['flags']:
        flags.setdefault(f['code'], []).append(f)
    blockers = []
    if 'TRUNCATED_FILE' in flags:
        blockers.append('The workbook file is damaged or cut off. Download it again and re-run.')
    sample = sample_mode(name)
    if 'PUBLISHED_DS' in flags and not approvals.get('accept_guessed_joins') and not sample:
        srcs = ', '.join(sorted({f['where'] for f in flags['PUBLISHED_DS']}))
        blockers.append(f'Published data source ({srcs}) has no .tds in input\\: its joins are not in the workbook. '
                        'Add the .tds, or set "accept_guessed_joins": true under "workflow" in decisions.json to build with '
                        'guessed joins (numbers will not be reliable).')
    if flags.keys() & {'RLS', 'PER_USER_COLUMN'} and not approvals.get('rls_approved_by'):
        blockers.append('Row-level security / USERNAME() logic found. The Architect must approve the design '
                        '(docs\\SECURITY_RLS.md). Record it as "rls_approved_by": "<name>" under "workflow" in decisions.json.')
    if flags.get('HARDCODED_USERS') and not approvals.get('rls_approved_by'):
        blockers.append('Logins are hard-coded in calculations (' + ', '.join(f['where'] for f in flags['HARDCODED_USERS']) +
                        '): needs the same Architect approval as RLS.')
    if 'PUBLISHED_DS' in flags and sample:
        flags['PUBLISHED_DS_SAMPLE'] = [{'where': '; '.join(f['where'] for f in flags['PUBLISHED_DS']),
                                         'detail': 'sample CSV mode: the build merges tables and infers joins from the CSV headers '
                                                   '(see DECISION NEEDED lines). A real warehouse build needs the published .tds.'}]
    compare = [{'code': c, 'why': COMPARE_FLAGS[c], 'where': '; '.join(f['where'] for f in flags[c])[:200],
                'detail': flags[c][0]['detail'][:200]} for c in COMPARE_FLAGS if c in flags]
    design = [{'code': c, 'why': DESIGN_FLAGS[c], 'where': '; '.join(f['where'] for f in flags[c])[:200]}
              for c in DESIGN_FLAGS if c in flags]
    low = [c['name'] for c in res['calculations'] if c.get('confidence') == 'low']
    people = {'low_confidence_calcs': low, 'decisions': decision_lines(name)}
    return blockers, compare, design, people


def run_cli(args, timeout=900):
    p = subprocess.run([sys.executable, *args], cwd=HERE, capture_output=True, text=True, encoding='utf-8',
                       errors='replace', timeout=timeout, env={**os.environ, 'PYTHONIOENCODING': 'utf-8'})
    return p.returncode, (p.stdout or '') + (p.stderr or '')


def stage_build(name):
    t0 = time.time()
    try:
        rc, out = run_cli(['bhayankar_build.py', name])
    except subprocess.TimeoutExpired:
        return {'ok': False, 'error': 'build timed out after 15 minutes'}
    m = re.search(r'validate: (\S+) \((\S+) errors, (\S+) warnings\)', out)
    log = os.path.join(out_dir(name), 'bhayankar_build_log.md')
    project = None
    if os.path.exists(log):
        with open(log, encoding='utf-8') as f:
            pm = re.search(r'Project: `([^`]+)`', f.read())
            project = pm.group(1) if pm else None
    if not m:
        return {'ok': False, 'error': (out.strip().splitlines() or ['build produced no result'])[-1][:300], 'project': project}
    errors, warnings = (int(x) if x.isdigit() else None for x in (m.group(2), m.group(3)))
    mc = re.search(r'model check: (\d+) errors, (\d+) warnings', out)
    model_errors = int(mc.group(1)) if mc else None
    return {'ok': rc == 0 and errors == 0 and not model_errors, 'validate': m.group(1), 'errors': errors,
            'warnings': warnings, 'model_errors': model_errors, 'model_warnings': int(mc.group(2)) if mc else None,
            'project': project, 'seconds': round(time.time() - t0, 1)}


def stage_test(name, workbook, project):
    if not project or not os.path.exists(project):
        return {'ok': False, 'error': 'no project to test'}
    dest = os.path.join(out_dir(name), 'test')
    try:
        rc, out = run_cli(['bhayankar_test.py', workbook, project, '--no-data', '--out', dest])
    except subprocess.TimeoutExpired:
        return {'ok': False, 'error': 'tester timed out after 15 minutes'}
    m = re.search(r'unit (\d+) pass / (\d+) fail / (\d+) warn, functional (\d+)/(\d+)/(\d+)', out)
    if not m:
        return {'ok': False, 'error': (out.strip().splitlines() or ['tester produced no result'])[-1][:300]}
    u = [int(x) for x in m.groups()]
    return {'ok': u[1] == 0 and u[4] == 0, 'unit': {'pass': u[0], 'fail': u[1], 'warn': u[2]},
            'functional': {'pass': u[3], 'fail': u[4], 'warn': u[5]}, 'report': os.path.join(dest, name, 'test_report.xlsx'),
            'data_tested': False, 'readiness': readiness(os.path.join(dest, name))}


def readiness(folder):
    """Share of the tester's structure checks that pass: PASS / (PASS + FAIL + WARN + MANUAL); INFO and SKIP do not
    count. One number to compare workbooks and toolkit versions. It says nothing about the numbers (no data)."""
    import csv
    c = {}
    for fn in ('unit_tests.csv', 'functional_tests.csv'):
        p = os.path.join(folder, fn)
        if os.path.exists(p):
            with open(p, encoding='utf-8-sig', newline='') as f:
                for r in csv.DictReader(f):
                    c[r['Status']] = c.get(r['Status'], 0) + 1
    n = sum(v for k, v in c.items() if k in ('PASS', 'FAIL', 'WARN', 'MANUAL'))
    return {'percent': round(100 * c.get('PASS', 0) / n, 1), 'checks': n, 'fail': c.get('FAIL', 0),
            'review': c.get('WARN', 0) + c.get('MANUAL', 0)} if n else None


# ---------- reporting ----------
def verdict(st):
    s = st['stages']
    if not s['extract'].get('ok'):
        return 'FAILED', 'The extractor could not read this workbook: ' + s['extract'].get('error', '')
    if st['blockers']:
        return 'BLOCKED', f'{len(st["blockers"])} thing(s) must be resolved before the build. See "Blockers".'
    b = s.get('build')
    if not b:
        return 'READY TO BUILD', 'Extracted and cleared the gate. Run again without --no-build to build.'
    if not b.get('ok'):
        if b.get('model_errors'):
            return 'BUILD FAILED', (f'The model check found {b["model_errors"]} error(s) that stop Power BI Desktop loading the '
                                    'model; see "Model check" in bhayankar_build_log.md.')
        return 'BUILD FAILED', b.get('error') or f'Validation found {b.get("errors")} error(s).'
    t = s.get('test')
    if t and not t.get('ok'):
        return 'TESTS FAILED', 'Built, but the structure tests report failures. See the test report.'
    return 'BUILT - NEEDS HUMAN CHECKS', ('Built and validated. It is NOT signed off: open it in Power BI Desktop, refresh, '
                                          'and compare numbers with Tableau (see "Compare with Tableau").')


def render(st):
    L = [f'# Workflow status: {st["workbook"]}', '', f'**Verdict: {st["verdict"]}**', '', st['verdict_detail'], '',
         f'Run at {st["time"]}. Tier suggested by the extractor: **{st.get("tier", "?")}** (the Architect confirms it).', '']
    s = st['stages']
    L += ['## Stages', '', '| Stage | Result |', '|---|---|']
    e = s['extract']
    L.append('| 1. Extract | ' + (f'done: {e.get("dashboards")} dashboards, {e.get("worksheets")} sheets, {e.get("calcs_used")} calcs used'
                                  if e.get('ok') else 'FAILED: ' + e.get('error', '')) + ' |')
    L.append('| 2. Gate | ' + (f'{len(st["blockers"])} blocker(s)' if st['blockers'] else 'clear') + ' |')
    b = s.get('build')
    L.append('| 3. Build | ' + ('not run' if not b else f'validate {b.get("validate")}: {b.get("errors")} errors, {b.get("warnings")} warnings'
                                 + (f'; model check: {b["model_errors"]} errors, {b.get("model_warnings")} warnings'
                                    if b.get('model_errors') is not None else '')
                                 if 'validate' in b else 'FAILED: ' + b.get('error', '')) + ' |')
    t = s.get('test')
    L.append('| 4. Test | ' + ('not run' if not t else (f'unit {t["unit"]["pass"]} pass / {t["unit"]["fail"]} fail / {t["unit"]["warn"]} warn; '
                                                        f'functional {t["functional"]["pass"]}/{t["functional"]["fail"]}/{t["functional"]["warn"]}; '
                                                        'structure only, no data compared') if 'unit' in t else 'FAILED: ' + t.get('error', '')) + ' |')
    L.append('| 5. Number check vs Tableau | **not done: a person does this** |')
    rd = (t or {}).get('readiness')
    if rd:
        L += ['', f'**Structural readiness: {rd["percent"]}%** of {rd["checks"]} structure checks pass; {rd["review"]} need a '
                  f'person (WARN / MANUAL), {rd["fail"]} fail. Structure only: it says nothing about the numbers.']
    if st['blockers']:
        L += ['', '## Blockers (resolve, then run again)', ''] + [f'- {x}' for x in st['blockers']]
    pe = st['people']
    if pe['decisions'] or pe['low_confidence_calcs']:
        L += ['', '## Needs a person', '']
        L += [f'- {x}' for x in pe['decisions'][:30]]
        if pe['low_confidence_calcs']:
            L.append('- Translations marked low confidence (check the DAX): ' + ', '.join(pe['low_confidence_calcs'][:30]))
    if st['compare']:
        L += ['', '## Compare with Tableau before sign-off', '']
        L += [f'- **{c["code"]}** ({c["where"]}): {c["why"]}. {c["detail"]}' for c in st['compare']]
    if st['design']:
        L += ['', '## Design follow-ups', '']
        L += [f'- **{d["code"]}** ({d["where"]}): {d["why"]}' for d in st['design']]
    L += ['', '## Next steps', ''] + [f'{i}. {x}' for i, x in enumerate(st['next'], 1)]
    if b and b.get('project'):
        L += ['', f'Project: `{b["project"]}`']
    L += ['', 'Details: `bhayankar_digest.md` (read first), `bhayankar_build_log.md`, `test\\`.']
    return '\n'.join(L) + '\n'


def next_steps(st):
    v = st['verdict']
    if v == 'FAILED':
        return ['Open the workbook in Tableau and save it again, or use --show on the failing item.', 'Run again.']
    if v == 'BLOCKED':
        return ['Resolve each blocker.', 'Run again: python bhayankar_run.py run "' + st['workbook'] + '"']
    if v == 'READY TO BUILD':
        return ['Run again without --no-build.']
    if v in ('BUILD FAILED', 'TESTS FAILED'):
        return ['Read bhayankar_build_log.md and the test report.', 'Fix the cause in the tool (add a test), then run again.']
    return ['Answer the "Needs a person" items and record them in input\\decisions.json.',
            'Open the .pbip in Power BI Desktop, refresh, and check the model has no errors.',
            'Compare real numbers with Tableau, one filter selection per page (always for the "Compare" items).',
            'Run Run_Verify.bat, then work through docs\\QA_CHECKLIST.md before hand-off.']


def run_one(name, build=True, test=True, force=False):
    wbs, tds, tmap = find_inputs()
    if name not in wbs:
        return {'workbook': name, 'verdict': 'NOT FOUND', 'verdict_detail': f'No workbook named "{name}" in input\\.',
                'state': 'failed'}
    old = read_status(name)
    fp = fingerprint(name)
    if old and old.get('fingerprint') == fp and not force and old.get('state') == 'done' and \
            (old['stages'].get('build') or not build):
        old['skipped'] = 'nothing changed since the last run (use --force to run again)'
        return old
    approvals = (load_json(os.path.join(INPUT_DIR, 'decisions.json'), {}) or {}).get(name, {}).get('workflow', {})
    res, ex = stage_extract(name, wbs[name], tds, tmap)
    st = {'workbook': name, 'file': os.path.relpath(wbs[name], HERE), 'fingerprint': fp, 'time': time.strftime('%Y-%m-%d %H:%M'),
          'stages': {'extract': ex}, 'blockers': [], 'compare': [], 'design': [], 'people': {'decisions': [], 'low_confidence_calcs': []}}
    if res:
        st['tier'] = res['tier']['suggested']
        st['blockers'], st['compare'], st['design'], st['people'] = gate(name, res, approvals)
        if build and not st['blockers']:
            st['stages']['build'] = stage_build(name)
            b = st['stages']['build']
            if test and b.get('ok'):
                st['stages']['test'] = stage_test(name, wbs[name], b.get('project'))
            st['people']['decisions'] = decision_lines(name)  # the build adds its own
    st['verdict'], st['verdict_detail'] = verdict(st)
    st['state'] = {'FAILED': 'failed', 'BLOCKED': 'blocked', 'BUILD FAILED': 'failed', 'TESTS FAILED': 'failed'}.get(st['verdict'], 'done')
    st['next'] = next_steps(st)
    os.makedirs(out_dir(name), exist_ok=True)
    with open(os.path.join(out_dir(name), STATUS_JSON), 'w', encoding='utf-8') as f:
        json.dump(st, f, indent=1, ensure_ascii=False)
    with open(os.path.join(out_dir(name), STATUS_MD), 'w', encoding='utf-8') as f:
        f.write(render(st))
    return st


def run_all(names=None, **kw):
    # --force with no name re-runs every workbook, "done" ones included (as the help says)
    wanted = names or [r['name'] for r in scan() if kw.get('force') or r['state'] in ('new', 'changed', 'blocked', 'failed')]
    return [run_one(n, **kw) for n in wanted]


def main():
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('cmd', choices=['scan', 'run', 'status'])
    ap.add_argument('workbook', nargs='?')
    ap.add_argument('--no-build', action='store_true')
    ap.add_argument('--no-test', action='store_true')
    ap.add_argument('--force', action='store_true')
    ap.add_argument('--json', action='store_true')
    a = ap.parse_args()
    if a.cmd == 'scan':
        rows = scan()
        if a.json:
            print(json.dumps(rows, indent=1))
        else:
            print('\n'.join(f'{r["state"]:8} {r["name"]}  ({r["verdict"] or "not run"})' for r in rows) or 'No workbooks in input\\.')
        return
    if a.cmd == 'status':
        p = os.path.join(out_dir(a.workbook or ''), STATUS_MD)
        print(read_text(p) if os.path.exists(p) else 'No status yet: run it first.')
        return
    results = run_all([a.workbook] if a.workbook else None, build=not a.no_build, test=not a.no_test, force=a.force)
    if not results:
        print('Nothing to do: every workbook in input\\ is up to date. Use --force to run again.')
    for st in results:
        print(f'{st["workbook"]}: {st["verdict"]}' + (f' ({st["skipped"]})' if st.get('skipped') else ''))
        if st.get('verdict') not in ('NOT FOUND',) and not st.get('skipped'):
            print(f'   status: output\\{st["workbook"]}\\{STATUS_MD}')
    sys.exit(0 if all(r.get('state') in ('done', 'blocked') for r in results) else 1)


if __name__ == '__main__':
    main()
