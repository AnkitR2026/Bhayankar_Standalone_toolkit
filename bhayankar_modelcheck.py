#!/usr/bin/env python3
r"""Bhayankar model check: offline checks of a PBIP semantic model (TMDL) for what makes Power BI Desktop refuse it
or show broken numbers, which the PBIR (report) validator does not look at.

Errors (Desktop does not load the model, or a measure / column is in error):
  NAME_CLASH          two objects whose names differ only in case (Power BI names are case-insensitive): columns and
                      measures of one table, measures of the whole model, tables
  UNKNOWN_REF         'Table'[Column] naming a table or column that is not in the model; [Name] that is no measure
  UNKNOWN_FUNCTION    a function DAX does not have (e.g. a Tableau name such as DAYOFYEAR( or ISO-WEEK()
  SYNTAX              unbalanced brackets or an unterminated string
  CIRCULAR            a measure that reads itself, directly or round a loop
  SORT_BY             sortByColumn naming a column that is not in the table, or the column itself
  RELATIONSHIP        an end that is not in the model, end columns of different types, two active relationships
                      between the same two tables
Warnings (loads, but check it):
  LOOP                active relationships forming a loop (Desktop may make one inactive)
  PLACEHOLDER         measures left as BLANK() in the _Needs fix folder
  REPORT_BINDING      a report visual / filter naming a field that is not in the model
  SORT_BY_CHAIN       a sort-by column that is itself sorted by another column

Usage:  python bhayankar_modelcheck.py <folder holding a .pbip, a .SemanticModel folder, or a .pbip file>
Python 3 standard library only.
"""
import glob
import json
import os
import re
import sys

DAX_FUNCTIONS = set('''
ABS ACOS ACOSH ACOT ACOTH ADDCOLUMNS ADDMISSINGITEMS ALL ALLCROSSFILTERED ALLEXCEPT ALLNOBLANKROW ALLSELECTED AND
APPROXIMATEDISTINCTCOUNT ASIN ASINH ATAN ATANH AVERAGE AVERAGEA AVERAGEX BETA.DIST BETA.INV BITAND BITLSHIFT BITOR
BITRSHIFT BITXOR BLANK CALCULATE CALCULATETABLE CALENDAR CALENDARAUTO CEILING CHISQ.DIST CHISQ.DIST.RT CHISQ.INV
CHISQ.INV.RT CLOSINGBALANCEMONTH CLOSINGBALANCEQUARTER CLOSINGBALANCEYEAR COALESCE COLLAPSE COLLAPSEALL COLUMNSTATISTICS
COMBIN COMBINA COMBINEVALUES CONCATENATE CONCATENATEX CONFIDENCE.NORM CONFIDENCE.T CONTAINS CONTAINSROW CONTAINSSTRING
CONTAINSSTRINGEXACT CONVERT COS COSH COT COTH COUNT COUNTA COUNTAX COUNTBLANK COUNTROWS COUNTX CROSSFILTER CROSSJOIN
CURRENCY CURRENTGROUP CUSTOMDATA DATATABLE DATE DATEADD DATEDIFF DATESBETWEEN DATESINPERIOD DATESMTD DATESQTD DATESYTD
DATEVALUE DAY DEGREES DETAILROWS DISTINCT DISTINCTCOUNT DISTINCTCOUNTNOBLANK DIVIDE EARLIER EARLIEST EDATE ENDOFMONTH
ENDOFQUARTER ENDOFYEAR EOMONTH ERROR EVALUATEANDLOG EVEN EXACT EXCEPT EXP EXPAND EXPANDALL EXPON.DIST FACT FALSE FILTER
FILTERS FIND FIRST FIRSTDATE FIRSTNONBLANK FIRSTNONBLANKVALUE FIXED FLOOR FORMAT GCD GENERATE GENERATEALL GENERATESERIES
GEOMEAN GEOMEANX GROUPBY HASONEFILTER HASONEVALUE HOUR IF IF.EAGER IFERROR IGNORE INDEX INT INTERSECT ISAFTER ISBLANK
ISCROSSFILTERED ISEMPTY ISERROR ISEVEN ISFILTERED ISINSCOPE ISLOGICAL ISNONTEXT ISNUMBER ISO.CEILING ISODD ISONORAFTER
ISSELECTEDMEASURE ISSUBTOTAL ISTEXT KEEPFILTERS LAST LASTDATE LASTNONBLANK LASTNONBLANKVALUE LCM LEFT LEN LINEST LINESTX
LN LOG LOG10 LOOKUP LOOKUPVALUE LOWER MATCHBY MAX MAXA MAXX MEDIAN MEDIANX MID MIN MINA MINUTE MINX MOD MONTH
MOVINGAVERAGE MROUND NATURALINNERJOIN NATURALLEFTOUTERJOIN NETWORKDAYS NEXT NEXTDAY NEXTMONTH NEXTQUARTER NEXTYEAR
NONVISUAL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NOT NOW ODD OFFSET OPENINGBALANCEMONTH OPENINGBALANCEQUARTER
OPENINGBALANCEYEAR OR ORDERBY PARALLELPERIOD PARTITIONBY PATH PATHCONTAINS PATHITEM PATHITEMREVERSE PATHLENGTH
PERCENTILE.EXC PERCENTILE.INC PERCENTILEX.EXC PERCENTILEX.INC PI POISSON.DIST POWER PREVIOUS PREVIOUSDAY PREVIOUSMONTH
PREVIOUSQUARTER PREVIOUSYEAR PRODUCT PRODUCTX QUARTER QUOTIENT RADIANS RAND RANDBETWEEN RANGE RANK RANK.EQ RANKX RELATED
RELATEDTABLE REMOVEFILTERS REPLACE REPT RIGHT ROLLUP ROLLUPADDISSUBTOTAL ROLLUPGROUP ROLLUPISSUBTOTAL ROUND ROUNDDOWN
ROUNDUP ROW ROWNUMBER RUNNINGSUM SAMEPERIODLASTYEAR SAMPLE SEARCH SECOND SELECTCOLUMNS SELECTEDMEASURE
SELECTEDMEASUREFORMATSTRING SELECTEDMEASURENAME SELECTEDVALUE SIGN SIN SINH SQRT SQRTPI STARTOFMONTH STARTOFQUARTER
STARTOFYEAR STDEV.P STDEV.S STDEVX.P STDEVX.S SUBSTITUTE SUBSTITUTEWITHINDEX SUM SUMMARIZE SUMMARIZECOLUMNS SUMX SWITCH
T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T TAN TANH TIME TIMEVALUE TOCSV TODAY TOJSON TOPN TOPNSKIP TOTALMTD TOTALQTD
TOTALYTD TREATAS TRIM TRUE TRUNC UNICHAR UNICODE UNION UPPER USERCULTURE USERELATIONSHIP USERNAME USEROBJECTID
USERPRINCIPALNAME UTCNOW UTCTODAY VALUE VALUES VAR.P VAR.S VARX.P VARX.S WEEKDAY WEEKNUM WINDOW XIRR XNPV YEAR YEARFRAC
'''.split())
KEYWORDS = {'RETURN', 'VAR', 'IN', 'ASC', 'DESC', 'SKIP', 'DENSE', 'DEFINE', 'EVALUATE', 'MEASURE'}


def unq(s):
    s = s.strip()
    return s[1:-1].replace("''", "'") if len(s) > 1 and s[0] == s[-1] == "'" else s


def strip_text(dax):
    """DAX with string literals and comments emptied (their content is not code)."""
    out, i, n = [], 0, len(dax)
    while i < n:
        ch = dax[i]
        if ch == '"':
            j = i + 1
            while j < n:
                if dax[j] == '"' and dax[j + 1:j + 2] == '"':
                    j += 2
                    continue
                if dax[j] == '"':
                    break
                j += 1
            if j >= n:
                return None  # unterminated string
            out.append('""')
            i = j + 1
        elif dax.startswith('//', i) or dax.startswith('--', i):
            j = dax.find('\n', i)
            i = n if j < 0 else j
        elif dax.startswith('/*', i):
            j = dax.find('*/', i + 2)
            i = n if j < 0 else j + 2
        else:
            out.append(ch)
            i += 1
    return ''.join(out)


def parse_tmdl(path):
    """One table .tmdl -> {'name', 'columns': {name: obj}, 'measures': {...}, 'partitions': [...]}; obj = {'name',
    'expr', 'props'}."""
    with open(path, encoding='utf-8-sig') as fh:
        lines = fh.read().split('\n')
    t = {'name': None, 'columns': {}, 'measures': {}, 'partitions': [], 'dups': []}
    cur, i = None, 0
    obj = re.compile(r"^\t(measure|column|partition|hierarchy|calculationGroup) ((?:'[^']*(?:''[^']*)*')|[^ =]+)(?: = ?(.*))?$")
    while i < len(lines):
        ln = lines[i]
        m = re.match(r"^table ((?:'[^']*(?:''[^']*)*')|.+?)\s*$", ln)
        if m:
            t['name'] = unq(m.group(1))
            i += 1
            continue
        m = obj.match(ln)
        if m:
            kind, name, expr = m.group(1), unq(m.group(2)), m.group(3)
            if expr is not None and expr.strip() == '```':
                body, i = [], i + 1
                while i < len(lines) and lines[i].strip() != '```':
                    body.append(lines[i])
                    i += 1
                expr = '\n'.join(body)
            elif expr is not None and expr.strip() == '':
                body = []
                while i + 1 < len(lines) and lines[i + 1].startswith('\t\t\t'):
                    i += 1
                    body.append(lines[i])
                expr = '\n'.join(body)
            cur = {'kind': kind, 'name': name, 'expr': expr, 'props': {}}
            if kind in ('column', 'measure'):
                bucket = t['columns' if kind == 'column' else 'measures']
                if name in bucket:
                    t['dups'].append(name)
                bucket[name] = cur
            elif kind == 'partition':
                t['partitions'].append(cur)
            i += 1
            continue
        m = re.match(r'^\t\t(\w+)(?:\s*[:=]\s*(.*))?$', ln)
        if m and cur is not None:
            val = m.group(2)
            if cur['kind'] == 'partition' and m.group(1) == 'source' and (val is None or val.strip() == ''):
                body = []
                while i + 1 < len(lines) and lines[i + 1].startswith('\t\t\t'):
                    i += 1
                    body.append(lines[i])
                val = '\n'.join(body)
            cur['props'][m.group(1)] = val if val is not None else True
        i += 1
    return t


def parse_relationships(path):
    rels, cur = [], None
    if not os.path.exists(path):
        return rels
    with open(path, encoding='utf-8-sig') as fh:
        text = fh.read()
    for ln in text.split('\n'):
        if ln.startswith('relationship '):
            cur = {'name': ln[13:].strip()}
            rels.append(cur)
        elif cur is not None and ln.startswith('\t') and ':' in ln:
            k, v = ln.strip().split(':', 1)
            cur[k.strip()] = v.strip()
    for r in rels:
        for side in ('fromColumn', 'toColumn'):
            m = re.match(r"^((?:'[^']*(?:''[^']*)*')|[^.]+)\.((?:'[^']*(?:''[^']*)*')|.+)$", r.get(side, ''))
            r[side + '_parsed'] = (unq(m.group(1)), unq(m.group(2))) if m else (None, None)
    return rels


def find_model(path):
    path = os.path.abspath(path)
    if path.lower().endswith('.pbip'):
        path = os.path.dirname(path)
    if path.lower().endswith('.semanticmodel'):
        return path, None
    sms = glob.glob(os.path.join(path, '*.SemanticModel'))
    reps = glob.glob(os.path.join(path, '*.Report'))
    return (sms[0] if sms else None), (reps[0] if reps else None)


def check(path, report_dir=None):
    """{'errors': [(code, where, message)], 'warnings': [...], 'stats': {...}} for the semantic model at path."""
    sm, rep = find_model(path)
    rep = report_dir or rep
    errors, warnings = [], []
    E = lambda code, where, msg: errors.append((code, where, msg))  # noqa: E731
    W = lambda code, where, msg: warnings.append((code, where, msg))  # noqa: E731
    if not sm:
        E('MISSING', path, 'no .SemanticModel folder')
        return {'errors': errors, 'warnings': warnings, 'stats': {}}
    tables = {}
    for f in sorted(glob.glob(os.path.join(sm, 'definition', 'tables', '*.tmdl'))):
        t = parse_tmdl(f)
        if t['name'] is None:
            E('SYNTAX', os.path.basename(f), 'no "table" line')
            continue
        for d in t['dups']:
            E('NAME_CLASH', t['name'], f'[{d}] is defined twice')
        tables[t['name']] = t
    low = {}
    for n in tables:
        low.setdefault(n.lower(), []).append(n)
    for names in low.values():
        if len(names) > 1:
            E('NAME_CLASH', ', '.join(names), 'table names differ only in case')
    all_meas = {}
    for t in tables.values():
        seen = {}
        for n in list(t['columns']) + list(t['measures']):
            if n.lower() in seen and seen[n.lower()] != n:
                E('NAME_CLASH', t['name'], f'[{seen[n.lower()]}] and [{n}] differ only in case')
            elif n in t['columns'] and n in t['measures']:
                E('NAME_CLASH', t['name'], f'[{n}] is a column and a measure')
            seen.setdefault(n.lower(), n)
        for n in t['measures']:
            all_meas.setdefault(n.lower(), []).append((t['name'], n))
    for lst in all_meas.values():
        if len(lst) > 1:
            E('NAME_CLASH', ', '.join(t for t, _ in lst), 'measure names must be unique in the model (case-insensitive): '
              + ', '.join(f'[{n}]' for _, n in lst))
    meas_names = {n.lower(): n for t in tables.values() for n in t['measures']}
    table_low = {n.lower(): t for n, t in tables.items()}

    def check_dax(where, dax, home, is_measure):
        text = strip_text(dax or '')
        if text is None:
            E('SYNTAX', where, 'unterminated string')
            return set()
        local = {x.lower() for x in re.findall(r'"([^"]+)"', dax or '')}  # names made by ADDCOLUMNS / SELECTCOLUMNS
        refs_meas = set()
        # names in order: 'table' ('' escapes) and [column] (]] escapes) may hold quotes, brackets and parens of their own
        toks, code, i, n = [], [], 0, len(text)
        while i < n:
            ch = text[i]
            if ch in "'[":
                close = "'" if ch == "'" else ']'
                j = i + 1
                while j < n and not (text[j] == close and text[j + 1:j + 2] != close):
                    j += 2 if text[j] == close else 1
                if j >= n:
                    E('SYNTAX', where, f'unterminated {"table name" if ch == chr(39) else "[name]"}')
                    return refs_meas
                name = text[i + 1:j].replace(close * 2, close)
                toks.append(('table' if ch == "'" else 'col', name, i))
                code.append("'T'" if ch == "'" else '[x]')
                i = j + 1
                continue
            code.append(ch)
            i += 1
        code = ''.join(code)
        bare = []
        for k, (kind, name, pos) in enumerate(toks):
            if kind == 'table':
                t = table_low.get(name.lower())
                if not t:
                    E('UNKNOWN_REF', where, f"table '{name}' is not in the model")
                    continue
                nxt = toks[k + 1] if k + 1 < len(toks) else None
                if nxt and nxt[0] == 'col' and text[pos:nxt[2]].endswith("'"):  # 'T'[c] written together
                    cn = nxt[1]
                    if cn.lower() not in {c.lower() for c in t['columns']} and cn.lower() not in {m.lower() for m in t['measures']}:
                        E('UNKNOWN_REF', where, f"'{name}'[{cn}] is not in the model")
                    elif cn.lower() in meas_names:
                        refs_meas.add(meas_names[cn.lower()])
            elif not (k and toks[k - 1][0] == 'table' and text[toks[k - 1][2]:pos].endswith("'")) and \
                    not re.match(r'\w', text[pos - 1:pos]):
                bare.append(name)
        if code.count('(') != code.count(')'):
            E('SYNTAX', where, f'unbalanced brackets ({code.count("(")} open, {code.count(")")} close)')
        if code.count('{') != code.count('}'):
            E('SYNTAX', where, 'unbalanced braces')
        for fn in re.findall(r'(?<![\w.@])([A-Za-z_][\w.]*)\s*\(', code):
            if fn.upper() not in DAX_FUNCTIONS and fn.upper() not in KEYWORDS:
                E('UNKNOWN_FUNCTION', where, f'{fn}() is not a DAX function')
        for r in bare:
            if r.startswith('@') or r.lower() in local:
                continue
            if r.lower() in meas_names:
                refs_meas.add(meas_names[r.lower()])
                continue
            if not is_measure and home and r.lower() in {c.lower() for c in tables[home]['columns']}:
                continue
            E('UNKNOWN_REF', where, f'[{r}] is not a measure' + ('' if is_measure else ' or a column of this table'))
        return refs_meas

    deps, placeholders = {}, 0
    for t in tables.values():
        for m in t['measures'].values():
            deps[m['name']] = check_dax(f"measure '{t['name']}'[{m['name']}]", m['expr'], t['name'], True)
            if (m['expr'] or '').strip() == 'BLANK()' and '_Needs fix' in str(m['props'].get('displayFolder') or ''):
                placeholders += 1
        for c in t['columns'].values():
            if c['expr']:
                check_dax(f"column '{t['name']}'[{c['name']}]", c['expr'], t['name'], False)
            sb = c['props'].get('sortByColumn')
            if sb:
                target = unq(str(sb))
                if target not in t['columns']:
                    E('SORT_BY', f"'{t['name']}'[{c['name']}]", f'sortByColumn [{target}] is not a column of the table')
                elif target == c['name']:
                    E('SORT_BY', f"'{t['name']}'[{c['name']}]", 'sorted by itself')
                elif t['columns'][target]['props'].get('sortByColumn'):
                    W('SORT_BY_CHAIN', f"'{t['name']}'[{c['name']}]", f'[{target}] is itself sorted by another column')
        for p in t['partitions']:
            src = p['props'].get('source')
            if p['expr'] and p['expr'].strip() == 'calculated' and isinstance(src, str):
                check_dax(f"table '{t['name']}' (calculated)", src, None, True)
    if placeholders:
        W('PLACEHOLDER', 'model', f'{placeholders} measure(s) are BLANK() placeholders in the _Needs fix folder')
    # circular measures
    for start in deps:
        stack, seen = [(start, [start])], {start}
        while stack:
            x, path = stack.pop()
            for y in deps.get(x, ()):
                if y == start:
                    E('CIRCULAR', f'[{start}]', ' -> '.join(f'[{p}]' for p in path + [start]))
                    stack = []
                    break
                if y not in seen:
                    seen.add(y)
                    stack.append((y, path + [y]))
    # relationships
    rels = parse_relationships(os.path.join(sm, 'definition', 'relationships.tmdl'))
    pairs, flows = {}, {}  # flows: table -> [(table it filters, relationship)]
    for r in rels:
        (ft, fc), (tt, tc) = r['fromColumn_parsed'], r['toColumn_parsed']
        ok = True
        for tn, cn in ((ft, fc), (tt, tc)):
            if tn not in tables or cn not in tables[tn]['columns']:
                E('RELATIONSHIP', r['name'], f"end '{tn}'[{cn}] is not in the model")
                ok = False
        if ok:
            d1 = tables[ft]['columns'][fc]['props'].get('dataType')
            d2 = tables[tt]['columns'][tc]['props'].get('dataType')
            if d1 and d2 and d1 != d2:
                E('RELATIONSHIP', r['name'], f'end columns have different types ({d1}, {d2})')
        if r.get('isActive', 'true').lower() == 'false' or not ok:
            continue
        key = frozenset((ft, tt))
        if key in pairs:
            E('RELATIONSHIP', r['name'], f'a second active relationship between {ft} and {tt} ({pairs[key]})')
        pairs[key] = r['name']
        flows.setdefault(tt, []).append((ft, r['name']))  # the one side filters the many side
        if r.get('crossFilteringBehavior', '').lower() == 'bothdirections':
            flows.setdefault(ft, []).append((tt, r['name']))
    # ambiguity = a table reached from another along two different filter paths (Desktop then makes a relationship
    # inactive). A shared dimension over two facts (Date and Region over Orders and Targets, a blend) is a loop when
    # the directions are ignored, but each filter has one path: not ambiguous, so not reported.
    warned = set()
    for start in sorted(flows):
        paths, stack, steps = {}, [(start, (start,), None)], 0
        while stack and steps < 50000:  # simple paths; models are small, the cap only guards a pathological one
            steps += 1
            n, path, _ = stack.pop()
            for m, rname in flows.get(n, ()):
                if m in path:
                    continue
                paths[m] = paths.get(m, 0) + 1
                if paths[m] == 2 and rname not in warned:
                    warned.add(rname)
                    W('LOOP', rname, f'{start} filters {m} along two paths (ambiguous): Desktop may make one relationship '
                                     'inactive')
                stack.append((m, path + (m,), rname))
    # report bindings
    if rep and os.path.isdir(rep):
        for f in glob.glob(os.path.join(rep, 'definition', 'pages', '*', 'visuals', '*', 'visual.json')) + \
                glob.glob(os.path.join(rep, 'definition', 'pages', '*', 'page.json')):
            try:
                with open(f, encoding='utf-8') as fh:
                    j = json.load(fh)
            except (OSError, ValueError) as e:
                E('SYNTAX', os.path.relpath(f, rep), f'not readable JSON ({e})')
                continue
            bad = set()

            def walk(x):
                if isinstance(x, dict):
                    if 'Property' in x and isinstance(x.get('Expression'), dict):
                        ent = (x['Expression'].get('SourceRef') or {}).get('Entity')
                        if ent is not None:
                            t = tables.get(ent)
                            if not t or (x['Property'] not in t['columns'] and x['Property'] not in t['measures']):
                                bad.add(f'{ent}[{x["Property"]}]')
                    for v in x.values():
                        walk(v)
                elif isinstance(x, list):
                    for v in x:
                        walk(v)
            walk(j)
            for b in sorted(bad):
                W('REPORT_BINDING', os.path.relpath(f, rep), f'{b} is not in the model')
    stats = {'tables': len(tables), 'columns': sum(len(t['columns']) for t in tables.values()),
             'measures': sum(len(t['measures']) for t in tables.values()), 'relationships': len(rels),
             'placeholders': placeholders}
    return {'errors': errors, 'warnings': warnings, 'stats': stats}


def summary(res):
    return f'model check: {len(res["errors"])} errors, {len(res["warnings"])} warnings'


def render(res):
    lines = [summary(res) + ' (' + ', '.join(f'{k} {v}' for k, v in res.get('stats', {}).items()) + ')']
    for level, rows in (('ERROR', res['errors']), ('WARNING', res['warnings'])):
        for code, where, msg in rows:
            lines.append(f'- {level} {code}: {where}: {msg}')
    return '\n'.join(lines)


def main():
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    rc = 0
    for p in sys.argv[1:]:
        res = check(p)
        print(f'{p}\n{render(res)}\n')
        rc = rc or (1 if res['errors'] else 0)
    return rc


if __name__ == '__main__':
    sys.exit(main())
