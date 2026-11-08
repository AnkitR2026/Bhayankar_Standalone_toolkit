"""Bhayankar tester, PBIP module: reads a Power BI project into plain dicts.

Model: TMDL folder (definition\\*.tmdl) or model.bim. Report: PBIR folder (definition\\pages) or the
legacy single report.json. Only what the tester compares is kept: tables, columns, measures,
partitions (M), relationships, roles, parameters, pages, visuals, fields, filters and navigation.
Python 3 standard library only.
"""
import json
import os
import re

TMDL_OBJECTS = {'model', 'table', 'column', 'measure', 'partition', 'hierarchy', 'level', 'relationship', 'role',
                'tablePermission', 'columnPermission', 'expression', 'culture', 'perspective', 'perspectiveTable',
                'perspectiveColumn', 'perspectiveMeasure', 'perspectiveHierarchy', 'annotation', 'calculationGroup',
                'calculationItem', 'ref', 'database', 'dataSource', 'linguisticMetadata', 'extendedProperty',
                'formatStringDefinition', 'detailRowsDefinition', 'queryGroup', 'function', 'calendar',
                'calendarColumnGroup', 'refreshPolicy', 'alternateOf'}
QUOTED = re.compile(r"'((?:[^']|'')*)'")
M_STRING = re.compile(r'"((?:[^"]|"")*)"')
M_PAIR = re.compile(r'\{\s*"((?:[^"]|"")*)"\s*,\s*"((?:[^"]|"")*)"\s*\}')


# ---------- TMDL ----------
def _depth(line):
    n = sp = 0
    for ch in line:
        if ch == '\t':
            n += 1
        elif ch == ' ':
            sp += 1
            if sp == 4:
                n, sp = n + 1, 0
        else:
            break
    return n


def _split_name(rest):
    """"'My Table' = expr" -> ('My Table', '= expr'); "Sales = expr" -> ('Sales', '= expr')."""
    rest = rest.strip()
    if rest.startswith("'"):
        out, i = '', 1
        while i < len(rest):
            if rest[i] == "'":
                if rest[i + 1:i + 2] == "'":
                    out, i = out + "'", i + 2
                    continue
                return out, rest[i + 1:].strip()
            out, i = out + rest[i], i + 1
        return out, ''
    if '=' in rest:
        name, _, after = rest.partition('=')
        return name.strip(), '=' + after
    return rest, ''


def parse_tmdl(text):
    """TMDL text -> list of nodes {kind, name, expr, props, children, description}."""
    items = []
    for ln in text.replace('\r', '').split('\n'):
        items.append((None, '') if not ln.strip() else (_depth(ln), ln.strip()))
    pos = [0]

    def gather(d_min, first):
        parts = []
        if first.startswith('```'):
            while pos[0] < len(items):
                t = items[pos[0]][1]
                pos[0] += 1
                if t.startswith('```'):
                    break
                parts.append(t)
            return '\n'.join(parts).strip()
        if first:
            parts.append(first)
        while pos[0] < len(items):
            dpt, t = items[pos[0]]
            if dpt is None:
                j = pos[0]
                while j < len(items) and items[j][0] is None:
                    j += 1
                if j < len(items) and items[j][0] >= d_min:
                    pos[0] = j
                    continue
                break
            if dpt < d_min:
                break
            parts.append(t)
            pos[0] += 1
        return '\n'.join(parts).strip()

    def block(d):
        nodes, props, desc = [], {}, []
        while pos[0] < len(items):
            dpt, t = items[pos[0]]
            if dpt is None:
                pos[0] += 1
                continue
            if dpt < d:
                break
            pos[0] += 1
            if dpt > d:
                continue
            if t.startswith('///'):
                desc.append(t[3:].strip())
                continue
            m = re.match(r'[A-Za-z_]\w*', t)
            w = m.group(0) if m else ''
            rest = t[len(w):]
            if w in TMDL_OBJECTS and not rest.startswith(':') and not rest.lstrip().startswith('='):
                name, after = _split_name(rest)
                expr = gather(d + 2, after[1:].strip()) if after.startswith('=') else None
                kids, kprops = block(d + 1)
                nodes.append({'kind': w, 'name': name, 'expr': expr, 'props': kprops, 'children': kids,
                              'description': ' '.join(desc) or None})
            elif rest.startswith(':'):
                props[w] = rest[1:].strip()
            elif rest.lstrip().startswith('='):
                props[w] = gather(d + 2, rest.lstrip()[1:].strip())
            else:
                props[w or t] = True
            desc = []
        return nodes, props

    return block(0)[0]


def col_ref(s):
    """"'Awareness Emails'.'Date Sent'" or "Date.Date" -> ('Awareness Emails', 'Date Sent')."""
    s = (s or '').strip()
    parts, cur, i = [], '', 0
    while i < len(s):
        ch = s[i]
        if ch == "'":
            j = i + 1
            while j < len(s):
                if s[j] == "'" and s[j + 1:j + 2] == "'":
                    cur += "'"
                    j += 2
                    continue
                if s[j] == "'":
                    break
                cur += s[j]
                j += 1
            i = j + 1
            continue
        if ch == '.' and len(parts) == 0:
            parts.append(cur)
            cur = ''
        else:
            cur += ch
        i += 1
    parts.append(cur)
    return (parts[0].strip(), parts[1].strip()) if len(parts) == 2 else (None, parts[0].strip())


def m_param_value(expr):
    """'"abc" meta [IsParameterQuery=true...]' -> 'abc'; 'true meta [...]' -> True; else None."""
    if not expr or 'meta' not in expr:
        return None
    head = expr.split('meta')[0].strip()
    m = re.match(r'^"((?:[^"]|"")*)"$', head)
    if m:
        return m.group(1).replace('""', '"')
    if head in ('true', 'false'):
        return head == 'true'
    try:
        return float(head)
    except ValueError:
        return head


def _new_table(name, description=None):
    return {'name': name, 'description': description, 'hidden': False, 'columns': {}, 'measures': {},
            'partitions': [], 'is_date_table': False, 'calc_group': False, 'hierarchies': {}}


def model_from_tmdl(folder):
    nodes = []
    for dp, dirs, fs in os.walk(folder):
        dirs[:] = [d for d in dirs if d.lower() != 'cultures']
        for fn in sorted(fs):
            if fn.lower().endswith('.tmdl'):
                with open(os.path.join(dp, fn), encoding='utf-8-sig') as f:
                    nodes.extend(parse_tmdl(f.read()))
    model = {'format': 'tmdl', 'path': folder, 'tables': {}, 'relationships': [], 'roles': [], 'expressions': {}}
    for n in nodes:
        k = n['kind']
        if k == 'table':
            t = _new_table(n['name'], n['description'])
            t['hidden'] = bool(n['props'].get('isHidden'))
            t['is_date_table'] = n['props'].get('dataCategory') == 'Time'
            for c in n['children']:
                p = c['props']
                if c['kind'] == 'column':
                    t['columns'][c['name']] = {
                        'name': c['name'], 'dataType': p.get('dataType'), 'sourceColumn': p.get('sourceColumn'),
                        'expression': c['expr'], 'kind': 'calculated' if c['expr'] else 'data',
                        'hidden': bool(p.get('isHidden')), 'formatString': p.get('formatString'),
                        'dataCategory': p.get('dataCategory'), 'summarizeBy': p.get('summarizeBy'),
                        'description': c['description'], 'isKey': bool(p.get('isKey')),
                        'sortByColumn': (p.get('sortByColumn') or '').strip("'") or None}
                elif c['kind'] == 'measure':
                    t['measures'][c['name']] = {
                        'name': c['name'], 'table': n['name'], 'expression': c['expr'] or '',
                        'formatString': p.get('formatString'), 'displayFolder': p.get('displayFolder'),
                        'description': c['description'], 'hidden': bool(p.get('isHidden')),
                        'dynamic_format': any(x['kind'] == 'formatStringDefinition' for x in c['children'])}
                elif c['kind'] == 'partition':
                    kind = (c['expr'] or '').strip()
                    t['partitions'].append({'name': c['name'], 'kind': kind, 'mode': p.get('mode'),
                                            'source': p.get('source') or p.get('expression') or ''})
                elif c['kind'] == 'calculationGroup':
                    t['calc_group'] = True
                elif c['kind'] == 'hierarchy':
                    t['hierarchies'][c['name']] = [lv['props'].get('column') or lv['name']
                                                   for lv in c['children'] if lv['kind'] == 'level']
            model['tables'][t['name']] = t
        elif k == 'relationship':
            p = n['props']
            ft, fc = col_ref(p.get('fromColumn'))
            tt, tc = col_ref(p.get('toColumn'))
            model['relationships'].append({
                'name': n['name'], 'from_table': ft, 'from_column': fc, 'to_table': tt, 'to_column': tc,
                'active': str(p.get('isActive', 'true')).lower() != 'false',
                'cross': p.get('crossFilteringBehavior') or 'oneDirection',
                'from_card': p.get('fromCardinality') or 'many', 'to_card': p.get('toCardinality') or 'one'})
        elif k == 'role':
            model['roles'].append({'name': n['name'], 'description': n['description'],
                                   'filters': {c['name']: c['expr'] or '' for c in n['children']
                                               if c['kind'] == 'tablePermission'},
                                   'members': [c['name'] for c in n['children'] if c['kind'] == 'member']})
        elif k == 'expression':
            model['expressions'][n['name']] = {'expr': n['expr'] or '', 'value': m_param_value(n['expr']),
                                               'description': n['description']}
    return model


def model_from_bim(path):
    with open(path, encoding='utf-8-sig') as f:
        bim = json.load(f)
    m = bim.get('model', bim)
    j = lambda v: '\n'.join(v) if isinstance(v, list) else (v or '')
    model = {'format': 'bim', 'path': path, 'tables': {}, 'relationships': [], 'roles': [], 'expressions': {}}
    for tb in m.get('tables', []):
        t = _new_table(tb['name'], tb.get('description'))
        t['hidden'] = bool(tb.get('isHidden'))
        t['is_date_table'] = tb.get('dataCategory') == 'Time'
        t['calc_group'] = 'calculationGroup' in tb
        for c in tb.get('columns', []):
            calc = c.get('type') == 'calculated'
            t['columns'][c['name']] = {'name': c['name'], 'dataType': c.get('dataType'), 'sourceColumn': c.get('sourceColumn'),
                                       'expression': j(c.get('expression')) if calc else None,
                                       'kind': 'calculated' if calc else 'data', 'hidden': bool(c.get('isHidden')),
                                       'formatString': c.get('formatString'), 'dataCategory': c.get('dataCategory'),
                                       'summarizeBy': c.get('summarizeBy'), 'description': c.get('description'),
                                       'isKey': bool(c.get('isKey'))}
        for ms in tb.get('measures', []):
            t['measures'][ms['name']] = {'name': ms['name'], 'table': tb['name'], 'expression': j(ms.get('expression')),
                                         'formatString': ms.get('formatString'), 'displayFolder': ms.get('displayFolder'),
                                         'description': ms.get('description'), 'hidden': bool(ms.get('isHidden')),
                                         'dynamic_format': 'formatStringDefinition' in ms}
        for pt in tb.get('partitions', []):
            src = pt.get('source', {})
            t['partitions'].append({'name': pt.get('name'), 'kind': src.get('type'), 'mode': pt.get('mode'),
                                    'source': j(src.get('expression') or src.get('query'))})
        for h in tb.get('hierarchies', []):
            t['hierarchies'][h['name']] = [lv.get('column') for lv in h.get('levels', [])]
        model['tables'][t['name']] = t
    for r in m.get('relationships', []):
        model['relationships'].append({'name': r.get('name'), 'from_table': r.get('fromTable'), 'from_column': r.get('fromColumn'),
                                       'to_table': r.get('toTable'), 'to_column': r.get('toColumn'),
                                       'active': r.get('isActive', True), 'cross': r.get('crossFilteringBehavior') or 'oneDirection',
                                       'from_card': r.get('fromCardinality') or 'many', 'to_card': r.get('toCardinality') or 'one'})
    for ro in m.get('roles', []):
        model['roles'].append({'name': ro['name'], 'description': ro.get('description'),
                               'filters': {tp['name']: j(tp.get('filterExpression')) for tp in ro.get('tablePermissions', [])},
                               'members': [x.get('memberName') for x in ro.get('members', [])]})
    for e in m.get('expressions', []):
        ex = j(e.get('expression'))
        model['expressions'][e['name']] = {'expr': ex, 'value': m_param_value(ex), 'description': e.get('description')}
    return model


def _call_args(text, start):
    """Text of a call's arguments from just after '(' to its matching ')', skipping M string literals."""
    depth, i, n = 1, start, len(text)
    while i < n and depth:
        ch = text[i]
        if ch == '"':
            j = text.find('"', i + 1)
            while j != -1 and text[j + 1:j + 2] == '"':
                j = text.find('"', j + 2)
            i = n if j == -1 else j + 1
            continue
        depth += ch == '('
        depth -= ch == ')'
        i += 1
    return text[start:i]


def enrich_model(model):
    """Derived lookups: M renames (so a renamed column still knows its source column), M-added columns,
    string literals per table (to find the source table/view) and a case-insensitive name index."""
    shared = '\n'.join(e['expr'] for e in model['expressions'].values())
    for t in model['tables'].values():
        m = '\n'.join(p['source'] for p in t['partitions'])
        t['m'] = m
        renames = {}
        text = m + '\n' + shared
        for call in re.finditer(r'Table\.RenameColumns\s*\(', text):
            for old, new in M_PAIR.findall(_call_args(text, call.end())):
                renames[new.replace('""', '"')] = old.replace('""', '"')
        t['m_renames'] = renames
        t['m_added'] = {a.replace('""', '"') for a in re.findall(r'Table\.AddColumn\s*\([^,]+,\s*"((?:[^"]|"")*)"', m + '\n' + shared)}
        t['literals'] = [s.replace('""', '"') for s in M_STRING.findall(m)]
        for c in t['columns'].values():
            sc = c.get('sourceColumn')
            if sc and sc.startswith('[') and sc.endswith(']'):  # calculated-table columns: [Value]
                sc = sc[1:-1]
            c['source_root'] = renames.get(sc, sc) if sc else None
        dt_cols = [c for c in t['columns'].values() if (c['dataType'] or '').lower() == 'datetime']
        if not t['is_date_table'] and re.match(r'^(date|dates|calendar|dim ?date)$', t['name'], re.I) and dt_cols:
            t['is_date_table'] = True
    model['measures'] = {}
    for t in model['tables'].values():
        for ms in t['measures'].values():
            model['measures'].setdefault(ms['name'], ms)
    return model


# ---------- report ----------
def literal(v):
    """PBIR literal string -> python value: "'Y'" -> 'Y', '12L' -> 12, '1.5D' -> 1.5, 'true' -> True."""
    if not isinstance(v, str):
        return v
    if v.startswith("'") and v.endswith("'"):
        return v[1:-1].replace("''", "'")
    if v in ('true', 'false'):
        return v == 'true'
    if v == 'null':
        return None
    m = re.match(r"^datetime'([^']+)'$", v)
    if m:
        return m.group(1)[:19]
    m = re.match(r'^(-?\d+(?:\.\d+)?(?:E[-+]?\d+)?)([LDM]?)$', v, re.I)
    if m:
        x = float(m.group(1))
        return int(x) if m.group(2).upper() == 'L' or x.is_integer() and m.group(2) == '' else x
    return v


def field_refs(obj, aliases=None):
    """Every model field referenced anywhere inside a PBIR JSON fragment -> list of dicts
    {kind: Column|Measure|Hierarchy, table, name, agg}."""
    out = []
    aliases = dict(aliases or {})

    def entity(expr):
        sr = (expr or {}).get('SourceRef') or {}
        return sr.get('Entity') or aliases.get(sr.get('Source'))

    def walk(o):
        if isinstance(o, dict):
            if 'From' in o and isinstance(o['From'], list):
                for fr in o['From']:
                    if isinstance(fr, dict) and fr.get('Name') and fr.get('Entity'):
                        aliases[fr['Name']] = fr['Entity']
            for kind in ('Column', 'Measure'):
                v = o.get(kind)
                if isinstance(v, dict) and 'Property' in v and isinstance(v.get('Expression'), dict):
                    ex = v['Expression']
                    if 'SourceRef' in ex:
                        out.append({'kind': kind, 'table': entity(ex), 'name': v['Property'], 'agg': None})
                    elif 'PropertyVariationSource' in ex:  # auto date hierarchy on a date column
                        pv = ex['PropertyVariationSource']
                        out.append({'kind': 'Column', 'table': entity(pv.get('Expression')), 'name': pv.get('Property'), 'agg': None})
            hl = o.get('HierarchyLevel')
            if isinstance(hl, dict):
                h = (hl.get('Expression') or {}).get('Hierarchy') or {}
                hx = h.get('Expression') or {}
                if 'PropertyVariationSource' in hx:
                    pv = hx['PropertyVariationSource']
                    out.append({'kind': 'Column', 'table': entity(pv.get('Expression')), 'name': pv.get('Property'),
                                'agg': None, 'level': hl.get('Level')})
                elif 'SourceRef' in hx:
                    out.append({'kind': 'Hierarchy', 'table': entity(hx), 'name': h.get('Hierarchy'), 'agg': None,
                                'level': hl.get('Level')})
                return
            ag = o.get('Aggregation')
            if isinstance(ag, dict) and isinstance(ag.get('Expression'), dict):
                inner = field_refs(ag['Expression'], aliases)
                for r in inner:
                    r['agg'] = ag.get('Function')
                out.extend(inner)
                return
            for k, v in o.items():
                if k in ('Column', 'Measure') and isinstance(v, dict) and 'Property' in v:
                    continue
                walk(v)
        elif isinstance(o, list):
            for x in o:
                walk(x)
    walk(obj)
    return out


def filter_info(f, aliases=None):
    """PBIR filter -> {field, type, values, not, how, hidden, condition}."""
    fl = f.get('filter') or {}
    # a Top N filter's 'subquery' source has no Entity (it is a query): only table sources are aliases
    al = {fr['Name']: fr['Entity'] for fr in fl.get('From', []) if fr.get('Name') and fr.get('Entity')}
    al.update(aliases or {})
    refs = field_refs(f.get('field') or {}, al)
    vals, neg, cond = [], False, None
    for w in fl.get('Where', []):
        c = w.get('Condition') or {}
        if 'Not' in c:
            neg = True
            c = c['Not'].get('Expression') or {}
        if 'In' in c:
            for row in c['In'].get('Values', []):
                for cell in row:
                    vals.append(literal((cell.get('Literal') or {}).get('Value')))
        elif c:
            cond = json.dumps(c)[:300]
    return {'field': refs[0] if refs else None, 'type': f.get('type'), 'values': vals, 'not': neg,
            'how': f.get('howCreated'), 'hidden': bool(f.get('isHiddenInViewMode')), 'condition': cond,
            'has_filter': bool(fl)}


def _lit_prop(objs, obj, prop):
    try:
        e = objs[obj][0]['properties'][prop]['expr']
    except (KeyError, IndexError, TypeError):
        return None
    if 'Literal' in e:
        return literal(e['Literal'].get('Value'))
    refs = field_refs(e)
    return f'[{refs[0]["name"]}]' if refs else None


def merge_card_groups(visuals):
    """Cards the Bhayankar build split from one multi-measure card (annotation bhayankarCardGroup) are read back as
    the one card they came from, so a checker compares fields against the Tableau sheet, not against a single part."""
    groups, out = {}, []
    for v in visuals:
        g = v.get('group')
        if not g:
            out.append(v)
        elif g not in groups:
            groups[g] = v
            out.append(v)
        else:
            lead = groups[g]
            top, bottom = min(lead['y'], v['y']), max(lead['y'] + lead['h'], v['y'] + v['h'])
            lead['y'], lead['h'] = top, bottom - top
            lead['fields'] += v['fields']
            lead['all_refs'] += v['all_refs']
            if v['title'] and not lead['title']:  # the part that carries the card's title may sort after the first one
                lead['title'], lead['title_show'] = v['title'], v['title_show']
            lead['filters'] += v['filters']
            for role, rs in v['roles'].items():
                lead['roles'].setdefault(role, []).extend(rs)
    return out


def visual_from_pbir(d, page_id):
    v = d.get('visual') or {}
    pos = d.get('position') or {}
    vco = v.get('visualContainerObjects') or {}
    objs = v.get('objects') or {}
    out = {'id': d.get('name'), 'page': page_id, 'type': v.get('visualType') or ('group' if 'visualGroup' in d else None),
           'x': pos.get('x', 0), 'y': pos.get('y', 0), 'w': pos.get('width', 0), 'h': pos.get('height', 0),
           'hidden': bool(d.get('isHidden')), 'parent': d.get('parentGroupName'),
           'group': next((a.get('value') for a in d.get('annotations') or [] if a.get('name') == 'bhayankarCardGroup'), None),
           'title': _lit_prop(vco, 'title', 'text'), 'title_show': _lit_prop(vco, 'title', 'show'),
           'roles': {}, 'fields': [], 'all_refs': field_refs(d), 'filters': [], 'nav': None, 'text': None,
           'objects': sorted(set(objs) | set(vco)), 'sort': None, 'slicer_selection': None,
           'slicer_header': _lit_prop(objs, 'header', 'text'), 'slicer_mode': _lit_prop(objs, 'data', 'mode')}
    qs = ((v.get('query') or {}).get('queryState')) or {}
    for role, spec in qs.items():
        for pj in (spec or {}).get('projections', []):
            for r in field_refs(pj.get('field') or {}):
                r['display'] = pj.get('displayName')
                r['role'] = role
                out['roles'].setdefault(role, []).append(r)
                out['fields'].append(r)
    sd = (v.get('query') or {}).get('sortDefinition')
    if sd:
        out['sort'] = [(r['name'], s.get('direction')) for s in sd.get('sort', []) for r in field_refs(s.get('field') or {})]
    for f in (d.get('filterConfig') or {}).get('filters', []):
        out['filters'].append(filter_info(f))
    sel = None
    try:
        sel = objs['general'][0]['properties']['filter']['filter']
    except (KeyError, IndexError, TypeError):
        pass
    if sel:
        out['slicer_selection'] = filter_info({'filter': sel, 'field': {}})['values']
    link = vco.get('visualLink') or objs.get('visualLink')
    if link:
        p = link[0].get('properties', {})
        g = lambda k: literal(((p.get(k) or {}).get('expr') or {}).get('Literal', {}).get('Value'))
        out['nav'] = {'type': g('type'), 'page': g('navigationSection'), 'drillthrough': g('drillthroughSection'),
                      'url': g('webUrl'), 'bookmark': g('bookmark'), 'show': g('show')}
    if out['type'] == 'textbox':
        try:
            out['text'] = ' '.join(r.get('value', '') for para in objs['general'][0]['properties']['paragraphs']
                                   for r in para.get('textRuns', []))
        except (KeyError, IndexError, TypeError):
            pass
    out['ref_lines'] = any(k in objs for k in ('y1AxisReferenceLine', 'xAxisReferenceLine', 'referenceLine', 'constantLine')) \
        or 'ReferenceLine' in json.dumps(objs)[:200000]
    return out


def report_from_pbir(definition):
    rep = {'format': 'pbir', 'path': definition, 'pages': [], 'filters': []}
    rj = os.path.join(definition, 'report.json')
    if os.path.exists(rj):
        with open(rj, encoding='utf-8-sig') as f:
            r = json.load(f)
        rep['filters'] = [filter_info(x) for x in (r.get('filterConfig') or {}).get('filters', [])]
    pdir = os.path.join(definition, 'pages')
    order = []
    if os.path.exists(os.path.join(pdir, 'pages.json')):
        with open(os.path.join(pdir, 'pages.json'), encoding='utf-8-sig') as f:
            order = json.load(f).get('pageOrder', [])
    ids = [d for d in (os.listdir(pdir) if os.path.isdir(pdir) else []) if os.path.isdir(os.path.join(pdir, d))]
    for pid in sorted(ids, key=lambda x: order.index(x) if x in order else 999):
        pj = os.path.join(pdir, pid, 'page.json')
        if not os.path.exists(pj):
            continue
        with open(pj, encoding='utf-8-sig') as f:
            p = json.load(f)
        pb = p.get('pageBinding') or {}
        page = {'id': pid, 'name': p.get('displayName') or pid, 'hidden': p.get('visibility') == 'HiddenInViewMode',
                'type': pb.get('type') or p.get('type') or ('Tooltip' if (p.get('displayOption') == 'ActualSize' and
                                                                          (p.get('height') or 0) <= 400) else None),
                'width': p.get('width'), 'height': p.get('height'),
                'filters': [filter_info(x) for x in (p.get('filterConfig') or {}).get('filters', [])],
                'interactions': p.get('visualInteractions') or [], 'visuals': []}
        if pb.get('type') == 'Tooltip' or p.get('type') == 'Tooltip':
            page['type'] = 'Tooltip'
        vdir = os.path.join(pdir, pid, 'visuals')
        for vid in (sorted(os.listdir(vdir)) if os.path.isdir(vdir) else []):
            vj = os.path.join(vdir, vid, 'visual.json')
            if os.path.exists(vj):
                with open(vj, encoding='utf-8-sig') as f:
                    page['visuals'].append(visual_from_pbir(json.load(f), pid))
        page['visuals'] = merge_card_groups(page['visuals'])
        rep['pages'].append(page)
    return rep


def report_from_legacy(path):
    """Legacy report.json (sections / visualContainers with JSON-string configs)."""
    with open(path, encoding='utf-8-sig') as f:
        r = json.load(f)
    rep = {'format': 'legacy', 'path': path, 'pages': [], 'filters': []}
    js = lambda s: json.loads(s) if isinstance(s, str) and s.strip() else (s or {})
    rep['filters'] = [filter_info(x) for x in js(r.get('filters')) or []] if isinstance(js(r.get('filters')), list) else []
    for i, s in enumerate(r.get('sections', [])):
        cfg = js(s.get('config'))
        page = {'id': s.get('name') or str(i), 'name': s.get('displayName') or s.get('name'),
                'hidden': cfg.get('visibility') == 1, 'type': {1: 'Drillthrough', 2: 'Tooltip'}.get(cfg.get('type')),
                'width': s.get('width'), 'height': s.get('height'),
                'filters': [filter_info(x) for x in (js(s.get('filters')) or [])], 'interactions': [], 'visuals': []}
        for vc in s.get('visualContainers', []):
            c = js(vc.get('config'))
            sv = c.get('singleVisual') or {}
            vis = {'id': c.get('name'), 'page': page['id'], 'type': sv.get('visualType') or ('group' if 'singleVisualGroup' in c else None),
                   'x': vc.get('x', 0), 'y': vc.get('y', 0), 'w': vc.get('width', 0), 'h': vc.get('height', 0),
                   'hidden': False, 'parent': c.get('parentGroupName'), 'title': _lit_prop(sv.get('vcObjects') or {}, 'title', 'text'),
                   'title_show': None, 'roles': {}, 'fields': [], 'all_refs': field_refs(c), 'objects': sorted(sv.get('objects') or {}),
                   'filters': [filter_info(x) for x in (js(vc.get('filters')) or [])], 'nav': None, 'text': None,
                   'sort': None, 'slicer_selection': None, 'slicer_header': None, 'slicer_mode': None, 'ref_lines': False}
            q = sv.get('prototypeQuery') or {}
            al = {fr['Name']: fr['Entity'] for fr in q.get('From', []) if fr.get('Name') and fr.get('Entity')}
            names = {}
            for sel in q.get('Select', []):
                for ref in field_refs(sel, al):
                    names[sel.get('Name')] = ref
            for role, lst in (sv.get('projections') or {}).items():
                for pj in lst:
                    ref = names.get(pj.get('queryRef'))
                    if ref:
                        ref = dict(ref, role=role, display=None)
                        vis['roles'].setdefault(role, []).append(ref)
                        vis['fields'].append(ref)
            page['visuals'].append(vis)
        rep['pages'].append(page)
    return rep


# ---------- project ----------
def locate(path):
    """A .pbip file, a project folder, or a .SemanticModel/.Report folder -> (pbip name, model path, report path)."""
    path = os.path.abspath(path)
    sm = rp = None
    name = None
    if path.lower().endswith('.pbip') and os.path.isfile(path):
        base = os.path.dirname(path)
        name = os.path.splitext(os.path.basename(path))[0]
        try:
            with open(path, encoding='utf-8-sig') as f:
                pj = json.load(f)
            for a in pj.get('artifacts', []):
                if 'report' in a:
                    rp = os.path.join(base, a['report'].get('path', ''))
        except (ValueError, OSError):
            pass
        rp = rp if rp and os.path.isdir(rp) else os.path.join(base, name + '.Report')
    elif path.lower().endswith('.report'):
        rp, name = path, os.path.splitext(os.path.basename(path))[0]
    elif path.lower().endswith('.semanticmodel'):
        sm, name = path, os.path.splitext(os.path.basename(path))[0]
    elif os.path.isdir(path):
        pb = [f for f in os.listdir(path) if f.lower().endswith('.pbip')]
        if pb:
            return locate(os.path.join(path, pb[0]))
        reps = [f for f in os.listdir(path) if f.lower().endswith('.report')]
        if reps:
            return locate(os.path.join(path, reps[0]))
    if rp and os.path.isdir(rp) and not sm:
        pbir = os.path.join(rp, 'definition.pbir')
        if os.path.exists(pbir):
            try:
                with open(pbir, encoding='utf-8-sig') as f:
                    ref = (json.load(f).get('datasetReference') or {}).get('byPath') or {}
                if ref.get('path'):
                    sm = os.path.normpath(os.path.join(rp, ref['path']))
            except (ValueError, OSError):
                pass
        if not sm or not os.path.isdir(sm):
            cand = os.path.join(os.path.dirname(rp), (name or '') + '.SemanticModel')
            sm = cand if os.path.isdir(cand) else None
    if sm and not rp:
        cand = os.path.join(os.path.dirname(sm), name + '.Report')
        rp = cand if os.path.isdir(cand) else None
    return name, sm, rp


def load_project(path):
    """-> {'name', 'model', 'report', 'errors'} (model/report None when missing)."""
    name, sm, rp = locate(path)
    proj = {'name': name, 'model_path': sm, 'report_path': rp, 'model': None, 'report': None, 'errors': []}
    if sm:
        try:
            if os.path.isdir(os.path.join(sm, 'definition')):
                proj['model'] = enrich_model(model_from_tmdl(os.path.join(sm, 'definition')))
            elif os.path.exists(os.path.join(sm, 'model.bim')):
                proj['model'] = enrich_model(model_from_bim(os.path.join(sm, 'model.bim')))
            else:
                proj['errors'].append(f'{sm}: no definition\\ (TMDL) or model.bim')
        except Exception as e:  # report and keep testing what can be read
            proj['errors'].append(f'semantic model unreadable: {e}')
    else:
        proj['errors'].append('no .SemanticModel folder found (live-connected report or wrong path)')
    if rp:
        try:
            if os.path.isdir(os.path.join(rp, 'definition', 'pages')) or os.path.exists(os.path.join(rp, 'definition', 'report.json')):
                proj['report'] = report_from_pbir(os.path.join(rp, 'definition'))
            elif os.path.exists(os.path.join(rp, 'report.json')):
                proj['report'] = report_from_legacy(os.path.join(rp, 'report.json'))
            else:
                proj['errors'].append(f'{rp}: no definition\\ (PBIR) or report.json')
        except Exception as e:
            proj['errors'].append(f'report unreadable: {e}')
    else:
        proj['errors'].append('no .Report folder found')
    return proj
