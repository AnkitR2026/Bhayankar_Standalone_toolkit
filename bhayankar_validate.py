#!/usr/bin/env python3
r"""Bhayankar offline PBIR validator: checks a Power BI project's report files with no internet and no Node.js.

Runs the rules of `powerbi-report-author validate` (Microsoft powerbi-report-authoring-cli 0.4.0, MIT) in Python:
  - JSON Schema validation against the published PBIR schemas, bundled in resources\schemas (nothing is downloaded)
  - project files: .pbip, .platform, definition.pbir, version.json, report.json, pages.json, page.json, visual.json
  - visuals: type known and not deprecated, roles and projections, formatting objects and properties, container objects
  - layout: visuals inside the page, dropdown slicer and textbox minimum heights
  - filters, drillthrough pages, visual interactions, resource packages
plus three checks the CLI does not make (code prefix BHY_): resource files on disk, the semantic model folder the
report binds to, and the Windows 260-character path limit (Power BI Desktop cannot open a longer path).

The visual metadata (roles, formatting objects, properties) is a snapshot in resources\pbir_metadata.json, made by
tools\snapshot_pbir_metadata.mjs. The result has the CLI's shape:
  {result, errorCount, warningCount, reportPath, diagnostics: {code: {severity, items: [{message, file, path}]}}}

Usage:
  python bhayankar_validate.py "<project>.pbip"                 (or the .Report folder)
  python bhayankar_validate.py "<project>.pbip" --format text
  python bhayankar_validate.py "<project>.pbip" --no-schema     (structure and visual rules only)
Python 3 standard library only.
"""
import argparse
import json
import math
import ntpath
import os
import posixpath
import re
import sys
from urllib.parse import urljoin

HERE = os.path.dirname(os.path.abspath(__file__))
RESOURCES = os.path.join(HERE, 'resources')
SCHEMA_DIR = os.path.join(RESOURCES, 'schemas', 'fabric')
META_FILE = os.path.join(RESOURCES, 'pbir_metadata.json')
FABRIC = 'https://developer.microsoft.com/json-schemas/fabric'
RDEF = FABRIC + '/item/report/definition'
MAX_PATH = 259  # Windows MAX_PATH (260) minus the terminating NUL
VALIDATOR = 'bhayankar_validate (offline; PBIR rules of powerbi-report-authoring-cli 0.4.0)'


def _family(label, root):
    return label, re.compile('^' + re.escape(root) + r'/\d+\.\d+\.\d+/schema\.json$')


FAMILIES = dict([
    _family('pbipProperties', FABRIC + '/pbip/pbipProperties'),
    _family('platformProperties', FABRIC + '/gitIntegration/platformProperties'),
    _family('definitionProperties', FABRIC + '/item/report/definitionProperties'),
    _family('versionMetadata', RDEF + '/versionMetadata'),
    _family('report', RDEF + '/report'),
    _family('pagesMetadata', RDEF + '/pagesMetadata'),
    _family('page', RDEF + '/page'),
    _family('visualContainer', RDEF + '/visualContainer'),
    _family('visualContainerMobileState', RDEF + '/visualContainerMobileState'),
    _family('bookmarksMetadata', RDEF + '/bookmarksMetadata'),
    _family('bookmark', RDEF + '/bookmark'),
    _family('reportExtension', RDEF + '/reportExtension'),
])
VALID_PAGE_OBJECTS = {
    'background': ['color', 'image', 'transparency'],
    'outspace': ['color', 'image', 'transparency'],
    'outspacePane': ['backgroundColor', 'transparency', 'foregroundColor', 'titleSize', 'headerSize', 'searchTextSize',
                     'fontFamily', 'border', 'borderColor', 'checkboxAndApplyColor', 'inputBoxColor', 'width'],
    'filterCard': ['backgroundColor', 'transparency', 'foregroundColor', 'textSize', 'fontFamily', 'inputBoxColor',
                   'border', 'borderColor'],
    'pageInformation': None, 'pageSize': None, 'displayArea': None, 'pageRefresh': None, 'personalizeVisual': None,
}
DATA_BOUND = {'Aggregation', 'Column', 'Conditional', 'FillRule', 'Measure'}
GUID = re.compile(r'^(?:[0-9a-fA-F]{32}|[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}|'
                  r'\{[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}|'
                  r'\([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\))$')
_META = None


def metadata():
    global _META
    if _META is None:
        with open(META_FILE, encoding='utf-8') as f:
            _META = json.load(f)
    return _META


def is_obj(v):
    return isinstance(v, dict)


def is_abs(p):
    return posixpath.isabs(p) or ntpath.isabs(p)


def number(v):
    """v when it is a JSON number, else None (JavaScript comparisons with a non-number are false)."""
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


# =====================================================================================
# diagnostics
# =====================================================================================
class Collector:
    def __init__(self):
        self.items = []

    def error(self, code, message, file=None, path=None):
        self.items.append(('error', code, message, file, path))

    def warning(self, code, message, file=None, path=None):
        self.items.append(('warning', code, message, file, path))

    def summary(self, report_path):
        grouped, errors, warnings = {}, 0, 0
        for sev, code, message, file, path in self.items:
            if sev == 'error':
                errors += 1
            else:
                warnings += 1
            g = grouped.setdefault(code, {'severity': sev, 'items': []})
            item = {'message': message}
            if file is not None:
                item['file'] = file
            if path is not None:
                item['path'] = path
            g['items'].append(item)
        result = 'failed' if errors else 'succeededWithWarnings' if warnings else 'succeeded'
        out = {'result': result, 'errorCount': errors, 'warningCount': warnings, 'reportPath': report_path,
               'validator': VALIDATOR}
        if result != 'succeeded':
            out['diagnostics'] = grouped
        return out


# =====================================================================================
# JSON Schema (draft-07 subset used by the PBIR schemas)
# =====================================================================================
def json_eq(a, b):
    """JSON equality: true is not 1, 1 equals 1.0."""
    if isinstance(a, bool) or isinstance(b, bool):
        return isinstance(a, bool) and isinstance(b, bool) and a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return a == b
    if type(a) is not type(b):
        return False
    if isinstance(a, dict):
        return a.keys() == b.keys() and all(json_eq(a[k], b[k]) for k in a)
    if isinstance(a, list):
        return len(a) == len(b) and all(json_eq(x, y) for x, y in zip(a, b))
    return a == b


TYPE_CHECK = {
    'object': lambda v: isinstance(v, dict),
    'array': lambda v: isinstance(v, list),
    'string': lambda v: isinstance(v, str),
    'boolean': lambda v: isinstance(v, bool),
    'null': lambda v: v is None,
    'integer': lambda v: (isinstance(v, int) and not isinstance(v, bool)) or (isinstance(v, float) and v.is_integer()),
    'number': lambda v: isinstance(v, (int, float)) and not isinstance(v, bool),
}


def ptr_escape(k):
    return str(k).replace('~', '~0').replace('/', '~1')


class SchemaStore:
    """The bundled schemas, addressed by their published URL (https://developer.microsoft.com/json-schemas/fabric/...)."""

    def __init__(self, root=SCHEMA_DIR):
        self.root = root
        self.docs = {}

    def load(self, url):
        if url in self.docs:
            return self.docs[url]
        doc = None
        if url.startswith(FABRIC + '/'):
            path = os.path.join(self.root, *url[len(FABRIC) + 1:].split('/'))
            if os.path.isfile(path):
                with open(path, encoding='utf-8-sig') as f:
                    doc = json.load(f)
        self.docs[url] = doc
        return doc

    def resolve(self, base, ref):
        url = urljoin(base, ref)
        doc_url, _, frag = url.partition('#')
        node = self.load(doc_url)
        for part in [p for p in frag.split('/') if p]:
            part = part.replace('~1', '/').replace('~0', '~')
            if isinstance(node, dict) and part in node:
                node = node[part]
            elif isinstance(node, list) and part.isdigit() and int(part) < len(node):
                node = node[int(part)]
            else:
                return doc_url, None
        return doc_url, node


class SchemaValidator:
    MAX_ERRORS = 50

    def __init__(self, store):
        self.store = store

    def validate(self, instance, url):
        """[(instancePath, message, additionalProperty)] for one document; None when the schema is not bundled."""
        schema = self.store.load(url)
        if schema is None:
            return None
        errs = []
        self.check(instance, schema, url, '', errs)
        return errs[:self.MAX_ERRORS]

    def check(self, inst, schema, base, path, errs):
        if schema is True or schema == {}:
            return
        if schema is False:
            errs.append((path, 'boolean schema is false', None))
            return
        if not isinstance(schema, dict):
            return
        if '$ref' in schema:  # draft-07: keywords next to $ref are ignored
            doc_url, target = self.store.resolve(base, schema['$ref'])
            if target is None:
                errs.append((path, f"can't resolve reference {schema['$ref']}", None))
            else:
                self.check(inst, target, doc_url, path, errs)
            return
        t = schema.get('type')
        if t is not None:
            types = t if isinstance(t, list) else [t]
            if not any(TYPE_CHECK.get(x, lambda v: True)(inst) for x in types):
                errs.append((path, f'must be {",".join(types)}', None))
                return
        if 'const' in schema and not json_eq(inst, schema['const']):
            errs.append((path, 'must be equal to constant', None))
        if 'enum' in schema and not any(json_eq(inst, e) for e in schema['enum']):
            errs.append((path, 'must be equal to one of the allowed values', None))
        if isinstance(inst, str):
            if 'minLength' in schema and len(inst) < schema['minLength']:
                errs.append((path, f'must NOT have fewer than {schema["minLength"]} characters', None))
            if 'maxLength' in schema and len(inst) > schema['maxLength']:
                errs.append((path, f'must NOT have more than {schema["maxLength"]} characters', None))
            if 'pattern' in schema and not re.search(schema['pattern'], inst):
                errs.append((path, f'must match pattern "{schema["pattern"]}"', None))
        if isinstance(inst, (int, float)) and not isinstance(inst, bool):
            for kw, ok, word in (('minimum', lambda a, b: a >= b, '>='), ('maximum', lambda a, b: a <= b, '<='),
                                 ('exclusiveMinimum', lambda a, b: a > b, '>'), ('exclusiveMaximum', lambda a, b: a < b, '<')):
                if kw in schema and isinstance(schema[kw], (int, float)) and not ok(inst, schema[kw]):
                    errs.append((path, f'must be {word} {schema[kw]}', None))
        if isinstance(inst, dict):
            for r in schema.get('required', []):
                if r not in inst:
                    errs.append((path, f"must have required property '{r}'", None))
            props = schema.get('properties', {})
            patterns = schema.get('patternProperties', {})
            extra = schema.get('additionalProperties', True)
            for k, v in inst.items():
                matched = False
                if k in props:
                    matched = True
                    self.check(v, props[k], base, f'{path}/{ptr_escape(k)}', errs)
                for pat, sub in patterns.items():
                    if re.search(pat, k):
                        matched = True
                        self.check(v, sub, base, f'{path}/{ptr_escape(k)}', errs)
                if not matched:
                    if extra is False:
                        errs.append((path, 'must NOT have additional properties', k))
                    elif isinstance(extra, dict):
                        self.check(v, extra, base, f'{path}/{ptr_escape(k)}', errs)
        if isinstance(inst, list):
            items = schema.get('items')
            if isinstance(items, dict):
                for i, x in enumerate(inst):
                    self.check(x, items, base, f'{path}/{i}', errs)
            elif isinstance(items, list):
                for i, x in enumerate(inst):
                    if i < len(items):
                        self.check(x, items[i], base, f'{path}/{i}', errs)
                    elif schema.get('additionalItems') is False:
                        errs.append((path, f'must NOT have more than {len(items)} items', None))
                        break
            if 'minItems' in schema and len(inst) < schema['minItems']:
                errs.append((path, f'must NOT have fewer than {schema["minItems"]} items', None))
            if 'maxItems' in schema and len(inst) > schema['maxItems']:
                errs.append((path, f'must NOT have more than {schema["maxItems"]} items', None))
            if schema.get('uniqueItems'):
                for i in range(len(inst)):
                    if any(json_eq(inst[i], inst[j]) for j in range(i)):
                        errs.append((path, 'must NOT have duplicate items', None))
                        break
        for sub in schema.get('allOf', []):
            self.check(inst, sub, base, path, errs)
        if 'anyOf' in schema:
            branches = []
            for sub in schema['anyOf']:
                e = []
                self.check(inst, sub, base, path, e)
                if not e:
                    break
                branches.append(e)
            else:
                errs.append((path, 'must match a schema in anyOf', None))
                errs.extend(self.closest(branches))
        if 'oneOf' in schema:
            branches, passed = [], 0
            for sub in schema['oneOf']:
                e = []
                self.check(inst, sub, base, path, e)
                if e:
                    branches.append(e)
                else:
                    passed += 1
            if passed != 1:
                errs.append((path, 'must match exactly one schema in oneOf', None))
                if not passed:
                    errs.extend(self.closest(branches))
        if 'not' in schema:
            e = []
            self.check(inst, schema['not'], base, path, e)
            if not e:
                errs.append((path, 'must NOT be valid', None))
        if 'if' in schema:
            e = []
            self.check(inst, schema['if'], base, path, e)
            branch = schema.get('then') if not e else schema.get('else')
            if branch is not None:
                self.check(inst, branch, base, path, errs)

    @staticmethod
    def closest(branches):
        """The errors of the anyOf/oneOf branch that matched furthest (deepest path, then fewest errors)."""
        if not branches:
            return []
        return min(branches, key=lambda e: (-max(len(p) for p, _, _ in e), len(e)))


# =====================================================================================
# project discovery
# =====================================================================================
def read_json(path, col):
    """Parsed JSON, or None (with a PBIR_JSON_PARSE_ERROR) when the file is not valid JSON."""
    try:
        with open(path, encoding='utf-8') as f:
            return json.load(f), True
    except ValueError as e:
        col.error('PBIR_JSON_PARSE_ERROR', f'Invalid JSON: {path} — {e}', file=path)
        return None, False


def locate(path):
    """(report_dir, pbip_path) for a .pbip file, a .Report folder or a folder holding definition\\ or one .pbip."""
    path = os.path.abspath(path)
    if os.path.isfile(path) and path.lower().endswith('.pbip'):
        try:
            with open(path, encoding='utf-8') as f:
                data = json.load(f)
        except (OSError, ValueError):
            data = {}
        rel = first_report_path(data) if is_obj(data) else None
        if rel and not is_abs(rel):
            cand = os.path.normpath(os.path.join(os.path.dirname(path), rel.replace('\\', '/')))
            if os.path.isdir(os.path.join(cand, 'definition')):
                return cand, path
        sib = os.path.splitext(path)[0] + '.Report'
        return sib, path
    if os.path.isdir(path):
        if path.lower().endswith('.report') or os.path.isdir(os.path.join(path, 'definition')):
            return path, sibling_pbip(path)
        pbips = [f for f in os.listdir(path) if f.lower().endswith('.pbip')]
        if len(pbips) == 1:
            return locate(os.path.join(path, pbips[0]))
        reports = [f for f in os.listdir(path) if f.lower().endswith('.report') and os.path.isdir(os.path.join(path, f))]
        if len(reports) == 1:
            return os.path.join(path, reports[0]), None
    return path, None


def sibling_pbip(report_dir):
    name = os.path.basename(report_dir.rstrip('\\/'))
    if not name.lower().endswith('.report') or len(name) <= len('.Report'):
        return None
    cand = os.path.join(os.path.dirname(report_dir.rstrip('\\/')), name[:-len('.Report')] + '.pbip')
    return cand if os.path.exists(cand) else None


def first_report_path(data):
    for a in data.get('artifacts') or []:
        if is_obj(a) and is_obj(a.get('report')) and isinstance(a['report'].get('path'), str) and a['report']['path'].strip():
            return a['report']['path']
    for r in data.get('reports') or []:
        if is_obj(r) and isinstance(r.get('path'), str) and r['path'].strip():
            return r['path']
    return None


# =====================================================================================
# the validator
# =====================================================================================
class Validator:
    def __init__(self, report_dir, pbip_path=None, schemas=True):
        self.report_dir = os.path.abspath(report_dir)
        self.def_dir = os.path.join(self.report_dir, 'definition')
        self.pbip_path = pbip_path if pbip_path is not None else sibling_pbip(self.report_dir)
        self.schemas = schemas
        self.col = Collector()
        self.targets = []  # (file, data) for schema validation
        self.filter_names = set()
        self.meta = metadata()

    # ---------- helpers ----------
    def push(self, path, data, family=None, missing_schema=True):
        self.targets.append((path, data))
        if missing_schema and is_obj(data) and not isinstance(data.get('$schema'), str):
            self.col.error('PBIR_JSON_FILE_NO_SCHEMA', 'JSON file has no "$schema" and cannot be schema-validated; Fabric '
                           f'rejects PBIR definition JSON files without "$schema": {path}', file=path)
        if family and is_obj(data) and isinstance(data.get('$schema'), str) and not FAMILIES[family].match(data['$schema']):
            self.col.error('PBIR_JSON_SCHEMA_URL_INVALID', f'JSON file has "$schema" "{data["$schema"]}", but expected a '
                           f'published {family} schema URL for this PBIR file: {path}', file=path, path='$schema')

    def push_if_json(self, path, family=None, missing_code=None, missing_schema=True):
        if not os.path.exists(path):
            if missing_code:
                self.col.error(missing_code, f'Missing: {path}', file=path)
            return None
        data, ok = read_json(path, self.col)
        if not ok:
            return None
        self.push(path, data, family, missing_schema)
        return data

    # ---------- run ----------
    def run(self):
        c = self.col
        if not os.path.isdir(self.def_dir):
            c.error('PBIR_DEFINITION_DIR_MISSING', f'Definition directory not found: {self.def_dir}', file=self.def_dir)
            return c.summary(self.report_dir)
        if self.pbip_path and os.path.exists(self.pbip_path):
            data = self.push_if_json(self.pbip_path, 'pbipProperties', missing_schema=False)
            if data is not None:
                self.pbip(data)
        plat = os.path.join(self.report_dir, '.platform')
        data = self.push_if_json(plat, 'platformProperties', 'PBIR_PLATFORM_MISSING', missing_schema=False)
        if data is not None:
            self.platform(plat, data)
        pbir = os.path.join(self.report_dir, 'definition.pbir')
        data = self.push_if_json(pbir, 'definitionProperties', 'PBIR_DEFINITION_PBIR_MISSING')
        if data is not None:
            self.dataset_reference(pbir, data)
        ver = os.path.join(self.def_dir, 'version.json')
        data = self.push_if_json(ver, 'versionMetadata', 'PBIR_VERSION_MISSING', missing_schema=False)
        if data is not None:
            self.version(ver, data)
        rj = os.path.join(self.def_dir, 'report.json')
        report = None
        if os.path.exists(rj):
            data, ok = read_json(rj, c)
            if ok:
                self.push(rj, data, 'report')
                if is_obj(data):
                    report = data
                else:
                    c.error('PBIR_JSON_ROOT_NOT_OBJECT', f'report.json root must be a JSON object: {rj}', file=rj)
            if report:
                self.filters(rj, report)
                self.resource_packages(rj, report)
        else:
            c.error('PBIR_REPORT_JSON_MISSING', f'Missing: {rj}', file=rj)
        custom = self.custom_visuals(report)
        pj = os.path.join(self.def_dir, 'pages', 'pages.json')
        pages_data, pages_ok = (None, False)
        if os.path.exists(pj):
            pages_data, pages_ok = read_json(pj, c)
            if pages_ok:
                self.push(pj, pages_data, 'pagesMetadata', missing_schema=False)
        self.pages_index(pj, pages_data, os.path.exists(pj), pages_ok)
        for fp in self.walk_json(os.path.join(self.def_dir, 'bookmarks')):
            name = os.path.basename(fp).lower()
            fam = 'bookmarksMetadata' if name == 'bookmarks.json' else 'bookmark' if name.endswith('.bookmark.json') else None
            self.push_if_json(fp, fam)
        self.push_if_json(os.path.join(self.def_dir, 'reportExtensions.json'), 'reportExtension')
        order = []
        if is_obj(pages_data) and isinstance(pages_data.get('pageOrder'), list):
            seen = set()
            for n in pages_data['pageOrder']:
                if not isinstance(n, str):
                    continue
                if n in seen:
                    c.error('PBIR_PAGE_ORDER_DUPLICATE', f'Duplicate page name in pageOrder: "{n}"', file=pj, path=n)
                seen.add(n)
                order.append(n)
        for page in self.page_records(order):
            self.page(page, custom)
        if self.schemas:
            self.schema_pass()
        self.project_checks()
        return c.summary(self.report_dir)

    @staticmethod
    def walk_json(d):
        out = []
        for dp, _, fs in os.walk(d) if os.path.isdir(d) else []:
            out += [os.path.join(dp, f) for f in sorted(fs) if f.lower().endswith('.json')]
        return out

    # ---------- project files ----------
    def pbip(self, data):
        c, p = self.col, self.pbip_path
        if not is_obj(data):
            c.error('PBIR_JSON_ROOT_NOT_OBJECT', f'.pbip root must be a JSON object: {p}', file=p)
            return
        if not isinstance(data.get('$schema'), str):
            c.error('PBIR_JSON_FILE_NO_SCHEMA', f'.pbip must include a "$schema" URL so the report host can validate the '
                    f'item shortcut: {p}', file=p, path='$schema')
        rel = first_report_path(data)
        if rel is None:
            c.error('PBIR_PBIP_ARTIFACTS_EMPTY', f'.pbip must reference a report; expected a non-empty '
                    f'"artifacts[].report.path" (e.g. "Sales.Report"): {p}', file=p, path='artifacts')
            return
        if '\\' in rel:
            c.warning('PBIR_PBIP_REPORT_PATH_SEPARATOR', f'.pbip report path "{rel}" uses a backslash; use "/": {p}',
                      file=p, path='artifacts[].report.path')
        if is_abs(rel):
            c.error('PBIR_PBIP_REPORT_PATH_NOT_RELATIVE', f'.pbip report path "{rel}" is absolute; it must be relative '
                    f'to the .pbip file: {p}', file=p, path='artifacts[].report.path')
            return
        resolved = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(p)), rel.replace('\\', '/')))
        if not os.path.isdir(os.path.join(resolved, 'definition')):
            c.error('PBIR_PBIP_REPORT_PATH_UNRESOLVED', f'.pbip report path "{rel}" does not resolve to a report folder '
                    f'containing "definition" ({resolved}): {p}', file=p, path='artifacts[].report.path')
            return
        if os.path.normcase(resolved) != os.path.normcase(self.report_dir):
            c.warning('PBIR_PBIP_REPORT_PATH_MISMATCH', f'.pbip report path "{rel}" resolves to {resolved}, which is not '
                      f'the report being validated ({self.report_dir}): {p}', file=p, path='artifacts[].report.path')

    def platform(self, p, data):
        c = self.col
        if not is_obj(data):
            c.error('PBIR_JSON_ROOT_NOT_OBJECT', f'.platform root must be a JSON object: {p}', file=p)
            return
        if not isinstance(data.get('$schema'), str) or not data['$schema']:
            c.error('PBIR_JSON_FILE_NO_SCHEMA', f'.platform must include a "$schema" URL: {p}', file=p, path='$schema')
        md = data.get('metadata')
        if not is_obj(md):
            c.error('PBIR_PLATFORM_METADATA_MISSING', f'.platform metadata object is missing: {p}', file=p, path='metadata')
        else:
            if md.get('type') != 'Report':
                c.error('PBIR_PLATFORM_METADATA_TYPE_INVALID', f'.platform metadata.type must be "Report": {p}',
                        file=p, path='metadata.type')
            if not isinstance(md.get('displayName'), str) or not md['displayName']:
                c.error('PBIR_PLATFORM_DISPLAY_NAME_MISSING', f'.platform metadata.displayName must be a non-empty '
                        f'string: {p}', file=p, path='metadata.displayName')
        cfg = data.get('config')
        if not is_obj(cfg):
            c.error('PBIR_PLATFORM_CONFIG_MISSING', f'.platform config object is missing: {p}', file=p, path='config')
        else:
            if not isinstance(cfg.get('version'), str) or not cfg['version']:
                c.error('PBIR_PLATFORM_VERSION_MISSING', f'.platform config.version must be a non-empty string: {p}',
                        file=p, path='config.version')
            lid = cfg.get('logicalId')
            if not isinstance(lid, str) or not lid:
                c.error('PBIR_PLATFORM_LOGICAL_ID_MISSING', f'.platform config.logicalId must be a non-empty string: {p}',
                        file=p, path='config.logicalId')
            elif not GUID.match(lid):
                c.error('PBIR_PLATFORM_LOGICAL_ID_INVALID', f'.platform config.logicalId "{lid}" is not a valid GUID: {p}',
                        file=p, path='config.logicalId')

    def dataset_reference(self, p, data):
        c = self.col
        if not is_obj(data):
            return
        ref = data.get('datasetReference')
        if not is_obj(ref):
            c.error('PBIR_DATASET_REFERENCE_MISSING', f'definition.pbir must include a "datasetReference" object binding '
                    f'the report to a semantic model: {p}', file=p, path='datasetReference')
            return
        by_path, by_conn = ref.get('byPath'), ref.get('byConnection')
        if by_path is None and by_conn is None:
            c.error('PBIR_DATASET_REFERENCE_FORM_MISSING', f'datasetReference must contain exactly one of "byPath" or '
                    f'"byConnection"; neither is present: {p}', file=p, path='datasetReference')
            return
        if by_path is not None and by_conn is not None:
            c.error('PBIR_DATASET_REFERENCE_FORM_AMBIGUOUS', f'datasetReference contains both "byPath" and '
                    f'"byConnection"; exactly one is allowed: {p}', file=p, path='datasetReference')
            return
        null_sibling = 'byPath' if 'byPath' in ref and ref['byPath'] is None else \
            'byConnection' if 'byConnection' in ref and ref['byConnection'] is None else None
        if null_sibling:
            c.error('PBIR_DATASET_REFERENCE_NULL_SIBLING', f'datasetReference has "{null_sibling}": null alongside the '
                    f'form in use; include only the form you are using: {p}', file=p, path=f'datasetReference.{null_sibling}')
        if by_path is not None:
            if not is_obj(by_path):
                return
            mp = by_path.get('path')
            if not isinstance(mp, str) or not mp.strip():
                c.error('PBIR_DATASET_BY_PATH_MISSING', f'datasetReference.byPath.path must be a non-empty relative path '
                        f'to the semantic model folder: {p}', file=p, path='datasetReference.byPath.path')
                return
            if '\\' in mp:
                c.warning('PBIR_DATASET_BY_PATH_SEPARATOR', f'datasetReference.byPath.path "{mp}" uses a backslash; '
                          f'use "/": {p}', file=p, path='datasetReference.byPath.path')
            if is_abs(mp):
                c.error('PBIR_DATASET_BY_PATH_NOT_RELATIVE', f'datasetReference.byPath.path "{mp}" is absolute; it must be '
                        f'relative to the .Report directory: {p}', file=p, path='datasetReference.byPath.path')
                return
            model = os.path.normpath(os.path.join(self.report_dir, mp.replace('\\', '/')))
            if not (os.path.isdir(os.path.join(model, 'definition')) or os.path.isfile(os.path.join(model, 'model.bim'))):
                c.error('BHY_MODEL_PATH_UNRESOLVED', f'datasetReference.byPath.path "{mp}" does not lead to a semantic model '
                        f'folder with a definition\\ (TMDL) folder or model.bim ({model}); Power BI Desktop cannot open the '
                        f'report: {p}', file=p, path='datasetReference.byPath.path')
        elif is_obj(by_conn) and isinstance(by_conn.get('connectionString'), str):
            cs = by_conn['connectionString']
            if not cs.strip():
                c.error('PBIR_DATASET_CONNECTION_STRING_EMPTY', f'datasetReference.byConnection.connectionString is empty: '
                        f'{p}', file=p, path='datasetReference.byConnection.connectionString')
                return
            keys = {seg.split('=', 1)[0].strip().lower() for seg in cs.split(';') if '=' in seg and seg.index('=') > 0}
            if 'semanticmodelid' in keys and 'data source' not in keys:
                c.warning('PBIR_DATASET_CONNECTION_API_ONLY_FORM', 'datasetReference.byConnection.connectionString is the '
                          f'Fabric REST API publish form and is not supported for local development: {p}', file=p,
                          path='datasetReference.byConnection.connectionString')

    def version(self, p, data):
        c = self.col
        if not is_obj(data):
            c.error('PBIR_JSON_ROOT_NOT_OBJECT', f'version.json root must be a JSON object: {p}', file=p)
            return
        if not isinstance(data.get('$schema'), str) or not data['$schema']:
            c.error('PBIR_JSON_FILE_NO_SCHEMA', f'version.json must include a "$schema" URL: {p}', file=p, path='$schema')
        if not isinstance(data.get('version'), str) or not data['version']:
            c.error('PBIR_VERSION_VALUE_MISSING', f'version.json must include a non-empty "version" string: {p}',
                    file=p, path='version')

    def resource_packages(self, p, report):
        for pkg in report.get('resourcePackages') or []:
            if not is_obj(pkg) or not isinstance(pkg.get('items'), list):
                continue
            seen, reported = set(), set()
            for item in pkg['items']:
                if not is_obj(item) or item.get('name') is None:
                    continue
                name = str(item['name'])
                if name in seen and name not in reported:
                    self.col.error('PBIR_RESOURCE_PACKAGE_DUPLICATE', f'Duplicate resource package item name "{name}" in '
                                   f'package "{pkg.get("name") or "(unnamed)"}": {p}', file=p, path=name)
                    reported.add(name)
                seen.add(name)
                if isinstance(item.get('path'), str) and pkg.get('name') in ('SharedResources', 'RegisteredResources'):
                    fp = os.path.join(self.report_dir, 'StaticResources', pkg['name'], *item['path'].split('/'))
                    if not os.path.isfile(fp):
                        self.col.error('BHY_RESOURCE_FILE_MISSING', f'Resource "{name}" in package "{pkg["name"]}" has no '
                                       f'file at {fp}: {p}', file=p, path=name)

    @staticmethod
    def custom_visuals(report):
        out = set()
        if not is_obj(report):
            return out
        out |= {g for g in report.get('publicCustomVisuals') or [] if isinstance(g, str) and g}
        for pkg in report.get('resourcePackages') or []:
            if is_obj(pkg) and isinstance(pkg.get('name'), str) and pkg.get('type') in ('CustomVisual', 'OrganizationalStoreCustomVisual'):
                out.add(pkg['name'])
        return out

    # ---------- pages ----------
    def pages_index(self, p, data, exists, ok):
        c = self.col
        if not exists:
            c.error('PBIR_PAGES_JSON_MISSING', f'Missing: {p}', file=p)
            return
        if not ok or data is None:
            return
        if not is_obj(data):
            c.error('PBIR_JSON_ROOT_NOT_OBJECT', f'pages.json root must be a JSON object: {p}', file=p)
            return
        order = data.get('pageOrder')
        if not isinstance(order, list):
            c.error('PBIR_PAGE_ORDER_NOT_ARRAY', f'pages.json "pageOrder" must be an array: {p}', file=p)
            return
        if not order:
            c.error('PBIR_PAGE_ORDER_EMPTY', f'Report has no pages; pages.json "pageOrder" is empty: {p}', file=p)
        if not isinstance(data.get('$schema'), str) or not data['$schema']:
            c.error('PBIR_JSON_FILE_NO_SCHEMA', f'pages.json has no "$schema" and cannot be schema-validated: {p}',
                    file=p, path='$schema')
        active = data.get('activePageName')
        if not isinstance(active, str) or not active:
            c.warning('PBIR_PAGE_ACTIVE_NAME_MISSING', f'pages.json activePageName is missing: {p}', file=p,
                      path='activePageName')
        elif active not in order:
            c.warning('PBIR_PAGE_ORDER_ACTIVE_NAME_MISSING', f'pages.json activePageName "{active}" is not present in '
                      f'pageOrder: {p}', file=p)

    def page_records(self, order):
        pages_dir = os.path.join(self.def_dir, 'pages')
        if not os.path.isdir(pages_dir):
            return []
        recs, seen = [], set()
        for i, n in enumerate(order):
            if n not in seen:
                seen.add(n)
                recs.append(self.page_record(pages_dir, n, i + 1))
        for n in sorted(os.listdir(pages_dir)):
            if os.path.isdir(os.path.join(pages_dir, n)) and n not in seen:
                seen.add(n)
                recs.append(self.page_record(pages_dir, n, None))
        return recs

    def page_record(self, pages_dir, name, ordinal):
        jp = os.path.join(pages_dir, name, 'page.json')
        rec = {'name': name, 'json': jp, 'ordinal': ordinal, 'has': os.path.isfile(jp), 'data': None, 'err': None}
        if rec['has']:
            try:
                with open(jp, encoding='utf-8') as f:
                    rec['data'] = json.load(f)
            except ValueError as e:
                rec['err'] = str(e)
        return rec

    def page(self, page, custom):
        c, jp = self.col, page['json']
        if page['ordinal'] is None:
            c.warning('PBIR_PAGE_DIR_NOT_IN_ORDER', f'Page directory exists but not in pageOrder: "{page["name"]}"',
                      file=jp, path=page['name'])
            return
        data = page['data']
        if page['has'] and not page['err'] and not is_obj(data):
            self.push(jp, data)
            c.error('PBIR_JSON_ROOT_NOT_OBJECT', f'page.json root must be a JSON object: {jp}', file=jp)
            return
        self.page_rules(page)
        if is_obj(data):
            self.push(jp, data, 'page')
            self.filters(jp, data)
        visuals = self.visual_records(page['name'])
        names_on_page, dirs_seen, json_names = set(), set(), set()
        for v in visuals:
            if v['dir'] in dirs_seen:
                c.error('PBIR_VISUAL_NAME_DUPLICATE_DIR', f'Duplicate visual name on page "{page["name"]}": "{v["dir"]}"',
                        file=v['json'], path=v['dir'])
            dirs_seen.add(v['dir'])
            if v['has'] and not v['err'] and not is_obj(v['data']):
                self.push(v['json'], v['data'])
                c.error('PBIR_JSON_ROOT_NOT_OBJECT', f'visual.json root must be a JSON object: {v["json"]}', file=v['json'])
                continue
            self.visual_identity(v)
            self.push_if_json(os.path.join(os.path.dirname(v['json']), 'mobile.json'), 'visualContainerMobileState')
            d = v['data']
            if not is_obj(d):
                continue
            self.push(v['json'], d, 'visualContainer')
            if isinstance(d.get('name'), str):
                if d['name'] in json_names:
                    c.warning('PBIR_VISUAL_NAME_DUPLICATE_JSON', f'Duplicate visual "name" value "{d["name"]}" on page '
                              f'"{page["name"]}": {v["json"]}', file=v['json'], path=d['name'])
                json_names.add(d['name'])
                names_on_page.add(d['name'])
            if d.get('position'):
                pd = data if is_obj(data) else {}
                self.layout(v, pd.get('width') or 1280, pd.get('height') or 720)
                self.dropdown_slicer(v)
            self.visual_type(v, custom)
            self.roles(v)
            self.formatting(v)
            self.textbox(v)
            self.filters(v['json'], d)
        if is_obj(data):
            self.interactions(page, names_on_page)
            self.drillthrough(page, visuals)

    def page_rules(self, page):
        c, jp, data = self.col, page['json'], page['data']
        if not page['has']:
            c.error('PBIR_PAGE_JSON_MISSING', f'Page in pageOrder but missing page.json: "{page["name"]}"', file=jp,
                    path=page['name'])
            return
        if page['err']:
            c.error('PBIR_JSON_PARSE_ERROR', f'Invalid JSON: {jp} — {page["err"]}', file=jp)
            return
        if not is_obj(data):
            return
        name = data.get('name')
        if name is None or name == '':
            c.warning('PBIR_PAGE_NAME_MISSING', f'Page name is missing or empty: {jp}', file=jp)
        elif not isinstance(name, str):
            c.error('PBIR_PAGE_NAME_NOT_STRING', f'Page name must be a string: {jp}', file=jp)
        else:
            if name != page['name']:
                c.warning('PBIR_PAGE_NAME_MISMATCH', f'page.json name "{name}" does not match page directory '
                          f'"{page["name"]}": {jp}', file=jp)
            if len(name) > self.meta['constants']['pageNameMax']:
                c.warning('PBIR_PAGE_NAME_TOO_LONG', f'Page name "{name}" is {len(name)} characters; maximum is '
                          f'{self.meta["constants"]["pageNameMax"]}: {jp}', file=jp)
        if not data.get('displayName'):
            c.warning('PBIR_PAGE_DISPLAY_NAME_MISSING', f'page.json missing "displayName": {jp}', file=jp)
        if data.get('visuals') is not None:
            c.error('PBIR_PAGE_VISUALS_IN_JSON', f'page.json must not contain "visuals": {jp}', file=jp)
        objs = data.get('objects')
        if is_obj(objs) and 'page' in objs:
            c.error('PBIR_PAGE_OBJECTS_HAS_PAGE', f'page.json objects must not contain "page": {jp}', file=jp)
        if data.get('displayOption') not in (None, 'Dynamic') and (data.get('width') is None or data.get('height') is None):
            c.warning('PBIR_PAGE_DISPLAY_OPTION_DIMENSIONS_MISSING', f'Page displayOption "{data["displayOption"]}" should '
                      f'include width and height: {jp}', file=jp)
        for obj, arr in (objs or {}).items() if is_obj(objs) else []:
            if obj not in VALID_PAGE_OBJECTS:
                c.error('PBIR_PAGE_OBJECT_UNKNOWN', f'Unknown page object "{obj}": {jp}', file=jp, path=obj)
                continue
            valid = VALID_PAGE_OBJECTS[obj]
            if not valid or not isinstance(arr, list):
                continue
            for entry in arr:
                props = entry.get('properties') if is_obj(entry) else None
                if not is_obj(props):
                    continue
                for pn in props:
                    if pn not in valid:
                        c.error('PBIR_PAGE_OBJECT_PROP_UNKNOWN', f'Invalid property "{pn}" in page object "{obj}" (valid: '
                                f'{", ".join(valid)}): {jp}', file=jp, path=f'{obj}.{pn}')
                if obj in ('background', 'outspace') and self.data_bound(props.get('color')):
                    c.error('PBIR_PAGE_COLOR_DATA_BOUND_UNSUPPORTED', f'Page object "{obj}.color" uses a data-bound '
                            f'expression, which Power BI Desktop does not evaluate for the page canvas: {jp}', file=jp,
                            path=f'{obj}.color')

    def data_bound(self, v):
        if isinstance(v, list):
            return any(self.data_bound(x) for x in v)
        if not is_obj(v):
            return False
        return any(k in DATA_BOUND or self.data_bound(x) for k, x in v.items())

    def visual_records(self, page_name):
        vdir = os.path.join(self.def_dir, 'pages', page_name, 'visuals')
        if not os.path.isdir(vdir):
            return []
        out = []
        for n in sorted(os.listdir(vdir)):
            fp = os.path.join(vdir, n)
            if os.path.isfile(fp) and n.lower().endswith('.json'):
                self.col.error('PBIR_VISUALS_JSON_MISPLACED', 'JSON files directly under a visuals directory are ignored; '
                               f'each visual must be in visuals/<visualName>/visual.json: {fp}', file=fp)
                continue
            if not os.path.isdir(fp):
                continue
            jp = os.path.join(fp, 'visual.json')
            rec = {'dir': n, 'json': jp, 'has': os.path.isfile(jp), 'data': None, 'err': None}
            if rec['has']:
                try:
                    with open(jp, encoding='utf-8') as f:
                        rec['data'] = json.load(f)
                except ValueError as e:
                    rec['err'] = str(e)
            out.append(rec)
        return out

    # ---------- visuals ----------
    def visual_identity(self, v):
        c, jp = self.col, v['json']
        if not v['has']:
            c.error('PBIR_VISUAL_DIR_WITHOUT_JSON', f'Visual directory without visual.json is ignored by the report host: '
                    f'{os.path.dirname(jp)}', file=jp)
            return
        if v['err']:
            c.error('PBIR_JSON_PARSE_ERROR', f'Invalid JSON: {jp} — {v["err"]}', file=jp)
            return
        d = v['data']
        if not is_obj(d):
            return
        name = d.get('name')
        if name is None or name == '':
            c.error('PBIR_VISUAL_NAME_MISSING', f'Visual name is missing or empty: {jp}', file=jp)
        elif not isinstance(name, str):
            c.error('PBIR_VISUAL_NAME_NOT_STRING', f'Visual name must be a string: {jp}', file=jp)
        elif len(name) > self.meta['constants']['visualNameMax']:
            c.warning('PBIR_VISUAL_NAME_TOO_LONG', f'Visual name "{name}" is {len(name)} characters; maximum is '
                      f'{self.meta["constants"]["visualNameMax"]}: {jp}', file=jp)
        if not d.get('position'):
            c.error('PBIR_VISUAL_POSITION_MISSING', f'visual.json missing "position": {jp}', file=jp)
        if 'visualContainerObjects' in d:
            c.error('PBIR_VISUAL_VCO_AT_ROOT', 'visualContainerObjects must be nested inside "visual" (sibling of '
                    f'"objects"), not at the root of visual.json: {jp}', file=jp)

    def layout(self, v, width, height):
        pos, jp, tol = v['data'].get('position'), v['json'], self.meta['constants']['layoutTolerance']
        if not is_obj(pos) or number(width) is None or number(height) is None:
            return
        x, y, w, h = (number(pos.get(k)) for k in ('x', 'y', 'width', 'height'))  # a non-number compares false in JS
        if x is not None and w is not None and x + w > width + tol:
            self.col.warning('PBIR_LAYOUT_OUT_OF_BOUNDS_WIDTH', f'Visual extends beyond page width (x:{x} + w:{w} = '
                             f'{x + w} > {width}): {jp}', file=jp)
        if y is not None and h is not None and y + h > height + tol:
            self.col.warning('PBIR_LAYOUT_OUT_OF_BOUNDS_HEIGHT', f'Visual extends beyond page height (y:{y} + h:{h} = '
                             f'{y + h} > {height}): {jp}', file=jp)
        if x is not None and x < 0:
            self.col.warning('PBIR_LAYOUT_NEGATIVE_X', f'Visual has negative x position ({x}): {jp}', file=jp)
        if y is not None and y < 0:
            self.col.warning('PBIR_LAYOUT_NEGATIVE_Y', f'Visual has negative y position ({y}): {jp}', file=jp)

    @staticmethod
    def lit_str(raw):
        v = (((raw or {}).get('expr') or {}).get('Literal') or {}).get('Value') if is_obj(raw) else None
        return v.replace("'", '') if isinstance(v, str) else None

    @staticmethod
    def lit_num(raw):
        v = (((raw or {}).get('expr') or {}).get('Literal') or {}).get('Value') if is_obj(raw) else None
        if v is None:
            return None
        m = re.match(r'^\s*([-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?)', re.sub(r'D$', '', str(v)))
        return float(m.group(1)) if m else None

    def dropdown_slicer(self, v):
        vis = v['data'].get('visual')
        pos = v['data'].get('position')
        if not is_obj(vis) or vis.get('visualType') != 'slicer' or not is_obj(pos) or number(pos.get('height')) is None:
            return
        objs = vis.get('objects') or {}

        def first(name, prop):
            arr = objs.get(name)
            return arr[0].get('properties', {}).get(prop) if isinstance(arr, list) and arr and is_obj(arr[0]) else None
        if self.lit_str(first('data', 'mode')) != 'Dropdown':
            return
        k = self.meta['constants']
        header = self.lit_str(first('header', 'show')) != 'false'
        top = self.lit_num(first('padding', 'top'))
        bottom = self.lit_num(first('padding', 'bottom'))
        top = k['defaultPadding'] if top is None else top
        bottom = k['defaultPadding'] if bottom is None else bottom
        sel = k['slicerSelectorFloor'] + top + bottom
        full = k['slicerHeaderFloor'] + sel
        h = pos['height']
        if h < sel:
            self.col.error('PBIR_SLICER_HEIGHT_BELOW_FLOOR', f'Dropdown slicer height {h}px < {fmt(sel)}px minimum; the '
                           f'selector will be clipped: {v["json"]}', file=v['json'])
        elif header and h < full:
            self.col.warning('PBIR_SLICER_HEADER_MAY_CLIP', f'Dropdown slicer height {h}px < {fmt(full)}px full height '
                             f'(header + selector + padding); the header may be clipped: {v["json"]}', file=v['json'])

    def textbox(self, v):
        vis = v['data'].get('visual')
        pos = v['data'].get('position')
        if not is_obj(vis) or vis.get('visualType') != 'textbox' or not is_obj(pos) or number(pos.get('height')) is None:
            return
        gen = (vis.get('objects') or {}).get('general')
        paras = gen[0].get('properties', {}).get('paragraphs') if isinstance(gen, list) and gen and is_obj(gen[0]) else None
        if not isinstance(paras, list):
            return
        max_pt = 0.0
        for para in paras:
            for run in (para or {}).get('textRuns') or []:
                raw = ((run or {}).get('textStyle') or {}).get('fontSize')
                try:
                    pt = float(raw) if isinstance(raw, (int, float)) else float(re.match(r'\s*([-+]?\d*\.?\d+)', raw).group(1)) \
                        if isinstance(raw, str) else None
                except (AttributeError, ValueError):
                    pt = None
                if pt is not None and pt > max_pt:
                    max_pt = pt
        if max_pt <= 0:
            return
        pad = (vis.get('visualContainerObjects') or {}).get('padding')
        props = pad[0].get('properties', {}) if isinstance(pad, list) and pad and is_obj(pad[0]) else {}
        k = self.meta['constants']
        top = self.lit_num(props.get('top'))
        bottom = self.lit_num(props.get('bottom'))
        top = k['defaultPadding'] if top is None else top
        bottom = k['defaultPadding'] if bottom is None else bottom
        need = max(18, math.ceil(max_pt * 25 / 16)) + top + bottom
        if pos['height'] < need:
            self.col.warning('PBIR_TEXTBOX_HEIGHT_BELOW_FLOOR', f'Textbox height {pos["height"]}px may be too small for '
                             f'{fmt(max_pt)}pt font (min {fmt(need)}px to avoid a scrollbar): {v["json"]}', file=v['json'])

    def visual_type(self, v, custom):
        vt = (v['data'].get('visual') or {}).get('visualType') if is_obj(v['data'].get('visual')) else None
        if not isinstance(vt, str) or not vt:
            return
        entry = self.meta['visuals'].get(vt)
        if not entry:
            if vt not in custom:
                self.col.warning('PBIR_VISUAL_TYPE_UNKNOWN', f'Unknown visualType "{vt}": {v["json"]}', file=v['json'],
                                 path='visual.visualType')
            return
        if entry.get('deprecated'):
            alt = f' Use "{entry["preferredAlternative"]}" instead.' if entry.get('preferredAlternative') else ''
            self.col.warning('PBIR_VISUAL_TYPE_DEPRECATED', f'Deprecated visualType "{vt}": '
                             f'{entry.get("deprecatedReason") or "Deprecated visual"}.{alt} — {v["json"]}',
                             file=v['json'], path='visual.visualType')

    def roles(self, v):
        vis = v['data'].get('visual')
        vt = vis.get('visualType') if is_obj(vis) else None
        entry = self.meta['visuals'].get(vt) if isinstance(vt, str) else None
        if not entry:
            return
        c, jp = self.col, v['json']
        qs = ((vis.get('query') or {}).get('queryState')) if is_obj(vis.get('query')) else None
        if not qs:
            if entry['requiredRoles']:
                c.error('PBIR_QUERY_STATE_MISSING', f'Visual has no queryState but visualType "{vt}" requires roles '
                        f'[{", ".join(entry["requiredRoles"])}]: {jp}', file=jp)
            return
        valid = entry['roles']
        for role in qs:
            if role not in valid:
                c.error('PBIR_ROLE_UNKNOWN', f'Unknown role "{role}" for visualType "{vt}": {jp}', file=jp, path=role)
        for req in entry['requiredRoles']:
            projs = (qs.get(req) or {}).get('projections') if is_obj(qs.get(req)) else None
            if not isinstance(projs, list) or not projs:
                c.error('PBIR_ROLE_REQUIRED_MISSING', f'Required role "{req}" missing or has no projections for visualType '
                        f'"{vt}": {jp}', file=jp, path=req)
        if vt == 'scatterChart' and not any(isinstance((qs.get(r) or {}).get('projections'), list) and qs[r]['projections']
                                            for r in ('Category', 'Series') if is_obj(qs.get(r))):
            c.warning('PBIR_SCATTER_NO_GROUPING', f'scatterChart has no grouping role (Category or Series); it will render '
                      f'a single aggregated point: {jp}', file=jp)
        for role, mx in entry['maxPerRole'].items():
            projs = (qs.get(role) or {}).get('projections') if is_obj(qs.get(role)) else None
            if isinstance(projs, list) and len(projs) > mx:
                c.error('PBIR_ROLE_MAX_EXCEEDED', f'Role "{role}" has {len(projs)} projections but max is {mx} for '
                        f'visualType "{vt}": {jp}', file=jp, path=role)
        for role, state in qs.items():
            if role not in valid or not is_obj(state) or not isinstance(state.get('projections'), list):
                continue
            for p in state['projections']:
                if not is_obj(p):
                    continue
                if not p.get('nativeQueryRef'):
                    c.error('PBIR_PROJECTION_MISSING_NATIVE_QUERY_REF', f'Projection in role "{role}" missing nativeQueryRef '
                            f'for visualType "{vt}": {jp}', file=jp, path=role)
                f = p.get('field')
                if is_obj(f):
                    kind = valid[role]
                    if kind == 'Grouping' and f.get('Measure'):
                        c.error('PBIR_ROLE_KIND_MISMATCH', f'Measure expression in Grouping-only role "{role}" for '
                                f'visualType "{vt}": {jp}', file=jp, path=role)
                    elif kind == 'Measure' and f.get('Column') and not f.get('Aggregation'):
                        c.error('PBIR_ROLE_KIND_MISMATCH', f'Column expression in Measure-only role "{role}" (use Measure '
                                f'or Aggregation) for visualType "{vt}": {jp}', file=jp, path=role)

    def formatting(self, v):
        vis = v['data'].get('visual')
        if not is_obj(vis):
            return
        vt = vis.get('visualType')
        if not isinstance(vt, str) or not vt:
            return
        entry = self.meta['visuals'].get(vt) or {}
        cap = self.meta['capabilities'].get(entry.get('capabilityKey') or vt)
        vcos = self.meta['vcos']
        jp = v['json']
        objs = vis.get('objects')
        if is_obj(objs):
            for name, arr in objs.items():
                odef = cap.get(name) if cap else None
                if cap is not None and odef is None:
                    hint = f' — did you mean to put "{name}" under visualContainerObjects?' if name in vcos else ''
                    self.col.error('PBIR_FORMATTING_OBJECT_UNKNOWN', f'Unknown formatting object "{name}" for visualType '
                                   f'"{vt}"{hint}: {jp}', file=jp, path=name)
                for item in arr if isinstance(arr, list) else []:
                    if not is_obj(item) or not is_obj(item.get('properties')):
                        continue
                    self.properties(item['properties'], odef, name, vt, jp, 'PBIR_FORMATTING_PROP_UNKNOWN')
                    self.fill_rule_stops(item['properties'], jp)
                    if name == 'columnFormatting' and self.has_gradient(item['properties'], 0):
                        self.col.error('PBIR_CF_GRADIENT_IN_COLUMN_FORMATTING', 'A gradient FillRule is stored under '
                                       f'"columnFormatting"; conditional gradients belong in "values": {jp}', file=jp,
                                       path='columnFormatting')
                    if name == 'columnWidth' and odef is not None:
                        expr = (item['properties'].get('value') or {}).get('expr') if is_obj(item['properties'].get('value')) else None
                        if is_obj(expr) and expr and 'Literal' not in expr:
                            self.col.error('PBIR_COLUMN_WIDTH_DATA_BOUND_UNSUPPORTED', 'Column width is not conditionally '
                                           f'formattable: "columnWidth.value" uses a "{next(iter(expr))}" expression: {jp}',
                                           file=jp, path='columnWidth.value')
        vco = vis.get('visualContainerObjects')
        if is_obj(vco):
            names = set(cap or {})
            for name, arr in vco.items():
                vdef = vcos.get(name)
                if vdef is None:
                    hint = f' — did you mean to put "{name}" under objects (it\'s a known formatting object for "{vt}")?' \
                        if name in names else ''
                    self.col.error('PBIR_VCO_UNKNOWN', f'Unknown visualContainerObject "{name}"{hint}: {jp}', file=jp,
                                   path=name)
                for item in arr if isinstance(arr, list) else []:
                    if not is_obj(item) or not is_obj(item.get('properties')):
                        continue
                    self.properties(item['properties'], vdef, name, 'VCO', jp, 'PBIR_VCO_PROP_UNKNOWN')
                    self.fill_rule_stops(item['properties'], jp)

    def properties(self, props, odef, obj, vt, jp, unknown_code):
        valid = set(odef) if is_obj(odef) else None
        for pn, pv in props.items():
            if valid is not None and pn not in valid:
                self.col.error(unknown_code, f'Unknown property "{pn}" in formatting object "{obj}" for {vt}: {jp}',
                               file=jp, path=f'{obj}.{pn}')
                continue
            if is_obj(pv):
                self.nested(pv, pn, valid, obj, jp, 0)
            pdef = odef.get(pn) if is_obj(odef) else None
            if is_obj(pdef) and pdef.get('kind') == 'enum' and pdef.get('enumValues') and pv is not None:
                raw = self.unwrap(pv)
                if raw is not None and raw not in pdef['enumValues']:
                    self.col.error('PBIR_FORMATTING_ENUM_INVALID', f'Invalid enum value "{raw}" for "{obj}.{pn}" (valid: '
                                   f'{", ".join(pdef["enumValues"])}): {jp}', file=jp, path=f'{obj}.{pn}')

    def nested(self, value, parent, valid, obj, jp, depth):
        if valid is None or depth > 3:
            return
        for k, child in value.items():
            if k in valid and k != parent:
                self.col.error('PBIR_FORMATTING_PROP_NESTED', f'Property "{k}" appears nested inside "{parent}"; place it '
                               f'as a sibling under "{obj}.properties": {jp}', file=jp, path=f'{obj}.{parent}')
            if is_obj(child):
                self.nested(child, parent, valid, obj, jp, depth + 1)

    @staticmethod
    def unwrap(v):
        lit = ((v.get('expr') or {}).get('Literal') if is_obj(v.get('expr')) else None) if is_obj(v) else None
        raw = lit.get('Value') if is_obj(lit) else None
        if not isinstance(raw, str):
            return None
        clean = re.sub(r"^'|'$", '', raw)
        m = re.match(r'^(-?\d+(?:\.\d+)?)[DLdl]$', clean)
        return m.group(1) if m else clean

    def fill_rule_stops(self, props, jp):
        cands = []
        for pn, pv in props.items():
            solid = (((pv.get('solid') or {}).get('color') or {}).get('expr') or {}) if is_obj(pv) and is_obj(pv.get('solid')) \
                and is_obj(pv['solid'].get('color')) and is_obj(pv['solid']['color'].get('expr')) else {}
            from_solid = (solid.get('FillRule') or {}).get('FillRule') if is_obj(solid.get('FillRule')) else None
            if from_solid:
                cands.append((pn, from_solid))
            expr = (pv.get('expr') or {}).get('FillRule') if is_obj(pv) and is_obj(pv.get('expr')) else None
            from_expr = (expr.get('FillRule') or expr) if is_obj(expr) else None
            if from_expr and from_expr is not from_solid:
                cands.append((pn, from_expr))
        if is_obj(props.get('fillRule')):
            cands.append(('fillRule', props['fillRule']))
        for pn, fr in cands:
            if not is_obj(fr):
                continue
            grad = fr.get('linearGradient2') or fr.get('linearGradient3')
            if not is_obj(grad):
                continue
            for stop in (('min', 'max') if fr.get('linearGradient2') else ('min', 'mid', 'max')):
                color = (grad.get(stop) or {}).get('color') if is_obj(grad.get(stop)) else None
                if is_obj(color) and 'expr' in color:
                    self.col.error('PBIR_FILLRULE_STOP_DOUBLE_WRAP', f'FillRule gradient "{stop}" color stop has an "expr" '
                                   f'wrapper; use {{"Literal": {{"Value": "\'#hex\'"}}}}: {jp}', file=jp, path=f'{pn}.{stop}')
                    continue
                ok = is_obj(color) and list(color) == ['Literal'] and is_obj(color['Literal']) and \
                    list(color['Literal']) == ['Value'] and isinstance(color['Literal']['Value'], str) and \
                    re.fullmatch(r"'(#[0-9a-fA-F]{3}|#[0-9a-fA-F]{6})'", color['Literal']['Value'])
                if not ok:
                    self.col.error('PBIR_FILLRULE_STOP_COLOR_INVALID', f'FillRule gradient "{stop}" color stop must be a '
                                   f'single-quoted Literal #RGB or #RRGGBB: {jp}', file=jp, path=f'{pn}.{stop}')
            ns = grad.get('nullColoringStrategy')
            if is_obj(ns):
                strat = ((ns.get('strategy') or {}).get('Literal') or {}).get('Value') if is_obj(ns.get('strategy')) else None
                if isinstance(strat, str) and strat.strip("'") == 'specificColor' and 'color' not in ns:
                    self.col.error('PBIR_FILLRULE_NULL_STRATEGY_COLOR_MISSING', 'FillRule nullColoringStrategy is '
                                   f'"specificColor" but has no "color": {jp}', file=jp, path=f'{pn}.nullColoringStrategy')

    def has_gradient(self, node, depth):
        if depth > 8:
            return False
        if isinstance(node, list):
            return any(self.has_gradient(x, depth + 1) for x in node)
        if not is_obj(node):
            return False
        return 'linearGradient2' in node or 'linearGradient3' in node or any(self.has_gradient(x, depth + 1) for x in node.values())

    # ---------- page-level relations ----------
    def filters(self, path, data):
        fc = data.get('filterConfig') if is_obj(data) else None
        flts = fc.get('filters') if is_obj(fc) else None
        if not isinstance(flts, list):
            return
        local = set()
        for f in flts:
            name = f.get('name') if is_obj(f) else None
            if name is None or name == '':
                self.col.warning('PBIR_FILTER_NAME_MISSING', f'Filter name is missing or empty: {path}', file=path)
            elif not isinstance(name, str):
                self.col.error('PBIR_FILTER_NAME_NOT_STRING', f'Filter name must be a string: {path}', file=path)
            else:
                if name in local:
                    self.col.warning('PBIR_FILTER_NAME_DUPLICATE_LOCAL', f'Duplicate filter name "{name}" within same '
                                     f'scope: {path}', file=path, path=name)
                local.add(name)
                if name in self.filter_names:
                    self.col.warning('PBIR_FILTER_NAME_DUPLICATE_GLOBAL', f'Duplicate filter name "{name}" across report '
                                     f'(also in another scope): {path}', file=path, path=name)
                self.filter_names.add(name)
            body = f.get('filter') if is_obj(f) else None
            where = body.get('Where') if is_obj(body) else None
            if isinstance(where, list):
                aliases = {x.get('Name') for x in body.get('From') or [] if is_obj(x) and isinstance(x.get('Name'), str)}
                for clause in where:
                    self.source_refs(clause, aliases, path)

    def source_refs(self, node, aliases, path):
        if isinstance(node, list):
            for x in node:
                self.source_refs(x, aliases, path)
            return
        if not is_obj(node):
            return
        ref = node.get('SourceRef')
        if is_obj(ref):
            if isinstance(ref.get('Entity'), str) and not ref.get('Source'):
                self.col.warning('PBIR_FILTER_ENTITY_IN_WHERE', 'Filter Where uses "Entity" instead of "Source" in '
                                 f'SourceRef (should reference From alias): {path}', file=path)
            elif isinstance(ref.get('Source'), str) and aliases and ref['Source'] not in aliases:
                self.col.warning('PBIR_FILTER_SOURCE_UNKNOWN', f'Filter Where SourceRef.Source "{ref["Source"]}" not found '
                                 f'in From aliases [{", ".join(sorted(aliases))}]: {path}', file=path)
            return
        for x in node.values():
            self.source_refs(x, aliases, path)

    def interactions(self, page, names):
        d, jp = page['data'], page['json']
        rels = [r for k in ('relationships', 'visualInteractions') for r in (d.get(k) or []) if isinstance(d.get(k), list)]
        reported = set()
        for r in rels:
            if not is_obj(r):
                continue
            for end in (r.get('source', r.get('sourceVisual')), r.get('target', r.get('targetVisual'))):
                n = end if isinstance(end, str) else next((end.get(k) for k in ('name', 'visualName', 'visualContainerName',
                                                                                'targetVisualName') if is_obj(end) and end.get(k)), None)
                if n and n not in names and n not in reported:
                    self.col.warning('PBIR_VISUAL_INTERACTION_REF_MISSING', f'Visual interaction references missing visual '
                                     f'"{n}": {jp}', file=jp, path=n)
                    reported.add(n)

    def drillthrough(self, page, visuals):
        d, jp = page['data'], page['json']
        disp = d.get('displayName') if isinstance(d.get('displayName'), str) else page['name']
        fc = d.get('filterConfig')
        flts = [f for f in (fc.get('filters') or [] if is_obj(fc) else []) if is_obj(f)]
        dts = [f for f in flts if f.get('howCreated') == 'Drillthrough']
        pb = d.get('pageBinding') if is_obj(d.get('pageBinding')) else None
        c = self.col
        if (pb or {}).get('type') != 'Drillthrough':
            if d.get('type') == 'Drillthrough' or dts:
                c.warning('PBIR_DRILLTHROUGH_PAGE_BINDING_INVALID', f'Drillthrough page "{disp}" must have pageBinding.type '
                          f'"Drillthrough": {jp}', file=jp, path='pageBinding.type')
            return
        params = [p for p in (pb.get('parameters') or []) if is_obj(p)] if isinstance(pb.get('parameters'), list) else []
        if not dts:
            c.warning('PBIR_DRILLTHROUGH_FILTER_MISSING', f'Drillthrough page "{disp}" has no page filter marked howCreated '
                      f'"Drillthrough": {jp}', file=jp, path='filterConfig.filters')
        if not params:
            c.warning('PBIR_DRILLTHROUGH_PARAMETER_MISSING', f'Drillthrough page "{disp}" has no pageBinding parameters: {jp}',
                      file=jp, path='pageBinding.parameters')
        by_name = {f['name']: f for f in flts if isinstance(f.get('name'), str)}
        for i, p in enumerate(params):
            path = f'pageBinding.parameters[{i}]'
            bf = p.get('boundFilter')
            if not isinstance(bf, str) or not bf:
                c.warning('PBIR_DRILLTHROUGH_BOUND_FILTER_MISSING', f'Drillthrough parameter {i + 1} on page "{disp}" has no '
                          f'boundFilter: {jp}', file=jp, path=f'{path}.boundFilter')
                continue
            f = by_name.get(bf)
            if f is None:
                c.error('PBIR_DRILLTHROUGH_BOUND_FILTER_REF_MISSING', f'Drillthrough parameter {i + 1} references missing '
                        f'page filter "{bf}": {jp}', file=jp, path=f'{path}.boundFilter')
                continue
            if f.get('howCreated') != 'Drillthrough':
                c.warning('PBIR_DRILLTHROUGH_BOUND_FILTER_NOT_DRILLTHROUGH', f'Drillthrough parameter {i + 1} references '
                          f'filter "{bf}", which is not marked howCreated "Drillthrough": {jp}', file=jp)
            if f.get('field') is None:
                c.warning('PBIR_DRILLTHROUGH_FILTER_FIELD_MISSING', f'Drillthrough filter "{bf}" has no field: {jp}',
                          file=jp)
            if p.get('fieldExpr') is None:
                c.warning('PBIR_DRILLTHROUGH_FIELD_EXPR_MISSING', f'Drillthrough parameter {i + 1} bound to "{bf}" has no '
                          f'fieldExpr: {jp}', file=jp, path=f'{path}.fieldExpr')
            elif f.get('field') is not None and not json_eq(p['fieldExpr'], f['field']):
                c.error('PBIR_DRILLTHROUGH_FIELD_MISMATCH', f'Drillthrough parameter {i + 1} fieldExpr does not match bound '
                        f'filter "{bf}": {jp}', file=jp, path=f'{path}.fieldExpr')
        bound = {p.get('boundFilter') for p in params if isinstance(p.get('boundFilter'), str)}
        for f in dts:
            if isinstance(f.get('name'), str) and f['name'] not in bound:
                c.warning('PBIR_DRILLTHROUGH_FILTER_UNBOUND', f'Drillthrough filter "{f["name"]}" is not referenced by any '
                          f'pageBinding parameter: {jp}', file=jp)
        if not any(self.back_button(v) for v in visuals):
            c.warning('PBIR_DRILLTHROUGH_BACK_BUTTON_MISSING', f'Drillthrough page "{disp}" has no visible actionButton with '
                      f'visualLink.type "Back": {jp}', file=jp, path='visuals')

    @staticmethod
    def back_button(v):
        d = v.get('data')
        if not is_obj(d) or d.get('isHidden') is True or not is_obj(d.get('visual')) or d['visual'].get('visualType') != 'actionButton':
            return False
        links = (d['visual'].get('visualContainerObjects') or {}).get('visualLink')
        for link in links if isinstance(links, list) else []:
            props = link.get('properties') if is_obj(link) else None
            if not is_obj(props):
                continue

            def lv(x):
                return (((x or {}).get('expr') or {}).get('Literal') or {}).get('Value') if is_obj(x) else None
            if lv(props.get('show')) == 'true' and lv(props.get('type')) == "'Back'":
                return True
        return False

    # ---------- schema pass ----------
    def schema_pass(self):
        sv = SchemaValidator(SchemaStore())
        missing = {}
        for path, data in self.targets:
            url = data.get('$schema') if is_obj(data) else None
            if not isinstance(url, str):
                continue
            errs = sv.validate(data, url)
            if errs is None:
                missing.setdefault(url, []).append(path)
                continue
            for ipath, msg, extra in errs:
                more = f' (property: "{extra}")' if extra else ''
                self.col.error('PBIR_SCHEMA_VALIDATION_ERROR', f'Schema validation: {ipath} {msg}{more}: {path}',
                               file=path, path=ipath)
        for url, paths in missing.items():
            self.col.warning('PBIR_SCHEMA_UNREACHABLE', f'JSON Schema "{url}" is not bundled with this toolkit (run '
                             f'tools\\fetch_schemas.py on a PC with internet to add it); schema validation skipped for: '
                             f'{", ".join(paths)}', file=paths[0])

    # ---------- whole project ----------
    def project_checks(self):
        root = os.path.dirname(self.report_dir)
        for dp, _, fs in os.walk(root):
            for f in fs:
                p = os.path.join(dp, f)
                full = os.path.realpath(p)  # long form of any 8.3 short name: the path Desktop opens
                if len(full) > MAX_PATH:
                    self.col.error('BHY_PATH_TOO_LONG', f'Path is {len(full)} characters; Power BI Desktop '
                                   f'cannot open files past the Windows {MAX_PATH + 1}-character limit: {p}', file=p)


def fmt(x):
    return str(int(x)) if float(x).is_integer() else str(x)


def validate_report(path, pbip_path=None, schemas=True):
    """Validate a project. path: a .pbip, a .Report folder, or a folder holding one project. Returns the CLI-shaped dict."""
    report_dir, found_pbip = locate(path)
    return Validator(report_dir, pbip_path or found_pbip, schemas=schemas).run()


def render_text(res):
    lines = []
    for code, g in (res.get('diagnostics') or {}).items():
        for it in g['items']:
            lines.append(f'{g["severity"].upper():7} {code}  {it["message"]}')
    lines.append(f'{res["errorCount"]} error(s), {res["warningCount"]} warning(s); result={res["result"]}')
    return '\n'.join(lines)


def main():
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('path', help='.pbip file, .Report folder, or a folder holding one project')
    ap.add_argument('--format', choices=['json', 'text'], default='json')
    ap.add_argument('--no-schema', action='store_true', help='skip JSON Schema validation')
    a = ap.parse_args()
    res = validate_report(a.path, schemas=not a.no_schema)
    print(json.dumps({'data': res}, indent=1, ensure_ascii=False) if a.format == 'json' else render_text(res))
    sys.exit(1 if res['result'] == 'failed' else 0)


if __name__ == '__main__':
    main()
