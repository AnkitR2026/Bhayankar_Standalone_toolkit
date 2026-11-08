"""Tests for what makes this the Standalone Toolkit: the offline scaffold, the offline PBIR validator with bundled
schemas, a build that never needs Node.js or the internet, the shared folder settings, and workbooks without dashboards.

Run (from the toolkit folder):  python -m unittest discover -s tests -v
"""
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
TOOL = os.path.dirname(HERE)
sys.path.insert(0, TOOL)
import bhayankar_build as bb  # noqa: E402
import bhayankar_extractor as bx  # noqa: E402
import bhayankar_scaffold as bs  # noqa: E402
import bhayankar_toolkit as bt  # noqa: E402
import bhayankar_validate as bv  # noqa: E402

FIX = os.path.join(HERE, 'fixtures')
CLI_REF = os.path.join(FIX, 'cli_scaffold_0.4.0')  # `powerbi-report-author scaffold . --name Ref --offline` (CLI 0.4.0)
THEME_SHA256 = '4e5801de9e975831c9fc97fd366db68f221e9d9a6db260870216bd039159a960'  # CY26SU10.json written by CLI 0.4.0
GUID = re.compile(r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$')


def load(path):
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def save(path, data):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


def build_fixture(root, name='mini', files=('mini.twb', 'orders_source.tds'), edit_spec=None):
    """Extract and build a fixture workbook under root\\input and root\\output; returns (Build, validation result)."""
    inp, out = os.path.join(root, 'input'), os.path.join(root, 'output')
    os.makedirs(inp, exist_ok=True)
    for fn in files:
        shutil.copy(os.path.join(FIX, fn), inp)
    bx.run_folder(inp, out, table_map={})
    if edit_spec:
        p = os.path.join(out, name, 'bhayankar_rebuild.json')
        spec = load(p)
        edit_spec(spec)
        save(p, spec)
    old = bb.INPUT_DIR, bb.OUTPUT_DIR
    bb.INPUT_DIR, bb.OUTPUT_DIR = inp, out
    try:
        b = bb.Build(name)
        return b, b.run()
    finally:
        bb.INPUT_DIR, bb.OUTPUT_DIR = old


def codes(res):
    return set((res.get('diagnostics') or {}))


class Scaffold(unittest.TestCase):
    """bhayankar_scaffold writes what `powerbi-report-author scaffold --offline` writes, without Node.js."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix='bhy_scaffold_')

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_same_files_as_the_cli_scaffold(self):
        bs.scaffold(self.tmp, 'Ref', '../Ref.SemanticModel')
        for rel in ('Ref.pbip', 'Ref.Report/definition.pbir', 'Ref.Report/definition/version.json',
                    'Ref.Report/definition/report.json'):
            self.assertEqual(load(os.path.join(self.tmp, rel)), load(os.path.join(CLI_REF, rel)), rel)
        ours, cli = load(os.path.join(self.tmp, 'Ref.Report', '.platform')), load(os.path.join(CLI_REF, 'Ref.Report', '.platform'))
        ours['config'].pop('logicalId')
        cli['config'].pop('logicalId')
        self.assertEqual(ours, cli)  # only the (random in the CLI) id differs
        pages = load(os.path.join(self.tmp, 'Ref.Report', 'definition', 'pages', 'pages.json'))
        self.assertEqual(pages['$schema'], load(os.path.join(CLI_REF, 'Ref.Report', 'definition', 'pages', 'pages.json'))['$schema'])
        self.assertEqual(pages['pageOrder'], [])  # no starter page: the build writes every page
        theme = os.path.join(self.tmp, 'Ref.Report', 'StaticResources', 'SharedResources', 'BaseThemes', 'CY26SU10.json')
        with open(theme, 'rb') as f:
            self.assertEqual(hashlib.sha256(f.read()).hexdigest(), THEME_SHA256)

    def test_logical_id_is_stable_and_a_guid(self):
        a, b = os.path.join(self.tmp, 'a'), os.path.join(self.tmp, 'b')
        bs.scaffold(a, 'Sales')
        bs.scaffold(b, 'Sales')
        ida = load(os.path.join(a, 'Sales.Report', '.platform'))['config']['logicalId']
        self.assertEqual(ida, load(os.path.join(b, 'Sales.Report', '.platform'))['config']['logicalId'])
        self.assertRegex(ida, GUID)
        self.assertNotEqual(ida, bs.logical_id('Report', 'Finance'))
        self.assertNotEqual(ida, load(os.path.join(a, 'Sales.SemanticModel', '.platform'))['config']['logicalId'])

    def test_short_folders_keep_the_workbook_name_on_the_pbip(self):
        bs.scaffold(self.tmp, 'Design Review', pbip_name='Design Review 1X1 Dashboard')
        pbip = load(os.path.join(self.tmp, 'Design Review 1X1 Dashboard.pbip'))
        self.assertEqual(pbip['artifacts'][0]['report']['path'], 'Design Review.Report')
        model = load(os.path.join(self.tmp, 'Design Review.SemanticModel', '.platform'))
        self.assertEqual(model['metadata'], {'type': 'SemanticModel', 'displayName': 'Design Review'})


class OfflineValidator(unittest.TestCase):
    """bhayankar_validate runs the CLI's PBIR rules offline; each mutation below was checked against the CLI 0.4.0."""

    @classmethod
    def setUpClass(cls):
        cls.root = tempfile.mkdtemp(prefix='bhy_val_')
        cls.b, cls.res = build_fixture(cls.root)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.root, ignore_errors=True)

    def copy(self):
        dst = tempfile.mkdtemp(prefix='bhy_mut_')
        self.addCleanup(shutil.rmtree, dst, True)
        proj = os.path.join(dst, 'p')
        shutil.copytree(self.b.pbip_dir, proj)
        return proj, os.path.join(proj, 'mini.pbip'), os.path.join(proj, 'mini.Report')

    @staticmethod
    def visuals(rep):
        out = []
        for dp, _, fs in os.walk(os.path.join(rep, 'definition', 'pages')):
            out += [os.path.join(dp, f) for f in fs if f == 'visual.json']
        return sorted(out)

    def test_build_validates_clean_offline(self):
        self.assertEqual((self.res['result'], self.res['errorCount'], self.res['warningCount']), ('succeeded', 0, 0))
        self.assertIn('offline', self.res['validator'])

    def mutate(self, fn):
        proj, pbip, rep = self.copy()
        fn(rep, proj)
        return bv.validate_report(pbip)

    def test_unknown_formatting_object_and_property(self):
        meta = bv.metadata()

        def m(rep, _):
            p = self.visuals(rep)[0]
            d = load(p)
            vt = d['visual']['visualType']
            obj = sorted(meta['capabilities'][meta['visuals'][vt]['capabilityKey']])[0]
            d['visual'].setdefault('objects', {})['notAnObject'] = [{'properties': {}}]
            d['visual']['objects'][obj] = [{'properties': {'notAProperty': {'expr': {'Literal': {'Value': 'true'}}}}}]
            save(p, d)
        res = self.mutate(m)
        self.assertTrue({'PBIR_FORMATTING_OBJECT_UNKNOWN', 'PBIR_FORMATTING_PROP_UNKNOWN'} <= codes(res), codes(res))
        self.assertEqual(res['result'], 'failed')

    def test_pivot_table_has_no_totals_property(self):
        """The Security Awareness build error: total.totals is a tableEx property, not a pivotTable one."""
        meta = bv.metadata()
        self.assertIn('totals', meta['capabilities'][meta['visuals']['tableEx']['capabilityKey']]['total'])
        self.assertNotIn('totals', meta['capabilities'][meta['visuals']['pivotTable']['capabilityKey']]['total'])

    def test_missing_position(self):
        def m(rep, _):
            p = self.visuals(rep)[0]
            d = load(p)
            d.pop('position')
            save(p, d)
        res = self.mutate(m)
        self.assertTrue({'PBIR_VISUAL_POSITION_MISSING', 'PBIR_SCHEMA_VALIDATION_ERROR'} <= codes(res), codes(res))

    def test_visual_below_the_page_is_a_warning(self):
        def m(rep, _):
            p = self.visuals(rep)[0]
            d = load(p)
            d['position']['y'] = 5000
            save(p, d)
        res = self.mutate(m)
        self.assertEqual(codes(res), {'PBIR_LAYOUT_OUT_OF_BOUNDS_HEIGHT'})
        self.assertEqual(res['result'], 'succeededWithWarnings')

    def test_unknown_visual_type(self):
        def m(rep, _):
            p = self.visuals(rep)[0]
            d = load(p)
            d['visual']['visualType'] = 'fooChart'
            save(p, d)
        self.assertIn('PBIR_VISUAL_TYPE_UNKNOWN', codes(self.mutate(m)))

    def test_platform_logical_id_must_be_a_guid(self):
        def m(rep, _):
            p = os.path.join(rep, '.platform')
            d = load(p)
            d['config']['logicalId'] = 'abc'
            save(p, d)
        self.assertIn('PBIR_PLATFORM_LOGICAL_ID_INVALID', codes(self.mutate(m)))

    def test_invalid_json(self):
        def m(rep, _):
            with open(self.visuals(rep)[0], 'w', encoding='utf-8') as f:
                f.write('{"name": ')
        self.assertIn('PBIR_JSON_PARSE_ERROR', codes(self.mutate(m)))

    def test_schema_version_not_bundled_is_a_warning_not_a_download(self):
        def m(rep, _):
            p = os.path.join(rep, 'definition', 'version.json')
            d = load(p)
            d['$schema'] = d['$schema'].replace('/1.0.0/', '/9.9.9/')
            save(p, d)
        res = self.mutate(m)
        self.assertIn('PBIR_SCHEMA_UNREACHABLE', codes(res))
        self.assertEqual(res['diagnostics']['PBIR_SCHEMA_UNREACHABLE']['severity'], 'warning')

    def test_missing_theme_file(self):
        def m(rep, _):
            os.remove(os.path.join(rep, 'StaticResources', 'SharedResources', 'BaseThemes', 'CY26SU10.json'))
        self.assertIn('BHY_RESOURCE_FILE_MISSING', codes(self.mutate(m)))

    def test_report_bound_to_a_missing_model(self):
        def m(_, proj):
            os.rename(os.path.join(proj, 'mini.SemanticModel'), os.path.join(proj, 'gone.SemanticModel'))
        self.assertIn('BHY_MODEL_PATH_UNRESOLVED', codes(self.mutate(m)))

    def test_long_paths_are_errors(self):
        old = bv.MAX_PATH
        bv.MAX_PATH = 40
        try:
            res = bv.validate_report(os.path.join(self.b.pbip_dir, 'mini.pbip'))
        finally:
            bv.MAX_PATH = old
        self.assertIn('BHY_PATH_TOO_LONG', codes(res))

    def test_result_has_the_cli_shape(self):
        def m(rep, _):
            p = self.visuals(rep)[0]
            d = load(p)
            d['visual']['visualType'] = 'fooChart'
            save(p, d)
        res = self.mutate(m)
        for k in ('result', 'errorCount', 'warningCount', 'reportPath', 'diagnostics'):
            self.assertIn(k, res)
        g = res['diagnostics']['PBIR_VISUAL_TYPE_UNKNOWN']
        self.assertEqual(g['severity'], 'warning')
        self.assertTrue(g['items'][0]['message'] and g['items'][0]['file'].endswith('visual.json'))

    def test_command_line(self):
        pbip = os.path.join(self.b.pbip_dir, 'mini.pbip')
        p = subprocess.run([sys.executable, os.path.join(TOOL, 'bhayankar_validate.py'), pbip], capture_output=True,
                           text=True, encoding='utf-8', timeout=120)
        self.assertEqual(p.returncode, 0, p.stderr)
        self.assertEqual(json.loads(p.stdout)['data']['result'], 'succeeded')
        proj, pbip2, rep = self.copy()
        with open(self.visuals(rep)[0], 'w', encoding='utf-8') as f:
            f.write('not json')
        p = subprocess.run([sys.executable, os.path.join(TOOL, 'bhayankar_validate.py'), pbip2, '--format', 'text'],
                           capture_output=True, text=True, encoding='utf-8', timeout=120)
        self.assertEqual(p.returncode, 1)
        self.assertIn('PBIR_JSON_PARSE_ERROR', p.stdout)


class BundledSchemas(unittest.TestCase):
    """The schemas in resources\\schemas are complete: no reference leaves the bundle, no keyword is unsupported."""

    SUPPORTED = {'$schema', '$id', '$ref', 'title', 'description', 'default', 'definitions', 'type', 'properties',
                 'required', 'additionalProperties', 'patternProperties', 'items', 'additionalItems', 'const', 'enum',
                 'anyOf', 'oneOf', 'allOf', 'not', 'if', 'then', 'else', 'pattern', 'minLength', 'maxLength',
                 'minimum', 'maximum', 'exclusiveMinimum', 'exclusiveMaximum', 'minItems', 'maxItems', 'uniqueItems',
                 'examples', 'deprecated', 'markdownDescription', 'format', '$comment'}

    def files(self):
        out = []
        for dp, _, fs in os.walk(bv.SCHEMA_DIR):
            for f in fs:
                p = os.path.join(dp, f)
                url = bv.FABRIC + '/' + os.path.relpath(p, bv.SCHEMA_DIR).replace(os.sep, '/')
                out.append((url, load(p)))
        return out

    def test_every_reference_resolves_inside_the_bundle(self):
        store = bv.SchemaStore()
        found = 0

        def walk(n, base):
            nonlocal found
            if isinstance(n, dict):
                for k, v in n.items():
                    if k == '$ref' and isinstance(v, str):
                        found += 1
                        self.assertIsNotNone(store.resolve(base, v)[1], f'{base}: {v}')
                    else:
                        walk(v, base)
            elif isinstance(n, list):
                for x in n:
                    walk(x, base)
        files = self.files()
        self.assertGreaterEqual(len(files), 13)
        for url, doc in files:
            walk(doc, url)
        self.assertGreater(found, 200)

    def test_only_supported_keywords(self):
        used = set()

        def walk(n, parent=None):
            if isinstance(n, dict):
                for k, v in n.items():
                    if parent not in ('properties', 'definitions', 'patternProperties'):
                        used.add(k)
                    walk(v, k)
            elif isinstance(n, list):
                for x in n:
                    walk(x, parent)
        for _, doc in self.files():
            walk(doc)
        self.assertEqual(used - self.SUPPORTED, set())

    def test_every_schema_the_build_writes_is_bundled(self):
        store = bv.SchemaStore()
        urls = [bb.VISUAL_SCHEMA, bb.PAGE_SCHEMA, *bs.SCHEMA.values(),
                'https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json']
        for u in urls:
            self.assertIsNotNone(store.load(u), u)

    def test_visual_metadata_covers_the_catalog(self):
        cat = load(os.path.join(TOOL, 'bhayankar_visual_catalog.json'))
        meta = bv.metadata()
        self.assertEqual(set(cat['visuals']) - set(meta['visuals']), set())
        self.assertEqual(meta['cli_version'], cat['cli_version'])


class NoNodeNeeded(unittest.TestCase):
    """A full build (scaffold, model, report, validation) runs with no powerbi-report-author on the PC."""

    def test_build_never_calls_the_cli(self):
        real = subprocess.run
        calls = []

        def guard(args, *a, **k):
            calls.append(args)
            if 'powerbi-report-author' in (args if isinstance(args, str) else ' '.join(map(str, args))):
                raise AssertionError('the build called the Node.js CLI')
            return real(args, *a, **k)
        root = tempfile.mkdtemp(prefix='bhy_nonode_')
        bb.subprocess.run = guard
        try:
            b, res = build_fixture(root)
        finally:
            bb.subprocess.run = real
            shutil.rmtree(root, ignore_errors=True)
        self.assertEqual(res['result'], 'succeeded')
        self.assertEqual(calls, [])

    def test_tester_validation_is_offline(self):
        import bhayankar_test as btst
        self.assertFalse(hasattr(btst, 'subprocess'))
        self.assertIs(btst.bv, bv)


class ToolkitIdentity(unittest.TestCase):
    def test_name_version_and_banner(self):
        self.assertEqual(bt.NAME, 'Bhayankar Standalone Toolkit')
        with open(os.path.join(TOOL, 'VERSION.txt'), encoding='utf-8') as f:
            self.assertIn(bt.VERSION, f.read())
        self.assertIn('Standalone', bt.banner())

    def test_folders_can_be_moved_with_environment_variables(self):
        keys = ('BHAYANKAR_INPUT', 'BHAYANKAR_OUTPUT', 'BHAYANKAR_TEST_INPUT', 'BHAYANKAR_TEST_OUTPUT')
        old = {k: os.environ.get(k) for k in keys}
        tmp = tempfile.gettempdir()
        try:
            for k in keys:
                os.environ.pop(k, None)
            self.assertEqual(bt.input_dir(), os.path.join(TOOL, 'input'))
            self.assertEqual(bt.test_output_dir(), os.path.join(TOOL, 'test_output'))
            for k in keys:
                os.environ[k] = os.path.join(tmp, k.lower())
            self.assertEqual(bt.input_dir(), os.path.join(tmp, 'bhayankar_input'))
            self.assertEqual(bt.output_dir(), os.path.join(tmp, 'bhayankar_output'))
            self.assertEqual(bt.test_input_dir(), os.path.join(tmp, 'bhayankar_test_input'))
        finally:
            for k, v in old.items():
                if v is None:
                    os.environ.pop(k, None)
                else:
                    os.environ[k] = v

    def test_build_log_names_the_toolkit(self):
        root = tempfile.mkdtemp(prefix='bhy_log_')
        try:
            b, _ = build_fixture(root)
            with open(os.path.join(b.wb_out, 'bhayankar_build_log.md'), encoding='utf-8') as f:
                text = f.read()
        finally:
            shutil.rmtree(root, ignore_errors=True)
        self.assertIn(bt.banner(), text)
        self.assertIn('offline', text)


class WorkbooksWithoutDashboards(unittest.TestCase):
    """Found while building the fixtures offline: a workbook with no dashboards crashed the build (order[0])."""

    def test_worksheets_become_pages_when_there_is_no_tab_list(self):
        root = tempfile.mkdtemp(prefix='bhy_nodash_')
        try:
            b, res = build_fixture(root, 'collision', ('collision.twb',))
            pages = load(os.path.join(b.report_dir, 'definition', 'pages', 'pages.json'))['pageOrder']
        finally:
            shutil.rmtree(root, ignore_errors=True)
        self.assertEqual(res['errorCount'], 0)
        self.assertEqual(len(pages), 2)  # Org View, Detail View

    def test_no_pages_at_all_gives_one_page_with_a_note(self):
        def empty(spec):
            spec['pages'] = []
            spec['tooltip_pages'] = []
        root = tempfile.mkdtemp(prefix='bhy_nopage_')
        try:
            b, res = build_fixture(root, edit_spec=empty)
            pages = load(os.path.join(b.report_dir, 'definition', 'pages', 'pages.json'))
        finally:
            shutil.rmtree(root, ignore_errors=True)
        self.assertEqual(len(pages['pageOrder']), 1)
        self.assertEqual(pages['activePageName'], pages['pageOrder'][0])
        self.assertEqual(res['errorCount'], 0)
        self.assertTrue(any('no dashboards' in x for x in b.log))


class FilterLiteralTests(unittest.TestCase):
    def test_tableau_null_member_is_a_null_literal(self):
        """IN(%null%, false) on a boolean column must not mix a text literal with a boolean (Power BI rejects the query)."""
        b = bb.Build.__new__(bb.Build)
        self.assertEqual(b.pbir_literal('%null%', 'boolean'), 'null')
        self.assertEqual(b.pbir_literal('false', 'boolean'), 'false')
        self.assertEqual(b.pbir_literal('%null%', 'string'), 'null')


class CardSplitTests(unittest.TestCase):
    def test_multi_measure_card_becomes_stacked_cards(self):
        b = bb.Build.__new__(bb.Build)
        b.log = []
        proj = lambda n: {'field': {'Measure': {}}, 'queryRef': 'T.' + n, 'nativeQueryRef': n}
        vis = {'name': 'abc', 'position': {'x': 10, 'y': 100, 'width': 225, 'height': 150, 'z': 1000, 'tabOrder': 1000},
               'visual': {'visualType': 'cardVisual', 'query': {'queryState': {'Data': {'projections': [proj('A'), proj('B'), proj('C')]}}}},
               'filterConfig': {'filters': [{'name': 'F1'}]}}
        parts = b.split_card(vis)
        self.assertEqual(len(parts), 3)
        self.assertEqual(len({p['name'] for p in parts}), 3)
        self.assertEqual([len(p['visual']['query']['queryState']['Data']['projections']) for p in parts], [1, 1, 1])
        ys = [p['position']['y'] for p in parts]
        self.assertEqual(ys, sorted(ys))
        self.assertTrue(all(p['position']['x'] == 10 and p['position']['width'] == 225 for p in parts))
        self.assertTrue(100 <= ys[0] and ys[-1] + parts[-1]['position']['height'] <= 250)
        short = dict(vis, position={'x': 0, 'y': 0, 'width': 200, 'height': 52, 'z': 1, 'tabOrder': 1})
        sp = b.split_card(short)
        self.assertTrue(max(p['position']['y'] + p['position']['height'] for p in sp) <= 52)
        self.assertEqual(len({p['filterConfig']['filters'][0]['name'] for p in parts[1:]} | {parts[0]['filterConfig']['filters'][0]['name']}), 3)

    def test_checker_reads_split_cards_back_as_one(self):
        import bhayankar_pbip as bp
        b = bb.Build.__new__(bb.Build)
        b.log = []
        proj = lambda n: {'field': {'Measure': {'Expression': {'SourceRef': {'Entity': 'T'}}, 'Property': n}}, 'queryRef': 'T.' + n, 'nativeQueryRef': n}
        vis = {'name': 'abc', 'position': {'x': 0, 'y': 0, 'width': 200, 'height': 150, 'z': 1, 'tabOrder': 1},
               'visual': {'visualType': 'cardVisual', 'query': {'queryState': {'Data': {'projections': [proj('A'), proj('B'), proj('C')]}}}}}
        read = [bp.visual_from_pbir(p, 'pg') for p in b.split_card(vis)]
        for r in read:  # the title text sits on one part only, whichever order the folders sort in
            r['title'] = None
        read[-1]['title'] = 'Card title'
        merged = bp.merge_card_groups(read)
        self.assertEqual(merged[0]['title'], 'Card title')
        self.assertEqual(len(merged), 1)
        self.assertEqual({f['name'] for f in merged[0]['fields']}, {'A', 'B', 'C'})
        self.assertEqual(merged[0]['h'], 150 - 0 if merged[0]['h'] == 150 else merged[0]['h'])

    def test_single_measure_card_is_unchanged(self):
        b = bb.Build.__new__(bb.Build)
        b.log = []
        vis = {'name': 'abc', 'position': {'x': 0, 'y': 0, 'width': 100, 'height': 50},
               'visual': {'visualType': 'cardVisual', 'query': {'queryState': {'Data': {'projections': [{'queryRef': 'T.A'}]}}}}}
        self.assertEqual(b.split_card(vis), [vis])


class TableFileNameTests(unittest.TestCase):
    def test_colon_in_a_table_name_does_not_reach_the_file_name(self):
        self.assertEqual(bb.table_file('Select a metric: Parameter'), 'Select a metric_ Parameter')
        for bad in '<>:"/\|?*':
            self.assertNotIn(bad, bb.table_file(f'a{bad}b'))
        self.assertEqual(bb.table_file('Orders'), 'Orders')


class ConstantAggregateTests(unittest.TestCase):
    def test_average_of_a_constant_is_the_constant(self):
        import bhayankar_dax as bd
        tr = bd.Translator(lambda tok: {'kind': 'unknown', 'table': None, 'name': tok, 'datatype': None})
        for formula, want in (('AVG(0)', '0'), ('MIN(5)', '5'), ('MAX(2)', '2')):
            dax = tr.translate(formula, role='measure')['dax']
            self.assertEqual(dax.strip(), want, formula)
            self.assertNotRegex(dax, r'AVERAGE|MIN\(|MAX\(')


class ForceRunTests(unittest.TestCase):
    def test_force_runs_workbooks_that_are_up_to_date(self):
        import bhayankar_run as br
        ran = []
        orig_scan, orig_one = br.scan, br.run_one
        try:
            br.scan = lambda: [{'name': 'A', 'state': 'done'}, {'name': 'B', 'state': 'new'}]
            br.run_one = lambda n, **kw: ran.append((n, kw.get('force'))) or {'name': n}
            br.run_all(None, force=False)
            self.assertEqual([n for n, _ in ran], ['B'])
            ran.clear()
            br.run_all(None, force=True)
            self.assertEqual([n for n, _ in ran], ['A', 'B'])
        finally:
            br.scan, br.run_one = orig_scan, orig_one


class SelectorGuardTests(unittest.TestCase):
    def test_unresolved_reference_inside_a_selector_is_found(self):
        b = bb.Build.__new__(bb.Build)
        b.model_cols = {'T': {'A'}}
        b.measures = {}
        sel = {'data': [{'scopeId': {'Comparison': {'Left': {'Column': {'Expression': {'SourceRef': {'Entity': 'T'}}, 'Property': 'Gone'}}}}}]}
        self.assertEqual(b.unresolved_ref(sel), 'T.Gone')
        ok = {'data': [{'scopeId': {'Comparison': {'Left': {'Column': {'Expression': {'SourceRef': {'Entity': 'T'}}, 'Property': 'A'}}}}}]}
        self.assertIsNone(b.unresolved_ref(ok))
        self.assertIsNone(b.unresolved_ref(None))


if __name__ == '__main__':
    unittest.main()
