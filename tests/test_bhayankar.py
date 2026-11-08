"""Regression tests for the Bhayankar extractor.

Run (from Bhayankar_Files):  python -m unittest discover -s tests -v
Optional corpus smoke test: set BHAYANKAR_CORPUS to a folder of .twb/.twbx files.
"""
import json
import os
import re
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
import bhayankar_extractor as bx  # noqa: E402
from bhayankar_dax import Translator  # noqa: E402

FIX = os.path.join(HERE, 'fixtures')
TWB = os.path.join(FIX, 'mini.twb')
TDS = os.path.join(FIX, 'orders_source.tds')


def run(tds=(), split=None, path=TWB):
    out = tempfile.mkdtemp(prefix='bhayankar_test_')
    try:
        res, files = bx.process(path, list(tds), out, split)
        with open(os.path.join(out, 'bhayankar_rebuild.json'), encoding='utf-8') as f:
            spec = json.load(f)
        with open(os.path.join(out, 'bhayankar_measures.dax'), encoding='utf-8') as f:
            dax = f.read()
        return res, files, spec, dax
    finally:
        shutil.rmtree(out, ignore_errors=True)


def calc(res, name):
    return next(c for c in res['calculations'] if c['name'] == name)


def page(spec, name):
    return next(p for p in spec['pages'] if p['name'] == name)


def visual(spec, sheet):
    return next(v for p in spec['pages'] for v in p['visuals'] if v['sheet'] == sheet)


class Extraction(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.res, cls.files, cls.spec, cls.dax = run()
        cls.res_t, _, cls.spec_t, _ = run(tds=[TDS])

    def test_used_only(self):
        self.assertIn('Orphan Calc', self.res['orphans']['calculations'])
        self.assertIn('Unused Param', self.res['orphans']['parameters'])
        self.assertNotIn('Never Used Column', [c['name'] for c in self.res['columns']])

    def test_noise_dropped(self):
        md = self.files['bhayankar_digest.md']
        self.assertNotIn('iVBORw0KGgo', md)              # thumbnail
        self.assertNotIn('Calculation_1]', md)           # internal ids resolved to captions
        self.assertIn('Order Count', md)

    def test_published_source(self):
        self.assertIn('PUBLISHED_DS', {f['code'] for f in self.res['flags']})
        self.assertNotIn('PUBLISHED_DS', {f['code'] for f in self.res_t['flags']})
        src = {c['name']: c for c in self.res_t['columns']}
        self.assertEqual(src['Order Number']['source'], 'ORDERS_EDW.ORDER_NUMBER')   # .tds overrides published names
        self.assertEqual(src['Hours']['source_column'], 'HOURS_NO_TRAVEL')
        rels = self.spec_t['model']['relationships']
        self.assertTrue(any(r['from'] == 'ORDERS_EDW' and r['to'] == 'LINES_EDW' for r in rels), rels)
        self.assertTrue(all(r['crossFilter'] == 'OneDirection' for r in rels))

    def test_flags(self):
        codes = {f['code'] for f in self.res['flags']}
        for c in ('LOD_FIXED', 'STRING_RANGE_FILTER', 'DATE_PARAM_VS_DATETIME', 'FILTER_ACTION', 'URL_ACTION',
                  'MEASURE_VALUES'):
            self.assertIn(c, codes)

    def test_dax_translation(self):
        self.assertEqual(calc(self.res, 'Order Count')['dax'], "DISTINCTCOUNTNOBLANK('ORDERS_EDW'[Order Number])")
        lod = calc(self.res, 'Type 7 Hours')['dax']
        self.assertTrue(lod.startswith("SUMX(VALUES('LINES_EDW'[Line Key]), VAR _key1 = 'LINES_EDW'[Line Key] RETURN"), lod)
        self.assertIn("REMOVEFILTERS('LINES_EDW')", lod)
        ratio = calc(self.res, 'Hours Ratio')['dax']   # x/y - 1 -> VAR + one DIVIDE (dax-guidelines)
        self.assertIn('RETURN DIVIDE(_num - _den, _den)', ratio)
        win = calc(self.res, 'In Window')
        self.assertEqual(win['pbi_kind'], 'calculated column')   # // comment stripped, parsed fine
        self.assertTrue(any('cannot react to slicers' in n for n in win['dax_notes']))

    def test_translator_units(self):
        cols = {'[A]': ('column', 'T', 'string'), '[B]': ('column', 'T', 'string'), '[N]': ('column', 'T', 'real'),
                '[M]': ('measure', None, 'real'), '[Parameters].[P]': ('param', None, None)}
        res = lambda t: ({'kind': cols[t][0], 'table': cols[t][1], 'name': t.strip('[]').split('].[')[-1],
                          'datatype': cols[t][2]} if t in cols else {'kind': 'unknown', 'name': t})
        tr = Translator(res)
        self.assertEqual(tr.translate("[A] + ' ' + [B]", 'dimension')['dax'], """'T'[A] & " " & 'T'[B]""")
        self.assertEqual(tr.translate("FIND([A], 'x')", 'dimension')['dax'], """FIND("x", 'T'[A], 1, 0)""")
        self.assertEqual(tr.translate("CASE [A] WHEN 'a' THEN 1 WHEN 'b', 'c' THEN 2 ELSE 3 END", 'dimension')['dax'],
                         """SWITCH('T'[A], "a", 1, "b", 2, "c", 2, 3)""")
        self.assertEqual(tr.translate('ZN(SUM([N]))', 'measure')['dax'], "COALESCE(SUM('T'[N]), 0)")
        self.assertEqual(tr.translate('SUM(IF [A] = "x" THEN [N] END)', 'measure')['dax'],
                         """SUMX('T', IF('T'[A] = "x", 'T'[N]))""")
        self.assertEqual(tr.translate("{EXCLUDE [A]: SUM([N])}", 'measure')['dax'],
                         "CALCULATE(SUM('T'[N]), REMOVEFILTERS('T'[A]))")
        self.assertEqual(tr.translate("1 - SUM([N]) / [M]", 'measure')['dax'],
                         "(VAR _num = SUM('T'[N]) VAR _den = [M] RETURN DIVIDE(_den - _num, _den))")
        fp = tr.translate("CASE [Parameters].[P] WHEN 'a' THEN [A] WHEN 'b' THEN [B] END", 'dimension')
        self.assertEqual(fp['kind'], 'field parameter')
        self.assertEqual(tr.translate('RUNNING_SUM(SUM([N]))', 'measure')['confidence'], 'low')
        self.assertIn('parse error', tr.translate('IF [A] THEN', 'dimension')['notes'][0])

    def test_owner_chart(self):
        v = visual(self.spec, 'Owner Bars')
        self.assertEqual(v['visualType'], 'clusteredBarChart')   # nested inner dim = colour -> legend
        r = {k: [f['name'] for f in fs] for k, fs in v['roles'].items()}
        self.assertEqual(r['Category'], ['Owner Name'])
        self.assertEqual(r['Series'], ['Year'])
        self.assertEqual(r['Y'], ['Type 7 Hours'])
        self.assertEqual(r['Tooltips'], ['Order Count'])
        self.assertEqual(v['catalog_errors'], [])
        self.assertEqual(v['position'], {'x': 0, 'y': 160, 'width': 1000, 'height': 640})
        self.assertEqual(v['title']['measure'], 'Owner Bars Title')
        self.assertEqual(v['colors']['map']['YEAR(Approved Date)']['2026'], '#dbdb8d')
        self.assertTrue(v['data_labels'])
        rc = [f for f in v.get('visual_filters', []) if f.get('type') == 'RowCondition']
        self.assertEqual(rc[0]['calc'], 'In Window')                 # param-driven row calc -> folded into the measures
        self.assertTrue(rc[0]['condition'].endswith('= "show"'))

    def test_slicers(self):
        p = page(self.spec, 'Main')
        between = [s for s in p['slicers'] if s.get('mode') == 'Between']
        self.assertEqual(len(between), 1)                            # Start/End pair merged
        self.assertEqual(between[0]['default'], {'from': '#2026-01-01#', 'to': '#2026-11-01#'})
        grp = next(s for s in p['slicers'] if s['field']['name'] == 'Dept Num (group)')
        self.assertEqual((grp['visualType'], grp['mode']), ('slicer', 'Dropdown'))
        self.assertEqual(grp['default']['values'], ['Design'])
        st = next(s for s in p['slicers'] if s['field']['name'] == 'Order Status')
        self.assertIn('DECISION NEEDED', st['default']['note'])    # never collapsed to "APPR"
        self.assertEqual(p['page_filters'], [])                     # slicer fields are not page filters

    def test_drillthrough_and_sync(self):
        d = page(self.spec, 'Detail')
        names = [f['name'] for f in d['drillthrough']['fields']]
        self.assertEqual(names, ['Owner Name', 'Year'])
        self.assertTrue(all(s.get('hidden') for s in d['slicers']))
        # only the Report Period (parameters) is synced: the two quick filters belong to "Owner Bars" alone, and
        # Tableau does not apply them to the detail sheet (DesignReviewRerunFixes.test_07)
        self.assertEqual([s['field']['name'] for s in d['slicers']], ['Date'])

    def test_detail_table(self):
        v = visual(self.spec, 'Detail Table')
        self.assertEqual(v['visualType'], 'tableEx')   # URL action needs a web URL on a dimension column
        self.assertEqual([f['name'] for f in v['roles']['Values']], ['Order Number', 'Order ID', 'Order Count', 'Hours Ratio'])
        self.assertEqual(v['url_action']['fields'], ['Order ID'])
        self.assertEqual(v['url_action']['link_on'], 'Order Number')
        self.assertEqual(self.spec['model']['url_columns'][0]['name'], 'Open in System URL')
        self.assertEqual(v['catalog_errors'], [])

    def test_static_and_help(self):
        self.assertEqual(visual(self.spec, 'Help')['visualType'], 'help-icon')
        self.assertIn('Help', self.spec['tooltip_pages'])
        p = page(self.spec, 'Main')
        self.assertTrue(any(s['type'] == 'image' and s['image'] == 'Image/logo.png' for s in p['static']))

    def test_model_rules(self):
        tables = {t['name']: t for t in self.spec['model']['tables']}
        self.assertIn('Order Count', [m['name'] for m in tables['ORDERS_EDW']['measures']])   # on its fact table
        oc = next(m for m in tables['ORDERS_EDW']['measures'] if m['name'] == 'Order Count')
        self.assertEqual(oc['suggested_name'], '# Order')
        self.assertEqual(oc['formatString'], '#,##0')
        self.assertEqual(self.spec['model']['date_table']['active_relationship_to'], 'Approved Date')
        self.assertTrue(any(c['name'] == 'Dept Num (group)' for c in tables['EMPLOYEE']['calculated_columns']))
        self.assertNotIn('Spacer', [m['name'] for t in tables.values() for m in t['measures']])

    def test_dax_script(self):
        self.assertIn('DEFINE', self.dax)
        self.assertIn("MEASURE 'ORDERS_EDW'[Order Count] = DISTINCTCOUNTNOBLANK('ORDERS_EDW'[Order Number])", self.dax)
        self.assertIn('EVALUATE', self.dax)

    def test_split_mode(self):
        _, files, _, _ = run(split=True)
        self.assertIn('bhayankar_digest.md', files)
        main = next(v for k, v in files.items() if k.startswith('bhayankar_digest_01_'))
        self.assertIn('Owner Bars', main)
        self.assertIn('Power BI rebuild plan', main)

    def test_truncated_file(self):
        tmp = tempfile.mkdtemp(prefix='bhayankar_trunc_')
        try:
            with open(TWB, encoding='utf-8') as f:
                text = f.read()
            cut = os.path.join(tmp, 'cut.twb')
            with open(cut, 'w', encoding='utf-8') as f:
                f.write(text[:text.index('<dashboard name=\'Detail\'>') + 40])
            res, _, _, _ = run(path=cut)
            self.assertEqual(res['flags'][0]['code'], 'TRUNCATED_FILE')
            self.assertEqual(res['workbook']['worksheets'], 3)       # everything before the cut is kept
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


class FolderMode(unittest.TestCase):
    def test_input_to_output(self):
        root = tempfile.mkdtemp(prefix='bhayankar_folder_')
        try:
            inp, out = os.path.join(root, 'input'), os.path.join(root, 'output')
            os.makedirs(inp)
            shutil.copy(TWB, inp)
            shutil.copy(TDS, inp)
            with open(TDS, encoding='utf-8') as f:   # a .tds for some other source must not be attached
                other = f.read().replace("formatted-name='ORDERS SOURCE'", "formatted-name='SOMETHING ELSE'")
            with open(os.path.join(inp, 'something_else.tds'), 'w', encoding='utf-8') as f:
                f.write(other)
            rows = bx.run_folder(inp, out, table_map={'ORDERS_EDW': 'Orders'})
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0]['tds_merged'], 'orders_source.tds')
            self.assertTrue(os.path.exists(os.path.join(out, 'inventory.csv')))
            wb_out = os.path.join(out, 'mini')
            for fn in ('bhayankar_digest.md', 'bhayankar_digest.json', 'bhayankar_rebuild.json', 'bhayankar_measures.dax'):
                self.assertTrue(os.path.exists(os.path.join(wb_out, fn)), fn)
            with open(os.path.join(wb_out, 'bhayankar_measures.dax'), encoding='utf-8') as f:
                self.assertIn("MEASURE 'Orders'[Order Count] = DISTINCTCOUNTNOBLANK('Orders'[Order Number])", f.read())
        finally:
            shutil.rmtree(root, ignore_errors=True)



class BuildHelpers(unittest.TestCase):
    def test_m_identifiers(self):
        import bhayankar_build as bb
        self.assertEqual(bb.mfield('WO_ID'), '[WO_ID]')
        self.assertEqual(bb.mfield('Dept Num-WO Owner Now'), '[#"Dept Num-WO Owner Now"]')   # Desktop: "Invalid identifier"
        self.assertEqual(bb.mid('375'), '#"375"')                                            # record keys starting with a digit
        self.assertEqual(bb.mid('Say "hi"'), '#"Say ""hi"""')

class Accuracy(unittest.TestCase):
    """Gaps found comparing the Security Awareness digest with its TWB (collision.twb reproduces them)."""
    @classmethod
    def setUpClass(cls):
        cls.res, cls.files, cls.spec, cls.dax = run(path=os.path.join(FIX, 'collision.twb'))
        cls.md = ''.join(cls.files.values())
        cls.flags = {}
        for f in cls.res['flags']:
            cls.flags.setdefault(f['code'], []).append(f)

    def test_relative_date_keeps_all_attributes(self):
        self.assertIn('include-future=true', self.md)
        self.assertIn('period-type-v2=month', self.md)

    def test_name_collision(self):
        versions = [c for c in self.res['calculations'] if c['name'] == 'Manager Level']
        self.assertEqual(len(versions), 2)
        self.assertEqual({c['datasource'] for c in versions}, {'EMAILS ORG', 'EMAILS DETAIL'})
        self.assertTrue(all('NAME_COLLISION' in c['flags'] for c in versions))
        detail = self.flags['NAME_COLLISION'][0]['detail']
        self.assertIn('[EMAILS ORG] used in Org View', detail)
        self.assertIn('[EMAILS DETAIL] used in Detail View', detail)
        self.assertIn('**Manager Level** from [EMAILS DETAIL]', self.md)

    def test_per_user_and_hardcoded_logins(self):
        self.assertIn('Manager Level', self.flags['PER_USER_COLUMN'][0]['where'])
        details = ' '.join(f['detail'] for f in self.flags['HARDCODED_USERS'])
        self.assertIn('E001, E002', details)

    def test_no_internal_ids(self):
        self.assertNotIn('Calculation_', self.md)
        self.assertNotIn('federated.', self.md)
        self.assertIn('EMAILS ORG, EMAILS DETAIL', self.flags['BLENDING'][0]['detail'])

    def test_measure_names_palette(self):
        pal = self.res['palettes']['Measure Names']
        self.assertEqual(pal.get('SUM(CLICKED)'), '#20415b')
        self.assertFalse(any('deleted field' in k for k in pal))

    def test_unused_copy_not_orphan(self):
        self.assertIn('Officers', self.res['orphans']['calculations'])
        self.assertNotIn('Manager Level', self.res['orphans']['calculations'])

    def test_measure_without_home_table(self):
        import bhayankar_report as br
        spec = {'model': {'tables': [{'measures': [
            {'name': 'A', 'table': 'Emails', 'dax': '1', 'confidence': 'high', 'description': ''},
            {'name': 'B', 'table': None, 'dax': '2', 'confidence': 'high', 'description': ''}]}]}}
        script = br.dax_script(spec)
        self.assertNotIn('<Fact table>', script)
        self.assertIn("MEASURE 'Emails'[B]", script)


class BuildSmoke(unittest.TestCase):
    """Extractor output -> TMDL model + PBIR report on disk (skips the powerbi-report-author validate call)."""

    def test_model_and_report_written(self):
        import bhayankar_build as bb
        root = tempfile.mkdtemp(prefix='bhayankar_build_')
        old = bb.INPUT_DIR, bb.OUTPUT_DIR
        try:
            inp, out = os.path.join(root, 'input'), os.path.join(root, 'output')
            os.makedirs(inp)
            shutil.copy(TWB, inp)
            shutil.copy(TDS, inp)
            bx.run_folder(inp, out, table_map={})
            bb.INPUT_DIR, bb.OUTPUT_DIR = inp, out
            b = bb.Build('mini')
            b.build_model()
            b.write_model()
            b.write_report()
            self.assertTrue(os.path.exists(os.path.join(b.pbip_dir, 'mini.pbip')))
            tmdl = [os.path.join(dp, f) for dp, _, fs in os.walk(b.model_dir) for f in fs if f.endswith('.tmdl')]
            self.assertTrue(any(f.endswith('model.tmdl') for f in tmdl), tmdl)
            self.assertTrue(any(f.endswith('.json') for _, _, fs in os.walk(b.report_dir) for f in fs))
        finally:
            bb.INPUT_DIR, bb.OUTPUT_DIR = old
            shutil.rmtree(root, ignore_errors=True)


class StandaloneBuild(unittest.TestCase):
    """Run_Build.bat works without Run_Extractor.bat, and never builds from a stale extract."""

    def test_extracts_on_demand(self):
        import bhayankar_build as bb
        root = tempfile.mkdtemp(prefix='bhayankar_alone_')
        old = bb.INPUT_DIR, bb.OUTPUT_DIR
        try:
            inp, out = os.path.join(root, 'input'), os.path.join(root, 'output')
            os.makedirs(inp)
            shutil.copy(TWB, inp)
            shutil.copy(TDS, inp)
            bb.INPUT_DIR, bb.OUTPUT_DIR = inp, out
            spec = os.path.join(out, 'mini', 'bhayankar_rebuild.json')
            self.assertEqual(bb.ensure_extracted([]), ['mini'])                    # nothing extracted yet
            self.assertTrue(os.path.exists(spec))
            first = os.path.getmtime(spec)
            self.assertEqual(bb.ensure_extracted(['mini']), ['mini'])              # up to date: no re-extract
            self.assertEqual(os.path.getmtime(spec), first)
            later = first + 60
            os.utime(os.path.join(inp, 'mini.twb'), (later, later))               # workbook changed
            bb.ensure_extracted(['mini'])
            self.assertGreater(os.path.getmtime(spec), first)
            self.assertEqual(bb.ensure_extracted(['nope']), [])                    # unknown workbook: reported, skipped
        finally:
            bb.INPUT_DIR, bb.OUTPUT_DIR = old
            shutil.rmtree(root, ignore_errors=True)


class LoadSafety(unittest.TestCase):
    """Security Awareness PBIP did not open in Desktop (2026-10-25): every rule here maps to one of the errors."""

    def tr(self):
        cols = {'[A]': ('column', 'T', 'string'), '[D]': ('column', 'T', 'datetime'), '[N]': ('column', 'T', 'real'),
                '[M]': ('measure', None, 'real')}
        res = lambda t: ({'kind': cols[t][0], 'table': cols[t][1], 'name': t.strip('[]'), 'datatype': cols[t][2]}
                         if t in cols else {'kind': 'unknown', 'name': t})
        return Translator(res)

    def test_attr_of_measure_is_not_selectedvalue(self):
        tr = self.tr()
        self.assertEqual(tr.translate('ATTR([A])', 'dimension')['dax'], "SELECTEDVALUE('T'[A])")
        out = tr.translate('IF ATTR([M]) = 1 THEN 1 ELSE 2 END', 'measure')['dax']   # "Column 'M' cannot be found"
        self.assertNotIn('SELECTEDVALUE', out)
        self.assertNotIn('SELECTEDVALUE', tr.translate('IF ATTR([M] = 1) THEN 1 END', 'measure')['dax'])  # "Column ''"

    def test_fixed_lod_with_expression_grain(self):
        out = self.tr().translate("{FIXED [A], DATEDIFF('month', [D], TODAY()) <= 12 : SUM([N])}", 'measure')['dax']
        # the grain expression is null-safe (Tableau: NULL date -> NULL), see DesignReviewGeneratorFixes
        dd = "IF(ISBLANK('T'[D]), BLANK(), DATEDIFF('T'[D], TODAY(), MONTH))"
        self.assertRegex(out, r"SUMMARIZE\(ADDCOLUMNS\('T', \"@d2\", \(VAR _v\d+ = " + re.escape(dd) +
                         r" RETURN NOT ISBLANK\(_v\d+\) && _v\d+ <= 12\)\), 'T'\[A\], \[@d2\]\)")
        self.assertRegex(out, r"FILTER\(ALL\('T'\), \(VAR _v\d+ = " + re.escape(dd) + r" RETURN .*\) = _key2\)")

    def guard_build(self):
        import bhayankar_build as bb
        b = bb.Build.__new__(bb.Build)
        b.log, b.decisions, b.report_period, b.slice_tables, b.param_tables = [], {}, None, {}, {}
        b.digest = {'flags': [{'code': 'HARDCODED_USERS', 'detail': 'logins E01, E02 hard-coded: move them'}]}
        b.tables = {'F': {'name': 'F', 'columns': [{'name': 'A'}, {'name': 'Login'}],
                          'calc_columns': [{'name': 'Mine', 'dax': "IF('F'[Login] = USERPRINCIPALNAME(), 1, 0)"}]}}
        m = lambda n, dax, t='F', notes=(): {'name': n, 'dax': dax, 'table': t, 'description': '', 'displayFolder': 'Measures',
                                             'formatString': None, 'hidden': False, 'notes': list(notes)}
        b.measures = {x['name']: x for x in [
            m('Ok', "COUNTROWS('F')"), m('Lost', '1', 'Year Parameter'), m('Missing', '[Nope] + 1'),
            m('VisCalc', '/* visual calculation */ PERCENT_OF_TOTAL([Ok])'), m('SelMeasure', 'SELECTEDVALUE([Ok])'),
            m('Grain', "SUMX(SUMMARIZE('F', 'F'[A], 'F'[A] = \"x\"), 1)"),
            m('Mixed', "IF([Ok] = 1, 'F'[A])", notes=['mixes row-level columns (A) with measures/aggregates: not valid DAX as written'])]}
        b.load_guard()
        return b

    def test_load_guard(self):
        b = self.guard_build()
        self.assertEqual(b.tables['F']['calc_columns'][0]['dax'], "IF('F'[Login] = \"E01\", 1, 0)")   # no UPN in a column
        self.assertEqual(b.measures['Ok']['dax'], "COUNTROWS('F')")
        self.assertEqual(b.measures['Lost']['table'], 'F')
        for n in ('Missing', 'VisCalc', 'SelMeasure', 'Grain', 'Mixed'):
            self.assertEqual(b.measures[n]['dax'], 'BLANK()', n)
            self.assertEqual(b.measures[n]['displayFolder'], 'Measures\\_Needs fix', n)
            self.assertIn('NOT TRANSLATED', b.measures[n]['description'], n)

    def test_slicer_without_table_is_skipped_not_a_crash(self):
        # Field from a published data source with no .tds: table None crashed the build (TypeError, 2026-10-26)
        b = self.guard_build()
        b.growth = []
        p = {'name': 'P'}
        s = {'id': 's1', 'visualType': 'slicer', 'mode': 'Dropdown', 'position': None, 'default': {},
             'field': {'label': 'Region', 'kind': 'Unknown', 'table': None, 'name': 'Region'}}
        self.assertIsNone(b.slicer(p, s))
        self.assertIn('slicer "Region" not generated', b.log[-1])
        b.param_tables = {'Year Parameter': ('DATATABLE(...)', 'Value', 'string')}  # parameter slicer -> its table
        sp = {**s, 'parameter': 'Year', 'field': {'label': 'Param:Year', 'kind': 'Measure', 'table': None, 'name': 'Year Value'}}
        vis = b.slicer(p, sp)
        self.assertEqual(vis['visual']['query']['queryState']['Values']['projections'][0]['queryRef'], 'Year Parameter.Value')

    def test_one_bad_item_does_not_stop_the_build(self):
        b = self.guard_build()
        self.assertIsNone(b.safe({'name': 'P'}, 'visual X', lambda: 1 / 0))
        self.assertIn('visual X NOT GENERATED', b.log[-1])

    def test_report_guard(self):
        import bhayankar_build as bb
        b = self.guard_build()
        good = bb.proj(bb.col_ref('F', 'A'))
        page = {'displayName': 'P', 'filterConfig': {'filters': [{'name': 'f1', 'field': bb.col_ref('F', 'Ok')}]}}
        visuals = [
            {'name': 'v1', 'visual': {'visualType': 'slicer', 'query': {'queryState': {'Values': {'projections': [
                bb.proj(bb.col_ref('Year Parameter', 'Value'))]}}}}},
            {'name': 'v2', 'visual': {'visualType': 'tableEx', 'query': {'queryState': {'Values': {'projections': [
                good, bb.proj(bb.meas_ref('F', 'Gone'))]}}}},
             'filterConfig': {'filters': [{'name': 'f2', 'field': bb.col_ref('F', 'Ok')}]}}]   # a measure as a column filter
        out = b.report_guard(page, visuals)
        self.assertEqual([v['name'] for v in out], ['v2'])
        self.assertEqual(out[0]['visual']['query']['queryState']['Values']['projections'], [good])
        self.assertEqual(out[0]['filterConfig']['filters'], [])
        self.assertEqual(page['filterConfig']['filters'], [])


class DesignReviewGeneratorFixes(unittest.TestCase):
    """Design Review 1X1 build (2026-10-29): 14 generator bugs that were hand-fixed in the output.
    One test per bug (the number is the one in docs\\PROJECT_LOG.md)."""

    def tr(self, extra=None):
        cols = {'[A]': ('column', 'T', 'string'), '[D]': ('column', 'T', 'datetime'), '[E]': ('column', 'T', 'datetime'),
                '[K]': ('column', 'T', 'string'), '[U]': ('column', 'W', 'boolean'), **(extra or {})}

        def res(t):
            if t not in cols:
                return {'kind': 'unknown', 'name': t}
            kind, table, dt = cols[t][:3]
            return {'kind': kind, 'table': table, 'name': t.strip('[]'), 'datatype': dt,
                    'row_dax': cols[t][3] if len(cols[t]) > 3 else None}
        return Translator(res)

    def build(self, tables=None, sources=None):
        import bhayankar_build as bb
        b = bb.Build.__new__(bb.Build)
        b.log, b.decisions, b.wb = [], {}, 'WB'
        b.sources = sources or {'mode': 'warehouse', 'tables': {}}
        b.relationships, b.moved_tables, b.moved_cols = [], {}, {}
        b.tables = tables or {}
        return b

    def rebuild(self, calcs=(), worksheets=()):
        from bhayankar_report import Rebuild
        r = Rebuild.__new__(Rebuild)
        r.res = {'worksheets': list(worksheets), 'actions': []}
        r.cat, r.notes, r.measures, r.date_roles, r.url_columns, r.param_dims = None, [], {}, [], [], {}
        r.cols = {'Task Number': {'name': 'Task Number', 'table': 'T', 'datatype': 'integer'},
                  'Met': {'name': 'Met', 'table': 'T', 'datatype': 'string'},
                  'Month Col': {'name': 'Month Col', 'table': 'T', 'datatype': 'string'}}
        r.calcs = {c['name']: c for c in calcs}
        r.groups, r.params, r.formats = {}, {}, {}
        return r

    def test_01_url_action_on_several_sheets_is_one_column(self):
        r = self.rebuild()
        r.res['actions'] = [{'kind': 'URL', 'name': 'Link', 'source': 'P / S1', 'url': 'https://x/?id=<[Met]>'},
                            {'kind': 'URL', 'name': 'Link', 'source': 'P / S2', 'url': 'https://x/?id=<[Met]>'}]
        pages = [{'name': 'P', 'static': [], 'notes': [],
                  'visuals': [{'sheet': s, 'visualType': 'cardVisual', 'roles': {}, 'notes': []} for s in ('S1', 'S2')]}]
        r.actions(pages, {})
        self.assertEqual(len(r.url_columns), 1)

    def test_02_parameter_default_with_a_comma(self):
        import bhayankar_build as bb
        args = bb.call_args('SELECTEDVALUE(\'DC Parameter\'[Value], "Demenko,Natalya S")', 'SELECTEDVALUE')
        self.assertEqual(args[1], '"Demenko,Natalya S"')

    def test_03_sum_around_a_measure(self):
        out = self.tr({'[M]': ('measure', 'T', 'real')}).translate('SUM([M]) / COUNTD([A])', 'measure')['dax']
        self.assertEqual(out, "DIVIDE([M], DISTINCTCOUNTNOBLANK('T'[A]))")

    def test_04_05_06_sample_rename_keeps_columns_and_source_names(self):
        b = self.build()
        t = {'name': 'W', 'source': 'W', 'steps': [], 'columns': [
            {'name': 'Ready Date', 'source_column': 'Const Ready', 'datatype': 'datetime'},
            {'name': 'Other', 'source_column': 'Other', 'datatype': 'string'},
            {'name': 'Coord', 'source_column': 'Coord', 'datatype': 'string'}]}
        # mapped by Power BI name (already the CSV header) and by source column (needs a rename)
        cfg = {'file': 'w.csv', 'columns': {'Ready Date': 'Const Ready', 'Coord': 'CSV Coordinator'}}
        self.assertEqual(b.sample_renames(t, cfg), [('CSV Coordinator', 'Coord')])
        b.sources = {'mode': 'sample', 'tables': {'W': cfg}}
        b.relations, b.digest = {}, {'groups': []}
        m = '\n'.join(b.m_partition(t))
        self.assertNotIn('Table.SelectColumns(SampleTable', m)          # 4: no pre-restriction to the mapped columns
        self.assertIn('Table.RenameColumns(', m)
        self.assertIn('{"Const Ready", ToDateTime', m)                   # 5: later steps use the M output name
        self.assertIn('"Other"', m)

    def test_07_tables_on_the_same_sample_row_are_merged_and_joins_inferred(self):
        tmp = tempfile.mkdtemp(prefix='bhayankar_orphan_')
        try:
            with open(os.path.join(tmp, 'wo.csv'), 'w', encoding='utf-8') as f:
                f.write('WO Number,Design Coordinator,Owner\n1,a,x\n2,b,y\n')
            with open(os.path.join(tmp, 'task.csv'), 'w', encoding='utf-8') as f:
                f.write('WO Number (TASK),Task\n1,t1\n1,t2\n2,t3\n')
            src = {'mode': 'sample', 'sample_folder': tmp, 'tables': {
                'WO': {'file': 'wo.csv'}, 'TASK': {'file': 'task.csv'},
                'COORD': {'file': 'wo.csv', 'columns': {'Coord Name': 'Design Coordinator'}},
                'OWNER': {'file': 'wo.csv', 'columns': {'Owner Now': 'Owner'}}}}
            col = lambda n: {'name': n, 'source_column': n, 'datatype': 'string'}
            tables = {'WO': {'name': 'WO', 'source': 'WO', 'steps': [], 'columns': [col('WO Number'), col('Owner')]},
                      'TASK': {'name': 'TASK', 'source': 'TASK', 'steps': [], 'columns': [col('Task')]},
                      'COORD': {'name': 'COORD', 'source': 'COORD', 'steps': [], 'columns': [col('Coord Name')]},
                      'OWNER': {'name': 'OWNER', 'source': 'OWNER', 'steps': [], 'columns': [col('Owner Now')]}}
            b = self.build(tables, src)
            b.merge_same_sample(tables)
            self.assertEqual(sorted(tables), ['TASK', 'WO'])
            self.assertEqual(b.remap('COORD', 'Coord Name'), ('WO', 'Coord Name'))
            self.assertEqual(b.remap('OWNER', 'Owner Now'), ('WO', 'Owner'))   # same CSV column: the existing one
            self.assertIn(('Design Coordinator', 'Coord Name'), tables['WO']['extra_renames'])
            self.assertEqual(b.remap_dax("FILTER('COORD', 'COORD'[Coord Name] = 1)"), "FILTER('WO', 'WO'[Coord Name] = 1)")
            b.infer_relationships(tables)
            self.assertEqual(b.relationships, [('TASK', 'WO Number (TASK)', 'WO', 'WO Number')])
            b.flag_orphans(tables)
            self.assertFalse([x for x in b.log if 'ORPHAN' in x])
            spec = [{'visuals': [{'roles': {'Rows': [{'kind': 'Column', 'table': 'OWNER', 'name': 'Owner Now'}]}}]}]
            b.remap_spec(spec)
            self.assertEqual(spec[0]['visuals'][0]['roles']['Rows'][0], {'kind': 'Column', 'table': 'WO', 'name': 'Owner'})
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_07_unrelated_table_is_flagged(self):
        b = self.build()
        b.flag_orphans({'A': {}, 'B': {}})
        self.assertEqual(len([x for x in b.log if 'WARNING ORPHAN TABLE' in x]), 2)

    def test_08_long_paths_get_short_folder_names(self):
        b = self.build()
        b.wb, b.spec = 'Design Review 1X1 Dashboard', {'model': {'tables': []}}
        b.pbip_dir = 'C:\\' + 'x' * 140 + '\\pbip'   # leaves room for ~20 characters of name
        item = b.item_name()
        self.assertLess(len(item), len(b.wb))
        self.assertTrue(b.wb.startswith(item))
        self.assertLessEqual(len(os.path.abspath(b.pbip_dir)) + 1 + len(item) + b.REPORT_TAIL, b.MAX_PATH)

    def test_08_fallback_folder_never_reuses_a_folder_it_did_not_make(self):
        tmp = tempfile.mkdtemp(prefix='bhayankar_pbi_')
        old = os.environ.get('SystemDrive')
        try:
            os.environ['SystemDrive'] = tmp
            os.makedirs(os.path.join(tmp, 'PBI', 'DR'))          # e.g. a hand-fixed copy of the report
            b = self.build()
            b.wb = 'Design Review'
            first = b.fallback_folder()
            self.assertEqual(os.path.basename(first), 'DR_2')
            self.assertEqual(os.listdir(os.path.join(tmp, 'PBI', 'DR')), [])  # untouched
            self.assertEqual(b.fallback_folder(), first)            # its own folder is reused on the next build
            b.wb = 'Data Release'                            # same initials, other workbook
            self.assertEqual(os.path.basename(b.fallback_folder()), 'DR_3')
        finally:
            if old is None:
                os.environ.pop('SystemDrive', None)
            else:
                os.environ['SystemDrive'] = old
            shutil.rmtree(tmp, ignore_errors=True)

    def test_09_parameter_driven_dimension(self):
        f = ('IF [D] >= [Param:Start] AND [A] = [Param:DC] THEN "Tasks APPR" ELSEIF [E] >= [Param:Start] '
             'THEN "Tasks INPRG" END')
        calc = {'name': 'Tasks Managed', 'pbi_kind': 'calculated column', 'formula': f, 'table': 'T', 'dax': 'X'}
        r = self.rebuild([calc], [{'name': 'S', 'manual_sorts': {'Tasks Managed': ['Tasks INPRG', 'Tasks APPR']}}])
        fr = r.column_ref('Tasks Managed')
        self.assertEqual((fr['table'], fr['name'], fr['param_dim']), ('Tasks Managed', 'Tasks Managed', True))
        self.assertEqual(r.param_dims['Tasks Managed']['values'], ['Tasks INPRG', 'Tasks APPR'])
        v = {'roles': {'Rows': [fr], 'Values': [{'kind': 'Measure', 'name': 'Task Count'}]}, 'notes': []}
        r.add_param_dim_filters(v)
        self.assertEqual(v['visual_filters'][0]['condition'], "X IN VALUES('Tasks Managed'[Tasks Managed])")
        t = {'roles': {'Values': [fr, {'kind': 'Column', 'table': 'T', 'name': 'A'}]}, 'notes': []}   # no measure
        r.add_param_dim_filters(t)
        self.assertEqual(t['roles']['Values'], [{'kind': 'Column', 'table': 'T', 'name': 'A'}])

    def test_10_row_calc_over_fixed_and_row_calcs_is_a_column(self):
        tr = self.tr()
        fixed = tr.translate("{FIXED [K]: MIN(DATEDIFF('day', [D], [E]))}", 'measure')
        self.assertIn("ALLEXCEPT('T', 'T'[K])", fixed['row_dax'])
        row = tr.translate("DATEDIFF('day', [D], [E])", 'measure', 'Sum')
        self.assertEqual(row['kind'], 'measure')
        self.assertTrue(row['row_dax'])
        tr2 = self.tr({'[F]': ('measure', 'T', 'integer', fixed['row_dax']), '[R]': ('measure', 'T', 'integer', row['row_dax'])})
        out = tr2.translate('IF [U] = TRUE AND [R] <= 1 THEN "Met" ELSE "Not Met" END', 'dimension')
        self.assertEqual(out['kind'], 'calculated column')
        self.assertNotIn('[R]', out['dax'])
        self.assertEqual(tr2.translate('IF [F] <= 40 THEN "Met" ELSE "Not Met" END', 'dimension')['kind'], 'calculated column')
        # cross-table column: placed on the many side, the one side read through RELATED
        b = self.build({'T': {'name': 'T', 'calc_columns': [{'name': 'Met', 'dax': "IF('W'[U] && 'T'[D] > 1, 1)"}]},
                        'W': {'name': 'W', 'calc_columns': []}})
        b.relationships = [('T', 'K', 'W', 'K')]
        b.place_calc_columns(b.tables)
        self.assertEqual(b.tables['T']['calc_columns'][0]['dax'], "IF(RELATED('W'[U]) && 'T'[D] > 1, 1)")

    def test_10_null_compares_like_tableau(self):
        out = self.tr().translate("IF DATEDIFF('day', [D], [E]) <= 40 THEN 'Met' ELSE 'Not Met' END", 'dimension')['dax']
        self.assertIn("IF(ISBLANK('T'[D]) || ISBLANK('T'[E]), BLANK(), DATEDIFF('T'[D], 'T'[E], DAY))", out)
        self.assertRegex(out, r'NOT ISBLANK\(_v\d+\) && _v\d+ <= 40')

    def test_11_week_start_monday(self):
        out = self.tr().translate("DATEDIFF('week', [D], [E], 'monday')", 'dimension')['dax']
        self.assertIn("DATEDIFF('T'[D] - 1, 'T'[E] - 1, WEEK)", out)
        self.assertIn("'T'[D] - WEEKDAY('T'[D] - 1, 1) + 1", self.tr().translate("DATETRUNC('week', [D], 'monday')", 'dimension')['dax'])

    def test_12_percent_of_total_follows_compute_using(self):
        met = {'label': 'Met', 'measure': False}
        r = self.rebuild()
        r.measures['Task Count'] = {'name': 'Task Count', 'table': 'T', 'dax': 'X'}
        r.cols['Task Count'] = None
        r.column_ref = lambda n: {'kind': 'Measure', 'table': 'T', 'name': n} if n == 'Task Count' else \
            {'kind': 'Column', 'table': 'T', 'name': n}
        r._ws = {'shelf': {'rows': [], 'cols': [met]}, 'panes': [{'enc': [{'label': 'Met', 'measure': False, 'channel': 'color'}]}]}
        fr = r.field('PERCENT_OF_TOTAL(Task Count)', True, tc='Rows')     # Table (across), nothing on Rows: whole view
        self.assertEqual(r.measures[fr['name']]['dax'], "DIVIDE([Task Count], CALCULATE([Task Count], ALLSELECTED('T'[Met])))")
        r._ws = {'shelf': {'rows': [], 'cols': [{'label': 'Month Col', 'measure': False}]},
                 'panes': [{'enc': [{'label': 'Met', 'measure': False, 'channel': 'color'}]}]}
        fr2 = r.field('PERCENT_OF_TOTAL(Task Count)', True, tc='Columns')  # Table (down): per column (month)
        self.assertIn("ALLSELECTED('T'[Met])", r.measures[fr2['name']]['dax'])
        self.assertNotIn('Month Col', r.measures[fr2['name']]['dax'])

    def test_13_filters_tableau_writes_without_quotes_or_as_null_rules(self):
        r = self.rebuild([{'name': 'Stage', 'pbi_kind': 'calculated column', 'table': 'T', 'dax': 'X',
                           'formula': 'IF [A] = "1" THEN "Tasks APPR" ELSEIF [A] = "2" THEN "Tasks INPRG" ELSE "Tasks COMP" END'}])
        spec = lambda field, rule: r.filter_spec({'field': field, 'rule': rule})
        self.assertEqual(spec('Task Number', 'IN(225, 226)  [values:database]')['values'], ['225', '226'])
        self.assertEqual(spec('Month Col', 'NOT ALL EXCEPT %null%  [values:relevant]')['type'], 'NotBlank')
        self.assertEqual(spec('Month Col', 'RANGE -inf..+inf (non-null)')['type'], 'NotBlank')
        rng = spec('Stage', 'RANGE("Tasks APPR".."Tasks INPRG")  [values:database]')
        self.assertEqual(rng['values'], ['Tasks APPR', 'Tasks INPRG', 'Tasks COMP'])
        b = self.build({'T': {'name': 'T', 'columns': [{'name': 'Task Number', 'datatype': 'integer'}], 'calc_columns': []}})
        f = b.categorical_filter({'type': 'Categorical', 'values': ['225'], 'field': {'table': 'T', 'name': 'Task Number'}}, 'k')
        self.assertEqual(f['filter']['Where'][0]['Condition']['In']['Values'][0][0]['Literal']['Value'], '225L')
        nb = b.categorical_filter({'type': 'NotBlank', 'field': {'table': 'T', 'name': 'Task Number'}}, 'k')
        self.assertEqual(nb['type'], 'Advanced')
        self.assertIn('Not', nb['filter']['Where'][0]['Condition'])

    def test_14_report_guard_sweeps_formatting_selectors(self):
        b = self.build()
        b.model_cols, b.measures = {'T': {'A'}}, {}
        vis = {'name': 'v', 'visual': {'visualType': 'tableEx', 'query': {'queryState': {'Values': {'projections': [
            {'field': {'Column': {'Expression': {'SourceRef': {'Entity': 'T'}}, 'Property': 'A'}}, 'queryRef': 'T.A'}]}}},
            'objects': {'columnWidth': [{'properties': {}, 'selector': {'metadata': 'T.A'}},
                                        {'properties': {}, 'selector': {'metadata': 'GONE.B'}}]}},
            'filterConfig': {'filters': [{'name': 'f', 'field': {'Column': {'Expression': {'SourceRef': {'Entity': 'T'}}, 'Property': 'A'}},
                                          'filter': {'From': [{'Name': 'f', 'Entity': 'GONE'}]}}]}}
        out = b.report_guard({'displayName': 'P'}, [vis])
        self.assertEqual([e['selector']['metadata'] for e in out[0]['visual']['objects']['columnWidth']], ['T.A'])
        self.assertEqual(out[0]['filterConfig']['filters'], [])

    def test_15_sheet_filters_without_quotes_ranges_and_relative_dates(self):
        """2026-11-04 corpus run: 147 sheet filters were never written (bare true/1, date ranges, relative dates),
        and ALL EXCEPT IN(...) produced garbled members."""
        r = self.rebuild([
            {'name': 'In Range', 'pbi_kind': 'calculated column', 'table': 'T', 'datatype': 'boolean',
             'formula': '[D] >= [Param:Start Date]', 'dax': "'T'[D] >= [Start Date Value]"},
            {'name': 'Year Flag', 'pbi_kind': 'calculated column', 'table': 'T', 'datatype': 'integer',
             'formula': 'IF YEAR([D]) = [Param:Year] THEN 1 ELSE 0 END', 'dax': "IF(YEAR('T'[D]) = [Year Value], 1, 0)"},
            {'name': 'Plain Flag', 'pbi_kind': 'calculated column', 'table': 'T', 'datatype': 'boolean',
             'formula': '[A] = "x"', 'dax': "'T'[A] = \"x\""}])
        spec = lambda field, rule: r.filter_spec({'field': field, 'rule': rule})
        # bare boolean on a parameter-driven row calc: folded into the measures as = TRUE(), not "Manual"
        rc = spec('In Range', 'true  [values:relevant]')
        self.assertEqual(rc['type'], 'RowCondition')
        self.assertEqual(rc['condition'], "('T'[D] >= [Start Date Value]) == TRUE()")
        self.assertEqual(spec('In Range', 'NOT true')['condition'], "NOT (('T'[D] >= [Start Date Value]) == TRUE())")
        # bug hunt 2026-11-07 #10: == so a blank row is not FALSE (Tableau drops NULL from a "False" filter)
        self.assertEqual(spec('In Range', 'false')['condition'], "('T'[D] >= [Start Date Value]) == FALSE()")
        # a number stays a number (was = "1": a text/number type error in DAX)
        self.assertEqual(spec('Year Flag', '1  [values:database]')['condition'], "IF(YEAR('T'[D]) = [Year Value], 1, 0) = 1")
        # a plain calc column / column filtered to a bare value: an ordinary categorical filter
        self.assertEqual(spec('Plain Flag', 'false  [values:database]')['values'], ['false'])
        self.assertEqual(spec('Task Number', '750  [values:relevant]')['values'], ['750'])
        # ALL EXCEPT IN(...): the members, not 'IN("a"'
        ex = spec('Met', 'NOT ALL EXCEPT IN("Other", "Withdrawn")  [values:database]')
        self.assertEqual((ex['type'], ex['values'], ex['inverted']), ('Categorical', ['Other', 'Withdrawn'], True))
        # quantitative and discrete ranges, relative dates
        rg = spec('Month Col', 'RANGE #2021-09-01 15:21:59#..+inf')
        self.assertEqual((rg['type'], rg['lo'], rg['hi']), ('Range', '#2021-09-01 15:21:59#', None))
        self.assertEqual(spec('Month Col', 'RANGE(1..4)  [values:relevant]')['values'], ['1', '2', '3', '4'])
        rel = spec('Month Col', 'RELATIVE first-period=-3 include-future=true include-null=false last-period=0 period-type-v2=year')
        self.assertEqual((rel['type'], rel['unit'], rel['first'], rel['last']), ('RelativeDate', 'year', -3, 0))
        self.assertNotIn('unit', spec('Month Col', 'RELATIVE first-period=-1 last-period=0 period-type-v2=quarter'))

        b = self.build({'T': {'name': 'T', 'columns': [{'name': 'D', 'datatype': 'datetime'}, {'name': 'N', 'datatype': 'integer'}],
                              'calc_columns': []}})
        b.renames, b.measures = {}, {'Total': {'table': 'T'}}
        where = lambda f: f['filter']['Where'][0]['Condition']
        f = b.categorical_filter({'type': 'Range', 'lo': '#2021-09-01 15:21:59#', 'hi': None, 'field': {'table': 'T', 'name': 'D'}}, 'k')
        self.assertEqual((f['type'], where(f)['Comparison']['ComparisonKind']), ('Advanced', 2))
        self.assertEqual(where(f)['Comparison']['Right']['Literal']['Value'], "datetime'2021-09-01T15:21:59'")
        f = b.categorical_filter({'type': 'Range', 'lo': '1', 'hi': '5', 'field': {'table': 'T', 'name': 'N'}}, 'k')
        self.assertEqual([where(f)['And'][s]['Comparison']['ComparisonKind'] for s in ('Left', 'Right')], [2, 4])  # >= and <=
        f = b.categorical_filter({'type': 'RelativeDate', 'unit': 'year', 'first': -3, 'last': 0, 'include_future': True,
                                  'field': {'table': 'T', 'name': 'D'}}, 'k')
        btw = where(f)['Between']
        self.assertEqual(f['type'], 'RelativeDate')
        self.assertEqual(btw['LowerBound'], {'DateSpan': {'Expression': {'DateAdd': {'Expression': {'Now': {}}, 'Amount': -3, 'TimeUnit': 3}}, 'TimeUnit': 3}})
        self.assertEqual(btw['UpperBound'], {'DateSpan': {'Expression': {'Now': {}}, 'TimeUnit': 3}})
        # a numeric filter on an aggregate calc: a visual-level measure comparison bound to the measure's table
        f = b.categorical_filter({'type': 'Categorical', 'values': ['1'], 'field': {'kind': 'Measure', 'table': None, 'name': 'Total'}}, 'k')
        self.assertIn('Measure', f['field'])
        self.assertEqual(f['filter']['From'][0]['Entity'], 'T')
        self.assertEqual(where(f)['Comparison']['Right']['Literal']['Value'], '1D')

    def test_16_joins_guessed_from_tableau_join_column_names(self):
        """Published source without .tds: Tableau renames the second copy of a join column '<name> (<table>) #n'."""
        import xml.etree.ElementTree as ET
        import bhayankar_build as bb
        from bhayankar_extractor import guess_joins
        rec = lambda fam, name: f'<metadata-record class="column"><remote-name>{name}</remote-name><family>{fam}</family></metadata-record>'
        ds = ET.fromstring('<datasource>' + ''.join([
            rec('WORK_ORDERS', 'WO Number'), rec('WORK_ORDERS', 'WO Status'), rec('WORK_ORDERS', 'Empl Num - PM (WORK_ORDERS)'),
            rec('TASKS', 'WO Number (Custom SQL Query) #2'), rec('TASKS', 'Task FWO'),
            rec('PM OWNER', 'Empl Num - PM (Custom SQL Query)'), rec('PM OWNER', 'Empl Name - PM'),
            rec('EMPLOYEE X', 'Empl Num-X'), rec('OTHER', 'Empl Num-X')]) + '</datasource>')
        got = {(g['many'], g['one'], g['physical_many']) for g in guess_joins(ds)}
        self.assertIn(('TASKS', 'WORK_ORDERS', 'WO Number'), got)          # "WO" is named after WORK_ORDERS
        self.assertIn(('WORK_ORDERS', 'PM OWNER', 'Empl Num - PM'), got)    # "PM" is named after PM OWNER
        self.assertFalse(any('EMPLOYEE X' in g for g in got))                    # same name, never renamed: no evidence
        self.assertTrue(bb.key_owner('EVENT_ID', 'Events'))
        self.assertFalse(bb.key_owner('INCIDENT_ID', 'Impacts'))
        self.assertTrue(bb.key_owner('WO Number', 'WORK_ORDERS'))
        self.assertFalse(bb.key_owner('WO Number', 'TASKS'))

    def test_17_both_directions_only_where_a_visual_needs_it_and_never_on_a_loop(self):
        b = self.build({'TASK': {'name': 'TASK', 'columns': [], 'calc_columns': []},
                        'WO': {'name': 'WO', 'columns': [], 'calc_columns': []}})
        b.relationships = [('TASK', 'WO Number', 'WO', 'WO Number')]
        b.measures = {'WO Count': {'dax': "DISTINCTCOUNT('WO'[WO Number])"}}
        b.renames, b.variants = {}, {}
        visual = lambda col_table: {'sheet': 'S', 'roles': {'Category': [{'kind': 'Column', 'table': col_table, 'name': 'X'}],
                                                             'Y': [{'kind': 'Measure', 'name': 'WO Count'}]}}
        b.spec = {'pages': [{'visuals': [visual('WO')]}]}
        b.filter_directions()
        self.assertEqual(b.both, set())                                  # one side groups its own measure: one direction
        b.spec = {'pages': [{'visuals': [visual('TASK')]}]}
        b.filter_directions()
        self.assertEqual(b.both, {b.relationships[0]})                    # WO count per task column: both directions
        b.relationships.append(('WO', 'K', 'TASK', 'K'))                  # a loop: never both directions
        b.filter_directions()
        self.assertEqual(b.both, set())

    def data_request_case(self):
        """DATA_REQUEST.md lists each table's CSV and headers and flags a missing header (header-only CSV: no rows)."""
        d = tempfile.mkdtemp()
        try:
            with open(os.path.join(d, 'ORDERS.csv'), 'w', encoding='utf-8') as f:
                f.write('ORDER_ID,STATUS\n')
            t = {'name': 'Orders', 'source': 'ORDERS', 'calc_columns': [], 'extra_renames': [],
                 'columns': [{'name': 'ORDER_ID', 'source_column': 'ORDER_ID'}, {'name': 'AMOUNT', 'source_column': 'AMOUNT'},
                             {'name': 'AMOUNT Sort', 'source_column': 'AMOUNT Sort', 'derived': True}]}
            b = self.build({'Orders': t}, {'mode': 'sample', 'sample_folder': d, 'tables': {}})
            b.digest, b.wb_out = {'groups': []}, d
            b.data_request()
            text = open(os.path.join(d, 'DATA_REQUEST.md'), encoding='utf-8').read()
            self.assertIn('**1 problem(s) found**', text)
            self.assertIn('| `AMOUNT` |  | **missing** |', text)
            self.assertNotIn('AMOUNT Sort', text)  # built by the model, not read from the source
            self.assertEqual(open(os.path.join(d, 'data_request', 'ORDERS.csv'), encoding='utf-8').read().strip(), 'ORDER_ID,AMOUNT')
        finally:
            shutil.rmtree(d, ignore_errors=True)

    def test_18_placeholders_that_have_a_translation(self):
        cols = {'[Task Number]': ('column', 'T', 'integer'), '[Fixed Labor Hours]': ('measure', None, 'real'),
                '[Event ID]': ('column', 'E', 'string')}
        res = lambda t: ({'kind': cols[t][0], 'table': cols[t][1], 'name': t.strip('[]'), 'datatype': cols[t][2]}
                         if t in cols else {'kind': 'unknown', 'name': t})
        tr = Translator(res)
        out = tr.translate('CASE [Task Number] when 90 then [Fixed Labor Hours] END', 'measure')
        self.assertEqual(out['dax'], "CALCULATE([Fixed Labor Hours], KEEPFILTERS('T'[Task Number] IN {90}))")
        out = tr.translate('IFNULL(ZN(LOOKUP(COUNT([Event ID]),0)),0)', 'measure')   # offset 0: the value itself
        self.assertNotIn('visual calc', out['dax'])
        self.assertEqual(out['kind'], 'measure')
        self.assertEqual(tr.translate('LOOKUP(COUNT([Event ID]),-1)', 'measure')['kind'], 'visual calculation')
        # dynamic title: a deleted field shows nothing; a caption differing in case finds its column
        r = self.rebuild()
        r.cols['Empl Name - Task COMP By'] = {'name': 'Empl Name - Task COMP By', 'table': 'T', 'datatype': 'string'}
        r.add_measure = lambda name, dax, *a, **k: r.measures.__setitem__(name, {'name': name, 'dax': dax})
        r.title({'title': 'Count <[Empl Name - Task COMP by]><[(deleted field Calc_1)]>'}, 'S')
        self.assertEqual(r.measures['S Title']['dax'], "\"Count \" & SELECTEDVALUE('T'[Empl Name - Task COMP By])")
        self.assertEqual(r.title({'title': '<[(deleted field Calc_1)]>'}, 'S2'), {'text': 'S2'})
        # quick table calc RUNNING_TOTAL, Table (across): a running sum along the Columns-shelf column; the colour
        # series partitions it by itself (filter context), so it is not addressed
        r = self.rebuild()
        r.measures['Count'] = {'name': 'Count', 'table': 'T', 'dax': 'X'}
        r.cols['Count'] = None
        r.column_ref = lambda n: {'kind': 'Measure', 'table': 'T', 'name': n} if n == 'Count' else {'kind': 'Column', 'table': 'T', 'name': n}
        r._ws = {'shelf': {'rows': [], 'cols': [{'label': 'Month Col', 'measure': False}]},
                 'panes': [{'enc': [{'label': 'Met', 'measure': False, 'channel': 'color'}]}]}
        self.data_request_case()
        fr = r.field('RUNNING_TOTAL(Count)', True, tc='Rows')
        self.assertIn("FILTER(ALLSELECTED('T'[Month Col]), 'T'[Month Col] <= _cur)", r.measures[fr['name']]['dax'])
        self.assertNotIn('Met', r.measures[fr['name']]['dax'])


class ExtractorCoverage(unittest.TestCase):
    """Tableau features the extractor did not read before 2026-10-30 (found by surveying 19 real workbooks).
    Fixture: tests/fixtures/features.twb (synthetic)."""

    @classmethod
    def setUpClass(cls):
        cls.res, cls.files, cls.spec, _ = run(path=os.path.join(FIX, 'features.twb'))
        cls.flags = {}
        for f in cls.res['flags']:
            cls.flags.setdefault(f['code'], []).append(f)
        cls.ws = {w['name']: w for w in cls.res['worksheets']}

    def test_parameter_action(self):
        a = next(a for a in self.res['actions'] if a['name'] == 'Set Region')
        self.assertEqual((a['kind'], a['parameter'], a['source_field']), ('PARAMETER', 'Param:Pick Region', 'Region'))
        self.assertIn('PARAMETER_ACTION', self.flags)
        notes = [n for p in self.spec['pages'] for n in p['notes'] if 'Set Region' in n]
        self.assertTrue(notes and 'Pick Region' in notes[0])  # sheet-level action still reaches its page

    def test_nav_action(self):
        a = next(a for a in self.res['actions'] if a['name'] == 'Go to Detail')
        self.assertEqual((a['kind'], a['target']), ('NAVIGATE', 'Detail'))

    def test_dynamic_zone_visibility_and_hidden_zone(self):
        items = {i['name']: i for i in next(d for d in self.res['dashboards'] if d['name'] == 'Overview')['items']}
        self.assertEqual(items['Detail']['shown_when'], 'Param:Pick Region')
        self.assertTrue(items['Help']['hidden'])
        self.assertIn('DYNAMIC_ZONE_VISIBILITY', self.flags)
        self.assertIn('HIDDEN_ZONE', self.flags)

    def test_story(self):
        st = next(d for d in self.res['dashboards'] if d['name'] == 'Story 1')['story']
        self.assertEqual([s['sheet'] for s in st], ['Overview', 'Detail'])

    def test_hierarchy_folder_sort_geo(self):
        self.assertEqual(self.res['hierarchies'][0]['levels'], ['Country', 'Region'])
        cols = {c['name']: c for c in self.res['columns']}
        self.assertEqual(cols['Priority']['sort_order'], ['High', 'Medium', 'Low'])
        self.assertEqual(cols['Priority']['folder'], 'Order Info')
        self.assertEqual(cols['Region']['geo_role'], '[State].[Name]')
        self.assertEqual(calc(self.res, 'Sum of Order Totals')['folder'], 'Order Info')

    def test_sheet_details(self):
        ws = self.ws['Sales by Region']
        self.assertIn('[Region] DESC by [SUM(Amount)]', ws['sorts'])  # shelf-sort-v2
        self.assertEqual(ws['subtotals'], ['Region'])
        self.assertEqual(ws['panes'][0]['trendline'], 'linear')
        rt = next(f for f in ws['shelf']['cols'] if f['label'].startswith('RUNNING_TOTAL'))
        self.assertEqual((rt['tc_ordering'], rt['tc_fields']), ('Field', ['Priority', 'Region']))

    def test_sheet_roles(self):
        roles = {n: w['role'] for n, w in self.ws.items()}
        self.assertEqual(roles, {'Sales by Region': 'data', 'Help': 'help', 'Link Shape': 'button', 'Detail': 'data'})
        self.assertTrue(calc(self.res, '" "')['placeholder'])

    def test_joined_grain(self):
        self.assertEqual(calc(self.res, 'Sum of Order Totals')['joined_grain']['fixed'], ['Order Total'])
        self.assertEqual(calc(self.res, 'Avg Days to Ship')['joined_grain']['tables'], ['LINES', 'ORDERS'])
        self.assertNotIn('joined_grain', calc(self.res, 'Order Total'))
        self.assertIn('PHYSICAL_JOIN', self.flags)

    def test_unhandled_xml_is_reported(self):
        details = ' '.join(f['detail'] for f in self.flags['UNHANDLED_XML'])
        self.assertIn('"future-feature"', details)
        self.assertIn('"hologram"', details)

    def test_digest_text(self):
        md = self.files['bhayankar_digest.md']
        for s in ('## Hierarchies', 'compute using Priority, Region', 'SHOWN WHEN [Param:Pick Region]',
                  'STORY POINTS', '(help sheet: no data visual)', 'subtotals by: Region'):
            self.assertIn(s, md)

    def test_home_table_is_deterministic(self):
        # cross-table calc: the first field in the formula decides, whatever the set order is
        import subprocess
        code = ('import bhayankar_extractor as bx;r,_=bx.load_root(r"%s");w=bx.Workbook(r);'
                'print(w.calc_table(("federated.ord","Calculation_3")))' % os.path.join(FIX, 'features.twb'))
        seen = {subprocess.run([sys.executable, '-c', code], cwd=os.path.dirname(HERE), capture_output=True, text=True,
                               env={**os.environ, 'PYTHONHASHSEED': str(s)}).stdout.strip() for s in range(6)}
        self.assertEqual(seen, {'ORDERS'})

    def test_extract_without_hyper_package(self):
        # a .twbx with a .hyper: listed with a reason when tableauhyperapi is missing (never crashes)
        import zipfile
        tmp = tempfile.mkdtemp(prefix='bhayankar_test_')
        try:
            p = os.path.join(tmp, 'x.twbx')
            with zipfile.ZipFile(p, 'w') as z:
                z.write(os.path.join(FIX, 'features.twb'), 'x.twb')
                z.writestr('Data/Extracts/x.hyper', b'not a real extract')
            saved = sys.modules.get('tableauhyperapi')
            sys.modules['tableauhyperapi'] = None  # import fails
            try:
                ext = bx.extract_tables(p)
            finally:
                if saved is None:
                    sys.modules.pop('tableauhyperapi', None)
                else:
                    sys.modules['tableauhyperapi'] = saved
            self.assertEqual(ext[0]['file'], 'Data/Extracts/x.hyper')
            self.assertIn('pip install tableauhyperapi', ext[0]['status'])
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


class DesignReviewRerunFixes(unittest.TestCase):
    """Design Review 1X1 re-run (2026-10-31): the TWB was read independently and compared with the
    digest, the rebuild spec and the built project. One test per bug (number = docs\\PROJECT_LOG.md)."""

    @staticmethod
    def res():
        f = lambda lab, m=False: {'label': lab, 'measure': m, 'continuous': m, 'date': False}
        ws = lambda name, rows, cols, enc, mark='Automatic', **kw: {
            'name': name, 'rows': '', 'cols': '', 'shelf': {'rows': rows, 'cols': cols, 'dual': None},
            'panes': [{'mark': mark, 'enc': enc}], 'filters': kw.pop('filters', []), **kw}
        return {
            'workbook': {}, 'datasources': [{'name': 'DS', 'filters': [], 'connection': {}}], 'parameters': [], 'groups': [], 'palettes': {},
            'columns': [{'name': n, 'table': 'W', 'datatype': 'string', 'source_column': n} for n in ('Job Plan', 'Met', 'Region Group')],
            'calculations': [{'name': 'WO Count', 'pbi_kind': 'measure', 'formula': 'COUNTD([WO Number])', 'datatype': 'integer',
                              'dax': "DISTINCTCOUNT('W'[WO Number])", 'table': None, 'confidence': 'high', 'dax_notes': []}],
            'worksheets': [
                ws('Bars', [f('Job Plan')], [f('WO Count', True)], [], 'Bar', sorts=['[Job Plan] DESC by [WO Count]'],
                   filters=[{'field': 'Region Group', 'rule': 'ALL (interactive)'}]),
                ws('Tiles', [], [f('Met')], [{**f('WO Count', True), 'channel': 'text'}]),
                ws('Detail Table', [f('Job Plan'), f('Met')], [], [{**f('WO Count', True), 'channel': 'text'}])],
            'dashboards': [  # XML order: the detail dashboard first
                {'name': 'Detail', 'px': [1000, 600], 'sizing': 'fixed', 'items': [
                    {'type': 'text', 'name': None, 'field': None, 'mode': None, 'box': [0, 0, 1000, 40], 'show_title': True,
                     'text': '<Sheet Name>'},
                    {'type': 'sheet', 'name': 'Detail Table', 'box': [0, 40, 1000, 560], 'show_title': False}]},
                {'name': 'Main', 'px': [1000, 600], 'sizing': 'fixed', 'items': [
                    {'type': 'sheet', 'name': 'Bars', 'box': [0, 80, 500, 500], 'show_title': True},
                    {'type': 'sheet', 'name': 'Tiles', 'box': [500, 80, 500, 60], 'show_title': True},
                    {'type': 'filter', 'name': 'Bars', 'field': 'Region Group', 'mode': 'checkdropdown', 'box': [0, 0, 200, 60],
                     'show_title': True}]}],
            'tabs': [{'name': 'Main', 'class': 'dashboard', 'hidden': False}, {'name': 'Detail', 'class': 'dashboard', 'hidden': False}],
            'actions': [{'name': 'To detail', 'kind': 'FILTER', 'source': 'Main / Tiles', 'target': 'Detail', 'trigger': 'on-select'}]}

    @classmethod
    def setUpClass(cls):
        from bhayankar_report import Rebuild
        cls.spec = Rebuild(cls.res(), catalog=None).build()

    def test_01_pages_follow_the_tableau_tab_order(self):
        self.assertEqual([p['name'] for p in self.spec['pages']], ['Main', 'Detail'])

    def test_02_dashboard_text_sheet_name_is_the_dashboard_name(self):
        self.assertEqual(page(self.spec, 'Detail')['static'][0]['text'], 'Detail')

    def test_03_single_sheet_action_takes_only_that_sheets_fields(self):
        self.assertEqual([f['name'] for f in page(self.spec, 'Detail')['drillthrough']['fields']], ['Met'])

    def test_04_text_table_columns_shelf_runs_across(self):
        v = visual(self.spec, 'Tiles')
        self.assertEqual((v['visualType'], [f['name'] for f in v['roles']['Columns']]), ('pivotTable', ['Met']))
        self.assertNotIn('Rows', v['roles'])
        self.assertEqual([f['name'] for f in visual(self.spec, 'Detail Table')['roles']['Rows']], ['Job Plan', 'Met'])

    def test_05_computed_sort_is_kept(self):
        self.assertEqual(visual(self.spec, 'Bars')['sort_by'], {'measure': 'WO Count', 'direction': 'Descending'})
        import bhayankar_build as bb
        b = bb.Build.__new__(bb.Build)
        b.log, b.renames, b.variants = [], {}, {}
        b.measures = {'WO Count': {'table': 'W'}}
        v = {**visual(self.spec, 'Bars'), 'position': {'x': 0, 'y': 0, 'width': 400, 'height': 300}}
        out = b.visual({'name': 'Main'}, v, [])
        srt = out['visual']['query']['sortDefinition']['sort'][0]
        self.assertEqual((srt['field']['Measure']['Property'], srt['direction']), ('WO Count', 'Descending'))

    def test_06_quick_table_calc_keeps_the_tableau_number_format(self):
        r = DesignReviewGeneratorFixes.rebuild(DesignReviewGeneratorFixes())
        r.measures['Task Count'] = {'name': 'Task Count', 'table': 'T', 'dax': 'X'}
        r.column_ref = lambda n: {'kind': 'Measure', 'table': 'T', 'name': n} if n == 'Task Count' else \
            {'kind': 'Column', 'table': 'T', 'name': n}
        r.formats = {'PERCENT_OF_TOTAL(Task Count)': 'p0%'}
        r._ws = {'shelf': {'rows': [], 'cols': [{'label': 'Met', 'measure': False}]}, 'panes': [{'enc': []}]}
        fr = r.field('PERCENT_OF_TOTAL(Task Count)', True, tc='Rows')
        self.assertEqual(r.measures[fr['name']]['formatString'], '0%')

    def test_07_quick_filter_reaches_only_its_sheets(self):
        main = page(self.spec, 'Main')
        rg = next(s for s in main['slicers'] if s['field']['name'] == 'Region Group')
        self.assertEqual(rg['scope_sheets'], ['Bars'])
        self.assertFalse([s for s in page(self.spec, 'Detail')['slicers'] if s['field']['name'] == 'Region Group'])
        import bhayankar_build as bb   # the build turns the scope into "no filter" interactions
        b = bb.Build.__new__(bb.Build)
        b.log, b.growth = [], []
        b.slicer = lambda p, s: {'name': 'slicer1', 'position': {}, 'visual': {}}
        b.visual = lambda p, v, tp: {'name': 'vis-' + v['sheet'], 'position': {}, 'visual': {}}
        b.static = lambda p, st, report: None
        written = {}
        b.write_page = lambda d, pg, vis: written.setdefault(pg['displayName'], pg)
        b.scaffold = lambda: tempfile.mkdtemp(prefix='bhayankar_pages_')
        b.report_dir = tempfile.mkdtemp(prefix='bhayankar_rep_')
        os.makedirs(os.path.join(b.report_dir, 'definition'))
        try:
            with open(os.path.join(b.report_dir, 'definition', 'report.json'), 'w') as fh:
                fh.write('{}')
            pages_dir = b.scaffold()
            with open(os.path.join(pages_dir, 'pages.json'), 'w') as fh:
                fh.write('{}')
            b.scaffold = lambda: pages_dir
            b.spec = {'pages': [main]}
            b.write_report()
        finally:
            shutil.rmtree(b.report_dir, ignore_errors=True)
        vi = written['Main']['visualInteractions']
        self.assertIn({'source': 'slicer1', 'target': 'vis-Tiles', 'type': 'NoFilter'}, vi)
        self.assertNotIn({'source': 'slicer1', 'target': 'vis-Bars', 'type': 'NoFilter'}, vi)

    def test_08_row_calc_shown_as_a_dimension_is_a_column(self):
        wb = bx.Workbook.__new__(bx.Workbook)
        wb.inst = {('ds', 'none:Days:ok'): ('Days', None, 'ordinal'), ('ds', 'sum:Hours:qk'): ('Hours', 'Sum', 'quantitative')}
        self.assertEqual(wb.view_role(('ds', 'Days'), 'measure'), 'dimension')     # a column of a detail table
        self.assertEqual(wb.view_role(('ds', 'Hours'), 'measure'), 'measure')      # only ever summed
        self.assertEqual(wb.view_role(('ds', 'Unused'), 'measure'), 'measure')
        tr = DesignReviewGeneratorFixes.tr(DesignReviewGeneratorFixes())
        self.assertEqual(tr.translate("DATEDIFF('day', [D], [E])", 'dimension')['kind'], 'calculated column')

    def test_09_blank_boolean_is_not_false(self):
        # checked in the engine 2026-10-31: a missing RELATED() boolean is blank; BLANK() = FALSE() is TRUE,
        # BLANK() == FALSE() is FALSE (Tableau: NULL = FALSE is NULL, so IF takes ELSE)
        tr = DesignReviewGeneratorFixes.tr(DesignReviewGeneratorFixes())
        out = tr.translate('IF [U] = FALSE AND [U] <> TRUE THEN "a" END', 'dimension')['dax']
        self.assertIn("'W'[U] == FALSE()", out)
        self.assertIn("(NOT ISBLANK('W'[U]) && 'W'[U] <> TRUE())", out)

    def test_10_parameter_table_name_is_not_doubled(self):
        from bhayankar_dax import param_table_name, parameter_pattern
        self.assertEqual(param_table_name('DC Parameter'), 'DC Parameter')
        self.assertEqual(param_table_name('Start Date'), 'Start Date Parameter')
        pp = parameter_pattern({'name': 'DC Parameter', 'datatype': 'string', 'value': '"a"', 'members': [['"a"', None]]})
        self.assertNotIn('Parameter Parameter', json.dumps(pp))

    def test_11_model_columns_keep_the_source_name(self):
        import bhayankar_build as bb
        b = DesignReviewGeneratorFixes.build(DesignReviewGeneratorFixes(), {'W': {'name': 'W', 'source': 'W', 'steps': [], 'calc_columns': [
            {'name': 'Met', 'dax': "IF('W'[Ready Date] > 1, 1)"}], 'columns': [
            {'name': 'Ready Date', 'source_column': 'WO Const Ready Date', 'datatype': 'datetime'},
            {'name': 'WO Number', 'source_column': 'WO Number', 'datatype': 'string'}]}})
        field = {'kind': 'Column', 'table': 'W', 'name': 'Ready Date', 'label': 'Ready Date'}
        b.spec = {'model': {'tables': [{'name': 'W', 'measures': [{'name': 'M', 'table': 'W', 'dax': "MAX('W'[Ready Date])"}],
                                        'calculated_columns': []}]},
                  'pages': [{'visuals': [{'roles': {'Rows': [field]}}]}]}
        b.digest = {'calculations': []}
        self.assertTrue(b.keep_source_names(b.tables))
        b.apply_remaps(b.spec['model'])
        self.assertEqual([c['name'] for c in b.tables['W']['columns']], ['WO Const Ready Date', 'WO Number'])
        self.assertEqual(b.spec['model']['tables'][0]['measures'][0]['dax'], "MAX('W'[WO Const Ready Date])")
        self.assertEqual(b.tables['W']['calc_columns'][0]['dax'], "IF('W'[WO Const Ready Date] > 1, 1)")
        self.assertEqual((field['name'], field['label']), ('WO Const Ready Date', 'Ready Date'))  # caption = header
        proj = bb.proj(bb.col_ref('W', field['name']), field['label'])
        self.assertEqual(proj['displayName'], 'Ready Date')
        # sources.json written before the rename names the column by its caption: no rename step back to the caption
        # (the DR refresh failed "column 'WO Const Ready Date' not found" before sample_renames knew the caption)
        cfg = {'file': 'w.csv', 'columns': {'Ready Date': 'WO Const Ready Date', 'WO Number': 'CSV WO'}}
        self.assertEqual(b.sample_renames(b.tables['W'], cfg), [('CSV WO', 'WO Number')])


@unittest.skipUnless(os.environ.get('BHAYANKAR_CORPUS'), 'set BHAYANKAR_CORPUS to run the corpus smoke test')
class Corpus(unittest.TestCase):
    def test_corpus(self):
        root = os.environ['BHAYANKAR_CORPUS']
        bad = []
        for fn in sorted(os.listdir(root)):
            if not fn.lower().endswith(('.twb', '.twbx')):
                continue
            try:
                _, _, spec, _ = run(path=os.path.join(root, fn))
                errs = [e for p in spec['pages'] for v in p['visuals'] for e in v.get('catalog_errors', [])]
                print(f'{fn}: {sum(len(p["visuals"]) for p in spec["pages"])} visuals, {len(errs)} catalog errors')
            except Exception as e:  # noqa: BLE001
                bad.append(f'{fn}: {e!r}')
        self.assertEqual(bad, [])


if __name__ == '__main__':
    unittest.main()
