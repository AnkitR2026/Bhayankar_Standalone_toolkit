"""Regression tests for the bug hunt of 2026-11-02: 17 real workbooks run through the Kiro toolkit.
One test (or more) per bug; the number is the one in docs\\PROJECT_LOG.md ("Bug hunt 2026-11-02").

Run (from bhayankar):  python -m unittest discover -s tests -v
"""
import copy
import json
import os
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, HERE)
import bhayankar_build as bb  # noqa: E402
from bhayankar_dax import Translator  # noqa: E402
import test_bhayankar as tb  # noqa: E402  (module import: its test classes are not collected twice)


def translator(extra=None):
    cols = {'[A]': ('column', 'T', 'string'), '[D]': ('column', 'T', 'datetime'), '[K]': ('column', 'T', 'string'),
            '[N]': ('column', 'T', 'real'), **(extra or {})}

    def res(t):
        if t not in cols:
            return {'kind': 'unknown', 'name': t}
        kind, table, dt = cols[t][:3]
        return {'kind': kind, 'table': table, 'name': t.strip('[]'), 'datatype': dt,
                'row_dax': cols[t][3] if len(cols[t]) > 3 else None}
    return Translator(res)


def guard_build(measures, param_rowcalcs=()):
    """A Build with one table F (columns A, K) and the given measures, ready for load_guard()."""
    b = bb.Build.__new__(bb.Build)
    b.log, b.decisions, b.report_period, b.slice_tables, b.param_tables, b.pdim_tables = [], {}, None, {}, {}, {}
    b.digest = {'flags': []}
    b.moved_tables, b.moved_cols = {}, {}
    b.param_rowcalcs = list(param_rowcalcs)
    b.tables = {'F': {'name': 'F', 'columns': [{'name': 'A'}, {'name': 'K'}], 'calc_columns': []}}
    b.measures = {n: {'name': n, 'dax': dax, 'table': 'F', 'description': '', 'displayFolder': 'Measures',
                      'formatString': None, 'hidden': False, 'notes': [], 'confidence': 'high'} for n, dax in measures.items()}
    b.load_guard()
    return b


class SilentWrongNumbers(unittest.TestCase):
    """Bugs 1-5: the report opens and shows a number, but not Tableau's number."""

    def test_01_measures_reading_a_placeholder_are_flagged(self):
        b = guard_build({'Ok': "COUNTROWS('F')", 'Broken': '[Nope] + 1', 'Total': '[Ok] + [Broken]', 'Ratio': 'DIVIDE([Total], [Ok])'})
        self.assertEqual(b.measures['Broken']['dax'], 'BLANK()')
        for n in ('Total', 'Ratio'):  # BLANK() + x = x: these looked fine before
            self.assertEqual(b.measures[n]['displayFolder'], 'Measures\\_Needs fix', n)
            self.assertEqual(b.measures[n]['confidence'], 'low', n)
            self.assertIn('WRONG UNTIL FIXED', b.measures[n]['description'], n)
        self.assertNotEqual(b.measures['Total']['dax'], 'BLANK()')  # kept, so the developer sees the formula
        self.assertEqual(b.measures['Ok']['displayFolder'], 'Measures')
        self.assertTrue(any('DECISION NEEDED' in ln and '"Total"' in ln for ln in b.log))

    def test_02_countd_ignores_blank_like_tableau(self):
        tr = translator()
        self.assertEqual(tr.translate('COUNTD([A])', 'measure')['dax'], "DISTINCTCOUNTNOBLANK('T'[A])")
        out = tr.translate('COUNTD(IF [K] = "x" THEN [A] END)', 'measure')['dax']
        self.assertTrue(out.startswith("COUNTROWS(FILTER(DISTINCT(SELECTCOLUMNS('T', \"@v\", "), out)
        self.assertTrue(out.endswith('NOT ISBLANK([@v])))'), out)
        # and the load guard keeps it (a "[v]" helper column was taken for a missing measure and stubbed)
        b = guard_build({'C': out.replace("'T'", "'F'")})
        self.assertNotEqual(b.measures['C']['dax'], 'BLANK()')

    def test_03_parameter_row_calc_is_inlined_into_measures(self):
        calc = {'name': 'Before Start?', 'dax': "'F'[A] < [Start Value]"}
        b = guard_build({'Start Value': '1', 'Cnt': "COUNTROWS(FILTER('F', 'F'[Before Start?] = TRUE()))"}, [('F', calc)])
        self.assertEqual(b.measures['Cnt']['dax'], "COUNTROWS(FILTER('F', ('F'[A] < [Start Value]) = TRUE()))")

    def test_04_tables_differing_only_in_case_are_one_table(self):
        b = bb.Build.__new__(bb.Build)
        b.log, b.decisions, b.wb = [], {}, 'WB'
        b.sources = {'mode': 'warehouse', 'tables': {}}
        b.digest = {'groups': [], 'calculations': [], 'parameters': [], 'flags': []}
        col = lambda n: {'name': n, 'source_column': n, 'datatype': 'string'}
        b.spec = {'pages': [], 'model': {'relationships': [], 'url_columns': [], 'parameter_dimensions': [], 'tables': [
            {'name': 'PM OWNER', 'source_table': 'PM OWNER', 'columns': [col('Dept'), col('Empl Name')],
             'calculated_columns': [], 'measures': []},
            {'name': 'PM Owner', 'source_table': 'PM Owner', 'columns': [col('Empl Name'), col('Extra')],
             'calculated_columns': [], 'measures': []}]}}
        b.tds_relations = lambda: {}
        b.build_measures = lambda: setattr(b, 'measures', {})
        b.slice_tables, b.param_tables, b.pdim_tables, b.report_period = {}, {}, {}, None
        b.build_model()
        self.assertEqual(list(b.tables), ['PM OWNER'])
        self.assertEqual([c['name'] for c in b.tables['PM OWNER']['columns']], ['Dept', 'Empl Name', 'Extra'])
        self.assertEqual(b.remap('PM Owner', 'Empl Name'), ('PM OWNER', 'Empl Name'))

    def test_05_fixed_lod_without_dimensions_has_a_row_value(self):
        out = translator().translate('{MAX([D])}', 'measure')
        self.assertEqual(out['row_dax'], "CALCULATE(MAX('T'[D]), REMOVEFILTERS('T'))")
        # a row calc over it becomes a calculated column (was: "not valid DAX" -> BLANK() measure)
        tr = translator({'[Latest]': ('measure', None, 'datetime', out['row_dax'])})
        row = tr.translate('YEAR([D]) = YEAR([Latest]) AND [D] <= [Latest]', 'dimension')
        self.assertEqual(row['kind'], 'calculated column')
        self.assertIn("CALCULATE(MAX('T'[D]), REMOVEFILTERS('T'))", row['dax'])


class BuildFailures(unittest.TestCase):
    """Bugs 6-10: the build crashed or the project failed validation."""

    def test_06_table_file_names_are_safe(self):
        self.assertEqual(bb.tmdl_file('Fame/Shame Index Parameter'), 'Fame%2FShame Index Parameter.tmdl')
        self.assertEqual(bb.tmdl_file('Divested? Parameter'), 'Divested%3F Parameter.tmdl')
        self.assertEqual(bb.tmdl_file('a:b*c"d<e>f|g\\h'), 'a%3Ab%2Ac%22d%3Ce%3Ef%7Cg%5Ch.tmdl')
        self.assertEqual(bb.tmdl_file('Ends with dot.'), 'Ends with dot%2E.tmdl')
        self.assertEqual(bb.tmdl_file('Orders'), 'Orders.tmdl')

    def help_build(self):
        b = bb.Build.__new__(bb.Build)
        b.log, b.renames, b.variants = [], {}, {}
        b.measures = {'Help Icon': {'table': 'F'}}
        b.rename = lambda n: n
        return b

    def test_07_sheet_on_two_dashboards_has_one_tooltip_page(self):
        b, tips = self.help_build(), []
        v = {'id': 'v1', 'sheet': 'Help', 'visualType': 'help-icon', 'tooltip_text': 'How to read this page',
             'position': {'x': 0, 'y': 0, 'width': 30, 'height': 30}}
        b.visual({'name': 'Summary'}, v, tips)
        b.visual({'name': 'Details'}, dict(v), tips)
        self.assertEqual(len(tips), 1)

    def test_08_workbook_without_dashboards_gets_one_page(self):
        root = tempfile.mkdtemp(prefix='bhayankar_bug08_')
        try:
            b = bb.Build.__new__(bb.Build)
            b.log, b.spec, b.images = [], {'pages': []}, {}
            b.report_dir = os.path.join(root, 'X.Report')
            pages_dir = os.path.join(b.report_dir, 'definition', 'pages')
            os.makedirs(pages_dir)
            with open(os.path.join(b.report_dir, 'definition', 'report.json'), 'w') as f:
                json.dump({}, f)
            with open(os.path.join(pages_dir, 'pages.json'), 'w') as f:
                json.dump({'pageOrder': []}, f)
            b.scaffold = lambda: pages_dir
            b.write_report()
            with open(os.path.join(pages_dir, 'pages.json')) as f:
                order = json.load(f)['pageOrder']
            self.assertEqual(len(order), 1)
            self.assertTrue(os.path.exists(os.path.join(pages_dir, order[0], 'page.json')))
        finally:
            shutil.rmtree(root, ignore_errors=True)

    def test_09_date_table_has_every_date_part_column(self):
        b = bb.Build.__new__(bb.Build)
        b.digest, b.measures = {'parameters': []}, {}
        tmdl = b.date_tmdl('Date', True)
        for c in ('Week Start', 'Week', 'Year Quarter', 'Day of Week', 'Month Day'):
            self.assertIn(f"column '{c}'", tmdl, c)
        self.assertIn('Date.StartOfWeek([Date], Day.Sunday)', tmdl)
        self.assertIn("sortByColumn: 'Year Quarter Number'", tmdl)
        from bhayankar_report import DATE_COL
        missing = set(DATE_COL.values()) - {'Hour'} - {c.split("'")[1] for c in tmdl.split('\n') if c.startswith('\tcolumn ')}
        self.assertEqual(missing, set())

    def test_10_chart_without_its_axis_becomes_a_card(self):
        b = bb.Build.__new__(bb.Build)
        b.log = []
        y = bb.proj(bb.meas_ref('F', 'Sales'))
        v = {'name': 'v1', 'visual': {'visualType': 'lineChart', 'query': {'queryState': {'Y': {'projections': [y]}}}}}
        self.assertTrue(b.required_roles_ok('P', v))
        self.assertEqual(v['visual']['visualType'], 'cardVisual')
        self.assertEqual(v['visual']['query']['queryState'], {'Data': {'projections': [y]}})
        self.assertIn('DECISION NEEDED', b.log[-1])
        empty = {'name': 'v2', 'visual': {'visualType': 'lineChart', 'query': {'queryState': {'Tooltips': {'projections': []}}}}}
        self.assertFalse(b.required_roles_ok('P', empty))
        ok = {'name': 'v3', 'visual': {'visualType': 'lineChart', 'query': {'queryState': {
            'Category': {'projections': [bb.proj(bb.col_ref('F', 'A'))]}, 'Y': {'projections': [y]}}}}}
        self.assertTrue(b.required_roles_ok('P', ok))
        self.assertEqual(ok['visual']['visualType'], 'lineChart')


class Behaviour(unittest.TestCase):
    """Bugs 11-16."""

    @classmethod
    def setUpClass(cls):
        from bhayankar_report import Rebuild
        res = tb.DesignReviewRerunFixes.res()
        f = lambda lab, m=False: {'label': lab, 'measure': m, 'continuous': m, 'date': False}
        res['worksheets'].append({'name': 'Total KPI', 'rows': '', 'cols': '', 'shelf': {'rows': [], 'cols': [], 'dual': None},
                                  'panes': [{'mark': 'Automatic', 'enc': [{**f('WO Count', True), 'channel': 'text'}]}], 'filters': []})
        res['dashboards'][1]['items'].append({'type': 'sheet', 'name': 'Total KPI', 'box': [500, 200, 300, 100], 'show_title': True})
        res['actions'] += [{'name': 'KPI to detail', 'kind': 'FILTER', 'source': 'Main / Total KPI', 'target': 'Detail',
                            'trigger': 'on-select', 'special-fields': 'all'}]
        cls.spec = Rebuild(res, catalog=None).build()
        res2 = copy.deepcopy(tb.DesignReviewRerunFixes.res())
        res2['actions'] = [{'name': 'Sheet level', 'kind': 'FILTER', 'source': ' / Tiles', 'target': 'Detail', 'trigger': 'on-select'}]
        cls.spec2 = Rebuild(res2, catalog=None).build()

    def test_11_action_from_a_kpi_sheet_is_navigation(self):
        main = tb.page(self.spec, 'Main')
        self.assertIn({'sheet': 'Total KPI', 'target': 'Detail', 'action': 'KPI to detail'}, main.get('nav_buttons', []))
        self.assertIn('Main', tb.page(self.spec, 'Detail')['nav_back'])
        self.assertEqual([f['name'] for f in tb.page(self.spec, 'Detail')['drillthrough']['fields']], ['Met'])  # Tiles still drills

    def test_11b_sheet_level_filter_action_finds_its_page(self):
        self.assertEqual([f['name'] for f in tb.page(self.spec2, 'Detail')['drillthrough']['fields']], ['Met'])

    @unittest.skipUnless(os.path.exists(os.path.join(os.path.dirname(HERE), 'bhayankar_mcp.py')), 'no MCP server in this edition')
    def test_13_mcp_rejects_paths_in_the_workbook_name(self):
        import bhayankar_mcp as mcp
        for bad in ('../../outside', '..\\x', 'C:\\Windows', '..', '', None, 'a/b'):
            self.assertFalse(mcp.known(bad), bad)
        with self.assertRaises(ValueError):
            mcp.call('bhayankar_status', {})

    @unittest.skipUnless(os.path.exists(os.path.join(os.path.dirname(HERE), 'bhayankar_mcp.py')), 'no MCP server in this edition')
    def test_13b_mcp_bad_part_is_a_plain_error(self):
        import bhayankar_mcp as mcp
        old = mcp.known, mcp.digest_text
        try:
            mcp.known, mcp.digest_text = (lambda n: True), (lambda n: '## Flags\nx')
            with self.assertRaises(ValueError):
                mcp.call('bhayankar_read_digest', {'workbook': 'W', 'part': 'bogus'})
        finally:
            mcp.known, mcp.digest_text = old

    def test_16_force_without_a_name_reruns_done_workbooks(self):
        import bhayankar_run as run
        old = run.scan, run.run_one
        try:
            run.scan = lambda: [{'name': 'A', 'state': 'done'}, {'name': 'B', 'state': 'new'}]
            run.run_one = lambda n, **kw: n
            self.assertEqual(run.run_all(force=True), ['A', 'B'])
            self.assertEqual(run.run_all(), ['B'])
        finally:
            run.scan, run.run_one = old


class SecondPass(unittest.TestCase):
    """Bugs 18-26: found once the workbooks that used to crash were built (Superstore 2026, Retail, Super Sample)."""

    def test_18_calc_reading_a_parameter_calc_is_inlined_too(self):
        b = bb.Build.__new__(bb.Build)
        b.log, b.decisions, b.report_period, b.slice_tables, b.param_tables, b.pdim_tables = [], {}, None, {}, {}, {}
        b.digest, b.moved_tables, b.moved_cols = {'flags': []}, {}, {}
        b.param_rowcalcs = [('F', {'name': 'Sel', 'dax': "'F'[A] = [P Value]"})]
        b.tables = {'F': {'name': 'F', 'columns': [{'name': 'A'}, {'name': 'K'}],
                          'calc_columns': [{'name': 'Filter', 'dax': "'F'[Sel] && 'F'[K] <> \"x\""}]}}
        b.measures = {n: {'name': n, 'dax': d, 'table': 'F', 'description': '', 'displayFolder': 'Measures', 'formatString': None,
                          'hidden': False, 'notes': [], 'confidence': 'high'}
                      for n, d in {'P Value': '1', 'N': "CALCULATE(COUNTROWS('F'), FILTER('F', 'F'[Filter]))"}.items()}
        b.load_guard()
        self.assertEqual(b.tables['F']['calc_columns'], [])  # not a column: it could never follow the slicer
        self.assertEqual(b.measures['N']['dax'], "CALCULATE(COUNTROWS('F'), FILTER('F', (('F'[A] = [P Value]) && 'F'[K] <> \"x\")))")

    def test_19_attr_of_an_inlined_calc_is_valid_dax(self):
        calc = {'name': 'Sel State', 'dax': "IF('F'[A] = [P Value], 'F'[K])"}
        b = guard_build({'P Value': '1', 'One': "SELECTEDVALUE('F'[Sel State])", 'Cnt': "DISTINCTCOUNTNOBLANK('F'[Sel State])"},
                        [('F', calc)])
        self.assertEqual(b.measures['One']['dax'], '(VAR _one = DISTINCT(SELECTCOLUMNS(\'F\', "@x", IF(\'F\'[A] = [P Value], '
                                                   '\'F\'[K]))) RETURN IF(COUNTROWS(_one) = 1, MAXX(_one, [@x])))')
        self.assertTrue(b.measures['Cnt']['dax'].startswith('COUNTROWS(FILTER(DISTINCT(SELECTCOLUMNS('))
        self.assertNotEqual(b.measures['One']['dax'], 'BLANK()')

    def test_20_colour_rule_on_a_missing_column_is_dropped(self):
        b = guard_build({})
        sel = {'data': [{'scopeId': {'Comparison': {'ComparisonKind': 0, 'Left': bb.col_ref('F', 'Gone'),
                                                     'Right': {'Literal': {'Value': "'x'"}}}}}]}
        v = {'name': 'v', 'visual': {'visualType': 'barChart', 'query': {'queryState': {
            'Category': {'projections': [bb.proj(bb.col_ref('F', 'A'))]}, 'Y': {'projections': [bb.proj(bb.meas_ref('F', 'M'))]}}},
            'objects': {'dataPoint': [{'properties': {}, 'selector': sel}]}}}
        b.measures['M'] = {'name': 'M', 'table': 'F', 'dax': '1'}
        out = b.report_guard({'displayName': 'P'}, [v])
        self.assertNotIn('dataPoint', out[0]['visual'].get('objects', {}))

    def test_21_inline_fixed_lod_in_a_row_formula_is_a_column(self):
        out = translator().translate('[D] = {MAX([D])}', 'dimension')
        self.assertEqual((out['kind'], out['dax']), ('calculated column', "'T'[D] = CALCULATE(MAX('T'[D]), REMOVEFILTERS('T'))"))
        out = translator().translate('[N] / {FIXED [A]: SUM([N])}', 'dimension')
        self.assertEqual(out['kind'], 'calculated column')
        self.assertIn("ALLEXCEPT('T', 'T'[A])", out['dax'])
        self.assertEqual(translator().translate('SUM([N]) / {FIXED : SUM([N])}', 'measure')['kind'], 'measure')

    def test_22_sets_become_row_formulas(self):
        import xml.etree.ElementTree as ET
        import bhayankar_extractor as bx
        g = lambda xml: ET.fromstring(xml)
        top = g("<group name='[Last 4 Weeks]'><groupfilter count='4' end='top' function='end' units='records'>"
                "<groupfilter direction='DESC' expression='MAX([StartDt])' function='order'>"
                "<groupfilter function='level-members' level='[none:StartDt:qk]'/></groupfilter></groupfilter></group>")
        self.assertEqual(bx.set_formula(top), "__SET_TOPN([StartDt], 4, 'DESC')")
        cond = g("<group name='[Last Week]'><groupfilter expression='[StartDt] = {MAX([StartDt])}' function='filter'>"
                 "<groupfilter function='level-members' level='[none:StartDt:qk]'/></groupfilter></group>")
        self.assertEqual(bx.set_formula(cond), '([StartDt] = {MAX([StartDt])})')
        mem = g("<group name='[Comp]'><groupfilter function='member' level='[CyComp]' member='\"Y\"'/></group>")
        self.assertEqual(bx.set_formula(mem), '[CyComp] = "Y"')
        agg = g("<group name='[Big]'><groupfilter expression='SUM([Sales]) &gt; 100' function='filter'>"
                "<groupfilter function='level-members' level='[Customer]'/></groupfilter></group>")
        self.assertIsNone(bx.set_formula(agg))  # a condition per member on an aggregate: left to a person
        dax = translator().translate("__SET_TOPN([D], 4, 'DESC')", 'dimension')
        self.assertEqual(dax['dax'], "('T'[D] IN SELECTCOLUMNS(TOPN(4, ALL('T'[D]), 'T'[D], DESC), \"@m\", 'T'[D]))")

    def test_23_calc_of_parameters_only_counts_as_a_parameter(self):
        tr = translator({'[Shift]': ('measure', None, 'integer')})
        res = tr.resolve
        tr.resolve = lambda t: {**res(t), 'param_only': True} if t == '[Shift]' else res(t)
        out = tr.translate('[D] >= [Start] - [Shift]', 'dimension')  # [Start] unknown -> stays a reference
        self.assertEqual(out['kind'], 'calculated column')
        self.assertTrue(any('reads a parameter' in n for n in out['notes']))

    def test_24_shape_marks_on_headers_become_a_matrix(self):
        from bhayankar_report import Rebuild
        r = Rebuild.__new__(Rebuild)
        r.cat, r.notes, r.measures, r.cols, r.calcs = None, [], {}, {'Order ID': {'name': 'Order ID', 'table': 'T', 'datatype': 'string'},
                                                                  'Status': {'name': 'Status', 'table': 'T', 'datatype': 'string'}}, {}
        r.groups, r.params, r.formats, r.date_roles, r.url_columns, r.param_dims = {}, {}, {}, [], [], {}
        f = lambda lab, m=False: {'label': lab, 'measure': m, 'continuous': m, 'date': False}
        ws = {'name': 'Orders Detailed Table', 'shelf': {'rows': [f('Order ID')], 'cols': [], 'dual': None},
              'panes': [{'mark': 'Shape', 'enc': [{**f('Status'), 'channel': 'shape'}]}], 'filters': []}
        v = r.infer_visual(ws)
        self.assertEqual(v['visualType'], 'tableEx')
        self.assertEqual([x['name'] for x in v['roles']['Values']], ['Order ID', 'Status'])

    def test_27_url_action_with_a_parameter_still_gets_its_column(self):
        from bhayankar_report import Rebuild
        res = tb.DesignReviewRerunFixes.res()
        res['actions'].append({'name': 'Open', 'kind': 'URL', 'source': 'Main / Bars', 'trigger': 'on-select',
                               'url': 'https://x.invalid/?jp=<[Job Plan]>&d=<[Param:Divested?]>'})
        spec = Rebuild(res, catalog=None).build()
        uc = spec['model']['url_columns']
        self.assertEqual([(u['table'], u['fields']) for u in uc], [('W', ['Job Plan'])])  # was table None: no column

    def test_28_measure_names_called_multiple_values_is_a_measure_pie(self):
        from bhayankar_report import Rebuild
        r = Rebuild.__new__(Rebuild)
        r.cat, r.notes, r.measures, r.cols, r.calcs = None, [], {}, {}, {}
        r.groups, r.params, r.formats, r.date_roles, r.url_columns, r.param_dims = {}, {}, {}, [], [], {}
        ws = {'name': 'Pie', 'shelf': {'rows': [], 'cols': [], 'dual': None}, 'measure_values': ['SUM(Sales)', 'SUM(Profit)'],
              'panes': [{'mark': 'Pie', 'enc': [{'channel': 'color', 'label': 'Multiple Values', 'measure': False, 'continuous': False,
                                                 'date': False}]}], 'filters': []}
        self.assertTrue(r.infer_visual(ws).get('measure_slices'))

    def test_25_show_takes_a_workbook_name(self):
        import subprocess
        p = subprocess.run([sys.executable, os.path.join(os.path.dirname(HERE), 'bhayankar_extractor.py'), 'no such workbook',
                            '--show', 'x'], capture_output=True, text=True)
        self.assertNotIn('Traceback', p.stderr)
        self.assertIn('not found', p.stderr)


if __name__ == '__main__':
    unittest.main()
