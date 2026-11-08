"""Tests for bhayankar_verify.py that need no Power BI Desktop (the engine bridge itself is exercised by a live run,
see docs/PROJECT_LOG.md), and for the generator fixes that verify found on the Facility Metrics workbook."""
import json
import os
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
TOOL = os.path.dirname(HERE)
sys.path.insert(0, TOOL)
import bhayankar_verify as bv  # noqa: E402
import bhayankar_report as br  # noqa: E402


class PortFile(unittest.TestCase):
    def test_reads_utf16_and_ascii_port_files(self):
        d = tempfile.mkdtemp(prefix='bhy_port_')
        for enc, raw in (('utf-16-le', '56285'), ('ascii', '56285\n')):
            p = os.path.join(d, 'p.txt')
            with open(p, 'wb') as f:
                f.write(raw.encode(enc))
            self.assertEqual(bv.read_port(p), 56285)

    def test_no_engines_when_nothing_listens(self):
        old = bv.WORKSPACE_ROOTS
        bv.WORKSPACE_ROOTS = [tempfile.mkdtemp(prefix='bhy_ws_')]
        try:
            self.assertEqual(bv.engines(), [])
        finally:
            bv.WORKSPACE_ROOTS = old


class Names(unittest.TestCase):
    def test_dax_quoting(self):
        self.assertEqual(bv.dq_table("Owner's Table"), "'Owner''s Table'")
        self.assertEqual(bv.dq_name('Share [%]'), '[Share [%]]]')

    def test_result_columns_lose_the_table_prefix(self):
        r = {'columns': ["Date[Year]", '[Total]', 'Plain'], 'rows': [[2026, 5, 'x']]}
        self.assertEqual(bv.records(r), [{'Year': 2026, 'Total': 5, 'Plain': 'x'}])


class Verdicts(unittest.TestCase):
    base = {'problems': [], 'counts': {'A': 5, 'B': 3}, 'measures': [{'ok': True, 'ms': 5}]}

    def test_order_of_severity(self):
        self.assertEqual(bv.verdict({**self.base, 'error': 'x'}), 'NOT VERIFIED')
        self.assertEqual(bv.verdict({**self.base, 'problems': [{}]}), 'MODEL ERRORS')
        self.assertEqual(bv.verdict({**self.base, 'counts': {'A': 0, 'B': 0, 'C': 7}}), 'NOT REFRESHED')
        self.assertEqual(bv.verdict({**self.base, 'measures': [{'ok': False}]}), 'MEASURE ERRORS')
        self.assertEqual(bv.verdict({**self.base, 'numbers': {'status': 'ok', 'fail': 2}}), 'NUMBERS DIFFER')
        self.assertEqual(bv.verdict(self.base), 'ENGINE CHECKS PASSED')

    def test_calculation_needed_state_is_not_an_error(self):
        self.assertTrue({1, 3, 4} <= bv.OK_STATES)   # Ready, NoData, CalculationNeeded: nothing refreshed yet
        self.assertFalse({5, 6, 7} & bv.OK_STATES)   # semantic / evaluation / dependency errors

    def test_report_says_what_to_do_when_not_refreshed(self):
        st = {**self.base, 'workbook': 'W', 'time': 't', 'project': 'p', 'counts': {'A': 0, 'B': 0}, 'verdict': 'NOT REFRESHED',
              'engine': {'port': 1, 'match': 1.0}}
        self.assertIn('Refresh', bv.render(st))

    def test_report_without_an_open_project_explains_how_to_open_it(self):
        st = {'workbook': 'W', 'time': 't', 'error': 'not open', 'verdict': 'NOT VERIFIED'}
        text = bv.render(st)
        self.assertIn('--open', text)
        self.assertIn('Could not verify', text)


class TableFiles(unittest.TestCase):
    def test_sources_json_file_names_become_oracle_table_files(self):
        d = tempfile.mkdtemp(prefix='bhy_src_')
        with open(os.path.join(d, 'sources.json'), 'w', encoding='utf-8') as f:
            json.dump({'mode': 'sample', 'tables': {'T1': {'file': 'x.csv'}, 'T2': {}}}, f)
        old = bv.INPUT_DIR
        bv.INPUT_DIR = d
        try:
            self.assertEqual(bv.table_files('C:/data'), {'T1': os.path.join('C:/data', 'x.csv')})
        finally:
            bv.INPUT_DIR = old


class QfpFindings(unittest.TestCase):
    """Generator bugs that the engine check found on the Facility Metrics workbook."""

    def rb(self, calc_type, fn='SUM'):
        rb = br.Rebuild.__new__(br.Rebuild)
        rb.cols, rb.calcs, rb.groups, rb.measures, rb.notes = {}, {'Met X': {'pbi_kind': 'calculated column', 'datatype': calc_type,
                                                                              'table': 'T', 'formula': ''}}, {}, {}, []
        return rb

    def test_sum_of_a_text_calc_is_the_field_not_a_sum(self):
        rb = self.rb('string')
        ref = rb.column_ref('Met X')
        self.assertEqual(ref['kind'], 'Column')
        self.assertIn('SUM', br.NUMERIC_ONLY_AGG)
        self.assertEqual({'string', 'boolean', 'date', 'datetime'} & {'integer', 'real'}, set())

    def test_unparsed_filter_rule_is_not_folded_into_measures(self):
        src = open(os.path.join(TOOL, 'bhayankar_report.py'), encoding='utf-8').read()
        i = src.index("if out.get('type') == 'Manual':")
        self.assertIn('return out', src[i:i + 200])
        self.assertLess(i, src.index("'Param:' in (c['formula'] or '')"))


class Launchers(unittest.TestCase):
    def test_every_launcher_exists_and_none_needs_node(self):
        for bat in ('Run_Build', 'Run_Extractor', 'Run_PowerBI_Report_Test', 'Run_SelfCheck', 'Run_Validate', 'Run_Verify',
                    'Run_Watch', 'Run_Workflow'):
            p = os.path.join(TOOL, bat + '.bat')
            self.assertTrue(os.path.exists(p), bat)
            text = open(p, encoding='utf-8').read()
            self.assertNotIn('powerbi-report-author', text)
            self.assertNotIn('npm', text)

    def test_no_kiro_or_node_dependency_left_in_the_code(self):
        for fn in os.listdir(TOOL):
            if fn.startswith('bhayankar_') and fn.endswith('.py') and fn not in ('bhayankar_report.py', 'bhayankar_validate.py'):
                text = open(os.path.join(TOOL, fn), encoding='utf-8').read()
                self.assertNotIn('steering', text, fn)
                self.assertNotIn('.kiro', text, fn)


if __name__ == '__main__':
    unittest.main()
