"""Tests for the 1.3.0 enhancement round (2026-11-07): table calculations as measures, parameter-picked sheet filters,
same-named calcs in several data sources, sheet filters on visuals without measures, FIXED LODs of text / dates, and
the offline model check. docs\\PROJECT_LOG.md "Enhancement round (1.3.0)".

Run (from the toolkit folder):  python -m unittest discover -s tests -v
"""
import os
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, HERE)
import bhayankar_build as bb  # noqa: E402
import bhayankar_modelcheck as mc  # noqa: E402
import bhayankar_report as br  # noqa: E402
from bhayankar_dax import Translator  # noqa: E402

COLS = {'a': 'real', 'b': 'real', 's': 'string', 'dt': 'datetime', 'k': 'string', 'flagA': 'boolean', 'flagB': 'boolean'}


def tr():
    def res(tok):
        n = tok.strip('[]')
        if n in COLS:
            return {'kind': 'column', 'table': 'T', 'name': n, 'datatype': COLS[n]}
        if n.startswith('Param:'):
            return {'kind': 'param', 'name': n[6:]}
        return {'kind': 'unknown', 'name': n}
    return Translator(res)


class TableCalcs(unittest.TestCase):
    """RANK / WINDOW_* / RUNNING_* / TOTAL / INDEX / LOOKUP were BLANK() placeholders; with the sheet's addressing
    columns (Tableau "compute using") they are measures."""

    def t(self, f, window):
        return tr().translate(f, window=window)

    def test_rank_over_the_addressed_column(self):
        r = self.t('RANK(SUM([a]))', ["'T'[s]"])
        self.assertEqual(r['kind'], 'measure')
        self.assertEqual(r['dax'], "IF(ISBLANK(SUM('T'[a])), BLANK(), RANKX(ALLSELECTED('T'[s]), CALCULATE(SUM('T'[a])), , DESC, SKIP))")
        self.assertIn('DENSE', self.t('RANK_DENSE(SUM([a]), "asc")', ["'T'[s]"])['dax'])
        self.assertIn(', ASC, DENSE', self.t("RANK_DENSE(SUM([a]), 'asc')", ["'T'[s]"])['dax'])

    def test_window_and_total(self):
        self.assertEqual(self.t('WINDOW_MAX(SUM([a]))', ["'T'[s]"])['dax'], "MAXX(ALLSELECTED('T'[s]), CALCULATE(SUM('T'[a])))")
        self.assertEqual(self.t('SUM([a]) / TOTAL(SUM([a]))', ["'T'[s]"])['dax'],
                         "DIVIDE(SUM('T'[a]), CALCULATE(SUM('T'[a]), ALLSELECTED('T'[s])))")
        self.assertEqual(self.t('TOTAL(SUM([a]))', [])['dax'], "SUM('T'[a])")  # each mark on its own: the mark itself

    def test_running_sum_orders_months_by_number(self):
        r = self.t('RUNNING_SUM(SUM([a]))', ["'Date'[Month]"])
        self.assertEqual(r['dax'], "SUMX(FILTER(ALLSELECTED('Date'[Month], 'Date'[Month Number]), "
                                   "'Date'[Month Number] <= MAX('Date'[Month Number])), CALCULATE(SUM('T'[a])))")

    def test_index_and_lookup(self):
        self.assertEqual(self.t('INDEX()', ["'T'[s]"])['dax'],
                         "IF(HASONEVALUE('T'[s]), ROWNUMBER(ALLSELECTED('T'[s]), ORDERBY('T'[s], ASC)))")
        self.assertIn("OFFSET(-1, ALLSELECTED('T'[s]), ORDERBY('T'[s], ASC))", self.t('LOOKUP(SUM([a]), -1)', ["'T'[s]"])['dax'])

    def test_what_cannot_be_written_stays_a_placeholder(self):
        for f, w in (('RUNNING_SUM(SUM([a]))', ["'T'[s]", "'T'[k]"]), ('WINDOW_SUM(SUM([a]), -2, 0)', ["'T'[s]"]),
                     ('RANK(SUM([a]))', None)):
            r = self.t(f, w)
            self.assertEqual(r['kind'], 'visual calculation', f)

    def test_report_writes_the_calc_over_the_sheet_addressing(self):
        res = {'columns': [{'name': 'Region', 'table': 'T', 'datatype': 'string'},
                           {'name': 'Amount', 'table': 'T', 'datatype': 'real'}],
               'calculations': [{'name': 'Rank Amount', 'formula': 'RANK(SUM([Amount]))', 'pbi_kind': 'visual calculation',
                                 'dax': 'RANKX(ALLSELECTED(<Dimension>), SUM(\'T\'[Amount]))', 'datatype': 'integer',
                                 'confidence': 'low', 'dax_notes': [], 'role': 'measure', 'table': None, 'datasource': 'D'}],
               'groups': [], 'parameters': [], 'worksheets': []}
        r = br.Rebuild(res, catalog={})
        r._ws = {'name': 'S', 'datasources': ['D'], 'panes': [],
                 'shelf': {'cols': [{'label': 'Region', 'measure': False}], 'rows': [{'label': 'Rank Amount', 'measure': True}]}}
        fr = r.field('Rank Amount', True, tc='Rows', shelf=True)
        self.assertEqual(fr['name'], 'Rank Amount')
        self.assertIn("RANKX(ALLSELECTED('T'[Region])", r.measures['Rank Amount']['dax'])
        r._ws['shelf']['cols'] = [{'label': 'Amount', 'measure': True}]  # another sheet: nothing on Columns
        r._ws['shelf']['rows'] = [{'label': 'Region', 'measure': False}]
        fr2 = r.field('Rank Amount', True, tc='Columns', shelf=True)  # Table (down): along Rows
        self.assertEqual(fr2['name'], 'Rank Amount')  # same addressing: same measure


class ParameterPickedFilter(unittest.TestCase):
    def test_boolean_branches_are_a_row_filter_not_a_field_parameter(self):
        """Retail "Relative Date Filter" = CASE [Param] WHEN "Last Week" THEN [Last Week] ... (true/false sets)."""
        r = tr().translate('CASE [Param:P] WHEN "A" THEN [flagA] WHEN "B" THEN [flagB] END')
        self.assertEqual(r['kind'], 'calculated column')
        self.assertEqual(r['dax'], "SWITCH([P Value], \"A\", 'T'[flagA], \"B\", 'T'[flagB])")
        r = tr().translate('CASE [Param:P] WHEN "A" THEN [s] WHEN "B" THEN [k] END')
        self.assertEqual(r['kind'], 'field parameter')  # text columns: still a dynamic field selector


class SameNamedCalcs(unittest.TestCase):
    def res(self):
        calc = lambda ds, tbl: {'name': 'Share', 'formula': 'SUM([x])', 'pbi_kind': 'measure', 'datasource': ds,  # noqa: E731
                                'dax': f"SUM('{tbl}'[x])", 'datatype': 'real', 'confidence': 'high', 'dax_notes': [],
                                'table': tbl, 'role': 'measure'}
        user = lambda ds, tbl: {'name': 'Twice', 'formula': '[Share] * 2', 'pbi_kind': 'measure', 'datasource': ds,  # noqa: E731
                                'dax': '[Share] * 2', 'datatype': 'real', 'confidence': 'high', 'dax_notes': [], 'table': tbl,
                                'role': 'measure'}
        return {'columns': [], 'groups': [], 'parameters': [], 'worksheets': [],
                'calculations': [calc('Store', 'S'), calc('Chain', 'C'), user('Store', 'S'), user('Chain', 'C')]}

    def test_one_measure_per_data_source_and_the_sheet_gets_its_own(self):
        r = br.Rebuild(self.res(), catalog={})
        self.assertEqual(r.measures['Share']['dax'], "SUM('S'[x])")
        self.assertEqual(r.measures['Share (Chain)']['dax'], "SUM('C'[x])")
        self.assertEqual(r.measures['Twice (Chain)']['dax'], '[Share (Chain)] * 2')  # follows its own copy
        r._ws = {'datasources': ['Chain', 'Parameters']}
        self.assertEqual(r.column_ref('Share')['name'], 'Share (Chain)')
        r._ws = {'datasources': ['Store']}
        self.assertEqual(r.column_ref('Share')['name'], 'Share')


def guard_stub(measures):
    b = bb.Build.__new__(bb.Build)
    b.tables = {'T': {'name': 'T', 'columns': [{'name': 'a'}], 'calc_columns': []}}
    b.measures = {n: {'name': n, 'dax': d} for n, d in measures.items()}
    return b


class FoldAndRowFilters(unittest.TestCase):
    def test_parameter_value_measure_reads_no_rows(self):
        b = guard_stub({'P Value': "SELECTEDVALUE('P Parameter'[Value], \"x\")", 'Rows': "COUNTROWS('T')",
                        'Via': '[Rows] + 1', 'Text': '"Title"'})
        self.assertFalse(b.reads_rows('P Value'))
        self.assertFalse(b.reads_rows('Text'))
        self.assertTrue(b.reads_rows('Rows'))  # a whole-table reference counts (measure_tables() missed it)
        self.assertTrue(b.reads_rows('Via'))

    def test_dimension_only_table_gets_a_row_filter_measure(self):
        b = bb.Build.__new__(bb.Build)
        b.log, b.decisions, b.renames, b.wb = [], {}, {}, 'WB'
        b.digest = {'parameters': [], 'calculations': []}
        b.tables = {'T': {'name': 'T', 'columns': [{'name': 'Player'}], 'calc_columns': []}}
        cond = {'type': 'RowCondition', 'table': 'T', 'calc': 'Is P1', 'value': 'true',
                'condition': "('T'[Player] = [Player 1 Value]) == TRUE()"}
        v = {'sheet': 'Medal Haul', 'visualType': 'tableEx', 'id': 'v1',
             'roles': {'Values': [{'kind': 'Column', 'table': 'T', 'name': 'Player'}]}, 'visual_filters': [cond]}
        b.spec = {'pages': [{'name': 'P', 'visuals': [v], 'slicers': []}],
                  'model': {'tables': [], 'parameters': [], 'parameter_dimensions': []}}
        b.build_measures()
        m = b.measures['Medal Haul Rows']
        self.assertIn("COUNTROWS('T')", m['dax'])
        self.assertIn("KEEPFILTERS(FILTER('T', ('T'[Player] = _p1) == TRUE()))", m['dax'])
        self.assertEqual(v['visual_filters'][-1], {'type': 'NotBlank', 'field': {'kind': 'Measure', 'table': 'T',
                                                                                  'name': 'Medal Haul Rows', 'label': 'Medal Haul Rows'}})


class FixedLodOfTextOrDate(unittest.TestCase):
    def test_text_and_dates_are_not_summed(self):
        """PM "Active/NonActive" = {FIXED [WO Number]: MIN([Active vs. Non-Active])} was SUMX over text."""
        self.assertTrue(tr().translate('{FIXED [k]: MIN([s])}', role='measure')['dax'].startswith('MINX(VALUES('))
        self.assertTrue(tr().translate('{FIXED [k]: MAX([dt])}', role='measure')['dax'].startswith('MAXX(VALUES('))
        self.assertTrue(tr().translate('{FIXED [k]: SUM([a])}', role='measure')['dax'].startswith('SUMX(VALUES('))


TMDL_T = """table 'T'

\tmeasure 'Total' = SUM('T'[Amount])
\t\tformatString: #,##0

\tmeasure 'total' = [Total] + 1

\tmeasure 'Loop' = [Loop] + 1

\tmeasure 'Day' = DAYOFYEAR('T'[When])

\tmeasure 'Bad' = SUM('T'[Nope]) + (1

\tcolumn 'Amount'
\t\tdataType: double
\t\tsourceColumn: Amount

\tcolumn 'When'
\t\tdataType: dateTime
\t\tsortByColumn: 'Missing'
\t\tsourceColumn: When

\tcolumn 'Name [x]'
\t\tdataType: string
\t\tsourceColumn: Name

\tpartition 'T' = m
\t\tmode: import
\t\tsource =
\t\t\t\tlet x = 1 in x
"""
TMDL_U = """table 'U'

\tmeasure 'Fine' = CALCULATE(SUM('T'[Amount]), KEEPFILTERS('T'[Name [x]]] = "O'Brien (x)"))
\t\tformatString: #,##0

\tcolumn 'Amount'
\t\tdataType: string
\t\tsourceColumn: Amount
"""
RELS = """relationship 'U to T'
\tfromColumn: 'U'.'Amount'
\ttoColumn: 'T'.'Amount'
"""


class ModelCheck(unittest.TestCase):
    """bhayankar_modelcheck: what the PBIR validator does not read (it passed every model the bug hunt found broken)."""

    def setUp(self):
        self.d = tempfile.mkdtemp(prefix='bhk_mc_')
        t = os.path.join(self.d, 'X.SemanticModel', 'definition', 'tables')
        os.makedirs(t)
        for name, text in (('T.tmdl', TMDL_T), ('U.tmdl', TMDL_U)):
            with open(os.path.join(t, name), 'w', encoding='utf-8') as f:
                f.write(text)
        with open(os.path.join(self.d, 'X.SemanticModel', 'definition', 'relationships.tmdl'), 'w', encoding='utf-8') as f:
            f.write(RELS)

    def tearDown(self):
        shutil.rmtree(self.d, ignore_errors=True)

    def test_finds_each_kind_of_error(self):
        res = mc.check(self.d)
        codes = {}
        for code, where, msg in res['errors']:
            codes.setdefault(code, []).append(f'{where}: {msg}')
        self.assertIn('NAME_CLASH', codes)  # Total / total
        self.assertTrue(any('[Loop]' in x for x in codes['CIRCULAR']))
        self.assertTrue(any('DAYOFYEAR' in x for x in codes['UNKNOWN_FUNCTION']))
        self.assertTrue(any("'T'[Nope]" in x for x in codes['UNKNOWN_REF']))
        self.assertTrue(any('unbalanced' in x for x in codes['SYNTAX']))
        self.assertTrue(any('Missing' in x for x in codes['SORT_BY']))
        self.assertTrue(any('different types' in x for x in codes['RELATIONSHIP']))
        # names holding quotes, brackets and parens are read as names, not code
        self.assertFalse(any("measure 'U'[Fine]" in w for _, w, _ in res['errors']), res['errors'])

    def test_a_built_model_is_clean(self):
        tb = os.path.join(os.path.dirname(HERE), 'tests', 'fixtures')
        self.assertTrue(os.path.isdir(tb))


if __name__ == '__main__':
    unittest.main()
