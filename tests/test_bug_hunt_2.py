"""Regression tests for the bug hunt of 2026-11-07 (Standalone 1.2.0 -> 1.2.1): 13 workbooks run through the toolkit.
One test (or more) per bug; the number is the one in docs\\PROJECT_LOG.md ("Bug hunt 2026-11-07").

Run (from the toolkit folder):  python -m unittest discover -s tests -v
"""
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, HERE)
import bhayankar_build as bb  # noqa: E402
from bhayankar_dax import Translator, strip_brackets  # noqa: E402
from test_bug_hunt import guard_build  # noqa: E402


def dax(formula):
    cols = {'a': 'real', 'b': 'real', 'c': 'real', 's': 'string', 't': 'string', 'dt': 'datetime', 'd2': 'date'}

    def res(tok):
        n = tok.strip('[]')
        if n in cols:
            return {'kind': 'column', 'table': 'T', 'name': n, 'datatype': cols[n]}
        return {'kind': 'unknown', 'name': n}
    return Translator(res).translate(formula)['dax']


class Brackets(unittest.TestCase):
    """Bug 1: the parse tree has no bracket nodes and the emitter wrote none back: ([a] + [b]) * [c] became a + b * c
    (Retail "Neg $ Impact" and Superstore "Profit Ratio Difference" showed wrong numbers)."""

    def test_01_lower_precedence_operand_is_bracketed(self):
        self.assertEqual(dax('([a] + [b]) * [c]'), "('T'[a] + 'T'[b]) * 'T'[c]")
        self.assertEqual(dax('[a] * ([b] - 1)'), "'T'[a] * ('T'[b] - 1)")
        self.assertEqual(dax('SUM([a]) * (SUM([b]) - SUM([c]))'), "SUM('T'[a]) * (SUM('T'[b]) - SUM('T'[c]))")
        self.assertEqual(dax('(SUM([a]) + SUM([b])) / 2'), "(SUM('T'[a]) + SUM('T'[b])) / 2")

    def test_01b_right_operand_of_minus_keeps_its_brackets(self):
        self.assertEqual(dax('[a] - ([b] - [c])'), "'T'[a] - ('T'[b] - 'T'[c])")
        self.assertEqual(dax('SUM([a]) - (SUM([b]) + SUM([c]))'), "SUM('T'[a]) - (SUM('T'[b]) + SUM('T'[c]))")
        self.assertEqual(dax('[a] - [b] - [c]'), "'T'[a] - 'T'[b] - 'T'[c]")  # left to right needs none

    def test_01c_negated_sum_is_bracketed(self):
        self.assertEqual(dax('-([a] + [b])'), "-('T'[a] + 'T'[b])")
        self.assertEqual(dax('-SUM([a]) * 2'), "-SUM('T'[a]) * 2")

    def test_01d_divide_arguments_have_no_extra_brackets(self):
        self.assertEqual(dax('([a]+[b])/([c]-1)'), "DIVIDE('T'[a] + 'T'[b], 'T'[c] - 1)")
        self.assertEqual(strip_brackets('(a) / (b)'), '(a) / (b)')
        self.assertEqual(strip_brackets("('T'[x (y)] + 1)"), "'T'[x (y)] + 1")
        self.assertEqual(strip_brackets('(")" & x)'), '")" & x')

    def test_02_function_results_written_as_operators_are_bracketed(self):
        """Bug 2: DATEADD / DATETRUNC week / MAKEDATETIME / STARTSWITH come out as a + b: [d2] - DATEADD('day', 7, [dt])
        was d2 - dt + 7."""
        self.assertEqual(dax("[d2] - DATEADD('day', 7, [dt])"), "'T'[d2] - ('T'[dt] + (7))")
        self.assertEqual(dax("[d2] - DATETRUNC('week', [dt])"), "'T'[d2] - ('T'[dt] - WEEKDAY('T'[dt], 1) + 1)")
        self.assertEqual(dax('STARTSWITH([s], "x") = FALSE'), '(LEFT(\'T\'[s], LEN("x")) = "x") == FALSE()')


class Functions(unittest.TestCase):
    def test_03_round_without_decimals_is_valid_dax(self):
        """Bug 3: Tableau ROUND(x) has optional decimals; DAX ROUND(x) is a syntax error."""
        self.assertEqual(dax('ROUND([a])'), "ROUND('T'[a], 0)")
        self.assertEqual(dax('ROUND([a], 2)'), "ROUND('T'[a], 2)")

    def test_04_dateadd_minutes_and_seconds(self):
        """Bug 4: every unit below a day was divided by 24 (30 minutes = 30 hours)."""
        self.assertEqual(dax("DATEADD('minute', 30, [dt])"), "('T'[dt] + (30) / 1440)")
        self.assertEqual(dax("DATEADD('second', 5, [dt])"), "('T'[dt] + (5) / 86400)")
        self.assertEqual(dax("DATEADD('hour', 2, [dt])"), "('T'[dt] + (2) / 24)")

    def test_05_date_parts_without_a_dax_function_of_their_name(self):
        """Bug 5: DATEPART('dayofyear') / 'iso-week' wrote DAYOFYEAR() / ISO-WEEK() (no such functions); DATENAME('week')
        wrote WEEK()."""
        self.assertEqual(dax("DATEPART('dayofyear', [dt])"), "(DATEDIFF(DATE(YEAR('T'[dt]), 1, 1), 'T'[dt], DAY) + 1)")
        self.assertEqual(dax("DATEPART('iso-week', [dt])"), "WEEKNUM('T'[dt], 21)")
        self.assertEqual(dax("DATENAME('week', [dt])"), 'FORMAT(WEEKNUM(\'T\'[dt]), "0")')
        self.assertIn('TODO', dax("DATEPART('iso-quarter', [dt])"))


class LoadGuard(unittest.TestCase):
    def test_06_measure_reading_itself_is_not_written(self):
        """Bug 6: a calc named like the field it reads ([Amount] = [Amount]) was written as a circular measure."""
        b = guard_build({'Amount': '[Amount]', 'Ping': '[Pong] + 1', 'Pong': '[Ping] * 2', 'Ok': "COUNTROWS('F')",
                         'Uses': '[Amount] + [Ok]'})
        for n in ('Amount', 'Ping', 'Pong'):
            self.assertEqual(b.measures[n]['dax'], 'BLANK()', n)
            self.assertIn('circular reference', b.measures[n]['description'], n)
        self.assertEqual(b.measures['Ok']['dax'], "COUNTROWS('F')")
        self.assertIn('WRONG UNTIL FIXED', b.measures['Uses']['description'])

    def test_07_untranslated_function_inside_a_measure_is_flagged(self):
        """Bug 7: SPLIT / REGEXP_* were written as /* TODO */ BLANK() inside a working-looking measure."""
        b = guard_build({'Secs': "SUMX('F', CONVERT(/* TODO SPLIT */ BLANK(), DOUBLE) * 60)", 'Avg': '[Secs] / 2'})
        self.assertEqual(b.measures['Secs']['dax'], 'BLANK()')
        self.assertEqual(b.measures['Secs']['displayFolder'], 'Measures\\_Needs fix')
        self.assertIn('not translated: SPLIT', b.measures['Secs']['description'])
        self.assertIn('WRONG UNTIL FIXED', b.measures['Avg']['description'])


def model_build(columns, calcs, visual_fields=()):
    """build_model() of one table T with the given source columns and calculated columns ({name: dax})."""
    b = bb.Build.__new__(bb.Build)
    b.log, b.decisions, b.wb = [], {}, 'WB'
    b.sources = {'mode': 'warehouse', 'tables': {}}
    b.digest = {'groups': [], 'calculations': [], 'parameters': [], 'flags': [], 'worksheets': []}
    col = lambda n: {'name': n, 'source_column': n, 'datatype': 'string'}  # noqa: E731
    b.spec = {'pages': [{'name': 'P', 'visuals': [{'sheet': 'S', 'roles': {'Category': [
        {'kind': 'Column', 'table': 'T', 'name': n} for n in visual_fields]}}]}],
        'model': {'relationships': [], 'url_columns': [], 'parameter_dimensions': [], 'tables': [
            {'name': 'T', 'source_table': 'T', 'columns': [col(c) for c in columns],
             'calculated_columns': [{'name': n, 'dax': d} for n, d in calcs.items()], 'measures': []}]}}
    b.tds_relations = lambda: {}
    b.build_measures = lambda: setattr(b, 'measures', {})
    b.slice_tables, b.param_tables, b.pdim_tables, b.report_period = {}, {}, {}, None
    b.param_rowcalcs = []
    b.build_model()
    return b


class NameClashes(unittest.TestCase):
    """Bug 8: Power BI names are case-insensitive. Putney "EVENT" = UPPER([Event]) next to the column Event, Retail calcs
    "POINT OF SALE" and "Point of Sale": the same name twice, and Desktop refuses the model."""

    def test_08_calc_differing_only_in_case_is_renamed_and_followed(self):
        b = model_build(['Event', 'Medal'], {'EVENT': "UPPER('T'[Event])", 'Shout': "'T'[EVENT] & \"!\""},
                        visual_fields=['EVENT', 'Event'])
        t = b.tables['T']
        names = [c['name'] for c in t['columns'] + t['calc_columns']]
        self.assertEqual(len({n.lower() for n in names}), len(names), names)
        self.assertIn('Event', names)  # the source column keeps its name
        self.assertIn('EVENT (calc)', names)
        calc = {c['name']: c['dax'] for c in t['calc_columns']}
        self.assertEqual(calc['EVENT (calc)'], "UPPER('T'[Event])")  # still reads the source column
        self.assertEqual(calc['Shout'], "'T'[EVENT (calc)] & \"!\"")
        fields = [f['name'] for f in b.spec['pages'][0]['visuals'][0]['roles']['Category']]
        self.assertEqual(fields, ['EVENT (calc)', 'Event'])
        self.assertTrue(any('renamed [EVENT (calc)]' in x for x in b.log))

    def test_08b_two_calcs_differing_only_in_case(self):
        b = model_build(['DataSourceDesc'], {'POINT OF SALE': "'T'[DataSourceDesc]", 'Point of Sale': "'T'[DataSourceDesc]"})
        names = [c['name'] for c in b.tables['T']['calc_columns']]
        self.assertEqual(names, ['POINT OF SALE', 'Point of Sale (calc)'])

    def test_08c_calc_repeating_the_column_of_its_name_is_written_once(self):
        b = model_build(['Region'], {'Region': "'T'[Region]"})
        t = b.tables['T']
        self.assertEqual([c['name'] for c in t['columns'] + t['calc_columns']], ['Region'])

    def test_09_measure_names_unique_in_the_model_ignoring_case(self):
        """Bug 9: Retail "Total Ty_units_Ship" (one table) and "Total TY_UNITS_Ship" (another), and a measure named like a
        column of its own table."""
        b = bb.Build.__new__(bb.Build)
        b.log, b.decisions, b.renames, b.measures = [], {'measure_names': {'Keep': 'Kept'}}, {'Keep': 'Kept'}, {}
        b.tables = {'A': {'columns': [{'name': 'Units'}], 'calc_columns': []},
                    'B': {'columns': [{'name': 'Region'}], 'calc_columns': []}}
        spec = {'tables': [
            {'measures': [{'name': 'Total Units', 'table': 'A', 'dax': '1'}, {'name': 'Keep', 'table': 'A', 'dax': '1'}]},
            {'measures': [{'name': 'Total UNITS', 'table': 'B', 'dax': '2'}, {'name': 'region', 'table': 'B', 'dax': '3'},
                          {'name': 'Total Units', 'table': 'A', 'dax': '1'}]}]}
        b.unique_measure_names(spec, ['A'])
        self.assertEqual(b.renames, {'Keep': 'Kept', 'Total UNITS': 'Total UNITS (B)', 'region': 'region (B)'})
        self.assertEqual(b.fix_refs('[Total UNITS] + [Total Units]'), '[Total UNITS (B)] + [Total Units]')
        self.assertEqual(b.decisions['measure_names'], {'Keep': 'Kept'})  # decisions.json content is not changed


if __name__ == '__main__':
    unittest.main()
