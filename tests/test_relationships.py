"""Tests for Tableau relationships (logical layer, 2020.2+) -> Power BI relationships (2026-11-08, 1.5.0).
Fixtures: tests\\fixtures\\relationships\\rel_uc*.twb (synthetic, no data; tools\\make_relationship_fixtures.py), in the
XML shape of the real relationships in "Business Scenario @ Superstore_v2026.1.twbx". Before 1.5.0 every one of them
was dropped ("relationship not parsed"): Tableau names keys by field ([Region] = [Region (PEOPLE)]), not by table.

Run (from the toolkit folder):  python -m unittest discover -s tests -v
"""
import glob
import os
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, HERE)
import bhayankar_build as bb  # noqa: E402
import bhayankar_extractor as bx  # noqa: E402
import bhayankar_modelcheck as mc  # noqa: E402

FIX = os.path.join(HERE, 'fixtures', 'relationships')


def build(name, sample=None):
    """Extract + build one fixture; sample = {table: csv text} switches to sample mode. Returns (Build, root)."""
    root = tempfile.mkdtemp(prefix='bhk_rel_')
    inp, out = os.path.join(root, 'i'), os.path.join(root, 'o')
    os.makedirs(inp)
    shutil.copy(os.path.join(FIX, name + '.twb'), inp)
    if sample:
        os.makedirs(os.path.join(inp, name + '_Data'))
        for t, text in sample.items():
            with open(os.path.join(inp, name + '_Data', t + '.csv'), 'w', encoding='utf-8') as f:
                f.write(text)
    bx.run_folder(inp, out, table_map={})
    old = bb.INPUT_DIR, bb.OUTPUT_DIR
    bb.INPUT_DIR, bb.OUTPUT_DIR = inp, out
    try:
        b = bb.Build(name)
        b.run()
    finally:
        bb.INPUT_DIR, bb.OUTPUT_DIR = old
    return b, root


def rels(b):
    return {(a, ka, t, kt) for a, ka, t, kt in b.relationships if t != 'Date'}


class Extractor(unittest.TestCase):
    def keys(self, name):
        root, _ = bx.load_root(os.path.join(FIX, name + '.twb'))
        ds = next(d for d in bx.Workbook(root).build()['datasources'] if d['connection'].get('logical'))
        return ds['connection']['logical']['relationships']

    def test_keys_resolve_to_table_and_column(self):
        r = self.keys('rel_uc1_star')[0]
        self.assertEqual(r['keys'], [{'op': '=', 'left': {'field': 'Region', 'table': 'ORDERS', 'column': 'Region'},
                                      'right': {'field': 'Region (PEOPLE)', 'table': 'PEOPLE', 'column': 'Region'}}])
        self.assertEqual(r['end_tables'], ['ORDERS', 'PEOPLE'])

    def test_compound_calc_inequality_and_performance_options(self):
        self.assertEqual(len(self.keys('rel_uc2_composite')[0]['keys']), 2)
        self.assertEqual(self.keys('rel_uc5_calc_key')[0]['keys'][0]['left'],
                         {'field': 'Order Year', 'calc': 'YEAR([Order Date])'})
        self.assertEqual(self.keys('rel_uc6_inequality')[0]['keys'][0]['op'], '>=')
        r = self.keys('rel_uc3_cardinality')[0]
        self.assertEqual((r['unique'], r['integrity']), ([False, True], 'all'))


class Build(unittest.TestCase):
    def tearDown(self):
        shutil.rmtree(getattr(self, 'root', ''), ignore_errors=True)

    def clean(self, name):
        sm = glob.glob(os.path.join(self.root, 'o', name, 'pbip', '*.SemanticModel'))[0]
        res = mc.check(sm)
        self.assertEqual(res['errors'], [])
        with open(os.path.join(sm, 'definition', 'relationships.tmdl'), encoding='utf-8') as f:
            return f.read()

    def test_star_in_tableau_key_format(self):
        b, self.root = build('rel_uc1_star')
        self.assertEqual(rels(b), {('ORDERS', 'Region', 'PEOPLE', 'Region'), ('RETURNS', 'Order ID', 'ORDERS', 'Order ID')})
        self.assertIn(('ORDERS', 'PEOPLE'), b.m2m)  # no cardinality, no data, key not named after a table: Tableau default
        tm = self.clean('rel_uc1_star')
        self.assertIn("relationship 'ORDERS to PEOPLE'\n\tfromColumn: 'ORDERS'.'Region'\n\ttoColumn: 'PEOPLE'.'Region'\n\ttoCardinality: many", tm)
        self.assertFalse(any('not parsed' in x for x in b.log if isinstance(x, str)))

    def test_composite_key_column_on_both_tables(self):
        b, self.root = build('rel_uc2_composite')
        self.assertEqual(rels(b), {('ORDERS', 'Region | Category (key)', 'TARGETS', 'Region | Category (key)')})
        step = dict(b.tables['TARGETS']['steps'])['Added Region | Category (key)']
        self.assertIn('Text.Combine({Text.From([Region]), Text.From([Category])}, "|")', step)
        self.assertIn('[Region] = null or [Category] = null then null', step)  # Tableau never matches null keys
        self.clean('rel_uc2_composite')

    def test_cardinality_one_decides_the_one_side(self):
        b, self.root = build('rel_uc3_cardinality')
        self.assertEqual(rels(b), {('ORDERS', 'Product ID', 'PRODUCTS', 'Product ID')})
        self.assertEqual(b.m2m, set())
        self.assertTrue(any('cardinality One' in x and 'referential integrity "all"' in x for x in b.log if isinstance(x, str)))

    def test_snowflake_chain(self):
        b, self.root = build('rel_uc4_snowflake')
        self.assertEqual(rels(b), {('ORDERS', 'Product ID', 'PRODUCTS', 'Product ID'),
                                   ('PRODUCTS', 'Category ID', 'CATEGORIES', 'Category ID')})  # CATEGORIES owns Category ID
        self.assertEqual(b.m2m, set())
        self.clean('rel_uc4_snowflake')

    def test_relationship_on_a_calculated_column(self):
        b, self.root = build('rel_uc5_calc_key')
        self.assertEqual(rels(b), {('ORDERS', 'Order Year', 'YEAR_GOALS', 'Year')})
        self.assertTrue(any(c['name'] == 'Order Year' for c in b.tables['ORDERS']['calc_columns']))
        self.clean('rel_uc5_calc_key')  # the unused Year key takes the calc's type (no int64 / string mismatch)

    def test_inequality_is_a_decision(self):
        b, self.root = build('rel_uc6_inequality')
        self.assertEqual(rels(b), set())
        self.assertTrue(any('operator >=' in x and 'DECISION NEEDED' in x for x in b.log if isinstance(x, str)))

    def test_sample_data_decides_and_many_to_many_keeps_every_row(self):
        b, self.root = build('rel_uc1_star', sample={
            'ORDERS': 'Order ID,Region,Category,Product ID,Order Date,Sales\n1,East,A,p1,2026-01-01,5\n2,East,B,p2,2026-01-02,7\n',
            'PEOPLE': 'Region,Regional Manager\nEast,Ann\nWest,Bo\n',
            'RETURNS': 'Order ID,Returned\n1,Yes\n1,Yes\n'})
        r = rels(b)
        self.assertIn(('ORDERS', 'Region', 'PEOPLE', 'Region'), r)  # PEOPLE unique in the sample: one-to-many
        self.assertNotIn(('ORDERS', 'PEOPLE'), b.m2m)
        lines = '\n'.join(b.m_partition(b.tables['PEOPLE']))
        self.assertIn('Kept one row per Region', lines)

    def test_key_owner_reads_camel_case_and_plural(self):
        self.assertTrue(bb.key_owner('ProductKey', 'Dim_Product'))
        self.assertTrue(bb.key_owner('Category ID', 'CATEGORIES'))
        self.assertFalse(bb.key_owner('ProductKey', 'Fact_Sales'))


if __name__ == '__main__':
    unittest.main()
