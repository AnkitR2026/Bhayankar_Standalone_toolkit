"""Tests for 1.6.0 (2026-11-08): Power Query connector per Tableau source type, Tableau data source filters, value
aliases, and sheet filters on worksheet-tab pages. Fixtures: tests\\fixtures\\sources\\src_uc*.twb (synthetic, made-up
servers; tools\\make_source_fixtures.py).

Run (from the toolkit folder):  python -m unittest discover -s tests -v
"""
import glob
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
import bhayankar_extractor as bx  # noqa: E402
import bhayankar_modelcheck as mc  # noqa: E402

FIX = os.path.join(HERE, 'fixtures', 'sources')


def build(name, sources=None):
    root = tempfile.mkdtemp(prefix='bhk_src_')
    inp, out = os.path.join(root, 'i'), os.path.join(root, 'o')
    os.makedirs(inp)
    shutil.copy(os.path.join(FIX, name + '.twb'), inp)
    if sources is not None:
        with open(os.path.join(inp, 'sources.json'), 'w', encoding='utf-8') as f:
            json.dump(sources, f)
    bx.run_folder(inp, out, table_map={})
    old = bb.INPUT_DIR, bb.OUTPUT_DIR
    bb.INPUT_DIR, bb.OUTPUT_DIR = inp, out
    try:
        b = bb.Build(name)
        b.run()
    finally:
        bb.INPUT_DIR, bb.OUTPUT_DIR = old
    sm = glob.glob(os.path.join(out, name, 'pbip', '*.SemanticModel'))[0]
    return b, root, sm


def read_m(b, table):
    return '\n'.join(b.m_partition(b.tables[table]))


class Connectors(unittest.TestCase):
    def tearDown(self):
        shutil.rmtree(getattr(self, 'root', ''), ignore_errors=True)

    def test_each_source_type_gets_its_connector(self):
        b, self.root, sm = build('src_uc1_connectors')
        self.assertIn('Snowflake.Databases("acct.example.snowflakecomputing.com", "REPORTING_WH"){[Name = "SALESDB"]}[Data]'
                      '{[Name = "MART"]}[Data]{[Name = "VW_SALES"]}[Data]', read_m(b, 'VW_SALES'))
        self.assertIn('Excel.Workbook(File.Contents("C:/Data/Budget.xlsx"), null, true){[Item = "Budget", Kind = "Sheet"]}[Data]',
                      read_m(b, 'Budget'))
        self.assertIn('Csv.Document(File.Contents("C:/Data/rates.csv")', read_m(b, 'rates.csv'))
        self.assertIn('Sql.Database("sqlprod.example.local", "SALES"){[Schema = "dbo", Item = "ORDERS"]}[Data]', read_m(b, 'ORDERS'))
        self.assertIn('if UseSampleData then SampleTable("ORDERS.csv") else', read_m(b, 'ORDERS'))  # sample switch kept
        self.assertEqual(mc.check(sm)['errors'], [])

    def test_a_configured_warehouse_keeps_the_warehouse(self):
        b, self.root, _ = build('src_uc1_connectors', {'mode': 'warehouse', 'warehouse': {'server': 'wh.example', 'database': 'DW'}})
        self.assertEqual(b.connector_mode(), 'warehouse')
        self.assertIn('SourceTable("VW_SALES", "VW_SALES.csv")', read_m(b, 'VW_SALES'))
        self.assertNotIn('Snowflake.Databases', read_m(b, 'VW_SALES'))

    def test_extractor_keeps_what_the_connectors_need(self):
        root, _ = bx.load_root(os.path.join(FIX, 'src_uc1_connectors.twb'))
        named = {n['class']: n for d in bx.Workbook(root).build()['datasources'] for n in d['connection'].get('named', [])}
        self.assertEqual(named['snowflake']['warehouse'], 'REPORTING_WH')
        self.assertEqual(named['excel-direct']['filename'], 'C:/Data/Budget.xlsx')
        self.assertEqual((named['textscan']['directory'], named['textscan']['filename']), ('C:/Data', 'rates.csv'))


class DataSourceFilters(unittest.TestCase):
    def tearDown(self):
        shutil.rmtree(getattr(self, 'root', ''), ignore_errors=True)

    def test_rows_are_filtered_in_power_query(self):
        b, self.root, sm = build('src_uc2_ds_filters')
        m = read_m(b, 'ORDERS')
        self.assertIn('List.Contains({"East", "West"}, [Region])', m)
        self.assertIn('[#"Order Date"] >= #date(2024, 1, 1) and [#"Order Date"] <= #date(2025, 12, 31)', m)  # a column no sheet uses
        self.assertIn('[Sales] <> null', m)
        self.assertIn('List.Contains({"Domestic"}, [Zone])', read_m(b, 'REGIONS'))
        self.assertEqual(mc.check(sm)['errors'], [])

    def test_a_filter_on_the_one_side_also_filters_every_page(self):
        b, self.root, _ = build('src_uc2_ds_filters')
        pf = [f['field']['name'] for p in b.spec['pages'] for f in p.get('page_filters', [])]
        self.assertIn('Zone', pf)  # REGIONS filters ORDERS: the facts must follow
        self.assertNotIn('Region', pf)  # ORDERS filters nothing: Power Query alone is enough

    def test_extractor_counts_filter_columns_as_used(self):
        root, _ = bx.load_root(os.path.join(FIX, 'src_uc2_ds_filters.twb'))
        w = bx.Workbook(root)
        self.assertIn(('federated.dsf1', 'Order Date'), w.used)


class Aliases(unittest.TestCase):
    def tearDown(self):
        shutil.rmtree(getattr(self, 'root', ''), ignore_errors=True)

    def test_visuals_show_the_alias_and_filter_the_source_value(self):
        b, self.root, sm = build('src_uc3_aliases')
        self.assertIn('Record.FieldOrDefault([East = "Eastern Region", West = "Western Region"], Text.From([Region]), '
                      'Text.From([Region]))', read_m(b, 'ORDERS'))
        col = next(c for c in b.tables['ORDERS']['columns'] if c['name'] == 'Region (alias)')
        self.assertEqual(col['sort_by'], 'Region')  # one alias per value: sorts like the source column
        vj = glob.glob(os.path.join(self.root, 'o', 'src_uc3_aliases', 'pbip', '*.Report', 'definition', 'pages', '*',
                                    'visuals', '*', 'visual.json'))[0]
        with open(vj, encoding='utf-8') as f:
            v = json.load(f)
        self.assertIn('Region (alias)', json.dumps(v['visual']['query']))
        flt = v['filterConfig']['filters']  # the worksheet-tab page keeps its sheet filter, on the source value
        self.assertEqual([x['field']['Column']['Property'] for x in flt], ['Region'])
        self.assertIn("'East'", json.dumps(flt))
        self.assertEqual(mc.check(sm)['errors'], [])


if __name__ == '__main__':
    unittest.main()
