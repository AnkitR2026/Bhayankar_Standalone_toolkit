"""Tests for Tableau data blending (2026-11-07, 1.4.0): blend links read from the workbook, built as Power BI shared
dimensions. Fixtures: tests\\fixtures\\blending\\blend_uc*.twb (synthetic, no data), one per Tableau blending pattern of
help.tableau.com "Blend Your Data"; the XML shapes follow the real blend in Retail Performance Dashboard.twb.

Run (from the toolkit folder):  python -m unittest discover -s tests -v
"""
import glob
import json
import os
import re
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

FIX = os.path.join(HERE, 'fixtures', 'blending')


def build(name, edit_twb=None):
    """Extract + build one blending fixture in a temp folder: (Build, digest, spec, output folder)."""
    root = tempfile.mkdtemp(prefix='bhk_blend_')
    inp, out = os.path.join(root, 'i'), os.path.join(root, 'o')
    os.makedirs(inp)
    with open(os.path.join(FIX, name + '.twb'), encoding='utf-8') as f:
        text = f.read()
    with open(os.path.join(inp, name + '.twb'), 'w', encoding='utf-8') as f:
        f.write(edit_twb(text) if edit_twb else text)
    bx.run_folder(inp, out, table_map={})
    old = bb.INPUT_DIR, bb.OUTPUT_DIR
    bb.INPUT_DIR, bb.OUTPUT_DIR = inp, out
    try:
        b = bb.Build(name)
        b.run()
    finally:
        bb.INPUT_DIR, bb.OUTPUT_DIR = old
    with open(os.path.join(out, name, 'bhayankar_digest.json'), encoding='utf-8') as f:
        digest = json.load(f)
    return b, digest, root


def rels(b):
    return {(a, ka, t, kt) for a, ka, t, kt in b.relationships}


def visual_fields(root, name):
    """[(visualType, ['Entity[Property]', ...], [filter field 'Entity[Property]'])] of the built report."""
    out = []
    for vj in sorted(glob.glob(os.path.join(root, 'o', name, 'pbip', '*.Report', 'definition', 'pages', '*', 'visuals', '*', 'visual.json'))):
        with open(vj, encoding='utf-8') as f:
            v = json.load(f)
        q = json.dumps(v['visual'].get('query', {}))
        fields = re.findall(r'"Entity": "([^"]+)"\}\}, "Property": "([^"]+)"', q)
        flt = [(x['field'].get('Measure') or x['field'].get('Column'))['Property'] for x in v.get('filterConfig', {}).get('filters', [])]
        out.append((v['visual']['visualType'], [f'{e}[{p}]' for e, p in fields], flt))
    return out


class Extractor(unittest.TestCase):
    """The digest names the primary, the secondaries and the ACTIVE links of each blended sheet."""

    @classmethod
    def setUpClass(cls):
        cls.wb = {}
        for p in sorted(glob.glob(os.path.join(FIX, '*.twb'))):
            root, _ = bx.load_root(p)
            cls.wb[os.path.basename(p)[:-4]] = bx.Workbook(root).build()

    def links(self, name):
        return [(b['primary_table'], b['primary_field'], b['secondary_table'], b['secondary_field'], b['level'], b['custom'])
                for b in self.wb[name].get('blends', [])]

    def test_automatic_link_on_the_same_name(self):
        self.assertEqual(self.links('blend_uc1_auto_link'), [('ORDERS', 'Region', 'TARGETS', 'Region', None, False)])
        ws = self.wb['blend_uc1_auto_link']['worksheets'][0]
        self.assertEqual(ws['blend']['primary'], 'Orders')
        self.assertEqual(ws['blend']['secondaries'][0]['datasource'], 'Sales Targets')

    def test_custom_link_between_different_names(self):
        # Data > Edit Blend Relationships: <datasource-relationships> maps Region to Sales Territory
        self.assertEqual(self.links('blend_uc2_custom_link'), [('ORDERS', 'Region', 'QUOTA', 'Sales Territory', None, True)])
        lk = self.wb['blend_uc2_custom_link']['blends'][0]
        self.assertEqual((lk['primary_column'], lk['secondary_column']), ('Region', 'Sales Territory'))

    def test_two_links_one_at_month_level(self):
        self.assertEqual(sorted(self.links('blend_uc3_multi_link_month')),
                         [('ORDERS', 'Order Date', 'MONTHLY_TARGETS', 'Month Start', 'month', True),
                          ('ORDERS', 'Region', 'MONTHLY_TARGETS', 'Region', None, True)])

    def test_each_sheet_has_its_own_primary(self):
        ws = {w['name']: w['blend'] for w in self.wb['blend_uc6_filter_and_swap']['worksheets']}
        self.assertEqual(ws['Stretch Targets']['primary'], 'Orders')
        self.assertEqual(ws['Targets First']['primary'], 'Sales Targets')

    def test_shelf_and_filter_fields_carry_their_data_source(self):
        ws = self.wb['blend_uc6_filter_and_swap']['worksheets'][0]
        self.assertEqual([f['ds'] for f in ws['shelf']['rows']], ['Orders'])
        self.assertEqual(ws['filters'][0]['ds'], 'Sales Targets')
        self.assertEqual([c['datasource'] for c in self.wb['blend_uc1_auto_link']['columns'] if c['name'] == 'Region'],
                         ['Orders', 'Sales Targets'])

    def test_flag_says_how_the_sources_are_linked(self):
        det = [f['detail'] for f in self.wb['blend_uc2_custom_link']['flags'] if f['code'] == 'BLENDING']
        self.assertEqual(det, ['Orders, Regional Quota | primary Orders; links: Regional Quota on Region = Sales Territory'])

    def test_link_is_inactive_while_its_primary_field_is_not_in_the_view(self):
        # Tableau: an automatic link is active only once the primary field is used in the view
        p = os.path.join(FIX, 'blend_uc1_auto_link.twb')
        with open(p, encoding='utf-8') as f:
            text = f.read()
        text = text.replace("<rows>[federated.0orders1blend].[none:Region:nk]</rows>", '<rows></rows>')
        text = text.replace("<column-instance column='[Region]' derivation='None' name='[none:Region:nk]' pivot='key' type='nominal' />", '', 1)
        res = bx.Workbook(bx.ET.fromstring(text.encode('utf-8'))).build()
        self.assertNotIn('blends', res)
        self.assertIn('BLEND_NO_LINK', [f['code'] for f in res['flags']])

    def test_a_self_mapping_links_nothing(self):
        # Retail saves a <datasource-relationship> whose source and target are the same data source
        with open(os.path.join(FIX, 'blend_uc2_custom_link.twb'), encoding='utf-8') as f:
            text = f.read().replace("value='[federated.2quota3blend].[none:Sales Territory:nk]'",
                                    "value='[federated.0orders1blend].[none:Region:nk]'")
        res = bx.Workbook(bx.ET.fromstring(text.encode('utf-8'))).build()
        self.assertNotIn('blends', res)  # no custom pair left and no same-named field: no active link


class BuildSharedDimensions(unittest.TestCase):
    """The build relates blended tables through a shared dimension (or the Date table for a date link)."""

    def tearDown(self):
        shutil.rmtree(getattr(self, 'root', ''), ignore_errors=True)

    def test_auto_link_makes_a_shared_dimension(self):
        b, _, self.root = build('blend_uc1_auto_link')
        self.assertEqual(rels(b), {('ORDERS', 'Region', 'Region', 'Region'), ('TARGETS', 'Region', 'Region', 'Region')})
        dim = b.tables['Region']
        self.assertTrue(dim['blend_dim'])
        m = '\n'.join(dim['m_lines'])
        self.assertIn('Table.SelectColumns(ORDERS, {"Region"})', m)
        self.assertIn('Table.SelectColumns(TARGETS, {"Region"})', m)
        self.assertIn('Comparer.OrdinalIgnoreCase', m)  # Power BI keys ignore case
        self.assertTrue(next(c for c in b.tables['ORDERS']['columns'] if c['name'] == 'Region')['hidden'])
        vt, fields, flt = visual_fields(self.root, 'blend_uc1_auto_link')[0]
        self.assertIn('Region[Region]', fields)  # the axis is the shared dimension, not ORDERS or TARGETS
        self.assertNotIn('TARGETS[Region]', fields)
        self.assertIn('# Orders Rows', flt)  # Tableau left join: only the primary's marks
        self.assertFalse(any('ORPHAN' in x for x in b.log if isinstance(x, str)))

    def test_custom_link_renames_the_secondary_key(self):
        b, _, self.root = build('blend_uc2_custom_link')
        self.assertEqual(rels(b), {('ORDERS', 'Region', 'Region', 'Region'), ('QUOTA', 'Sales Territory', 'Region', 'Region')})
        self.assertIn('Table.RenameColumns(Table.SelectColumns(QUOTA, {"Sales Territory"}), {{"Sales Territory", "Region"}})',
                      '\n'.join(b.tables['Region']['m_lines']))

    def test_month_link_goes_through_the_date_table(self):
        b, _, self.root = build('blend_uc3_multi_link_month')
        r = rels(b)
        self.assertIn(('ORDERS', 'Order Date (Date)', 'Date', 'Date'), r)
        self.assertIn(('MONTHLY_TARGETS', 'Month Start (Date)', 'Date', 'Date'), r)
        self.assertIn(('MONTHLY_TARGETS', 'Region', 'Region', 'Region'), r)  # Date over both facts is not ambiguous
        _, fields, _ = visual_fields(self.root, 'blend_uc3_multi_link_month')[0]
        self.assertIn('Date[Year Month]', fields)
        self.assertIn('Region[Region]', fields)  # the colour legend too
        sm = glob.glob(os.path.join(self.root, 'o', 'blend_uc3_multi_link_month', 'pbip', '*.SemanticModel'))[0]
        res = mc.check(sm)
        self.assertEqual((res['errors'], res['warnings']), ([], []))

    def test_cross_source_calc_divides_per_region(self):
        b, _, self.root = build('blend_uc4_cross_source_calc')
        self.assertEqual(b.measures['% of Target']['dax'], "DIVIDE(SUM('ORDERS'[Sales]), SUM('TARGETS'[Target]))")
        self.assertIn(('TARGETS', 'Region', 'Region', 'Region'), rels(b))  # so SUM(Target) is per region

    def test_secondary_dimension_is_an_attr_measure(self):
        b, _, self.root = build('blend_uc5_secondary_dim')
        m = b.measures['Regional Manager (ATTR)']
        self.assertEqual(m['dax'], "IF(DISTINCTCOUNTNOBLANK('PEOPLE'[Regional Manager]) > 1, \"*\", SELECTEDVALUE('PEOPLE'[Regional Manager]))")
        _, fields, _ = visual_fields(self.root, 'blend_uc5_secondary_dim')[0]
        self.assertIn('PEOPLE[Regional Manager (ATTR)]', fields)  # was dropped from the visual
        self.assertNotIn('PEOPLE[Regional Manager]', fields)  # a column would cross every manager with every region

    def test_both_directions_share_one_dimension_and_each_sheet_keeps_its_primary(self):
        b, _, self.root = build('blend_uc6_filter_and_swap')
        self.assertEqual([t for t in b.tables if b.tables[t].get('blend_dim')], ['Region'])
        vis = visual_fields(self.root, 'blend_uc6_filter_and_swap')
        flts = sorted(tuple(f) for _, _, f in vis)
        self.assertEqual(flts, [('# Targets Rows',), ('Target Type', '# Orders Rows')])

    def test_tables_already_related_get_no_second_path(self):
        b, _, self.root = build('blend_uc1_auto_link')
        b.relationships = [('TARGETS', 'Region', 'ORDERS', 'Region')]  # TARGETS -> ORDERS exists already
        b.log = []
        b.blend_dimensions(b.tables)
        self.assertTrue(any('two filter paths' in x for x in b.log))


class ModelCheckLoops(unittest.TestCase):
    """LOOP warns only where a table is filtered along two paths, not for a shared dimension over two facts."""

    def model(self, rel_lines):
        self.d = tempfile.mkdtemp(prefix='bhk_mcl_')
        t = os.path.join(self.d, 'X.SemanticModel', 'definition', 'tables')
        os.makedirs(t)
        for name in ('A', 'B', 'C', 'D'):
            with open(os.path.join(t, f'{name}.tmdl'), 'w', encoding='utf-8') as f:
                f.write(f'table {name}\n\n\tcolumn K\n\t\tdataType: string\n\t\tsourceColumn: K\n\n')
        with open(os.path.join(t, '..', 'relationships.tmdl'), 'w', encoding='utf-8') as f:
            for i, (frm, to, both) in enumerate(rel_lines):
                f.write(f"relationship r{i}\n\tfromColumn: {frm}.K\n\ttoColumn: {to}.K\n"
                        + ('\tcrossFilteringBehavior: bothDirections\n' if both else '') + '\n')
        return mc.check(self.d)

    def tearDown(self):
        shutil.rmtree(getattr(self, 'd', ''), ignore_errors=True)

    def test_two_dimensions_over_two_facts_is_fine(self):
        res = self.model([('C', 'A', False), ('D', 'A', False), ('C', 'B', False), ('D', 'B', False)])
        self.assertEqual(res['warnings'], [])

    def test_two_filter_paths_warn(self):
        # A filters B and C, both filter D: D is reached from A twice
        res = self.model([('B', 'A', False), ('C', 'A', False), ('D', 'B', False), ('D', 'C', False)])
        self.assertEqual([w[0] for w in res['warnings']], ['LOOP'])

    def test_both_directions_make_a_loop(self):
        res = self.model([('C', 'A', True), ('D', 'A', False), ('C', 'B', False), ('D', 'B', False)])
        self.assertIn('LOOP', [w[0] for w in res['warnings']])


if __name__ == '__main__':
    unittest.main()
