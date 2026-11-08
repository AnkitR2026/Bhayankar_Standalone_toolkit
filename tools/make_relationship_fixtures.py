"""Writes the Tableau relationship (logical layer, 2020.2+) use-case workbooks into the folder given as argv[1].
Synthetic TWB XML, no data. The shape follows the real relationships in "Business Scenario @ Superstore_v2026.1.twbx":
a federated data source with <relation type='collection'>, <cols><map key='[Region (PEOPLE)]' value='[PEOPLE].[Region]'/>
and <object-graph><relationships><relationship><expression op='='><expression op='[Region]'/>
<expression op='[Region (PEOPLE)]'/></expression><first-end-point object-id=.../><second-end-point object-id=.../>.

  rel_uc1_star          ORDERS -> PEOPLE (Region), ORDERS -> RETURNS (Order ID): Tableau's own key format
  rel_uc2_composite     ORDERS -> TARGETS on Region AND Category (compound relationship)
  rel_uc3_cardinality   Performance options: PRODUCTS key unique ("One"), referential integrity "All records match"
  rel_uc4_snowflake     ORDERS -> PRODUCTS -> CATEGORIES (a chain)
  rel_uc5_calc_key      a relationship on a calculated field: [Order Year] = [Year] (Order Year = YEAR([Order Date]))
  rel_uc6_inequality    ORDERS -> PRICE_BANDS on [Sales] >= [Band Min] (comparison operator)
"""
import os
import sys

HEAD = "<?xml version='1.0' encoding='utf-8' ?>\n<!-- Bhayankar relationship fixture: synthetic workbook, not real data -->\n" \
       "<workbook source-build='2025.1.0' version='18.1' xmlns:user='http://www.tableausoftware.com/xml/user'>\n"
DS = 'federated.rel0model1'

TABLES = {
    'ORDERS': [('Order ID', 'string'), ('Region', 'string'), ('Category', 'string'), ('Product ID', 'string'),
               ('Order Date', 'date'), ('Sales', 'real')],
    'PEOPLE': [('Region', 'string'), ('Regional Manager', 'string')],
    'RETURNS': [('Order ID', 'string'), ('Returned', 'string')],
    'TARGETS': [('Region', 'string'), ('Category', 'string'), ('Target', 'real')],
    'PRODUCTS': [('Product ID', 'string'), ('Product Name', 'string'), ('Category ID', 'string')],
    'CATEGORIES': [('Category ID', 'string'), ('Category Name', 'string')],
    'YEAR_GOALS': [('Year', 'integer'), ('Goal', 'real')],
    'PRICE_BANDS': [('Band Min', 'real'), ('Band', 'string')],
}
MEASURES = {'Sales', 'Target', 'Goal'}


def field(table, col, first):
    """Tableau's local field name: the first table keeps the column name, a later clash gets ' (TABLE)'."""
    return col if first else f'{col} ({table})'


def esc(s):
    return s.replace('&', '&amp;').replace("'", '&apos;').replace('<', '&lt;').replace('>', '&gt;')


def workbook(tables, rels, sheet_fields, calcs=()):
    """tables: model order; rels: [(t1, t2, [(f1, op, f2)], end-point attrs dict for second end)];
    sheet_fields: [(local field, 'dim'|'sum')] on the sheet."""
    seen, fields = set(), {}  # local field -> (table, column, datatype)
    for t in tables:
        for c, dt in TABLES[t]:
            f = field(t, c, c not in seen)
            seen.add(c)
            fields[f] = (t, c, dt)
    L = [HEAD, '  <datasources>', f"    <datasource caption='Sales Model' inline='true' name='{DS}' version='18.1'>",
         "      <connection class='federated'>",
         "        <named-connections><named-connection caption='sqlprod' name='sqlserver.0aa'><connection class='sqlserver' "
         "dbname='SALES' schema='dbo' server='sqlprod' /></named-connection></named-connections>",
         "        <relation type='collection'>"]
    L += [f"          <relation connection='sqlserver.0aa' name='{t}' table='[dbo].[{t}]' type='table' />" for t in tables]
    L += ['        </relation>', '        <cols>']
    L += [f"          <map key='[{esc(f)}]' value='[{t}].[{c}]' />" for f, (t, c, _) in fields.items()]
    L += ['        </cols>', '        <metadata-records>']
    L += [f"          <metadata-record class='column'><remote-name>{c}</remote-name><local-name>[{esc(f)}]</local-name>"
          f"<parent-name>[{t}]</parent-name><local-type>{dt}</local-type></metadata-record>" for f, (t, c, dt) in fields.items()]
    L += ['        </metadata-records>', '      </connection>']
    for f, (t, c, dt) in fields.items():
        role = 'measure' if c in MEASURES else 'dimension'
        typ = 'quantitative' if role == 'measure' else ('ordinal' if dt in ('date', 'integer') else 'nominal')
        L.append(f"      <column{' aggregation=' + chr(39) + 'Sum' + chr(39) if role == 'measure' else ''} datatype='{dt}' "
                 f"name='[{esc(f)}]' role='{role}' type='{typ}' />")
    for name, cap, dt, formula in calcs:
        L.append(f"      <column caption='{cap}' datatype='{dt}' name='[{name}]' role='dimension' type='ordinal'>"
                 f"<calculation class='tableau' formula='{esc(formula)}' /></column>")
    L += ['      <object-graph>', '        <objects>']
    L += [f"          <object caption='{t}' id='{t}_0A1B'><properties context=''><relation connection='sqlserver.0aa' "
          f"name='{t}' table='[dbo].[{t}]' type='table' /></properties></object>" for t in tables]
    L += ['        </objects>', '        <relationships>']
    for t1, t2, pairs, second in rels:
        L.append('          <relationship>')
        exprs = [f"<expression op='{esc(op)}'><expression op='[{esc(a)}]' /><expression op='[{esc(b)}]' /></expression>"
                 for a, op, b in pairs]
        L.append('            ' + (exprs[0] if len(exprs) == 1 else "<expression op='AND'>" + ''.join(exprs) + '</expression>'))
        L.append(f"            <first-end-point object-id='{t1}_0A1B' />")
        attrs = ''.join(f" {k}='{v}'" for k, v in sorted(second.items()))
        L.append(f"            <second-end-point object-id='{t2}_0A1B'{attrs} />")
        L.append('          </relationship>')
    L += ['        </relationships>', '      </object-graph>', '    </datasource>', '  </datasources>']
    # one sheet: every listed field on Rows (dims) or Columns (SUM)
    inst = []
    for f, kind in sheet_fields:
        n = f if f in fields else next(x[0] for x in calcs if x[1] == f)
        inst.append((n, kind, (f'none:{n}:nk' if kind == 'dim' else f'sum:{n}:qk')))
    deps = [f"            <column-instance column='[{esc(n)}]' derivation='{'None' if k == 'dim' else 'Sum'}' "
            f"name='[{esc(i)}]' pivot='key' type='{'nominal' if k == 'dim' else 'quantitative'}' />" for n, k, i in inst]
    rows = ' / '.join(f'[{DS}].[{esc(i)}]' for _, k, i in inst if k == 'dim')
    cols = ' + '.join(f'[{DS}].[{esc(i)}]' for _, k, i in inst if k == 'sum')
    L += ['  <worksheets>', "    <worksheet name='Model Check'>", '      <table>', '        <view>',
          f"          <datasources><datasource caption='Sales Model' name='{DS}' /></datasources>",
          f"          <datasource-dependencies datasource='{DS}'>", *deps, '          </datasource-dependencies>',
          "          <aggregation value='true' />", '        </view>', '        <style />',
          "        <panes><pane><view><breakdown value='auto' /></view><mark class='Text' /></pane></panes>",
          f'        <rows>{rows}</rows>', f'        <cols>{cols}</cols>', '      </table>', '    </worksheet>', '  </worksheets>',
          "  <windows>\n    <window class='worksheet' name='Model Check' />\n  </windows>\n</workbook>\n"]
    return '\n'.join(L)


CASES = {
    'rel_uc1_star': (['ORDERS', 'PEOPLE', 'RETURNS'],
                     [('ORDERS', 'PEOPLE', [('Region', '=', 'Region (PEOPLE)')], {}),
                      ('ORDERS', 'RETURNS', [('Order ID', '=', 'Order ID (RETURNS)')], {})],
                     [('Regional Manager', 'dim'), ('Returned', 'dim'), ('Sales', 'sum')]),
    'rel_uc2_composite': (['ORDERS', 'TARGETS'],
                          [('ORDERS', 'TARGETS', [('Region', '=', 'Region (TARGETS)'), ('Category', '=', 'Category (TARGETS)')], {})],
                          [('Region', 'dim'), ('Category', 'dim'), ('Sales', 'sum'), ('Target', 'sum')]),
    'rel_uc3_cardinality': (['ORDERS', 'PRODUCTS'],
                            [('ORDERS', 'PRODUCTS', [('Product ID', '=', 'Product ID (PRODUCTS)')],
                              {'unique-key': 'true', 'referential-integrity': 'all'})],
                            [('Product Name', 'dim'), ('Sales', 'sum')]),
    'rel_uc4_snowflake': (['ORDERS', 'PRODUCTS', 'CATEGORIES'],
                          [('ORDERS', 'PRODUCTS', [('Product ID', '=', 'Product ID (PRODUCTS)')], {}),
                           ('PRODUCTS', 'CATEGORIES', [('Category ID', '=', 'Category ID (CATEGORIES)')], {})],
                          [('Category Name', 'dim'), ('Product Name', 'dim'), ('Sales', 'sum')]),
    'rel_uc5_calc_key': (['ORDERS', 'YEAR_GOALS'],
                         [('ORDERS', 'YEAR_GOALS', [('Calculation_77', '=', 'Year')], {})],
                         [('Order Year', 'dim'), ('Sales', 'sum'), ('Goal', 'sum')],
                         [('Calculation_77', 'Order Year', 'integer', 'YEAR([Order Date])')]),
    'rel_uc6_inequality': (['ORDERS', 'PRICE_BANDS'],
                           [('ORDERS', 'PRICE_BANDS', [('Sales', '>=', 'Band Min')], {})],
                           [('Band', 'dim'), ('Sales', 'sum')]),
}

if __name__ == '__main__':
    out = sys.argv[1]
    os.makedirs(out, exist_ok=True)
    for name, case in CASES.items():
        with open(os.path.join(out, name + '.twb'), 'w', encoding='utf-8') as fh:
            fh.write(workbook(*case))
        print('wrote', name)
