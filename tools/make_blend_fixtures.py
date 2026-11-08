"""Writes the data-blending use-case workbooks (synthetic TWB XML, no data) into the folder given as argv[1].

Each workbook is one Tableau blending pattern from help.tableau.com "Blend Your Data":
  blend_uc1_auto_link        Sales vs target, link on a field with the same name (automatic link)
  blend_uc2_custom_link      fields with different names linked in Data > Edit Blend Relationships
  blend_uc3_multi_link_month two links: Region + the MONTH of a date (date link level)
  blend_uc4_cross_source_calc  a calc in the primary reading the secondary: SUM([Sales]) / SUM([Targets].[Target])
  blend_uc5_secondary_dim    a secondary dimension on the view (ATTR, one value per primary mark)
  blend_uc6_filter_and_swap  a filter on a secondary field + a second sheet where the secondary is the primary
The XML follows the real blend in Retail Performance Dashboard.twb: <datasource-relationships> at workbook level
(after <datasources>), <column-mapping><map key='[primary].[instance]' value='[secondary].[instance]'/>; the
linking field appears as a column-instance in the secondary's datasource-dependencies of the sheet.
"""
import os
import sys

HEAD = "<?xml version='1.0' encoding='utf-8' ?>\n<!-- Bhayankar blending fixture: synthetic workbook, not real data -->\n" \
       "<workbook source-build='2025.1.0' version='18.1' xmlns:user='http://www.tableausoftware.com/xml/user'>\n"

# name, caption, connection xml, relation name, table, [(col, datatype, role, type)]
SOURCES = {
    'orders': ('federated.0orders1blend', 'Orders',
               "<connection class='federated'><named-connections><named-connection caption='sqlprod' name='sqlserver.0aa'>"
               "<connection class='sqlserver' dbname='SALES' schema='dbo' server='sqlprod' /></named-connection></named-connections>"
               "<relation connection='sqlserver.0aa' name='ORDERS' table='[dbo].[ORDERS]' type='table' />",
               'ORDERS',
               [('Order ID', 'string', 'dimension', 'nominal'), ('Region', 'string', 'dimension', 'nominal'),
                ('Category', 'string', 'dimension', 'nominal'), ('Order Date', 'date', 'dimension', 'ordinal'),
                ('Sales', 'real', 'measure', 'quantitative')]),
    'targets': ('federated.1targets2blend', 'Sales Targets',
                "<connection class='federated'><named-connections><named-connection caption='Targets' name='excel-direct.0bb'>"
                "<connection class='excel-direct' filename='C:/Data/Targets.xlsx' /></named-connection></named-connections>"
                "<relation connection='excel-direct.0bb' name='TARGETS' table='[TARGETS$]' type='table' />",
                'TARGETS',
                [('Region', 'string', 'dimension', 'nominal'), ('Target Type', 'string', 'dimension', 'nominal'),
                 ('Target', 'real', 'measure', 'quantitative')]),
    'quota': ('federated.2quota3blend', 'Regional Quota',
              "<connection class='federated'><named-connections><named-connection caption='quota' name='textscan.0cc'>"
              "<connection class='textscan' directory='C:/Data' filename='QUOTA.csv' /></named-connection></named-connections>"
              "<relation connection='textscan.0cc' name='QUOTA' table='[QUOTA#csv]' type='table' />",
              'QUOTA',
              [('Sales Territory', 'string', 'dimension', 'nominal'), ('Quota', 'real', 'measure', 'quantitative')]),
    'monthly': ('federated.3monthly4blend', 'Monthly Targets',
                "<connection class='federated'><named-connections><named-connection caption='Monthly' name='excel-direct.0dd'>"
                "<connection class='excel-direct' filename='C:/Data/MonthlyTargets.xlsx' /></named-connection></named-connections>"
                "<relation connection='excel-direct.0dd' name='MONTHLY_TARGETS' table='[MONTHLY_TARGETS$]' type='table' />",
                'MONTHLY_TARGETS',
                [('Region', 'string', 'dimension', 'nominal'), ('Month Start', 'date', 'dimension', 'ordinal'),
                 ('Monthly Target', 'real', 'measure', 'quantitative')]),
    'people': ('federated.4people5blend', 'Region Managers',
               "<connection class='federated'><named-connections><named-connection caption='People' name='excel-direct.0ee'>"
               "<connection class='excel-direct' filename='C:/Data/People.xlsx' /></named-connection></named-connections>"
               "<relation connection='excel-direct.0ee' name='PEOPLE' table='[PEOPLE$]' type='table' />",
               'PEOPLE',
               [('Region', 'string', 'dimension', 'nominal'), ('Regional Manager', 'string', 'dimension', 'nominal')]),
}


def esc(s):
    return s.replace('&', '&amp;').replace("'", '&apos;').replace('<', '&lt;').replace('>', '&gt;')


def datasource(key, calcs=()):
    name, cap, conn, rel, cols = SOURCES[key]
    out = [f"    <datasource caption='{cap}' inline='true' name='{name}' version='18.1'>", '      ' + conn, '      <metadata-records>']
    for c, dt, _, _ in cols:
        out.append(f"        <metadata-record class='column'><remote-name>{c}</remote-name><local-name>[{c}]</local-name>"
                   f"<parent-name>[{rel}]</parent-name><local-type>{dt}</local-type></metadata-record>")
    out.append('      </metadata-records>')
    out.append('      </connection>')
    for c, dt, role, typ in cols:
        agg = " aggregation='Sum'" if role == 'measure' else ''
        out.append(f"      <column{agg} datatype='{dt}' name='[{c}]' role='{role}' type='{typ}' />")
    for cname, cap2, dt, role, formula in calcs:
        out.append(f"      <column caption='{cap2}' datatype='{dt}' name='[{cname}]' role='{role}' type='quantitative'>"
                   f"<calculation class='tableau' formula='{esc(formula)}' /></column>")
    out.append('    </datasource>')
    return '\n'.join(out)


def deps(key, cols, instances, calcs=()):
    """cols: [(col, datatype, role, type)]; instances: [(column, derivation, instance name, type)]"""
    name = SOURCES[key][0]
    out = [f"          <datasource-dependencies datasource='{name}'>"]
    for c, dt, role, typ in cols:
        out.append(f"            <column datatype='{dt}' name='[{c}]' role='{role}' type='{typ}' />")
    for cname, cap2, dt, role, formula in calcs:
        out.append(f"            <column caption='{cap2}' datatype='{dt}' name='[{cname}]' role='{role}' type='quantitative'>"
                   f"<calculation class='tableau' formula='{esc(formula)}' /></column>")
    for col, der, inst, typ in instances:
        out.append(f"            <column-instance column='[{col}]' derivation='{der}' name='[{inst}]' pivot='key' type='{typ}' />")
    out.append('          </datasource-dependencies>')
    return '\n'.join(out)


def col(key, c):
    return next(x for x in SOURCES[key][4] if x[0] == c)


def worksheet(name, primary, secondaries, dep_xml, rows, cols, mark='Bar', encodings='', filters=''):
    dss = ''.join(f"<datasource caption='{SOURCES[k][1]}' name='{SOURCES[k][0]}' />" for k in [primary] + secondaries)
    return f"""    <worksheet name='{name}'>
      <table>
        <view>
          <datasources>{dss}</datasources>
{dep_xml}
{filters}          <aggregation value='true' />
        </view>
        <style />
        <panes>
          <pane selection-relaxation-option='selection-relaxation-allow'>
            <view><breakdown value='auto' /></view>
            <mark class='{mark}' />
            <encodings>{encodings}</encodings>
          </pane>
        </panes>
        <rows>{rows}</rows>
        <cols>{cols}</cols>
      </table>
    </worksheet>"""


def relationships(maps):
    """maps: [(primary key, secondary key, [(primary instance, secondary instance)])]"""
    out = ['  <datasource-relationships>']
    for p, s, pairs in maps:
        out.append(f"    <datasource-relationship source='{SOURCES[p][0]}' target='{SOURCES[s][0]}'>")
        out.append('      <column-mapping>')
        for a, b in pairs:
            out.append(f"        <map key='[{SOURCES[p][0]}].[{a}]' value='[{SOURCES[s][0]}].[{b}]' />")
        out.append('      </column-mapping>')
        out.append('    </datasource-relationship>')
    out.append('  </datasource-relationships>')
    return '\n'.join(out)


def workbook(sources, sheets, rel_xml='', calcs=None):
    calcs = calcs or {}
    dash_zones = ''.join(f"<zone h='{100000 // len(sheets)}' id='{i + 2}' name='{s[0]}' w='100000' x='0' y='{i * (100000 // len(sheets))}' />"
                         for i, s in enumerate(sheets))
    parts = [HEAD, '  <datasources>', *[datasource(k, calcs.get(k, ())) for k in sources], '  </datasources>']
    if rel_xml:
        parts.append(rel_xml)
    parts += ['  <worksheets>', *[s[1] for s in sheets], '  </worksheets>',
              "  <dashboards>\n    <dashboard name='Blend'>\n      <size maxheight='800' maxwidth='1000' minheight='800' minwidth='1000' sizing-mode='fixed' />",
              f"      <zones><zone h='100000' id='1' type-v2='layout-basic' w='100000' x='0' y='0'>{dash_zones}</zone></zones>",
              '    </dashboard>\n  </dashboards>',
              "  <windows>\n    <window class='dashboard' name='Blend' />",
              *[f"    <window class='worksheet' hidden='true' name='{s[0]}' />" for s in sheets],
              '  </windows>\n</workbook>\n']
    return '\n'.join(parts)


O, T, Q, M, P = (SOURCES[k][0] for k in ('orders', 'targets', 'quota', 'monthly', 'people'))


def uc1():
    d = deps('orders', [col('orders', 'Region'), col('orders', 'Sales')],
             [('Region', 'None', 'none:Region:nk', 'nominal'), ('Sales', 'Sum', 'sum:Sales:qk', 'quantitative')]) + '\n' + \
        deps('targets', [col('targets', 'Region'), col('targets', 'Target')],
             [('Region', 'None', 'none:Region:nk', 'nominal'), ('Target', 'Sum', 'sum:Target:qk', 'quantitative')])
    ws = worksheet('Sales vs Target', 'orders', ['targets'], d, f'[{O}].[none:Region:nk]', f'[{O}].[sum:Sales:qk]',
                   encodings=f"<tooltip column='[{T}].[sum:Target:qk]' />")
    return workbook(['orders', 'targets'], [('Sales vs Target', ws)])


def uc2():
    d = deps('orders', [col('orders', 'Region'), col('orders', 'Sales')],
             [('Region', 'None', 'none:Region:nk', 'nominal'), ('Sales', 'Sum', 'sum:Sales:qk', 'quantitative')]) + '\n' + \
        deps('quota', [col('quota', 'Sales Territory'), col('quota', 'Quota')],
             [('Sales Territory', 'None', 'none:Sales Territory:nk', 'nominal'), ('Quota', 'Sum', 'sum:Quota:qk', 'quantitative')])
    ws = worksheet('Sales vs Quota', 'orders', ['quota'], d, f'[{O}].[none:Region:nk]', f'[{O}].[sum:Sales:qk]',
                   encodings=f"<tooltip column='[{Q}].[sum:Quota:qk]' />")
    return workbook(['orders', 'quota'], [('Sales vs Quota', ws)],
                    relationships([('orders', 'quota', [('none:Region:nk', 'none:Sales Territory:nk')])]))


def uc3():
    d = deps('orders', [col('orders', 'Region'), col('orders', 'Order Date'), col('orders', 'Sales')],
             [('Region', 'None', 'none:Region:nk', 'nominal'), ('Order Date', 'TruncMonth', 'tmn:Order Date:ok', 'ordinal'),
              ('Sales', 'Sum', 'sum:Sales:qk', 'quantitative')]) + '\n' + \
        deps('monthly', [col('monthly', 'Region'), col('monthly', 'Month Start'), col('monthly', 'Monthly Target')],
             [('Region', 'None', 'none:Region:nk', 'nominal'), ('Month Start', 'TruncMonth', 'tmn:Month Start:ok', 'ordinal'),
              ('Monthly Target', 'Sum', 'sum:Monthly Target:qk', 'quantitative')])
    ws = worksheet('Monthly Sales vs Target', 'orders', ['monthly'], d, f'[{O}].[sum:Sales:qk]', f'[{O}].[tmn:Order Date:ok]',
                   mark='Line', encodings=f"<color column='[{O}].[none:Region:nk]' /><tooltip column='[{M}].[sum:Monthly Target:qk]' />")
    return workbook(['orders', 'monthly'], [('Monthly Sales vs Target', ws)],
                    relationships([('orders', 'monthly', [('none:Region:nk', 'none:Region:nk'),
                                                          ('tmn:Order Date:ok', 'tmn:Month Start:ok')])]))


def uc4():
    calc = [('Calculation_901', '% of Target', 'real', 'measure', f'SUM([Sales]) / SUM([{T}].[Target])')]
    d = deps('orders', [col('orders', 'Region'), col('orders', 'Sales')],
             [('Region', 'None', 'none:Region:nk', 'nominal'), ('Calculation_901', 'User', 'usr:Calculation_901:qk', 'quantitative')],
             calcs=calc) + '\n' + \
        deps('targets', [col('targets', 'Region'), col('targets', 'Target')],
             [('Region', 'None', 'none:Region:nk', 'nominal')])
    ws = worksheet('Attainment', 'orders', ['targets'], d, f'[{O}].[none:Region:nk]', f'[{O}].[usr:Calculation_901:qk]')
    return workbook(['orders', 'targets'], [('Attainment', ws)], calcs={'orders': calc})


def uc5():
    d = deps('orders', [col('orders', 'Region'), col('orders', 'Sales')],
             [('Region', 'None', 'none:Region:nk', 'nominal'), ('Sales', 'Sum', 'sum:Sales:qk', 'quantitative')]) + '\n' + \
        deps('people', [col('people', 'Region'), col('people', 'Regional Manager')],
             [('Region', 'None', 'none:Region:nk', 'nominal'),
              ('Regional Manager', 'Attribute', 'attr:Regional Manager:nk', 'nominal')])
    ws = worksheet('Sales by Manager', 'orders', ['people'], d, f'([{O}].[none:Region:nk] / [{P}].[attr:Regional Manager:nk])',
                   f'[{O}].[sum:Sales:qk]')
    return workbook(['orders', 'people'], [('Sales by Manager', ws)])


def uc6():
    f1 = (f"          <filter class='categorical' column='[{T}].[none:Target Type:nk]'>\n"
          f"            <groupfilter function='member' level='[none:Target Type:nk]' member='&quot;Stretch&quot;' user:ui-enumeration='inclusive' />\n"
          f"          </filter>\n")
    d1 = deps('orders', [col('orders', 'Region'), col('orders', 'Sales')],
              [('Region', 'None', 'none:Region:nk', 'nominal'), ('Sales', 'Sum', 'sum:Sales:qk', 'quantitative')]) + '\n' + \
        deps('targets', [col('targets', 'Region'), col('targets', 'Target Type'), col('targets', 'Target')],
             [('Region', 'None', 'none:Region:nk', 'nominal'), ('Target Type', 'None', 'none:Target Type:nk', 'nominal'),
              ('Target', 'Sum', 'sum:Target:qk', 'quantitative')])
    ws1 = worksheet('Stretch Targets', 'orders', ['targets'], d1, f'[{O}].[none:Region:nk]', f'[{T}].[sum:Target:qk]',
                    encodings=f"<tooltip column='[{O}].[sum:Sales:qk]' />", filters=f1)
    # targets is the primary here: every region with a target shows, also those without orders
    d2 = deps('targets', [col('targets', 'Region'), col('targets', 'Target')],
              [('Region', 'None', 'none:Region:nk', 'nominal'), ('Target', 'Sum', 'sum:Target:qk', 'quantitative')]) + '\n' + \
        deps('orders', [col('orders', 'Region'), col('orders', 'Sales')],
             [('Region', 'None', 'none:Region:nk', 'nominal'), ('Sales', 'Sum', 'sum:Sales:qk', 'quantitative')])
    ws2 = worksheet('Targets First', 'targets', ['orders'], d2, f'[{T}].[none:Region:nk]', f'[{T}].[sum:Target:qk]',
                    encodings=f"<tooltip column='[{O}].[sum:Sales:qk]' />")
    return workbook(['orders', 'targets'], [('Stretch Targets', ws1), ('Targets First', ws2)])


if __name__ == '__main__':
    out = sys.argv[1]
    os.makedirs(out, exist_ok=True)
    for fn, f in (('blend_uc1_auto_link', uc1), ('blend_uc2_custom_link', uc2), ('blend_uc3_multi_link_month', uc3),
                  ('blend_uc4_cross_source_calc', uc4), ('blend_uc5_secondary_dim', uc5), ('blend_uc6_filter_and_swap', uc6)):
        with open(os.path.join(out, fn + '.twb'), 'w', encoding='utf-8') as fh:
            fh.write(f())
        print('wrote', fn)
