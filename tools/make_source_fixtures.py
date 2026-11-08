"""Writes the source-connector / data-source-filter / value-alias use-case workbooks into argv[1] (synthetic TWB XML,
no data; server and file names are made up).

  src_uc1_connectors   four data sources: Snowflake (DB.SCHEMA.VIEW), Excel sheet, text file, SQL Server
  src_uc2_ds_filters   data source filters: Region IN (East, West), Order Date range, Sales non-null, and a filter on a
                       column of REGIONS, the one side of a relationship (so the facts must follow it)
  src_uc3_aliases      value aliases on Region ("East" shown as "Eastern Region") with a sheet filter on the raw value
"""
import os
import sys

HEAD = "<?xml version='1.0' encoding='utf-8' ?>\n<!-- Bhayankar source fixture: synthetic workbook, not real data -->\n" \
       "<workbook source-build='2025.1.0' version='18.1' xmlns:user='http://www.tableausoftware.com/xml/user'>\n"


def ds(name, caption, conn_xml, relation, table, cols, extra='', object_graph=''):
    """cols: [(column, datatype, role, table, alias xml)]"""
    L = [f"    <datasource caption='{caption}' inline='true' name='{name}' version='18.1'>",
         "      <connection class='federated'>", f'        <named-connections>{conn_xml}</named-connections>',
         f'        {relation}', '        <metadata-records>']
    L += [f"          <metadata-record class='column'><remote-name>{c}</remote-name><local-name>[{c}]</local-name>"
          f"<parent-name>[{t or table}]</parent-name><local-type>{dt}</local-type></metadata-record>" for c, dt, _, t, _ in cols]
    L += ['        </metadata-records>', '      </connection>']
    for c, dt, role, _, alias in cols:
        typ = 'quantitative' if role == 'measure' else 'nominal'
        agg = " aggregation='Sum'" if role == 'measure' else ''
        L.append(f"      <column{agg} datatype='{dt}' name='[{c}]' role='{role}' type='{typ}'>{alias}</column>")
    L += [extra, object_graph, '    </datasource>']
    return '\n'.join(x for x in L if x)


def sheet(name, dsname, caption, dims, measures, filters=''):
    inst = [f"<column-instance column='[{d}]' derivation='None' name='[none:{d}:nk]' pivot='key' type='nominal' />" for d in dims] + \
           [f"<column-instance column='[{m}]' derivation='Sum' name='[sum:{m}:qk]' pivot='key' type='quantitative' />" for m in measures]
    rows = ' / '.join(f'[{dsname}].[none:{d}:nk]' for d in dims)
    cols = ' + '.join(f'[{dsname}].[sum:{m}:qk]' for m in measures)
    return f"""    <worksheet name='{name}'>
      <table>
        <view>
          <datasources><datasource caption='{caption}' name='{dsname}' /></datasources>
          <datasource-dependencies datasource='{dsname}'>{''.join(inst)}</datasource-dependencies>
{filters}          <aggregation value='true' />
        </view>
        <style />
        <panes><pane><view><breakdown value='auto' /></view><mark class='Bar' /></pane></panes>
        <rows>{rows}</rows>
        <cols>{cols}</cols>
      </table>
    </worksheet>"""


def workbook(datasources, sheets):
    return '\n'.join([HEAD, '  <datasources>', *datasources, '  </datasources>', '  <worksheets>', *sheets, '  </worksheets>',
                      '  <windows>', *[f"    <window class='worksheet' name='{n}' />" for n in
                                       [s.split("name='")[1].split("'")[0] for s in sheets]], '  </windows>', '</workbook>', ''])


def uc1():
    sf = ds('federated.sf1', 'Snowflake Sales',
            "<named-connection caption='acct' name='snowflake.0a'><connection class='snowflake' dbname='SALESDB' schema='MART' "
            "server='acct.example.snowflakecomputing.com' warehouse='REPORTING_WH' /></named-connection>",
            "<relation connection='snowflake.0a' name='VW_SALES' table='[SALESDB].[MART].[VW_SALES]' type='table' />", 'VW_SALES',
            [('Region', 'string', 'dimension', None, ''), ('Amount', 'real', 'measure', None, '')])
    xl = ds('federated.xl1', 'Budget Workbook',
            "<named-connection caption='Budget' name='excel-direct.0b'><connection class='excel-direct' filename='C:/Data/Budget.xlsx' />"
            "</named-connection>",
            "<relation connection='excel-direct.0b' name='Budget' table='[Budget$]' type='table' />", 'Budget',
            [('Dept', 'string', 'dimension', None, ''), ('Budget', 'real', 'measure', None, '')])
    tx = ds('federated.tx1', 'Rates CSV',
            "<named-connection caption='rates' name='textscan.0c'><connection class='textscan' directory='C:/Data' filename='rates.csv' />"
            "</named-connection>",
            "<relation connection='textscan.0c' name='rates.csv' table='[rates#csv]' type='table' />", 'rates.csv',
            [('Currency', 'string', 'dimension', None, ''), ('Rate', 'real', 'measure', None, '')])
    sq = ds('federated.sq1', 'SQL Orders',
            "<named-connection caption='sqlprod' name='sqlserver.0d'><connection class='sqlserver' dbname='SALES' schema='dbo' "
            "server='sqlprod.example.local' /></named-connection>",
            "<relation connection='sqlserver.0d' name='ORDERS' table='[dbo].[ORDERS]' type='table' />", 'ORDERS',
            [('Category', 'string', 'dimension', None, ''), ('Sales', 'real', 'measure', None, '')])
    return workbook([sf, xl, tx, sq], [sheet('Snowflake Sheet', 'federated.sf1', 'Snowflake Sales', ['Region'], ['Amount']),
                                       sheet('Excel Sheet', 'federated.xl1', 'Budget Workbook', ['Dept'], ['Budget']),
                                       sheet('CSV Sheet', 'federated.tx1', 'Rates CSV', ['Currency'], ['Rate']),
                                       sheet('SQL Sheet', 'federated.sq1', 'SQL Orders', ['Category'], ['Sales'])])


def uc2():
    n = 'federated.dsf1'
    filters = (f"      <filter class='categorical' column='[{n}].[Region]'><groupfilter function='union' user:ui-enumeration='inclusive'>"
               f"<groupfilter function='member' level='[Region]' member='&quot;East&quot;' />"
               f"<groupfilter function='member' level='[Region]' member='&quot;West&quot;' /></groupfilter></filter>\n"
               f"      <filter class='quantitative' column='[{n}].[Order Date]' included-values='in-range'><min>#2024-01-01#</min>"
               f"<max>#2025-12-31#</max></filter>\n"
               f"      <filter class='quantitative' column='[{n}].[Sales]' included-values='non-null' />\n"
               f"      <filter class='categorical' column='[{n}].[Zone]'><groupfilter function='member' level='[Zone]' "
               f"member='&quot;Domestic&quot;' user:ui-enumeration='inclusive' /></filter>")
    og = ("      <object-graph><objects><object caption='ORDERS' id='O_1'><properties context=''><relation connection='sqlserver.0d' "
          "name='ORDERS' table='[dbo].[ORDERS]' type='table' /></properties></object><object caption='REGIONS' id='R_1'>"
          "<properties context=''><relation connection='sqlserver.0d' name='REGIONS' table='[dbo].[REGIONS]' type='table' />"
          "</properties></object></objects><relationships><relationship><expression op='='><expression op='[Region]' />"
          "<expression op='[Region (REGIONS)]' /></expression><first-end-point object-id='O_1' />"
          "<second-end-point object-id='R_1' unique-key='true' /></relationship></relationships></object-graph>")
    d = ds(n, 'Filtered Orders',
           "<named-connection caption='sqlprod' name='sqlserver.0d'><connection class='sqlserver' dbname='SALES' schema='dbo' "
           "server='sqlprod.example.local' /></named-connection>",
           "<relation type='collection'><relation connection='sqlserver.0d' name='ORDERS' table='[dbo].[ORDERS]' type='table' />"
           "<relation connection='sqlserver.0d' name='REGIONS' table='[dbo].[REGIONS]' type='table' /></relation>"
           "<cols><map key='[Region]' value='[ORDERS].[Region]' /><map key='[Order Date]' value='[ORDERS].[Order Date]' />"
           "<map key='[Sales]' value='[ORDERS].[Sales]' /><map key='[Region (REGIONS)]' value='[REGIONS].[Region]' />"
           "<map key='[Zone]' value='[REGIONS].[Zone]' /></cols>", 'ORDERS',
           [('Region', 'string', 'dimension', 'ORDERS', ''), ('Order Date', 'date', 'dimension', 'ORDERS', ''),
            ('Sales', 'real', 'measure', 'ORDERS', ''), ('Region (REGIONS)', 'string', 'dimension', 'REGIONS', ''),
            ('Zone', 'string', 'dimension', 'REGIONS', '')], extra=filters, object_graph=og)
    d = d.replace('<remote-name>Region (REGIONS)</remote-name>', '<remote-name>Region</remote-name>')
    return workbook([d], [sheet('Sales by Zone', n, 'Filtered Orders', ['Zone', 'Region'], ['Sales'])])


def uc3():
    n = 'federated.al1'
    alias = ("<aliases><alias key='&quot;East&quot;' value='Eastern Region' /><alias key='&quot;West&quot;' value='Western Region' />"
             "</aliases>")
    d = ds(n, 'Aliased Orders',
           "<named-connection caption='sqlprod' name='sqlserver.0d'><connection class='sqlserver' dbname='SALES' schema='dbo' "
           "server='sqlprod.example.local' /></named-connection>",
           "<relation connection='sqlserver.0d' name='ORDERS' table='[dbo].[ORDERS]' type='table' />", 'ORDERS',
           [('Region', 'string', 'dimension', None, alias), ('Sales', 'real', 'measure', None, '')])
    flt = (f"          <filter class='categorical' column='[{n}].[none:Region:nk]'><groupfilter function='member' "
           f"level='[none:Region:nk]' member='&quot;East&quot;' user:ui-enumeration='inclusive' /></filter>\n")
    return workbook([d], [sheet('Sales by Region', n, 'Aliased Orders', ['Region'], ['Sales'], flt)])


if __name__ == '__main__':
    out = sys.argv[1]
    os.makedirs(out, exist_ok=True)
    for fn, f in (('src_uc1_connectors', uc1), ('src_uc2_ds_filters', uc2), ('src_uc3_aliases', uc3)):
        with open(os.path.join(out, fn + '.twb'), 'w', encoding='utf-8') as fh:
            fh.write(f())
        print('wrote', fn)
