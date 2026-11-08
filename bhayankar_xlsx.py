"""Bhayankar tester, Excel module: writes a styled .xlsx with the standard library only.

Each sheet gets a bold header, frozen first row, autofilter and column widths; cells in a column
named "Status" are coloured (PASS green, FAIL red, WARN amber, MANUAL blue, SKIP grey).
"""
import re
import zipfile
from xml.sax.saxutils import escape

STATUS_STYLE = {'PASS': 2, 'FAIL': 3, 'WARN': 4, 'MANUAL': 5, 'SKIP': 6, 'INFO': 7}
BAD_XML = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]')

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="3"><font><sz val="10"/><name val="Segoe UI"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/></font><font><b/><sz val="12"/><name val="Segoe UI"/></font></fonts>
<fills count="9"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFEDEDED"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF7F7F7"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="9">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top"/></xf>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top"/></xf>
<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top"/></xf>
<xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top"/></xf>
<xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top"/></xf>
<xf numFmtId="0" fontId="0" fillId="7" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top"/></xf>
<xf numFmtId="0" fontId="0" fillId="8" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs></styleSheet>'''


def col_letter(i):
    s = ''
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def cell(ref, v, style=0):
    st = f' s="{style}"' if style else ''
    if v is None or v == '':
        return f'<c r="{ref}"{st}/>' if style else ''
    if isinstance(v, bool):
        v = 'TRUE' if v else 'FALSE'
    if isinstance(v, (int, float)) and v == v and abs(v) != float('inf'):
        return f'<c r="{ref}"{st}><v>{repr(float(v)) if isinstance(v, float) else v}</v></c>'
    s = BAD_XML.sub('', str(v))[:32000]
    return f'<c r="{ref}"{st} t="inlineStr"><is><t xml:space="preserve">{escape(s)}</t></is></c>'


def sheet_xml(headers, rows, widths=None, title_rows=0):
    status_col = next((i for i, h in enumerate(headers) if h == 'Status'), None)
    widths = widths or []
    calc = []
    for i, h in enumerate(headers):
        w = max([len(str(h))] + [len(str(r[i])) if i < len(r) and r[i] is not None else 0 for r in rows[:2000]])
        calc.append(min(max(widths[i] if i < len(widths) else w + 2, 8), 70))
    cols = ''.join(f'<col min="{i + 1}" max="{i + 1}" width="{w}" customWidth="1"/>' for i, w in enumerate(calc))
    out = []
    hdr = ''.join(cell(f'{col_letter(i)}1', h, 1) for i, h in enumerate(headers))
    out.append(f'<row r="1">{hdr}</row>')
    for n, r in enumerate(rows, start=2):
        cs = []
        for i, v in enumerate(r):
            style = 0
            if i == status_col and isinstance(v, str):
                style = STATUS_STYLE.get(v.split()[0] if v else '', 0)
            cs.append(cell(f'{col_letter(i)}{n}', v, style))
        out.append(f'<row r="{n}">{"".join(cs)}</row>')
    last = f'{col_letter(max(len(headers) - 1, 0))}{max(len(rows) + 1, 1)}'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            '</sheetView></sheetViews>'
            f'<cols>{cols}</cols><sheetData>{"".join(out)}</sheetData>'
            + (f'<autoFilter ref="A1:{last}"/>' if rows else '') + '</worksheet>')


def safe_name(n, used):
    n = re.sub(r'[\[\]\*\?/\\:]', ' ', n)[:31].strip() or 'Sheet'
    base, i = n, 2
    while n.lower() in used:
        n = f'{base[:28]} {i}'
        i += 1
    used.add(n.lower())
    return n


def defined_names(names, sheets):
    """Excel expects one hidden _FilterDatabase name per sheet that has an autofilter."""
    items = [f'<definedName name="_xlnm._FilterDatabase" localSheetId="{i}" hidden="1">'
             f"{escape(chr(39) + n.replace(chr(39), chr(39) * 2) + chr(39))}!$A$1:${col_letter(len(s[1]) - 1)}${len(s[2]) + 1}</definedName>"
             for i, (n, s) in enumerate(zip(names, sheets)) if s[2]]
    return f'<definedNames>{"".join(items)}</definedNames>' if items else ''


def write_xlsx(path, sheets):
    """sheets: list of (name, headers, rows[, widths])."""
    used = set()
    names = [safe_name(s[0], used) for s in sheets]
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                   + ''.join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" '
                             'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                             for i in range(len(sheets))) + '</Types>')
        z.writestr('_rels/.rels',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                   '</Relationships>')
        z.writestr('xl/workbook.xml',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
                   + ''.join(f'<sheet name="{escape(n, {chr(34): "&quot;"})}" sheetId="{i + 1}" r:id="rId{i + 1}"/>'
                             for i, n in enumerate(names))
                   + '</sheets>' + defined_names(names, sheets) + '</workbook>')
        z.writestr('xl/_rels/workbook.xml.rels',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   + ''.join(f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                             f'Target="worksheets/sheet{i + 1}.xml"/>' for i in range(len(sheets)))
                   + f'<Relationship Id="rId{len(sheets) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                   '</Relationships>')
        z.writestr('xl/styles.xml', STYLES)
        for i, s in enumerate(sheets):
            z.writestr(f'xl/worksheets/sheet{i + 1}.xml', sheet_xml(s[1], s[2], s[3] if len(s) > 3 else None))
