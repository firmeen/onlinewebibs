"""Write valid .xlsx files with Python's standard library only.

This intentionally needs no Excel package on the collector host. Excel's
floating-point precision is limited, so metric_raw remains an exact string.
"""
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZipFile, ZIP_DEFLATED
import math

URI_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
URI_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
URI_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'


def xml_text(value):
    s = str(value)
    s = ''.join(c for c in s if c in '\t\n\r' or ord(c) >= 32)
    return escape(s, {'"':'&quot;', "'":'&apos;'})


def column_letter(index):
    parts = ''
    while index:
        index, rem = divmod(index-1, 26)
        parts = chr(65+rem) + parts
    return parts


def make_content_types(count):
    sheets=''.join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
                   f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                   for i in range(1,count+1))
    return ('''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
''' + sheets + '</Types>')


def make_styles():
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="{URI_MAIN}">
<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.##########;[Red](#,##0.##########)"/></numFmts>
<fonts count="2">
<font><sz val="11"/><name val="Aptos"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font>
</fonts>
<fills count="3">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF102F38"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''


def write_cell(zipfile_stream, ref, value, is_header=False, numeric=False):
    if value is None or value == '':
        return
    if numeric and isinstance(value,(float,int)) and math.isfinite(value):
        zipfile_stream.write(f'<c r="{ref}" s="2"><v>{format(value, ".15g")}</v></c>'.encode('utf-8'))
    else:
        sty = ' s="1"' if is_header else ''
        content = xml_text(value)
        zipfile_stream.write(f'<c r="{ref}" t="inlineStr"{sty}><is><t xml:space="preserve">{content}</t></is></c>'.encode('utf-8'))


def write_sheet(z, sheet_path, fields, rows):
    endrow = len(rows) + 1
    endcol = column_letter(len(fields))
    with z.open(sheet_path, 'w', force_zip64=True) as stream:
        start = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 f'<worksheet xmlns="{URI_MAIN}">'
                 f'<dimension ref="A1:{endcol}{endrow}"/>'
                 f'<sheetViews><sheetView workbookViewId="0">'
                 f'<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
                 f'</sheetView></sheetViews><sheetFormatPr defaultRowHeight="17"/>'
                 f'<cols>'
                 f'<col min="1" max="1" width="18" customWidth="1"/>'
                 f'<col min="2" max="5" width="19" customWidth="1"/>'
                 f'<col min="6" max="14" width="27" customWidth="1"/>'
                 f'<col min="15" max="16" width="21" customWidth="1"/>'
                 f'<col min="17" max="18" width="23" customWidth="1"/>'
                 f'<col min="19" max="20" width="22" customWidth="1"/>'
                 f'</cols><sheetData>')
        stream.write(start.encode('utf-8'))
        stream.write(b'<row r="1" ht="24" customHeight="1">')
        for j,field in enumerate(fields,1):
            write_cell(stream,column_letter(j)+'1',field,is_header=True)
        stream.write(b'</row>')
        for i,item in enumerate(rows,2):
            stream.write(f'<row r="{i}">'.encode('utf-8'))
            for j,field in enumerate(fields,1):
                val=item.get(field)
                is_metric=(field == 'metric_value')
                if is_metric:
                    val=float(val) if val is not None else None
                write_cell(stream,column_letter(j)+str(i),val,numeric=is_metric)
            stream.write(b'</row>')
        stream.write(f'</sheetData><autoFilter ref="A1:{endcol}{endrow}"/></worksheet>'.encode('utf-8'))


def write_excel(data_by_code, fields, out_file):
    out_file = Path(out_file)
    out_file.parent.mkdir(parents=True, exist_ok=True)
    codes = list(data_by_code)
    with ZipFile(out_file,mode='w',compression=ZIP_DEFLATED,compresslevel=6) as z:
        z.writestr('[Content_Types].xml',make_content_types(len(codes)))
        z.writestr('_rels/.rels',f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{URI_PKG}"><Relationship Id="rId1" Type="{URI_REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>''')
        sheets = ''.join(f'<sheet name="{xml_text(code)}" sheetId="{i}" r:id="rId{i}"/>' for i,code in enumerate(codes,1))
        z.writestr('xl/workbook.xml',f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="{URI_MAIN}" xmlns:r="{URI_REL}"><sheets>{sheets}</sheets></workbook>''')
        rels=''.join(f'<Relationship Id="rId{i}" Type="{URI_REL}/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1,len(codes)+1))
        rels += f'<Relationship Id="rId{len(codes)+1}" Type="{URI_REL}/styles" Target="styles.xml"/>'
        z.writestr('xl/_rels/workbook.xml.rels',f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{URI_PKG}">{rels}</Relationships>''')
        z.writestr('xl/styles.xml',make_styles())
        for i,code in enumerate(codes,1):
            write_sheet(z,f'xl/worksheets/sheet{i}.xml',fields,data_by_code[code])