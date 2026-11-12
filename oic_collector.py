#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""OIC XMLA -> Excel + Microsoft SQL Server (python-tds / pytds ONLY).

Examples:
  python oic_collector.py --mode offline --source-file capture.txt --output excel
  python oic_collector.py --mode live --output both

The input capture may contain credentials: never commit it to Git.
"""

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import sys
import time
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation, ROUND_HALF_EVEN, localcontext
from xml.etree import ElementTree as ET

ENDPOINT = 'https://onlinewebibs.oic.or.th/xmla-proxy/api/xmla-proxy'
XMLA_NS = 'urn:schemas-microsoft-com:xml-analysis'
MD_NS = 'urn:schemas-microsoft-com:xml-analysis:mddataset'
SOAP_NS = 'http://schemas.xmlsoap.org/soap/envelope/'
PROFILE_FILE = Path(__file__).with_name('request_profiles.json')

FIELDS = [
    'record_key', 'insurance_code', 'insurance_type', 'catalog', 'cube_name',
    'column_dimension', 'column_member', 'column_unique_name',
    'row_dimension_1', 'row_member_1', 'row_unique_name_1',
    'row_dimension_2', 'row_member_2', 'row_unique_name_2',
    'aggregation_level', 'measure_code', 'metric_value', 'metric_raw',
    'cube_last_update', 'retrieved_utc',
]


def local_name(tag):
    return tag.rsplit('}', 1)[-1]


def profiles_load():
    return json.loads(PROFILE_FILE.read_text(encoding='utf-8'))


def soap_request(profile):
    envelope = ET.Element(f'{{{SOAP_NS}}}Envelope')
    body = ET.SubElement(envelope, f'{{{SOAP_NS}}}Body')
    exe = ET.SubElement(body, f'{{{XMLA_NS}}}Execute')
    cmd = ET.SubElement(exe, f'{{{XMLA_NS}}}Command')
    ET.SubElement(cmd, f'{{{XMLA_NS}}}Statement').text = profile['mdx']
    props = ET.SubElement(exe, f'{{{XMLA_NS}}}Properties')
    prop_list = ET.SubElement(props, f'{{{XMLA_NS}}}PropertyList')
    for key, value in (
        ('Catalog', profile['catalog']),
        ('ShowHiddenCubes', 'True'),
        ('SspropInitAppName', 'Microsoft SQL Server Management Studio'),
        ('Timeout', '3600'), ('LocaleIdentifier', '9'),
    ):
        ET.SubElement(prop_list, f'{{{XMLA_NS}}}{key}').text = value
    return ET.tostring(envelope, encoding='utf-8', xml_declaration=False)


def acquire_live(profile, timeout, session):
    token = os.environ.get('OIC_BEARER_TOKEN', '').strip()
    if token.lower().startswith('bearer '):
        token = token[7:].strip()
    if not token:
        raise RuntimeError('Missing OIC_BEARER_TOKEN. Use a valid authorized browser session token.')
    headers = {
        'Content-Type': 'text/xml; charset=utf-8',
        'Accept': 'text/plain, */*; q=0.01',
        'Authorization': 'Bearer ' + token,
        'Origin': 'https://onlinewebibs.oic.or.th',
        'Referer': ('https://onlinewebibs.oic.or.th/WebForm/Dashboard/Pivot/'
                    f'DynamicPivotCube.aspx?Class={profile["class"]}'),
        'X-Requested-With': 'XMLHttpRequest',
    }
    cookie = os.environ.get('OIC_COOKIE', '').strip()
    if cookie:
        headers['Cookie'] = cookie
    for attempt in range(1, 4):
        try:
            res = session.post(ENDPOINT, data=soap_request(profile), headers=headers,
                               timeout=(15, timeout), allow_redirects=False)
        except Exception as exc:
            if attempt >= 3:
                raise RuntimeError(f'Network request failed after {attempt} attempts: {type(exc).__name__}: {exc}') from exc
            time.sleep(2 ** attempt)
            continue
        if res.status_code == 200:
            if 'xml' not in res.headers.get('Content-Type', '').lower() and not res.content.lstrip().startswith(b'<'):
                raise RuntimeError('HTTP 200 but response is not XML (possible HTML login/session page).')
            return res.content
        if res.status_code in (401, 403):
            raise RuntimeError(f'HTTP {res.status_code}: token/session is invalid or lacks permission; reauthenticate normally.')
        if res.status_code in (429, 500, 502, 503, 504) and attempt < 3:
            wait = res.headers.get('Retry-After', '')
            wait = min(60, int(wait)) if wait.isdecimal() else 2 ** attempt
            time.sleep(wait)
            continue
        raise RuntimeError(f'HTTP {res.status_code}: endpoint rejected the XMLA request.')
    raise RuntimeError('Request failed')


def read_offline(source_file, codes):
    """Extract exactly the captured XML SOAP responses; do NOT reuse sensitive headers."""
    lines = Path(source_file).read_text(encoding='utf-8-sig').splitlines()
    response_lines = [line.strip() for line in lines if line.lstrip().startswith(('<soap:Envelope', '<Envelope'))
                      and ('ExecuteResponse' in line[:300] or '<soap:Fault' in line[:300])]
    original_codes = list(profiles_load())
    if len(response_lines) < len(original_codes):
        raise RuntimeError(f'Expected {len(original_codes)} SOAP responses; found {len(response_lines)} in {source_file}')
    return {code: response_lines[i].encode('utf-8') for i, code in enumerate(original_codes) if code in codes}


def members_from_tuple(tup):
    out = []
    for mem in tup.findall(f'{{{MD_NS}}}Member'):
        uname = mem.findtext(f'{{{MD_NS}}}UName', default='')
        caption = mem.findtext(f'{{{MD_NS}}}Caption', default='')
        out.append({'dimension': mem.attrib.get('Hierarchy', ''),
                    'caption': caption, 'uname': uname,
                    'is_all': uname.endswith('[All]')})
    return out


def metric_decimal(raw):
    if raw is None or raw.strip() == '':
        return None
    try:
        d = Decimal(raw)
        if not d.is_finite():
            return None
        return d
    except InvalidOperation:
        return None


def sql_decimal(value):
    """Numeric helper for SQL DECIMAL(38,12), retaining full original in metric_raw."""
    if value is None:
        return None
    with localcontext() as context:
        context.prec = 70
        candidate = value.quantize(Decimal('0.000000000001'), rounding=ROUND_HALF_EVEN)
    # DECIMAL(38,12) allows at most 26 digits to the left of the decimal.
    # The original string stays in metric_raw even when SQL numeric overflows.
    if candidate.copy_abs() >= Decimal('1E26'):
        return None
    return candidate


def parse_mddataset(xml_bytes, code, profile):
    try:
        envelope = ET.fromstring(xml_bytes)
    except ET.ParseError as exc:
        raise RuntimeError(f'{code}: response XML is invalid: {exc}') from exc
    for node in envelope.iter():
        if local_name(node.tag) == 'Fault':
            fault = ' '.join((child.text or '') for child in node.iter() if child.text)
            raise RuntimeError(f'{code}: XMLA SOAP Fault: {fault[:350]}')

    mdroot = next((x for x in envelope.iter() if x.tag == f'{{{MD_NS}}}root'), None)
    if mdroot is None:
        raise RuntimeError(f'{code}: no XMLA mddataset/root in SOAP response')
    axes_el = mdroot.find(f'{{{MD_NS}}}Axes')
    if axes_el is None:
        raise RuntimeError(f'{code}: XMLA response has no Axes')
    axes = {axis.attrib.get('name'): axis for axis in axes_el.findall(f'{{{MD_NS}}}Axis')}
    if 'Axis0' not in axes or 'Axis1' not in axes:
        raise RuntimeError(f'{code}: expected Axis0 and Axis1')
    axis0_tuples = axes['Axis0'].findall(f'.//{{{MD_NS}}}Tuples/{{{MD_NS}}}Tuple')
    axis1_tuples = axes['Axis1'].findall(f'.//{{{MD_NS}}}Tuples/{{{MD_NS}}}Tuple')
    if not axis0_tuples:
        raise RuntimeError(f'{code}: Axis0 has no tuples')
    axis0 = [members_from_tuple(t) for t in axis0_tuples]
    axis1 = [members_from_tuple(t) for t in axis1_tuples]

    cube_name = ''
    cube_last = ''
    for node in mdroot.iter():
        if local_name(node.tag) == 'CubeName' and node.text:
            cube_name = node.text
        if local_name(node.tag) == 'LastDataUpdate' and node.text:
            cube_last = node.text
    fetched_at = datetime.now(timezone.utc).replace(tzinfo=None)
    retrieved_iso = fetched_at.isoformat(timespec='seconds') + 'Z'
    output = []
    cells = mdroot.find(f'{{{MD_NS}}}CellData')
    if cells is None:
        raise RuntimeError(f'{code}: XMLA response has no CellData')
    ncols = len(axis0)
    for cell in cells.findall(f'{{{MD_NS}}}Cell'):
        raw = cell.findtext(f'{{{MD_NS}}}Value')
        if raw is None or raw.strip() == '':
            continue  # empty XMLA cells are not measurements
        ordinal = int(cell.attrib['CellOrdinal'])
        row_idx, col_idx = divmod(ordinal, ncols)
        if row_idx >= len(axis1):
            raise RuntimeError(f'{code}: invalid CellOrdinal {ordinal}')
        c_members = axis0[col_idx]
        r_members = axis1[row_idx]
        measure = next((m['caption'] for m in c_members if m['dimension'] == '[Measures]'), '')
        col_members = [m for m in c_members if m['dimension'] != '[Measures]']
        col = col_members[0] if col_members else {'dimension':'', 'caption':'', 'uname':''}
        r1 = r_members[0] if r_members else {'dimension':'', 'caption':'', 'uname':''}
        r2 = r_members[1] if len(r_members) >= 2 else {'dimension':'', 'caption':'', 'uname':''}
        if r_members and all(m['is_all'] for m in r_members):
            level = 'TOTAL'
        elif any(m['is_all'] for m in r_members):
            level = 'SUBTOTAL'
        else:
            level = 'DETAIL'
        stable_key = json.dumps([profile['catalog'], cube_name,
                                 [m['uname'] for m in c_members],
                                 [m['uname'] for m in r_members]],
                                ensure_ascii=False, separators=(',', ':'))
        key = hashlib.sha256(stable_key.encode('utf-8')).hexdigest()
        output.append({
            'record_key':key, 'insurance_code':code,
            'insurance_type':profile['name_th'], 'catalog':profile['catalog'],
            'cube_name':cube_name, 'column_dimension':col['dimension'],
            'column_member':col['caption'], 'column_unique_name':col['uname'],
            'row_dimension_1':r1['dimension'], 'row_member_1':r1['caption'],
            'row_unique_name_1':r1['uname'], 'row_dimension_2':r2['dimension'],
            'row_member_2':r2['caption'], 'row_unique_name_2':r2['uname'],
            'aggregation_level':level, 'measure_code':measure,
            'metric_value':metric_decimal(raw), 'metric_raw':raw,
            'cube_last_update':cube_last, 'retrieved_utc':retrieved_iso,
            '_retrieved_datetime':fetched_at,
        })
    if not output:
        raise RuntimeError(f'{code}: XMLA response returned zero non-empty values; not writing an empty replacement')
    return output


def save_excel(all_data, out_path):
    from xlsx_writer import write_excel
    write_excel(all_data, FIELDS, out_path)


def safe_table_name(name):
    parts = name.split('.')
    if len(parts) != 2 or any(not re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]*', p) for p in parts):
        raise ValueError('SQL table must look like dbo.OIC_INSURANCE_METRICS')
    return f'[{parts[0]}].[{parts[1]}]', f'{parts[0]}.{parts[1]}'


def save_sql(all_data, table):
    try:
        import pytds  # python-tds only; pyodbc is NOT used.
    except ImportError as exc:
        raise RuntimeError('Missing python-tds. Install: python -m pip install python-tds') from exc
    fq_table, plain_table = safe_table_name(table)
    server = os.getenv('SQL_SERVER','').strip()
    database = os.getenv('SQL_DATABASE','').strip()
    username = os.getenv('SQL_USER','').strip()
    password = os.getenv('SQL_PASSWORD','')
    if not (server and database and username and password):
        raise RuntimeError('For SQL output set SQL_SERVER, SQL_DATABASE, SQL_USER and SQL_PASSWORD')
    port = int(os.getenv('SQL_PORT','1433'))
    conn = pytds.connect(server=server, database=database, user=username,
                         password=password, port=port, login_timeout=20,
                         timeout=180, autocommit=False, appname='OIC_XMLA_Collector')
    try:
        ddl = f'''IF OBJECT_ID(N'{plain_table}', N'U') IS NULL
BEGIN
    CREATE TABLE {fq_table} (
        [record_key] CHAR(64) NOT NULL PRIMARY KEY,
        [insurance_code] NVARCHAR(8) NOT NULL,
        [insurance_type] NVARCHAR(120) NOT NULL,
        [catalog] NVARCHAR(120) NOT NULL,
        [cube_name] NVARCHAR(120) NULL,
        [column_dimension] NVARCHAR(300) NULL,
        [column_member] NVARCHAR(1000) NULL,
        [column_unique_name] NVARCHAR(1500) NULL,
        [row_dimension_1] NVARCHAR(300) NULL,
        [row_member_1] NVARCHAR(1500) NULL,
        [row_unique_name_1] NVARCHAR(1800) NULL,
        [row_dimension_2] NVARCHAR(300) NULL,
        [row_member_2] NVARCHAR(1500) NULL,
        [row_unique_name_2] NVARCHAR(1800) NULL,
        [aggregation_level] NVARCHAR(20) NOT NULL,
        [measure_code] NVARCHAR(120) NOT NULL,
        [metric_value] DECIMAL(38,12) NULL,
        [metric_raw] NVARCHAR(200) NOT NULL,
        [cube_last_update] NVARCHAR(50) NULL,
        [retrieved_utc] DATETIME2(0) NOT NULL
    );
END'''
        with conn.cursor() as cur:
            cur.execute(ddl)
        conn.commit()
        sql_fields = FIELDS
        cols = ', '.join(f'[{field}]' for field in sql_fields)
        placeholders = ', '.join(['%s'] * len(sql_fields))
        insert_sql = f'INSERT INTO {fq_table} ({cols}) VALUES ({placeholders})'
        for code, rows in all_data.items():
            try:
                with conn.cursor() as cur:
                    # Replace one complete insurance class transactionally, without accumulating duplicates.
                    cur.execute(f'DELETE FROM {fq_table} WHERE [insurance_code] = %s', (code,))
                    for start in range(0,len(rows),250):
                        batch = []
                        for item in rows[start:start+250]:
                            vals = []
                            for key in sql_fields:
                                if key == 'metric_value':
                                    vals.append(sql_decimal(item[key]))
                                elif key == 'retrieved_utc':
                                    vals.append(item['_retrieved_datetime'])
                                else:
                                    vals.append(item[key] or None)
                            batch.append(tuple(vals))
                        cur.executemany(insert_sql, batch)
                conn.commit()
                print(f'  SQL committed {code}: {len(rows):,} rows')
            except Exception:
                conn.rollback()
                raise
    finally:
        conn.close()


def main():
    parser = argparse.ArgumentParser(description='Extract OIC XMLA data to Excel and/or SQL Server (pytds only)')
    parser.add_argument('--mode', choices=['live','offline'], default='offline')
    parser.add_argument('--source-file', help='Captured .txt containing 5 XMLA responses (offline mode)')
    parser.add_argument('--output', choices=['excel','sql','both'], default='excel')
    parser.add_argument('--codes', nargs='+', choices=list(profiles_load()), default=list(profiles_load()))
    parser.add_argument('--excel-path', default='', help='Custom .xlsx output path')
    parser.add_argument('--table', default='dbo.OIC_INSURANCE_METRICS')
    parser.add_argument('--timeout', type=int, default=180)
    parser.add_argument('--delay', type=float, default=2.0, help='Delay between live requests (seconds)')
    args = parser.parse_args()
    profiles = profiles_load()
    codes = list(dict.fromkeys(args.codes))
    data = {}
    if args.mode == 'offline':
        if not args.source_file:
            parser.error('--source-file is required when --mode offline')
        offline_data = read_offline(args.source_file, codes)
        for code in codes:
            data[code] = parse_mddataset(offline_data[code], code, profiles[code])
            print(f'[OK] {code}: {len(data[code]):,} data cells (offline)')
    else:
        try:
            import requests
        except ImportError as exc:
            raise RuntimeError('Missing requests: python -m pip install requests') from exc
        with requests.Session() as session:
            for i, code in enumerate(codes):
                if i and args.delay:
                    time.sleep(args.delay)
                raw = acquire_live(profiles[code], args.timeout, session)
                data[code] = parse_mddataset(raw, code, profiles[code])
                print(f'[OK] {code}: {len(data[code]):,} data cells (live)')
    total = sum(len(rows) for rows in data.values())
    if args.output in ('excel','both'):
        path = Path(args.excel_path) if args.excel_path else Path('output') / ('oic_insurance_' + datetime.now().strftime('%Y%m%d_%H%M%S') + '.xlsx')
        save_excel(data, path)
        print(f'[EXCEL] {path.resolve()}')
    if args.output in ('sql','both'):
        save_sql(data, args.table)
        print(f'[SQL] Done: {args.table}')
    print(f'[DONE] {len(data)} categories, {total:,} non-empty measurements')


if __name__ == '__main__':
    try:
        main()
    except (ValueError, RuntimeError, OSError) as exc:
        print(f'[ERROR] {exc}', file=sys.stderr)
        sys.exit(1)