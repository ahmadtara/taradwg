import re
from datetime import datetime
from math import dist


def extract_value(text, pattern, default=None, as_int=False):
    match = re.search(pattern, text)
    if match:
        val = match.group(1)
        return int(val) if as_int else val
    return default


def find_nearest_pole(source, poles):
    return min(poles, key=lambda p: dist([source['lat'], source['lon']], [p['lat'], p['lon']]), default={}).get('name', '')


def append_to_fdt_sheet(sheet, fdt_points, poles, kmz_filename, district, subdistrict, vendor):
    headers = sheet.row_values(1)
    header_map = {h.strip().lower(): i for i, h in enumerate(headers)}
    prev_row = next((r for r in reversed(sheet.get_all_values()) if any(r)), [""] * len(headers))
    today = datetime.today()
    formatted_date = today.strftime("%d/%m/%Y") if '/' in prev_row[header_map['ah']] else today.strftime("%Y-%m-%d")

    rows = []
    for fdt in fdt_points:
        name = fdt['name']
        row = [""] * len(headers)
        row[1:5] = prev_row[1:5]
        row[5] = district.upper()
        row[6] = subdistrict.upper()
        row[7] = kmz_filename
        row[8] = name
        row[9] = name
        row[10] = fdt['lat']
        row[11] = fdt['lon']
        row[13:15] = prev_row[13:15]
        row[24:27] = prev_row[24:27]
        row[28] = vendor.upper()
        row[33] = vendor.upper()
        row[34] = formatted_date
        row[18] = find_nearest_pole(fdt, [p for p in poles if p['folder'] == '7m4inch'])

        if 'templatecode' in header_map:
            row[header_map['templatecode']] = fdt['description']

        # Rules for M (12), R (17), AP (40)
        core = extract_value(name, r'FDT\s*(\d+)', as_int=True)
        if core:
            row[12] = {48: '2', 72: '3', 96: '4'}.get(core, '')
            row[17] = {48: '4', 72: '6', 96: '8'}.get(core, '')
            row[40] = {48: 'FDT TYPE 48 CORE', 72: 'FDT TYPE 72 CORE', 96: 'FDT TYPE 96 CORE'}.get(core, '')

        rows.append(row)

    sheet.append_rows(rows)


def append_to_cable_cluster_sheet(sheet, cables, vendor, kmz_filename):
    headers = sheet.row_values(1)
    header_map = {h.strip().lower(): i for i, h in enumerate(headers)}
    prev_row = next((r for r in reversed(sheet.get_all_values()) if any(r)), [""] * len(headers))
    today = datetime.today()
    formatted_date = today.strftime("%d/%m/%Y") if '/' in prev_row[header_map['y']] else today.strftime("%Y-%m-%d")

    rows = []
    for c in cables:
        row = [""] * len(headers)
        row[0] = row[1] = c['name']
        row[2:6] = prev_row[2:6]
        row[10] = prev_row[10]
        row[20:22] = prev_row[20:22]
        row[38] = vendor.upper()
        row[36] = kmz_filename
        row[24] = formatted_date

        row[9] = str(extract_value(c['name'], r'FO\s*(\d+)/(\d+)T', default=''))
        row[16] = str(extract_value(c['name'], r'AE[-\s]*(\d+)', default=''))
        if 'path_length' in c:
            row[15] = f"{c['path_length']:.2f}"

        rows.append(row)

    sheet.append_rows(rows)


def append_to_subfeeder_cable_sheet(sheet, cables, vendor, kmz_filename):
    headers = sheet.row_values(1)
    header_map = {h.strip().lower(): i for i, h in enumerate(headers)}
    prev_row = next((r for r in reversed(sheet.get_all_values()) if any(r)), [""] * len(headers))
    today = datetime.today()
    formatted_date = today.strftime("%d/%m/%Y") if '/' in prev_row[header_map['y']] else today.strftime("%Y-%m-%d")

    rows = []
    for c in cables:
        row = [""] * len(headers)
        row[0] = row[1] = c['name']
        row[2:6] = prev_row[2:6]
        row[10:12] = prev_row[10:12]
        row[20:22] = prev_row[20:22]
        row[35] = prev_row[35]
        row[38] = vendor.upper()
        row[36] = kmz_filename
        row[24] = formatted_date

        row[9] = str(extract_value(c['name'], r'FO\s*(\d+)/(\d+)T', default=''))
        row[12] = str(extract_value(c['name'], r'FO\s*(\d+)', default=''))
        row[16] = str(extract_value(c['name'], r'AE[-\s]*(\d+)', default=''))
        if 'path_length' in c:
            row[15] = f"{c['path_length']:.2f}"

        rows.append(row)

    sheet.append_rows(rows)
