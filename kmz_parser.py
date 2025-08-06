from datetime import datetime
from math import dist
import re

def get_latest_row(sheet):
    values = sheet.get_all_values()
    for i in range(len(values)-1, 0, -1):
        if any(values[i]):
            return values[i]
    return [""] * len(sheet.row_values(1))

def format_date(prev_row, headers):
    today = datetime.today()
    date_col = 'installationdate' if 'installationdate' in headers else 'installation_date'
    idx = headers.get(date_col)
    return today.strftime("%d/%m/%Y") if idx and prev_row[idx].count("/") == 2 else today.strftime("%Y-%m-%d")

def get_header_map(sheet):
    headers = sheet.row_values(1)
    return {name.strip().lower(): i for i, name in enumerate(headers)}

def find_nearest_pole(fdt, poles):
    return min(poles, key=lambda p: dist([fdt['lat'], fdt['lon']], [p['lat'], p['lon']]))['name']

def fill_fdt_pekanbaru(sheet, fdt_points, poles_74, district, subdistrict, vendor, filename):
    header_map = get_header_map(sheet)
    prev_row = get_latest_row(sheet)
    today_str = format_date(prev_row, header_map)
    rows = []

    for fdt in fdt_points:
        name = fdt['name'].upper()
        row = [""] * len(header_map)

        for col in ['b', 'c', 'd', 'e', 'n', 'o', 'y', 'z', 'aa', 'ab', 'ad', 'ae', 's']:
            idx = header_map.get(col.lower())
            if idx: row[idx] = prev_row[idx]

        row[header_map['f']] = district.upper()
        row[header_map['g']] = subdistrict.upper()
        row[header_map['h']] = filename
        row[header_map['i']] = name
        row[header_map['j']] = name
        row[header_map['k']] = fdt['lat']
        row[header_map['l']] = fdt['lon']
        row[header_map['ah']] = today_str
        row[header_map['af']] = vendor.upper()
        row[header_map['as']] = vendor.upper()

        if desc := fdt.get('desc'):
            idx = header_map.get('templatecode')
            if idx: row[idx] = desc

        parent_id = find_nearest_pole(fdt, poles_74)
        if 'an' in header_map:
            row[header_map['an']] = parent_id

        fdt_core = re.search(r'\bFDT\s*(\d+)\b', name)
        if fdt_core:
            core = int(fdt_core.group(1))
            m_val = {48: 2, 72: 3, 96: 4}.get(core, '')
            r_val = {48: 4, 72: 6, 96: 8}.get(core, '')
            ap_val = {48: 'FDT TYPE 48 CORE', 72: 'FDT TYPE 72 CORE', 96: 'FDT TYPE 96 CORE'}.get(core, '')
            if 'm' in header_map: row[header_map['m']] = m_val
            if 'r' in header_map: row[header_map['r']] = r_val
            if 'ap' in header_map: row[header_map['ap']] = ap_val

        rows.append(row)

    sheet.append_rows(rows)
    return len(rows)

def extract_core_count(text):
    match = re.search(r'FO\s*(\d+)', text.upper())
    return int(match.group(1)) if match else None

def extract_length(text):
    match = re.search(r'AE[-\s]?(\d+)', text.upper())
    return int(match.group(1)) if match else 0

def fill_cable_pekanbaru(sheet, cables, vendor, filename):
    header_map = get_header_map(sheet)
    prev_row = get_latest_row(sheet)
    today_str = format_date(prev_row, header_map)
    rows = []

    for path in cables:
        name = path['name'].upper()
        row = [""] * len(header_map)
        row[0] = name
        row[1] = name
        for col in ['c', 'd', 'e', 'f', 'k', 'u', 'v']:
            idx = header_map.get(col.lower())
            if idx: row[idx] = prev_row[idx]
        if 'am' in header_map: row[header_map['am']] = vendor.upper()
        if 'ak' in header_map: row[header_map['ak']] = filename
        if 'y' in header_map: row[header_map['y']] = today_str

        core = extract_core_count(name)
        if core:
            if 'j' in header_map:
                row[header_map['j']] = core // 12 if core <= 48 else 4
        if 'q' in header_map:
            row[header_map['q']] = extract_length(name)
        if 'p' in header_map:
            row[header_map['p']] = int(path['length'])

        rows.append(row)

    sheet.append_rows(rows)
    return len(rows)

def fill_subfeeder_pekanbaru(sheet, cables, vendor, filename):
    header_map = get_header_map(sheet)
    prev_row = get_latest_row(sheet)
    today_str = format_date(prev_row, header_map)
    rows = []

    for path in cables:
        name = path['name'].upper()
        row = [""] * len(header_map)
        row[0] = name
        row[1] = name
        for col in ['c', 'd', 'e', 'f', 'k', 'l', 'u', 'v', 'aj']:
            idx = header_map.get(col.lower())
            if idx: row[idx] = prev_row[idx]
        if 'am' in header_map: row[header_map['am']] = vendor.upper()
        if 'ak' in header_map: row[header_map['ak']] = filename
        if 'y' in header_map: row[header_map['y']] = today_str

        core = extract_core_count(name)
        if core:
            if 'j' in header_map:
                row[header_map['j']] = core // 12 if core <= 48 else 12
            if 'm' in header_map:
                row[header_map['m']] = core
        if 'q' in header_map:
            row[header_map['q']] = extract_length(name)
        if 'p' in header_map:
            row[header_map['p']] = int(path['length'])

        rows.append(row)

    sheet.append_rows(rows)
    return len(rows)
