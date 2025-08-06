# sheet_writer.py
import gspread
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

def authenticate_google(creds_dict):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(credentials)
    return client

def get_latest_row(sheet):
    values = sheet.get_all_values()
    for i in range(len(values) - 1, 0, -1):
        if any(values[i]):
            return values[i]
    return [""] * len(sheet.row_values(1))

def get_date_format(ref_value):
    if ref_value.count("/") == 2:
        return "%d/%m/%Y"
    return "%Y-%m-%d"

def fill_fdt_pekanbaru(sheet, fdt_points, poles_74, district, subdistrict, vendor, kmz_name):
    headers = sheet.row_values(1)
    header_map = {name.strip().lower(): idx for idx, name in enumerate(headers)}
    prev_row = get_latest_row(sheet)
    date_fmt = get_date_format(prev_row[header_map.get('installationdate', 0)])
    today = datetime.today().strftime(date_fmt)

    result_rows = []
    for fdt in fdt_points:
        row = [""] * len(headers)

        # Duplicate columns B-E, N-O, Y, Z, AA, AB, AD, AE, S
        for col in ['b','c','d','e','n','o','y','z','aa','ab','ad','ae','s']:
            idx = header_map.get(col.lower())
            if idx is not None:
                row[idx] = prev_row[idx]

        # Static / mapped columns
        row[header_map['i']] = fdt['name']
        row[header_map['j']] = fdt['name']
        row[header_map['k']] = fdt['lat']
        row[header_map['l']] = fdt['lon']
        row[header_map['f']] = district.upper()
        row[header_map['g']] = subdistrict.upper()
        row[header_map['af']] = vendor.upper()
        row[header_map['as']] = vendor.upper()
        row[header_map['ah']] = today
        row[header_map['h']] = kmz_name

        # Parent ID 1 from nearest pole
        if 'parentid 1' in header_map:
            row[header_map['parentid 1']] = find_nearest_pole(fdt, poles_74)

        # Templatecode
        if 'templatecode' in header_map:
            row[header_map['templatecode']] = fdt['desc']

        # Kolom M, R, AP dari nama FDT
        a_value = fdt['name'].upper()
        if 'fdt 48' in a_value:
            row[header_map.get('m')] = '2'
            row[header_map.get('r')] = '4'
            row[header_map.get('ap')] = 'FDT TYPE 48 CORE'
        elif 'fdt 72' in a_value:
            row[header_map.get('m')] = '3'
            row[header_map.get('r')] = '6'
            row[header_map.get('ap')] = 'FDT TYPE 72 CORE'
        elif 'fdt 96' in a_value:
            row[header_map.get('m')] = '4'
            row[header_map.get('r')] = '8'
            row[header_map.get('ap')] = 'FDT TYPE 96 CORE'

        result_rows.append(row)

    sheet.append_rows(result_rows)
    return len(result_rows)

def fill_cable_pekanbaru(sheet, path_data, vendor, kmz_name):
    headers = sheet.row_values(1)
    header_map = {name.strip().lower(): idx for idx, name in enumerate(headers)}
    prev_row = get_latest_row(sheet)
    date_fmt = get_date_format(prev_row[header_map.get('y', 0)])
    today = datetime.today().strftime(date_fmt)

    result_rows = []
    for path in path_data:
        row = [""] * len(headers)

        row[header_map.get('a')] = path['name']
        row[header_map.get('b')] = path['name']

        for col in ['c','d','e','f','k','u','v']:
            idx = header_map.get(col.lower())
            if idx is not None:
                row[idx] = prev_row[idx]

        row[header_map.get('am')] = vendor.upper()
        row[header_map.get('ak')] = kmz_name
        row[header_map.get('y')] = today

        # FO Capacity (J)
        ports, _ = parse_fo_capacity(path['name'])
        if ports and 'j' in header_map:
            row[header_map['j']] = str(ports)

        # AE Distance (Q)
        dist = parse_distance(path['name'])
        if dist and 'q' in header_map:
            row[header_map['q']] = str(dist)

        # Path length (P)
        if 'p' in header_map:
            row[header_map['p']] = str(measure_path_length(path['coords']))

        result_rows.append(row)

    sheet.append_rows(result_rows)
    return len(result_rows)

def fill_subfeeder_pekanbaru(sheet, path_data, vendor, kmz_name):
    headers = sheet.row_values(1)
    header_map = {name.strip().lower(): idx for idx, name in enumerate(headers)}
    prev_row = get_latest_row(sheet)
    date_fmt = get_date_format(prev_row[header_map.get('y', 0)])
    today = datetime.today().strftime(date_fmt)

    result_rows = []
    for path in path_data:
        row = [""] * len(headers)

        row[header_map.get('a')] = path['name']
        row[header_map.get('b')] = path['name']

        for col in ['c','d','e','f','k','l','u','v','aj']:
            idx = header_map.get(col.lower())
            if idx is not None:
                row[idx] = prev_row[idx]

        row[header_map.get('am')] = vendor.upper()
        row[header_map.get('ak')] = kmz_name
        row[header_map.get('y')] = today

        # FO Capacity
        ports, cores = parse_fo_capacity(path['name'])
        if ports and 'j' in header_map:
            row[header_map['j']] = str(ports)
        if cores and 'm' in header_map:
            row[header_map['m']] = str(cores)

        dist = parse_distance(path['name'])
        if dist and 'q' in header_map:
            row[header_map['q']] = str(dist)

        if 'p' in header_map:
            row[header_map['p']] = str(measure_path_length(path['coords']))

        result_rows.append(row)

    sheet.append_rows(result_rows)
    return len(result_rows)
