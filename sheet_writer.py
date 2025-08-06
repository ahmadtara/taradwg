from datetime import datetime
from math import dist

def fill_fdt_pekanbaru(sheet, fdt_points, poles_74, district, subdistrict, vendor, kmz_filename):
    headers = sheet.row_values(1)
    header_map = {name.strip().lower(): i for i, name in enumerate(headers)}

    values = sheet.get_all_values()
    for i in range(len(values) - 1, 0, -1):
        if any(values[i]):
            prev_row = values[i]
            break
    else:
        prev_row = [""] * len(headers)

    today = datetime.today()
    formatted_date = today.strftime("%d/%m/%Y") if "/" in prev_row[header_map.get('ah', 0)] else today.strftime("%Y-%m-%d")

    def parse_fdt_type(name):
        if "FDT 48" in name.upper():
            return ("2", "4", "FDT TYPE 48 CORE")
        elif "FDT 72" in name.upper():
            return ("3", "6", "FDT TYPE 72 CORE")
        elif "FDT 96" in name.upper():
            return ("4", "8", "FDT TYPE 96 CORE")
        return ("", "", "")

    def find_nearest_pole(pt):
        min_d, nearest_name = float('inf'), ""
        for pole in poles_74:
            d = dist([pt['lat'], pt['lon']], [pole['lat'], pole['lon']])
            if d < min_d:
                min_d, nearest_name = d, pole['name']
        return nearest_name

    rows = []
    for pt in fdt_points:
        row = [""] * len(headers)
        row[1:5] = prev_row[1:5]  # B-E
        row[5] = district.upper()  # F
        row[6] = subdistrict.upper()  # G
        row[7] = kmz_filename  # H
        row[8] = pt['name']  # I
        row[9] = pt['name']  # J
        row[10] = pt['lat']  # K
        row[11] = pt['lon']  # L

        row[12:14] = prev_row[12:14]  # M-N
        row[13:15] = prev_row[13:15]  # N-O
        row[24:28] = prev_row[24:28]  # Y-Z-AA-AB
        row[30:32] = prev_row[30:32]  # AD-AE
        row[18] = prev_row[18]  # S

        row[33] = formatted_date  # AH
        row[31] = vendor.upper()  # AF
        row[44] = vendor.upper()  # AS

        row[39] = find_nearest_pole(pt)  # AN

        m, r, ap = parse_fdt_type(pt['name'])
        row[12] = m  # M
        row[17] = r  # R
        row[41] = ap  # AP

        if 'templatecode' in pt:
            row[header_map.get("templatecode")] = pt['templatecode']

        rows.append(row)

    sheet.append_rows(rows)
    return len(rows)

def fill_cable_pekanbaru(sheet, cable_paths, vendor, kmz_filename):
    headers = sheet.row_values(1)
    header_map = {name.strip().lower(): i for i, name in enumerate(headers)}
    values = sheet.get_all_values()
    for i in range(len(values) - 1, 0, -1):
        if any(values[i]):
            prev_row = values[i]
            break
    else:
        prev_row = [""] * len(headers)

    today = datetime.today()
    formatted_date = today.strftime("%d/%m/%Y") if "/" in prev_row[header_map.get('y', 0)] else today.strftime("%Y-%m-%d")

    def parse_fo(text):
        if "FO 24/2T" in text.upper(): return "2"
        if "FO 48/4T" in text.upper(): return "4"
        if "FO 36/3T" in text.upper(): return "3"
        return ""

    def parse_ae(text):
        import re
        match = re.search(r"AE\s*-?\s*(\d+)\s*M", text.upper())
        return match.group(1) if match else ""

    rows = []
    for path in cable_paths:
        row = [""] * len(headers)
        row[0] = path['name']
        row[1] = path['name']
        for i in ['c', 'd', 'e', 'f', 'k', 'u', 'v']:
            idx = header_map.get(i)
            if idx is not None:
                row[idx] = prev_row[idx]

        row[9] = parse_fo(path['name'])  # J
        row[16] = str(round(path['length'], 2))  # P
        row[17] = parse_ae(path['name'])  # Q
        row[24] = formatted_date  # Y
        row[26] = vendor  # AM
        row[22] = kmz_filename  # AK

        rows.append(row)

    sheet.append_rows(rows)
    return len(rows)

def fill_subfeeder_pekanbaru(sheet, sub_paths, vendor, kmz_filename):
    headers = sheet.row_values(1)
    header_map = {name.strip().lower(): i for i, name in enumerate(headers)}
    values = sheet.get_all_values()
    for i in range(len(values) - 1, 0, -1):
        if any(values[i]):
            prev_row = values[i]
            break
    else:
        prev_row = [""] * len(headers)

    today = datetime.today()
    formatted_date = today.strftime("%d/%m/%Y") if "/" in prev_row[header_map.get('y', 0)] else today.strftime("%Y-%m-%d")

    def parse_jm(text):
        if "FO 24/2T" in text.upper(): return ("2", "24")
        if "FO 48/4T" in text.upper(): return ("4", "48")
        if "FO 12/2T" in text.upper(): return ("12", "12")
        if "FO 96/9T" in text.upper(): return ("", "96")
        return ("", "")

    def parse_ae(text):
        import re
        match = re.search(r"AE\s*-?\s*(\d+)\s*M", text.upper())
        return match.group(1) if match else ""

    rows = []
    for path in sub_paths:
        row = [""] * len(headers)
        row[0] = path['name']
        row[1] = path['name']
        for col in ['c', 'd', 'e', 'f', 'k', 'l', 'u', 'v', 'aj']:
            idx = header_map.get(col)
            if idx is not None:
                row[idx] = prev_row[idx]

        j, m = parse_jm(path['name'])
        row[9] = j  # J
        row[12] = m  # M
        row[16] = str(round(path['length'], 2))  # P
        row[17] = parse_ae(path['name'])  # Q
        row[24] = formatted_date  # Y
        row[26] = vendor  # AM
        row[22] = kmz_filename  # AK

        rows.append(row)

    sheet.append_rows(rows)
    return len(rows)
