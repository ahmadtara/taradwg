from datetime import datetime
import re
from math import dist

def duplicate_columns(row, prev_row, indices):
    for idx in indices:
        if idx < len(row) and idx < len(prev_row):
            row[idx] = prev_row[idx]
    return row

def format_date(prev_date):
    today = datetime.today()
    return today.strftime("%d/%m/%Y") if "/" in prev_date else today.strftime("%Y-%m-%d")

def parse_fdt_capacity(name):
    if "FDT 48" in name:
        return 48, 2, 4, "FDT TYPE 48 CORE"
    elif "FDT 72" in name:
        return 72, 3, 6, "FDT TYPE 72 CORE"
    elif "FDT 96" in name:
        return 96, 4, 8, "FDT TYPE 96 CORE"
    return 0, "", "", ""

def parse_cable_j(name):
    if "FO 24/2T" in name:
        return 2
    elif "FO 36/3T" in name:
        return 3
    elif "FO 48/4T" in name:
        return 4
    elif "FO 12/2T" in name:
        return 12
    elif "FO 96/9T" in name:
        return 96
    return ""

def parse_distance_q(text):
    match = re.search(r'AE[-\s]*([\d]+)\s*M', text)
    return match.group(1) if match else ""

def append_to_fdt_sheet(sheet, fdt_points, poles_74, kmz_name, district, subdistrict, vendor):
    headers = sheet.row_values(1)
    values = sheet.get_all_values()
    prev_row = next((row for row in reversed(values) if any(row)), [""] * len(headers))
    header_map = {h.lower(): i for i, h in enumerate(headers)}

    all_rows = []
    for fdt in fdt_points:
        row = [""] * len(headers)

        row[header_map.get("templatecode", 0)] = fdt["description"]
        row[header_map.get("district", 5)] = district.upper()
        row[header_map.get("subdistrict", 6)] = subdistrict.upper()
        row[header_map.get("fdtname", 8)] = fdt["name"]
        row[header_map.get("fdtid", 9)] = fdt["name"]
        row[header_map.get("latitude", 10)] = fdt["lat"]
        row[header_map.get("longitude", 11)] = fdt["lon"]
        row[header_map.get("kmz file", 7)] = kmz_name

        row = duplicate_columns(row, prev_row, [1, 2, 3, 4, 13, 14, 24, 25, 26, 27, 29, 30, 18])
        row[header_map.get("installation date", 33)] = format_date(prev_row[header_map.get("installation date", 33)])
        row[header_map.get("vendor name", 31)] = vendor.upper()
        row[header_map.get("vendor", 44)] = vendor.upper()

        capacity, m_val, r_val, fdt_type = parse_fdt_capacity(fdt["name"])
        row[header_map.get("m", 12)] = m_val
        row[header_map.get("r", 17)] = r_val
        row[header_map.get("fdt type", 41)] = fdt_type

        # Find nearest 7-4 pole
        nearest = min(poles_74, key=lambda p: dist([fdt["lat"], fdt["lon"]], [p["lat"], p["lon"]]))
        row[header_map.get("parentid 1", 39)] = nearest["name"]

        all_rows.append(row)

    sheet.append_rows(all_rows)

def append_to_cable_cluster_sheet(sheet, cables, vendor, kmz_name):
    headers = sheet.row_values(1)
    values = sheet.get_all_values()
    prev_row = next((row for row in reversed(values) if any(row)), [""] * len(headers))
    header_map = {h.lower(): i for i, h in enumerate(headers)}

    all_rows = []
    for cable in cables:
        row = [""] * len(headers)

        row[header_map.get("name", 0)] = cable["name"]
        row[header_map.get("id", 1)] = cable["name"]

        row = duplicate_columns(row, prev_row, [2, 3, 4, 5, 10, 20, 21])
        row[header_map.get("kmz file", 36)] = kmz_name
        row[header_map.get("vendor name", 38)] = vendor.upper()
        row[header_map.get("installation date", 24)] = format_date(prev_row[header_map.get("installation date", 24)])

        row[header_map.get("j", 9)] = parse_cable_j(cable["name"])
        row[header_map.get("q", 16)] = parse_distance_q(cable["name"])
        row[header_map.get("p", 15)] = str(cable.get("length", ""))

        all_rows.append(row)

    sheet.append_rows(all_rows)

def append_to_subfeeder_cable_sheet(sheet, cables, vendor, kmz_name):
    headers = sheet.row_values(1)
    values = sheet.get_all_values()
    prev_row = next((row for row in reversed(values) if any(row)), [""] * len(headers))
    header_map = {h.lower(): i for i, h in enumerate(headers)}

    all_rows = []
    for cable in cables:
        row = [""] * len(headers)

        row[header_map.get("name", 0)] = cable["name"]
        row[header_map.get("id", 1)] = cable["name"]

        row = duplicate_columns(row, prev_row, [2, 3, 4, 5, 10, 11, 20, 21, 35])
        row[header_map.get("vendor name", 38)] = vendor.upper()
        row[header_map.get("kmz file", 36)] = kmz_name
        row[header_map.get("installation date", 24)] = format_date(prev_row[header_map.get("installation date", 24)])

        row[header_map.get("j", 9)] = parse_cable_j(cable["name"])
        row[header_map.get("m", 12)] = parse_cable_j(cable["name"]) if "96" not in cable["name"] else 96
        row[header_map.get("q", 16)] = parse_distance_q(cable["name"])
        row[header_map.get("p", 15)] = str(cable.get("length", ""))

        all_rows.append(row)

    sheet.append_rows(all_rows)
