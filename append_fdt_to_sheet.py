from datetime import datetime
from math import dist
import streamlit as st

def find_nearest_pole(fdt_point, poles):
    min_dist = float('inf')
    nearest_name = ""
    for pole in poles:
        d = dist([fdt_point['lat'], fdt_point['lon']], [pole['lat'], pole['lon']])
        if d < min_dist:
            min_dist = d
            nearest_name = pole['name']
    return nearest_name

def templatecode_to_kolom_m(templatecode):
    mapping = {
        "FDT 48": "2",
        "FDT 72": "3",
        "FDT 96": "4"
    }
    return mapping.get(templatecode.strip().upper(), "")

def templatecode_to_kolom_r(templatecode):
    mapping = {
        "FDT 48": "4",
        "FDT 72": "6",
        "FDT 96": "8"
    }
    return mapping.get(templatecode.strip().upper(), "")

def templatecode_to_kolom_ap(templatecode):
    mapping = {
        "FDT 48": "FDT TYPE 48 CORE",
        "FDT 72": "FDT TYPE 72 CORE",
        "FDT 96": "FDT TYPE 96 CORE"
    }
    return mapping.get(templatecode.strip().upper(), "")

def append_fdt_to_sheet(sheet, fdt_data, poles, district, subdistrict, vendor, kmz_name):
    existing_rows = sheet.get_all_values()
    headers = sheet.row_values(1)
    header_map = {name.strip().lower(): i for i, name in enumerate(headers)}
    template_row = existing_rows[-2] if len(existing_rows) > 2 else []
    rows = []

    idx_parentid = header_map.get('parentid 1')
    if idx_parentid is None:
        st.error("Kolom 'Parentid 1' tidak ditemukan di header spreadsheet.")
        return 0

    for fdt in fdt_data:
        name = fdt['name']
        lat = fdt['lat']
        lon = fdt['lon']
        desc = fdt.get('description', '')

        kolom_m = templatecode_to_kolom_m(template_row[0])
        kolom_r = templatecode_to_kolom_r(template_row[0])
        kolom_ap = templatecode_to_kolom_ap(template_row[0])

        row = [""] * len(existing_rows[0])
        row[0] = desc
        row[1:5] = template_row[1:5]
        row[5] = district
        row[6] = subdistrict
        row[7] = kmz_name
        row[8] = name
        row[9] = name
        row[10] = lat
        row[11] = lon
        row[12] = kolom_m
        row[13:16] = template_row[13:16]
        row[17] = kolom_r
        row[18] = template_row[18]
        row[24:28] = template_row[24:28]
        row[29] = template_row[29]
        row[30] = template_row[30]
        row[40] = template_row[40]
        row[41] = kolom_ap
        row[33] = datetime.today().strftime("%d/%m/%Y")
        row[31] = vendor
        row[44] = vendor

        if idx_parentid is not None:
            row[idx_parentid] = find_nearest_pole(fdt, [
                p for p in poles if p['folder'] in ['7m4inch', '7m3inch', 'ext7m3inch', 'ext7m4inch', 'ext9m4inch']
            ])

        rows.append(row)

    sheet.append_rows(rows, value_input_option="USER_ENTERED")
    return len(rows)
