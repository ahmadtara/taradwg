from datetime import datetime
import streamlit as st

def append_poles_to_main_sheet(sheet, poles, district, subdistrict, vendor):
    headers = sheet.row_values(1)
    header_map = {h.strip().lower(): i for i, h in enumerate(headers)}

    prev = sheet.get_all_values()[-1] if sheet.get_all_values() else [""] * len(headers)
    today = datetime.today()
    date_str = today.strftime("%d/%m/%Y") if prev[header_map.get('installationdate', 0)].count("/") == 2 else today.strftime("%Y-%m-%d")

    rows = []
    for p in poles:
        row = [""] * len(headers)
        row[0:4] = prev[0:4]
        row[4] = district.upper()
        row[5] = subdistrict.upper()
        row[6] = p['name']
        row[7] = p['name']
        row[8] = p['lat']
        row[9] = p['lon']

        for col in ['constructionstage', 'accessibility', 'activationstage', 'hierarchytype']:
            if col in header_map:
                row[header_map[col]] = prev[header_map[col]]

        for col in ['pole height', 'vendorname', 'installationyear', 'productionyear', 'installationdate', 'remarks']:
            if col in header_map:
                if col == 'pole height':
                    row[header_map[col]] = p['height']
                elif col == 'vendorname':
                    row[header_map[col]] = vendor.upper()
                elif col in ['installationyear', 'productionyear']:
                    row[header_map[col]] = str(today.year)
                elif col == 'installationdate':
                    row[header_map[col]] = date_str
                elif col == 'remarks':
                    row[header_map[col]] = p.get('remarks', "")

        if 'poletype' in header_map:
            row[header_map['poletype']] = p['folder']

        rows.append(row)

    sheet.append_rows(rows)
    st.success(f"✅ {len(rows)} pole dikirim ke spreadsheet utama.")

def append_fat_to_sheet(sheet, fat_points, poles, district, subdistrict, vendor):
    from math import dist

    def find_nearest(p_fat):
        min_d = float('inf')
        nearest = ""
        for pole in poles:
            d = dist([p_fat['lat'], p_fat['lon']], [pole['lat'], pole['lon']])
            if d < min_d:
                min_d = d
                nearest = pole['name']
        return nearest

    headers = sheet.row_values(1)
    header_map = {h.strip().lower(): i for i, h in enumerate(headers)}

    prev = sheet.get_all_values()[-1] if sheet.get_all_values() else [""] * len(headers)
    today = datetime.today()
    date_str = today.strftime("%d/%m/%Y") if prev[header_map.get('installation_date', 0)].count("/") == 2 else today.strftime("%Y-%m-%d")

    rows = []
    for fat in fat_points:
        row = [""] * len(headers)
        row[0:5] = prev[0:5]
        row[5] = district.upper()
        row[6] = subdistrict.upper()
        row[7] = fat['name']
        row[8] = fat['name']
        row[9] = fat['lat']
        row[10] = fat['lon']

        for i in range(11, 24):
            row[i] = prev[i] if i < len(prev) else ""

        row[24] = vendor.upper()
        row[26] = date_str

        if 'parentid 1' in header_map:
            row[header_map['parentid 1']] = find_nearest(fat)

        for key in ['parent_type 1', 'fat type']:
            if key in header_map:
                row[header_map[key]] = prev[header_map[key]]

        if 'vendor name' in header_map:
            row[header_map['vendor name']] = vendor.upper()

        rows.append(row)

    sheet.append_rows(rows)
    st.success(f"✅ {len(rows)} FAT dikirim ke spreadsheet ke-2.")

def append_to_fdt_sheet(sheet, fdt_points):
    headers = sheet.row_values(1)
    header_map = {h.strip().lower(): i for i, h in enumerate(headers)}
    rows = []

    for p in fdt_points:
        row = [""] * len(headers)
        row[0] = p['name']
        row[1] = p['lat']
        row[2] = p['lon']
        rows.append(row)

    sheet.append_rows(rows)
    st.success(f"✅ {len(rows)} FDT dikirim ke spreadsheet 3.")

def append_to_cable_cluster_sheet(sheet, cable_cluster):
    headers = sheet.row_values(1)
    header_map = {h.strip().lower(): i for i, h in enumerate(headers)}
    rows = []

    for p in cable_cluster:
        row = [""] * len(headers)
        row[0] = p['name']
        row[1] = p['lat']
        row[2] = p['lon']
        rows.append(row)

    sheet.append_rows(rows)
    st.success(f"✅ {len(rows)} Cable (CLUSTER) dikirim ke spreadsheet 4.")

def append_to_subfeeder_cable_sheet(sheet, cable_subfeeder):
    headers = sheet.row_values(1)
    header_map = {h.strip().lower(): i for i, h in enumerate(headers)}
    rows = []

    for p in cable_subfeeder:
        row = [""] * len(headers)
        row[0] = p['name']
        row[1] = p['lat']
        row[2] = p['lon']
        rows.append(row)

    sheet.append_rows(rows)
    st.success(f"✅ {len(rows)} Cable (SUBFEEDER) dikirim ke spreadsheet 5.")
