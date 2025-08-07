from datetime import datetime
import streamlit as st

def find_nearest_pole(fat, poles):
    if not poles:
        return ""
    fat_lat, fat_lon = float(fat['lat']), float(fat['lon'])

    def distance(pole):
        pole_lat, pole_lon = float(pole['lat']), float(pole['lon'])
        return (fat_lat - pole_lat) ** 2 + (fat_lon - pole_lon) ** 2

    nearest = min(poles, key=distance)
    return nearest['name']

def append_fat_to_sheet(sheet, fat_points, poles, district, subdistrict, vendor):
    headers = sheet.row_values(1)
    header_map = {name.strip().lower(): i for i, name in enumerate(headers)}

    values = sheet.get_all_values()
    for i in range(len(values)-1, 0, -1):
        if any(values[i]):
            prev_row = values[i]
            break
    else:
        prev_row = [""] * len(headers)

    today = datetime.today()
    formatted_date = today.strftime("%d/%m/%Y") if prev_row[header_map.get('installation_date', 0)].count("/") == 2 else today.strftime("%Y-%m-%d")

    all_rows = []
    for fat in fat_points:
        row = [""] * len(headers)

        # Kolom A - E
        for col_idx in range(5):
            row[col_idx] = prev_row[col_idx] if col_idx < len(prev_row) else ""

        # Kolom F - J
        row[5] = district.upper()
        row[6] = subdistrict.upper()
        row[7] = fat['name']
        row[8] = fat['name']
        row[9] = fat['lat']
        row[10] = fat['lon']

        # Kolom K - X
        for idx in range(11, 24):
            row[idx] = prev_row[idx] if idx < len(prev_row) else ""

        # Kolom Y (vendor)
        row[24] = vendor.upper()

        # Kolom AA (installation_date)
        row[26] = formatted_date

        # Parent ID 1 (AG)
        idx_ag = header_map.get('parentid 1')
        if idx_ag is not None:
            row[idx_ag] = find_nearest_pole(fat, [p for p in poles if p['folder'] == '7m3inch'])

        # Parent Type 1 & FAT Type
        for col in ['parent_type 1', 'fat type']:
            idx = header_map.get(col.lower())
            if idx is not None:
                row[idx] = prev_row[idx]

        # Vendor Name
        idx_al = header_map.get('vendor name')
        if idx_al is not None:
            row[idx_al] = vendor.upper()

        all_rows.append(row)

    sheet.append_rows(all_rows)
    st.success(f"✅ {len(fat_points)} FAT")
