import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import tempfile
from datetime import datetime

dist = __import__('math').dist

SPREADSHEET_ID = "1yXBIuX2LjUWxbpnNqf6A9YimtG7d77V_AHLidhWKIS8"
SPREADSHEET_ID_2 = "1WI0Gb8ul5GPUND4ADvhFgH4GSlgwq1_4rRgfOnPz-yc"
SPREADSHEET_ID_3 = "1EnteHGDnRhwthlCO9B12zvHUuv3wtq5L2AKlV11qAOU"
SHEET_NAME = "Pole Pekanbaru"
SHEET_NAME_2 = "Fat Pekanbaru"
SHEET_NAME_3 = "FDT Pekanbaru"

_cached_headers = None
_cached_prev_row = None

def authenticate_google():
    creds_dict = st.secrets["gcp_service_account"]
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(credentials)
    return client

def extract_points_from_kmz(kmz_path):
    fat_points, poles, poles_subfeeder = [], [], []

    def recurse_folder(folder, ns, path=""):
        items = []
        name_el = folder.find("kml:name", ns)
        folder_name = name_el.text.upper() if name_el is not None else "UNKNOWN"
        new_path = f"{path}/{folder_name}" if path else folder_name
        for sub in folder.findall("kml:Folder", ns):
            items += recurse_folder(sub, ns, new_path)
        for pm in folder.findall("kml:Placemark", ns):
            nm = pm.find("kml:name", ns)
            coord = pm.find(".//kml:coordinates", ns)
            if nm is not None and coord is not None and ',' in coord.text:
                lon, lat = coord.text.strip().split(",")[:2]
                items.append({"name": nm.text.strip(), "lat": float(lat), "lon": float(lon), "path": new_path})
        return items

    with zipfile.ZipFile(kmz_path, 'r') as zf:
        kml_file = next((f for f in zf.namelist() if f.lower().endswith(".kml")), None)
        if not kml_file:
            st.error("❌ Tidak ditemukan file .kml dalam .kmz")
            return [], [], []

        root = ET.parse(zf.open(kml_file)).getroot()
        ns = {"kml": "http://www.opengis.net/kml/2.2"}
        all_pm = []
        for folder in root.findall(".//kml:Folder", ns):
            all_pm += recurse_folder(folder, ns)

    for p in all_pm:
        base_folder = p["path"].split("/")[0].upper()
        if base_folder == "FAT":
            fat_points.append(p)
        elif base_folder == "NEW POLE 7-3":
            poles.append({**p, "folder": "7m3inch", "height": "7", "remarks": "CLUSTER"})
            poles_subfeeder.append({**p, "folder": "7m3inch", "height": "7"})
        elif base_folder == "NEW POLE 7-4":
            poles.append({**p, "folder": "7m4inch", "height": "7"})
        elif base_folder == "NEW POLE 9-4":
            poles.append({**p, "folder": "9m4inch", "height": "9"})

    return fat_points, poles, poles_subfeeder

def find_nearest_pole(fat_point, poles):
    min_dist = float('inf')
    nearest_name = ""
    for pole in poles:
        d = dist([fat_point['lat'], fat_point['lon']], [pole['lat'], pole['lon']])
        if d < min_dist:
            min_dist = d
            nearest_name = pole['name']
    return nearest_name

def append_poles_to_main_sheet(sheet, poles, district, subdistrict, vendor):
    global _cached_headers, _cached_prev_row

    headers = _cached_headers or sheet.row_values(1)
    _cached_headers = headers
    header_map = {name.strip().lower(): i for i, name in enumerate(headers)}

    values = sheet.get_all_values()
    for i in range(len(values)-1, 0, -1):
        if any(values[i]):
            prev_row = values[i]
            break
    else:
        prev_row = [""] * len(headers)

    _cached_prev_row = prev_row

    today = datetime.today()
    formatted_date = today.strftime("%d/%m/%Y") if prev_row[header_map.get('installationdate', 0)].count("/") == 2 else today.strftime("%Y-%m-%d")

    count_types = {"7m3inch": 0, "7m4inch": 0, "9m4inch": 0}

    district = district.upper()
    subdistrict = subdistrict.upper()
    vendor = vendor.upper()

    all_rows = []
    for pole in poles:
        count_types[pole['folder']] += 1

        row = [""] * len(headers)
        row[0:4] = prev_row[0:4]
        row[4] = district
        row[5] = subdistrict
        row[6] = pole['name']
        row[7] = pole['name']
        row[8] = pole['lat']
        row[9] = pole['lon']

        for col in ['constructionstage', 'accessibility', 'activationstage', 'hierarchytype']:
            if col in header_map:
                row[header_map[col]] = prev_row[header_map[col]]

        for col in ['pole height', 'vendorname', 'installationyear', 'productionyear', 'installationdate', 'remarks']:
            idx = header_map.get(col.lower())
            if idx is not None:
                if col.lower() == 'pole height':
                    row[idx] = pole['height']
                elif col.lower() == 'vendorname':
                    row[idx] = vendor
                elif col.lower() in ['installationyear', 'productionyear']:
                    row[idx] = str(today.year)
                elif col.lower() == 'installationdate':
                    row[idx] = formatted_date
                elif col.lower() == 'remarks':
                    if pole['folder'] in ['7m4inch', '9m4inch']:
                        row[idx] = "SUBFEEDER"
                    else:
                        row[idx] = "CLUSTER"

        if 'poletype' in header_map:
            row[header_map['poletype']] = pole['folder']

        all_rows.append(row)

    sheet.append_rows(all_rows)

    st.info(f"""
📊 **Ringkasan Pengunggahan**:
- 7m3inch: {count_types['7m3inch']} titik
- 7m4inch: {count_types['7m4inch']} titik
- 9m4inch: {count_types['9m4inch']} titik
""")

def extract_fdt_type_info(name):
    name_upper = name.upper()
    if "FDT 48" in name_upper:
        return ("2", "4", "FDT TYPE 48 CORE")
    elif "FDT 72" in name_upper:
        return ("3", "6", "FDT TYPE 72 CORE")
    elif "FDT 96" in name_upper:
        return ("4", "8", "FDT TYPE 96 CORE")
    else:
        return ("", "", "")

def append_fdt_to_sheet(sheet, fat_points, poles, district, subdistrict, vendor):
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
    formatted_date = today.strftime("%d/%m/%Y") if prev_row[header_map.get('installationdate', 0)].count("/") == 2 else today.strftime("%Y-%m-%d")

    all_rows = []
    for fat in fat_points:
        if 'FDT' not in fat['name'].upper():
            continue

        row = [""] * len(headers)
        row[0] = fat['name']
        for col in [1, 2, 3, 4, 13, 14, 24, 25, 26, 27, 29, 30, 18]:
            if col < len(prev_row):
                row[col] = prev_row[col]
        row[5] = district.upper()
        row[6] = subdistrict.upper()
        if 'vendor name' in header_map:
            row[header_map['vendor name']] = vendor.upper()
        if 'vendorname' in header_map:
            row[header_map['vendorname']] = vendor.upper()
        row[7] = fat['path'].split("/")[0]
        path_parts = fat['path'].split("/")
        row[8] = path_parts[-1] if len(path_parts) > 1 else ""
        row[9] = path_parts[-1] if len(path_parts) > 1 else ""
        row[10] = fat['lat']
        row[11] = fat['lon']
        m_val, r_val, ap_val = extract_fdt_type_info(fat['name'])
        row[12] = m_val
        row[17] = r_val
        row[41] = ap_val
        if 'installationdate' in header_map:
            row[header_map['installationdate']] = formatted_date
        if 'parentid 1' in header_map:
            nearby_pole = find_nearest_pole(fat, [p for p in poles if p['folder'] == '7m4inch'])
            row[header_map['parentid 1']] = nearby_pole
        all_rows.append(row)

    sheet.append_rows(all_rows)
    st.success(f"✅ {len(all_rows)} FDT berhasil dikirim ke Spreadsheet ke-3 🛰️")
