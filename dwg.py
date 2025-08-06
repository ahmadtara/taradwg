# ✅ Kode lengkap akan dibuat berdasarkan struktur sebelumnya
# ✅ Fokus pada 3 spreadsheet: FDT, DISTRIBUTION CABLE, dan CABLE dari SUBFEEDER
# ✅ Penamaan variabel mengikuti konsistensi dan kemudahan traceability
# ✅ Gunakan fungsi modular agar mudah debugging dan pemeliharaan

import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import tempfile
from datetime import datetime
from math import dist
import re

SPREADSHEET_ID_3 = "1EnteHGDnRhwthlCO9B12zvHUuv3wtq5L2AKlV11qAOU"
SHEET_NAME_3 = "FDT Pekanbaru"

SPREADSHEET_ID_4 = "1D_OMm46yr-e80s3sCyvbSSsf8wrUCwpwiYsVBKPgszw"
SHEET_NAME_4 = "Cable Pekanbaru"

SPREADSHEET_ID_5 = "1paa8sT3nTZh_xxwHeKV8pwVIWacq7lC8U9A8BlX6LUw"
SHEET_NAME_5 = "Sheet1"

def authenticate_google():
    creds_dict = st.secrets["gcp_service_account"]
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(credentials)
    return client

def extract_data_from_kmz(kmz_path):
    folders = {}
    
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
            desc = pm.find("kml:description", ns)
            if nm is not None and coord is not None and ',' in coord.text:
                lon, lat = coord.text.strip().split(",")[:2]
                items.append({
                    "name": nm.text.strip(),
                    "lat": float(lat),
                    "lon": float(lon),
                    "path": new_path,
                    "desc": desc.text.strip() if desc is not None else ""
                })
        return items

    with zipfile.ZipFile(kmz_path, 'r') as zf:
        kml_file = next((f for f in zf.namelist() if f.lower().endswith(".kml")), None)
        if not kml_file:
            st.error("❌ Tidak ditemukan file .kml dalam .kmz")
            return {}

        root = ET.parse(zf.open(kml_file)).getroot()
        ns = {"kml": "http://www.opengis.net/kml/2.2"}
        for folder in root.findall(".//kml:Folder", ns):
            items = recurse_folder(folder, ns)
            for item in items:
                top_folder = item["path"].split("/")[0]
                if top_folder not in folders:
                    folders[top_folder] = []
                folders[top_folder].append(item)

    return folders

def parse_core_capacity(name):
    if "FDT 48" in name:
        return 48, 2, 4, "FDT TYPE 48 CORE"
    elif "FDT 72" in name:
        return 72, 3, 6, "FDT TYPE 72 CORE"
    elif "FDT 96" in name:
        return 96, 4, 8, "FDT TYPE 96 CORE"
    return 0, "", "", ""

def extract_fo_count(text):
    if "FO 24/2T" in text:
        return 2, 24
    elif "FO 48/4T" in text:
        return 4, 48
    elif "FO 36/3T" in text:
        return 3, 36
    elif "FO 12/2T" in text:
        return 2, 12
    elif "FO 96/9T" in text:
        return 9, 96
    return "", ""

def extract_ae_distance(text):
    match = re.search(r"AE\s*-?\s*(\d+)\s*M", text.upper())
    if match:
        return match.group(1)
    return ""

def calculate_path_length(points):
    total = 0.0
    for i in range(1, len(points)):
        total += dist([points[i-1]['lat'], points[i-1]['lon']], [points[i]['lat'], points[i]['lon']])
    return round(total * 111139, 2)  # degrees to meter approx

def find_nearest_pole(fdt_point, pole_list):
    min_dist = float('inf')
    nearest_name = ""
    for pole in pole_list:
        d = dist([fdt_point['lat'], fdt_point['lon']], [pole['lat'], pole['lon']])
        if d < min_dist:
            min_dist = d
            nearest_name = pole['name']
    return nearest_name

def append_fdt_to_sheet(sheet, fdt_points, poles_7_4, district, subdistrict, vendor, kmz_name):
    headers = sheet.row_values(1)
    header_map = {name.strip().lower(): i for i, name in enumerate(headers)}

    values = sheet.get_all_values()
    for row in reversed(values[1:]):
        if any(cell.strip() for cell in row):
            prev_row = row
            break
    else:
        prev_row = ["" for _ in headers]

    today = datetime.today().strftime("%d/%m/%Y") if "/" in prev_row[header_map.get("installation_date", 0)] else datetime.today().strftime("%Y-%m-%d")

    new_rows = []
    for point in fdt_points:
        row = prev_row.copy()
        templatecode = point.get("desc", "")
        row[header_map.get("templatecode")] = templatecode

        row[header_map.get("district")] = district.upper()
        row[header_map.get("subdistrict")] = subdistrict.upper()
        row[header_map.get("site name")] = point["name"]
        row[header_map.get("location name")] = point["name"]
        row[header_map.get("latitude")] = point["lat"]
        row[header_map.get("longitude")] = point["lon"]
        row[header_map.get("vendor name")] = vendor.upper()
        row[header_map.get("as built vendor")] = vendor.upper()
        row[header_map.get("installation_date")] = today
        row[header_map.get("cluster name")] = kmz_name

        _, col_m, col_r, fdt_type = parse_core_capacity(point["name"])
        row[header_map.get("col m")] = col_m
        row[header_map.get("col r")] = col_r
        row[header_map.get("fat type")] = fdt_type

        row[header_map.get("parentid 1")] = find_nearest_pole(point, poles_7_4)

        new_rows.append(row)

    sheet.append_rows(new_rows)
    st.success(f"✅ {len(new_rows)} FDT berhasil ditambahkan ke {SHEET_NAME_3}")

def append_cable_pekanbaru(sheet, items, district, subdistrict, vendor, kmz_name):
    headers = sheet.row_values(1)
    header_map = {name.strip().lower(): i for i, name in enumerate(headers)}
    
    new_rows = []
    for item in items:
        count, core = extract_fo_count(item['name'])
        ae = extract_ae_distance(item['name'])

        row = ["" for _ in headers]
        row[header_map.get("cluster name")] = kmz_name
        row[header_map.get("fo type")] = item['name']
        row[header_map.get("vendor")] = vendor.upper()
        row[header_map.get("district")] = district.upper()
        row[header_map.get("subdistrict")] = subdistrict.upper()
        row[header_map.get("core")] = core
        row[header_map.get("ae")] = ae
        row[header_map.get("tray")] = count
        new_rows.append(row)

    sheet.append_rows(new_rows)
    st.success(f"✅ {len(new_rows)} cable distribusi berhasil ditambahkan ke {SHEET_NAME_4}")

def append_subfeeder_cable(sheet, items, district, subdistrict, vendor, kmz_name):
    headers = sheet.row_values(1)
    header_map = {name.strip().lower(): i for i, name in enumerate(headers)}

    new_rows = []
    for item in items:
        count, core = extract_fo_count(item['name'])
        ae = extract_ae_distance(item['name'])
        row = ["" for _ in headers]
        row[header_map.get("fo type")] = item['name']
        row[header_map.get("vendor")] = vendor.upper()
        row[header_map.get("district")] = district.upper()
        row[header_map.get("subdistrict")] = subdistrict.upper()
        row[header_map.get("cluster name")] = kmz_name
        row[header_map.get("core")] = core
        row[header_map.get("ae")] = ae
        row[header_map.get("tray")] = count
        new_rows.append(row)

    sheet.append_rows(new_rows)
    st.success(f"✅ {len(new_rows)} kabel subfeeder berhasil ditambahkan ke {SHEET_NAME_5}")

# 🔽 MAIN STREAMLIT UI

def main():
    st.title("📌 KMZ to Google Sheets - Auto Mapper")

    kmz_file = st.file_uploader("📤 Upload file .kmz", type="kmz")
    district = st.text_input("🗺️ District")
    subdistrict = st.text_input("🏙️ Subdistrict")
    vendor = st.text_input("🏗️ Vendor")

    if kmz_file and district and subdistrict and vendor:
        with st.spinner("🔍 Memproses KMZ..."):
            folders = extract_data_from_kmz(kmz_file)
            kmz_name = kmz_file.name.replace(".kmz", "")
            client = authenticate_google()

            if 'FDT' in folders:
                sheet = client.open_by_key(SPREADSHEET_ID_3).worksheet(SHEET_NAME_3)
                poles_7_4 = folders.get("EXISTING POLE EMR 7-4", [])
                append_fdt_to_sheet(sheet, folders['FDT'], poles_7_4, district, subdistrict, vendor, kmz_name)
            if 'DISTRIBUTION CABLE' in folders:
                sheet = client.open_by_key(SPREADSHEET_ID_4).worksheet(SHEET_NAME_4)
                append_cable_pekanbaru(sheet, folders['DISTRIBUTION CABLE'], district, subdistrict, vendor, kmz_name)
            if 'SUBFEEDER CABLE' in folders:
                sheet = client.open_by_key(SPREADSHEET_ID_5).worksheet(SHEET_NAME_5)
                append_subfeeder_cable(sheet, folders['SUBFEEDER CABLE'], district, subdistrict, vendor, kmz_name)

            st.success("✅ Semua data berhasil diproses dan dikirim ke Spreadsheet!")

if __name__ == "__main__":
    main()
