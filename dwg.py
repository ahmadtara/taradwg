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

def extract_data_from_kmz(kmz_file):
    folders = {}
    with zipfile.ZipFile(kmz_file, 'r') as z:
        kml_filename = [f for f in z.namelist() if f.endswith('.kml')][0]
        with z.open(kml_filename) as kml_file:
            tree = ET.parse(kml_file)
            root = tree.getroot()

            ns = {'kml': 'http://www.opengis.net/kml/2.2'}
            for folder in root.findall(".//kml:Folder", ns):
                folder_name = folder.find("kml:name", ns).text.strip()
                placemarks = folder.findall("kml:Placemark", ns)

                results = []
                for placemark in placemarks:
                    name = placemark.find("kml:name", ns)
                    name = name.text.strip() if name is not None else ""

                    coords_tag = placemark.find(".//kml:coordinates", ns)
                    coords = coords_tag.text.strip() if coords_tag is not None else ""
                    coords = coords.split(",") if coords else ["", "", ""]

                    description_tag = placemark.find("kml:description", ns)
                    description = description_tag.text.strip() if description_tag is not None else ""

                    results.append({
                        'name': name,
                        'lon': coords[0],
                        'lat': coords[1],
                        'alt': coords[2] if len(coords) > 2 else "",
                        'description': description
                    })

                folders[folder_name] = results
    return folders

def append_fdt_to_sheet(sheet, fdt_data, pole_data, district, subdistrict, vendor, kmz_name):
    existing_rows = sheet.get_all_values()
    template_row = existing_rows[-1] if len(existing_rows) > 1 else []
    rows = []
    for fdt in fdt_data:
        name = fdt['name']
        lat = fdt['lat']
        lon = fdt['lon']
        desc = fdt.get('description', '')

        nearest_pole = min(pole_data, key=lambda p: dist([
            float(lat), float(lon)
        ], [float(p['lat']), float(p['lon'])])) if pole_data else {}

        # Kolom logika berdasarkan kapasitas FDT
        capacity = int(re.findall(r'\d+', name)[0]) if re.findall(r'\d+', name) else 0
        kolom_m = capacity // 24 if capacity else ""
        kolom_r = (capacity * 2) // 24 if capacity else ""
        kolom_ap = f"FDT TYPE {capacity} CORE" if capacity else ""

        row = [""] * 45
        row[0] = desc                       # A
        row[1:5] = template_row[1:5]         # B-E
        row[5] = district                    # F
        row[6] = subdistrict                # G
        row[7] = kmz_name                   # H
        row[8] = name                       # I
        row[9] = name                       # J
        row[10] = lat                       # K
        row[11] = lon                       # L
        row[12] = kolom_m                   # M
        row[13:14] = template_row[13:14]    # N
        row[14:16] = template_row[14:16]    # O, P
        row[17] = kolom_r                   # R
        row[18] = template_row[18]          # S
        row[24:27] = template_row[24:27]  
        row[26:30] = template_row[26:30]    # AA, AB, AC, AD
        row[40] = template_row[40]          # AE
        row[33] = datetime.today().strftime("%d/%m/%Y")  # AH
        row[39] = nearest_pole.get('name', '')  # AN / Parentid 1
        row[44] = vendor                    # AS
        row[29:30] = template_row[29:30]           

        rows.append(row)

    sheet.append_rows(rows, value_input_option="USER_ENTERED")
    return len(rows)

def append_cable_pekanbaru(sheet, cable_data, district, subdistrict, vendor, kmz_name):
    rows = []
    for cable in cable_data:
        row = [
            cable['name'],
            cable['lat'],
            cable['lon'],
            district,
            subdistrict,
            vendor,
            kmz_name,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ]
        rows.append(row)

    sheet.append_rows(rows, value_input_option="USER_ENTERED")
    return len(rows)

def append_subfeeder_cable(sheet, cable_data, district, subdistrict, vendor, kmz_name):
    rows = []
    for cable in cable_data:
        row = [
            cable['name'],
            cable['lat'],
            cable['lon'],
            district,
            subdistrict,
            vendor,
            kmz_name,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ]
        rows.append(row)

    sheet.append_rows(rows, value_input_option="USER_ENTERED")
    return len(rows)

# 🔽 MAIN STREAMLIT UI

def main():
    st.title("📌 KMZ to Google Sheets - Auto Mapper")

    col1, col2 = st.columns(2)
    with col1:
        kmz_fdt_file = st.file_uploader("📤 Upload file .kmz Cluster (FDT)", type="kmz", key="fdt")
    with col2:
        kmz_subfeeder_file = st.file_uploader("📤 Upload file .kmz Subfeeder", type="kmz", key="subfeeder")

    district = st.text_input("🗺️ District")
    subdistrict = st.text_input("🏙️ Subdistrict")
    vendor = st.text_input("🏗️ Vendor")

    client = None

    count_fdt = 0
    count_cable = 0
    count_subfeeder = 0

    if kmz_fdt_file and district and subdistrict and vendor:
        with st.spinner("🔍 Memproses KMZ FDT..."):
            folders = extract_data_from_kmz(kmz_fdt_file)
            kmz_name = kmz_fdt_file.name.replace(".kmz", "")
            if client is None:
                client = authenticate_google()

            if 'FDT' in folders:
                sheet = client.open_by_key(SPREADSHEET_ID_3).worksheet(SHEET_NAME_3)
                poles_7_4 = folders.get("NEW POLE 7-4", [])
                count_fdt = append_fdt_to_sheet(sheet, folders['FDT'], poles_7_4, district, subdistrict, vendor, kmz_name)

            if 'DISTRIBUTION CABLE' in folders:
                sheet = client.open_by_key(SPREADSHEET_ID_4).worksheet(SHEET_NAME_4)
                count_cable = append_cable_pekanbaru(sheet, folders['DISTRIBUTION CABLE'], district, subdistrict, vendor, kmz_name)

    if kmz_subfeeder_file and district and subdistrict and vendor:
        with st.spinner("🔍 Memproses KMZ Subfeeder..."):
            folders = extract_data_from_kmz(kmz_subfeeder_file)
            kmz_name = kmz_subfeeder_file.name.replace(".kmz", "")
            if client is None:
                client = authenticate_google()

            if 'SUBFEEDER CABLE' in folders:
                sheet = client.open_by_key(SPREADSHEET_ID_5).worksheet(SHEET_NAME_5)
                count_subfeeder = append_subfeeder_cable(sheet, folders['SUBFEEDER CABLE'], district, subdistrict, vendor, kmz_name)

    if (kmz_fdt_file or kmz_subfeeder_file) and district and subdistrict and vendor:
        st.success("✅ Semua data berhasil diproses dan dikirim ke Spreadsheet!")
        st.info(f"🛰️ {count_fdt} FDT dikirim ke spreadsheet FDT Pekanbaru")
        st.info(f"📦 {count_cable} kabel distribusi dikirim ke Cable Pekanbaru")
        st.info(f"🔌 {count_subfeeder} kabel subfeeder dikirim ke Sheet1")

if __name__ == "__main__":
    main()






