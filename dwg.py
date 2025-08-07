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

def extract_kmz_data_combined(kmz_file):
    folders = {}
    poles = []

    def recurse_folder(folder, ns, path=""):
        items = []
        name_el = folder.find("kml:name", ns)
        folder_name = name_el.text.strip().upper() if name_el is not None else "UNKNOWN"
        new_path = f"{path}/{folder_name}" if path else folder_name

        # Tambahkan folder ke dictionary utama
        if folder_name not in folders:
            folders[folder_name] = []

        # Cek semua placemark di folder ini
        for placemark in folder.findall("kml:Placemark", ns):
            name_tag = placemark.find("kml:name", ns)
            name = name_tag.text.strip() if name_tag is not None else ""

            coords_tag = placemark.find(".//kml:coordinates", ns)
            coords = coords_tag.text.strip().split(",") if coords_tag is not None and coords_tag.text else ["", ""]
            lon, lat = coords[:2] if len(coords) >= 2 else ("", "")

            description_tag = placemark.find("kml:description", ns)
            description = description_tag.text.strip() if description_tag is not None else ""

            item = {
                "name": name,
                "lon": float(lon) if lon else None,
                "lat": float(lat) if lat else None,
                "description": description,
                "folder": folder_name,
                "full_path": new_path
            }

            folders[folder_name].append(item)
            items.append(item)

            # Deteksi jenis tiang dari folder_name
            if folder_name == "NEW POLE 7-3":
                poles.append({**item, "folder": "7m3inch", "height": "7"})
            elif folder_name == "NEW POLE 7-4":
                poles.append({**item, "folder": "7m4inch", "height": "7"})
            elif folder_name == "NEW POLE 9-4":
                poles.append({**item, "folder": "9m4inch", "height": "9"})
            elif folder_name == "EXISTING POLE EMR 7-3":
                poles.append({**item, "folder": "ext7m3inch", "height": "7"})
            elif folder_name == "EXISTING POLE EMR 7-4":
                poles.append({**item, "folder": "ext7m4inch", "height": "7"})
            elif folder_name == "EXISTING POLE EMR 9-4":
                poles.append({**item, "folder": "ext9m4inch", "height": "9"})

        # Rekursi ke subfolder
        for subfolder in folder.findall("kml:Folder", ns):
            items += recurse_folder(subfolder, ns, new_path)

        return items

    # Buka KMZ dan parse file KML di dalamnya
    with zipfile.ZipFile(kmz_file, 'r') as z:
        kml_filename = next((f for f in z.namelist() if f.lower().endswith('.kml')), None)
        if not kml_filename:
            raise ValueError("❌ Tidak ditemukan file .kml dalam .kmz")

        with z.open(kml_filename) as kml_file:
            tree = ET.parse(kml_file)
            root = tree.getroot()
            ns = {'kml': 'http://www.opengis.net/kml/2.2'}

            # Telusuri semua Folder utama
            for folder in root.findall(".//kml:Folder", ns):
                recurse_folder(folder, ns)

    return folders, poless
    
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

def is_float(value):
    try:
        float(value)
        return True
    except (TypeError, ValueError):
        return False

def append_fdt_to_sheet(sheet, fdt_data, poles, district, subdistrict, vendor, kmz_name):
    existing_rows = sheet.get_all_values()
    headers = sheet.row_values(1)
    header_map = {name.strip().lower(): i for i, name in enumerate(headers)}
    template_row = existing_rows[-1] if len(existing_rows) > 1 else []
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
        row[24:26] = template_row[24:26]
        row[26] = template_row[26]
        row[29] = template_row[29]
        row[30] = template_row[30]
        row[40] = template_row[40]
        row[41] = kolom_ap
        row[33] = datetime.today().strftime("%d/%m/%Y")
        row[31] = vendor
        row[44] = vendor

        idx_an = header_map.get('parentid 1')
        if idx_an is not None:
            row[idx_an] = find_nearest_pole(fdt, [
                p for p in poles if p['folder'] in ['7m4inch', '7m3inch', 'ext7m3inch', 'ext7m4inch', 'ext9m4inch']
            ])

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
            folders = extract_kmz_data_combined(kmz_fdt_file)
            kmz_name = kmz_fdt_file.name.replace(".kmz", "")
            if client is None:
                client = authenticate_google()

            poles = folders.get("NEW POLE 7-4", [])
            if poles:
                st.success(f"✅ {len(poles)} titik tiang dari 'NEW POLE 7-4' berhasil diambil.")
            else:
                st.warning("⚠️ Tidak ditemukan titik tiang di folder 'NEW POLE 7-4' dalam KMZ.")

            if 'FDT' in folders:
                sheet = client.open_by_key(SPREADSHEET_ID_3).worksheet(SHEET_NAME_3)
                count_fdt = append_fdt_to_sheet(sheet, folders['FDT'], poles, district, subdistrict, vendor, kmz_name)

            if 'DISTRIBUTION CABLE' in folders:
                sheet = client.open_by_key(SPREADSHEET_ID_4).worksheet(SHEET_NAME_4)
                count_cable = append_cable_pekanbaru(sheet, folders['DISTRIBUTION CABLE'], district, subdistrict, vendor, kmz_name)

    if kmz_subfeeder_file and district and subdistrict and vendor:
        with st.spinner("🔍 Memproses KMZ Subfeeder..."):
            folders = extract_kmz_data_combined(kmz_subfeeder_file)
            kmz_name = kmz_subfeeder_file.name.replace(".kmz", "")
            if client is None:
                client = authenticate_google()

            if 'CABLE' in folders:
                sheet = client.open_by_key(SPREADSHEET_ID_5).worksheet(SHEET_NAME_5)
                count_subfeeder = append_subfeeder_cable(sheet, folders['CABLE'], district, subdistrict, vendor, kmz_name)

    if (kmz_fdt_file or kmz_subfeeder_file) and district and subdistrict and vendor:
        st.success("✅ Semua data berhasil diproses dan dikirim ke Spreadsheet!")
        st.info(f"🛰️ {count_fdt} FDT dikirim ke spreadsheet FDT Pekanbaru")
        st.info(f"📦 {count_cable} kabel distribusi dikirim ke Cable Pekanbaru")
        st.info(f"🔌 {count_subfeeder} kabel subfeeder dikirim ke Sheet1")

if __name__ == "__main__":
    main()


