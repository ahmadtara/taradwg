import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import gspread
from fastkml import kml
from oauth2client.service_account import ServiceAccountCredentials
from append_fdt_to_sheet import append_fdt_to_sheet
from append_cable_pekanbaru import append_cable_pekanbaru
from append_subfeeder_cable import append_subfeeder_cable
from append_fat_to_sheet import append_fat_to_sheet
from append_poles_to_main_sheet import append_poles_to_main_sheet
from datetime import datetime
import tempfile

dist = __import__('math').dist

SPREADSHEET_ID_3 = "1EnteHGDnRhwthlCO9B12zvHUuv3wtq5L2AKlV11qAOU"
SHEET_NAME_3 = "FDT Pekanbaru"

SPREADSHEET_ID_4 = "1D_OMm46yr-e80s3sCyvbSSsf8wrUCwpwiYsVBKPgszw"
SHEET_NAME_4 = "Cable Pekanbaru"

SPREADSHEET_ID_5 = "1paa8sT3nTZh_xxwHeKV8pwVIWacq7lC8U9A8BlX6LUw"
SHEET_NAME_5 = "Sheet1"

SPREADSHEET_ID = "1yXBIuX2LjUWxbpnNqf6A9YimtG7d77V_AHLidhWKIS8"
SHEET_NAME = "Pole Pekanbaru"

SPREADSHEET_ID_2 = "1WI0Gb8ul5GPUND4ADvhFgH4GSlgwq1_4rRgfOnPz-yc"
SHEET_NAME_2 = "FAT Pekanbaru"

_cached_headers = None
_cached_prev_row = None

def authenticate_google():
    creds_dict = st.secrets["gcp_service_account"]
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(credentials)
    return client

def extract_kmz_data_combined(kmz_file):
    folders = {}
    poles = []
    seen_items = set()

    def recurse_folder(folder, ns, path=""):
        name_el = folder.find("kml:name", ns)
        folder_name = name_el.text.strip().upper() if name_el is not None else "UNKNOWN"
        current_path = f"{path}/{folder_name}" if path else folder_name

        if folder_name not in folders:
            folders[folder_name] = []

        for placemark in folder.findall("kml:Placemark", ns):
            name_tag = placemark.find("kml:name", ns)
            name = name_tag.text.strip() if name_tag is not None else ""

            coords_tag = placemark.find(".//kml:coordinates", ns)
            coords = coords_tag.text.strip().split(",") if coords_tag is not None and coords_tag.text else ["", ""]
            lon, lat = coords[:2] if len(coords) >= 2 else (None, None)

            description_tag = placemark.find("kml:description", ns)
            description = description_tag.text.strip() if description_tag is not None else ""

            unique_key = (name, lon, lat, folder_name)
            if unique_key in seen_items:
                continue
            seen_items.add(unique_key)

            item = {
                "name": name,
                "lon": float(lon) if lon else None,
                "lat": float(lat) if lat else None,
                "description": description,
                "folder": folder_name,
                "full_path": current_path
            }

            folders[folder_name].append(item)

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

        for subfolder in folder.findall("kml:Folder", ns):
            recurse_folder(subfolder, ns, current_path)

    with zipfile.ZipFile(kmz_file, 'r') as z:
        kml_filename = next((f for f in z.namelist() if f.lower().endswith('.kml')), None)
        if not kml_filename:
            raise ValueError("❌ Tidak ditemukan file .kml dalam .kmz")

        with z.open(kml_filename) as kml_file:
            tree = ET.parse(kml_file)
            root = tree.getroot()
            ns = {'kml': 'http://www.opengis.net/kml/2.2'}

            for folder in root.findall(".//kml:Folder", ns):
                recurse_folder(folder, ns)

    return folders, poles

def main():
    st.title("📌 KMZ to Google Sheets - Auto Mapper")

    col1, col2 = st.columns(2)
    with col1:
        kmz_fdt_file = st.file_uploader("📤 Upload file .kmz Cluster (FDT, FAT & NEW POLE)", type="kmz", key="fdt")
    with col2:
        kmz_subfeeder_file = st.file_uploader("📤 Upload file .kmz Subfeeder (NEW POLE 7-4 / 9-4)", type="kmz", key="subfeeder")

    col1, col2, col3 = st.columns(3)
    with col1:
        district = st.text_input("🗺️ District (E)")
    with col2:
        subdistrict = st.text_input("🏙️ Subdistrict (F)")
    with col3:
        vendor = st.text_input("🏗️ Vendor Name (AB)")

    submit = st.button("🚀 Submit & Kirim ke Spreadsheet")

    if submit:
        client = None
        count_fdt = 0
        count_cable = 0
        count_subfeeder = 0
        count_fat = 0
        count_pole = 0

        if kmz_fdt_file and district and subdistrict and vendor:
            with st.spinner("🔍 Memproses KMZ FDT..."):
                folders, poles = extract_kmz_data_combined(kmz_fdt_file)
                kmz_name = kmz_fdt_file.name.replace(".kmz", "")
                if client is None:
                    client = authenticate_google()

                if 'FDT' in folders:
                    sheet = client.open_by_key(SPREADSHEET_ID_3).worksheet(SHEET_NAME_3)
                    count_fdt = append_fdt_to_sheet(sheet, folders['FDT'], poles, district, subdistrict, vendor, kmz_name)

                if 'DISTRIBUTION CABLE' in folders:
                    sheet = client.open_by_key(SPREADSHEET_ID_4).worksheet(SHEET_NAME_4)
                    count_cable = append_cable_pekanbaru(sheet, folders['DISTRIBUTION CABLE'], district, subdistrict, vendor, kmz_name)

                if 'FAT' in folders:
                    sheet = client.open_by_key(SPREADSHEET_ID_2).worksheet(SHEET_NAME_2)
                    append_fat_to_sheet(sheet, folders['FAT'], poles, district, subdistrict, vendor)
                    count_fat = len(folders['FAT'])

                pole_keys = ['NEW POLE 7-3', 'NEW POLE 7-4', 'NEW POLE 9-4', 'EXISTING POLE EMR 7-3', 'EXISTING POLE EMR 7-4', 'EXISTING POLE EMR 9-4']
                poles_to_append = [item for key in pole_keys if key in folders for item in folders[key]]
                if poles_to_append:
                    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
                    append_poles_to_main_sheet(sheet, poles_to_append, district, subdistrict, vendor)
                    count_pole = len(poles_to_append)

        if kmz_subfeeder_file and district and subdistrict and vendor:
            with st.spinner("🔍 Memproses KMZ Subfeeder..."):
                folders, poles = extract_kmz_data_combined(kmz_subfeeder_file)
                kmz_name = kmz_subfeeder_file.name.replace(".kmz", "")
                if client is None:
                    client = authenticate_google()

                if 'CABLE' in folders:
                    sheet = client.open_by_key(SPREADSHEET_ID_5).worksheet(SHEET_NAME_5)
                    count_subfeeder = append_subfeeder_cable(sheet, folders['CABLE'], district, subdistrict, vendor, kmz_name)

                pole_keys = ['NEW POLE 7-4', 'NEW POLE 9-4']
                poles_to_append = [item for key in pole_keys if key in folders for item in folders[key]]
                if poles_to_append:
                    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
                    append_poles_to_main_sheet(sheet, poles_to_append, district, subdistrict, vendor)
                    count_pole += len(poles_to_append)

        if (kmz_fdt_file or kmz_subfeeder_file) and district and subdistrict and vendor:
            st.success("✅ Semua data berhasil diproses dan dikirim ke Spreadsheet!")
            st.info(f"🛰️ {count_fdt} FDT dikirim ke spreadsheet FDT Pekanbaru")
            st.info(f"📦 {count_cable} kabel distribusi dikirim ke Cable Pekanbaru")
            st.info(f"🔌 {count_subfeeder} kabel subfeeder dikirim ke Sheet1")
            st.info(f"🗼 {count_pole} pole dikirim ke spreadsheet Pole Pekanbaru")
            st.info(f"🧭 {count_fat} FAT dikirim ke spreadsheet FAT Pekanbaru")
        else:
            st.warning("⚠️ Mohon lengkapi semua input dan upload file yang diperlukan sebelum Submit.")

if __name__ == "__main__":
    main()
