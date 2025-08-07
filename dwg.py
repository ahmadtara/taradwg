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
from datetime import datetime

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

def extract_kmz_data_combined(kmz_path):
    with zipfile.ZipFile(kmz_path, 'r') as zf:
        kml_filename = [f for f in zf.namelist() if f.endswith('.kml')][0]
        with zf.open(kml_filename) as f:
            doc = f.read()

    k = kml.KML()
    k.from_string(doc)

    def extract_features(features):
        items = []
        for feature in features:
            if hasattr(feature, 'geometry'):
                geom = feature.geometry
                props = {
                    'name': feature.name,
                    'description': feature.description or "",
                    'geometry': geom
                }
                items.append(props)
            if hasattr(feature, 'features'):
                items.extend(extract_features(feature.features()))
        return items

    all_features = extract_features(k.features())
    return all_features

def parse_folder(folder):
    folders = {
        "FDT": [],
        "HP COVER": [],
        "FAT": [],
        "NEW POLE 7-3": [],
        "NEW POLE 7-4": [],
        "EXISTING POLE EMR 7-3": [],
        "EXISTING POLE EMR 7-4": [],
        "DISTRIBUTION": [],
        "SUBFEEDER": []
    }

    for root, dirs, files in os.walk(folder):
        for file in files:
            if file.endswith('.kmz'):
                path = os.path.join(root, file)
                parts = os.path.normpath(path).split(os.sep)
                for key in folders.keys():
                    if key in parts:
                        features = extract_kmz_data_combined(path)
                        for feature in features:
                            folders[key].append({
                                "name": feature['name'],
                                "description": feature['description'],
                                "geometry": feature['geometry'],
                                "full_path": path,
                                "folder_name": key
                            })
    return folders

    
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

    if kmz_subfeeder_file and district and subdistrict and vendor:
        with st.spinner("🔍 Memproses KMZ Subfeeder..."):
            folders, _ = extract_kmz_data_combined(kmz_subfeeder_file)
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








