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

# ... fungsi lainnya tidak berubah ...

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

    if kmz_fdt_file and district and subdistrict and vendor:
        with st.spinner("🔍 Memproses KMZ FDT..."):
            folders = extract_data_from_kmz(kmz_fdt_file)
            kmz_name = kmz_fdt_file.name.replace(".kmz", "")
            if client is None:
                client = authenticate_google()

            if 'FDT' in folders:
                sheet = client.open_by_key(SPREADSHEET_ID_3).worksheet(SHEET_NAME_3)
                poles_7_4 = folders.get("EXISTING POLE EMR 7-4", [])
                append_fdt_to_sheet(sheet, folders['FDT'], poles_7_4, district, subdistrict, vendor, kmz_name)
            if 'DISTRIBUTION CABLE' in folders:
                sheet = client.open_by_key(SPREADSHEET_ID_4).worksheet(SHEET_NAME_4)
                append_cable_pekanbaru(sheet, folders['DISTRIBUTION CABLE'], district, subdistrict, vendor, kmz_name)

    if kmz_subfeeder_file and district and subdistrict and vendor:
        with st.spinner("🔍 Memproses KMZ Subfeeder..."):
            folders = extract_data_from_kmz(kmz_subfeeder_file)
            kmz_name = kmz_subfeeder_file.name.replace(".kmz", "")
            if client is None:
                client = authenticate_google()

            if 'SUBFEEDER CABLE' in folders:
                sheet = client.open_by_key(SPREADSHEET_ID_5).worksheet(SHEET_NAME_5)
                append_subfeeder_cable(sheet, folders['SUBFEEDER CABLE'], district, subdistrict, vendor, kmz_name)

    if (kmz_fdt_file or kmz_subfeeder_file) and district and subdistrict and vendor:
        st.success("✅ Semua data berhasil diproses dan dikirim ke Spreadsheet!")

if __name__ == "__main__":
    main()
