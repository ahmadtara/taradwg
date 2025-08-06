from datetime import datetime
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import streamlit as st
import zipfile
import os
import tempfile
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account

from sheet_writer import (
    append_poles_to_main_sheet,
    append_fat_to_sheet,
    append_to_fdt_sheet,
    append_to_cable_cluster_sheet,
    append_to_subfeeder_cable_sheet
)
from kmz_parser import extract_points_from_kmz

SPREADSHEET_ID = "1yXBIuX2LjUWxbpnNqf6A9YimtG7d77V_AHLidhWKIS8"
SPREADSHEET_ID_2 = "1WI0Gb8ul5GPUND4ADvhFgH4GSlgwq1_4rRgfOnPz-yc"
SPREADSHEET_ID_3 = "1EnteHGDnRhwthlCO9B12zvHUuv3wtq5L2AKlV11qAOU"
SPREADSHEET_ID_4 = "1D_OMm46yr-e80s3sCyvbSSsf8wrUCwpwiYsVBKPgszw"
SPREADSHEET_ID_5 = "1paa8sT3nTZh_xxwHeKV8pwVIWacq7lC8U9A8BlX6LUw"

SHEET_NAME = "Pole Pekanbaru"
SHEET_NAME_2 = "Fat Pekanbaru"
SHEET_NAME_3 = "FDT Pekanbaru"
SHEET_NAME_4 = "Cable Pekanbaru"
SHEET_NAME_5 = "Sheet1"

GDRIVE_FOLDERS = {
    "DISTRIBUTION CABLE": "1XkWqvRX4SUYMrtMQ7vt8197oSja4r9p-",
    "BOUNDARY CLUSTER": "1IMpaQWnpG8c8P5j3phUMP1G9zTPBDQMi",
    "CABLE": "16aesqK-OIqYIDAIn_ymLzf1-VkLyXonl"
}

def get_client():
    creds_dict = st.secrets["gcp_service_account"]
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(credentials)
    return client

def upload_kml_to_drive(kmz_path):
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=['https://www.googleapis.com/auth/drive']
    )
    drive_service = build('drive', 'v3', credentials=creds)

    with zipfile.ZipFile(kmz_path, 'r') as zf:
        for folder_name in GDRIVE_FOLDERS:
            for file in zf.namelist():
                if folder_name in file.upper() and file.lower().endswith(".kml"):
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".kml") as tmp:
                        tmp.write(zf.read(file))
                        tmp_path = tmp.name

                    file_metadata = {
                        'name': os.path.basename(file),
                        'parents': [GDRIVE_FOLDERS[folder_name]]
                    }
                    media = MediaFileUpload(tmp_path, mimetype='application/vnd.google-earth.kml+xml')

                    drive_service.files().create(
                        body=file_metadata,
                        media_body=media,
                        fields='id'
                    ).execute()
                    st.success(f"📄 File {file} berhasil diupload ke folder {folder_name} di Google Drive.")

# === STREAMLIT INTERFACE ===
st.set_page_config(page_title="Uploader Pole KMZ", layout="centered")
st.title("📡 Uploader Pole KMZ (CLUSTER + SUBFEEDER + FAT SPLITTER)")

col1, col2, col3 = st.columns(3)
with col1:
    district_input = st.text_input("District (E)")
with col2:
    subdistrict_input = st.text_input("Subdistrict (F)")
with col3:
    vendor_input = st.text_input("Vendor Name (AB)")

uploaded_cluster = st.file_uploader("📤 Upload file .KMZ CLUSTER (berisi FAT & NEW POLE)", type=["kmz"])
uploaded_subfeeder = st.file_uploader("📤 Upload file .KMZ SUBFEEDER (berisi NEW POLE 7-4 / 9-4)", type=["kmz"])

submit_clicked = st.button("🚀 Submit dan Kirim ke Google Sheet")

if submit_clicked:
    if not district_input or not subdistrict_input or not vendor_input:
        st.warning("⚠️ Harap isi semua kolom input manual.")
    elif not uploaded_cluster:
        st.warning("⚠️ Harap upload file KMZ CLUSTER.")
    else:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".kmz") as tmp:
            tmp.write(uploaded_cluster.read())
            kmz_path = tmp.name

        with st.spinner("🔍 Membaca data dari KMZ CLUSTER..."):
            fat_points, poles_cluster, poles_subfeeder, fdt_points, cable_cluster, _ = extract_points_from_kmz(kmz_path)

        try:
            client = get_client()
            sheet1 = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
            if poles_cluster:
                append_poles_to_main_sheet(sheet1, poles_cluster, district_input, subdistrict_input, vendor_input)

            if fat_points:
                sheet2 = client.open_by_key(SPREADSHEET_ID_2).worksheet(SHEET_NAME)
                append_fat_to_sheet(sheet2, fat_points, poles_subfeeder, district_input, subdistrict_input, vendor_input)
            else:
                st.warning("⚠️ Tidak ditemukan folder FAT dalam file KMZ.")

            if fdt_points:
                sheet3 = client.open_by_key(SPREADSHEET_ID_3).worksheet(SHEET_NAME_3)
                append_to_fdt_sheet(sheet3, fdt_points, poles_cluster, kmz_path, district_input, subdistrict_input, vendor_input)

            if cable_cluster:
                sheet4 = client.open_by_key(SPREADSHEET_ID_4).worksheet(SHEET_NAME_4)
                append_to_cable_cluster_sheet(sheet4, cable_cluster, vendor_input, kmz_path)

            upload_kml_to_drive(kmz_path)

        except Exception as e:
            st.error(f"❌ Gagal mengirim data dari CLUSTER: {e}")

    if uploaded_subfeeder:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".kmz") as tmp:
            tmp.write(uploaded_subfeeder.read())
            kmz_path = tmp.name

        with st.spinner("🔍 Membaca data dari KMZ SUBFEEDER..."):
            _, poles_subonly, _, _, _, cable_subfeeder = extract_points_from_kmz(kmz_path)

        try:
            client = get_client()
            sheet1 = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
            if poles_subonly:
                append_poles_to_main_sheet(sheet1, poles_subonly, district_input, subdistrict_input, vendor_input)

            if cable_subfeeder:
                sheet5 = client.open_by_key(SPREADSHEET_ID_5).worksheet(SHEET_NAME_5)
                append_to_subfeeder_cable_sheet(sheet5, cable_subfeeder, vendor_input, kmz_path)

            upload_kml_to_drive(kmz_path)

        except Exception as e:
            st.error(f"❌ Gagal mengirim data SUBFEEDER: {e}")

