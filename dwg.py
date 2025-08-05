import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import tempfile
from datetime import datetime
import os

try:
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
    from google.oauth2 import service_account
except ModuleNotFoundError as e:
    st.error(f"❌ Modul tidak ditemukan: {e}. Coba pastikan `google-api-python-client` sudah terinstall.")
    st.stop()

st.set_page_config(page_title="Uploader FAT Splitter", layout="centered")

st.title("📡 Uploader FAT Splitter")

col1, col2, col3 = st.columns(3)
with col1:
    district_input = st.text_input("District (E)")
with col2:
    subdistrict_input = st.text_input("Subdistrict (F)")
with col3:
    vendor_input = st.text_input("Vendor Name (AB)")

uploaded_cluster = st.file_uploader("📄 Upload file .KMZ CLUSTER (berisi FAT & NEW POLE)", type=["kmz"])
uploaded_subfeeder = st.file_uploader("📄 Upload file .KMZ SUBFEEDER (berisi NEW POLE 7-4 / 9-4)", type=["kmz"])

submit_clicked = st.button("🚀 Submit dan Kirim ke Google Sheet")

st.write("---")
st.write("✅ Aplikasi dimulai...")

dist = __import__('math').dist

SPREADSHEET_ID = "1yXBIuX2LjUWxbpnNqf6A9YimtG7d77V_AHLidhWKIS8"
SPREADSHEET_ID_2 = "1WI0Gb8ul5GPUND4ADvhFgH4GSlgwq1_4rRgfOnPz-yc"
SHEET_NAME = "Pole Pekanbaru"
SHEET_NAME_2 = "FAT Pekanbaru"

GDRIVE_FOLDERS = {
    "DISTRIBUTION CABLE": "1XkWqvRX4SUYMrtMQ7vt8197oSja4r9p-",
    "BOUNDARY CLUSTER": "1IMpaQWnpG8c8P5j3phUMP1G9zTPBDQMi",
    "CABLE": "16aesqK-OIqYIDAIn_ymLzf1-VkLyXonl"
}

_cached_headers = None
_cached_prev_row = None

def upload_kml_to_drive(kmz_path):
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=['https://www.googleapis.com/auth/drive']
    )
    drive_service = build('drive', 'v3', credentials=creds)

    with zipfile.ZipFile(kmz_path, 'r') as zf:
        kml_filename = next((f for f in zf.namelist() if f.lower().endswith(".kml")), None)
        if not kml_filename:
            st.warning("⚠️ File .kml tidak ditemukan dalam KMZ.")
            return

        with tempfile.NamedTemporaryFile(delete=False, suffix=".kml") as tmp_kml:
            tmp_kml.write(zf.read(kml_filename))
            tmp_kml_path = tmp_kml.name

    new_filename = os.path.splitext(os.path.basename(kmz_path))[0] + ".kml"

    for folder_name, folder_id in GDRIVE_FOLDERS.items():
        file_metadata = {
            'name': new_filename,
            'parents': [folder_id]
        }
        media = MediaFileUpload(tmp_kml_path, mimetype='application/vnd.google-earth.kml+xml')

        drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()

    st.success(f"📄 File {new_filename} berhasil diupload ke semua folder Google Drive.")

if submit_clicked:
    if uploaded_cluster:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".kmz") as tmp:
            tmp.write(uploaded_cluster.read())
            kmz_path = tmp.name
        upload_kml_to_drive(kmz_path)

# ... (fungsi lainnya tetap tidak berubah sampai bawah)
