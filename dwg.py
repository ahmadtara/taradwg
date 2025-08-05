import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import tempfile
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account
import os

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

def authenticate_google():
    creds_dict = st.secrets["gcp_service_account"]
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(credentials)
    return client

def upload_kml_to_drive(kmz_path, folder_name):
    try:
        creds_dict = st.secrets["gcp_service_account"]
        creds = service_account.Credentials.from_service_account_info(
            creds_dict, scopes=['https://www.googleapis.com/auth/drive']
        )
        drive_service = build('drive', 'v3', credentials=creds)

        with zipfile.ZipFile(kmz_path, 'r') as zf:
            kml_file = next((f for f in zf.namelist() if f.lower().endswith(".kml")), None)
            if not kml_file:
                st.error("❌ Tidak ditemukan file .kml dalam .kmz")
                return

            content = zf.read(kml_file)
            tree = ET.fromstring(content)
            ns = {"kml": "http://www.opengis.net/kml/2.2"}

            folders = tree.findall(".//kml:Folder", ns)
            for f in folders:
                name_el = f.find("kml:name", ns)
                if name_el is not None and name_el.text.upper() == folder_name:
                    folder_only = ET.Element(tree.tag, tree.attrib)
                    doc = ET.SubElement(folder_only, "kml:Document", nsmap=tree.nsmap if hasattr(tree, 'nsmap') else {})
                    doc.append(f)

                    tmp_kml = tempfile.NamedTemporaryFile(delete=False, suffix=".kml")
                    ET.ElementTree(folder_only).write(tmp_kml.name, encoding="utf-8", xml_declaration=True)

                    file_metadata = {
                        "name": f"{folder_name}.kml",
                        "parents": [GDRIVE_FOLDERS[folder_name]]
                    }
                    media = MediaFileUpload(tmp_kml.name, mimetype='application/vnd.google-earth.kml+xml')
                    drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                    os.unlink(tmp_kml.name)
                    st.success(f"✅ Folder '{folder_name}' berhasil dikirim ke Google Drive")
                    return
            st.warning(f"⚠️ Folder '{folder_name}' tidak ditemukan dalam KMZ")
    except Exception as e:
        st.error(f"❌ Gagal upload KML ke Google Drive: {e}")

# (kode lainnya tidak berubah)

# Tambahkan bagian setelah upload
        with st.spinner("☁️ Uploading KML folders ke Google Drive..."):
            upload_kml_to_drive(kmz_path, "DISTRIBUTION CABLE")
            upload_kml_to_drive(kmz_path, "BOUNDARY CLUSTER")

    if uploaded_subfeeder:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".kmz") as tmp:
            tmp.write(uploaded_subfeeder.read())
            kmz_path = tmp.name

        with st.spinner("🔍 Membaca data dari KMZ SUBFEEDER..."):
            _, poles_subonly, _ = extract_points_from_kmz(kmz_path, remarks_default="SUBFEEDER")

        try:
            client = authenticate_google()
            sheet1 = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
            if poles_subonly:
                append_poles_to_main_sheet(sheet1, poles_subonly, district_input, subdistrict_input, vendor_input)
        except Exception as e:
            st.error(f"❌ Gagal mengirim data SUBFEEDER ke spreadsheet utama: {e}")

        with st.spinner("☁️ Uploading CABLE folder ke Google Drive..."):
            upload_kml_to_drive(kmz_path, "CABLE")
