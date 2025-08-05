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

SPREADSHEET_ID = "1yXBIuX2LjUWxbpnNqf6A9YimtG7d77V_AHLidhWKIS8"
SPREADSHEET_ID_2 = "1WI0Gb8ul5GPUND4ADvhFgH4GSlgwq1_4rRgfOnPz-yc"
SHEET_NAME = "Pole Pekanbaru"

SPREADSHEET_ID_3 = "1EnteHGDnRhwthlCO9B12zvHUuv3wtq5L2AKlV11qAOU"
SPREADSHEET_ID_4 = "1D_OMm46yr-e80s3sCyvbSSsf8wrUCwpwiYsVBKPgszw"
SPREADSHEET_ID_5 = "1paa8sT3nTZh_xxwHeKV8pwVIWacq7lC8U9A8BlX6LUw"

SHEET_NAME_3 = "FDT Pekanbaru"
SHEET_NAME_4 = "Cable Pekanbaru"
SHEET_NAME_5 = "Sheet1"

GDRIVE_FOLDERS = {
    "DISTRIBUTION CABLE": "1XkWqvRX4SUYMrtMQ7vt8197oSja4r9p-",
    "BOUNDARY CLUSTER": "1IMpaQWnpG8c8P5j3phUMP1G9zTPBDQMi",
    "CABLE": "16aesqK-OIqYIDAIn_ymLzf1-VkLyXonl"
}

_cached_headers = None
_cached_prev_row = None

def get_client():
    creds_dict = st.secrets["gcp_service_account"]
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(credentials)
    return client

def get_prev_row(sheet):
    values = sheet.get_all_values()
    for i in range(len(values)-1, 0, -1):
        if any(values[i]):
            return values[i]
    return sheet.row_values(1)

def extract_number(text, pattern):
    match = re.search(pattern, text)
    if match:
        return match.group(1)
    return ""

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

from .spreadsheet_main import append_poles_to_main_sheet, append_fat_to_sheet
from .spreadsheet_extra import append_to_fdt_sheet, append_to_cable_cluster_sheet, append_to_subfeeder_cable_sheet
