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

def extract_points_from_kmz(kmz_path, remarks_default="CLUSTER"):
    with zipfile.ZipFile(kmz_path, 'r') as zf:
        kml_file = next((f for f in zf.namelist() if f.lower().endswith(".kml")), None)
        if not kml_file:
            return [], [], []

        content = zf.read(kml_file)
        tree = ET.fromstring(content)
        ns = {"kml": "http://www.opengis.net/kml/2.2"}

        folders = tree.findall(".//kml:Folder", ns)
        all_points = []
        subfeeder_points = []
        fat_points = []

        for folder in folders:
            name_el = folder.find("kml:name", ns)
            folder_name = name_el.text.upper() if name_el is not None else ""

            placemarks = folder.findall(".//kml:Placemark", ns)
            for pm in placemarks:
                name = pm.find("kml:name", ns).text if pm.find("kml:name", ns) is not None else ""
                coord_el = pm.find(".//kml:coordinates", ns)
                if coord_el is None:
                    continue
                coords = coord_el.text.strip().split(",")
                if len(coords) < 2:
                    continue
                lon, lat = float(coords[0]), float(coords[1])
                row = {"Pole_Id": name, "Latitude": lat, "Longitude": lon, "Remarks": remarks_default}
                if folder_name.startswith("NEW POLE"):
                    all_points.append(row)
                    if "7-4" in folder_name or "9-4" in folder_name:
                        subfeeder_points.append(row)
                elif folder_name == "FAT":
                    fat_points.append(row)

        return all_points, subfeeder_points, fat_points

def append_poles_to_main_sheet(sheet, poles, district, subdistrict, vendor):
    global _cached_headers, _cached_prev_row

    if not poles:
        return

    if _cached_headers is None:
        _cached_headers = sheet.row_values(1)

    headers = _cached_headers

    if _cached_prev_row is None:
        all_rows = sheet.get_all_values()
        _cached_prev_row = all_rows[-1] if all_rows else []

    prev_row = _cached_prev_row
    today = datetime.now().strftime("%d/%m/%Y")

    rows = []
    for pole in poles:
        row = [""] * len(headers)
        row[headers.index("Pole_Id")] = pole["Pole_Id"]
        row[headers.index("PoleName")] = pole["Pole_Id"]
        row[headers.index("Latitude")] = pole["Latitude"]
        row[headers.index("Longitude")] = pole["Longitude"]
        row[headers.index("District")] = district
        row[headers.index("Subdistrict")] = subdistrict
        row[headers.index("VendorName")] = vendor
        row[headers.index("PoleType")] = "7m3inch" if "7-3" in pole["Pole_Id"] else "7m4inch"
        row[headers.index("Pole Height")] = "7" if "7" in pole["Pole_Id"] else "9"
        row[headers.index("InstallationDate")] = today
        row[headers.index("remark")] = pole["Remarks"]

        for col in ["Region", "SubRegion", "ProvinceName", "City", "ConstructionStage",
                    "accessibility", "ActivationStage", "HierarchyType", "InstallationYear",
                    "ProductionYear"]:
            if col in headers and headers.index(col) < len(prev_row):
                row[headers.index(col)] = prev_row[headers.index(col)]

        rows.append(row)

    sheet.append_rows(rows)
