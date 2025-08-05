import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import tempfile
from datetime import datetime
dist = __import__('math').dist

SPREADSHEET_ID = "1yXBIuX2LjUWxbpnNqf6A9YimtG7d77V_AHLidhWKIS8"
SPREADSHEET_ID_2 = "1WI0Gb8ul5GPUND4ADvhFgH4GSlgwq1_4rRgfOnPz-yc"
SHEET_NAME = "Pole Pekanbaru"

_cached_headers = None
_cached_prev_row = None

def authenticate_google():
    creds_dict = st.secrets["gcp_service_account"]
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(credentials)
    return client

def extract_points_from_kmz(kmz_path):
    fat_points, poles = [], []

    def recurse_folder(folder, ns, path=""):
        items = []
        name_el = folder.find("kml:name", ns)
        folder_name = name_el.text.upper() if name_el is not None else "UNKNOWN"
        new_path = f"{path}/{folder_name}" if path else folder_name
        for sub in folder.findall("kml:Folder", ns):
            items += recurse_folder(sub, ns, new_path)
        for pm in folder.findall("kml:Placemark", ns):
            nm = pm.find("kml:name", ns)
            coord = pm.find(".//kml:coordinates", ns)
            if nm is not None and coord is not None and ',' in coord.text:
                lon, lat = coord.text.strip().split(",")[:2]
                items.append({"name": nm.text.strip(), "lat": float(lat), "lon": float(lon), "path": new_path})
        return items

    with zipfile.ZipFile(kmz_path, 'r') as zf:
        kml_file = next((f for f in zf.namelist() if f.lower().endswith(".kml")), None)
        if not kml_file:
            st.error("❌ Tidak ditemukan file .kml dalam .kmz")
            return [], []

        root = ET.parse(zf.open(kml_file)).getroot()
        ns = {"kml": "http://www.opengis.net/kml/2.2"}
        all_pm = []
        for folder in root.findall(".//kml:Folder", ns):
            all_pm += recurse_folder(folder, ns)

    for p in all_pm:
        base_folder = p["path"].split("/")[0].upper()
        if base_folder == "FAT":
            fat_points.append(p)
        elif base_folder == "NEW POLE 7-3":
            poles.append(p)

    return fat_points, poles

def find_nearest_pole(fat_point, poles):
    min_dist = float('inf')
    nearest_name = ""
    for pole in poles:
        d = dist([fat_point['lat'], fat_point['lon']], [pole['lat'], pole['lon']])
        if d < min_dist:
            min_dist = d
            nearest_name = pole['name']
    return nearest_name

def append_fat_to_sheet(sheet, fat_points, poles, district, subdistrict, vendor):
    headers = sheet.row_values(1)
    header_map = {name.strip().lower(): i for i, name in enumerate(headers)}

    values = sheet.get_all_values()
    for i in range(len(values)-1, 0, -1):
        if any(values[i]):
            prev_row = values[i]
            break
    else:
        prev_row = [""] * len(headers)

    today = datetime.today()
    formatted_date = today.strftime("%d/%m/%Y") if prev_row[header_map.get('installation_date', 0)].count("/") == 2 else today.strftime("%Y-%m-%d")

    all_rows = []
    for fat in fat_points:
        row = [""] * len(headers)

        for col_idx in range(5):  # kolom A-E
            row[col_idx] = prev_row[col_idx] if col_idx < len(prev_row) else ""

        row[5] = district.upper()       # Kolom F
        row[6] = subdistrict.upper()   # Kolom G
        row[7] = fat['name']           # Kolom H
        row[8] = fat['name']           # Kolom I
        row[9] = fat['lat']            # Kolom J
        row[10] = fat['lon']           # Kolom K

        for idx in range(11, 24):      # Kolom L-X
            row[idx] = prev_row[idx] if idx < len(prev_row) else ""

        row[24] = vendor.upper()       # Kolom Y
        row[26] = formatted_date       # Kolom AA

        # Kolom AG
        idx_ag = header_map.get('parentid 1')
        if idx_ag:
            row[idx_ag] = find_nearest_pole(fat, poles)

        # Kolom AH-AI
        for col in ['parent_type 1', 'fat type']:
            idx = header_map.get(col.lower())
            if idx:
                row[idx] = prev_row[idx]

        # Kolom AL
        idx_al = header_map.get('vendor name')
        if idx_al:
            row[idx_al] = vendor.upper()

        all_rows.append(row)

    sheet.append_rows(all_rows)
    st.success(f"✅ {len(fat_points)} FAT berhasil dikirim ke Spreadsheet ke-2 🛰️")

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

uploaded_cluster = st.file_uploader("📤 Upload file .KMZ CLUSTER (berisi FAT & NEW POLE 7-3)", type=["kmz"])

submit_clicked = st.button("🚀 Submit dan Kirim ke Google Sheet")

if submit_clicked:
    if not district_input or not subdistrict_input or not vendor_input:
        st.warning("⚠️ Harap isi semua kolom input manual.")
    elif not uploaded_cluster:
        st.warning("⚠️ Harap upload file KMZ.")
    else:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".kmz") as tmp:
            tmp.write(uploaded_cluster.read())
            kmz_path = tmp.name

        with st.spinner("🔍 Membaca data FAT dan Pole dari KMZ..."):
            fat_points, poles = extract_points_from_kmz(kmz_path)

        if fat_points:
            try:
                client = authenticate_google()
                sheet2 = client.open_by_key(SPREADSHEET_ID_2).worksheet(SHEET_NAME)
                append_fat_to_sheet(sheet2, fat_points, poles, district_input, subdistrict_input, vendor_input)
            except Exception as e:
                st.error(f"❌ Gagal mengirim ke spreadsheet kedua: {e}")
        else:
            st.warning("⚠️ Tidak ditemukan folder FAT dalam file KMZ.")
