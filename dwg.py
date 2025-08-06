import zipfile
import xml.etree.ElementTree as ET
import streamlit as st
import math
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Konstanta kolom untuk Spreadsheet FDT
KOLOM_TEMPLATECODE = 0
KOLOM_NAMA = 1
KOLOM_LAT = 2
KOLOM_LON = 3
KOLOM_KETINGGIAN = 4
KOLOM_PARENT_ID_1 = 39
KOLOM_DUPLIKAT_AD = 29
KOLOM_DUPLIKAT_AE = 30
KOLOM_DUPLIKAT_AO = 40
KOLOM_AO_ASLI = 41
KOLOM_AF_HASIL = 31

# Spreadsheet ID
SPREADSHEET_ID_3 = "1EnteHGDnRhwthlCO9B12zvHUuv3wtq5L2AKlV11qAOU"
SPREADSHEET_ID_4 = "1lVjvOP5b3nzKw_qWZpRP2F9EupzzvFGzeDW1kLdUV60"

# Folder
FOLDER_FDT = "FDT"

# Autentikasi Google Sheets
@st.cache_resource
def get_gsheet_client():
    credentials = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return build("sheets", "v4", credentials=credentials).spreadsheets()

def haversine(lat1, lon1, lat2, lon2):
    R = 6371.0
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi/2)**2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda/2)**2
    return R * (2 * math.atan2(math.sqrt(a), math.sqrt(1 - a)))

def extract_points_from_kmz(kmz_file):
    folders = {}
    poles = []
    poles_subfeeder = []
    with zipfile.ZipFile(kmz_file, 'r') as kmz:
        with kmz.open('doc.kml', 'r') as kmlfile:
            tree = ET.parse(kmlfile)
            root = tree.getroot()
            ns = {'kml': 'http://www.opengis.net/kml/2.2'}
            for folder in root.findall(".//kml:Folder", ns):
                name_elem = folder.find("kml:name", ns)
                if name_elem is None:
                    continue
                folder_name = name_elem.text.strip()
                placemarks = []
                for pm in folder.findall(".//kml:Placemark", ns):
                    coords = pm.find(".//kml:coordinates", ns)
                    name = pm.find("kml:name", ns)
                    desc = pm.find("kml:description", ns)
                    if coords is not None and name is not None:
                        lon, lat, *_ = map(float, coords.text.strip().split(","))
                        placemark = {
                            "name": name.text.strip(),
                            "lat": lat,
                            "lon": lon,
                            "description": desc.text.strip() if desc is not None else ""
                        }
                        placemarks.append(placemark)

                        base_folder = folder_name.upper()
                        if base_folder == "SUBFEEDER CABLE":
                            poles_subfeeder.append(placemark)
                        elif base_folder == "NEW POLE 7-3":
                            poles.append({**placemark, "folder": "7m3inch", "height": "7"})
                        elif base_folder == "NEW POLE 7-4":
                            poles.append({**placemark, "folder": "7m4inch", "height": "7"})
                        elif base_folder == "NEW POLE 9-4":
                            poles.append({**placemark, "folder": "9m4inch", "height": "9"})
                        elif base_folder == "EXISTING POLE EMR 7-4":
                            poles.append({**placemark, "folder": "ext7m4inch", "height": "7"})
                        elif base_folder == "EXISTING POLE EMR 7-3":
                            poles.append({**placemark, "folder": "ext7m3inch", "height": "7"})
                        elif base_folder == "EXISTING POLE EMR 9-4":
                            poles.append({**placemark, "folder": "ext9m4inch", "height": "9"})

                folders[folder_name] = placemarks
    return folders, poles, poles_subfeeder

def find_nearest_pole(point, poles):
    min_dist = float('inf')
    nearest_pole = None
    for pole in poles:
        dist = haversine(point['lat'], point['lon'], pole['lat'], pole['lon'])
        if dist < min_dist:
            min_dist = dist
            nearest_pole = pole
    return nearest_pole['name'] if nearest_pole else ""

def append_fdt_to_sheet(sheet, fdt_points, poles):
    rows = []
    for fdt in fdt_points:
        row = ["" for _ in range(50)]
        row[KOLOM_TEMPLATECODE] = fdt['description']
        row[KOLOM_NAMA] = fdt['name']
        row[KOLOM_LAT] = fdt['lat']
        row[KOLOM_LON] = fdt['lon']
        row[KOLOM_KETINGGIAN] = 0
        row[KOLOM_PARENT_ID_1] = find_nearest_pole(fdt, poles)
        row[KOLOM_AF_HASIL] = fdt['description']
        row[KOLOM_DUPLIKAT_AD] = "=INDIRECT(""AD"" & ROW()-1)"
        row[KOLOM_DUPLIKAT_AE] = "=INDIRECT(""AE"" & ROW()-1)"
        row[KOLOM_DUPLIKAT_AO] = "=INDIRECT(""AO"" & ROW()-1)"
        rows.append(row)
    sheet.values().append(
        spreadsheetId=SPREADSHEET_ID_3,
        range="FDT Pekanbaru!A2",
        valueInputOption="USER_ENTERED",
        body={"values": rows}
    ).execute()
    st.success(f"{len(rows)} FDT berhasil dikirim ke Spreadsheet ke-3")

def append_cable_to_sheet(sheet, poles_subfeeder):
    rows = [[p['name'], p['lat'], p['lon']] for p in poles_subfeeder]
    sheet.values().append(
        spreadsheetId=SPREADSHEET_ID_4,
        range="Cable Pekanbaru!A2",
        valueInputOption="USER_ENTERED",
        body={"values": rows}
    ).execute()
    st.success(f"{len(rows)} kabel berhasil dikirim ke Spreadsheet ke-4")

def main():
    st.title("KMZ to Google Sheets - Auto Mapper")

    client = get_gsheet_client()

    kmz_fdt_file = st.file_uploader("Upload KMZ untuk FDT", type="kmz")
    if kmz_fdt_file:
        with st.spinner("Memproses KMZ FDT..."):
            folders, poles, poles_subfeeder = extract_points_from_kmz(kmz_fdt_file)

            if FOLDER_FDT not in folders:
                st.warning("Folder FDT tidak ditemukan di file KMZ!")
                return

            fdt_points = folders[FOLDER_FDT]
            append_fdt_to_sheet(client, fdt_points, poles)

    kmz_cable_file = st.file_uploader("Upload KMZ untuk Distribution/Subfeeder Cable", type="kmz")
    if kmz_cable_file:
        with st.spinner("Memproses KMZ Kabel..."):
            folders, poles, poles_subfeeder = extract_points_from_kmz(kmz_cable_file)
            append_cable_to_sheet(client, poles_subfeeder)

if __name__ == "__main__":
    main()
