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

def extract_data_from_kmz(kmz_file):
    def parse_folder(folder_element, ns, folder_path=""):
        folder_name_el = folder_element.find("kml:name", ns)
        folder_name = folder_name_el.text.strip() if folder_name_el is not None else "Unknown"
        full_folder_path = f"{folder_path}/{folder_name}" if folder_path else folder_name

        placemarks = []
        for placemark in folder_element.findall("kml:Placemark", ns):
            name = placemark.find("kml:name", ns)
            name = name.text.strip() if name is not None else ""

            coords_tag = placemark.find(".//kml:coordinates", ns)
            coords = coords_tag.text.strip() if coords_tag is not None else ""
            coords = coords.split(",") if coords else ["", "", ""]

            description_tag = placemark.find("kml:description", ns)
            description = description_tag.text.strip() if description_tag is not None else ""

            placemarks.append({
                'name': name,
                'lon': coords[0],
                'lat': coords[1],
                'alt': coords[2] if len(coords) > 2 else "",
                'description': description,
                'folder': full_folder_path.split("/")[-1],
                'full_path': full_folder_path
            })

        for subfolder in folder_element.findall("kml:Folder", ns):
            placemarks.extend(parse_folder(subfolder, ns, full_folder_path))

        return placemarks

    folders = {}
    new_pole_found = False
    with zipfile.ZipFile(kmz_file, 'r') as z:
        kml_filename = [f for f in z.namelist() if f.endswith('.kml')][0]
        with z.open(kml_filename) as kml_file:
            tree = ET.parse(kml_file)
            root = tree.getroot()

            ns = {'kml': 'http://www.opengis.net/kml/2.2'}
            for folder in root.findall(".//kml:Folder", ns):
                folder_name_el = folder.find("kml:name", ns)
                folder_name = folder_name_el.text.strip() if folder_name_el is not None else "Unknown"
                placemarks = parse_folder(folder, ns)
                if folder_name not in folders:
                    folders[folder_name] = []
                folders[folder_name].extend(placemarks)
                if folder_name.upper().startswith("NEW POLE") and placemarks:
                    new_pole_found = True

    if not new_pole_found:
        st.warning("⚠️ Tidak ditemukan titik tiang di folder yang diawali dengan 'NEW POLE' dalam KMZ.")
    else:
        st.success("✅ Titik tiang dari folder 'NEW POLE' berhasil ditemukan dalam KMZ.")

    return folders
    
def extract_points_from_kmz(kmz_path):
    fat_points, poles, poles_subfeeder = [], [], []

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
            return [], [], []

        root = ET.parse(zf.open(kml_file)).getroot()
        ns = {"kml": "http://www.opengis.net/kml/2.2"}
        all_pm = []
        for folder in root.findall(".//kml:Folder", ns):
            all_pm += recurse_folder(folder, ns)

    for p in all_pm:
        base_folder = p["path"].split("/")[0].upper()
        if base_folder == "FAT":
            fat_points.append(p)
        elif base_folder == "NEW POLE 7-4":
            poles.append({**p, "folder": "7m4inch", "height": "7"})
        elif base_folder == "NEW POLE 9-4":
            poles.append({**p, "folder": "9m4inch", "height": "9"})

    return fat_points, poles, poles_subfeeder

def find_nearest_pole(fdt_point, poles):
    min_dist = float('inf')
    nearest_name = ""
    for pole in poles:
        d = dist([fdt_point['lat'], fdt_point['lon']], [pole['lat'], pole['lon']])
        if d < min_dist:
            min_dist = d
            nearest_name = pole['name']
    return nearest_name

def templatecode_to_kolom_m(templatecode):
    mapping = {
        "FDT 48": "2",
        "FDT 72": "3",
        "FDT 96": "4"
    }
    return mapping.get(templatecode.strip().upper(), "")

def templatecode_to_kolom_r(templatecode):
    mapping = {
        "FDT 48": "4",
        "FDT 72": "6",
        "FDT 96": "8"
    }
    return mapping.get(templatecode.strip().upper(), "")

def templatecode_to_kolom_ap(templatecode):
    mapping = {
        "FDT 48": "FDT TYPE 48 CORE",
        "FDT 72": "FDT TYPE 72 CORE",
        "FDT 96": "FDT TYPE 96 CORE"
    }
    return mapping.get(templatecode.strip().upper(), "")

def is_float(value):
    try:
        float(value)
        return True
    except (TypeError, ValueError):
        return False


def append_fdt_to_sheet(sheet, fdt_data, poles, district, subdistrict, vendor, kmz_name):
    existing_rows = sheet.get_all_values()
    header_map = {name.strip().lower(): i for i, name in enumerate(headers)}
    template_row = existing_rows[-1] if len(existing_rows) > 1 else []
    rows = []

    idx_parentid = header_map.get('parentid 1')
    if idx_parentid is None:
        st.error("Kolom 'Parentid 1' tidak ditemukan di header spreadsheet.")
        return 0

    for fdt in fdt_data:
        name = fdt['name']
        lat = fdt['lat']
        lon = fdt['lon']
        desc = fdt.get('description', '')

        kolom_m = templatecode_to_kolom_m(template_row[0])
        kolom_r = templatecode_to_kolom_r(template_row[0])
        kolom_ap = templatecode_to_kolom_ap(template_row[0])

        row = [""] * len(existing_rows[0])
        row[0] = desc
        row[1:5] = template_row[1:5]
        row[5] = district
        row[6] = subdistrict
        row[7] = kmz_name
        row[8] = name
        row[9] = name
        row[10] = lat
        row[11] = lon
        row[12] = kolom_m
        row[13:16] = template_row[13:16]
        row[17] = kolom_r
        row[18] = template_row[18]
        row[24:26] = template_row[24:26]
        row[26] = template_row[26]
        row[29] = template_row[29]
        row[30] = template_row[30]
        row[40] = template_row[40]
        row[41] = kolom_ap
        row[33] = datetime.today().strftime("%d/%m/%Y")
        row[31] = vendor
        row[44] = vendor
        idx_an = header_map.get('parentid 1')
        if idx_an is not None:
            row[idx_an] = find_nearest_pole(fdt, [p for p in poles if p['folder'] == '7m4inch'])
        
        rows.append(row)

    sheet.append_rows(rows, value_input_option="USER_ENTERED")
    return len(rows)

def append_cable_pekanbaru(sheet, cable_data, district, subdistrict, vendor, kmz_name):
    rows = []
    for cable in cable_data:
        row = [
            cable['name'],
            cable['lat'],
            cable['lon'],
            district,
            subdistrict,
            vendor,
            kmz_name,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ]
        rows.append(row)

    sheet.append_rows(rows, value_input_option="USER_ENTERED")
    return len(rows)

def append_subfeeder_cable(sheet, cable_data, district, subdistrict, vendor, kmz_name):
    rows = []
    for cable in cable_data:
        row = [
            cable['name'],
            cable['lat'],
            cable['lon'],
            district,
            subdistrict,
            vendor,
            kmz_name,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ]
        rows.append(row)

    sheet.append_rows(rows, value_input_option="USER_ENTERED")
    return len(rows)

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
            folders = extract_data_from_kmz(kmz_fdt_file)
            kmz_name = kmz_fdt_file.name.replace(".kmz", "")
            if client is None:
                client = authenticate_google()

            poles_7_4 = folders.get("NEW POLE 7-4", [])
            if poles_7_4:
                st.success(f"✅ {len(poles_7_4)} titik tiang dari 'NEW POLE 7-4' berhasil diambil.")
            else:
                st.warning("⚠️ Tidak ditemukan titik tiang di folder 'NEW POLE 7-4' dalam KMZ.")

            if 'FDT' in folders:
                sheet = client.open_by_key(SPREADSHEET_ID_3).worksheet(SHEET_NAME_3)
                count_fdt = append_fdt_to_sheet(sheet, folders['FDT'], poles_7_4, district, subdistrict, vendor, kmz_name)

            if 'DISTRIBUTION CABLE' in folders:
                sheet = client.open_by_key(SPREADSHEET_ID_4).worksheet(SHEET_NAME_4)
                count_cable = append_cable_pekanbaru(sheet, folders['DISTRIBUTION CABLE'], district, subdistrict, vendor, kmz_name)

    if kmz_subfeeder_file and district and subdistrict and vendor:
        with st.spinner("🔍 Memproses KMZ Subfeeder..."):
            folders = extract_data_from_kmz(kmz_subfeeder_file)
            kmz_name = kmz_subfeeder_file.name.replace(".kmz", "")
            if client is None:
                client = authenticate_google()

            if 'SUBFEEDER CABLE' in folders:
                sheet = client.open_by_key(SPREADSHEET_ID_5).worksheet(SHEET_NAME_5)
                count_subfeeder = append_subfeeder_cable(sheet, folders['SUBFEEDER CABLE'], district, subdistrict, vendor, kmz_name)

    if (kmz_fdt_file or kmz_subfeeder_file) and district and subdistrict and vendor:
        st.success("✅ Semua data berhasil diproses dan dikirim ke Spreadsheet!")
        st.info(f"🛰️ {count_fdt} FDT dikirim ke spreadsheet FDT Pekanbaru")
        st.info(f"📦 {count_cable} kabel distribusi dikirim ke Cable Pekanbaru")
        st.info(f"🔌 {count_subfeeder} kabel subfeeder dikirim ke Sheet1")

if __name__ == "__main__":
    main()



