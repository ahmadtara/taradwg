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
import tempfile  # ✅ tambahkan ini
dist = __import__('math').dist

SPREADSHEET_ID_3 = "1EnteHGDnRhwthlCO9B12zvHUuv3wtq5L2AKlV11qAOU"
SHEET_NAME_3 = "FDT Pekanbaru"

SPREADSHEET_ID_4 = "1D_OMm46yr-e80s3sCyvbSSsf8wrUCwpwiYsVBKPgszw"
SHEET_NAME_4 = "Cable Pekanbaru"

SPREADSHEET_ID_5 = "1paa8sT3nTZh_xxwHeKV8pwVIWacq7lC8U9A8BlX6LUw"
SHEET_NAME_5 = "Sheet1"

SPREADSHEET_ID = "1yXBIuX2LjUWxbpnNqf6A9YimtG7d77V_AHLidhWKIS8"
SPREADSHEET_ID_2 = "1WI0Gb8ul5GPUND4ADvhFgH4GSlgwq1_4rRgfOnPz-yc"
SHEET_NAME = "Pole Pekanbaru"
SHEET_NAME_2 = "FAT Pekanbaru"
_cached_headers = None
_cached_prev_row = None

def authenticate_google():
    creds_dict = st.secrets["gcp_service_account"]
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(credentials)
    return client

def extract_kmz_data_combined(kmz_file):
    import xml.etree.ElementTree as ET
    import zipfile

    folders = {}
    poles = []
    seen_items = set()
    
    def recurse_folder(folder, ns, path=""):
        name_el = folder.find("kml:name", ns)
        folder_name = name_el.text.strip().upper() if name_el is not None else "UNKNOWN"
        current_path = f"{path}/{folder_name}" if path else folder_name

        if folder_name not in folders:
            folders[folder_name] = []

        for placemark in folder.findall("kml:Placemark", ns):
            name_tag = placemark.find("kml:name", ns)
            name = name_tag.text.strip() if name_tag is not None else ""

            coords_tag = placemark.find(".//kml:coordinates", ns)
            coords = coords_tag.text.strip().split(",") if coords_tag is not None and coords_tag.text else ["", ""]
            lon, lat = coords[:2] if len(coords) >= 2 else (None, None)

            description_tag = placemark.find("kml:description", ns)
            description = description_tag.text.strip() if description_tag is not None else ""

            unique_key = (name, lon, lat, folder_name)
            if unique_key in seen_items:
                continue
            seen_items.add(unique_key)

            item = {
                "name": name,
                "lon": float(lon) if lon else None,
                "lat": float(lat) if lat else None,
                "description": description,
                "folder": folder_name,
                "full_path": current_path
            }

            folders[folder_name].append(item)

            if folder_name == "NEW POLE 7-3":
                poles.append({**item, "folder": "7m3inch", "height": "7"})
            elif folder_name == "NEW POLE 7-4":
                poles.append({**item, "folder": "7m4inch", "height": "7"})
            elif folder_name == "NEW POLE 9-4":
                poles.append({**item, "folder": "9m4inch", "height": "9"})
            elif folder_name == "EXISTING POLE EMR 7-3":
                poles.append({**item, "folder": "ext7m3inch", "height": "7"})
            elif folder_name == "EXISTING POLE EMR 7-4":
                poles.append({**item, "folder": "ext7m4inch", "height": "7"})
            elif folder_name == "EXISTING POLE EMR 9-4":
                poles.append({**item, "folder": "ext9m4inch", "height": "9"})

        for subfolder in folder.findall("kml:Folder", ns):
            recurse_folder(subfolder, ns, current_path)

    with zipfile.ZipFile(kmz_file, 'r') as z:
        kml_filename = next((f for f in z.namelist() if f.lower().endswith('.kml')), None)
        if not kml_filename:
            raise ValueError("❌ Tidak ditemukan file .kml dalam .kmz")

        with z.open(kml_filename) as kml_file:
            tree = ET.parse(kml_file)
            root = tree.getroot()
            ns = {'kml': 'http://www.opengis.net/kml/2.2'}

            for folder in root.findall(".//kml:Folder", ns):
                recurse_folder(folder, ns)

    return folders, poles

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
        elif base_folder == "NEW POLE 7-3":
            poles.append({**p, "folder": "7m3inch", "height": "7", "remarks": "CLUSTER"})
            poles_subfeeder.append({**p, "folder": "7m3inch", "height": "7"})
        elif base_folder == "NEW POLE 7-4":
            poles.append({**p, "folder": "7m4inch", "height": "7"})
        elif base_folder == "NEW POLE 9-4":
            poles.append({**p, "folder": "9m4inch", "height": "9"})

    return fat_points, poles, poles_subfeeder

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

        for col_idx in range(5):
            row[col_idx] = prev_row[col_idx] if col_idx < len(prev_row) else ""

        row[5] = district.upper()
        row[6] = subdistrict.upper()
        row[7] = fat['name']
        row[8] = fat['name']
        row[9] = fat['lat']
        row[10] = fat['lon']

        for idx in range(11, 24):
            row[idx] = prev_row[idx] if idx < len(prev_row) else ""

        row[24] = vendor.upper()
        row[26] = formatted_date

        idx_ag = header_map.get('parentid 1')
        if idx_ag is not None:
            row[idx_ag] = find_nearest_pole(fat, [p for p in poles if p['folder'] == '7m3inch'])

        for col in ['parent_type 1', 'fat type']:
            idx = header_map.get(col.lower())
            if idx is not None:
                row[idx] = prev_row[idx]

        idx_al = header_map.get('vendor name')
        if idx_al is not None:
            row[idx_al] = vendor.upper()

        all_rows.append(row)

    sheet.append_rows(all_rows)
    st.success(f"✅ {len(fat_points)} FAT ")

def append_poles_to_main_sheet(sheet, poles, district, subdistrict, vendor):
    global _cached_headers, _cached_prev_row

    headers = _cached_headers or sheet.row_values(1)
    _cached_headers = headers
    header_map = {name.strip().lower(): i for i, name in enumerate(headers)}

    values = sheet.get_all_values()
    for i in range(len(values)-1, 0, -1):
        if any(values[i]):
            prev_row = values[i]
            break
    else:
        prev_row = [""] * len(headers)

    _cached_prev_row = prev_row

    today = datetime.today()
    formatted_date = today.strftime("%d/%m/%Y") if prev_row[header_map.get('installationdate', 0)].count("/") == 2 else today.strftime("%Y-%m-%d")

    count_types = {"7m3inch": 0, "7m4inch": 0, "9m4inch": 0}

    district = district.upper()
    subdistrict = subdistrict.upper()
    vendor = vendor.upper()

    all_rows = []
    for pole in poles:
        count_types[pole['folder']] += 1

        row = [""] * len(headers)
        row[0:4] = prev_row[0:4]
        row[4] = district
        row[5] = subdistrict
        row[6] = pole['name']
        row[7] = pole['name']
        row[8] = pole['lat']
        row[9] = pole['lon']

        for col in ['constructionstage', 'accessibility', 'activationstage', 'hierarchytype']:
            if col in header_map:
                row[header_map[col]] = prev_row[header_map[col]]

        for col in ['pole height', 'vendorname', 'installationyear', 'productionyear', 'installationdate', 'remarks']:
            idx = header_map.get(col.lower())
            if idx is not None:
                if col.lower() == 'pole height':
                    row[idx] = pole['height']
                elif col.lower() == 'vendorname':
                    row[idx] = vendor
                elif col.lower() in ['installationyear', 'productionyear']:
                    row[idx] = str(today.year)
                elif col.lower() == 'installationdate':
                    row[idx] = formatted_date
                elif col.lower() == 'remarks':
                    if pole['folder'] in ['7m4inch', '9m4inch']:
                        row[idx] = "SUBFEEDER"
                    else:
                        row[idx] = "CLUSTER"

        if 'poletype' in header_map:
            row[header_map['poletype']] = pole['folder']

        all_rows.append(row)

    sheet.append_rows(all_rows)

    st.info(f"""
📊 **Ringkasan Pengunggahan**:
✅ 7m3inch: {count_types['7m3inch']} titik
✅ 7m4inch: {count_types['7m4inch']} titik
✅ 9m4inch: {count_types['9m4inch']} titik
""")
    
def main():
    st.title("🚀 Webgis Teknologia - By. Tara")
    st.markdown("<h2>👋 Hai, <span style='color:#0A84FF'>bro assalamualaikum</span></h2>", unsafe_allow_html=True)
    st.markdown("""    ⚠️ <span style='font-weight:bold;'>CATATAN PENTING :</span><br> """, unsafe_allow_html=True)

    st.markdown("""    
    ✅ Deskripsi dari "FDT" wajib isi : contoh <span style='color:#FF6B6B;'> FDT 48 , FDT 72 , FDT 96  </span> <br> 
    ✅ Deskripsi dari "cable distribusi & subfeeder" wajib isi : contoh <span style='color:#FF6B6B;'>Total Route : xxxM. </span> <br>
    ✅ Wajib isi sesuai contoh diatas <br>
    ✅ Pastikan .KMZ dari Cluster & Subfeeder yang di upload udah sesuai sama template EMR <br>
    ✅ Nama KMZ Wajib Capital semua dan sesuai dengan nama RFS </span>.<br> 
    ✅ Wajib ikut keterangan karena program mengikuti template tersebut agar berhasil <br>""", unsafe_allow_html=True)
    
    
    col1, col2 = st.columns(2)
    with col1:
        kmz_fdt_file = st.file_uploader("📤 Upload file .kmz Cluster (FDT)", type="kmz", key="fdt")
    with col2:
        kmz_subfeeder_file = st.file_uploader("📤 Upload file .kmz Subfeeder", type="kmz", key="subfeeder")

    district = st.text_input("🗺️ District")
    subdistrict = st.text_input("🏙️ Subdistrict")
    vendor = st.text_input("🏗️ Vendor")

    submit = st.button("🚀 Submit & Kirim ke Spreadsheet")

    if submit:
        if not district or not subdistrict or not vendor:
            st.warning("⚠️ Harap isi semua kolom input manual.")
            return

        client = None
        count_fdt = 0
        count_cable = 0
        count_subfeeder = 0

        # === PROSES CLUSTER (FAT + POLE 7m3) ===
        if kmz_fdt_file:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".kmz") as tmp:
                tmp.write(kmz_fdt_file.read())
                kmz_path = tmp.name

            with st.spinner("🔍 Membaca data dari KMZ CLUSTER..."):
                fat_points, poles_cluster, poles_subfeeder = extract_points_from_kmz(kmz_path)

            try:
                client = authenticate_google()
                sheet1 = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
                if poles_cluster:
                    append_poles_to_main_sheet(sheet1, poles_cluster, district, subdistrict, vendor)
            except Exception as e:
                st.error(f"❌ Gagal mengirim ke spreadsheet utama: {e}")

            if fat_points:
                try:
                    sheet2 = client.open_by_key(SPREADSHEET_ID_2).worksheet(SHEET_NAME_2)
                    append_fat_to_sheet(sheet2, fat_points, poles_subfeeder, district, subdistrict, vendor)
                except Exception as e:
                    st.error(f"❌ Gagal mengirim ke spreadsheet kedua: {e}")
            else:
                st.warning("⚠️ Tidak ditemukan folder FAT dalam file KMZ.")

        # === PROSES SUBFEEDER ===
        if kmz_subfeeder_file:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".kmz") as tmp:
                tmp.write(kmz_subfeeder_file.read())
                kmz_path = tmp.name

            with st.spinner("🔍 Membaca data dari KMZ SUBFEEDER..."):
                _, poles_subonly, _ = extract_points_from_kmz(kmz_path)

            try:
                if client is None:
                    client = authenticate_google()
                sheet1 = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
                if poles_subonly:
                    append_poles_to_main_sheet(sheet1, poles_subonly, district, subdistrict, vendor)
            except Exception as e:
                st.error(f"❌ Gagal mengirim data SUBFEEDER ke spreadsheet utama: {e}")

        # === PROSES FDT & CABLE ===
        if (kmz_fdt_file or kmz_subfeeder_file):
            if kmz_fdt_file:
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
                        cable_names = [item['name'] for item in folders['DISTRIBUTION CABLE'] if item.get('name')]
                        st.write(cable_names)

            if kmz_subfeeder_file:
                with st.spinner("🔍 Memproses KMZ Subfeeder..."):
                    folders, poles = extract_kmz_data_combined(kmz_subfeeder_file)
                    kmz_name = kmz_subfeeder_file.name.replace(".kmz", "")
                    if client is None:
                        client = authenticate_google()

                    if 'CABLE' in folders:
                        sheet = client.open_by_key(SPREADSHEET_ID_5).worksheet(SHEET_NAME_5)
                        count_subfeeder = append_subfeeder_cable(sheet, folders['CABLE'], district, subdistrict, vendor, kmz_name)
                        cable_names2 = [item['name'] for item in folders['CABLE'] if item.get('name')]
                        st.write(cable_names2)

        # === RINGKASAN HASIL ===
        if (kmz_fdt_file or kmz_subfeeder_file):
            st.success("✅ Semua data berhasil diproses dan dikirim ke Spreadsheet!")
            if count_fdt:
                st.info(f"✅ {count_fdt} FDT ")
            if count_cable:
                st.info(f"✅ {count_cable} Kabel distribusi")
            if count_subfeeder:
                st.info(f"✅ {count_subfeeder} Kabel SubFeeder")
        else:
            st.warning("⚠️ Mohon upload minimal satu file KMZ CLUSTER atau SUBFEEDER.")

if __name__ == "__main__":
    main()

















