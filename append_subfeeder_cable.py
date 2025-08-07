from datetime import datetime
from pyproj import Geod

# Hitung panjang lintasan berdasarkan koordinat
geod = Geod(ellps="WGS84")

def append_subfeeder_cable(sheet, cable_data, district, subdistrict, vendor, kmz_name):
    existing_rows = sheet.get_all_values()
    rows = []
    
    # Template dari baris terakhir
    template_row = existing_rows[-2] if len(existing_rows) > 2 else []
    for cable in cable_data:
    name = cable.get("name", "")
    coords = []
    # Ambil koordinat dari geometry LineString
    # Ambil koordinat dari geometry LineString
    if "geometry" in cable and hasattr(cable["geometry"], "coords"):
        coords = list(cable["geometry"].coords)

    length = 0
    if coords and len(coords) > 1:
        # Hitung total panjang lintasan
        for i in range(len(coords) - 1):
            lon1, lat1 = coords[i]
            lon2, lat2 = coords[i + 1]
            segment_length = geod.inverse(lon1, lat1, lon2, lat2)[2]  # hasil dalam meter
            length += segment_length
            
        # Buat row kosong sesuai panjang baris template
        row = [""] * len(template_row)
        
        row[0] = name
        row[1] = name
        row[2:6] = template_row[2:6]
        row[10:11] = template_row[10:11]
        row[20:21] = template_row[20:21]
        row[22] = vendor
        row[24] = datetime.today().strftime("%d/%m/%Y")
        row[35] = template_row[35]
        row[36] = kmz_name  # opsional jika Anda ingin mencatat sumber file
        row[38] = vendor

        # ===== Kolom J (index 9) =====
        if "FO 24/2T" in name:
            row[9] = "2"
        elif "FO 48/4T" in name:
            row[9] = "4"
        elif "FO 96/9T" in name:
            row[9] = "9"
        elif "FO 12/2T" in name:
            row[9] = "12"

        # ===== Kolom M (index 12) =====
        if "FO 24/2T" in name:
            row[12] = "24"
        elif "FO 48/4T" in name:
            row[12] = "48"
        elif "FO 96/9T" in name:
            row[12] = "96"

        # ===== Kolom Q (index 16) =====
        match = re.search(r"AE\s*[-]?\s*(\d+)\s*M", name.upper())
        if match:
            ae_number = match.group(1)
            row[16] = ae_number  # Kolom Q

        # ===== Kolom P (index 15) - panjang lintasan =====
        row[15] = str(round(length, 2))

        rows.append(row)

    # Tambahkan ke sheet
    if rows:
        sheet.append_rows(rows, value_input_option="USER_ENTERED")
    return len(rows)
