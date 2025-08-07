import re
from datetime import datetime
from pyproj import Geod

# Inisialisasi Geod WGS84
geod = Geod(ellps="WGS84")

def append_subfeeder_cable(sheet, cable_data, district, subdistrict, vendor, kmz_name):
    existing_rows = sheet.get_all_values()
    rows = []

    # Gunakan baris sebelum terakhir sebagai template
    template_row = existing_rows[-2] if len(existing_rows) > 2 else []

    for cable in cable_data:
        name = cable.get("name", "")
        normalized_name = name.upper().replace(" ", "").replace("-", "")
        coords = []

        # Ambil koordinat dari LineString geometry
        if "geometry" in cable and hasattr(cable["geometry"], "coords"):
            coords = list(cable["geometry"].coords)

        # Hitung panjang lintasan
        length = 0
        if coords and len(coords) > 1:
            for i in range(len(coords) - 1):
                lon1, lat1 = coords[i]
                lon2, lat2 = coords[i + 1]
                segment_length = geod.inverse(lon1, lat1, lon2, lat2)[2]
                length += segment_length

        # Siapkan baris baru dengan panjang template
        row = [""] * len(template_row)
        row[0] = name
        row[1] = name
        row[2:6] = template_row[2:6]
        row[10] = template_row[10]
        row[11] = template_row[11]
        row[20] = template_row[20]  # Kolom U (index 20)
        row[21] = template_row[21]  # Kolom V (index 21)
        row[22] = vendor            # Kolom W
        row[24] = datetime.today().strftime("%d/%m/%Y")  # Kolom Y
        row[35] = template_row[35]  # Kolom AJ
        row[36] = kmz_name          # Kolom AK
        row[38] = vendor            # Kolom AM

        # ===== Kolom J (index 9) =====
        if "FO24/2T" in normalized_name:
            row[9] = "2"
        elif "FO48/4T" in normalized_name:
            row[9] = "4"
        elif "FO96/9T" in normalized_name:
            row[9] = "9"
        elif "FO12/2T" in normalized_name:
            row[9] = "12"

        # ===== Kolom M (index 12) =====
        if "FO24/2T" in normalized_name:
            row[12] = "24"
        elif "FO48/4T" in normalized_name:
            row[12] = "48"
        elif "FO96/9T" in normalized_name:
            row[12] = "96"
        elif "FO12/2T" in normalized_name:
            row[12] = "12"

        # === Kolom Q (index 16) === Ambil angka setelah AE xxxx M
        match = re.search(r"AE\s*[-]?\s*(\d+)\s*M", name.upper())
        if match:
            row[16] = match.group(1)

        # === Kolom P (index 15) === Isi panjang lintasan (dalam meter)
        row[15] = str(round(length, 2))

        rows.append(row)

    # Tambah ke sheet jika ada baris baru
    if rows:
        sheet.append_rows(rows, value_input_option="USER_ENTERED")

    return len(rows)
