import re
from datetime import datetime
from pyproj import Transformer
from shapely.geometry import LineString

# Ubah koordinat WGS84 ke UTM Zone 60S (EPSG:32760)
transformer = Transformer.from_crs("epsg:4326", "epsg:32760", always_xy=True)


def append_subfeeder_cable(sheet, cable_data, district, subdistrict, vendor, kmz_name):
    existing_rows = sheet.get_all_values()
    rows = []

    # Gunakan baris sebelum terakhir sebagai template
    template_row = existing_rows[-2] if len(existing_rows) > 2 else []

    for cable in cable_data:
        name = cable.get("name", "")
        normalized_name = name.upper().replace(" ", "").replace("-", "")
        coords = []

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
        row[9] = cable.get("no_of_tube", "")
    
        # ===== Kolom M (index 12) =====
        row[12] = cable.get("total_core", "")
        

        # === Kolom Q (index 16) === Ambil angka setelah AE xxxx M
        match = re.search(r"AE\s*[-]?\s*(\d+)\s*M", name.upper())
        if match:
            row[16] = match.group(1)

        # === Kolom P (index 15): Hitung panjang lintasan (meter) ===
        coords = cable.get("coordinates", [])
        if coords and len(coords) >= 2:
            utm_coords = []
            for lon, lat in coords:
                x, y = transformer.transform(lon, lat)
                utm_coords.append((x, y))
            line = LineString(utm_coords)
            row[15] = round(line.length, 2)
        else:
            row[15] = "")

        rows.append(row)

    # Tambah ke sheet jika ada baris baru
    if rows:
        sheet.append_rows(rows, value_input_option="USER_ENTERED")

    return len(rows)
