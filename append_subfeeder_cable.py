from datetime import datetime

def append_subfeeder_cable(sheet, cable_data, district, subdistrict, vendor, kmz_name):
    existing_rows = sheet.get_all_values()
    rows = []
    
    # Template dari baris terakhir
    template_row = existing_rows[-2] if len(existing_rows) > 2 else []
    for cable in cable_data:
        name = cable.get("name", "")  # pastikan 'cable_data' adalah list of dict

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
        match = re.search(r"AE[\s\-]*([0-9]+)[\s]*M", name.upper())
        if match:
            row[16] = match.group(1)

        # ===== Kolom P (index 15) - panjang lintasan =====
        if "path" in cable:
            path = cable["path"]
            if isinstance(path, list) and len(path) >= 2:
                total_distance = 0
                for i in range(1, len(path)):
                    lat1, lon1 = path[i-1]
                    lat2, lon2 = path[i]
                    # Rumus haversine sederhana (flat approximation, cukup presisi untuk city scale)
                    from math import radians, sin, cos, sqrt, atan2
                    R = 6371000  # Earth radius in meter
                    dlat = radians(lat2 - lat1)
                    dlon = radians(lon2 - lon1)
                    a = sin(dlat/2)**2 + cos(radians(lat1)) * cos(radians(lat2)) * sin(dlon/2)**2
                    c = 2 * atan2(sqrt(a), sqrt(1-a))
                    distance = R * c
                    total_distance += distance
                row[15] = round(total_distance)

        rows.append(row)

    # Tambahkan ke sheet
    if rows:
        sheet.append_rows(rows, value_input_option="USER_ENTERED")
    return len(rows)
