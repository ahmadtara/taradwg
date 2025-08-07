import re
from datetime import datetime


def append_cable_pekanbaru(sheet, cable_data, district, subdistrict, vendor, kmz_name):
    existing_rows = sheet.get_all_values()
    rows = []

    # Gunakan baris sebelum terakhir sebagai template
    template_row = existing_rows[-2] if len(existing_rows) > 2 else []

    for cable in cable_data:
        name = cable.get("name", "")
        description = cable.get("description", "") 
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
        row[24] = datetime.today().strftime("%d/%m/%Y")  # Kolom Y
        row[35] = template_row[35]  # Kolom AJ
        row[36] = kmz_name          # Kolom AK
        row[38] = vendor            # Kolom AM
        
        match_fo = re.search(r"\(FO\s*(\d+)C/(\d+)T\)", row[0].upper())
        if match_fo:
            row[9] = match_fo.group(2)   # Kolom J
            row[12] = match_fo.group(1)  # Kolom M
        
        # === Kolom Q (index 16) === Ambil angka setelah AE xxxx M
        match = re.search(r"AE\s*[-]?\s*(\d+)\s*M", name.upper())
        if match:
            row[16] = match.group(1)

        length_from_desc = ""
        desc_match = re.search(r"Total\s+Route\s*:\s*(\d+)\s*M", description, re.IGNORECASE)
        if desc_match:
            length_from_desc = desc_match.group(1)
            row[15] = length_from_desc  # gunakan hasil dari deskripsi

        rows.append(row)

    if rows:
        sheet.append_rows(rows, value_input_option="USER_ENTERED")
    return len(rows)
