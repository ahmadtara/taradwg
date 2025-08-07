from datetime import datetime

def append_subfeeder_cable(sheet, cable_data, district, subdistrict, vendor, kmz_name):
    existing_rows = sheet.get_all_values()
    rows = []
    
    # Template dari baris terakhir
    template_row = existing_rows[-1] if len(existing_rows) > 1 else [""] * 30

    for cable in cable_data:
        name = cable.get("name", "")  # pastikan 'cable_data' adalah list of dict

        # Buat row kosong sesuai panjang baris template
        row = [""] * len(template_row)
        
        row[0] = name
        row[1] = name
        row[3] = district
        row[4] = subdistrict
        row[20] = vendor
        row[24] = datetime.today().strftime("%d/%m/%Y")
        row[29] = kmz_name  # opsional jika Anda ingin mencatat sumber file

        rows.append(row)

    # Tambahkan ke sheet
    if rows:
        sheet.append_rows(rows, value_input_option="USER_ENTERED")
    return len(rows)
