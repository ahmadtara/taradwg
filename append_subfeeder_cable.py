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
        row[11] = template_row[11]
        row[20:21] = template_row[20:21]
        row[22] = vendor
        row[24] = datetime.today().strftime("%d/%m/%Y")
        row[35] = template_row[35]
        row[36] = kmz_name  # opsional jika Anda ingin mencatat sumber file
        row[38] = vendor

        rows.append(row)

    # Tambahkan ke sheet
    if rows:
        sheet.append_rows(rows, value_input_option="USER_ENTERED")
    return len(rows)
