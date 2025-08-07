from datetime import datetime

def append_cable_pekanbaru(sheet, cable_data, district, subdistrict, vendor, kmz_name):
    existing_rows = sheet.get_all_values()
    rows = []
    
    # Ambil baris template dari baris terakhir atau buat default kosong
    template_row = existing_rows[-1] if len(existing_rows) > 1 else [""] * 30

    for cable in cable_data:
        name = cable.get("name", "")  # asumsi cable_data adalah list of dict

        row = [""] * len(template_row)
        row[0] = name
        row[1] = name
        row[3] = district
        row[4] = subdistrict
        row[20] = vendor
        row[24] = datetime.today().strftime("%d/%m/%Y")
        row[29] = kmz_name

        rows.append(row)

    if rows:
        sheet.append_rows(rows, value_input_option="USER_ENTERED")
    return len(rows)
