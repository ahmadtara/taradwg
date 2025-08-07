from datetime import datetime

def append_subfeeder_cable(sheet, cable_data, district, subdistrict, vendor, kmz_name):
    existing_rows = sheet.get_all_values()
    rows = []
    template_row = existing_rows[-1] if len(existing_rows) > 1 else []
        row = [""] * len(existing_rows[0])
        row[0] = name
        row[1] = name
        row[3:5] = template_row[3:5]
        row[20:21] = template_row[20:21]
        row[24] = datetime.today().strftime("%d/%m/%Y")
        
        rows.append(row)

    sheet.append_rows(rows, value_input_option="USER_ENTERED")
    return len(rows)
