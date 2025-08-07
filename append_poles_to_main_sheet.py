from datetime import datetime

_cached_headers = None
_cached_prev_row = None

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

        # Copy beberapa kolom dari prev_row
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
