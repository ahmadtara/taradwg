def fill_subfeeder_pekanbaru(sheet, sub_paths, vendor, filename):
    from datetime import datetime
    import re

    headers = sheet.row_values(1)
    header_map = {name.strip().lower(): idx for idx, name in enumerate(headers)}

    values = sheet.get_all_values()
    for i in range(len(values)-1, 0, -1):
        if any(values[i]):
            prev_row = values[i]
            break
    else:
        prev_row = [""] * len(headers)

    today = datetime.today()
    formatted_date = parse_date_format(prev_row[header_map.get('y', 0)], today)

    rows = []
    for path in sub_paths:
        row = [""] * len(headers)
        parts = path["name"].split("-")
        row[0] = parts[0] if len(parts) > 0 else path["name"]
        row[1] = parts[1] if len(parts) > 1 else path["name"]

        # FO pattern: FO 24C/2T atau FO 48/4T
        fo_match = re.search(r"FO\s*(\d+)[A-Z]*\/(\d+)T", path["name"], re.IGNORECASE)
        if fo_match:
            m_val = fo_match.group(1)
            j_val = fo_match.group(2)
            row[header_map.get('j')] = j_val
            row[header_map.get('m')] = m_val

        # AE pattern: AE-775M atau AE 775 M
        ae_match = re.search(r"AE[-\s]?(\d+)\s*M", path["name"], re.IGNORECASE)
        if ae_match:
            row[header_map.get('q')] = ae_match.group(1)

        # Panjang path
        if 'p' in header_map:
            row[header_map['p']] = path["length_m"]

        # Kolom vendor dan filename
        row[header_map.get('am')] = vendor
        row[header_map.get('ak')] = filename

        # Tanggal
        row[header_map.get('y')] = formatted_date

        # Duplikat kolom tertentu dari baris atas
        for col in ['c', 'd', 'e', 'f', 'k', 'l', 'u', 'v', 'aj']:
            if col in header_map:
                row[header_map[col]] = prev_row[header_map[col]]

        rows.append(row)

    sheet.append_rows(rows)
    return len(rows)
