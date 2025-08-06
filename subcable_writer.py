def fill_subfeeder_pekanbaru(sheet, sub_paths, vendor, filename):
    from datetime import datetime
    import re
    from sheet_utils import parse_date_format

    # Ambil header dan buat mapping nama kolom ke indeks
    headers = sheet.row_values(1)
    header_map = {name.strip().lower(): idx for idx, name in enumerate(headers)}

    # Ambil data sebelumnya untuk diduplikasi
    values = sheet.get_all_values()
    for i in range(len(values) - 1, 0, -1):
        if any(values[i]):
            prev_row = values[i]
            break
    else:
        prev_row = [""] * len(headers)

    # Format tanggal sesuai baris sebelumnya
    today = datetime.today()
    formatted_date = parse_date_format(prev_row[header_map.get('y', 0)], today)

    rows = []
    for path in sub_paths:
        row = [""] * len(headers)

        # Isi nama titik A & B
        parts = path["name"].split("-")
        row[0] = parts[0].strip() if len(parts) > 0 else path["name"]
        row[1] = parts[1].strip() if len(parts) > 1 else path["name"]

        # FO pattern: FO 24/2T atau FO 48C/4T
        fo_match = re.search(r"FO\s*(\d+)[A-Z]*\/(\d+)T", path["name"], re.IGNORECASE)
        if fo_match:
            m_val = fo_match.group(1)  # jumlah core
            j_val = fo_match.group(2)  # jumlah tray
            if 'j' in header_map:
                row[header_map['j']] = j_val
            if 'm' in header_map:
                row[header_map['m']] = m_val

        # AE pattern: AE-775M atau AE 775 M
        ae_match = re.search(r"AE[-\s]?(\d+)\s*M", path["name"], re.IGNORECASE)
        if ae_match and 'q' in header_map:
            row[header_map['q']] = ae_match.group(1)

        # Panjang path dari hasil hitung geometris
        if 'p' in header_map:
            row[header_map['p']] = path["length_m"]

        # Kolom vendor dan filename
        if 'am' in header_map:
            row[header_map['am']] = vendor
        if 'ak' in header_map:
            row[header_map['ak']] = filename

        # Kolom tanggal
        if 'y' in header_map:
            row[header_map['y']] = formatted_date

        # Kolom duplikat dari baris atas
        for col in ['c', 'd', 'e', 'f', 'k', 'l', 'u', 'v', 'aj']:
            if col in header_map:
                row[header_map[col]] = prev_row[header_map[col]]

        rows.append(row)

    sheet.append_rows(rows)
    return len(rows)
