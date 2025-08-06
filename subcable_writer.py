import re
from datetime import datetime
from sheet_utils import get_latest_row_data, parse_date_format, extract_distance

def fill_subfeeder_pekanbaru(sheet, paths, vendor_name, kmz_filename):
    headers = sheet.row_values(1)
    header_map = {h.strip().lower(): i for i, h in enumerate(headers)}

    prev_row = get_latest_row_data(sheet)
    date_fmt = parse_date_format(prev_row[header_map['installationdate']])
    today = datetime.today().strftime(date_fmt)

    all_rows = []
    for path in paths:
        row = [""] * len(headers)

        # A-B: titik nama
        row[0] = path['start_name']
        row[1] = path['end_name']

        # C-F, K, L, U, V, AJ: copy dari baris sebelumnya
        for col in ['c', 'd', 'e', 'f', 'k', 'l', 'u', 'v', 'aj']:
            idx = header_map.get(col.lower())
            if idx is not None:
                row[idx] = prev_row[idx]

        # AM: Vendor Name input
        idx_am = header_map.get('am')
        if idx_am is not None:
            row[idx_am] = vendor_name.upper()

        # AK: nama file kmz
        idx_ak = header_map.get('ak')
        if idx_ak is not None:
            row[idx_ak] = kmz_filename

        # Y: tanggal hari ini
        idx_y = header_map.get('y')
        if idx_y is not None:
            row[idx_y] = today

        # J, M: parsing FO xx/xT
        match = re.search(r'fo\s*(\d+)\s*/\s*(\d+)', path['start_name'].lower())
        if match:
            core = int(match.group(1))
            split = int(match.group(2))
            row[header_map['j']] = str(split) if 'j' in header_map else ""
            row[header_map['m']] = str(core) if 'm' in header_map else ""

        # Q: panjang dari AE xxx M
        ae_match = re.search(r'ae[-\s]*(\d+)\s*m', path['start_name'].lower())
        if ae_match:
            row[header_map['q']] = ae_match.group(1)

        # P: panjang path (meter)
        if 'p' in header_map:
            row[header_map['p']] = round(path.get('length_m', 0), 2)

        all_rows.append(row)

    sheet.append_rows(all_rows)
    return len(all_rows)
