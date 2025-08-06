# subcable_writer.py
from fdt_cable_writer import parse_fo_capacity, parse_distance, measure_path_length

def process_cable_data(path_data, prev_row, header_map, vendor, kmz_name, is_subfeeder=False):
    from datetime import datetime
    date_fmt = "%d/%m/%Y" if '/' in prev_row[header_map.get('y', 0)] else "%Y-%m-%d"
    today = datetime.today().strftime(date_fmt)

    result_rows = []
    for path in path_data:
        row = [""] * len(header_map)

        name = path['name']
        row[header_map.get('a')] = name
        row[header_map.get('b')] = name

        if is_subfeeder:
            cols = ['c','d','e','f','k','l','u','v','aj']
        else:
            cols = ['c','d','e','f','k','u','v']

        for col in cols:
            idx = header_map.get(col.lower())
            if idx is not None:
                row[idx] = prev_row[idx]

        if 'am' in header_map:
            row[header_map['am']] = vendor.upper()
        if 'ak' in header_map:
            row[header_map['ak']] = kmz_name
        if 'y' in header_map:
            row[header_map['y']] = today

        ports, cores = parse_fo_capacity(name)
        if ports:
            if 'j' in header_map:
                row[header_map['j']] = str(ports)
        if cores:
            if is_subfeeder and 'm' in header_map:
                row[header_map['m']] = str(cores)

        dist = parse_distance(name)
        if dist and 'q' in header_map:
            row[header_map['q']] = str(dist)

        if 'p' in header_map and 'coords' in path:
            row[header_map['p']] = str(measure_path_length(path['coords']))

        result_rows.append(row)

    return result_rows
