# ==== kmz_parser.py ====
import zipfile
import xml.etree.ElementTree as ET
from shapely.geometry import LineString
from fastkml import kml


def extract_points_from_kmz(kmz_path):
    def parse_kml_coordinates(coords_str):
        coords = [c.split(',') for c in coords_str.strip().split() if ',' in c]
        return [(float(lat), float(lon)) for lon, lat, *_ in coords]

    fat_points, poles_cluster, poles_subfeeder = [], [], []
    fdt_points, cable_cluster, cable_subfeeder = [], [], []

    with zipfile.ZipFile(kmz_path, 'r') as zf:
        kml_file = next(f for f in zf.namelist() if f.lower().endswith(".kml"))
        kml_data = zf.read(kml_file).decode("utf-8")

    k = kml.KML()
    k.from_string(kml_data.encode("utf-8"))
    ns = "{http://www.opengis.net/kml/2.2}"

    def walk_kml(features, folder_path=""):
        items = []
        for f in features:
            name = f.name.upper() if hasattr(f, 'name') else "UNKNOWN"
            path = f"{folder_path}/{name}" if folder_path else name
            if hasattr(f, 'features'):
                items += walk_kml(f.features(), path)
            elif hasattr(f, 'geometry'):
                geom = f.geometry
                if hasattr(f, 'description'):
                    desc = f.description
                else:
                    desc = ""
                items.append({"name": name, "geometry": geom, "path": path, "description": desc})
        return items

    all_features = walk_kml(k.features())

    for item in all_features:
        path = item['path']
        geom = item['geometry']
        desc = item.get("description", "")

        if geom.geom_type == 'Point':
            lat, lon = geom.y, geom.x
            if path.startswith("FAT"):
                fat_points.append({"name": item['name'], "lat": lat, "lon": lon, "path": path})
            elif path.startswith("NEW POLE 7-3"):
                poles_cluster.append({"name": item['name'], "lat": lat, "lon": lon, "path": path, "folder": "7m3inch", "height": "7"})
            elif path.startswith("NEW POLE 7-4"):
                poles_cluster.append({"name": item['name'], "lat": lat, "lon": lon, "path": path, "folder": "7m4inch", "height": "7"})
                poles_subfeeder.append({"name": item['name'], "lat": lat, "lon": lon, "path": path})
            elif path.startswith("NEW POLE 9-4"):
                poles_cluster.append({"name": item['name'], "lat": lat, "lon": lon, "path": path, "folder": "9m4inch", "height": "9"})
            elif path.startswith("FDT"):
                fdt_points.append({"name": item['name'], "lat": lat, "lon": lon, "path": path, "description": desc})
            elif path.startswith("CABLE"):
                cable_subfeeder.append({"name": item['name'], "lat": lat, "lon": lon, "path": path})
            elif path.startswith("DISTRIBUTION CABLE"):
                cable_cluster.append({"name": item['name'], "lat": lat, "lon": lon, "path": path})

        elif geom.geom_type in ['LineString', 'MultiLineString']:
            coords = list(geom.coords)
            path_length = LineString(coords).length * 111_319.9  # approx meter conversion
            if path.startswith("DISTRIBUTION CABLE"):
                cable_cluster.append({"name": item['name'], "path_length": round(path_length, 2), "path": path})
            elif path.startswith("CABLE"):
                cable_subfeeder.append({"name": item['name'], "path_length": round(path_length, 2), "path": path})

    return fat_points, poles_cluster, poles_subfeeder, fdt_points, cable_cluster, cable_subfeeder


# ==== sheet_writer.py ====
import re
from datetime import datetime
from math import dist


def extract_value(text, pattern, default=None, as_int=False):
    match = re.search(pattern, text)
    if match:
        val = match.group(1)
        return int(val) if as_int else val
    return default


def find_nearest_pole(source, poles):
    return min(poles, key=lambda p: dist([source['lat'], source['lon']], [p['lat'], p['lon']]), default={}).get('name', '')


def append_to_fdt_sheet(sheet, fdt_points, poles, kmz_filename, district, subdistrict, vendor):
    headers = sheet.row_values(1)
    header_map = {h.strip().lower(): i for i, h in enumerate(headers)}
    prev_row = next((r for r in reversed(sheet.get_all_values()) if any(r)), [""] * len(headers))
    today = datetime.today()
    formatted_date = today.strftime("%d/%m/%Y") if '/' in prev_row[header_map['ah']] else today.strftime("%Y-%m-%d")

    rows = []
    for fdt in fdt_points:
        name = fdt['name']
        row = [""] * len(headers)
        row[1:5] = prev_row[1:5]
        row[5] = district.upper()
        row[6] = subdistrict.upper()
        row[7] = kmz_filename
        row[8] = name
        row[9] = name
        row[10] = fdt['lat']
        row[11] = fdt['lon']
        row[13:15] = prev_row[13:15]
        row[24:27] = prev_row[24:27]
        row[28] = vendor.upper()
        row[33] = vendor.upper()
        row[34] = formatted_date
        row[18] = find_nearest_pole(fdt, [p for p in poles if p['folder'] == '7m4inch'])

        if 'templatecode' in header_map:
            row[header_map['templatecode']] = fdt['description']

        # Rules for M (12), R (17), AP (40)
        core = extract_value(name, r'FDT\s*(\d+)', as_int=True)
        if core:
            row[12] = {48: '2', 72: '3', 96: '4'}.get(core, '')
            row[17] = {48: '4', 72: '6', 96: '8'}.get(core, '')
            row[40] = {48: 'FDT TYPE 48 CORE', 72: 'FDT TYPE 72 CORE', 96: 'FDT TYPE 96 CORE'}.get(core, '')

        rows.append(row)

    sheet.append_rows(rows)


def append_to_cable_cluster_sheet(sheet, cables, vendor, kmz_filename):
    headers = sheet.row_values(1)
    header_map = {h.strip().lower(): i for i, h in enumerate(headers)}
    prev_row = next((r for r in reversed(sheet.get_all_values()) if any(r)), [""] * len(headers))
    today = datetime.today()
    formatted_date = today.strftime("%d/%m/%Y") if '/' in prev_row[header_map['y']] else today.strftime("%Y-%m-%d")

    rows = []
    for c in cables:
        row = [""] * len(headers)
        row[0] = row[1] = c['name']
        row[2:6] = prev_row[2:6]
        row[10] = prev_row[10]
        row[20:22] = prev_row[20:22]
        row[38] = vendor.upper()
        row[36] = kmz_filename
        row[24] = formatted_date

        row[9] = str(extract_value(c['name'], r'FO\s*(\d+)/(\d+)T', default=''))
        row[16] = str(extract_value(c['name'], r'AE[-\s]*(\d+)', default=''))
        if 'path_length' in c:
            row[15] = f"{c['path_length']:.2f}"

        rows.append(row)

    sheet.append_rows(rows)


def append_to_subfeeder_cable_sheet(sheet, cables, vendor, kmz_filename):
    headers = sheet.row_values(1)
    header_map = {h.strip().lower(): i for i, h in enumerate(headers)}
    prev_row = next((r for r in reversed(sheet.get_all_values()) if any(r)), [""] * len(headers))
    today = datetime.today()
    formatted_date = today.strftime("%d/%m/%Y") if '/' in prev_row[header_map['y']] else today.strftime("%Y-%m-%d")

    rows = []
    for c in cables:
        row = [""] * len(headers)
        row[0] = row[1] = c['name']
        row[2:6] = prev_row[2:6]
        row[10:12] = prev_row[10:12]
        row[20:22] = prev_row[20:22]
        row[35] = prev_row[35]
        row[38] = vendor.upper()
        row[36] = kmz_filename
        row[24] = formatted_date

        row[9] = str(extract_value(c['name'], r'FO\s*(\d+)/(\d+)T', default=''))
        row[12] = str(extract_value(c['name'], r'FO\s*(\d+)', default=''))
        row[16] = str(extract_value(c['name'], r'AE[-\s]*(\d+)', default=''))
        if 'path_length' in c:
            row[15] = f"{c['path_length']:.2f}"

        rows.append(row)

    sheet.append_rows(rows)
