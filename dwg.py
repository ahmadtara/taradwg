import os
import zipfile
from xml.etree import ElementTree as ET
import ezdxf
from pyproj import Transformer

# Transformer: EPSG 4326 (latlon) → EPSG 32760 (UTM zone 60S)
transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)

TARGET_FOLDERS = {
    'FDT', 'FAT', 'HP COVER',
    'NEW POLE 7-3', 'NEW POLE 7-4', 'EXISTING POLE EMR 7-4', 'EXISTING POLE EMR 7-3',
    'BOUNDARY', 'DISTRIBUTION CABLE', 'SLING WIRE'
}

LAYER_MAPPING = {
    "BOUNDARY": "FAT AREA",
    "DISTRIBUTION CABLE": "FO 36 CORE",
    "SLING WIRE": "FO STRAND AE",
}

BLOCK_MAPPING = {
    'NEW POLE 7-3': 'NW',
    'NEW POLE 7-4': 'NW',
    'EXISTING POLE EMR 7-4': 'NW',
    'EXISTING POLE EMR 7-3': 'NW',
}

def extract_kmz(kmz_path, extract_dir):
    with zipfile.ZipFile(kmz_path, 'r') as kmz_file:
        kmz_file.extractall(extract_dir)
    return os.path.join(extract_dir, "doc.kml")

def parse_kml(kml_path):
    ns = {'kml': 'http://www.opengis.net/kml/2.2'}
    with open(kml_path, 'rb') as f:
        tree = ET.parse(f)
    root = tree.getroot()
    folders = root.findall('.//kml:Folder', ns)

    points = []
    paths = []

    for folder in folders:
        folder_name_tag = folder.find('kml:name', ns)
        if folder_name_tag is None:
            continue
        folder_name = folder_name_tag.text.strip().upper()
        if folder_name not in TARGET_FOLDERS:
            continue

        placemarks = folder.findall('.//kml:Placemark', ns)
        for pm in placemarks:
            name = pm.find('kml:name', ns)
            coord = pm.find('.//kml:coordinates', ns)
            linestring = pm.find('.//kml:LineString', ns)
            polygon = pm.find('.//kml:Polygon', ns)

            if (linestring or polygon) and coord is not None:
                coord_list = coord.text.strip().split()
                path = []
                for pair in coord_list:
                    lon, lat, *_ = pair.split(',')
                    path.append({
                        "latitude": float(lat),
                        "longitude": float(lon),
                        "folder": folder_name
                    })
                paths.append(path)
            elif name is not None and coord is not None:
                lon, lat, *_ = coord.text.strip().split(',')
                points.append({
                    'name': name.text.strip(),
                    'latitude': float(lat),
                    'longitude': float(lon),
                    'folder': folder_name
                })
    return points, paths

def latlon_to_xy(lat, lon):
    x, y = transformer.transform(lon, lat)
    return x, y

def apply_offset(points_xy):
    xs = [x for x, y in points_xy]
    ys = [y for x, y in points_xy]
    cx, cy = sum(xs)/len(xs), sum(ys)/len(ys)
    return [(x - cx, y - cy) for x, y in points_xy], (cx, cy)

def draw_elements(template_path, points, paths):
    doc = ezdxf.readfile(template_path)
    msp = doc.modelspace()

    # Ambil referensi properti dari TEXT
    matchprop_hp = matchprop_pole = matchprop_sr = None
    for e in msp.query('TEXT'):
        txt = e.dxf.text.upper()
        if 'NN-' in txt:
            matchprop_hp = e.dxf
        elif 'MR.SRMRW16' in txt:
            matchprop_pole = e.dxf
        elif 'SRMRW16.067.B01' in txt:
            matchprop_sr = e.dxf

    # Ambil properti layer PATH
    layer_matchprops = {}
    for e in msp:
        if e.dxf.layer in LAYER_MAPPING.values():
            layer_matchprops[e.dxf.layer] = e.dxf

    # Gabungkan semua koordinat
    all_points = [(p['longitude'], p['latitude']) for p in points]
    all_paths = [(pt['longitude'], pt['latitude']) for path in paths for pt in path]
    all_coords = all_points + all_paths

    if not all_coords:
        raise ValueError("❌ Tidak ada koordinat ditemukan.")

    xy_coords = [latlon_to_xy(lat, lon) for lon, lat in all_coords]
    shifted_coords, offset = apply_offset(xy_coords)

    # Gambar titik
    point_idx = 0
    for p in points:
        x, y = shifted_coords[point_idx]
        folder = p['folder']

        # Gambar simbol
        if folder in BLOCK_MAPPING:
            msp.add_blockref(BLOCK_MAPPING[folder], (x, y), dxfattribs={"layer": folder})
        else:
            msp.add_circle((x, y), radius=2, dxfattribs={"layer": folder})

        # Tambahkan teks nama
        if folder == "HP COVER":
            matchprop = matchprop_hp
        elif "POLE" in folder:
            matchprop = matchprop_pole
        elif folder in ["FAT", "FDT"]:
            matchprop = matchprop_sr
        else:
            matchprop = None

        text_attribs = {
            "insert": (x + 2, y),
            "layer": folder,
            "height": matchprop.height if matchprop else 1.5,
            "color": matchprop.color if matchprop else 7
        }
        msp.add_text(p["name"], dxfattribs=text_attribs)
        point_idx += 1

    # Gambar path (line/polygon)
    for path in paths:
        if not path:
            continue
        folder = path[0]['folder']
        if folder not in LAYER_MAPPING:
            continue

        template_layer = LAYER_MAPPING[folder]
        prop = layer_matchprops.get(template_layer, None)
        attribs = {"layer": template_layer}
        if prop:
            attribs.update({
                "color": prop.color,
                "linetype": prop.linetype
            })

        shifted_xy = [latlon_to_xy(p['latitude'], p['longitude']) for p in path]
        shifted_xy_offset, _ = apply_offset(shifted_xy)
        msp.add_lwpolyline(shifted_xy_offset, close=False, dxfattribs=attribs)

    return doc
