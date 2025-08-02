import os
import zipfile
import ezdxf
from fastkml import kml
from shapely.geometry import Point, LineString, Polygon
from pyproj import Transformer

# Folder ke Layer mapping
LAYER_MAPPING = {
    "BOUNDARY": "FAT AREA",
    "DISTRIBUTION CABLE": "FO 36 CORE",
    "SLING WIRE": "FO STRAND AE"
}

# Folder ke block NW
BLOCK_FOLDERS = [
    "NEW POLE 7-3", "NEW POLE 7-4",
    "EXISTING POLE EMR 7-3", "EXISTING POLE EMR 7-4"
]

# Folder ke kategori text matchprop
TEXT_FOLDER_MAP = {
    "FDT": "NN-",
    "FAT": "SRMRW16.067.B01",
    "HP COVER": "MR.SRMRW16",
    "NEW POLE 7-3": "MR.SRMRW16",
    "NEW POLE 7-4": "MR.SRMRW16",
    "EXISTING POLE EMR 7-3": "MR.SRMRW16",
    "EXISTING POLE EMR 7-4": "MR.SRMRW16"
}

TARGET_EPSG = "EPSG:32760"  # UTM Zone 60S
transformer = Transformer.from_crs("EPSG:4326", TARGET_EPSG, always_xy=True)


def extract_kml_from_kmz(kmz_path):
    with zipfile.ZipFile(kmz_path, 'r') as zf:
        for name in zf.namelist():
            if name.endswith('.kml'):
                return zf.read(name).decode('utf-8')
    return None


def parse_kml_geometries(kml_string):
    k = kml.KML()
    k.from_string(kml_string)
    result = []

    def _parse_features(features, folder_name=""):
        for f in features:
            if isinstance(f, kml.Placemark):
                geom = f.geometry
                if isinstance(geom, Point):
                    coords = transformer.transform(geom.x, geom.y)
                    result.append((folder_name, 'POINT', coords))
                elif isinstance(geom, LineString):
                    coords = [transformer.transform(x, y) for x, y in geom.coords]
                    result.append((folder_name, 'LINE', coords))
                elif isinstance(geom, Polygon):
                    coords = [transformer.transform(x, y) for x, y in geom.exterior.coords]
                    result.append((folder_name, 'POLYGON', coords))
            elif hasattr(f, 'features'):
                new_folder_name = f.name or folder_name
                _parse_features(f.features(), new_folder_name)

    _parse_features(k.features())
    return result


def draw_to_dxf(template_dxf, output_dxf, parsed):
    doc = ezdxf.readfile(template_dxf)
    msp = doc.modelspace()

    matchprop_hp = None
    matchprop_pole = None
    matchprop_sr = None

    for e in msp.query('TEXT'):
        txt = e.dxf.text.upper()
        if 'NN-' in txt:
            matchprop_hp = e.dxf
        elif 'MR.SRMRW16' in txt:
            matchprop_pole = e.dxf
        elif 'SRMRW16.067.B01' in txt:
            matchprop_sr = e.dxf

    # Hitung offset tengah untuk geser semua objek
    all_coords = [pt for _, _, pts in parsed for pt in (pts if isinstance(pts, list) else [pts])]
    if all_coords:
        avg_x = sum(x for x, _ in all_coords) / len(all_coords)
        avg_y = sum(y for _, y in all_coords) / len(all_coords)
    else:
        avg_x = avg_y = 0

    for folder, gtype, coords in parsed:
        layer_name = LAYER_MAPPING.get(folder.upper(), folder.upper())

        if gtype == 'POINT':
            x, y = coords[0] - avg_x, coords[1] - avg_y
            if folder.upper() in [f.upper() for f in BLOCK_FOLDERS]:
                msp.add_blockref('NW', insert=(x, y), dxfattribs={"layer": layer_name})
            elif folder.upper() in TEXT_FOLDER_MAP:
                if 'FDT' in folder.upper() and matchprop_hp:
                    msp.add_text("FDT", dxfattribs=matchprop_hp).set_pos((x, y))
                elif 'FAT' in folder.upper() and matchprop_sr:
                    msp.add_text("FAT", dxfattribs=matchprop_sr).set_pos((x, y))
                elif 'HP COVER' in folder.upper() and matchprop_pole:
                    msp.add_text("HP", dxfattribs=matchprop_pole).set_pos((x, y))
                else:
                    msp.add_circle((x, y), radius=2, dxfattribs={"layer": layer_name})
            else:
                msp.add_circle((x, y), radius=2, dxfattribs={"layer": layer_name})

        elif gtype in ['LINE', 'POLYGON']:
            offset_coords = [(x - avg_x, y - avg_y) for x, y in coords]
            msp.add_lwpolyline(offset_coords, dxfattribs={"layer": layer_name})

    doc.saveas(output_dxf)


if __name__ == '__main__':
    kmz_path = 'input.kmz'
    template_dxf = 'template_ref.dxf'
    output_dxf = 'template_ref.dxf'  # Overwrite template directly

    kml_data = extract_kml_from_kmz(kmz_path)
    parsed = parse_kml_geometries(kml_data)
    draw_to_dxf(template_dxf, output_dxf, parsed)
