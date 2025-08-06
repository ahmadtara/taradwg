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
