import streamlit as st
import zipfile
import os
from xml.etree import ElementTree as ET
import ezdxf
from pyproj import Transformer

st.set_page_config(page_title="KMZ → DXF Converter with Matchprop", layout="wide")

transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)

# Folder yang didukung
target_folders = {
    'FDT', 'FAT', 'HP COVER',
    'NEW POLE 7-3', 'NEW POLE 7-4', 'EXISTING POLE EMR 7-4', 'EXISTING POLE EMR 7-3',
    'BOUNDARY', 'DISTRIBUTION CABLE', 'SLING WIRE'
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
        if folder_name not in target_folders:
            continue

        placemarks = folder.findall('.//kml:Placemark', ns)
        for pm in placemarks:
            name_tag = pm.find('kml:name', ns)
            name = name_tag.text.strip() if name_tag is not None else ''

            # Cek koordinat untuk <Point>
            coord_tag = pm.find('.//kml:Point/kml:coordinates', ns)
            if coord_tag is not None:
                lon, lat, *_ = coord_tag.text.strip().split(',')
                points.append({
                    'name': name,
                    'latitude': float(lat),
                    'longitude': float(lon),
                    'folder': folder_name
                })

            # Cek koordinat untuk <LineString> dan <Polygon>
            for path_tag in pm.findall('.//kml:LineString/kml:coordinates', ns) + \
                            pm.findall('.//kml:Polygon/kml:outerBoundaryIs/kml:LinearRing/kml:coordinates', ns):
                raw_coords = path_tag.text.strip().split()
                coord_list = []
                for c in raw_coords:
                    lon, lat, *_ = c.split(',')
                    coord_list.append((float(lat), float(lon)))
                paths.append({
                    'name': name,
                    'folder': folder_name,
                    'path': coord_list
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

def classify_points(points):
    classified = {
        "FDT": [], "FAT": [], "HP_COVER": [], "NEW_POLE": [], "EXISTING_POLE": [], "POLE": []
    }
    for p in points:
        folder = p['folder']
        if "FDT" in folder:
            classified["FDT"].append(p)
        elif "FAT" in folder:
            classified["FAT"].append(p)
        elif "HP COVER" in folder:
            classified["HP_COVER"].append(p)
        elif "NEW POLE" in folder:
            classified["NEW_POLE"].append(p)
        elif "EXISTING" in folder or "EMR" in folder:
            classified["EXISTING_POLE"].append(p)
        else:
            classified["POLE"].append(p)
    return classified

def classify_paths(paths):
    classified_paths = {
        "BOUNDARY": [], "DISTRIBUTION_CABLE": [], "SLING_WIRE": []
    }
    for p in paths:
        folder = p['folder']
        if "BOUNDARY" in folder:
            classified_paths["BOUNDARY"].append(p)
        elif "DISTRIBUTION" in folder:
            classified_paths["DISTRIBUTION_CABLE"].append(p)
        elif "SLING" in folder:
            classified_paths["SLING_WIRE"].append(p)
    return classified_paths

def draw_to_dxf(classified_points, classified_paths, template_path):
    template_doc = ezdxf.readfile(template_path)
    msp = template_doc.modelspace()

    all_points_xy = [latlon_to_xy(p['latitude'], p['longitude']) for c in classified_points.values() for p in c]
    if not all_points_xy:
        st.error("❌ Tidak ada titik ditemukan di KMZ!")
        return None

    shifted_points, (cx, cy) = apply_offset(all_points_xy)

    idx = 0
    for category in classified_points.values():
        for i in range(len(category)):
            category[i]['xy'] = shifted_points[idx]
            idx += 1

    for layer_name, data in classified_points.items():
        for obj in data:
            x, y = obj['xy']
            if layer_name in ["NEW_POLE", "EXISTING_POLE"]:
                msp.add_blockref('NW', insert=(x, y), dxfattribs={"layer": layer_name})
            else:
                msp.add_circle((x, y), radius=2, dxfattribs={"layer": layer_name})
            msp.add_text(obj["name"], dxfattribs={"layer": layer_name, "insert": (x + 2, y)})

    for layer, items in classified_paths.items():
        for p in items:
            points_xy = [latlon_to_xy(lat, lon) for lat, lon in p['path']]
            shifted = [(x - cx, y - cy) for x, y in points_xy]
            msp.add_lwpolyline(shifted, close=False, dxfattribs={"layer": layer})

    return template_doc

# UI
st.title("🏗️ KMZ → DXF Converter with Matchprop")
st.write("Konversi file KMZ menjadi DXF dengan properti teks dan garis dari template DXF.")

uploaded_kmz = st.file_uploader("📂 Upload File KMZ", type=["kmz"])
uploaded_template = st.file_uploader("📀 Upload Template DXF", type=["dxf"])

if uploaded_kmz and uploaded_template:
    extract_dir = "temp_kmz"
    os.makedirs(extract_dir, exist_ok=True)
    output_dxf = "converted_output.dxf"

    with open("template_ref.dxf", "wb") as f:
        f.write(uploaded_template.read())

    with st.spinner("🔍 Memproses data..."):
        kml_path = extract_kmz(uploaded_kmz, extract_dir)
        points, paths = parse_kml(kml_path)
        classified_points = classify_points(points)
        classified_paths = classify_paths(paths)

        updated_doc = draw_to_dxf(classified_points, classified_paths, "template_ref.dxf")
        if updated_doc:
            updated_doc.saveas(output_dxf)

    if os.path.exists(output_dxf):
        st.success("✅ Konversi berhasil! DXF sudah dibuat.")
        with open(output_dxf, "rb") as f:
            st.download_button("⬇️ Download DXF", f, file_name="output_from_kmz.dxf")

        st.markdown("### 📊 Ringkasan Objek")
        for layer_name, objs in classified_points.items():
            st.write(f"- **{layer_name}**: {len(objs)} titik")
        for layer_name, objs in classified_paths.items():
            st.write(f"- **{layer_name}**: {len(objs)} path")
