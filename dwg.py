import streamlit as st
import zipfile
import os
from xml.etree import ElementTree as ET
import ezdxf
from pyproj import Transformer

st.set_page_config(page_title="KMZ → DXF Converter with Matchprop", layout="wide")

transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)

target_folders = {
    'FDT', 'FAT', 'HP COVER',
    'NEW POLE 7-3', 'NEW POLE 7-4', 'EXISTING POLE EMR 7-4', 'EXISTING POLE EMR 7-3',
    'BOUNDARY', 'DISTRIBUTION CABLE', 'SLING WIRE'
}

LAYER_MAPPING = {
    "BOUNDARY": "FAT AREA",
    "DISTRIBUTION CABLE": "FO 36 CORE",
    "SLING WIRE": "FO STRAND AE",
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
            name = pm.find('kml:name', ns)
            coord = pm.find('.//kml:coordinates', ns)
            linestring = pm.find('.//kml:LineString', ns)
            polygon = pm.find('.//kml:Polygon', ns)
            if (linestring or polygon) and coord is not None:
                coord_list = coord.text.strip().split()
                path = []
                for pair in coord_list:
                    lon, lat, *_ = pair.split(',')
                    path.append({"latitude": float(lat), "longitude": float(lon), "folder": folder_name})
                paths.append(path)
            elif name is not None and coord is not None:
                lon, lat, *_ = coord.text.strip().split(',')
                points.append({'name': name.text.strip(), 'latitude': float(lat), 'longitude': float(lon), 'folder': folder_name})
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
        "FDT": [], "FAT": [], "HP_COVER": [], "NEW_POLE": [], "EXISTING_POLE": [], "POLE": [],
        "BOUNDARY": [], "DISTRIBUTION CABLE": [], "SLING WIRE": []
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
        elif folder in ["BOUNDARY", "DISTRIBUTION CABLE", "SLING WIRE"]:
            classified[folder].append(p)
        else:
            classified["POLE"].append(p)
    return classified

def draw_to_dxf(classified, paths, template_path):
    doc = ezdxf.readfile(template_path)
    msp = doc.modelspace()

    matchprop_hp = matchprop_pole = matchprop_sr = None
    layer_matchprops = {}

    for e in msp.query('TEXT'):
        txt = e.dxf.text.upper()
        if 'NN-' in txt:
            matchprop_hp = e.dxf
        elif 'MR.SRMRW16' in txt:
            matchprop_pole = e.dxf
        elif 'SRMRW16.067.B01' in txt:
            matchprop_sr = e.dxf
        elif e.dxf.layer in LAYER_MAPPING.values():
            layer_matchprops[e.dxf.layer] = e.dxf

    all_points_xy = []
    for category in classified.values():
        for p in category:
            all_points_xy.append(latlon_to_xy(p['latitude'], p['longitude']))

    if not all_points_xy:
        st.error("❌ Tidak ada titik ditemukan di KMZ!")
        return None

    shifted_points, (cx, cy) = apply_offset(all_points_xy)

    idx = 0
    for category_name, category in classified.items():
        for i in range(len(category)):
            category[i]['xy'] = shifted_points[idx]
            idx += 1

    for layer_name, data in classified.items():
        for obj in data:
            x, y = obj['xy']
            if layer_name in ["NEW_POLE", "EXISTING_POLE"]:
                msp.add_blockref("NW", (x, y), dxfattribs={"layer": layer_name})
            elif layer_name != "HP_COVER":
                msp.add_circle((x, y), radius=2, dxfattribs={"layer": layer_name})

            if layer_name == "HP_COVER":
                matchprop = matchprop_hp
            elif layer_name in ["NEW_POLE", "EXISTING_POLE"]:
                matchprop = matchprop_pole
            elif layer_name in ["FAT", "FDT"]:
                matchprop = matchprop_sr
            else:
                matchprop = None

            if matchprop:
                attribs = {
                    "height": matchprop.height,
                    "layer": layer_name,
                    "color": matchprop.color,
                    "insert": (x + 2, y)
                }
            else:
                attribs = {"height": 1.5, "layer": layer_name, "insert": (x + 2, y)}

            msp.add_text(obj["name"], dxfattribs=attribs)

    for path in paths:
        if not path:
            continue
        folder = path[0]['folder']
        if folder not in LAYER_MAPPING:
            continue
        target_layer = folder
        xy_path = [latlon_to_xy(p['latitude'], p['longitude']) for p in path]
        shifted_xy, _ = apply_offset(xy_path)
        dxf_layer = folder
        template_layer = LAYER_MAPPING[folder]
        prop = layer_matchprops.get(template_layer, None)
        attribs = {"layer": dxf_layer}
        if prop:
            attribs.update({"color": prop.color, "linetype": prop.linetype})
        msp.add_lwpolyline(shifted_xy, close=False, dxfattribs=attribs)

    return doc

st.title("🏗️ KMZ → DXF Converter with Matchprop")
st.write("Konversi file KMZ menjadi DXF dengan properti teks dan layer dari template.")

uploaded_kmz = st.file_uploader("📂 Upload File KMZ", type=["kmz"])
uploaded_template = st.file_uploader("📐 Upload Template DXF", type=["dxf"])

if uploaded_kmz and uploaded_template:
    extract_dir = "temp_kmz"
    os.makedirs(extract_dir, exist_ok=True)
    output_dxf = "converted_output.dxf"

    with open("template_ref.dxf", "wb") as f:
        f.write(uploaded_template.read())

    with st.spinner("🔍 Memproses data..."):
        kml_path = extract_kmz(uploaded_kmz, extract_dir)
        points, paths = parse_kml(kml_path)
        classified = classify_points(points)

        updated_doc = draw_to_dxf(classified, paths, "template_ref.dxf")
        if updated_doc:
            updated_doc.saveas(output_dxf)

    if os.path.exists(output_dxf):
        st.success("✅ Konversi berhasil! DXF sudah dibuat.")
        with open(output_dxf, "rb") as f:
            st.download_button("⬇️ Download DXF", f, file_name="output_from_kmz.dxf")

        st.markdown("### 📊 Ringkasan Objek")
        for layer_name, objs in classified.items():
            st.write(f"- **{layer_name}**: {len(objs)} titik")
