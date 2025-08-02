import streamlit as st 
import zipfile
import os
from xml.etree import ElementTree as ET
import ezdxf
from pyproj import Transformer

st.set_page_config(page_title="KMZ → DXF Converter with Matchprop", layout="wide")

transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)

target_folders = {
    'FDT', 'FAT', 'HP COVER', 'NEW POLE 7-3', 'NEW POLE 7-4', 'EXISTING POLE EMR 7-4', 'EXISTING POLE EMR 7-3'
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
            if name is not None and coord is not None:
                name_text = name.text.strip()
                lon, lat, *_ = coord.text.strip().split(',')
                points.append({'name': name_text, 'latitude': float(lat), 'longitude': float(lon), 'folder': folder_name})
    return points

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

def draw_to_dxf(classified, template_path):
    template_doc = ezdxf.readfile(template_path)
    template_msp = template_doc.modelspace()

    matchprop_hp = None
    matchprop_pole = None
    matchprop_sr = None

    for e in template_msp.query('TEXT'):
        txt = e.dxf.text.upper()
        if 'NN-' in txt:
            matchprop_hp = e.dxf
        elif 'MR.SRMRW16' in txt:
            matchprop_pole = e.dxf
        elif 'SRMRW16.067.B01' in txt:
            matchprop_sr = e.dxf

    doc = ezdxf.new(dxfversion="R2010")
    msp = doc.modelspace()

    # ✅ Salin block NW dari template ke dokumen output
    if "NW" in template_doc.blocks:
        source_block = template_doc.blocks.get("NW")
        dest_block = doc.blocks.new(name="NW", base_point=source_block.block.dxf.base_point)
        for entity in source_block:
            dest_block.add_entity(entity.copy())
    else:
        st.warning("⚠️ Block 'NW' tidak ditemukan di template.")

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
        if layer_name not in doc.layers:
            doc.layers.add(name=layer_name)
        for obj in data:
            x, y = obj['xy']

            # Tambahkan block NW jika POLE
            if layer_name in ["NEW_POLE", "EXISTING_POLE"]:
                try:
                    msp.add_blockref("NW", (x, y), dxfattribs={"layer": layer_name})
                except Exception as e:
                    st.error(f"❌ Gagal insert block NW: {e}")
            elif layer_name != "HP_COVER":
                msp.add_circle((x, y), radius=2, dxfattribs={"layer": layer_name})

            # Tambahkan teks semua titik
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

    return doc

st.title("🏗️ KMZ → DXF Converter with Matchprop")
st.write("Konversi file KMZ menjadi DXF dengan properti teks yang ditiru dari template (matchprop).")

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
        points = parse_kml(kml_path)
        classified = classify_points(points)

        updated_doc = draw_to_dxf(classified, "template_ref.dxf")
        if updated_doc:
            updated_doc.saveas(output_dxf)

    if os.path.exists(output_dxf):
        st.success("✅ Konversi berhasil! DXF sudah dibuat.")
        with open(output_dxf, "rb") as f:
            st.download_button("⬇️ Download DXF", f, file_name="output_from_kmz.dxf")

        st.markdown("### 📊 Ringkasan Objek")
        for layer_name, objs in classified.items():
            st.write(f"- **{layer_name}**: {len(objs)} titik")
