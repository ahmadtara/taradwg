import streamlit as st
from fastkml import kml
from pyproj import Transformer
import ezdxf
import tempfile
import os

# Set folder target
target_folders = {
    'FDT',
    'NEW POLE 7-3',
    'NEW POLE 7-4',
    'EXISTING POLE EMR 7-4',
    'FAT',
    'HP COVER'
}

# Transformer dari WGS84 ke UTM zona 60S (EPSG:32760 = UTM zona 60S)
transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)

def parse_kml(kml_file):
    k = kml.KML()
    k.from_string(kml_file.read().decode("utf-8"))

    points = []

    def extract(features):
        for f in features:
            if hasattr(f, 'features'):
                fname = f.name.strip() if f.name else ""
                if fname in target_folders:
                    for placemark in f.features():
                        if hasattr(placemark, 'geometry') and placemark.geometry.geom_type == 'Point':
                            lon, lat = placemark.geometry.x, placemark.geometry.y
                            easting, northing = transformer.transform(lon, lat)
                            points.append({
                                'folder': fname,
                                'placemark': placemark.name,
                                'easting': round(easting, 3),
                                'northing': round(northing, 3)
                            })
                extract(f.features())

    extract(k.features())
    return points

def generate_dxf(data):
    doc = ezdxf.new(dxfversion='R2010')
    msp = doc.modelspace()

    for item in data:
        x, y = item['easting'], item['northing']
        label = f"{item['placemark']}\n({item['folder']})"
        msp.add_point((x, y))
        msp.add_text(label, dxfattribs={'height': 2.5}).set_pos((x, y + 3), align='CENTER')

    # Simpan ke file sementara
    temp_dxf = tempfile.NamedTemporaryFile(delete=False, suffix=".dxf")
    doc.saveas(temp_dxf.name)
    return temp_dxf.name

# Streamlit UI
st.title("Konversi KML ke DXF (UTM 60S)")
uploaded_file = st.file_uploader("Upload file .kml", type=["kml"])

if uploaded_file:
    with st.spinner("🔄 Memproses file..."):
        data = parse_kml(uploaded_file)
        if not data:
            st.warning("⚠️ Tidak ditemukan titik dalam folder target.")
        else:
            dxf_path = generate_dxf(data)
            with open(dxf_path, "rb") as f:
                st.success("✅ Berhasil dikonversi ke DXF!")
                st.download_button(
                    label="⬇️ Download DXF",
                    data=f,
                    file_name=os.path.basename(dxf_path),
                    mime="application/dxf"
                )
