import streamlit as st
from fastkml import kml
from pyproj import Transformer
import ezdxf
import io

# Folder target
target_folders = {
    'FDT',
    'NEW POLE 7-3',
    'NEW POLE 7-4',
    'EXISTING POLE EMR 7-4',
    'FAT',
    'HP COVER'
}

# Setup UTM zone 60S transformer
transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)

def parse_kml_bytes(kml_bytes):
    k = kml.KML()
    k.from_string(kml_bytes)  # langsung pakai bytes, BUKAN .decode()
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
    buffer = io.BytesIO()
    doc.write(buffer)
    buffer.seek(0)
    return buffer

# Streamlit UI
st.title("📍 Konversi KML ke DXF (UTM Zone 60S)")

uploaded_file = st.file_uploader("📤 Upload file .kml kamu di sini", type=["kml"])

if uploaded_file is not None:
    try:
        kml_bytes = uploaded_file.read()
        data = parse_kml_bytes(kml_bytes)

        if not data:
            st.warning("⚠️ Tidak ditemukan placemark dalam folder target.")
        else:
            dxf_buffer = generate_dxf(data)
            st.success("✅ File DXF berhasil dibuat!")
            st.download_button("⬇️ Download DXF", dxf_buffer, file_name="output.dxf", mime="application/dxf")
    except Exception as e:
        st.error(f"❌ Terjadi error saat memproses file:\n\n{e}")
