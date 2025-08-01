import streamlit as st
from fastkml import kml
from pyproj import Transformer
import ezdxf
import io
import base64

# Folder yang ingin diambil
target_folders = {
    'FDT',
    'NEW POLE 7-3',
    'NEW POLE 7-4',
    'EXISTING POLE EMR 7-4',
    'FAT',
    'HP COVER'
}

# Transformer: WGS84 ke UTM Zona 60S
transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)

# Fungsi parsing file .kml
def parse_kml(kml_file):
    raw_data = kml_file.getvalue()
    try:
        text = raw_data.decode("utf-8")
    except UnicodeDecodeError:
        text = raw_data.decode("latin-1")

    k = kml.KML()
    k.from_string(text)

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

# Fungsi generate file .dxf
def generate_dxf(data):
    doc = ezdxf.new(dxfversion='R2010')
    msp = doc.modelspace()

    for item in data:
        x, y = item['easting'], item['northing']
        label = f"{item['placemark']}\n({item['folder']})"
        msp.add_point((x, y))
        msp.add_text(label, dxfattribs={'height': 2.5}).set_pos((x, y + 3), align='CENTER')

    buffer = io.BytesIO()
    doc.saveas(buffer)
    buffer.seek(0)
    return buffer

# Streamlit UI
st.title("📐 Konversi KML ke DXF (UTM Zona 60S)")
st.markdown("Upload file `.kml`, program akan mengambil titik dari folder tertentu dan mengonversi ke file AutoCAD `.dxf`.")

uploaded_file = st.file_uploader("📤 Upload file .kml", type=["kml"])

if uploaded_file:
    try:
        data = parse_kml(uploaded_file)

        if not data:
            st.warning("⚠️ Tidak ada placemark ditemukan dalam folder target.")
        else:
            st.success(f"✅ Ditemukan {len(data)} titik dari folder target.")
            dxf_bytes = generate_dxf(data)
            st.download_button(
                label="⬇️ Download File DXF",
                data=dxf_bytes,
                file_name="output_kml_to_dxf.dxf",
                mime="application/dxf"
            )
    except Exception as e:
        st.error(f"❌ Terjadi error saat memproses file:\n```\n{str(e)}\n```")
