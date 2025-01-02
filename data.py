import pandas as pd
import streamlit as st
from io import BytesIO
import time

# Heading Aplikasi
st.set_page_config(page_title="Aplikasi Tracing Perubahan Kolektabilitas", layout="wide")

# Gaya Umum Aplikasi
st.markdown("""
    <style>
    /* Styling untuk background halaman */
    .reportview-container {
        background-color: #f4f6f9;
    }
    /* Judul utama */
    .title {
        font-size: 36px;
        font-weight: bold;
        color: #1a3a5c;
        text-align: center;
        padding: 10px 0;
    }
    /* Styling teks penjelasan */
    .description {
        font-size: 18px;
        color: #555555;
        line-height: 1.8;
        text-align: center;
    }
    /* Sidebar */
    .sidebar .sidebar-content {
        background-color: #2C3E50;
        color: white;
    }
    .sidebar .sidebar-header {
        font-size: 20px;
        font-weight: bold;
        color: white;
    }
    /* Styling tombol dan elemen interaktif */
    .stButton>button {
        background-color: #3498db;
        color: white;
        font-weight: bold;
        border-radius: 5px;
        padding: 10px 20px;
        font-size: 16px;
    }
    /* Styling tabel */
    .dataframe {
        font-size: 14px;
        color: #2c3e50;
    }
    /* Footer */
    .footer {
        text-align: center;
        font-size: 16px;
        color: #7f8c8d;
        padding: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# Judul Aplikasi
st.markdown('<div class="title">Aplikasi Tracing Perubahan Kolektabilitas</div>', unsafe_allow_html=True)

# Penjelasan Aplikasi
st.markdown("""
    <p class="description">
        Aplikasi ini digunakan untuk memantau kolektabilitas Nasabah.
    </p>
""", unsafe_allow_html=True)

# Menambahkan teks "Development by A. Febriansyah" di bawah judul aplikasi dengan bold
st.markdown("""
    <p class="footer">Development by A. Febriansyah</p>
""", unsafe_allow_html=True)

# Sidebar untuk Unggah File Excel
st.sidebar.header("Unggah File Excel")
uploaded_file_bulan_lalu = st.sidebar.file_uploader("Unggah file Excel untuk **Bulan Lalu** (.xls atau .xlsx)", type=["xls", "xlsx"], key="bulan_lalu")
uploaded_file_data_saat_ini = st.sidebar.file_uploader("Unggah file Excel untuk **Data Saat Ini** (.xls atau .xlsx)", type=["xls", "xlsx"], key="data_saat_ini")

# Fungsi untuk format Rp IDR
def format_rp(value):
    try:
        return f"Rp {value:,.0f}"
    except:
        return value

# Fungsi untuk highlight berdasarkan perubahan kolektabilitas
def highlight_kocek(val):
    color = 'background-color: white;'  # Default background
    try:
        if val["_KOLEK_BULAN_LALU"] == 0:
            color = 'background-color: #ADD8E6;'  # Biru terang jika bulan lalu 0
        elif val["_KOLEK_SAAT_INI"] > val["_KOLEK_BULAN_LALU"]:
            color = 'background-color: red; color: white;'  # Merah jika naik
        elif val["_KOLEK_SAAT_INI"] < val["_KOLEK_BULAN_LALU"]:
            color = 'background-color: green; color: white;'  # Hijau jika turun
    except:
        pass
    return [color] * len(val)

# Styling DataFrame
def style_dataframe(df):
    return df.style.apply(highlight_kocek, axis=1)

# Periksa apakah file diunggah
if uploaded_file_bulan_lalu and uploaded_file_data_saat_ini:
    try:
        # Menampilkan loading spinner
        with st.spinner('Memproses file **Bulan Lalu** dan **Data Saat Ini**...'):
            # Deteksi jenis file
            file_extension_bulan_lalu = uploaded_file_bulan_lalu.name.split(".")[-1]
            file_extension_data_saat_ini = uploaded_file_data_saat_ini.name.split(".")[-1]

            # Tentukan engine untuk masing-masing file
            if file_extension_bulan_lalu == "xls":
                engine_bulan_lalu = "xlrd"
            else:
                engine_bulan_lalu = "openpyxl"

            if file_extension_data_saat_ini == "xls":
                engine_data_saat_ini = "xlrd"
            else:
                engine_data_saat_ini = "openpyxl"

            # Membaca file
            df_bulan_lalu = pd.read_excel(uploaded_file_bulan_lalu, engine=engine_bulan_lalu)
            df_data_saat_ini = pd.read_excel(uploaded_file_data_saat_ini, engine=engine_data_saat_ini)

            # Kolom yang ingin diambil
            selected_columns = ["NOREKENING", "_PRODUK", "NAMA", "_KOLEK", "PLAFOND", "BAKIDEBET", "PETUGAS"]
            required_columns = ["NOREKENING", "_KOLEK"]

            # Validasi kolom wajib
            if not all(col in df_bulan_lalu.columns for col in required_columns):
                st.error("Kolom wajib (NOREKENING, _KOLEK) tidak ditemukan di file Bulan Lalu.")
                st.stop()
            if not all(col in df_data_saat_ini.columns for col in required_columns):
                st.error("Kolom wajib (NOREKENING, _KOLEK) tidak ditemukan di file Data Saat Ini.")
                st.stop()

            # Filter data
            df_bulan_lalu_filtered = df_bulan_lalu[[col for col in selected_columns if col in df_bulan_lalu.columns]]
            df_data_saat_ini_filtered = df_data_saat_ini[[col for col in selected_columns if col in df_data_saat_ini.columns]]

            # Konversi _KOLEK ke integer
            df_bulan_lalu_filtered["_KOLEK"] = pd.to_numeric(df_bulan_lalu_filtered["_KOLEK"], errors='coerce').fillna(0).astype(int)
            df_data_saat_ini_filtered["_KOLEK"] = pd.to_numeric(df_data_saat_ini_filtered["_KOLEK"], errors='coerce').fillna(0).astype(int)

            # Progress bar
            progress_bar = st.progress(0)
            for i in range(100):
                time.sleep(0.02)
                progress_bar.progress(i + 1)

            # Gabungkan data berdasarkan NOREKENING
            merged_df = pd.merge(
                df_data_saat_ini_filtered,
                df_bulan_lalu_filtered[["NOREKENING", "_KOLEK"]],
                on="NOREKENING",
                how="left",
                suffixes=("_SAAT_INI", "_BULAN_LALU")
            )

            # Pastikan _KOLEK_BULAN_LALU tetap integer
            merged_df["_KOLEK_BULAN_LALU"] = pd.to_numeric(merged_df["_KOLEK_BULAN_LALU"], errors='coerce').fillna(0).astype(int)

            # Tambahkan kolom status
            def add_status(row):
                if row["_KOLEK_SAAT_INI"] > row["_KOLEK_BULAN_LALU"]:
                    return "Naik"
                elif row["_KOLEK_SAAT_INI"] < row["_KOLEK_BULAN_LALU"]:
                    return "Turun"
                elif row["_KOLEK_BULAN_LALU"] == 0:
                    return "Baru"
                return "Tidak Berubah"
            
            merged_df["Status"] = merged_df.apply(add_status, axis=1)

            # Format kolom sebagai IDR
            if "BAKIDEBET" in merged_df.columns:
                merged_df["BAKIDEBET"] = merged_df["BAKIDEBET"].apply(format_rp)
            if "PLAFOND" in merged_df.columns:
                merged_df["PLAFOND"] = merged_df["PLAFOND"].apply(format_rp)

            # Styling DataFrame
            styled_df = style_dataframe(merged_df)

            # Tampilkan hasil
            st.write("### Hasil Tracing Perubahan Kolek **Data Saat Ini** dan **Bulan Lalu**")
            st.markdown("Perbandingan antara data kolektabilitas **Data Saat Ini** dengan **Bulan Lalu** berdasarkan **NOREKENING**.")
            st.dataframe(styled_df)

            # Membuat file Excel untuk diunduh
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                merged_df.to_excel(writer, index=False, sheet_name="Sheet1")
            processed_data = output.getvalue()

            # Tombol unduh file
            st.download_button(
                label="Unduh Data",
                data=processed_data,
                file_name="gabungan_data_perubahan_kolektabilitas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
else:
    st.info("Silakan unggah kedua file Excel untuk **Bulan Lalu** dan **Data Saat Ini** untuk memulai perbandingan.")
