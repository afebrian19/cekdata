    import pandas as pd
    import streamlit as st
    from io import BytesIO

    # Heading Aplikasi
    st.markdown("# App development by A. Febriansyah")  # Tambahkan H1 di sini
    st.title("Aplikasi Filter Kolom Data Excel (.xls & .xlsx)")

    # Unggah File
    uploaded_file = st.file_uploader("Unggah file Excel Anda (.xls atau .xlsx)", type=["xls", "xlsx"])

    if uploaded_file:
        try:
            # Deteksi jenis file berdasarkan ekstensi
            file_extension = uploaded_file.name.split(".")[-1]
            
            # Pilih engine berdasarkan format file
            if file_extension == "xls":
                engine = "xlrd"  # Untuk file Excel 97-2003
            elif file_extension == "xlsx":
                engine = "openpyxl"  # Untuk file Excel modern
            else:
                st.error("Format file tidak didukung. Harap unggah file .xls atau .xlsx.")
                st.stop()

            # Baca file Excel
            df = pd.read_excel(uploaded_file, engine=engine)
            
            # Kolom yang ingin diambil
            selected_columns = ["NOREKENING", "_PRODUK", "NAMA", "_KOLEK", "PLAFOND", "BAKIDEBET", "PETUGAS"]
            
            # Periksa apakah kolom tersedia
            available_columns = [col for col in selected_columns if col in df.columns]
            
            if available_columns:
                # Filter data berdasarkan kolom yang dipilih
                filtered_df = df[available_columns]
                
                # Ubah format kolom 'NOREKENING' jika ada di dalam data
                if "NOREKENING" in filtered_df.columns:
                    filtered_df["NOREKENING"] = filtered_df["NOREKENING"].astype(str).str.replace(",", ".", regex=False)
                
                # Tampilkan hasil
                st.write("### Data yang Difilter")
                st.dataframe(filtered_df)
                
                # Membuat file Excel untuk diunduh
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    filtered_df.to_excel(writer, index=False, sheet_name="Sheet1")
                processed_data = output.getvalue()

                # Tombol unduh file
                st.download_button(
                    label="Unduh Data yang Difilter",
                    data=processed_data,
                    file_name="filtered_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Kolom yang dipilih tidak ditemukan dalam file Excel Anda.")
        
        except Exception as e:
            st.error(f"Terjadi kesalahan saat membaca file: {e}")