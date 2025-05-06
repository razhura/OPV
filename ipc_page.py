# ipc_page.py

import streamlit as st
import pandas as pd
import io

def parse_kekerasan_excel(file):
    """
    Parsing untuk template Excel pengujian Kekerasan (khusus format stacking)
    """
    try:
        df = pd.read_excel(file, header=None)

        # Ambil baris pertama sebagai header
        header_row = df.iloc[0]
        df = df[1:]
        df.columns = header_row
        df.reset_index(drop=True, inplace=True)

        # Hapus baris yang mengandung 'Rata-rata', 'SD', atau 'RSD' di kolom A
        df = df[~df.iloc[:, 0].astype(str).str.contains("Rata|SD|RSD", na=False)]

        # Ambil semua kolom nomor batch unik dari kolom A
        batch_series = df.iloc[:, 0].dropna().unique()

        result_df = pd.DataFrame()

        for batch in batch_series:
            subset = df[df.iloc[:, 0] == batch]
            if subset.empty:
                continue

            # Ambil data dari kolom E (data 1-5) dan F (data 6-10)
            # Asumsi kolom ke-5 dan ke-6 (indeks 4 dan 5)
            values_e = subset.iloc[0:5, 4]
            values_f = subset.iloc[0:5, 5]
            stacked = pd.concat([values_e, values_f], ignore_index=True)

           result_df[batch] = stacked

        # Set index mulai dari 1 (agar tidak bentrok saat loop)
        result_df.index = range(1, len(result_df) + 1)


        st.write("Data Kekerasan Terstruktur:")
        st.dataframe(result_df)
        return result_df

    except Exception as e:
        st.error(f"Gagal memproses file Kekerasan: {e}")
        return None

def parse_keseragaman_bobot_excel(file):
    """
    Parsing untuk template Excel pengujian Keseragaman Bobot
    """
    try:
        df = pd.read_excel(file, header=None)

        # Ambil baris pertama sebagai header
        header_row = df.iloc[0]
        df = df[1:]
        df.columns = header_row
        df.reset_index(drop=True, inplace=True)

        # Hapus baris yang mengandung 'Rata-rata', 'SD', atau 'RSD' di kolom A
        df = df[~df.iloc[:, 0].astype(str).str.contains("Rata|SD|RSD", na=False)]

        # Ambil semua kolom nomor batch unik dari kolom A
        batch_series = df.iloc[:, 0].dropna().unique()

        result_df = pd.DataFrame()

        for batch in batch_series:
            subset = df[df.iloc[:, 0] == batch]
            if subset.empty:
                continue

            # Ambil data dari kolom E (data 1-5) dan F (data 6-10)
            # Asumsi kolom ke-5 dan ke-6 (indeks 4 dan 5)
            values_e = subset.iloc[0:5, 4]
            values_f = subset.iloc[0:5, 5]
            values_g = subset.iloc[0:5, 6]
            values_h = subset.iloc[0:5, 7]
            stacked = pd.concat([values_e, values_f, values_g, values_h], ignore_index=True)

             result_df[batch] = stacked

        # Set index mulai dari 1 (agar tidak bentrok saat loop)
        result_df.index = range(1, len(result_df) + 1)

        st.write("Data Keseragaman Bobot Terstruktur:")
        st.dataframe(result_df)
        return result_df

    except Exception as e:
        st.error(f"Gagal memproses file Keseragaman Bobot: {e}")
        return None

def parse_tebal_excel(file):
    """
    Parsing untuk template Excel pengujian Tebal
    """
    try:
        df = pd.read_excel(file, header=None)

        # Ambil baris pertama sebagai header
        header_row = df.iloc[0]
        df = df[1:]
        df.columns = header_row
        df.reset_index(drop=True, inplace=True)

        # Ambil semua kolom nomor batch unik dari kolom A
        batch_series = df.iloc[:, 0].dropna().unique()

        result_df = pd.DataFrame()

        for batch in batch_series:
            subset = df[df.iloc[:, 0] == batch]
            if subset.empty:
                continue

            # Ambil data dari kolom E (data 1-3) dan F (data 4-6)
            # Kolom ke-5 dan ke-6 (indeks 4 dan 5)
            values_e = subset.iloc[0:3, 4]
            values_f = subset.iloc[0:3, 5]

            stacked = pd.concat([values_e, values_f], ignore_index=True)

            result_df[batch] = stacked

        # Set index mulai dari 1 (agar tidak bentrok saat loop)
        result_df.index = range(1, len(result_df) + 1)

        st.write("Data Tebal Terstruktur:")
        st.dataframe(result_df)
        return result_df

    except Exception as e:
        st.error(f"Gagal memproses file Tebal: {e}")
        return None
    
def parse_waktu_hancur_friability_excel(file):
    """
    Parsing untuk template Excel pengujian Waktu Hancur dan Friability
    """
    try:
        df = pd.read_excel(file, header=None)

        # Ambil baris pertama sebagai header
        header_row = df.iloc[0]
        df = df[1:]
        df.columns = header_row
        df.reset_index(drop=True, inplace=True)

        # Ambil semua nomor batch unik dari kolom A
        batch_series = df.iloc[:, 0].dropna().unique()

        # Ambil label baris dari kolom E (kolom ke-5 -> index 4)
        parameter_labels = df.iloc[0:2, 4].tolist()

        result_rows = []

        for batch in batch_series:
            subset = df[df.iloc[:, 0] == batch]
            if subset.empty:
                continue

            # Ambil 2 nilai dari kolom E (baris 0 dan 1 relatif terhadap subset)
            values = subset.iloc[0:2, 4].tolist()

            # Gabungkan batch + data
            row = [batch] + values
            result_rows.append(row)

  
            result_df[batch] = stacked

        # Set index mulai dari 1 (agar tidak bentrok saat loop)
        result_df.index = range(1, len(result_df) + 1)

        st.write("Data Tebal Terstruktur:")
        st.dataframe(result_df)
        return result_df

    except Exception as e:
        st.error(f"Gagal memproses file Tebal: {e}")
        return None

def tampilkan_ipc():
    st.title("Halaman IPC")
    st.write("Ini adalah tampilan khusus IPC.")
    
    # Menambahkan radio button untuk memilih opsi
    selected_option = st.radio(
        "Pilih jenis pengujian:",
        [
            "Kekerasan",
            "Keseragaman Bobot",
            "Tebal",
            "Waktu Hancur dan Friability"
        ],
        horizontal=True
    )
    
    # Menampilkan informasi template yang harus digunakan
    template_info = {
        "Kekerasan": "Template Excel untuk pengujian kekerasan",
        "Keseragaman Bobot": "Template Excel untuk pengujian keseragaman bobot",
        "Tebal": "Template Excel untuk pengujian tebal",
        "Waktu Hancur dan Friability": "Template Excel untuk pengujian waktu hancur dan friability"
    }
    
    st.info(f"Upload file Excel dengan format: {template_info[selected_option]}")

        # Fungsi untuk mengeksport DataFrame ke Excel
    def export_dataframe(df, filename="data_export"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        output.seek(0)
        b64 = base64.b64encode(output.read()).decode()
        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}.xlsx">📥 Download Excel File</a>'
        return href
        
    # File uploader tunggal
    uploaded_file = st.file_uploader("Upload file Excel sesuai template", type=["xlsx","ods"], key=f"uploader_{selected_option}")
    
    if uploaded_file:
        # Simpan salinan file untuk diproses (karena file uploader bisa digunakan sekali)
        file_copy = io.BytesIO(uploaded_file.getvalue())
        
        st.success(f"File untuk pengujian {selected_option} berhasil diupload")
        st.subheader(f"Hasil Pengujian {selected_option}")
        
        # Parsing file berdasarkan jenis pengujian yang dipilih
        if selected_option == "Kekerasan":
            df = parse_kekerasan_excel(file_copy)
            if df is not None:
                # Tambahkan visualisasi atau analisis tambahan khusus pengujian Kekerasan
                st.subheader("Visualisasi Data Kekerasan")
                if 'Nilai Kekerasan' in df.columns:
                    st.bar_chart(df['Nilai Kekerasan'])
        
        elif selected_option == "Keseragaman Bobot":
            df = parse_keseragaman_bobot_excel(file_copy)
            if df is not None:
                # Tambahkan visualisasi atau analisis tambahan khusus pengujian Keseragaman Bobot
                st.subheader("Visualisasi Data Keseragaman Bobot")
                if 'Bobot' in df.columns:
                    st.line_chart(df['Bobot'])
        
        elif selected_option == "Tebal":
            df = parse_tebal_excel(file_copy)
            if df is not None:
                # Tambahkan visualisasi atau analisis tambahan khusus pengujian Tebal
                st.subheader("Visualisasi Data Tebal")
                if 'Tebal' in df.columns:
                    st.bar_chart(df['Tebal'])
        
        elif selected_option == "Waktu Hancur dan Friability":
            df = parse_waktu_hancur_friability_excel(file_copy)
            if df is not None:
                # Tambahkan visualisasi atau analisis tambahan khusus pengujian Waktu Hancur dan Friability
                st.subheader("Visualisasi Data")
                col1, col2 = st.columns(2)
                with col1:
                    if 'Waktu Hancur' in df.columns:
                        st.subheader("Waktu Hancur")
                        st.bar_chart(df['Waktu Hancur'])
                with col2:
                    if 'Friability' in df.columns:
                        st.subheader("Friability")
                        st.bar_chart(df['Friability'])
