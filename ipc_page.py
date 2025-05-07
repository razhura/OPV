# ipc_page.py

import streamlit as st
import pandas as pd
import numpy as np
import io
import base64

def calculate_statistics(df):
    """
    Calculate statistics (MIN, MAX, MEAN, SD, RSD) for each batch in the dataframe
    """
    # Create a new dataframe for statistics
    stats_df = pd.DataFrame()
    
    # For each column (batch) in the dataframe
    for column in df.columns:
        # Calculate statistics
        min_val = df[column].min()
        max_val = df[column].max()
        mean_val = df[column].mean()
        std_val = df[column].std()
        # Calculate RSD (Relative Standard Deviation) as (SD/Mean) * 100
        rsd_val = (std_val / mean_val * 100) if mean_val != 0 else 0
        
        # Add to statistics dataframe
        stats_df[column] = [min_val, max_val, mean_val, std_val, rsd_val]
    
    # Set index names
    stats_df.index = ['MIN', 'MAX', 'MEAN', 'SD', 'RSD (%)']
    
    return stats_df

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
            try:
                values_e = subset.iloc[0:5, 4]
                values_f = subset.iloc[0:5, 5]
                
                # Proses dan clean data untuk mengatasi masalah konversi string ke numeric
                def clean_value(val):
                    if isinstance(val, str):
                        # Cek apakah string memiliki format yang salah (angka-angka tanpa spasi)
                        if len(val) > 8 and not ' ' in val:
                            # Coba pisahkan angka berdasarkan pola
                            # Asumsikan angka dengan pola 2 digit
                            parts = []
                            i = 0
                            while i < len(val):
                                # Cek karakter berikutnya untuk melihat pola angka
                                if i < len(val) - 1:
                                    if val[i] == '1' and val[i+1] in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']:
                                        # Kemungkinan angka belasan
                                        parts.append(val[i:i+2])
                                        i += 2
                                    else:
                                        # Angka satuan atau puluhan
                                        parts.append(val[i])
                                        i += 1
                                else:
                                    parts.append(val[i])
                                    i += 1
                            
                            # Ambil nilai pertama saja untuk kesederhanaan
                            # Gunakan nilai tengah jika ada banyak angka
                            if len(parts) > 0:
                                middle_index = len(parts) // 2
                                return float(parts[middle_index])
                            else:
                                return np.nan
                        else:
                            try:
                                return float(val)
                            except:
                                return np.nan
                    else:
                        return val
                
                # Apply cleaning function
                values_e = values_e.apply(clean_value)
                values_f = values_f.apply(clean_value)
                
                stacked = pd.concat([values_e, values_f], ignore_index=True)
                
                # Pastikan semua nilai bisa dikonversi ke numeric
                stacked = pd.to_numeric(stacked, errors='coerce')
                
                # Hapus nilai NaN jika ada
                stacked = stacked.dropna()
                
                result_df[batch] = stacked
                
            except Exception as e:
                st.warning(f"Ada masalah saat memproses batch {batch}: {e}")
                # Lanjutkan ke batch berikutnya jika ada masalah
                continue
            
        # Periksa apakah dataframe kosong
        if result_df.empty:
            st.error("Tidak ada data valid yang dapat diproses.")
            return None
            
        # Set index mulai dari 1 (agar tidak bentrok saat loop)    
        result_df.index = range(1, len(result_df) + 1)

        # Tampilkan dataframe hasil
        st.write("Data Kekerasan Terstruktur:")
        st.dataframe(result_df)
        
        # Hitung dan tampilkan statistik
        stats_df = calculate_statistics(result_df)
        st.write("Statistik Data Kekerasan:")
        st.dataframe(stats_df.style.format("{:.4f}"))
        
        # Gabungkan dataframe untuk ekspor
        export_df = pd.concat([result_df, stats_df])
        
        return export_df

    except Exception as e:
        st.error(f"Gagal memproses file Kekerasan: {e}")
        st.write("Detail error:", str(e))
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

            try:
                # Ambil data dari kolom E (data 1-5), F (data 6-10), G (data 11-15), H (data 16-20)
                # Asumsi kolom ke-5, ke-6, ke-7, ke-8 (indeks 4, 5, 6, 7)
                values_e = subset.iloc[0:5, 4]
                values_f = subset.iloc[0:5, 5]
                values_g = subset.iloc[0:5, 6]
                values_h = subset.iloc[0:5, 7]
                
                # Fungsi untuk membersihkan dan memproses data
                def clean_numeric_value(val):
                    if isinstance(val, str):
                        # Coba pisahkan nilai jika string terlalu panjang dan tanpa spasi
                        if len(val) > 8 and not ' ' in val:
                            # Coba ekstrak angka yang valid
                            import re
                            # Ekstrak pola angka dengan desimal
                            numbers = re.findall(r'\d+\.\d+|\d+', val)
                            if numbers:
                                # Ambil nilai tengah jika ada beberapa angka
                                middle_index = len(numbers) // 2
                                return float(numbers[middle_index])
                            else:
                                return np.nan
                        else:
                            try:
                                return float(val)
                            except:
                                return np.nan
                    else:
                        return val
                
                # Apply cleaning function ke semua nilai
                values_e = values_e.apply(clean_numeric_value)
                values_f = values_f.apply(clean_numeric_value)
                values_g = values_g.apply(clean_numeric_value)
                values_h = values_h.apply(clean_numeric_value)
                
                stacked = pd.concat([values_e, values_f, values_g, values_h], ignore_index=True)
                
                # Pastikan semua nilai bisa dikonversi ke numeric
                stacked = pd.to_numeric(stacked, errors='coerce')
                
                # Hapus nilai NaN jika ada
                stacked = stacked.dropna()
                
                result_df[batch] = stacked
                
            except Exception as e:
                st.warning(f"Ada masalah saat memproses batch {batch}: {e}")
                # Lanjutkan ke batch berikutnya
                continue
            
        # Periksa apakah dataframe kosong
        if result_df.empty:
            st.error("Tidak ada data valid yang dapat diproses.")
            return None
            
        # Set index mulai dari 1 (agar tidak bentrok saat loop)    
        result_df.index = range(1, len(result_df) + 1)

        # Tampilkan dataframe hasil
        st.write("Data Keseragaman Bobot Terstruktur:")
        st.dataframe(result_df)
        
        # Hitung dan tampilkan statistik
        stats_df = calculate_statistics(result_df)
        st.write("Statistik Data Keseragaman Bobot:")
        st.dataframe(stats_df.style.format("{:.4f}"))
        
        # Gabungkan dataframe untuk ekspor
        export_df = pd.concat([result_df, stats_df])
        
        return export_df

    except Exception as e:
        st.error(f"Gagal memproses file Keseragaman Bobot: {e}")
        st.write("Detail error:", str(e))
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

            try:
                # Ambil data dari kolom E (data 1-3) dan F (data 4-6)
                # Kolom ke-5 dan ke-6 (indeks 4 dan 5)
                values_e = subset.iloc[0:3, 4]
                values_f = subset.iloc[0:3, 5]
                
                # Fungsi untuk membersihkan dan memproses data
                def clean_numeric_value(val):
                    if isinstance(val, str):
                        # Coba pisahkan nilai jika string terlalu panjang dan tanpa spasi
                        if len(val) > 8 and not ' ' in val:
                            # Ekstrak angka-angka dari string menggunakan regex
                            import re
                            numbers = re.findall(r'\d+\.\d+|\d+', val)
                            if numbers:
                                # Ambil nilai tengah jika ada beberapa angka
                                middle_index = len(numbers) // 2
                                return float(numbers[middle_index])
                            else:
                                return np.nan
                        else:
                            try:
                                return float(val)
                            except:
                                return np.nan
                    else:
                        return val
                
                # Apply cleaning function ke semua nilai
                values_e = values_e.apply(clean_numeric_value)
                values_f = values_f.apply(clean_numeric_value)
                
                stacked = pd.concat([values_e, values_f], ignore_index=True)
                
                # Pastikan semua nilai bisa dikonversi ke numeric
                stacked = pd.to_numeric(stacked, errors='coerce')
                
                # Hapus nilai NaN jika ada
                stacked = stacked.dropna()
                
                result_df[batch] = stacked
                
            except Exception as e:
                st.warning(f"Ada masalah saat memproses batch {batch}: {e}")
                # Lanjutkan ke batch berikutnya
                continue
            
        # Periksa apakah dataframe kosong
        if result_df.empty:
            st.error("Tidak ada data valid yang dapat diproses.")
            return None

        # Set index mulai dari 1 (agar tidak bentrok saat loop)
        result_df.index = range(1, len(result_df) + 1)

        # Tampilkan dataframe hasil
        st.write("Data Tebal Terstruktur:")
        st.dataframe(result_df)
        
        # Hitung dan tampilkan statistik
        stats_df = calculate_statistics(result_df)
        st.write("Statistik Data Tebal:")
        st.dataframe(stats_df.style.format("{:.4f}"))
        
        # Gabungkan dataframe untuk ekspor
        export_df = pd.concat([result_df, stats_df])
        
        return export_df

    except Exception as e:
        st.error(f"Gagal memproses file Tebal: {e}")
        st.write("Detail error:", str(e))
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

            try:
                # Ambil 2 nilai dari kolom E (baris 0 dan 1 relatif terhadap subset)
                values = subset.iloc[0:2, 4].tolist()
                
                # Fungsi untuk membersihkan dan memproses data
                def clean_numeric_value(val):
                    if isinstance(val, str):
                        # Coba pisahkan nilai jika string terlalu panjang dan tanpa spasi
                        if len(val) > 8 and not ' ' in val:
                            # Ekstrak angka-angka dari string menggunakan regex
                            import re
                            numbers = re.findall(r'\d+\.\d+|\d+', val)
                            if numbers:
                                # Ambil nilai tengah jika ada beberapa angka
                                middle_index = len(numbers) // 2
                                return float(numbers[middle_index])
                            else:
                                return np.nan
                        else:
                            try:
                                return float(val)
                            except:
                                return np.nan
                    else:
                        return val
                
                # Apply cleaning function ke semua nilai
                cleaned_values = [clean_numeric_value(val) for val in values]
                
                # Gabungkan batch + data
                row = [batch] + cleaned_values
                result_rows.append(row)
                
            except Exception as e:
                st.warning(f"Ada masalah saat memproses batch {batch}: {e}")
                # Lanjutkan ke batch berikutnya
                continue
            
        # Periksa apakah ada data yang berhasil diproses
        if not result_rows:
            st.error("Tidak ada data valid yang dapat diproses.")
            return None
            
        # Buat dataframe hasil
        columns = ["Batch"] + parameter_labels
        result_df = pd.DataFrame(result_rows, columns=columns)
        
        # Pastikan kolom numerik dikonversi dengan benar
        for param in parameter_labels:
            result_df[param] = pd.to_numeric(result_df[param], errors='coerce')

        # Tampilkan dataframe hasil
        st.write("Data Waktu Hancur dan Friability Terstruktur:")
        st.dataframe(result_df)
        
        # Untuk waktu hancur dan friability, kita perlu menghitung statistik untuk masing-masing parameter
        # Karena formatnya berbeda (data sudah dalam format lebar, bukan panjang)
        stats = {}
        
        # Loop melalui setiap parameter (kecuali kolom Batch)
        for param in parameter_labels:
            try:
                param_data = result_df[param]
                
                # Hitung statistik
                stats[f"{param}_MIN"] = param_data.min()
                stats[f"{param}_MAX"] = param_data.max()
                stats[f"{param}_MEAN"] = param_data.mean()
                stats[f"{param}_SD"] = param_data.std()
                # Calculate RSD
                stats[f"{param}_RSD (%)"] = (param_data.std() / param_data.mean() * 100) if param_data.mean() != 0 else 0
            except Exception as e:
                st.warning(f"Tidak dapat menghitung statistik untuk parameter {param}: {e}")
        
        # Buat dataframe statistik
        stats_df = pd.DataFrame([stats])
        
        st.write("Statistik Data Waktu Hancur dan Friability:")
        st.dataframe(stats_df.style.format("{:.4f}"))
        
        # Untuk ekspor, gabungkan kedua dataframe dengan cara yang berbeda
        # Ubah format stats_df agar cocok untuk digabungkan
        export_df = pd.concat([result_df, stats_df.T.reset_index()], axis=1)
        
        return export_df

    except Exception as e:
        st.error(f"Gagal memproses file Waktu Hancur dan Friability: {e}")
        st.write("Detail error:", str(e))
        return None

# Fungsi untuk visualisasi box plot
def create_boxplot(df):
    """
    Membuat box plot untuk dataframe yang diberikan
    """
    # Melting dataframe untuk format yang sesuai dengan box plot
    melted_df = df.melt(var_name='Batch', value_name='Nilai')
    
    # Plot box plot
    fig = {
        'data': [
            {
                'type': 'box',
                'y': melted_df[melted_df['Batch'] == batch]['Nilai'],
                'name': batch
            } for batch in df.columns
        ],
        'layout': {
            'title': 'Box Plot Data per Batch',
            'yaxis': {'title': 'Nilai'},
            'boxmode': 'group'
        }
    }
    
    return fig

# Fungsi untuk mengeksport DataFrame ke Excel
def export_dataframe(df, filename="data_export"):
    """
    Fungsi untuk mengekspor DataFrame ke file Excel yang dapat diunduh
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=True)
    output.seek(0)
    b64 = base64.b64encode(output.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}.xlsx">📥 Download Excel File</a>'
    return href

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
                # Data untuk visualisasi (hapus baris statistik)
                viz_df = df.iloc[:10] if len(df) > 10 else df.copy()
                
                # Tambahkan visualisasi
                st.subheader("Visualisasi Data Kekerasan")
                # Bar chart untuk rata-rata per batch
                st.write("Rata-rata Kekerasan per Batch:")
                means = df.loc["MEAN"] if "MEAN" in df.index else df.mean()
                st.bar_chart(means)
                
                # Box plot untuk distribusi data
                st.write("Box Plot Kekerasan per Batch:")
                st.plotly_chart(create_boxplot(viz_df))
                
        elif selected_option == "Keseragaman Bobot":
            df = parse_keseragaman_bobot_excel(file_copy)
            if df is not None:
                # Data untuk visualisasi (hapus baris statistik)
                viz_df = df.iloc[:20] if len(df) > 20 else df.copy()
                
                # Tambahkan visualisasi
                st.subheader("Visualisasi Data Keseragaman Bobot")
                # Line chart untuk trend data
                st.write("Trend Keseragaman Bobot:")
                st.line_chart(viz_df)
                
                # Box plot untuk distribusi data
                st.write("Box Plot Keseragaman Bobot per Batch:")
                st.plotly_chart(create_boxplot(viz_df))
                
        elif selected_option == "Tebal":
            df = parse_tebal_excel(file_copy)
            if df is not None:
                # Data untuk visualisasi (hapus baris statistik)
                viz_df = df.iloc[:6] if len(df) > 6 else df.copy()
                
                # Tambahkan visualisasi
                st.subheader("Visualisasi Data Tebal")
                # Bar chart untuk rata-rata per batch
                st.write("Rata-rata Tebal per Batch:")
                means = df.loc["MEAN"] if "MEAN" in df.index else df.mean()
                st.bar_chart(means)
                
                # Box plot untuk distribusi data
                st.write("Box Plot Tebal per Batch:")
                st.plotly_chart(create_boxplot(viz_df))
        
        elif selected_option == "Waktu Hancur dan Friability":
            df = parse_waktu_hancur_friability_excel(file_copy)
            if df is not None:
                # Tambahkan visualisasi
                st.subheader("Visualisasi Data")
                
                # Karena struktur data berbeda, kita perlu pendekatan visualisasi yang berbeda
                if "Waktu Hancur" in df.columns and "Friability" in df.columns:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.write("Waktu Hancur per Batch:")
                        # Filter hanya baris data (bukan statistik)
                        data_df = df[~df["Batch"].astype(str).str.contains("_", na=False)]
                        st.bar_chart(data_df.set_index("Batch")["Waktu Hancur"])
                    
                    with col2:
                        st.write("Friability per Batch:")
                        st.bar_chart(data_df.set_index("Batch")["Friability"])

        # Tampilkan tombol download jika data berhasil diproses
        if df is not None:
            filename = f"data_{selected_option.lower().replace(' ', '_')}"
            st.markdown(export_dataframe(df, filename), unsafe_allow_html=True)
            st.success(f"Data {selected_option} siap diunduh. Klik tombol di atas untuk mengunduh file Excel.")
            
            # Tambahkan opsi untuk menampilkan semua data dalam bentuk tabel
            if st.checkbox("Tampilkan semua data dalam bentuk tabel"):
                st.write("Data Lengkap (termasuk statistik):")
                st.dataframe(df)
