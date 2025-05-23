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
    try:
        df = pd.read_excel(file, header=None, engine='odf')

        # Validasi ukuran minimal file
        if df.shape[0] < 7 or df.shape[1] < 7:
            st.error("Template tidak sesuai: data minimal tidak terpenuhi.")
            return None

        # Ambil nama batch dari baris ke-2 (index 1), mulai dari kolom E (index 4)
        batch_names = df.iloc[1, 4:].dropna().values
        result_df = pd.DataFrame()

        for i, batch in enumerate(batch_names):
            try:
                # Kolom E hingga G (index 4, 5, 6), 5 data pertama (row index 2–6)
                data_1_5 = df.iloc[2:7, 4 + i*2]
                # Kolom F hingga H (index 5, 6, 7), 5 data kedua (row index 2–6)
                data_6_10 = df.iloc[2:7, 5 + i*2]

                values = pd.concat([data_1_5, data_6_10], ignore_index=True)
                values = pd.to_numeric(values, errors='coerce').dropna()

                if len(values) == 10:
                    result_df[batch] = values
                else:
                    st.warning(f"Data batch {batch} tidak lengkap. Diabaikan.")
            except Exception as e:
                st.warning(f"Batch {batch} tidak valid: {e}")
                continue

        if result_df.empty:
            st.error("Tidak ada data valid yang dapat diproses.")
            return None

        result_df.index = range(1, len(result_df) + 1)
        st.write("Data Keseragaman Bobot Terstruktur:")
        st.dataframe(result_df)

        stats_df = calculate_statistics(result_df)
        st.write("Statistik Data Kekerasan:")
        st.dataframe(stats_df.style.format("{:.4f}"))

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
        import pandas as pd
        import numpy as np
        import streamlit as st
        
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
                values_e = subset.iloc[0:3, 4].copy()  # Gunakan copy() untuk menghindari SettingWithCopyWarning
                values_f = subset.iloc[0:3, 5].copy()
                
                # Fungsi untuk membersihkan dan memproses data
                def clean_numeric_value(val):
                    if pd.isna(val):
                        return np.nan
                    if isinstance(val, (int, float)):
                        return float(val)
                    if isinstance(val, str):
                        # Coba pisahkan nilai jika string terlalu panjang dan tanpa spasi
                        if len(val) > 8 and ' ' not in val:
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
                        return np.nan
                
                # Apply cleaning function ke semua nilai
                values_e = values_e.apply(clean_numeric_value)
                values_f = values_f.apply(clean_numeric_value)
                
                stacked = pd.concat([values_e, values_f], ignore_index=True)
                
                # Pastikan semua nilai bisa dikonversi ke numeric
                stacked = pd.to_numeric(stacked, errors='coerce')
                
                # Hapus nilai NaN jika ada
                stacked = stacked.dropna()
                
                # Tambahkan ke result_df hanya jika ada data valid
                if not stacked.empty:
                    result_df[str(batch)] = stacked  # Konversi batch ke string untuk menghindari masalah tipe data
            
            except Exception as e:
                st.warning(f"Ada masalah saat memproses batch {batch}: {e}")
                # Lanjutkan ke batch berikutnya
                continue
            
        # Periksa apakah dataframe kosong
        if result_df.empty:
            st.error("Tidak ada data valid yang dapat diproses.")
            return None

        # Set index mulai dari 1
        result_df.index = range(1, len(result_df) + 1)

        # Hitung statistik
        stats = {}
        stats["MIN"] = result_df.min()
        stats["MAX"] = result_df.max()
        stats["MEAN"] = result_df.mean()
        stats["SD"] = result_df.std()
        stats["RSD (%)"] = (stats["SD"] / stats["MEAN"]) * 100
        
        # Buat DataFrame statistik
        stats_df = pd.DataFrame(stats)
        
        # Transpose agar statsnya menjadi baris
        stats_df = stats_df.T
        
        # Gabungkan kedua dataframe secara vertikal
        # Pertama, kita perlu memastikan bahwa stats_df memiliki indeks yang tidak bentrok
        max_index = result_df.index.max()
        stats_df.index = range(max_index + 1, max_index + 1 + len(stats_df))
        
        # Gabungkan dataframe data dengan dataframe statistik
        combined_df = pd.concat([result_df, stats_df])
        
        # Tambahkan kolom label di awal untuk memberikan nama pada baris statistik
        # Buat kolom label dengan NaN untuk baris data
        labels = pd.Series([""] * len(result_df), index=result_df.index)  # Gunakan string kosong daripada NaN
        
        # Tambahkan label untuk baris statistik
        for i, stat_name in enumerate(["MIN", "MAX", "MEAN", "SD", "RSD (%)"], start=max_index + 1):
            labels[i] = stat_name
            
        # Tambahkan kolom label ke dataframe gabungan
        combined_df.insert(0, "", labels)
        
        # Tampilkan dataframe gabungan
        st.write("Data Tebal Terstruktur dengan Statistik:")
        
        # Format angka dalam dataframe dengan 4 angka desimal
        # Gunakan fungsi format khusus untuk menangani nilai non-numerik
        def format_values(val):
            if isinstance(val, (int, float)) and not pd.isna(val):
                return f"{val:.4f}"
            return val
        
        # Gunakan applymap untuk menerapkan format ke seluruh dataframe
        formatted_df = combined_df.applymap(format_values)
        
        # Tampilkan dataframe yang sudah diformat
        st.dataframe(formatted_df)
        


    except Exception as e:
        st.error(f"Gagal memproses file Tebal: {e}")
        st.write("Detail error:", str(e))
        return None
    
    
def parse_waktu_hancur_friability_excel(file):
    """
    Parses Excel template for Waktu Hancur (Disintegration Time) and Friability testing data.
    
    This function reads an Excel file with pharmaceutical test data and creates two separate DataFrames
    for Waktu Hancur and Friability data, with both displayed in a consolidated format that includes
    batch data and statistics in a single table.
    
    Args:
        file: The uploaded Excel file object
        
    Returns:
        tuple: Two DataFrames (waktu_hancur_df, friability_df) for display
    """
    try:
        # Read the Excel file - automatically handle different formats
        df = pd.read_excel(file, header=None)
        
        # Debug information
        st.write("Excel file loaded. Shape:", df.shape)
        
        # First, try to find header row by looking for "Nomor Batch"
        header_row_idx = None
        batch_col = None
        value_col = None
        
        # Look for header row containing "Nomor Batch"
        for i in range(min(10, len(df))):  # Check first 10 rows
            row = df.iloc[i]
            for j, cell in enumerate(row):
                if isinstance(cell, str) and "Nomor Batch" in cell:
                    header_row_idx = i
                    batch_col = j
                    break
            if header_row_idx is not None:
                break
        
        # If not found by string matching, use the first row
        if header_row_idx is None:
            header_row_idx = 0
            # Try to guess batch column (usually first column)
            batch_col = 0
            
        # Look for "Sample Data" column
        if header_row_idx is not None:
            header_row = df.iloc[header_row_idx]
            for j, cell in enumerate(header_row):
                if isinstance(cell, str) and "Sample Data" in cell:
                    value_col = j
                    break
        
        # If still not found, try column E (index 4) which is common for Sample Data
        if value_col is None:
            value_col = 4  # Assume column E (index 4) contains values
            
        # Debug info
        st.write(f"Using header row: {header_row_idx}, Batch column: {batch_col}, Value column: {value_col}")
        
        # Get data rows (skip header)
        data_df = df.iloc[header_row_idx+1:].copy()
        
        # Filter out rows with no batch numbers
        data_df = data_df[~data_df.iloc[:, batch_col].isna()]
        
        # Initialize dictionaries to store batch numbers and their values
        friability_data = {}
        waktu_hancur_data = {}
        
        # Lists to collect all values for overall statistics
        all_friability_values = []
        all_waktu_hancur_values = []
        
        # Process each row
        for _, row in data_df.iterrows():
            batch = row.iloc[batch_col]
            value = row.iloc[value_col] if not pd.isna(row.iloc[value_col]) else None
            
            # Skip rows with no value
            if value is None:
                continue
                
            try:
                # Convert to float for comparison
                value_float = float(value)
                
                # Sort based on Sample Data value - friability is typically less than 2.5%
                if value_float < 2.5:
                    # Add to batch-specific data
                    if str(batch) not in friability_data:
                        friability_data[str(batch)] = value_float
                    
                    # Add to all values for overall statistics
                    all_friability_values.append(value_float)
                else:
                    # Add to batch-specific data
                    if str(batch) not in waktu_hancur_data:
                        waktu_hancur_data[str(batch)] = value_float
                    
                    # Add to all values for overall statistics
                    all_waktu_hancur_values.append(value_float)
            except (ValueError, TypeError):
                # If value can't be converted to float, skip this row
                st.warning(f"Skipping row with batch {batch}, invalid value: {value}")
                continue
        
        # Calculate overall statistics for Waktu Hancur
        waktu_hancur_min = min(all_waktu_hancur_values) if all_waktu_hancur_values else None
        waktu_hancur_max = max(all_waktu_hancur_values) if all_waktu_hancur_values else None
        waktu_hancur_mean = sum(all_waktu_hancur_values)/len(all_waktu_hancur_values) if all_waktu_hancur_values else None
        
        # Calculate SD and RSD only if there are multiple values to compare
        if len(all_waktu_hancur_values) > 1:
            waktu_hancur_sd = np.std(all_waktu_hancur_values, ddof=1)
            waktu_hancur_rsd = (waktu_hancur_sd / waktu_hancur_mean * 100) if waktu_hancur_mean else None
        else:
            waktu_hancur_sd = None
            waktu_hancur_rsd = None
        
        # Calculate overall statistics for Friability
        friability_min = min(all_friability_values) if all_friability_values else None
        friability_max = max(all_friability_values) if all_friability_values else None
        friability_mean = sum(all_friability_values)/len(all_friability_values) if all_friability_values else None
        
        # Calculate SD and RSD only if there are multiple values to compare
        if len(all_friability_values) > 1:
            friability_sd = np.std(all_friability_values, ddof=1)
            friability_rsd = (friability_sd / friability_mean * 100) if friability_mean else None
        else:
            friability_sd = None
            friability_rsd = None
        
        # Create consolidated friability DataFrame
        friability_df = pd.DataFrame()
        if friability_data:
            # Create DataFrame with batch data
            batch_df = pd.DataFrame({
                "Batch": list(friability_data.keys()),
                "Friability": list(friability_data.values())
            })
            
            # Sort by Batch for consistent display
            batch_df = batch_df.sort_values("Batch").reset_index(drop=True)
            
            # Create statistics DataFrame
            stats_df = pd.DataFrame({
                "Batch": ["Minimum", "Maximum", "Rata-rata", "Standar Deviasi", "RSD (%)"],
                "Friability": [friability_min, friability_max, friability_mean, friability_sd, friability_rsd]
            })
            
            # Combine batch data with statistics at the bottom
            friability_df = pd.concat([batch_df, stats_df], ignore_index=True)
        
        # Create consolidated waktu hancur DataFrame
        waktu_hancur_df = pd.DataFrame()
        if waktu_hancur_data:
            # Create DataFrame with batch data
            batch_df = pd.DataFrame({
                "Batch": list(waktu_hancur_data.keys()),
                "Waktu Hancur": list(waktu_hancur_data.values())
            })
            
            # Sort by Batch for consistent display
            batch_df = batch_df.sort_values("Batch").reset_index(drop=True)
            
            # Create statistics DataFrame
            stats_df = pd.DataFrame({
                "Batch": ["Minimum", "Maximum", "Rata-rata", "Standar Deviasi", "RSD (%)"],
                "Waktu Hancur": [waktu_hancur_min, waktu_hancur_max, waktu_hancur_mean, waktu_hancur_sd, waktu_hancur_rsd]
            })
            
            # Combine batch data with statistics at the bottom
            waktu_hancur_df = pd.concat([batch_df, stats_df], ignore_index=True)
        
        # Display the consolidated tables
        st.write("Tabel Waktu Hancur dengan Statistik:")
        st.dataframe(waktu_hancur_df.style.format({
            "Waktu Hancur": "{:.4f}"
        }))
        
        st.write("Tabel Friability dengan Statistik:")
        st.dataframe(friability_df.style.format({
            "Friability": "{:.4f}"
        }))
        
        # Return both DataFrames
        return waktu_hancur_df, friability_df
    
    except Exception as e:
        st.error(f"Gagal memproses file Waktu Hancur dan Friability: {e}")
        st.exception(e)
        # Return empty DataFrames
        return pd.DataFrame(), pd.DataFrame()

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
                
                # Tampilkan tombol download
                filename = "data_kekerasan"
                st.markdown(export_dataframe(df, filename), unsafe_allow_html=True)
                st.success("Data Kekerasan siap diunduh. Klik tombol di atas untuk mengunduh file Excel.")
                
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
                
                # Tampilkan tombol download
                filename = "data_keseragaman_bobot"
                st.markdown(export_dataframe(df, filename), unsafe_allow_html=True)
                st.success("Data Keseragaman Bobot siap diunduh. Klik tombol di atas untuk mengunduh file Excel.")
                
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
                
                # Tampilkan tombol download
                filename = "data_tebal"
                st.markdown(export_dataframe(df, filename), unsafe_allow_html=True)
                st.success("Data Tebal siap diunduh. Klik tombol di atas untuk mengunduh file Excel.")
        
        elif selected_option == "Waktu Hancur dan Friability":
            # Fixed: Unpacking the tuple returned by parse_waktu_hancur_friability_excel
            waktu_hancur_df, friability_df = parse_waktu_hancur_friability_excel(file_copy)
            

                
                # Tampilkan tombol download untuk masing-masing dataframe
            if not waktu_hancur_df.empty:
                    st.markdown(export_dataframe(waktu_hancur_df, "data_waktu_hancur"), unsafe_allow_html=True)
                    st.success("Data Waktu Hancur siap diunduh. Klik tombol di atas untuk mengunduh file Excel.")
                
            if not friability_df.empty:
                    st.markdown(export_dataframe(friability_df, "data_friability"), unsafe_allow_html=True)
                    st.success("Data Friability siap diunduh. Klik tombol di atas untuk mengunduh file Excel.")

        # Tambahkan opsi untuk menampilkan semua data dalam bentuk tabel
        # (Kecuali untuk Waktu Hancur dan Friability yang sudah ditampilkan)
        if selected_option != "Waktu Hancur dan Friability" and 'df' in locals() and df is not None:
            if st.checkbox("Tampilkan semua data dalam bentuk tabel"):
                st.write("Data Lengkap (termasuk statistik):")
                st.dataframe(df)
