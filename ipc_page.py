import streamlit as st
import pandas as pd
import numpy as np
import io
import re

# --- Fungsi Helper ---
def _clean_numeric_value_helper(val):
    """
    Membersihkan dan mengkonversi nilai menjadi float.
    Menangani string angka yang mungkin tergabung atau mengandung karakter non-numerik.
    """
    if isinstance(val, str):
        # Mencoba mengekstrak angka dari string yang mungkin tergabung (heuristik)
        if len(val) > 7 and '.' in val and val.count('.') > 1 and ' ' not in val:
            numbers = re.findall(r'-?\d+\.?\d*', val) # Mencari semua bagian yang mirip angka
            if numbers:
                try:
                    return float(numbers[0]) # Mengambil angka pertama yang ditemukan
                except ValueError:
                    pass # Jika gagal, lanjutkan ke upaya konversi standar
        try:
            return float(val)
        except ValueError:
            return np.nan
    elif isinstance(val, (int, float)):
        return float(val)
    return np.nan

def calculate_statistics(df):
    """
    Menghitung statistik (MIN, MAX, MEAN, SD, RSD) untuk setiap batch dalam dataframe.
    """
    stats_df = pd.DataFrame()
    for column in df.columns:
        numeric_data = pd.to_numeric(df[column], errors='coerce').dropna()
        if numeric_data.empty:
            min_val, max_val, mean_val, std_val, rsd_val = np.nan, np.nan, np.nan, np.nan, np.nan
        else:
            min_val = numeric_data.min()
            max_val = numeric_data.max()
            mean_val = numeric_data.mean()
            std_val = numeric_data.std(ddof=1) # ddof=1 untuk sample standard deviation
            rsd_val = (std_val / mean_val * 100) if mean_val != 0 else np.nan # Hindari ZeroDivisionError dan hasil 0 jika mean 0
        
        stats_df[column] = [min_val, max_val, mean_val, std_val, rsd_val]
    stats_df.index = ['MIN', 'MAX', 'MEAN', 'SD', 'RSD (%)']
    return stats_df

# --- Fungsi Parsing untuk Setiap Jenis Pengujian ---

def parse_kekerasan_excel(file):
    try:
        # Hapus engine='odf' agar pandas otomatis mendeteksi untuk xlsx dan ods
        df = pd.read_excel(file, header=None) 
        if df.shape[0] < 10 or df.shape[1] < 6: # Validasi dasar ukuran template
            st.error("Template tidak sesuai: data minimal tidak terpenuhi (Kekerasan).")
            return None
        
        result_df = pd.DataFrame()
        current_row = 2 # Data dimulai dari baris ke-3 (indeks 2)

        while current_row < df.shape[0] - 7: # Memastikan cukup baris untuk satu blok batch
            try:
                batch_name_val = df.iloc[current_row, 0] 
                if pd.isna(batch_name_val) or str(batch_name_val).strip() == '':
                    break 
                
                batch_name = str(batch_name_val)

                # Data dari Kolom E (indeks 4) dan F (indeks 5)
                data_1_5 = df.iloc[current_row:current_row+5, 4] if df.shape[1] > 4 else pd.Series(dtype='float64')
                data_6_10 = df.iloc[current_row:current_row+5, 5] if df.shape[1] > 5 else pd.Series(dtype='float64')
                
                values = pd.concat([data_1_5, data_6_10], ignore_index=True)
                values = pd.to_numeric(values, errors='coerce').dropna()
                
                if len(values) >= 8: # Minimal 8 data poin valid
                    values = values[:10] 
                    if len(values) < 10:
                        values = pd.concat([values, pd.Series([np.nan] * (10 - len(values)))], ignore_index=True)
                    result_df[batch_name] = values
                else:
                    st.warning(f"Data batch {batch_name} tidak lengkap ({len(values)} data valid). Diabaikan.")
                
                current_row += 8 
            except IndexError: 
                st.warning(f"Error memproses struktur data pada baris sekitar {current_row}. Periksa format file.")
                current_row += 8 
                continue
            except Exception as e: 
                st.warning(f"Error memproses batch pada baris {current_row}: {e}")
                current_row += 8
                continue
        
        if result_df.empty:
            st.error("Tidak ada data valid yang dapat diproses dari file Kekerasan.")
            return None
        
        max_data_points = 10
        result_df.index = range(1, max_data_points + 1)
        
        stats_df = calculate_statistics(result_df)
        final_df = pd.concat([result_df, stats_df])
        
        st.write("Data Kekerasan dengan Statistik:")
        
        # Format tampilan yang lebih aman
        display_df_styled = final_df.style.format(
            lambda x: f"{x:.2f}" if isinstance(x, (int, float)) and pd.notna(x) else str(x) if pd.notna(x) else "",
            na_rep=""
        )
        st.dataframe(display_df_styled)
        return final_df
    except Exception as e:
        st.error(f"Gagal memproses file Kekerasan: {e}")
        st.exception(e)
        return None

def parse_keseragaman_bobot_excel(file):
    try:
        df = pd.read_excel(file, header=None)
        if df.empty:
            st.error("File Keseragaman Bobot kosong.")
            return None

        header_row_index = 0 
        header_row = df.iloc[header_row_index]
        df_data = df[header_row_index+1:].copy()
        df_data.columns = header_row
        df_data.reset_index(drop=True, inplace=True)

        batch_column_name = df_data.columns[0]
        
        df_data = df_data[~df_data.iloc[:, 0].astype(str).str.contains("Rata|SD|RSD", na=False, case=False)]
        
        batch_series = df_data[batch_column_name].dropna().unique()
        result_df = pd.DataFrame()

        data_cols_indices = [4, 5, 6, 7] # Kolom E, F, G, H

        for batch_val in batch_series:
            batch = str(batch_val)
            subset = df_data[df_data[batch_column_name] == batch_val]
            if subset.empty:
                continue
            
            all_batch_values = []
            for col_idx in data_cols_indices:
                if col_idx < subset.shape[1]: 
                    col_values = subset.iloc[0:5, col_idx] # Ambil 5 baris pertama dari setiap kolom data
                    all_batch_values.append(col_values)
            
            if not all_batch_values:
                st.warning(f"Tidak ada kolom data (E,F,G,H) ditemukan untuk batch {batch}.")
                continue

            stacked_values = pd.concat(all_batch_values, ignore_index=True)
            # Gunakan helper global
            cleaned_values = stacked_values.apply(_clean_numeric_value_helper).dropna() 
            
            if not cleaned_values.empty:
                target_length = 20
                if len(cleaned_values) < target_length:
                    padding = pd.Series([np.nan] * (target_length - len(cleaned_values)))
                    cleaned_values = pd.concat([cleaned_values, padding], ignore_index=True)
                result_df[batch] = cleaned_values[:target_length]
            else:
                st.warning(f"Tidak ada data numerik valid yang ditemukan untuk batch {batch} setelah pembersihan.")

        if result_df.empty:
            st.error("Tidak ada data Keseragaman Bobot valid yang dapat diproses.")
            return None
        
        result_df.index = range(1, len(result_df) + 1)

        st.write("Data Keseragaman Bobot Terstruktur:")
        # Format dengan presisi 4 desimal untuk data, string untuk lainnya jika ada
        st.dataframe(result_df.style.format(lambda x: f"{x:.4f}" if isinstance(x, (int, float)) else str(x), na_rep=""))
        
        stats_df = calculate_statistics(result_df)
        st.write("Statistik Data Keseragaman Bobot:")
        st.dataframe(stats_df.style.format(lambda x: f"{x:.4f}" if isinstance(x, (int, float)) else str(x), na_rep=""))
        
        export_df = pd.concat([result_df, stats_df])
        return export_df
    except Exception as e:
        st.error(f"Gagal memproses file Keseragaman Bobot: {e}")
        st.exception(e)
        return None

def parse_tebal_excel(file):
    try:
        df = pd.read_excel(file, header=None)
        if df.empty:
            st.error("File Tebal kosong.")
            return None

        header_row_index = 0
        header_row = df.iloc[header_row_index]
        df_data = df[header_row_index+1:].copy()
        df_data.columns = header_row
        df_data.reset_index(drop=True, inplace=True)
        
        batch_column_name = df_data.columns[0] 
        batch_series = df_data[batch_column_name].dropna().unique()
        result_df = pd.DataFrame()

        data_cols_indices = [4, 5] # Kolom E, F
        num_values_per_col = 3

        for batch_val in batch_series:
            batch = str(batch_val)
            subset = df_data[df_data[batch_column_name] == batch_val]
            if subset.empty:
                continue

            all_batch_values = []
            for col_idx in data_cols_indices:
                if col_idx < subset.shape[1]: 
                    col_values = subset.iloc[0:num_values_per_col, col_idx]
                    all_batch_values.append(col_values)
            
            if not all_batch_values:
                st.warning(f"Tidak ada kolom data (E,F) ditemukan untuk batch {batch}.")
                continue
            
            stacked_values = pd.concat(all_batch_values, ignore_index=True)
            # Gunakan helper global
            cleaned_values = stacked_values.apply(_clean_numeric_value_helper).dropna()

            if not cleaned_values.empty:
                target_length = 6 
                if len(cleaned_values) < target_length:
                    padding = pd.Series([np.nan] * (target_length - len(cleaned_values)))
                    cleaned_values = pd.concat([cleaned_values, padding], ignore_index=True)
                result_df[batch] = cleaned_values[:target_length]
            else:
                st.warning(f"Tidak ada data numerik valid yang ditemukan untuk batch tebal {batch} setelah pembersihan.")

        if result_df.empty:
            st.error("Tidak ada data Tebal valid yang dapat diproses.")
            return None
        
        result_df.index = range(1, len(result_df) + 1)
        
        stats_df = calculate_statistics(result_df)
        exportable_df = pd.concat([result_df, stats_df]) 

        # Untuk display, tambahkan label statistik sebagai kolom pertama
        display_df = exportable_df.copy()
        # Membuat kolom "Keterangan" dari index `exportable_df`
        # Baris data akan memiliki nomor (1, 2, ...), baris stat akan memiliki nama stat ('MIN', 'MAX', ...)
        keterangan_values = [str(idx) for idx in display_df.index]
        display_df.insert(0, "Keterangan", keterangan_values)
        display_df.reset_index(drop=True, inplace=True) # Reset index untuk tampilan bersih

        st.write("Data Tebal Terstruktur dengan Statistik:")
        
        # *** INI BAGIAN YANG DIPERBAIKI ***
        numeric_cols_display = [col for col in display_df.columns if col != "Keterangan"]
        formatter_display = {col: "{:.4f}" for col in numeric_cols_display}
        formatter_display["Keterangan"] = "{}" # Format kolom "Keterangan" sebagai string

        st.dataframe(
            display_df.style.format(formatter_display, na_rep="")
            .set_properties(**{'text-align': 'left'})
            .set_table_styles([dict(selector='th', props=[('text-align', 'left')])])
        )
        
        return exportable_df

    except Exception as e:
        st.error(f"Gagal memproses file Tebal: {e}")
        st.exception(e)
        return None

def parse_waktu_hancur_friability_excel(file):
    try:
        df = pd.read_excel(file, header=None)
        if df.empty:
            st.error("File Waktu Hancur/Friability kosong.")
            return pd.DataFrame(), pd.DataFrame()
            
        header_row_idx = None
        batch_col_idx = None 
        value_col_idx = None 
        
        for i in range(min(10, len(df))): 
            row_values = df.iloc[i].astype(str) 
            if "Nomor Batch" in row_values.values:
                header_row_idx = i
                batch_col_idx = row_values[row_values.str.contains("Nomor Batch", case=False, na=False)].index[0]
                if "Sample Data" in row_values.values:
                        value_col_idx = row_values[row_values.str.contains("Sample Data", case=False, na=False)].index[0]
                break
        
        if header_row_idx is None: 
            st.warning("Header 'Nomor Batch' tidak ditemukan, menggunakan asumsi default (Baris 1 header, Kolom A Batch).")
            header_row_idx = 0
            batch_col_idx = 0 
        
        if value_col_idx is None: 
            potential_value_col = 4 
            if df.shape[1] > potential_value_col:
                st.warning(f"Kolom 'Sample Data' tidak ditemukan, mengasumsikan data ada di kolom E (indeks {potential_value_col}).")
                value_col_idx = potential_value_col
            else: 
                st.error("Tidak dapat menemukan kolom data ('Sample Data' atau default). Harap periksa template.")
                return pd.DataFrame(), pd.DataFrame()

        df.columns = df.iloc[header_row_idx]
        data_df = df.iloc[header_row_idx+1:].copy()
        
        batch_col_name = data_df.columns[batch_col_idx]
        value_col_name = data_df.columns[value_col_idx]

        data_df = data_df[~data_df[batch_col_name].isna()]
        
        friability_data_dict = {} 
        waktu_hancur_data_dict = {} 
        
        all_friability_values = []
        all_waktu_hancur_values = []
        
        for _, row in data_df.iterrows():
            batch = str(row[batch_col_name])
            value_raw = row[value_col_name]
            
            if pd.isna(value_raw):
                continue
            
            try:
                value_float = float(value_raw)
                # Heuristik pemisahan Friability dan Waktu Hancur
                if value_float < 2.5 and value_float >= 0: # Friability biasanya persentase kecil dan non-negatif
                    if batch not in friability_data_dict: 
                        friability_data_dict[batch] = value_float
                    all_friability_values.append(value_float)
                else: # Diasumsikan Waktu Hancur
                    if batch not in waktu_hancur_data_dict: 
                        waktu_hancur_data_dict[batch] = value_float
                    all_waktu_hancur_values.append(value_float)
            except (ValueError, TypeError):
                st.warning(f"Melewatkan baris untuk batch {batch}, nilai tidak valid: {value_raw}")
                continue
        
        waktu_hancur_df = pd.DataFrame()
        if waktu_hancur_data_dict:
            batch_df_wh = pd.DataFrame(list(waktu_hancur_data_dict.items()), columns=["Batch", "Waktu Hancur"]).sort_values("Batch").reset_index(drop=True)
            wh_min = np.min(all_waktu_hancur_values) if all_waktu_hancur_values else np.nan
            wh_max = np.max(all_waktu_hancur_values) if all_waktu_hancur_values else np.nan
            wh_mean = np.mean(all_waktu_hancur_values) if all_waktu_hancur_values else np.nan
            wh_sd = np.std(all_waktu_hancur_values, ddof=1) if len(all_waktu_hancur_values) > 1 else np.nan
            wh_rsd = (wh_sd / wh_mean * 100) if wh_mean and wh_mean != 0 else np.nan
            stats_df_wh = pd.DataFrame({
                "Batch": ["Minimum", "Maximum", "Rata-rata", "Standar Deviasi", "RSD (%)"],
                "Waktu Hancur": [wh_min, wh_max, wh_mean, wh_sd, wh_rsd]
            })
            waktu_hancur_df = pd.concat([batch_df_wh, stats_df_wh], ignore_index=True)

        friability_df = pd.DataFrame()
        if friability_data_dict:
            batch_df_fr = pd.DataFrame(list(friability_data_dict.items()), columns=["Batch", "Friability"]).sort_values("Batch").reset_index(drop=True)
            fr_min = np.min(all_friability_values) if all_friability_values else np.nan
            fr_max = np.max(all_friability_values) if all_friability_values else np.nan
            fr_mean = np.mean(all_friability_values) if all_friability_values else np.nan
            fr_sd = np.std(all_friability_values, ddof=1) if len(all_friability_values) > 1 else np.nan
            fr_rsd = (fr_sd / fr_mean * 100) if fr_mean and fr_mean != 0 else np.nan
            stats_df_fr = pd.DataFrame({
                "Batch": ["Minimum", "Maximum", "Rata-rata", "Standar Deviasi", "RSD (%)"],
                "Friability": [fr_min, fr_max, fr_mean, fr_sd, fr_rsd]
            })
            friability_df = pd.concat([batch_df_fr, stats_df_fr], ignore_index=True)

        if not waktu_hancur_df.empty:
            st.write("Tabel Waktu Hancur dengan Statistik:")
            st.dataframe(waktu_hancur_df.style.format({"Waktu Hancur": "{:.4f}", "Batch": "{}"}, na_rep=""))
        else:
            st.info("Tidak ada data Waktu Hancur yang diproses atau ditemukan.")
            
        if not friability_df.empty:
            st.write("Tabel Friability dengan Statistik:")
            st.dataframe(friability_df.style.format({"Friability": "{:.4f}", "Batch": "{}"}, na_rep=""))
        else:
            st.info("Tidak ada data Friability yang diproses atau ditemukan.")
            
        return waktu_hancur_df, friability_df
    except Exception as e:
        st.error(f"Gagal memproses file Waktu Hancur dan Friability: {e}")
        st.exception(e)
        return pd.DataFrame(), pd.DataFrame()


def get_excel_for_download(df, index=True): # Tambahkan parameter index
    """
    Mempersiapkan DataFrame untuk diunduh sebagai file Excel.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=index) # Gunakan parameter index
    output.seek(0) 
    return output


def tampilkan_ipc():
    st.title("Halaman IPC")
    st.write("Ini adalah tampilan khusus IPC.")
    
    selected_option = st.radio(
        "Pilih jenis pengujian:",
        [
            "Kekerasan",
            "Keseragaman Bobot",
            "Tebal",
            "Waktu Hancur dan Friability"
        ],
        horizontal=True,
        key="ipc_test_selection" 
    )
    
    template_info = {
        "Kekerasan": "Template Excel untuk pengujian kekerasan.",
        "Keseragaman Bobot": "Template Excel untuk pengujian keseragaman bobot.",
        "Tebal": "Template Excel untuk pengujian tebal.",
        "Waktu Hancur dan Friability": "Template Excel untuk pengujian waktu hancur dan friability (pastikan ada kolom 'Nomor Batch' dan 'Sample Data')."
    }
    
    st.info(f"Upload file Excel dengan format: {template_info[selected_option]}")
    
    uploader_key = f"uploader_{selected_option.replace(' ', '_').lower()}"
    uploaded_file = st.file_uploader("Upload file Excel sesuai template", type=["xlsx","ods"], key=uploader_key)
    
    if uploaded_file:
        file_copy = io.BytesIO(uploaded_file.getvalue()) 
        st.success(f"File untuk pengujian {selected_option} berhasil diupload: {uploaded_file.name}")
        st.subheader(f"Hasil Pengujian {selected_option}")
        
        df_result = None 
        df_result_wh = None 
        df_result_fr = None

        if selected_option == "Kekerasan":
            df_result = parse_kekerasan_excel(file_copy)
            if df_result is not None and not df_result.empty:
                filename = "data_kekerasan"
                excel_bytes_io = get_excel_for_download(df_result, index=True) # Index penting di sini
                st.download_button(
                    label=f"📥 Download Data Kekerasan",
                    data=excel_bytes_io,
                    file_name=f"{filename}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            elif df_result is not None: 
                st.info("Tidak ada data kekerasan valid yang ditemukan dalam file untuk diproses.")

        elif selected_option == "Keseragaman Bobot":
            df_result = parse_keseragaman_bobot_excel(file_copy)
            if df_result is not None and not df_result.empty:
                filename = "data_keseragaman_bobot"
                excel_bytes_io = get_excel_for_download(df_result, index=True) # Index penting di sini
                st.download_button(
                    label=f"📥 Download Data Keseragaman Bobot",
                    data=excel_bytes_io,
                    file_name=f"{filename}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            elif df_result is not None:
                st.info("Tidak ada data keseragaman bobot valid yang ditemukan dalam file untuk diproses.")

        elif selected_option == "Tebal":
            df_result = parse_tebal_excel(file_copy) 
            if df_result is not None and not df_result.empty:
                filename = "data_tebal"
                excel_bytes_io = get_excel_for_download(df_result, index=True) # Index penting di sini (mengandung label MIN, MAX, dll)
                st.download_button(
                    label=f"📥 Download Data Tebal",
                    data=excel_bytes_io,
                    file_name=f"{filename}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            elif df_result is not None:
                st.info("Tidak ada data tebal valid yang ditemukan dalam file untuk diproses.")
        
        elif selected_option == "Waktu Hancur dan Friability":
            df_result_wh, df_result_fr = parse_waktu_hancur_friability_excel(file_copy)
            
            if df_result_wh is not None and not df_result_wh.empty:
                # Untuk Waktu Hancur & Friability, index default (0,1,2,..) kurang informatif, Batch sudah jadi kolom
                excel_bytes_io_wh = get_excel_for_download(df_result_wh, index=False) 
                st.download_button(
                    label="📥 Download Data Waktu Hancur",
                    data=excel_bytes_io_wh,
                    file_name="data_waktu_hancur.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_wh"
                )
            elif df_result_wh is not None: 
                st.info("Tidak ada data Waktu Hancur yang valid untuk diunduh (mungkin tidak ada data yang memenuhi kriteria).")
            
            if df_result_fr is not None and not df_result_fr.empty:
                excel_bytes_io_fr = get_excel_for_download(df_result_fr, index=False)
                st.download_button(
                    label="📥 Download Data Friability",
                    data=excel_bytes_io_fr,
                    file_name="data_friability.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_fr"
                )
            elif df_result_fr is not None:
                st.info("Tidak ada data Friability yang valid untuk diunduh (mungkin tidak ada data yang memenuhi kriteria).")

        # Opsi untuk menampilkan tabel data yang akan diekspor (preview)
        if selected_option != "Waktu Hancur dan Friability":
            if df_result is not None and not df_result.empty:
                show_table_key = f"show_table_{selected_option.replace(' ', '_').lower()}"
                if st.checkbox("Tampilkan tabel data yang akan diekspor (format dasar)", key=show_table_key):
                    st.write("Data Lengkap (sesuai format ekspor):")
                    # Gunakan format yang aman juga untuk preview jika memungkinkan
                    # Namun, untuk kesederhanaan, format default pandas seringkali cukup
                    st.dataframe(df_result.style.format(na_rep="")) 
            elif df_result is None and uploaded_file: 
                st.warning("Tidak ada tabel untuk ditampilkan karena terjadi kesalahan saat pemrosesan file atau file tidak valid.")
    elif st.session_state.get(uploader_key) is not None: 
        st.info("Unggah file Excel baru untuk memulai analisis.")


if __name__ == '__main__':
    tampilkan_ipc()
