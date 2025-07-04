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
        if len(val) > 7 and '.' in val and val.count('.') > 1 and ' ' not in val:
            numbers = re.findall(r'-?\d+\.?\d*', val) 
            if numbers:
                try:
                    return float(numbers[0]) 
                except ValueError:
                    pass 
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
            std_val = numeric_data.std(ddof=1) 
            rsd_val = (std_val / mean_val * 100) if mean_val != 0 else np.nan
        
        stats_df[column] = [min_val, max_val, mean_val, std_val, rsd_val]
    stats_df.index = ['MIN', 'MAX', 'MEAN', 'SD', 'RSD (%)']
    return stats_df

# --- Helper untuk Styling Tabel ---
# Daftar label yang mengindikasikan baris statistik
STAT_ROW_LABELS = ['MIN', 'MAX', 'MEAN', 'SD', 'RSD (%)', # Dari calculate_statistics
                   'Minimum', 'Maximum', 'Rata-rata', 'Standar Deviasi'] # Dari Waktu Hancur/Friability

def data_cell_formatter(val):
    """Formatter untuk sel data asli: tampilkan 'apa adanya'."""
    if pd.isna(val): return ""
    # str(val) akan menampilkan 3.0 sebagai "3.0", 3.4 sebagai "3.4"
    return str(val)

def stat_cell_formatter(val, decimals=4):
    """Formatter untuk sel statistik: N angka desimal."""
    if pd.isna(val): return ""
    if isinstance(val, (int, float)):
        return f"{val:.{decimals}f}"
    return str(val)

def set_common_table_properties(styler_obj):
    """Menerapkan properti tampilan umum ke Styler."""
    styler_obj.set_properties(**{'text-align': 'left'}) \
              .set_table_styles([dict(selector='th', props=[('text-align', 'left')])])

def apply_conditional_formatting(df_to_format, id_source_type, # 'index' atau nama kolom
                                 numeric_cols_to_format,
                                 stat_decimals_map, # Dictionary: {nama_fungsi_asal: jumlah_desimal_stat}
                                 parser_origin_name # Nama fungsi parser asal untuk lookup desimal
                                 ):
    """
    Membuat dan mengembalikan objek Pandas Styler dengan format kondisional.
    """
    styler = df_to_format.style
    
    current_stat_decimals = stat_decimals_map.get(parser_origin_name, 4) # Default 4 jika tidak dispesifikkan

    if id_source_type == 'index':
        source_for_mask = df_to_format.index.map(str) # Bandingkan string dengan string
    else: # id_source_type adalah nama kolom
        source_for_mask = df_to_format[id_source_type].astype(str)
        styler.format({id_source_type: "{}"}, na_rep="") # Format kolom ID sebagai string

    data_rows_mask = ~source_for_mask.isin(STAT_ROW_LABELS)
    stat_rows_mask = source_for_mask.isin(STAT_ROW_LABELS)

    # Buat fungsi formatter statistik dengan jumlah desimal yang tepat
    current_stat_formatter = lambda x: stat_cell_formatter(x, decimals=current_stat_decimals)

    # Terapkan formatter ke subset yang sesuai
    styler.format(data_cell_formatter, subset=pd.IndexSlice[data_rows_mask, numeric_cols_to_format], na_rep="")
    styler.format(current_stat_formatter, subset=pd.IndexSlice[stat_rows_mask, numeric_cols_to_format], na_rep="")
    
    set_common_table_properties(styler) # Terapkan properti umum tabel
    return styler

# Definisikan jumlah desimal untuk statistik per jenis pengujian
STAT_DECIMALS_PER_TEST = {
    "Kekerasan": 2,
    "Keseragaman Bobot": 4, # Untuk stats_df yang ditampilkan
    "Tebal": 4,
    "Waktu Hancur": 4,
    "Friability": 4
}

# --- Fungsi Parsing untuk Setiap Jenis Pengujian (dengan styling terpusat) ---

def parse_kekerasan_excel(file):
    try:
        df = pd.read_excel(file, header=None) 
        if df.shape[0] < 10 or df.shape[1] < 6:
            st.error("Template tidak sesuai (Kekerasan).")
            return None
        # ... (logika parsing Kekerasan tetap sama hingga final_df) ...
        result_df = pd.DataFrame()
        current_row = 2
        while current_row < df.shape[0] - 7: 
            try:
                batch_name_val = df.iloc[current_row, 0] 
                if pd.isna(batch_name_val) or str(batch_name_val).strip() == '': break 
                batch_name = str(batch_name_val)
                data_1_5 = df.iloc[current_row:current_row+5, 4] if df.shape[1] > 4 else pd.Series(dtype='float64')
                data_6_10 = df.iloc[current_row:current_row+5, 5] if df.shape[1] > 5 else pd.Series(dtype='float64')
                values = pd.concat([data_1_5, data_6_10], ignore_index=True)
                values = pd.to_numeric(values, errors='coerce').dropna()
                if len(values) >= 8: 
                    values = values[:10] 
                    if len(values) < 10:
                        values = pd.concat([values, pd.Series([np.nan] * (10 - len(values)))], ignore_index=True)
                    result_df[batch_name] = values
                else: st.warning(f"Batch {batch_name} tidak lengkap. Diabaikan.")
                current_row += 8 
            except IndexError: st.warning(f"Error struktur data baris ~{current_row}."); current_row += 8; continue
            except Exception as e: st.warning(f"Error batch baris {current_row}: {e}"); current_row += 8; continue
        if result_df.empty: st.error("Tidak ada data valid (Kekerasan)."); return None
        result_df.index = range(1, 11) # Kekerasan biasanya 10 data
        stats_df = calculate_statistics(result_df)
        final_df = pd.concat([result_df, stats_df])
        
        st.write("Data Kekerasan dengan Statistik:")
        numeric_cols = final_df.columns.tolist()
        styled_df = apply_conditional_formatting(
            df_to_format=final_df,
            id_source_type='index', # Gunakan index (1,2,..,'MIN','MAX') untuk bedakan baris
            numeric_cols_to_format=numeric_cols,
            stat_decimals_map=STAT_DECIMALS_PER_TEST,
            parser_origin_name="Kekerasan"
        )
        st.dataframe(styled_df)
        return final_df
    except Exception as e:
        st.error(f"Gagal memproses file Kekerasan: {e}")
        st.exception(e)
        return None

def parse_keseragaman_bobot_excel(file):
    try:
        df = pd.read_excel(file, header=None)
        if df.empty: st.error("File Keseragaman Bobot kosong."); return None
        # ... (logika parsing Keseragaman Bobot tetap sama hingga result_df dan stats_df) ...
        header_row_index = 0 
        header_row = df.iloc[header_row_index]
        df_data = df[header_row_index+1:].copy()
        df_data.columns = header_row
        df_data.reset_index(drop=True, inplace=True)
        batch_column_name = df_data.columns[0]
        df_data = df_data[~df_data.iloc[:, 0].astype(str).str.contains("Rata|SD|RSD", na=False, case=False)]
        batch_series = df_data[batch_column_name].dropna().unique()
        result_df = pd.DataFrame()
        data_cols_indices = [4, 5, 6, 7]
        for batch_val in batch_series:
            batch = str(batch_val)
            subset = df_data[df_data[batch_column_name] == batch_val]
            if subset.empty: continue
            all_batch_values = []
            for col_idx in data_cols_indices:
                if col_idx < subset.shape[1]: 
                    all_batch_values.append(subset.iloc[0:5, col_idx])
            if not all_batch_values: st.warning(f"Kolom data E,F,G,H tidak ditemukan batch {batch}."); continue
            stacked_values = pd.concat(all_batch_values, ignore_index=True)
            cleaned_values = stacked_values.apply(_clean_numeric_value_helper).dropna() 
            if not cleaned_values.empty:
                target_length = 20
                if len(cleaned_values) < target_length:
                    padding = pd.Series([np.nan] * (target_length - len(cleaned_values)))
                    cleaned_values = pd.concat([cleaned_values, padding], ignore_index=True)
                result_df[batch] = cleaned_values[:target_length]
            else: st.warning(f"Tidak ada data numerik valid batch {batch}.")
        if result_df.empty: st.error("Tidak ada data Keseragaman Bobot valid."); return None
        result_df.index = range(1, len(result_df) + 1)

        st.write("Data Keseragaman Bobot Terstruktur:")
        # Untuk result_df (data asli), semua sel diformat dengan data_cell_formatter
        styled_result_df = result_df.style.format(data_cell_formatter, na_rep="")
        set_common_table_properties(styled_result_df)
        st.dataframe(styled_result_df)
        
        stats_df = calculate_statistics(result_df)
        st.write("Statistik Data Keseragaman Bobot:")
        # Untuk stats_df, semua sel diformat dengan stat_cell_formatter
        stat_formatter_for_kb = lambda x: stat_cell_formatter(x, decimals=STAT_DECIMALS_PER_TEST["Keseragaman Bobot"])
        styled_stats_df = stats_df.style.format(stat_formatter_for_kb, na_rep="")
        set_common_table_properties(styled_stats_df)
        st.dataframe(styled_stats_df)
        
        export_df = pd.concat([result_df, stats_df])
        return export_df
    except Exception as e:
        st.error(f"Gagal memproses file Keseragaman Bobot: {e}")
        st.exception(e)
        return None

def parse_keseragaman_bobot_effervescent_excel(file):
    try:
        df = pd.read_excel(file, header=None)
        if df.empty:
            st.error("File Keseragaman Bobot Effervescent kosong.")
            return None

        # Ambil hanya kolom yang diperlukan: Nomor Batch dan Data 1-10
        df_needed = df.iloc[:, [0] + list(range(4, df.shape[1]))].copy()
        df_needed.columns = ['Nomor Batch'] + [f'Data{i}' for i in range(1, df_needed.shape[1])]

        # Group by Nomor Batch
        grouped = df_needed.groupby('Nomor Batch')

        batch_dict = {}
        for batch_name, group in grouped:
            data_values = []

            for _, row in group.iterrows():
                row_values = row.iloc[1:].tolist()
                cleaned_values = [v for v in row_values if pd.notna(v)]
                data_values.extend(cleaned_values)

            batch_dict[batch_name] = data_values

        # Membuat dataframe hasil: 1 batch = 1 kolom
        max_length = max(len(v) for v in batch_dict.values())
        result_df = pd.DataFrame()

        for batch, values in batch_dict.items():
            # Padding jika jumlah data kurang dari batch lain
            if len(values) < max_length:
                values.extend([np.nan] * (max_length - len(values)))
            result_df[batch] = values

        # Tampilkan hasil di Streamlit
        st.write("Data Keseragaman Bobot Effervescent Transpose (1 batch = 1 kolom):")
        styled_df = result_df.style.format(na_rep="")
        set_common_table_properties(styled_df)
        st.dataframe(styled_df)

        return result_df

    except Exception as e:
        st.error(f"Gagal memproses file Keseragaman Bobot Effervescent: {e}")
        st.exception(e)
        return None



def parse_tebal_excel(file):
    try:
        df = pd.read_excel(file, header=None)
        if df.empty: st.error("File Tebal kosong."); return None
        # ... (logika parsing Tebal tetap sama hingga exportable_df) ...
        header_row_index = 0
        header_row = df.iloc[header_row_index]
        df_data = df[header_row_index+1:].copy()
        df_data.columns = header_row
        df_data.reset_index(drop=True, inplace=True)
        batch_column_name = df_data.columns[0] 
        batch_series = df_data[batch_column_name].dropna().unique()
        result_df = pd.DataFrame()
        data_cols_indices = [4, 5] 
        num_values_per_col = 3
        for batch_val in batch_series:
            batch = str(batch_val)
            subset = df_data[df_data[batch_column_name] == batch_val]
            if subset.empty: continue
            all_batch_values = []
            for col_idx in data_cols_indices:
                if col_idx < subset.shape[1]: 
                    all_batch_values.append(subset.iloc[0:num_values_per_col, col_idx])
            if not all_batch_values: st.warning(f"Kolom data E,F tidak ditemukan batch {batch}."); continue
            stacked_values = pd.concat(all_batch_values, ignore_index=True)
            cleaned_values = stacked_values.apply(_clean_numeric_value_helper).dropna()
            if not cleaned_values.empty:
                target_length = 6 
                if len(cleaned_values) < target_length:
                    padding = pd.Series([np.nan] * (target_length - len(cleaned_values)))
                    cleaned_values = pd.concat([cleaned_values, padding], ignore_index=True)
                result_df[batch] = cleaned_values[:target_length]
            else: st.warning(f"Tidak ada data numerik valid batch tebal {batch}.")
        if result_df.empty: st.error("Tidak ada data Tebal valid."); return None
        result_df.index = range(1, len(result_df) + 1)
        stats_df = calculate_statistics(result_df)
        exportable_df = pd.concat([result_df, stats_df]) 

        display_df = exportable_df.copy()
        display_df.insert(0, "Keterangan", display_df.index.map(str)) 
        display_df = display_df.reset_index(drop=True)

        st.write("Data Tebal Terstruktur dengan Statistik:")
        numeric_cols = [col for col in display_df.columns if col != "Keterangan"]
        styled_df = apply_conditional_formatting(
            df_to_format=display_df,
            id_source_type='Keterangan', # Gunakan kolom "Keterangan" untuk bedakan baris
            numeric_cols_to_format=numeric_cols,
            stat_decimals_map=STAT_DECIMALS_PER_TEST,
            parser_origin_name="Tebal"
        )
        st.dataframe(styled_df)
        return exportable_df
    except Exception as e:
        st.error(f"Gagal memproses file Tebal: {e}")
        st.exception(e)
        return None

def parse_waktu_hancur_friability_excel(file):
    try:
        df = pd.read_excel(file, header=None)
        if df.empty: st.error("File Waktu Hancur/Friability kosong."); return pd.DataFrame(), pd.DataFrame()
        # ... (logika parsing Waktu Hancur & Friability tetap sama hingga waktu_hancur_df dan friability_df) ...
        header_row_idx = None; batch_col_idx = None; value_col_idx = None 
        for i in range(min(10, len(df))): 
            row_values = df.iloc[i].astype(str) 
            if "Nomor Batch" in row_values.values:
                header_row_idx = i
                batch_col_idx = row_values[row_values.str.contains("Nomor Batch", case=False, na=False)].index[0]
                if "Sample Data" in row_values.values: value_col_idx = row_values[row_values.str.contains("Sample Data", case=False, na=False)].index[0]
                break
        if header_row_idx is None: st.warning("Header 'Nomor Batch' tidak ditemukan..."); header_row_idx = 0; batch_col_idx = 0 
        if value_col_idx is None: 
            potential_value_col = 4 
            if df.shape[1] > potential_value_col: st.warning(f"Kolom 'Sample Data' tidak ditemukan..."); value_col_idx = potential_value_col
            else: st.error("Tidak dapat menemukan kolom data."); return pd.DataFrame(), pd.DataFrame()
        df.columns = df.iloc[header_row_idx]; data_df = df.iloc[header_row_idx+1:].copy()
        batch_col_name = data_df.columns[batch_col_idx]; value_col_name = data_df.columns[value_col_idx]
        data_df = data_df[~data_df[batch_col_name].isna()]
        friability_data_dict = {}; waktu_hancur_data_dict = {} 
        all_friability_values = []; all_waktu_hancur_values = []
        for _, row in data_df.iterrows():
            batch = str(row[batch_col_name]); value_raw = row[value_col_name]
            if pd.isna(value_raw): continue
            try:
                value_float = float(value_raw)
                if 0 <= value_float < 2.5 : # Friability
                    if batch not in friability_data_dict: friability_data_dict[batch] = value_float
                    all_friability_values.append(value_float)
                else: # Waktu Hancur
                    if batch not in waktu_hancur_data_dict: waktu_hancur_data_dict[batch] = value_float
                    all_waktu_hancur_values.append(value_float)
            except (ValueError, TypeError): st.warning(f"Melewatkan baris batch {batch}, nilai tidak valid: {value_raw}"); continue
        waktu_hancur_df = pd.DataFrame()
        if waktu_hancur_data_dict:
            batch_df_wh = pd.DataFrame(list(waktu_hancur_data_dict.items()), columns=["Batch", "Waktu Hancur"]).sort_values("Batch").reset_index(drop=True)
            wh_min = np.min(all_waktu_hancur_values) if all_waktu_hancur_values else np.nan; wh_max = np.max(all_waktu_hancur_values) if all_waktu_hancur_values else np.nan
            wh_mean = np.mean(all_waktu_hancur_values) if all_waktu_hancur_values else np.nan; wh_sd = np.std(all_waktu_hancur_values, ddof=1) if len(all_waktu_hancur_values) > 1 else np.nan
            wh_rsd = (wh_sd / wh_mean * 100) if wh_mean and wh_mean != 0 else np.nan
            stats_df_wh = pd.DataFrame({"Batch": ["Minimum", "Maximum", "Rata-rata", "Standar Deviasi", "RSD (%)"], "Waktu Hancur": [wh_min, wh_max, wh_mean, wh_sd, wh_rsd]})
            waktu_hancur_df = pd.concat([batch_df_wh, stats_df_wh], ignore_index=True)
        friability_df = pd.DataFrame()
        if friability_data_dict:
            batch_df_fr = pd.DataFrame(list(friability_data_dict.items()), columns=["Batch", "Friability"]).sort_values("Batch").reset_index(drop=True)
            fr_min = np.min(all_friability_values) if all_friability_values else np.nan; fr_max = np.max(all_friability_values) if all_friability_values else np.nan
            fr_mean = np.mean(all_friability_values) if all_friability_values else np.nan; fr_sd = np.std(all_friability_values, ddof=1) if len(all_friability_values) > 1 else np.nan
            fr_rsd = (fr_sd / fr_mean * 100) if fr_mean and fr_mean != 0 else np.nan
            stats_df_fr = pd.DataFrame({"Batch": ["Minimum", "Maximum", "Rata-rata", "Standar Deviasi", "RSD (%)"], "Friability": [fr_min, fr_max, fr_mean, fr_sd, fr_rsd]})
            friability_df = pd.concat([batch_df_fr, stats_df_fr], ignore_index=True)
        
        if not waktu_hancur_df.empty:
            st.write("Tabel Waktu Hancur dengan Statistik:")
            styled_wh_df = apply_conditional_formatting(
                df_to_format=waktu_hancur_df,
                id_source_type='Batch', # Gunakan kolom "Batch" untuk bedakan baris
                numeric_cols_to_format=["Waktu Hancur"],
                stat_decimals_map=STAT_DECIMALS_PER_TEST,
                parser_origin_name="Waktu Hancur"
            )
            st.dataframe(styled_wh_df)
        else:
            st.info("Tidak ada data Waktu Hancur yang diproses atau ditemukan.")
            
        if not friability_df.empty:
            st.write("Tabel Friability dengan Statistik:")
            styled_fr_df = apply_conditional_formatting(
                df_to_format=friability_df,
                id_source_type='Batch', # Gunakan kolom "Batch" untuk bedakan baris
                numeric_cols_to_format=["Friability"],
                stat_decimals_map=STAT_DECIMALS_PER_TEST,
                parser_origin_name="Friability"
            )
            st.dataframe(styled_fr_df)
        else:
            st.info("Tidak ada data Friability yang diproses atau ditemukan.")
            
        return waktu_hancur_df, friability_df
    except Exception as e:
        st.error(f"Gagal memproses file Waktu Hancur dan Friability: {e}")
        st.exception(e)
        return pd.DataFrame(), pd.DataFrame()

def get_excel_for_download(df, index=True):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=index) 
    output.seek(0) 
    return output

def tampilkan_ipc():
    st.title("Halaman IPC")
    st.write("Ini adalah tampilan khusus IPC")
    
    selected_option = st.radio(
        "Pilih jenis pengujian:",
        ["Kekerasan", "Keseragaman Bobot", "Keseragaman Bobot Effervescent", "Tebal", "Waktu Hancur dan Friability"],
        horizontal=True, key="ipc_test_selection" 
    )
    
    template_info = {
        "Kekerasan": "Template Excel untuk pengujian kekerasan.",
        "Keseragaman Bobot": "Template Excel untuk pengujian keseragaman bobot.",
        "Keseragaman Bobot Effervescent": "Template Excel untuk pengujian keseragaman bobot effervescent.",
        "Tebal": "Template Excel untuk pengujian tebal.",
        "Waktu Hancur dan Friability": "Template Excel untuk pengujian waktu hancur dan friability (kolom 'Nomor Batch' & 'Sample Data')."
    }
    st.info(f"Upload file Excel dengan format: {template_info[selected_option]}")
    template_links = {
        "Kekerasan": "https://drive.google.com/file/d/1fNFMrq6eiLsfRq-_9sM_lsAJcIvEdzfa/view?usp=drive_link",
        "Keseragaman Bobot": "https://drive.google.com/file/d/1Qf13SEbM34IHvOcJ72jgnuU7xoYuydMJ/view?usp=drive_link",
        "Keseragaman Bobot Effervescent": "#",  # Ganti nanti jika sudah ada link template
        "Tebal": "https://drive.google.com/file/d/1US8atXnBTN6zBLLBPhojGKEcf7o3_Vd8/view?usp=drive_link",
        "Waktu Hancur dan Friability": "https://drive.google.com/file/d/1_L87a1pB8eK7JwaKHaQKqfhhXHTqtZHs/view?usp=drive_link"
    }
    
    st.markdown(
        f"[📥 Download Template {selected_option} di sini]({template_links[selected_option]})",
        unsafe_allow_html=True
    )
    uploader_key = f"uploader_{selected_option.replace(' ', '_').lower()}"
    uploaded_file = st.file_uploader("Upload file Excel sesuai template", type=["xlsx","ods"], key=uploader_key)
    
    if uploaded_file:
        file_copy = io.BytesIO(uploaded_file.getvalue()) 
        st.success(f"File untuk pengujian {selected_option} berhasil diupload: {uploaded_file.name}")
        st.subheader(f"Hasil Pengujian {selected_option}")
        
        df_result = None; df_result_wh = None; df_result_fr = None

        if selected_option == "Kekerasan":
            df_result = parse_kekerasan_excel(file_copy)
            if df_result is not None and not df_result.empty:
                excel_bytes_io = get_excel_for_download(df_result, index=True)
                st.download_button(label=f"📥 Download Data Kekerasan", data=excel_bytes_io, file_name="data_kekerasan.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            elif df_result is not None: st.info("Tidak ada data kekerasan valid.")
        elif selected_option == "Keseragaman Bobot":
            df_result = parse_keseragaman_bobot_excel(file_copy)
            if df_result is not None and not df_result.empty:
                excel_bytes_io = get_excel_for_download(df_result, index=True)
                st.download_button(label=f"📥 Download Data Keseragaman Bobot", data=excel_bytes_io, file_name="data_keseragaman_bobot.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            elif df_result is not None: st.info("Tidak ada data keseragaman bobot valid.")
        
        elif selected_option == "Keseragaman Bobot Effervescent":
            df_result = parse_keseragaman_bobot_effervescent_excel(file_copy)
            if df_result is not None and not df_result.empty:
                excel_bytes_io = get_excel_for_download(df_result, index=False)
                st.download_button(
                    label=f"📥 Download Data Keseragaman Bobot Effervescent",
                    data=excel_bytes_io,
                    file_name="data_keseragaman_bobot_effervescent.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            elif df_result is not None:
                st.info("Tidak ada data Keseragaman Bobot Effervescent valid.")

        elif selected_option == "Tebal":
            df_result = parse_tebal_excel(file_copy) 
            if df_result is not None and not df_result.empty:
                excel_bytes_io = get_excel_for_download(df_result, index=True)
                st.download_button(label=f"📥 Download Data Tebal", data=excel_bytes_io, file_name="data_tebal.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            elif df_result is not None: st.info("Tidak ada data tebal valid.")
        elif selected_option == "Waktu Hancur dan Friability":
            df_result_wh, df_result_fr = parse_waktu_hancur_friability_excel(file_copy)
            if df_result_wh is not None and not df_result_wh.empty:
                excel_bytes_io_wh = get_excel_for_download(df_result_wh, index=False) 
                st.download_button(label="📥 Download Data Waktu Hancur", data=excel_bytes_io_wh, file_name="data_waktu_hancur.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_wh")
            elif df_result_wh is not None: st.info("Tidak ada data Waktu Hancur valid untuk diunduh.")
            if df_result_fr is not None and not df_result_fr.empty:
                excel_bytes_io_fr = get_excel_for_download(df_result_fr, index=False)
                st.download_button(label="📥 Download Data Friability", data=excel_bytes_io_fr, file_name="data_friability.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_fr")
            elif df_result_fr is not None: st.info("Tidak ada data Friability valid untuk diunduh.")

        if selected_option != "Waktu Hancur dan Friability":
            if df_result is not None and not df_result.empty:
                show_table_key = f"show_table_{selected_option.replace(' ', '_').lower()}"
                if st.checkbox("Tampilkan tabel data yang akan diekspor (format dasar)", key=show_table_key):
                    st.write("Data Lengkap (sesuai format ekspor):")
                    # Untuk preview, styling sederhana mungkin cukup, atau bisa juga terapkan conditional formatting
                    # Namun, df_result di sini adalah DataFrame mentah sebelum display_df (untuk Tebal)
                    # jadi conditional formatting akan lebih kompleks diterapkan di sini.
                    st.dataframe(df_result.style.format(na_rep="")) 
            elif df_result is None and uploaded_file: 
                st.warning("Tidak ada tabel untuk ditampilkan karena kesalahan pemrosesan atau file tidak valid.")
    elif st.session_state.get(uploader_key) is not None: 
        st.info("Unggah file Excel baru untuk memulai analisis.")

if __name__ == '__main__':
    tampilkan_ipc()
