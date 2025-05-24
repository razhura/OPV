import streamlit as st
import pandas as pd
import numpy as np
import io
# import base64 # No longer needed for download links
import re

# --- (calculate_statistics, parse_kekerasan_excel, parse_keseragaman_bobot_excel, parse_tebal_excel, parse_waktu_hancur_friability_excel functions remain the same as in the previous version) ---
# For brevity, I'm assuming these functions are unchanged from the last provided version.
# If you need them repeated, let me know. I'll include the modified get_excel_for_download and tampilkan_ipc.

def calculate_statistics(df):
    """
    Calculate statistics (MIN, MAX, MEAN, SD, RSD) for each batch in the dataframe
    """
    stats_df = pd.DataFrame()
    for column in df.columns:
        # Ensure column data is numeric before calculating stats, handle potential errors
        numeric_data = pd.to_numeric(df[column], errors='coerce').dropna()
        if numeric_data.empty:
            min_val, max_val, mean_val, std_val, rsd_val = np.nan, np.nan, np.nan, np.nan, np.nan
        else:
            min_val = numeric_data.min()
            max_val = numeric_data.max()
            mean_val = numeric_data.mean()
            std_val = numeric_data.std()
            rsd_val = (std_val / mean_val * 100) if mean_val != 0 else 0
        
        stats_df[column] = [min_val, max_val, mean_val, std_val, rsd_val]
    stats_df.index = ['MIN', 'MAX', 'MEAN', 'SD', 'RSD (%)']
    return stats_df

def parse_kekerasan_excel(file):
    try:
        df = pd.read_excel(file, header=None, engine='odf') # Specify engine for .ods if needed
        if df.shape[0] < 10 or df.shape[1] < 6: # Basic validation
            st.error("Template tidak sesuai: data minimal tidak terpenuhi (Kekerasan).")
            return None
        
        result_df = pd.DataFrame()
        # batch_names = [] # Not actively used, can be removed if not needed later
        current_row = 2 # Data starts from the 3rd row (index 2)
        # batch_counter = 1 # Not actively used

        while current_row < df.shape[0] - 7: # Ensure enough rows for one batch block (8 rows pattern)
            try:
                batch_name_val = df.iloc[current_row, 0] # Batch name from Column A
                if pd.isna(batch_name_val) or str(batch_name_val).strip() == '':
                    break # Stop if no more batch names
                
                batch_name = str(batch_name_val)

                # Data from Column E (index 4) and F (index 5)
                data_1_5 = df.iloc[current_row:current_row+5, 4] if df.shape[1] > 4 else pd.Series(dtype='float64')
                data_6_10 = df.iloc[current_row:current_row+5, 5] if df.shape[1] > 5 else pd.Series(dtype='float64')
                
                values = pd.concat([data_1_5, data_6_10], ignore_index=True)
                values = pd.to_numeric(values, errors='coerce').dropna()
                
                if len(values) >= 8: # Minimum 8 valid data points
                    values = values[:10] # Take only the first 10 if more are present
                    # Pad with NaN if less than 10 for consistent column length in result_df
                    if len(values) < 10:
                        values = pd.concat([values, pd.Series([np.nan] * (10 - len(values)))], ignore_index=True)
                    result_df[batch_name] = values
                    # batch_names.append(batch_name)
                else:
                    st.warning(f"Data batch {batch_name} tidak lengkap ({len(values)} data valid). Diabaikan.")
                
                current_row += 8 # Move to the next batch block
                # batch_counter += 1
            except IndexError: # Handle cases where rows/columns don't exist as expected
                st.warning(f"Error memproses struktur data pada baris sekitar {current_row}. Periksa format file.")
                current_row += 8 # Attempt to skip to next block
                continue
            except Exception as e: # Catch other errors during batch processing
                st.warning(f"Error memproses batch pada baris {current_row}: {e}")
                current_row += 8
                continue
        
        if result_df.empty:
            st.error("Tidak ada data valid yang dapat diproses dari file Kekerasan.")
            return None
        
        # Data is already structured to have up to 10 rows per batch.
        # The number of rows in result_df will be 10 (or less if all batches had fewer than 10 after padding)
        # Set index from 1 to 10 (or max_len if consistently padded)
        max_data_points = 10
        result_df.index = range(1, max_data_points + 1)
        
        stats_df = calculate_statistics(result_df)
        final_df = pd.concat([result_df, stats_df])
        
        st.write("Data Kekerasan dengan Statistik:")
        
        # Display formatting
        display_df_styled = final_df.style.format(
            lambda x: f"{x:.2f}" if isinstance(x, (int, float)) and pd.notna(x) else "",
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

        header_row_index = 0 # Assuming header is the first row
        # Find header more dynamically if needed, e.g., by looking for "Nomor Batch"
        header_row = df.iloc[header_row_index]
        df_data = df[header_row_index+1:].copy()
        df_data.columns = header_row
        df_data.reset_index(drop=True, inplace=True)

        # Assuming first column is batch identifier
        batch_column_name = df_data.columns[0]
        
        # Remove summary rows (Rata-rata, SD, RSD) if they exist based on first column string match
        df_data = df_data[~df_data.iloc[:, 0].astype(str).str.contains("Rata|SD|RSD", na=False, case=False)]
        
        batch_series = df_data[batch_column_name].dropna().unique()
        result_df = pd.DataFrame()

        # Define column indices for data (E, F, G, H -> 4, 5, 6, 7)
        data_cols_indices = [4, 5, 6, 7] # Corresponds to E, F, G, H

        for batch_val in batch_series:
            batch = str(batch_val)
            subset = df_data[df_data[batch_column_name] == batch_val]
            if subset.empty:
                continue
            
            all_batch_values = []
            for col_idx in data_cols_indices:
                if col_idx < subset.shape[1]: # Check if column exists
                    # Assuming up to 5 values per column block for this test (20 total)
                    col_values = subset.iloc[0:5, col_idx] 
                    all_batch_values.append(col_values)
            
            if not all_batch_values:
                st.warning(f"Tidak ada kolom data (E,F,G,H) ditemukan untuk batch {batch}.")
                continue

            stacked_values = pd.concat(all_batch_values, ignore_index=True)

            def clean_numeric_value(val):
                if isinstance(val, str):
                    # Try to extract number from concatenated strings like "123.456123.789"
                    if len(val) > 8 and not ' ' in val and val.count('.') > 1: 
                        numbers = re.findall(r'\d+\.?\d*', val) # find all number-like parts
                        if numbers:
                            # Heuristic: use the first one, or average, or middle one?
                            # For now, let's try taking the one that seems most plausible if multiple are extracted
                            # This case is tricky and depends on how data is malformed.
                            # A simple approach might be to just try converting the first part.
                            try: return float(numbers[0]) 
                            except: pass # fallback if first part isn't a full number
                    try:
                        return float(val)
                    except ValueError:
                        return np.nan
                elif isinstance(val, (int, float)):
                    return float(val)
                return np.nan

            cleaned_values = stacked_values.apply(clean_numeric_value).dropna()
            
            if not cleaned_values.empty:
                # Pad with NaN if less than 20 data points for consistent column length
                target_length = 20
                if len(cleaned_values) < target_length:
                    padding = pd.Series([np.nan] * (target_length - len(cleaned_values)))
                    cleaned_values = pd.concat([cleaned_values, padding], ignore_index=True)
                result_df[batch] = cleaned_values[:target_length] # Ensure max 20 data points
            else:
                st.warning(f"Tidak ada data numerik valid yang ditemukan untuk batch {batch} setelah pembersihan.")

        if result_df.empty:
            st.error("Tidak ada data Keseragaman Bobot valid yang dapat diproses.")
            return None
        
        result_df.index = range(1, len(result_df) + 1)

        st.write("Data Keseragaman Bobot Terstruktur:")
        st.dataframe(result_df.style.format("{:.4f}", na_rep=""))
        
        stats_df = calculate_statistics(result_df)
        st.write("Statistik Data Keseragaman Bobot:")
        st.dataframe(stats_df.style.format("{:.4f}", na_rep=""))
        
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

        header_row_index = 0 # Assuming header is the first row
        header_row = df.iloc[header_row_index]
        df_data = df[header_row_index+1:].copy()
        df_data.columns = header_row
        df_data.reset_index(drop=True, inplace=True)
        
        batch_column_name = df_data.columns[0] # Assuming first column is batch
        batch_series = df_data[batch_column_name].dropna().unique()
        result_df = pd.DataFrame()

        # Data from Column E (index 4) and F (index 5), typically 3 values each for 6 total
        data_cols_indices = [4, 5] 
        num_values_per_col = 3

        for batch_val in batch_series:
            batch = str(batch_val)
            subset = df_data[df_data[batch_column_name] == batch_val]
            if subset.empty:
                continue

            all_batch_values = []
            for col_idx in data_cols_indices:
                if col_idx < subset.shape[1]: # Check column exists
                    col_values = subset.iloc[0:num_values_per_col, col_idx]
                    all_batch_values.append(col_values)
            
            if not all_batch_values:
                st.warning(f"Tidak ada kolom data (E,F) ditemukan untuk batch {batch}.")
                continue
            
            stacked_values = pd.concat(all_batch_values, ignore_index=True)

            def clean_numeric_value(val): # Same cleaner as Keseragaman Bobot
                if isinstance(val, str):
                    if len(val) > 8 and not ' ' in val and val.count('.') > 1:
                        numbers = re.findall(r'\d+\.?\d*', val)
                        if numbers:
                            try: return float(numbers[0])
                            except: pass
                    try:
                        return float(val)
                    except ValueError:
                        return np.nan
                elif isinstance(val, (int, float)):
                    return float(val)
                return np.nan
            
            cleaned_values = stacked_values.apply(clean_numeric_value).dropna()

            if not cleaned_values.empty:
                target_length = 6 # Typically 6 data points for Tebal
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
        exportable_df = pd.concat([result_df, stats_df]) # This is for export

        # For display, add the statistics labels as the first column
        display_df = exportable_df.copy()
        stat_labels = [""] * len(result_df) + list(stats_df.index)
        # Ensure the index of display_df is suitable for inserting labels
        # If result_df has e.g. 6 rows, index 0-5. stats_df has 5 rows, index 0-4.
        # Combined, 11 rows. stat_labels should match.
        display_df.insert(0, "Keterangan", stat_labels)


        st.write("Data Tebal Terstruktur dengan Statistik:")
        st.dataframe(display_df.style.format("{:.4f}", na_rep="").set_properties(**{'text-align': 'left'}).set_table_styles([dict(selector='th', props=[('text-align', 'left')])]))
        
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
        batch_col_idx = None # Column index for "Nomor Batch"
        value_col_idx = None # Column index for "Sample Data"
        
        # Try to find header row and key columns ("Nomor Batch", "Sample Data")
        for i in range(min(10, len(df))): # Check first 10 rows for header
            row_values = df.iloc[i].astype(str) # Convert row to string for searching
            if "Nomor Batch" in row_values.values:
                header_row_idx = i
                batch_col_idx = row_values[row_values.str.contains("Nomor Batch", case=False, na=False)].index[0]
                # Try to find "Sample Data" in the same header row
                if "Sample Data" in row_values.values:
                     value_col_idx = row_values[row_values.str.contains("Sample Data", case=False, na=False)].index[0]
                break
        
        if header_row_idx is None: # Fallback if "Nomor Batch" not explicitly found
            st.warning("Header 'Nomor Batch' tidak ditemukan, menggunakan asumsi default (Baris 1 header, Kolom A Batch).")
            header_row_idx = 0
            batch_col_idx = 0 
        
        if value_col_idx is None: # Fallback for "Sample Data" column
            # Common position for sample data is column E (index 4)
            potential_value_col = 4 
            if df.shape[1] > potential_value_col:
                st.warning(f"Kolom 'Sample Data' tidak ditemukan, mengasumsikan data ada di kolom E (indeks {potential_value_col}).")
                value_col_idx = potential_value_col
            else: # Not enough columns for default
                st.error("Tidak dapat menemukan kolom data ('Sample Data' atau default). Harap periksa template.")
                return pd.DataFrame(), pd.DataFrame()

        # Set column names from identified header row
        df.columns = df.iloc[header_row_idx]
        data_df = df.iloc[header_row_idx+1:].copy()
        
        # Get actual column names using the indices
        batch_col_name = data_df.columns[batch_col_idx]
        value_col_name = data_df.columns[value_col_idx]

        # Filter out rows where batch number is NaN
        data_df = data_df[~data_df[batch_col_name].isna()]
        
        friability_data_dict = {} # Store as {batch: value}
        waktu_hancur_data_dict = {} # Store as {batch: value}
        
        all_friability_values = []
        all_waktu_hancur_values = []
        
        for _, row in data_df.iterrows():
            batch = str(row[batch_col_name])
            value_raw = row[value_col_name]
            
            if pd.isna(value_raw):
                continue
            
            try:
                value_float = float(value_raw)
                # Heuristic: Friability values are typically small percentages (< 2.5%)
                # Waktu Hancur values are typically larger (time in minutes/seconds)
                if value_float < 2.5: # Assumed Friability
                    if batch not in friability_data_dict: # Take first value for a batch
                        friability_data_dict[batch] = value_float
                    all_friability_values.append(value_float)
                else: # Assumed Waktu Hancur
                    if batch not in waktu_hancur_data_dict: # Take first value for a batch
                        waktu_hancur_data_dict[batch] = value_float
                    all_waktu_hancur_values.append(value_float)
            except (ValueError, TypeError):
                st.warning(f"Melewatkan baris untuk batch {batch}, nilai tidak valid: {value_raw}")
                continue
        
        # Create Waktu Hancur DataFrame
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

        # Create Friability DataFrame
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

        # Display tables
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


def get_excel_for_download(df):
    """
    Prepares a DataFrame for Excel download by writing it to an in-memory BytesIO object.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=True) # Keep index for export, as it often contains valuable labels
    output.seek(0) # Rewind the buffer to the beginning
    return output # Return the BytesIO object


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
        key="ipc_test_selection" # Unique key for the radio button
    )
    
    template_info = {
        "Kekerasan": "Template Excel untuk pengujian kekerasan.",
        "Keseragaman Bobot": "Template Excel untuk pengujian keseragaman bobot.",
        "Tebal": "Template Excel untuk pengujian tebal.",
        "Waktu Hancur dan Friability": "Template Excel untuk pengujian waktu hancur dan friability (pastikan ada kolom 'Nomor Batch' dan 'Sample Data')."
    }
    
    st.info(f"Upload file Excel dengan format: {template_info[selected_option]}")
    
    # Ensure unique key for file_uploader based on selected_option to reset it when option changes
    uploader_key = f"uploader_{selected_option.replace(' ', '_').lower()}"
    uploaded_file = st.file_uploader("Upload file Excel sesuai template", type=["xlsx","ods"], key=uploader_key)
    
    if uploaded_file:
        file_copy = io.BytesIO(uploaded_file.getvalue()) # Work with a copy
        st.success(f"File untuk pengujian {selected_option} berhasil diupload: {uploaded_file.name}")
        st.subheader(f"Hasil Pengujian {selected_option}")
        
        df_result = None 
        df_result_wh = None 
        df_result_fr = None

        if selected_option == "Kekerasan":
            df_result = parse_kekerasan_excel(file_copy)
            if df_result is not None and not df_result.empty:
                filename = "data_kekerasan"
                excel_bytes_io = get_excel_for_download(df_result)
                st.download_button(
                    label=f"📥 Download Data Kekerasan",
                    data=excel_bytes_io,
                    file_name=f"{filename}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            elif df_result is not None: # Empty dataframe, means no valid data found by parser
                 st.info("Tidak ada data kekerasan valid yang ditemukan dalam file untuk diproses.")
            # If df_result is None, error already shown by parser

        elif selected_option == "Keseragaman Bobot":
            df_result = parse_keseragaman_bobot_excel(file_copy)
            if df_result is not None and not df_result.empty:
                filename = "data_keseragaman_bobot"
                excel_bytes_io = get_excel_for_download(df_result)
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
                excel_bytes_io = get_excel_for_download(df_result)
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
                excel_bytes_io_wh = get_excel_for_download(df_result_wh)
                st.download_button(
                    label="📥 Download Data Waktu Hancur",
                    data=excel_bytes_io_wh,
                    file_name="data_waktu_hancur.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_wh"
                )
            elif df_result_wh is not None: # Parser returned an empty (but valid) DataFrame
                st.info("Tidak ada data Waktu Hancur yang valid untuk diunduh (mungkin tidak ada data yang memenuhi kriteria).")
            
            if df_result_fr is not None and not df_result_fr.empty:
                excel_bytes_io_fr = get_excel_for_download(df_result_fr)
                st.download_button(
                    label="📥 Download Data Friability",
                    data=excel_bytes_io_fr,
                    file_name="data_friability.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_fr"
                )
            elif df_result_fr is not None:
                st.info("Tidak ada data Friability yang valid untuk diunduh (mungkin tidak ada data yang memenuhi kriteria).")

        # Option to display processed data (before download)
        if selected_option != "Waktu Hancur dan Friability":
            if df_result is not None and not df_result.empty:
                show_table_key = f"show_table_{selected_option.replace(' ', '_').lower()}"
                if st.checkbox("Tampilkan tabel data yang akan diekspor", key=show_table_key):
                    st.write("Data Lengkap (sesuai format ekspor):")
                    st.dataframe(df_result.style.format(na_rep="")) # Basic display for checkbox
            elif df_result is None and uploaded_file: 
                st.warning("Tidak ada tabel untuk ditampilkan karena terjadi kesalahan saat pemrosesan file atau file tidak valid.")
    elif st.session_state.get(uploader_key) is not None: # If a file was previously uploaded but now cleared
        st.info("Unggah file Excel baru untuk memulai analisis.")


if __name__ == '__main__':
    tampilkan_ipc()
