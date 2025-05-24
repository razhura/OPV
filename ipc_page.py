import streamlit as st
import pandas as pd
import numpy as np
import io
import base64
import re # Ensure re is imported if used in clean_numeric_value

def calculate_statistics(df):
    """
    Calculate statistics (MIN, MAX, MEAN, SD, RSD) for each batch in the dataframe
    """
    stats_df = pd.DataFrame()
    for column in df.columns:
        min_val = df[column].min()
        max_val = df[column].max()
        mean_val = df[column].mean()
        std_val = df[column].std()
        rsd_val = (std_val / mean_val * 100) if mean_val != 0 else 0
        stats_df[column] = [min_val, max_val, mean_val, std_val, rsd_val]
    stats_df.index = ['MIN', 'MAX', 'MEAN', 'SD', 'RSD (%)']
    return stats_df

def parse_kekerasan_excel(file):
    try:
        df = pd.read_excel(file, header=None, engine='odf')
        if df.shape[0] < 10 or df.shape[1] < 6:
            st.error("Template tidak sesuai: data minimal tidak terpenuhi.")
            return None
        
        result_df = pd.DataFrame()
        batch_names = []
        current_row = 2
        batch_counter = 1
        
        while current_row < df.shape[0] - 7:
            try:
                batch_name = df.iloc[current_row, 0]
                if pd.isna(batch_name) or batch_name == '':
                    break
                
                data_1_5 = df.iloc[current_row:current_row+5, 4]
                data_6_10 = df.iloc[current_row:current_row+5, 5]
                
                values = pd.concat([data_1_5, data_6_10], ignore_index=True)
                values = pd.to_numeric(values, errors='coerce').dropna()
                
                if len(values) >= 8:
                    values = values[:10] if len(values) > 10 else values
                    result_df[str(batch_name)] = values
                    batch_names.append(str(batch_name))
                else:
                    st.warning(f"Data batch {batch_name} tidak lengkap ({len(values)} data). Diabaikan.")
                
                current_row += 8
                batch_counter += 1
            except Exception as e:
                st.warning(f"Error pada baris {current_row}: {e}")
                current_row += 8
                continue
        
        if result_df.empty:
            st.error("Tidak ada data valid yang dapat diproses.")
            return None
        
        max_len = 0
        if not result_df.empty: # Check if result_df has columns
            max_len = max([len(result_df[col]) for col in result_df.columns]) if result_df.columns.size > 0 else 0
        
        for col in result_df.columns:
            if len(result_df[col]) < max_len:
                padding = [np.nan] * (max_len - len(result_df[col])) # Use np.nan for padding
                current_values = list(result_df[col])
                current_values.extend(padding)
                result_df[col] = current_values
        
        if max_len > 0: # Set index only if there's data
            result_df.index = range(1, max_len + 1)
        
        stats_df = calculate_statistics(result_df)
        final_df = pd.concat([result_df, stats_df])
        
        st.write("Data Kekerasan dengan Statistik:")
        
        display_df = final_df.copy()
        
        # Formatting for display
        # Determine subset for data rows dynamically
        data_rows_subset = pd.IndexSlice[result_df.index, :] if not result_df.empty else pd.IndexSlice[:, :]

        styled_df = display_df.style.format({
            col: lambda x: f"{x:.2f}" if pd.notna(x) and isinstance(x, (int, float)) else ""
            for col in display_df.columns
        }, subset=data_rows_subset)
        
        styled_df = styled_df.format({
            col: lambda x: f"{x:.2f}" if isinstance(x, (int, float)) and pd.notna(x) else str(x)
            for col in display_df.columns
        }, subset=pd.IndexSlice[stats_df.index, :])
        
        st.dataframe(styled_df)
        return final_df
    except Exception as e:
        st.error(f"Gagal memproses file Kekerasan: {e}")
        st.write("Detail error:", str(e))
        return None

def parse_keseragaman_bobot_excel(file):
    try:
        df = pd.read_excel(file, header=None)
        header_row = df.iloc[0]
        df = df[1:]
        df.columns = header_row
        df.reset_index(drop=True, inplace=True)
        df = df[~df.iloc[:, 0].astype(str).str.contains("Rata|SD|RSD", na=False)]
        batch_series = df.iloc[:, 0].dropna().unique()
        result_df = pd.DataFrame()

        for batch in batch_series:
            subset = df[df.iloc[:, 0] == batch]
            if subset.empty:
                continue
            try:
                values_e = subset.iloc[0:5, 4] if subset.shape[1] > 4 else pd.Series(dtype='float64')
                values_f = subset.iloc[0:5, 5] if subset.shape[1] > 5 else pd.Series(dtype='float64')
                values_g = subset.iloc[0:5, 6] if subset.shape[1] > 6 else pd.Series(dtype='float64')
                values_h = subset.iloc[0:5, 7] if subset.shape[1] > 7 else pd.Series(dtype='float64')
                
                def clean_numeric_value(val):
                    if isinstance(val, str):
                        if len(val) > 8 and not ' ' in val:
                            numbers = re.findall(r'\d+\.\d+|\d+', val)
                            if numbers:
                                middle_index = len(numbers) // 2
                                return float(numbers[middle_index])
                            else:
                                return np.nan
                        else:
                            try:
                                return float(val)
                            except:
                                return np.nan
                    elif isinstance(val, (int, float)):
                         return float(val)
                    else:
                        return np.nan # Ensure non-numeric, non-string become NaN
                
                values_e = values_e.apply(clean_numeric_value)
                values_f = values_f.apply(clean_numeric_value)
                values_g = values_g.apply(clean_numeric_value)
                values_h = values_h.apply(clean_numeric_value)
                
                stacked = pd.concat([values_e, values_f, values_g, values_h], ignore_index=True)
                stacked = pd.to_numeric(stacked, errors='coerce').dropna()
                
                if not stacked.empty: # Add only if there is data
                     result_df[str(batch)] = stacked # Ensure batch is string for column name
            except Exception as e:
                st.warning(f"Ada masalah saat memproses batch {batch}: {e}")
                continue
        
        if result_df.empty:
            st.error("Tidak ada data valid yang dapat diproses.")
            return None
        
        # Pad columns to the same length
        max_len = 0
        if not result_df.empty:
             max_len = max(len(result_df[col]) for col in result_df.columns) if result_df.columns.size > 0 else 0

        for col in result_df.columns:
            if len(result_df[col]) < max_len:
                padding = [np.nan] * (max_len - len(result_df[col]))
                current_values = list(result_df[col])
                current_values.extend(padding)
                result_df[col] = pd.Series(current_values) # Ensure it's a Series

        if max_len > 0:
            result_df.index = range(1, max_len + 1)

        st.write("Data Keseragaman Bobot Terstruktur:")
        st.dataframe(result_df.style.format("{:.4f}", na_rep="")) # Format for display
        
        stats_df = calculate_statistics(result_df)
        st.write("Statistik Data Keseragaman Bobot:")
        st.dataframe(stats_df.style.format("{:.4f}", na_rep=""))
        
        export_df = pd.concat([result_df, stats_df])
        return export_df
    except Exception as e:
        st.error(f"Gagal memproses file Keseragaman Bobot: {e}")
        st.write("Detail error:", str(e))
        return None

def parse_tebal_excel(file):
    try:
        df = pd.read_excel(file, header=None)
        header_row = df.iloc[0]
        df = df[1:]
        df.columns = header_row
        df.reset_index(drop=True, inplace=True)
        batch_series = df.iloc[:, 0].dropna().unique()
        result_df = pd.DataFrame()

        for batch in batch_series:
            subset = df[df.iloc[:, 0] == batch]
            if subset.empty:
                continue
            try:
                # Ensure columns exist before trying to access them
                values_e = subset.iloc[0:3, 4].copy() if subset.shape[1] > 4 else pd.Series(dtype='float64')
                values_f = subset.iloc[0:3, 5].copy() if subset.shape[1] > 5 else pd.Series(dtype='float64')
                
                def clean_numeric_value(val):
                    if pd.isna(val):
                        return np.nan
                    if isinstance(val, (int, float)):
                        return float(val)
                    if isinstance(val, str):
                        if len(val) > 8 and ' ' not in val:
                            numbers = re.findall(r'\d+\.\d+|\d+', val)
                            if numbers:
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
                
                values_e = values_e.apply(clean_numeric_value)
                values_f = values_f.apply(clean_numeric_value)
                
                stacked = pd.concat([values_e, values_f], ignore_index=True)
                stacked = pd.to_numeric(stacked, errors='coerce').dropna()
                
                if not stacked.empty:
                    result_df[str(batch)] = stacked
            except Exception as e:
                st.warning(f"Ada masalah saat memproses batch {batch}: {e}")
                continue
        
        if result_df.empty:
            st.error("Tidak ada data valid yang dapat diproses.")
            return None

        # Pad columns to the same length
        max_len = 0
        if not result_df.empty:
            max_len = max(len(result_df[col]) for col in result_df.columns) if result_df.columns.size > 0 else 0
        
        for col in result_df.columns:
            if len(result_df[col]) < max_len:
                padding = [np.nan] * (max_len - len(result_df[col]))
                current_values = list(result_df[col])
                current_values.extend(padding)
                result_df[col] = pd.Series(current_values) # Ensure it's a Series

        if max_len > 0:
            result_df.index = range(1, max_len + 1)

        stats = {
            "MIN": result_df.min(),
            "MAX": result_df.max(),
            "MEAN": result_df.mean(),
            "SD": result_df.std(),
            "RSD (%)": (result_df.std() / result_df.mean()) * 100 if not result_df.mean().eq(0).any() else 0
        }
        stats_df = pd.DataFrame(stats).T
        
        # Ensure correct index for stats_df if result_df was empty or had no numeric data for stats
        if not result_df.empty and max_len > 0:
            max_idx_val = result_df.index.max() if result_df.index.size > 0 else 0
        else: # Handle case where result_df might be empty or all NaN after processing
            max_idx_val = 0
            # If result_df is completely empty or all NaN, stats_df might also be all NaN or empty
            # For display purposes, we create an empty result_df if it's None
            if result_df.empty and not isinstance(result_df, pd.DataFrame):
                result_df = pd.DataFrame()


        # If stats_df is not empty, assign its index
        if not stats_df.empty:
             stats_df.index = ['MIN', 'MAX', 'MEAN', 'SD', 'RSD (%)']


        combined_df = pd.concat([result_df, stats_df])
        
        # Prepare for display - add empty first column with labels for stats
        # Create labels column
        labels_list = [""] * len(result_df) + list(stats_df.index)
        # Ensure combined_df has a compatible index before inserting
        combined_df.reset_index(drop=True, inplace=True) # Reset index to simple range
        combined_df.insert(0, "Statistik", labels_list[:len(combined_df)]) # Match length
        
        # Rename the index of combined_df for display purposes if needed
        # Or just use the new "Statistik" column and hide index for st.dataframe
        # For export, we might want the original structure without this "Statistik" column
        # So, let's create a display_df and return the original combined_df for export
        
        display_final_df = combined_df.copy()
        # The first column which was index now has labels, original index is gone
        # For display, if the first column is named "Statistik", it's fine.
        # For export, we might want the original format.
        # The export_dataframe function handles index=True by default.
        
        # Let's return the dataframe that is logically combined_df before inserting the 'Statistik' column for display
        exportable_df = pd.concat([result_df, stats_df])


        st.write("Data Tebal Terstruktur dengan Statistik:")
        def format_values(val):
            if isinstance(val, (int, float)) and not pd.isna(val):
                return f"{val:.4f}"
            return str(val) if pd.notna(val) else "" # ensure NaNs are empty strings
        
        # Use the display_final_df for st.dataframe
        # Set the "Statistik" column as index for better display alignment
        if "Statistik" in display_final_df.columns:
            display_final_df.set_index("Statistik", inplace=True)

        st.dataframe(display_final_df.applymap(format_values))
        
        return exportable_df # Return the original combined data for export

    except Exception as e:
        st.error(f"Gagal memproses file Tebal: {e}")
        st.write("Detail error:", str(e))
        return None

def parse_waktu_hancur_friability_excel(file):
    try:
        df = pd.read_excel(file, header=None)
        # st.write("Excel file loaded. Shape:", df.shape) # Debug
        
        header_row_idx = None
        batch_col = None
        value_col = None
        
        for i in range(min(10, len(df))):
            row = df.iloc[i]
            for j, cell in enumerate(row):
                if isinstance(cell, str) and "Nomor Batch" in cell:
                    header_row_idx = i
                    batch_col = j
                    break
            if header_row_idx is not None:
                break
        
        if header_row_idx is None:
            header_row_idx = 0
            batch_col = 0 # Assume first column for batch if "Nomor Batch" not found
            
        if header_row_idx is not None:
            header_row = df.iloc[header_row_idx]
            for j, cell in enumerate(header_row):
                if isinstance(cell, str) and "Sample Data" in cell: # Assuming "Sample Data" is the values column
                    value_col = j
                    break
        
        if value_col is None: # Default if "Sample Data" not found
             # Try to find a column that looks like data (numeric, often after batch and other info)
             # A common pattern is column E (index 4)
            if df.shape[1] > 4: # Check if column E exists
                 value_col = 4
            elif df.shape[1] > batch_col + 1: # Or the column right after batch_col if available
                 value_col = batch_col + 1
            else: # Fallback to a guess or raise error
                 st.error("Tidak dapat menemukan kolom data (Sample Data). Harap periksa template.")
                 return pd.DataFrame(), pd.DataFrame()


        # st.write(f"Using header row: {header_row_idx}, Batch column: {batch_col}, Value column: {value_col}") # Debug
        
        data_df = df.iloc[header_row_idx+1:].copy()
        data_df = data_df[~data_df.iloc[:, batch_col].isna()] # Filter out rows with no batch numbers
        
        friability_data = {}
        waktu_hancur_data = {}
        
        all_friability_values = []
        all_waktu_hancur_values = []
        
        for _, row in data_df.iterrows():
            # Ensure row has enough columns before accessing
            if row.shape[0] <= batch_col or row.shape[0] <= value_col:
                st.warning(f"Melewatkan baris karena kekurangan kolom: {row}")
                continue

            batch = row.iloc[batch_col]
            value = row.iloc[value_col] if not pd.isna(row.iloc[value_col]) else None
            
            if value is None:
                continue
            
            try:
                value_float = float(value)
                # Friability is typically a percentage loss, usually small (e.g., < 1.0% or < 2.5%)
                # Disintegration time is in minutes or seconds, usually > 2.5 (unless it's seconds and very fast)
                # This threshold (2.5) is a heuristic and might need adjustment based on typical data ranges.
                if value_float < 2.5: # Heuristic for friability (percentage)
                    if str(batch) not in friability_data: # Take first value encountered for a batch
                         friability_data[str(batch)] = value_float
                    all_friability_values.append(value_float)
                else: # Heuristic for disintegration time (minutes or larger values)
                    if str(batch) not in waktu_hancur_data: # Take first value encountered for a batch
                        waktu_hancur_data[str(batch)] = value_float
                    all_waktu_hancur_values.append(value_float)
            except (ValueError, TypeError):
                st.warning(f"Skipping row with batch {batch}, invalid value for conversion to float: {value}")
                continue
        
        # Waktu Hancur processing
        waktu_hancur_df = pd.DataFrame()
        if waktu_hancur_data:
            batch_df_wh = pd.DataFrame({
                "Batch": list(waktu_hancur_data.keys()),
                "Waktu Hancur": list(waktu_hancur_data.values())
            }).sort_values("Batch").reset_index(drop=True)

            wh_min = min(all_waktu_hancur_values) if all_waktu_hancur_values else np.nan
            wh_max = max(all_waktu_hancur_values) if all_waktu_hancur_values else np.nan
            wh_mean = np.mean(all_waktu_hancur_values) if all_waktu_hancur_values else np.nan
            wh_sd = np.std(all_waktu_hancur_values, ddof=1) if len(all_waktu_hancur_values) > 1 else np.nan
            wh_rsd = (wh_sd / wh_mean * 100) if wh_mean and wh_mean != 0 else np.nan
            
            stats_df_wh = pd.DataFrame({
                "Batch": ["Minimum", "Maximum", "Rata-rata", "Standar Deviasi", "RSD (%)"],
                "Waktu Hancur": [wh_min, wh_max, wh_mean, wh_sd, wh_rsd]
            })
            waktu_hancur_df = pd.concat([batch_df_wh, stats_df_wh], ignore_index=True)

        # Friability processing
        friability_df = pd.DataFrame()
        if friability_data:
            batch_df_fr = pd.DataFrame({
                "Batch": list(friability_data.keys()),
                "Friability": list(friability_data.values())
            }).sort_values("Batch").reset_index(drop=True)

            fr_min = min(all_friability_values) if all_friability_values else np.nan
            fr_max = max(all_friability_values) if all_friability_values else np.nan
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
            st.info("Tidak ada data Waktu Hancur yang diproses.")
            
        if not friability_df.empty:
            st.write("Tabel Friability dengan Statistik:")
            st.dataframe(friability_df.style.format({"Friability": "{:.4f}", "Batch": "{}"}, na_rep=""))
        else:
            st.info("Tidak ada data Friability yang diproses.")
            
        return waktu_hancur_df, friability_df
    except Exception as e:
        st.error(f"Gagal memproses file Waktu Hancur dan Friability: {e}")
        st.exception(e)
        return pd.DataFrame(), pd.DataFrame()

def export_dataframe(df, filename="data_export"):
    output = io.BytesIO()
    # When exporting, we don't want the Styler proxy, but the actual DataFrame.
    # Also, ensure the index is handled as desired for export.
    # If the DataFrame has a meaningful index (like batch names or stats labels), export it.
    # If the index is just a range, it might be omitted (index=False).
    # For these tables, the first column often acts as a label, or the actual index is meaningful.
    
    # If 'Statistik' column was added and set as index for display (like in Tebal),
    # we might want to reset it for export or ensure the original df is passed.
    # The parsing functions should return the "clean" DataFrame for export.

    # Check if the first column is 'Statistik' and set as index for export
    # This is a bit tricky as 'df' here could be from various parsers.
    # Generally, the DataFrames passed here should be ready for export.
    # The `index=True` is default and usually good if index has meaning.

    df_to_export = df.copy()
    # If 'Statistik' is the name of the index, it will be exported as the first column.
    # If 'Statistik' is a regular column and index is a range, that's also fine.

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_to_export.to_excel(writer, index=True) # Keep index for most cases
    output.seek(0)
    b64 = base64.b64encode(output.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}.xlsx">📥 Download {filename.replace("_", " ").title()} Excel File</a>'
    return href

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
        horizontal=True
    )
    
    template_info = {
        "Kekerasan": "Template Excel untuk pengujian kekerasan",
        "Keseragaman Bobot": "Template Excel untuk pengujian keseragaman bobot",
        "Tebal": "Template Excel untuk pengujian tebal",
        "Waktu Hancur dan Friability": "Template Excel untuk pengujian waktu hancur dan friability (kolom 'Nomor Batch' dan 'Sample Data')"
    }
    
    st.info(f"Upload file Excel dengan format: {template_info[selected_option]}")
    
    uploaded_file = st.file_uploader("Upload file Excel sesuai template", type=["xlsx","ods"], key=f"uploader_{selected_option.replace(' ', '_')}") # Ensure unique key
    
    if uploaded_file:
        file_copy = io.BytesIO(uploaded_file.getvalue())
        st.success(f"File untuk pengujian {selected_option} berhasil diupload")
        st.subheader(f"Hasil Pengujian {selected_option}")
        
        df_result = None # For single df results
        df_result_wh = None # For waktu hancur
        df_result_fr = None # For friability

        if selected_option == "Kekerasan":
            df_result = parse_kekerasan_excel(file_copy)
            if df_result is not None and not df_result.empty:
                filename = "data_kekerasan"
                st.markdown(export_dataframe(df_result, filename), unsafe_allow_html=True)
                st.success(f"{filename.replace('_', ' ').title()} siap diunduh. Klik tombol di atas.")
            elif df_result is not None and df_result.empty: # Parser returned empty df (valid case, no data)
                 st.info("Tidak ada data valid yang ditemukan dalam file untuk diproses.")
            # If df_result is None, error was already shown by parser

        elif selected_option == "Keseragaman Bobot":
            df_result = parse_keseragaman_bobot_excel(file_copy)
            if df_result is not None and not df_result.empty:
                filename = "data_keseragaman_bobot"
                st.markdown(export_dataframe(df_result, filename), unsafe_allow_html=True)
                st.success(f"{filename.replace('_', ' ').title()} siap diunduh. Klik tombol di atas.")
            elif df_result is not None and df_result.empty:
                 st.info("Tidak ada data valid yang ditemukan dalam file untuk diproses.")

        elif selected_option == "Tebal":
            df_result = parse_tebal_excel(file_copy) # This should return the data for export
            if df_result is not None and not df_result.empty:
                filename = "data_tebal"
                st.markdown(export_dataframe(df_result, filename), unsafe_allow_html=True)
                st.success(f"{filename.replace('_', ' ').title()} siap diunduh. Klik tombol di atas.")
            elif df_result is not None and df_result.empty:
                 st.info("Tidak ada data valid yang ditemukan dalam file untuk diproses.")
        
        elif selected_option == "Waktu Hancur dan Friability":
            df_result_wh, df_result_fr = parse_waktu_hancur_friability_excel(file_copy)
            
            if df_result_wh is not None and not df_result_wh.empty:
                st.markdown(export_dataframe(df_result_wh, "data_waktu_hancur"), unsafe_allow_html=True)
                st.success("Data Waktu Hancur siap diunduh. Klik tombol di atas.")
            elif df_result_wh is not None and df_result_wh.empty : # Parser returned empty but valid df
                st.info("Tidak ada data Waktu Hancur yang valid untuk diunduh (mungkin tidak ada data atau semua data tidak memenuhi kriteria).")
            
            if df_result_fr is not None and not df_result_fr.empty:
                st.markdown(export_dataframe(df_result_fr, "data_friability"), unsafe_allow_html=True)
                st.success("Data Friability siap diunduh. Klik tombol di atas.")
            elif df_result_fr is not None and df_result_fr.empty:
                st.info("Tidak ada data Friability yang valid untuk diunduh (mungkin tidak ada data atau semua data tidak memenuhi kriteria).")

        # Option to display all data in a table (for single df results)
        if selected_option != "Waktu Hancur dan Friability":
            if df_result is not None and not df_result.empty:
                if st.checkbox("Tampilkan semua data dalam bentuk tabel (data yang akan diexport)", key=f"show_table_{selected_option.replace(' ', '_')}"):
                    st.write("Data Lengkap (termasuk statistik):")
                    # Use a simple display for the checkbox data, formatting is already done by parser or can be minimal
                    st.dataframe(df_result) 
            elif df_result is None and uploaded_file: # If parsing failed and returned None
                st.warning("Tidak ada tabel untuk ditampilkan karena terjadi kesalahan saat pemrosesan file atau file tidak valid.")


# Main app execution (if you run this script directly)
if __name__ == '__main__':
    # You would typically have a main app structure here if this is part of a larger app
    # For now, just calling tampilkan_ipc() to make it runnable as a standalone page
    tampilkan_ipc()
