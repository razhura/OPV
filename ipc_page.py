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
        
        result_df.index = range(1, len(result_df) + 1) # BARIS INI MUNGKIN SUDAH ADA/MIRIP
        
        stats_df = calculate_statistics(result_df) # BARIS INI SUDAH ADA
        exportable_df = pd.concat([result_df, stats_df]) # BARIS INI SUDAH ADA

        # Untuk display, tambahkan label statistik sebagai kolom pertama
        display_df = exportable_df.copy() # BARIS INI SUDAH ADA
      # Bagian ini mungkin sedikit berbeda di kode Anda, sesuaikan:
        # stat_labels = [""] * len(result_df) + list(stats_df.index) # VERSI LAMA ANDA
        # display_df.insert(0, "Keterangan", stat_labels) # VERSI LAMA ANDA

      # --- AWAL BAGIAN YANG PERLU DIMODIFIKASI/DITAMBAHKAN ---
      # Buat kolom "Keterangan" dari index exportable_df
        keterangan_col_values = []
        for idx_val in display_df.index: # display_df.index sama dengan exportable_df.index
            if isinstance(idx_val, str): # Jika indexnya string (misal 'MIN', 'MAX')
                keterangan_col_values.append(idx_val)
            else: # Jika indexnya angka (misal 1, 2, 3 dari result_df)
                keterangan_col_values.append(str(idx_val)) # Ubah jadi string, atau "" jika ingin kosong

        display_df.insert(0, "Keterangan", keterangan_col_values)
      # Reset index agar tampilan lebih bersih karena "Keterangan" sudah jadi penjelas baris
        display_df = display_df.reset_index(drop=True)
      # --- AKHIR BAGIAN YANG PERLU DIMODIFIKASI/DITAMBAHKAN ---

        st.write("Data Tebal Terstruktur dengan Statistik:") # BARIS INI SUDAH ADA
        
      # --- AWAL BAGIAN YANG PERLU DIMODIFIKASI (BARIS st.dataframe) ---
      # Kode LAMA Anda yang menyebabkan error:
      # st.dataframe(display_df.style.format("{:.4f}", na_rep="").set_properties(**{'text-align': 'left'}).set_table_styles([dict(selector='th', props=[('text-align', 'left')])]))
      
      # Kode BARU (SOLUSI):
        numeric_cols = [col for col in display_df.columns if col != "Keterangan"]
        formatter = {col: "{:.4f}" for col in numeric_cols}
        formatter["Keterangan"] = "{}" # Format kolom "Keterangan" sebagai string

        st.dataframe(
            display_df.style.format(formatter, na_rep="")
            .set_properties(**{'text-align': 'left'})
            .set_table_styles([dict(selector='th', props=[('text-align', 'left')])])
        )
      # --- AKHIR BAGIAN YANG PERLU DIMODIFIKASI (BARIS st.dataframe) ---
        
        return exportable_df # BARIS INI SUDAH ADA

    except Exception as e:
        st.error(f"Gagal memproses file Tebal: {e}")
        st.exception(e)
        return None
