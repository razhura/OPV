import pandas as pd
import numpy as np
import io
import streamlit as st
from openpyxl import load_workbook
import re   


def extract_headers_from_rows_10_and_11(excel_file):
    wb = load_workbook(excel_file, data_only=True)
    ws = wb.active

    headers = []
    seen = {}
    max_col = ws.max_column

    for col in range(1, max_col + 1):
        cell_10 = ws.cell(row=1, column=col)
        cell_11 = ws.cell(row=2, column=col)

        for merged_range in ws.merged_cells.ranges:
            if cell_10.coordinate in merged_range:
                cell_10 = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                break
        for merged_range in ws.merged_cells.ranges:
            if cell_11.coordinate in merged_range:
                cell_11 = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                break

        val_10 = str(cell_10.value).strip() if cell_10.value else ""
        val_11 = str(cell_11.value).strip() if cell_11.value else ""

        if not val_11 or val_10 == val_11 or col <= 2:
            header = val_10
        else:
            header = f"{val_10} > {val_11}"

        if header in seen:
            seen[header] += 1
            header = f"{header}_{seen[header]}"
        else:
            seen[header] = 1

        headers.append(header)

    return headers


def normalize_columns(df):
    mapping = {
        'Nomor Batch': 'Nomor Batch',
        'No. Order Produksi': 'No. Order Produksi',
        'Jalur': 'Jalur',
        'Kode Bahan': 'Kode Bahan',
        'Nama Bahan': 'Nama Bahan',
        'Kuantiti > Terpakai': 'Kuantiti > Terpakai',
        'Kuantiti > Rusak': 'Kuantiti > Rusak',
        'No Lot Supplier': 'No Lot Supplier',
        'Label QC': 'Label QC'
    }


    from difflib import get_close_matches

    new_columns = {}
    for expected_col in mapping:
        matches = get_close_matches(expected_col, df.columns, n=1, cutoff=0.6)
        if matches:
            new_columns[matches[0]] = mapping[expected_col]

    df = df.rename(columns=new_columns)
    return df



def transform_batch_data(df):
    selected_cols = [
        'Nomor Batch',
        'No. Order Produksi',
        'Jalur',
        'Kode Bahan',
        'Nama Bahan',
        'Kuantiti > Terpakai',
        'Kuantiti > Rusak',
        'No Lot Supplier',
        'Label QC'
    ]

    missing = [col for col in selected_cols if col not in df.columns]
    if missing:
        raise ValueError(f"Kolom berikut tidak ditemukan dalam data: {missing}")

    df = df[selected_cols].copy()
    
    # Dapatkan urutan unik batch berdasarkan kemunculan pertama dalam data asli
    batch_order = df['Nomor Batch'].drop_duplicates().tolist()
    
    # Group berdasarkan batch, tapi pertahankan urutan asli
    grouped = df.groupby('Nomor Batch', sort=False)

    transformed_rows = []
    max_items = 0

    # Proses batch sesuai urutan kemunculan asli
    for batch in batch_order:
        if batch in grouped.groups:
            group = grouped.get_group(batch)
            
            # Ambil No. Order Produksi dan Jalur dari baris pertama kelompok
            order_produksi = group.iloc[0]['No. Order Produksi']
            jalur = group.iloc[0]['Jalur']

            row_data = [batch, order_produksi, jalur]

            for _, item in group.iterrows():
                row_data.extend([
                    item['Kode Bahan'],
                    item['Nama Bahan'],
                    item['Kuantiti > Terpakai'],
                    item['Kuantiti > Rusak'],
                    item['No Lot Supplier'],
                    item['Label QC']
                ])

            max_items = max(max_items, len(group))
            transformed_rows.append(row_data)

    full_row_len = 3 + max_items * 6
    for row in transformed_rows:
        row.extend([''] * (full_row_len - len(row)))

    headers = ['Nomor Batch', 'No. Order Produksi', 'Jalur']
    for i in range(1, max_items + 1):
        headers.extend([
            f"Kode Bahan {i}",
            f"Nama Bahan {i}",
            f"Kuantiti > Terpakai {i}",
            f"Kuantiti > Rusak {i}",
            f"No Lot Supplier {i}",
            f"Label QC {i}"
        ])

    return pd.DataFrame(transformed_rows, columns=headers)


def simplify_headers(df):
    # Hapus penomoran di akhir kolom seperti "Kode Bahan 1" → "Kode Bahan"
    new_cols = []
    for col in df.columns:
        if col == "Nomor Batch":
            new_cols.append(col)
        else:
            # Hilangkan angka dan spasi di akhir, tapi simpan seluruh bagian awal
            simplified = re.sub(r"\s\d+$", "", col)
            new_cols.append(simplified)
    df.columns = new_cols
    return df


def create_filtered_table_by_name(df, selected_name):
    # Temukan semua kolom "Nama Bahan" di dataframe
    nama_bahan_cols = [col for col in df.columns if col.startswith('Nama Bahan ')]
    
    # Dapatkan indeks yang sesuai dengan nama bahan yang dipilih
    filtered_indices = []
    for col in nama_bahan_cols:
        # Dapatkan indeks dari nama kolom, misalnya "Nama Bahan 1" → 1
        index = int(col.split()[-1])
        
        # Periksa setiap baris untuk nilai yang cocok dengan nama bahan yang dipilih
        mask = df[col] == selected_name
        # Jika ada baris yang cocok, tambahkan indeks ini ke daftar
        if mask.any():
            filtered_indices.append(index)
    
    # Gabungkan semua dataframe terfilter untuk setiap indeks yang ditemukan
    filtered_dfs = []
    for index in filtered_indices:
        # Kolom yang akan dipertahankan
        columns_to_keep = [
            'Nomor Batch', 
            'No. Order Produksi', 
            'Jalur', 
            f'Nama Bahan {index}',
            f'Kode Bahan {index}',
            f'Kuantiti > Terpakai {index}',
            f'Kuantiti > Rusak {index}',
            f'No Lot Supplier {index}',
            f'Label QC {index}'
        ]
        
        # Filter kolom yang ada di dataframe
        available_columns = [col for col in columns_to_keep if col in df.columns]
        
        # Buat dataframe baru dengan kolom yang tersedia
        temp_df = df[available_columns].copy()
        
        # Filter baris dimana nama bahan sesuai dengan yang dipilih
        temp_df = temp_df[temp_df[f'Nama Bahan {index}'] == selected_name]
        
        # Ganti nama kolom untuk menghilangkan nomor indeks
        new_column_names = {}
        for col in temp_df.columns:
            if col not in ['Nomor Batch', 'No. Order Produksi', 'Jalur']:
                new_name = re.sub(r"\s\d+$", "", col)
                new_column_names[col] = new_name
        
        # Terapkan perubahan nama kolom
        temp_df = temp_df.rename(columns=new_column_names)
        
        # Tambahkan ke daftar dataframe terfilter
        if not temp_df.empty:
            filtered_dfs.append(temp_df)
    
    # Gabungkan semua dataframe terfilter
    if filtered_dfs:
        return pd.concat(filtered_dfs, ignore_index=True)
    else:
        # Jika tidak ada yang cocok, kembalikan dataframe kosong dengan kolom yang sesuai
        return pd.DataFrame(columns=['Nomor Batch', 'No. Order Produksi', 'Jalur', 
                                    'Nama Bahan', 'Kode Bahan', 'Kuantiti > Terpakai', 
                                    'Kuantiti > Rusak', 'No Lot Supplier', 'Label QC'])


def get_unique_bahan_names(df):
    # Temukan semua kolom "Nama Bahan" di dataframe
    nama_bahan_cols = [col for col in df.columns if col.startswith('Nama Bahan ')]
    
    # Kumpulkan semua nilai unik dari kolom-kolom tersebut
    unique_names = set()
    for col in nama_bahan_cols:
        # Hanya tambahkan nilai yang tidak null/NaN dan bukan string kosong
        values = df[col].dropna()
        values = values[values != '']
        unique_names.update(values)
    
    # Kembalikan sebagai list yang diurutkan
    return sorted(list(unique_names))


def merge_same_materials(df):
    """
    Memindahkan kelompok data dengan kode bahan yang sama ke baris baru
    Jika dalam satu baris ada kode bahan yang sama di kelompok berbeda,
    kelompok kedua akan dipindah ke baris baru (tanpa nomor batch, no order, jalur)
    """
    import pandas as pd
    
    # Buat list untuk menyimpan semua baris hasil
    result_rows = []
    
    # Dapatkan semua kolom kode bahan
    kode_bahan_cols = [col for col in df.columns if col.startswith('Kode Bahan ')]
    
    # Dapatkan indeks dari nama kolom
    indices = []
    for col in kode_bahan_cols:
        try:
            index = int(col.split()[-1])
            indices.append(index)
        except:
            continue
    
    indices.sort()
    
    # Untuk setiap baris dalam dataframe asli
    for row_idx in df.index:
        # Kumpulkan semua kelompok data bahan dalam baris ini
        materials_groups = []
        
        for idx in indices:
            kode_col = f'Kode Bahan {idx}'
            nama_col = f'Nama Bahan {idx}'
            
            # Periksa apakah ada data kode bahan
            if (kode_col in df.columns and 
                pd.notna(df.loc[row_idx, kode_col]) and 
                str(df.loc[row_idx, kode_col]).strip() != ''):
                
                # Kumpulkan semua data dalam kelompok ini
                group_data = {
                    'original_index': idx,
                    'kode': str(df.loc[row_idx, kode_col]).strip(),
                    'nama': df.loc[row_idx, nama_col] if nama_col in df.columns else '',
                    'terpakai': df.loc[row_idx, f'Kuantiti > Terpakai {idx}'] if f'Kuantiti > Terpakai {idx}' in df.columns else '',
                    'rusak': df.loc[row_idx, f'Kuantiti > Rusak {idx}'] if f'Kuantiti > Rusak {idx}' in df.columns else '',
                    'lot': df.loc[row_idx, f'No Lot Supplier {idx}'] if f'No Lot Supplier {idx}' in df.columns else '',
                    'qc': df.loc[row_idx, f'Label QC {idx}'] if f'Label QC {idx}' in df.columns else ''
                }
                
                materials_groups.append(group_data)
        
        if not materials_groups:
            # Jika tidak ada data material, copy baris asli
            result_rows.append(df.loc[row_idx].copy())
            continue
        
        # Identifikasi duplikasi kode bahan (100% sama)
        seen_codes = {}
        groups_to_keep = []  # Kelompok yang tetap di baris asli
        groups_to_move = []  # Kelompok yang akan dipindah ke baris baru
        
        for group in materials_groups:
            kode = group['kode'].strip()  # Kode bahan exact match
            if kode in seen_codes:
                # Kode bahan sudah ada sebelumnya, tandai untuk dipindah
                groups_to_move.append(group)
            else:
                # Kode bahan pertama kali muncul, tetap di baris asli
                seen_codes[kode] = True
                groups_to_keep.append(group)
        
        # Jika tidak ada duplikasi, semua tetap di baris asli
        if not groups_to_move:
            result_rows.append(df.loc[row_idx].copy())
            continue
        
        # Buat baris asli dengan kelompok yang tidak dipindah
        current_row = df.loc[row_idx].copy()
        
        # Kosongkan semua kolom bahan terlebih dahulu
        for idx in indices:
            for col_type in ['Nama Bahan', 'Kode Bahan', 'Kuantiti > Terpakai', 'Kuantiti > Rusak', 'No Lot Supplier', 'Label QC']:
                col_name = f'{col_type} {idx}'
                if col_name in current_row.index:
                    current_row[col_name] = ''
        
        # Isi kembali dengan kelompok yang tersisa, bergeser ke kiri mulai dari posisi 1
        for new_idx, group in enumerate(groups_to_keep, 1):
            if new_idx <= len(indices):
                # Isi semua data kelompok
                if f'Nama Bahan {new_idx}' in current_row.index:
                    current_row[f'Nama Bahan {new_idx}'] = group['nama']
                if f'Kode Bahan {new_idx}' in current_row.index:
                    current_row[f'Kode Bahan {new_idx}'] = group['kode']
                if f'Kuantiti > Terpakai {new_idx}' in current_row.index:
                    current_row[f'Kuantiti > Terpakai {new_idx}'] = group['terpakai']
                if f'Kuantiti > Rusak {new_idx}' in current_row.index:
                    current_row[f'Kuantiti > Rusak {new_idx}'] = group['rusak']
                if f'No Lot Supplier {new_idx}' in current_row.index:
                    current_row[f'No Lot Supplier {new_idx}'] = group['lot']
                if f'Label QC {new_idx}' in current_row.index:
                    current_row[f'Label QC {new_idx}'] = group['qc']
        
        result_rows.append(current_row)
        
        # Buat baris baru untuk setiap kelompok yang dipindah
        for group in groups_to_move:
            # Buat baris kosong berdasarkan struktur dataframe asli
            new_row = pd.Series(index=df.columns, dtype=object)
            
            # Kosongkan semua kolom (termasuk nomor batch, no order, jalur)
            for col in new_row.index:
                new_row[col] = ''
            
            # Cari posisi kelompok dengan kode bahan yang sama di groups_to_keep
            target_position = None
            for kept_group in groups_to_keep:
                if kept_group['kode'].strip() == group['kode'].strip():
                    # Cari posisi kelompok ini di baris yang sudah diatur ulang
                    for pos, check_group in enumerate(groups_to_keep, 1):
                        if check_group == kept_group:
                            target_position = pos
                            break
                    break
            
            # Jika tidak ditemukan posisi yang sama, letakkan di posisi aslinya
            if target_position is None:
                target_position = group['original_index']
            
            # Isi data kelompok yang dipindah di posisi yang sesuai
            if f'Nama Bahan {target_position}' in new_row.index:
                new_row[f'Nama Bahan {target_position}'] = group['nama']
            if f'Kode Bahan {target_position}' in new_row.index:
                new_row[f'Kode Bahan {target_position}'] = group['kode']
            if f'Kuantiti > Terpakai {target_position}' in new_row.index:
                new_row[f'Kuantiti > Terpakai {target_position}'] = group['terpakai']
            if f'Kuantiti > Rusak {target_position}' in new_row.index:
                new_row[f'Kuantiti > Rusak {target_position}'] = group['rusak']
            if f'No Lot Supplier {target_position}' in new_row.index:
                new_row[f'No Lot Supplier {target_position}'] = group['lot']
            if f'Label QC {target_position}' in new_row.index:
                new_row[f'Label QC {target_position}'] = group['qc']
            
            result_rows.append(new_row)
    
    # Buat DataFrame baru dari hasil
    result_df = pd.DataFrame(result_rows)
    result_df.reset_index(drop=True, inplace=True)
    
    return result_df


# Tambahkan fungsi ini ke dalam fungsi tampilkan_bahan()
def tampilkan_bahan():
    st.title("Halaman CPP BAHAN")
    st.write("Ini adalah tampilan khusus CPP BAHAN.")

    uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx"])

    if uploaded_file is not None:
        combined_headers = extract_headers_from_rows_10_and_11(uploaded_file)
        df_asli = pd.read_excel(uploaded_file, skiprows=2, header=None)
        df_asli.columns = combined_headers

        try:
            st.subheader("📄 Data Excel Asli")
            st.dataframe(df_asli)
            st.info(f"Kolom yang terdeteksi: {', '.join(df_asli.columns.tolist())}")

            if st.button("🔍 Ekstrak Data Batch"):
                with st.spinner("Memproses data..."):
                    df_asli = normalize_columns(df_asli)
                    result_df = transform_batch_data(df_asli)

                    st.session_state.result_df = result_df
                    st.session_state.processed = True
                    
                    unique_bahan_names = get_unique_bahan_names(result_df)
                    st.session_state.unique_bahan_names = unique_bahan_names

                    st.subheader("🔢 Hasil Ekstraksi Data Batch")
                    st.dataframe(result_df)

                    # Ekspor CSV
                    csv_df = simplify_headers(result_df.copy())
                    csv = csv_df.to_csv(index=False)
                    st.download_button(
                        label="📥 Download Data Hasil Ekstraksi (CSV)",
                        data=csv,
                        file_name="data_batch_extracted.csv",
                        mime="text/csv"
                    )

                    # Ekspor Excel
                    excel_df = simplify_headers(result_df.copy())
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        excel_df.to_excel(writer, index=False, sheet_name='Batch Data')
                    buffer.seek(0)

                    st.download_button(
                        label="📥 Download Data Hasil Ekstraksi (Excel)",
                        data=buffer,
                        file_name="data_batch_extracted.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            # Tampilkan tombol merge jika data telah diproses
            if 'processed' in st.session_state and st.session_state.processed:
                # Tambahkan tombol untuk merge data bahan yang sama
                if st.button("🔄 Kelompokkan Bahan yang Sama"):
                    with st.spinner("Mengelompokkan data bahan yang sama..."):
                        merged_df = merge_same_materials(st.session_state.result_df)
                        st.session_state.result_df = merged_df
                        
                        # Update unique bahan names
                        unique_bahan_names = get_unique_bahan_names(merged_df)
                        st.session_state.unique_bahan_names = unique_bahan_names
                        
                        st.subheader("✅ Data Setelah Pengelompokan Bahan")
                        st.dataframe(merged_df)
                        st.success("Data bahan yang sama telah dikelompokkan!")
                        
                        # Ekspor hasil merge
                        csv_merged = simplify_headers(merged_df.copy()).to_csv(index=False)
                        st.download_button(
                            label="📥 Download Data Terkelompok (CSV)",
                            data=csv_merged,
                            file_name="data_batch_merged.csv",
                            mime="text/csv",
                            key="csv_merged"
                        )
                        
                        buffer_merged = io.BytesIO()
                        with pd.ExcelWriter(buffer_merged, engine='openpyxl') as writer:
                            simplify_headers(merged_df.copy()).to_excel(writer, index=False, sheet_name='Merged Data')
                        buffer_merged.seek(0)
                        
                        st.download_button(
                            label="📥 Download Data Terkelompok (Excel)",
                            data=buffer_merged,
                            file_name="data_batch_merged.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="excel_merged"
                        )

                st.subheader("🔍 Filter Data Berdasarkan Nama Bahan")
                
                unique_bahan_names = st.session_state.unique_bahan_names
                
                col1, col2 = st.columns([1, 4])
                with col1:
                    if st.button("Pilih Semua"):
                        st.session_state.selected_bahan_names = unique_bahan_names
                
                if 'selected_bahan_names' not in st.session_state:
                    st.session_state.selected_bahan_names = []
                
                with col2:
                    selected_bahan_names = st.multiselect(
                        "Pilih Nama Bahan:",
                        unique_bahan_names,
                        default=st.session_state.selected_bahan_names
                    )
                    st.session_state.selected_bahan_names = selected_bahan_names
                
                if selected_bahan_names:
                    for selected_name in selected_bahan_names:
                        filtered_df = create_filtered_table_by_name(st.session_state.result_df, selected_name)
                        safe_filename = re.sub(r'[^\w\s-]', '', selected_name).strip().replace(' ', '_')
                        
                        if not filtered_df.empty:
                            st.subheader(f"📊 Tabel Terfilter - {selected_name}")
                            st.dataframe(filtered_df)
                            
                            csv_filtered = filtered_df.to_csv(index=False)
                            st.download_button(
                                label=f"📥 Download Tabel {selected_name} (CSV)",
                                data=csv_filtered,
                                file_name=f"filtered_{safe_filename}.csv",
                                mime="text/csv",
                                key=f"csv_{safe_filename}"
                            )
                            
                            buffer_filtered = io.BytesIO()
                            with pd.ExcelWriter(buffer_filtered, engine='openpyxl') as writer:
                                filtered_df.to_excel(writer, index=False, sheet_name='Filtered Data')
                            buffer_filtered.seek(0)
                            
                            st.download_button(
                                label=f"📥 Download Tabel {selected_name} (Excel)",
                                data=buffer_filtered,
                                file_name=f"filtered_{safe_filename}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"excel_{safe_filename}"
                            )
                        else:
                            st.warning(f"Tidak ada data untuk {selected_name}")
                        
                        st.markdown("---")

        except Exception as e:
            st.error(f"Terjadi kesalahan saat ekstraksi data: {e}")


if __name__ == "__main__":
    tampilkan_bahan()
