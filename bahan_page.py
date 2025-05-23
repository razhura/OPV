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

                    st.session_state.result_df = result_df  # Simpan hasil ke session state
                    st.session_state.processed = True  # Tandai bahwa data telah diproses
                    
                    # Dapatkan daftar unik nama bahan
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

            # Tampilkan dropdown dan tabel terfilter jika data telah diproses
            if 'processed' in st.session_state and st.session_state.processed:
                st.subheader("🔍 Filter Data Berdasarkan Nama Bahan")
                
                # Dapatkan daftar unik nama bahan
                unique_bahan_names = st.session_state.unique_bahan_names
                
                # Tambahkan tombol "Pilih Semua"
                col1, col2 = st.columns([1, 4])
                with col1:
                    if st.button("Pilih Semua"):
                        st.session_state.selected_bahan_names = unique_bahan_names
                
                # Siapkan multiselect dengan default dari session state jika ada
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
                    # Untuk setiap nama bahan yang dipilih, buat tabel terfilter
                    for selected_name in selected_bahan_names:
                        # Buat tabel terfilter untuk nama bahan yang dipilih
                        filtered_df = create_filtered_table_by_name(st.session_state.result_df, selected_name)
                        
                        # Buat nama file yang aman untuk digunakan (tanpa karakter khusus)
                        safe_filename = re.sub(r'[^\w\s-]', '', selected_name).strip().replace(' ', '_')
                        
                        if not filtered_df.empty:
                            st.subheader(f"📊 Tabel Terfilter - {selected_name}")
                            st.dataframe(filtered_df)
                            
                            # Ekspor tabel terfilter ke CSV
                            csv_filtered = filtered_df.to_csv(index=False)
                            st.download_button(
                                label=f"📥 Download Tabel {selected_name} (CSV)",
                                data=csv_filtered,
                                file_name=f"filtered_{safe_filename}.csv",
                                mime="text/csv",
                                key=f"csv_{safe_filename}"  # Unique key untuk setiap button
                            )
                            
                            # Ekspor tabel terfilter ke Excel
                            buffer_filtered = io.BytesIO()
                            with pd.ExcelWriter(buffer_filtered, engine='openpyxl') as writer:
                                filtered_df.to_excel(writer, index=False, sheet_name='Filtered Data')
                            buffer_filtered.seek(0)
                            
                            st.download_button(
                                label=f"📥 Download Tabel {selected_name} (Excel)",
                                data=buffer_filtered,
                                file_name=f"filtered_{safe_filename}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"excel_{safe_filename}"  # Unique key untuk setiap button
                            )
                        else:
                            st.warning(f"Tidak ada data untuk {selected_name}")
                        
                        st.markdown("---")  # Separator between tables

        except Exception as e:
            st.error(f"Terjadi kesalahan saat ekstraksi data: {e}")


if __name__ == "__main__":
    tampilkan_bahan()
