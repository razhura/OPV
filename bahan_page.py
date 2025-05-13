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
    grouped = df.groupby('Nomor Batch')

    transformed_rows = []
    max_items = 0

    for batch, group in grouped:
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


def create_filtered_table(df, selected_index):
    # Kolom yang akan dipertahankan
    columns_to_keep = [
        'Nomor Batch', 
        'No. Order Produksi', 
        'Jalur', 
        f'Nama Bahan {selected_index}',
        f'Kode Bahan {selected_index}',
        f'Kuantiti > Terpakai {selected_index}',
        f'Kuantiti > Rusak {selected_index}',
        f'No Lot Supplier {selected_index}',
        f'Label QC {selected_index}'
    ]
    
    # Filter kolom yang ada di dataframe
    available_columns = [col for col in columns_to_keep if col in df.columns]
    
    # Buat dataframe baru dengan kolom yang tersedia
    filtered_df = df[available_columns].copy()
    
    # Ganti nama kolom untuk menghilangkan nomor indeks
    new_column_names = {}
    for col in filtered_df.columns:
        if col not in ['Nomor Batch', 'No. Order Produksi', 'Jalur']:
            new_name = re.sub(r"\s\d+$", "", col)
            new_column_names[col] = new_name
    
    # Terapkan perubahan nama kolom
    filtered_df = filtered_df.rename(columns=new_column_names)
    
    # Hapus baris yang nama bahannya kosong
    filtered_df = filtered_df[filtered_df['Nama Bahan'].notna() & (filtered_df['Nama Bahan'] != '')]
    
    return filtered_df


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

                    # Cari berapa banyak kolom "Nama Bahan" yang ada
                    nama_bahan_columns = [col for col in result_df.columns if "Nama Bahan" in col]
                    
                    # Siapkan opsi untuk memilih indeks Nama Bahan
                    available_indices = [int(col.split()[-1]) for col in nama_bahan_columns]
                    
                    st.session_state.available_indices = available_indices  # Simpan indeks yang tersedia

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
                
                # Dapatkan indeks yang tersedia
                available_indices = st.session_state.available_indices
                
                # Tambahkan tombol "Pilih Semua"
                col1, col2 = st.columns([1, 4])
                with col1:
                    if st.button("Pilih Semua"):
                        st.session_state.selected_indices = available_indices
                
                # Siapkan multiselect dengan default dari session state jika ada
                if 'selected_indices' not in st.session_state:
                    st.session_state.selected_indices = []
                
                with col2:
                    selected_indices = st.multiselect(
                        "Pilih Indeks Nama Bahan:",
                        available_indices,
                        default=st.session_state.selected_indices,
                        format_func=lambda x: f"Nama Bahan {x}"
                    )
                    st.session_state.selected_indices = selected_indices
                
                if selected_indices:
                    # Untuk setiap indeks yang dipilih, buat tabel terfilter
                    for selected_index in selected_indices:
                        # Buat tabel terfilter untuk indeks yang dipilih
                        filtered_df = create_filtered_table(st.session_state.result_df, selected_index)
                        
                        st.subheader(f"📊 Tabel Terfilter - Nama Bahan {selected_index}")
                        st.dataframe(filtered_df)
                        
                        # Ekspor tabel terfilter ke CSV
                        csv_filtered = filtered_df.to_csv(index=False)
                        st.download_button(
                            label=f"📥 Download Tabel Terfilter Nama Bahan {selected_index} (CSV)",
                            data=csv_filtered,
                            file_name=f"filtered_nama_bahan_{selected_index}.csv",
                            mime="text/csv",
                            key=f"csv_{selected_index}"  # Unique key untuk setiap button
                        )
                        
                        # Ekspor tabel terfilter ke Excel
                        buffer_filtered = io.BytesIO()
                        with pd.ExcelWriter(buffer_filtered, engine='openpyxl') as writer:
                            filtered_df.to_excel(writer, index=False, sheet_name=f'Nama Bahan {selected_index}')
                        buffer_filtered.seek(0)
                        
                        st.download_button(
                            label=f"📥 Download Tabel Terfilter Nama Bahan {selected_index} (Excel)",
                            data=buffer_filtered,
                            file_name=f"filtered_nama_bahan_{selected_index}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"excel_{selected_index}"  # Unique key untuk setiap button
                        )
                        
                        st.markdown("---")  # Separator between tables

        except Exception as e:
            st.error(f"Terjadi kesalahan saat ekstraksi data: {e}")


if __name__ == "__main__":
    tampilkan_bahan()
