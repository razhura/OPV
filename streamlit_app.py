import streamlit as st
import pandas as pd
import numpy as np
import re
from utils import combine_duplicate_columns
from header_parser import extract_multi_level_headers
from openpyxl import load_workbook
from navbar import render_navbar


st.set_page_config(page_title="Excel QCA Parser", layout="wide")
render_navbar()
st.title("📊 OPV KONIMEX V2.1")

# Fungsi parsing header bertingkat dari baris 4-6
def extract_multi_level_headers(excel_file, start_row=4, num_levels=3):
    wb = load_workbook(excel_file, data_only=True)
    ws = wb.active

    headers = []
    max_col = ws.max_column

    def simplify_main_header(header_text):
        if "-" in header_text:
            return header_text.split("-")[0].strip()
        return header_text.strip()

    for col in range(1, max_col + 1):
        levels = []
        for row in range(start_row, start_row + num_levels):
            cell = ws.cell(row=row, column=col)

            for merged_range in ws.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                    break

            value = str(cell.value).strip() if cell.value else ""
            levels.append(value)

        # Kolom A, B, C khusus ambil 1 level pertama
        if col <= 3:
            headers.append(levels[0])
        else:
            # Header utama disingkat
            simplified_main = simplify_main_header(levels[0])
            combined = " > ".join([simplified_main] + levels[1:])
            headers.append(combined)

    return headers



query_params = st.experimental_get_query_params()
page = query_params.get("page", ["QCA"])[0]

if page == "IPC":
    st.title("⚙️ In Process Control (IPC)")
    st.info("Halaman IPC masih dalam pengembangan.")
    st.stop()
else:
    st.title("📊 Critical Quality Attribute (QCA)")

# Upload file Excel
uploaded_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx", "ods"])

if uploaded_file is not None:
    # --- Proses Data ---

    # Ambil header dari baris 4-6
    combined_headers = extract_multi_level_headers(uploaded_file, start_row=4, num_levels=3)

    # Baca data Excel (mulai dari baris ke-7)
    df = pd.read_excel(uploaded_file, skiprows=6, header=None)
    df.columns = combined_headers
    
     # Deteksi dan konversi otomatis ke float
    for col in df.columns:
        sample = df[col].dropna().astype(str)

        decimal_like_count = sample.apply(lambda x: bool(re.match(r'^\d+([.,]\d+)?$', x.strip()))).sum()

        if len(sample) > 0 and decimal_like_count / len(sample) > 0.3:
            def safe_float(x):
                x = x.replace(",", ".").strip()
                if re.match(r'^\d+(\.\d+)?$', x):
                    return float(x)
                return pd.NA

            df[col] = sample.apply(safe_float)
    
    # Gabungkan kolom duplikat
    df = combine_duplicate_columns(df)

    # 2. Reset index
    df = df.reset_index(drop=True) 

    # 3. Perbaiki data kosong dalam batch
    current_batch = None
    batch_rows = []

    for idx in range(len(df)):
        batch_value = df.loc[idx, "Nomor Batch"]

        if pd.notna(batch_value):
            if batch_rows:
                for col in df.columns:
                    for row_idx in batch_rows:
                        if pd.isna(df.loc[row_idx, col]):
                            for search_idx in batch_rows:
                                if search_idx > row_idx and pd.notna(df.loc[search_idx, col]):
                                    df.loc[row_idx, col] = df.loc[search_idx, col]
                                    df.loc[search_idx, col] = None
                                    break
                batch_rows = []
            current_batch = batch_value
            batch_rows = [idx]
        else:
            if current_batch is not None:
                batch_rows.append(idx)

    if batch_rows:
        for col in df.columns:
            for row_idx in batch_rows:
                if pd.isna(df.loc[row_idx, col]):
                    for search_idx in batch_rows:
                        if search_idx > row_idx and pd.notna(df.loc[search_idx, col]):
                            df.loc[row_idx, col] = df.loc[search_idx, col]
                            df.loc[search_idx, col] = None
                            break

    # 4. Isi Nomor Batch kosong dari atas
    if "Nomor Batch" in df.columns:
        df["Nomor Batch"] = df["Nomor Batch"].fillna(method="ffill")

    # 5. Hapus baris kosong semua
    df = df.dropna(how="all")

    # 6. Hapus baris yang cuma punya Nomor Batch saja
    cols_to_check = [col for col in df.columns if col != "Nomor Batch"]
    df = df.dropna(subset=cols_to_check, how="all")


    # Fungsi bantu: deteksi string desimal (misal '1.19')
    def is_possible_decimal(value):
        if isinstance(value, str) and re.match(r'^\d+\.\d+$', value.strip()):
            return True
        return False

    # --- Tampilkan Data Awal ---
    st.subheader("📄 Data Akhir (Setelah Parsing dan Gabung Kolom):")
    st.dataframe(df)

    # --- Tombol untuk memilih fitur ---
    st.subheader("🔍 Pilih Fitur yang Ingin Digunakan")
    
    # Defaultnya fitur tersembunyi (None)
    feature_choice = st.radio(
        "Pilih fitur:",
        [None, "Pilih Kolom", "Pilih Batch"],
        format_func=lambda x: "Pilih fitur..." if x is None else x,
        index=0  # Default tidak ada yang dipilih
    )
    
    # === FITUR: PILIH KOLOM ===
    if feature_choice == "Pilih Kolom":
        st.subheader("🔍 Pilih Kolom yang Ingin Ditampilkan")
        
        # Jika "Nomor Batch" ada dalam kolom, jadikan default pilihan pertama
        default_columns = []
        if "Nomor Batch" in df.columns:
            default_columns = ["Nomor Batch"]
        
        selected_columns = st.multiselect("Pilih kolom:", df.columns, default=default_columns)

        if selected_columns:
            df_filtered = df[selected_columns]
            st.subheader("📄 Data dari Kolom yang Dipilih:")
            st.dataframe(df_filtered)

            # Statistik numerik
            st.subheader("📊 Statistik Ringkasan (Numerik)")
            numeric_cols = df_filtered.select_dtypes(include=np.number).columns.tolist()

            if numeric_cols:
                stats = df_filtered[numeric_cols].agg(['min', 'max', 'mean']).T
                stats.columns = ['Min', 'Max', 'Mean']
                st.dataframe(stats)
            else:
                st.info("Tidak ada kolom numerik dalam data yang difilter.")
        else:
            st.info("Silakan pilih minimal satu kolom untuk ditampilkan.")
    
    # === FITUR: PILIH BATCH ===
    elif feature_choice == "Pilih Batch":
        st.subheader("🎯 Pilih Batch dan Tampilkan Data Ke Kanan")

        # Lock kolom batch ke "Nomor Batch"
        batch_column = "Nomor Batch"
        
        # Periksa apakah kolom batch ada di dataframe
        if batch_column in df.columns:
            unique_batches = df[batch_column].dropna().unique()
            
            # Memungkinkan pemilihan multiple batch
            selected_batches = st.multiselect("Pilih Nilai Batch:", unique_batches)

            if selected_batches:
                # Filter data berdasarkan batch yang dipilih
                batch_rows = df[df[batch_column].isin(selected_batches)]

                if not batch_rows.empty:
                    batch_idx = df.columns.get_loc(batch_column)
                    selected_columns_right = df.columns[batch_idx:]

                    st.subheader(f"📄 Data dari Batch yang Dipilih dan Kolom Ke Kanan:")
                    st.dataframe(batch_rows[selected_columns_right])
                else:
                    st.warning("Batch yang dipilih tidak ditemukan di data.")
            else:
                st.info("Silakan pilih minimal satu batch untuk ditampilkan.")
        else:
            st.warning(f"Kolom '{batch_column}' tidak ditemukan di data. Pastikan kolom batch ada dalam file Excel.")
    elif feature_choice is None:
        st.info("Silakan pilih fitur yang ingin digunakan di atas.")

else:
    st.warning("⚠️ SILAKAN UPLOAD FILE TERLEBIH DAHULU.")
