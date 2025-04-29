import streamlit as st
import pandas as pd
import numpy as np
import re
from utils import combine_duplicate_columns
from header_parser import extract_multi_level_headers
from openpyxl import load_workbook

st.set_page_config(page_title="Excel QCA Parser", layout="wide")
st.title("📊 OPV KONIMEX")

def is_possible_decimal(value):
    if isinstance(value, str) and re.match(r'^\d+\.\d+$', value.strip()):
        return True
    return False

# Upload file Excel
uploaded_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    # --- Proses Data ---

    # Ambil header dari baris 4-6
    combined_headers = extract_multi_level_headers(uploaded_file, start_row=4, num_levels=3)

    # Baca data Excel (mulai dari baris ke-7)
    df = pd.read_excel(uploaded_file, skiprows=6, header=None)
    df.columns = combined_headers

    # Gabungkan kolom duplikat
    df = combine_duplicate_columns(df)

    # Deteksi dan konversi otomatis ke float
    for col in df.columns:
        sample = df[col].dropna().astype(str)
        decimal_like_count = sample.apply(is_possible_decimal).sum()
        if len(sample) > 0 and decimal_like_count / len(sample) > 0.3:
            df[col] = df[col].astype(str).str.replace(",", ".").astype(float)

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
