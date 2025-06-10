import pandas as pd
import numpy as np
import io
import streamlit as st
from openpyxl import load_workbook
import re
import os



def filter_labelqc():
    st.title("📤 UPLOAD HASIL JADI DARI CPpP BAHAN")
    uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx", "csv"])

    if uploaded_file is not None:
        try:
            # Baca file
            if uploaded_file.name.endswith('.csv'):
                df_asli = pd.read_csv(uploaded_file)
            else:
                df_asli = pd.read_excel(uploaded_file)

            df_asli.columns = df_asli.columns.str.strip()
            st.success("✅ File berhasil dimuat.")
            st.subheader("📄 Data Excel Asli")
            st.dataframe(df_asli)

            # Temukan semua pasangan kolom "Kode Bahan.X" dan "Label QC.X"
            kode_bahan_pairs = []
            for col in df_asli.columns:
                if col.startswith("Kode Bahan"):
                    suffix = col.split("Kode Bahan")[-1]  # bisa '', '.1', '.2', dsb
                    label_qc_col = "Label QC" + suffix
                    if label_qc_col in df_asli.columns:
                        kode_bahan_pairs.append((col, label_qc_col))

            # Cari kolom No Batch jika ada
            batch_cols = []
            for col in df_asli.columns:
                if "Batch" in col or "No Batch" in col or "Nomor Batch" in col:
                    batch_cols.append(col)

            # Jika tidak ada kolom batch yang spesifik, tanyakan ke pengguna
            if not batch_cols:
                st.warning("⚠️ Kolom Batch tidak ditemukan secara otomatis.")
                all_cols = list(df_asli.columns)
                batch_cols = st.multiselect(
                    "Pilih kolom yang berisi informasi Batch:", 
                    all_cols,
                    key="batch_column_selector"
                )

            # Gabungkan semua kode bahan menjadi satu list unik
            all_kode_bahan = pd.Series(dtype=str)
            for kode_col, _ in kode_bahan_pairs:
                all_kode_bahan = pd.concat([ 
                    all_kode_bahan, 
                    df_asli[kode_col].dropna().astype(str).apply(lambda x: x.strip()) 
                ])

            kode_bahan_list = sorted(all_kode_bahan.dropna().unique())
            
            # Buat dataframe yang berisi semua pasangan Kode Bahan dan Label QC
            all_data = []
            for kode_col, label_col in kode_bahan_pairs:
                valid_rows = df_asli[[kode_col, label_col] + batch_cols].dropna(subset=[kode_col])
                for _, row in valid_rows.iterrows():
                    batch_info = {batch_col: row[batch_col] if pd.notna(row[batch_col]) else "" for batch_col in batch_cols}
                    all_data.append({
                        "Kode Bahan": str(row[kode_col]).strip(),
                        "Label QC": row[label_col] if pd.notna(row[label_col]) else "",
                        **batch_info
                    })
            
            # Buat dataframe untuk semua data
            complete_df = pd.DataFrame(all_data)
            
            # Buat ringkasan untuk semua kode bahan
            grouped_all_df = (
                complete_df
                .drop_duplicates()
                .groupby("Kode Bahan")["Label QC"]
                .unique()
                .reset_index()
            )
            grouped_all_df["Label QC"] = grouped_all_df["Label QC"].apply(lambda x: ", ".join(sorted([str(item) for item in x if str(item).strip()])))
            
            # Tampilkan ringkasan untuk semua kode bahan
            st.subheader("🧾 Ringkasan Label QC untuk Semua Kode Bahan")
            st.dataframe(grouped_all_df)
            
            # Fitur Download Ringkasan untuk semua kode bahan
            if not grouped_all_df.empty:
                def to_excel(df):
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        df.to_excel(writer, index=False, sheet_name="Label QC")
                    output.seek(0)
                    return output

                # Tambahan fungsi ekspor dengan warna
                def to_excel_with_color(df, color_column="Label QC"):
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        df.to_excel(writer, index=False, sheet_name="Label QC")
                        wb = writer.book
                        ws = writer.sheets["Label QC"]
                
                        from openpyxl.styles import PatternFill
                
                        col_idx = df.columns.get_loc(color_column) + 1
                
                        label_colors = {}
                        for row_idx, val in enumerate(df[color_column], start=2):
                            label_str = str(val).strip()
                            if label_str not in label_colors:
                                hash_val = hash(label_str) % 200
                                lightness = 30 + hash_val % 50
                                blue_rgb = int((255 * lightness) / 100)
                                hex_color = f"FF{blue_rgb:02X}{blue_rgb:02X}FF"
                                label_colors[label_str] = PatternFill(
                                    start_color=hex_color,
                                    end_color=hex_color,
                                    fill_type="solid"
                                )
                            ws.cell(row=row_idx, column=col_idx).fill = label_colors[label_str]
                    output.seek(0)
                    return output
                
                excel_all_grouped = to_excel(grouped_all_df)
                st.download_button(
                    label="📥 Download Ringkasan Semua Label QC (Excel)",
                    data=excel_all_grouped,
                    file_name="ringkasan_semua_label_qc.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            # Fitur filter untuk detail kode bahan tertentu
            st.header("🔍 Filter Berdasarkan Kode Bahan")
            selected_kode = st.selectbox("Pilih Kode Bahan untuk Detail", kode_bahan_list)

            # Filter berdasarkan pasangan kode dan label yang sesuai
            hasil_data = []

            for kode_col, label_col in kode_bahan_pairs:
                mask = df_asli[kode_col].astype(str).str.strip() == selected_kode
                filtered_rows = df_asli[mask]
                for _, row in filtered_rows.iterrows():
                    batch_info = {batch_col: row[batch_col] if pd.notna(row[batch_col]) else "" for batch_col in batch_cols}
                    hasil_data.append({
                        "Kode Bahan": selected_kode,
                        "Label QC": row[label_col] if pd.notna(row[label_col]) else "",
                        **batch_info
                    })

            hasil_df = pd.DataFrame(hasil_data)

            st.subheader(f"🏷️ Detail Label QC untuk Kode Bahan: {selected_kode}")
            st.dataframe(hasil_df)

            # Fitur Download Dataframe ke Excel
            if not hasil_df.empty:
                excel_data = to_excel(hasil_df)
                st.download_button(
                    label="📥 Download Detail Label QC (Excel)",
                    data=excel_data,
                    file_name=f"detail_label_qc_{selected_kode}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # FITUR BARU: Filter berdasarkan Label QC (MULTIPLE SELECTION)
            st.header("🔍 Filter Berdasarkan Label QC")
            
            # Dapatkan semua label QC unik
            all_labels = complete_df["Label QC"].dropna().astype(str).unique()
            all_labels = sorted([label for label in all_labels if label.strip()])
            
            # Tambahkan checkbox untuk "Pilih Semua"
            select_all = st.checkbox("Pilih Semua Label QC")
            
            # Buat multiselect dengan default semua terpilih jika select_all dicentang
            if select_all:
                default_selection = all_labels
            else:
                default_selection = []
                
            # Ubah dari selectbox menjadi multiselect untuk memilih lebih dari satu Label QC
            selected_labels = st.multiselect(
                "Pilih Label QC untuk Melihat Batch", 
                all_labels,
                default=default_selection
            )
            
            if selected_labels:
                # Filter data berdasarkan label QC yang dipilih (multiple)
                label_filtered_df = complete_df[complete_df["Label QC"].astype(str).isin(selected_labels)]
                
                if label_filtered_df.empty:
                    st.warning(f"Tidak ada data dengan Label QC yang dipilih")
                else:
                    # Konversi kolom batch ke string untuk menghindari masalah sorting dengan NaN values
                    for batch_col in batch_cols:
                        if batch_col in label_filtered_df.columns:
                            label_filtered_df[batch_col] = label_filtered_df[batch_col].fillna("").astype(str)
                    
                    # Urutkan hasil sesuai permintaan: Nomor Batch, Kode Bahan, Label QC
                    sort_columns = batch_cols + ["Kode Bahan", "Label QC"]
                    label_filtered_df = label_filtered_df.sort_values(by=sort_columns)
                    
                    # Tampilkan judul dengan label yang dipilih
                    label_list_str = ", ".join(selected_labels)
                    st.subheader(f"📋 Batch dengan Label QC: {label_list_str}")
                    
                    # Reorganisasi kolom untuk menampilkan dengan urutan: Nomor Batch > Kode Bahan > Label QC
                    column_order = batch_cols + ["Kode Bahan", "Label QC"]
                    label_filtered_df = label_filtered_df[column_order]
                    
                    # Tampilkan hasil filter dengan urutan kolom yang sudah diatur
                    st.dataframe(label_filtered_df)
                    
                    # Opsi untuk pewarnaan Label QC di file Excel
                    warna_excel = st.checkbox("🎨 Warnai kolom Label QC di file Excel")
                    
                    # Tentukan fungsi ekspor berdasarkan pilihan user
                    if warna_excel:
                        excel_label_data = to_excel_with_color(label_filtered_df)
                    else:
                        excel_label_data = to_excel(label_filtered_df)

                    # Buat nama file yang sesuai dengan label yang dipilih
                    if len(selected_labels) == 1:
                        filename = f"batch_dengan_label_qc_{selected_labels[0]}.xlsx"
                    else:
                        # Jika multi-label, gunakan 'multiple_labels'
                        filename = f"batch_dengan_multiple_label_qc.xlsx"
                    
                    st.download_button(
                        label="📥 Download Data dengan Label QC ini (Excel)",
                        data=excel_label_data,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except Exception as e:
            st.error(f"❌ Terjadi kesalahan saat membaca file: {e}")

if __name__ == "__main__":
    filter_labelqc()
