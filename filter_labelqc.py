import pandas as pd
import numpy as np
import io
import streamlit as st
from openpyxl import load_workbook
import re
import os


def filter_labelqc():
    st.subheader("Upload File Ekstra Data Batch CPP Bahan")
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
                    import re
                    from openpyxl.styles import PatternFill
                    import colorsys
                    output = io.BytesIO()
                
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        df.to_excel(writer, index=False, sheet_name="Label QC")
                        wb = writer.book
                        ws = writer.sheets["Label QC"]
                
                        col_idx = df.columns.get_loc(color_column) + 1
                
                        for row_idx, val in enumerate(df[color_column], start=2):
                            label_str = str(val).strip().upper()
                
                            # Ambil angka & huruf dari label (contoh: "23A" → angka=23, huruf=A)
                            match = re.match(r"(\d+)([A-Z]?)", label_str)
                            if not match:
                                continue
                
                            angka = int(match.group(1))
                            huruf = match.group(2)
                
                            # === WARNA DASAR BERDASARKAN ANGKA ===
                            # Hue antara 190 (biru kehijauan) sampai 270 (biru keungu-unguan)
                            hue = 190 + (angka % 10) * 8  # hasilnya 190–270
                            # Lightness berdasarkan huruf: A=75%, B=70%, ..., Z=45%
                            lightness = 75 - (ord(huruf) - ord("A")) * 2.5 if huruf else 75
                            lightness = max(45, min(75, lightness))  # dibatasi biar ga terlalu gelap/terang
                            saturation = 0.9  # selalu 90% saturasi
                
                            # Konversi HSL ke RGB (0–255)
                            r, g, b = colorsys.hls_to_rgb(hue / 360, lightness / 100, saturation)
                            r = int(r * 255)
                            g = int(g * 255)
                            b = int(b * 255)
                            hex_color = f"FF{r:02X}{g:02X}{b:02X}"
                
                            # Apply warna ke cell
                            fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
                            ws.cell(row=row_idx, column=col_idx).fill = fill
                
                    output.seek(0)
                    return output

                # Tombol download Excel Kode Bahan
                def to_excel_styled(df):
                    from openpyxl.styles import PatternFill
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        df.to_excel(writer, index=False, sheet_name="Kode Bahan")
                        wb = writer.book
                        ws = writer.sheets["Kode Bahan"]
                
                        gray1 = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                        gray2 = PatternFill(start_color="BBBBBB", end_color="BBBBBB", fill_type="solid")
                
                        jumlah_idx = df.columns.get_loc("Jumlah Batch") + 1
                        fill = gray1
                
                        for row in range(2, len(df) + 2):
                            val = ws.cell(row=row, column=jumlah_idx).value
                            for col in range(1, len(df.columns) + 1):
                                ws.cell(row=row, column=col).fill = fill
                            if val:
                                fill = gray2 if fill == gray1 else gray1
                
                    output.seek(0)
                    return output
                
                excel_all_grouped = to_excel(grouped_all_df)
                st.download_button(
                    label="📥 Download Ringkasan Semua Label QC (Excel)",
                    data=excel_all_grouped,
                    file_name="ringkasan_semua_label_qc.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # Tampilkan semua data lengkap: Kode Bahan - Nomor Batch - Label QC
            st.header("📋 Semua Kode Bahan dengan Batch dan Label QC")
            
            batch_col_primary = batch_cols[0] if batch_cols else "Nomor Batch"
            summary_by_kode = (
                complete_df[["Kode Bahan", batch_col_primary, "Label QC"]]
                .drop_duplicates()
                .sort_values(by=["Kode Bahan", "Label QC", batch_col_primary])
            )

            # Tambahkan kolom prefix bulan dari nomor batch (misal: 'AUG24' dari 'AUG24A01')
            summary_by_kode["Prefix Bulan"] = summary_by_kode[batch_col_primary].str.extract(r'^([A-Z]{3}\d{2})')
            
            # Inisialisasi kolom Jumlah Batch kosong
            summary_by_kode["Jumlah Batch"] = ""
            
            # Grup berdasarkan: Kode Bahan, Label QC, dan Prefix Bulan
            grouped = summary_by_kode.groupby(["Kode Bahan", "Label QC", "Prefix Bulan"])
            
            # Isi jumlah batch hanya di baris terakhir per grup
            for _, group_indices in grouped.groups.items():
                group_rows = summary_by_kode.loc[group_indices]
                jumlah = group_rows[batch_col_primary].nunique()
                last_index = group_rows.index[-1]
                summary_by_kode.at[last_index, "Jumlah Batch"] = f"{jumlah} batch"
            
            # Hapus kolom bantu
            summary_by_kode = summary_by_kode.drop(columns=["Prefix Bulan"])
            # Urutkan ulang kolom biar rapi
            summary_by_kode = summary_by_kode[["Kode Bahan", "Label QC", batch_col_primary, "Jumlah Batch"]]
            
            st.dataframe(summary_by_kode)
            
            excel_summary = to_excel_styled(summary_by_kode)
            st.download_button(
                label="📥 Download Semua Data Kode Bahan (Excel)",
                data=excel_summary,
                file_name="semua_kode_bahan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            
            ##### FITUR BARU: Filter berdasarkan Label QC (MULTIPLE SELECTION) #####
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

def tampilkan_filter_labelqc():
    # st.title("Filter Data CPP Bahan")
    st.write("Ini adalah tampilan halaman Filter Data CPP Bahan")

    selected_option = st.radio(
        "Pilih jenis pengujian:",
        ["Filter Label QC", "Kuantiti"], 
        horizontal=True, 
        key="filter_qc_selection"  # <- pastikan key-nya unik
    )

    if selected_option == "Filter Label QC":
        filter_labelqc()
    elif selected_option == "Kuantiti":
        st.info("🔧 Fitur 'Kuantiti' masih dalam pengembangan.")

if __name__ == "__main__":
    tampilkan_filter_labelqc()

