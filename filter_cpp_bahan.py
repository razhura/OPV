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

#KUANTITI
def rapikan(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    batch_indices = df[df["Nomor Batch"].notna()].index.tolist()
    batch_indices.append(len(df))

    hasil = []

    for i in range(len(batch_indices) - 1):
        start = batch_indices[i]
        end = batch_indices[i + 1]
        blok = df.iloc[start:end].reset_index(drop=True)

        n_rows = len(blok)
        blok_baru = pd.DataFrame(columns=blok.columns)

        for col in blok.columns:
            if col == "Nomor Batch":
                isi = [blok.at[0, col]] + ["" for _ in range(n_rows - 1)]
            else:
                data_isi = blok[col].dropna().astype(str).str.strip()
                data_isi = data_isi[data_isi != ""]
                isi = data_isi.tolist() + [None] * (n_rows - len(data_isi))

            blok_baru[col] = isi

        hasil.append(blok_baru)

    df_bersih = pd.concat(hasil, ignore_index=True)

    # === PEMBERSIHAN BARIS KOSONG YANG LEBIH EFEKTIF ===
    
    # 1. Ganti string kosong dan whitespace dengan NaN (gunakan np.nan bukan pd.NA)
    df_bersih = df_bersih.replace(r'^\s*$', np.nan, regex=True)
    df_bersih = df_bersih.replace('', np.nan)
    
    # 2. Fungsi untuk mengecek apakah baris memiliki data bermakna
    def has_meaningful_data(row):
        # Jika ada Nomor Batch, cek apakah ada data lain yang tidak kosong
        has_batch = pd.notna(row["Nomor Batch"]) and str(row["Nomor Batch"]).strip() not in ["", "nan", "None"]
        
        # Cek kolom selain Nomor Batch
        other_cols = [col for col in row.index if col != "Nomor Batch"]
        has_other_data = any(
            pd.notna(row[col]) and str(row[col]).strip() not in ["", "nan", "None"] 
            for col in other_cols
        )
        
        # Baris valid jika: (punya batch DAN punya data lain) ATAU (tidak punya batch tapi punya data lain)
        return (has_batch and has_other_data) or (not has_batch and has_other_data)
    
    # 3. Filter hanya baris yang memiliki data bermakna
    df_final = df_bersih[df_bersih.apply(has_meaningful_data, axis=1)].reset_index(drop=True)
    
    # 4. Pembersihan final: hapus baris yang benar-benar kosong
    df_final = df_final.dropna(how='all')
    
    # 5. PENTING: Konversi semua pd.NA ke np.nan untuk kompatibilitas Excel
    df_final = df_final.fillna(np.nan)
    
    return df_final


def parse_kuantiti(value):
    """
    Fungsi untuk mengekstrak nilai numerik dari string kuantiti
    Contoh: "800 GRAM" -> 800.0, "1.600 GRAM" -> 1600.0
    """
    if pd.isna(value) or value == "":
        return 0.0
    
    # Konversi ke string dan bersihkan
    value_str = str(value).strip().upper()
    
    # Hapus unit (GRAM, KG, dll) dan ekstrak angka
    # Pola untuk menangkap angka dengan titik/koma sebagai pemisah ribuan
    import re
    number_pattern = r'(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?)'
    match = re.search(number_pattern, value_str)
    
    if match:
        number_str = match.group(1)
        # Ganti koma dengan titik dan hapus titik pemisah ribuan
        if ',' in number_str:
            # Jika ada koma, anggap sebagai pemisah desimal
            if number_str.count(',') == 1 and not number_str.endswith(','):
                number_str = number_str.replace('.', '').replace(',', '.')
            else:
                number_str = number_str.replace(',', '')
        
        try:
            return float(number_str)
        except ValueError:
            return 0.0
    
    return 0.0


def hitung_total_kuantiti(df_cleaned):
    """
    Fungsi untuk menghitung total kuantiti per bahan per batch dan per label QC
    """
    # Dapatkan semua kolom bahan
    bahan_cols = [col for col in df_cleaned.columns if col.startswith("Nama Bahan Formula")]
    
    # Siapkan data untuk perhitungan
    hasil_perhitungan = []
    
    for _, row in df_cleaned.iterrows():
        nomor_batch = row["Nomor Batch"]
        
        # Proses setiap kolom bahan
        for col in bahan_cols:
            if pd.notna(row[col]) and str(row[col]).strip() not in ["", "nan", "None"]:
                suffix = col.replace("Nama Bahan Formula", "")
                
                # Ambil data terkait
                nama_bahan = str(row[col]).strip()
                kode_bahan = str(row[f"Kode Bahan{suffix}"]).strip() if f"Kode Bahan{suffix}" in df_cleaned.columns else ""
                
                # Ambil kuantiti
                kuantiti_terpakai_str = row[f"Kuantiti > Terpakai{suffix}"] if f"Kuantiti > Terpakai{suffix}" in df_cleaned.columns else ""
                kuantiti_rusak_str = row[f"Kuantiti > Rusak{suffix}"] if f"Kuantiti > Rusak{suffix}" in df_cleaned.columns else ""
                label_qc = str(row[f"Label QC{suffix}"]).strip() if f"Label QC{suffix}" in df_cleaned.columns else ""
                
                # Parse kuantiti
                kuantiti_terpakai = parse_kuantiti(kuantiti_terpakai_str)
                kuantiti_rusak = parse_kuantiti(kuantiti_rusak_str)
                
                # Simpan hasil
                hasil_perhitungan.append({
                    "Nomor Batch": nomor_batch,
                    "Nama Bahan": nama_bahan,
                    "Kode Bahan": kode_bahan,
                    "Kuantiti Terpakai": kuantiti_terpakai,
                    "Kuantiti Rusak": kuantiti_rusak,
                    "Label QC": label_qc,
                    "Kuantiti Terpakai Str": kuantiti_terpakai_str,
                    "Kuantiti Rusak Str": kuantiti_rusak_str
                })
    
    return pd.DataFrame(hasil_perhitungan)


def buat_summary_kuantiti(df_kuantiti):
    """
    Membuat summary kuantiti per bahan per batch dan per label QC
    """
    # Summary 1: Total per Bahan per Batch
    summary_batch = df_kuantiti.groupby(["Nomor Batch", "Nama Bahan", "Kode Bahan"]).agg({
        "Kuantiti Terpakai": "sum",
        "Kuantiti Rusak": "sum",
        "Label QC": lambda x: ", ".join(sorted(set([str(v) for v in x if str(v).strip() not in ["", "nan", "None"]])))
    }).reset_index()
    
    # Format kuantiti kembali ke string dengan unit
    summary_batch["Total Kuantiti Terpakai"] = summary_batch["Kuantiti Terpakai"].apply(
        lambda x: f"{x:,.0f} GRAM" if x > 0 else "0 GRAM"
    )
    summary_batch["Total Kuantiti Rusak"] = summary_batch["Kuantiti Rusak"].apply(
        lambda x: f"{x:,.0f} GRAM" if x > 0 else "0 GRAM"
    )
    
    # Summary 2: Total per Label QC
    # Pecah label QC yang digabung dengan koma
    df_label_expanded = []
    for _, row in df_kuantiti.iterrows():
        labels = str(row["Label QC"]).split(",") if str(row["Label QC"]).strip() not in ["", "nan", "None"] else [""]
        for label in labels:
            label = label.strip()
            if label:
                df_label_expanded.append({
                    "Nomor Batch": row["Nomor Batch"],
                    "Nama Bahan": row["Nama Bahan"],
                    "Kode Bahan": row["Kode Bahan"],
                    "Label QC": label,
                    "Kuantiti Terpakai": row["Kuantiti Terpakai"],
                    "Kuantiti Rusak": row["Kuantiti Rusak"]
                })
    
    df_label_expanded = pd.DataFrame(df_label_expanded)
    
    if not df_label_expanded.empty:
        summary_label = df_label_expanded.groupby(["Label QC", "Nama Bahan", "Kode Bahan"]).agg({
            "Kuantiti Terpakai": "sum",
            "Kuantiti Rusak": "sum",
            "Nomor Batch": lambda x: len(set(x))
        }).reset_index()
        
        summary_label["Total Kuantiti Terpakai"] = summary_label["Kuantiti Terpakai"].apply(
            lambda x: f"{x:,.0f} GRAM" if x > 0 else "0 GRAM"
        )
        summary_label["Total Kuantiti Rusak"] = summary_label["Kuantiti Rusak"].apply(
            lambda x: f"{x:,.0f} GRAM" if x > 0 else "0 GRAM"
        )
        summary_label["Jumlah Batch"] = summary_label["Nomor Batch"].apply(lambda x: f"{x} batch")
        
        # Reorder kolom
        summary_label = summary_label[["Label QC", "Nama Bahan", "Kode Bahan", "Total Kuantiti Terpakai", "Total Kuantiti Rusak", "Jumlah Batch"]]
    else:
        summary_label = pd.DataFrame()
    
    # Reorder kolom untuk summary batch
    summary_batch = summary_batch[["Nomor Batch", "Nama Bahan", "Kode Bahan", "Total Kuantiti Terpakai", "Total Kuantiti Rusak", "Label QC"]]
    
    return summary_batch, summary_label


def kuantiti():
    st.subheader("Upload Data Kuantiti Bahan")
    uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx", "xls"], key="kuantiti_uploader")

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)

            # Hapus kolom yang tidak diperlukan
            drop_cols = ["No. Order Produksi", "Jalur"]
            drop_cols += [col for col in df.columns if "No Lot Supplier" in col]
            df_cleaned = df.drop(columns=[col for col in drop_cols if col in df.columns])
            
            # Rapikan data dengan pembersihan baris kosong yang lebih efektif
            df_cleaned = rapikan(df_cleaned)
            
            # === PEMBERSIHAN TAMBAHAN SETELAH RAPIKAN ===
            # Hapus baris yang mungkin masih kosong setelah proses rapikan
            
            # Method 1: Hapus baris yang semua kolomnya NaN
            df_cleaned = df_cleaned.dropna(how='all')
            
            # Method 2: Hapus baris yang hanya berisi string kosong
            def not_empty_row(row):
                return any(
                    pd.notna(val) and str(val).strip() not in ["", "nan", "None"] 
                    for val in row
                )
            
            df_cleaned = df_cleaned[df_cleaned.apply(not_empty_row, axis=1)].reset_index(drop=True)
            
            # Method 3: Pembersihan final berdasarkan kolom kunci
            # Pastikan setiap baris memiliki setidaknya satu data yang bermakna
            key_columns = ["Nomor Batch"] + [col for col in df_cleaned.columns if "Nama Bahan Formula" in col]
            
            def has_key_data(row):
                return any(
                    pd.notna(row[col]) and str(row[col]).strip() not in ["", "nan", "None"] 
                    for col in key_columns if col in row.index
                )
            
            df_cleaned = df_cleaned[df_cleaned.apply(has_key_data, axis=1)].reset_index(drop=True)
            
            # === PERBAIKAN ERROR <NA>: Konversi semua nilai NA untuk kompatibilitas Excel ===
            # Ganti semua pd.NA dengan np.nan atau string kosong
            df_cleaned = df_cleaned.fillna("")  # Atau gunakan np.nan jika ingin tetap NaN
            
            # Pastikan tidak ada dtype object yang bermasalah
            for col in df_cleaned.columns:
                if df_cleaned[col].dtype == 'object':
                    df_cleaned[col] = df_cleaned[col].astype(str).replace('nan', '').replace('<NA>', '')
            
            st.success("✅ File berhasil dimuat dan dirapikan.")
            st.subheader("📄 Data Setelah Dirapikan")
            st.dataframe(df_cleaned)
            
            # Hitung kuantiti
            df_kuantiti = hitung_total_kuantiti(df_cleaned)
            
            if not df_kuantiti.empty:
                st.subheader("📊 Detail Kuantiti per Bahan")
                st.dataframe(df_kuantiti)
                
                # Buat summary
                summary_batch, summary_label = buat_summary_kuantiti(df_kuantiti)
                
                # Tampilkan Summary per Batch
                st.subheader("📋 Summary Kuantiti per Batch")
                st.dataframe(summary_batch)
                
                # Tampilkan Summary per Label QC (jika ada)
                if not summary_label.empty:
                    st.subheader("🏷️ Summary Kuantiti per Label QC")
                    st.dataframe(summary_label)
                
                # Fungsi untuk ekspor Excel
                def to_excel_kuantiti(df_list, sheet_names):
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        for df, sheet_name in zip(df_list, sheet_names):
                            df.to_excel(writer, index=False, sheet_name=sheet_name)
                    output.seek(0)
                    return output
                
                # Tombol download
                if not summary_label.empty:
                    excel_data = to_excel_kuantiti(
                        [df_cleaned, df_kuantiti, summary_batch, summary_label],
                        ["Data Bersih", "Detail Kuantiti", "Summary per Batch", "Summary per Label QC"]
                    )
                else:
                    excel_data = to_excel_kuantiti(
                        [df_cleaned, df_kuantiti, summary_batch],
                        ["Data Bersih", "Detail Kuantiti", "Summary per Batch"]
                    )
                
                st.download_button(
                    label="📥 Download Semua Data Kuantiti (Excel)",
                    data=excel_data,
                    file_name="analisis_kuantiti_bahan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Filter berdasarkan Label QC
                st.header("🔍 Filter Berdasarkan Label QC")
                
                # Dapatkan semua label QC unik dari data kuantiti
                all_labels_kuantiti = []
                for labels_str in df_kuantiti["Label QC"].dropna():
                    if str(labels_str).strip() not in ["", "nan", "None"]:
                        labels = str(labels_str).split(",")
                        all_labels_kuantiti.extend([label.strip() for label in labels if label.strip()])
                
                all_labels_kuantiti = sorted(list(set(all_labels_kuantiti)))
                
                if all_labels_kuantiti:
                    # Checkbox untuk pilih semua
                    select_all_kuantiti = st.checkbox("Pilih Semua Label QC", key="select_all_kuantiti")
                    
                    if select_all_kuantiti:
                        default_selection_kuantiti = all_labels_kuantiti
                    else:
                        default_selection_kuantiti = []
                    
                    selected_labels_kuantiti = st.multiselect(
                        "Pilih Label QC untuk Analisis Kuantiti:", 
                        all_labels_kuantiti,
                        default=default_selection_kuantiti,
                        key="kuantiti_label_selector"
                    )
                    
                    if selected_labels_kuantiti:
                        # Filter data berdasarkan label yang dipilih
                        filtered_kuantiti = []
                        for _, row in df_kuantiti.iterrows():
                            row_labels = str(row["Label QC"]).split(",") if str(row["Label QC"]).strip() not in ["", "nan", "None"] else []
                            row_labels = [label.strip() for label in row_labels]
                            
                            # Cek apakah ada label yang dipilih di baris ini
                            if any(label in selected_labels_kuantiti for label in row_labels):
                                filtered_kuantiti.append(row)
                        
                        if filtered_kuantiti:
                            df_filtered_kuantiti = pd.DataFrame(filtered_kuantiti)
                            
                            # Tampilkan data yang difilter
                            label_list_str = ", ".join(selected_labels_kuantiti)
                            st.subheader(f"📊 Data Kuantiti dengan Label QC: {label_list_str}")
                            st.dataframe(df_filtered_kuantiti)
                            
                            # Buat summary untuk data yang difilter
                            summary_batch_filtered, summary_label_filtered = buat_summary_kuantiti(df_filtered_kuantiti)
                            
                            # Tampilkan summary yang difilter
                            st.subheader("📋 Summary Kuantiti per Batch (Filtered)")
                            st.dataframe(summary_batch_filtered)
                            
                            if not summary_label_filtered.empty:
                                st.subheader("🏷️ Summary Kuantiti per Label QC (Filtered)")
                                st.dataframe(summary_label_filtered)
                            
                            # Download untuk data yang difilter
                            if not summary_label_filtered.empty:
                                excel_filtered = to_excel_kuantiti(
                                    [df_filtered_kuantiti, summary_batch_filtered, summary_label_filtered],
                                    ["Detail Kuantiti", "Summary per Batch", "Summary per Label QC"]
                                )
                            else:
                                excel_filtered = to_excel_kuantiti(
                                    [df_filtered_kuantiti, summary_batch_filtered],
                                    ["Detail Kuantiti", "Summary per Batch"]
                                )
                            
                            # Buat nama file yang sesuai
                            if len(selected_labels_kuantiti) == 1:
                                filename = f"kuantiti_label_qc_{selected_labels_kuantiti[0]}.xlsx"
                            else:
                                filename = f"kuantiti_multiple_label_qc.xlsx"
                            
                            st.download_button(
                                label="📥 Download Data Kuantiti Filtered (Excel)",
                                data=excel_filtered,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="download_filtered_kuantiti"
                            )
                        else:
                            st.warning("Tidak ada data kuantiti dengan Label QC yang dipilih.")
                else:
                    st.info("Tidak ada Label QC yang ditemukan dalam data kuantiti.")
            else:
                st.warning("Tidak ada data kuantiti yang dapat diproses.")

        except Exception as e:
            st.error(f"❌ Terjadi kesalahan saat memproses file: {e}")
            st.error("Detail error untuk debugging:")
            st.exception(e)


# Fungsi utama untuk menjalankan aplikasi
def main():
    st.set_page_config(page_title="Analisis Data Bahan", layout="wide")
    st.title("🧪 Aplikasi Analisis Data Bahan")
    
    # Sidebar untuk memilih fitur
    st.sidebar.title("🔧 Pilih Fitur")
    fitur = st.sidebar.radio(
        "Pilih analisis yang ingin dilakukan:",
        ["Filter Label QC", "Analisis Kuantiti"]
    )
    
    if fitur == "Filter Label QC":
        filter_labelqc()
    elif fitur == "Analisis Kuantiti":
        kuantiti()


if __name__ == "__main__":
    main()
