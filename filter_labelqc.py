import pandas as pd
import numpy as np
import io
import streamlit as st
from openpyxl import load_workbook
import re
import os


def filter_labelqc():
    st.title("📤 UPLOAD HASIL JADI DARI CPP BAHAN")
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

                # Fungsi download excel kuantiti
                def to_excel_merged_blocks(df):
                    import openpyxl
                    from openpyxl.styles import Alignment
                    output = io.BytesIO()
                
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        bahan_unik = df["Nama Bahan Formula"].dropna().unique()
                        col_start = 0
                
                        for bahan in bahan_unik:
                            df_bahan = df[df["Nama Bahan Formula"] == bahan].copy().reset_index(drop=True)
                
                            # Rename header hanya untuk merge
                            df_bahan = df_bahan.rename(columns={
                                "Kuantiti: Terpakai": "Kuantiti: Terpakai",
                                "Kuantiti: Rusak": "Kuantiti: Rusak"
                            })
                
                            # Tulis data mulai dari row 2
                            df_bahan.to_excel(writer, index=False, sheet_name="Rekap Per Bahan",
                                              startrow=1, startcol=col_start)
                
                            # Handle merge header
                            wb = writer.book
                            ws = writer.sheets["Rekap Per Bahan"]
                
                            col_map = {col: idx+col_start+1 for idx, col in enumerate(df_bahan.columns)}
                            terpakai_idx = col_map["Kuantiti: Terpakai"]
                            rusak_idx = col_map["Kuantiti: Rusak"]
                
                            # Merge "Kuantiti" di baris 1
                            ws.merge_cells(start_row=1, start_column=terpakai_idx, end_row=1, end_column=rusak_idx)
                            ws.cell(row=1, column=terpakai_idx).value = "Kuantiti"
                            ws.cell(row=1, column=terpakai_idx).alignment = Alignment(horizontal="center")
                
                            # Isi header row 2 (nama kolom asli)
                            for col_name, col_idx in col_map.items():
                                ws.cell(row=2, column=col_idx).value = col_name
                                ws.cell(row=2, column=col_idx).alignment = Alignment(horizontal="center")
                
                            # Geser posisi ke kanan untuk blok bahan selanjutnya
                            col_start += len(df_bahan.columns) + 2
                
                    output.seek(0)
                    return output

                
                
                excel_all_grouped = to_excel(grouped_all_df)
                st.download_button(
                    label="📥 Download Ringkasan Semua Label QC (Excel)",
                    data=excel_all_grouped,
                    file_name="ringkasan_semua_label_qc.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
######################################################################################## 
            # Tampilkan semua data lengkap: Kode Bahan - Nomor Batch - Label QC
            st.header("📋 Semua Kode Bahan dengan Batch dan Label QC.")
            
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

########################################################################################  
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

            # Kuantiti
            st.header("📦 Rekap Kuantiti per Nama Bahan per Batch")
            kuantiti_data = []
            for i in range(20):  # maksimal 20 pasangan
                nama_col = f"Nama Bahan Formula{'' if i == 0 else f'.{i}'}"
                terpakai_col = f"Kuantiti > Terpakai{'' if i == 0 else f'.{i}'}"
                rusak_col = f"Kuantiti > Rusak{'' if i == 0 else f'.{i}'}"
                label_col = f"Label QC{'' if i == 0 else f'.{i}'}"
            
                if nama_col in df_asli.columns and terpakai_col in df_asli.columns:
                    for _, row in df_asli.iterrows():
                        batch = row[batch_cols[0]] if batch_cols else ""
                        nama = row[nama_col]
                        terpakai = row[terpakai_col]
                        rusak = row[rusak_col] if rusak_col in df_asli.columns else 0
                        label = row[label_col] if label_col in df_asli.columns else ""
            
                        if pd.notna(nama) and pd.notna(terpakai):
                            kuantiti_data.append({
                                "Nomor Batch": batch,
                                "Nama Bahan Formula": nama,
                                "Kuantiti: Terpakai": terpakai,
                                "Kuantiti: Rusak": rusak,
                                "Label QC": label
                            })
            
            df_kuantiti = pd.DataFrame(kuantiti_data)
            
            # Ambil angka dari kolom kuantiti
            def extract_angka(x):
                try:
                    num_str = str(x).split()[0]
                    num_str = num_str.replace(".", "")      # hapus titik ribuan
                    num_str = num_str.replace(",", ".")     # ubah koma jadi titik desimal
                    return float(num_str)
                except:
                    return 0

            
            df_kuantiti["Angka Terpakai"] = df_kuantiti["Kuantiti: Terpakai"].apply(extract_angka)
            df_kuantiti["Angka Rusak"] = df_kuantiti["Kuantiti: Rusak"].apply(extract_angka)
            
            hasil = []
            grouped = df_kuantiti.groupby(["Nomor Batch", "Nama Bahan Formula"])
            
            for (batch, bahan), group in grouped:
                for idx, row in group.iterrows():
                    hasil.append({
                        "Nomor Batch": batch if idx == group.index[0] else "",
                        "Nama Bahan Formula": bahan if idx == group.index[0] else "",
                        "Kuantiti: Terpakai": row["Kuantiti: Terpakai"],
                        "Kuantiti: Rusak": row["Kuantiti: Rusak"],
                        "Label QC": row["Label QC"]
                    })
            
                if len(group) > 1:
                    total_terpakai = group["Angka Terpakai"].sum()
                    total_rusak = group["Angka Rusak"].sum()
            
                    total_terpakai_str = f"{int(total_terpakai):,}".replace(",", ".") + " GRAM"
                    total_rusak_str = f"{int(total_rusak):,}".replace(",", ".")
            
                    hasil.append({
                        "Nomor Batch": "",
                        "Nama Bahan Formula": "",
                        "Kuantiti: Terpakai": total_terpakai_str,
                        "Kuantiti: Rusak": total_rusak_str,
                        "Label QC": ""
                    })

            
            df_hasil = pd.DataFrame(hasil)
            #st.dataframe(df_hasil)
            # from functools import reduce
            
            # bahan_unik = df_hasil["Nama Bahan Formula"].dropna().unique()
            # dfs_horizontal = []
            
            # for idx, bahan in enumerate(bahan_unik):
            #     df_bahan = df_hasil[df_hasil["Nama Bahan Formula"] == bahan].copy()
            #     df_bahan.reset_index(drop=True, inplace=True)
            
            #     # Tambahkan suffix .1, .2 dst khusus untuk tampilan
            #     suffix = f".{idx}" if idx > 0 else ""
            #     df_bahan.columns = [f"{col}{suffix}" for col in df_bahan.columns]
            
            #     dfs_horizontal.append(df_bahan)
            
            # if dfs_horizontal:
            #     df_preview_horizontal = reduce(lambda left, right: pd.concat([left, right], axis=1), dfs_horizontal)
            #     st.subheader("👀 Preview Rekap Blok Tiap Bahan (Horizontal)")
            #     st.dataframe(df_preview_horizontal)
            # else:
            #     st.info("Tidak ada data yang dapat ditampilkan.")


        except Exception as e:
            st.error(f"❌ Terjadi kesalahan saat membaca file: {e}")

if __name__ == "__main__":
    filter_labelqc()
