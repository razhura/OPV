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
                valid_rows = df_asli[[kode_col, label_col]].dropna(subset=[kode_col])
                for _, row in valid_rows.iterrows():
                    all_data.append({
                        "Kode Bahan": str(row[kode_col]).strip(),
                        "Label QC": row[label_col] if pd.notna(row[label_col]) else ""
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
                
                excel_all_grouped = to_excel(grouped_all_df)
                st.download_button(
                    label="📥 Download Ringkasan Semua Label QC (Excel)",
                    data=excel_all_grouped,
                    file_name="ringkasan_semua_label_qc.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            # Fitur filter untuk detail kode bahan tertentu (tetap seperti sebelumnya)
            selected_kode = st.selectbox("🔍 Pilih Kode Bahan untuk Detail", kode_bahan_list)

            # Filter berdasarkan pasangan kode dan label yang sesuai
            hasil_data = []

            for kode_col, label_col in kode_bahan_pairs:
                mask = df_asli[kode_col].astype(str).str.strip() == selected_kode
                filtered_rows = df_asli[mask]
                for _, row in filtered_rows.iterrows():
                    hasil_data.append({
                        "Kode Bahan": selected_kode,
                        "Label QC": row[label_col] if pd.notna(row[label_col]) else ""
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

        except Exception as e:
            st.error(f"❌ Terjadi kesalahan saat membaca file: {e}")

if __name__ == "__main__":
    filter_labelqc()
