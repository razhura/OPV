import streamlit as st
import pandas as pd
import io
from datetime import datetime
import openpyxl

def process_multiple_excel_files():
    st.title("CQA EKSTRAK")
    st.write("Upload File yang banyak 🤑")

    # Inisialisasi session state jika belum ada
    if 'files_for_processing' not in st.session_state:
        st.session_state.files_for_processing = []
    if 'last_uploaded_file_names_sorted' not in st.session_state:
        st.session_state.last_uploaded_file_names_sorted = []

    # File uploader
    newly_uploaded_files_from_widget = st.file_uploader(
        "Pilih Beberapa File Excel Yang Akan Kamu Ekstrak",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="Anda dapat memilih beberapa file Excel sekaligus. Urutan awal mungkin berdasarkan abjad."
    )

    # Logika untuk memperbarui st.session_state.files_for_processing
    # berdasarkan perubahan pada file uploader.
    if newly_uploaded_files_from_widget is not None:
        current_uploader_names_sorted = sorted([f.name for f in newly_uploaded_files_from_widget])
        if current_uploader_names_sorted != st.session_state.last_uploaded_file_names_sorted:
            st.session_state.files_for_processing = newly_uploaded_files_from_widget
            st.session_state.last_uploaded_file_names_sorted = current_uploader_names_sorted
            if newly_uploaded_files_from_widget:
                st.info("Daftar file telah diperbarui dari uploader. Urutan awal mungkin berdasarkan abjad. Anda dapat mengaturnya di bawah ini.")
            # st.rerun() tidak diperlukan di sini, Streamlit akan menangani alur secara otomatis.

    if st.session_state.files_for_processing:
        st.write(f"📁 {len(st.session_state.files_for_processing)} file siap diproses. Anda dapat mengubah urutannya:")
        st.markdown("---")

        # Tampilkan file dengan tombol untuk mengatur ulang urutan
        for i in range(len(st.session_state.files_for_processing)):
            file_obj = st.session_state.files_for_processing[i]
            
            # Gunakan kolom untuk tata letak setiap item file
            col1, col2, col3 = st.columns([0.7, 0.15, 0.15]) # Sesuaikan rasio sesuai kebutuhan

            with col1:
                col1.write(f"{i + 1}. {file_obj.name}")

            with col2:
                # Tombol "Naik" (⬆️)
                # Dinonaktifkan jika file sudah paling atas
                if col2.button("⬆️", key=f"up_{i}", help="Pindahkan ke Atas", disabled=(i == 0)):
                    # Tukar dengan item sebelumnya
                    item_to_move = st.session_state.files_for_processing.pop(i)
                    st.session_state.files_for_processing.insert(i - 1, item_to_move)
                    st.rerun() # Jalankan ulang aplikasi untuk mencerminkan urutan baru

            with col3:
                # Tombol "Turun" (⬇️)
                # Dinonaktifkan jika file sudah paling bawah
                if col3.button("⬇️", key=f"down_{i}", help="Pindahkan ke Bawah", disabled=(i == len(st.session_state.files_for_processing) - 1)):
                    # Tukar dengan item berikutnya
                    item_to_move = st.session_state.files_for_processing.pop(i)
                    st.session_state.files_for_processing.insert(i + 1, item_to_move)
                    st.rerun() # Jalankan ulang aplikasi untuk mencerminkan urutan baru
        
        st.markdown("---")
        st.info("📋 Akan mengekstrak kolom A, F, G mulai dari baris ke-3 (header adalah sel yang digabung pada baris 1-2)")
        
        # Tombol Proses
        if st.button("🔄 Proses File", type="primary"):
            if st.session_state.files_for_processing: # Pastikan ada file untuk diproses
                process_files(st.session_state.files_for_processing) # Kirim daftar file yang sudah diurutkan
            else:
                st.warning("Tidak ada file untuk diproses. Silakan unggah file terlebih dahulu.")
                    
    else:
        # Kondisi ini terjadi jika st.session_state.files_for_processing kosong.
        # Periksa apakah uploader juga kosong atau belum pernah ada interaksi.
        if newly_uploaded_files_from_widget is None or not newly_uploaded_files_from_widget:
            st.info("👆 Silakan unggah file Excel untuk memulai")

def read_excel_with_merged_headers(file, target_columns=[0, 5, 6]):
    try:
        workbook = openpyxl.load_workbook(file, data_only=True)
        sheet = workbook.active
        
        headers = []
        for col_idx in target_columns:
            header_value = sheet.cell(row=1, column=col_idx+1).value
            if header_value is None or str(header_value).strip() == '':
                header_value = sheet.cell(row=2, column=col_idx+1).value
            
            if header_value is None:
                header_value = f"Kolom_{chr(65+col_idx)}"
            
            headers.append(str(header_value).strip())
        
        workbook.close()
        
        df = pd.read_excel(file, header=None, skiprows=2)
        df_selected = df.iloc[:, target_columns].copy()
        df_selected.columns = headers
        df_selected = df_selected.dropna(how='all').reset_index(drop=True)
        
        return df_selected, headers
        
    except Exception as e:
        st.error(f"Error saat membaca file: {str(e)}")
        return None, None

def process_files(uploaded_files_ordered): # Nama parameter diubah untuk kejelasan
    all_data = []
    error_files = []
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i, uploaded_file in enumerate(uploaded_files_ordered): # Gunakan daftar yang sudah diurutkan
        try:
            progress = (i + 1) / len(uploaded_files_ordered)
            progress_bar.progress(progress)
            status_text.text(f"Memproses: {uploaded_file.name}")
            
            df_data, headers = read_excel_with_merged_headers(uploaded_file)
            
            if df_data is None or df_data.empty:
                error_files.append({
                    'file': uploaded_file.name,
                    'error': 'Tidak ada data ditemukan atau file kosong'
                })
                continue
                        
            all_data.append({
                'filename': uploaded_file.name,
                'data': df_data,
                'headers': headers
            })
            
        except Exception as e:
            error_files.append({
                'file': uploaded_file.name,
                'error': str(e)
            })
            st.error(f"Error memproses {uploaded_file.name}: {str(e)}")
    
    progress_bar.empty()
    status_text.empty()
    
    if all_data:
        st.success(f"✅ Berhasil memproses {len(all_data)} file")
        
        with st.expander("Pratinjau data dari setiap file (setelah diurutkan)"):
            for file_data in all_data:
                st.write(f"**{file_data['filename']}**")
                st.write(f"Header: {file_data['headers']}")
                st.write(f"Bentuk: {file_data['data'].shape}")
                st.dataframe(file_data['data'].head())
                st.write("---")
        
        combined_df = pd.concat([file_data['data'] for file_data in all_data], ignore_index=True)
        
        st.write(f"📊 Bentuk data gabungan: {combined_df.shape[0]} baris × {combined_df.shape[1]} kolom")
        
        with st.expander("Pratinjau data gabungan"):
            st.dataframe(combined_df.head(20))
        
        st.subheader("🔄 Data Hasil Transpose")
        
        transposed_df = combined_df.T
        transposed_df.reset_index(inplace=True)
        transposed_df.columns = [f'Baris_{j}' for j in range(len(transposed_df.columns))] # Mengubah i menjadi j untuk menghindari konflik
        transposed_df.rename(columns={'Baris_0': 'Kolom_Asli'}, inplace=True)
        
        st.subheader("📋 Hasil Akhir")
        st.write(f"Bentuk data transpose: {transposed_df.shape[0]} baris × {transposed_df.shape[1]} kolom")
        
        with st.expander("Pratinjau data yang sudah di-transpose"):
            st.dataframe(transposed_df.head(20))
        
        st.subheader("📥 Unduh Hasil")
        
        output_buffer = io.BytesIO()
        
        sheet_option = st.radio(
            "Pengaturan sheet Excel:",
            ["Satu sheet (hanya hasil akhir)", "Beberapa sheet (asli + transpose)"]
        )
        
        if sheet_option == "Satu sheet (hanya hasil akhir)":
            transposed_df.to_excel(output_buffer, index=False, sheet_name='Data_Final')
        else:
            with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                combined_df.to_excel(writer, sheet_name='Gabungan_Asli', index=False)
                transposed_df.to_excel(writer, sheet_name='Hasil_Final', index=False)
        
        output_buffer.seek(0)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"processed_AFG_columns_{timestamp}.xlsx"
        
        st.download_button(
            label="📥 Unduh File Excel Hasil Proses",
            data=output_buffer.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
    else:
        st.error("❌ Tidak ada file yang berhasil diproses")
    
    if error_files:
        st.subheader("⚠️ Error Pemrosesan")
        error_df = pd.DataFrame(error_files)
        st.dataframe(error_df)

if __name__ == "__main__":
    process_multiple_excel_files()
