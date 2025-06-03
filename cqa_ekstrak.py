import streamlit as st
import pandas as pd
import io
from datetime import datetime
import openpyxl

def process_data_with_stacking(all_data):

    if not all_data:
        return pd.DataFrame()
    
    all_a_values = []
    seen_values = set()
    
    for file_data in all_data:
        df = file_data['data']
        if not df.empty:
            col_a = df.columns[0]
            for val in df[col_a].dropna():
                val_str = str(val).strip()
                if val_str != '' and val_str not in seen_values:
                    all_a_values.append(val_str)
                    seen_values.add(val_str)
    
    if not all_a_values:
        return pd.DataFrame()
    
    sorted_a_values = all_a_values
    
    transpose_data = {'Header': []}
    
    for a_val in sorted_a_values:
        transpose_data[a_val] = []
    
    for file_idx, file_data in enumerate(all_data):
        df = file_data['data']
        filename = file_data['filename']
        
        if df.empty:
            continue
            
        col_names = df.columns.tolist()
        col_a = col_names[0]
        col_g = col_names[1]
        col_h = col_names[2]
        
        g_header = col_g
        h_header = col_h
        
        transpose_data['Header'].extend([g_header, h_header])
        
        for a_val in sorted_a_values:
            filtered_data = df[df[col_a] == a_val]
            
            g_values = filtered_data[col_g].dropna().tolist()
            h_values = filtered_data[col_h].dropna().tolist()
            
            g_combined = ', '.join([str(val) for val in g_values if str(val).strip() != '']) if g_values else ''
            h_combined = ', '.join([str(val) for val in h_values if str(val).strip() != '']) if h_values else ''
            
            transpose_data[a_val].extend([g_combined, h_combined])
    
    result_df = pd.DataFrame(transpose_data)
    
    return result_df

def read_excel_with_merged_headers(file, target_columns=[0, 6, 7]):

    try:
        workbook = openpyxl.load_workbook(file, data_only=True)
        sheet = workbook.active
        
        headers = []
        for col_idx in target_columns:
            header_value = sheet.cell(row=1, column=col_idx+1).value
            if header_value is None or str(header_value).strip() == '':
                header_value = sheet.cell(row=2, column=col_idx+1).value
            
            if header_value is None:
                header_value = f"Column_{chr(65+col_idx)}"
            
            headers.append(str(header_value).strip())
        
        workbook.close()
        
        df = pd.read_excel(file, header=None, skiprows=2)
        df_selected = df.iloc[:, target_columns].copy()
        df_selected.columns = headers
        df_selected = df_selected.dropna(how='all').reset_index(drop=True)
        
        return df_selected, headers
        
    except Exception as e:
        st.error(f"Error membaca file: {str(e)}")
        return None, None

def process_files(files_to_process_ordered): 

    all_data = []
    error_files = []
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    

    for i, uploaded_file in enumerate(files_to_process_ordered): 
        try:
            progress = (i + 1) / len(files_to_process_ordered)
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
        
        with st.expander("Pratinjau data dari setiap file (sesuai urutan yang diatur)"):
            for file_data in all_data:
                st.write(f"**{file_data['filename']}**")
                st.write(f"Header: {file_data['headers']}")
                st.write(f"Bentuk: {file_data['data'].shape}")
                st.dataframe(file_data['data'].head())
                st.write("---")
        
        combined_df = pd.concat([file_data['data'] for file_data in all_data], ignore_index=True)
        
        st.write(f"📊 Bentuk data gabungan: {combined_df.shape[0]} baris × {combined_df.shape[1]} kolom")
        
        processed_df = process_data_with_stacking(all_data)
        
        if not processed_df.empty:
            st.subheader("📋 Hasil Akhir (Transposed)")
            st.write(f"Bentuk data final: {processed_df.shape[0]} baris × {processed_df.shape[1]} kolom")
            st.write("**Format:** Kolom A yang sama digabung jadi header, Data G & H dari setiap file terpisah")
            
            with st.expander("Pratinjau data hasil akhir"):
                styled_df = processed_df.style.set_properties(
                    subset=['Header'], 
                    **{'font-weight': 'bold', 'width': 'fit-content'}
                )
                st.dataframe(styled_df, use_container_width=True)
            
            st.subheader("📥 Unduh Hasil")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Data Gabungan (Belum Transpose)**")
                combined_buffer = io.BytesIO()
                combined_df.to_excel(combined_buffer, index=False, sheet_name='Data_Gabungan')
                combined_buffer.seek(0)
                combined_filename = f"data_gabungan_{timestamp}.xlsx"
                st.download_button(
                    label="📥 Unduh Data Gabungan",
                    data=combined_buffer.getvalue(),
                    file_name=combined_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Download data yang sudah digabung dari semua file tapi belum di-transpose"
                )
            
            with col2:
                st.write("**Data Hasil Transpose**")
                sheet_option = st.radio(
                    "Organisasi sheet Excel:",
                    ["Sheet tunggal (hasil akhir saja)", "Multiple sheet (asli + transpose)"],
                    key="sheet_option_cqa" # Mengubah key agar unik
                )
                output_buffer = io.BytesIO()
                if sheet_option == "Sheet tunggal (hasil akhir saja)":
                    processed_df.to_excel(output_buffer, index=False, sheet_name='Data_Transpose')
                else:
                    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                        combined_df.to_excel(writer, sheet_name='Data_Asli_Gabungan', index=False)
                        processed_df.to_excel(writer, sheet_name='Hasil_Transpose', index=False)
                output_buffer.seek(0)
                filename = f"transposed_AGH_columns_{timestamp}.xlsx"
                st.download_button(
                    label="📥 Unduh Data Transpose",
                    data=output_buffer.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    help="Download data yang sudah di-transpose dengan benar"
                )
        else:
            st.error("❌ Tidak ada data yang dapat diproses untuk transpose")
    else:
        st.error("❌ Tidak ada file yang berhasil diproses")
    
    if error_files:
        st.subheader("⚠️ Error Pemrosesan")
        error_df = pd.DataFrame(error_files)
        st.dataframe(error_df)

def process_multiple_excel_files():
    st.title("CQA EKSTRAK")
    st.write("Upload File yang banyak 🤑")

    # Inisialisasi session state jika belum ada
    if 'files_for_cqa_processing' not in st.session_state:
        st.session_state.files_for_cqa_processing = []
    if 'last_uploaded_cqa_file_names_sorted' not in st.session_state:
        st.session_state.last_uploaded_cqa_file_names_sorted = []

    newly_uploaded_files_from_widget = st.file_uploader(
        "Pilih Beberapa File Excel Yang Akan Kamu Ekstrak", 
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="Anda dapat memilih beberapa file Excel sekaligus. Urutan awal mungkin berdasarkan abjad."
    )

    if newly_uploaded_files_from_widget is not None:
        current_uploader_names_sorted = sorted([f.name for f in newly_uploaded_files_from_widget])
        if current_uploader_names_sorted != st.session_state.last_uploaded_cqa_file_names_sorted:
            st.session_state.files_for_cqa_processing = newly_uploaded_files_from_widget
            st.session_state.last_uploaded_cqa_file_names_sorted = current_uploader_names_sorted
            if newly_uploaded_files_from_widget:
                st.info("Daftar file telah diperbarui dari uploader. Urutan awal mungkin berdasarkan abjad. Anda dapat mengaturnya di bawah ini.")

    if st.session_state.files_for_cqa_processing:
        st.write(f"📁 {len(st.session_state.files_for_cqa_processing)} file siap diproses. Anda dapat mengubah urutannya:")
        st.markdown("---")

        for i in range(len(st.session_state.files_for_cqa_processing)):
            file_obj = st.session_state.files_for_cqa_processing[i]
            col1, col2, col3 = st.columns([0.7, 0.15, 0.15]) 

            with col1:
                col1.write(f"{i + 1}. {file_obj.name}")
            with col2:
                if col2.button("⬆️", key=f"cqa_up_{i}", help="Pindahkan ke Atas", disabled=(i == 0)):
                    item_to_move = st.session_state.files_for_cqa_processing.pop(i)
                    st.session_state.files_for_cqa_processing.insert(i - 1, item_to_move)
                    st.rerun()
            with col3:
                if col3.button("⬇️", key=f"cqa_down_{i}", help="Pindahkan ke Bawah", disabled=(i == len(st.session_state.files_for_cqa_processing) - 1)):
                    item_to_move = st.session_state.files_for_cqa_processing.pop(i)
                    st.session_state.files_for_cqa_processing.insert(i + 1, item_to_move)
                    st.rerun()
        
        st.markdown("---")
        st.info("📋 Akan mengekstrak kolom A, G, H mulai dari baris 3 (header adalah merged cells di baris 1-2)")
        
        if st.button("🔄 Proses File", type="primary"):
            if st.session_state.files_for_cqa_processing:
                # Panggil process_files dengan daftar file yang sudah diurutkan dari session_state
                process_files(st.session_state.files_for_cqa_processing) 
            else:
                st.warning("Tidak ada file untuk diproses. Silakan unggah file terlebih dahulu.")
                    
    else:
        if newly_uploaded_files_from_widget is None or not newly_uploaded_files_from_widget:
            st.info("👆 Silakan upload file Excel untuk memulai")

if __name__ == "__main__":
    process_multiple_excel_files()
