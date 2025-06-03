import streamlit as st
import pandas as pd
import io
from datetime import datetime
import openpyxl

def process_data_with_stacking(all_data):
    """
    Create proper transpose where:
    - Column A data becomes column headers (merged/unique)
    - Original headers (G, H) become the first column (row labels) 
    - Data from G, H columns stay separate per file (NOT combined)
    - Filename is NOT shown in the Header column
    - Order follows first appearance in data, not alphabetical
    """
    if not all_data:
        return pd.DataFrame()
    
    # Get all unique values from column A in order of first appearance
    all_a_values = []
    seen_values = set()
    
    for file_data in all_data:
        df = file_data['data']
        if not df.empty:
            col_a = df.columns[0]
            # Iterate through data to preserve order of first appearance
            for val in df[col_a].dropna():
                val_str = str(val).strip()
                if val_str != '' and val_str not in seen_values:
                    all_a_values.append(val_str)
                    seen_values.add(val_str)
    
    if not all_a_values:
        return pd.DataFrame()
    
    # Use order of first appearance instead of alphabetical sorting
    sorted_a_values = all_a_values
    
    # Create transpose structure
    transpose_data = {'Header': []}
    
    # Initialize columns for each unique A value
    for a_val in sorted_a_values:
        transpose_data[a_val] = []
    
    # Process each file separately (G, H data NOT combined)
    for file_idx, file_data in enumerate(all_data):
        df = file_data['data']
        filename = file_data['filename']
        
        if df.empty:
            continue
            
        col_names = df.columns.tolist()
        col_a = col_names[0]  # Column A
        col_g = col_names[1]  # Column G  
        col_h = col_names[2]  # Column H
        
        # Add rows for this file's G and H data (WITHOUT filename in header)
        g_header = col_g  # Remove filename from header
        h_header = col_h  # Remove filename from header
        
        transpose_data['Header'].extend([g_header, h_header])
        
        # For each unique A value, find corresponding G and H data from this file
        for a_val in sorted_a_values:
            # Filter data for this A value in this specific file
            filtered_data = df[df[col_a] == a_val]
            
            # Get G and H values for this A value in this file
            g_values = filtered_data[col_g].dropna().tolist()
            h_values = filtered_data[col_h].dropna().tolist()
            
            # Combine multiple values with comma if exist
            g_combined = ', '.join([str(val) for val in g_values if str(val).strip() != '']) if g_values else ''
            h_combined = ', '.join([str(val) for val in h_values if str(val).strip() != '']) if h_values else ''
            
            # Add to transpose data
            transpose_data[a_val].extend([g_combined, h_combined])
    
    # Create DataFrame
    result_df = pd.DataFrame(transpose_data)
    
    return result_df

def process_multiple_excel_files():
    """
    Upload multiple Excel files, extract columns A, F, G from row 3 onwards, and combine with transpose
    """
    st.title("CQA EKSTRAK")
    st.write("Upload File yang banyak 🤑")
    
    # File uploader for multiple files
    uploaded_files = st.file_uploader(
        "Pilih Beberapa File Excel Yang Akan Kamu Ekstrak", 
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="Anda dapat memilih beberapa file Excel sekaligus"
    )
    
    if uploaded_files:
        st.write(f"📁 {len(uploaded_files)} file terupload")
        
        # Show uploaded file names
        with st.expander("Lihat file yang diupload"):
            for i, file in enumerate(uploaded_files, 1):
                st.write(f"{i}. {file.name}")
        
        st.info("📋 Akan mengekstrak kolom A, G, H mulai dari baris 3 (header adalah merged cells di baris 1-2)")
        
        # Process button
        if st.button("🔄 Proses File", type="primary"):
            process_files(uploaded_files)
                    
    else:
        st.info("👆 Silakan upload file Excel untuk memulai")

def read_excel_with_merged_headers(file, target_columns=[0, 6, 7]):  # A=0, G=6, H=7
    """
    Read Excel file with merged headers in rows 1-2, data starts from row 3
    """
    try:
        # Read the file to get merged header information
        workbook = openpyxl.load_workbook(file, data_only=True)
        sheet = workbook.active
        
        # Get headers from merged cells (row 1 or 2)
        headers = []
        for col_idx in target_columns:
            # Try to get header from row 1 first, then row 2
            header_value = sheet.cell(row=1, column=col_idx+1).value
            if header_value is None or str(header_value).strip() == '':
                header_value = sheet.cell(row=2, column=col_idx+1).value
            
            if header_value is None:
                header_value = f"Column_{chr(65+col_idx)}"  # A, G, H
            
            headers.append(str(header_value).strip())
        
        workbook.close()
        
        # Read data starting from row 3 (index 2), no header
        df = pd.read_excel(file, header=None, skiprows=2)
        
        # Extract only the target columns (A, G, H = index 0, 6, 7)
        df_selected = df.iloc[:, target_columns].copy()
        
        # Set the column names
        df_selected.columns = headers
        
        # Remove empty rows
        df_selected = df_selected.dropna(how='all').reset_index(drop=True)
        
        return df_selected, headers
        
    except Exception as e:
        st.error(f"Error membaca file: {str(e)}")
        return None, None

def process_files(uploaded_files):
    """
    Process multiple Excel files and combine selected columns with transpose
    """
    all_data = []
    error_files = []
    
    # Progress bar
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i, uploaded_file in enumerate(uploaded_files):
        try:
            # Update progress
            progress = (i + 1) / len(uploaded_files)
            progress_bar.progress(progress)
            status_text.text(f"Memproses: {uploaded_file.name}")
            
            # Read Excel file with merged headers
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
    
    # Clear progress indicators
    progress_bar.empty()
    status_text.empty()
    
    if all_data:
        # Display processing results
        st.success(f"✅ Berhasil memproses {len(all_data)} file")
        
        # Show preview of each file's data
        with st.expander("Pratinjau data dari setiap file"):
            for file_data in all_data:
                st.write(f"**{file_data['filename']}**")
                st.write(f"Header: {file_data['headers']}")
                st.write(f"Bentuk: {file_data['data'].shape}")
                st.dataframe(file_data['data'].head())
                st.write("---")
        
        # Combine all data
        combined_df = pd.concat([file_data['data'] for file_data in all_data], ignore_index=True)
        
        st.write(f"📊 Bentuk data gabungan: {combined_df.shape[0]} baris × {combined_df.shape[1]} kolom")
        

        
        # Show processed data with stacking and transpose
        processed_df = process_data_with_stacking(all_data)  # Pass all_data instead of combined_df
        
        if not processed_df.empty:
            
            # Use processed data as final result (now properly transposed)
            st.subheader("📋 Hasil Akhir (Transposed)")
            st.write(f"Bentuk data final: {processed_df.shape[0]} baris × {processed_df.shape[1]} kolom")
            st.write("**Format:** Kolom A yang sama digabung jadi header, Data G & H dari setiap file terpisah")
            
            with st.expander("Pratinjau data hasil akhir"):
                # Apply styling to make the first column bold
                styled_df = processed_df.style.set_properties(
                    subset=['Header'], 
                    **{'font-weight': 'bold', 'width': 'fit-content'}
                )
                st.dataframe(styled_df, use_container_width=True)
            
            # Download section
            st.subheader("📥 Unduh Hasil")
            
            # Generate timestamp for filenames
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Create download buttons in columns
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Data Gabungan (Belum Transpose)**")
                # Create Excel file for combined data
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
                # Option for sheet organization
                sheet_option = st.radio(
                    "Organisasi sheet Excel:",
                    ["Sheet tunggal (hasil akhir saja)", "Multiple sheet (asli + transpose)"],
                    key="sheet_option"
                )
                
                # Create Excel file in memory
                output_buffer = io.BytesIO()
                
                if sheet_option == "Sheet tunggal (hasil akhir saja)":
                    processed_df.to_excel(output_buffer, index=False, sheet_name='Data_Transpose')
                else:
                    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                        # Original combined data
                        combined_df.to_excel(writer, sheet_name='Data_Asli_Gabungan', index=False)
                        # Final processed data
                        processed_df.to_excel(writer, sheet_name='Hasil_Transpose', index=False)
                
                output_buffer.seek(0)
                
                filename = f"transposed_AGH_columns_{timestamp}.xlsx"
                
                # Download button
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
    
    # Show error summary if any
    if error_files:
        st.subheader("⚠️ Error Pemrosesan")
        error_df = pd.DataFrame(error_files)
        st.dataframe(error_df)

# Run the app
if __name__ == "__main__":
    process_multiple_excel_files()
