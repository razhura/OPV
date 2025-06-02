import streamlit as st
import pandas as pd
import io
from datetime import datetime
import openpyxl

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
        help="You can select multiple Excel files at once"
    )
    
    if uploaded_files:
        st.write(f"📁 {len(uploaded_files)} files uploaded")
        
        # Show uploaded file names
        with st.expander("View uploaded files"):
            for i, file in enumerate(uploaded_files, 1):
                st.write(f"{i}. {file.name}")
        
        st.info("📋 Will extract columns A, F, G starting from row 3 (headers are merged cells in rows 1-2)")
        
        # Process button
        if st.button("🔄 Process Files", type="primary"):
            process_files(uploaded_files)
                    
    else:
        st.info("👆 Please upload Excel files to get started")

def read_excel_with_merged_headers(file, target_columns=[0, 5, 6]):  # A=0, F=5, G=6
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
                header_value = f"Column_{chr(65+col_idx)}"  # A, F, G
            
            headers.append(str(header_value).strip())
        
        workbook.close()
        
        # Read data starting from row 3 (index 2), no header
        df = pd.read_excel(file, header=None, skiprows=2)
        
        # Extract only the target columns (A, F, G = index 0, 5, 6)
        df_selected = df.iloc[:, target_columns].copy()
        
        # Set the column names
        df_selected.columns = headers
        
        # Remove empty rows
        df_selected = df_selected.dropna(how='all').reset_index(drop=True)
        
        return df_selected, headers
        
    except Exception as e:
        st.error(f"Error reading file: {str(e)}")
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
            status_text.text(f"Processing: {uploaded_file.name}")
            
            # Read Excel file with merged headers
            df_data, headers = read_excel_with_merged_headers(uploaded_file)
            
            if df_data is None or df_data.empty:
                error_files.append({
                    'file': uploaded_file.name,
                    'error': 'No data found or empty file'
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
            st.error(f"Error processing {uploaded_file.name}: {str(e)}")
    
    # Clear progress indicators
    progress_bar.empty()
    status_text.empty()
    
    if all_data:
        # Display processing results
        st.success(f"✅ Successfully processed {len(all_data)} files")
        
        # Show preview of each file's data
        with st.expander("Preview data from each file"):
            for file_data in all_data:
                st.write(f"**{file_data['filename']}**")
                st.write(f"Headers: {file_data['headers']}")
                st.write(f"Shape: {file_data['data'].shape}")
                st.dataframe(file_data['data'].head())
                st.write("---")
        
        # Combine all data
        combined_df = pd.concat([file_data['data'] for file_data in all_data], ignore_index=True)
        
        st.write(f"📊 Combined data shape: {combined_df.shape[0]} rows × {combined_df.shape[1]} columns")
        
        # Show combined data preview
        with st.expander("Preview combined data"):
            st.dataframe(combined_df.head(20))
        
        # Auto transpose the combined data
        st.subheader("🔄 Data Hasil Transpose")
        
        # Create transposed version
        transposed_df = combined_df.T
        transposed_df.reset_index(inplace=True)
        transposed_df.columns = [f'Baris_{i}' for i in range(len(transposed_df.columns))]
        transposed_df.rename(columns={'Baris_0': 'Kolom_Asli'}, inplace=True)
        
        # Show final result
        st.subheader("📋 Hasil Akhir")
        st.write(f"Bentuk data transpose: {transposed_df.shape[0]} baris × {transposed_df.shape[1]} kolom")
        
        with st.expander("Preview data yang sudah di-transpose"):
            st.dataframe(transposed_df.head(20))
        
        # Download section
        st.subheader("📥 Download Results")
        
        # Create Excel file in memory
        output_buffer = io.BytesIO()
        
        # Option for sheet organization
        sheet_option = st.radio(
            "Excel sheet organization:",
            ["Single sheet (final result only)", "Multiple sheets (original + transposed)"]
        )
        
        if sheet_option == "Single sheet (final result only)":
            transposed_df.to_excel(output_buffer, index=False, sheet_name='Final_Data')
        else:
            with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                # Original combined data
                combined_df.to_excel(writer, sheet_name='Original_Combined', index=False)
                # Final processed data
                transposed_df.to_excel(writer, sheet_name='Final_Result', index=False)
        
        output_buffer.seek(0)
        
        # Generate filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"processed_AFG_columns_{timestamp}.xlsx"
        
        # Download button
        st.download_button(
            label="📥 Download Processed Excel File",
            data=output_buffer.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
    else:
        st.error("❌ No files were successfully processed")
    
    # Show error summary if any
    if error_files:
        st.subheader("⚠️ Processing Errors")
        error_df = pd.DataFrame(error_files)
        st.dataframe(error_df)

# Run the app
if __name__ == "__main__":
    process_multiple_excel_files()
