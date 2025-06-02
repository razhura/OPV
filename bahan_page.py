import pandas as pd
import numpy as np
import io
import streamlit as st
from openpyxl import load_workbook
import re


def extract_headers_from_rows_10_and_11(excel_file):
    wb = load_workbook(excel_file, data_only=True)
    ws = wb.active

    headers = []
    seen = {}
    max_col = ws.max_column

    # Adjusting to read from row 10 (index 9) and row 11 (index 10)
    # openpyxl is 1-indexed for rows and columns
    for col in range(1, max_col + 1):
        cell_10_val_obj = ws.cell(row=10, column=col) # Row 10
        cell_11_val_obj = ws.cell(row=11, column=col) # Row 11

        # Handle merged cells for row 10
        for merged_range in ws.merged_cells.ranges:
            if cell_10_val_obj.coordinate in merged_range:
                # Get the top-left cell of the merged range
                cell_10_val_obj = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                break
        
        # Handle merged cells for row 11
        for merged_range in ws.merged_cells.ranges:
            if cell_11_val_obj.coordinate in merged_range:
                # Get the top-left cell of the merged range
                cell_11_val_obj = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                break

        val_10 = str(cell_10_val_obj.value).strip() if cell_10_val_obj.value else ""
        val_11 = str(cell_11_val_obj.value).strip() if cell_11_val_obj.value else ""

        # Logic for combining headers (first 2 columns use header from row 10 directly)
        if col <= 2: # For the first two columns (e.g., 'Nomor Batch', 'No. Order Produksi')
            header = val_10
        elif not val_11 or val_10 == val_11 : # If row 11 is empty or same as row 10
            header = val_10
        else: # Combine headers
            header = f"{val_10} > {val_11}"

        if header in seen:
            seen[header] += 1
            header = f"{header}_{seen[header]}" # Append suffix for duplicates
        else:
            seen[header] = 1

        headers.append(header)

    return headers


def normalize_columns(df):
    mapping = {
        'Nomor Batch': 'Nomor Batch',
        'No. Order Produksi': 'No. Order Produksi',
        'Jalur': 'Jalur',
        'Kode Bahan': 'Kode Bahan',
        'Nama Bahan': 'Nama Bahan',
        'Kuantiti > Terpakai': 'Kuantiti > Terpakai',
        'Kuantiti > Rusak': 'Kuantiti > Rusak',
        'No Lot Supplier': 'No Lot Supplier',
        'Label QC': 'Label QC'
    }

    from difflib import get_close_matches

    new_columns = {}
    original_df_cols = df.columns.tolist() # Work with a list of current column names

    for expected_col_pattern in mapping.keys():
        # Find matches for the base pattern (e.g., "Kode Bahan" for "Kode Bahan_2", "Kode Bahan_3")
        # This handles cases where headers were made unique like "Kode Bahan_2"
        # We want to normalize the *base* part of the header.
        
        # If the exact expected_col_pattern is already a column, map it.
        if expected_col_pattern in original_df_cols:
            # Check if we are trying to map it to itself (which is fine)
            # or if it's a specific mapping like 'Old Name': 'New Name'
            if matches_actual_col_in_df(expected_col_pattern, df.columns):
                 new_columns[expected_col_pattern] = mapping[expected_col_pattern]
                 continue # Move to next mapping key

        # Try to find close matches for the expected_col_pattern
        # among the original DataFrame columns.
        matches = get_close_matches(expected_col_pattern, original_df_cols, n=1, cutoff=0.6)
        if matches:
            # If a close match is found, map this matched column name
            # to the target name from the mapping.
            matched_col_in_df = matches[0]
            target_col_name = mapping[expected_col_pattern]
            new_columns[matched_col_in_df] = target_col_name
    
    # Rename only the columns that were successfully matched.
    df = df.rename(columns=new_columns)
    return df

def matches_actual_col_in_df(pattern, df_columns):
    """Helper to check if a pattern exactly matches any column, or base of a numbered column."""
    if pattern in df_columns:
        return True
    # Check for patterns like "Kode Bahan" when columns might be "Kode Bahan_2"
    # This specific logic might be too broad if not careful.
    # For now, we rely on get_close_matches for more complex scenarios.
    return False


def transform_batch_data(df):
    selected_cols = [
        'Nomor Batch',
        'No. Order Produksi',
        'Jalur',
        'Kode Bahan',
        'Nama Bahan',
        'Kuantiti > Terpakai',
        'Kuantiti > Rusak',
        'No Lot Supplier',
        'Label QC'
    ]

    # Identify which of the selected_cols are actually present after normalization
    # This is important because normalize_columns might not find all of them,
    # or they might have suffixes if they were duplicates (e.g., 'Nama Bahan_2')
    
    present_selected_cols = []
    # Base columns that are expected to be unique and directly present
    base_unique_cols = ['Nomor Batch', 'No. Order Produksi', 'Jalur']
    # Repeated groups of columns
    repeated_col_patterns = ['Kode Bahan', 'Nama Bahan', 'Kuantiti > Terpakai', 'Kuantiti > Rusak', 'No Lot Supplier', 'Label QC']

    for col in base_unique_cols:
        if col in df.columns:
            present_selected_cols.append(col)
        else:
            # Try to find a suffixed version if direct match fails (e.g. 'Nomor Batch_2')
            # This shouldn't happen for these base columns if header extraction is correct
            suffixed_cols = [c for c in df.columns if c.startswith(col)]
            if suffixed_cols:
                present_selected_cols.append(suffixed_cols[0]) # take the first match

    # For repeated columns, find all their occurrences (e.g., 'Kode Bahan', 'Kode Bahan_2', etc.)
    for pattern in repeated_col_patterns:
        # Find all columns in df that start with this pattern
        matching_cols = [col for col in df.columns if col.startswith(pattern)]
        # Sort them to maintain order if they have numerical suffixes (e.g. _2, _3 ...)
        # Basic sort works for _2, _10. For complex cases, natural sort might be needed.
        matching_cols.sort() 
        present_selected_cols.extend(matching_cols)
    
    # Ensure no duplicates from the above process, though it should be fine.
    present_selected_cols = sorted(list(set(present_selected_cols)), key=present_selected_cols.index)


    missing = [col for col in ['Nomor Batch', 'No. Order Produksi', 'Jalur'] if col not in df.columns] # Check essential base columns
    if missing:
        raise ValueError(f"Kolom dasar berikut tidak ditemukan setelah normalisasi: {missing}. Kolom terdeteksi: {df.columns.tolist()}")

    # Use only the columns that are actually present in the DataFrame for selection
    # This df_subset will contain all relevant 'Nomor Batch', 'No. Order Produksi', 'Jalur',
    # and all variants of 'Kode Bahan X', 'Nama Bahan X', etc.
    df_subset = df[present_selected_cols].copy()
    
    # Ensure 'Nomor Batch' exists for grouping
    if 'Nomor Batch' not in df_subset.columns:
        raise ValueError("'Nomor Batch' tidak ditemukan di kolom DataFrame yang dipilih.")

    # Get unique batch order based on first appearance in the original data
    # Handle cases where 'Nomor Batch' might be NaN or None, drop them.
    valid_batches = df_subset['Nomor Batch'].dropna().drop_duplicates()
    batch_order = valid_batches.tolist()
    
    # Filter the DataFrame to only include rows with valid batch numbers
    df_subset_filtered = df_subset[df_subset['Nomor Batch'].isin(batch_order)].copy()

    # Group by 'Nomor Batch', maintaining the original order
    grouped = df_subset_filtered.groupby('Nomor Batch', sort=False)

    transformed_rows = []
    max_material_groups = 0 # Tracks the maximum number of material groups for any batch

    # Identify all columns related to materials (Kode Bahan, Nama Bahan, etc.)
    # These will have numerical suffixes if multiple material sets exist per original row.
    material_related_cols = {} # Stores lists of 'Kode Bahan 1', 'Kode Bahan 2', ...
    
    # Find all unique base material column names (e.g. "Kode Bahan", "Nama Bahan")
    # and their numbered variants present in the dataframe
    
    # Example: "Kode Bahan", "Nama Bahan", "Kuantiti > Terpakai", etc.
    material_base_names = ['Kode Bahan', 'Nama Bahan', 'Kuantiti > Terpakai', 'Kuantiti > Rusak', 'No Lot Supplier', 'Label QC']
    
    # Find all actual column names in df_subset_filtered that match these patterns
    # e.g., "Kode Bahan", "Kode Bahan_2", "Nama Bahan", "Nama Bahan_2"
    # We need to determine the maximum index (like _2, _3)
    
    # Let's find the max index for material groups
    # This logic assumes suffixes like "_2", "_3". If no suffix, it's group 1.
    max_suffix_num = 0
    for col_name in df_subset_filtered.columns:
        for base_name in material_base_names:
            if col_name.startswith(base_name):
                suffix = col_name[len(base_name):].lstrip('_') # Remove base_name and leading '_'
                if suffix.isdigit():
                    max_suffix_num = max(max_suffix_num, int(suffix))
                elif not suffix: # No suffix means it's the first group
                    max_suffix_num = max(max_suffix_num, 1)
                    
    if max_suffix_num == 0 and any(any(base_name in col for col in df_subset_filtered.columns) for base_name in material_base_names):
        # This case means columns like "Kode Bahan" exist but no "Kode Bahan_2" etc.
        # So there is effectively 1 group of material columns.
        max_suffix_num = 1
        
    num_material_sets_per_row = max_suffix_num


    for batch_id in batch_order:
        if batch_id in grouped.groups:
            group_df = grouped.get_group(batch_id)
            
            # Take 'No. Order Produksi' and 'Jalur' from the first row of the group
            # Use .iloc[0] safely
            order_produksi = group_df['No. Order Produksi'].iloc[0] if 'No. Order Produksi' in group_df.columns and not group_df.empty else ""
            jalur = group_df['Jalur'].iloc[0] if 'Jalur' in group_df.columns and not group_df.empty else ""

            current_batch_row_data = [batch_id, order_produksi, jalur]
            num_items_in_this_batch = 0

            for _, item_row in group_df.iterrows():
                # Iterate through the number of material sets identified (e.g., _1, _2)
                for i in range(1, num_material_sets_per_row + 1):
                    suffix = f"_{i}" if i > 1 else "" # No suffix for the first set
                    
                    kode_bahan_col = f'Kode Bahan{suffix}'
                    nama_bahan_col = f'Nama Bahan{suffix}'
                    terpakai_col = f'Kuantiti > Terpakai{suffix}'
                    rusak_col = f'Kuantiti > Rusak{suffix}'
                    lot_col = f'No Lot Supplier{suffix}'
                    qc_col = f'Label QC{suffix}'

                    # Check if at least 'Kode Bahan' for this set exists and has a value
                    if kode_bahan_col in item_row and pd.notna(item_row[kode_bahan_col]) and str(item_row[kode_bahan_col]).strip() != "":
                        current_batch_row_data.extend([
                            item_row.get(kode_bahan_col, ""),
                            item_row.get(nama_bahan_col, ""),
                            item_row.get(terpakai_col, ""),
                            item_row.get(rusak_col, ""),
                            item_row.get(lot_col, ""),
                            item_row.get(qc_col, "")
                        ])
                        num_items_in_this_batch +=1
                    elif i == 1 and not (kode_bahan_col in item_row and pd.notna(item_row[kode_bahan_col]) and str(item_row[kode_bahan_col]).strip() != ""):
                        # If even the first set of material (no suffix) is empty for this item_row,
                        # still add placeholders if this item_row is the *first* for the batch
                        # to ensure the structure is somewhat maintained for the first item.
                        # This part is tricky and depends on exact desired output for sparse data.
                        # For now, we only add if Kode Bahan is present.
                        pass


            transformed_rows.append(current_batch_row_data)
            max_material_groups = max(max_material_groups, num_items_in_this_batch)


    # Dynamically create headers for the transformed DataFrame
    final_headers = ['Nomor Batch', 'No. Order Produksi', 'Jalur']
    for i in range(1, max_material_groups + 1):
        final_headers.extend([
            f"Kode Bahan {i}",
            f"Nama Bahan {i}",
            f"Kuantiti > Terpakai {i}",
            f"Kuantiti > Rusak {i}",
            f"No Lot Supplier {i}",
            f"Label QC {i}"
        ])
    
    # Pad rows to ensure they all have the same number of columns as final_headers
    # This must be done carefully. The length of final_headers is 3 (base) + max_material_groups * 6.
    expected_row_length = len(final_headers)
    padded_rows = []
    for row_data in transformed_rows:
        current_len = len(row_data)
        if current_len < expected_row_length:
            row_data.extend([''] * (expected_row_length - current_len))
        elif current_len > expected_row_length: # Should not happen if logic is correct
            row_data = row_data[:expected_row_length]
        padded_rows.append(row_data)

    if not padded_rows: # If no data was processed
        return pd.DataFrame(columns=final_headers)
        
    return pd.DataFrame(padded_rows, columns=final_headers)


def simplify_headers(df):
    new_cols = []
    for col in df.columns:
        if col in ["Nomor Batch", "No. Order Produksi", "Jalur"]: # Keep these as is
            new_cols.append(col)
        else:
            simplified = re.sub(r"\s\d+$", "", col) # Remove trailing space and number
            new_cols.append(simplified)
    df.columns = new_cols
    return df


def create_filtered_table_by_name(df, selected_name):
    nama_bahan_cols = [col for col in df.columns if col.startswith('Nama Bahan ')]
    
    filtered_dfs = []
    for col_name_with_index in nama_bahan_cols: # e.g., "Nama Bahan 1", "Nama Bahan 2"
        # Get the index from the column name, e.g., "Nama Bahan 1" -> 1
        match = re.search(r'\s(\d+)$', col_name_with_index)
        if not match: continue # Skip if no index found (e.g. a base "Nama Bahan" if it exists without index)
        
        index_str = match.group(1)
        
        # Rows where the current `Nama Bahan {index}` column matches selected_name
        mask = df[col_name_with_index] == selected_name
        if not mask.any():
            continue

        temp_df = df[mask].copy() # Get rows that match the selected_name in this specific "Nama Bahan X" column

        # Select relevant columns for this specific index
        columns_to_keep_for_this_index = [
            'Nomor Batch', 
            'No. Order Produksi', 
            'Jalur', 
            f'Nama Bahan {index_str}',
            f'Kode Bahan {index_str}',
            f'Kuantiti > Terpakai {index_str}',
            f'Kuantiti > Rusak {index_str}',
            f'No Lot Supplier {index_str}',
            f'Label QC {index_str}'
        ]
        
        # Ensure all columns_to_keep_for_this_index exist in temp_df before selecting
        available_cols_for_index = [col for col in columns_to_keep_for_this_index if col in temp_df.columns]
        
        # Create a new DataFrame with only these columns and the filtered rows
        specific_filtered_df = temp_df[available_cols_for_index].copy()
        
        # Rename columns to remove the index (e.g., "Nama Bahan 1" -> "Nama Bahan")
        renamed_cols = {}
        for col_in_specific_df in specific_filtered_df.columns:
            if col_in_specific_df not in ['Nomor Batch', 'No. Order Produksi', 'Jalur']:
                new_name = re.sub(r'\s\d+$', '', col_in_specific_df)
                renamed_cols[col_in_specific_df] = new_name
        
        specific_filtered_df.rename(columns=renamed_cols, inplace=True)
        
        if not specific_filtered_df.empty:
            filtered_dfs.append(specific_filtered_df)
    
    if filtered_dfs:
        # Concatenate all DataFrames. They should have the same simplified column names now.
        final_df = pd.concat(filtered_dfs, ignore_index=True)
        # Reorder columns to a standard order
        standard_order = ['Nomor Batch', 'No. Order Produksi', 'Jalur', 
                          'Nama Bahan', 'Kode Bahan', 'Kuantiti > Terpakai', 
                          'Kuantiti > Rusak', 'No Lot Supplier', 'Label QC']
        # Filter to only include columns that are actually in final_df
        existing_standard_order = [col for col in standard_order if col in final_df.columns]
        return final_df[existing_standard_order]
    else:
        return pd.DataFrame(columns=['Nomor Batch', 'No. Order Produksi', 'Jalur', 
                                     'Nama Bahan', 'Kode Bahan', 'Kuantiti > Terpakai', 
                                     'Kuantiti > Rusak', 'No Lot Supplier', 'Label QC'])


def get_unique_bahan_names(df):
    nama_bahan_cols = [col for col in df.columns if col.startswith('Nama Bahan ')]
    unique_names = set()
    for col in nama_bahan_cols:
        values = df[col].dropna()
        values = values[values != '']
        unique_names.update(values)
    return sorted(list(unique_names))

# --- NEW FUNCTION ---
def get_unique_batch_numbers(df):
    """Extracts unique 'Nomor Batch' from the DataFrame."""
    if 'Nomor Batch' in df.columns:
        return sorted(list(df['Nomor Batch'].dropna().unique()))
    return []
# --- END NEW FUNCTION ---

def merge_same_materials(df):
    """
    Memindahkan kelompok data dengan kode bahan yang sama ke baris baru
    Jika dalam satu baris ada kode bahan yang sama di kelompok berbeda,
    kelompok kedua akan dipindah ke baris baru (tanpa nomor batch, no order, jalur)
    """
    result_rows = []
    
    # Determine the maximum index for material groups (e.g., "Kode Bahan 1", "Kode Bahan 2" -> max_index = 2)
    max_index = 0
    for col in df.columns:
        if col.startswith('Kode Bahan '):
            try:
                idx = int(col.split()[-1])
                if idx > max_index:
                    max_index = idx
            except ValueError: # Handle cases like "Kode Bahan" (no index)
                if 1 > max_index: max_index = 1 
    
    if max_index == 0: # No material columns found
        return df.copy()

    indices = list(range(1, max_index + 1))

    for _, row in df.iterrows():
        materials_in_row = {} # Key: kode bahan, Value: list of groups with this kode
        
        # Collect all material groups from the current row
        all_material_groups_in_row = []
        for i in indices:
            kode_col = f'Kode Bahan {i}'
            
            if kode_col in row and pd.notna(row[kode_col]) and
