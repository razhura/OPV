import pandas as pd
import re
from collections import defaultdict

def combine_duplicate_columns(df, mode="gabung"):
    """
    Menggabungkan atau memisahkan kolom duplikat berdasarkan mode yang dipilih.
    
    Parameters:
    df (DataFrame): DataFrame yang akan diproses
    mode (str): "gabung" untuk menggabungkan kolom [Teks] & [Nilai], 
                "pisah" untuk tetap memisahkannya (kondisi asli)
    
    Returns:
    DataFrame: DataFrame yang telah diproses sesuai mode
    """
    
    if mode == "pisah":
        # Mode pisah: Hanya tangani kolom yang benar-benar duplikat (nama persis sama)
        # Kolom dengan sufiks [Teks] dan [Nilai] tetap terpisah seperti kondisi asli
        
        new_columns = []
        seen = defaultdict(list)

        # Kelompokkan kolom berdasarkan nama yang persis sama
        for idx, col in enumerate(df.columns):
            seen[col].append(idx)

        final_data = {}

        # Proses setiap grup kolom
        for col, indexes in seen.items():
            if len(indexes) == 1:
                # Kolom unik, langsung ambil
                final_data[col] = df.iloc[:, indexes[0]]
            else:
                # Kolom duplikat (nama persis sama), gabungkan datanya
                combined_series = df.iloc[:, indexes[0]].copy()
                for i in indexes[1:]:
                    combined_series = combined_series.combine_first(df.iloc[:, i])
                final_data[col] = combined_series

        # Kembalikan dengan urutan kolom asli (tanpa duplikat)
        unique_columns = []
        for col in df.columns:
            if col not in unique_columns:
                unique_columns.append(col)
        
        return pd.DataFrame({col: final_data[col] for col in unique_columns})
    
    elif mode == "gabung":
        # Mode gabung: Lakukan penggabungan kolom duplikat + gabungkan kolom dengan sufiks
        
        # Tahap 1: Tangani kolom yang benar-benar duplikat (nama persis sama)
        seen = defaultdict(list)
        for idx, col in enumerate(df.columns):
            seen[col].append(idx)

        temp_data = {}
        for col, indexes in seen.items():
            if len(indexes) == 1:
                temp_data[col] = df.iloc[:, indexes[0]]
            else:
                combined_series = df.iloc[:, indexes[0]].copy()
                for i in indexes[1:]:
                    combined_series = combined_series.combine_first(df.iloc[:, i])
                temp_data[col] = combined_series

        temp_df = pd.DataFrame(temp_data)

        # Tahap 2: Gabungkan kolom yang nama dasarnya sama (hilangkan [teks], [nilai])
        def clean_column_name(name):
            return re.sub(r"\s*\[.*?\]\s*$", "", name).strip()

        base_name_map = defaultdict(list)
        original_order = {}  # Untuk menjaga urutan berdasarkan kemunculan pertama
        
        for idx, col in enumerate(temp_df.columns):
            base_name = clean_column_name(col)
            base_name_map[base_name].append(idx)
            
            # Simpan posisi kemunculan pertama dari base name ini
            if base_name not in original_order:
                original_order[base_name] = idx

        final_data = {}

        # Proses berdasarkan urutan kemunculan pertama
        for base_name in sorted(base_name_map.keys(), key=lambda x: original_order[x]):
            indexes = base_name_map[base_name]
            
            if len(indexes) == 1:
                final_data[base_name] = temp_df.iloc[:, indexes[0]]
            else:
                # Gabungkan kolom dengan base name yang sama
                combined_series = temp_df.iloc[:, indexes[0]].copy()
                for i in indexes[1:]:
                    combined_series = combined_series.combine_first(temp_df.iloc[:, i])
                final_data[base_name] = combined_series

        return pd.DataFrame(final_data)
    
    else:
        # Mode tidak dikenali, kembalikan dataframe asli tanpa perubahan
        return df
