import pandas as pd
from collections import defaultdict

def combine_duplicate_columns(df):
    new_columns = []
    seen = defaultdict(list)

    # Kumpulkan posisi kolom berdasarkan nama
    for idx, col in enumerate(df.columns):
        seen[col].append(idx)

    final_data = {}

    for col, indexes in seen.items():
        if len(indexes) == 1:
            # Jika hanya satu kolom, langsung ambil
            final_data[col] = df.iloc[:, indexes[0]]
        else:
            # Gabungkan kolom duplikat dengan prioritas kiri
            combined_series = df.iloc[:, indexes[0]].copy()
            for i in indexes[1:]:
                combined_series = combined_series.combine_first(df.iloc[:, i])
            final_data[col] = combined_series

    # Buat DataFrame baru dengan kolom unik
    new_df = pd.DataFrame(final_data)
    return new_df
