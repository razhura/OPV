import pandas as pd
import re
from collections import defaultdict

def combine_duplicate_columns(df):
    # Tahap 1: Gabungkan kolom yang nama persis sama
    new_columns = []
    seen = defaultdict(list)

    for idx, col in enumerate(df.columns):
        seen[col].append(idx)

    final_data = {}

    for col, indexes in seen.items():
        if len(indexes) == 1:
            final_data[col] = df.iloc[:, indexes[0]]
        else:
            combined_series = df.iloc[:, indexes[0]].copy()
            for i in indexes[1:]:
                combined_series = combined_series.combine_first(df.iloc[:, i])
            final_data[col] = combined_series

    temp_df = pd.DataFrame(final_data)

    # Tahap 2: Gabungkan kolom yang nama dasarnya sama (hilangkan [teks], [nilai])
    def clean_column_name(name):
        return re.sub(r"\s*\[.*?\]\s*$", "", name).strip()

    base_name_map = defaultdict(list)
    for idx, col in enumerate(temp_df.columns):
        base_name = clean_column_name(col)
        base_name_map[base_name].append(idx)

    final_final_data = {}

    for base_name, indexes in base_name_map.items():
        if len(indexes) == 1:
            final_final_data[base_name] = temp_df.iloc[:, indexes[0]]
        else:
            combined_series = temp_df.iloc[:, indexes[0]].copy()
            for i in indexes[1:]:
                combined_series = combined_series.combine_first(temp_df.iloc[:, i])
            final_final_data[base_name] = combined_series

    final_df = pd.DataFrame(final_final_data)

    return final_df
