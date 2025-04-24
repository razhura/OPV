from openpyxl import load_workbook

def extract_multi_level_headers(excel_file, start_row=4, num_levels=3):
    wb = load_workbook(excel_file, data_only=True)
    ws = wb.active

    headers = []
    max_col = ws.max_column

    # Buat list header per kolom
    for col_idx in range(1, max_col + 1):
        col_header = []
        for row_offset in range(num_levels):
            row = start_row + row_offset
            cell = ws.cell(row=row, column=col_idx)

            # Deteksi merge cell
            for merged_range in ws.merged_cells.ranges:
                if (cell.coordinate in merged_range):
                    cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                    break

            value = str(cell.value).strip() if cell.value else ""
            col_header.append(value)

        # Gabung jadi satu string header bertingkat
        combined = " > ".join([h for h in col_header if h])
        headers.append(combined)

    return headers
