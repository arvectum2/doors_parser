"""Parse Линейки производителей 3.xlsx correctly - find exact structure"""
import openpyxl

XLSX = r"C:\Users\GN\Desktop\doors_parser_ready_fix\data\Линейки производителей 3.xlsx"
wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb.active

print(f"Sheet: {ws.title}")
print(f"Dimensions: {ws.dimensions}")
print(f"Max row: {ws.max_row}, Max col: {ws.max_column}")

# Print rows 1-10 to understand structure
print("\n--- Rows 1-10 ---")
for row_num in range(1, 11):
    cells = []
    for col_num in range(1, min(ws.max_column+1, 30)):
        val = ws.cell(row=row_num, column=col_num).value
        if val is not None:
            cells.append(f"C{col_num}={val}")
    if cells:
        print(f"Row {row_num}: {'; '.join(cells)}")

# Print ALL non-empty cells in rows 4-5 to find producer names
print("\n--- Row 4 (producers) ---")
for col_num in range(1, ws.max_column+1):
    val = ws.cell(row=4, column=col_num).value
    if val is not None:
        print(f"  C{col_num}: {val}")

print("\n--- Row 5 (first data row) ---")
for col_num in range(1, ws.max_column+1):
    val = ws.cell(row=5, column=col_num).value
    if val is not None:
        print(f"  C{col_num}: {val}")
