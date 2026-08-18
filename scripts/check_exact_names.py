"""Check exact producer names in xlsx row 4"""
import openpyxl

XLSX = r"C:\Users\GN\Desktop\doors_parser_ready_fix\data\Линейки производителей 3.xlsx"
wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb.active

print("Row 4 (producers) - all cells:")
for col in range(1, ws.max_column+1):
    val = ws.cell(row=4, column=col).value
    if val:
        print(f"  C{col}: '{val}' (type={type(val).__name__})")

print("\nRow 5 (coverings) - all cells:")
for col in range(1, ws.max_column+1):
    val = ws.cell(row=5, column=col).value
    if val:
        print(f"  C{col}: '{val}'")

# Check for merged cells
print(f"\nMerged cells: {ws.merged_cells.ranges}")

# Check if there are hidden rows/columns
print(f"\nHidden columns: {[col for col in range(1, ws.max_column+1) if ws.column_dimensions[openpyxl.utils.get_column_letter(col)].hidden]}")

# Also check row 3 for title
print("\nRow 3:")
for col in range(1, ws.max_column+1):
    val = ws.cell(row=3, column=col).value
    if val:
        print(f"  C{col}: '{val}'")
