"""Check duplicate producer columns in xlsx"""
import openpyxl

XLSX = r"C:\Users\GN\Desktop\doors_parser_ready_fix\data\Линейки производителей 3.xlsx"
wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb.active

# Map each data column to its producer
producers = {}
for col in range(2, ws.max_column+1):
    val = ws.cell(row=4, column=col).value
    if val:
        producers[col] = str(val).strip()

# Check which producers appear in multiple columns
from collections import Counter
prod_counts = Counter(producers.values())
dups = {p: c for p, c in prod_counts.items() if c > 1}
print("Duplicate producers:")
for p, c in dups.items():
    cols = [col for col, name in producers.items() if name == p]
    print(f"  {p}: columns {cols}")

# For each duplicate, check if lines overlap or are distinct
for p, cols in [(p, [col for col, name in producers.items() if name == p]) for p in dups]:
    print(f"\n--- {p} (cols {cols}) ---")
    all_lines = {}
    for col in cols:
        covering = ws.cell(row=5, column=col).value or ""
        lines = []
        for row in range(6, ws.max_row+1):
            val = ws.cell(row=row, column=col).value
            if val and str(val).strip():
                lines.append(str(val).strip())
        print(f"  Col {col} (covering={covering}): {len(lines)} lines")
        for l in lines[:3]:
            print(f"    {l}")
        if len(lines) > 3:
            print(f"    ... and {len(lines)-3} more")
        for l in lines:
            if l not in all_lines:
                all_lines[l] = []
            all_lines[l].append(col)
    
    # Check overlap
    overlap = [l for l, cols_list in all_lines.items() if len(cols_list) > 1]
    print(f"  Overlapping lines: {len(overlap)}")
    if overlap:
        for o in overlap[:3]:
            print(f"    {o}")

# Also check row 5 covering values for each column
print("\n--- Covering values (row 5) ---")
for col, prod in sorted(producers.items()):
    cov = ws.cell(row=5, column=col).value
    if cov:
        print(f"  C{col} ({prod}): {cov}")
