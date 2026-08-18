"""Check if Калевские exists anywhere in the xlsx"""
import openpyxl, os

XLSX = r"C:\Users\GN\Desktop\doors_parser_ready_fix\data\Линейки производителей 3.xlsx"
wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb.active

# Search every cell for Калевские
found = []
for row in range(1, ws.max_row+1):
    for col in range(1, ws.max_column+1):
        val = ws.cell(row=row, column=col).value
        if val and "алевск" in str(val).lower():
            found.append({"row": row, "col": col, "value": str(val)})

print(f"'Калевские' found in xlsx: {len(found)} cells")
for f in found:
    print(f"  Row {f['row']}, Col {f['col']}: '{f['value']}'")

# Also check reference_lines_clean.json
import json
REF = r"C:\Users\GN\Desktop\doors_parser_ready_fix\output\final_rework_ready_v4\reference_lines_clean.json"
if os.path.exists(REF):
    with open(REF, encoding="utf-8") as f:
        ref = json.load(f)
    kalev_in_ref = {k: v for k, v in ref.items() if "алевск" in k.lower()}
    print(f"\nreference_lines_clean.json: Калевские producers = {list(kalev_in_ref.keys())}")
else:
    print("\nreference_lines_clean.json not found")

# Also check the old decoded JSON
REF_OLD = r"C:\Users\GN\Desktop\doors_parser_ready_fix\output\reference_lines_decoded.json"
with open(REF_OLD, encoding="utf-8") as f:
    ref_old = json.load(f)
kalev_in_old = {k: v for k, v in ref_old.items() if "алевск" in k.lower()}
print(f"reference_lines_decoded.json: Калевские producers = {list(kalev_in_old.keys())}")

# Check Slab decor in xlsx
slab_found = []
for row in range(1, ws.max_row+1):
    for col in range(1, ws.max_column+1):
        val = ws.cell(row=row, column=col).value
        if val and "slab" in str(val).lower():
            slab_found.append({"row": row, "col": col, "value": str(val)})
print(f"\n'Slab decor' found in xlsx: {len(slab_found)} cells")
for s in slab_found:
    print(f"  Row {s['row']}, Col {s['col']}: '{s['value']}'")

# Check Hausdoors 4 lines in xlsx
hd_lines = ["\u0424\u0430\u043a\u0442\u0443\u0440\u043d\u044b\u0435 \u043b\u0438\u043d\u0438\u0438", "\u041d\u043e\u0432\u0430\u044f \u043a\u043b\u0430\u0441\u0441\u0438\u043a\u0430",
            "\u041c\u0438\u043d\u0438\u043c\u0430\u043b\u0438\u0437\u043c", "\u0421\u043e\u0432\u0440\u0435\u043c\u0435\u043d\u043d\u0430\u044f \u043a\u043b\u0430\u0441\u0441\u0438\u043a\u0430"]
print(f"\nHausdoors 4 lines in xlsx:")
for line in hd_lines:
    found_hd = []
    for row in range(1, ws.max_row+1):
        for col in range(1, ws.max_column+1):
            val = ws.cell(row=row, column=col).value
            if val and line.lower() in str(val).lower():
                found_hd.append({"row": row, "col": col, "value": str(val)})
    print(f"  '{line}': {len(found_hd)} cells")
    for f in found_hd:
        print(f"    Row {f['row']}, Col {f['col']}: '{f['value']}'")
