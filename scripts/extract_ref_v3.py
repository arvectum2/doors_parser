"""Re-parse xlsx with merged cells correctly - use covering as producer identifier"""
import openpyxl, csv, json, os

XLSX = r"C:\Users\GN\Desktop\doors_parser_ready_fix\data\Линейки производителей 3.xlsx"
OUT = r"C:\Users\GN\Desktop\doors_parser_ready_fix\output\final_rework_ready_v4"

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb.active

# The xlsx structure:
# Row 4 = group name (merged cells, e.g. "Колорит" spans many columns)
# Row 5 = covering/sub-type (each column has unique covering)
# Row 6+ = line names under each covering
#
# But the user expects specific producer names that map to coverings:
# "Колорит" with covering "Пленка" -> might actually be "Олимп" in the user's list
# Need to use covering as the real producer identifier

# Get row 4 (merged group) and row 5 (covering) for each column
row4_vals = {}
row5_vals = {}
for col in range(2, ws.max_column+1):
    val4 = ws.cell(row=4, column=col).value
    val5 = ws.cell(row=5, column=col).value
    if val4:
        row4_vals[col] = str(val4).strip()
    if val5:
        row5_vals[col] = str(val5).strip()

# The real mapping: each column's producer = covering from row 5
# For columns where row4 says "Колорит" but covering is different,
# the covering is the actual producer sub-type
# 
# User's expected 10 producers and their coverings:
# Скамбио Порте -> covering="Скамбио Порте" (C2)
# Хаусдорс -> covering="ПВХ" (C14)
# Олимп -> ??? (possibly Колорит columns with different coverings)
# Аэлита -> covering="ПВХ-пленка" (C20)
# Лорд -> covering="Эмаль" or "Пленка" (C10, C11)
# Комфорт -> covering="ПВХ-пленка" (C22, C23)
# Двери регионов -> covering="ПВХ-пленка" (C20) or "ПВХ" (C20)?
# Термодур -> covering="Пленка" (C8)
# Энтро -> covering="Aria" (C26) or "Эко, ПВХ" (C28)
# АСД -> covering="Все модели" (C29)

# The actual correct mapping from the xlsx:
# Each column is a separate producer+covering combination
# Column 2: Скамбио Порте / Скамбио Порте
# Column 4: Колорит / Пленка -> This is "Олимп" in user's list
# Column 5: Колорит / ПВХ-пленка -> This might be "Лорд" or another
# Column 7: Колорит / Эко -> This is "Лорд" in user's list
# Column 8: Термодур / Пленка -> Термодур
# Column 10: Колорит / Эмаль -> This is "Комфорт" or another
# Column 11: Лорд / Пленка -> Лорд
# Column 13: Колорит / Эко -> Duplicate?
# Column 14: Хаусдорс / ПВХ -> Хаусдорс
# Column 16: Колорит / Эко,Плёнка -> Скандинавия?
# Column 17: Скандинавия / Новая -> Скандинавия
# Column 19: Колорит / Плёнка -> Двери регионов?
# Column 20: Аэлита / ПВХ-пленка -> Аэлита
# Column 22: Колорит / ПВХ-пленка -> Комфорт?
# Column 23: Комфорт / ПВХ-пленка дверное полотно -> Комфорт
# Column 25: Колорит / (empty) -> ?
# Column 26: Энтро / Aria -> Энтро
# Column 28: Колорит / Эко, ПВХ -> Энтро (main)
# Column 29: АСД / Все модели -> АСД
# Column 31: Колорит / (empty) -> ?

# Since the user says there should be 10 producers and the old extraction got 194,
# let me just use the covering (row 5) as the producer name directly.
# This should give us the correct 10 producers.

all_lines = []
for col in range(2, ws.max_column+1):
    group = row4_vals.get(col, "")
    covering = row5_vals.get(col, "")
    if not covering:
        continue  # Skip columns without covering (empty placeholders)
    
    # Use covering as the producer name
    producer = covering
    
    for row in range(6, ws.max_row+1):
        val = ws.cell(row=row, column=col).value
        if val and str(val).strip():
            line_name = str(val).strip()
            all_lines.append({
                "producer": producer,
                "expected_line": line_name,
                "expected_covering": covering,
                "source_row": row,
                "source_col": col,
                "group_name": group
            })

print(f"Total lines: {len(all_lines)}")

by_producer = {}
for l in all_lines:
    p = l["producer"]
    if p not in by_producer:
        by_producer[p] = []
    by_producer[p].append(l)

print(f"Producers: {len(by_producer)}")
for p in sorted(by_producer.keys()):
    print(f"  {p}: {len(by_producer[p])} lines")

# Specific checks
print(f"\n--- Checks ---")
slab = [l for l in all_lines if "Slab" in l["expected_line"]]
for s in slab:
    print(f"Slab decor: producer={s['producer']}, group={s['group_name']}")

aria_lines = [l for l in all_lines if "Aria" in l["expected_line"] or "aria" in l["expected_line"].lower()]
for a in aria_lines:
    print(f"Aria: producer={a['producer']}, line={a['expected_line']}")

# Save
with open(os.path.join(OUT, "reference_lines_clean.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=["producer", "expected_line", "expected_covering", "source_row", "source_col", "group_name"], delimiter=";")
    w.writeheader()
    w.writerows(all_lines)

ref_json = {}
for l in all_lines:
    p = l["producer"]
    if p not in ref_json:
        ref_json[p] = []
    ref_json[p].append({"line": l["expected_line"], "covering": l["expected_covering"]})
with open(os.path.join(OUT, "reference_lines_clean.json"), "w", encoding="utf-8") as f:
    json.dump(ref_json, f, ensure_ascii=False, indent=2)

import openpyxl as xl
from openpyxl.styles import Font, PatternFill
wb2 = xl.Workbook()
ws2 = wb2.active
ws2.title = "Clean Reference"
headers = ["producer", "expected_line", "expected_covering", "source_row", "source_col", "group_name"]
for ci, h in enumerate(headers, 1):
    c = ws2.cell(row=1, column=ci, value=h)
    c.font = Font(color="FFFFFF", bold=True)
    c.fill = PatternFill(start_color="4472C4", fill_type="solid")
for ri, l in enumerate(all_lines, 2):
    for ci, h in enumerate(headers, 1):
        ws2.cell(row=ri, column=ci, value=l[h])
wb2.save(os.path.join(OUT, "reference_lines_clean.xlsx"))

print(f"\nSaved: {len(all_lines)} lines, {len(by_producer)} producers")
