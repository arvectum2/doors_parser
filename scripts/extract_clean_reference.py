"""Extract reference lines correctly from xlsx"""
import openpyxl, csv, json, os

XLSX = r"C:\Users\GN\Desktop\doors_parser_ready_fix\data\Линейки производителей 3.xlsx"
OUT = r"C:\Users\GN\Desktop\doors_parser_ready_fix\output\final_rework_ready_v4"
os.makedirs(OUT, exist_ok=True)

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb.active

# Row 4 has producer names, row 5 has covering info
# Structure: producer in row 4, covering in row 5 (same column or nearby)
# Lines start from row 6

# Map column -> producer from row 4
producers = {}
for col in range(2, ws.max_column+1):
    val = ws.cell(row=4, column=col).value
    if val:
        producers[col] = str(val).strip()

print("Producers found:")
for col, name in sorted(producers.items()):
    print(f"  C{col}: {name}")

# Map column -> covering from row 5
coverings = {}
for col in range(2, ws.max_column+1):
    val = ws.cell(row=5, column=col).value
    if val:
        coverings[col] = str(val).strip()

# Extract lines (rows 6+)
# For each column, the producer is from row 4, lines from row 6+
all_lines = []
for col, producer in sorted(producers.items()):
    covering = coverings.get(col, "")
    for row in range(6, ws.max_row+1):
        val = ws.cell(row=row, column=col).value
        if val and str(val).strip():
            line_name = str(val).strip()
            all_lines.append({
                "producer": producer,
                "expected_line": line_name,
                "expected_covering": covering,
                "source_row": row,
                "source_col": col
            })

print(f"\nTotal expected lines: {len(all_lines)}")

# Group by producer
by_producer = {}
for l in all_lines:
    p = l["producer"]
    if p not in by_producer:
        by_producer[p] = []
    by_producer[p].append(l)

print(f"\nBy producer ({len(by_producer)}):")
for p in sorted(by_producer.keys()):
    lines = by_producer[p]
    line_names = [l["expected_line"] for l in lines]
    print(f"  {p}: {len(lines)} lines")
    for ln in line_names[:5]:
        print(f"    {ln}")
    if len(line_names) > 5:
        print(f"    ... and {len(line_names)-5} more")

# Verify no Калевские
kalev = [l for l in all_lines if "Калевские" in l["producer"] or "алевск" in l["producer"].lower()]
print(f"\nКалевские check: {len(kalev)} lines")

# Verify Slab decor
slab = [l for l in all_lines if "Slab" in l["expected_line"] or "slab" in l["expected_line"].lower()]
print(f"Slab decor check:")
for s in slab:
    print(f"  producer={s['producer']}, line={s['expected_line']}")

# Save clean reference
with open(os.path.join(OUT, "reference_lines_clean.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=["producer", "expected_line", "expected_covering", "source_row", "source_col"], delimiter=";")
    w.writeheader()
    w.writerows(all_lines)

# Save as JSON for easy use
ref_json = {}
for l in all_lines:
    p = l["producer"]
    if p not in ref_json:
        ref_json[p] = []
    ref_json[p].append({"line": l["expected_line"], "covering": l["expected_covering"]})

with open(os.path.join(OUT, "reference_lines_clean.json"), "w", encoding="utf-8") as f:
    json.dump(ref_json, f, ensure_ascii=False, indent=2)

# Save as xlsx
import openpyxl as xl
wb2 = xl.Workbook()
ws2 = wb2.active
ws2.title = "Clean Reference"
headers = ["producer", "expected_line", "expected_covering", "source_row", "source_col"]
from openpyxl.styles import Font, PatternFill
for ci, h in enumerate(headers, 1):
    c = ws2.cell(row=1, column=ci, value=h)
    c.font = Font(color="FFFFFF", bold=True)
    c.fill = PatternFill(start_color="4472C4", fill_type="solid")
for ri, l in enumerate(all_lines, 2):
    for ci, h in enumerate(headers, 1):
        ws2.cell(row=ri, column=ci, value=l[h])
wb2.save(os.path.join(OUT, "reference_lines_clean.xlsx"))

print(f"\nSaved: reference_lines_clean.csv, .json, .xlsx")
print(f"Total: {len(all_lines)} lines, {len(by_producer)} producers")
