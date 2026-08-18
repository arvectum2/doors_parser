"""Extract clean reference: merge duplicate producer columns, get unique lines"""
import openpyxl, csv, json, os

XLSX = r"C:\Users\GN\Desktop\doors_parser_ready_fix\data\Линейки производителей 3.xlsx"
OUT = r"C:\Users\GN\Desktop\doors_parser_ready_fix\output\final_rework_ready_v4"
os.makedirs(OUT, exist_ok=True)

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb.active

# Get producers from row 4
producers = {}
for col in range(2, ws.max_column+1):
    val = ws.cell(row=4, column=col).value
    if val:
        producers[col] = str(val).strip()

# Get covering from row 5
coverings = {}
for col in range(2, ws.max_column+1):
    val = ws.cell(row=5, column=col).value
    if val:
        coverings[col] = str(val).strip()

# Group columns by producer
prod_cols = {}
for col, prod in producers.items():
    if prod not in prod_cols:
        prod_cols[prod] = []
    prod_cols[prod].append(col)

print("Producer -> columns mapping:")
for prod, cols in sorted(prod_cols.items()):
    covs = [coverings.get(c, "") for c in cols]
    print(f"  {prod}: cols={cols}, coverings={covs}")

# Extract lines: for each producer, merge all columns
# If a line appears in multiple columns under the same producer, keep it once
# but record all coverings it appears under
all_lines = []
for prod, cols in sorted(prod_cols.items()):
    seen_lines = set()
    for col in cols:
        covering = coverings.get(col, "")
        for row in range(6, ws.max_row+1):
            val = ws.cell(row=row, column=col).value
            if val and str(val).strip():
                line_name = str(val).strip()
                # Use (producer, line_name) as unique key
                key = (prod, line_name)
                if key not in seen_lines:
                    seen_lines.add(key)
                    all_lines.append({
                        "producer": prod,
                        "expected_line": line_name,
                        "expected_covering": covering,
                        "source_row": row,
                        "source_col": col
                    })

print(f"\nTotal unique lines: {len(all_lines)}")

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
    print(f"  {p}: {len(lines)} lines")
    for l in lines[:3]:
        print(f"    {l['expected_line']} (covering={l['expected_covering']})")
    if len(lines) > 3:
        print(f"    ... and {len(lines)-3} more")

# Specific checks
print(f"\n--- Checks ---")
kalev = [l for l in all_lines if "алевск" in l["producer"].lower()]
print(f"Калевские: {len(kalev)} lines (should be 0)")

slab = [l for l in all_lines if "Slab" in l["expected_line"]]
for s in slab:
    print(f"Slab decor: producer={s['producer']}, line={s['expected_line']}")

aria = [l for l in all_lines if "Aria" in l["expected_line"]]
for a in aria:
    print(f"Aria: producer={a['producer']}, line={a['expected_line']}")

# Check expected 10 producers from user's list
expected_producers = ["Скамбио Порте", "Хаусдорс", "Олимп", "Аэлита", "Лорд",
                      "Комфорт", "Двери регионов", "Термодур", "Энтро", "АСД"]
found_producers = list(by_producer.keys())
print(f"\nExpected producers: {expected_producers}")
print(f"Found producers: {found_producers}")
missing = [p for p in expected_producers if p not in found_producers]
extra = [p for p in found_producers if p not in expected_producers]
print(f"Missing: {missing}")
print(f"Extra: {extra}")

# Save
with open(os.path.join(OUT, "reference_lines_clean.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=["producer", "expected_line", "expected_covering", "source_row", "source_col"], delimiter=";")
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
headers = ["producer", "expected_line", "expected_covering", "source_row", "source_col"]
for ci, h in enumerate(headers, 1):
    c = ws2.cell(row=1, column=ci, value=h)
    c.font = Font(color="FFFFFF", bold=True)
    c.fill = PatternFill(start_color="4472C4", fill_type="solid")
for ri, l in enumerate(all_lines, 2):
    for ci, h in enumerate(headers, 1):
        ws2.cell(row=ri, column=ci, value=l[h])
wb2.save(os.path.join(OUT, "reference_lines_clean.xlsx"))

print(f"\nSaved clean reference: {len(all_lines)} lines, {len(by_producer)} producers")
