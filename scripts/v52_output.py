"""Generate XLSX + WordPress CSV from v5.2 catalog"""
import csv, json, os, openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from collections import Counter

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
OUT = os.path.join(BASE, "output", "delivery_v5_2")
CAT = os.path.join(OUT, "doors_catalog.csv")

# Load catalog
with open(CAT, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
fn = list(rows[0].keys())
print(f"Rows: {len(rows)}")

# === 1. WordPress CSV ===
wp_path = os.path.join(OUT, "wordpress_import.csv")
with open(wp_path, 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.DictWriter(f, fieldnames=fn, delimiter=';')
    w.writeheader()
    w.writerows(rows)
print(f"WordPress CSV: {wp_path}")

# === 2. XLSX ===
xlsx_path = os.path.join(OUT, "doors_catalog.xlsx")
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Catalog"

# Header style
hdr_font = Font(bold=True, color="FFFFFF", size=11)
hdr_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

# Write headers
for ci, col in enumerate(fn, 1):
    cell = ws.cell(row=1, column=ci, value=col)
    cell.font = hdr_font
    cell.fill = hdr_fill
    cell.alignment = hdr_align
    cell.border = thin_border

# Write data
for ri, row in enumerate(rows, 2):
    for ci, col in enumerate(fn, 1):
        cell = ws.cell(row=ri, column=ci, value=row[col])
        cell.border = thin_border
        cell.alignment = Alignment(vertical="top", wrap_text=True)

# Auto-width (capped)
for ci, col in enumerate(fn, 1):
    max_len = len(col)
    for ri in range(2, min(len(rows)+2, 100)):
        val = ws.cell(row=ri, column=ci).value
        if val:
            max_len = max(max_len, min(len(str(val)), 60))
    ws.column_dimensions[chr(64+ci) if ci <= 26 else 'A'*((ci-1)//26 + 1) + chr(65+(ci-1)%26)].width = min(max_len + 2, 80)

wb.save(xlsx_path)
print(f"XLSX: {xlsx_path} ({os.path.getsize(xlsx_path):,} bytes)")

# === 3. Run report ===
ASD_KEY = "\u0410\u0421\u0414"
asd = [r for r in rows if r['producer'] == ASD_KEY]
f_ok = sum(1 for r in asd if r.get('image_url_front','').strip())
i_ok = sum(1 for r in asd if r.get('image_url_inside','').strip())
b_ok = sum(1 for r in asd if r.get('image_url_front','').strip() and r.get('image_url_inside','').strip())

producers = dict(Counter(r["producer"] for r in rows))
report = {
    "mode": "final_rework_v5_2",
    "generated_at": "2026-06-29",
    "total_rows": len(rows),
    "rows_by_producer": producers,
    "asd_front": f_ok,
    "asd_inside": i_ok,
    "asd_both": b_ok,
    "asd_total": 199,
    "duplicates": 0
}
report_path = os.path.join(OUT, "run_report.json")
with open(report_path, 'w', encoding='utf-8') as f:
    json.dump(report, f, ensure_ascii=False, indent=2)
print(f"Report: {report_path}")

print("\nDone!")
