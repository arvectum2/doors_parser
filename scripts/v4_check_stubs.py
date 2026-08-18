"""Step 1: Load v4 catalog, find ALL [MISSING] stubs"""
import csv, zipfile, io, json, os

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
ZIP = os.path.join(BASE, "output", "doors_parser_final_REWORK_READY_v4_2026-06-23_17-08-14.zip")

with zipfile.ZipFile(ZIP) as zf:
    csv_name = [n for n in zf.namelist() if n.endswith("doors_catalog.csv")][0]
    with zf.open(csv_name) as f:
        catalog = list(csv.DictReader(io.StringIO(f.read().decode("utf-8-sig")), delimiter=";"))

print(f"Total rows: {len(catalog)}")

# Find ALL [MISSING] stubs
missing = [r for r in catalog if "[MISSING]" in r.get("model_name", "")]
print(f"\n[MISSING] stubs: {len(missing)}")
for r in missing:
    print(f"  producer={r['producer'][:15]} | target_line={r['target_line'][:25]} | model={r['model_name'][:35]} | status={r.get('status','')} | comment={r.get('comment','')[:40]}")

# Find all needs_review with source_line keywords
needs_review_source = [r for r in catalog if r.get("status","") == "needs_review" and 
    any(kw in (r.get("comment","") or "").lower() for kw in ["source_line", "not_found", "missing", "no_models"])]
print(f"\nneeds_review with source keywords: {len(needs_review_source)}")
for r in needs_review_source:
    if "[MISSING]" not in r.get("model_name", ""):
        print(f"  producer={r['producer'][:15]} | target_line={r['target_line'][:25]} | model={r['model_name'][:35]}")

# Check Hausdoors specifically
hd = [r for r in catalog if r["producer"] == "\u0425\u0430\u0443\u0441\u0434\u043e\u0440\u0441"]
print(f"\nHausdorss total: {len(hd)}")
hd_missing = [r for r in hd if "[MISSING]" in r.get("model_name", "")]
print(f"Hausdorss [MISSING]: {len(hd_missing)}")
for r in hd_missing:
    print(f"  {r['target_line'][:25]} | {r['model_name'][:35]} | {r.get('status','')}")

# Check Калевские
kalev = [r for r in catalog if "\u041a\u0430\u043b\u0435\u0432\u0441\u043a" in r.get("producer", "")]
print(f"\nКалевские rows: {len(kalev)}")
for r in kalev:
    print(f"  {r['target_line'][:25]} | {r['model_name'][:35]} | {r.get('status','')}")

# Check Slab decor
slab = [r for r in catalog if "Slab" in r.get("target_line", "") or "slab" in r.get("target_line", "").lower()]
print(f"\nSlab decor rows: {len(slab)}")
for r in slab:
    print(f"  producer={r['producer'][:15]} | target_line={r['target_line'][:25]} | model={r['model_name'][:35]}")
