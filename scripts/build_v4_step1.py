"""Load old decoded JSON and current catalog, find gaps, build v4"""
import csv, json, os, hashlib
from collections import Counter
from datetime import datetime
from zipfile import ZipFile
import openpyxl
from openpyxl.styles import Font, PatternFill

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CAT_V3 = os.path.join(BASE, "output", "final_rework_ready_v2", "doors_catalog.csv")
REF_OLD = os.path.join(BASE, "output", "reference_lines_decoded.json")
OUT = os.path.join(BASE, "output", "final_rework_ready_v4")
os.makedirs(OUT, exist_ok=True)
os.makedirs(os.path.join(OUT, "check_by_producer"), exist_ok=True)

# Load v3 catalog
with open(CAT_V3, encoding="utf-8-sig") as f:
    catalog = list(csv.DictReader(f, delimiter=";"))
field_names = list(catalog[0].keys())

# Load reference
with open(REF_OLD, encoding="utf-8") as f:
    ref_data = json.load(f)

# Build expected lines list
all_expected = []
for producer, entries in ref_data.items():
    for entry in entries:
        line = entry["line"] if isinstance(entry, dict) else str(entry)
        covering = entry.get("covering", "") if isinstance(entry, dict) else ""
        all_expected.append({"producer": producer, "line": line, "covering": covering})

print(f"Reference: {len(all_expected)} expected lines, {len(ref_data)} producers")

# Known fix: Slab decor belongs to Энтро, not Скамбио Порте
for exp in all_expected:
    if exp["line"] == "Slab decor" and exp["producer"] != "\u042d\u043d\u0442\u0440\u043e":
        old_prod = exp["producer"]
        exp["producer"] = "\u042d\u043d\u0442\u0440\u043e"
        print(f"FIXED: Slab decor moved from {old_prod} to Энтро")

# Check which lines are covered
covered = 0
uncovered_list = []
for exp in all_expected:
    producer = exp["producer"]
    line = exp["line"]
    ls = line.strip()
    
    # Find matching rows in catalog
    rows = [r for r in catalog if r["producer"] == producer and r["target_line"] == line]
    if not rows:
        rows = [r for r in catalog if r["producer"] == producer and r.get("target_line", "").strip() == ls]
    if not rows:
        rows = [r for r in catalog if r["producer"] == producer and ls.lower() in r.get("target_line", "").lower()]
    if not rows:
        rows = [r for r in catalog if r["producer"] == producer and r.get("target_line", "").lower() in ls.lower() and len(r.get("target_line", "")) > 2]
    
    # "Все модели" placeholder
    is_placeholder = "все модели" in ls.lower() or ls.lower().startswith("все ")
    if is_placeholder and not rows:
        rows = [r for r in catalog if r["producer"] == producer]
    
    if rows:
        covered += 1
    else:
        uncovered_list.append(exp)

print(f"Coverage: {covered}/{len(all_expected)}, missing: {len(uncovered_list)}")
for u in uncovered_list:
    print(f"  MISSING: {u['producer']}: {u['line']}")

# Add [MISSING] stubs for all uncovered lines
added = 0
for exp in uncovered_list:
    catalog.append({
        "door_type": "",
        "producer": exp["producer"],
        "target_line": exp["line"],
        "model_name": f"[MISSING] {exp['line']}",
        "description": "",
        "material": "",
        "coating": exp["covering"],
        "image_url": "",
        "source_url": "",
        "status": "needs_review",
        "comment": "source_line_not_found_or_no_models_on_site_after_scan",
        "image_url_front": "",
        "image_url_inside": ""
    })
    added += 1

print(f"Added {added} [MISSING] stubs")

# Verify Hausdoors 4 lines exist
hd_lines = ["\u0424\u0430\u043a\u0442\u0443\u0440\u043d\u044b\u0435 \u043b\u0438\u043d\u0438\u0438", "\u041d\u043e\u0432\u0430\u044f \u043a\u043b\u0430\u0441\u0441\u0438\u043a\u0430",
            "\u041c\u0438\u043d\u0438\u043c\u0430\u043b\u0438\u0437\u043c", "\u0421\u043e\u0432\u0440\u0435\u043c\u0435\u043d\u043d\u0430\u044f \u043a\u043b\u0430\u0441\u0441\u0438\u043a\u0430"]
print("\nHausdoors 4 lines check:")
for line in hd_lines:
    rows = [r for r in catalog if r["producer"] == "\u0425\u0430\u0443\u0441\u0434\u043e\u0440\u0441" and r["target_line"] == line]
    print(f"  {line}: {len(rows)} rows")

# Verify Entro Aria = 11
aria = [r for r in catalog if r["producer"] == "\u042d\u043d\u0442\u0440\u043e" and "aria" in r.get("target_line", "").lower()]
print(f"\nEntro Aria: {len(aria)} rows")

# Verify ASD = 199
asd = [r for r in catalog if r["producer"] == "\u0410\u0421\u0414"]
print(f"ASD: {len(asd)} rows")

# Dedup
seen = set()
deduped = []
for r in catalog:
    key = (r["producer"], r["target_line"], r["model_name"])
    if key not in seen:
        seen.add(key)
        deduped.append(r)
catalog = deduped
print(f"After dedup: {len(catalog)} rows")

# Final coverage check
covered_final = 0
for exp in all_expected:
    producer = exp["producer"]
    line = exp["line"]
    rows = [r for r in catalog if r["producer"] == producer and (r["target_line"] == line or exp["line"].lower() in r.get("target_line", "").lower())]
    is_placeholder = "все модели" in line.lower() or line.lower().startswith("все ")
    if is_placeholder:
        rows = [r for r in catalog if r["producer"] == producer]
    if rows:
        covered_final += 1

print(f"\nFinal coverage: {covered_final}/{len(all_expected)}")

# Save
with open(os.path.join(OUT, "doors_catalog.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=field_names, delimiter=";")
    w.writeheader()
    w.writerows(catalog)

print(f"Saved: {len(catalog)} rows")
