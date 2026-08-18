"""Build v16.2 — clear 5 fragment images retained in v16.1"""
import csv, os, json, hashlib, shutil
from collections import Counter
from datetime import datetime
from zipfile import ZipFile

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
V161_DIR = os.path.join(BASE, "output", "final_hybrid_client_ready_v16_asd_full_door_only")
V161_CSV = os.path.join(V161_DIR, "doors_catalog.csv")
V161_XLSX = os.path.join(V161_DIR, "doors_catalog.xlsx")
V161_WP = os.path.join(V161_DIR, "wordpress_import.csv")
V161_REPORT = os.path.join(V161_DIR, "run_report.json")
V161_CHECK_DIR = os.path.join(V161_DIR, "check_by_producer")
V161_CLEAN_DIR = os.path.join(V161_DIR, "output", "asd_photo_patch_v2_clean")
V161_NOT_FOUND = os.path.join(V161_DIR, "not_found.csv")

CLEAN_CSV = os.path.join(BASE, "output", "asd_photo_patch_v2_clean", "asd_photo_patch_v2_clean.csv")

# The 5 fragment rows (exact model_name + source_url)
# We identify by model_name only and verify the URL matches the fragment pattern
FRAGMENT_NAMES = [
    "Стальная дверь \"Аякс\". Заказная модель",
    "Стальная дверь \"Барселона\"",
    "Стальная дверь \"Милан\". Заказная модель",
    "Стальная дверь \"Стелла\". Заказная модель",
    "Стальная дверь \"Стелла\" с зеркалом. Заказная модель",
]

print("Loading v16.1 catalog...")
with open(V161_CSV, encoding="utf-8-sig") as f:
    v161_rows = list(csv.DictReader(f, delimiter=";"))

FIELD_NAMES = v161_rows[0].keys() if v161_rows else [
    "door_type","producer","target_line","model_name","description",
    "material","coating","image_url","source_url","status","comment"
]

# Find and fix the 5 fragment rows
fix_count = 0
for r in v161_rows:
    if r["producer"] == "АСД":
        mn = r.get("model_name","") or ""
        if mn in FRAGMENT_NAMES:
            old_url = r.get("image_url","") or ""
            # Verify it's actually a cache/fragment URL (safety check)
            if "cache_image" in old_url or "kollazh" in old_url.lower() or "collage" in old_url.lower() or old_url.endswith(".webp"):
                print(f"  FIXED: {mn[:35]:35} cleared: {old_url[:55]}")
                r["image_url"] = ""
                r["status"] = "needs_review"
                old_comment = r.get("comment","") or ""
                new_comment = old_fragment_image_cleaned = f"old_fragment_image_cleared_no_full_door_candidate; {old_comment}" if old_comment else "old_fragment_image_cleared_no_full_door_candidate"
                r["comment"] = new_comment
                fix_count += 1
            else:
                print(f"  SKIPPED (not fragment): {mn[:35]} url={old_url[:55]}")

print(f"\nFixed rows: {fix_count}")

# ─── Verifications ───
errors = []
if len(v161_rows) != 3084:
    errors.append(f"Total rows: {len(v161_rows)} (expected 3084)")

producers = [r["producer"] for r in v161_rows]
producer_counts = Counter(producers)
if len(producer_counts) != 10:
    errors.append(f"Producer count: {len(producer_counts)} (expected 10)")

seen, dups = set(), []
for r in v161_rows:
    key = (r.get("producer",""), r.get("target_line",""), r.get("model_name",""))
    if key in seen:
        dups.append(key)
    seen.add(key)
if dups:
    errors.append(f"Duplicates: {len(dups)}")

# Non-ASD unchanged check
v14_csv = os.path.join(BASE, "output", "final_hybrid_client_ready_v14", "doors_catalog.csv")
with open(v14_csv, encoding="utf-8-sig") as f:
    v14_rows = list(csv.DictReader(f, delimiter=";"))

for p, c in Counter(r["producer"] for r in v14_rows).items():
    if p != "АСД":
        v162_c = producer_counts.get(p, 0)
        if v162_c != c:
            errors.append(f"Producer {p}: v14={c} v162={v162_c}")

for i, (v14_r, v162_r) in enumerate(zip(v14_rows, v161_rows)):
    if v14_r["producer"] != "АСД":
        for col in ["door_type","producer","target_line","model_name","image_url","source_url"]:
            if v14_r.get(col,"") != v162_r.get(col,""):
                errors.append(f"Row {i+2} non-ASD changed: {col}")
                break

# Check ASD bad patterns
BAD_PATTERNS = ["interier","interior","banner","badge","hit","logo","icon",
                "ruchka","ручка","handle","lock","zamok","замок","fragment"]
for r in v161_rows:
    if r["producer"] == "АСД":
        url = r.get("image_url","") or ""
        if any(b in url.lower() for b in BAD_PATTERNS):
            errors.append(f"ASD bad pattern: {r['model_name'][:25]} {url[:55]}")

# Check no fragment visual_decision in remaining ASD URLs
debug_csv = os.path.join(BASE, "output", "asd_photo_patch_v2", "asd_photo_candidates_debug.csv")
with open(debug_csv, encoding="utf-8-sig") as f:
    debugs = list(csv.DictReader(f, delimiter=";"))

debug_selected = {}
for d in debugs:
    if d.get("is_selected") == "YES":
        debug_selected[(d["model_name"], d["source_url"])] = d

fragment_remaining = []
for r in v161_rows:
    if r["producer"] == "АСД":
        key = (r["model_name"], r["source_url"])
        url = r.get("image_url","") or ""
        if url and key in debug_selected:
            d = debug_selected[key]
            if d.get("visual_decision","") not in ("full_door", "full_door_short"):
                fragment_remaining.append((r, d))

if fragment_remaining:
    errors.append(f"ASD rows with fragment visual_decision remaining: {len(fragment_remaining)}")
    for r, d in fragment_remaining:
        errors.append(f"  remaining fragment: {r['model_name'][:35]} url={r.get('image_url','')[:50]}")

# ASD counts
asd_v162 = [r for r in v161_rows if r["producer"] == "АСД"]
asd_with_img = [r for r in asd_v162 if r.get("image_url","")]
asd_without = [r for r in asd_v162 if not r.get("image_url","")]
print(f"ASD with image: {len(asd_with_img)}, without: {len(asd_without)}")

if errors:
    print(f"\n=== ERRORS ({len(errors)}) ===")
    for e in errors:
        print(f"  FAIL: {e}")
else:
    print("\n=== ALL VERIFICATIONS PASSED ===")

# ─── Save doors_catalog.csv ───
with open(V161_CSV, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=list(FIELD_NAMES), delimiter=";")
    w.writeheader()
    w.writerows(v161_rows)
print(f"\ndoors_catalog.csv: {len(v161_rows)} rows")

# XLSX
try:
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Catalog v16.2"
    for ci, h in enumerate(FIELD_NAMES, 1):
        ws.cell(row=1, column=ci, value=h)
    for ri, r in enumerate(v161_rows, 2):
        for ci, h in enumerate(FIELD_NAMES, 1):
            ws.cell(row=ri, column=ci, value=r.get(h, ""))
    wb.save(V161_XLSX)
    print("doors_catalog.xlsx: saved")
except ImportError:
    print("XLSX: skipped")

# WordPress import
if os.path.exists(V161_WP):
    with open(V161_WP, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f, delimiter=";")
        wp_fields = reader.fieldnames
        wp_rows = list(reader)
    asd_lookup = {(r["model_name"], r["source_url"]): r for r in v161_rows if r["producer"] == "АСД"}
    for wp_r in wp_rows:
        key = (wp_r.get("model_name",""), wp_r.get("source_url",""))
        if key in asd_lookup:
            wp_r["image_url"] = asd_lookup[key].get("image_url","")
    with open(V161_WP, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=wp_fields, delimiter=";")
        w.writeheader()
        w.writerows(wp_rows)
    wp_count = len(wp_rows)
    print(f"wordpress_import.csv: {wp_count} rows")
    if wp_count != len(v161_rows):
        errors.append(f"WP rows {wp_count} != catalog {len(v161_rows)}")

# Run report
report_data = {
    "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "version": "v16.2_ASD_FULL_DOOR_ONLY",
    "base": "v16.1",
    "fixes": f"cleared_5_fragment_images_from_not_found_rows",
    "total_rows": len(v161_rows),
    "producers_count": len(producer_counts),
    "producers": dict(producer_counts),
    "asd": {
        "total": len(asd_v162),
        "with_image_url": len(asd_with_img),
        "without_image_url": len(asd_without),
        "fragment_images_cleared": fix_count,
    },
    "errors": errors
}
with open(V161_REPORT, "w", encoding="utf-8") as f:
    json.dump(report_data, f, ensure_ascii=False, indent=2)
print("run_report.json: saved")

# Update check_by_producer/АСД.csv
asd_check_csv = os.path.join(V161_CHECK_DIR, "АСД.csv")
if os.path.exists(asd_check_csv):
    with open(asd_check_csv, encoding="utf-8-sig") as f:
        check_rows = list(csv.DictReader(f, delimiter=";"))
    check_fields = list(check_rows[0].keys()) if check_rows else []
    asd_lookup_cat = {(r["model_name"], r["source_url"]): r for r in v161_rows if r["producer"] == "АСД"}
    for cr in check_rows:
        key = (cr.get("model_name",""), cr.get("source_url",""))
        if key in asd_lookup_cat:
            cat_r = asd_lookup_cat[key]
            cr["image_url"] = cat_r.get("image_url","")
            cr["status"] = cat_r.get("status","")
            cr["comment"] = cat_r.get("comment","")
    with open(asd_check_csv, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=check_fields, delimiter=";")
        w.writeheader()
        w.writerows(check_rows)
    print("check_by_producer/АСД.csv: updated")

# Update check_by_producer/АСД.xlsx
asd_check_xlsx = os.path.join(V161_CHECK_DIR, "АСД.xlsx")
if os.path.exists(asd_check_xlsx):
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "АСД v16.2"
        for ci, h in enumerate(check_fields, 1):
            ws.cell(row=1, column=ci, value=h)
        for ri, cr in enumerate(check_rows, 2):
            for ci, h in enumerate(check_fields, 1):
                ws.cell(row=ri, column=ci, value=cr.get(h, ""))
        wb.save(asd_check_xlsx)
        print("check_by_producer/АСД.xlsx: updated")
    except ImportError:
        pass

# Update clean patch CSV to reflect these 5 rows were cleared
if os.path.exists(CLEAN_CSV):
    with open(CLEAN_CSV, encoding="utf-8-sig") as f:
        clean_rows = list(csv.DictReader(f, delimiter=";"))
    clean_fields = list(clean_rows[0].keys()) if clean_rows else []
    updated_clean = 0
    for cr in clean_rows:
        mn = cr.get("model_name","") or ""
        if mn in FRAGMENT_NAMES:
            if cr.get("selected_image_url","") and ("cache_image" in cr["selected_image_url"] or "kollazh" in cr["selected_image_url"].lower()):
                # These should already have empty selected_image_url (cleared by sanitization)
                pass
            # Add a note about fragment clearing
            cr["patch_comment"] = (cr.get("patch_comment","") or "") + "; old_v14_fragment_image_cleared_in_v162"
            updated_clean += 1
    if updated_clean:
        with open(CLEAN_CSV, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=clean_fields, delimiter=";")
            w.writeheader()
            w.writerows(clean_rows)
        print(f"Clean patch: {updated_clean} rows updated")
else:
    print("Clean patch CSV: not found")

# ─── Create ZIP ───
V162_BASE = V161_DIR  # Same directory, updated content
zip_name = f"doors_parser_final_HYBRID_CLIENT_READY_v16_2_ASD_FULL_DOOR_ONLY_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.zip"
zip_path = os.path.join(BASE, "output", zip_name)

with ZipFile(zip_path, "w") as zf:
    for root, dirs, files in os.walk(V162_BASE):
        for fn in files:
            fp = os.path.join(root, fn)
            arcname = os.path.relpath(fp, V162_BASE)
            zf.write(fp, arcname=os.path.join("output", "final_hybrid_client_ready_v16_asd_full_door_only", arcname))
    # Also add asd_photo_patch_v2_clean at top level
    v16_clean_dir = os.path.join(V162_BASE, "output", "asd_photo_patch_v2_clean")
    if os.path.exists(v16_clean_dir):
        for fn in os.listdir(v16_clean_dir):
            fp = os.path.join(v16_clean_dir, fn)
            zf.write(fp, arcname=os.path.join("output", "asd_photo_patch_v2_clean", fn))

sha = hashlib.sha256()
with open(zip_path, "rb") as f:
    for chunk in iter(lambda: f.read(65536), b""):
        sha.update(chunk)

print(f"\n{'='*60}")
print(f"ZIP: {zip_path}")
print(f"SHA256: {sha.hexdigest()}")
print(f"{'='*60}")

# Final output
print(f"\n=== FINAL REPORT ===")
print(f"Archive: {zip_name}")
print(f"SHA256: {sha.hexdigest()}")
print(f"Total rows: {len(v161_rows)}")
print(f"Producers: {len(producer_counts)}")
print(f"ASD: {len(asd_v162)} total, {len(asd_with_img)} with image, {len(asd_without)} without")
print(f"Fragment images cleared: {fix_count}")
print(f"Expected: 187 with image, 12 without -> {'MATCH' if len(asd_with_img)==187 and len(asd_without)==12 else f'GOT {len(asd_with_img)}/{len(asd_without)}'}")
print(f"Non-ASD unchanged: {'YES' if not any('non-ASD' in str(e) for e in errors) else 'CHECK ERRORS'}")
print(f"Fragment remaining in image_url: {len(fragment_remaining)}")
