"""Stage 3: Final assembly with patches applied"""
import csv, json, os, hashlib
from collections import Counter, defaultdict
from datetime import datetime
from zipfile import ZipFile

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CATALOG_CSV = os.path.join(BASE, "output", "final_hybrid_client_ready_v16_3_asd_full_door_only", "doors_catalog.csv")
PATCH_DIR = os.path.join(BASE, "output", "final_rework_patches")
AUDIT_DIR = os.path.join(BASE, "output", "final_rework_audit")
OUT_DIR = os.path.join(BASE, "output", "final_rework_ready")
os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(os.path.join(OUT_DIR, "check_by_producer"), exist_ok=True)

# Load catalog
with open(CATALOG_CSV, encoding="utf-8-sig") as f:
    catalog = list(csv.DictReader(f, delimiter=";"))

# Load patches
def read_csv_auto(path):
    with open(path, encoding="utf-8-sig") as f:
        sample = f.read(2000)
    delim = ";" if sample.count(";") > sample.count(",") else ","
    with open(path, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f, delimiter=delim))

photo_patches = read_csv_auto(os.path.join(PATCH_DIR, "patch_photo_front_inside.csv"))
model_patches = read_csv_auto(os.path.join(PATCH_DIR, "patch_missing_models.csv"))

# Load audit summary
with open(os.path.join(AUDIT_DIR, "current_vs_expected_summary.json"), encoding="utf-8") as f:
    audit = json.load(f)

# Build photo patch lookup: (producer, model_name) -> patch
photo_lookup = {}
for p in photo_patches:
    key = (p["producer"], p["model_name"])
    photo_lookup[key] = p

# New field names
FIELD_NAMES = list(catalog[0].keys()) + ["image_url_front", "image_url_inside"]

# Apply patches
for r in catalog:
    producer = r["producer"]
    model = r.get("model_name","")
    img = r.get("image_url","") or ""
    
    # Initialize new columns
    r["image_url_front"] = ""
    r["image_url_inside"] = ""
    
    # Check photo patch
    key = (producer, model)
    if key in photo_lookup:
        patch = photo_lookup[key]
        decision = patch.get("decision","")
        
        if decision == "not_found":
            # No image found - clear image_url
            r["image_url"] = ""
            r["status"] = "needs_review"
            r["comment"] = (r.get("comment","") or "") + "; no_valid_full_door_photo"
        elif decision == "manual_check":
            # Keep existing but mark for review
            if img:
                r["image_url_front"] = img
            r["status"] = "needs_review"
            r["comment"] = (r.get("comment","") or "") + "; " + patch.get("reason","photo_issue")
        elif decision == "keep_old" and img:
            r["image_url_front"] = img
    
    # For entry doors: set front from image_url if not already set
    door_type = r.get("door_type","") or ""
    is_entry = "входн" in door_type.lower()
    
    if is_entry:
        if not r["image_url_front"] and img:
            r["image_url_front"] = img
        # image_url_inside stays empty (needs manual check)
        if not r["image_url_inside"]:
            if r["status"] != "needs_review":
                r["status"] = "needs_review"
                r["comment"] = (r.get("comment","") or "") + "; entry_door_inside_photo_missing"
    
    # Compatibility: image_url = image_url_front if front exists
    if r["image_url_front"]:
        r["image_url"] = r["image_url_front"]

# Special rules
# АСД: considered fully covered (199 models)
# Термодур: keep old thermobreak data as-is

# Add needs_review entries for missing lines
missing_line_rows = []
for p in model_patches:
    if p["reason"].startswith("Line"):
        # Skip АСД — reference says "Все модели" x11, catalog has 199 models
        if p["producer"] == "АСД":
            continue
        missing_line_rows.append({
            "door_type": "входная",
            "producer": p["producer"],
            "target_line": p["target_line"],
            "model_name": f"[MISSING] {p['target_line']}",
            "description": "",
            "material": "",
            "coating": p["covering"],
            "image_url": "",
            "source_url": "",
            "status": "needs_review",
            "comment": f"missing_line_from_reference_xlsx; needs_site_parsing",
            "image_url_front": "",
            "image_url_inside": ""
        })

catalog.extend(missing_line_rows)

# Verification
errors = []
if len(catalog) < 3084:
    errors.append(f"Total rows: {len(catalog)} (expected >= 3084, got {len(catalog)})")

producers = [r["producer"] for r in catalog]
producer_counts = Counter(producers)

# Check duplicates
seen = set()
dups = []
for r in catalog:
    key = (r.get("producer",""), r.get("target_line",""), r.get("model_name",""))
    if key in seen:
        dups.append(key)
    seen.add(key)
if dups:
    errors.append(f"Duplicates: {len(dups)}")

# ASD check
asd = [r for r in catalog if r["producer"] == "АСД"]
asd_with = [r for r in asd if r.get("image_url","")]
asd_without = [r for r in asd if not r.get("image_url","")]

# WordPress row count
wp_count = len(catalog)

# Save doors_catalog.csv
with open(os.path.join(OUT_DIR, "doors_catalog.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=FIELD_NAMES, delimiter=";")
    w.writeheader()
    w.writerows(catalog)
print(f"doors_catalog.csv: {len(catalog)} rows")

# Save doors_catalog.xlsx
try:
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Catalog Final"
    for ci, h in enumerate(FIELD_NAMES, 1):
        ws.cell(row=1, column=ci, value=h)
    for ri, r in enumerate(catalog, 2):
        for ci, h in enumerate(FIELD_NAMES, 1):
            ws.cell(row=ri, column=ci, value=r.get(h, ""))
    wb.save(os.path.join(OUT_DIR, "doors_catalog.xlsx"))
    print("doors_catalog.xlsx: saved")
except ImportError:
    pass

# Save wordpress_import.csv
with open(os.path.join(OUT_DIR, "wordpress_import.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=FIELD_NAMES, delimiter=";")
    w.writeheader()
    w.writerows(catalog)
print(f"wordpress_import.csv: {len(catalog)} rows")

# Save check_by_producer
for producer in sorted(producer_counts.keys()):
    producer_rows = [r for r in catalog if r["producer"] == producer]
    safe_name = producer.replace("/", "_").replace("\\", "_")
    
    csv_path = os.path.join(OUT_DIR, "check_by_producer", f"{safe_name}.csv")
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELD_NAMES, delimiter=";")
        w.writeheader()
        w.writerows(producer_rows)
    
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = safe_name[:30]
        for ci, h in enumerate(FIELD_NAMES, 1):
            ws.cell(row=1, column=ci, value=h)
        for ri, r in enumerate(producer_rows, 2):
            for ci, h in enumerate(FIELD_NAMES, 1):
                ws.cell(row=ri, column=ci, value=r.get(h, ""))
        wb.save(os.path.join(OUT_DIR, "check_by_producer", f"{safe_name}.xlsx"))
    except ImportError:
        pass

print(f"check_by_producer/: {len(producer_counts)} producers")

# Run report
needs_review_by_producer = Counter()
for r in catalog:
    if r.get("status","") == "needs_review":
        needs_review_by_producer[r["producer"]] += 1

report = {
    "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "mode": "final_rework_against_manufacturer_lines_xlsx",
    "source_catalog_used": "doors_parser_final_HYBRID_CLIENT_READY_v16_3_ASD_FULL_DOOR_ONLY_2026-06-02_18-38-29.zip",
    "xlsx_reference_used": "Линейки производителей 3.xlsx",
    "total_rows": len(catalog),
    "producers_count": len(producer_counts),
    "rows_by_producer": dict(producer_counts),
    "expected_lines_by_producer": {p: s["expected"] for p, s in audit["by_producer"].items()},
    "missing_lines_count_by_producer": {p: s["missing_line"] for p, s in audit["by_producer"].items()},
    "missing_models_count_by_producer": {p: s["missing_models"] for p, s in audit["by_producer"].items()},
    "front_inside_photo_stats": {
        "total_entry_doors": sum(1 for r in catalog if "входн" in (r.get("door_type","") or "").lower()),
        "with_front": sum(1 for r in catalog if r.get("image_url_front","")),
        "with_inside": sum(1 for r in catalog if r.get("image_url_inside","")),
    },
    "needs_review_by_producer": dict(needs_review_by_producer),
    "duplicates_count": len(dups),
    "thermobreak_policy": "keep_old",
    "errors": errors
}

with open(os.path.join(OUT_DIR, "run_report.json"), "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)
print("run_report.json: saved")

# Copy audit and patch files
import shutil
for fn in os.listdir(AUDIT_DIR):
    src = os.path.join(AUDIT_DIR, fn)
    dst = os.path.join(OUT_DIR, fn)
    if os.path.isfile(src):
        shutil.copy2(src, dst)

for fn in os.listdir(PATCH_DIR):
    src = os.path.join(PATCH_DIR, fn)
    dst = os.path.join(OUT_DIR, fn)
    if os.path.isfile(src):
        shutil.copy2(src, dst)

# Build ZIP
zip_name = f"doors_parser_final_REWORK_READY_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.zip"
zip_path = os.path.join(BASE, "output", zip_name)
with ZipFile(zip_path, "w") as zf:
    # final_rework_ready/
    for root, dirs, files in os.walk(OUT_DIR):
        for fn in files:
            fp = os.path.join(root, fn)
            arcname = os.path.relpath(fp, OUT_DIR)
            zf.write(fp, arcname=os.path.join("output", "final_rework_ready", arcname))
    # final_rework_audit/
    for root, dirs, files in os.walk(AUDIT_DIR):
        for fn in files:
            fp = os.path.join(root, fn)
            arcname = os.path.relpath(fp, AUDIT_DIR)
            zf.write(fp, arcname=os.path.join("output", "final_rework_audit", arcname))
    # final_rework_patches/
    for root, dirs, files in os.walk(PATCH_DIR):
        for fn in files:
            fp = os.path.join(root, fn)
            arcname = os.path.relpath(fp, PATCH_DIR)
            zf.write(fp, arcname=os.path.join("output", "final_rework_patches", arcname))
    # scripts/
    scripts_dir = os.path.join(BASE, "scripts")
    for fn in ["build_catalog_completeness_audit.py", "build_rework_patches.py", "build_final_rework.py"]:
        fp = os.path.join(scripts_dir, fn)
        if os.path.exists(fp):
            zf.write(fp, arcname=os.path.join("scripts", fn))
    # config/
    config_dir = os.path.join(BASE, "config")
    if os.path.exists(os.path.join(config_dir, "sites.yaml")):
        zf.write(os.path.join(config_dir, "sites.yaml"), arcname="config/sites.yaml")
    # data/
    data_dir = os.path.join(BASE, "data")
    ref_xlsx = os.path.join(data_dir, "reference_lines.xlsx")
    if os.path.exists(ref_xlsx):
        zf.write(ref_xlsx, arcname="data/reference_lines.xlsx")

sha = hashlib.sha256()
with open(zip_path, "rb") as f:
    for chunk in iter(lambda: f.read(65536), b""):
        sha.update(chunk)

print(f"\n{'='*60}")
print(f"STAGE 3 FINAL ASSEMBLY COMPLETE")
print(f"{'='*60}")
print(f"Archive: {zip_path}")
print(f"SHA256: {sha.hexdigest()}")
print(f"\nTotal rows: {len(catalog)}")
print(f"Producers: {len(producer_counts)}")
print(f"ASD: {len(asd)} total, {len(asd_with)} with image, {len(asd_without)} without")
print(f"Duplicates: {len(dups)}")
print(f"WordPress rows: {wp_count}")
print(f"Needs review by producer:")
for p, c in sorted(needs_review_by_producer.items()):
    print(f"  {p}: {c}")
print(f"\nExpected lines covered: {audit['total_in_catalog']}/{audit['total_expected_lines']}")
print(f"Missing lines added as needs_review: {len(missing_line_rows)}")
print(f"Errors: {errors if errors else 'None'}")
