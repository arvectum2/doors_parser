"""Build v16.3 — clear remaining fragment images"""
import csv, os, json, hashlib
from collections import Counter
from datetime import datetime
from zipfile import ZipFile

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
V16_DIR = os.path.join(BASE, "output", "final_hybrid_client_ready_v16_asd_full_door_only")
V16_CSV = os.path.join(V16_DIR, "doors_catalog.csv")
V16_XLSX = os.path.join(V16_DIR, "doors_catalog.xlsx")
V16_WP = os.path.join(V16_DIR, "wordpress_import.csv")
V16_REPORT = os.path.join(V16_DIR, "run_report.json")
V16_CHECK_DIR = os.path.join(V16_DIR, "check_by_producer")
V16_CLEAN_DIR = os.path.join(V16_DIR, "output", "asd_photo_patch_v2_clean")
V16_NOT_FOUND = os.path.join(V16_DIR, "not_found.csv")

DEBUG_CSV = os.path.join(BASE, "output", "asd_photo_patch_v2", "asd_photo_candidates_debug.csv")

with open(V16_CSV, encoding="utf-8-sig") as f:
    rows = list(csv.DictReader(f, delimiter=";"))

with open(DEBUG_CSV, encoding="utf-8-sig") as f:
    debugs = list(csv.DictReader(f, delimiter=";"))

# Build debug lookup: for each URL that exists in debug, get its visual_decision
# Key by (model_name, source_url) → visual_decision for the URL matching the catalog
# But we need to check the catalog image_url against ALL candidate URLs in debug
debug_urls_by_key = {}  # (model_name, source_url) -> list of (candidate_url, visual_decision, ar)
for d in debugs:
    key = (d["model_name"], d["source_url"])
    if key not in debug_urls_by_key:
        debug_urls_by_key[key] = []
    debug_urls_by_key[key].append({
        "url": d.get("candidate_url",""),
        "vd": d.get("visual_decision",""),
        "ar": d.get("aspect_ratio",""),
        "w": d.get("width",""),
        "h": d.get("height",""),
    })

# Find ASD rows where catalog image_url matches a debug entry with fragment vd
fix_count = 0
fragment_details = []

for r in rows:
    if r["producer"] != "АСД":
        continue
    catalog_url = r.get("image_url","") or ""
    if not catalog_url:
        continue
    key = (r["model_name"], r["source_url"])
    if key not in debug_urls_by_key:
        continue
    for entry in debug_urls_by_key[key]:
        if entry["url"] == catalog_url and entry["vd"] not in ("full_door", "full_door_short"):
            # Fragment found in catalog
            print(f"  FIX: {r['model_name'][:35]:35} vd={entry['vd']:15} ar={entry['ar']:6} url={catalog_url[:50]}")
            r["image_url"] = ""
            r["status"] = "needs_review"
            old_comment = r.get("comment","") or ""
            add = "old_fragment_image_cleared_no_full_door_candidate"
            r["comment"] = f"{add}; {old_comment}" if old_comment else add
            fix_count += 1
            fragment_details.append({
                "model": r["model_name"],
                "old_url": catalog_url,
                "visual_decision": entry["vd"],
                "ar": entry["ar"],
                "dimensions": f"{entry['w']}x{entry['h']}"
            })
            break

# Verifications
FIELD_NAMES = list(rows[0].keys()) if rows else []
errors = []
if len(rows) != 3084:
    errors.append(f"Total rows: {len(rows)} (expected 3084)")
producers = [r["producer"] for r in rows]
producer_counts = Counter(producers)
if len(producer_counts) != 10:
    errors.append(f"Producer count: {len(producer_counts)} (expected 10)")
seen, dups = set(), []
for r in rows:
    key = (r.get("producer",""), r.get("target_line",""), r.get("model_name",""))
    if key in seen:
        dups.append(key)
    seen.add(key)
if dups:
    errors.append(f"Duplicates: {len(dups)}")

v14_csv = os.path.join(BASE, "output", "final_hybrid_client_ready_v14", "doors_catalog.csv")
with open(v14_csv, encoding="utf-8-sig") as f:
    v14 = list(csv.DictReader(f, delimiter=";"))

for p, c in Counter(r["producer"] for r in v14).items():
    if p != "АСД":
        v163_c = producer_counts.get(p, 0)
        if v163_c != c:
            errors.append(f"Producer {p}: v14={c} v163={v163_c}")
for i, (v14_r, v163_r) in enumerate(zip(v14, rows)):
    if v14_r["producer"] != "АСД":
        for col in ["door_type","producer","target_line","model_name","image_url","source_url"]:
            if v14_r.get(col,"") != v163_r.get(col,""):
                errors.append(f"Row {i+2} non-ASD changed: {col}")
                break

# Final fragment check
remaining_fragment = []
for r in rows:
    if r["producer"] != "АСД":
        continue
    url = r.get("image_url","") or ""
    if not url:
        continue
    key = (r["model_name"], r["source_url"])
    if key not in debug_urls_by_key:
        continue
    for entry in debug_urls_by_key[key]:
        if entry["url"] == url and entry["vd"] not in ("full_door", "full_door_short"):
            remaining_fragment.append((r, entry))

if remaining_fragment:
    errors.append(f"Fragment remaining in ASD: {len(remaining_fragment)}")
    for r, e in remaining_fragment:
        errors.append(f"  {r['model_name'][:30]} vd={e['vd']} url={r.get('image_url','')[:50]}")

# ASD counts
asd = [r for r in rows if r["producer"] == "АСД"]
asd_with = [r for r in asd if r.get("image_url","")]
asd_without = [r for r in asd if not r.get("image_url","")]

print(f"\nFixed: {fix_count}")
print(f"ASD: {len(asd)} total, {len(asd_with)} with image, {len(asd_without)} without")

if errors:
    print(f"\n=== ERRORS ({len(errors)}) ===")
    for e in errors:
        print(f"  FAIL: {e}")
else:
    print("\n=== ALL VERIFICATIONS PASSED ===")

# Save
with open(V16_CSV, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=FIELD_NAMES, delimiter=";")
    w.writeheader()
    w.writerows(rows)
print(f"doors_catalog.csv: {len(rows)} rows")

try:
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Catalog v16.3"
    for ci, h in enumerate(FIELD_NAMES, 1):
        ws.cell(row=1, column=ci, value=h)
    for ri, r in enumerate(rows, 2):
        for ci, h in enumerate(FIELD_NAMES, 1):
            ws.cell(row=ri, column=ci, value=r.get(h, ""))
    wb.save(V16_XLSX)
    print("doors_catalog.xlsx: saved")
except ImportError:
    pass

# WordPress
with open(V16_WP, encoding="utf-8-sig") as f:
    reader = csv.DictReader(f, delimiter=";")
    wp_fields = reader.fieldnames
    wp_rows = list(reader)
asd_lookup = {(r["model_name"], r["source_url"]): r for r in rows if r["producer"] == "АСД"}
for wp_r in wp_rows:
    key = (wp_r.get("model_name",""), wp_r.get("source_url",""))
    if key in asd_lookup:
        wp_r["image_url"] = asd_lookup[key].get("image_url","")
with open(V16_WP, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=wp_fields, delimiter=";")
    w.writeheader()
    w.writerows(wp_rows)
print(f"wordpress_import.csv: {len(wp_rows)} rows")

# Report
report = {
    "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "version": "v16.3_ASD_FULL_DOOR_ONLY",
    "base": "v16.2",
    "fixes": {
        "fragment_images_cleared_this_run": fix_count,
        "total_fragment_images_cleared_across_versions": 5 + 2 + fix_count,
        "cleared_rows": fragment_details
    },
    "total_rows": len(rows),
    "producers_count": len(producer_counts),
    "producers": dict(producer_counts),
    "asd": {
        "total": len(asd),
        "with_image_url": len(asd_with),
        "without_image_url": len(asd_without),
    },
    "errors": errors
}
with open(V16_REPORT, "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)
print("run_report.json: saved")

# check_by_producer/АСД.csv
asd_check_csv = os.path.join(V16_CHECK_DIR, "АСД.csv")
if os.path.exists(asd_check_csv):
    with open(asd_check_csv, encoding="utf-8-sig") as f:
        check_rows = list(csv.DictReader(f, delimiter=";"))
    check_fields = list(check_rows[0].keys()) if check_rows else []
    cat_lookup = {(r["model_name"], r["source_url"]): r for r in rows if r["producer"] == "АСД"}
    for cr in check_rows:
        key = (cr.get("model_name",""), cr.get("source_url",""))
        if key in cat_lookup:
            cr["image_url"] = cat_lookup[key].get("image_url","")
            cr["status"] = cat_lookup[key].get("status","")
            cr["comment"] = cat_lookup[key].get("comment","")
    with open(asd_check_csv, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=check_fields, delimiter=";")
        w.writeheader()
        w.writerows(check_rows)
    print("check_by_producer/АСД.csv: updated")

    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "АСД v16.3"
        for ci, h in enumerate(check_fields, 1):
            ws.cell(row=1, column=ci, value=h)
        for ri, cr in enumerate(check_rows, 2):
            for ci, h in enumerate(check_fields, 1):
                ws.cell(row=ri, column=ci, value=cr.get(h, ""))
        wb.save(os.path.join(V16_CHECK_DIR, "АСД.xlsx"))
        print("check_by_producer/АСД.xlsx: updated")
    except ImportError:
        pass

# ZIP
zip_name = f"doors_parser_final_HYBRID_CLIENT_READY_v16_3_ASD_FULL_DOOR_ONLY_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.zip"
zip_path = os.path.join(BASE, "output", zip_name)
with ZipFile(zip_path, "w") as zf:
    for root, dirs, files in os.walk(V16_DIR):
        for fn in files:
            fp = os.path.join(root, fn)
            arcname = os.path.relpath(fp, V16_DIR)
            zf.write(fp, arcname=os.path.join("output", "final_hybrid_client_ready_v16_asd_full_door_only", arcname))
    v16_clean = os.path.join(V16_DIR, "output", "asd_photo_patch_v2_clean")
    if os.path.exists(v16_clean):
        for fn in os.listdir(v16_clean):
            zf.write(os.path.join(v16_clean, fn), arcname=os.path.join("output", "asd_photo_patch_v2_clean", fn))

sha = hashlib.sha256()
with open(zip_path, "rb") as f:
    for chunk in iter(lambda: f.read(65536), b""):
        sha.update(chunk)

print(f"\n{'='*60}")
print(f"ZIP: {zip_path}")
print(f"SHA256: {sha.hexdigest()}")
print(f"{'='*60}")

print(f"\n=== FINAL REPORT ===")
print(f"Archive: {os.path.basename(zip_path)}")
print(f"SHA256: {sha.hexdigest()}")
print(f"Total rows: {len(rows)}")
print(f"Producers: {len(producer_counts)}")
print(f"ASD: {len(asd)} total, {len(asd_with)} with image, {len(asd_without)} without")
print(f"Expected: 185 with, 14 without -> {'MATCH' if len(asd_with)==185 and len(asd_without)==14 else f'GOT {len(asd_with)}/{len(asd_without)}'}")
print(f"Fragments cleared this run: {fix_count}")
print(f"Fragment remaining in image_url: {len(remaining_fragment)}")
print(f"Non-ASD unchanged: {'YES' if not any('non-ASD' in str(e) for e in errors) else 'CHECK'}")
print(f"WordPress rows = catalog rows: {'YES' if len(wp_rows)==len(rows) else f'NO ({len(wp_rows)} vs {len(rows)})'}")
