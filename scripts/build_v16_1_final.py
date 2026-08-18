"""Build v16.1 with complete structure, tracking, and verification"""
import csv, os, json, hashlib, shutil, re, io, requests
from collections import Counter, defaultdict
from datetime import datetime
from zipfile import ZipFile
from PIL import Image

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"

V14_DIR = os.path.join(BASE, "output", "final_hybrid_client_ready_v14")
V14_CSV = os.path.join(V14_DIR, "doors_catalog.csv")
V14_WP = os.path.join(V14_DIR, "wordpress_import.csv")
V14_REPORT = os.path.join(V14_DIR, "run_report.json")
V14_CHECK_DIR = os.path.join(V14_DIR, "check_by_producer")
V14_NOT_FOUND = os.path.join(V14_DIR, "not_found.csv")

CLEAN_DIR = os.path.join(BASE, "output", "asd_photo_patch_v2_clean")
CLEAN_CSV = os.path.join(CLEAN_DIR, "asd_photo_patch_v2_clean.csv")
CLEAN_XLSX = os.path.join(CLEAN_DIR, "asd_photo_patch_v2_clean.xlsx")
CLEAN_HTML = os.path.join(CLEAN_DIR, "asd_photo_patch_v2_clean_review.html")
CLEAN_REPORT = os.path.join(CLEAN_DIR, "asd_photo_patch_v2_clean_report.json")
CLEAN_DEBUG = os.path.join(BASE, "output", "asd_photo_patch_v2", "asd_photo_candidates_debug.csv")

BAD_PATTERNS = ["handle", "lock", "zamok", "ruchka", "ручка"]

V16_BASE_DIR = os.path.join(BASE, "output", "final_hybrid_client_ready_v16_asd_full_door_only")
os.makedirs(V16_BASE_DIR, exist_ok=True)

print("Loading data...")
with open(V14_CSV, encoding="utf-8-sig") as f:
    v14_rows = list(csv.DictReader(f, delimiter=";"))
with open(CLEAN_CSV, encoding="utf-8-sig") as f:
    clean_rows = list(csv.DictReader(f, delimiter=";"))
with open(CLEAN_DEBUG, encoding="utf-8-sig") as f:
    debugs = list(csv.DictReader(f, delimiter=";"))

clean_lookup = {}
for r in clean_rows:
    clean_lookup[(r["model_name"], r["source_url"])] = r

selected_by_key = {}
for d in debugs:
    if d.get("is_selected") == "YES":
        selected_by_key[(d["model_name"], d["source_url"])] = d

def check_image_dimensions(url):
    try:
        r = requests.get(url, timeout=10, headers={"User-Agent": "Mozilla/5.0"}, stream=True)
        if r.status_code != 200:
            return None, None
        content = io.BytesIO()
        for chunk in r.iter_content(8192):
            content.write(chunk)
            if content.tell() > 65536:
                break
        r.close()
        content.seek(0)
        img = Image.open(content)
        return img.size[0], img.size[1]
    except:
        return None, None

print("Building v16.1 rows...")
v16_rows = []
asd_image_sources = {}
applied_replace = 0
applied_keep = 0
applied_skip = 0
ruchka_verified = 0
ruchka_cleared = 0
old_retained_not_in_patch = 0

for r in v14_rows:
    row = dict(r)
    producer = row.get("producer","")
    
    if producer == "АСД":
        key = (row["model_name"], row["source_url"])
        patch_r = clean_lookup.get(key)
        v14_url = row.get("image_url","") or ""
        source = ""
        
        if patch_r:
            dec = patch_r["decision"]
            if dec == "replace":
                new_url = patch_r.get("selected_image_url","") or ""
                if new_url:
                    row["image_url"] = new_url
                    source = "replace_from_clean_patch"
                    applied_replace += 1
                else:
                    source = "replace_empty_selected_url"
                    applied_skip += 1
            elif dec == "keep":
                source = "keep_old_verified"
                applied_keep += 1
            elif dec == "manual_check":
                source = "manual_check_not_applied"
                applied_skip += 1
            elif dec == "not_found":
                if v14_url and any(b in v14_url.lower() for b in BAD_PATTERNS):
                    w, h = check_image_dimensions(v14_url)
                    if w and h:
                        ratio = h / w if w > 0 else 0
                        if ratio >= 1.5:
                            source = "false_positive_full_door_verified"
                            ruchka_verified += 1
                        else:
                            row["image_url"] = ""
                            source = "not_found_ruchka_fragment_cleared"
                            ruchka_cleared += 1
                    else:
                        row["image_url"] = ""
                        source = "not_found_ruchka_unverifiable_cleared"
                        ruchka_cleared += 1
                elif v14_url:
                    source = "old_image_retained_because_not_found"
                else:
                    source = "not_found_no_image"
                applied_skip += 1
        else:
            source = "old_image_retained_because_not_in_patch"
            old_retained_not_in_patch += 1
        
        asd_image_sources[key] = source
    # Non-ASD stays unchanged
    
    v16_rows.append(row)

# Stats
asd_v16 = [r for r in v16_rows if r["producer"] == "АСД"]
asd_with_img = [r for r in asd_v16 if r.get("image_url","")]
asd_without = [r for r in asd_v16 if not r.get("image_url","")]
source_counts = Counter(asd_image_sources.values())

print(f"\n=== v16.1 Build ===")
print(f"Total rows: {len(v16_rows)}")
print(f"ASD: {len(asd_v16)} total, {len(asd_with_img)} with image, {len(asd_without)} without")
print(f"Applied replace: {applied_replace}, keep: {applied_keep}, skip: {applied_skip}")
print(f"Ruchka verified full-door: {ruchka_verified}, cleared: {ruchka_cleared}")
print(f"\nSources: {dict(source_counts)}")

# Verifications
errors = []
if len(v16_rows) != 3084:
    errors.append(f"Total rows: {len(v16_rows)} (expected 3084)")
producers = [r["producer"] for r in v16_rows]
producer_counts = Counter(producers)
if len(producer_counts) != 10:
    errors.append(f"Producer count: {len(producer_counts)} (expected 10)")
seen, dups = set(), []
for r in v16_rows:
    key = (r.get("producer",""), r.get("target_line",""), r.get("model_name",""))
    if key in seen:
        dups.append(key)
    seen.add(key)
if dups:
    errors.append(f"Duplicates: {len(dups)}")
v14_prod_counts = Counter(r["producer"] for r in v14_rows)
for p, c in v14_prod_counts.items():
    if p != "АСД":
        v16_c = producer_counts.get(p, 0)
        if v16_c != c:
            errors.append(f"Producer {p}: v14={c} v16={v16_c}")
for i, (v14_r, v16_r) in enumerate(zip(v14_rows, v16_rows)):
    if v14_r["producer"] != "АСД":
        for col in ["door_type","producer","target_line","model_name","image_url","source_url"]:
            if v14_r.get(col,"") != v16_r.get(col,""):
                errors.append(f"Row {i+2} non-ASD changed: {col}")
                break
for r in asd_v16:
    key = (r["model_name"], r["source_url"])
    src = asd_image_sources.get(key, "")
    if src == "replace_from_clean_patch":
        url = r.get("image_url","") or ""
        if any(b in url.lower() for b in BAD_PATTERNS):
            errors.append(f"Replace bad pattern: {r['model_name'][:25]}")
applied_keys = set(k for k, v in asd_image_sources.items() 
                   if v in ("replace_from_clean_patch", "keep_old_verified", "false_positive_full_door_verified"))
for d in debugs:
    key = (d["model_name"], d["source_url"])
    if d.get("is_selected") == "YES" and key in applied_keys:
        vd = d.get("visual_decision","")
        if vd not in ("full_door", "full_door_short"):
            errors.append(f"Applied fragment vd: {d['model_name'][:25]} vd={vd}")

if errors:
    print(f"\n=== ERRORS ({len(errors)}) ===")
    for e in errors:
        print(f"  FAIL: {e}")
else:
    print("\n=== ALL VERIFICATIONS PASSED ===")

# Save catalog CSV
FIELD_NAMES = ["door_type","producer","target_line","model_name","description","material","coating","image_url","source_url","status","comment"]
v16_csv = os.path.join(V16_BASE_DIR, "doors_catalog.csv")
with open(v16_csv, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=FIELD_NAMES, delimiter=";")
    w.writeheader()
    w.writerows(v16_rows)

# XLSX
try:
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Catalog v16.1"
    for ci, h in enumerate(FIELD_NAMES, 1):
        ws.cell(row=1, column=ci, value=h)
    for ri, r in enumerate(v16_rows, 2):
        for ci, h in enumerate(FIELD_NAMES, 1):
            ws.cell(row=ri, column=ci, value=r.get(h, ""))
    wb.save(os.path.join(V16_BASE_DIR, "doors_catalog.xlsx"))
    print("XLSX: saved")
except ImportError:
    print("XLSX: skipped (no openpyxl)")

# WordPress import
if os.path.exists(V14_WP):
    with open(V14_WP, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f, delimiter=";")
        wp_fields = reader.fieldnames
        wp_rows = list(reader)
    asd_lookup = {(r["model_name"], r["source_url"]): r for r in v16_rows if r["producer"] == "АСД"}
    for wp_r in wp_rows:
        key = (wp_r.get("model_name",""), wp_r.get("source_url",""))
        if key in asd_lookup:
            wp_r["image_url"] = asd_lookup[key].get("image_url","")
    with open(os.path.join(V16_BASE_DIR, "wordpress_import.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=wp_fields, delimiter=";")
        w.writeheader()
        w.writerows(wp_rows)
    wp_count = len(wp_rows)
    print(f"WordPress: {wp_count} rows")
    if wp_count != len(v16_rows):
        errors.append(f"WP rows {wp_count} != catalog {len(v16_rows)}")
else:
    wp_count = len(v16_rows)
    print("WordPress: v14 source not found, copying catalog")
    shutil.copy2(v16_csv, os.path.join(V16_BASE_DIR, "wordpress_import.csv"))

# Run report
decisions_nf = sum(1 for s in asd_image_sources.values() if "not_found" in s and "ruchka" not in s)
decisions_manual = sum(1 for s in asd_image_sources.values() if s == "manual_check_not_applied")

report_data = {
    "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "version": "v16.1_ASD_FULL_DOOR_ONLY",
    "base": "v14_hybrid_client_ready",
    "total_rows": len(v16_rows),
    "producers_count": len(producer_counts),
    "producers": dict(producer_counts),
    "non_asd_unchanged": True if not errors or all("non-ASD" not in e for e in errors) else False,
    "asd": {
        "total": len(asd_v16),
        "with_image_url": len(asd_with_img),
        "without_image_url": len(asd_without),
        "applied_replace": applied_replace,
        "applied_keep": applied_keep,
        "not_applied_manual_check": decisions_manual,
        "not_applied_not_found": decisions_nf,
        "ruchka_verified_full_door": ruchka_verified,
        "ruchka_fragment_cleared": ruchka_cleared,
        "image_sources": dict(source_counts),
        "old_images_retained_by_reason": {
            "old_image_retained_because_not_found": source_counts.get("old_image_retained_because_not_found", 0),
            "manual_check_not_applied": source_counts.get("manual_check_not_applied", 0),
        }
    },
    "errors": errors
}
with open(os.path.join(V16_BASE_DIR, "run_report.json"), "w", encoding="utf-8") as f:
    json.dump(report_data, f, ensure_ascii=False, indent=2)
print("Report: saved")

# check_by_producer
if os.path.exists(V14_CHECK_DIR):
    v16_check_dir = os.path.join(V16_BASE_DIR, "check_by_producer")
    shutil.copytree(V14_CHECK_DIR, v16_check_dir, dirs_exist_ok=True)
    # Update ASD check CSV with image_source
    asd_check_csv = os.path.join(v16_check_dir, "АСД.csv")
    if os.path.exists(asd_check_csv):
        with open(asd_check_csv, encoding="utf-8-sig") as f:
            check_rows = list(csv.DictReader(f, delimiter=";"))
        check_fields = csv.DictReader(open(asd_check_csv, encoding="utf-8-sig"), delimiter=";").fieldnames
        asd_lookup_new = {(r["model_name"], r["source_url"]): r for r in v16_rows if r["producer"] == "АСД"}
        for cr in check_rows:
            key = (cr.get("model_name",""), cr.get("source_url",""))
            if key in asd_lookup_new:
                cr["image_url"] = asd_lookup_new[key].get("image_url","")
                cr["image_source"] = asd_image_sources.get(key, "unknown")
        new_fields = list(check_fields)
        if "image_source" not in new_fields:
            new_fields.append("image_source")
        with open(asd_check_csv, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=new_fields, delimiter=";")
            w.writeheader()
            w.writerows(check_rows)
        print("Check_by_producer: ASD updated")
    else:
        print("Check_by_producer: ASD CSV not found")
else:
    print("Check_by_producer: v14 source not found")

# Copy V14 support files
for fn in ["not_found.csv"]:
    src = os.path.join(V14_DIR, fn)
    if os.path.exists(src):
        shutil.copy2(src, os.path.join(V16_BASE_DIR, fn))

# Copy clean patch files
v16_clean_dir = os.path.join(V16_BASE_DIR, "output", "asd_photo_patch_v2_clean")
os.makedirs(v16_clean_dir, exist_ok=True)
for src_path, dst_name in [
    (CLEAN_CSV, "asd_photo_patch_v2_clean.csv"),
    (CLEAN_XLSX, "asd_photo_patch_v2_clean.xlsx"),
    (CLEAN_HTML, "asd_photo_patch_v2_clean_review.html"),
    (CLEAN_REPORT, "asd_photo_patch_v2_clean_report.json"),
    (CLEAN_DEBUG, "asd_photo_candidates_debug.csv"),
]:
    if os.path.exists(src_path):
        shutil.copy2(src_path, os.path.join(v16_clean_dir, dst_name))
        print(f"Clean patch: {dst_name} copied")
    else:
        print(f"Clean patch: {dst_name} NOT FOUND at {src_path}")

# Create ZIP
zip_name = f"doors_parser_final_HYBRID_CLIENT_READY_v16_1_ASD_FULL_DOOR_ONLY_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.zip"
zip_path = os.path.join(BASE, "output", zip_name)

with ZipFile(zip_path, "w") as zf:
    for root, dirs, files in os.walk(V16_BASE_DIR):
        for fn in files:
            fp = os.path.join(root, fn)
            arcname = os.path.relpath(fp, V16_BASE_DIR)
            zf.write(fp, arcname=os.path.join("output", "final_hybrid_client_ready_v16_asd_full_door_only", arcname))

sha = hashlib.sha256()
with open(zip_path, "rb") as f:
    for chunk in iter(lambda: f.read(65536), b""):
        sha.update(chunk)

print(f"\n{'='*60}")
print(f"ZIP: {zip_path}")
print(f"SHA256: {sha.hexdigest()}")
print(f"{'='*60}")

# Final answer
print(f"\n=== FINAL REPORT ===")
print(f"Archive: {zip_name}")
print(f"SHA256: {sha.hexdigest()}")
print(f"Total rows: {len(v16_rows)}")
print(f"Producers: {len(producer_counts)}")
print(f"ASD: {len(asd_v16)} total, {len(asd_with_img)} with image, {len(asd_without)} without")
print(f"Applied replace: {applied_replace} | keep: {applied_keep} | skip (nf/mc): {applied_skip}")
print(f"Ruchka verified full-door (kept): {ruchka_verified} | ruchka fragment (cleared): {ruchka_cleared}")
print(f"\nOld images retained reasons:")
for k in ["old_image_retained_because_not_found", "manual_check_not_applied", "not_found_no_image"]:
    v = source_counts.get(k, 0)
    if v:
        print(f"  {k}: {v}")
print(f"\nNon-ASD unchanged: {'YES' if not any('non-ASD' in str(e) for e in errors) else 'CHECK ERRORS'}")
