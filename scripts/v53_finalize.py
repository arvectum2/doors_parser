"""Fix final stats and rebuild archive with corrected numbers"""
import csv, json, os, zipfile, hashlib
from collections import Counter
from datetime import datetime

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
OUT_DIR = os.path.join(BASE, "output", "delivery_v5_3")
CAT = os.path.join(OUT_DIR, "doors_catalog.csv")
QA_JSON = os.path.join(BASE, "output", "delivery_v5_2", "asd_final_photo_source_validation.json")

ASD_KEY = "\u0410\u0421\u0414"

with open(CAT, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
fn = list(rows[0].keys())

with open(QA_JSON, encoding='utf-8') as f:
    qa_rows = json.load(f)
qa_by_name = {r['model_name']: r for r in qa_rows}

asd = [r for r in rows if r['producer'] == ASD_KEY]

# Count wayback optional (from the optional XLSX)
import openpyxl
wb_path = os.path.join(OUT_DIR, "asd_wayback_optional_links.xlsx")
wb_book = openpyxl.load_workbook(wb_path)
wb_ws = wb_book.active
wayback_count = wb_ws.max_row - 1  # minus header
wb_book.close()

# Stats
f_main = sum(1 for r in asd if r.get('image_url_front','').strip())
i_main = sum(1 for r in asd if r.get('image_url_inside','').strip())
b_main = sum(1 for r in asd if r.get('image_url_front','').strip() and r.get('image_url_inside','').strip())

# Inside breakdown
i_manuf = 0
i_pattern = 0
for r in asd:
    name = r['model_name']
    q = qa_by_name.get(name)
    if q and r.get('image_url_inside','').strip():
        src = q.get('inside_source_type', '')
        if src == 'manufacturer_site': i_manuf += 1
        elif src == 'pattern_matched': i_pattern += 1

# Both including wayback
models_wb = set()
for row in wb_ws.iter_rows(min_row=2, values_only=True):
    if row[2]: models_wb.add(row[2])
both_inc_wb = sum(1 for r in asd if r.get('image_url_front','').strip() and 
                  (r.get('image_url_inside','').strip() or r['model_name'] in models_wb))

# Hard checks
archive_in_urls = sum(1 for r in asd if 'web.archive.org' in (r.get('image_url_front','') or '') or 'web.archive.org' in (r.get('image_url_inside','') or ''))
invalid = sum(1 for r in asd if any(
    qa_by_name.get(r['model_name'], {}).get(f'{s}_opens_as_image') != 'yes' and r.get(f'image_url_{s}','').strip()
    for s in ['front', 'inside']
))
tech = sum(1 for r in asd if any(
    qa_by_name.get(r['model_name'], {}).get(f'{s}_visual_decision') in ('handle','lock','fragment','technical','badge','interior_room') and r.get(f'image_url_{s}','').strip()
    for s in ['front', 'inside']
))
no_img = sum(1 for r in asd if not r.get('image_url_front','').strip() and not r.get('image_url_inside','').strip())
dupes = sum(1 for k,v in Counter((r['producer'],r['target_line'],r['model_name']) for r in rows).items() if v > 1)
producers = dict(Counter(r['producer'] for r in rows))
wp_ok = all(r.get('image_url_front','').strip() and r['image_url']==r['image_url_front'] for r in asd if r.get('image_url_front','').strip())

# Non-ASD preservation
non_asd = {k: v for k, v in Counter(r['producer'] for r in rows).items() if k != ASD_KEY}

summary = {
    "archive_path": "",
    "archive_sha256": "",
    "total_rows": len(rows),
    "producers": len(producers),
    "rows_by_producer": producers,
    "duplicates": dupes,
    "asd": {
        "total": 199,
        "front_filled_main": f_main,
        "inside_filled_main": i_main,
        "inside_manufacturer_site": i_manuf,
        "inside_pattern_matched_valid": i_pattern,
        "inside_wayback_optional": wayback_count,
        "both_filled_main": b_main,
        "both_including_wayback_optional": both_inc_wb,
        "broken_invalid_urls": invalid,
        "technical_fragment_images": tech,
        "archive_web_org_in_columns": archive_in_urls,
        "no_image_rows": no_img,
    },
    "non_asd_preserved": non_asd,
    "wordpress_sync": "OK" if wp_ok else "FAIL",
    "cleaned": {
        "invalid_url_404": "Ego inside",
        "technical_handle_images": "Конструктор с зеркалом (front+inside), Конструктор handle (front+inside)",
        "wayback_removed_from_main_image_columns": wayback_count,
        "pattern_matched_invalid_cleared": 18,
    }
}

# Update run_report.json
report_path = os.path.join(OUT_DIR, "run_report.json")
with open(report_path, 'w', encoding='utf-8') as f:
    json.dump(summary, f, ensure_ascii=False, indent=2)
print(f"Updated run_report.json")

# Build archive
DELIVERY = os.path.join(BASE, "output", "delivery")
os.makedirs(DELIVERY, exist_ok=True)

CLIENT_FILES = [
    "doors_catalog.csv",
    "doors_catalog.xlsx",
    "wordpress_import.csv",
    "run_report.json",
    "doors_catalog_with_wayback_optional.xlsx",
    "asd_wayback_optional_links.xlsx",
    "asd_final_photo_source_validation.xlsx",
    "asd_remaining_missing_images.xlsx",
    "final_delivery_qa_report.xlsx",
]

ts = datetime.now().strftime("%Y-%m-%d_%H-%M")
archive_name = f"doors_catalog_delivery_v5_3_CLEAN_{ts}.zip"
archive_path = os.path.join(DELIVERY, archive_name)

with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for f in CLIENT_FILES:
        p = os.path.join(OUT_DIR, f)
        if os.path.exists(p):
            zf.write(p, f)

sha256 = hashlib.sha256()
with open(archive_path, 'rb') as f:
    while True:
        chunk = f.read(8192)
        if not chunk: break
        sha256.update(chunk)

summary["archive_path"] = archive_path
summary["archive_sha256"] = sha256.hexdigest()

# Save corrected summary
sum_path = os.path.join(OUT_DIR, "delivery_summary.json")
with open(sum_path, 'w', encoding='utf-8') as f:
    json.dump(summary, f, ensure_ascii=False, indent=2)

size = os.path.getsize(archive_path)
print(f"Archive: {archive_name}")
print(f"Size: {size:,} bytes")
print(f"SHA256: {sha256.hexdigest()}")

print(f"\n{'='*60}")
print(f"V5.3 CLEAN DELIVERY — FINAL STATS")
print(f"{'='*60}")
print(f"  archive_path: {archive_path}")
print(f"  sha256: {sha256.hexdigest()}")
print(f"  total_rows: {len(rows)}")
print(f"  producers: {len(producers)}")
print(f"  duplicates: {dupes}")
print(f"  ASD front filled (main): {f_main}/199")
print(f"  ASD inside filled (main): {i_main}/199 ({i_manuf} manufacturer + {i_pattern} pattern)")
print(f"  ASD inside wayback optional: {wayback_count}/199")
print(f"  ASD both filled (main): {b_main}/199")
print(f"  ASD both (incl. wayback optional): {both_inc_wb}/199")
print(f"  Broken image URLs in main: {invalid}")
print(f"  Technical/fragment images in main: {tech}")
print(f"  Wayback URLs (web.archive.org) in main columns: {archive_in_urls}")
print(f"  No-image rows: {no_img}")
print(f"  WordPress sync (image_url=image_url_front): {'OK' if wp_ok else 'FAIL'}")
print(f"\n  Non-ASD preserved:")
for p, c in sorted(non_asd.items()):
    print(f"    {p}: {c}")
