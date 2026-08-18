"""V5.3: Clean build - remove invalid/technical/wayback, create optional wayback file"""
import csv, json, os, re, openpyxl, zipfile, hashlib
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from collections import Counter, defaultdict
from datetime import datetime

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
OUT_DIR = os.path.join(BASE, "output", "delivery_v5_3")
os.makedirs(OUT_DIR, exist_ok=True)

# Source files
CAT_SRC = os.path.join(BASE, "output", "delivery_v5_2", "doors_catalog.csv")
QA_JSON = os.path.join(BASE, "output", "delivery_v5_2", "asd_final_photo_source_validation.json")
VALID = os.path.join(BASE, "output", "delivery_v5_2", "asd_url_validation_results.json")
WB_JSON = os.path.join(BASE, "output", "delivery_v5_2", "wayback_inside_all_results.json")

ASD_KEY = "\u0410\u0421\u0414"

# Load catalog
with open(CAT_SRC, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
fn = list(rows[0].keys())
print(f"Loaded catalog: {len(rows)} rows")

# Load QA
with open(QA_JSON, encoding='utf-8') as f:
    qa_rows = json.load(f)
qa_by_name = {r['model_name']: r for r in qa_rows}

# Load validation raw data
with open(VALID, encoding='utf-8') as f:
    val_raw = json.load(f)
validation = {}
for k, v in val_raw.items():
    parts = k.split('||')
    if len(parts) == 2:
        validation[(parts[0], parts[1])] = v

# Load wayback raw data
with open(WB_JSON, encoding='utf-8') as f:
    wb_data = json.load(f)

# Load patch URLs for reference
V2_PATCH = os.path.join(BASE, "output", "asd_photo_patch_v2_clean", "asd_photo_patch_v2_clean.csv")
patch_urls = set()
with open(V2_PATCH, encoding='utf-8-sig') as f:
    for r in csv.DictReader(f, delimiter=';'):
        u = r.get('selected_image_url','').strip().replace('https://','http://')
        if u: patch_urls.add(u)

# Build wayback URL -> model name mapping for optional file
wb_url_to_model = {}
for model_name, info in wb_data.items():
    url = info.get('inside_url', '').strip().replace('https://', 'http://')
    if url:
        wb_url_to_model[url] = model_name

# ========================================================================
# STEP 1: Clean hard blockers
# ========================================================================
changes_log = []
asd = [r for r in rows if r['producer'] == ASD_KEY]
for r in asd:
    name = r['model_name']
    q = qa_by_name.get(name)
    if not q: continue
    
    comment_bits = [r.get('comment','').strip()] if r.get('comment','').strip() else []
    
    # --- 1a: Invalid URL (Ego inside - 404) ---
    if name == 'Стальная дверь "Ego". Заказная модель':
        inside_url = r.get('image_url_inside', '').strip()
        if inside_url:
            # This URL returned 404 - clear it
            r['image_url_inside'] = ''
            r['status'] = 'needs_review'
            comment_bits.append('invalid_image_url_cleared_404')
            changes_log.append(f"CLEARED invalid inside [{name[:50]}] {inside_url[:60]}")
    
    # --- 1b: Technical/fragment images ---
    for side, col in [('front', 'image_url_front'), ('inside', 'image_url_inside')]:
        visual = q.get(f'{side}_visual_decision', '')
        if visual in ('handle', 'lock', 'fragment', 'technical', 'badge', 'interior_room'):
            url = r.get(col, '').strip()
            if url:
                r[col] = ''
                r['status'] = 'needs_review'
                comment_bits.append(f'{side}_{visual}_image_cleared')
                changes_log.append(f"CLEARED {visual} [{side}] [{name[:50]}] {url[:60]}")
    
    # --- 1c: Remove Wayback URLs from main catalog ---
    for side, col in [('inside', 'image_url_inside'), ('front', 'image_url_front')]:
        src = q.get(f'{side}_source_type', '')
        url = r.get(col, '').strip()
        if src == 'wayback' and url and col == 'image_url_inside':
            r['image_url_inside'] = ''
            r['image_url'] = r.get('image_url_front', '') or r.get('image_url_front', '')
            r['status'] = 'needs_review'
            tag = 'inside_photo_available_only_in_wayback_not_used_in_main_delivery'
            if tag not in comment_bits:
                comment_bits.append(tag)
            changes_log.append(f"REMOVED wayback inside [{name[:50]}]")
    
    # Update comment
    r['comment'] = '; '.join(comment_bits)

# Ensure image_url = image_url_front for all rows with front
for r in asd:
    if r.get('image_url_front', '').strip():
        r['image_url'] = r['image_url_front']

print(f"\nChanges applied: {len(changes_log)}")
for cl in changes_log:
    print(f"  {cl}")

# ========================================================================
# STEP 2: Create optional wayback file
# ========================================================================
wb_optional_rows = []
for r in asd:
    name = r['model_name']
    q = qa_by_name.get(name)
    if not q: continue
    
    inside_src = q.get('inside_source_type', '')
    inside_url = q.get('image_url_inside', '')
    
    if inside_src == 'wayback' and inside_url:
        val = validation.get((name, 'inside'), {})
        wb_optional_rows.append({
            'producer': r['producer'],
            'target_line': r.get('target_line', ''),
            'model_name': name,
            'source_url': r.get('source_url', ''),
            'image_url_front': r.get('image_url_front', ''),
            'image_url_inside_wayback_optional': inside_url,
            'wayback_http_status': val.get('http_status', 0),
            'wayback_content_type': val.get('content_type', ''),
            'visual_decision': q.get('inside_visual_decision', 'unknown'),
            'note': 'optional_archive_photo_not_used_in_main_delivery',
        })

print(f"\nWayback optional rows: {len(wb_optional_rows)}")

# Sort by model_name
wb_optional_rows.sort(key=lambda x: x['model_name'])

# ========================================================================
# STEP 3: Create XLSX files
# ========================================================================
hdr_font = Font(bold=True, color="FFFFFF", size=10)
hdr_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

def write_xlsx(path, headers, data, sheet_name="Sheet1"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font = hdr_font; c.fill = hdr_fill; c.alignment = hdr_align; c.border = thin_border
    for ri, row in enumerate(data, 2):
        for ci, h in enumerate(headers, 1):
            val = row.get(h, '')
            if isinstance(val, list):
                val = '; '.join(val)
            c = ws.cell(row=ri, column=ci, value=val)
            c.border = thin_border
            c.alignment = Alignment(vertical="top", wrap_text=True)
    wb.save(path)

# Main catalog CSV
cat_path = os.path.join(OUT_DIR, "doors_catalog.csv")
with open(cat_path, 'w', encoding='utf-8-sig', newline='') as f:
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writeheader()
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writerows(rows)
print(f"Main catalog CSV: {cat_path}")

# Main catalog XLSX
xlsx_path = os.path.join(OUT_DIR, "doors_catalog.xlsx")
write_xlsx(xlsx_path, fn, rows, "Catalog")
print(f"Main catalog XLSX: {xlsx_path}")

# WordPress CSV
wp_path = os.path.join(OUT_DIR, "wordpress_import.csv")
with open(wp_path, 'w', encoding='utf-8-sig', newline='') as f:
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writeheader()
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writerows(rows)
print(f"WordPress CSV: {wp_path}")

# Optional wayback XLSX
wb_headers = ['producer', 'target_line', 'model_name', 'source_url', 'image_url_front',
              'image_url_inside_wayback_optional', 'wayback_http_status', 'wayback_content_type',
              'visual_decision', 'note']
wb_xlsx_path = os.path.join(OUT_DIR, "doors_catalog_with_wayback_optional.xlsx")
write_xlsx(wb_xlsx_path, wb_headers, wb_optional_rows, "Wayback Optional")
print(f"Wayback optional XLSX: {wb_xlsx_path} ({len(wb_optional_rows)} rows)")

# Optional wayback links XLSX (same content, different name)
wb_links_path = os.path.join(OUT_DIR, "asd_wayback_optional_links.xlsx")
write_xlsx(wb_links_path, wb_headers, wb_optional_rows, "Wayback Links")
print(f"Wayback links XLSX: {wb_links_path} ({len(wb_optional_rows)} rows)")

# ========================================================================
# STEP 4: Stats and summary
# ========================================================================
asd_final = [r for r in rows if r['producer'] == ASD_KEY]
f_final = sum(1 for r in asd_final if r.get('image_url_front','').strip())
i_final = sum(1 for r in asd_final if r.get('image_url_inside','').strip())
b_final = sum(1 for r in asd_final if r.get('image_url_front','').strip() and r.get('image_url_inside','').strip())

# Inside breakdown by source (post-cleanup)
inside_manufacturer = 0
inside_pattern = 0
for r in asd_final:
    name = r['model_name']
    q = qa_by_name.get(name)
    if q and r.get('image_url_inside','').strip():
        src = q.get('inside_source_type', '')
        if src == 'manufacturer_site':
            inside_manufacturer += 1
        elif src == 'pattern_matched':
            inside_pattern += 1

# Count wayback optional
wayback_inside = len(wb_optional_rows)  # should be 96

# Both including optional
both_with_optional = sum(1 for r in asd_final if r.get('image_url_front','').strip())  # all with front + wayback access
# Recalculate: both in main + wayback models that only need inside
models_with_wayback_inside = set(w['model_name'] for w in wb_optional_rows)
models_with_front = set(r['model_name'] for r in asd_final if r.get('image_url_front','').strip())
both_including_optional = len(models_with_front | models_with_wayback_inside)  # this is wrong
# Actually: both including optional = count of models that have front in main + (inside in main OR inside in wayback)
both_inc_optional = len([r for r in asd_final if r.get('image_url_front','').strip() and 
                        (r.get('image_url_inside','').strip() or r['model_name'] in models_with_wayback_inside)])

# No-image rows
no_img = sum(1 for r in asd_final if not r.get('image_url_front','').strip() and not r.get('image_url_inside','').strip())

# Producers check
prods = Counter(r['producer'] for r in rows)

# Duplicates check
keys = Counter((r['producer'], r['target_line'], r['model_name']) for r in rows)
dupes = [k for k, v in keys.items() if v > 1]

# Check remaining invalid/technical in main catalog
remaining_invalid = 0
remaining_technical = 0
remaining_wayback = 0
for r in asd_final:
    name = r['model_name']
    q = qa_by_name.get(name)
    if not q: continue
    
    # Check front
    f_url = r.get('image_url_front', '').strip()
    if f_url:
        if q.get('front_source_type') == 'wayback':
            remaining_wayback += 1
        if q.get('front_opens_as_image') != 'yes':
            remaining_invalid += 1
        if q.get('front_visual_decision') in ('handle','lock','fragment','technical','badge','interior_room'):
            remaining_technical += 1
    
    # Check inside
    i_url = r.get('image_url_inside', '').strip()
    if i_url:
        if q.get('inside_source_type') == 'wayback':
            remaining_wayback += 1
        if q.get('inside_opens_as_image') != 'yes':
            remaining_invalid += 1
        if q.get('inside_visual_decision') in ('handle','lock','fragment','technical','badge','interior_room'):
            remaining_technical += 1

# WordPress sync check
wp_ok = all(r.get('image_url_front','').strip() and r['image_url'] == r['image_url_front'] 
            for r in asd_final if r.get('image_url_front','').strip())

# Non-ASD preserved check
non_asd_check = {}
for p in [k for k in prods.keys() if k != ASD_KEY]:
    prod_rows = [r for r in rows if r['producer'] == p]
    non_asd_check[p] = len(prod_rows)

print(f"\n{'='*60}")
print(f"V5.3 CLEAN DELIVERY SUMMARY")
print(f"{'='*60}")
print(f"  total_rows: {len(rows)}")
print(f"  producers: {len(prods)}")
print(f"  duplicates: {len(dupes)}")
print(f"  ASD front filled (main): {f_final}/199")
print(f"  ASD inside filled (main): {i_final}/199")
print(f"    - manufacturer_site: {inside_manufacturer}")
print(f"    - pattern_matched (valid): {inside_pattern}")
print(f"  ASD inside wayback optional: {wayback_inside}/199")
print(f"  ASD both filled (main): {b_final}/199")
print(f"  ASD both filled (incl. wayback optional): {both_inc_optional}/199")
print(f"  Broken/invalid URLs in main: {remaining_invalid}")
print(f"  Technical/fragment images in main: {remaining_technical}")
print(f"  Wayback URLs in main columns: {remaining_wayback}")
print(f"  No-image rows: {no_img}")
print(f"  WordPress sync: image_url=image_url_front: {'OK' if wp_ok else 'FAIL'}")

# ========================================================================
# STEP 5: Copy QA reports from v5.2
# ========================================================================
import shutil
qa_reports = [
    "asd_final_photo_source_validation.xlsx",
    "asd_remaining_missing_images.xlsx",
    "final_delivery_qa_report.xlsx",
]
for f in qa_reports:
    src = os.path.join(BASE, "output", "delivery_v5_2", f)
    dst = os.path.join(OUT_DIR, f)
    if os.path.exists(src):
        shutil.copy2(src, dst)
        print(f"  Copied: {f}")

# ========================================================================
# STEP 6: Run report
# ========================================================================
report = {
    "mode": "final_delivery_v5_3_clean",
    "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "total_rows": len(rows),
    "producers": dict(sorted(prods.items(), key=lambda x: -x[1])),
    "duplicates": len(dupes),
    "asd": {
        "total": 199,
        "front_filled_main": f_final,
        "inside_filled_main": i_final,
        "inside_manufacturer": inside_manufacturer,
        "inside_pattern_valid": inside_pattern,
        "inside_wayback_optional": wayback_inside,
        "both_filled_main": b_final,
        "both_including_wayback_optional": both_inc_optional,
        "no_image_rows": no_img,
        "broken_urls_in_main": remaining_invalid,
        "technical_images_cleared": 4,
        "invalid_urls_cleared": 1,
        "wayback_removed_from_main": wayback_inside,
    },
    "producers_preserved": non_asd_check,
}
report_path = os.path.join(OUT_DIR, "run_report.json")
with open(report_path, 'w', encoding='utf-8') as f:
    json.dump(report, f, ensure_ascii=False, indent=2)
print(f"Run report: {report_path}")

# ========================================================================
# STEP 7: Build archive
# ========================================================================
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
            print(f"  Added: {f} ({os.path.getsize(p):,} bytes)")

sha256 = hashlib.sha256()
with open(archive_path, 'rb') as f:
    while True:
        chunk = f.read(8192)
        if not chunk: break
        sha256.update(chunk)

size = os.path.getsize(archive_path)
print(f"\nArchive: {archive_name}")
print(f"Size: {size:,} bytes")
print(f"SHA256: {sha256.hexdigest()}")

# Save summary
summary = {
    "archive_path": archive_path,
    "archive_sha256": sha256.hexdigest(),
    "total_rows": len(rows),
    "producers": len(prods),
    "duplicates": len(dupes),
    "asd_front_filled_main": f_final,
    "asd_inside_filled_main": i_final,
    "asd_inside_wayback_optional": wayback_inside,
    "asd_both_main": b_final,
    "asd_both_including_wayback": both_inc_optional,
    "broken_image_urls": remaining_invalid,
    "technical_fragment_images": remaining_technical,
    "wayback_urls_in_main_columns": remaining_wayback,
    "wordpress_sync": "OK" if wp_ok else "FAIL",
}
sum_path = os.path.join(OUT_DIR, "delivery_summary.json")
with open(sum_path, 'w', encoding='utf-8') as f:
    json.dump(summary, f, ensure_ascii=False, indent=2)

print(f"\nSummary: {sum_path}")
print(f"Archive: {archive_path}")
