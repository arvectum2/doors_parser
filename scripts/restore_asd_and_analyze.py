import csv
import os
import json
import urllib.request
import urllib.error
import ssl
import time
from collections import Counter, defaultdict
from datetime import datetime
from urllib.parse import urlparse
import hashlib

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CATALOG_V5 = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")
V4_PATH = os.path.join(BASE, "output", "final_rework_ready_v4", "doors_catalog.csv")
PATCH_PATH = os.path.join(BASE, "output", "asd_photo_patch_v2_clean", "asd_photo_patch_v2_clean.csv")
OUT_DIR = os.path.join(BASE, "output", "delivery")

EXPECTED_PRODUCERS = [
    "Скамбио Порте", "Хаусдорс", "Олимп", "Аэлита", "Лорд",
    "Комфорт", "Двери регионов", "Термодур", "Энтро", "АСД"
]

# Load data
with open(CATALOG_V5, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
fieldnames = list(rows[0].keys())

with open(V4_PATH, encoding='utf-8-sig') as f:
    v4_rows = list(csv.DictReader(f, delimiter=';'))
v4_asd = {r['model_name']: r for r in v4_rows if r['producer'] == 'АСД'}

with open(PATCH_PATH, encoding='utf-8-sig') as f:
    patch_rows = list(csv.DictReader(f, delimiter=';'))
patch_map = {r['model_name']: r.get('selected_image_url', '').strip() for r in patch_rows if r.get('selected_image_url', '').strip()}

print(f"Loaded: {len(rows)} catalog rows, {len(v4_asd)} v4 ASD, {len(patch_map)} patch URLs")

# ============================================================
# STEP 1: Restore ASD front images from v2 patch
# ============================================================
print(f"\n=== STEP 1: RESTORE ASD FRONT IMAGES ===")

asd_rows = [r for r in rows if r['producer'] == 'АСД']
asd_restored = 0
asd_already_ok = 0
asd_no_patch = 0

for r in asd_rows:
    name = r['model_name']
    current_front = r.get('image_url_front', '').strip()
    
    if current_front:
        asd_already_ok += 1
        continue
    
    patch_url = patch_map.get(name, '')
    if patch_url:
        # Use HTTP instead of HTTPS (HTTPS has redirect loop)
        http_url = patch_url.replace('https://', 'http://')
        r['image_url_front'] = http_url
        if not r.get('image_url', '').strip():
            r['image_url'] = http_url
        asd_restored += 1
    else:
        asd_no_patch += 1

print(f"ASD front restored: {asd_restored}")
print(f"ASD front already OK: {asd_already_ok}")
print(f"ASD front no patch URL: {asd_no_patch}")

# ============================================================
# STEP 2: Check ASD inside images
# ============================================================
print(f"\n=== STEP 2: ASD INSIDE IMAGES ===")

asd_inside_ok = sum(1 for r in asd_rows if r.get('image_url_inside', '').strip())
asd_inside_missing = len(asd_rows) - asd_inside_ok
print(f"ASD inside OK: {asd_inside_ok}/199")
print(f"ASD inside missing: {asd_inside_missing}/199")

# List missing inside rows
missing_inside = [(r['model_name'], r.get('image_url_front', '')) for r in asd_rows if not r.get('image_url_inside', '').strip()]
print(f"\nASD rows missing inside (first 10):")
for name, front in missing_inside[:10]:
    print(f"  {name[:50]}: front={'YES' if front else 'NO'}")

# ============================================================
# STEP 3: Global image analysis
# ============================================================
print(f"\n=== STEP 3: GLOBAL IMAGE ANALYSIS ===")

total = len(rows)
producers = Counter(r['producer'] for r in rows)

# Per-row image analysis
no_image_rows = []
has_image_rows = []
for r in rows:
    image_url = r.get('image_url', '').strip()
    image_url_front = r.get('image_url_front', '').strip()
    image_url_inside = r.get('image_url_inside', '').strip()
    has_any = bool(image_url or image_url_front or image_url_inside)
    
    if has_any:
        has_image_rows.append(r)
    else:
        no_image_rows.append(r)

print(f"Total rows: {total}")
print(f"With any image: {len(has_image_rows)}")
print(f"Without any image: {len(no_image_rows)}")

# No-image by producer
no_img_by_prod = Counter(r['producer'] for r in no_image_rows)
print(f"\nNo-image by producer:")
for p in EXPECTED_PRODUCERS:
    count = no_img_by_prod.get(p, 0)
    if count > 0:
        print(f"  {p}: {count}")

# No-image by reason
no_img_reasons = defaultdict(list)
for r in no_image_rows:
    comment = r.get('comment', '')
    if '404' in comment:
        no_img_reasons['image_404_on_manufacturer_site'].append(r)
    elif '302' in comment:
        no_img_reasons['image_302_redirect_cleared'].append(r)
    elif 'cleared' in comment:
        no_img_reasons['image_url_cleared_broken'].append(r)
    elif 'network' in comment.lower():
        no_img_reasons['network_unreachable'].append(r)
    elif '[MISSING]' in r.get('model_name', ''):
        no_img_reasons['stub_model_no_source'].append(r)
    elif not comment or comment == 'found':
        no_img_reasons['empty_or_found_comment'].append(r)
    else:
        # Extract first meaningful part of comment
        first_part = comment.split(';')[0].strip()[:50]
        no_img_reasons[f'other:{first_part}'].append(r)

print(f"\nNo-image by reason:")
for reason, items in sorted(no_img_reasons.items(), key=lambda x: -len(x[1])):
    prods = Counter(i['producer'] for i in items)
    print(f"  {reason}: {len(items)} — {dict(prods)}")

# ============================================================
# STEP 4: dveri-lord.ru analysis
# ============================================================
print(f"\n=== STEP 4: DVERI-LORD.RU ===")
lord_rows = [r for r in rows if r['producer'] == 'Лорд']
lord_no_image = [r for r in lord_rows if not r.get('image_url', '').strip() and not r.get('image_url_front', '').strip()]
lord_with_dveri_lord = [r for r in lord_rows if 'dveri-lord.ru' in (r.get('image_url', '') or '') or 'dveri-lord.ru' in (r.get('image_url_front', '') or '')]
print(f"Lord total: {len(lord_rows)}")
print(f"Lord no image: {len(lord_no_image)}")
print(f"Lord with dveri-lord.ru URLs: {len(lord_with_dveri_lord)}")

# ============================================================
# STEP 5: as-doors.ru 302 analysis
# ============================================================
print(f"\n=== STEP 5: AS-DOORS.RU 302 ANALYSIS ===")
as_doors_302 = [r for r in rows if '302' in r.get('comment', '')]
print(f"Rows with 302 in comment: {len(as_doors_302)}")
as_doors_302_by_prod = Counter(r['producer'] for r in as_doors_302)
print(f"By producer: {dict(as_doors_302_by_prod)}")

# ============================================================
# STEP 6: Domain analysis
# ============================================================
print(f"\n=== STEP 6: DOMAIN ANALYSIS ===")
all_urls = set()
url_domains = Counter()
for r in rows:
    for field in ['image_url', 'image_url_front', 'image_url_inside']:
        url = r.get(field, '').strip()
        if url:
            all_urls.add(url)
            try:
                parsed = urlparse(url)
                domain = parsed.netloc or parsed.path.split('/')[0]
                url_domains[domain] += 1
            except:
                pass

print(f"Unique image URLs: {len(all_urls)}")
print(f"URLs by domain:")
for domain, count in url_domains.most_common():
    print(f"  {domain}: {count}")

# ============================================================
# STEP 7: Duplicates check
# ============================================================
print(f"\n=== STEP 7: DUPLICATES CHECK ===")
keys = set()
dupes = 0
for r in rows:
    key = (r['producer'], r['target_line'], r['model_name'])
    if key in keys:
        dupes += 1
    keys.add(key)
print(f"Duplicates: {dupes}")

# ============================================================
# STEP 8: WordPress sync
# ============================================================
print(f"\n=== STEP 8: WORDPRESS SYNC ===")
wp_path = os.path.join(BASE, "output", "final_rework_ready_v5", "wordpress_import.csv")
xlsx_path = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.xlsx")
report_path = os.path.join(BASE, "output", "final_rework_ready_v5", "run_report.json")

wp_rows_count = 0
if os.path.exists(wp_path):
    with open(wp_path, encoding='utf-8-sig') as f:
        wp_rows_count = len(list(csv.DictReader(f, delimiter=';')))
xlsx_exists = os.path.exists(xlsx_path)
with open(report_path, encoding='utf-8') as f:
    report = json.load(f)
report_rows = report.get('total_rows', 0)

print(f"CSV rows: {total}")
print(f"XLSX exists: {xlsx_exists}")
print(f"WP CSV rows: {wp_rows_count}")
print(f"Report rows: {report_rows}")
print(f"All match: {total == wp_rows_count == report_rows}")

# Save catalog with restored ASD front images
with open(CATALOG_V5, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=';')
    writer.writeheader()
    writer.writerows(rows)

print(f"\n=== SAVED: {CATALOG_V5} ===")
print(f"ASD front restored: {asd_restored}")
print(f"ASD inside still missing: {asd_inside_missing}")

# Save summary for report generation
summary = {
    'total_rows': total,
    'producers_count': len(set(r['producer'] for r in rows)),
    'duplicates': dupes,
    'rows_with_any_image': len(has_image_rows),
    'rows_without_any_image': len(no_image_rows),
    'no_image_by_producer': dict(no_img_by_prod),
    'asd': {
        'total': len(asd_rows),
        'front_ok': asd_already_ok + asd_restored,
        'front_restored': asd_restored,
        'front_no_patch': asd_no_patch,
        'inside_ok': asd_inside_ok,
        'inside_missing': asd_inside_missing,
    },
    'as_doors_302_count': len(as_doors_302),
    'lord_no_image': len(lord_no_image),
    'unique_urls': len(all_urls),
    'wordpress_sync': total == wp_rows_count == report_rows,
}

with open(os.path.join(OUT_DIR, 'reconciliation_summary.json'), 'w', encoding='utf-8') as f:
    json.dump(summary, f, ensure_ascii=False, indent=2)

print(f"\nSaved reconciliation_summary.json")
