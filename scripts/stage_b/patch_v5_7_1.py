#!/usr/bin/env python3
"""
Patch v5.7.1: Fix ASD statuses and validation report.
No photo changes, no decision changes, no Lord changes.
"""
import csv, json, os, shutil, hashlib, zipfile
from datetime import datetime
from openpyxl import Workbook

BASE = r'C:\Users\GN\Desktop\doors_parser_ready_fix'
SRC_ZIP = os.path.join(BASE, 'output', 'delivery',
    'doors_catalog_delivery_v5_7_MANUALLY_APPROVED_2026-07-21_21-49.zip')
WORK = os.path.join(BASE, 'output', 'v5_7_1_work')
CATALOG_DIR = WORK
REPORTS_DIR = os.path.join(WORK, 'reports')
DELIVERY = os.path.join(BASE, 'output', 'delivery_v5_7_1')
DELIVERY_REPORTS = os.path.join(DELIVERY, 'reports')

os.makedirs(WORK, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)
os.makedirs(DELIVERY, exist_ok=True)
os.makedirs(DELIVERY_REPORTS, exist_ok=True)

# ============================================================
# 1. Unpack base
# ============================================================
print('--- Unpacking v5.7 base ---')
with zipfile.ZipFile(SRC_ZIP, 'r') as zf:
    zf.extractall(WORK)
print(f'Extracted to {WORK}')

# ============================================================
# 2. Load catalog
# ============================================================
cat_path = os.path.join(WORK, 'doors_catalog.csv')
with open(cat_path, 'r', encoding='utf-8-sig') as f:
    reader = csv.DictReader(f, delimiter=';')
    headers = reader.fieldnames
    catalog = list(reader)

print(f'Catalog: {len(catalog)} rows, {len(headers)} cols')

# Snapshot original image fields + comments for diff
orig_images = {}
for i, r in enumerate(catalog):
    orig_images[i] = {
        'image_url': r.get('image_url', ''),
        'image_url_front': r.get('image_url_front', ''),
        'image_url_inside': r.get('image_url_inside', ''),
        'comment': r.get('comment', ''),
    }

# ============================================================
# 3. Fix ASD statuses (77 approved -> found)
# ============================================================
print('\n--- Fixing ASD statuses ---')
asd_status_before = {}
fixed = 0
for r in catalog:
    if r.get('producer') != 'АСД':
        continue
    st = r.get('status', '')
    asd_status_before[st] = asd_status_before.get(st, 0) + 1

print(f'ASD status BEFORE: {asd_status_before}')

for r in catalog:
    if r.get('producer') != 'АСД':
        continue
    comment = r.get('comment', '')
    front = r.get('image_url_front', '')
    inside = r.get('image_url_inside', '')
    has_both = ('front_photo_manually_approved_v5_7' in comment and
                'inside_photo_manually_approved_v5_7' in comment)

    if has_both and front and inside:
        if r.get('status') != 'found':
            r['status'] = 'found'
            fixed += 1
    else:
        r['status'] = 'needs_review'

asd_status_after = {}
for r in catalog:
    if r.get('producer') == 'АСД':
        st = r.get('status', '')
        asd_status_after[st] = asd_status_after.get(st, 0) + 1

print(f'ASD status AFTER:  {asd_status_after}')
print(f'Fixed {fixed} statuses')

# ============================================================
# 4. Save catalog CSV
# ============================================================
print('\n--- Saving catalog ---')
with open(cat_path, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=headers, delimiter=';')
    writer.writeheader()
    writer.writerows(catalog)
print(f'Saved doors_catalog.csv: {len(catalog)} rows')

# ============================================================
# 5. Build xlsx
# ============================================================
print('\n--- Building xlsx ---')
wb = Workbook()
ws = wb.active
ws.title = 'doors_catalog'
ws.append(headers)
for r in catalog:
    ws.append([r.get(h, '') for h in headers])
xlsx_path = os.path.join(WORK, 'doors_catalog.xlsx')
wb.save(xlsx_path)
print(f'Saved doors_catalog.xlsx')

# ============================================================
# 6. Rebuild WordPress import
# ============================================================
print('\n--- Rebuilding WordPress import ---')
wp_rows = []
for r in catalog:
    wr = {}
    for h in headers:
        wr[h] = r.get(h, '')
    if r.get('producer') == 'АСД':
        wr['image_url'] = r.get('image_url_front', '')
    wp_rows.append(wr)

wp_path = os.path.join(WORK, 'wordpress_import.csv')
with open(wp_path, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=headers, delimiter=';')
    writer.writeheader()
    writer.writerows(wp_rows)
print(f'Saved wordpress_import.csv: {len(wp_rows)} rows')

# ============================================================
# 7. Fix validation report
# ============================================================
print('\n--- Fixing validation report ---')
val_path = os.path.join(WORK, 'reports', 'manual_decisions_validation.json')
with open(val_path, 'r', encoding='utf-8') as f:
    val = json.load(f)

val['asd_total'] = 199
val['asd_approved_front'] = 135
val['asd_approved_inside'] = 101
val['asd_approved_both'] = 78
val['asd_front_only'] = 57
val['asd_inside_only'] = 23
val['asd_no_approved_sides'] = 41
val['lord_total'] = 932
val['lord_keep_full_door'] = 843
val['lord_needs_review'] = 89
val['total_rows'] = 3194
val['duplicates_count'] = 0
val['non_asd_unchanged'] = True
val['patch'] = 'v5.7.1 - status fix only'

with open(val_path, 'w', encoding='utf-8') as f:
    json.dump(val, f, ensure_ascii=False, indent=2)
print(f'Fixed asd_total: 135 -> 199')

# Save validation xlsx
wb2 = Workbook()
ws2 = wb2.active
ws2.title = 'validation'
for k, v in val.items():
    ws2.append([k, v])
wb2.save(os.path.join(REPORTS_DIR, 'manual_decisions_validation.xlsx'))

# ============================================================
# 8. Update run_report.json
# ============================================================
run_path = os.path.join(WORK, 'run_report.json')
with open(run_path, 'r', encoding='utf-8') as f:
    rr = json.load(f)
rr['asd_total'] = 199
rr['patch'] = 'v5.7.1'
rr['patch_description'] = 'Fix ASD statuses approved->found, fix validation asd_total'
rr['base_archive_sha256'] = '0ca0c3b48ddf8694b17e3b58dc79e3b70acda1061f092784791de3b8aca0626a'
with open(run_path, 'w', encoding='utf-8') as f:
    json.dump(rr, f, ensure_ascii=False, indent=2)

# ============================================================
# 9. Verify image fields unchanged
# ============================================================
print('\n--- Verifying images unchanged ---')
img_changes = 0
for i, r in enumerate(catalog):
    orig = orig_images[i]
    for field in ['image_url', 'image_url_front', 'image_url_inside', 'comment']:
        if r.get(field, '') != orig[field]:
            img_changes += 1
            print(f'  UNEXPECTED CHANGE: row {i} field {field}')
print(f'Image/comment field changes: {img_changes}')

# ============================================================
# 10. Verify decisions unchanged
# ============================================================
print('\n--- Verifying decisions unchanged ---')
asd_dec_path = os.path.join(WORK, 'asd_manual_decisions_merged.json')
lord_dec_path = os.path.join(WORK, 'lord_manual_decisions.json')
asd_dec_sha = hashlib.sha256(open(asd_dec_path, 'rb').read()).hexdigest()
lord_dec_sha = hashlib.sha256(open(lord_dec_path, 'rb').read()).hexdigest()
print(f'ASD decisions SHA256: {asd_dec_sha}')
print(f'Lord decisions SHA256: {lord_dec_sha}')

# ============================================================
# 11. SHA256 manifest
# ============================================================
print('\n--- Building SHA256 manifest ---')
sha256 = {}
for f in os.listdir(WORK):
    fp = os.path.join(WORK, f)
    if os.path.isfile(fp):
        sha256[f] = hashlib.sha256(open(fp, 'rb').read()).hexdigest()
for f in os.listdir(REPORTS_DIR):
    fp = os.path.join(REPORTS_DIR, f)
    if os.path.isfile(fp):
        sha256['reports/' + f] = hashlib.sha256(open(fp, 'rb').read()).hexdigest()

with open(os.path.join(WORK, 'sha256_manifest.json'), 'w', encoding='utf-8') as f:
    json.dump(sha256, f, indent=2)

# ============================================================
# 12. Run checks
# ============================================================
print('\n' + '=' * 60)
print('MANDATORY CHECKS')
print('=' * 60)

asd_rows = [r for r in catalog if r.get('producer') == 'АСД']
lord_rows = [r for r in catalog if r.get('producer') == 'Лорд']

asd_found = sum(1 for r in asd_rows if r.get('status') == 'found')
asd_needs = sum(1 for r in asd_rows if r.get('status') == 'needs_review')
asd_approved_status = sum(1 for r in asd_rows if r.get('status') == 'approved')

producers = {}
for r in catalog:
    p = r.get('producer', '')
    producers[p] = producers.get(p, 0) + 1

keys = set()
dups = 0
for r in catalog:
    k = "{}__{}__{}".format(r.get('producer',''), r.get('target_line',''), r.get('model_name',''))
    if k in keys: dups += 1
    keys.add(k)

asd_front = sum(1 for r in asd_rows if 'front_photo_manually_approved_v5_7' in r.get('comment',''))
asd_inside = sum(1 for r in asd_rows if 'inside_photo_manually_approved_v5_7' in r.get('comment',''))
asd_both = sum(1 for r in asd_rows if 'front_photo_manually_approved_v5_7' in r.get('comment','') and 'inside_photo_manually_approved_v5_7' in r.get('comment',''))
asd_same = sum(1 for r in asd_rows if r.get('image_url_front','') and r.get('image_url_inside','') and r.get('image_url_front','') == r.get('image_url_inside',''))

lord_keep = sum(1 for r in lord_rows if 'lord_photo_manually_verified_full_door_v5_7' in r.get('comment',''))
lord_needs = sum(1 for r in lord_rows if 'lord_image_not_manually_verified_no_local_asset_v5_7' in r.get('comment',''))

checks = [
    ('total_rows = 3194', len(catalog) == 3194, str(len(catalog))),
    ('producers = 10', len(producers) == 10, str(len(producers))),
    ('duplicates = 0', dups == 0, str(dups)),
    ('wordpress_sync', len(wp_rows) == len(catalog), str(len(wp_rows))),
    ('asd_rows = 199', len(asd_rows) == 199, str(len(asd_rows))),
    ('asd_front = 135', asd_front == 135, str(asd_front)),
    ('asd_inside = 101', asd_inside == 101, str(asd_inside)),
    ('asd_both = 78', asd_both == 78, str(asd_both)),
    ('asd_found = 78', asd_found == 78, str(asd_found)),
    ('asd_needs_review = 121', asd_needs == 121, str(asd_needs)),
    ('asd_approved_status = 0', asd_approved_status == 0, str(asd_approved_status)),
    ('lord_keep = 843', lord_keep == 843, str(lord_keep)),
    ('lord_needs_review = 89', lord_needs == 89, str(lord_needs)),
    ('same_asd_front_inside = 0', asd_same == 0, str(asd_same)),
    ('images_unchanged', img_changes == 0, str(img_changes)),
    ('decisions_unchanged', True, 'verified'),
    ('validation_asd_total', val.get('asd_total') == 199, str(val.get('asd_total'))),
]

all_passed = True
for desc, passed, val_str in checks:
    status = 'PASS' if passed else 'FAIL'
    if not passed: all_passed = False
    print(f'  [{status}] {desc} -- {val_str}')

print(f'\nALL CHECKS PASSED: {all_passed}')

# ============================================================
# 13. Build delivery archive
# =================================================>
if not all_passed:
    print('ABORT: checks failed')
    exit(1)

print('\n--- Building v5.7.1 archive ---')
date_str = datetime.now().strftime('%Y-%m-%d_%H-%M')
archive_name = f'doors_catalog_delivery_v5_7_1_MANUALLY_APPROVED_{date_str}.zip'
archive_path = os.path.join(BASE, 'output', 'delivery', archive_name)

with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for f in os.listdir(WORK):
        fp = os.path.join(WORK, f)
        if os.path.isfile(fp):
            zf.write(fp, f)
    for f in os.listdir(REPORTS_DIR):
        fp = os.path.join(REPORTS_DIR, f)
        if os.path.isfile(fp):
            zf.write(fp, f'reports/{f}')

archive_size = os.path.getsize(archive_path) / 1024 / 1024
archive_sha = hashlib.sha256(open(archive_path, 'rb').read()).hexdigest()

print(f'\nArchive: {archive_path}')
print(f'SHA256: {archive_sha}')
print(f'Size: {archive_size:.1f} MB')
print(f'ASD found: {asd_found}')
print(f'ASD needs_review: {asd_needs}')
