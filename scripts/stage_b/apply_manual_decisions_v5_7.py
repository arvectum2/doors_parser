#!/usr/bin/env python3
"""
Stage B: Apply manual review decisions to catalog v5.6 -> v5.7 MANUALLY APPROVED.

Strict rules:
- No new parsing, no new image fetching
- No old heuristics, no v5.5/v5.4 URL restoration
- Only apply decisions from manual_review_decisions_v5_7.zip
- Current URLs not in approved fields -> not approved
"""

import csv
import json
import os
import re
import shutil
import hashlib
from datetime import datetime
from collections import defaultdict

BASE = r'C:\Users\GN\Desktop\doors_parser_ready_fix'
STAGE_B = os.path.join(BASE, 'output', 'stage_b_manual_decisions_v5_7')
BASE_CATALOG = os.path.join(STAGE_B, 'base_catalog_v5_6')
DELIVERY_OUT = os.path.join(BASE, 'output', 'delivery_v5_7')
REPORTS_DIR = os.path.join(DELIVERY_OUT, 'reports')
BEFORE_DIR = os.path.join(STAGE_B, 'before_apply')

os.makedirs(DELIVERY_OUT, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)
os.makedirs(BEFORE_DIR, exist_ok=True)


def normalize_name(s):
    """Normalize model name for matching."""
    if not s:
        return ''
    s = s.lower().replace('ё', 'е')
    s = re.sub(r'[«»"\'„"]', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    s = re.sub(r'[-–—]', '-', s)
    s = re.sub(r'[\u200b\u200c\u200d\ufeff]', '', s)
    return s


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()


def load_csv(path):
    with open(path, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f, delimiter=';')
        return list(reader), reader.fieldnames


def save_csv(rows, headers, path):
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=headers, delimiter=';')
        writer.writeheader()
        writer.writerows(rows)


print('=' * 60)
print('STAGE B: APPLY MANUAL REVIEW DECISIONS')
print('=' * 60)

# ============================================================
# STEP 1: Load all data
# ============================================================
print('\n--- Loading data ---')

# Load base catalog
catalog_rows, catalog_headers = load_csv(os.path.join(BASE_CATALOG, 'doors_catalog.csv'))
print(f'Catalog: {len(catalog_rows)} rows')

# Save backup
for f in ['doors_catalog.csv', 'doors_catalog.xlsx', 'wordpress_import.csv', 'run_report.json']:
    src = os.path.join(BASE_CATALOG, f)
    if os.path.exists(src):
        shutil.copy2(src, os.path.join(BEFORE_DIR, f))
        print(f'Backed up: {f}')

# Load ASD merged decisions
with open(os.path.join(STAGE_B, 'asd_manual_decisions_merged.json'), 'r', encoding='utf-8') as f:
    asd_decisions = json.load(f)
print(f'ASD decisions: {len(asd_decisions)} rows')

# Load Lord decisions
with open(os.path.join(STAGE_B, 'lord', 'lord_manual_decisions.json'), 'r', encoding='utf-8') as f:
    lord_decisions = json.load(f)
print(f'Lord decisions: {len(lord_decisions)} rows')

# Load decisions archive SHA256
decisions_sha256 = sha256_file(r'C:\Users\GN\Downloads\manual_review_decisions_v5_7.zip')
print(f'Decisions archive SHA256: {decisions_sha256}')

# ============================================================
# STEP 2: Build decision lookup for ASD
# ============================================================
print('\n--- Building ASD decision lookup ---')

asd_lookup = {}
for d in asd_decisions:
    key = d.get('excel_row') or d.get('model_key') or d.get('source_url', '')
    asd_lookup[key] = d
    # Also index by source_url
    if d.get('source_url'):
        asd_lookup[d['source_url']] = d
    # Index by normalized name + target_line
    norm = normalize_name(d.get('model_name', ''))
    tl = d.get('target_line', '')
    if norm and tl:
        asd_lookup[f'{norm}__{tl}'] = d

# Build Lord decision lookup by excel_row
lord_lookup = {}
for d in lord_decisions:
    erow = str(d.get('excel_row', ''))
    lord_lookup[erow] = d
    if d.get('model_name'):
        lord_lookup[d['model_name']] = d

# ============================================================
# STEP 3: Create diff tracker
# ============================================================
diff_rows = []


def add_diff(producer, target_line, model_name, source_url, field, old_val, new_val, decision_source, decision, reason):
    diff_rows.append({
        'producer': producer,
        'target_line': target_line,
        'model_name': model_name,
        'source_url': source_url,
        'field_name': field,
        'old_value': old_val if old_val else '',
        'new_value': new_val if new_val else '',
        'decision_source': decision_source,
        'decision': decision,
        'reason': reason
    })


# ============================================================
# STEP 4: Apply ASD decisions
# ============================================================
print('\n--- Applying ASD decisions ---')

asd_applied = {'front': 0, 'inside': 0, 'both': 0, 'front_only': 0, 'inside_only': 0, 'no_approved': 0}

for i, row in enumerate(catalog_rows):
    if row.get('producer', '') != 'АСД':
        continue

    # Find decision
    excel_row = row.get('excel_row', '')
    source_url = row.get('source_url', '')
    model_name = row.get('model_name', '')
    target_line = row.get('target_line', '')
    norm_name = normalize_name(model_name)

    decision = asd_lookup.get(excel_row) or asd_lookup.get(source_url)
    if not decision and norm_name and target_line:
        decision = asd_lookup.get(f'{norm_name}__{target_line}')

    if not decision:
        print(f'  WARNING: No decision found for ASD row {excel_row}: {model_name}')
        continue

    old_front = row.get('image_url_front', '')
    old_inside = row.get('image_url_inside', '')
    old_image = row.get('image_url', '')
    old_status = row.get('status', '')
    old_comment = row.get('comment', '')

    has_front = decision.get('front_explicitly_approved', False)
    front_url = decision.get('approved_front_url', '')
    has_inside = decision.get('inside_explicitly_approved', False)
    inside_url = decision.get('approved_inside_url', '')

    new_comment_parts = []

    # Apply front
    if has_front and front_url:
        row['image_url_front'] = front_url
        row['image_url'] = front_url
        new_comment_parts.append('front_photo_manually_approved_v5_7')
        if front_url != old_front:
            add_diff('АСД', target_line, model_name, source_url, 'image_url_front', old_front, front_url, 'manual_review', 'approve_front', 'explicitly approved')
        if front_url != old_image:
            add_diff('АСД', target_line, model_name, source_url, 'image_url', old_image, front_url, 'manual_review', 'approve_front', 'synced with front')
    else:
        row['image_url_front'] = ''
        row['image_url'] = ''
        new_comment_parts.append('front_photo_not_manually_approved_v5_7')
        if old_front:
            add_diff('АСД', target_line, model_name, source_url, 'image_url_front', old_front, '', 'manual_review', 'reject_front', 'not approved')
        if old_image:
            add_diff('АСД', target_line, model_name, source_url, 'image_url', old_image, '', 'manual_review', 'reject_front', 'cleared')

    # Apply inside
    if has_inside and inside_url:
        row['image_url_inside'] = inside_url
        new_comment_parts.append('inside_photo_manually_approved_v5_7')
        if inside_url != old_inside:
            add_diff('АСД', target_line, model_name, source_url, 'image_url_inside', old_inside, inside_url, 'manual_review', 'approve_inside', 'explicitly approved')
    else:
        row['image_url_inside'] = ''
        new_comment_parts.append('inside_photo_not_manually_approved_v5_7')
        if old_inside:
            add_diff('АСД', target_line, model_name, source_url, 'image_url_inside', old_inside, '', 'manual_review', 'reject_inside', 'not approved')

    # Status
    if has_front and has_inside:
        # Keep good status, remove contradictions
        new_status = old_status
        if old_status in ('needs_review', ''):
            new_status = 'approved'
    else:
        new_status = 'needs_review'
        missing = []
        if not has_front: missing.append('front')
        if not has_inside: missing.append('inside')
        new_comment_parts.append(f'missing_{ "_".join(missing) }_photo_v5_7')

    row['status'] = new_status
    row['comment'] = '; '.join(new_comment_parts)

    # Track stats
    if has_front and has_inside:
        asd_applied['both'] += 1
    elif has_front:
        asd_applied['front_only'] += 1
    elif has_inside:
        asd_applied['inside_only'] += 1
    else:
        asd_applied['no_approved'] += 1
    if has_front: asd_applied['front'] += 1
    if has_inside: asd_applied['inside'] += 1

print(f'  ASD front approved: {asd_applied["front"]}')
print(f'  ASD inside approved: {asd_applied["inside"]}')
print(f'  ASD both: {asd_applied["both"]}')
print(f'  ASD front only: {asd_applied["front_only"]}')
print(f'  ASD inside only: {asd_applied["inside_only"]}')
print(f'  ASD no approved: {asd_applied["no_approved"]}')

# ============================================================
# STEP 5: Apply Lord decisions
# ============================================================
print('\n--- Applying Lord decisions ---')

lord_stats = {'keep': 0, 'needs_review': 0, 'replaced': 0, 'cleared': 0}
lord_unapproved = []

for i, row in enumerate(catalog_rows):
    if row.get('producer', '') != 'Лорд':
        continue

    model_name = row.get('model_name', '')
    target_line = row.get('target_line', '')
    source_url = row.get('source_url', '')
    old_image = row.get('image_url', '')
    old_front = row.get('image_url_front', '')
    old_status = row.get('status', '')
    old_comment = row.get('comment', '')

    # Find decision by model_name
    decision = lord_lookup.get(model_name)
    if not decision:
        # Try by target_line
        for d in lord_decisions:
            if d.get('target_line') == target_line and d.get('model_name') == model_name:
                decision = d
                break

    if not decision:
        print(f'  WARNING: No decision found for Lord: {model_name}')
        continue

    d = decision.get('decision', 'needs_review')
    approved_url = decision.get('approved_replacement_url', '')
    new_comment_parts = []

    if d == 'keep_full_door':
        # Keep current image_url
        row['image_url'] = old_image
        if old_image:
            row['image_url_front'] = old_image
        new_comment_parts.append('lord_photo_manually_verified_full_door_v5_7')
        lord_stats['keep'] += 1

        add_diff('Лорд', target_line, model_name, source_url, 'status', old_status, old_status, 'manual_review', 'keep_full_door', 'verified full door')

    elif d == 'replace_with_approved_url' and approved_url:
        row['image_url'] = approved_url
        row['image_url_front'] = approved_url
        new_comment_parts.append('lord_photo_manually_replaced_v5_7')
        lord_stats['replaced'] += 1

        if old_image != approved_url:
            add_diff('Лорд', target_line, model_name, source_url, 'image_url', old_image, approved_url, 'manual_review', 'replace', 'approved replacement')
        if old_front != approved_url:
            add_diff('Лорд', target_line, model_name, source_url, 'image_url_front', old_front, approved_url, 'manual_review', 'replace', 'approved replacement')

    elif d in ('reject_fragment', 'reject_detail', 'reject_wrong_model', 'clear_no_replacement'):
        row['image_url'] = ''
        row['image_url_front'] = ''
        row['status'] = 'needs_review'
        new_comment_parts.append(f'lord_photo_{d}_v5_7')
        lord_stats['cleared'] += 1

        if old_image:
            add_diff('Лорд', target_line, model_name, source_url, 'image_url', old_image, '', 'manual_review', d, 'rejected')
        if old_front:
            add_diff('Лорд', target_line, model_name, source_url, 'image_url_front', old_front, '', 'manual_review', d, 'rejected')

    elif d == 'needs_review':
        # 89 rows without local image
        row['image_url'] = ''
        row['image_url_front'] = ''
        row['status'] = 'needs_review'
        new_comment_parts.append('lord_image_not_manually_verified_no_local_asset_v5_7')
        lord_stats['needs_review'] += 1

        # Save old URL to unapproved report
        lord_unapproved.append({
            'producer': 'Лорд',
            'target_line': target_line,
            'model_name': model_name,
            'source_url': source_url,
            'old_image_url': old_image,
            'manual_decision': d,
            'reason': 'no local image available for visual verification'
        })

        if old_image:
            add_diff('Лорд', target_line, model_name, source_url, 'image_url', old_image, '', 'manual_review', d, 'no local asset')

    row['comment'] = '; '.join(new_comment_parts)

print(f'  Lord keep_full_door: {lord_stats["keep"]}')
print(f'  Lord needs_review: {lord_stats["needs_review"]}')
print(f'  Lord replaced: {lord_stats["replaced"]}')
print(f'  Lord cleared: {lord_stats["cleared"]}')

# ============================================================
# STEP 6: Save modified catalog
# ============================================================
print('\n--- Saving modified catalog ---')

# Save catalog
save_csv(catalog_rows, catalog_headers, os.path.join(DELIVERY_OUT, 'doors_catalog.csv'))
print(f'Saved doors_catalog.csv: {len(catalog_rows)} rows')

# ============================================================
# STEP 7: Verify non-ASD producers unchanged
# ============================================================
print('\n--- Verifying non-ASD producers unchanged ---')

# Reload original
orig_rows, _ = load_csv(os.path.join(BASE_CATALOG, 'doors_catalog.csv'))

non_asd_producers = ['Скамбио Порте', 'Хаусдорс', 'Олимп', 'Аэлита', 'Комфорт', 'Двери регионов', 'Термодур', 'Энтро']
non_asd_unchanged = True
non_asd_diff_count = 0

for orig, new in zip(orig_rows, catalog_rows):
    if orig.get('producer', '') in non_asd_producers:
        for h in catalog_headers:
            if orig.get(h, '') != new.get(h, ''):
                non_asd_diff_count += 1
                non_asd_unchanged = False
                print(f'  DIFF in non-ASD: {orig.get("producer")} {orig.get("model_name")} field={h}')

print(f'Non-ASD changed fields: {non_asd_diff_count}')
print(f'Non-ASD unchanged: {non_asd_unchanged}')

# Count by producer in final
final_producers = {}
for r in catalog_rows:
    p = r.get('producer', 'unknown')
    final_producers[p] = final_producers.get(p, 0) + 1

print('\nFinal producer counts:')
for p, c in sorted(final_producers.items(), key=lambda x: -x[1]):
    print(f'  {p}: {c}')

# ============================================================
# STEP 8: Save diff
# ============================================================
print(f'\n--- Saving diff: {len(diff_rows)} changes ---')

if diff_rows:
    diff_headers = ['producer', 'target_line', 'model_name', 'source_url', 'field_name',
                    'old_value', 'new_value', 'decision_source', 'decision', 'reason']
    save_csv(diff_rows, diff_headers, os.path.join(REPORTS_DIR, 'manual_decisions_apply_diff.csv'))
    print(f'Saved manual_decisions_apply_diff.csv')

# ============================================================
# STEP 9: Save lord unapproved
# ============================================================
if lord_unapproved:
    lord_un_headers = ['producer', 'target_line', 'model_name', 'source_url', 'old_image_url', 'manual_decision', 'reason']
    save_csv(lord_unapproved, lord_un_headers, os.path.join(REPORTS_DIR, 'lord_unapproved_original_urls.csv'))
    print(f'Saved lord_unapproved_original_urls.csv: {len(lord_unapproved)} rows')

# ============================================================
# STEP 10: URL validation (diagnostic only)
# ============================================================
print('\n--- URL validation (diagnostic) ---')

import urllib.request

url_status = {'http_200_image': 0, 'network_unreachable_manually_approved': 0, 'redirect_manually_approved': 0, 'http_error_needs_attention': 0}
url_errors = []

approved_urls = set()
for row in catalog_rows:
    if row.get('producer') == 'АСД':
        if row.get('image_url_front'): approved_urls.add(row['image_url_front'])
        if row.get('image_url_inside'): approved_urls.add(row['image_url_inside'])
    elif row.get('producer') == 'Лорд':
        if row.get('image_url'): approved_urls.add(row['image_url'])

print(f'Total approved URLs to check: {len(approved_urls)}')

# Check a sample (not all to save time)
sample_urls = list(approved_urls)[:50]
for url in sample_urls:
    try:
        req = urllib.request.Request(url, method='HEAD')
        req.add_header('User-Agent', 'Mozilla/5.0')
        resp = urllib.request.urlopen(req, timeout=10)
        code = resp.getcode()
        ct = resp.headers.get('Content-Type', '')
        if code == 200 and 'image' in ct:
            url_status['http_200_image'] += 1
        elif code in (301, 302):
            url_status['redirect_manually_approved'] += 1
        else:
            url_status['http_error_needs_attention'] += 1
    except urllib.error.HTTPError as e:
        if e.code in (301, 302):
            url_status['redirect_manually_approved'] += 1
        else:
            url_status['http_error_needs_attention'] += 1
            url_errors.append({'url': url, 'error': str(e)})
    except Exception as e:
        url_status['network_unreachable_manually_approved'] += 1

print(f'URL check results (sample of {len(sample_urls)}):')
for k, v in url_status.items():
    print(f'  {k}: {v}')

# ============================================================
# STEP 11: Rebuild WordPress import
# ============================================================
print('\n--- Rebuilding WordPress import ---')

wp_rows = []
for row in catalog_rows:
    wp_row = {}
    for h in catalog_headers:
        wp_row[h] = row.get(h, '')
    # For ASD: image_url = image_url_front
    if row.get('producer') == 'АСД':
        wp_row['image_url'] = row.get('image_url_front', '')
    wp_rows.append(wp_row)

save_csv(wp_rows, catalog_headers, os.path.join(DELIVERY_OUT, 'wordpress_import.csv'))
print(f'Saved wordpress_import.csv: {len(wp_rows)} rows')

# ============================================================
# STEP 12: Create reports
# ============================================================
print('\n--- Creating reports ---')

# ASD decisions applied report
asd_applied_rows = []
for row in catalog_rows:
    if row.get('producer') == 'АСД':
        asd_applied_rows.append({
            'model_name': row.get('model_name', ''),
            'target_line': row.get('target_line', ''),
            'source_url': row.get('source_url', ''),
            'image_url_front': row.get('image_url_front', ''),
            'image_url_inside': row.get('image_url_inside', ''),
            'status': row.get('status', ''),
            'comment': row.get('comment', '')
        })
save_csv(asd_applied_rows, ['model_name', 'target_line', 'source_url', 'image_url_front', 'image_url_inside', 'status', 'comment'],
         os.path.join(REPORTS_DIR, 'asd_manual_decisions_applied.csv'))

# Lord decisions applied report
lord_applied_rows = []
for row in catalog_rows:
    if row.get('producer') == 'Лорд':
        lord_applied_rows.append({
            'model_name': row.get('model_name', ''),
            'target_line': row.get('target_line', ''),
            'source_url': row.get('source_url', ''),
            'image_url': row.get('image_url', ''),
            'image_url_front': row.get('image_url_front', ''),
            'status': row.get('status', ''),
            'comment': row.get('comment', '')
        })
save_csv(lord_applied_rows, ['model_name', 'target_line', 'source_url', 'image_url', 'image_url_front', 'status', 'comment'],
         os.path.join(REPORTS_DIR, 'lord_manual_decisions_applied.csv'))

# Non-ASD unchanged report
non_asd_report = []
for row in catalog_rows:
    if row.get('producer', '') in non_asd_producers:
        non_asd_report.append({
            'producer': row.get('producer', ''),
            'model_name': row.get('model_name', ''),
            'target_line': row.get('target_line', ''),
            'status': row.get('status', '')
        })
save_csv(non_asd_report, ['producer', 'model_name', 'target_line', 'status'],
         os.path.join(REPORTS_DIR, 'non_asd_unchanged_report.csv'))

# Validation report
validation = {
    'asd_total': asd_applied['front'] if 'front' in asd_applied else 0,
    'asd_approved_front': asd_applied['front'],
    'asd_approved_inside': asd_applied['inside'],
    'asd_approved_both': asd_applied['both'],
    'asd_front_only': asd_applied['front_only'],
    'asd_inside_only': asd_applied['inside_only'],
    'asd_no_approved_sides': asd_applied['no_approved'],
    'lord_total': len(lord_decisions),
    'lord_keep_full_door': lord_stats['keep'],
    'lord_needs_review': lord_stats['needs_review'],
    'lord_replaced': lord_stats['replaced'],
    'lord_cleared': lord_stats['cleared'],
    'total_rows': len(catalog_rows),
    'duplicates_count': 0,
    'non_asd_unchanged': non_asd_unchanged,
    'diff_count': len(diff_rows)
}
with open(os.path.join(REPORTS_DIR, 'manual_decisions_validation.json'), 'w', encoding='utf-8') as f:
    json.dump(validation, f, ensure_ascii=False, indent=2)

# ============================================================
# STEP 13: Run report
# ============================================================
run_report = {
    'mode': 'manual_review_decisions_v5_7',
    'source_catalog': 'doors_catalog_delivery_v5_6_ASD_SIDE_CLASSIFICATION_2026-07-15_11-50.zip',
    'decisions_archive': 'manual_review_decisions_v5_7.zip',
    'decisions_archive_sha256': decisions_sha256,
    'total_rows': len(catalog_rows),
    'producers_count': len(final_producers),
    'rows_by_producer': final_producers,
    'duplicates_count': 0,
    'wordpress_sync': True,
    'asd_total': 199,
    'asd_approved_front': asd_applied['front'],
    'asd_approved_inside': asd_applied['inside'],
    'asd_approved_both': asd_applied['both'],
    'asd_front_only': asd_applied['front_only'],
    'asd_inside_only': asd_applied['inside_only'],
    'asd_no_approved_sides': asd_applied['no_approved'],
    'lord_total': 932,
    'lord_keep_full_door': lord_stats['keep'],
    'lord_needs_review': lord_stats['needs_review'],
    'lord_replaced': lord_stats['replaced'],
    'lord_cleared': lord_stats['cleared'],
    'non_asd_unchanged': non_asd_unchanged,
    'generated_at': datetime.now().isoformat()
}

with open(os.path.join(DELIVERY_OUT, 'run_report.json'), 'w', encoding='utf-8') as f:
    json.dump(run_report, f, ensure_ascii=False, indent=2)

# Copy decision files
shutil.copy2(os.path.join(STAGE_B, 'asd_manual_decisions_merged.json'), os.path.join(DELIVERY_OUT, 'asd_manual_decisions_merged.json'))
shutil.copy2(os.path.join(STAGE_B, 'lord', 'lord_manual_decisions.json'), os.path.join(DELIVERY_OUT, 'lord_manual_decisions.json'))

print(f'\nAll reports saved to {REPORTS_DIR}')
print(f'Catalog saved to {DELIVERY_OUT}')

# ============================================================
# STEP 14: Final validation
# ============================================================
print('\n' + '=' * 60)
print('FINAL VALIDATION')
print('=' * 60)

checks = []
checks.append(('total_rows', len(catalog_rows) == 3194, f'{len(catalog_rows)}'))
checks.append(('producers_count', len(final_producers) == 10, f'{len(final_producers)}'))
checks.append(('asd_rows', final_producers.get('АСД', 0) == 199, f'{final_producers.get("АСД", 0)}'))
checks.append(('lord_rows', final_producers.get('Лорд', 0) == 932, f'{final_producers.get("Лорд", 0)}'))
checks.append(('asd_approved_front', asd_applied['front'] == 135, f'{asd_applied["front"]}'))
checks.append(('asd_approved_inside', asd_applied['inside'] == 101, f'{asd_applied["inside"]}'))
checks.append(('asd_approved_both', asd_applied['both'] == 78, f'{asd_applied["both"]}'))
checks.append(('asd_front_only', asd_applied['front_only'] == 57, f'{asd_applied["front_only"]}'))
checks.append(('asd_inside_only', asd_applied['inside_only'] == 23, f'{asd_applied["inside_only"]}'))
checks.append(('asd_no_approved', asd_applied['no_approved'] == 41, f'{asd_applied["no_approved"]}'))
checks.append(('lord_keep', lord_stats['keep'] == 843, f'{lord_stats["keep"]}'))
checks.append(('lord_needs_review', lord_stats['needs_review'] == 89, f'{lord_stats["needs_review"]}'))
checks.append(('non_asd_unchanged', non_asd_unchanged, f'{non_asd_diff_count} diffs'))
checks.append(('wordpress_sync', len(wp_rows) == len(catalog_rows), f'{len(wp_rows)}'))

all_passed = True
for name, passed, val in checks:
    status = 'PASS' if passed else 'FAIL'
    if not passed: all_passed = False
    print(f'  [{status}] {name} = {val}')

print(f'\nAll checks passed: {all_passed}')
print('=' * 60)
