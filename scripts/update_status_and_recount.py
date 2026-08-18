import csv
import json
import os
from collections import Counter

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CATALOG = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")
OUT_DIR = os.path.join(BASE, "output", "delivery")

with open(CATALOG, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

# Update status for rows with cleared images
cleared_count = 0
for r in rows:
    has_img = bool(r.get('image_url', '').strip())
    has_front = bool(r.get('image_url_front', '').strip())
    if not has_img and not has_front:
        if r.get('status', '') == 'found':
            r['status'] = 'needs_review'
        if not r.get('comment', ''):
            r['comment'] = 'image_url_cleared_broken_or_unavailable'

# Also update Olimp Paris status
for r in rows:
    if r['producer'] == 'Олимп' and 'Париж' in r.get('target_line', ''):
        r['status'] = 'needs_review'
        r['comment'] = 'image_404_on_manufacturer_site_olymp-doors.ru.tilda.ws'

# Save updated CSV
fieldnames = list(rows[0].keys())
with open(CATALOG, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=';')
    writer.writeheader()
    writer.writerows(rows)

# Recount
no_image = sum(1 for r in rows if not r.get('image_url', '').strip() and not r.get('image_url_front', '').strip())
has_image = len(rows) - no_image
needs_review = sum(1 for r in rows if r.get('status', '') == 'needs_review')

print(f"Total rows: {len(rows)}")
print(f"Rows with images: {has_image}")
print(f"Rows without images: {no_image}")
print(f"Rows with needs_review status: {needs_review}")

# Count unique image URLs
all_urls = set()
for r in rows:
    for field in ['image_url', 'image_url_front', 'image_url_inside']:
        url = r.get(field, '').strip()
        if url:
            all_urls.add(url)
print(f"Unique image URLs: {len(all_urls)}")

# No-image breakdown by producer
no_img_rows = [r for r in rows if not r.get('image_url', '').strip() and not r.get('image_url_front', '').strip()]
no_img_by_prod = Counter(r['producer'] for r in no_img_rows)
print(f"\nNo-image by producer:")
for p, c in sorted(no_img_by_prod.items()):
    print(f"  {p}: {c}")

# Update no_image_rows report
no_img_details = []
for r in no_img_rows:
    reason = ""
    comment = r.get('comment', '')
    if '404' in comment: reason = "image_404_on_manufacturer_site"
    elif 'unavailable' in comment: reason = "image_url_unavailable"
    elif 'cleared' in comment: reason = "image_url_cleared_broken"
    elif '[MISSING]' in r.get('model_name', ''): reason = "stub_model_no_source"
    else: reason = "unknown"
    
    no_img_details.append({
        'producer': r['producer'],
        'target_line': r['target_line'],
        'model_name': r['model_name'],
        'source_url': r.get('source_url', ''),
        'image_url': r.get('image_url', ''),
        'image_url_front': r.get('image_url_front', ''),
        'image_url_inside': r.get('image_url_inside', ''),
        'status': r.get('status', ''),
        'comment': r.get('comment', ''),
        'reason_no_image': reason,
    })

# Save no-image report
no_img_path = os.path.join(OUT_DIR, 'no_image_rows_v5.csv')
if no_img_details:
    fieldnames_ni = ['producer', 'target_line', 'model_name', 'source_url', 'image_url', 'image_url_front', 'image_url_inside', 'status', 'comment', 'reason_no_image']
    with open(no_img_path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames_ni, delimiter=';')
        writer.writeheader()
        writer.writerows(no_img_details)

reason_counts = Counter(r['reason_no_image'] for r in no_img_details)
print(f"\nNo-image by reason:")
for reason, count in sorted(reason_counts.items(), key=lambda x: -x[1]):
    print(f"  {reason}: {count}")
