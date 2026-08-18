import csv
import json
import os
from collections import Counter, OrderedDict

EXPECTED_PRODUCERS = [
    "Скамбио Порте", "Хаусдорс", "Олимп", "Аэлита", "Лорд",
    "Комфорт", "Двери регионов", "Термодур", "Энтро", "АСД"
]

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CATALOG = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")
OUT_DIR = os.path.join(BASE, "output", "delivery")

with open(CATALOG, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

print(f"Total rows: {len(rows)}")
print(f"Columns: {list(rows[0].keys())}")

# ========== STEP 1: Producers check ==========
all_producers = sorted(set(r['producer'] for r in rows))
print(f"\n=== STEP 1: Producers check ===")
print(f"Unique producers: {len(all_producers)}")
for p in all_producers:
    print(f"  {p}")

unexpected = [p for p in all_producers if p not in EXPECTED_PRODUCERS]
missing = [p for p in EXPECTED_PRODUCERS if p not in all_producers]
print(f"Unexpected producers: {unexpected}")
print(f"Missing producers: {missing}")

# Check if Форс/Дверь/Стальная appear in model_name/target_line
for keyword in ['Форс', 'Дверь', 'Стальная']:
    in_producer = sum(1 for r in rows if keyword.lower() in r.get('producer', '').lower())
    in_target = sum(1 for r in rows if keyword.lower() in r.get('target_line', '').lower())
    in_model = sum(1 for r in rows if keyword.lower() in r.get('model_name', '').lower())
    if in_producer or in_target or in_model:
        print(f"  '{keyword}' found in: producer={in_producer}, target_line={in_target}, model_name={in_model}")

# ========== STEP 2: Rows by producer ==========
print(f"\n=== STEP 2: Rows by producer ===")
producer_counts = Counter(r['producer'] for r in rows)
for p in EXPECTED_PRODUCERS:
    count = producer_counts.get(p, 0)
    print(f"  {p}: {count}")
print(f"  TOTAL: {sum(producer_counts.get(p, 0) for p in EXPECTED_PRODUCERS)}")
unexpected_counts = {p: c for p, c in producer_counts.items() if p not in EXPECTED_PRODUCERS}
if unexpected_counts:
    print(f"  UNEXPECTED: {unexpected_counts}")

# ========== STEP 3: No-image rows ==========
print(f"\n=== STEP 3: No-image rows ===")
no_image = []
for r in rows:
    has_img = bool(r.get('image_url', '').strip())
    has_front = bool(r.get('image_url_front', '').strip())
    has_inside = bool(r.get('image_url_inside', '').strip())
    if not has_img and not has_front:
        # Determine reason
        reason = ""
        producer = r.get('producer', '')
        source = r.get('source_url', '')
        status = r.get('status', '')
        comment = r.get('comment', '')

        if '404' in comment or 'not available' in comment.lower():
            reason = "image_404_on_manufacturer_site"
        elif '[MISSING]' in r.get('model_name', ''):
            reason = "stub_model_no_source"
        elif not source:
            reason = "no_source_url"
        elif 'needs_review' in status.lower():
            reason = "needs_manual_review"
        elif 'found' in status.lower() and not has_img:
            reason = "scraped_without_image_url"
        else:
            reason = "unknown"

        no_image.append({
            'producer': producer,
            'target_line': r.get('target_line', ''),
            'model_name': r.get('model_name', ''),
            'source_url': source,
            'image_url': r.get('image_url', ''),
            'image_url_front': r.get('image_url_front', ''),
            'image_url_inside': r.get('image_url_inside', ''),
            'status': status,
            'comment': comment,
            'reason_no_image': reason,
        })

print(f"  Rows without images: {len(no_image)}")
no_img_by_producer = Counter(r['producer'] for r in no_image)
for p, c in sorted(no_img_by_producer.items()):
    print(f"    {p}: {c}")

# Save no-image report
if no_image:
    fieldnames = ['producer', 'target_line', 'model_name', 'source_url', 'image_url', 'image_url_front', 'image_url_inside', 'status', 'comment', 'reason_no_image']
    no_img_path = os.path.join(OUT_DIR, 'no_image_rows_v5.csv')
    with open(no_img_path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=';')
        writer.writeheader()
        writer.writerows(no_image)
    print(f"  Saved: {no_img_path}")

# Group by producer for summary
print(f"\n  No-image by reason:")
reason_counts = Counter(r['reason_no_image'] for r in no_image)
for reason, count in sorted(reason_counts.items(), key=lambda x: -x[1]):
    print(f"    {reason}: {count}")
