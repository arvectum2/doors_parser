import csv
import json
import os
import sys
from collections import Counter, defaultdict
from datetime import datetime
import hashlib

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CATALOG = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")

EXPECTED_PRODUCERS = [
    "Скамбио Порте", "Хаусдорс", "Олимп", "Аэлита", "Лорд",
    "Комфорт", "Двери регионов", "Термодур", "Энтро", "АСД"
]

with open(CATALOG, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

print(f"=== STEP 0: BASE METRICS ===")
print(f"Total rows: {len(rows)}")
print(f"Columns: {list(rows[0].keys())}")

# Unique producers
all_producers = sorted(set(r['producer'] for r in rows))
print(f"Producers ({len(all_producers)}): {all_producers}")

# Producer counts
prod_counts = Counter(r['producer'] for r in rows)
for p in EXPECTED_PRODUCERS:
    print(f"  {p}: {prod_counts.get(p, 0)}")

# Duplicates
keys = set()
dupes = []
for i, r in enumerate(rows):
    key = (r['producer'], r['target_line'], r['model_name'])
    if key in keys:
        dupes.append((i, r))
    keys.add(key)
print(f"\nDuplicates: {len(dupes)}")

# ============================================================
# STEP 1: Image analysis per row
# ============================================================
print(f"\n=== STEP 1: IMAGE ANALYSIS ===")

# Per-row analysis
row_analysis = []
for r in rows:
    image_url = r.get('image_url', '').strip()
    image_url_front = r.get('image_url_front', '').strip()
    image_url_inside = r.get('image_url_inside', '').strip()
    
    has_url = bool(image_url)
    has_front = bool(image_url_front)
    has_inside = bool(image_url_inside)
    has_any = has_url or has_front or has_inside
    
    row_analysis.append({
        'producer': r['producer'],
        'target_line': r['target_line'],
        'model_name': r['model_name'],
        'source_url': r.get('source_url', ''),
        'image_url': image_url,
        'image_url_front': image_url_front,
        'image_url_inside': image_url_inside,
        'has_url': has_url,
        'has_front': has_front,
        'has_inside': has_inside,
        'has_any': has_any,
        'comment': r.get('comment', ''),
        'status': r.get('status', ''),
    })

total = len(rows)
rows_with_any = sum(1 for ra in row_analysis if ra['has_any'])
rows_without_any = total - rows_with_any
rows_with_url = sum(1 for ra in row_analysis if ra['has_url'])
rows_with_front = sum(1 for ra in row_analysis if ra['has_front'])
rows_with_inside = sum(1 for ra in row_analysis if ra['has_inside'])
rows_with_both = sum(1 for ra in row_analysis if ra['has_front'] and ra['has_inside'])

print(f"rows_with_any_image: {rows_with_any}")
print(f"rows_without_any_image: {rows_without_any}")
print(f"rows_with_image_url: {rows_with_url}")
print(f"rows_with_image_url_front: {rows_with_front}")
print(f"rows_with_image_url_inside: {rows_with_inside}")
print(f"rows_with_both_front_inside: {rows_with_both}")

# Breakdown by domain
all_urls = set()
url_domains = Counter()
for ra in row_analysis:
    for field in ['image_url', 'image_url_front', 'image_url_inside']:
        url = ra[field.replace('image_url_front', 'image_url_front').replace('image_url_inside', 'image_url_inside')]
        if not url:
            url = ra[field] if field in ra else ''
        if url:
            all_urls.add(url)
            try:
                from urllib.parse import urlparse
                domain = urlparse(url).netloc
                url_domains[domain] += 1
            except:
                url_domains['parse_error'] += 1

print(f"\nUnique image URLs: {len(all_urls)}")
print(f"\nURLs by domain:")
for domain, count in url_domains.most_common():
    print(f"  {domain}: {count}")

# ============================================================
# STEP 2: ASD deep check
# ============================================================
print(f"\n=== STEP 2: ASD DEEP CHECK ===")
asd_rows = [ra for ra in row_analysis if ra['producer'] == 'АСД']
print(f"ASD total: {len(asd_rows)}")
asd_front = sum(1 for ra in asd_rows if ra['has_front'])
asd_inside = sum(1 for ra in asd_rows if ra['has_inside'])
asd_both = sum(1 for ra in asd_rows if ra['has_front'] and ra['has_inside'])
asd_url = sum(1 for ra in asd_rows if ra['has_url'])
asd_no_front = sum(1 for ra in asd_rows if not ra['has_front'])
asd_no_inside = sum(1 for ra in asd_rows if not ra['has_inside'])

# Check image_url != image_url_front
asd_url_mismatch = sum(1 for ra in asd_rows if ra['has_url'] and ra['has_front'] and ra['image_url'] != ra['image_url_front'])

print(f"ASD image_url filled: {asd_url}")
print(f"ASD image_url_front filled: {asd_front}")
print(f"ASD image_url_inside filled: {asd_inside}")
print(f"ASD both filled: {asd_both}")
print(f"ASD without front: {asd_no_front}")
print(f"ASD without inside: {asd_no_inside}")
print(f"ASD image_url != image_url_front: {asd_url_mismatch}")

if asd_no_front > 0 or asd_no_inside > 0:
    print(f"\n*** ASD WARNING: Missing front/inside! ***")
    for ra in asd_rows:
        if not ra['has_front'] or not ra['has_inside']:
            print(f"  {ra['model_name']}: front={ra['has_front']} inside={ra['has_inside']} url={ra['image_url'][:60] if ra['image_url'] else 'EMPTY'}")

# Check if any ASD URLs are from as-doors.ru
asd_as_doors = [ra for ra in asd_rows if 'as-doors.ru' in ra.get('image_url', '') or 'as-doors.ru' in ra.get('image_url_front', '') or 'as-doors.ru' in ra.get('image_url_inside', '')]
print(f"\nASD rows with as-doors.ru URLs: {len(asd_as_doors)}")
for ra in asd_as_doors[:5]:
    print(f"  {ra['model_name']}: url={ra['image_url'][:60]} front={ra['image_url_front'][:60] if ra['image_url_front'] else 'EMPTY'}")

# ============================================================
# STEP 3: as-doors.ru 302 audit
# ============================================================
print(f"\n=== STEP 3: AS-DOORS.RU 302 AUDIT ===")

as_doors_rows = [ra for ra in row_analysis if 'as-doors.ru' in (ra.get('image_url', '') or '') or 'as-doors.ru' in (ra.get('image_url_front', '') or '') or 'as-doors.ru' in (ra.get('image_url_inside', '') or '')]
print(f"Rows with as-doors.ru URLs: {len(as_doors_rows)}")

# Rows that have empty image_url but comment mentions 302
cleared_302 = [ra for ra in row_analysis if ra.get('comment', '') and '302' in ra.get('comment', '')]
print(f"Rows with 302 in comment: {len(cleared_302)}")
for ra in cleared_302[:5]:
    print(f"  {ra['producer']}: {ra['model_name']}: comment={ra['comment'][:60]}")

# Check which producers are affected by 302 clearing
cleared_302_by_prod = Counter(ra['producer'] for ra in cleared_302)
print(f"\n302 cleared by producer:")
for p, c in cleared_302_by_prod.most_common():
    print(f"  {p}: {c}")

# ============================================================
# STEP 4: No-image breakdown
# ============================================================
print(f"\n=== STEP 4: NO-IMAGE BREAKDOWN ===")

no_image = [ra for ra in row_analysis if not ra['has_any']]
no_img_by_prod = Counter(ra['producer'] for ra in no_image)
print(f"Total no-image rows: {len(no_image)}")
print(f"\nNo-image by producer:")
for p in EXPECTED_PRODUCERS:
    count = no_img_by_prod.get(p, 0)
    print(f"  {p}: {count}")

# Why are they no-image?
no_img_reasons = defaultdict(list)
for ra in no_image:
    comment = ra.get('comment', '')
    if '404' in comment:
        no_img_reasons['image_404'].append(ra)
    elif '302' in comment:
        no_img_reasons['image_302_cleared'].append(ra)
    elif 'cleared' in comment:
        no_img_reasons['image_cleared_other'].append(ra)
    elif 'network' in comment.lower():
        no_img_reasons['network_unreachable'].append(ra)
    elif '[MISSING]' in ra.get('model_name', ''):
        no_img_reasons['stub_model'].append(ra)
    elif not comment or comment == 'found':
        no_img_reasons['empty_or_found'].append(ra)
    else:
        no_img_reasons[f'other:{comment[:40]}'].append(ra)

print(f"\nNo-image reasons:")
for reason, items in sorted(no_img_reasons.items(), key=lambda x: -len(x[1])):
    prods = Counter(i['producer'] for i in items)
    print(f"  {reason}: {len(items)} — {dict(prods)}")

# ============================================================
# STEP 5: Lord / dveri-lord.ru network check
# ============================================================
print(f"\n=== STEP 5: LORD / DVERI-LORD.RU ===")
lord_rows = [ra for ra in row_analysis if ra['producer'] == 'Лорд']
print(f"Lord total: {len(lord_rows)}")

lord_with_url = [ra for ra in lord_rows if ra['has_url'] or ra['has_front'] or ra['has_inside']]
lord_no_image = [ra for ra in lord_rows if not ra['has_any']]
lord_dveri_lord = [ra for ra in lord_rows if 'dveri-lord.ru' in (ra.get('image_url', '') or '') or 'dveri-lord.ru' in (ra.get('image_url_front', '') or '')]

print(f"Lord with any image: {len(lord_with_url)}")
print(f"Lord without any image: {len(lord_no_image)}")
print(f"Lord with dveri-lord.ru URLs: {len(lord_dveri_lord)}")

# Lord domains
lord_domains = Counter()
for ra in lord_rows:
    for field in ['image_url', 'image_url_front', 'image_url_inside']:
        url = ra.get(field, '')
        if url:
            try:
                from urllib.parse import urlparse
                domain = urlparse(url).netloc
                lord_domains[domain] += 1
            except:
                pass

print(f"\nLord URL domains:")
for domain, count in lord_domains.most_common():
    print(f"  {domain}: {count}")

# Lord empty image rows
print(f"\nLord rows without any image:")
for ra in lord_no_image:
    print(f"  {ra['model_name'][:50]}: url={ra['image_url'][:50] if ra['image_url'] else 'EMPTY'} front={ra['image_url_front'][:50] if ra['image_url_front'] else 'EMPTY'} comment={ra.get('comment', '')[:50]}")
