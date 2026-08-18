import csv
import os
import json
from collections import Counter

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CATALOG = os.path.join(BASE, "output", "delivery_v5_2", "doors_catalog.csv")

with open(CATALOG, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

asd = [r for r in rows if r['producer'] == 'АСД']
asd_front = sum(1 for r in asd if r.get('image_url_front', '').strip())
asd_inside = sum(1 for r in asd if r.get('image_url_inside', '').strip())
asd_both = sum(1 for r in asd if r.get('image_url_front', '').strip() and r.get('image_url_inside', '').strip())
asd_no_front = [r['model_name'] for r in asd if not r.get('image_url_front', '').strip()]
asd_no_inside = [r['model_name'] for r in asd if not r.get('image_url_inside', '').strip()]

print(f"Total rows: {len(rows)}")
print(f"ASD total: {len(asd)}")
print(f"ASD front: {asd_front}/199")
print(f"ASD inside: {asd_inside}/199")
print(f"ASD both: {asd_both}/199")
print(f"\nASD missing front ({len(asd_no_front)}):")
for name in asd_no_front:
    print(f"  {name[:60]}")
print(f"\nASD missing inside ({len(asd_no_inside)}):")
for name in asd_no_inside:
    print(f"  {name[:60]}")

# Producers
producers = Counter(r['producer'] for r in rows)
print(f"\nProducers:")
for p, c in sorted(producers.items()):
    print(f"  {p}: {c}")

# Duplicates
keys = set()
dupes = 0
for r in rows:
    key = (r['producer'], r['target_line'], r['model_name'])
    if key in keys:
        dupes += 1
    keys.add(key)
print(f"\nDuplicates: {dupes}")

# No-image
no_image = [r for r in rows if not r.get('image_url', '').strip() and not r.get('image_url_front', '').strip() and not r.get('image_url_inside', '').strip()]
print(f"No-image rows: {len(no_image)}")
