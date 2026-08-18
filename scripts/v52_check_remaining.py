import csv
import os
from collections import Counter

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CATALOG = os.path.join(BASE, "output", "delivery_v5_2", "doors_catalog.csv")

with open(CATALOG, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

asd = [r for r in rows if r['producer'] == 'АСД']
missing_inside = [r for r in asd if not r.get('image_url_inside', '').strip()]

print(f"ASD missing inside: {len(missing_inside)}")
for r in missing_inside:
    front = r.get('image_url_front', '').strip()
    print(f"  {r['model_name'][:50]}: front={'YES' if front else 'NO'}")

# Check: for models with front but no inside, can we use front as inside?
with_front_no_inside = [r for r in missing_inside if r.get('image_url_front', '').strip()]
print(f"\nWith front but no inside: {len(with_front_no_inside)}")
for r in with_front_no_inside:
    print(f"  {r['model_name'][:50]}: front={r['image_url_front'][:60]}")
