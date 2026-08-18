import csv
import os
from urllib.parse import urlparse
from collections import Counter

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
v5_path = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")
v4_path = os.path.join(BASE, "output", "final_rework_ready_v4", "doors_catalog.csv")

with open(v5_path, encoding='utf-8-sig') as f:
    v5_rows = list(csv.DictReader(f, delimiter=';'))
with open(v4_path, encoding='utf-8-sig') as f:
    v4_rows = list(csv.DictReader(f, delimiter=';'))

v5_asd = {r['model_name']: r for r in v5_rows if r['producer'] == 'АСД'}
v4_asd = {r['model_name']: r for r in v4_rows if r['producer'] == 'АСД'}

# Check v5 inside images
v5_inside = [(name, r) for name, r in v5_asd.items() if r.get('image_url_inside', '').strip()]
print(f"V5 ASD inside images: {len(v5_inside)}")

# Check domains
inside_domains = Counter()
for name, r in v5_inside:
    url = r.get('image_url_inside', '')
    parsed = urlparse(url)
    inside_domains[parsed.netloc] += 1

print(f"Inside URL domains: {dict(inside_domains)}")

# Check if these match v4 inside URLs
match_v4 = 0
mismatch = 0
for name, r in v5_inside:
    v4r = v4_asd.get(name)
    if v4r:
        v4_inside = v4r.get('image_url_inside', '').strip()
        v5_inside_url = r.get('image_url_inside', '').strip()
        if v4_inside == v5_inside_url:
            match_v4 += 1
        else:
            mismatch += 1

print(f"\nV5 inside == V4 inside: {match_v4}")
print(f"V5 inside != V4 inside: {mismatch}")

# Show some mismatched ones
print(f"\nMismatched inside URLs (first 5):")
count = 0
for name, r in v5_inside:
    v4r = v4_asd.get(name)
    if v4r:
        v4_inside = v4r.get('image_url_inside', '').strip()
        v5_inside_url = r.get('image_url_inside', '').strip()
        if v4_inside != v5_inside_url:
            if count < 5:
                print(f"  {name[:40]}")
                print(f"    V4: {v4_inside[:70]}")
                print(f"    V5: {v5_inside_url[:70]}")
                count += 1

# Check which v5 inside URLs are from as-doors.ru
asd_inside = [(name, r) for name, r in v5_inside if 'as-doors.ru' in r.get('image_url_inside', '')]
print(f"\nV5 ASD inside from as-doors.ru: {len(asd_inside)}")

# Check if these are old or new format
old_format_inside = sum(1 for name, r in asd_inside if 'assets/cache_image' not in r.get('image_url_inside', ''))
new_format_inside = sum(1 for name, r in asd_inside if 'assets/cache_image' in r.get('image_url_inside', ''))
print(f"  Old format: {old_format_inside}")
print(f"  New format: {new_format_inside}")
