import csv
import os

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"

# Read v4 catalog
v4_path = os.path.join(BASE, "output", "final_rework_ready_v4", "doors_catalog.csv")
v5_path = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")

with open(v4_path, encoding='utf-8-sig') as f:
    v4_rows = list(csv.DictReader(f, delimiter=';'))

with open(v5_path, encoding='utf-8-sig') as f:
    v5_rows = list(csv.DictReader(f, delimiter=';'))

print(f"V4 rows: {len(v4_rows)}")
print(f"V5 rows: {len(v5_rows)}")

# ASD in v4
v4_asd = [r for r in v4_rows if r['producer'] == 'АСД']
print(f"\nV4 ASD rows: {len(v4_asd)}")

# ASD in v5
v5_asd = [r for r in v5_rows if r['producer'] == 'АСД']
print(f"V5 ASD rows: {len(v5_asd)}")

# Compare
v4_asd_keys = {(r['producer'], r['target_line'], r['model_name']): r for r in v4_asd}
v5_asd_keys = {(r['producer'], r['target_line'], r['model_name']): r for r in v5_asd}

# Check which ASD rows lost images
lost_front = 0
lost_inside = 0
lost_both = 0
still_have = 0
new_in_v5 = 0

for key, v4r in v4_asd_keys.items():
    v5r = v5_asd_keys.get(key)
    if not v5r:
        new_in_v5 += 1
        continue
    
    v4_has_front = bool(v4r.get('image_url_front', '').strip())
    v4_has_inside = bool(v4r.get('image_url_inside', '').strip())
    v5_has_front = bool(v5r.get('image_url_front', '').strip())
    v5_has_inside = bool(v5r.get('image_url_inside', '').strip())
    
    if v4_has_front and not v5_has_front:
        lost_front += 1
    if v4_has_inside and not v5_has_inside:
        lost_inside += 1
    if (v4_has_front or v4_has_inside) and not v5_has_front and not v5_has_inside:
        lost_both += 1
    if v5_has_front and v5_has_inside:
        still_have += 1

print(f"\nASD comparison v4 to v5:")
print(f"  Still have both front+inside: {still_have}")
print(f"  Lost front: {lost_front}")
print(f"  Lost inside: {lost_inside}")
print(f"  Lost both: {lost_both}")
print(f"  New in v5: {new_in_v5}")

# Show first 10 ASD rows from v4 with images
print(f"\nV4 ASD rows with images (first 10):")
count = 0
for key, v4r in v4_asd_keys.items():
    if v4r.get('image_url_front', '').strip() and count < 10:
        v5r = v5_asd_keys.get(key)
        v5_front = v5r.get('image_url_front', '') if v5r else 'NOT IN V5'
        print(f"  {v4r['model_name'][:50]}")
        print(f"    V4 front: {v4r['image_url_front'][:80]}")
        print(f"    V5 front: {v5_front[:80] if v5_front else 'EMPTY'}")
        count += 1

# Show ASD rows that lost images
print(f"\nASD rows that lost front (first 10):")
count = 0
for key, v4r in v4_asd_keys.items():
    v5r = v5_asd_keys.get(key)
    if not v5r:
        continue
    if v4r.get('image_url_front', '').strip() and not v5r.get('image_url_front', '').strip():
        if count < 10:
            print(f"  {v4r['model_name'][:50]}")
            print(f"    V4 front: {v4r['image_url_front'][:80]}")
            print(f"    V5 front: EMPTY")
            count += 1
