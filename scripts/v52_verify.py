"""Comprehensive verify v5.2 catalog"""
import csv, json, os
from collections import Counter, defaultdict

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CAT = os.path.join(BASE, "output", "delivery_v5_2", "doors_catalog.csv")

with open(CAT, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
fn = list(rows[0].keys())

print("=" * 60)
print("V5.2 VERIFICATION")
print("=" * 60)

# 1. Row count
print(f"\nTotal rows: {len(rows)}")

# 2. Producer breakdown
prods = Counter(r['producer'] for r in rows)
print(f"\nProducers ({len(prods)}):")
for p, c in sorted(prods.items(), key=lambda x: -x[1]):
    print(f"  {p}: {c}")

# 3. ASD detail
ASD_KEY = "\u0410\u0421\u0414"
asd = [r for r in rows if r['producer'] == ASD_KEY]
f_ok = sum(1 for r in asd if r.get('image_url_front','').strip())
i_ok = sum(1 for r in asd if r.get('image_url_inside','').strip())
b_ok = sum(1 for r in asd if r.get('image_url_front','').strip() and r.get('image_url_inside','').strip())
print(f"\nASD: {len(asd)} rows")
print(f"  Front: {f_ok}/199")
print(f"  Inside: {i_ok}/199")
print(f"  Both: {b_ok}/199")

# List missing
for r in asd:
    has_f = bool(r.get('image_url_front','').strip())
    has_i = bool(r.get('image_url_inside','').strip())
    if not has_f or not has_i:
        status = []
        if not has_f: status.append('NO_FRONT')
        if not has_i: status.append('NO_INSIDE')
        print(f"  MISSING {','.join(status)}: {r['model_name'][:60]}")

# 4. Duplicates
keys = Counter((r['producer'], r['target_line'], r['model_name']) for r in rows)
dupes = {k: v for k, v in keys.items() if v > 1}
print(f"\nDuplicates: {len(dupes)}")
for d, c in dupes.items():
    print(f"  x{c}: {d}")

# 5. Empty critical fields
critical = ['producer', 'target_line', 'model_name']
missing = defaultdict(int)
for r in rows:
    for f in critical:
        if not r.get(f, '').strip():
            missing[f] += 1
print(f"\nEmpty critical fields:")
for f, c in missing.items():
    print(f"  {f}: {c}")

# 6. URL consistency - check http:// vs https://
http_front = sum(1 for r in rows if r.get('image_url_front','').strip().startswith('http://'))
https_front = sum(1 for r in rows if r.get('image_url_front','').strip().startswith('https://'))
http_inside = sum(1 for r in rows if r.get('image_url_inside','').strip().startswith('http://'))
https_inside = sum(1 for r in rows if r.get('image_url_inside','').strip().startswith('https://'))
print(f"\nURL scheme:")
print(f"  Front: http={http_front}, https={https_front}")
print(f"  Inside: http={http_inside}, https={https_inside}")

# 7. Source URL count
src = sum(1 for r in rows if r.get('source_url','').strip())
desc = sum(1 for r in rows if r.get('description','').strip())
print(f"\nOther:")
print(f"  With source_url: {src}/{len(rows)}")
print(f"  With description: {desc}/{len(rows)}")

# 8. Compare with v5.1
print(f"\nV5.1 baseline check:")
cat_v5_1 = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")
with open(cat_v5_1, encoding='utf-8-sig') as f:
    v5_1 = list(csv.DictReader(f, delimiter=';'))
print(f"  v5.1 rows: {len(v5_1)}")
print(f"  v5.2 rows: {len(rows)}")
print(f"  Net change: {len(rows) - len(v5_1)}")

# Check which rows changed for ASD
asd_v5_1 = [r for r in v5_1 if r['producer'] == ASD_KEY]
asd_v5_2 = asd
new_front = 0
new_inside = 0
for r2 in asd_v5_2:
    name = r2['model_name']
    r1 = next((r for r in asd_v5_1 if r['model_name'] == name), None)
    if r1:
        if r2.get('image_url_front','').strip() and not r1.get('image_url_front','').strip():
            new_front += 1
        if r2.get('image_url_inside','').strip() and not r1.get('image_url_inside','').strip():
            new_inside += 1

print(f"\n  ASD improvements vs v5.1:")
print(f"    New front URLs: {new_front}")
print(f"    New inside URLs: {new_inside}")

print("\n" + "=" * 60)
print("VERIFICATION COMPLETE")
print("=" * 60)
