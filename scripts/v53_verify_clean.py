"""V5.3: Post-cleanup verification and fix review"""
import csv, json, os
from collections import Counter

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CAT = os.path.join(BASE, "output", "delivery_v5_3", "doors_catalog.csv")
QA_JSON = os.path.join(BASE, "output", "delivery_v5_2", "asd_final_photo_source_validation.json")
OUT = os.path.join(BASE, "output", "delivery_v5_2")

ASD_KEY = "\u0410\u0421\u0414"

with open(CAT, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
asd = [r for r in rows if r['producer'] == ASD_KEY]

with open(QA_JSON, encoding='utf-8') as f:
    qa_rows = json.load(f)
qa_by_name = {r['model_name']: r for r in qa_rows}

print("=" * 60)
print("V5.3 POST-CLEANUP VERIFICATION")
print("=" * 60)

# 1. Count actual web.archive.org in main columns
arch_inside = 0
arch_front = 0
for r in asd:
    if 'web.archive.org' in (r.get('image_url_inside','') or ''):
        arch_inside += 1
        print(f"  ARCHIVE INSIDE: {r['model_name'][:50]}")
    if 'web.archive.org' in (r.get('image_url_front','') or ''):
        arch_front += 1
        print(f"  ARCHIVE FRONT: {r['model_name'][:50]}")
print(f"\n  web.archive.org in image_url_inside: {arch_inside}")
print(f"  web.archive.org in image_url_front: {arch_front}")
print(f"  TOTAL archive URLs in main columns: {arch_front + arch_inside}")

# 2. Check all producers
prods = Counter(r['producer'] for r in rows)
print(f"\n  Producers: {len(prods)}")
for p, c in sorted(prods.items(), key=lambda x: -x[1]):
    print(f"    {p}: {c}")

# 3. ASD stats
f_filled = sum(1 for r in asd if r.get('image_url_front','').strip())
i_filled = sum(1 for r in asd if r.get('image_url_inside','').strip())
b_filled = sum(1 for r in asd if r.get('image_url_front','').strip() and r.get('image_url_inside','').strip())
print(f"\n  ASD front: {f_filled}/199")
print(f"  ASD inside: {i_filled}/199")
print(f"  ASD both: {b_filled}/199")

# 4. Verify all non-ASD preserved
print(f"\n  Total rows: {len(rows)}")
# Check specific producers
komfort = sum(1 for r in rows if r['producer'] == '\u041a\u043e\u043c\u0444\u043e\u0440\u0442')
hausdoors = sum(1 for r in rows if r['producer'] == '\u0425\u0430\u0443\u0441\u0434\u043e\u0440\u0441')
olimp = sum(1 for r in rows if r['producer'] == '\u041e\u043b\u0438\u043c\u043f')
entro = sum(1 for r in rows if r['producer'] == '\u042d\u043d\u0442\u0440\u043e')
print(f"  Komfort: {komfort}")
print(f"  Hausdoors: {hausdoors}")
print(f"  Olimp: {olimp}")
print(f"  Entro: {entro}")

# 5. Check image_url = image_url_front
wp_ok = all(
    r.get('image_url_front','').strip() and r['image_url'] == r['image_url_front']
    for r in asd if r.get('image_url_front','').strip()
)
print(f"  WordPress sync: {'OK' if wp_ok else 'FAIL'}")

# 6. Duplicates
keys = Counter((r['producer'], r['target_line'], r['model_name']) for r in rows)
dupes = {k:v for k,v in keys.items() if v > 1}
print(f"  Duplicates: {len(dupes)}")
for d, c in dupes.items():
    print(f"    x{c}: {d}")

# 7. Which models lost front?
for q in qa_rows:
    name = q['model_name']
    old_f = q.get('image_url_front', '')
    r = next((x for x in asd if x['model_name'] == name), None)
    if r:
        new_f = r.get('image_url_front', '').strip()
        if old_f and not new_f:
            print(f"  LOST FRONT: {name[:60]} | was: {old_f[:60]}")

# 8. Which models lost inside (beyond wayback removal)?
wb_removed = 0
pattern_invalid_removed = 0
for q in qa_rows:
    name = q['model_name']
    old_i = q.get('image_url_inside', '')
    r = next((x for x in asd if x['model_name'] == name), None)
    if r:
        new_i = r.get('image_url_inside', '').strip()
        if old_i and not new_i:
            src = q.get('inside_source_type', '')
            if src == 'wayback':
                wb_removed += 1
            else:
                pattern_invalid_removed += 1
                print(f"  REMOVED non-wayback inside [{name[:50]}] src={src} old={old_i[:60]}")

print(f"\n  Inside removed (wayback): {wb_removed}")
print(f"  Inside removed (other, e.g. invalid/technical): {pattern_invalid_removed}")

# 9. Still invalid?
for r in asd:
    for col in ('image_url_front', 'image_url_inside'):
        url = r.get(col, '').strip()
        if url and 'web.archive.org' in url:
            print(f"  WARNING: {col} has archive URL: {r['model_name'][:50]}")
