import csv
from collections import Counter

with open('output/final_rework_ready_v5/doors_catalog.csv', encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

print("Total rows:", len(rows))
producers = Counter(r['producer'] for r in rows)
for p, c in sorted(producers.items()):
    print(f"  {p}: {c}")

# Duplicate check
dupes = []
keys = set()
for r in rows:
    key = (r['producer'], r['target_line'], r['model_name'])
    if key in keys:
        dupes.append(key)
    keys.add(key)
print(f"\nDuplicates: {len(dupes)}")
for d in dupes[:10]:
    print(f"  {d}")

# Find producer names
print("\n--- Hausdoors ---")
hd_names = set(r['producer'] for r in rows if 'ausdor' in r['producer'].lower() or 'усдор' in r['producer'] or 'Хаус' in r['producer'])
print("Names found:", hd_names)
for hn in hd_names:
    hd = [r for r in rows if r['producer'] == hn]
    print(f"  Total: {len(hd)}")
    hd_lines = Counter(r['target_line'] for r in hd)
    for line, cnt in sorted(hd_lines.items()):
        print(f"    {line}: {cnt}")

print("\n--- Komfort ---")
kf_names = set(r['producer'] for r in rows if 'омфорт' in r['producer'] or 'KOMFORT' in r['producer'].upper())
print("Names found:", kf_names)
for kn in kf_names:
    kf = [r for r in rows if r['producer'] == kn]
    print(f"  Total: {len(kf)}")
    kf_lines = Counter(r['target_line'] for r in kf)
    for line, cnt in sorted(kf_lines.items()):
        print(f"    {line}: {cnt}")

print("\n--- Olimp ---")
ol_names = set(r['producer'] for r in rows if 'лимп' in r['producer'] or 'OLIMP' in r['producer'].upper())
print("Names found:", ol_names)
for on in ol_names:
    ol = [r for r in rows if r['producer'] == on]
    print(f"  Total: {len(ol)}")
    ol_lines = Counter(r['target_line'] for r in ol)
    for line, cnt in sorted(ol_lines.items()):
        print(f"    {line}: {cnt}")
    paris = [r for r in ol if 'ариж' in r['target_line'] or 'Paris' in r['target_line']]
    print(f"  Paris models: {len(paris)}")
    for r in paris:
        print(f"    {r['model_name']} | {r['target_line']}")

print("\n--- ASD ---")
asd_names = set(r['producer'] for r in rows if r['producer'] == 'АСД' or 'asd' in r['producer'].lower() or 'СД' in r['producer'])
print("Names found:", asd_names)
for an in asd_names:
    asd = [r for r in rows if r['producer'] == an]
    asd_front = sum(1 for r in asd if r.get('image_url_front', ''))
    asd_inside = sum(1 for r in asd if r.get('image_url_inside', ''))
    print(f"  Total: {len(asd)}, front: {asd_front}, inside: {asd_inside}")

# Image URL validation summary
print("\n--- Image URL stats ---")
has_front = sum(1 for r in rows if r.get('image_url_front', ''))
has_inside = sum(1 for r in rows if r.get('image_url_inside', ''))
has_url = sum(1 for r in rows if r.get('image_url', ''))
print(f"  image_url: {has_url}/{len(rows)}")
print(f"  image_url_front: {has_front}/{len(rows)}")
print(f"  image_url_inside: {has_inside}/{len(rows)}")

# Check for [MISSING] or stub markers
stubs = [r for r in rows if '[MISSING]' in r.get('model_name', '') or '[MISSING]' in r.get('target_line', '')]
print(f"\n--- Stubs/MISSING: {len(stubs)} ---")
for s in stubs:
    print(f"  {s['producer']} | {s['target_line']} | {s['model_name']}")
