import csv
from collections import Counter

with open('output/final_rework_ready_v5/doors_catalog.csv', encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

hd = [r for r in rows if 'Хаус' in r['producer'] or 'аусд' in r['producer']]
print(f"Hausdoors total: {len(hd)}")

lines = Counter(r['target_line'] for r in hd)
for l, c in sorted(lines.items()):
    print(f"  {l}: {c}")

pvc = ['Фактурные линии', 'Новая классика', 'Минимализм', 'Современная классика']
print("\nPVC collections:")
for p in pvc:
    models = [r for r in hd if r['target_line'] == p]
    print(f"\n  {p}: {len(models)}")
    for m in models:
        print(f"    {m['model_name']} | {m['coating']} | {m['material']}")

# Check images for Hausdoors PVC
print("\nHausdoors PVC images:")
for p in pvc:
    models = [r for r in hd if r['target_line'] == p]
    front = sum(1 for m in models if m.get('image_url_front', ''))
    print(f"  {p}: front={front}/{len(models)}")

# Check all 10 producers present
all_producers = set(r['producer'] for r in rows)
print(f"\nAll producers ({len(all_producers)}):")
for p in sorted(all_producers):
    print(f"  {p}: {sum(1 for r in rows if r['producer'] == p)}")
