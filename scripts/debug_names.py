import csv, json

ASD = '\u0410\u0421\u0414'
with open('output/final_rework_ready_v5/doors_catalog.csv', encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
asd_rows = [r for r in rows if r['producer'] == ASD]

print("Catalog ASD model_name samples:")
for r in asd_rows[:15]:
    print(f'  [{r["model_name"]}]')

with open('output/delivery_v5_2/wayback_inside_all_results.json') as f:
    wb = json.load(f)
wb_keys = list(wb.keys())
print("\nWayback keys samples:")
for k in wb_keys[:5]:
    print(f'  [{k}]')

# Matching
print("\nMatching attempts:")
for r in asd_rows[:20]:
    cn = r['model_name'].strip().lower()
    found = None
    for wk in wb_keys:
        if cn in wk.lower() and 'заказная' in wk.lower():
            found = wk
            break
    status = f'MATCH: {found[:60]}' if found else 'NO MATCH'
    print(f'  {r["model_name"][:40]:40s} -> {status}')
