import json, csv

ASD_KEY = '\u0410\u0421\u0414'

# Load catalog
with open('output/final_rework_ready_v5/doors_catalog.csv', encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
asd = [r for r in rows if r['producer'] == ASD_KEY]
cat_names = [r['model_name'] for r in asd]

# Load fixed wayback JSON
with open('output/delivery_v5_2/wayback_inside_all_results_fixed.json') as f:
    wb = json.load(f)

# Check how wb_inside looks
wb_inside = {k: v['inside_url'] for k, v in wb.items() if v.get('inside_url')}
print(f'wb_inside keys count: {len(wb_inside)}')

# Check match
matched = 0
unmatched = []
for cn in cat_names:
    if cn in wb_inside:
        matched += 1
    else:
        unmatched.append(cn)

print(f'Direct match: {matched}/199')

# Check one unmatched
if unmatched:
    u = unmatched[0]
    print(f'\nFirst unmatched: {repr(u[:80])}')
    print(f'Unmatched hex: {u.encode("utf-8").hex(":")}')
    
    # Try to find it in wb_inside keys
    for wbk in wb_inside.keys():
        if u[:30] == wbk[:30] or u[-30:] == wbk[-30:]:
            print(f'  Found by prefix/suffix: {repr(wbk[:80])}')
            print(f'  WB key hex: {wbk.encode("utf-8").hex(":")}')
            break
    else:
        # Try with the first wb_inside key
        wbk = list(wb_inside.keys())[0]
        print(f'\nFirst wb_inside key: {repr(wbk[:80])}')
        print(f'WB key hex: {wbk.encode("utf-8").hex(":")}')
        print(f'Cat[0] hex: {cat_names[0].encode("utf-8").hex(":")}')
        print(f'Cat[0]==WB: {cat_names[0] == wbk}')

# Check if any catalog name matches any wb_inside key at all
print('\nChecking first 5 catalog names:')
for cn in cat_names[:5]:
    in_wb = cn in wb_inside
    in_full = cn in wb
    print(f'  {cn[:40]}... -> in_wb_inside: {in_wb}, in_wb_full: {in_full}')
