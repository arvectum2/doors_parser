import json, csv

ASD = '\u0410\u0421\u0414'
with open('output/final_rework_ready_v5/doors_catalog.csv', encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
asd = [r for r in rows if r['producer'] == ASD]

with open('output/delivery_v5_2/wayback_inside_all_results.json') as f:
    wb = json.load(f)

cat_name = asd[1]['model_name']  # Avant
print('Cat name repr:', repr(cat_name))
print('Cat name hex:', cat_name.encode('utf-8').hex(':'))

print('In full wb:', cat_name in wb)
if cat_name in wb:
    print('WB value:', wb[cat_name])

wb_inside = {k: v.get('inside_url', '') for k, v in wb.items() if v.get('inside_url')}
print('wb_inside keys:', len(wb_inside))
wbk0 = list(wb_inside.keys())[0]
print('First wb_inside key repr:', repr(wbk0))
print('== check:', cat_name == wbk0)

# Find Avant
found = False
for wbk in wb_inside.keys():
    if 'Avant' in wbk:
        print('Found Avant in wb_inside:', repr(wbk[:80]))
        print('  Matches ==:', cat_name == wbk)
        print('  Cat hex:', cat_name.encode('utf-8').hex(':'))
        print('  WB  hex:', wbk.encode('utf-8').hex(':'))
        found = True
        break
if not found:
    print('Avant NOT found in wb_inside keys')
    print('Searching all wb keys for Avant...')
    for k in wb.keys():
        if 'Avant' in k:
            print('  Found in full wb:', repr(k[:80]))
            print('  Has inside_url:', bool(wb[k].get('inside_url')))
