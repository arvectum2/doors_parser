import json, csv, unicodedata

ASD = '\u0410\u0421\u0414'
with open('output/final_rework_ready_v5/doors_catalog.csv', encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
asd = [r for r in rows if r['producer'] == ASD]
cat_names = [r['model_name'] for r in asd]

with open('output/delivery_v5_2/wayback_inside_all_results.json') as f:
    wb = json.load(f)

# Check each catalog name against wb keys
matched = 0
unmatched = []
for cn in cat_names:
    if cn in wb:
        matched += 1
    else:
        unmatched.append(cn)

print(f'Direct match: {matched}/199')
print(f'Unmatched: {len(unmatched)}')
if unmatched:
    cn = unmatched[0]
    print(f'First unmatched: {repr(cn[:60])}')
    h = cn.encode('utf-8').hex(':')
    print(f'Its hex: {h}')
    # Look for it in wb by different methods
    for wbk in wb.keys():
        if unicodedata.normalize('NFC', cn) == unicodedata.normalize('NFC', wbk):
            print('  Found via NFC normalize!')
            print(f'  WB hex: {wbk.encode("utf-8").hex(":")}')
            break
    else:
        # Try case-insensitive
        cn_lower = cn.lower()
        for wbk in wb.keys():
            if cn_lower == wbk.lower():
                print('  Found via case-insensitive!')
                print(f'  WB hex: {wbk.encode("utf-8").hex(":")}')
                break
        else:
            # Try finding by key content
            for wbk in wb.keys():
                if len(wbk) > 10 and wbk[:20] == cn[:20]:
                    print('  Found via first 20 chars!')
                    print(f'  WB hex: {wbk.encode("utf-8").hex(":")}')
                    print(f'  Cat hex: {cn[:20].encode("utf-8").hex(":")}')
                    break
            else:
                # Look for 'Avant' substring
                for wbk in wb.keys():
                    if 'Avant' in wbk:
                        print('  Found Avant in wb keys:')
                        print(f'  WB: {repr(wbk[:80])}')
                        print(f'  Cat: {repr(cn[:80])}')
                        print(f'  WB hex: {wbk.encode("utf-8").hex(":")}')
                        print(f'  Cat hex: {cn.encode("utf-8").hex(":")}')
                        break
