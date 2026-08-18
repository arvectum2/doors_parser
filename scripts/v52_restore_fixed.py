import json, csv

# Fix the corrupted JSON keys (UTF-8 -> CP1251 -> UTF-8 mojibake)
with open('output/delivery_v5_2/wayback_inside_all_results.json') as f:
    wb = json.load(f)

def fix_mojibake(s):
    try:
        return s.encode('cp1251').decode('utf-8')
    except:
        return s

wb_fixed = {}
for orig_k, v in wb.items():
    fixed_k = fix_mojibake(orig_k)
    wb_fixed[fixed_k] = v

# Save fixed version
with open('output/delivery_v5_2/wayback_inside_all_results_fixed.json', 'w', encoding='utf-8') as f:
    json.dump(wb_fixed, f, ensure_ascii=False, indent=2)

# Count working inside URLs
working = sum(1 for v in wb_fixed.values() if v.get('inside_url'))
print(f'Fixed JSON: {len(wb_fixed)} entries, {working} with inside URLs')

# Now restore
CAT = 'output/final_rework_ready_v5/doors_catalog.csv'
PATCH_CSV = 'output/asd_photo_patch_v2_clean/asd_photo_patch_v2_clean.csv'
WB_JSON = 'output/delivery_v5_2/wayback_inside_all_results_fixed.json'
OUT = 'output/delivery_v5_2/doors_catalog.csv'

ASD_KEY = '\u0410\u0421\u0414'

with open(CAT, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
fn = list(rows[0].keys())
print(f'V5.1 rows: {len(rows)}')

# Patch front
patch_front = {}
with open(PATCH_CSV, encoding='utf-8-sig') as f:
    for r in csv.DictReader(f, delimiter=';'):
        url = r.get('selected_image_url', '').strip()
        if url:
            patch_front[r['model_name']] = url.replace('https://', 'http://')
print(f'Patch front URLs: {len(patch_front)}')

# Wayback inside (fixed)
with open(WB_JSON) as f:
    wb_data = json.load(f)
wb_inside = {k: v['inside_url'].replace('https://', 'http://') 
             for k, v in wb_data.items() if v.get('inside_url')}
print(f'Wayback inside URLs: {len(wb_inside)}')

# Restore
front_restored = 0
inside_restored = 0
front_missing = 0
inside_missing = 0

for r in rows:
    if r['producer'] != ASD_KEY:
        continue
    name = r['model_name']
    
    if not r.get('image_url_front', '').strip() and patch_front.get(name):
        r['image_url_front'] = patch_front[name]
        if not r.get('image_url', '').strip():
            r['image_url'] = patch_front[name]
        front_restored += 1
    elif not r.get('image_url_front', '').strip():
        front_missing += 1
    
    if not r.get('image_url_inside', '').strip() and wb_inside.get(name):
        r['image_url_inside'] = wb_inside[name]
        inside_restored += 1
    elif not r.get('image_url_inside', '').strip():
        inside_missing += 1

# Stats
asd_rows = [r for r in rows if r['producer'] == ASD_KEY]
f_ok = sum(1 for r in asd_rows if r.get('image_url_front', '').strip())
i_ok = sum(1 for r in asd_rows if r.get('image_url_inside', '').strip())
b_ok = sum(1 for r in asd_rows if r.get('image_url_front', '').strip() and r.get('image_url_inside', '').strip())

print(f'\nASD: {len(asd_rows)} rows')
print(f'Front: {f_ok}/199 (restored: {front_restored}, missing: {front_missing})')
print(f'Inside: {i_ok}/199 (restored: {inside_restored}, missing: {inside_missing})')
print(f'Both: {b_ok}/199')

# Save
with open(OUT, 'w', encoding='utf-8-sig', newline='') as f:
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writeheader()
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writerows(rows)
print(f'Saved: {OUT}')
