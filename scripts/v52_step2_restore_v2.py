import csv, json, os

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CAT = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")
PATCH_CSV = os.path.join(BASE, "output", "asd_photo_patch_v2_clean", "asd_photo_patch_v2_clean.csv")
WB_JSON = os.path.join(BASE, "output", "delivery_v5_2", "wayback_inside_all_results.json")
OUT = os.path.join(BASE, "output", "delivery_v5_2", "doors_catalog.csv")

ASD_KEY = "\u0410\u0421\u0414"  # АСД

# NOTE: ALL open() calls MUST use encoding='utf-8' because default on this system is cp1251

with open(CAT, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
fn = list(rows[0].keys())
print(f"V5.1 rows: {len(rows)}")

# Patch front URLs
patch_front = {}
with open(PATCH_CSV, encoding='utf-8-sig') as f:
    for r in csv.DictReader(f, delimiter=';'):
        url = r.get('selected_image_url', '').strip()
        if url:
            patch_front[r['model_name']] = url.replace('https://', 'http://')
print(f"Patch front URLs: {len(patch_front)}")

# Wayback inside URLs - CRITICAL: use encoding='utf-8'
with open(WB_JSON, encoding='utf-8') as f:
    wb_data = json.load(f)
wb_inside = {k: v['inside_url'].replace('https://', 'http://')
             for k, v in wb_data.items() if v.get('inside_url')}
print(f"Wayback inside URLs: {len(wb_inside)}")

# Restore
front_restored = 0
inside_restored = 0
front_ok = 0
inside_ok = 0

for r in rows:
    if r['producer'] != ASD_KEY:
        continue
    name = r['model_name']

    if r.get('image_url_front', '').strip():
        front_ok += 1
    elif patch_front.get(name):
        r['image_url_front'] = patch_front[name]
        if not r.get('image_url', '').strip():
            r['image_url'] = patch_front[name]
        front_restored += 1

    if r.get('image_url_inside', '').strip():
        inside_ok += 1
    elif wb_inside.get(name):
        r['image_url_inside'] = wb_inside[name]
        inside_restored += 1

# Stats
asd_rows = [r for r in rows if r['producer'] == ASD_KEY]
f_total = sum(1 for r in asd_rows if r.get('image_url_front', '').strip())
i_total = sum(1 for r in asd_rows if r.get('image_url_inside', '').strip())
b_total = sum(1 for r in asd_rows if r.get('image_url_front', '').strip() and r.get('image_url_inside', '').strip())

print(f"\nASD: {len(asd_rows)} rows")
print(f"Front: {f_total}/199 (already had: {front_ok}, restored: {front_restored})")
print(f"Inside: {i_total}/199 (already had: {inside_ok}, restored: {inside_restored})")
print(f"Both: {b_total}/199")

# Save catalog
with open(OUT, 'w', encoding='utf-8-sig', newline='') as f:
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writeheader()
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writerows(rows)
print(f"Saved: {OUT}")
