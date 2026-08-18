import csv, json, os, re

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CAT = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")
PATCH_CSV = os.path.join(BASE, "output", "asd_photo_patch_v2_clean", "asd_photo_patch_v2_clean.csv")
WB_JSON = os.path.join(BASE, "output", "delivery_v5_2", "wayback_inside_all_results.json")
OUT = os.path.join(BASE, "output", "delivery_v5_2", "doors_catalog.csv")

ASD_KEY = "\u0410\u0421\u0414"  # АСД

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

# Wayback inside URLs
with open(WB_JSON) as f:
    wb_inside = {k: v.get('inside_url', '').replace('https://', 'http://') 
                 for k, v in json.load(f).items() if v.get('inside_url')}
print(f"Wayback inside URLs: {len(wb_inside)}")

# Restore
front_ok = 0
front_restored = 0
front_still = 0
inside_ok = 0
inside_restored = 0
inside_still = 0

for r in rows:
    if r['producer'] != ASD_KEY:
        continue
    
    name = r['model_name']
    
    # Front
    if r.get('image_url_front', '').strip():
        front_ok += 1
    elif patch_front.get(name):
        r['image_url_front'] = patch_front[name]
        r['image_url'] = r['image_url'] or patch_front[name]
        front_restored += 1
    else:
        front_still += 1
    
    # Inside
    if r.get('image_url_inside', '').strip():
        inside_ok += 1
    elif wb_inside.get(name):
        r['image_url_inside'] = wb_inside[name]
        inside_restored += 1
    else:
        inside_still += 1

total_asd = sum(1 for r in rows if r['producer'] == ASD_KEY)
both = sum(1 for r in rows if r['producer'] == ASD_KEY and r.get('image_url_front','').strip() and r.get('image_url_inside','').strip())

print(f"\nASD: total={total_asd}, both={both}/199")
print(f"Front: kept={front_ok}, restored={front_restored}, missing={front_still}")
print(f"Inside: kept={inside_ok}, restored={inside_restored}, missing={inside_still}")
print(f"Final front: {front_ok+front_restored}/199")
print(f"Final inside: {inside_ok+inside_restored}/199")

# Save
with open(OUT, 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.DictWriter(f, fieldnames=fn, delimiter=';')
    w.writeheader()
    w.writerows(rows)
print(f"\nSaved: {OUT}")
