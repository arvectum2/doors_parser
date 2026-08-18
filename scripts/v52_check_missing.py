import csv, json, os

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CAT = os.path.join(BASE, "output", "delivery_v5_2", "doors_catalog.csv")
PATCH_CSV = os.path.join(BASE, "output", "asd_photo_patch_v2_clean", "asd_photo_patch_v2_clean.csv")

ASD_KEY = "\u0410\u0421\u0414"

with open(CAT, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
asd = [r for r in rows if r['producer'] == ASD_KEY]

# 6 without inside
no_inside = [r for r in asd if not r.get('image_url_inside','').strip()]
print("6 ASD doors WITHOUT inside image:")
for r in no_inside:
    f_url = r.get('image_url_front','')[:60] if r.get('image_url_front') else 'NONE'
    print(f"  [{r['model_name'][:50]:50s}] front={f_url}")

# 15 without front
no_front = [r for r in asd if not r.get('image_url_front','').strip()]
print(f"\n15 ASD doors WITHOUT front image:")
for r in no_front:
    i_url = r.get('image_url_inside','')[:60] if r.get('image_url_inside') else 'NONE'
    print(f"  [{r['model_name'][:50]:50s}] inside={i_url}")

# Check patch CSV for missing front
print("\nChecking patch CSV for missing front URLs...")
patch = {}
with open(PATCH_CSV, encoding='utf-8-sig') as f:
    for r in csv.DictReader(f, delimiter=';'):
        url = r.get('selected_image_url','').strip()
        if url:
            patch[r['model_name']] = url

for r in no_front:
    if r['model_name'] in patch:
        print(f"  PATCH AVAILABLE for: {r['model_name'][:50]}")
    else:
        print(f"  NO PATCH for: {r['model_name'][:50]}")
