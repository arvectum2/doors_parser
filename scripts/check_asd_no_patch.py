import csv
import os
import json
from collections import Counter

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CATALOG = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")
PATCH_PATH = os.path.join(BASE, "output", "asd_photo_patch_v2_clean", "asd_photo_patch_v2_clean.csv")

with open(CATALOG, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

with open(PATCH_PATH, encoding='utf-8-sig') as f:
    patch_rows = list(csv.DictReader(f, delimiter=';'))

patch_names = {r['model_name'] for r in patch_rows if r.get('selected_image_url', '').strip()}

asd_rows = [r for r in rows if r['producer'] == 'АСД']

# Find ASD rows without front image
no_front = [r for r in asd_rows if not r.get('image_url_front', '').strip()]
print(f"ASD rows without front image: {len(no_front)}")
for r in no_front:
    has_patch = r['model_name'] in patch_names
    print(f"  {r['model_name'][:60]}: patch={has_patch} comment={r.get('comment', '')[:50]}")

# Find ASD rows without inside image
no_inside = [r for r in asd_rows if not r.get('image_url_inside', '').strip()]
print(f"\nASD rows without inside image: {len(no_inside)}")
for r in no_inside[:10]:
    front = r.get('image_url_front', '')
    print(f"  {r['model_name'][:60]}: front={'YES' if front else 'NO'}")

# Check the 15 without patch - what's their comment?
print(f"\n15 ASD rows without patch URL:")
no_patch = [r for r in asd_rows if r['model_name'] not in patch_names]
for r in no_patch:
    print(f"  {r['model_name'][:60]}: front={bool(r.get('image_url_front', '').strip())} comment={r.get('comment', '')[:60]}")
