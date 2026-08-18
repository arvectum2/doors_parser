import csv
import os
import json
from collections import Counter

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CATALOG = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")
V4_PATH = os.path.join(BASE, "output", "final_rework_ready_v4", "doors_catalog.csv")
PATCH_PATH = os.path.join(BASE, "output", "asd_photo_patch_v2_clean", "asd_photo_patch_v2_clean.csv")

with open(CATALOG, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
with open(V4_PATH, encoding='utf-8-sig') as f:
    v4_rows = list(csv.DictReader(f, delimiter=';'))
with open(PATCH_PATH, encoding='utf-8-sig') as f:
    patch_rows = list(csv.DictReader(f, delimiter=';'))

v4_asd = {r['model_name']: r for r in v4_rows if r['producer'] == 'АСД'}
patch_map = {r['model_name']: r for r in patch_rows}

asd_rows = [r for r in rows if r['producer'] == 'АСД']
no_front = [r for r in asd_rows if not r.get('image_url_front', '').strip()]

print(f"ASD rows without front: {len(no_front)}")

# For each, show v4 URL and patch info
for r in no_front:
    name = r['model_name']
    v4r = v4_asd.get(name, {})
    v4_front = v4r.get('image_url_front', '')
    v4_inside = v4r.get('image_url_inside', '')
    patch_info = patch_map.get(name, {})
    patch_selected = patch_info.get('selected_image_url', '')
    patch_decision = patch_info.get('decision', '')
    patch_reason = patch_info.get('reason', '')
    
    print(f"\n{name[:60]}")
    print(f"  V4 front: {v4_front[:70] if v4_front else 'EMPTY'}")
    print(f"  V4 inside: {v4_inside[:70] if v4_inside else 'EMPTY'}")
    print(f"  Patch decision: {patch_decision}")
    print(f"  Patch reason: {patch_reason[:60]}")
    print(f"  Patch selected: {patch_selected[:70] if patch_selected else 'NONE'}")
