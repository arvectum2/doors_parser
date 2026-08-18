import csv
import os
import json
import urllib.request
import urllib.error
import ssl
import time
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import Counter

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE
opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
V52_CATALOG = os.path.join(BASE, "output", "delivery_v5_2", "doors_catalog.csv")
V4_CATALOG = os.path.join(BASE, "output", "final_rework_ready_v4", "doors_catalog.csv")
PATCH_PATH = os.path.join(BASE, "output", "asd_photo_patch_v2_clean", "asd_photo_patch_v2_clean.csv")

with open(V52_CATALOG, encoding='utf-8-sig') as f:
    v52_rows = list(csv.DictReader(f, delimiter=';'))
with open(V4_CATALOG, encoding='utf-8-sig') as f:
    v4_rows = list(csv.DictReader(f, delimiter=';'))
with open(PATCH_PATH, encoding='utf-8-sig') as f:
    patch_rows = list(csv.DictReader(f, delimiter=';'))

v4_asd = {r['model_name']: r for r in v4_rows if r['producer'] == 'АСД'}
patch_map = {r['model_name']: r for r in patch_rows}

# Find ASD rows still missing front or inside
asd_v52 = [(i, r) for i, r in enumerate(v52_rows) if r['producer'] == 'АСД']
missing_front = [(i, r) for i, r in asd_v52 if not r.get('image_url_front', '').strip()]
missing_inside = [(i, r) for i, r in asd_v52 if not r.get('image_url_inside', '').strip()]

print(f"ASD missing front: {len(missing_front)}")
print(f"ASD missing inside: {len(missing_inside)}")

# For missing front: try to find from v4 inside URLs (inside URL pattern often mirrors front)
def check_image_url(url, timeout=10):
    for try_url in [url.replace('https://', 'http://'), url]:
        try:
            req = urllib.request.Request(try_url, headers={'User-Agent': 'Mozilla/5.0'})
            resp = opener.open(req, timeout=timeout)
            ct = resp.headers.get('Content-Type', '')
            data = resp.read(4)
            is_img = data[:3] == b'\xff\xd8\xff' or data[:4] == b'\x89PNG' or data[:4] == b'RIFF'
            if is_img:
                return True
        except:
            pass
    return False

# For missing front: try v4 front URL with different filename patterns
print(f"\n=== TRYING TO FIND MISSING FRONT ===")
for idx, r in missing_front:
    name = r['model_name']
    v4r = v4_asd.get(name, {})
    v4_front = v4r.get('image_url_front', '')
    source = r.get('source_url', '')
    
    print(f"\n  {name[:50]}")
    print(f"    V4 front: {v4_front[:60]}")
    
    # Try v4 front URL directly
    if v4_front:
        http_url = v4_front.replace('https://', 'http://')
        if check_image_url(http_url):
            print(f"    *** V4 FRONT WORKS: {http_url[:60]}")
            v52_rows[idx]['image_url_front'] = http_url
            if not v52_rows[idx].get('image_url', '').strip():
                v52_rows[idx]['image_url'] = http_url
            continue
    
    # Try extracting image path from v4 front and constructing alternative URLs
    if v4_front:
        # Pattern: /doors/zakaznyie/xxx.html/img/year/model/filename.jpg
        # Try: /img/year/model/filename.jpg (without page prefix)
        match = re.search(r'/(?:img|images)/(\d{4}/[^?#]+)', v4_front)
        if match:
            img_path = match.group(1)
            for prefix in ['https://as-doors.ru/img/', 'http://as-doors.ru/img/', 'https://as-doors.ru/images/', 'http://as-doors.ru/images/']:
                test_url = prefix + img_path
                if check_image_url(test_url):
                    print(f"    *** CONSTRUCTED URL WORKS: {test_url[:60]}")
                    v52_rows[idx]['image_url_front'] = test_url
                    if not v52_rows[idx].get('image_url', '').strip():
                        v52_rows[idx]['image_url'] = test_url
                    break

# For missing inside: try to construct from front URL pattern
print(f"\n=== TRYING TO FIND MISSING INSIDE ===")
inside_found_extra = 0
for idx, r in missing_inside:
    name = r['model_name']
    front = r.get('image_url_front', '').strip()
    v4r = v4_asd.get(name, {})
    v4_inside = v4r.get('image_url_inside', '')
    
    if not front:
        continue
    
    print(f"\n  {name[:50]}")
    
    # Try v4 inside URL directly
    if v4_inside:
        http_url = v4_inside.replace('https://', 'http://')
        if check_image_url(http_url):
            print(f"    *** V4 INSIDE WORKS: {http_url[:60]}")
            v52_rows[idx]['image_url_inside'] = http_url
            inside_found_extra += 1
            continue
    
    # Try constructing inside from front URL by changing filename
    # Front URL pattern: /img/year/model/filename-XX.jpg
    # Inside might be: /img/year/model/filename-YY.jpg (different number)
    front_match = re.search(r'/([^/]+?)(?:-\d+)?\.(jpg|png|webp)$', front, re.I)
    if front_match:
        base_name = front_match.group(1)
        ext = front_match.group(2)
        dir_path = front[:front.rfind('/') + 1]
        
        # Try different numbers for inside
        for num in range(1, 20):
            test_url = f"{dir_path}{base_name}-{num}.{ext}"
            if test_url != front and check_image_url(test_url):
                print(f"    *** CONSTRUCTED INSIDE: {test_url[:60]}")
                v52_rows[idx]['image_url_inside'] = test_url
                inside_found_extra += 1
                break

print(f"\nExtra inside found: {inside_found_extra}")

# Save updated catalog
fieldnames = list(v52_rows[0].keys())
with open(V52_CATALOG, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=';')
    writer.writeheader()
    writer.writerows(v52_rows)

# Final stats
asd_final = [r for r in v52_rows if r['producer'] == 'АСД']
asd_front = sum(1 for r in asd_final if r.get('image_url_front', '').strip())
asd_inside = sum(1 for r in asd_final if r.get('image_url_inside', '').strip())
asd_both = sum(1 for r in asd_final if r.get('image_url_front', '').strip() and r.get('image_url_inside', '').strip())
print(f"\nFinal ASD: front={asd_front}/199, inside={asd_inside}/199, both={asd_both}/199")

# List remaining missing
remaining_missing_front = [r['model_name'] for r in asd_final if not r.get('image_url_front', '').strip()]
remaining_missing_inside = [r['model_name'] for r in asd_final if not r.get('image_url_inside', '').strip()]
print(f"\nRemaining missing front ({len(remaining_missing_front)}):")
for name in remaining_missing_front:
    print(f"  {name[:60]}")
print(f"\nRemaining missing inside ({len(remaining_missing_inside)}):")
for name in remaining_missing_inside:
    print(f"  {name[:60]}")
