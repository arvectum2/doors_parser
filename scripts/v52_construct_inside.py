import csv
import os
import json
import urllib.request
import urllib.error
import ssl
import re
from collections import Counter

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE
opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CATALOG = os.path.join(BASE, "output", "delivery_v5_2", "doors_catalog.csv")
V4_CATALOG = os.path.join(BASE, "output", "final_rework_ready_v4", "doors_catalog.csv")

with open(CATALOG, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
fieldnames = list(rows[0].keys())

with open(V4_CATALOG, encoding='utf-8-sig') as f:
    v4_rows = list(csv.DictReader(f, delimiter=';'))
v4_asd = {r['model_name']: r for r in v4_rows if r['producer'] == 'АСД'}

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

# Find ASD rows missing inside
asd_indices = [(i, r) for i, r in enumerate(rows) if r['producer'] == 'АСД']
missing_inside = [(i, r) for i, r in asd_indices if not r.get('image_url_inside', '').strip()]
print(f"Missing inside: {len(missing_inside)}")

inside_found = 0
for idx, r in missing_inside:
    name = r['model_name']
    front = r.get('image_url_front', '').strip()
    v4r = v4_asd.get(name, {})
    v4_inside = v4r.get('image_url_inside', '')
    
    if not front:
        continue
    
    # Try constructing inside from front URL by changing filename number
    front_match = re.search(r'/([^/]+?)(?:-(\d+))?\.(jpg|png|webp)$', front, re.I)
    if front_match:
        base_name = front_match.group(1)
        ext = front_match.group(3)
        dir_path = front[:front.rfind('/') + 1]
        
        for num in range(1, 20):
            test_url = f"{dir_path}{base_name}-{num}.{ext}"
            if test_url != front and check_image_url(test_url):
                rows[idx]['image_url_inside'] = test_url
                inside_found += 1
                break

print(f"Inside found by pattern: {inside_found}")

# Save
with open(CATALOG, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=';')
    writer.writeheader()
    writer.writerows(rows)

# Final stats
asd = [r for r in rows if r['producer'] == 'АСД']
asd_front = sum(1 for r in asd if r.get('image_url_front', '').strip())
asd_inside = sum(1 for r in asd if r.get('image_url_inside', '').strip())
asd_both = sum(1 for r in asd if r.get('image_url_front', '').strip() and r.get('image_url_inside', '').strip())
print(f"\nFinal: front={asd_front}/199, inside={asd_inside}/199, both={asd_both}/199")

asd_no_inside = [r['model_name'] for r in asd if not r.get('image_url_inside', '').strip()]
print(f"Still missing inside ({len(asd_no_inside)}):")
for name in asd_no_inside:
    print(f"  {name[:60]}")
