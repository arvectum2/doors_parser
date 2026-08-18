import csv
import os
import json
import urllib.request
import urllib.error
import ssl
from collections import Counter

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
v4_path = os.path.join(BASE, "output", "final_rework_ready_v4", "doors_catalog.csv")
v5_path = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")
patch_path = os.path.join(BASE, "output", "asd_photo_patch_v2_clean", "asd_photo_patch_v2_clean.csv")

with open(v4_path, encoding='utf-8-sig') as f:
    v4_rows = list(csv.DictReader(f, delimiter=';'))
with open(v5_path, encoding='utf-8-sig') as f:
    v5_rows = list(csv.DictReader(f, delimiter=';'))
with open(patch_path, encoding='utf-8-sig') as f:
    patch_rows = list(csv.DictReader(f, delimiter=';'))

v4_asd = {r['model_name']: r for r in v4_rows if r['producer'] == 'АСД'}
v5_asd = {r['model_name']: r for r in v5_rows if r['producer'] == 'АСД'}
patch_map = {r['model_name']: r for r in patch_rows}

# Check v4 inside images
v4_inside_count = sum(1 for name, r in v4_asd.items() if r.get('image_url_inside', '').strip())
print(f"V4 ASD inside images: {v4_inside_count}/199")

# Check which v4 inside URLs work
v4_inside_urls = {}
for name, r in v4_asd.items():
    inside = r.get('image_url_inside', '').strip()
    if inside:
        v4_inside_urls[name] = inside

print(f"\nV4 ASD inside URLs: {len(v4_inside_urls)}")

# Test a few v4 inside URLs
def check_url(url, timeout=10):
    http_url = url.replace('https://', 'http://')
    for try_url in [http_url, url]:
        try:
            req = urllib.request.Request(try_url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
            resp = opener.open(req, timeout=timeout)
            ct = resp.headers.get('Content-Type', '')
            data = resp.read(4)
            is_img = data[:3] == b'\xff\xd8\xff' or data[:4] == b'\x89PNG' or data[:4] == b'RIFF'
            if is_img:
                return try_url, True, ct
        except:
            pass
    return url, False, ''

print(f"\nTesting first 10 v4 inside URLs...")
inside_working = 0
for name, url in list(v4_inside_urls.items())[:10]:
    _, ok, ct = check_url(url)
    if ok:
        inside_working += 1
        print(f"  OK: {name[:40]}: {url[-60:]}")
    else:
        print(f"  FAIL: {name[:40]}: {url[-60:]}")

print(f"\nInside working: {inside_working}/10")

# Summary
asd_v5_with_front = sum(1 for r in v5_rows if r['producer'] == 'АСД' and r.get('image_url_front', '').strip())
asd_v5_with_inside = sum(1 for r in v5_rows if r['producer'] == 'АСД' and r.get('image_url_inside', '').strip())
print(f"\nV5 ASD current state:")
print(f"  Front: {asd_v5_with_front}/199")
print(f"  Inside: {asd_v5_with_inside}/199")

# Check patch coverage
patch_covered = sum(1 for name in v5_asd if name in patch_map)
print(f"  Patch covers: {patch_covered}/199")
