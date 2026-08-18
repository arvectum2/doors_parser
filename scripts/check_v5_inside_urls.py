import csv
import os
import urllib.request
import urllib.error
import ssl
from collections import Counter

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
v5_path = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")

with open(v5_path, encoding='utf-8-sig') as f:
    v5_rows = list(csv.DictReader(f, delimiter=';'))

v5_asd = [r for r in v5_rows if r['producer'] == 'АСД']
v5_inside = [(r['model_name'], r.get('image_url_inside', '')) for r in v5_asd if r.get('image_url_inside', '').strip()]

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
                return True, ct
        except:
            pass
    return False, ''

# Test all 40 inside URLs
print(f"Testing {len(v5_inside)} v5 inside URLs...")
ok = 0
fail = 0
for name, url in v5_inside:
    working, ct = check_url(url)
    if working:
        ok += 1
    else:
        fail += 1
        print(f"  FAIL: {name[:40]}: {url[-60:]}")

print(f"\nResults: {ok} working, {fail} failed out of {len(v5_inside)}")

# Now check v4 inside URLs that are NOT in v5 (i.e., 159 missing ones)
v4_path = os.path.join(BASE, "output", "final_rework_ready_v4", "doors_catalog.csv")
with open(v4_path, encoding='utf-8-sig') as f:
    v4_rows = list(csv.DictReader(f, delimiter=';'))

v4_asd = {r['model_name']: r for r in v4_rows if r['producer'] == 'АСД'}
v5_asd_names = {r['model_name'] for r in v5_asd}

# Missing inside URLs from v4
missing_inside = []
for name, r in v4_asd.items():
    if name not in v5_asd_names:
        continue
    v5r = next((x for x in v5_asd if x['model_name'] == name), None)
    if v5r and not v5r.get('image_url_inside', '').strip():
        inside = r.get('image_url_inside', '').strip()
        if inside:
            missing_inside.append((name, inside))

print(f"\nV4 inside URLs missing from v5: {len(missing_inside)}")

# Test first 10 missing
print(f"\nTesting first 10 missing inside URLs...")
missing_ok = 0
for name, url in missing_inside[:10]:
    working, ct = check_url(url)
    if working:
        missing_ok += 1
        print(f"  OK: {name[:40]}: {url[-60:]}")
    else:
        print(f"  FAIL: {name[:40]}: {url[-60:]}")

print(f"Missing inside working: {missing_ok}/10")
