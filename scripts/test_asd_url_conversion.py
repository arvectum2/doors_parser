import csv
import os
import urllib.request
import urllib.error
import ssl
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import Counter
import re

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
v4_path = os.path.join(BASE, "output", "final_rework_ready_v4", "doors_catalog.csv")

with open(v4_path, encoding='utf-8-sig') as f:
    v4_rows = list(csv.DictReader(f, delimiter=';'))

v4_asd = [r for r in v4_rows if r['producer'] == 'АСД']

# Separate old vs new format
old_format = []
new_format = []
for r in v4_asd:
    url = r.get('image_url_front', '')
    if 'assets/cache_image' in url:
        new_format.append(r)
    elif url:
        old_format.append(r)

print(f"Old format: {len(old_format)}")
print(f"New format: {len(new_format)}")

# Analyze old format URLs to find pattern
# Example: https://as-doors.ru/doors/zakaznyie/forest.html/img/2025/forest/forest-16.jpg
# The part after .html is the image path
# Try converting to: https://as-doors.ru/assets/cache_image/img/2025/forest/forest-16.jpg

# Test first 20 old format URLs with new prefix
test_old = old_format[:20]
print(f"\nTesting old->new URL conversion...")

def try_new_url(old_url):
    # Extract image path after .html
    match = re.search(r'\.html(/.+)', old_url)
    if match:
        img_path = match.group(1)
        new_url = f"https://as-doors.ru/assets/cache_image{img_path}"
        return new_url
    # Try onstock pattern
    match = re.search(r'onstock/[^/]+\.html(/.+)', old_url)
    if match:
        img_path = match.group(1)
        new_url = f"https://as-doors.ru/assets/cache_image{img_path}"
        return new_url
    return None

def check_url(url, timeout=10):
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
        resp = opener.open(req, timeout=timeout)
        ct = resp.headers.get('Content-Type', '')
        data = resp.read(4)
        is_img = data[:3] == b'\xff\xd8\xff' or data[:4] == b'\x89PNG' or data[:4] == b'RIFF'
        return url, resp.getcode(), ct, is_img, None
    except Exception as e:
        return url, 0, '', False, str(e)[:80]

working = 0
for r in test_old:
    old_url = r.get('image_url_front', '')
    new_url = try_new_url(old_url)
    if new_url:
        url, status, ct, is_img, err = check_url(new_url)
        if is_img:
            working += 1
            print(f"  WORKING: {old_url[-50:]}")
            print(f"    -> {new_url[-50:]}")
        else:
            print(f"  FAIL: {old_url[-50:]} -> {err[:40]}")
    else:
        print(f"  NO CONVERSION: {old_url[-50:]}")

print(f"\nWorking conversions: {working}/{len(test_old)}")

# If conversion works, test all 160
if working > 0:
    print(f"\nTesting all {len(old_format)} old URLs with conversion...")
    all_working = 0
    all_fail = 0
    
    with ThreadPoolExecutor(max_workers=20) as executor:
        futures = {}
        for r in old_format:
            old_url = r.get('image_url_front', '')
            new_url = try_new_url(old_url)
            if new_url:
                futures[executor.submit(check_url, new_url, 10)] = (r['model_name'], old_url, new_url)
        
        for future in as_completed(futures):
            name, old_url, new_url = futures[future]
            url, status, ct, is_img, err = future.result()
            if is_img:
                all_working += 1
            else:
                all_fail += 1
    
    print(f"  All working: {all_working}/{len(old_format)}")
    print(f"  All failed: {all_fail}/{len(old_format)}")
