import csv
import os
import urllib.request
import urllib.error
import ssl
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import Counter
import re
import json

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
v4_path = os.path.join(BASE, "output", "final_rework_ready_v4", "doors_catalog.csv")
v5_path = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")

with open(v4_path, encoding='utf-8-sig') as f:
    v4_rows = list(csv.DictReader(f, delimiter=';'))
with open(v5_path, encoding='utf-8-sig') as f:
    v5_rows = list(csv.DictReader(f, delimiter=';'))

v4_asd = {r['model_name']: r for r in v4_rows if r['producer'] == 'АСД'}
v5_asd = {r['model_name']: r for r in v5_rows if r['producer'] == 'АСД'}

# Find ASD rows that lost front image
lost = []
for name, v5r in v5_asd.items():
    v4r = v4_asd.get(name)
    if not v4r:
        continue
    if v4r.get('image_url_front', '').strip() and not v5r.get('image_url_front', '').strip():
        lost.append((name, v4r, v5r))

print(f"ASD rows that lost front image: {len(lost)}")

def check_url(url, timeout=10):
    """Try both HTTP and HTTPS"""
    # Try HTTP first (works better for as-doors.ru)
    http_url = url.replace('https://', 'http://')
    try:
        req = urllib.request.Request(http_url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
        resp = opener.open(req, timeout=timeout)
        ct = resp.headers.get('Content-Type', '')
        data = resp.read(4)
        is_img = data[:3] == b'\xff\xd8\xff' or data[:4] == b'\x89PNG' or data[:4] == b'RIFF'
        if is_img:
            return http_url, True, ct, None
    except:
        pass
    
    # Try HTTPS
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
        resp = opener.open(req, timeout=timeout)
        ct = resp.headers.get('Content-Type', '')
        data = resp.read(4)
        is_img = data[:3] == b'\xff\xd8\xff' or data[:4] == b'\x89PNG' or data[:4] == b'RIFF'
        if is_img:
            return url, True, ct, None
    except:
        pass
    
    return url, False, '', 'both_failed'

# Test all lost ASD front URLs via HTTP
print(f"\nTesting {len(lost)} lost URLs via HTTP...")
results = []
start = time.time()

with ThreadPoolExecutor(max_workers=20) as executor:
    futures = {}
    for name, v4r, v5r in lost:
        old_url = v4r.get('image_url_front', '')
        if old_url:
            futures[executor.submit(check_url, old_url, 10)] = (name, v4r, v5r, old_url)
    
    done = 0
    for future in as_completed(futures):
        done += 1
        name, v4r, v5r, old_url = futures[future]
        new_url, ok, ct, err = future.result()
        results.append({
            'model_name': name,
            'old_url': old_url,
            'new_url': new_url if ok else '',
            'working': ok,
            'content_type': ct,
            'error': err,
        })
        if done % 20 == 0:
            print(f"  {done}/{len(lost)} ({time.time()-start:.1f}s)")

working = sum(1 for r in results if r['working'])
print(f"\nResults ({time.time()-start:.1f}s):")
print(f"  Working via HTTP: {working}/{len(results)}")
print(f"  Failed: {len(results) - working}")

# Show working ones
print(f"\nWorking URLs:")
for r in results:
    if r['working']:
        print(f"  {r['model_name'][:40]}: {r['new_url'][-60:]}")

# Save results
with open(os.path.join(BASE, "output", "delivery", "asd_http_test_results.json"), 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print(f"\nSaved results to asd_http_test_results.json")
