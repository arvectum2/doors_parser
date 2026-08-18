import csv
import os
import json
import urllib.request
import urllib.error
import ssl
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import Counter

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
V4_CATALOG = os.path.join(BASE, "output", "final_rework_ready_v4", "doors_catalog.csv")

with open(V4_CATALOG, encoding='utf-8-sig') as f:
    v4_rows = list(csv.DictReader(f, delimiter=';'))

v4_asd = [r for r in v4_rows if r['producer'] == 'АСД']
print(f"V4 ASD: {len(v4_asd)}")

# Collect all unique inside URLs
inside_urls = {}
for r in v4_asd:
    inside = r.get('image_url_inside', '').strip()
    if inside:
        inside_urls[r['model_name']] = inside

print(f"V4 ASD inside URLs: {len(inside_urls)}")

def check_url(url, timeout=10):
    """Try both HTTP and HTTPS, return working URL or None"""
    for try_url in [url.replace('https://', 'http://'), url]:
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

# Test ALL v4 inside URLs
print(f"\nTesting all {len(inside_urls)} v4 inside URLs...")
results = {}
start = time.time()

with ThreadPoolExecutor(max_workers=20) as executor:
    futures = {}
    for name, url in inside_urls.items():
        futures[executor.submit(check_url, url, 10)] = (name, url)
    
    done = 0
    for future in as_completed(futures):
        done += 1
        name, original_url = futures[future]
        url, ok, ct = future.result()
        results[name] = {'working': ok, 'url': url if ok else '', 'original': original_url}
        if done % 30 == 0:
            print(f"  {done}/{len(inside_urls)} ({time.time()-start:.1f}s)")

working = sum(1 for r in results.values() if r['working'])
print(f"\nResults ({time.time()-start:.1f}s):")
print(f"  Working: {working}/{len(results)}")
print(f"  Failed: {len(results) - working}")

# Categorize results
working_urls = {name: r['url'] for name, r in results.items() if r['working']}
failed_urls = {name: r['original'] for name, r in results.items() if not r['working']}

print(f"\nWorking inside URLs (first 10):")
for name, url in list(working_urls.items())[:10]:
    print(f"  {name[:40]}: {url[-60:]}")

print(f"\nFailed inside URLs (first 10):")
for name, url in list(failed_urls.items())[:10]:
    print(f"  {name[:40]}: {url[-60:]}")

# Save results
with open(os.path.join(BASE, "output", "delivery_v5_2", "v4_inside_url_check.json"), 'w', encoding='utf-8') as f:
    json.dump({
        'working': working_urls,
        'failed': failed_urls,
        'stats': {'working': working, 'failed': len(failed_urls), 'total': len(inside_urls)}
    }, f, ensure_ascii=False, indent=2)

print(f"\nSaved v4_inside_url_check.json")
