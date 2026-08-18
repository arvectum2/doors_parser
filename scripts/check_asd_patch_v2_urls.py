import csv
import os
import urllib.request
import urllib.error
import ssl
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
import json

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
patch_path = os.path.join(BASE, "output", "asd_photo_patch_v2_clean", "asd_photo_patch_v2_clean.csv")

with open(patch_path, encoding='utf-8-sig') as f:
    patch_rows = list(csv.DictReader(f, delimiter=';'))

print(f"Patch rows: {len(patch_rows)}")

# Get unique selected_image_url
selected_urls = {}
for r in patch_rows:
    model = r['model_name']
    selected = r.get('selected_image_url', '').strip()
    old = r.get('old_image_url', '').strip()
    if selected:
        selected_urls[model] = {'old': old, 'new': selected, 'decision': r.get('decision', '')}

print(f"Models with selected URL: {len(selected_urls)}")

def check_url(url, timeout=10):
    # Try HTTP first
    http_url = url.replace('https://', 'http://')
    for try_url in [http_url, url]:
        try:
            req = urllib.request.Request(try_url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
            resp = opener.open(req, timeout=timeout)
            ct = resp.headers.get('Content-Type', '')
            data = resp.read(4)
            is_img = data[:3] == b'\xff\xd8\xff' or data[:4] == b'\x89PNG' or data[:4] == b'RIFF'
            if is_img:
                return try_url, True, ct, None
        except:
            pass
    return url, False, '', 'both_failed'

# Test all selected URLs
print(f"\nTesting {len(selected_urls)} selected URLs...")
results = []
start = time.time()

with ThreadPoolExecutor(max_workers=15) as executor:
    futures = {}
    for model, info in selected_urls.items():
        futures[executor.submit(check_url, info['new'], 10)] = (model, info)
    
    done = 0
    for future in as_completed(futures):
        done += 1
        model, info = futures[future]
        url, ok, ct, err = future.result()
        results.append({
            'model_name': model,
            'old_url': info['old'],
            'selected_url': info['new'],
            'working': ok,
            'content_type': ct,
            'error': err,
        })

working = sum(1 for r in results if r['working'])
print(f"\nResults ({time.time()-start:.1f}s):")
print(f"  Working: {working}/{len(results)}")
print(f"  Failed: {len(results) - working}")

# Show working ones
print(f"\nWorking URLs:")
for r in results:
    if r['working']:
        print(f"  {r['model_name'][:40]}: {r['selected_url'][-60:]}")

# Show failed ones
print(f"\nFailed URLs:")
for r in results:
    if not r['working']:
        print(f"  {r['model_name'][:40]}: {r['selected_url'][-60:]}")

# Save results
with open(os.path.join(BASE, "output", "delivery", "asd_patch_v2_url_check.json"), 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
