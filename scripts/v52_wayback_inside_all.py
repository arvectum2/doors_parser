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
V4_CATALOG = os.path.join(BASE, "output", "final_rework_ready_v4", "doors_catalog.csv")
V51_CATALOG = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")

with open(V4_CATALOG, encoding='utf-8-sig') as f:
    v4_rows = list(csv.DictReader(f, delimiter=';'))
with open(V51_CATALOG, encoding='utf-8-sig') as f:
    v51_rows = list(csv.DictReader(f, delimiter=';'))

v4_asd = {r['model_name']: r for r in v4_rows if r['producer'] == 'АСД'}
v51_asd = {r['model_name']: r for r in v51_rows if r['producer'] == 'АСД'}

# Find ASD rows where v5.1 inside is empty but v4 had inside
needs_inside = []
for name, v51r in v51_asd.items():
    v51_inside = v51r.get('image_url_inside', '').strip()
    v4r = v4_asd.get(name)
    if v4r:
        v4_inside = v4r.get('image_url_inside', '').strip()
        if not v51_inside and v4_inside:
            needs_inside.append((name, v51r, v4r, v4_inside))

print(f"ASD rows needing inside restore: {len(needs_inside)}")

def check_wayback(url, timeout=15):
    wb_api = f"https://archive.org/wayback/available?url={url}"
    try:
        req = urllib.request.Request(wb_api, headers={'User-Agent': 'Mozilla/5.0'})
        resp = opener.open(req, timeout=timeout)
        data = json.loads(resp.read())
        snapshots = data.get('archived_snapshots', {})
        if snapshots.get('closest'):
            return snapshots['closest']['url']
    except:
        pass
    return None

def extract_images_from_wayback(wb_url, timeout=15):
    try:
        req = urllib.request.Request(wb_url, headers={'User-Agent': 'Mozilla/5.0'})
        resp = opener.open(req, timeout=timeout)
        html = resp.read().decode('utf-8', errors='ignore')
        
        imgs = re.findall(r'(?:src|data-src)="([^"]+)"', html, re.I)
        door_imgs = []
        for img in imgs:
            img_lower = img.lower()
            if any(x in img_lower for x in ['jpg', 'png', 'webp']):
                if not any(x in img_lower for x in ['logo', 'icon', 'star', 'arrow', 'banner', 'badge', 'loading', 'toolbar', 'novinka_round']):
                    if img.startswith('/web/'):
                        match = re.search(r'/web/\d+im_/(https?://.+)', img)
                        if match:
                            img = match.group(1)
                    door_imgs.append(img)
        return door_imgs
    except:
        pass
    return []

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

# Process all 159 models
results = {}
found_count = 0
start = time.time()

for i, (name, v51r, v4r, v4_inside) in enumerate(needs_inside):
    source = v4r.get('source_url', '')
    
    # Try Wayback with source URL
    wb_url = check_wayback(source)
    if wb_url:
        imgs = extract_images_from_wayback(wb_url)
        # Find inside image (usually 2nd product photo, skip novinka_round2.png)
        product_imgs = [img for img in imgs if 'novinka_round2' not in img]
        if len(product_imgs) >= 2:
            inside_candidate = product_imgs[1]  # 2nd image is often inside
            if check_image_url(inside_candidate):
                results[name] = {'inside_url': inside_candidate, 'source': 'wayback'}
                found_count += 1
            elif len(product_imgs) >= 3:
                # Try 3rd image
                inside_candidate = product_imgs[2]
                if check_image_url(inside_candidate):
                    results[name] = {'inside_url': inside_candidate, 'source': 'wayback'}
                    found_count += 1
                else:
                    results[name] = {'inside_url': '', 'source': 'wayback_no_inside'}
            else:
                results[name] = {'inside_url': '', 'source': 'wayback_no_inside'}
        elif len(product_imgs) == 1:
            results[name] = {'inside_url': '', 'source': 'wayback_only_one_img'}
        else:
            results[name] = {'inside_url': '', 'source': 'wayback_no_product_imgs'}
    else:
        results[name] = {'inside_url': '', 'source': 'no_wayback_snapshot'}
    
    if (i+1) % 20 == 0:
        print(f"  {i+1}/{len(needs_inside)} ({time.time()-start:.1f}s) found: {found_count}", flush=True)

print(f"\nDone ({time.time()-start:.1f}s): found {found_count}/{len(needs_inside)}")

# Save results
with open(os.path.join(BASE, "output", "delivery_v5_2", "wayback_inside_all_results.json"), 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

# Show found ones
print(f"\nFound inside URLs:")
for name, info in results.items():
    if info['inside_url']:
        print(f"  {name[:50]}: {info['inside_url'][-60:]}")

# Show not found
not_found = [(name, info) for name, info in results.items() if not info['inside_url']]
print(f"\nNot found: {len(not_found)}")
for name, info in not_found[:10]:
    print(f"  {name[:50]}: {info['source']}")
