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

# For each, try Wayback Machine
def check_wayback(url, timeout=15):
    """Check if Wayback Machine has this URL"""
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

def extract_inside_from_wayback(wb_url, timeout=15):
    """Fetch Wayback page and extract inside image"""
    try:
        req = urllib.request.Request(wb_url, headers={'User-Agent': 'Mozilla/5.0'})
        resp = opener.open(req, timeout=timeout)
        html = resp.read().decode('utf-8', errors='ignore')
        
        # Find all image URLs
        imgs = re.findall(r'(?:src|data-src)="([^"]+)"', html, re.I)
        # Filter for door images (not icons, logos, etc)
        door_imgs = []
        for img in imgs:
            img_lower = img.lower()
            if any(x in img_lower for x in ['jpg', 'png', 'webp']):
                if not any(x in img_lower for x in ['logo', 'icon', 'star', 'arrow', 'banner', 'badge', 'loading', 'toolbar']):
                    # Normalize URL
                    if img.startswith('/web/'):
                        # Extract original URL from Wayback format
                        match = re.search(r'/web/\d+im_/(https?://.+)', img)
                        if match:
                            img = match.group(1)
                    door_imgs.append(img)
        
        # The first door image is usually the main product photo (front)
        # The second might be inside
        if len(door_imgs) >= 2:
            return door_imgs[0], door_imgs[1]
        elif len(door_imgs) == 1:
            return door_imgs[0], None
    except:
        pass
    return None, None

def check_image_url(url, timeout=10):
    """Check if URL returns an image"""
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

# Test first 20 with Wayback
print(f"\nTesting Wayback Machine for first 20 models...")
found_inside = 0
results = []

for name, v51r, v4r, v4_inside in needs_inside[:20]:
    source = v4r.get('source_url', '')
    print(f"\n  {name[:50]}")
    print(f"    Source: {source[:60]}")
    print(f"    V4 inside: {v4_inside[:60]}")
    
    # Try Wayback with source URL
    wb_url = check_wayback(source)
    if wb_url:
        print(f"    Wayback: {wb_url[:60]}")
        front, inside = extract_inside_from_wayback(wb_url)
        if front:
            print(f"    Front: {front[:60]}")
        if inside:
            print(f"    Inside: {inside[:60]}")
            # Check if inside works
            if check_image_url(inside):
                found_inside += 1
                results.append({'model': name, 'inside_url': inside, 'source': 'wayback'})
                print(f"    *** INSIDE FOUND: {inside[:60]}")
            else:
                print(f"    Inside URL not accessible")
        else:
            print(f"    No inside found on page")
    else:
        print(f"    No Wayback snapshot")
    time.sleep(0.5)

print(f"\nFound inside: {found_inside}/20")

# Save results
with open(os.path.join(BASE, "output", "delivery_v5_2", "wayback_inside_results.json"), 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
