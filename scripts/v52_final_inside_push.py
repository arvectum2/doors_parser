import csv
import os
import json
import urllib.request
import urllib.error
import ssl
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

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

# Find ASD rows missing inside
asd_indices = [(i, r) for i, r in enumerate(rows) if r['producer'] == 'АСД']
missing_inside = [(i, r) for i, r in asd_indices if not r.get('image_url_inside', '').strip()]

print(f"Missing inside: {len(missing_inside)}")

inside_found = 0
for idx, r in missing_inside:
    name = r['model_name']
    front = r.get('image_url_front', '').strip()
    v4r = v4_asd.get(name, {})
    source = v4r.get('source_url', '')
    
    print(f"\n  {name[:50]}")
    
    # Try Wayback with source URL
    wb_url = check_wayback(source)
    if wb_url:
        imgs = extract_images_from_wayback(wb_url)
        product_imgs = [img for img in imgs if 'novinka_round2' not in img]
        if len(product_imgs) >= 2:
            inside_candidate = product_imgs[1]
            if check_image_url(inside_candidate):
                http_url = inside_candidate.replace('https://', 'http://')
                rows[idx]['image_url_inside'] = http_url
                inside_found += 1
                print(f"    WAYBACK INSIDE: {http_url[:60]}")
                continue
    
    # Try more aggressive pattern matching on front URL
    if front:
        # Try different directory structures
        front_match = re.search(r'(/(?:img|images)/\d{4}/[^/]+/)', front)
        if front_match:
            dir_path = front_match.group(1)
            for num in range(1, 30):
                for ext in ['jpg', 'png', 'webp']:
                    test_url = f"http://as-doors.ru{dir_path}*-{num}.{ext}"
                    # Actually try specific filenames
                    base = front.split('/')[-1]
                    base_match = re.match(r'(.+?)(?:-\d+)?\.(.+)$', base)
                    if base_match:
                        base_name = base_match.group(1)
                        test_url = f"http://as-doors.ru{dir_path}{base_name}-{num}.{ext}"
                        if test_url != front and check_image_url(test_url):
                            rows[idx]['image_url_inside'] = test_url
                            inside_found += 1
                            print(f"    PATTERN INSIDE: {test_url[:60]}")
                            break
                else:
                    continue
                break

print(f"\nExtra inside found: {inside_found}")

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
