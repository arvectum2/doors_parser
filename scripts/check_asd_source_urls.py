import csv
import os
import urllib.request
import urllib.error
import ssl
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import Counter
import re
from html.parser import HTMLParser

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

# Find ASD rows that lost images
lost = []
for name, v5r in v5_asd.items():
    v4r = v4_asd.get(name)
    if not v4r:
        continue
    if v4r.get('image_url_front', '').strip() and not v5r.get('image_url_front', '').strip():
        lost.append((name, v4r, v5r))

print(f"ASD rows that lost front image: {len(lost)}")

# For each lost row, try to scrape source_url to find new image
class ImageFinder(HTMLParser):
    def __init__(self):
        super().__init__()
        self.images = []
    def handle_starttag(self, tag, attrs):
        if tag == 'img':
            for attr, val in attrs:
                if attr == 'src' and val:
                    self.images.append(val)

def find_images_on_page(url, timeout=15):
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
        resp = opener.open(req, timeout=timeout)
        ct = resp.headers.get('Content-Type', '')
        if 'text/html' not in ct:
            return url, None, f"not_html: {ct}"
        html = resp.read().decode('utf-8', errors='ignore')
        parser = ImageFinder()
        parser.feed(html)
        return url, parser.images, None
    except Exception as e:
        return url, None, str(e)[:100]

# Test first 10 source URLs
print(f"\nTesting source URLs for image discovery...")
found_images = 0
for name, v4r, v5r in lost[:10]:
    source = v4r.get('source_url', '')
    if not source:
        print(f"  {name[:40]}: no source_url")
        continue
    url, images, err = find_images_on_page(source)
    if err:
        print(f"  {name[:40]}: ERROR {err[:40]}")
    elif images:
        # Find images that look like door photos
        door_imgs = [img for img in images if any(x in img.lower() for x in ['img', 'image', 'photo', 'door', 'jpg', 'png', 'webp'])]
        if door_imgs:
            found_images += 1
            print(f"  {name[:40]}: {len(door_imgs)} door images found")
            for img in door_imgs[:3]:
                print(f"    {img[:80]}")
        else:
            print(f"  {name[:40]}: {len(images)} images, none match door patterns")
    else:
        print(f"  {name[:40]}: no images on page")

print(f"\nFound images on {found_images}/{min(10, len(lost))} source pages")

# Also check which v4 URLs use the NEW format (assets/cache_image)
new_format = [r for name, r in v4_asd.items() if 'assets/cache_image' in r.get('image_url_front', '')]
old_format = [r for name, r in v4_asd.items() if 'assets/cache_image' not in r.get('image_url_front', '') and r.get('image_url_front', '')]
print(f"\nV4 ASD URL formats:")
print(f"  New format (assets/cache_image): {len(new_format)}")
print(f"  Old format (doors/zakaznyie/...): {len(old_format)}")

# Check if the new format URLs work
new_urls = [r.get('image_url_front', '') for r in new_format][:10]
print(f"\nTesting new format URLs...")
new_ok = 0
for url in new_urls:
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        resp = opener.open(req, timeout=10)
        ct = resp.headers.get('Content-Type', '')
        data = resp.read(4)
        is_img = data[:3] == b'\xff\xd8\xff' or data[:4] == b'\x89PNG' or data[:4] == b'RIFF'
        if is_img:
            new_ok += 1
    except:
        pass

print(f"  New format working: {new_ok}/{len(new_urls)}")

# Check the 37 ASD rows that still have images in v5
still_have = [(name, v5r) for name, v5r in v5_asd.items() if v5r.get('image_url_front', '').strip()]
print(f"\nASD rows still having images in v5: {len(still_have)}")
for name, r in still_have[:5]:
    print(f"  {name[:40]}: {r['image_url_front'][:60]}")
