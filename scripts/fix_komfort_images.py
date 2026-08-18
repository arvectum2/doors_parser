import csv
import json
import re
import urllib.request
import ssl
from collections import Counter, OrderedDict

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

def fetch(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    resp = urllib.request.urlopen(req, timeout=15, context=ctx)
    return resp.read().decode('utf-8', errors='replace')

# Scrape all 3 catalog pages for Emal
slug_to_img = {}
for page in range(1, 4):
    url = f"https://komfortdoors.ru/emal/?page={page}"
    print(f"Fetching {url}...")
    html = fetch(url)
    
    # Find all product blocks: <a href="COLLECTION/SLUG"><img src="/assets/cache_image/products/ID/FILENAME" ...>
    pattern = r'<a\s+href="((?:new-line|slope|true|gift|art|premier|malta|neo|molding|turin)/[^"]+)"[^>]*>\s*<img\s+[^>]*src="(/assets/cache_image/products/\d+/[^"]+_0x0_9d7\.webp)"'
    matches = re.findall(pattern, html)
    print(f"  Found {len(matches)} products")
    for slug, img_path in matches:
        img_url = f"https://komfortdoors.ru{img_path}"
        slug_to_img[slug] = img_url
        print(f"  {slug} -> {img_url}")

print(f"\nTotal slug->image mappings: {len(slug_to_img)}")

# Save mapping
with open('output/final_rework_ready_v5/komfort_slug_image_map.json', 'w', encoding='utf-8') as f:
    json.dump(slug_to_img, f, ensure_ascii=False, indent=2)

# Now read current CSV and fix Komfort image URLs
with open('output/final_rework_ready_v5/doors_catalog.csv', encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

fixed = 0
for r in rows:
    if r['producer'] == 'Комфорт':
        source_url = r.get('source_url', '')
        # source_url is like https://komfortdoors.ru/emal/turin/turin-33-dg
        # Extract slug after /emal/ or just the collection/model part
        slug_match = re.search(r'komfortdoors\.ru/(.+)', source_url)
        if slug_match:
            slug = slug_match.group(1)
            if slug in slug_to_img:
                old_img = r.get('image_url', '')
                r['image_url'] = slug_to_img[slug]
                r['image_url_front'] = slug_to_img[slug]
                if old_img != slug_to_img[slug]:
                    fixed += 1
                    print(f"  FIXED: {r['model_name']} | {old_img} -> {slug_to_img[slug]}")
            else:
                print(f"  NOT FOUND: {r['model_name']} slug={slug}")

print(f"\nFixed {fixed} Komfort image URLs")

# Write updated CSV
fieldnames = list(rows[0].keys())
with open('output/final_rework_ready_v5/doors_catalog.csv', 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=';')
    writer.writeheader()
    writer.writerows(rows)
print("CSV updated")
