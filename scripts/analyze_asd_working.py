import csv
import os
import json
from collections import Counter
from urllib.parse import urlparse

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
v5_path = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")

with open(v5_path, encoding='utf-8-sig') as f:
    v5_rows = list(csv.DictReader(f, delimiter=';'))

v5_asd = [r for r in v5_rows if r['producer'] == 'АСД']
print(f"V5 ASD rows: {len(v5_asd)}")

# Analyze which have images and their URL patterns
with_images = [r for r in v5_asd if r.get('image_url_front', '').strip()]
without_images = [r for r in v5_asd if not r.get('image_url_front', '').strip()]

print(f"With front image: {len(with_images)}")
print(f"Without front image: {len(without_images)}")

# Analyze URL patterns of working ones
print(f"\nWorking URL patterns:")
for r in with_images:
    url = r.get('image_url_front', '')
    parsed = urlparse(url)
    path_parts = parsed.path.split('/')
    # Show the pattern
    if 'assets/cache_image' in url:
        pattern = 'assets/cache_image/...'
    elif 'doors/zakaznyie' in url:
        pattern = 'doors/zakaznyie/...'
    elif 'onstock' in url:
        pattern = 'onstock/...'
    else:
        pattern = f"{parsed.netloc}/{path_parts[1] if len(path_parts) > 1 else ''}"
    print(f"  {r['model_name'][:40]}: {pattern}")
    print(f"    URL: {url[:80]}")

# Check if the without_images rows have source_urls we can scrape
print(f"\nWithout images - source_url analysis:")
source_patterns = Counter()
for r in without_images[:10]:
    source = r.get('source_url', '')
    if source:
        parsed = urlparse(source)
        source_patterns[parsed.netloc] += 1
    print(f"  {r['model_name'][:40]}: source={source[:60] if source else 'EMPTY'}")

print(f"\nSource URL domains: {dict(source_patterns)}")

# Check comments
comments = Counter(r.get('comment', '')[:50] for r in without_images)
print(f"\nComments for rows without images:")
for comment, count in comments.most_common():
    print(f"  {count}x: {comment}")
