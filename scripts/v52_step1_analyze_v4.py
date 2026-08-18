import csv
import os
import json
import hashlib
import time
import urllib.request
import urllib.error
import ssl
import re
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from urllib.parse import urlparse
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
V4_CATALOG = os.path.join(BASE, "output", "final_rework_ready_v4", "doors_catalog.csv")
V51_CATALOG = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")
PATCH_PATH = os.path.join(BASE, "output", "asd_photo_patch_v2_clean", "asd_photo_patch_v2_clean.csv")
OUT_DIR = os.path.join(BASE, "output", "delivery_v5_2")
DELIVERY_DIR = os.path.join(BASE, "output", "delivery")
os.makedirs(OUT_DIR, exist_ok=True)

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE
opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))

EXPECTED_PRODUCERS = [
    "Скамбио Порте", "Хаусдорс", "Олимп", "Аэлита", "Лорд",
    "Комфорт", "Двери регионов", "Термодур", "Энтро", "АСД"
]

# Load catalogs
with open(V4_CATALOG, encoding='utf-8-sig') as f:
    v4_rows = list(csv.DictReader(f, delimiter=';'))
with open(V51_CATALOG, encoding='utf-8-sig') as f:
    v51_rows = list(csv.DictReader(f, delimiter=';'))
fieldnames = list(v51_rows[0].keys())

print(f"V4 rows: {len(v4_rows)}, V5.1 rows: {len(v51_rows)}")

# Build v4 ASD lookup
v4_asd = {}
for r in v4_rows:
    if r['producer'] == 'АСД':
        key = (r['producer'], r['target_line'], r['model_name'])
        v4_asd[key] = r

# Build v5.1 ASD rows
v51_asd = {}
for r in v51_rows:
    if r['producer'] == 'АСД':
        key = (r['producer'], r['target_line'], r['model_name'])
        v51_asd[key] = r

print(f"V4 ASD: {len(v4_asd)}, V5.1 ASD: {len(v51_asd)}")

# Check v4 ASD inside status
v4_inside_ok = sum(1 for r in v4_asd.values() if r.get('image_url_inside', '').strip())
v4_front_ok = sum(1 for r in v4_asd.values() if r.get('image_url_front', '').strip())
print(f"V4 ASD front: {v4_front_ok}/199, inside: {v4_inside_ok}/199")

# Check v4 inside URLs - which are old format and which are new
v4_inside_old_format = 0
v4_inside_new_format = 0
for r in v4_asd.values():
    inside = r.get('image_url_inside', '').strip()
    if inside:
        if 'assets/cache_image' in inside or 'img/' in inside or 'images/' in inside:
            v4_inside_new_format += 1
        else:
            v4_inside_old_format += 1

print(f"V4 inside format: new={v4_inside_new_format}, old={v4_inside_old_format}")

# Show v4 inside URL examples
print(f"\nV4 inside URL examples (first 5):")
for r in list(v4_asd.values())[:5]:
    inside = r.get('image_url_inside', '')
    if inside:
        print(f"  {r['model_name'][:40]}: {inside[:80]}")

# Show v4 inside URLs that use old format
print(f"\nV4 inside URLs with old format (doors/zakaznyie):")
old_inside = [(r['model_name'], r.get('image_url_inside', '')) for r in v4_asd.values() 
              if r.get('image_url_inside', '').strip() and 'doors/zakaznyie' in r.get('image_url_inside', '')]
print(f"  Count: {len(old_inside)}")
for name, url in old_inside[:5]:
    print(f"  {name[:40]}: {url[:80]}")

# Show v4 inside URLs that use onstock format
old_inside_onstock = [(r['model_name'], r.get('image_url_inside', '')) for r in v4_asd.values() 
                      if r.get('image_url_inside', '').strip() and 'onstock' in r.get('image_url_inside', '')]
print(f"\nV4 inside URLs with onstock format:")
print(f"  Count: {len(old_inside_onstock)}")
for name, url in old_inside_onstock[:5]:
    print(f"  {name[:40]}: {url[:80]}")

# Show v4 inside URLs that use new format
new_inside = [(r['model_name'], r.get('image_url_inside', '')) for r in v4_asd.values() 
              if r.get('image_url_inside', '').strip() and ('assets/cache_image' in r.get('image_url_inside', '') or '/img/' in r.get('image_url_inside', '') or '/images/' in r.get('image_url_inside', ''))]
print(f"\nV4 inside URLs with new format:")
print(f"  Count: {len(new_inside)}")
for name, url in new_inside[:5]:
    print(f"  {name[:40]}: {url[:80]}")
