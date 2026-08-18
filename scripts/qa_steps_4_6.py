import csv
import json
import os
from collections import Counter
from urllib.parse import urlparse
import urllib.request
import ssl
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CATALOG = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")
OUT_DIR = os.path.join(BASE, "output", "delivery")

with open(CATALOG, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

EXPECTED_PRODUCERS = [
    "Скамбио Порте", "Хаусдорс", "Олимп", "Аэлита", "Лорд",
    "Комфорт", "Двери регионов", "Термодур", "Энтро", "АСД"
]

# ========== STEP 4: Olimp Paris check ==========
print("=== STEP 4: Olimp Paris check ===")
parizh = [r for r in rows if r['producer'] == 'Олимп' and 'Париж' in r.get('target_line', '')]
print(f"  Paris rows: {len(parizh)}")
parizh_ok = True
for r in parizh:
    has_img = bool(r.get('image_url', '').strip()) or bool(r.get('image_url_front', '').strip())
    has_source = bool(r.get('source_url', '').strip())
    status = r.get('status', '')
    comment = r.get('comment', '')
    img_ok = not has_img  # images should be cleared
    source_ok = has_source
    status_ok = status == 'needs_review' or '404' in comment or 'not available' in comment.lower()
    comment_ok = '404' in comment or 'not available' in comment.lower() or 'image' in comment.lower()

    if not (img_ok and source_ok and status_ok and comment_ok):
        parizh_ok = False
        print(f"  PROBLEM: {r['model_name']}")
        print(f"    image_url: '{r.get('image_url', '')}' (should be empty)")
        print(f"    source_url: '{r.get('source_url', '')}' (should be filled)")
        print(f"    status: '{status}' (should be needs_review)")
        print(f"    comment: '{comment}' (should mention 404)")

if parizh_ok:
    print(f"  All {len(parizh)} Paris rows OK: images cleared, source_url filled, status=needs_review, comment mentions 404")

# ========== STEP 5: Komfort check ==========
print("\n=== STEP 5: Komfort check ===")
kf = [r for r in rows if r['producer'] == 'Комфорт']
kf_lines = Counter(r['target_line'] for r in kf)
print(f"  Total Komfort: {len(kf)} (expected 119)")
print(f"  Collections: {dict(sorted(kf_lines.items()))}")

# Check no duplicates
kf_keys = set()
kf_dupes = 0
for r in kf:
    key = (r['producer'], r['target_line'], r['model_name'])
    if key in kf_keys:
        kf_dupes += 1
    kf_keys.add(key)
print(f"  Duplicates: {kf_dupes}")

# Check expected collections
expected_collections = ['Art', 'Gift', 'Malta', 'Molding', 'Neo', 'New Line', 'Premier', 'Slope', 'True', 'Turin']
unexpected_collections = [c for c in kf_lines if c not in expected_collections]
print(f"  Expected collections present: {all(c in kf_lines for c in expected_collections)}")
if unexpected_collections:
    print(f"  Unexpected collections: {unexpected_collections}")

# Check Экошпон/Эмаль presence
eko = sum(1 for r in kf if 'кошпон' in r.get('target_line', '').lower() or 'ekoshpon' in r.get('target_line', '').lower())
emal = sum(1 for r in kf if 'маль' in r.get('target_line', '').lower() or 'emal' in r.get('target_line', '').lower())
print(f"  Экошпон rows: {eko} (should be 0 - not on site)")
print(f"  Эмаль rows: {emal} (should be 0 - parent category, not individual)")

# Check image URLs have komfortdoors.ru domain
kf_img_urls = [r.get('image_url', '') for r in kf if r.get('image_url', '').strip()]
kf_domains = Counter(urlparse(u).netloc for u in kf_img_urls)
print(f"  Image URL domains: {dict(kf_domains)}")
print(f"  All images from komfortdoors.ru: {all(d == 'komfortdoors.ru' for d in kf_domains)}")

# ========== STEP 6: Hausdoors check ==========
print("\n=== STEP 6: Hausdoors check ===")
hd = [r for r in rows if r['producer'] == 'Хаусдорс']
hd_lines = Counter(r['target_line'] for r in hd)
print(f"  Total Hausdoors: {len(hd)} (expected 58)")
print(f"  Collections: {dict(sorted(hd_lines.items()))}")

pvc_collections = ['Фактурные линии', 'Новая классика', 'Минимализм', 'Современная классика']
pvc_models = [r for r in hd if r['target_line'] in pvc_collections]
non_pvc = [r for r in hd if r['target_line'] not in pvc_collections]
print(f"  PVC models: {len(pvc_models)} (expected 31)")
print(f"  Non-PVC models: {len(non_pvc)} (expected 27)")

# Check PVC model counts
for c in pvc_collections:
    count = sum(1 for r in pvc_models if r['target_line'] == c)
    print(f"    {c}: {count}")

# Check no stubs remain
stubs = [r for r in hd if '[MISSING]' in r.get('model_name', '')]
print(f"  Remaining stubs: {len(stubs)} (expected 0)")

# Check PVC images
pvc_with_img = sum(1 for r in pvc_models if r.get('image_url_front', '').strip() or r.get('image_url', '').strip())
print(f"  PVC models with images: {pvc_with_img}/{len(pvc_models)}")

# Check Hausdoors image domains
hd_img_urls = [r.get('image_url', '') for r in hd if r.get('image_url', '').strip()]
hd_domains = Counter(urlparse(u).netloc for u in hd_img_urls)
print(f"  Image URL domains: {dict(hd_domains)}")
