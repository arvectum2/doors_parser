import csv
import json
import os
import urllib.request
import urllib.error
import ssl
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import Counter
from urllib.parse import urlparse

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

class SmartRedirectHandler(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        try:
            return urllib.request.Request(newurl, headers=dict(req.headers))
        except:
            return None

opener = urllib.request.build_opener(SmartRedirectHandler, urllib.request.HTTPSHandler(context=ctx))

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CATALOG = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")
OUT_DIR = os.path.join(BASE, "output", "delivery")

with open(CATALOG, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

# Collect all image URLs
all_urls = set()
for r in rows:
    for field in ['image_url', 'image_url_front', 'image_url_inside']:
        url = r.get(field, '').strip()
        if url:
            all_urls.add(url)

# Only check as-doors.ru and entro-design.ru (known problematic)
check_domains = ['as-doors.ru', 'entro-design.ru', 'termo-door.ru', 'aelita-doors.ru', 'www.dveriregionov.ru']
check_urls = [u for u in all_urls if urlparse(u).netloc in check_domains]
print(f"URLs to check: {len(check_urls)} (from {check_domains})")

def check_url(url, timeout=15):
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
        resp = opener.open(req, timeout=timeout)
        ct = resp.headers.get('Content-Type', '')
        cl = resp.headers.get('Content-Length', '0')
        data = resp.read(4)
        actual_type = 'unknown'
        if data[:3] == b'\xff\xd8\xff': actual_type = 'image/jpeg'
        elif data[:4] == b'\x89PNG': actual_type = 'image/png'
        elif data[:4] == b'RIFF': actual_type = 'image/webp'
        elif data[:3] == b'GIF': actual_type = 'image/gif'
        elif b'<!DOCTYPE' in data or b'<html' in data: actual_type = 'text/html'
        is_image = actual_type.startswith('image/')
        return url, resp.getcode(), ct, cl, actual_type, is_image, None
    except urllib.error.HTTPError as e:
        return url, e.code, '', '0', '', False, f"HTTP {e.code}"
    except Exception as e:
        return url, 0, '', '0', '', False, str(e)[:100]

results = {}
errors = []

print(f"Checking {len(check_urls)} URLs...")
start = time.time()
with ThreadPoolExecutor(max_workers=30) as executor:
    futures = {executor.submit(check_url, u, 15): u for u in check_urls}
    done = 0
    for future in as_completed(futures):
        done += 1
        if done % 100 == 0:
            print(f"  {done}/{len(check_urls)} ({time.time()-start:.1f}s)", flush=True)
        url, status, ct, cl, at, is_img, err = future.result()
        results[url] = {'status': status, 'ct': ct, 'cl': cl, 'at': at, 'is_img': is_img, 'ok': err is None}
        if err:
            errors.append((url, err))

elapsed = time.time() - start
ok_count = sum(1 for v in results.values() if v['ok'])
broken_count = len(errors)
print(f"Done ({elapsed:.1f}s): {ok_count} OK, {broken_count} errors")

# Analyze errors by domain
error_by_domain = Counter()
for url, err in errors:
    d = urlparse(url).netloc
    if '404' in err: error_by_domain[f"{d}:404"] += 1
    elif '302' in err: error_by_domain[f"{d}:302"] += 1
    elif '503' in err: error_by_domain[f"{d}:503"] += 1
    elif 'timed out' in err.lower(): error_by_domain[f"{d}:timeout"] += 1
    else: error_by_domain[f"{d}:{err[:20]}"] += 1

print(f"\nErrors by domain:")
for k, c in error_by_domain.most_common():
    print(f"  {k}: {c}")

# Find broken URLs to clear
broken_urls = set()
for url, err in errors:
    if '404' in err or '302' in err or '503' in err:
        broken_urls.add(url)

print(f"\nURLs to clear (404/302/503): {len(broken_urls)}")

# Clear broken URLs in CSV
cleared = 0
for r in rows:
    for field in ['image_url', 'image_url_front', 'image_url_inside']:
        url = r.get(field, '').strip()
        if url in broken_urls:
            r[field] = ''
            r['status'] = 'needs_review'
            r['comment'] = f'image_url_cleared_{[e for u, e in errors if u == url][0][:20]}'
            cleared += 1

print(f"Cleared {cleared} broken URLs in CSV")

# Save updated CSV
fieldnames = list(rows[0].keys())
with open(CATALOG, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=';')
    writer.writeheader()
    writer.writerows(rows)

# Recount
no_image = sum(1 for r in rows if not r.get('image_url', '').strip() and not r.get('image_url_front', '').strip())
print(f"\nFinal: {len(rows)} rows, {no_image} without images, {len(rows) - no_image} with images")
