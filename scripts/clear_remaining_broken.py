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

# Find all remaining as-doors.ru URLs
as_urls = set()
for r in rows:
    for field in ['image_url', 'image_url_front', 'image_url_inside']:
        url = r.get(field, '').strip()
        if url and 'as-doors.ru' in url:
            as_urls.add(url)

print(f"Remaining as-doors.ru URLs: {len(as_urls)}")

def check_url(url, timeout=15):
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
        resp = opener.open(req, timeout=timeout)
        ct = resp.headers.get('Content-Type', '')
        data = resp.read(4)
        is_img = data[:3] == b'\xff\xd8\xff' or data[:4] == b'\x89PNG' or data[:4] == b'RIFF' or data[:4] == b'RIFF'
        return url, resp.getcode(), ct, is_img, None
    except urllib.error.HTTPError as e:
        return url, e.code, '', False, f"HTTP {e.code}"
    except Exception as e:
        return url, 0, '', False, str(e)[:100]

broken_urls = set()
start = time.time()
with ThreadPoolExecutor(max_workers=30) as executor:
    futures = {executor.submit(check_url, u, 15): u for u in as_urls}
    done = 0
    for future in as_completed(futures):
        done += 1
        url, status, ct, is_img, err = future.result()
        if err:
            broken_urls.add(url)
            if done % 20 == 0:
                print(f"  {done}/{len(as_urls)} ({time.time()-start:.1f}s) broken: {len(broken_urls)}", flush=True)

print(f"Done ({time.time()-start:.1f}s): {len(as_urls) - len(broken_urls)} OK, {len(broken_urls)} broken")

# Clear broken URLs
cleared = 0
for r in rows:
    for field in ['image_url', 'image_url_front', 'image_url_inside']:
        url = r.get(field, '').strip()
        if url in broken_urls:
            r[field] = ''
            r['status'] = 'needs_review'
            if not r.get('comment', '') or r['comment'] == 'found':
                r['comment'] = 'image_url_cleared_302_redirect'
            cleared += 1

print(f"Cleared {cleared} broken URLs")

# Save CSV
fieldnames = list(rows[0].keys())
with open(CATALOG, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=';')
    writer.writeheader()
    writer.writerows(rows)

# Final counts
no_image = sum(1 for r in rows if not r.get('image_url', '').strip() and not r.get('image_url_front', '').strip())
has_image = len(rows) - no_image

# Count unique image URLs
all_urls_final = set()
for r in rows:
    for field in ['image_url', 'image_url_front', 'image_url_inside']:
        url = r.get(field, '').strip()
        if url:
            all_urls_final.add(url)

print(f"\nFinal: {len(rows)} rows")
print(f"  With images: {has_image}")
print(f"  Without images: {no_image}")
print(f"  Unique image URLs: {len(all_urls_final)}")

# No-image by producer
no_img_rows = [r for r in rows if not r.get('image_url', '').strip() and not r.get('image_url_front', '').strip()]
no_img_by_prod = Counter(r['producer'] for r in no_img_rows)
print(f"\nNo-image by producer:")
for p, c in sorted(no_img_by_prod.items()):
    print(f"  {p}: {c}")

# Save no-image report
no_img_details = []
for r in no_img_rows:
    comment = r.get('comment', '')
    if '404' in comment: reason = "image_404_on_manufacturer_site"
    elif '302' in comment: reason = "image_302_redirect_cleared"
    elif 'cleared' in comment: reason = "image_url_cleared_broken"
    elif '[MISSING]' in r.get('model_name', ''): reason = "stub_model_no_source"
    elif 'network' in comment.lower(): reason = "network_unreachable"
    else: reason = r.get('comment', 'unknown')
    
    no_img_details.append({
        'producer': r['producer'],
        'target_line': r['target_line'],
        'model_name': r['model_name'],
        'source_url': r.get('source_url', ''),
        'image_url': r.get('image_url', ''),
        'image_url_front': r.get('image_url_front', ''),
        'image_url_inside': r.get('image_url_inside', ''),
        'status': r.get('status', ''),
        'comment': r.get('comment', ''),
        'reason_no_image': reason,
    })

fieldnames_ni = ['producer', 'target_line', 'model_name', 'source_url', 'image_url', 'image_url_front', 'image_url_inside', 'status', 'comment', 'reason_no_image']
with open(os.path.join(OUT_DIR, 'no_image_rows_v5.csv'), 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames_ni, delimiter=';')
    writer.writeheader()
    writer.writerows(no_img_details)

reason_counts = Counter(r['reason_no_image'] for r in no_img_details)
print(f"\nNo-image by reason:")
for reason, count in sorted(reason_counts.items(), key=lambda x: -x[1]):
    print(f"  {reason}: {count}")
