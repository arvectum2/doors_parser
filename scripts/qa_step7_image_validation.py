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

# Collect all image URLs with their source info
url_sources = {}
for r in rows:
    for field in ['image_url', 'image_url_front', 'image_url_inside']:
        url = r.get(field, '').strip()
        if url:
            if url not in url_sources:
                url_sources[url] = []
            url_sources[url].append({
                'producer': r['producer'],
                'model_name': r['model_name'],
                'field': field,
            })

all_urls = list(url_sources.keys())
print(f"Total unique image URLs to check: {len(all_urls)}", flush=True)

def check_url(url):
    """Check URL: follow redirects, verify content is image, check size."""
    try:
        req = urllib.request.Request(url, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        resp = opener.open(req, timeout=15)
        ct = resp.headers.get('Content-Type', '')
        cl = resp.headers.get('Content-Length', '0')
        
        # Read first 32 bytes to verify it's actually an image
        data = resp.read(32)
        
        actual_type = 'unknown'
        if data[:3] == b'\xff\xd8\xff':
            actual_type = 'image/jpeg'
        elif data[:4] == b'\x89PNG':
            actual_type = 'image/png'
        elif data[:4] == b'RIFF':
            actual_type = 'image/webp'
        elif data[:3] == b'GIF':
            actual_type = 'image/gif'
        elif data[:4] == b'\x00\x00\x00\x1c' or data[:4] == b'\x00\x00\x00\x18':
            actual_type = 'image/heic'
        elif b'<!DOCTYPE' in data or b'<html' in data:
            actual_type = 'text/html'
        
        is_image = actual_type.startswith('image/')
        is_html = actual_type == 'text/html' or (ct and 'html' in ct.lower())
        
        return url, resp.getcode(), ct, cl, actual_type, is_image, is_html, None
    except urllib.error.HTTPError as e:
        return url, e.code, '', '0', '', False, False, f"HTTP {e.code}"
    except Exception as e:
        return url, 0, '', '0', '', False, False, str(e)[:100]

start = time.time()
results = {}
errors = []

print(f"Checking with 40 threads, 15s timeout, redirect following...", flush=True)

with ThreadPoolExecutor(max_workers=40) as executor:
    futures = {executor.submit(check_url, url): url for url in all_urls}
    done = 0
    for future in as_completed(futures):
        done += 1
        if done % 500 == 0:
            elapsed = time.time() - start
            print(f"  {done}/{len(all_urls)} ({elapsed:.1f}s)", flush=True)
        url, status, ct, cl, actual_type, is_image, is_html, err = future.result()
        results[url] = {
            'status': status, 'content_type': ct, 'content_length': cl,
            'actual_type': actual_type, 'is_image': is_image, 'is_html': is_html,
            'ok': err is None
        }
        if err:
            errors.append((url, err))

elapsed = time.time() - start
print(f"\nDone in {elapsed:.1f}s", flush=True)

# Analyze results
ok_count = sum(1 for v in results.values() if v['ok'])
http_404 = [e for e in errors if '404' in e[1]]
http_other = [e for e in errors if '404' not in e[1]]
non_image = [(url, v['actual_type'], v['content_type']) for url, v in results.items() if v['ok'] and not v['is_image']]
html_urls = [(url, v['actual_type']) for url, v in results.items() if v['is_html']]
small_files = [(url, v['content_length']) for url, v in results.items() if v['ok'] and v['content_length'].isdigit() and int(v['content_length']) < 2048]

print(f"Total checked: {len(all_urls)}")
print(f"OK (reachable): {ok_count}")
print(f"HTTP 404: {len(http_404)}")
print(f"Other errors: {len(http_other)}")
print(f"Non-image content: {len(non_image)}")
print(f"HTML responses: {len(html_urls)}")
print(f"Small files (<2KB): {len(small_files)}")

# Categorize broken URLs
broken_urls = set()
broken_rows = []
for url, err in errors:
    broken_urls.add(url)
    for src in url_sources[url]:
        broken_rows.append({**src, 'url': url, 'error': err})

# Non-image URLs
non_image_rows = []
for url, at, ct in non_image:
    for src in url_sources[url]:
        non_image_rows.append({**src, 'url': url, 'actual_type': at, 'content_type': ct})

# Group errors by domain
error_domains = Counter(urlparse(u).netloc for u, _ in errors)
print(f"\nErrors by domain:")
for d, c in error_domains.most_common(10):
    print(f"  {d}: {c}")

# Group non-image by domain
non_img_domains = Counter(urlparse(u).netloc for u, _, _ in non_image)
print(f"\nNon-image by domain:")
for d, c in non_img_domains.most_common(10):
    print(f"  {d}: {c}")

# Save report
report = {
    'total_image_urls_checked': len(all_urls),
    'ok_count': ok_count,
    'broken_image_urls_count': len(errors),
    'non_image_urls_count': len(non_image),
    'html_urls_count': len(html_urls),
    'small_files_count': len(small_files),
    'elapsed_seconds': round(elapsed, 1),
    'errors': [{'url': u, 'error': e} for u, e in errors],
    'non_image': [{'url': u, 'actual_type': at, 'content_type': ct} for u, at, ct in non_image],
    'broken_rows': broken_rows,
    'non_image_rows': non_image_rows,
}

# Rows with broken images
rows_with_broken = set()
for r in broken_rows:
    rows_with_broken.add((r['producer'], r['model_name'], r['field']))

rows_with_non_image = set()
for r in non_image_rows:
    rows_with_non_image.add((r['producer'], r['model_name'], r['field']))

report['affected_rows'] = {
    'broken_image_rows': len(rows_with_broken),
    'non_image_rows': len(rows_with_non_image),
}

print(f"\nAffected rows:")
print(f"  Rows with broken images: {len(rows_with_broken)}")
print(f"  Rows with non-image URLs: {len(rows_with_non_image)}")

# Show some broken examples
print(f"\nBroken URL examples (first 10):")
for url, err in errors[:10]:
    srcs = url_sources[url]
    print(f"  {url}")
    print(f"    Error: {err}")
    print(f"    Used by: {srcs[0]['producer']} | {srcs[0]['model_name']} | {srcs[0]['field']}")

# Show non-image examples
print(f"\nNon-image URL examples (first 10):")
for url, at, ct in non_image[:10]:
    srcs = url_sources[url]
    print(f"  {url}")
    print(f"    Actual: {at}, Content-Type: {ct}")
    print(f"    Used by: {srcs[0]['producer']} | {srcs[0]['model_name']} | {srcs[0]['field']}")

with open(os.path.join(OUT_DIR, 'image_url_validation_report.json'), 'w', encoding='utf-8') as f:
    json.dump(report, f, ensure_ascii=False, indent=2)
print(f"\nReport saved to {os.path.join(OUT_DIR, 'image_url_validation_report.json')}")
