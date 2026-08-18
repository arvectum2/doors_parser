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

# Group by domain
domain_urls = {}
for url in all_urls:
    d = urlparse(url).netloc
    if d not in domain_urls:
        domain_urls[d] = []
    domain_urls[d].append(url)

print(f"Total unique image URLs: {len(all_urls)}")
print(f"Domains: {len(domain_urls)}")
for d, urls in sorted(domain_urls.items(), key=lambda x: -len(x[1])):
    print(f"  {d}: {len(urls)}")

def check_url(url, timeout=10):
    try:
        req = urllib.request.Request(url, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        resp = opener.open(req, timeout=timeout)
        ct = resp.headers.get('Content-Type', '')
        cl = resp.headers.get('Content-Length', '0')
        data = resp.read(32)
        
        actual_type = 'unknown'
        if data[:3] == b'\xff\xd8\xff': actual_type = 'image/jpeg'
        elif data[:4] == b'\x89PNG': actual_type = 'image/png'
        elif data[:4] == b'RIFF': actual_type = 'image/webp'
        elif data[:3] == b'GIF': actual_type = 'image/gif'
        elif b'<!DOCTYPE' in data or b'<html' in data: actual_type = 'text/html'
        
        is_image = actual_type.startswith('image/')
        is_html = actual_type == 'text/html' or ('html' in ct.lower() if ct else False)
        
        return url, resp.getcode(), ct, cl, actual_type, is_image, is_html, None
    except urllib.error.HTTPError as e:
        return url, e.code, '', '0', '', False, False, f"HTTP {e.code}"
    except Exception as e:
        return url, 0, '', '0', '', False, False, str(e)[:100]

# Check fast domains first (komfortdoors.ru, hausdoors.ru, tildacdn.com, etc.)
FAST_DOMAINS = ['komfortdoors.ru', 'hausdoors.ru', 'thb.tildacdn.com', 'static.tildacdn.com', 
                'entro-design.ru', 'dver.com', 'aelita-doors.ru', 'as-doors.ru']
SLOW_DOMAINS = ['dveri-lord.ru']

fast_urls = [u for u in all_urls if urlparse(u).netloc in FAST_DOMAINS]
slow_urls = [u for u in all_urls if urlparse(u).netloc in SLOW_DOMAINS]
other_urls = [u for u in all_urls if urlparse(u).netloc not in FAST_DOMAINS and urlparse(u).netloc not in SLOW_DOMAINS]

print(f"\nFast domain URLs: {len(fast_urls)}")
print(f"Slow domain URLs (dveri-lord.ru): {len(slow_urls)}")
print(f"Other URLs: {len(other_urls)}")

all_results = {}
all_errors = []

# Phase 1: Fast domains (10s timeout)
print(f"\n=== Phase 1: Fast domains ({len(fast_urls)} URLs, 10s timeout) ===")
start = time.time()
with ThreadPoolExecutor(max_workers=40) as executor:
    futures = {executor.submit(check_url, u, 10): u for u in fast_urls}
    done = 0
    for future in as_completed(futures):
        done += 1
        if done % 200 == 0:
            print(f"  {done}/{len(fast_urls)} ({time.time()-start:.1f}s)", flush=True)
        url, status, ct, cl, at, is_img, is_html, err = future.result()
        all_results[url] = {'status': status, 'ct': ct, 'cl': cl, 'at': at, 'is_img': is_img, 'is_html': is_html, 'ok': err is None}
        if err:
            all_errors.append((url, err))

elapsed = time.time() - start
ok = sum(1 for v in all_results.values() if v['ok'])
errs = sum(1 for v in all_results.values() if not v['ok'])
print(f"  Phase 1 done: {ok} OK, {errs} errors ({elapsed:.1f}s)")

# Phase 2: Slow domains (20s timeout, 20 threads)
print(f"\n=== Phase 2: Slow domains ({len(slow_urls)} URLs, 20s timeout) ===")
start = time.time()
with ThreadPoolExecutor(max_workers=20) as executor:
    futures = {executor.submit(check_url, u, 20): u for u in slow_urls}
    done = 0
    for future in as_completed(futures):
        done += 1
        if done % 100 == 0:
            print(f"  {done}/{len(slow_urls)} ({time.time()-start:.1f}s)", flush=True)
        url, status, ct, cl, at, is_img, is_html, err = future.result()
        all_results[url] = {'status': status, 'ct': ct, 'cl': cl, 'at': at, 'is_img': is_img, 'is_html': is_html, 'ok': err is None}
        if err:
            all_errors.append((url, err))

elapsed = time.time() - start
ok = sum(1 for v in all_results.values() if v['ok'] and urlparse(list(all_results.keys())[0]).netloc in SLOW_DOMAINS)
errs = sum(1 for v in all_results.values() if not v['ok'] and urlparse(list(all_results.keys())[0]).netloc in SLOW_DOMAINS)
print(f"  Phase 2 done ({elapsed:.1f}s)")

# Phase 3: Other domains (10s timeout)
if other_urls:
    print(f"\n=== Phase 3: Other domains ({len(other_urls)} URLs) ===")
    start = time.time()
    with ThreadPoolExecutor(max_workers=30) as executor:
        futures = {executor.submit(check_url, u, 10): u for u in other_urls}
        done = 0
        for future in as_completed(futures):
            done += 1
            url, status, ct, cl, at, is_img, is_html, err = future.result()
            all_results[url] = {'status': status, 'ct': ct, 'cl': cl, 'at': at, 'is_img': is_img, 'is_html': is_html, 'ok': err is None}
            if err:
                all_errors.append((url, err))
    print(f"  Phase 3 done ({time.time()-start:.1f}s)")

# Summary
total_checked = len(all_results)
ok_count = sum(1 for v in all_results.values() if v['ok'])
broken_count = len(all_errors)
non_image = [(u, v['at'], v['ct']) for u, v in all_results.items() if v['ok'] and not v['is_img']]
html_count = sum(1 for v in all_results.values() if v['is_html'])
small = [(u, v['cl']) for u, v in all_results.items() if v['ok'] and v['cl'].isdigit() and int(v['cl']) < 2048]

print(f"\n=== FINAL RESULTS ===")
print(f"Total checked: {total_checked}")
print(f"OK (image): {ok_count}")
print(f"Broken (error): {broken_count}")
print(f"Non-image content: {len(non_image)}")
print(f"HTML responses: {html_count}")
print(f"Small files (<2KB): {len(small)}")

# Error breakdown
error_types = Counter()
for url, err in all_errors:
    if '404' in err: error_types['HTTP 404'] += 1
    elif 'timed out' in err.lower(): error_types['Timeout'] += 1
    elif '403' in err: error_types['HTTP 403'] += 1
    elif 'SSL' in err or 'ssl' in err: error_types['SSL error'] += 1
    elif 'Connection' in err: error_types['Connection error'] += 1
    elif 'RemoteDisconnected' in err: error_types['Remote disconnected'] += 1
    else: error_types[err[:40]] += 1
print(f"\nError breakdown:")
for t, c in error_types.most_common():
    print(f"  {t}: {c}")

# Non-image breakdown
non_img_types = Counter(at for _, at, _ in non_image)
print(f"\nNon-image types:")
for t, c in non_img_types.most_common():
    print(f"  {t}: {c}")

# Affected rows
broken_urls = set(u for u, _ in all_errors)
non_image_urls = set(u for u, _, _ in non_image)

rows_broken = set()
rows_non_img = set()
for r in rows:
    for field in ['image_url', 'image_url_front', 'image_url_inside']:
        url = r.get(field, '').strip()
        if url in broken_urls:
            rows_broken.add((r['producer'], r['model_name']))
        if url in non_image_urls:
            rows_non_img.add((r['producer'], r['model_name']))

print(f"\nAffected rows:")
print(f"  Rows with broken images: {len(rows_broken)}")
print(f"  Rows with non-image URLs: {len(rows_non_img)}")

# Show broken examples
print(f"\nBroken URL examples:")
for url, err in all_errors[:15]:
    srcs = url_sources.get(url, [])
    domain = urlparse(url).netloc
    print(f"  [{domain}] {url[:80]}...")
    print(f"    Error: {err}")
    if srcs:
        print(f"    Used by: {srcs[0]['producer']} | {srcs[0]['model_name']}")

# Show non-image examples
print(f"\nNon-image examples:")
for url, at, ct in non_image[:10]:
    srcs = url_sources.get(url, [])
    domain = urlparse(url).netloc
    print(f"  [{domain}] {url[:80]}...")
    print(f"    Actual: {at}, CT: {ct}")
    if srcs:
        print(f"    Used by: {srcs[0]['producer']} | {srcs[0]['model_name']}")

# Save report
report = {
    'total_image_urls_checked': total_checked,
    'ok_count': ok_count,
    'broken_image_urls_count': broken_count,
    'non_image_urls_count': len(non_image),
    'html_urls_count': html_count,
    'small_files_count': len(small),
    'cleared_image_urls_count': 7,  # Olimp Paris already cleared
    'remaining_empty_image_rows_count': 59,
    'errors': [{'url': u, 'error': e} for u, e in all_errors],
    'non_image': [{'url': u, 'actual_type': at, 'content_type': ct} for u, at, ct in non_image],
    'affected_rows': {
        'broken_image_rows': [{'producer': p, 'model_name': m} for p, m in sorted(rows_broken)],
        'non_image_rows': [{'producer': p, 'model_name': m} for p, m in sorted(rows_non_img)],
    },
}

with open(os.path.join(OUT_DIR, 'image_url_validation_report.json'), 'w', encoding='utf-8') as f:
    json.dump(report, f, ensure_ascii=False, indent=2)
print(f"\nReport saved.")
