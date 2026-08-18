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

url_sources = {}
for r in rows:
    for field in ['image_url', 'image_url_front', 'image_url_inside']:
        url = r.get(field, '').strip()
        if url:
            if url not in url_sources:
                url_sources[url] = []
            url_sources[url].append({'producer': r['producer'], 'model_name': r['model_name'], 'field': field})

all_urls = list(url_sources.keys())

# Separate domains
DVERI_LORD_DOMAIN = 'dveri-lord.ru'
dveri_lord_urls = [u for u in all_urls if urlparse(u).netloc == DVERI_LORD_DOMAIN]
other_urls = [u for u in all_urls if urlparse(u).netloc != DVERI_LORD_DOMAIN]

print(f"Total: {len(all_urls)}, dveri-lord.ru: {len(dveri_lord_urls)}, other: {len(other_urls)}")

def check_url(url, timeout=10):
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
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

all_results = {}
all_errors = []

# Phase 1: Non-dveri-lord URLs (40 threads, 10s timeout)
print(f"\n=== Phase 1: {len(other_urls)} non-dveri-lord URLs ===")
start = time.time()
with ThreadPoolExecutor(max_workers=40) as executor:
    futures = {executor.submit(check_url, u, 10): u for u in other_urls}
    done = 0
    for future in as_completed(futures):
        done += 1
        if done % 500 == 0:
            print(f"  {done}/{len(other_urls)} ({time.time()-start:.1f}s)", flush=True)
        url, status, ct, cl, at, is_img, is_html, err = future.result()
        all_results[url] = {'status': status, 'ct': ct, 'cl': cl, 'at': at, 'is_img': is_img, 'is_html': is_html, 'ok': err is None, 'domain': urlparse(url).netloc}
        if err:
            all_errors.append((url, err))

other_ok = sum(1 for u in other_urls if all_results.get(u, {}).get('ok'))
other_err = len(other_urls) - other_ok
print(f"  Done: {other_ok} OK, {other_err} errors ({time.time()-start:.1f}s)")

# Phase 2: dveri-lord.ru - sample 20 URLs to check if they work at all
print(f"\n=== Phase 2: dveri-lord.ru sampling ===")
sample_size = min(20, len(dveri_lord_urls))
sample = dveri_lord_urls[:sample_size]
start = time.time()
dveri_lord_ok = 0
dveri_lord_err = 0
with ThreadPoolExecutor(max_workers=10) as executor:
    futures = {executor.submit(check_url, u, 15): u for u in sample}
    for future in as_completed(futures):
        url, status, ct, cl, at, is_img, is_html, err = future.result()
        all_results[url] = {'status': status, 'ct': ct, 'cl': cl, 'at': at, 'is_img': is_img, 'is_html': is_html, 'ok': err is None, 'domain': 'dveri-lord.ru'}
        if err:
            dveri_lord_err += 1
            all_errors.append((url, err))
        else:
            dveri_lord_ok += 1

dveri_lord_sample_rate = dveri_lord_ok / sample_size if sample_size > 0 else 0
dveri_lord_estimated_ok = int(len(dveri_lord_urls) * dveri_lord_sample_rate)
dveri_lord_estimated_err = len(dveri_lord_urls) - dveri_lord_estimated_ok
print(f"  Sample {sample_size}: {dveri_lord_ok} OK, {dveri_lord_err} errors ({time.time()-start:.1f}s)")
print(f"  OK rate: {dveri_lord_sample_rate:.1%}")
print(f"  Estimated for all {len(dveri_lord_urls)}: ~{dveri_lord_estimated_ok} OK, ~{dveri_lord_estimated_err} errors")

# For dveri-lord.ru, check specific error types
dveri_lord_errors_sampled = [(u, e) for u, e in all_errors if urlparse(u).netloc == 'dveri-lord.ru']
if dveri_lord_errors_sampled:
    print(f"  dveri-lord error types (sampled):")
    err_types = Counter(e[:40] for _, e in dveri_lord_errors_sampled)
    for t, c in err_types.most_common():
        print(f"    {t}: {c}")

# Summary
total_checked = len(all_results)
ok_count = sum(1 for v in all_results.values() if v['ok'])
broken_count = len(all_errors)
non_image = [(u, v['at'], v['ct']) for u, v in all_results.items() if v['ok'] and not v['is_img']]
html_count = sum(1 for v in all_results.values() if v['is_html'])
small = [(u, v['cl']) for u, v in all_results.items() if v['ok'] and v['cl'].isdigit() and int(v['cl']) < 2048]

# Add dveri-lord estimated results
total_with_estimate = total_checked + (len(dveri_lord_urls) - sample_size)
ok_with_estimate = ok_count + dveri_lord_estimated_ok - dveri_lord_ok

print(f"\n=== FINAL RESULTS ===")
print(f"Total unique URLs: {len(all_urls)}")
print(f"Actually checked: {total_checked}")
print(f"OK (image, sampled): {ok_count}")
print(f"Broken (error, sampled): {broken_count}")
print(f"Non-image content: {len(non_image)}")
print(f"HTML responses: {html_count}")
print(f"Small files (<2KB): {len(small)}")
print(f"")
print(f"dveri-lord.ru estimated: ~{dveri_lord_estimated_ok} OK, ~{dveri_lord_estimated_err} errors")
print(f"Total estimated OK: ~{ok_with_estimate}/{len(all_urls)}")
print(f"Total estimated broken: ~{broken_count + dveri_lord_estimated_err}")

# Error breakdown
error_types_all = Counter()
for url, err in all_errors:
    if '404' in err: error_types_all['HTTP 404'] += 1
    elif 'timed out' in err.lower(): error_types_all['Timeout'] += 1
    elif '403' in err: error_types_all['HTTP 403'] += 1
    elif 'SSL' in err: error_types_all['SSL error'] += 1
    elif 'Connection' in err: error_types_all['Connection error'] += 1
    else: error_types_all[err[:40]] += 1
print(f"\nError breakdown (sampled):")
for t, c in error_types_all.most_common():
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

# Show broken examples by producer
broken_by_prod = Counter()
for p, m in rows_broken:
    broken_by_prod[p] += 1
print(f"\nBroken images by producer:")
for p, c in broken_by_prod.most_common():
    print(f"  {p}: {c}")

# Show non-image examples
print(f"\nNon-image URL details:")
for url, at, ct in non_image:
    srcs = url_sources.get(url, [])
    print(f"  {url[:80]}...")
    print(f"    Actual: {at}, CT: {ct}, Domain: {urlparse(url).netloc}")
    if srcs:
        print(f"    Used by: {srcs[0]['producer']} | {srcs[0]['model_name']}")

# Save comprehensive report
report = {
    'total_image_urls_checked': len(all_urls),
    'actually_checked': total_checked,
    'ok_count': ok_count,
    'broken_image_urls_count': broken_count,
    'non_image_urls_count': len(non_image),
    'html_urls_count': html_count,
    'small_files_count': len(small),
    'dveri_lord_sampled': sample_size,
    'dveri_lord_ok_rate': round(dveri_lord_sample_rate, 3),
    'dveri_lord_estimated_ok': dveri_lord_estimated_ok,
    'dveri_lord_estimated_errors': dveri_lord_estimated_err,
    'total_estimated_ok': ok_with_estimate,
    'total_estimated_broken': broken_count + dveri_lord_estimated_err,
    'cleared_image_urls_count': 7,
    'remaining_empty_image_rows_count': 59,
    'errors_sampled': [{'url': u, 'error': e} for u, e in all_errors],
    'non_image': [{'url': u, 'actual_type': at, 'content_type': ct, 'domain': urlparse(u).netloc} for u, at, ct in non_image],
    'affected_rows': {
        'broken_image_rows': [{'producer': p, 'model_name': m} for p, m in sorted(rows_broken)],
        'non_image_rows': [{'producer': p, 'model_name': m} for p, m in sorted(rows_non_img)],
    },
}

with open(os.path.join(OUT_DIR, 'image_url_validation_report.json'), 'w', encoding='utf-8') as f:
    json.dump(report, f, ensure_ascii=False, indent=2)
print(f"\nReport saved.")
