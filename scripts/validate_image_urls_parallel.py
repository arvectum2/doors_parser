import csv
import json
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

with open('output/final_rework_ready_v5/doors_catalog.csv', encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

all_urls = set()
for r in rows:
    for field in ['image_url', 'image_url_front', 'image_url_inside']:
        url = r.get(field, '').strip()
        if url:
            all_urls.add(url)

domains = Counter(urlparse(u).netloc for u in all_urls)
print(f"Total unique URLs: {len(all_urls)}")
for d, c in domains.most_common():
    print(f"  {d}: {c}")

def check_url(url):
    try:
        req = urllib.request.Request(url, method='HEAD', headers={'User-Agent': 'Mozilla/5.0'})
        resp = urllib.request.urlopen(req, timeout=10, context=ctx)
        ct = resp.headers.get('Content-Type', '')
        size = resp.headers.get('Content-Length', '0')
        return url, resp.getcode(), ct, size, None
    except urllib.error.HTTPError as e:
        return url, e.code, '', '0', f"HTTP {e.code}"
    except Exception as e:
        return url, 0, '', '0', str(e)[:100]

print(f"\nChecking {len(all_urls)} URLs with 20 threads...")
start = time.time()
results = {}
errors = []

with ThreadPoolExecutor(max_workers=20) as executor:
    futures = {executor.submit(check_url, url): url for url in all_urls}
    done = 0
    for future in as_completed(futures):
        done += 1
        if done % 200 == 0:
            print(f"  {done}/{len(all_urls)} checked ({time.time()-start:.1f}s)")
        url, status, ct, size, err = future.result()
        results[url] = {'status': status, 'content_type': ct, 'size': size, 'ok': err is None}
        if err:
            errors.append((url, err))

elapsed = time.time() - start
print(f"\nDone in {elapsed:.1f}s")

ok_count = sum(1 for v in results.values() if v['ok'])
http_404 = [e for e in errors if '404' in e[1]]
http_other = [e for e in errors if '404' not in e[1]]

print(f"\nOK (reachable): {ok_count}/{len(all_urls)}")
print(f"404 errors: {len(http_404)}")
print(f"Other errors: {len(http_other)}")

non_image = [(url, v['content_type']) for url, v in results.items() if v['ok'] and not v['content_type'].startswith('image/')]
small = [(url, v['size']) for url, v in results.items() if v['ok'] and v['size'].isdigit() and int(v['size']) < 1000]

print(f"Non-image content-type: {len(non_image)}")
print(f"Small files (<1KB): {len(small)}")

if http_404:
    print("\n404 URLs (first 20):")
    for url, err in http_404[:20]:
        using = []
        for r in rows:
            for f in ['image_url', 'image_url_front', 'image_url_inside']:
                if url == r.get(f, '').strip():
                    using.append((r['producer'], r['model_name'], f))
                    break
        print(f"  {url}")
        for prod, model, field in using[:2]:
            print(f"    -> {prod} | {model} | {field}")

if http_other:
    print("\nOther errors (first 20):")
    for url, err in http_other[:20]:
        print(f"  {url}: {err}")

if non_image:
    print("\nNon-image content-type (first 10):")
    for url, ct in non_image[:10]:
        print(f"  {url}: {ct}")

if small:
    print("\nSmall files <1KB (first 10):")
    for url, sz in small[:10]:
        print(f"  {url}: {sz} bytes")

# Find which rows are affected
broken_urls = set(e[0] for e in errors)
broken_rows = []
for r in rows:
    for f in ['image_url', 'image_url_front', 'image_url_inside']:
        if r.get(f, '').strip() in broken_urls:
            broken_rows.append(r)
            break

print(f"\nRows with broken images: {len(broken_rows)}")
for r in broken_rows[:10]:
    print(f"  {r['producer']} | {r['target_line']} | {r['model_name']}")

report = {
    'total_urls': len(all_urls),
    'ok': ok_count,
    'failed_404': len(http_404),
    'failed_other': len(http_other),
    'non_image_ct': len(non_image),
    'small_files': len(small),
    'elapsed_seconds': round(elapsed, 1),
    'errors': [{'url': u, 'error': e} for u, e in errors],
    'non_image': [{'url': u, 'ct': ct} for u, ct in non_image],
    'small': [{'url': u, 'size': s} for u, s in small],
    'broken_rows': [{'producer': r['producer'], 'target_line': r['target_line'], 'model_name': r['model_name']} for r in broken_rows],
}
with open('output/final_rework_ready_v5/image_url_validation_report.json', 'w', encoding='utf-8') as f:
    json.dump(report, f, ensure_ascii=False, indent=2)
print(f"\nReport saved.")
