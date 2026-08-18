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

print(f"Total unique URLs: {len(all_urls)}", flush=True)

def check_url(url):
    try:
        req = urllib.request.Request(url, method='GET', headers={
            'User-Agent': 'Mozilla/5.0',
            'Range': 'bytes=0-1023'
        })
        resp = urllib.request.urlopen(req, timeout=5, context=ctx)
        ct = resp.headers.get('Content-Type', '')
        cl = resp.headers.get('Content-Length', resp.headers.get('Content-Range', '0'))
        return url, resp.getcode(), ct, str(cl), None
    except urllib.error.HTTPError as e:
        return url, e.code, '', '0', f"HTTP {e.code}"
    except Exception as e:
        return url, 0, '', '0', str(e)[:80]

start = time.time()
results = {}
errors = []

print(f"Checking with 40 threads, 5s timeout...", flush=True)

with ThreadPoolExecutor(max_workers=40) as executor:
    futures = {executor.submit(check_url, url): url for url in all_urls}
    done = 0
    for future in as_completed(futures):
        done += 1
        if done % 500 == 0:
            print(f"  {done}/{len(all_urls)} ({time.time()-start:.1f}s)", flush=True)
        url, status, ct, size, err = future.result()
        results[url] = {'status': status, 'content_type': ct, 'size': size, 'ok': err is None}
        if err:
            errors.append((url, err))

elapsed = time.time() - start
ok_count = sum(1 for v in results.values() if v['ok'])
http_404 = [e for e in errors if '404' in e[1]]
http_other = [e for e in errors if '404' not in e[1]]

print(f"\nDone in {elapsed:.1f}s", flush=True)
print(f"OK: {ok_count}/{len(all_urls)}")
print(f"404: {len(http_404)}")
print(f"Other errors: {len(http_other)}")

non_image = [(url, v['content_type']) for url, v in results.items() if v['ok'] and not v['content_type'].startswith('image/')]
print(f"Non-image CT: {len(non_image)}")

broken_urls = set(e[0] for e in errors)
broken_rows = []
for r in rows:
    for f in ['image_url', 'image_url_front', 'image_url_inside']:
        if r.get(f, '').strip() in broken_urls:
            broken_rows.append((r['producer'], r['target_line'], r['model_name'], f, r.get(f,'').strip()))
            break

print(f"Rows with broken images: {len(broken_rows)}")
for prod, line, model, field, url in broken_rows[:20]:
    print(f"  {prod} | {line} | {model} | {field}")
    print(f"    {url}")

report = {
    'total_urls': len(all_urls),
    'ok': ok_count,
    'failed_404': len(http_404),
    'failed_other': len(http_other),
    'non_image_ct': len(non_image),
    'elapsed_seconds': round(elapsed, 1),
    'errors': [{'url': u, 'error': e} for u, e in errors],
    'non_image': [{'url': u, 'ct': ct} for u, ct in non_image],
    'broken_rows': [{'producer': p, 'target_line': l, 'model_name': m, 'field': f, 'url': u} for p, l, m, f, u in broken_rows],
}
with open('output/final_rework_ready_v5/image_url_validation_report.json', 'w', encoding='utf-8') as fh:
    json.dump(report, fh, ensure_ascii=False, indent=2)
print("Report saved.")
