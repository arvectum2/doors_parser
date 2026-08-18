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

class SmartRedirectHandler(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        try:
            return urllib.request.Request(newurl, headers=dict(req.headers))
        except:
            return None

opener = urllib.request.build_opener(SmartRedirectHandler, urllib.request.HTTPSHandler(context=ctx))

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
        req = urllib.request.Request(url, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        resp = opener.open(req, timeout=10)
        ct = resp.headers.get('Content-Type', '')
        data = resp.read(16)
        actual_type = ''
        if data[:3] == b'\xff\xd8\xff':
            actual_type = 'image/jpeg'
        elif data[:4] == b'\x89PNG':
            actual_type = 'image/png'
        elif data[:4] == b'RIFF':
            actual_type = 'image/webp'
        elif data[:3] == b'GIF':
            actual_type = 'image/gif'
        else:
            actual_type = ct
        return url, resp.getcode(), actual_type, None
    except urllib.error.HTTPError as e:
        return url, e.code, '', f"HTTP {e.code}"
    except Exception as e:
        return url, 0, '', str(e)[:100]

start = time.time()
results = {}
errors = []

print(f"Checking with 40 threads, 10s timeout, redirect following...", flush=True)

with ThreadPoolExecutor(max_workers=40) as executor:
    futures = {executor.submit(check_url, url): url for url in all_urls}
    done = 0
    for future in as_completed(futures):
        done += 1
        if done % 500 == 0:
            print(f"  {done}/{len(all_urls)} ({time.time()-start:.1f}s)", flush=True)
        url, status, ct, err = future.result()
        results[url] = {'status': status, 'content_type': ct, 'ok': err is None}
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

broken_urls = set(e[0] for e in errors)
broken_rows = []
for r in rows:
    for f in ['image_url', 'image_url_front', 'image_url_inside']:
        if r.get(f, '').strip() in broken_urls:
            broken_rows.append((r['producer'], r['target_line'], r['model_name'], f, r.get(f,'').strip()))
            break

print(f"Rows with broken images: {len(broken_rows)}")

# Group by producer
prod_counts = Counter(p for p, l, m, f, u in broken_rows)
print("\nBroken by producer:")
for p, c in prod_counts.most_common():
    print(f"  {p}: {c}")

if http_404:
    print(f"\n404 URLs ({len(http_404)}):")
    for url, err in http_404[:30]:
        using = []
        for r in rows:
            for f in ['image_url', 'image_url_front', 'image_url_inside']:
                if url == r.get(f, '').strip():
                    using.append((r['producer'], r['target_line'], r['model_name'], f))
                    break
        print(f"  {url}")
        for prod, line, model, field in using[:2]:
            print(f"    -> {prod} | {line} | {model}")

if http_other:
    print(f"\nOther errors ({len(http_other)}):")
    err_types = Counter(e[1][:50] for e in http_other)
    for t, c in err_types.most_common():
        print(f"  {t}: {c}")

report = {
    'total_urls': len(all_urls),
    'ok': ok_count,
    'failed_404': len(http_404),
    'failed_other': len(http_other),
    'elapsed_seconds': round(elapsed, 1),
    'errors': [{'url': u, 'error': e} for u, e in errors],
    'broken_rows': [{'producer': p, 'target_line': l, 'model_name': m, 'field': f, 'url': u} for p, l, m, f, u in broken_rows],
}
with open('output/final_rework_ready_v5/image_url_validation_report.json', 'w', encoding='utf-8') as fh:
    json.dump(report, fh, ensure_ascii=False, indent=2)
print("\nReport saved.")
