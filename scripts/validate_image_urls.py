import csv
import urllib.request
import urllib.error
import ssl
import time
import json
from collections import Counter

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

with open('output/final_rework_ready_v5/doors_catalog.csv', encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

# Collect all unique image URLs
all_urls = set()
for r in rows:
    for field in ['image_url', 'image_url_front', 'image_url_inside']:
        url = r.get(field, '').strip()
        if url:
            all_urls.add(url)

print(f"Total rows: {len(rows)}")
print(f"Unique image URLs to check: {len(all_urls)}")

results = {}
errors = []
checked = 0

for url in sorted(all_urls):
    checked += 1
    if checked % 50 == 0:
        print(f"  Checked {checked}/{len(all_urls)}...")
    try:
        req = urllib.request.Request(url, method='HEAD', headers={'User-Agent': 'Mozilla/5.0'})
        resp = urllib.request.urlopen(req, timeout=10, context=ctx)
        ct = resp.headers.get('Content-Type', '')
        size = resp.headers.get('Content-Length', '0')
        status = resp.getcode()
        results[url] = {'status': status, 'content_type': ct, 'size': size, 'ok': True}
    except urllib.error.HTTPError as e:
        results[url] = {'status': e.code, 'content_type': '', 'size': '0', 'ok': False}
        errors.append((url, f"HTTP {e.code}"))
    except Exception as e:
        results[url] = {'status': 0, 'content_type': '', 'size': '0', 'ok': False}
        errors.append((url, str(e)[:80]))
    time.sleep(0.05)

print(f"\n=== RESULTS ===")
print(f"Checked: {checked}")
ok_count = sum(1 for v in results.values() if v['ok'])
print(f"OK (reachable): {ok_count}")
print(f"Failed: {len(errors)}")

# Group errors
http_404 = [e for e in errors if '404' in e[1]]
http_other = [e for e in errors if '404' not in e[1]]
print(f"\n404 errors: {len(http_404)}")
print(f"Other errors: {len(http_other)}")

if http_404:
    print("\n404 URLs:")
    for url, err in http_404[:20]:
        # Find which rows use this URL
        using = [(r['producer'], r['model_name'], r.get('image_url',''), r.get('image_url_front',''), r.get('image_url_inside','')) for r in rows if url in [r.get('image_url',''), r.get('image_url_front',''), r.get('image_url_inside','')]]
        print(f"  {url}")
        for prod, model, u, f, i in using[:3]:
            field = 'image_url' if url == u else ('image_url_front' if url == f else 'image_url_inside')
            print(f"    -> {prod} | {model} | {field}")

if http_other:
    print("\nOther errors:")
    for url, err in http_other[:20]:
        print(f"  {url}: {err}")

# Non-image content types
non_image = [(url, v['content_type']) for url, v in results.items() if v['ok'] and not v['content_type'].startswith('image/')]
print(f"\nNon-image content-type: {len(non_image)}")
for url, ct in non_image[:10]:
    print(f"  {url}: {ct}")

# Small files (likely placeholders)
small = [(url, v['size']) for url, v in results.items() if v['ok'] and v['size'].isdigit() and int(v['size']) < 1000]
print(f"\nSmall files (<1KB): {len(small)}")
for url, sz in small[:10]:
    print(f"  {url}: {sz} bytes")

# Save full report
report = {
    'total_urls': len(all_urls),
    'checked': checked,
    'ok': ok_count,
    'failed_404': len(http_404),
    'failed_other': len(http_other),
    'non_image_ct': len(non_image),
    'small_files': len(small),
    'errors': [{'url': u, 'error': e} for u, e in errors],
    'non_image': [{'url': u, 'ct': ct} for u, ct in non_image],
    'small': [{'url': u, 'size': s} for u, s in small],
}
with open('output/final_rework_ready_v5/image_url_validation_report.json', 'w', encoding='utf-8') as f:
    json.dump(report, f, ensure_ascii=False, indent=2)
print(f"\nReport saved to output/final_rework_ready_v5/image_url_validation_report.json")
