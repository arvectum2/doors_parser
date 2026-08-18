import csv
import os
import urllib.request
import urllib.error
import ssl
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
import json

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
v4_path = os.path.join(BASE, "output", "final_rework_ready_v4", "doors_catalog.csv")

with open(v4_path, encoding='utf-8-sig') as f:
    v4_rows = list(csv.DictReader(f, delimiter=';'))

v4_asd = [r for r in v4_rows if r['producer'] == 'АСД']
print(f"V4 ASD rows: {len(v4_asd)}")

# Collect unique front URLs from v4 ASD
v4_front_urls = set()
for r in v4_asd:
    url = r.get('image_url_front', '').strip()
    if url:
        v4_front_urls.add(url)

print(f"V4 ASD unique front URLs: {len(v4_front_urls)}")

# Test first 30 with 20s timeout
test_count = min(30, len(v4_front_urls))
test_urls = list(v4_front_urls)[:test_count]

def check_url(url, timeout=20):
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
        resp = opener.open(req, timeout=timeout)
        ct = resp.headers.get('Content-Type', '')
        data = resp.read(4)
        is_img = data[:3] == b'\xff\xd8\xff' or data[:4] == b'\x89PNG' or data[:4] == b'RIFF'
        final_url = resp.url
        return url, resp.getcode(), ct, is_img, final_url, None
    except urllib.error.HTTPError as e:
        return url, e.code, '', False, url, f"HTTP {e.code}"
    except Exception as e:
        return url, 0, '', False, url, str(e)[:100]

print(f"Testing {test_count} URLs with 20s timeout...")
ok = 0
redirect = 0
errors = []
with ThreadPoolExecutor(max_workers=10) as executor:
    futures = {executor.submit(check_url, u, 20): u for u in test_urls}
    for future in as_completed(futures):
        url, status, ct, is_img, final_url, err = future.result()
        if err:
            errors.append((url, err, final_url))
            print(f"  ERROR: {status} {url[-60:]} -> {err[:40]}")
        elif is_img:
            ok += 1
            print(f"  OK: {url[-60:]} CT={ct}")
        else:
            redirect += 1
            print(f"  NOT IMG: {url[-60:]} -> {final_url[-60:]} CT={ct}")

print(f"\nResults: {ok} OK, {redirect} not-image, {len(errors)} errors")
print(f"Estimated for all {len(v4_front_urls)}: ~{int(len(v4_front_urls) * ok / test_count)} OK")

# Save all v4 ASD URLs with their status
print(f"\nSaving all v4 ASD front URLs for bulk restore...")
all_v4_asd = []
for r in v4_asd:
    all_v4_asd.append({
        'model_name': r['model_name'],
        'target_line': r['target_line'],
        'image_url_front': r.get('image_url_front', ''),
        'image_url_inside': r.get('image_url_inside', ''),
        'source_url': r.get('source_url', ''),
    })

# Save for restore script
with open(os.path.join(BASE, "output", "delivery", "v4_asd_urls.json"), 'w', encoding='utf-8') as f:
    json.dump(all_v4_asd, f, ensure_ascii=False, indent=2)

print(f"Saved {len(all_v4_asd)} ASD rows to v4_asd_urls.json")
