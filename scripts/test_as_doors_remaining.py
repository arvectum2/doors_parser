import urllib.request
import ssl
import time
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

import csv
import os

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CATALOG = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")

with open(CATALOG, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

# Find as-doors.ru URLs that still exist (not cleared)
as_urls = set()
for r in rows:
    for field in ['image_url', 'image_url_front', 'image_url_inside']:
        url = r.get(field, '').strip()
        if url and 'as-doors.ru' in url:
            as_urls.add(url)

print(f"Remaining as-doors.ru URLs: {len(as_urls)}")

# Test first 20 with 30s timeout
test_count = min(20, len(as_urls))
test_urls = list(as_urls)[:test_count]
print(f"Testing {test_count} with 30s timeout...")

ok = 0
err_count = 0
for url in test_urls:
    start = time.time()
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
        resp = opener.open(req, timeout=30)
        ct = resp.headers.get('Content-Type', '')
        data = resp.read(4)
        is_img = data[:3] == b'\xff\xd8\xff' or data[:4] == b'\x89PNG' or data[:4] == b'RIFF'
        elapsed = time.time() - start
        if is_img:
            ok += 1
            print(f"  OK ({elapsed:.1f}s): {url[-60:]} CT={ct}")
        else:
            err_count += 1
            print(f"  NOT IMAGE ({elapsed:.1f}s): {url[-60:]} CT={ct}")
    except Exception as e:
        elapsed = time.time() - start
        err_count += 1
        print(f"  ERROR ({elapsed:.1f}s): {url[-60:]} {str(e)[:60]}")

print(f"\nResults: {ok} OK, {err_count} errors out of {test_count}")
print(f"Estimated for all {len(as_urls)}: ~{int(len(as_urls) * ok / test_count)} OK, ~{int(len(as_urls) * err_count / test_count)} errors")
