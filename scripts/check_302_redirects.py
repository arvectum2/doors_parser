import csv
import os
import urllib.request
import urllib.error
import ssl
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import Counter

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

# Follow redirects manually
class DebugRedirectHandler(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None  # Don't follow - we want to see the redirect

opener = urllib.request.build_opener(DebugRedirectHandler, urllib.request.HTTPSHandler(context=ctx))

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
v4_path = os.path.join(BASE, "output", "final_rework_ready_v4", "doors_catalog.csv")

with open(v4_path, encoding='utf-8-sig') as f:
    v4_rows = list(csv.DictReader(f, delimiter=';'))

v4_asd = [r for r in v4_rows if r['producer'] == 'АСД']

# Test 20 URLs that returned 302 to see where they redirect
test_urls = []
for r in v4_asd[:50]:
    url = r.get('image_url_front', '').strip()
    if url:
        test_urls.append((r['model_name'], url))

print(f"Testing 302 redirect destinations for {min(20, len(test_urls))} URLs...")

redirect_targets = []
for name, url in test_urls[:20]:
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
        resp = opener.open(req, timeout=15)
        location = resp.headers.get('Location', 'NO LOCATION')
        status = resp.getcode()
        print(f"  {status} -> {location[-80:]}")
        redirect_targets.append((name, url, location, status))
    except urllib.error.HTTPError as e:
        location = e.headers.get('Location', 'NO LOCATION') if hasattr(e, 'headers') else 'NO HEADERS'
        print(f"  {e.code} -> {location[-80:]}")
        redirect_targets.append((name, url, location, e.code))
    except Exception as e:
        print(f"  ERROR: {str(e)[:60]}")
        redirect_targets.append((name, url, 'ERROR', 0))

# Analyze redirect patterns
print(f"\nRedirect target analysis:")
patterns = Counter()
for name, url, loc, status in redirect_targets:
    if loc.startswith('http'):
        from urllib.parse import urlparse
        parsed = urlparse(loc)
        patterns[f"{parsed.netloc}{parsed.path[:30]}"] += 1
    else:
        patterns[loc[:50]] += 1

for pattern, count in patterns.most_common(10):
    print(f"  {count}x: {pattern}")

# Check if the redirect targets work
print(f"\nFollowing redirects to find working URLs...")
working = 0
for name, url, loc, status in redirect_targets:
    if loc.startswith('http'):
        try:
            req = urllib.request.Request(loc, headers={'User-Agent': 'Mozilla/5.0'})
            resp = opener.open(req, timeout=15)
            ct = resp.headers.get('Content-Type', '')
            data = resp.read(4)
            is_img = data[:3] == b'\xff\xd8\xff' or data[:4] == b'\x89PNG' or data[:4] == b'RIFF'
            if is_img:
                working += 1
                print(f"  WORKING: {url[-50:]} -> {loc[-50:]} CT={ct}")
            else:
                print(f"  NOT IMG: {url[-50:]} -> {loc[-50:]} CT={ct}")
        except Exception as e:
            print(f"  FAILED: {url[-50:]} -> {str(e)[:40]}")

print(f"\nWorking redirect targets: {working}/{len(redirect_targets)}")
