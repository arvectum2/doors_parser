import csv
import json
import os
import urllib.request
import urllib.error
import ssl
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import Counter
from urllib.parse import urlparse, quote

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

# Find URLs with non-ASCII characters
ascii_problem_urls = set()
for r in rows:
    for field in ['image_url', 'image_url_front', 'image_url_inside']:
        url = r.get(field, '').strip()
        if url:
            try:
                url.encode('ascii')
            except UnicodeEncodeError:
                ascii_problem_urls.add(url)

print(f"URLs with non-ASCII characters: {len(ascii_problem_urls)}")

def encode_url(url):
    """Properly encode non-ASCII characters in URL."""
    parsed = urlparse(url)
    # Encode the path component
    path_encoded = quote(parsed.path, safe='/:@!$&\'()*+,;=')
    # Reconstruct URL
    encoded_url = f"{parsed.scheme}://{parsed.netloc}{path_encoded}"
    if parsed.query:
        encoded_url += f"?{parsed.query}"
    return encoded_url

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
        return url, resp.getcode(), ct, cl, actual_type, is_image, is_html, None, url
    except urllib.error.HTTPError as e:
        return url, e.code, '', '0', '', False, False, f"HTTP {e.code}", url
    except Exception as e:
        return url, 0, '', '0', '', False, False, str(e)[:100], url

# Test encoding fix on ASCII problem URLs
print("\n=== Testing ASCII encoding fix ===")
fixed_results = {}
for url in list(ascii_problem_urls)[:10]:
    encoded = encode_url(url)
    print(f"  Original: {url[:80]}...")
    print(f"  Encoded:  {encoded[:80]}...")
    try:
        req = urllib.request.Request(encoded, headers={'User-Agent': 'Mozilla/5.0'})
        resp = opener.open(req, timeout=10)
        ct = resp.headers.get('Content-Type', '')
        data = resp.read(4)
        is_img = data[:3] == b'\xff\xd8\xff' or data[:4] == b'\x89PNG' or data[:4] == b'RIFF'
        print(f"  Result: {resp.getcode()} CT={ct} IsImage={is_img}")
        fixed_results[url] = {'encoded': encoded, 'ok': True, 'ct': ct, 'is_img': is_img}
    except Exception as e:
        print(f"  Error: {str(e)[:60]}")
        fixed_results[url] = {'encoded': encoded, 'ok': False, 'error': str(e)[:60]}

ok_fixes = sum(1 for v in fixed_results.values() if v.get('ok'))
print(f"\nFixed {ok_fixes}/{len(fixed_results)} ASCII URLs")

# Now apply fix to all ASCII problem URLs in the CSV
print("\n=== Applying fixes to CSV ===")
fixed_count = 0
for r in rows:
    for field in ['image_url', 'image_url_front', 'image_url_inside']:
        url = r.get(field, '').strip()
        if url and url in ascii_problem_urls:
            try:
                url.encode('ascii')
            except UnicodeEncodeError:
                encoded = encode_url(url)
                # Test if encoded URL works
                try:
                    req = urllib.request.Request(encoded, headers={'User-Agent': 'Mozilla/5.0'})
                    resp = opener.open(req, timeout=10)
                    ct = resp.headers.get('Content-Type', '')
                    data = resp.read(4)
                    is_img = data[:3] == b'\xff\xd8\xff' or data[:4] == b'\x89PNG' or data[:4] == b'RIFF'
                    if is_img:
                        r[field] = encoded
                        fixed_count += 1
                        print(f"  FIXED: {r['model_name']} | {field}")
                    else:
                        r[field] = ''
                        print(f"  CLEARED: {r['model_name']} | {field} (not image: {ct})")
                except:
                    r[field] = ''
                    print(f"  CLEARED: {r['model_name']} | {field} (encoded URL failed)")

print(f"\nFixed {fixed_count} ASCII URLs in CSV")

# Save updated CSV
fieldnames = list(rows[0].keys())
with open(CATALOG, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=';')
    writer.writeheader()
    writer.writerows(rows)
print("CSV updated")
