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

# Test dveri-lord.ru with 30s timeout
print("=== dveri-lord.ru test (30s timeout) ===")
test_urls_dveri = [
    "https://dveri-lord.ru/wp-content/uploads/2023/06/altro-1f.png",
    "https://dveri-lord.ru/wp-content/uploads/2023/06/altro-mf1.png",
    "https://dveri-lord.ru/wp-content/uploads/2025/10/avanti-1-nalichnik-teleskop-italy-antiskrejch-milk-800h2000.png",
]
for url in test_urls_dveri:
    start = time.time()
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
        resp = opener.open(req, timeout=30)
        ct = resp.headers.get('Content-Type', '')
        cl = resp.headers.get('Content-Length', '0')
        data = resp.read(4)
        is_img = data[:3] == b'\xff\xd8\xff' or data[:4] == b'\x89PNG' or data[:4] == b'RIFF'
        elapsed = time.time() - start
        print(f"  OK ({elapsed:.1f}s): {url[-50:]} CT={ct} CL={cl} IsImage={is_img}")
    except Exception as e:
        elapsed = time.time() - start
        print(f"  ERROR ({elapsed:.1f}s): {url[-50:]} {str(e)[:60]}")

# Test as-doors.ru with proper redirect following
print("\n=== as-doors.ru test (with redirect following) ===")
test_urls_as = [
    "https://as-doors.ru/images/2023/avant/avant-z1-35.jpg",
    "https://as-doors.ru/onstock/stalnaya-dver-everest.html/img/everest22/everest4.jpg",
    "https://as-doors.ru/doors/avanti/avanti-1.jpg",
]
for url in test_urls_as:
    start = time.time()
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
        resp = opener.open(req, timeout=15)
        ct = resp.headers.get('Content-Type', '')
        cl = resp.headers.get('Content-Length', '0')
        data = resp.read(4)
        is_img = data[:3] == b'\xff\xd8\xff' or data[:4] == b'\x89PNG' or data[:4] == b'RIFF'
        elapsed = time.time() - start
        print(f"  OK ({elapsed:.1f}s): {url[-60:]} CT={ct} CL={cl} IsImage={is_img}")
    except Exception as e:
        elapsed = time.time() - start
        print(f"  ERROR ({elapsed:.1f}s): {url[-60:]} {str(e)[:60]}")

# Test as-doors.ru - find working image URLs
print("\n=== as-doors.ru - testing various image paths ===")
test_urls_as2 = [
    "https://as-doors.ru/images/2023/avant/avant-z1-1.jpg",
    "https://as-doors.ru/images/2023/avant/avant-z1-2.jpg",
    "https://as-doors.ru/images/2023/avant/avant-1.jpg",
    "https://as-doors.ru/images/2023/avant/avant-2.jpg",
]
for url in test_urls_as2:
    start = time.time()
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        resp = opener.open(req, timeout=15)
        ct = resp.headers.get('Content-Type', '')
        data = resp.read(4)
        is_img = data[:3] == b'\xff\xd8\xff' or data[:4] == b'\x89PNG' or data[:4] == b'RIFF'
        elapsed = time.time() - start
        print(f"  OK ({elapsed:.1f}s): {url[-60:]} CT={ct} IsImage={is_img}")
    except Exception as e:
        elapsed = time.time() - start
        print(f"  ERROR ({elapsed:.1f}s): {url[-60:]} {str(e)[:60]}")

# Test entro-design.ru
print("\n=== entro-design.ru test ===")
test_entro = [
    "https://entro-design.ru/upload/iblock/7e6/7e6a5b5b5b5b5b5b5b5b5b5b5b5b5b5b.jpg",
]
for url in test_entro:
    start = time.time()
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        resp = opener.open(req, timeout=15)
        ct = resp.headers.get('Content-Type', '')
        data = resp.read(4)
        is_img = data[:3] == b'\xff\xd8\xff' or data[:4] == b'\x89PNG' or data[:4] == b'RIFF'
        elapsed = time.time() - start
        print(f"  OK ({elapsed:.1f}s): CT={ct} IsImage={is_img}")
    except Exception as e:
        elapsed = time.time() - start
        print(f"  ERROR ({elapsed:.1f}s): {str(e)[:60]}")
