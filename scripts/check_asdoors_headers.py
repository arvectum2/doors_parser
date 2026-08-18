import urllib.request
import urllib.error
import ssl
import time

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

# Try different approaches
urls = [
    "https://as-doors.ru/",
    "https://as-doors.ru/assets/cache_image/img/2026/ledo-t888/ledo-t888-2026-2-14.jpg",
    "https://as-doors.ru/assets/cache_image/img/2025/forest/forest-16.jpg",
]

headers_sets = [
    {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'},
    {'User-Agent': 'Mozilla/5.0'},
    {'User-Agent': 'Mozilla/5.0', 'Accept': 'image/webp,image/apng,image/*,*/*;q=0.8'},
    {'User-Agent': 'Mozilla/5.0', 'Accept': '*/*', 'Referer': 'https://as-doors.ru/'},
]

for url in urls:
    print(f"\nURL: {url}")
    for i, headers in enumerate(headers_sets):
        try:
            req = urllib.request.Request(url, headers=headers)
            resp = urllib.request.urlopen(req, timeout=15, context=ctx)
            ct = resp.headers.get('Content-Type', '')
            print(f"  Headers {i}: OK CT={ct}")
            break
        except Exception as e:
            print(f"  Headers {i}: {str(e)[:50]}")
        time.sleep(0.5)
