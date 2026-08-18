import urllib.request
import urllib.error
import ssl
import re

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))

# Test: try to access as-doors.ru and find product pages
test_urls = [
    "https://as-doors.ru/doors/zakaznyie/forest.html",
    "https://as-doors.ru/onstock/forest.html",
    "https://as-doors.ru/assets/cache_image/img/2025/forest/forest-16.jpg",
    "https://as-doors.ru/assets/cache_image/img/2026/ledo-t888/ledo-t888-2026-2-14.jpg",
    "https://as-doors.ru/",
]

for url in test_urls:
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
        resp = opener.open(req, timeout=15)
        ct = resp.headers.get('Content-Type', '')
        final_url = resp.url
        print(f"OK: {url[:60]}")
        print(f"  Final: {final_url[:80]}")
        print(f"  CT: {ct}")
        if 'text/html' in ct:
            html = resp.read(5000).decode('utf-8', errors='ignore')
            # Find og:image or product image
            og_match = re.search(r'og:image["\s]+content="([^"]+)"', html)
            if og_match:
                print(f"  og:image: {og_match.group(1)[:80]}")
            # Find first product image
            img_match = re.search(r'<img[^>]+src="([^"]*(?:jpg|png|webp)[^"]*)"', html, re.I)
            if img_match:
                print(f"  First img: {img_match.group(1)[:80]}")
        print()
    except Exception as e:
        print(f"ERROR: {url[:60]}: {str(e)[:60]}")
        print()
