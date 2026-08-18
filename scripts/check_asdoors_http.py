import urllib.request
import urllib.error
import ssl

# Try HTTP (not HTTPS)
urls = [
    "http://as-doors.ru/",
    "http://as-doors.ru/assets/cache_image/img/2026/ledo-t888/ledo-t888-2026-2-14.jpg",
    "http://as-doors.ru/assets/cache_image/img/2025/forest/forest-16.jpg",
    "http://as-doors.ru/assets/cache_image/img/2026/brig-t888/brig-t888-2026-2-17_1058x0_3a2.webp",
]

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))

for url in urls:
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'})
        resp = opener.open(req, timeout=15)
        ct = resp.headers.get('Content-Type', '')
        data = resp.read(4)
        is_img = data[:3] == b'\xff\xd8\xff' or data[:4] == b'\x89PNG' or data[:4] == b'RIFF'
        print(f"OK: {url[:70]} CT={ct} img={is_img}")
    except Exception as e:
        print(f"ERR: {url[:70]}: {str(e)[:60]}")
