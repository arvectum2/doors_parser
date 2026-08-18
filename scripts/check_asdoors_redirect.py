import urllib.request
import urllib.error
import ssl

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

class RedirectCatcher(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        print(f"  REDIRECT {code} -> {newurl}")
        return None

opener = urllib.request.build_opener(RedirectCatcher, urllib.request.HTTPSHandler(context=ctx))

urls = [
    "https://as-doors.ru/",
    "https://as-doors.ru/assets/cache_image/img/2026/ledo-t888/ledo-t888-2026-2-14.jpg",
]

for url in urls:
    print(f"\nURL: {url}")
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
        resp = opener.open(req, timeout=15)
        print(f"  Status: {resp.getcode()}")
    except urllib.error.HTTPError as e:
        location = e.headers.get('Location', 'NONE') if hasattr(e, 'headers') else 'NO HEADERS'
        print(f"  HTTPError {e.code}: Location={location}")
        # Try to read response body
        try:
            body = e.read(500).decode('utf-8', errors='ignore')
            print(f"  Body: {body[:200]}")
        except:
            pass
    except Exception as e:
        print(f"  Error: {str(e)[:100]}")
