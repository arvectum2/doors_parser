import urllib.request
import urllib.error
import ssl
import re
import json
from html.parser import HTMLParser

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))

class ImageExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.images = []
    def handle_starttag(self, tag, attrs):
        if tag == 'img':
            d = dict(attrs)
            src = d.get('src', '') or d.get('data-src', '')
            if src:
                self.images.append(src)

# Try a few product pages via HTTP to find inside images
test_pages = [
    ("Forest", "http://as-doors.ru/onstock/forest.html"),
    ("Ferrum", "http://as-doors.ru/doors/zakaznyie/stalnaya-dver-ferrum.html"),
    ("Britannia", "http://as-doors.ru/doors/zakaznyie/britannia.html"),
    ("Vegas", "http://as-doors.ru/doors/zakaznyie/vegas.html"),
]

for name, url in test_pages:
    print(f"\n=== {name}: {url} ===")
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
        resp = opener.open(req, timeout=15)
        ct = resp.headers.get('Content-Type', '')
        if 'text/html' not in ct:
            print(f"  Not HTML: {ct}")
            continue
        html = resp.read().decode('utf-8', errors='ignore')
        
        # Extract all images
        extractor = ImageExtractor()
        extractor.feed(html)
        
        # Filter for door images
        door_imgs = [img for img in extractor.images if any(x in img.lower() for x in ['jpg', 'png', 'webp', 'img/', 'images/'])]
        print(f"  Total images: {len(extractor.images)}, door images: {len(door_imgs)}")
        
        # Show first 10
        for img in door_imgs[:10]:
            print(f"    {img[:80]}")
            
    except Exception as e:
        print(f"  ERROR: {str(e)[:80]}")
