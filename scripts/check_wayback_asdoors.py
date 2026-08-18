import urllib.request
import urllib.error
import ssl
import re
import json
import time

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))

# Try Wayback Machine
test_urls = [
    ("Forest", "https://as-doors.ru/onstock/forest.html"),
    ("Ferrum", "https://as-doors.ru/doors/zakaznyie/stalnaya-dver-ferrum.html"),
]

for name, url in test_urls:
    # Check Wayback Machine availability
    wb_url = f"https://archive.org/wayback/available?url={url}"
    print(f"\n=== {name} ===")
    print(f"Checking Wayback: {wb_url}")
    try:
        req = urllib.request.Request(wb_url, headers={'User-Agent': 'Mozilla/5.0'})
        resp = opener.open(req, timeout=15)
        data = json.loads(resp.read())
        snapshots = data.get('archived_snapshots', {})
        if snapshots.get('closest'):
            snap = snapshots['closest']
            print(f"  Found: {snap['url']} (timestamp={snap['timestamp']})")
            
            # Fetch the archived page
            archive_resp_url = snap['url']
            req2 = urllib.request.Request(archive_resp_url, headers={'User-Agent': 'Mozilla/5.0'})
            resp2 = opener.open(req2, timeout=15)
            html = resp2.read().decode('utf-8', errors='ignore')
            
            # Find images
            imgs = re.findall(r'<img[^>]+src="([^"]+)"', html, re.I)
            door_imgs = [img for img in imgs if any(x in img.lower() for x in ['jpg', 'png', 'webp', 'img/', 'images/'])]
            print(f"  Images found: {len(door_imgs)}")
            for img in door_imgs[:10]:
                print(f"    {img[:80]}")
        else:
            print(f"  No snapshots found")
    except Exception as e:
        print(f"  ERROR: {str(e)[:80]}")
    time.sleep(1)
