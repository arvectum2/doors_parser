import re
import json
import urllib.request
import ssl

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

def fetch(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    resp = urllib.request.urlopen(req, timeout=15, context=ctx)
    return resp.read().decode('utf-8', errors='replace')

html = fetch("https://komfortdoors.ru/emal/?page=1")

# Save raw HTML for analysis
with open('output/final_rework_ready_v5/komfort_page1.html', 'w', encoding='utf-8') as f:
    f.write(html)

# Find all cache_image URLs
cache_urls = re.findall(r'(/assets/cache_image/products/\d+/[^"\'<>\s]+)', html)
print(f"Found {len(cache_urls)} cache_image URLs")
for u in cache_urls[:10]:
    print(f"  {u}")

# Find all product links
product_links = re.findall(r'href="((?:emal/)?(?:new-line|slope|true|gift|art|premier|malta|neo|molding|turin)/[^"]+)"', html)
print(f"\nFound {len(product_links)} product links")
for l in product_links[:10]:
    print(f"  {l}")

# Find img tags
img_tags = re.findall(r'<img[^>]+>', html)
print(f"\nFound {len(img_tags)} img tags")
for t in img_tags[:5]:
    print(f"  {t[:200]}")

# Try broader pattern for product blocks
# Look for all links within the product grid
product_blocks = re.findall(r'<a[^>]*href="([^"]*(?:new-line|slope|true|gift|art|premier|malta|neo|molding|turin)/[^"]*)"[^>]*>.*?</a>', html, re.DOTALL)
print(f"\nProduct blocks (broad): {len(product_blocks)}")
