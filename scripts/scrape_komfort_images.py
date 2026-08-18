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

# Scrape all 3 catalog pages for Emal
all_products = {}
for page in range(1, 4):
    url = f"https://komfortdoors.ru/emal/?page={page}"
    print(f"Fetching {url}...")
    html = fetch(url)
    
    # Pattern: <img src="/assets/cache_image/products/{id}/{filename}_0x0_9d7.webp" ...>
    # Followed by link to product page
    # Find all product blocks: img + title + link
    pattern = r'<img\s+[^>]*src="(/assets/cache_image/products/\d+/[^"]+_0x0_9d7\.webp)"[^>]*>\s*</a>\s*<a[^>]*href="([^"]+)"[^>]*>\s*([^<]+)</a>'
    matches = re.findall(pattern, html)
    
    print(f"  Found {len(matches)} products on page {page}")
    for img_path, link, title in matches:
        title = title.strip()
        img_url = f"https://komfortdoors.ru{img_path}"
        # Extract the URL slug from link
        slug = link.strip('/')
        # Extract product ID from img path
        prod_id = re.search(r'/products/(\d+)/', img_path).group(1)
        all_products[title] = {'img_url': img_url, 'slug': slug, 'prod_id': prod_id}
        print(f"  {title} -> {img_url}")

print(f"\nTotal products scraped: {len(all_products)}")

# Save to JSON
with open('output/final_rework_ready_v5/komfort_images.json', 'w', encoding='utf-8') as f:
    json.dump(all_products, f, ensure_ascii=False, indent=2)
print("Saved to komfort_images.json")
