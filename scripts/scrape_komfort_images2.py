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

all_products = {}

for page in range(1, 4):
    url = f"https://komfortdoors.ru/emal/?page={page}"
    print(f"Fetching {url}...")
    html = fetch(url)
    
    # Match markdown-style pattern in HTML: [![title](img_url "title")](link)
    # The HTML has: <a href="link"><img src="/assets/cache_image/..." title="Title" alt="Title"></a>
    # <a href="link"><img src="/assets/cache_image/products/XXX/FILENAME_0x0_9d7.webp" title="TITLE" alt="TITLE"></a>
    pattern = r'<a\s+href="((?:new-line|slope|true|gift|art|premier|malta|neo|molding|turin)/[^"]+)"[^>]*>\s*<img\s+[^>]*src="(/assets/cache_image/products/\d+/[^"]+_0x0_9d7\.webp)"[^>]*title="([^"]*)"'
    matches = re.findall(pattern, html)
    
    if not matches:
        # Try alternative: img first, then link
        pattern2 = r'<img\s+[^>]*src="(/assets/cache_image/products/\d+/[^"]+_0x0_9d7\.webp)"[^>]*title="([^"]*)"'
        pattern3 = r'<a\s+href="((?:new-line|slope|true|gift|art|premier|malta|neo|molding|turin)/[^"]+)"'
        
        imgs = re.findall(pattern2, html)
        links = re.findall(pattern3, html)
        
        print(f"  Page {page}: {len(imgs)} images, {len(links)} links")
        
        # Try to pair them (they should be in order: link, img, link, img...)
        # Actually let's try a more flexible pattern
        # Find all blocks that have both img and link
        blocks = re.findall(
            r'<a\s+href="((?:new-line|slope|true|gift|art|premier|malta|neo|molding|turin)/[^"]+)"[^>]*>\s*(?:<[^>]+>\s*)*<img\s+[^>]*src="(/assets/cache_image/products/\d+/[^"]+_0x0_9d7\.webp)"',
            html, re.DOTALL
        )
        matches = blocks
        if not matches:
            # Even more flexible: find all a tags with product links that contain img
            # Let's parse the raw HTML differently
            # Split by product link anchors
            parts = re.split(r'(?=<a\s+href="(new-line|slope|true|gift|art|premier|malta|neo|molding|turin)/)', html)
            for i in range(1, len(parts), 2):
                link_path = re.match(r'([^/]+/[^"]+)', parts[i])
                if link_path:
                    full_link = link_path.group(1)
                    img_match = re.search(r'src="(/assets/cache_image/products/\d+/[^"]+_0x0_9d7\.webp)"', parts[i+1][:2000] if i+1 < len(parts) else '')
                    if img_match:
                        title_match = re.search(r'title="([^"]+)"', parts[i+1][:2000] if i+1 < len(parts) else '')
                        title = title_match.group(1) if title_match else full_link.split('/')[-1]
                        img_url = f"https://komfortdoors.ru{img_match.group(1)}"
                        all_products[title] = {'img_url': img_url, 'slug': full_link}
            continue
    
    print(f"  Page {page}: {len(matches)} products")
    for link, img_path, title in matches:
        img_url = f"https://komfortdoors.ru{img_path}"
        all_products[title] = {'img_url': img_url, 'slug': link}
        print(f"    {title} -> {img_url}")

print(f"\nTotal: {len(all_products)} products")

# Show all products
for title, data in sorted(all_products.items()):
    print(f"  {title}: {data['slug']} -> {data['img_url']}")

with open('output/final_rework_ready_v5/komfort_images.json', 'w', encoding='utf-8') as f:
    json.dump(all_products, f, ensure_ascii=False, indent=2)
