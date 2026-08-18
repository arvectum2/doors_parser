"""Deep scrape Entro Aria page to find all model cards"""
import requests, json, os
from bs4 import BeautifulSoup
from urllib.parse import urlparse

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"

def resolve_url(base, src):
    if src.startswith("http"): return src
    if src.startswith("//"): return "https:" + src
    if src.startswith("/"):
        p = urlparse(base)
        return f"{p.scheme}://{p.netloc}{src}"
    return base.rstrip("/") + "/" + src.lstrip("/")

# Fetch the main Aria page
url = "https://entro-design.ru/aria"
resp = requests.get(url, timeout=15, headers=HEADERS)
print(f"Status: {resp.status_code}, URL: {resp.url}")
soup = BeautifulSoup(resp.text, "html.parser")

# Save raw HTML for inspection
with open(os.path.join(BASE, "output", "aria_page.html"), "w", encoding="utf-8") as f:
    f.write(resp.text)

# Find ALL links on the page
all_links = soup.select("a[href]")
print(f"\nTotal links: {len(all_links)}")

# Find product-related links
product_links = []
for a in all_links:
    href = a.get("href", "")
    text = a.get_text(strip=True)[:80]
    full = resolve_url(url, href)
    # Check for aria-related
    if any(kw in href.lower() for kw in ["aria", "product", "door", "dver"]):
        product_links.append({"href": full, "text": text})

print(f"\nProduct-related links: {len(product_links)}")
for pl in product_links:
    print(f"  {pl['href'][:60]} | {pl['text'][:40]}")

# Find all images on the page
all_imgs = soup.select("img")
print(f"\nAll images: {len(all_imgs)}")
for img in all_imgs:
    src = img.get("src") or img.get("data-src") or ""
    alt = img.get("alt", "")
    if src:
        full = resolve_url(url, src)
        print(f"  {full[:70]} | alt={alt[:30]}")

# Find product cards (divs, sections)
cards = soup.select(".product-card, .catalog-item, .card, [class*=product], .item, .door, .model")
print(f"\nProduct cards: {len(cards)}")
for c in cards[:5]:
    print(f"  tag={c.name}, classes={c.get('class', [])}")
    a = c.find("a")
    img = c.find("img")
    if a: print(f"    link: {a.get('href', '')[:50]}")
    if img: print(f"    img: {(img.get('src') or img.get('data-src', ''))[:50]}")

# Find JSON data (some sites use JSON-LD or embedded data)
scripts = soup.find_all("script", type="application/ld+json")
for s in scripts:
    try:
        data = json.loads(s.string)
        print(f"\nJSON-LD: {json.dumps(data, ensure_ascii=False)[:200]}")
    except:
        pass

# Also look for inline JS data
for s in soup.find_all("script"):
    if s.string and ("aria" in (s.string or "").lower() or "product" in (s.string or "").lower()):
        text = s.string[:500]
        print(f"\nJS data with aria/product: {text[:200]}")
