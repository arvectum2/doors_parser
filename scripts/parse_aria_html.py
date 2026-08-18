"""Parse the Aria page HTML to find product cards with model names"""
from bs4 import BeautifulSoup
import os, re

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
with open(os.path.join(BASE, "output", "aria_page.html"), encoding="utf-8") as f:
    html = f.read()

soup = BeautifulSoup(html, "html.parser")

# Find product cards
cards = soup.select(".product-card")
print(f"product-card elements: {len(cards)}")
for i, card in enumerate(cards):
    title = card.select_one(".product-card-title, h2, h3")
    img = card.select_one("img")
    link = card.select_one("a[href]")
    desc = card.select_one(".product-card-description, p")
    print(f"\nCard {i+1}:")
    if title: print(f"  title: {title.get_text(strip=True)[:60]}")
    if img:
        src = img.get("src") or img.get("data-src") or ""
        alt = img.get("alt", "")
        print(f"  img: {src[:70]}")
        print(f"  alt: {alt[:40]}")
    if link: print(f"  link: {link.get('href', '')[:60]}")
    if desc: print(f"  desc: {desc.get_text(strip=True)[:60]}")

# Also check product-top-container
tops = soup.select(".product-top-container")
print(f"\nproduct-top-container: {len(tops)}")
for i, t in enumerate(tops):
    img = t.select_one("img")
    link = t.select_one("a[href]")
    if img: print(f"  img: {(img.get('src') or img.get('data-src', ''))[:70]}")
    if link: print(f"  link: {link.get('href', '')[:60]}")

# Check for h2 elements with model names
h2s = soup.select("h2")
print(f"\nAll h2: {len(h2s)}")
for h in h2s:
    text = h.get_text(strip=True)[:60]
    if text:
        print(f"  {text}")

# Check for all divs with class containing 'product'
prods = soup.select("[class*=product]")
print(f"\nElements with 'product' in class: {len(prods)}")
classes_seen = set()
for p in prods:
    cls = " ".join(p.get("class", []))
    if cls not in classes_seen:
        classes_seen.add(cls)
        print(f"  {p.name}.{cls}")

# Search for model-related text
text = soup.get_text()
for pattern in ["Aria 1", "Aria 2", "Aria 3", "Aria 4", "Aria 5", "Aria-1", "Aria-2", "Aria-3", "Aria-4", "Aria-5", "ARIA-1", "ARIA-2", "ARIA-3"]:
    count = text.lower().count(pattern.lower())
    if count > 0:
        print(f"  '{pattern}' found {count} times")

# Look for any numbered model patterns
aria_matches = re.findall(r'(?i)aria[\s\-_]?\d', text)
print(f"\nAria+number matches in text: {aria_matches}")

# Check 450x images (product thumbnails)
imgs_450 = [img for img in soup.select("img") if "450x" in (img.get("src") or "")]
print(f"\n450x product images: {len(imgs_450)}")
for img in imgs_450:
    src = img.get("src") or ""
    alt = img.get("alt", "")
    print(f"  {src[:70]} | alt={alt[:30]}")
