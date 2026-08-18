"""Find all Aria model names and their images from the page"""
from bs4 import BeautifulSoup
import os, re

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
with open(os.path.join(BASE, "output", "aria_page.html"), encoding="utf-8") as f:
    html = f.read()

soup = BeautifulSoup(html, "html.parser")

# Find the section with model listings
# Look for product-wrap elements (the model grid)
wraps = soup.select(".product-wrap")
print(f"product-wrap: {len(wraps)}")
for i, w in enumerate(wraps):
    title = w.select_one("h2, h3, .product-card-title")
    img = w.select_one("img")
    link = w.select_one("a[href]")
    print(f"\nWrap {i+1}:")
    if title: print(f"  title: {title.get_text(strip=True)}")
    if img:
        src = img.get("src") or img.get("data-src") or ""
        print(f"  img: {src[:80]}")
    if link: print(f"  link: {link.get('href', '')[:60]}")

# Search for Aria model text with surrounding context
text = soup.get_text()
# Find all lines containing "Aria" followed by a number
lines = text.split("\n")
aria_lines = []
for i, line in enumerate(lines):
    if re.search(r'(?i)aria\s*\d', line.strip()):
        context = line.strip()[:100]
        aria_lines.append((i, context))

print(f"\nLines with Aria+number: {len(aria_lines)}")
for idx, ctx in aria_lines[:20]:
    print(f"  line {idx}: {ctx}")

# Find all text near 450x images
for img in soup.select("img"):
    src = img.get("src") or ""
    if "450x" in src:
        # Find parent and siblings
        parent = img.parent
        for _ in range(5):
            if parent is None: break
            text_content = parent.get_text(strip=True)[:100]
            if text_content and len(text_content) > 3:
                print(f"\n  Near {src[:40]}:")
                print(f"    Parent text: {text_content}")
                break
            parent = parent.parent

# Find the full text of the page section that lists models
# Look for the section with "Модельный ряд" or model listing
for section in soup.select("section, div[class*=section], div[class*=model]"):
    text = section.get_text(strip=True)[:200]
    if any(kw in text.lower() for kw in ["aria", "model", "\u043c\u043e\u0434\u0435\u043b"]):
        print(f"\nSection: {section.name}.{'.'.join(section.get('class', []))}")
        print(f"  Text: {text[:150]}")
