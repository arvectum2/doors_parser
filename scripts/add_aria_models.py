"""Get full-size Aria model images and add all 11 models to catalog"""
import csv, json, os, re, hashlib
from datetime import datetime
from zipfile import ZipFile
import requests
from bs4 import BeautifulSoup
from urllib.parse import urlparse

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CAT = os.path.join(BASE, "output", "final_rework_ready_v2", "doors_catalog.csv")
OUT = os.path.join(BASE, "output", "final_rework_ready_v2")
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

def resolve_url(base, src):
    if src.startswith("http"): return src
    if src.startswith("//"): return "https:" + src
    if src.startswith("/"):
        p = urlparse(base)
        return f"{p.scheme}://{p.netloc}{src}"
    return base.rstrip("/") + "/" + src.lstrip("/")

# Load catalog
with open(CAT, encoding="utf-8-sig") as f:
    catalog = list(csv.DictReader(f, delimiter=";"))
field_names = list(catalog[0].keys())

# Load aria models found
with open(os.path.join(OUT, "aria_models_found.json"), encoding="utf-8") as f:
    models = json.load(f)

# Filter to only Aria 1-11 (skip plain "Aria")
aria_models = [m for m in models if re.match(r'^Aria \d+$', m["name"])]
print(f"Aria models to add: {len(aria_models)}")

# For each model, try to get full-size image
# The page uses JavaScript to switch between models
# The 0x60 thumbnails are tiny - we need to find the actual product images
# Let's try fetching the page and looking for data attributes or JS variables
print("\n--- Getting full-size images ---")
base_page = "https://entro-design.ru/aria"
resp = requests.get(base_page, timeout=15, headers=HEADERS)
soup = BeautifulSoup(resp.text, "html.parser")

# Look for data attributes on model-sku divs
model_skus = soup.select("div.model-sku, .model-sku")
for ms in model_skus:
    name = ms.get_text(strip=True)
    if not name or not re.match(r'^Aria \d+$', name): continue
    
    # Check all attributes
    attrs = dict(ms.attrs)
    # print(f"  {name}: attrs={attrs}")
    
    # Check parent for images
    parent = ms.parent
    for _ in range(5):
        if parent is None: break
        # Check for data attributes with image URLs
        for attr, val in parent.attrs.items():
            if isinstance(val, str) and (".jpg" in val or ".png" in val or "cache" in val):
                print(f"  {name}: found img in attr {attr}: {val[:60]}")
        # Check for img elements
        imgs = parent.select("img")
        for img in imgs:
            src = img.get("src") or img.get("data-src") or ""
            if src and "0x60" not in src:
                full = resolve_url(base_page, src)
                print(f"  {name}: found non-tiny img: {full[:60]}")
        parent = parent.parent

# The page likely uses AJAX/JS to load model images
# Let's check the page source for image URLs in script tags
scripts = soup.select("script")
for s in scripts:
    if s.string:
        # Look for image URLs in JS
        urls = re.findall(r'(https?://[^\s"\'<>]+\.(?:jpg|jpeg|png|webp))', s.string)
        for u in urls:
            if "aria" in u.lower() or "cache/thumb" in u:
                # print(f"  JS img: {u[:70]}")
                pass

# Try to get larger versions by modifying the cache URL
# Pattern: /wp-content/cache/thumb/{hash}/{fullhash}_0x60.jpg
# Try: /wp-content/cache/thumb/{hash}/{fullhash}_910x.jpg or _450x.jpg
print("\n--- Try full-size URLs ---")
for m in aria_models:
    name = m["name"]
    thumb = m["img"]
    if not thumb:
        print(f"  {name}: no thumbnail")
        continue
    
    # Try different sizes
    for suffix in ["", "_910x0.jpg", "_450x0.jpg", "_800x0.jpg"]:
        if "_0x60" in thumb:
            full_url = thumb.replace("_0x60.jpg", suffix) if suffix else thumb.replace("_0x60.jpg", ".jpg")
        else:
            full_url = thumb + suffix if suffix else thumb
        
        try:
            resp = requests.head(full_url, timeout=5, headers=HEADERS)
            if resp.status_code == 200:
                ct = resp.headers.get("Content-Type", "")
                cl = resp.headers.get("Content-Length", "0")
                if "image" in ct and int(cl) > 1000:
                    m["full_img"] = full_url
                    print(f"  {name}: {full_url[:60]} ({cl} bytes)")
                    break
        except:
            pass
    else:
        # Use the thumbnail as fallback
        m["full_img"] = thumb
        print(f"  {name}: using thumbnail {thumb[:60]}")

# Now add models to catalog
# First remove existing Aria row
existing_aria = [r for r in catalog if r["producer"] == "\u042d\u043d\u0442\u0440\u043e" and "aria" in r.get("target_line", "").lower()]
print(f"\nExisting Aria rows: {len(existing_aria)}")
for r in existing_aria:
    print(f"  {r['model_name']} | {r.get('source_url', '')[:50]}")

# Remove existing Aria rows
catalog = [r for r in catalog if not (r["producer"] == "\u042d\u043d\u0442\u0440\u043e" and "aria" in r.get("target_line", "").lower())]

# Add all 11 models
for m in aria_models:
    name = m["name"]
    img = m.get("full_img", m.get("img", ""))
    catalog.append({
        "door_type": "\u043c\u0435\u0436\u043a\u043e\u043c\u043d\u0430\u0442\u043d\u0430\u044f",
        "producer": "\u042d\u043d\u0442\u0440\u043e",
        "target_line": "Aria",
        "model_name": name,
        "description": "",
        "material": "",
        "coating": "\u041f\u0412\u0425, \u042d\u043c\u0430\u043b\u044c",
        "image_url": img,
        "source_url": "https://entro-design.ru/aria",
        "status": "found",
        "comment": "aria_model_from_site",
        "image_url_front": img,
        "image_url_inside": ""
    })
    print(f"  Added: {name} | img={img[:50]}")

print(f"\nTotal after: {len(catalog)} rows")

# Dedup
seen = set()
deduped = []
for r in catalog:
    key = (r["producer"], r["target_line"], r["model_name"])
    if key not in seen:
        seen.add(key)
        deduped.append(r)
catalog = deduped
print(f"After dedup: {len(catalog)} rows")

# Save
with open(CAT, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=field_names, delimiter=";")
    w.writeheader()
    w.writerows(catalog)

# Verify
entro_aria = [r for r in catalog if r["producer"] == "\u042d\u043d\u0442\u0440\u043e" and "aria" in r.get("target_line", "").lower()]
print(f"\nFinal Entro Aria rows: {len(entro_aria)}")
for r in entro_aria:
    print(f"  {r['model_name']} | front={r.get('image_url_front','')[:50]}")
