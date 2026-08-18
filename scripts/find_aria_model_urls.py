"""Find URLs and images for all 11 Aria models"""
from bs4 import BeautifulSoup
import os, re, requests
from urllib.parse import urlparse

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

def resolve_url(base, src):
    if src.startswith("http"): return src
    if src.startswith("//"): return "https:" + src
    if src.startswith("/"):
        p = urlparse(base)
        return f"{p.scheme}://{p.netloc}{src}"
    return base.rstrip("/") + "/" + src.lstrip("/")

with open(os.path.join(BASE, "output", "aria_page.html"), encoding="utf-8") as f:
    html = f.read()
soup = BeautifulSoup(html, "html.parser")

# Find all model-sku divs - each represents a model
model_skus = soup.select("div.model-sku, .model-sku")
print(f"model-sku elements: {len(model_skus)}")

models = []
for ms in model_skus:
    name = ms.get_text(strip=True)
    if not name: continue
    # Find parent that contains image and link
    parent = ms.parent
    for _ in range(8):
        if parent is None: break
        imgs = parent.select("img")
        links = parent.select("a[href]")
        if imgs or links:
            img_src = ""
            link = ""
            for img in imgs:
                src = img.get("src") or img.get("data-src") or ""
                if src and "logo" not in src and "icon" not in src and "arrow" not in src:
                    img_src = resolve_url("https://entro-design.ru/aria", src)
                    break
            for a in links:
                href = a.get("href", "")
                if href and "aria" in href.lower():
                    link = resolve_url("https://entro-design.ru/aria", href)
                    break
            if img_src or link:
                models.append({"name": name, "img": img_src, "link": link})
                break
        parent = parent.parent

print(f"\nModels found: {len(models)}")
for m in models:
    print(f"  {m['name']} | img={m['img'][:50] if m['img'] else 'N/A'} | link={m['link'][:50] if m['link'] else 'N/A'}")

# Now check: each model might be a clickable item on the page that loads different content
# Let's check if there are individual pages for each model
print("\n--- Check individual model pages ---")
for model_num in range(1, 12):
    for pattern in [
        f"https://entro-design.ru/aria/aria-{model_num}",
        f"https://entro-design.ru/catalog/aria-{model_num}",
        f"https://entro-design.ru/product/aria-{model_num}",
        f"https://entro-design.ru/aria-{model_num}",
    ]:
        try:
            resp = requests.get(pattern, timeout=8, headers=HEADERS, allow_redirects=True)
            if resp.status_code == 200:
                soup2 = BeautifulSoup(resp.text, "html.parser")
                h1 = soup2.find("h1")
                h1_text = h1.get_text(strip=True) if h1 else ""
                # Check if it's actually about Aria model_num
                if f"aria {model_num}" in h1_text.lower() or f"aria-{model_num}" in resp.url.lower():
                    img = soup2.select_one(".product img, .gallery img, [data-fancybox] img, figure img")
                    img_src = ""
                    if img:
                        img_src = img.get("src") or img.get("data-src") or ""
                        if img_src: img_src = resolve_url(resp.url, img_src)
                    print(f"  Aria {model_num}: {resp.url} | h1={h1_text[:30]} | img={'Y' if img_src else 'N'}")
                    if img_src:
                        # Update model
                        for m in models:
                            if m["name"] == f"Aria {model_num}":
                                if not m["img"]:
                                    m["img"] = img_src
                                if not m["link"]:
                                    m["link"] = resp.url
                    break
        except:
            pass

# Final summary
print(f"\n--- Final Aria models ---")
for m in models:
    print(f"  {m['name']} | img={m['img'][:60] if m['img'] else 'N/A'} | link={m['link'][:60] if m['link'] else 'N/A'}")

import json
with open(os.path.join(BASE, "output", "final_rework_ready_v2", "aria_models_found.json"), "w", encoding="utf-8") as f:
    json.dump(models, f, ensure_ascii=False, indent=2)
