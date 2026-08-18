"""Step 1: Load catalog, scrape Entro Aria, load reference lines"""
import csv, json, os, zipfile, io, re
import requests
from bs4 import BeautifulSoup
from urllib.parse import urlparse

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
ZIP = os.path.join(BASE, "output", "doors_parser_final_REWORK_READY_v2_2026-06-23_16-25-55.zip")
OUT = os.path.join(BASE, "output", "final_rework_ready_v2")
os.makedirs(OUT, exist_ok=True)

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

# Load catalog from zip
with zipfile.ZipFile(ZIP) as zf:
    csv_name = [n for n in zf.namelist() if n.endswith("doors_catalog.csv")][0]
    with zf.open(csv_name) as f:
        catalog = list(csv.DictReader(io.StringIO(f.read().decode("utf-8-sig")), delimiter=";"))
print(f"Catalog: {len(catalog)} rows")

# Load reference lines
ref_path = os.path.join(BASE, "output", "reference_lines_decoded.json")
with open(ref_path, encoding="utf-8") as f:
    ref_data = json.load(f)
print(f"Reference: {sum(len(v) for v in ref_data.values())} lines across {len(ref_data)} producers")

# Save catalog locally for validation
with open(os.path.join(OUT, "doors_catalog.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=list(catalog[0].keys()), delimiter=";")
    w.writeheader()
    w.writerows(catalog)

# ============================================================
# ENTRO ARIA DEEP SCRAPE
# ============================================================
print("\n=== ENTRO ARIA DEEP SCRAPE ===")

def resolve_url(base, src):
    if src.startswith("http"): return src
    if src.startswith("//"): return "https:" + src
    if src.startswith("/"):
        p = urlparse(base)
        return f"{p.scheme}://{p.netloc}{src}"
    return base.rstrip("/") + "/" + src.lstrip("/")

aria_results = []

# Step 1: Scrape the main Aria catalog page
print("\n--- Main Aria catalog page ---")
main_urls = [
    "https://entro-design.ru/catalog/aria",
    "https://entro-design.ru/catalog/areya",
]
for url in main_urls:
    try:
        resp = requests.get(url, timeout=15, headers=HEADERS, allow_redirects=True)
        print(f"  {url} -> {resp.status_code} (final: {resp.url})")
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, "html.parser")
            title = soup.find("title")
            h1 = soup.find("h1")
            print(f"    title: {title.get_text(strip=True)[:60] if title else 'N/A'}")
            print(f"    h1: {h1.get_text(strip=True)[:60] if h1 else 'N/A'}")
            
            # Find product cards / model links
            cards = soup.select(".product-card, .catalog-item, .card, [class*=product], .item")
            links = soup.select("a[href]")
            aria_links = []
            for a in links:
                href = a.get("href", "")
                text = a.get_text(strip=True)[:80]
                if any(kw in href.lower() for kw in ["aria", "areya"]):
                    full = resolve_url(url, href)
                    aria_links.append({"href": full, "text": text})
            print(f"    Product cards: {len(cards)}")
            print(f"    Aria links: {len(aria_links)}")
            for al in aria_links[:20]:
                print(f"      {al['href'][:60]} | {al['text'][:40]}")
            
            # Find all images on the page
            imgs = soup.select("img")
            aria_imgs = []
            for img in imgs:
                src = img.get("src") or img.get("data-src") or ""
                if src:
                    full = resolve_url(url, src)
                    if "aria" in full.lower() or "areya" in full.lower():
                        aria_imgs.append(full)
            print(f"    Aria images: {len(aria_imgs)}")
            for ai in aria_imgs[:10]:
                print(f"      {ai[:70]}")
    except Exception as e:
        print(f"    Error: {e}")

# Step 2: Try individual Aria model pages
print("\n--- Individual Aria model pages ---")
aria_pages = []
for prefix in ["aria", "areya"]:
    for num in range(1, 10):
        for suffix in ["", "-1", "-2", "-3", "-4", "-5"]:
            aria_pages.append(f"https://entro-design.ru/catalog/{prefix}-{num}{suffix}")
            aria_pages.append(f"https://entro-design.ru/product/{prefix}-{num}{suffix}")
            aria_pages.append(f"https://entro-design.ru/catalog/{prefix}{suffix}")

# Also check entro-design.ru/dveri/aria
aria_pages.extend([
    "https://entro-design.ru/dveri/aria",
    "https://entro-design.ru/dveri/aria-1",
    "https://entro-design.ru/dveri/aria-2",
    "https://entro-design.ru/dveri/aria-3",
])

found_pages = []
for url in aria_pages:
    try:
        resp = requests.get(url, timeout=8, headers=HEADERS, allow_redirects=True)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, "html.parser")
            h1 = soup.find("h1")
            title = soup.find("title")
            h1_text = h1.get_text(strip=True) if h1 else ""
            title_text = title.get_text(strip=True) if title else ""
            # Check if this is actually an Aria product page
            if any(kw in (h1_text + title_text).lower() for kw in ["aria", "areya", "\u0410\u0440\u0435\u044f"]):
                img = soup.select_one(".product img, .gallery img, [data-fancybox] img, .main-image img")
                img_src = ""
                if img:
                    img_src = img.get("src") or img.get("data-src") or ""
                    if img_src: img_src = resolve_url(url, img_src)
                found_pages.append({
                    "url": resp.url,
                    "h1": h1_text[:60],
                    "title": title_text[:60],
                    "img": img_src[:80] if img_src else "",
                })
                print(f"  FOUND: {resp.url[:60]} | h1={h1_text[:30]} | img={'Y' if img_src else 'N'}")
    except:
        pass

print(f"\nTotal Aria pages found: {len(found_pages)}")
for p in found_pages:
    print(f"  {p['url']}")
    print(f"    h1: {p['h1']}")
    print(f"    img: {p['img']}")

# Step 3: Check existing Entro rows in catalog
entro = [r for r in catalog if r["producer"] == "\u042d\u043d\u0442\u0440\u043e"]
aria_catalog = [r for r in entro if "aria" in r.get("target_line", "").lower() or "aria" in r.get("model_name", "").lower() or "\u0410\u0440\u0435\u044f" in r.get("model_name", "")]
print(f"\n--- Existing Entro Aria in catalog ---")
print(f"  Entro total: {len(entro)}")
print(f"  Aria rows: {len(aria_catalog)}")
for r in aria_catalog:
    print(f"    {r['model_name'][:30]} | line={r['target_line'][:20]} | status={r.get('status','')} | front={r.get('image_url_front','')[:40]}")

# Save results
results = {
    "found_pages": found_pages,
    "aria_catalog_rows": [{"model": r["model_name"], "line": r["target_line"], "status": r.get("status",""), "source": r.get("source_url","")[:60]} for r in aria_catalog],
    "entro_total": len(entro),
}
with open(os.path.join(OUT, "aria_scrape_results.json"), "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print("\nDone. Results saved to aria_scrape_results.json")
