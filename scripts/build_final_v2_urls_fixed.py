"""Fix ASD: find correct /doors/ source URLs for onstock rows, re-scrape galleries"""
import csv, json, os, time, hashlib, re
from collections import Counter
from datetime import datetime
from zipfile import ZipFile
import requests
from bs4 import BeautifulSoup

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
V163 = os.path.join(BASE, "output", "final_hybrid_client_ready_v16_3_asd_full_door_only", "doors_catalog.csv")
OUT = os.path.join(BASE, "output", "final_rework_ready_v2")
os.makedirs(OUT, exist_ok=True)

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

def resolve_url(base, src):
    if src.startswith("http"): return src
    if src.startswith("//"): return "https:" + src
    if src.startswith("/"):
        from urllib.parse import urlparse
        p = urlparse(base)
        return f"{p.scheme}://{p.netloc}{src}"
    return base.rstrip("/") + "/" + src.lstrip("/")

def fetch_gallery(source_url):
    if not source_url: return []
    try:
        resp = requests.get(source_url, timeout=15, headers=HEADERS)
        if resp.status_code != 200: return []
        soup = BeautifulSoup(resp.text, "html.parser")
        urls, seen = [], set()
        for a in soup.select("a[data-fancybox], a.fancybox"):
            href = a.get("href", "")
            if href and href not in seen:
                seen.add(href)
                urls.append(resolve_url(source_url, href))
        for img in soup.select(".gallery img, [data-fancybox] img"):
            src = img.get("src") or img.get("data-src") or ""
            if src and src not in seen:
                seen.add(src)
                urls.append(resolve_url(source_url, src))
        return urls
    except:
        return []

def is_door_image(src):
    sl = src.lower()
    if sl.endswith((".html", ".htm")): return False
    if "/onstock/" in sl: return False
    return True

with open(V163, encoding="utf-8-sig") as f:
    catalog = list(csv.DictReader(f, delimiter=";"))

field_names = list(catalog[0].keys()) + ["image_url_front", "image_url_inside"]
for r in catalog:
    r["image_url_front"] = ""
    r["image_url_inside"] = ""

# Step 1: Find correct source URLs for onstock ASD rows
print("=== Finding correct ASD source URLs ===")
asd = [r for r in catalog if r["producer"] == "АСД"]

# Build search index from ASD catalog and site
onstock_rows = [r for r in asd if "/onstock/" in (r.get("source_url", "") or "")]
print(f"  onstock rows to fix: {len(onstock_rows)}")

# Try to find /doors/ pages by scraping the main catalog page
print("  Scraping ASD catalog...")
catalog_urls = {}
for page_url in ["https://as-doors.ru/doors/", "https://as-doors.ru/doors/zakaznyie/"]:
    try:
        resp = requests.get(page_url, timeout=15, headers=HEADERS)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, "html.parser")
            for a in soup.find_all("a", href=True):
                href = a["href"]
                if "/doors/" in href or "/zakaznyie/" in href:
                    full = resolve_url(page_url, href)
                    # Extract model name from URL
                    name_part = href.split("/")[-1].replace(".html", "").replace("-", " ").lower()
                    catalog_urls[name_part] = full
            print(f"    {page_url}: found {len(catalog_urls)} links")
    except Exception as e:
        print(f"    Error: {e}")

# Also try individual known patterns
for r in onstock_rows:
    model = r["model_name"]
    # Try to construct a /doors/ URL from model name
    slug = re.sub(r'[^a-z0-9]+', '-', model.lower()).strip('-')
    # Try common patterns
    for prefix in ["https://as-doors.ru/doors/zakaznyie/", "https://as-doors.ru/doors/"]:
        try_url = prefix + slug + ".html"
        try:
            resp = requests.head(try_url, timeout=5, headers=HEADERS, allow_redirects=True)
            if resp.status_code == 200:
                catalog_urls[slug] = try_url
                break
        except:
            pass

# Step 2: ASD scan with corrected URLs
print("\n=== ASD scan ===")
inside_count = 0
front_count = 0
for i, r in enumerate(asd):
    src = r.get("source_url", "")
    if not src: continue
    
    # If source is onstock, try to find a /doors/ page
    if "/onstock/" in src:
        model_slug = re.sub(r'[^a-z0-9]+', '-', r["model_name"].lower()).strip('-')
        # Search catalog_urls for a match
        found_url = None
        for key, url in catalog_urls.items():
            if model_slug[:10] in key or key[:10] in model_slug:
                found_url = url
                break
        if found_url:
            src = found_url
            r["source_url"] = src

    gallery = fetch_gallery(src)
    # Filter to only door images (skip accessory links)
    gallery = [u for u in gallery if is_door_image(u)]
    
    if len(gallery) >= 2:
        r["image_url"] = gallery[0]
        r["image_url_front"] = gallery[0]
        r["image_url_inside"] = gallery[1]
        r["status"] = "found"
        front_count += 1
        inside_count += 1
    elif len(gallery) == 1:
        r["image_url"] = gallery[0]
        r["image_url_front"] = gallery[0]
        front_count += 1
    
    if (i + 1) % 50 == 0:
        print(f"  {i+1}/{len(asd)}, front={front_count}, inside={inside_count}")
        time.sleep(0.3)

print(f"Scan done: front={front_count}, inside={inside_count}")

# Fix Bruklin
for r in catalog:
    if r["producer"] == "АСД" and "Бруклин" in r.get("model_name", ""):
        gallery = fetch_gallery("https://as-doors.ru/doors/zakaznyie/bruklin.html")
        gallery = [u for u in gallery if is_door_image(u)]
        if gallery:
            r["image_url"] = gallery[0]
            r["image_url_front"] = gallery[0]
            r["image_url_inside"] = gallery[1] if len(gallery) > 1 else ""
            r["source_url"] = "https://as-doors.ru/doors/zakaznyie/bruklin.html"
            print(f"Bruklin fixed")
        break

# Set front for non-ASD
for r in catalog:
    if r["producer"] != "АСД" and r.get("image_url") and not r.get("image_url_front"):
        r["image_url_front"] = r["image_url"]

# Dedup
seen = set()
deduped = []
for r in catalog:
    key = (r["producer"], r["target_line"], r["model_name"])
    if key not in seen:
        seen.add(key)
        deduped.append(r)
catalog = deduped

# Save
csv_path = os.path.join(OUT, "doors_catalog.csv")
with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=field_names, delimiter=";")
    w.writeheader()
    w.writerows(catalog)

# Verify
with open(csv_path, encoding="utf-8-sig") as f:
    verify = list(csv.DictReader(f, delimiter=";"))
vasd = [r for r in verify if r["producer"] == "АСД"]

BAD = ["cache_image", "fragment", "handle", "ruchka", "lock", "zamok", "interior", "badge", "banner", "logo", "404"]
bad_count = 0
for r in vasd:
    for bp in BAD:
        if bp in (r.get("image_url_front", "") or "").lower() or bp in (r.get("image_url_inside", "") or "").lower():
            bad_count += 1
            break

print(f"\nPost-save: ASD={len(vasd)}, front={sum(1 for r in vasd if r.get('image_url_front',''))}, inside={sum(1 for r in vasd if r.get('image_url_inside',''))}, bad={bad_count}")

# Show remaining bad
for r in vasd:
    for bp in BAD:
        if bp in (r.get("image_url_front", "") or "").lower() or bp in (r.get("image_url_inside", "") or "").lower():
            print(f"  BAD: {r['model_name'][:25]} | {bp}")
            break

# Show no front
for r in vasd:
    if not r.get("image_url_front", ""):
        print(f"  NO FRONT: {r['model_name'][:30]} | src={r.get('source_url','')[:60]}")

# XLSX
try:
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    for ci, h in enumerate(field_names, 1): ws.cell(row=1, column=ci, value=h)
    for ri, r in enumerate(verify, 2):
        for ci, h in enumerate(field_names, 1): ws.cell(row=ri, column=ci, value=r.get(h, ""))
    wb.save(os.path.join(OUT, "doors_catalog.xlsx"))
except: pass

with open(os.path.join(OUT, "wordpress_import.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=field_names, delimiter=";")
    w.writeheader()
    w.writerows(verify)

producers = Counter(r["producer"] for r in verify)
os.makedirs(os.path.join(OUT, "check_by_producer"), exist_ok=True)
for prod in sorted(producers.keys()):
    pr = [r for r in verify if r["producer"] == prod]
    safe = prod.replace("/", "_").replace("\\", "_")
    with open(os.path.join(OUT, "check_by_producer", f"{safe}.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=field_names, delimiter=";")
        w.writeheader()
        w.writerows(pr)

report = {"generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "total": len(verify), "producers": dict(producers),
    "asd": {"total": len(vasd), "front": sum(1 for r in vasd if r.get("image_url_front","")),
            "inside": sum(1 for r in vasd if r.get("image_url_inside","")), "bad": bad_count}}
with open(os.path.join(OUT, "run_report.json"), "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)

zn = f"doors_parser_final_REWORK_READY_v2_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.zip"
zp = os.path.join(BASE, "output", zn)
with ZipFile(zp, "w") as zf:
    for root, dirs, files in os.walk(OUT):
        for fn in files:
            fp = os.path.join(root, fn)
            zf.write(fp, arcname=os.path.join("output", "final_rework_ready_v2", os.path.relpath(fp, OUT)))

sha = hashlib.sha256()
with open(zp, "rb") as f:
    for chunk in iter(lambda: f.read(65536), b""): sha.update(chunk)

print(f"\n{'='*60}")
print(f"Archive: {zp}")
print(f"SHA256: {sha.hexdigest()}")
print(f"Total: {len(verify)}, ASD: front={sum(1 for r in vasd if r.get('image_url_front',''))}/{len(vasd)}, inside={sum(1 for r in vasd if r.get('image_url_inside',''))}/{len(vasd)}, bad={bad_count}")
