"""Fix only the 38 bad ASD rows: re-scrape with strict image-only filter"""
import csv, json, os, time, hashlib, re
from collections import Counter
from datetime import datetime
from zipfile import ZipFile
import requests
from bs4 import BeautifulSoup
from urllib.parse import urlparse

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CAT = os.path.join(BASE, "output", "final_rework_ready_v2", "doors_catalog.csv")
OUT = os.path.join(BASE, "output", "final_rework_ready_v2")
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
BAD = ["cache_image", "fragment", "handle", "ruchka", "lock", "zamok", "interior", "badge", "banner", "logo", "404"]

def resolve_url(base, src):
    if src.startswith("http"): return src
    if src.startswith("//"): return "https:" + src
    if src.startswith("/"):
        p = urlparse(base)
        return f"{p.scheme}://{p.netloc}{src}"
    return base.rstrip("/") + "/" + src.lstrip("/")

def fetch_door_images(source_url):
    """Fetch only actual door images, skip .html, /onstock/ accessories, etc."""
    if not source_url: return []
    try:
        resp = requests.get(source_url, timeout=15, headers=HEADERS)
        if resp.status_code != 200: return []
        soup = BeautifulSoup(resp.text, "html.parser")
        urls, seen = [], set()
        
        # Gallery img tags are most reliable
        for img in soup.select(".gallery img, [data-fancybox] img"):
            src = img.get("src") or img.get("data-src") or ""
            if not src or src in seen: continue
            full = resolve_url(source_url, src)
            sl = full.lower()
            # Skip non-images
            if sl.endswith((".html", ".htm")): continue
            if "/onstock/" in sl: continue
            # Skip tiny images (icons, arrows, etc.)
            w = img.get("width", "")
            h = img.get("height", "")
            if w and w.isdigit() and int(w) < 100: continue
            if h and h.isdigit() and int(h) < 100: continue
            seen.add(src)
            urls.append(full)
        
        # Fancybox links as backup
        for a in soup.select("a[data-fancybox], a.fancybox"):
            href = a.get("href", "")
            if not href or href in seen: continue
            full = resolve_url(source_url, href)
            sl = full.lower()
            if sl.endswith((".html", ".htm")): continue
            if "/onstock/" in sl: continue
            seen.add(href)
            urls.append(full)
        
        return urls
    except:
        return []

with open(CAT, encoding="utf-8-sig") as f:
    catalog = list(csv.DictReader(f, delimiter=";"))
field_names = list(catalog[0].keys())

# Find bad rows
asd = [r for r in catalog if r["producer"] == "\u0410\u0421\u0414"]
bad_rows = []
for r in asd:
    front = (r.get("image_url_front", "") or "").lower()
    inside = (r.get("image_url_inside", "") or "").lower()
    for bp in BAD:
        if bp in front or bp in inside:
            bad_rows.append(r)
            break

no_front = [r for r in asd if not r.get("image_url_front", "")]
print(f"Bad rows to fix: {len(bad_rows)}")
print(f"No front rows: {len(no_front)}")

# Also check rows with /onstock/ in their image URLs
onstock_img = [r for r in asd if "/onstock/" in (r.get("image_url_front", "") or "")]
print(f"Onstock image rows: {len(onstock_img)}")

fixed = 0
for i, r in enumerate(bad_rows):
    src = r.get("source_url", "")
    model = r["model_name"][:25]
    gallery = fetch_door_images(src)
    if len(gallery) >= 2:
        r["image_url"] = gallery[0]
        r["image_url_front"] = gallery[0]
        r["image_url_inside"] = gallery[1]
        r["status"] = "found"
        fixed += 1
        print(f"  [{i+1}/{len(bad_rows)}] FIXED {model}: {len(gallery)} imgs")
    elif len(gallery) == 1:
        r["image_url"] = gallery[0]
        r["image_url_front"] = gallery[0]
        r["image_url_inside"] = ""
        r["status"] = "found"
        fixed += 1
        print(f"  [{i+1}/{len(bad_rows)}] FRONT ONLY {model}")
    else:
        print(f"  [{i+1}/{len(bad_rows)}] NO IMAGES {model} | src={src[:50]}")
    time.sleep(0.3)

# Fix no-front rows
for r in no_front:
    src = r.get("source_url", "")
    model = r["model_name"][:25]
    gallery = fetch_door_images(src)
    if gallery:
        r["image_url"] = gallery[0]
        r["image_url_front"] = gallery[0]
        if len(gallery) > 1:
            r["image_url_inside"] = gallery[1]
        r["status"] = "found"
        fixed += 1
        print(f"  FRONT FIXED: {model}")
    time.sleep(0.3)

print(f"\nFixed: {fixed}/{len(bad_rows) + len(no_front)}")

# Verify
asd2 = [r for r in catalog if r["producer"] == "\u0410\u0421\u0414"]
bad2 = 0
for r in asd2:
    front = (r.get("image_url_front", "") or "").lower()
    inside = (r.get("image_url_inside", "") or "").lower()
    for bp in BAD:
        if bp in front or bp in inside:
            bad2 += 1
            break
print(f"After fix: ASD={len(asd2)}, front={sum(1 for r in asd2 if r.get('image_url_front',''))}, inside={sum(1 for r in asd2 if r.get('image_url_inside',''))}, bad={bad2}")

# Save
with open(CAT, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=field_names, delimiter=";")
    w.writeheader()
    w.writerows(catalog)

# XLSX
try:
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    for ci, h in enumerate(field_names, 1): ws.cell(row=1, column=ci, value=h)
    for ri, r in enumerate(catalog, 2):
        for ci, h in enumerate(field_names, 1): ws.cell(row=ri, column=ci, value=r.get(h, ""))
    wb.save(os.path.join(OUT, "doors_catalog.xlsx"))
except: pass

with open(os.path.join(OUT, "wordpress_import.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=field_names, delimiter=";")
    w.writeheader()
    w.writerows(catalog)

producers = Counter(r["producer"] for r in catalog)
os.makedirs(os.path.join(OUT, "check_by_producer"), exist_ok=True)
for prod in sorted(producers.keys()):
    pr = [r for r in catalog if r["producer"] == prod]
    safe = prod.replace("/", "_").replace("\\", "_")
    with open(os.path.join(OUT, "check_by_producer", f"{safe}.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=field_names, delimiter=";")
        w.writeheader()
        w.writerows(pr)

report = {"generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "total": len(catalog), "producers": dict(producers),
    "asd": {"total": len(asd2), "front": sum(1 for r in asd2 if r.get("image_url_front","")),
            "inside": sum(1 for r in asd2 if r.get("image_url_inside","")), "bad": bad2}}
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
print(f"Total: {len(catalog)}, ASD: front={sum(1 for r in asd2 if r.get('image_url_front',''))}/{len(asd2)}, inside={sum(1 for r in asd2 if r.get('image_url_inside',''))}/{len(asd2)}, bad={bad2}")
