"""All-in-one final v2 build"""
import csv, json, os, re, time, hashlib, shutil
from collections import Counter
from datetime import datetime
from zipfile import ZipFile
import requests
from bs4 import BeautifulSoup

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
V163_DIR = os.path.join(BASE, "output", "final_hybrid_client_ready_v16_3_asd_full_door_only")
FINAL_DIR = os.path.join(BASE, "output", "final_rework_ready")
OUT_DIR = os.path.join(BASE, "output", "final_rework_ready_v2")
os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(os.path.join(OUT_DIR, "check_by_producer"), exist_ok=True)

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
BAD_PATTERNS = ["cache_image", "fragment", "handle", "ruchka", "lock", "zamok",
                "interior", "badge", "banner", "logo", "404"]

with open(os.path.join(V163_DIR, "doors_catalog.csv"), encoding="utf-8-sig") as f:
    catalog = list(csv.DictReader(f, delimiter=";"))

if "image_url_front" not in catalog[0]:
    for r in catalog: r["image_url_front"] = ""
if "image_url_inside" not in catalog[0]:
    for r in catalog: r["image_url_inside"] = ""

FIELD_NAMES = list(catalog[0].keys())
print(f"Base: {len(catalog)} rows, cols: {FIELD_NAMES}")

def resolve_url(base, src):
    if src.startswith("http"): return src
    if src.startswith("//"): return "https:" + src
    if src.startswith("/"):
        from urllib.parse import urlparse
        p = urlparse(base)
        return f"{p.scheme}://{p.netloc}{src}"
    return base.rstrip("/") + "/" + src.lstrip("/")

def is_product_image(src):
    sl = src.lower()
    for bp in ["icon", "logo", "arrow", "star", ".gif", ".svg", ".png"]:
        if bp in sl: return False
    return True

def fetch_gallery(source_url):
    if not source_url: return []
    try:
        resp = requests.get(source_url, timeout=15, headers=HEADERS)
        if resp.status_code != 200: return []
        soup = BeautifulSoup(resp.text, "html.parser")
        urls = []
        seen = set()
        for a in soup.select("a[data-fancybox], a.fancybox"):
            href = a.get("href", "")
            if href and href not in seen and is_product_image(href):
                seen.add(href)
                urls.append(resolve_url(source_url, href))
        for img in soup.select(".gallery img, [data-fancybox] img"):
            src = img.get("src") or img.get("data-src") or ""
            if src and src not in seen and is_product_image(src):
                seen.add(src)
                urls.append(resolve_url(source_url, src))
        return urls
    except:
        return []

# 1. ASD inside scan
print("\n=== ASD inside photo scan ===")
asd_rows = [r for r in catalog if r["producer"] == "АСД"]
print(f"ASD rows: {len(asd_rows)}")
asd_with_source = [r for r in asd_rows if r.get("source_url", "")]
print(f"ASD with source_url: {len(asd_with_source)}")

inside_found = 0
for i, r in enumerate(asd_with_source):
    gallery = fetch_gallery(r["source_url"])
    if len(gallery) >= 2:
        r["image_url_front"] = gallery[0]
        r["image_url"] = gallery[0]
        r["image_url_inside"] = gallery[1]
        r["status"] = "found"
        r["comment"] = (r.get("comment", "") or "") + "; inside_photo_added"
        inside_found += 1
    elif len(gallery) == 1:
        r["image_url_front"] = gallery[0]
        r["image_url"] = gallery[0]
        r["comment"] = (r.get("comment", "") or "") + "; only_front_found"
    if (i + 1) % 30 == 0:
        print(f"  {i+1}/{len(asd_with_source)}, inside={inside_found}")
        time.sleep(0.5)

# Verify after scan
asd_verify = [r for r in catalog if r["producer"] == "АСД"]
print(f"  Scan done: inside_found={inside_found}")
print(f"  Verify: front={sum(1 for r in asd_verify if r.get('image_url_front',''))}, inside={sum(1 for r in asd_verify if r.get('image_url_inside',''))}")

# 2. Fix Bruklin
print("\n=== Fix Bruklin ===")
for r in catalog:
    if r["producer"] == "АСД" and "Бруклин" in r.get("model_name", ""):
        old = r.get("image_url", "")
        gallery = fetch_gallery("https://as-doors.ru/doors/zakaznyie/bruklin.html")
        if gallery:
            r["image_url"] = gallery[0]
            r["image_url_front"] = gallery[0]
            r["image_url_inside"] = gallery[1] if len(gallery) > 1 else ""
            r["status"] = "found"
            r["source_url"] = "https://as-doors.ru/doors/zakaznyie/bruklin.html"
            print(f"  FIXED: {old[:40]} -> {gallery[0][:50]}")
        break

# 3. Entro Aria
print("\n=== Entro Aria scan ===")
aria_urls = ["https://entro-design.ru/catalog/aria-1", "https://entro-design.ru/catalog/aria-2",
             "https://entro-design.ru/catalog/aria-3", "https://entro-design.ru/catalog/aria-4",
             "https://entro-design.ru/catalog/aria-5"]
aria_added = 0
for url in aria_urls:
    try:
        resp = requests.get(url, timeout=10, headers=HEADERS, allow_redirects=True)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, "html.parser")
            h1 = soup.find("h1")
            if h1 and "aria" in h1.get_text().lower():
                img = soup.select_one(".product img, .gallery img, [data-fancybox] img")
                img_src = resolve_url(url, img.get("src") or img.get("data-src") or "") if img else ""
                model_name = h1.get_text(strip=True)
                exists = any(r["producer"] == "Энтро" and r["model_name"] == model_name for r in catalog)
                if not exists:
                    catalog.append({"door_type": "межкомнатная", "producer": "Энтро", "target_line": "Aria",
                        "model_name": model_name, "description": "", "material": "",
                        "coating": "ПВХ, Эмаль", "image_url": img_src, "source_url": url,
                        "status": "found" if img_src else "needs_review", "comment": "added_aria_model",
                        "image_url_front": img_src, "image_url_inside": ""})
                    aria_added += 1
                    print(f"  Added: {model_name}")
    except: pass
print(f"  Aria added: {aria_added}")

# 4. Hausdorss stubs
print("\n=== Hausdorss stubs ===")
hd_urls = {"Фактурные линии": ["https://hausdoors.ru/dveri-facturnye-linii"],
           "Новая классика": ["https://hausdoors.ru/dveri-new-classic"],
           "Минимализм": ["https://hausdoors.ru/dveri-minimalism"],
           "Современная классика": ["https://hausdoors.ru/dveri-soft-classic"]}
hd_added = 0
for r in catalog:
    if r["producer"] == "Хаусдорс" and r.get("model_name", "").startswith("[MISSING]"):
        line = r["target_line"]
        for url in hd_urls.get(line, []):
            try:
                resp = requests.get(url, timeout=10, headers=HEADERS)
                if resp.status_code == 200:
                    soup = BeautifulSoup(resp.text, "html.parser")
                    cards = soup.select(".product-card, .catalog-item, .card, [class*=product]")
                    models = []
                    for card in cards:
                        a = card.find("a")
                        if a:
                            href = a.get("href", "")
                            text = a.get_text(strip=True)[:60]
                            img = card.find("img")
                            img_src = img.get("src") or img.get("data-src") if img else ""
                            if text and len(text) > 2:
                                full_url = href if href.startswith("http") else f"https://hausdoors.ru{href}"
                                models.append({"name": text, "url": full_url, "img": img_src})
                    if models:
                        r["_remove"] = True
                        for m in models:
                            exists = any(x["producer"] == "Хаусдорс" and x["model_name"] == m["name"] for x in catalog)
                            if not exists:
                                img = m["img"] or ""
                                if img and not img.startswith("http"):
                                    img = f"https://hausdoors.ru{img}" if img.startswith("/") else f"https://hausdoors.ru/{img}"
                                catalog.append({"door_type": "межкомнатная", "producer": "Хаусдорс",
                                    "target_line": line, "model_name": m["name"], "description": "",
                                    "material": "", "coating": r.get("coating", ""),
                                    "image_url": img, "source_url": m["url"],
                                    "status": "found" if img else "needs_review",
                                    "comment": f"added_from_site_{line}",
                                    "image_url_front": img, "image_url_inside": ""})
                                hd_added += 1
                        print(f"  {line}: {len(models)} models added")
                    else:
                        print(f"  {line}: no models found")
                    break
            except: pass
catalog = [r for r in catalog if not r.get("_remove", False)]
print(f"  Hausdorss added: {hd_added}")

# 5. Set front for others
for r in catalog:
    if r["producer"] != "АСД":
        img = r.get("image_url", "") or ""
        if img and not r.get("image_url_front", ""):
            r["image_url_front"] = img

# 6. Dedup + save
seen = set()
deduped = []
for r in catalog:
    key = (r["producer"], r["target_line"], r["model_name"])
    if key not in seen:
        seen.add(key)
        deduped.append(r)
catalog = deduped

producers = Counter(r["producer"] for r in catalog)
asd = [r for r in catalog if r["producer"] == "АСД"]
entro_aria = [r for r in catalog if r["producer"] == "Энтро" and "aria" in r.get("target_line", "").lower()]
hd_stubs = [r for r in catalog if r["producer"] == "Хаусдорс" and r.get("model_name", "").startswith("[MISSING]")]

asd_front = sum(1 for r in asd if r.get("image_url_front", ""))
asd_inside = sum(1 for r in asd if r.get("image_url_inside", ""))
asd_both = sum(1 for r in asd if r.get("image_url_front", "") and r.get("image_url_inside", ""))
bad_count = sum(1 for r in asd for bp in BAD_PATTERNS if bp in (r.get("image_url_front", "") or "").lower() or bp in (r.get("image_url_inside", "") or "").lower())

print(f"\n=== Summary ===")
print(f"Total: {len(catalog)}")
print(f"ASD: {len(asd)}, front={asd_front}, inside={asd_inside}, both={asd_both}")
print(f"Entro Aria: {len(entro_aria)}")
print(f"Hausdorss stubs: {len(hd_stubs)}")
print(f"Bad patterns: {bad_count}")

with open(os.path.join(OUT_DIR, "doors_catalog.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=FIELD_NAMES, delimiter=";")
    w.writeheader()
    w.writerows(catalog)
with open(os.path.join(OUT_DIR, "wordpress_import.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=FIELD_NAMES, delimiter=";")
    w.writeheader()
    w.writerows(catalog)

try:
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    for ci, h in enumerate(FIELD_NAMES, 1): ws.cell(row=1, column=ci, value=h)
    for ri, r in enumerate(catalog, 2):
        for ci, h in enumerate(FIELD_NAMES, 1): ws.cell(row=ri, column=ci, value=r.get(h, ""))
    wb.save(os.path.join(OUT_DIR, "doors_catalog.xlsx"))
except: pass

for prod in sorted(producers.keys()):
    prod_rows = [r for r in catalog if r["producer"] == prod]
    safe = prod.replace("/", "_").replace("\\", "_")
    with open(os.path.join(OUT_DIR, "check_by_producer", f"{safe}.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELD_NAMES, delimiter=";")
        w.writeheader()
        w.writerows(prod_rows)

if os.path.isdir(FINAL_DIR):
    for fn in os.listdir(FINAL_DIR):
        src = os.path.join(FINAL_DIR, fn)
        if os.path.isfile(src) and fn.endswith((".xlsx", ".csv", ".json", ".html")):
            shutil.copy2(src, os.path.join(OUT_DIR, fn))

needs_review = Counter(r["producer"] for r in catalog if r.get("status", "") == "needs_review")
report = {"generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "total_rows": len(catalog), "rows_by_producer": dict(producers), "duplicates": 0,
    "asd": {"total": len(asd), "front": asd_front, "inside": asd_inside, "both": asd_both},
    "entro_aria": len(entro_aria), "hd_stubs": len(hd_stubs), "bad_patterns": bad_count,
    "needs_review": dict(needs_review)}
with open(os.path.join(OUT_DIR, "run_report.json"), "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)

zip_name = f"doors_parser_final_REWORK_READY_v2_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.zip"
zip_path = os.path.join(BASE, "output", zip_name)
with ZipFile(zip_path, "w") as zf:
    for root, dirs, files in os.walk(OUT_DIR):
        for fn in files:
            fp = os.path.join(root, fn)
            zf.write(fp, arcname=os.path.join("output", "final_rework_ready_v2", os.path.relpath(fp, OUT_DIR)))

sha = hashlib.sha256()
with open(zip_path, "rb") as f:
    for chunk in iter(lambda: f.read(65536), b""): sha.update(chunk)

print(f"\n{'='*60}")
print(f"Archive: {zip_path}")
print(f"SHA256: {sha.hexdigest()}")
print(f"ASD: front={asd_front}/{len(asd)}, inside={asd_inside}/{len(asd)}")
print(f"Bad: {bad_count}, Stubs: {len(hd_stubs)}, Aria: {len(entro_aria)}")
