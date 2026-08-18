"""Fix Бруклин + check Энтро Aria + rebuild v2 final"""
import csv, json, os, re, time, hashlib, shutil
from collections import Counter
from datetime import datetime
from zipfile import ZipFile
import requests
from bs4 import BeautifulSoup

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
OUT_DIR = os.path.join(BASE, "output", "final_rework_ready_v2")

with open(os.path.join(OUT_DIR, "doors_catalog.csv"), encoding="utf-8-sig") as f:
    catalog = list(csv.DictReader(f, delimiter=";"))

FIELD_NAMES = list(catalog[0].keys())
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

# ============================================================
# FIX 1: Бруклин — fetch correct image from source page
# ============================================================
print("=== FIX Бруклин ===")
for r in catalog:
    if r["producer"] == "АСД" and "Бруклин" in r.get("model_name", ""):
        source_url = r.get("source_url", "")
        print(f"  source: {source_url}")
        try:
            resp = requests.get(source_url, timeout=15, headers=HEADERS)
            if resp.status_code == 200:
                soup = BeautifulSoup(resp.text, "html.parser")
                # Find gallery images
                gallery_imgs = soup.select(".gallery img, [data-fancybox] img")
                fancybox_links = soup.select("a[data-fancybox], a.fancybox")
                
                gallery_urls = []
                for a in fancybox_links:
                    href = a.get("href", "")
                    if href and not any(x in href.lower() for x in ["fragment", "detail", "icon", "logo"]):
                        if href.startswith("http"):
                            gallery_urls.append(href)
                        elif href.startswith("/"):
                            gallery_urls.append("https://as-doors.ru" + href)
                        else:
                            gallery_urls.append("https://as-doors.ru/" + href)
                
                for img in gallery_imgs:
                    src = img.get("src") or img.get("data-src") or ""
                    if src and src not in gallery_urls:
                        if src.startswith("/"):
                            src = "https://as-doors.ru" + src
                        elif not src.startswith("http"):
                            src = "https://as-doors.ru/" + src
                        if not any(x in src.lower() for x in ["fragment", "detail", "icon", "logo"]):
                            gallery_urls.append(src)
                
                print(f"  Gallery URLs: {len(gallery_urls)}")
                for u in gallery_urls[:5]:
                    print(f"    {u[:70]}")
                
                if gallery_urls:
                    # Use first gallery image as front
                    r["image_url"] = gallery_urls[0]
                    r["image_url_front"] = gallery_urls[0]
                    if len(gallery_urls) > 1:
                        r["image_url_inside"] = gallery_urls[1]
                    r["status"] = "found"
                    r["comment"] = (r.get("comment", "") or "") + "; cache_image_replaced_from_gallery"
                    print(f"  FIXED: front={gallery_urls[0][:60]}")
        except Exception as e:
            print(f"  Error: {e}")
        break

# ============================================================
# FIX 2: Энтро Aria — deeper scan
# ============================================================
print("\n=== FIX Энтро Aria ===")
aria_rows = [r for r in catalog if r["producer"] == "Энтро" and "aria" in r.get("target_line", "").lower()]
print(f"  Current Aria rows: {len(aria_rows)}")

# Try deeper scan of entro-design.ru
try:
    # First check the catalog page
    resp = requests.get("https://entro-design.ru/catalog", timeout=15, headers=HEADERS)
    if resp.status_code == 200:
        soup = BeautifulSoup(resp.text, "html.parser")
        # Find all links that might be Aria
        all_links = soup.find_all("a", href=True)
        aria_links = []
        for a in all_links:
            href = a.get("href", "")
            text = a.get_text(strip=True)
            if "aria" in href.lower() or "aria" in text.lower() or "арея" in text.lower():
                aria_links.append((href, text))
        
        print(f"  Aria-related links found: {len(aria_links)}")
        for href, text in aria_links[:10]:
            print(f"    {text[:40]} -> {href[:50]}")
        
        # Also check if there's a sub-catalog for Aria
        # Look for category/collection pages
        cat_links = [a for a in all_links if "/catalog/" in a.get("href", "")]
        print(f"  Catalog sub-links: {len(cat_links)}")
        for a in cat_links[:15]:
            href = a.get("href", "")
            text = a.get_text(strip=True)[:30]
            print(f"    {text} -> {href[:50]}")
except Exception as e:
    print(f"  Error: {e}")

# Try direct Aria product pages
aria_product_urls = [
    "https://entro-design.ru/catalog/aria/aria-1",
    "https://entro-design.ru/catalog/aria/aria-2",
    "https://entro-design.ru/catalog/aria/aria-3",
    "https://entro-design.ru/catalog/aria-1",
    "https://entro-design.ru/catalog/aria-2",
    "https://entro-design.ru/catalog/aria-3",
    "https://entro-design.ru/product/aria-1",
    "https://entro-design.ru/product/aria-2",
    "https://entro-design.ru/product/aria-3",
]

aria_found = []
for url in aria_product_urls:
    try:
        resp = requests.get(url, timeout=10, headers=HEADERS, allow_redirects=True)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, "html.parser")
            title = soup.find("title")
            title_text = title.get_text() if title else ""
            h1 = soup.find("h1")
            h1_text = h1.get_text(strip=True) if h1 else ""
            img = soup.select_one(".product img, .gallery img, [data-fancybox] img")
            img_src = ""
            if img:
                img_src = img.get("src") or img.get("data-src") or ""
            
            if "aria" in (h1_text + title_text).lower():
                print(f"  FOUND: {url}")
                print(f"    title: {title_text[:50]}")
                print(f"    h1: {h1_text[:50]}")
                print(f"    img: {img_src[:50]}")
                aria_found.append({
                    "url": url,
                    "name": h1_text or title_text,
                    "img": img_src
                })
    except:
        pass

print(f"\n  Aria product pages found: {len(aria_found)}")

# Add found Aria models to catalog
aria_added = 0
for af in aria_found:
    model_name = af["name"]
    exists = any(r["producer"] == "Энтро" and r["model_name"] == model_name for r in catalog)
    if not exists:
        img = af["img"]
        if img and not img.startswith("http"):
            img = "https://entro-design.ru" + img if img.startswith("/") else "https://entro-design.ru/" + img
        catalog.append({
            "door_type": "межкомнатная",
            "producer": "Энтро",
            "target_line": "Aria",
            "model_name": model_name,
            "description": "",
            "material": "",
            "coating": "ПВХ, Эмаль",
            "image_url": img,
            "source_url": af["url"],
            "status": "found" if img else "needs_review",
            "comment": "added_from_site_scan_aria_models",
            "image_url_front": img,
            "image_url_inside": ""
        })
        aria_added += 1
        print(f"  Added: {model_name[:40]}")

print(f"  Aria models added: {aria_added}")

# ============================================================
# Deduplicate + save final v2
# ============================================================
print("\n=== FINAL SAVE ===")
seen = set()
deduped = []
for r in catalog:
    key = (r["producer"], r["target_line"], r["model_name"])
    if key not in seen:
        seen.add(key)
        deduped.append(r)
catalog = deduped

# Check bad patterns
BAD_PATTERNS = ["cache_image", "fragment", "handle", "ruchka", "lock", "zamok",
                "interior", "badge", "banner", "logo", "404"]
bad_count = 0
for r in catalog:
    if r["producer"] == "АСД":
        for bp in BAD_PATTERNS:
            if bp in (r.get("image_url_front", "") or "").lower() or bp in (r.get("image_url_inside", "") or "").lower():
                bad_count += 1
                break

producers = Counter(r["producer"] for r in catalog)
asd = [r for r in catalog if r["producer"] == "АСД"]
entro_aria = [r for r in catalog if r["producer"] == "Энтро" and "aria" in r.get("target_line", "").lower()]
hd_stubs = [r for r in catalog if r["producer"] == "Хаусдорс" and r.get("model_name", "").startswith("[MISSING]")]

print(f"Total: {len(catalog)}")
print(f"ASD: {len(asd)}, front={sum(1 for r in asd if r.get('image_url_front',''))}, inside={sum(1 for r in asd if r.get('image_url_inside',''))}")
print(f"Энтро Aria: {len(entro_aria)}")
print(f"Хаусdorf stubs: {len(hd_stubs)}")
print(f"Bad patterns: {bad_count}")

# Save
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
    ws.title = "Catalog v2"
    for ci, h in enumerate(FIELD_NAMES, 1):
        ws.cell(row=1, column=ci, value=h)
    for ri, r in enumerate(catalog, 2):
        for ci, h in enumerate(FIELD_NAMES, 1):
            ws.cell(row=ri, column=ci, value=r.get(h, ""))
    wb.save(os.path.join(OUT_DIR, "doors_catalog.xlsx"))
except ImportError:
    pass

# check_by_producer
for producer in sorted(producers.keys()):
    producer_rows = [r for r in catalog if r["producer"] == producer]
    safe_name = producer.replace("/", "_").replace("\\", "_")
    csv_path = os.path.join(OUT_DIR, "check_by_producer", f"{safe_name}.csv")
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELD_NAMES, delimiter=";")
        w.writeheader()
        w.writerows(producer_rows)
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = safe_name[:30]
        for ci, h in enumerate(FIELD_NAMES, 1):
            ws.cell(row=1, column=ci, value=h)
        for ri, r in enumerate(producer_rows, 2):
            for ci, h in enumerate(FIELD_NAMES, 1):
                ws.cell(row=ri, column=ci, value=r.get(h, ""))
        wb.save(os.path.join(OUT_DIR, "check_by_producer", f"{safe_name}.xlsx"))
    except ImportError:
        pass

# Run report
needs_review = Counter()
for r in catalog:
    if r.get("status", "") == "needs_review":
        needs_review[r["producer"]] += 1

report = {
    "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "mode": "final_rework_v2_front_inside_and_missing_models",
    "total_rows": len(catalog),
    "rows_by_producer": dict(producers),
    "duplicates_count": 0,
    "expected_lines_total": 194,
    "expected_lines_covered": 194 - len(hd_stubs),
    "expected_lines_stubs_remaining": len(hd_stubs),
    "asd_front_inside_stats": {
        "total_asd_rows": len(asd),
        "front_filled": sum(1 for r in asd if r.get("image_url_front", "")),
        "inside_filled": sum(1 for r in asd if r.get("image_url_inside", "")),
        "both_filled": sum(1 for r in asd if r.get("image_url_front", "") and r.get("image_url_inside", "")),
        "front_missing": sum(1 for r in asd if not r.get("image_url_front", "")),
        "inside_missing": sum(1 for r in asd if not r.get("image_url_inside", "")),
    },
    "entro_aria_rows_before": 1,
    "entro_aria_rows_after": len(entro_aria),
    "hausdoors_missing_stubs_before": 4,
    "hausdoors_missing_stubs_after": len(hd_stubs),
    "bad_patterns_count": bad_count,
    "needs_review_by_producer": dict(needs_review),
}

with open(os.path.join(OUT_DIR, "run_report.json"), "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)

# ZIP
zip_name = f"doors_parser_final_REWORK_READY_v2_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.zip"
zip_path = os.path.join(BASE, "output", zip_name)
with ZipFile(zip_path, "w") as zf:
    for root, dirs, files in os.walk(OUT_DIR):
        for fn in files:
            fp = os.path.join(root, fn)
            arcname = os.path.relpath(fp, OUT_DIR)
            zf.write(fp, arcname=os.path.join("output", "final_rework_ready_v2", arcname))

sha = hashlib.sha256()
with open(zip_path, "rb") as f:
    for chunk in iter(lambda: f.read(65536), b""):
        sha.update(chunk)

print(f"\n{'='*60}")
print(f"FINAL v2 REPORT")
print(f"{'='*60}")
print(f"Archive: {zip_path}")
print(f"SHA256: {sha.hexdigest()}")
print(f"Total rows: {len(catalog)}")
print(f"ASD: front={report['asd_front_inside_stats']['front_filled']}/{len(asd)}, inside={report['asd_front_inside_stats']['inside_filled']}/{len(asd)}, both={report['asd_front_inside_stats']['both_filled']}")
print(f"Энтро Aria: before=1, after={len(entro_aria)}")
print(f"Хаусдорс stubs: before=4, after={len(hd_stubs)}")
print(f"Expected lines covered: {report['expected_lines_covered']}/194")
print(f"Bad patterns: {bad_count}")
print(f"Needs review: {dict(needs_review)}")
