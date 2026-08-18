"""Stage 4: Targeted fixes — АСД inside photo, Бруклин, Энтро Aria, Хаусдорс missing lines"""
import csv, json, os, re, time, hashlib
from collections import Counter, defaultdict
from datetime import datetime
from zipfile import ZipFile
import requests
from bs4 import BeautifulSoup

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
FINAL_DIR = os.path.join(BASE, "output", "final_rework_ready")
OUT_DIR = os.path.join(BASE, "output", "final_rework_ready_v2")
os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(os.path.join(OUT_DIR, "check_by_producer"), exist_ok=True)

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
TIMEOUT = 15

# Load catalog
with open(os.path.join(FINAL_DIR, "doors_catalog.csv"), encoding="utf-8-sig") as f:
    catalog = list(csv.DictReader(f, delimiter=";"))

FIELD_NAMES = list(catalog[0].keys())

# ============================================================
# 1. АСД — fetch inside photos from source_url
# ============================================================
print("=" * 60)
print("STAGE 4.1: АСД inside photo scan")
print("=" * 60)

asd_rows = [r for r in catalog if r["producer"] == "АСД"]
asd_with_source = [r for r in asd_rows if r.get("source_url", "")]
print(f"АСД rows: {len(asd_rows)}, with source_url: {len(asd_with_source)}")

BAD_PATTERNS = ["cache_image", "fragment", "handle", "ruchka", "lock", "zamok",
                "interior", "badge", "banner", "logo", "404"]

def is_full_door_image(src):
    """Check if URL looks like a full door image (not fragment/technical)"""
    sl = src.lower()
    for bp in BAD_PATTERNS:
        if bp in sl:
            return False
    if any(x in sl for x in [".png", ".gif", ".svg", "icon", "logo", "arrow", "star"]):
        return False
    return True

def fetch_asd_gallery(source_url):
    """Fetch АСД page and extract gallery images"""
    try:
        resp = requests.get(source_url, timeout=TIMEOUT, headers=HEADERS)
        if resp.status_code != 200:
            return [], 0, 0, 0
        soup = BeautifulSoup(resp.text, "html.parser")
        
        # Find gallery container images (the main product images)
        gallery_imgs = soup.select(".gallery img, [data-fancybox] img")
        
        # Also find fancybox links for full-size versions
        fancybox_links = soup.select("a[data-fancybox], a.fancybox")
        link_hrefs = [a.get("href", "") for a in fancybox_links if a.get("href", "")]
        
        # Collect unique image URLs from gallery
        seen = set()
        gallery_urls = []
        
        # First try gallery links (full-size)
        for href in link_hrefs:
            if href not in seen and is_full_door_image(href):
                seen.add(href)
                gallery_urls.append(href)
        
        # Then add gallery img src (may be thumbnails)
        for img in gallery_imgs:
            src = img.get("src") or img.get("data-src") or ""
            if src and src not in seen and is_full_door_image(src):
                seen.add(src)
                gallery_urls.append(src)
        
        # Count rejected
        rejected_tech = 0
        rejected_frag = 0
        rejected_interior = 0
        
        all_links = soup.select("a[href]")
        for a in all_links:
            href = a.get("href", "")
            if href and href not in seen:
                sl = href.lower()
                if any(x in sl for x in ["fragment", "detail", "close", "macro"]):
                    rejected_frag += 1
                elif any(x in sl for x in ["interior", "room", "inside_view"]):
                    rejected_interior += 1
                elif any(x in sl for x in ["technical", "scheme", "diagram"]):
                    rejected_tech += 1
        
        return gallery_urls, rejected_tech, rejected_frag, rejected_interior
    except Exception as e:
        return [], 0, 0, 0

def resolve_url(base_url, src):
    """Resolve relative URL"""
    if src.startswith("http"):
        return src
    if src.startswith("//"):
        return "https:" + src
    if src.startswith("/"):
        from urllib.parse import urlparse
        parsed = urlparse(base_url)
        return f"{parsed.scheme}://{parsed.netloc}{src}"
    return base_url.rstrip("/") + "/" + src.lstrip("/")

# Process АСД rows
asd_fixes = []
fixed_count = 0
for i, r in enumerate(asd_with_source):
    source_url = r["source_url"]
    old_img = r.get("image_url", "") or ""
    
    gallery, tech, frag, interior = fetch_asd_gallery(source_url)
    
    front = ""
    inside = ""
    front_decision = "not_found"
    inside_decision = "not_found"
    
    if len(gallery) >= 2:
        front = resolve_url(source_url, gallery[0])
        inside = resolve_url(source_url, gallery[1])
        front_decision = "full_door_gallery_1"
        inside_decision = "full_door_gallery_2"
    elif len(gallery) == 1:
        front = resolve_url(source_url, gallery[0])
        front_decision = "full_door_gallery_1_only"
        inside_decision = "not_found"
    
    # Update catalog
    if front:
        r["image_url_front"] = front
        r["image_url"] = front
    if inside:
        r["image_url_inside"] = inside
    
    if inside:
        r["status"] = "found"
        old_comment = r.get("comment", "") or ""
        r["comment"] = f"inside_photo_added; {old_comment}" if old_comment else "inside_photo_added"
        fixed_count += 1
    
    asd_fixes.append({
        "producer": "АСД",
        "target_line": r.get("target_line", ""),
        "model_name": r.get("model_name", ""),
        "source_url": source_url,
        "old_image_url": old_img,
        "image_url_front": front,
        "image_url_inside": inside,
        "front_decision": front_decision,
        "inside_decision": inside_decision,
        "rejected_technical_count": tech,
        "rejected_fragment_count": frag,
        "rejected_interior_count": interior,
        "notes": f"gallery_size={len(gallery)}"
    })
    
    if (i + 1) % 20 == 0:
        print(f"  Processed {i+1}/{len(asd_with_source)}, inside found: {fixed_count}")
        time.sleep(1)

print(f"  АСД total processed: {len(asd_fixes)}")
print(f"  Inside photos found: {fixed_count}")

# Save АСД audit
asd_audit_fields = ["producer","target_line","model_name","source_url","old_image_url",
    "image_url_front","image_url_inside","front_decision","inside_decision",
    "rejected_technical_count","rejected_fragment_count","rejected_interior_count","notes"]

with open(os.path.join(OUT_DIR, "asd_front_inside_audit.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=asd_audit_fields)
    w.writeheader()
    w.writerows(asd_fixes)

# ============================================================
# 2. АСД — fix Бруклин cache_image
# ============================================================
print("\n" + "=" * 60)
print("STAGE 4.2: АСД Бруклин cache_image fix")
print("=" * 60)

bruklin_row = None
for r in catalog:
    if r["producer"] == "АСД" and "Бруклин" in r.get("model_name", ""):
        bruklin_row = r
        break

if bruklin_row:
    url = bruklin_row.get("image_url", "") or ""
    print(f"  Found: {bruklin_row['model_name'][:40]}")
    print(f"  URL: {url[:60]}")
    
    # Check if it's a full door by fetching dimensions
    try:
        resp = requests.get(url, timeout=TIMEOUT, headers=HEADERS)
        if resp.status_code == 200 and len(resp.content) > 1000:
            # Check content type
            ct = resp.headers.get("content-type", "")
            print(f"  Content-Type: {ct}, size: {len(resp.content)}")
            
            # Try to get dimensions
            from io import BytesIO
            try:
                from PIL import Image
                img = Image.open(BytesIO(resp.content))
                w, h = img.size
                ratio = h / w if w > 0 else 0
                print(f"  Dimensions: {w}x{h}, ratio: {ratio:.2f}")
                
                if ratio >= 1.5:
                    # Full door image
                    bruklin_row["comment"] = (bruklin_row.get("comment", "") or "") + "; cache_image_full_door_verified"
                    print(f"  Verdict: FULL DOOR — keeping")
                else:
                    # Fragment
                    bruklin_row["image_url"] = ""
                    bruklin_row["image_url_front"] = ""
                    bruklin_row["status"] = "needs_review"
                    bruklin_row["comment"] = (bruklin_row.get("comment", "") or "") + "; cache_image_fragment_cleared"
                    print(f"  Verdict: FRAGMENT — cleared")
            except ImportError:
                # No PIL, try to find correct image on source page
                print(f"  PIL not available, checking source page for correct image")
                source_url = bruklin_row.get("source_url", "")
                if source_url:
                    gallery, _, _, _ = fetch_asd_gallery(source_url)
                    if gallery:
                        correct_url = resolve_url(source_url, gallery[0])
                        bruklin_row["image_url"] = correct_url
                        bruklin_row["image_url_front"] = correct_url
                        bruklin_row["comment"] = (bruklin_row.get("comment", "") or "") + "; cache_image_replaced_from_gallery"
                        print(f"  Replaced with: {correct_url[:60]}")
        else:
            print(f"  Failed to fetch or too small")
    except Exception as e:
        print(f"  Error: {e}")
else:
    print("  Бруклин row not found")

# ============================================================
# 3. Энтро — restore Aria models
# ============================================================
print("\n" + "=" * 60)
print("STAGE 4.3: Энтро Aria models restore")
print("=" * 60)

entro_aria_rows = [r for r in catalog if r["producer"] == "Энтро" and "aria" in r.get("target_line", "").lower()]
print(f"  Current Aria rows: {len(entro_aria_rows)}")

# Try to find Aria page on entro-design.ru
aria_models_found = []
try:
    # Common Aria URLs
    aria_urls = [
        "https://entro-design.ru/catalog/aria",
        "https://entro-design.ru/catalog/aria/",
        "https://entro-design.ru/catalog/dveri-aria",
        "https://entro-design.ru/catalog/dveri-aria/",
    ]
    
    for aria_url in aria_urls:
        resp = requests.get(aria_url, timeout=TIMEOUT, headers=HEADERS)
        if resp.status_code == 200:
            print(f"  Found Aria page: {aria_url}")
            soup = BeautifulSoup(resp.text, "html.parser")
            
            # Find model cards/links
            model_links = soup.select("a[href*=aria], .product-card a, .catalog-item a")
            seen_models = set()
            for a in model_links:
                href = a.get("href", "")
                text = a.get_text(strip=True)
                if text and "aria" in text.lower() and text not in seen_models:
                    seen_models.add(text)
                    # Find image
                    img = a.find("img")
                    img_src = ""
                    if img:
                        img_src = img.get("src") or img.get("data-src") or ""
                    full_url = href if href.startswith("http") else f"https://entro-design.ru{href}"
                    aria_models_found.append({
                        "model_name": text,
                        "source_url": full_url,
                        "image_url": img_src
                    })
            
            if not aria_models_found:
                # Try finding all product items
                items = soup.select(".product-item, .catalog-item, .card, [class*=product]")
                print(f"  Product items found: {len(items)}")
                for item in items[:10]:
                    a = item.find("a")
                    if a:
                        href = a.get("href", "")
                        text = a.get_text(strip=True)[:50]
                        img = item.find("img")
                        img_src = img.get("src", "") if img else ""
                        print(f"    {text} -> {href[:50]} img={img_src[:40]}")
            break
except Exception as e:
    print(f"  Error scanning Aria: {e}")

print(f"  Aria models found on site: {len(aria_models_found)}")
for m in aria_models_found:
    print(f"    {m['model_name'][:30]} | {m['source_url'][:50]}")

# If models found, add them to catalog
aria_added = 0
if aria_models_found:
    for m in aria_models_found:
        # Check if already exists
        exists = any(r["producer"] == "Энтро" and r["model_name"] == m["model_name"] for r in catalog)
        if not exists:
            catalog.append({
                "door_type": "межкомнатная",
                "producer": "Энтро",
                "target_line": "Aria",
                "model_name": m["model_name"],
                "description": "",
                "material": "",
                "coating": "ПВХ, Эмаль",
                "image_url": m.get("image_url", ""),
                "source_url": m.get("source_url", ""),
                "status": "found" if m.get("image_url") else "needs_review",
                "comment": "added_from_site_scan_aria_collection",
                "image_url_front": m.get("image_url", ""),
                "image_url_inside": ""
            })
            aria_added += 1

print(f"  Aria models added: {aria_added}")

# Save Энтро Aria report
entro_report = []
for r in catalog:
    if r["producer"] == "Энтро" and "aria" in r.get("target_line", "").lower():
        entro_report.append({
            "model_name": r["model_name"],
            "source_url": r.get("source_url", ""),
            "image_url_front": r.get("image_url_front", ""),
            "found_in_site": "yes",
            "added_to_catalog": "yes" if "added_from_site" in (r.get("comment", "") or "") else "already_existed",
            "status": r.get("status", ""),
            "comment": r.get("comment", "")
        })

entro_fields = ["model_name","source_url","image_url_front","found_in_site","added_to_catalog","status","comment"]
with open(os.path.join(OUT_DIR, "entro_aria_rebuild_report.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=entro_fields)
    w.writeheader()
    w.writerows(entro_report)

# ============================================================
# 4. Хаусдорс — replace [MISSING] stubs
# ============================================================
print("\n" + "=" * 60)
print("STAGE 4.4: Хаусдорс missing lines rebuild")
print("=" * 60)

hausdoors_stubs = [r for r in catalog if r["producer"] == "Хаусдорс" and r.get("model_name", "").startswith("[MISSING]")]
print(f"  Current stubs: {len(hausdoors_stubs)}")

# Known Хаусдорс line URLs from config
hd_line_urls = {
    "Фактурные линии": ["https://hausdoors.ru/dveri-facturnye-linii", "https://hausdoors.ru/dveri-fakturnye-linii"],
    "Новая классика": ["https://hausdoors.ru/dveri-new-classic"],
    "Минимализм": ["https://hausdoors.ru/dveri-minimalism"],
    "Современная классика": ["https://hausdoors.ru/dveri-soft-classic"],
}

hausdoors_report = []
for stub in hausdoors_stubs:
    line_name = stub["target_line"]
    urls = hd_line_urls.get(line_name, [])
    
    models_found = []
    source_found = ""
    
    for url in urls:
        try:
            resp = requests.get(url, timeout=TIMEOUT, headers=HEADERS)
            if resp.status_code == 200:
                source_found = url
                soup = BeautifulSoup(resp.text, "html.parser")
                
                # Find product cards
                cards = soup.select(".product-card, .catalog-item, .card, [class*=product], .item")
                for card in cards:
                    a = card.find("a")
                    if a:
                        href = a.get("href", "")
                        text = a.get_text(strip=True)[:80]
                        img = card.find("img")
                        img_src = img.get("src") or img.get("data-src") if img else ""
                        if text and len(text) > 2:
                            full_url = href if href.startswith("http") else f"https://hausdoors.ru{href}"
                            models_found.append({
                                "model_name": text,
                                "source_url": full_url,
                                "image_url": img_src
                            })
                
                if not models_found:
                    # Try finding any links with door-related text
                    all_links = soup.select("a[href*=dveri], a[href*=door]")
                    for a in all_links:
                        text = a.get_text(strip=True)[:80]
                        href = a.get("href", "")
                        if text and len(text) > 3 and line_name.split()[0].lower() in text.lower():
                            full_url = href if href.startswith("http") else f"https://hausdoors.ru{href}"
                            models_found.append({
                                "model_name": text,
                                "source_url": full_url,
                                "image_url": ""
                            })
                break
        except Exception as e:
            continue
    
    # Update stub
    if models_found:
        # Replace stub with real models
        stub["status"] = "found"
        stub["comment"] = f"site_scan_found_{len(models_found)}_models"
        
        for m in models_found:
            exists = any(r["producer"] == "Хаусдорс" and r["model_name"] == m["model_name"] for r in catalog)
            if not exists:
                catalog.append({
                    "door_type": "межкомнатная",
                    "producer": "Хаусдорс",
                    "target_line": line_name,
                    "model_name": m["model_name"],
                    "description": "",
                    "material": "",
                    "coating": stub.get("coating", ""),
                    "image_url": m.get("image_url", ""),
                    "source_url": m.get("source_url", ""),
                    "status": "found" if m.get("image_url") else "needs_review",
                    "comment": f"added_from_site_scan_{line_name}",
                    "image_url_front": m.get("image_url", ""),
                    "image_url_inside": ""
                })
        
        hausdoors_report.append({
            "expected_line": line_name,
            "expected_covering": stub.get("coating", ""),
            "source_url_found": source_found,
            "models_found_count": len(models_found),
            "models_added_count": len(models_found),
            "stub_kept": "no",
            "reason": f"found_{len(models_found)}_models_on_site"
        })
    else:
        stub["comment"] = (stub.get("comment", "") or "") + "; source_line_not_found_on_site_after_scan"
        hausdoors_report.append({
            "expected_line": line_name,
            "expected_covering": stub.get("coating", ""),
            "source_url_found": source_found,
            "models_found_count": 0,
            "models_added_count": 0,
            "stub_kept": "yes",
            "reason": "source_line_not_found_on_site_after_scan"
        })

hd_fields = ["expected_line","expected_covering","source_url_found","models_found_count",
             "models_added_count","stub_kept","reason"]
with open(os.path.join(OUT_DIR, "hausdoors_missing_lines_rebuild_report.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=hd_fields)
    w.writeheader()
    w.writerows(hausdoors_report)

print(f"  Хаусдорс stubs processed: {len(hausdoors_report)}")
for hr in hausdoors_report:
    print(f"    {hr['expected_line']}: {hr['reason']}")

# ============================================================
# 5. Rebuild final catalog v2
# ============================================================
print("\n" + "=" * 60)
print("STAGE 4.5: Rebuild final catalog v2")
print("=" * 60)

# Remove [MISSING] stubs that were replaced
catalog = [r for r in catalog if not (r["producer"] == "Хаусдорс" and r.get("model_name", "").startswith("[MISSING]") and r.get("status") == "found")]

# Deduplicate
seen = set()
deduped = []
for r in catalog:
    key = (r["producer"], r["target_line"], r["model_name"])
    if key not in seen:
        seen.add(key)
        deduped.append(r)
catalog = deduped

# Stats
producers = Counter(r["producer"] for r in catalog)
asd_final = [r for r in catalog if r["producer"] == "АСД"]
entro_aria_final = [r for r in catalog if r["producer"] == "Энтро" and "aria" in r.get("target_line", "").lower()]
hd_stubs_remaining = [r for r in catalog if r["producer"] == "Хаусдорс" and r.get("model_name", "").startswith("[MISSING]")]

print(f"Total rows: {len(catalog)}")
print(f"Producers: {dict(producers)}")
print(f"ASD: {len(asd_final)} rows")
print(f"ASD with front: {sum(1 for r in asd_final if r.get('image_url_front',''))}")
print(f"ASD with inside: {sum(1 for r in asd_final if r.get('image_url_inside',''))}")
print(f"Энтро Aria: {len(entro_aria_final)} rows")
print(f"Хаусdorf stubs remaining: {len(hd_stubs_remaining)}")

# Check bad patterns
bad_count = 0
for r in asd_final:
    for bp in BAD_PATTERNS:
        if bp in (r.get("image_url_front", "") or "").lower() or bp in (r.get("image_url_inside", "") or "").lower():
            bad_count += 1
            break
print(f"Bad patterns in АСД front/inside: {bad_count}")

# Save final catalog
with open(os.path.join(OUT_DIR, "doors_catalog.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=FIELD_NAMES, delimiter=";")
    w.writeheader()
    w.writerows(catalog)

# XLSX
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

# WordPress
with open(os.path.join(OUT_DIR, "wordpress_import.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=FIELD_NAMES, delimiter=";")
    w.writeheader()
    w.writerows(catalog)

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
needs_review_by_producer = Counter()
for r in catalog:
    if r.get("status", "") == "needs_review":
        needs_review_by_producer[r["producer"]] += 1

report = {
    "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "mode": "final_rework_v2_front_inside_and_missing_models",
    "total_rows": len(catalog),
    "producers_count": len(producers),
    "rows_by_producer": dict(producers),
    "duplicates_count": 0,
    "expected_lines_total": 194,
    "expected_lines_covered": 194 - len(hd_stubs_remaining),
    "expected_lines_stubs_remaining": len(hd_stubs_remaining),
    "asd_front_inside_stats": {
        "total_asd_rows": len(asd_final),
        "front_filled": sum(1 for r in asd_final if r.get("image_url_front", "")),
        "inside_filled": sum(1 for r in asd_final if r.get("image_url_inside", "")),
        "both_filled": sum(1 for r in asd_final if r.get("image_url_front", "") and r.get("image_url_inside", "")),
        "front_missing": sum(1 for r in asd_final if not r.get("image_url_front", "")),
        "inside_missing": sum(1 for r in asd_final if not r.get("image_url_inside", "")),
    },
    "entro_aria_rows_before": 1,
    "entro_aria_rows_after": len(entro_aria_final),
    "hausdoors_missing_stubs_before": 4,
    "hausdoors_missing_stubs_after": len(hd_stubs_remaining),
    "bad_patterns_count": bad_count,
    "needs_review_by_producer": dict(needs_review_by_producer),
    "asd_inside_photos_added": fixed_count,
}

with open(os.path.join(OUT_DIR, "run_report.json"), "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)

# Copy audit/patch files
import shutil
for fn in os.listdir(FINAL_DIR):
    src = os.path.join(FINAL_DIR, fn)
    dst = os.path.join(OUT_DIR, fn)
    if os.path.isfile(src) and fn.endswith((".xlsx", ".csv", ".json", ".html")):
        shutil.copy2(src, dst)

# Build ZIP
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
print(f"STAGE 4 COMPLETE — FINAL REPORT")
print(f"{'='*60}")
print(f"Archive: {zip_path}")
print(f"SHA256: {sha.hexdigest()}")
print(f"Total rows: {len(catalog)}")
print(f"Rows by producer: {dict(producers)}")
print(f"\nАСД:")
print(f"  Front filled: {report['asd_front_inside_stats']['front_filled']}/{len(asd_final)}")
print(f"  Inside filled: {report['asd_front_inside_stats']['inside_filled']}/{len(asd_final)}")
print(f"  Both filled: {report['asd_front_inside_stats']['both_filled']}")
print(f"  Inside missing: {report['asd_front_inside_stats']['inside_missing']}")
print(f"  Inside photos added this run: {fixed_count}")
print(f"  Bad patterns: {bad_count}")
print(f"\nЭнтро Aria:")
print(f"  Before: 1 row")
print(f"  After: {len(entro_aria_final)} rows")
print(f"  Added: {aria_added}")
print(f"\nХаусдорс:")
print(f"  Stubs before: 4")
print(f"  Stubs after: {len(hd_stubs_remaining)}")
print(f"\nExpected lines covered: {report['expected_lines_covered']}/{194}")
print(f"Remaining needs_review by producer:")
for p, c in sorted(needs_review_by_producer.items()):
    print(f"  {p}: {c}")
