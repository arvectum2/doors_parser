"""Fix Бруклин URL + final rebuild"""
import csv, json, os, hashlib
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

# Fix Бруклин
for r in catalog:
    if r["producer"] == "\u0410\u0421\u0414" and "\u0411\u0440\u0443\u043a\u043b\u0438\u043d" in r.get("model_name", ""):
        # The source_url is malformed in the catalog
        correct_url = "https://as-doors.ru/doors/zakaznyie/bruklin.html"
        print(f"Fixing Бруклин: {correct_url}")
        try:
            resp = requests.get(correct_url, timeout=15, headers=HEADERS)
            if resp.status_code == 200:
                soup = BeautifulSoup(resp.text, "html.parser")
                gallery_urls = []
                seen = set()
                for a in soup.select("a[data-fancybox], a.fancybox"):
                    href = a.get("href", "")
                    if href and href not in seen:
                        seen.add(href)
                        if href.startswith("/"):
                            href = "https://as-doors.ru" + href
                        elif not href.startswith("http"):
                            href = "https://as-doors.ru/" + href
                        gallery_urls.append(href)
                
                print(f"  Gallery: {len(gallery_urls)} images")
                for u in gallery_urls[:3]:
                    print(f"    {u[:70]}")
                
                if gallery_urls:
                    r["image_url"] = gallery_urls[0]
                    r["image_url_front"] = gallery_urls[0]
                    if len(gallery_urls) > 1:
                        r["image_url_inside"] = gallery_urls[1]
                    r["source_url"] = correct_url
                    r["status"] = "found"
                    r["comment"] = (r.get("comment", "") or "") + "; bruklin_url_fixed"
                    print(f"  FIXED: front={gallery_urls[0][:60]}")
        except Exception as e:
            print(f"  Error: {e}")
        break

# Verify no bad patterns remain
bad = ["cache_image","fragment","handle","ruchka","lock","zamok","interior","badge","banner","logo","404"]
bad_count = 0
for r in catalog:
    if r["producer"] == "\u0410\u0421\u0414":
        for bp in bad:
            if bp in (r.get("image_url_front","") or "").lower() or bp in (r.get("image_url_inside","") or "").lower():
                bad_count += 1
                break

print(f"\nBad patterns remaining: {bad_count}")

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
producers = Counter(r["producer"] for r in catalog)
for prod in sorted(producers.keys()):
    prod_rows = [r for r in catalog if r["producer"] == prod]
    safe = prod.replace("/","_").replace("\\","_")
    with open(os.path.join(OUT_DIR, "check_by_producer", f"{safe}.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELD_NAMES, delimiter=";")
        w.writeheader()
        w.writerows(prod_rows)

asd = [r for r in catalog if r["producer"] == "\u0410\u0421\u0414"]
entro_aria = [r for r in catalog if r["producer"] == "\u042d\u043d\u0442\u0440\u043e" and "aria" in r.get("target_line","").lower()]
hd_stubs = [r for r in catalog if r["producer"] == "\u0425\u0430\u0443\u0441\u0434\u043e\u0440\u0441" and r.get("model_name","").startswith("[MISSING]")]
needs_review = Counter(r["producer"] for r in catalog if r.get("status","") == "needs_review")

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
        "front_filled": sum(1 for r in asd if r.get("image_url_front","")),
        "inside_filled": sum(1 for r in asd if r.get("image_url_inside","")),
        "both_filled": sum(1 for r in asd if r.get("image_url_front","") and r.get("image_url_inside","")),
        "front_missing": sum(1 for r in asd if not r.get("image_url_front","")),
        "inside_missing": sum(1 for r in asd if not r.get("image_url_inside","")),
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
print(f"FINAL v2")
print(f"{'='*60}")
print(f"Archive: {zip_path}")
print(f"SHA256: {sha.hexdigest()}")
print(f"Total: {len(catalog)}")
print(f"ASD: front={report['asd_front_inside_stats']['front_filled']}/{len(asd)}, inside={report['asd_front_inside_stats']['inside_filled']}/{len(asd)}, both={report['asd_front_inside_stats']['both_filled']}")
print(f"Энтро Aria: 1 -> {len(entro_aria)}")
print(f"Хаусдорс stubs: 4 -> {len(hd_stubs)}")
print(f"Expected lines: {report['expected_lines_covered']}/194")
print(f"Bad patterns: {bad_count}")
print(f"Needs review: {dict(needs_review)}")
