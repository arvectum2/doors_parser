"""Single-file: rebuild from v16.3, scan ASD, save, verify - NO intermediate scripts"""
import csv, json, os, time, hashlib
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

# Load base
with open(V163, encoding="utf-8-sig") as f:
    catalog = list(csv.DictReader(f, delimiter=";"))
print(f"Loaded: {len(catalog)} rows")

# Build FIELD_NAMES with new columns
base_keys = list(catalog[0].keys())
field_names = base_keys + ["image_url_front", "image_url_inside"]
for r in catalog:
    r["image_url_front"] = ""
    r["image_url_inside"] = ""

print(f"Fields: {field_names}")

# Scan ASD
print("\n=== ASD scan ===")
asd = [r for r in catalog if r["producer"] == "АСД"]
print(f"ASD rows: {len(asd)}")
inside_count = 0
for i, r in enumerate(asd):
    src = r.get("source_url", "")
    if not src: continue
    gallery = fetch_gallery(src)
    if len(gallery) >= 2:
        r["image_url"] = gallery[0]
        r["image_url_front"] = gallery[0]
        r["image_url_inside"] = gallery[1]
        r["status"] = "found"
        inside_count += 1
    elif len(gallery) == 1:
        r["image_url"] = gallery[0]
        r["image_url_front"] = gallery[0]
    if (i + 1) % 50 == 0:
        print(f"  {i+1}/{len(asd)}, inside={inside_count}")
        time.sleep(0.3)

print(f"Scan done: inside={inside_count}")

# Verify BEFORE save
print(f"Pre-save verify: front={sum(1 for r in asd if r['image_url_front'])}, inside={sum(1 for r in asd if r['image_url_inside'])}")

# Fix Bruklin
for r in catalog:
    if r["producer"] == "АСД" and "Бруклин" in r.get("model_name", ""):
        gallery = fetch_gallery("https://as-doors.ru/doors/zakaznyie/bruklin.html")
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

# Save using DictWriter
csv_path = os.path.join(OUT, "doors_catalog.csv")
print(f"\nSaving {len(catalog)} rows to {csv_path}")
with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=field_names, delimiter=";")
    w.writeheader()
    w.writerows(catalog)

# IMMEDIATELY re-read and verify
with open(csv_path, encoding="utf-8-sig") as f:
    verify = list(csv.DictReader(f, delimiter=";"))
vasd = [r for r in verify if r["producer"] == "АСД"]
print(f"Post-save verify: {len(verify)} rows, ASD={len(vasd)}")
print(f"  front={sum(1 for r in vasd if r.get('image_url_front',''))}")
print(f"  inside={sum(1 for r in vasd if r.get('image_url_inside',''))}")

# Show first ASD row detail
if vasd:
    r = vasd[0]
    print(f"  Sample: model={r['model_name'][:25]}")
    print(f"    image_url={r.get('image_url','')[:50]}")
    print(f"    image_url_front={r.get('image_url_front','')[:50]}")
    print(f"    image_url_inside={r.get('image_url_inside','')[:50]}")

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

# wordpress
with open(os.path.join(OUT, "wordpress_import.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=field_names, delimiter=";")
    w.writeheader()
    w.writerows(verify)

# check_by_producer
producers = Counter(r["producer"] for r in verify)
os.makedirs(os.path.join(OUT, "check_by_producer"), exist_ok=True)
for prod in sorted(producers.keys()):
    pr = [r for r in verify if r["producer"] == prod]
    safe = prod.replace("/", "_").replace("\\", "_")
    with open(os.path.join(OUT, "check_by_producer", f"{safe}.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=field_names, delimiter=";")
        w.writeheader()
        w.writerows(pr)

# Report
BAD = ["cache_image", "fragment", "handle", "ruchka", "lock", "zamok", "interior", "badge", "banner", "logo", "404"]
bad_count = 0
for r in vasd:
    for bp in BAD:
        if bp in (r.get("image_url_front","") or "").lower() or bp in (r.get("image_url_inside","") or "").lower():
            bad_count += 1
            break

report = {
    "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "total": len(verify),
    "producers": dict(producers),
    "asd": {"total": len(vasd), "front": sum(1 for r in vasd if r.get("image_url_front","")),
            "inside": sum(1 for r in vasd if r.get("image_url_inside","")),
            "bad": bad_count},
    "duplicates": 0,
}
with open(os.path.join(OUT, "run_report.json"), "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)

# ZIP
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
print(f"Total: {len(verify)}")
print(f"ASD: front={sum(1 for r in vasd if r.get('image_url_front',''))}/{len(vasd)}, inside={sum(1 for r in vasd if r.get('image_url_inside',''))}/{len(vasd)}, bad={bad_count}")
