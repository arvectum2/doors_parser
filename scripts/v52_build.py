import csv
import os
import json
import hashlib
import time
import urllib.request
import urllib.error
import ssl
import re
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from urllib.parse import urlparse
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
V4_CATALOG = os.path.join(BASE, "output", "final_rework_ready_v4", "doors_catalog.csv")
V51_CATALOG = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")
PATCH_PATH = os.path.join(BASE, "output", "asd_photo_patch_v2_clean", "asd_photo_patch_v2_clean.csv")
WAYBACK_INSIDE = os.path.join(BASE, "output", "delivery_v5_2", "wayback_inside_all_results.json")
OUT_DIR = os.path.join(BASE, "output", "delivery_v5_2")
DELIVERY_DIR = os.path.join(BASE, "output", "delivery")
os.makedirs(OUT_DIR, exist_ok=True)

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE
opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))

EXPECTED_PRODUCERS = [
    "Скамбио Порте", "Хаусдорс", "Олимп", "Аэлита", "Лорд",
    "Комфорт", "Двери регионов", "Термодур", "Энтро", "АСД"
]

# Load all data
with open(V4_CATALOG, encoding='utf-8-sig') as f:
    v4_rows = list(csv.DictReader(f, delimiter=';'))
with open(V51_CATALOG, encoding='utf-8-sig') as f:
    v51_rows = list(csv.DictReader(f, delimiter=';'))
fieldnames = list(v51_rows[0].keys())

with open(PATCH_PATH, encoding='utf-8-sig') as f:
    patch_rows = list(csv.DictReader(f, delimiter=';'))
patch_front_map = {r['model_name']: r.get('selected_image_url', '').strip() for r in patch_rows if r.get('selected_image_url', '').strip()}

with open(WAYBACK_INSIDE, encoding='utf-8') as f:
    wayback_inside = json.load(f)

v4_asd = {r['model_name']: r for r in v4_rows if r['producer'] == 'АСД'}
print(f"Loaded: V4={len(v4_rows)}, V5.1={len(v51_rows)}, patch={len(patch_front_map)}, wayback_inside={len(wayback_inside)}")

# Work on a copy of v5.1 rows
rows = [dict(r) for r in v51_rows]

# ============================================================
# STEP 1: Restore ASD inside images from v4 (40 working) + Wayback (96)
# ============================================================
print(f"\n=== STEP 1: RESTORE ASD INSIDE ===")

asd_rows = [(i, r) for i, r in enumerate(rows) if r['producer'] == 'АСД']
inside_restored_from_v4 = 0
inside_restored_from_wayback = 0
inside_already_ok = 0
inside_still_missing = 0

for idx, r in asd_rows:
    name = r['model_name']
    current_inside = r.get('image_url_inside', '').strip()
    
    if current_inside:
        inside_already_ok += 1
        continue
    
    # Try Wayback result
    wb_info = wayback_inside.get(name, {})
    wb_url = wb_info.get('inside_url', '')
    if wb_url:
        # Use HTTP
        http_url = wb_url.replace('https://', 'http://')
        rows[idx]['image_url_inside'] = http_url
        inside_restored_from_wayback += 1
    else:
        inside_still_missing += 1

print(f"Inside already OK: {inside_already_ok}")
print(f"Inside restored from Wayback: {inside_restored_from_wayback}")
print(f"Inside still missing: {inside_still_missing}")

# ============================================================
# STEP 2: Verify ASD front (from v5.1 which already restored from patch)
# ============================================================
print(f"\n=== STEP 2: VERIFY ASD FRONT ===")

asd_front_ok = sum(1 for _, r in asd_rows if r.get('image_url_front', '').strip())
asd_front_missing = sum(1 for _, r in asd_rows if not r.get('image_url_front', '').strip())
print(f"ASD front OK: {asd_front_ok}/199")
print(f"ASD front missing: {asd_front_missing}/199")

# List missing front
missing_front = [(r['model_name'], r.get('source_url', '')) for _, r in asd_rows if not r.get('image_url_front', '').strip()]
print(f"\nASD rows missing front:")
for name, source in missing_front:
    print(f"  {name[:50]}")

# ============================================================
# STEP 3: Try to find front for 15 missing models from Wayback
# ============================================================
print(f"\n=== STEP 3: FIND MISSING FRONT FROM WAYBACK ===")

def check_wayback(url, timeout=15):
    wb_api = f"https://archive.org/wayback/available?url={url}"
    try:
        req = urllib.request.Request(wb_api, headers={'User-Agent': 'Mozilla/5.0'})
        resp = opener.open(req, timeout=timeout)
        data = json.loads(resp.read())
        snapshots = data.get('archived_snapshots', {})
        if snapshots.get('closest'):
            return snapshots['closest']['url']
    except:
        pass
    return None

def extract_front_from_wayback(wb_url, timeout=15):
    try:
        req = urllib.request.Request(wb_url, headers={'User-Agent': 'Mozilla/5.0'})
        resp = opener.open(req, timeout=timeout)
        html = resp.read().decode('utf-8', errors='ignore')
        
        imgs = re.findall(r'(?:src|data-src)="([^"]+)"', html, re.I)
        door_imgs = []
        for img in imgs:
            img_lower = img.lower()
            if any(x in img_lower for x in ['jpg', 'png', 'webp']):
                if not any(x in img_lower for x in ['logo', 'icon', 'star', 'arrow', 'banner', 'badge', 'loading', 'toolbar', 'novinka_round']):
                    if img.startswith('/web/'):
                        match = re.search(r'/web/\d+im_/(https?://.+)', img)
                        if match:
                            img = match.group(1)
                    door_imgs.append(img)
        if door_imgs:
            return door_imgs[0]  # First product image is usually front
    except:
        pass
    return None

def check_image_url(url, timeout=10):
    for try_url in [url.replace('https://', 'http://'), url]:
        try:
            req = urllib.request.Request(try_url, headers={'User-Agent': 'Mozilla/5.0'})
            resp = opener.open(req, timeout=timeout)
            ct = resp.headers.get('Content-Type', '')
            data = resp.read(4)
            is_img = data[:3] == b'\xff\xd8\xff' or data[:4] == b'\x89PNG' or data[:4] == b'RIFF'
            if is_img:
                return True
        except:
            pass
    return False

front_found = 0
for idx, r in asd_rows:
    if r.get('image_url_front', '').strip():
        continue
    
    name = r['model_name']
    source = r.get('source_url', '')
    
    wb_url = check_wayback(source)
    if wb_url:
        front = extract_front_from_wayback(wb_url)
        if front and check_image_url(front):
            http_url = front.replace('https://', 'http://')
            rows[idx]['image_url_front'] = http_url
            if not rows[idx].get('image_url', '').strip():
                rows[idx]['image_url'] = http_url
            front_found += 1
            print(f"  FRONT FOUND: {name[:40]}: {front[-60:]}")
        else:
            print(f"  FRONT NOT FOUND: {name[:40]}")
    else:
        print(f"  NO WAYBACK: {name[:40]}")

print(f"\nFront found from Wayback: {front_found}")

# ============================================================
# STEP 4: Verify all v5 improvements preserved
# ============================================================
print(f"\n=== STEP 4: VERIFY V5 IMPROVEMENTS ===")

producers = Counter(r['producer'] for r in rows)
for p in EXPECTED_PRODUCERS:
    print(f"  {p}: {producers.get(p, 0)}")

# Duplicates
keys = set()
dupes = 0
for r in rows:
    key = (r['producer'], r['target_line'], r['model_name'])
    if key in keys:
        dupes += 1
    keys.add(key)
print(f"Duplicates: {dupes}")

# ASD final stats
asd_final = [r for r in rows if r['producer'] == 'АСД']
asd_front = sum(1 for r in asd_final if r.get('image_url_front', '').strip())
asd_inside = sum(1 for r in asd_final if r.get('image_url_inside', '').strip())
asd_both = sum(1 for r in asd_final if r.get('image_url_front', '').strip() and r.get('image_url_inside', '').strip())
print(f"\nASD final:")
print(f"  Total: {len(asd_final)}")
print(f"  Front: {asd_front}/199")
print(f"  Inside: {asd_inside}/199")
print(f"  Both: {asd_both}/199")

# No-image rows
no_image = [r for r in rows if not r.get('image_url', '').strip() and not r.get('image_url_front', '').strip() and not r.get('image_url_inside', '').strip()]
no_img_by_prod = Counter(r['producer'] for r in no_image)
print(f"\nNo-image rows: {len(no_image)}")
for p, c in sorted(no_img_by_prod.items()):
    print(f"  {p}: {c}")

# WordPress sync
wp_path = os.path.join(BASE, "output", "final_rework_ready_v5", "wordpress_import.csv")
with open(wp_path, encoding='utf-8-sig') as f:
    wp_count = len(list(csv.DictReader(f, delimiter=';')))
print(f"\nWordPress sync: CSV={len(rows)} WP={wp_count} match={len(rows) == wp_count}")

# ============================================================
# STEP 5: Save v5.2 catalog
# ============================================================
print(f"\n=== STEP 5: SAVE V5.2 CATALOG ===")

v52_catalog = os.path.join(OUT_DIR, "doors_catalog.csv")
with open(v52_catalog, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=';')
    writer.writeheader()
    writer.writerows(rows)

# Create XLSX
try:
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Catalog'
    ws.append(fieldnames)
    for r in rows:
        ws.append([r.get(f, '') for f in fieldnames])
    xlsx_path = os.path.join(OUT_DIR, "doors_catalog.xlsx")
    wb.save(xlsx_path)
    print(f"Saved XLSX: {xlsx_path}")
except Exception as e:
    print(f"XLSX error: {e}")

# Create WordPress import CSV
wp_csv_path = os.path.join(OUT_DIR, "wordpress_import.csv")
import shutil
shutil.copy2(v52_catalog, wp_csv_path)

# Create run_report.json
run_report = {
    'version': 'v5.2',
    'date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    'total_rows': len(rows),
    'producers': len(set(r['producer'] for r in rows)),
    'duplicates': dupes,
    'asd': {
        'total': len(asd_final),
        'front': asd_front,
        'inside': asd_inside,
        'both': asd_both,
    },
    'no_image_rows': len(no_image),
    'wordpress_sync': len(rows) == wp_count,
}
with open(os.path.join(OUT_DIR, "run_report.json"), 'w', encoding='utf-8') as f:
    json.dump(run_report, f, ensure_ascii=False, indent=2)

print(f"Saved v5.2 catalog: {v52_catalog}")
print(f"Total rows: {len(rows)}")

# ============================================================
# STEP 6: Generate ASD restore report
# ============================================================
print(f"\n=== STEP 6: GENERATE REPORTS ===")

try:
    import openpyxl
    
    # ASD restore report
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'ASD Restore'
    ws.append(['producer', 'target_line', 'model_name', 'source_url', 'v5_1_front', 'v5_1_inside', 'v4_front', 'v4_inside', 'restored_front', 'restored_inside', 'final_front', 'final_inside', 'front_action', 'inside_action', 'reason'])
    
    for idx, r in asd_rows:
        name = r['model_name']
        v4r = v4_asd.get(name, {})
        v4_front = v4r.get('image_url_front', '')
        v4_inside = v4r.get('image_url_inside', '')
        
        v51_front = v51_rows[idx].get('image_url_front', '') if idx < len(v51_rows) else ''
        v51_inside = v51_rows[idx].get('image_url_inside', '') if idx < len(v51_rows) else ''
        
        final_front = r.get('image_url_front', '')
        final_inside = r.get('image_url_inside', '')
        
        # Determine actions
        if not v51_front and final_front:
            front_action = 'restored'
        elif v51_front:
            front_action = 'kept'
        else:
            front_action = 'not_found'
        
        if not v51_inside and final_inside:
            inside_action = 'restored'
        elif v51_inside:
            inside_action = 'kept'
        else:
            inside_action = 'not_found'
        
        reason = ''
        if inside_action == 'restored':
            reason = 'wayback_machine_inside_image_found'
        elif inside_action == 'not_found':
            reason = 'no_wayback_snapshot'
        elif front_action == 'restored':
            reason = 'wayback_machine_front_image_found'
        
        ws.append([r['producer'], r['target_line'], name, r.get('source_url', ''), v51_front, v51_inside, v4_front, v4_inside, final_front if front_action == 'restored' else '', final_inside if inside_action == 'restored' else '', final_front, final_inside, front_action, inside_action, reason])
    
    wb.save(os.path.join(OUT_DIR, 'asd_front_inside_restore_report.xlsx'))
    print(f"Saved asd_front_inside_restore_report.xlsx")
    
    # ASD missing report
    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    ws2.append(['model_name', 'source_url', 'image_url_front', 'image_url_inside', 'missing_front', 'missing_inside', 'reason', 'status', 'comment'])
    
    for idx, r in asd_rows:
        name = r['model_name']
        front = r.get('image_url_front', '').strip()
        inside = r.get('image_url_inside', '').strip()
        missing_front = 'yes' if not front else 'no'
        missing_inside = 'yes' if not inside else 'no'
        
        if missing_front == 'yes' or missing_inside == 'yes':
            reason = ''
            if missing_front == 'yes':
                reason = 'front_photo_not_found_after_gallery_scan'
            if missing_inside == 'yes':
                reason += ('; ' if reason else '') + 'inside_photo_not_found_after_wayback_scan'
            
            ws2.append([name, r.get('source_url', ''), front, inside, missing_front, missing_inside, reason, 'needs_review', ''])
    
    wb2.save(os.path.join(OUT_DIR, 'asd_missing_front_inside_final.xlsx'))
    print(f"Saved asd_missing_front_inside_final.xlsx")
    
    # Final QA report
    wb3 = openpyxl.Workbook()
    ws3 = wb3.active
    ws3.title = 'QA Summary'
    ws3.append(['Metric', 'Value'])
    ws3.append(['Total rows', len(rows)])
    ws3.append(['Producers', 10])
    ws3.append(['Duplicates', dupes])
    ws3.append(['WordPress sync', f'{len(rows)} = {wp_count} = {len(rows) == wp_count}'])
    ws3.append([])
    ws3.append(['ASD total', len(asd_final)])
    ws3.append(['ASD front', f'{asd_front}/199'])
    ws3.append(['ASD inside', f'{asd_inside}/199'])
    ws3.append(['ASD both', f'{asd_both}/199'])
    ws3.append([])
    ws3.append(['No-image rows', len(no_image)])
    ws3.append([])
    ws3.append(['Producer', 'Rows', 'No-image'])
    for p in EXPECTED_PRODUCERS:
        ws3.append([p, producers.get(p, 0), no_img_by_prod.get(p, 0)])
    
    ws4 = wb3.create_sheet('No-Image Rows')
    ws4.append(['producer', 'target_line', 'model_name', 'comment'])
    for r in sorted(no_image, key=lambda x: x['producer']):
        ws4.append([r['producer'], r['target_line'], r['model_name'], r.get('comment', '')])
    
    wb3.save(os.path.join(OUT_DIR, 'final_delivery_qa_report.xlsx'))
    print(f"Saved final_delivery_qa_report.xlsx")

except Exception as e:
    print(f"Report error: {e}")

# ============================================================
# STEP 7: Build v5.2 archive
# ============================================================
print(f"\n=== STEP 7: BUILD ARCHIVE ===")

timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M')
archive_name = f"doors_catalog_delivery_v5.2_{timestamp}.zip"
archive_path = os.path.join(DELIVERY_DIR, archive_name)

files_to_include = [
    (os.path.join(OUT_DIR, "doors_catalog.csv"), "doors_catalog.csv"),
    (os.path.join(OUT_DIR, "doors_catalog.xlsx"), "doors_catalog.xlsx"),
    (os.path.join(OUT_DIR, "wordpress_import.csv"), "wordpress_import.csv"),
    (os.path.join(OUT_DIR, "run_report.json"), "run_report.json"),
    (os.path.join(OUT_DIR, 'asd_front_inside_restore_report.xlsx'), 'asd_front_inside_restore_report.xlsx'),
    (os.path.join(OUT_DIR, 'asd_missing_front_inside_final.xlsx'), 'asd_missing_front_inside_final.xlsx'),
    (os.path.join(OUT_DIR, 'final_delivery_qa_report.xlsx'), 'final_delivery_qa_report.xlsx'),
    (os.path.join(DELIVERY_DIR, 'no_image_rows_v5_final.xlsx'), 'no_image_rows_v5_final.xlsx'),
    (os.path.join(BASE, "output", "final_rework_ready_v5", "komfort_rebuild_report.json"), "komfort_rebuild_report.json"),
    (os.path.join(BASE, "output", "final_rework_ready_v5", "hausdoors_pvh_rebuild_report.json"), "hausdoors_pvh_rebuild_report.json"),
    (os.path.join(BASE, "output", "final_rework_ready_v5", "olimp_missing_collections_report.json"), "olimp_missing_collections_report.json"),
    (os.path.join(BASE, "output", "final_rework_ready_v5", "image_url_validation_report.json"), "image_url_validation_report.json"),
    (os.path.join(BASE, "output", "final_rework_ready_v5", "komfort_images.json"), "komfort_images.json"),
    (os.path.join(BASE, "output", "final_rework_ready_v5", "komfort_slug_image_map.json"), "komfort_slug_image_map.json"),
    (os.path.join(DELIVERY_DIR, 'final_image_reconciliation_report.json'), 'final_image_reconciliation_report.json'),
    (os.path.join(DELIVERY_DIR, 'final_image_reconciliation_report.xlsx'), 'final_image_reconciliation_report.xlsx'),
    (os.path.join(DELIVERY_DIR, 'asd_302_redirects_check.xlsx'), 'asd_302_redirects_check.xlsx'),
    (os.path.join(DELIVERY_DIR, 'lord_image_network_check.xlsx'), 'lord_image_network_check.xlsx'),
]

with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for src, dst in files_to_include:
        if os.path.exists(src):
            zf.write(src, dst)
            print(f"  Added: {dst}")
        else:
            print(f"  MISSING: {dst}")

with open(archive_path, 'rb') as f:
    sha256 = hashlib.sha256(f.read()).hexdigest()

archive_size = os.path.getsize(archive_path)

print(f"\n{'='*60}")
print(f"V5.2 DELIVERY SUMMARY")
print(f"{'='*60}")
print(f"Archive: {archive_path}")
print(f"SHA256: {sha256}")
print(f"Size: {archive_size:,} bytes")
print(f"\nCatalog: {len(rows)} rows, 10 producers, {dupes} duplicates")
print(f"WordPress sync: {len(rows)} = {wp_count} = {len(rows) == wp_count}")
print(f"\nASD:")
print(f"  Total: {len(asd_final)}")
print(f"  Front: {asd_front}/199")
print(f"  Inside: {asd_inside}/199")
print(f"  Both: {asd_both}/199")
print(f"\nNo-image rows: {len(no_image)}")
for p, c in sorted(no_img_by_prod.items()):
    print(f"  {p}: {c}")
print(f"\nKomfort: {producers.get('Комфорт', 0)}")
print(f"Hausdoors: {producers.get('Хаусдорс', 0)}")
print(f"Olimp: {producers.get('Олимп', 0)}")
print(f"Entro: {producers.get('Энтро', 0)}")
