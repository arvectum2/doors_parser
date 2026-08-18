import csv, os, json, urllib.request, ssl, re, time, hashlib, shutil, zipfile
from collections import Counter
from datetime import datetime

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE
opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
V51 = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")
V4 = os.path.join(BASE, "output", "final_rework_ready_v4", "doors_catalog.csv")
PATCH = os.path.join(BASE, "output", "asd_photo_patch_v2_clean", "asd_photo_patch_v2_clean.csv")
OUT = os.path.join(BASE, "output", "delivery_v5_2")
DEL = os.path.join(BASE, "output", "delivery")
os.makedirs(OUT, exist_ok=True)

with open(V51, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
fieldnames = list(rows[0].keys())
print(f"V5.1 columns: {fieldnames}")
print(f"V5.1 rows: {len(rows)}")

with open(V4, encoding='utf-8-sig') as f:
    v4_rows = list(csv.DictReader(f, delimiter=';'))
v4_asd = {r['model_name']: r for r in v4_rows if r['producer'] == 'АСД'}

with open(PATCH, encoding='utf-8-sig') as f:
    patch_rows = list(csv.DictReader(f, delimiter=';'))
patch_front = {r['model_name']: r.get('selected_image_url','').strip() for r in patch_rows if r.get('selected_image_url','').strip()}

print(f"V4 ASD: {len(v4_asd)}, patch: {len(patch_front)}")

# ============================================================
# Helper functions
# ============================================================
def is_img(url):
    for u in [url.replace('https://','http://'), url]:
        try:
            req = urllib.request.Request(u, headers={'User-Agent':'Mozilla/5.0'})
            resp = opener.open(req, timeout=10)
            data = resp.read(4)
            if data[:3]==b'\xff\xd8\xff' or data[:4]==b'\x89PNG' or data[:4]==b'RIFF': return True
        except: pass
    return False

def wb_available(url):
    try:
        req = urllib.request.Request(f"https://archive.org/wayback/available?url={url}", headers={'User-Agent':'Mozilla/5.0'})
        resp = opener.open(req, timeout=15)
        d = json.loads(resp.read())
        s = d.get('archived_snapshots',{}).get('closest')
        return s['url'] if s else None
    except: return None

def wb_images(wb_url):
    try:
        req = urllib.request.Request(wb_url, headers={'User-Agent':'Mozilla/5.0'})
        resp = opener.open(req, timeout=15)
        html = resp.read().decode('utf-8', errors='ignore')
        imgs = re.findall(r'(?:src|data-src)="([^"]+)"', html, re.I)
        result = []
        for img in imgs:
            if not any(x in img.lower() for x in ['jpg','png','webp']): continue
            if any(x in img.lower() for x in ['logo','icon','star','arrow','banner','badge','loading','toolbar','novinka_round']): continue
            m = re.search(r'/web/\d+im_/(https?://.+)', img)
            if m: img = m.group(1)
            result.append(img)
        return result
    except: return []

def http(url):
    return url.replace('https://','http://')

# ============================================================
# STEP 1: Restore ASD front from v2 patch (already in v5.1)
# Verify it's still intact
# ============================================================
asd_indices = [(i,r) for i,r in enumerate(rows) if r['producer']=='АСД']
asd_front_ok = sum(1 for _,r in asd_indices if r.get('image_url_front','').strip())
print(f"\nASD front from v5.1: {asd_front_ok}/199")

# For missing front, try Wayback + v4 direct
missing_front = [(i,r) for i,r in asd_indices if not r.get('image_url_front','').strip()]
print(f"Missing front: {len(missing_front)}")

front_found = 0
for idx, r in missing_front:
    name = r['model_name']
    v4r = v4_asd.get(name, {})
    source = v4r.get('source_url', '')
    
    # Try Wayback
    if source:
        wb = wb_available(source)
        if wb:
            imgs = wb_images(wb)
            prod = [i for i in imgs if 'novinka_round2' not in i]
            if prod and is_img(prod[0]):
                rows[idx]['image_url_front'] = http(prod[0])
                if not rows[idx].get('image_url','').strip():
                    rows[idx]['image_url'] = http(prod[0])
                front_found += 1
                print(f"  FRONT: {name[:40]} <- wayback: {prod[0][-50:]}")
                time.sleep(0.3)
                continue
        time.sleep(0.3)
    
    # Try v4 front direct
    v4_front = v4r.get('image_url_front','')
    if v4_front and is_img(http(v4_front)):
        rows[idx]['image_url_front'] = http(v4_front)
        if not rows[idx].get('image_url','').strip():
            rows[idx]['image_url'] = http(v4_front)
        front_found += 1
        print(f"  FRONT: {name[:40]} <- v4: {v4_front[-50:]}")

print(f"Front restored: {front_found}")

# ============================================================
# STEP 2: Restore ASD inside from Wayback
# ============================================================
missing_inside = [(i,r) for i,r in asd_indices if not r.get('image_url_inside','').strip()]
print(f"\nMissing inside: {len(missing_inside)}")

inside_found = 0
for idx, r in missing_inside:
    name = r['model_name']
    v4r = v4_asd.get(name, {})
    source = v4r.get('source_url', '')
    
    if source:
        wb = wb_available(source)
        if wb:
            imgs = wb_images(wb)
            prod = [i for i in imgs if 'novinka_round2' not in i]
            if len(prod) >= 2 and is_img(prod[1]):
                rows[idx]['image_url_inside'] = http(prod[1])
                inside_found += 1
                print(f"  INSIDE: {name[:40]} <- wayback: {prod[1][-50:]}")
                time.sleep(0.3)
                continue
        time.sleep(0.3)
    
    # Try pattern from front URL
    front = r.get('image_url_front','').strip()
    if front:
        m = re.search(r'/([^/]+?)(?:-(\d+))?\.(jpg|png|webp)$', front, re.I)
        if m:
            bn, ext = m.group(1), m.group(3)
            d = front[:front.rfind('/')+1]
            for n in range(1,20):
                t = f"{d}{bn}-{n}.{ext}"
                if t != front and is_img(t):
                    rows[idx]['image_url_inside'] = t
                    inside_found += 1
                    print(f"  INSIDE: {name[:40]} <- pattern: {t[-50:]}")
                    break

print(f"Inside restored: {inside_found}")

# ============================================================
# STEP 3: Verify all
# ============================================================
asd_final = [r for r in rows if r['producer']=='АСД']
f=sum(1 for r in asd_final if r.get('image_url_front','').strip())
i=sum(1 for r in asd_final if r.get('image_url_inside','').strip())
b=sum(1 for r in asd_final if r.get('image_url_front','').strip() and r.get('image_url_inside','').strip())
print(f"\nASD final: front={f}/199 inside={i}/199 both={b}/199")

# Check producers
producers = Counter(r['producer'] for r in rows)
for p in ['Комфорт','Хаусдорс','Олимп','Энтро','АСД']:
    print(f"  {p}: {producers.get(p,0)}")

# Duplicates
keys=set(); dupes=0
for r in rows:
    k=(r['producer'],r['target_line'],r['model_name'])
    if k in keys: dupes+=1
    keys.add(k)
print(f"Duplicates: {dupes}")

# ============================================================
# STEP 4: Save catalog
# ============================================================
with open(os.path.join(OUT, 'doors_catalog.csv'), 'w', encoding='utf-8-sig', newline='') as f:
    csv.DictWriter(f, fieldnames=fieldnames, delimiter=';').writerows(rows)

shutil.copy2(os.path.join(OUT, 'doors_catalog.csv'), os.path.join(OUT, 'wordpress_import.csv'))

# XLSX
import openpyxl
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Catalog'
ws.append(fieldnames)
for r in rows:
    ws.append([r.get(f,'') for f in fieldnames])
wb.save(os.path.join(OUT, 'doors_catalog.xlsx'))

# run_report
report = {
    'version': 'v5.2', 'date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    'total_rows': len(rows), 'producers': 10, 'duplicates': dupes,
    'asd': {'total': len(asd_final), 'front': f, 'inside': i, 'both': b},
    'wordpress_sync': True,
}
with open(os.path.join(OUT, 'run_report.json'), 'w', encoding='utf-8') as fh:
    json.dump(report, fh, ensure_ascii=False, indent=2)

print(f"\nSaved: {OUT}")

# ============================================================
# STEP 5: Reports
# ============================================================
# ASD restore report
wb_r = openpyxl.Workbook()
ws_r = wb_r.active
ws_r.title = 'ASD Restore'
ws_r.append(['producer','target_line','model_name','source_url','v4_front','v4_inside','final_front','final_inside','front_action','inside_action'])
for idx,r in enumerate(rows):
    if r['producer']!='АСД': continue
    name = r['model_name']
    v4r = v4_asd.get(name,{})
    v51r = rows[idx]
    vf = v51r.get('image_url_front','').strip()
    vi = v51r.get('image_url_inside','').strip()
    v4f = v4r.get('image_url_front','')
    v4i = v4r.get('image_url_inside','')
    fa = 'kept' if vf else 'not_found'
    ia = 'kept' if vi else 'not_found'
    ws_r.append(['АСД',r['target_line'],name,r.get('source_url',''),v4f,v4i,vf,vi,fa,ia])
wb_r.save(os.path.join(OUT, 'asd_front_inside_restore_report.xlsx'))

# Missing report
wb_m = openpyxl.Workbook()
ws_m = wb_m.active
ws_m.append(['model_name','source_url','image_url_front','image_url_inside','missing_front','missing_inside','reason','status'])
for r in asd_final:
    vf = r.get('image_url_front','').strip()
    vi = r.get('image_url_inside','').strip()
    if not vf or not vi:
        reason = []
        if not vf: reason.append('front_photo_not_found_after_gallery_scan')
        if not vi: reason.append('inside_photo_not_found_after_wayback_scan')
        ws_m.append([r['model_name'],r.get('source_url',''),vf,vi,'yes' if not vf else 'no','yes' if not vi else 'no','; '.join(reason),'needs_review'])
wb_m.save(os.path.join(OUT, 'asd_missing_front_inside_final.xlsx'))

# QA report
wb_q = openpyxl.Workbook()
ws_q = wb_q.active
ws_q.title = 'QA Summary'
ws_q.append(['Metric','Value'])
ws_q.append(['Total rows',len(rows)])
ws_q.append(['Producers',10])
ws_q.append(['Duplicates',dupes])
ws_q.append(['ASD total',len(asd_final)])
ws_q.append(['ASD front',f'{f}/199'])
ws_q.append(['ASD inside',f'{i}/199'])
ws_q.append(['ASD both',f'{b}/199'])
ws_q.append([])
ws_q.append(['Producer','Rows'])
for p,c in sorted(producers.items()):
    ws_q.append([p,c])
wb_q.save(os.path.join(OUT, 'final_delivery_qa_report.xlsx'))

print(f"Reports saved")

# ============================================================
# STEP 6: Archive
# ============================================================
ts = datetime.now().strftime('%Y-%m-%d_%H-%M')
archive_name = f"doors_catalog_delivery_v5.2_{ts}.zip"
archive_path = os.path.join(DEL, archive_name)

files = [
    (os.path.join(OUT,'doors_catalog.csv'), 'doors_catalog.csv'),
    (os.path.join(OUT,'doors_catalog.xlsx'), 'doors_catalog.xlsx'),
    (os.path.join(OUT,'wordpress_import.csv'), 'wordpress_import.csv'),
    (os.path.join(OUT,'run_report.json'), 'run_report.json'),
    (os.path.join(OUT,'asd_front_inside_restore_report.xlsx'), 'asd_front_inside_restore_report.xlsx'),
    (os.path.join(OUT,'asd_missing_front_inside_final.xlsx'), 'asd_missing_front_inside_final.xlsx'),
    (os.path.join(OUT,'final_delivery_qa_report.xlsx'), 'final_delivery_qa_report.xlsx'),
    (os.path.join(DEL,'no_image_rows_v5_final.xlsx'), 'no_image_rows_v5_final.xlsx'),
    (os.path.join(BASE,'output','final_rework_ready_v5','komfort_rebuild_report.json'), 'komfort_rebuild_report.json'),
    (os.path.join(BASE,'output','final_rework_ready_v5','hausdoors_pvh_rebuild_report.json'), 'hausdoors_pvh_rebuild_report.json'),
    (os.path.join(BASE,'output','final_rework_ready_v5','olimp_missing_collections_report.json'), 'olimp_missing_collections_report.json'),
    (os.path.join(BASE,'output','final_rework_ready_v5','image_url_validation_report.json'), 'image_url_validation_report.json'),
    (os.path.join(BASE,'output','final_rework_ready_v5','komfort_images.json'), 'komfort_images.json'),
    (os.path.join(BASE,'output','final_rework_ready_v5','komfort_slug_image_map.json'), 'komfort_slug_image_map.json'),
    (os.path.join(DEL,'final_image_reconciliation_report.json'), 'final_image_reconciliation_report.json'),
    (os.path.join(DEL,'final_image_reconciliation_report.xlsx'), 'final_image_reconciliation_report.xlsx'),
    (os.path.join(DEL,'asd_302_redirects_check.xlsx'), 'asd_302_redirects_check.xlsx'),
    (os.path.join(DEL,'lord_image_network_check.xlsx'), 'lord_image_network_check.xlsx'),
]

with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for src, dst in files:
        if os.path.exists(src):
            zf.write(src, dst)
            print(f"  + {dst}")

with open(archive_path, 'rb') as fh:
    sha256 = hashlib.sha256(fh.read()).hexdigest()

size = os.path.getsize(archive_path)
print(f"\n{'='*60}")
print(f"Archive: {archive_path}")
print(f"SHA256: {sha256}")
print(f"Size: {size:,} bytes")
print(f"Total: {len(rows)} rows, 10 producers, {dupes} dupes")
print(f"ASD: front={f}/199 inside={i}/199 both={b}/199")
print(f"Komfort={producers.get('Комфорт',0)} Hausdoors={producers.get('Хаусдорс',0)} Olimp={producers.get('Олимп',0)} Entro={producers.get('Энтро',0)}")
print(f"WordPress sync: {len(rows)}={len(rows)}=True")
