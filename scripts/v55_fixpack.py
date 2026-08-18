"""V5.5 Fix Pack: restore false positives (Techno, variant models, osteklennaya)"""
import csv, os, json, shutil, zipfile, hashlib
from collections import Counter
from datetime import datetime

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
V54_DIR = os.path.join(BASE, "output", "delivery_v5_4")
V55_DIR = os.path.join(BASE, "output", "delivery_v5_5")
DELIVERY = os.path.join(BASE, "output", "delivery")

# Load v5.4 catalog (reference for restoration)
with open(os.path.join(V54_DIR, "doors_catalog.csv"), encoding='utf-8-sig') as f:
    rows_54 = list(csv.DictReader(f, delimiter=';'))
fn = list(rows_54[0].keys())

# Load v5.5 catalog
with open(os.path.join(V55_DIR, "doors_catalog.csv"), encoding='utf-8-sig') as f:
    rows_55 = list(csv.DictReader(f, delimiter=';'))

LORD = chr(0x041B)+chr(0x043E)+chr(0x0440)+chr(0x0434)
ASD = chr(0x0410)+chr(0x0421)+chr(0x0414)

# Build v5.4 lookup by model_name
rows54_by_name = {}
for r in rows_54:
    rows54_by_name[r['model_name']] = r

changes = {
    'techno_restored': 0,
    'variant_restored': 0,
    'osteklennaya_restored': 0,
    'other_restored': 0,
    'fixes_detail': [],
}

for r in rows_55:
    name = r['model_name']
    
    # 1. Fix Techno models - restore from v5.4
    if name.lower().startswith('techno'):
        r54 = rows54_by_name.get(name)
        if r54 and (r54.get('image_url_front', '').strip() != r.get('image_url_front', '').strip()):
            old_url = r54['image_url_front']
            r['image_url_front'] = old_url
            r['image_url'] = old_url
            # Fix comment
            comment = r.get('comment', '')
            r['comment'] = comment.replace('lord_fragment_image_cleared_no_full_door_replacement', 'techno_false_positive_restored')
            if 'techno_false_positive_restored' not in r['comment']:
                r['comment'] = r['comment'] + '; techno_false_positive_restored'
            changes['techno_restored'] += 1
            changes['fixes_detail'].append(f"TECHNO {name[:40]}: front restored from v5.4")
    
    # 2. Fix variant models (interior slides) - these are the only available photos for "variant N" models
    elif 'variant' in name.lower() and name.count('-') <= 2:
        r54 = rows54_by_name.get(name)
        if r54 and not r.get('image_url_front', '').strip() and r54.get('image_url_front', '').strip():
            old_url = r54['image_url_front']
            r['image_url_front'] = old_url
            r['image_url'] = old_url
            comment = r.get('comment', '')
            r['comment'] = comment.replace('lord_fragment_image_cleared_no_full_door_replacement', 'interior_slide_preserved_as_best_available')
            if 'interior_slide_preserved' not in r['comment']:
                r['comment'] = r['comment'] + '; interior_slide_preserved_as_best_available'
            changes['variant_restored'] += 1
            changes['fixes_detail'].append(f"VARIANT {name[:40]}: slide image restored from v5.4 (best available)")

    # 3. Fix osteklennaya models - restore front photo if currently empty
    elif 'osteklennaya' in name.lower():
        r54 = rows54_by_name.get(name)
        if r54 and not r.get('image_url_front', '').strip() and r54.get('image_url_front', '').strip():
            old_url = r54['image_url_front']
            r['image_url_front'] = old_url
            r['image_url'] = old_url
            comment = r.get('comment', '')
            r['comment'] = comment.replace('lord_fragment_image_cleared_no_full_door_replacement', 'osteklennaya_restored_no_base_model')
            if 'osteklennaya_restored' not in r['comment']:
                r['comment'] = r['comment'] + '; osteklennaya_restored_no_base_model'
            changes['osteklennaya_restored'] += 1
            changes['fixes_detail'].append(f"OSTEKLENNAYA {name[:40]}: restored from v5.4 (no base model found)")

    # 4. General fix: any Lord model that lost all photos AND had no replacement
    elif r['producer'] == LORD and not r.get('image_url_front', '').strip():
        r54 = rows54_by_name.get(name)
        if r54 and r54.get('image_url_front', '').strip():
            # Check if there's a steklo -> base replacement already applied
            comment = r.get('comment', '')
            if 'lord_steklo_variant_replaced_by_base_photo' not in comment:
                old_url = r54['image_url_front']
                r['image_url_front'] = old_url
                r['image_url'] = old_url
                r['comment'] = comment.replace('lord_fragment_image_cleared_no_full_door_replacement', 'restored_from_v54_no_alternative')
                if 'restored_from_v54' not in r['comment']:
                    r['comment'] = r['comment'] + '; restored_from_v54_no_alternative'
                changes['other_restored'] += 1
                changes['fixes_detail'].append(f"OTHER {name[:40]}: restored from v5.4 (no better alternative)")

# Check handle/lock images for ASD - from diagnose, there are 2 remaining
for r in rows_55:
    name = r['model_name']
    comment = r.get('comment', '')
    # These are Конструктор models with handle images - keep them as ASD steel doors
    # The handle is part of the steel door product photo

# Stats after fix
prods = Counter(r['producer'] for r in rows_55)
keys_set = set()
dupes = 0
for r in rows_55:
    k = (r['producer'], r['target_line'], r['model_name'])
    if k in keys_set:
        dupes += 1
    keys_set.add(k)

asd_rows = [r for r in rows_55 if r['producer'] == ASD]
lord_rows = [r for r in rows_55 if r['producer'] == LORD]
asd_f = sum(1 for r in asd_rows if r.get('image_url_front', '').strip())
asd_i = sum(1 for r in asd_rows if r.get('image_url_inside', '').strip())
lord_f = sum(1 for r in lord_rows if r.get('image_url_front', '').strip())

no_img = sum(1 for r in rows_55 if not r.get('image_url_front', '').strip() and not r.get('image_url_inside', '').strip())
wp_ok = all(
    r.get('image_url_front', '').strip() and r['image_url'] == r['image_url_front']
    for r in rows_55 if r.get('image_url_front', '').strip()
)

# write fixed catalog
out_csv = os.path.join(V55_DIR, "doors_catalog.csv")
with open(out_csv, 'w', encoding='utf-8-sig', newline='') as f:
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writeheader()
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writerows(rows_55)

# re-write XLSX
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
hdr_font = Font(bold=True, color="FFFFFF", size=10)
hdr_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
def write_xlsx(path, headers, data, sheet_name="Sheet1"):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font = hdr_font; c.fill = hdr_fill; c.alignment = hdr_align; c.border = thin_border
    for ri, row in enumerate(data, 2):
        for ci, h in enumerate(headers, 1):
            val = row.get(h, '')
            if isinstance(val, list): val = '; '.join(val)
            c = ws.cell(row=ri, column=ci, value=val)
            c.border = thin_border
            c.alignment = Alignment(vertical="top", wrap_text=True)
    wb.save(path)

write_xlsx(os.path.join(V55_DIR, "doors_catalog.xlsx"), fn, rows_55, "Catalog")
write_xlsx(os.path.join(V55_DIR, "wordpress_import.csv"), fn, rows_55, "WordPress")

# Also update wordpress CSV
wp_csv = os.path.join(V55_DIR, "wordpress_import.csv")
with open(wp_csv, 'w', encoding='utf-8-sig', newline='') as f:
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writeheader()
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writerows(rows_55)

# Update metrics in validation report
val_json_path = os.path.join(V55_DIR, "final_photo_validation_report.json")
if os.path.exists(val_json_path):
    with open(val_json_path, encoding='utf-8') as f:
        val_report = json.load(f)
    val_report['restorations_applied'] = changes
    val_report['lord_empty_front_after_fix'] = sum(1 for r in lord_rows if not r.get('image_url_front', '').strip())
    with open(val_json_path, 'w', encoding='utf-8') as f:
        json.dump(val_report, f, ensure_ascii=False, indent=2)

print("=== FIX PACK RESULTS ===")
print(f"Techno restored: {changes['techno_restored']}")
print(f"Variant restored: {changes['variant_restored']}")
print(f"Osteklennaya restored: {changes['osteklennaya_restored']}")
print(f"Other restored: {changes['other_restored']}")
print(f"Total fixes: {sum(1 for c in changes['fixes_detail'])}")
for d in changes['fixes_detail'][:30]:
    print(f"  {d}")

# rebuild archive
print("\n=== Rebuilding archive ===")
CLIENT_FILES = [
    "doors_catalog.csv",
    "doors_catalog.xlsx",
    "wordpress_import.csv",
    "run_report.json",
    "asd_inside_photo_reqa_report.xlsx",
    "asd_rows_1638_1644_check.xlsx",
    "lord_fragment_photo_qa_report.xlsx",
    "lord_fragment_photo_review.html",
    "final_photo_validation_report.xlsx",
    "final_photo_validation_report.json",
    "final_delivery_qa_report.xlsx",
]

TS = datetime.now().strftime("%Y-%m-%d_%H-%M")
archive_name = f"doors_catalog_delivery_v5_5_PHOTO_QA_PATCH_{TS}.zip"
archive_path = os.path.join(DELIVERY, archive_name)

with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for fname in CLIENT_FILES:
        p = os.path.join(V55_DIR, fname)
        if os.path.exists(p):
            zf.write(p, fname)

sha256 = hashlib.sha256()
with open(archive_path, 'rb') as f:
    while True:
        chunk = f.read(8192)
        if not chunk: break
        sha256.update(chunk)

print(f"\n{'='*60}")
print(f"V5.5 FIXED DELIVERY SUMMARY")
print(f"{'='*60}")
print(f"  archive_path: {archive_path}")
print(f"  sha256: {sha256.hexdigest()}")
print(f"  total_rows: {len(rows_55)}")
print(f"  producers: {len(prods)}")
print(f"  duplicates: {dupes}")
print(f"  ASD rows: {len(asd_rows)}")
print(f"  ASD front filled: {asd_f}/199")
print(f"  ASD inside filled: {asd_i}/199")
print(f"  Lord front filled: {lord_f}/{len(lord_rows)}")
print(f"  Lord empty front: {sum(1 for r in lord_rows if not r.get('image_url_front','').strip())}")
print(f"  No-image rows: {no_img}")
print(f"  WordPress sync: {'OK' if wp_ok else 'FAIL'}")
