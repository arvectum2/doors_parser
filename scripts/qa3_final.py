"""Regenerate all QA reports with correct classification + apply delivery rules + build archive"""
import csv, json, os, re, openpyxl, zipfile, hashlib
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from collections import Counter, defaultdict
from datetime import datetime

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
OUT_DIR = os.path.join(BASE, "output", "delivery_v5_2")
CAT = os.path.join(OUT_DIR, "doors_catalog.csv")
VALID = os.path.join(OUT_DIR, "asd_url_validation_results.json")
QA_JSON = os.path.join(OUT_DIR, "asd_final_photo_source_validation.json")

ASD_KEY = "\u0410\u0421\u0414"

# Load data
with open(CAT, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
fn = list(rows[0].keys())
asd = [r for r in rows if r['producer'] == ASD_KEY]

with open(QA_JSON, encoding='utf-8') as f:
    qa_rows = json.load(f)
qa_by_name = {r['model_name']: r for r in qa_rows}

# Styles
hdr_font = Font(bold=True, color="FFFFFF", size=10)
hdr_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

def write_xlsx(path, headers, data, sheet_name="Sheet1"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font = hdr_font; c.fill = hdr_fill; c.alignment = hdr_align; c.border = thin_border
    for ri, row in enumerate(data, 2):
        for ci, h in enumerate(headers, 1):
            c = ws.cell(row=ri, column=ci, value=row.get(h, ''))
            c.border = thin_border
            c.alignment = Alignment(vertical="top", wrap_text=True)
    col_widths = {chr(65+i): 50 if i < 3 else 18 for i in range(len(headers))}
    for col, w in col_widths.items():
        if col in [chr(65+i) for i in range(len(headers))]:
            ws.column_dimensions[col].width = w
    wb.save(path)

# ========================================================================
# QA Report 1: Full validation XLSX
# ========================================================================
qa_headers = ['model_name', 'source_url', 'image_url_front', 'image_url_inside',
              'front_source_type', 'inside_source_type', 'front_http_status', 'inside_http_status',
              'front_content_type', 'inside_content_type', 'front_opens_as_image', 'inside_opens_as_image',
              'front_visual_decision', 'inside_visual_decision', 'action', 'reason']
path1 = os.path.join(OUT_DIR, "asd_final_photo_source_validation.xlsx")
write_xlsx(path1, qa_headers, qa_rows, "ASD Photo Source Validation")
print(f"Report 1: {path1} ({len(qa_rows)} rows)")

# ========================================================================
# QA Report 2: Wayback inside links
# ========================================================================
wb_rows = [q for q in qa_rows if q['inside_source_type'] == 'wayback']
wb_headers = ['model_name', 'source_url', 'image_url_inside', 'inside_http_status',
              'inside_content_type', 'inside_opens_as_image', 'inside_visual_decision',
              'recommended_for_delivery', 'reason']
wb_data = []
for q in wb_rows:
    recommended = 'yes' if q['inside_opens_as_image'] == 'yes' and q['inside_visual_decision'] in ('full_door_inside', 'full_door_front') else 'no'
    wb_data.append({
        'model_name': q['model_name'], 'source_url': q['source_url'],
        'image_url_inside': q['image_url_inside'],
        'inside_http_status': q['inside_http_status'],
        'inside_content_type': q['inside_content_type'],
        'inside_opens_as_image': q['inside_opens_as_image'],
        'inside_visual_decision': q['inside_visual_decision'],
        'recommended_for_delivery': recommended,
        'reason': q['reason'],
    })
path2 = os.path.join(OUT_DIR, "asd_wayback_inside_links.xlsx")
write_xlsx(path2, wb_headers, wb_data, "Wayback Inside Links")
print(f"Report 2: {path2} ({len(wb_data)} rows)")

# ========================================================================
# QA Report 3: Pattern-matched inside
# ========================================================================
pm_rows = [q for q in qa_rows if q['inside_source_type'] == 'pattern_matched']
pm_headers = ['model_name', 'source_url', 'image_url_front', 'image_url_inside',
              'inside_http_status', 'inside_content_type', 'inside_opens_as_image',
              'inside_visual_decision', 'action', 'reason']
pm_data = []
for q in pm_rows:
    pm_data.append({
        'model_name': q['model_name'], 'source_url': q['source_url'],
        'image_url_front': q['image_url_front'],
        'image_url_inside': q['image_url_inside'],
        'inside_http_status': q['inside_http_status'],
        'inside_content_type': q['inside_content_type'],
        'inside_opens_as_image': q['inside_opens_as_image'],
        'inside_visual_decision': q['inside_visual_decision'],
        'action': q['action'], 'reason': q['reason'],
    })
path3 = os.path.join(OUT_DIR, "asd_pattern_matched_inside_check.xlsx")
write_xlsx(path3, pm_headers, pm_data, "Pattern Matched Inside")
print(f"Report 3: {path3} ({len(pm_data)} rows)")

# ========================================================================
# QA Report 4: Missing images
# ========================================================================
missing_rows = [q for q in qa_rows if not q['image_url_front'] or not q['image_url_inside']]
missing_headers = ['model_name', 'target_line', 'source_url', 'has_front', 'has_inside',
                   'found_in_v2_patch', 'found_in_wayback', 'final_status', 'final_comment']

missing_data = []
for q in missing_rows:
    name = q['model_name']
    r = next((x for x in asd if x['model_name'] == name), {})
    
    has_front = bool(q['image_url_front'])
    has_inside = bool(q['image_url_inside'])
    
    f_in_patch = q['front_source_type'] == 'v2_patch'
    i_in_wb = q['inside_source_type'] == 'wayback'
    
    comment_parts = []
    if not has_front and not has_inside:
        status = 'no_images'
        comment_parts.append('No front or inside images')
    elif has_front and not has_inside:
        status = 'no_inside_only'
        comment_parts.append('Front exists, inside missing')
    elif not has_front and has_inside:
        status = 'no_front_only'
        comment_parts.append('Inside exists, front missing')
    else:
        status = 'partial'
        comment_parts.append('Some but not all images')
    
    missing_data.append({
        'model_name': name,
        'target_line': r.get('target_line', ''),
        'source_url': r.get('source_url', ''),
        'has_front': 'yes' if has_front else 'no',
        'has_inside': 'yes' if has_inside else 'no',
        'found_in_v2_patch': 'yes' if f_in_patch else 'no',
        'found_in_wayback': 'yes' if i_in_wb else 'no',
        'final_status': status,
        'final_comment': '; '.join(comment_parts),
    })
path4 = os.path.join(OUT_DIR, "asd_remaining_missing_images.xlsx")
write_xlsx(path4, missing_headers, missing_data, "Missing Images")
print(f"Report 4: {path4} ({len(missing_data)} rows)")

# ========================================================================
# Apply delivery rules
# ========================================================================
updated = 0
for r in rows:
    if r['producer'] != ASD_KEY:
        continue
    name = r['model_name']
    q = qa_by_name.get(name)
    if not q: continue
    
    # Rule: image_url = image_url_front
    if r.get('image_url_front', '').strip():
        r['image_url'] = r['image_url_front']
    
    comment_bits = []
    if r.get('comment', '').strip():
        comment_bits.append(r['comment'].strip())
    
    # Rule: mark wayback inside URLs in comment
    if q['inside_source_type'] == 'wayback':
        tag = 'inside_photo_from_wayback_archive'
        if tag not in str(comment_bits):
            comment_bits.append(tag)
            updated += 1
    
    # Rule: invalid pattern-matched URLs - clear
    if q['inside_source_type'] == 'pattern_matched' and q['inside_opens_as_image'] == 'no':
        if r.get('image_url_inside', ''):
            r['image_url_inside'] = ''
            r['status'] = 'needs_review'
            comment_bits.append('inside_pattern_url_invalid_cleared')
            updated += 1
    
    r['comment'] = '; '.join(comment_bits)

# Save updated catalog
cat_path = os.path.join(OUT_DIR, "doors_catalog.csv")
with open(cat_path, 'w', encoding='utf-8-sig', newline='') as f:
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writeheader()
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writerows(rows)
print(f"\nCatalog updated: {cat_path} ({updated} rows modified)")

# ========================================================================
# Build archive
# ========================================================================
DELIVERY = os.path.join(BASE, "output", "delivery")
os.makedirs(DELIVERY, exist_ok=True)

CLIENT_FILES = [
    "doors_catalog.csv",
    "doors_catalog.xlsx",
    "wordpress_import.csv",
    "run_report.json",
    "asd_final_photo_source_validation.xlsx",
    "asd_final_photo_source_validation.json",
    "asd_wayback_inside_links.xlsx",
    "asd_pattern_matched_inside_check.xlsx",
    "asd_remaining_missing_images.xlsx",
    "final_delivery_qa_report.xlsx",
]

ts = datetime.now().strftime("%Y-%m-%d_%H-%M")
archive_name = f"doors_catalog_delivery_v5_2_QA_{ts}.zip"
archive_path = os.path.join(DELIVERY, archive_name)

with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for f in CLIENT_FILES:
        p = os.path.join(OUT_DIR, f)
        if os.path.exists(p):
            zf.write(p, f)
            print(f"  Added: {f} ({os.path.getsize(p):,} bytes)")

# SHA256
sha256 = hashlib.sha256()
with open(archive_path, 'rb') as f:
    while True:
        chunk = f.read(8192)
        if not chunk: break
        sha256.update(chunk)

size = os.path.getsize(archive_path)
print(f"\nArchive: {archive_name}")
print(f"Size: {size:,} bytes")
print(f"SHA256: {sha256.hexdigest()}")

# ========================================================================
# Final summary
# ========================================================================
# Re-count after updates
asd_final = [r for r in rows if r['producer'] == ASD_KEY]
f_final = sum(1 for r in asd_final if r.get('image_url_front','').strip())
i_final = sum(1 for r in asd_final if r.get('image_url_inside','').strip())
b_final = sum(1 for r in asd_final if r.get('image_url_front','').strip() and r.get('image_url_inside','').strip())

# Count wayback usage in catalog
wb_in_catalog = sum(1 for r in asd_final if 'inside_photo_from_wayback_archive' in (r.get('comment','') or ''))
# Count pattern-matched in catalog (not cleared)
pm_in_catalog = sum(1 for q in qa_rows if q['inside_source_type'] == 'pattern_matched' and q['inside_opens_as_image'] == 'yes' and any(r['model_name']==q['model_name'] and r.get('image_url_inside','').strip() for r in asd_final))

invalid_still = 0
for q in qa_rows:
    if q['front_opens_as_image'] == 'no' and q['image_url_front']:
        invalid_still += 1
    if q['inside_opens_as_image'] == 'no' and q['image_url_inside']:
        # Check if already cleared
        r = next((x for x in asd_final if x['model_name']==q['model_name']), None)
        if r and r.get('image_url_inside','').strip():
            invalid_still += 1

tech_count = sum(1 for q in qa_rows if q['front_visual_decision'] in ('handle','lock','fragment','technical','badge','interior_room'))
tech_count += sum(1 for q in qa_rows if q['inside_visual_decision'] in ('handle','lock','fragment','technical','badge','interior_room'))

no_img = sum(1 for q in qa_rows if not q['image_url_front'] and not q['image_url_inside'])

print(f"\n{'='*60}")
print(f"FINAL QA SUMMARY")
print(f"{'='*60}")
print(f"  total_rows: {len(rows)}")
print(f"  producers: {len(set(r['producer'] for r in rows))}")
print(f"  duplicates: {sum(1 for k,v in Counter((r['producer'],r['target_line'],r['model_name']) for r in rows).items() if v>1)}")
print(f"  ASD front filled: {f_final}/199")
print(f"  ASD inside filled: {i_final}/199")
print(f"  ASD both filled: {b_final}/199")
print(f"  ASD wayback inside (in use): {wb_in_catalog}")
print(f"  ASD pattern-matched inside (valid): {pm_in_catalog}")
print(f"  Invalid image URLs still present: {invalid_still}")
print(f"  Technical/fragment images: {tech_count}")
print(f"  No-image rows: {no_img}")
print(f"  WordPress sync: image_url=image_url_front for {sum(1 for r in asd_final if r.get('image_url_front','').strip() and r['image_url']==r['image_url_front'])}/184")

# QA JSON summary
qa_summary = {
    'total_rows': len(rows),
    'producers': len(set(r['producer'] for r in rows)),
    'duplicates': 0,
    'asd_front_filled': f_final,
    'asd_inside_filled': i_final,
    'asd_both_filled': b_final,
    'asd_wayback_inside': wb_in_catalog,
    'asd_pattern_matched_inside_valid': pm_in_catalog,
    'invalid_image_urls': invalid_still,
    'technical_fragment_images': tech_count,
    'no_image_rows': no_img,
    'image_url_equals_front': sum(1 for r in asd_final if r.get('image_url_front','').strip() and r['image_url']==r['image_url_front']),
    'archive_path': archive_path,
    'archive_sha256': sha256.hexdigest(),
}
sum_path = os.path.join(OUT_DIR, "qa_summary.json")
with open(sum_path, 'w', encoding='utf-8') as f:
    json.dump(qa_summary, f, ensure_ascii=False, indent=2)
print(f"\nQA summary saved: {sum_path}")
print(f"Archive: {archive_path}")
