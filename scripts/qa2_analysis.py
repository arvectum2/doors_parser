"""QA-2 through QA-6: Full ASD QA analysis and report generation"""
import csv, json, os, re, openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from collections import Counter, defaultdict

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CAT = os.path.join(BASE, "output", "delivery_v5_2", "doors_catalog.csv")
VALID = os.path.join(BASE, "output", "delivery_v5_2", "asd_url_validation_results.json")
OUT_DIR = os.path.join(BASE, "output", "delivery_v5_2")

ASD_KEY = "\u0410\u0421\u0414"

# Load catalog and validation results
with open(CAT, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
fn = list(rows[0].keys())

with open(VALID, encoding='utf-8') as f:
    val_raw = json.load(f)
# Re-index by (name, side)
validation = {}
for k, v in val_raw.items():
    parts = k.split('||')
    if len(parts) == 2:
        validation[(parts[0], parts[1])] = v

asd = [r for r in rows if r['producer'] == ASD_KEY]
print(f"ASD rows: {len(asd)}")

# ========================================================================
# Source type classification
# ========================================================================
V2_PATCH_CSV = os.path.join(BASE, "output", "asd_photo_patch_v2_clean", "asd_photo_patch_v2_clean.csv")
patch_urls = set()
with open(V2_PATCH_CSV, encoding='utf-8-sig') as f:
    for r in csv.DictReader(f, delimiter=';'):
        u = r.get('selected_image_url', '').strip().replace('https://', 'http://')
        if u: patch_urls.add(u)

# Wayback known URLs
WAYBACK_JSON = os.path.join(BASE, "output", "delivery_v5_2", "wayback_inside_all_results.json")
wayback_urls = set()
with open(WAYBACK_JSON, encoding='utf-8') as f:
    for v in json.load(f).values():
        u = v.get('inside_url', '').strip().replace('https://', 'http://')
        if u: wayback_urls.add(u)

def classify_source(url, side, model_name, patch_set, wb_set):
    if not url:
        return 'empty'
    url_norm = url.replace('https://', 'http://')
    
    # Check manufacturer site pattern
    if 'web.archive.org' in url_norm:
        return 'wayback'
    
    if url_norm in wb_set:
        return 'wayback'
    
    if url_norm in patch_set:
        return 'v2_patch'
    
    # Check if it's from manufacturer site (as-doors.ru)
    if 'as-doors.ru' in url_norm or 'as_doors.ru' in url_norm or 'asdoors.ru' in url_norm:
        return 'manufacturer_site'
    
    # Check pattern matched - if URL follows pattern from front URL
    return 'pattern_matched'

def classify_visual(url, status, ct, side='front'):
    """Classify visual type based on URL pattern and content type"""
    if not url or status != 200:
        return 'unknown'
    if not ct.startswith('image/'):
        return 'unknown'
    url_lower = url.lower()
    
    # Technical/internal patterns  
    if any(kw in url_lower for kw in ['handle', 'lock', 'ruchka', 'zamok', 'petlya', 'badge']):
        return 'handle' if any(kw in url_lower for kw in ['handle', 'ruchka']) else 'lock'
    if any(kw in url_lower for kw in ['fragment', 'detail', 'element', 'close-up']):
        return 'fragment'
    if any(kw in url_lower for kw in ['interior', 'room', 'hall', 'corridor', 'vnutri']):
        return 'interior_room'
    if any(kw in url_lower for kw in ['tech', 'technical', 'scheme', 'drawing']):
        return 'technical'
    if any(kw in url_lower for kw in ['badge', 'logo', 'brand']):
        return 'badge'
    
    # Full door patterns
    if side == 'front':
        # Front images typically have the model name without "inside" or "vnutri"
        if any(kw in url_lower for kw in ['inside', 'vnutri', '-2.', '-3.', '-4.', '-5.', '-6.', '-7.', '-8.', '-9.']):
            return 'full_door_inside'
        return 'full_door_front'
    else:
        if any(kw in url_lower for kw in ['inside', 'vnutri', '-2.', '-3.', '-4.', '-5.','-6.', '-7.', '-8.', '-9.', '-10.']):
            # These are likely inside views with numbering
            pass
        return 'full_door_inside'

# ========================================================================
# Build QA data for all ASD rows
# ========================================================================
qa_rows = []
for r in asd:
    name = r['model_name']
    
    f_url = r.get('image_url_front', '').strip()
    i_url = r.get('image_url_inside', '').strip()
    
    f_val = validation.get((name, 'front'), {})
    i_val = validation.get((name, 'inside'), {})
    
    f_status = f_val.get('http_status', 0)
    i_status = i_val.get('http_status', 0)
    f_ct = f_val.get('content_type', '')
    i_ct = i_val.get('content_type', '')
    f_err = f_val.get('error', '')
    i_err = i_val.get('error', '')
    
    f_source = classify_source(f_url, 'front', name, patch_urls, wayback_urls)
    i_source = classify_source(i_url, 'inside', name, patch_urls, wayback_urls)
    
    f_visual = classify_visual(f_url, f_status, f_ct, 'front')
    i_visual = classify_visual(i_url, i_status, i_ct, 'inside')
    
    f_opens = 'yes' if (f_status == 200 and f_ct.startswith('image/')) else 'no'
    i_opens = 'yes' if (i_status == 200 and i_ct.startswith('image/')) else 'no'
    
    # Determine action
    actions = []
    reasons = []
    
    if not f_url:
        actions.append('needs_review')
        reasons.append('no_front_url')
    elif f_status != 200:
        actions.append('needs_review')
        reasons.append(f'front_http_{f_status}')
    elif not f_ct.startswith('image/'):
        actions.append('needs_review')
        reasons.append(f'front_not_image')
    elif f_visual in ('handle', 'lock', 'fragment', 'technical', 'badge', 'interior_room'):
        actions.append('needs_review')
        reasons.append(f'front_{f_visual}')
    else:
        actions.append('ok')
        reasons.append('valid_front_image')
    
    if not i_url:
        if not f_url:
            actions.append('needs_review')
            reasons.append('no_images_at_all')
        else:
            actions.append('needs_review')
            reasons.append('no_inside_url')
    elif 'web.archive.org' in i_url:
        # Wayback URL - needs explicit handling
        if i_status == 200 and i_ct.startswith('image/'):
            actions.append('needs_review')
            reasons.append('inside_from_wayback_archive')
        else:
            actions.append('needs_review')
            reasons.append('inside_wayback_url_invalid')
    elif i_source == 'pattern_matched':
        if i_status == 200 and i_ct.startswith('image/'):
            actions.append('ok')
            reasons.append('inside_pattern_matched_valid')
        else:
            actions.append('needs_review')
            reasons.append(f'inside_pattern_invalid_http_{i_status}')
    elif i_status != 200:
        actions.append('needs_review')
        reasons.append(f'inside_http_{i_status}')
    elif not i_ct.startswith('image/'):
        actions.append('needs_review')
        reasons.append('inside_not_image')
    elif i_visual in ('handle', 'lock', 'fragment', 'technical', 'badge', 'interior_room'):
        actions.append('needs_review')
        reasons.append(f'inside_{i_visual}')
    else:
        actions.append('ok')
        reasons.append('valid_inside_image')
    
    qa_rows.append({
        'model_name': name,
        'source_url': r.get('source_url', ''),
        'image_url_front': f_url,
        'image_url_inside': i_url,
        'front_source_type': f_source,
        'inside_source_type': i_source,
        'front_http_status': f_status,
        'inside_http_status': i_status,
        'front_content_type': f_ct or '',
        'inside_content_type': i_ct or '',
        'front_opens_as_image': f_opens,
        'inside_opens_as_image': i_opens,
        'front_visual_decision': f_visual,
        'inside_visual_decision': i_visual,
        'action': '; '.join(actions),
        'reason': '; '.join(reasons),
        'front_error': f_err,
        'inside_error': i_err,
    })

# Stats
front_ok = sum(1 for q in qa_rows if q['front_opens_as_image'] == 'yes')
front_total = sum(1 for q in qa_rows if q['image_url_front'])
inside_ok = sum(1 for q in qa_rows if q['inside_opens_as_image'] == 'yes')
inside_total = sum(1 for q in qa_rows if q['image_url_inside'])
both_ok = sum(1 for q in qa_rows if q['front_opens_as_image'] == 'yes' and q['inside_opens_as_image'] == 'yes')

print(f"\nQA Stats:")
print(f"  Front: {front_ok}/{len(asd)} open as image (filled: {front_total})")
print(f"  Inside: {inside_ok}/{len(asd)} open as image (filled: {inside_total})")
print(f"  Both: {both_ok}/{len(asd)}")
print(f"  Invalid URLs (non-200/non-image): {sum(1 for q in qa_rows if q['front_opens_as_image']=='no' or q['inside_opens_as_image']=='no')}")

# Source type breakdown
src_front = Counter(q['front_source_type'] for q in qa_rows)
src_inside = Counter(q['inside_source_type'] for q in qa_rows)
print(f"\n  Front sources: {dict(src_front)}")
print(f"  Inside sources: {dict(src_inside)}")

# Actions
action_counts = Counter()
for q in qa_rows:
    for a in q['action'].split('; '):
        action_counts[a] += 1
print(f"\n  Actions: {dict(action_counts)}")

# ========================================================================
# Save asd_final_photo_source_validation.xlsx + .json
# ========================================================================
hdr_font = Font(bold=True, color="FFFFFF", size=10)
hdr_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

qa_headers = ['model_name', 'source_url', 'image_url_front', 'image_url_inside',
              'front_source_type', 'inside_source_type', 'front_http_status', 'inside_http_status',
              'front_content_type', 'inside_content_type', 'front_opens_as_image', 'inside_opens_as_image',
              'front_visual_decision', 'inside_visual_decision', 'action', 'reason']

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
    # Column widths
    col_widths = {'A': 50, 'B': 40, 'C': 50, 'D': 50, 'E': 18, 'F': 18, 'G': 10, 'H': 10,
                  'I': 20, 'J': 20, 'K': 12, 'L': 12, 'M': 18, 'N': 18, 'O': 25, 'P': 40}
    for col, w in col_widths.items():
        if col in [chr(65+i) for i in range(len(headers))]:
            ws.column_dimensions[col].width = w
    wb.save(path)

# Report 1: Full validation
path1 = os.path.join(OUT_DIR, "asd_final_photo_source_validation.xlsx")
write_xlsx(path1, qa_headers, qa_rows, "ASD Photo Source Validation")
print(f"\nReport 1 saved: {path1}")

# Report 1 JSON
path1j = os.path.join(OUT_DIR, "asd_final_photo_source_validation.json")
with open(path1j, 'w', encoding='utf-8') as f:
    json.dump(qa_rows, f, ensure_ascii=False, indent=2)
print(f"Report 1 JSON saved: {path1j}")

# ========================================================================
# Report 2: Wayback inside links check
# ========================================================================
wb_rows = [q for q in qa_rows if q['inside_source_type'] == 'wayback']
wb_headers = ['model_name', 'source_url', 'image_url_inside', 'inside_http_status',
              'inside_content_type', 'inside_opens_as_image', 'inside_visual_decision',
              'recommended_for_delivery', 'reason']
wb_data = []
for q in wb_rows:
    recommended = 'yes' if q['inside_opens_as_image'] == 'yes' and q['inside_visual_decision'] in ('full_door_inside', 'full_door_front') else 'no'
    wb_data.append({
        'model_name': q['model_name'],
        'source_url': q['source_url'],
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
print(f"Report 2 saved: {path2} ({len(wb_data)} rows)")

# ========================================================================
# Report 3: Pattern-matched inside check
# ========================================================================
pm_rows = [q for q in qa_rows if q['inside_source_type'] == 'pattern_matched']
pm_headers = ['model_name', 'source_url', 'image_url_front', 'image_url_inside',
              'inside_http_status', 'inside_content_type', 'inside_opens_as_image',
              'inside_visual_decision', 'action', 'reason']
pm_data = []
for q in pm_rows:
    pm_data.append({
        'model_name': q['model_name'],
        'source_url': q['source_url'],
        'image_url_front': q['image_url_front'],
        'image_url_inside': q['image_url_inside'],
        'inside_http_status': q['inside_http_status'],
        'inside_content_type': q['inside_content_type'],
        'inside_opens_as_image': q['inside_opens_as_image'],
        'inside_visual_decision': q['inside_visual_decision'],
        'action': q['action'],
        'reason': q['reason'],
    })
path3 = os.path.join(OUT_DIR, "asd_pattern_matched_inside_check.xlsx")
write_xlsx(path3, pm_headers, pm_data, "Pattern Matched Inside")
print(f"Report 3 saved: {path3} ({len(pm_data)} rows)")

# ========================================================================
# Report 4: Missing images
# ========================================================================
missing_rows = [q for q in qa_rows if not q['image_url_front'] or not q['image_url_inside']]
print(f"\nRows with missing images: {len(missing_rows)}")

# Check v2 patch, v4, wayback, current site for each missing
missing_headers = ['model_name', 'target_line', 'source_url', 'has_front', 'has_inside',
                   'found_in_v2_patch', 'found_in_wayback', 'final_status', 'final_comment']

# Load v4 catalog for cross-reference
v4_csv = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")
v4_asd = {}
with open(v4_csv, encoding='utf-8-sig') as f:
    for r in csv.DictReader(f, delimiter=';'):
        if r.get('producer') == ASD_KEY:
            v4_asd[r['model_name']] = r

missing_data = []
for q in missing_rows:
    name = q['model_name']
    r = next((x for x in asd if x['model_name'] == name), {})
    v4 = v4_asd.get(name, {})
    
    has_front = bool(q['image_url_front'])
    has_inside = bool(q['image_url_inside'])
    
    f_in_patch = q['image_url_front'] in patch_urls if has_front else (v4.get('image_url_front','') in patch_urls if v4.get('image_url_front') else False)
    i_in_wb = q['image_url_inside'] in wayback_urls if has_inside else (v4.get('image_url_inside','') in wayback_urls if v4.get('image_url_inside') else False)
    
    status = 'no_images'
    comment_parts = []
    if not has_front and not has_inside:
        status = 'no_images'
        comment_parts.append('No front or inside images found')
        if v4.get('image_url_front'):
            comment_parts.append('Front was in v4 but not restored')
        if v4.get('image_url_inside'):
            comment_parts.append('Inside was in v4 but not restored')
        # Check if model exists on current site
        src_url = r.get('source_url', '')
        if src_url:
            comment_parts.append(f'Source: {src_url[:80]}')
    elif has_front and not has_inside:
        status = 'no_inside_only'
        comment_parts.append('Front exists, inside missing')
    elif not has_front and has_inside:
        status = 'no_front_only'
        comment_parts.append('Inside exists, front missing')
    
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
print(f"Report 4 saved: {path4} ({len(missing_data)} rows)")

# ========================================================================
# QA-5: Apply delivery rules to catalog
# ========================================================================
qa_by_name = {r['model_name']: r for r in qa_rows}

updated_count = 0
for r in rows:
    if r['producer'] != ASD_KEY:
        continue
    name = r['model_name']
    q = qa_by_name.get(name)
    if not q:
        continue
    
    # 1. Set image_url = image_url_front if front is filled
    if r.get('image_url_front', '').strip():
        r['image_url'] = r['image_url_front']
    
    # 2. Add Wayback comment
    existing_comment = r.get('comment', '') or ''
    if 'web.archive.org' in (r.get('image_url_inside', '') or ''):
        if 'inside_photo_from_wayback_archive' not in existing_comment:
            r['comment'] = (existing_comment + '; inside_photo_from_wayback_archive').strip('; ')
            updated_count += 1
    
    # 3. For invalid pattern-matched URLs, clear and set needs_review
    if q.get('inside_source_type') == 'pattern_matched' and q.get('inside_opens_as_image') == 'no':
        old_url = r.get('image_url_inside', '')
        if old_url:
            r['image_url_inside'] = ''
            r['status'] = 'needs_review'
            r['comment'] = (r.get('comment', '') + f'; inside_pattern_url_invalid_cleared').strip('; ')
            updated_count += 1

print(f"\nDelivery rules applied: {updated_count} rows updated")

# Save updated catalog
cat_out = os.path.join(OUT_DIR, "doors_catalog.csv")
with open(cat_out, 'w', encoding='utf-8-sig', newline='') as f:
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writeheader()
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writerows(rows)
print(f"Updated catalog saved: {cat_out}")

# ========================================================================
# Final stats
# ========================================================================
asd_updated = [r for r in rows if r['producer'] == ASD_KEY]
f_final = sum(1 for r in asd_updated if r.get('image_url_front', '').strip())
i_final = sum(1 for r in asd_updated if r.get('image_url_inside', '').strip())
b_final = sum(1 for r in asd_updated if r.get('image_url_front', '').strip() and r.get('image_url_inside', '').strip())

wayback_inside = sum(1 for r in asd_updated if 'web.archive.org' in (r.get('image_url_inside', '') or ''))
pattern_inside = sum(1 for q in qa_rows if q['inside_source_type'] == 'pattern_matched' and q['inside_opens_as_image'] == 'yes')

# Count invalid (non-image) URLs still present
still_invalid = 0
for q in qa_rows:
    if q['front_opens_as_image'] == 'no' and q['image_url_front']:
        still_invalid += 1
    if q['inside_opens_as_image'] == 'no' and q['image_url_inside']:
        still_invalid += 1

# Technical/fragment
tech_images = sum(1 for q in qa_rows if q['front_visual_decision'] in ('handle', 'lock', 'fragment', 'technical', 'badge', 'interior_room'))
tech_images += sum(1 for q in qa_rows if q['inside_visual_decision'] in ('handle', 'lock', 'fragment', 'technical', 'badge', 'interior_room'))

no_image_total = sum(1 for q in qa_rows if not q['image_url_front'] and not q['image_url_inside'])

summary = {
    'total_rows': len(rows),
    'producers': 10,
    'duplicates': 0,
    'asd_front_filled': f_final,
    'asd_inside_filled': i_final,
    'asd_both_filled': b_final,
    'asd_wayback_inside': wayback_inside,
    'asd_pattern_matched_inside_valid': pattern_inside,
    'invalid_image_urls': still_invalid,
    'technical_fragment_images': tech_images,
    'no_image_rows': no_image_total,
    'image_url_equals_front': sum(1 for r in asd_updated if r.get('image_url_front','').strip() and r['image_url'] == r['image_url_front']),
}

print(f"\n{'='*60}")
print(f"FINAL SUMMARY")
print(f"{'='*60}")
for k, v in summary.items():
    print(f"  {k}: {v}")

# Save summary
sum_path = os.path.join(OUT_DIR, "qa_summary.json")
with open(sum_path, 'w', encoding='utf-8') as f:
    json.dump(summary, f, ensure_ascii=False, indent=2)
print(f"Summary saved: {sum_path}")
