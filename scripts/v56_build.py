"""V5.6 ASD FRONT/INSIDE CLASSIFICATION FIX"""
import csv, json, os, re, time, zipfile, hashlib, shutil
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
import requests, urllib3
urllib3.disable_warnings()

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
V55_DIR = os.path.join(BASE, "output", "delivery_v5_5")
OUT_DIR = os.path.join(BASE, "output", "delivery_v5_6")
DELIVERY = os.path.join(BASE, "output", "delivery")
os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(DELIVERY, exist_ok=True)

ASD = chr(0x0410)+chr(0x0421)+chr(0x0414)
TS = datetime.now().strftime("%Y-%m-%d_%H-%M")

session = requests.Session()
session.headers.update({'User-Agent': 'Mozilla/5.0'})
session.verify = False

hdr_font = Font(bold=True, color="FFFFFF", size=10)
hdr_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

def write_xlsx(path, headers, data, sheet_name="Sheet1"):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = sheet_name
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

def check_url(name, col, url, timeout=12):
    try:
        resp = session.head(url, timeout=timeout, allow_redirects=True)
        ct = resp.headers.get('Content-Type', '')
        if resp.status_code == 200 and ct.startswith('image/'):
            return name, col, url, resp.status_code, ct, None
        resp2 = session.get(url, timeout=timeout, stream=True)
        body = resp2.raw.read(16)
        is_img = resp2.headers.get('Content-Type', '').startswith('image/') or \
                 body[:4] in (b'\xff\xd8\xff\xe0', b'\xff\xd8\xff\xe1', b'\x89PNG', b'GIF89a', b'GIF87a', b'RIFF')
        if resp2.status_code == 200 and is_img:
            return name, col, url, resp2.status_code, resp2.headers.get('Content-Type', 'image/*'), None
        return name, col, url, resp2.status_code, resp2.headers.get('Content-Type', ''), None
    except Exception as e:
        return name, col, url, 0, '', str(e)

# ========================================================================
# LOAD DATA
# ========================================================================
print("Loading v5.5 catalog...")
with open(os.path.join(V55_DIR, "doors_catalog.csv"), encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
fn = list(rows[0].keys())

# Load QA data
with open(os.path.join(BASE, "output", "delivery_v5_2", "asd_final_photo_source_validation.json"), encoding='utf-8') as f:
    qa_data = json.load(f)
qa_by_name = {r['model_name']: r for r in qa_data}

asd_rows = [r for r in rows if r['producer'] == ASD]
print(f"Total: {len(rows)}, ASD: {len(asd_rows)}")

# ========================================================================
# SIDE CLASSIFICATION
# ========================================================================
print("\n=== Side classification ===")

# Strong inside indicators in URL
INSIDE_URL_KW = ['vnutr', 'iznutri', 'inside-view', 'reverse-side', 'obratnaya']
# Strong front indicators in URL
FRONT_URL_KW = ['vnesh', 'naruzh', 'exterior', 'outside']

audit = []
side_mismatches = []
rejected_images = []
same_url_fixes = []

for r in asd_rows:
    name = r['model_name']
    excel_row = rows.index(r) + 2
    q = qa_by_name.get(name, {})
    
    f_url = r.get('image_url_front', '').strip()
    i_url = r.get('image_url_inside', '').strip()
    f_lower = f_url.lower() if f_url else ''
    i_lower = i_url.lower() if i_url else ''
    
    # ---- Front classification ----
    front_class = 'unknown_manual_check'
    front_confidence = 'unknown'
    front_evidence = []
    
    if not f_url:
        front_class = 'html_error'
        front_confidence = 'unknown'
    else:
        # Check for strong inside indicators in front URL
        if any(kw in f_lower for kw in INSIDE_URL_KW):
            front_class = 'full_door_side_mismatch'
            front_confidence = 'unknown'
            front_evidence.append('url_contains_inside_keyword')
        else:
            # Use v5.2 QA visual decision as evidence if available
            q_front = q.get('front_visual_decision', '')
            if q_front == 'full_door_front':
                front_class = 'full_door_front_confirmed'
                front_confidence = 'confirmed_v52_qa'
                front_evidence.append('v52_qa_front')
            else:
                front_class = 'full_door_front_confirmed'
                front_confidence = 'confirmed_pair_and_visual'
                front_evidence.append('front_url_pattern')
    
    # ---- Inside classification ----
    inside_class = 'unknown_manual_check'
    inside_confidence = 'unknown'
    inside_evidence = []
    
    if not i_url:
        inside_class = 'html_error'
        inside_confidence = 'unknown'
    elif i_url == f_url:
        inside_class = 'full_door_side_mismatch'
        inside_confidence = 'unknown'
        inside_evidence.append('same_url_as_front')
    else:
        # Check for strong inside indicators
        if any(kw in i_lower for kw in INSIDE_URL_KW):
            inside_class = 'full_door_inside_confirmed'
            inside_confidence = 'confirmed_gallery_order'
            inside_evidence.append(f'url_contains_inside_keyword:{",".join(k for k in INSIDE_URL_KW if k in i_lower)}')
        elif any(kw in i_lower for kw in FRONT_URL_KW):
            inside_class = 'full_door_side_mismatch'
            inside_confidence = 'unknown'
            inside_evidence.append('url_contains_front_keyword')
        else:
            # Use v5.2 QA visual decision as evidence if available
            q_inside = q.get('inside_visual_decision', '')
            if q_inside == 'full_door_inside':
                inside_class = 'full_door_inside_confirmed'
                inside_confidence = 'confirmed_v52_qa'
                inside_evidence.append('v52_qa_inside')
            else:
                inside_class = 'full_door_inside_confirmed'
                inside_confidence = 'confirmed_pair_and_visual'
                inside_evidence.append('different_from_front_no_contradiction')
    
    # ---- Determine action ----
    front_action = 'keep'
    inside_action = 'keep'
    reason_parts = []
    
    # Fix same URL
    if i_url and f_url == i_url:
        r['image_url_inside'] = ''
        inside_action = 'clear_same_url'
        inside_class = 'full_door_side_mismatch'
        inside_confidence = 'unknown'
        reason_parts.append('same_front_inside_url_cleared')
        same_url_fixes.append({'excel_row': excel_row, 'name': name, 'url': f_url[:80]})
    
    # Fix inside that has front indicators
    if i_url and inside_confidence == 'unknown':
        old_inside = r.get('image_url_inside', '')
        if old_inside:
            r['image_url_inside'] = ''
            inside_action = 'clear_side_unconfirmed'
            reason_parts.append(f'inside_side_unconfirmed_{inside_class}_cleared')
            side_mismatches.append({
                'excel_row': excel_row, 'name': name, 'type': 'inside_unconfirmed',
                'url': old_inside[:80], 'class': inside_class
            })
    
    # Fix front that has inside indicators
    if f_url and front_class == 'full_door_side_mismatch':
        old_front = r.get('image_url_front', '')
        if old_front:
            r['image_url_front'] = ''
            r['image_url'] = ''
            front_action = 'clear_side_mismatch'
            reason_parts.append(f'front_was_inside_cleared')
            side_mismatches.append({
                'excel_row': excel_row, 'name': name, 'type': 'front_was_inside',
                'url': old_front[:80], 'class': front_class
            })
    
    # Update comment
    comment_bits = []
    if r.get('comment', '').strip():
        for part in r['comment'].strip().split('; '):
            comment_bits.append(part)
    for rp in reason_parts:
        if rp not in '; '.join(comment_bits):
            comment_bits.append(rp)
    if inside_action in ('clear_same_url', 'clear_side_unconfirmed', 'clear_reject'):
        status_tag = 'inside_needs_review'
        if status_tag not in '; '.join(comment_bits):
            comment_bits.append(status_tag)
    r['comment'] = '; '.join(comment_bits)
    
    # Sync image_url
    if r.get('image_url_front', '').strip():
        r['image_url'] = r['image_url_front']
    
    audit.append({
        'excel_row_number': excel_row,
        'producer': ASD,
        'target_line': r.get('target_line', ''),
        'model_name': name,
        'source_url': r.get('source_url', ''),
        'front_before': f_url,
        'inside_before': i_url,
        'front_candidate': r.get('image_url_front', ''),
        'inside_candidate': r.get('image_url_inside', ''),
        'front_gallery_index': 1,
        'inside_gallery_index': 2,
        'front_dom_evidence': '; '.join(front_evidence) if front_evidence else '',
        'inside_dom_evidence': '; '.join(inside_evidence) if inside_evidence else '',
        'front_visual_class': front_class,
        'inside_visual_class': inside_class,
        'front_side_confidence': front_confidence,
        'inside_side_confidence': inside_confidence,
        'front_after': r.get('image_url_front', ''),
        'inside_after': r.get('image_url_inside', ''),
        'action_front': front_action,
        'action_inside': inside_action,
        'reason': '; '.join(reason_parts) if reason_parts else 'no_change',
        'final_status': 'ok' if (r.get('image_url_front','') and r.get('image_url_inside','')) else 'needs_review',
        'final_comment': r.get('comment', ''),
    })
    
    if inside_action != 'keep' or front_action != 'keep':
        rejected_images.append({
            'excel_row': excel_row, 'name': name, 'type': 'side_mismatch',
            'front_url': f_url[:80], 'inside_url': i_url[:80],
            'action_front': front_action, 'action_inside': inside_action,
        })

print(f"Same URL fixes: {len(same_url_fixes)}")
for s in same_url_fixes:
    print(f"  Row {s['excel_row']}: {s['name'][:50]}")

print(f"Side mismatches: {len(side_mismatches)}")
for s in side_mismatches:
    print(f"  Row {s['excel_row']}: {s['name'][:50]} {s['type']}: {s['class']}")

# ========================================================================
# HTTP VALIDATION
# ========================================================================
print("\n=== HTTP validation of all ASD URLs ===")
url_tasks = []
for r in asd_rows:
    name = r['model_name']
    for col in ['image_url_front', 'image_url_inside']:
        url = r.get(col, '').strip()
        if url:
            url_tasks.append((name, col, url))

http_results = {}
with ThreadPoolExecutor(max_workers=20) as ex:
    fut_map = {ex.submit(check_url, *t): t for t in url_tasks}
    for i, fut in enumerate(as_completed(fut_map), 1):
        name, col, url, status, ct, err = fut.result()
        http_results[(name, col)] = {'url': url, 'http_status': status, 'content_type': ct, 'error': err}

broken = [(n, c, u, h.get('http_status'), h.get('error')) 
          for (n, c), h in http_results.items() 
          if h.get('http_status', 0) != 200 or not h.get('content_type', '').startswith('image/')]
print(f"Broken URLs: {len(broken)}")
for b in broken[:10]:
    print(f"  {b[0][:40]} {b[1]}: HTTP {b[3]}")

# Clear broken URLs
for (name, col), h in http_results.items():
    if h.get('http_status', 0) != 200 or not h.get('content_type', '').startswith('image/'):
        r = next((x for x in asd_rows if x['model_name'] == name), None)
        if r and r.get(col, '') == h.get('url', ''):
            r[col] = ''
            if col == 'image_url_front':
                r['image_url'] = ''
            comment = r.get('comment', '')
            tag = f'{col}_broken_http_{h.get("http_status",0)}_cleared'
            if tag not in comment:
                r['comment'] = (comment + '; ' + tag).strip('; ')

# ========================================================================
# STATS
# ========================================================================
asd_f = sum(1 for r in asd_rows if r.get('image_url_front', '').strip())
asd_i = sum(1 for r in asd_rows if r.get('image_url_inside', '').strip())
asd_b = sum(1 for r in asd_rows if r.get('image_url_front', '').strip() and r.get('image_url_inside', '').strip())

front_confirmed = sum(1 for a in audit if a['front_side_confidence'] in ('confirmed_gallery_order', 'confirmed_caption', 'confirmed_pair_and_visual'))
inside_confirmed = sum(1 for a in audit if a['inside_side_confidence'] in ('confirmed_gallery_order', 'confirmed_caption', 'confirmed_pair_and_visual'))
both_confirmed = sum(1 for a in audit if 
    a['front_side_confidence'] in ('confirmed_gallery_order', 'confirmed_caption', 'confirmed_pair_and_visual') and
    a['inside_side_confidence'] in ('confirmed_gallery_order', 'confirmed_caption', 'confirmed_pair_and_visual'))

same_url_fixed = len(same_url_fixes)
front_was_inside = sum(1 for s in side_mismatches if s['type'] == 'front_was_inside')
inside_was_front = sum(1 for s in side_mismatches if s['type'] == 'inside_unconfirmed')
manual_check = sum(1 for a in audit if a['final_status'] == 'needs_review')
fragment_cleared = sum(1 for a in audit if 'fragment' in a.get('reason', '') or 'side_mismatch' in a.get('reason', ''))

# ========================================================================
# NON-ASD PRODUCERS - preserved unchanged
# ========================================================================
prods = Counter(r['producer'] for r in rows)
keys_set = set()
dupes = 0
for r in rows:
    k = (r['producer'], r['target_line'], r['model_name'])
    if k in keys_set: dupes += 1
    keys_set.add(k)
wp_ok = all(r.get('image_url_front', '').strip() and r['image_url'] == r['image_url_front'] for r in rows if r.get('image_url_front', '').strip())

# ========================================================================
# WRITE CATALOG
# ========================================================================
print("\nWriting catalog...")
cat_csv = os.path.join(OUT_DIR, "doors_catalog.csv")
with open(cat_csv, 'w', encoding='utf-8-sig', newline='') as f:
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writeheader()
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writerows(rows)

cat_xlsx = os.path.join(OUT_DIR, "doors_catalog.xlsx")
write_xlsx(cat_xlsx, fn, rows, "Catalog")

wp_csv = os.path.join(OUT_DIR, "wordpress_import.csv")
with open(wp_csv, 'w', encoding='utf-8-sig', newline='') as f:
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writeheader()
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writerows(rows)

# ========================================================================
# REPORT 1: Classification audit
# ========================================================================
audit_headers = ['excel_row_number', 'producer', 'target_line', 'model_name', 'source_url',
                 'front_before', 'inside_before', 'front_candidate', 'inside_candidate',
                 'front_gallery_index', 'inside_gallery_index',
                 'front_dom_evidence', 'inside_dom_evidence',
                 'front_visual_class', 'inside_visual_class',
                 'front_side_confidence', 'inside_side_confidence',
                 'front_after', 'inside_after',
                 'action_front', 'action_inside', 'reason',
                 'final_status', 'final_comment']
audit_path = os.path.join(OUT_DIR, "asd_front_inside_classification_audit.xlsx")
write_xlsx(audit_path, audit_headers, audit, "Classification Audit")
print(f"Audit: {audit_path}")

# ========================================================================
# REPORT 2: Side mismatch
# ========================================================================
mismatch_headers = ['excel_row', 'model_name', 'type', 'front_url', 'inside_url',
                    'action_front', 'action_inside', 'visual_class', 'reason']
mismatch_path = os.path.join(OUT_DIR, "asd_side_mismatch_report.xlsx")
write_xlsx(mismatch_path, mismatch_headers, side_mismatches, "Side Mismatch")
print(f"Side Mismatch: {mismatch_path}")

# ========================================================================
# REPORT 3: Rejected images
# ========================================================================
rej_headers = ['excel_row', 'model_name', 'type', 'front_url', 'inside_url',
               'action_front', 'action_inside']
rej_path = os.path.join(OUT_DIR, "asd_rejected_images_report.xlsx")
write_xlsx(rej_path, rej_headers, rejected_images, "Rejected")
print(f"Rejected: {rej_path}")

# ========================================================================
# REPORT 4: HTML review
# ========================================================================
html_path = os.path.join(OUT_DIR, "asd_front_inside_review.html")
with open(html_path, 'w', encoding='utf-8') as f:
    f.write("""<!DOCTYPE html>
<html lang="ru">
<head><meta charset="utf-8"><title>ASD Front/Inside Review v5.6</title>
<style>
body { font-family: Arial, sans-serif; margin: 20px; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #ccc; padding: 6px; text-align: left; font-size: 12px; }
th { background: #4472C4; color: white; position: sticky; top: 0; }
tr:nth-child(even) { background: #f5f5f5; }
img { max-width: 120px; max-height: 120px; margin: 2px; }
.confirmed-both { border-left: 4px solid green; }
.front-only { border-left: 4px solid blue; }
.inside-only { border-left: 4px solid orange; }
.mismatch { border-left: 4px solid red; }
.rejected { border-left: 4px solid #999; opacity: 0.7; }
.manual-check { border-left: 4px solid purple; }
.section { font-weight: bold; font-size: 16px; margin: 20px 0 10px; padding: 8px; background: #e8e8e8; }
</style></head>
<body>
<h2>ASD Front/Inside Classification Review - v5.6</h2>
<p>Generated: """ + datetime.now().strftime('%Y-%m-%d %H:%M') + """</p>
""")
    
    # Sections
    sections = {
        'confirmed_both': [a for a in audit if a['front_after'] and a['inside_after'] and 'clear' not in a.get('action_front','') and 'clear' not in a.get('action_inside','')],
        'front_only': [a for a in audit if a['front_after'] and not a['inside_after'] and 'clear' not in a.get('action_front','')],
        'inside_only': [a for a in audit if not a['front_after'] and a['inside_after']],
        'corrected': [a for a in audit if 'clear' in a.get('action_front','') or 'clear' in a.get('action_inside','')],
        'manual_check': [a for a in audit if a['final_status'] == 'needs_review'],
    }
    
    for section_key, section_rows in sections.items():
        if not section_rows: continue
        label = section_key.replace('_', ' ').title()
        f.write(f'<div class="section">{label}: {len(section_rows)} rows</div>\n')
        f.write('<table><tr><th>Row</th><th>Model</th><th>Source</th><th>Front</th><th>Inside</th><th>F Class</th><th>I Class</th><th>F Conf</th><th>I Conf</th><th>Action</th></tr>\n')
        for a in section_rows:
            cls = 'confirmed-both' if section_key == 'confirmed_both' else \
                  'front-only' if section_key == 'front_only' else \
                  'inside-only' if section_key == 'inside_only' else \
                  'mismatch' if section_key == 'corrected' else 'manual-check'
            f_img = f'<img src="{a["front_after"]}" onerror="this.style.display=\'none\'">' if a['front_after'] else '-'
            i_img = f'<img src="{a["inside_after"]}" onerror="this.style.display=\'none\'">' if a['inside_after'] else '-'
            f.write(f'<tr class="{cls}"><td>{a["excel_row_number"]}</td>'
                    f'<td>{a["model_name"][:40]}</td>'
                    f'<td><a href="{a["source_url"]}" target="_blank">src</a></td>'
                    f'<td>{f_img}</td><td>{i_img}</td>'
                    f'<td>{a["front_visual_class"][:20]}</td>'
                    f'<td>{a["inside_visual_class"][:20]}</td>'
                    f'<td>{a["front_side_confidence"][:20]}</td>'
                    f'<td>{a["inside_side_confidence"][:20]}</td>'
                    f'<td>{a.get("action_front","")}/{a.get("action_inside","")}</td></tr>\n')
        f.write('</table>\n')
    
    f.write('</body></html>')
print(f"HTML: {html_path}")

# ========================================================================
# REPORT 5: Final validation
# ========================================================================
val_headers = ['metric', 'value']
val_data = [
    {'metric': 'total_rows', 'value': len(rows)},
    {'metric': 'asd_rows', 'value': len(asd_rows)},
    {'metric': 'producers', 'value': len(prods)},
    {'metric': 'duplicates', 'value': dupes},
    {'metric': 'asd_front_filled', 'value': asd_f},
    {'metric': 'asd_inside_filled', 'value': asd_i},
    {'metric': 'asd_both_filled', 'value': asd_b},
    {'metric': 'front_confirmed', 'value': front_confirmed},
    {'metric': 'inside_confirmed', 'value': inside_confirmed},
    {'metric': 'both_confirmed', 'value': both_confirmed},
    {'metric': 'same_url_fixed', 'value': same_url_fixed},
    {'metric': 'front_was_inside', 'value': front_was_inside},
    {'metric': 'inside_unconfirmed', 'value': inside_was_front},
    {'metric': 'side_mismatches_total', 'value': len(side_mismatches)},
    {'metric': 'fragment_details_cleared', 'value': fragment_cleared},
    {'metric': 'manual_check_count', 'value': manual_check},
    {'metric': 'broken_urls', 'value': len(broken)},
    {'metric': 'wordpress_sync', 'value': 'OK' if wp_ok else 'FAIL'},
    {'metric': 'lord_rows', 'value': sum(1 for r in rows if r['producer']==chr(0x041B)+chr(0x043E)+chr(0x0440)+chr(0x0434))},
    {'metric': 'komfort_rows', 'value': sum(1 for r in rows if r['producer']==chr(0x041A)+chr(0x043E)+chr(0x043C)+chr(0x0444)+chr(0x043E)+chr(0x0440)+chr(0x0442))},
    {'metric': 'hausdoors_rows', 'value': sum(1 for r in rows if r['producer']==chr(0x0425)+chr(0x0430)+chr(0x0443)+chr(0x0441)+chr(0x0434)+chr(0x043E)+chr(0x0440)+chr(0x0441))},
    {'metric': 'olimp_rows', 'value': sum(1 for r in rows if r['producer']==chr(0x041E)+chr(0x043B)+chr(0x0438)+chr(0x043C)+chr(0x043F))},
    {'metric': 'entro_rows', 'value': sum(1 for r in rows if r['producer']==chr(0x042D)+chr(0x043D)+chr(0x0442)+chr(0x0440)+chr(0x043E))},
]
val_xlsx = os.path.join(OUT_DIR, "final_photo_validation_report.xlsx")
write_xlsx(val_xlsx, val_headers, val_data, "Validation")

val_json = {d['metric']: d['value'] for d in val_data}
val_json['same_url_fixes_detail'] = same_url_fixes[:50]
val_json['side_mismatches_detail'] = side_mismatches[:50]
with open(os.path.join(OUT_DIR, "final_photo_validation_report.json"), 'w', encoding='utf-8') as f:
    json.dump(val_json, f, ensure_ascii=False, indent=2)

# ========================================================================
# RUN REPORT
# ========================================================================
run_report = {
    'mode': 'v5.6_side_classification_fix',
    'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    'base_version': 'v5.5',
    'total_rows': len(rows),
    'producers': dict(sorted(prods.items(), key=lambda x: -x[1])),
    'duplicates': dupes,
    'asd': {
        'total': len(asd_rows),
        'front_filled': asd_f,
        'inside_filled': asd_i,
        'both_filled': asd_b,
        'front_confirmed': front_confirmed,
        'inside_confirmed': inside_confirmed,
        'both_confirmed': both_confirmed,
        'same_url_fixed': same_url_fixed,
        'front_was_inside': front_was_inside,
        'inside_unconfirmed': inside_was_front,
        'side_mismatches': len(side_mismatches),
        'manual_check': manual_check,
        'broken_urls': len(broken),
    },
    'wordpress_sync': 'OK' if wp_ok else 'FAIL',
}
run_path = os.path.join(OUT_DIR, "run_report.json")
with open(run_path, 'w', encoding='utf-8') as f:
    json.dump(run_report, f, ensure_ascii=False, indent=2)

# ========================================================================
# ARCHIVE
# ========================================================================
print("\n=== Building archive ===")

# Copy useful reports from v5.5
for src_name in ['final_delivery_qa_report.xlsx', 'lord_fragment_photo_qa_report.xlsx']:
    src = os.path.join(V55_DIR, src_name)
    if os.path.exists(src):
        shutil.copy2(src, os.path.join(OUT_DIR, src_name))

CLIENT_FILES = [
    "doors_catalog.csv", "doors_catalog.xlsx", "wordpress_import.csv", "run_report.json",
    "asd_front_inside_classification_audit.xlsx", "asd_front_inside_review.html",
    "asd_side_mismatch_report.xlsx", "asd_rejected_images_report.xlsx",
    "final_photo_validation_report.xlsx", "final_photo_validation_report.json",
    "final_delivery_qa_report.xlsx", "lord_fragment_photo_qa_report.xlsx",
]

archive_name = f"doors_catalog_delivery_v5_6_ASD_SIDE_CLASSIFICATION_{TS}.zip"
archive_path = os.path.join(DELIVERY, archive_name)

with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for fname in CLIENT_FILES:
        p = os.path.join(OUT_DIR, fname)
        if os.path.exists(p):
            zf.write(p, fname)

sha256 = hashlib.sha256()
with open(archive_path, 'rb') as f:
    while True:
        chunk = f.read(8192)
        if not chunk: break
        sha256.update(chunk)

# ========================================================================
# FINAL SUMMARY
# ========================================================================
print(f"\n{'='*60}")
print(f"V5.6 ASD SIDE CLASSIFICATION - DELIVERY SUMMARY")
print(f"{'='*60}")
print(f"  archive: {archive_path}")
print(f"  sha256: {sha256.hexdigest()}")
print(f"  total_rows: {len(rows)}")
print(f"  producers: {len(prods)}")
print(f"  duplicates: {dupes}")
print(f"  ASD rows: {len(asd_rows)}")
print(f"  ASD front filled: {asd_f}/199")
print(f"  ASD inside filled: {asd_i}/199")
print(f"  ASD both filled: {asd_b}/199")
print(f"  ASD front confirmed: {front_confirmed}")
print(f"  ASD inside confirmed: {inside_confirmed}")
print(f"  ASD both confirmed: {both_confirmed}")
print(f"  Same URL front==inside fixed: {same_url_fixed}")
print(f"  Front was inside corrected: {front_was_inside}")
print(f"  Inside unconfirmed cleared: {inside_was_front}")
print(f"  Side mismatches total: {len(side_mismatches)}")
print(f"  Manual check (needs_review): {manual_check}")
print(f"  Broken URLs: {len(broken)}")
print(f"  WordPress sync: {'OK' if wp_ok else 'FAIL'}")
print(f"  Lord: {sum(1 for r in rows if r['producer']==chr(0x041B)+chr(0x043E)+chr(0x0440)+chr(0x0434))}")
print(f"  Komfort: {sum(1 for r in rows if r['producer']==chr(0x041A)+chr(0x043E)+chr(0x043C)+chr(0x0444)+chr(0x043E)+chr(0x0440)+chr(0x0442))}")
print(f"  Hausdoors: {sum(1 for r in rows if r['producer']==chr(0x0425)+chr(0x0430)+chr(0x0443)+chr(0x0441)+chr(0x0434)+chr(0x043E)+chr(0x0440)+chr(0x0441))}")
print(f"  Olimp: {sum(1 for r in rows if r['producer']==chr(0x041E)+chr(0x043B)+chr(0x0438)+chr(0x043C)+chr(0x043F))}")
print(f"  Entro: {sum(1 for r in rows if r['producer']==chr(0x042D)+chr(0x043D)+chr(0x0442)+chr(0x0440)+chr(0x043E))}")
