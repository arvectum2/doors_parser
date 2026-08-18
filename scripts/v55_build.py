"""V5.5 PHOTO QA PATCH — усиленные правила + ASD inside re-QA + Lord fragment fix + global validation"""
import csv, json, os, re, openpyxl, zipfile, hashlib, time, shutil
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
import requests, urllib3
urllib3.disable_warnings()

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
SRC_DIR = os.path.join(BASE, "output", "delivery_v5_4")
OUT_DIR = os.path.join(BASE, "output", "delivery_v5_5")
DELIVERY = os.path.join(BASE, "output", "delivery")
os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(DELIVERY, exist_ok=True)

ASD_KEY = "\u0410\u0421\u0414"
LORD_KEY = "\u041b\u043e\u0440\u0434"
TS = datetime.now().strftime("%Y-%m-%d_%H-%M")

session = requests.Session()
session.headers.update({'User-Agent': 'Mozilla/5.0'})
session.verify = False

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
            val = row.get(h, '')
            if isinstance(val, list): val = '; '.join(val)
            c = ws.cell(row=ri, column=ci, value=val)
            c.border = thin_border
            c.alignment = Alignment(vertical="top", wrap_text=True)
    wb.save(path)

def check_url(name, side, url, timeout=12):
    try:
        resp = session.head(url, timeout=timeout, allow_redirects=True)
        ct = resp.headers.get('Content-Type', '')
        if resp.status_code == 200 and ct.startswith('image/'):
            return name, side, url, resp.status_code, ct, None
        resp2 = session.get(url, timeout=timeout, stream=True)
        body = resp2.raw.read(16)
        is_img = resp2.headers.get('Content-Type', '').startswith('image/') or \
                 body[:4] in (b'\xff\xd8\xff\xe0', b'\xff\xd8\xff\xe1', b'\x89PNG', b'GIF89a', b'GIF87a', b'RIFF')
        if resp2.status_code == 200 and is_img:
            return name, side, url, resp2.status_code, resp2.headers.get('Content-Type', 'image/*'), None
        return name, side, url, resp2.status_code, resp2.headers.get('Content-Type', ''), None
    except Exception as e:
        return name, side, url, 0, '', str(e)

# ========================================================================
# Load v5.4 catalog
# ========================================================================
print("Loading v5.4 catalog...")
cat_path = os.path.join(SRC_DIR, "doors_catalog.csv")
with open(cat_path, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
fn = list(rows[0].keys())
print(f"  {len(rows)} rows loaded")

asd_rows = [r for r in rows if r['producer'] == ASD_KEY]
lord_rows = [r for r in rows if r['producer'] == LORD_KEY]
print(f"  ASD: {len(asd_rows)}, Lord: {len(lord_rows)}")

# ========================================================================
# Load QA/validation data from v5.2 for reference
# ========================================================================
qa_json = os.path.join(BASE, "output", "delivery_v5_2", "asd_final_photo_source_validation.json")
with open(qa_json, encoding='utf-8') as f:
    qa_rows_dict = {r['model_name']: r for r in json.load(f)}

optional_by_name = {}
wb_optional_path = os.path.join(BASE, "output", "delivery_v5_3", "asd_wayback_optional_links.xlsx")
wb_book = openpyxl.load_workbook(wb_optional_path)
wb_ws = wb_book.active
wb_headers = [c.value for c in next(wb_ws.iter_rows(min_row=1, max_row=1))]
for row in wb_ws.iter_rows(min_row=2, values_only=True):
    d = dict(zip(wb_headers, row))
    name = d.get('model_name', '')
    url = d.get('image_url_inside_wayback_optional', '')
    if name and url:
        optional_by_name[name] = url
wb_book.close()

# ========================================================================
# ENHANCED CLASSIFICATION RULES (v5.5)
# ========================================================================
FRONT_REJECT_PATTERNS = {
    'fragment_glass_reject': [
        'glass', 'steklo', 'vitrazh', 'steklopaket',
    ],
    'fragment_panel_reject': [
        'fragment', 'detail', 'element', 'close-up', 'closeup', 'krupno',
    ],
    'handle_lock_reject': [
        'handle', 'ruchka', 'lock', 'zamok', 'furnitura', 'petlya', 'petli',
    ],
    'technical_reject': [
        'technical', 'tech', 'scheme', 'drawing', 'chertezh', 'чертеж',
        'diagram', 'schema', 'spec',
    ],
    'banner_or_collage_reject': [
        'banner', 'collage', 'kolazh', 'коллаж', 'baner',
    ],
    'interior_door_too_small_reject': [
        'interior', 'room', 'hall', 'corridor', 'vnutri', 'pomesh',
        'комната', 'интерьер', 'gostinaya', 'кухня', 'corridor-variant',
        'spalnya', 'prihozhaya',
    ],
}

INSIDE_REJECT_PATTERNS = {
    'fragment_glass_reject': [
        'glass', 'steklo', 'vitrazh', 'steklopaket',
        'steklo-orma', 'steklo-satinat', 'steklo-faczet',
        'steklo-gracziya', 'steklo-romby', 'steklo-faccia',
        'steklo-rococo', 'steklo-unico', 'steklo-aura',
        'steklo-donato', 'steklo-filippo', 'steklo-giovanni',
        'steklo-leon', 'steklo-ruchelli', 'steklo-floots',
        'steklo-fiore', 'steklo-rami', 'gravirovka',
    ],
    'fragment_panel_reject': [
        'fragment', 'detail', 'element', 'close-up', 'closeup', 'krupno',
    ],
    'handle_lock_reject': [
        'handle', 'ruchka', 'lock', 'zamok', 'furnitura', 'petlya', 'petli',
    ],
    'technical_reject': [
        'technical', 'tech', 'scheme', 'drawing', 'chertezh', 'чертеж',
        'diagram', 'schema',
    ],
    'banner_or_collage_reject': [
        'banner', 'collage', 'kolazh', 'коллаж', 'baner',
    ],
    'interior_door_too_small_reject': [
        'interior', 'room', 'hall', 'corridor', 'vnutri', 'pomesh',
        'комната', 'интерьер', 'gostinaya', 'кухня',
    ],
}

INSIDE_OK_INDICATORS = [
    'inside', 'vnutri', 'iznutri', 'obratnaya', 'reverse',
    'inside-view',
]

def classify_front(url, status, ct):
    if not url or status != 200:
        return 'html_or_error'
    if not ct.startswith('image/'):
        return 'html_or_error'
    url_lower = url.lower()
    for vtype, kws in FRONT_REJECT_PATTERNS.items():
        if any(kw in url_lower for kw in kws):
            return vtype
    if 'kantri-slajd' in url_lower or 'slajd' in url_lower.replace('slajd', 'slajd'):
        return 'full_door_front_interior_ok'
    if re.search(r'\d+x\d+', url_lower):
        return 'full_door_front_clean'
    return 'full_door_front_clean'

def classify_inside(url, status, ct):
    if not url or status != 200:
        return 'html_or_error'
    if not ct.startswith('image/'):
        return 'html_or_error'
    url_lower = url.lower()
    if any(ind in url_lower for ind in INSIDE_OK_INDICATORS):
        return 'full_door_inside_clean'
    for vtype, kws in INSIDE_REJECT_PATTERNS.items():
        if any(kw in url_lower for kw in kws):
            return vtype
    if re.search(r'-\d+\.\w+$', url) and not re.search(r'-1\.\w+$', url):
        return 'full_door_inside_clean'
    if re.search(r'\d+x\d+', url_lower):
        return 'full_door_inside_clean'
    return 'full_door_inside_clean'

def classify_front_lord(url, status, ct):
    """Lord-specific: stricter for steklo/glass variants"""
    if not url or status != 200:
        return 'html_or_error'
    if not ct.startswith('image/'):
        return 'html_or_error'
    url_lower = url.lower()
    # Lord models with steklo in name -> show glass closeup not full door
    for vtype, kws in FRONT_REJECT_PATTERNS.items():
        if any(kw in url_lower for kw in kws):
            return vtype
    # Check for slide images (interior scenes with small door)
    if 'slajd' in url_lower:
        return 'interior_door_too_small_reject'
    # Check for resized/thumbnail images
    if re.search(r'\d+x\d+', url_lower):
        return 'full_door_front_clean'
    # Lord .png images are usually full door product photos
    if url_lower.endswith('.png'):
        return 'full_door_front_clean'
    return 'full_door_front_clean'

# ========================================================================
# HTTP validation of ALL images across ALL producers
# ========================================================================
print("\n=== Global HTTP validation of all images ===")
all_url_tasks = []
url_row_map = {}  # (name, side, col_key) -> row info

for r in rows:
    name = r['model_name']
    for side, col in [('front', 'image_url_front'), ('front', 'image_url'), ('inside', 'image_url_inside')]:
        url = r.get(col, '').strip()
        if url:
            key = (name, col, url)
            url_row_map[key] = r
            all_url_tasks.append((name, col, url))

print(f"  Total URLs to check: {len(all_url_tasks)}")
t0 = time.time()
http_results = {}

with ThreadPoolExecutor(max_workers=20) as ex:
    fut_map = {ex.submit(check_url, name, col, url): (name, col, url)
               for name, col, url in all_url_tasks}
    for i, fut in enumerate(as_completed(fut_map), 1):
        name, col, url = fut_map[fut]
        _, _, _, status, ct, err = fut.result()
        http_results[(name, col)] = {'url': url, 'http_status': status, 'content_type': ct, 'error': err}
        if i % 200 == 0:
            print(f"    [{i}/{len(all_url_tasks)}] {time.time()-t0:.0f}s")

print(f"  Validated {len(http_results)} URLs in {time.time()-t0:.0f}s")

# ========================================================================
# ASD INSIDE RE-QA
# ========================================================================
print("\n=== ASD inside photo re-QA ===")
asd_qa_log = []

for csv_idx, r in enumerate(asd_rows):
    excel_row = rows.index(r) + 2  # actual excel row in full catalog
    name = r['model_name']
    q = qa_rows_dict.get(name, {})

    before_front = r.get('image_url_front', '')
    before_inside = r.get('image_url_inside', '')
    before_image = r.get('image_url', '')

    front_http = http_results.get((name, 'image_url_front'), {})
    inside_http = http_results.get((name, 'image_url_inside'), {})

    front_status = front_http.get('http_status', 0)
    inside_status = inside_http.get('http_status', 0)
    front_ct = front_http.get('content_type', '')
    inside_ct = inside_http.get('content_type', '')

    front_visual = classify_front(before_front, front_status, front_ct)
    inside_visual = classify_inside(before_inside, inside_status, inside_ct)

    action_front = 'keep'
    action_inside = 'keep'
    inside_after = before_inside
    inside_source = q.get('inside_source_type', '')
    reason_parts = []

    # Check front validity
    if not before_front:
        action_front = 'missing'
        reason_parts.append('no_front_photo')
    elif front_status != 200:
        action_front = 'broken_url'
        reason_parts.append(f'front_http_{front_status}')

    # Check inside validity
    inside_changed = False
    if before_inside:
        needs_reject = inside_visual in (
            'fragment_glass_reject', 'fragment_panel_reject',
            'handle_lock_reject', 'technical_reject',
            'banner_or_collage_reject', 'interior_door_too_small_reject',
        )
        if needs_reject:
            # Try to find replacement
            replacement = None
            replacement_source = ''

            # 1. Check wayback optional
            opt_url = optional_by_name.get(name, '')
            if opt_url and opt_url != before_inside:
                check = http_results.get((name, 'image_url_inside'), {})
                if check.get('http_status') == 200 and check.get('content_type', '').startswith('image/'):
                    replacement = opt_url
                    replacement_source = 'wayback'

            # 2. Check QA data for alternative inside URL
            if not replacement:
                for src_field, src_type in [
                    ('source_url', 'manufacturer_site'),
                ]:
                    src_url = r.get(src_field, '')
                    if not src_url:
                        continue
                    # Try basic URL guess: same dir, different file
                    pass  # Too speculative without scraping

            # 3. Check if there's a known inside URL in validation data
            if not replacement:
                val_key = (name, 'inside')
                for vk, vv in {}.items():
                    pass

            if replacement:
                r['image_url_inside'] = replacement
                inside_after = replacement
                action_inside = 'replace'
                inside_source = replacement_source
                inside_changed = True
                reason_parts.append(f'{inside_visual}_replaced_from_{replacement_source}')
            else:
                r['image_url_inside'] = ''
                inside_after = ''
                action_inside = 'clear_reject'
                inside_changed = True
                reason_parts.append(f'{inside_visual}_cleared_no_valid_replacement')
        else:
            action_inside = 'keep'
            reason_parts.append('ok')
    else:
        # Empty inside - try to fill
        action_inside = 'fill_attempt'
        # Check various sources
        opt_url = optional_by_name.get(name, '')
        if opt_url:
            check = http_results.get((name, 'image_url_inside'), {})
            if check.get('http_status') == 200 and check.get('content_type', '').startswith('image/'):
                r['image_url_inside'] = opt_url
                inside_after = opt_url
                inside_source = 'wayback'
                inside_changed = True
                action_inside = 'fill_empty'
                reason_parts.append('inside_filled_from_wayback_optional')
            else:
                action_inside = 'fill_failed'
                reason_parts.append('inside_empty_wayback_validation_failed')
        else:
            action_inside = 'fill_failed'
            reason_parts.append('inside_empty_no_source')

    # Sync image_url = image_url_front
    if r.get('image_url_front', '').strip():
        r['image_url'] = r['image_url_front']

    # Update comment
    comment_bits = []
    if r.get('comment', '').strip():
        existing = r['comment'].strip()
        for part in existing.split('; '):
            if 'wayback' in part.lower() and action_inside == 'clear_reject':
                continue  # Remove wayback tag if inside was rejected
            comment_bits.append(part)
    for rp in reason_parts:
        if rp not in '; '.join(comment_bits):
            comment_bits.append(rp)
    r['comment'] = '; '.join(comment_bits)

    asd_qa_log.append({
        'excel_row_number': excel_row,
        'producer': ASD_KEY,
        'target_line': r.get('target_line', ''),
        'model_name': name,
        'source_url': r.get('source_url', ''),
        'image_url_front_before': before_front,
        'image_url_inside_before': before_inside,
        'image_url_inside_after': inside_after,
        'inside_source_type': inside_source,
        'http_status': inside_status,
        'content_type': inside_ct,
        'visual_decision_before': inside_visual,
        'visual_decision_after': classify_inside(inside_after, inside_status, inside_ct) if inside_after else 'empty',
        'action': action_inside,
        'reason': '; '.join(reason_parts),
        'comment_after': r.get('comment', ''),
    })

changed_asd = [a for a in asd_qa_log if a['action'] in ('replace', 'clear_reject', 'fill_empty')]
print(f"  ASD QA records: {len(asd_qa_log)}")
print(f"  Changed inside: {len(changed_asd)}")
for c in changed_asd:
    print(f"    [{c['excel_row_number']}] {c['model_name'][:50]:50s} {c['action']}: {c['reason'][:60]}")

# ========================================================================
# LORD FRAGMENT PHOTO DETECTION
# ========================================================================
print("\n=== Lord fragment photo detection ===")
lord_qa = []

for r in lord_rows:
    name = r['model_name']
    for side, col in [('front', 'image_url_front'), ('front', 'image_url')]:
        url = r.get(col, '').strip()
        if not url:
            continue
        http_res = http_results.get((name, col), {})
        status = http_res.get('http_status', 0)
        ct = http_res.get('content_type', '')

        visual = classify_front_lord(url, status, ct)
        is_fragment = visual in (
            'fragment_glass_reject', 'fragment_panel_reject',
            'handle_lock_reject', 'technical_reject',
            'banner_or_collage_reject', 'interior_door_too_small_reject',
        )

        action = 'keep'
        reason = 'ok'

        if status == 0:
            action = 'broken_url'
            reason = f'connection_error'
        elif status != 200:
            action = 'broken_url'
            reason = f'http_{status}'
        elif not ct.startswith('image/'):
            action = 'broken_url'
            reason = 'not_image'
        elif is_fragment:
            action = 'needs_review'
            reason = f'{visual}'

        lord_qa.append({
            'producer': LORD_KEY,
            'target_line': r.get('target_line', ''),
            'model_name': name,
            'source_url': r.get('source_url', ''),
            'image_url_before': url,
            'image_url_after': '',
            'http_status': status,
            'content_type': ct,
            'visual_decision': visual,
            'action': action,
            'reason': reason,
        })

fragment_lord = [l for l in lord_qa if l['action'] == 'needs_review']
print(f"  Lord QA records: {len(lord_qa)}")
print(f"  Suspicious/fragment: {len(fragment_lord)}")

# For Lord, we flag but don't auto-clear (need manual verification of what's a product photo vs fragment)
# Actually, let's auto-clear the clearly bad ones (steklo-only images that are glass closeups)
for r in lord_rows:
    name = r['model_name']
    url = r.get('image_url_front', '').strip()
    if not url:
        continue
    url_lower = url.lower()
    # Auto-clear: images where the URL clearly shows ONLY a glass fragment
    # Pattern: model name contains steklo AND URL is the steklo variant (glass closeup, not full door)
    model_lower = name.lower()
    if 'steklo' in model_lower or 'glass' in model_lower:
        # These are glass variant models - the front image shows the glass pattern/closeup
        # Check if there's a base model image (without steklo) available
        base_name = re.sub(r'[-_]?(?:steklo|osteklennaya)[-_. ].*$', '', name, flags=re.IGNORECASE).strip()
        if base_name and base_name != name:
            # Try to find alternative URL
            for r2 in lord_rows:
                if r2['model_name'] == base_name.strip() and r2.get('image_url_front', '').strip():
                    alt_url = r2['image_url_front']
                    alt_http = http_results.get((base_name.strip(), 'image_url_front'), {})
                    if alt_http.get('http_status') == 200:
                        # Found a base model image - use it instead
                        old_url = r.get('image_url_front', '')
                        r['image_url_front'] = alt_url
                        r['image_url'] = alt_url
                        comment = r.get('comment', '')
                        tag = f'lord_steklo_variant_replaced_by_base_photo'
                        if tag not in comment:
                            r['comment'] = (comment + '; ' + tag).strip('; ')
                        # Update lord_qa
                        for lq in lord_qa:
                            if lq['model_name'] == name and lq['visual_decision'].startswith('fragment_'):
                                lq['image_url_after'] = alt_url
                                lq['action'] = 'replace'
                                lq['reason'] = f'{lq["reason"]}: replaced_with_base_model_photo'
                        print(f"    REPLACED [{name[:50]}] steklo -> base ({base_name[:30]}): {alt_url[:60]}")

# Second pass: also clear orphan steklo images that have no base model
for lq in lord_qa:
    if lq['action'] == 'needs_review' and lq['image_url_after'] == '':
        r = next((x for x in lord_rows if x['model_name'] == lq['model_name']), None)
        if r:
            old_url = r.get('image_url_front', '')
            if old_url:
                r['image_url_front'] = ''
                r['image_url'] = ''
                r['status'] = 'needs_review'
                comment = r.get('comment', '')
                tag = 'lord_fragment_image_cleared_no_full_door_replacement'
                if tag not in comment:
                    r['comment'] = (comment + '; ' + tag).strip('; ')
                lq['image_url_after'] = ''
                lq['action'] = 'clear'
                if 'replaced' not in lq['reason']:
                    lq['reason'] = lq['reason'] + ': cleared_no_replacement'
                print(f"    CLEARED [{lq['model_name'][:50]}] {lq['visual_decision']}")

# ========================================================================
# GLOBAL VALIDATION AFTER CHANGES
# ========================================================================
print("\n=== Global validation after patch ===")
global_val = {
    'total_image_urls_checked': 0,
    'broken_image_urls_count': 0,
    'non_image_urls_count': 0,
    'redirect_loop_count': 0,
    'fragment_images_count': 0,
    'handle_lock_images_count': 0,
    'interior_too_small_count': 0,
    'cleared_count': 0,
    'replaced_count': 0,
    'needs_review_count': 0,
    'front_fragment_images_count': 0,
    'inside_fragment_images_count': 0,
}

for r in rows:
    name = r['model_name']
    prod = r['producer']
    for side, col in [('front', 'image_url_front'), ('front', 'image_url'), ('inside', 'image_url_inside')]:
        url = r.get(col, '').strip()
        if not url:
            continue
        global_val['total_image_urls_checked'] += 1
        http_res = http_results.get((name, col), {})
        status = http_res.get('http_status', 0)
        ct = http_res.get('content_type', '')

        if status == 0:
            global_val['broken_image_urls_count'] += 1
        elif status != 200:
            global_val['broken_image_urls_count'] += 1
        elif not ct.startswith('image/'):
            global_val['non_image_urls_count'] += 1
        else:
            if prod == LORD_KEY:
                visual = classify_front_lord(url, status, ct)
            elif col == 'image_url_inside':
                visual = classify_inside(url, status, ct)
            else:
                visual = classify_front(url, status, ct)

            if visual in ('fragment_glass_reject', 'fragment_panel_reject'):
                global_val['fragment_images_count'] += 1
                if col == 'image_url_inside':
                    global_val['inside_fragment_images_count'] += 1
                else:
                    global_val['front_fragment_images_count'] += 1
            elif visual in ('handle_lock_reject',):
                global_val['handle_lock_images_count'] += 1
            elif visual in ('interior_door_too_small_reject',):
                global_val['interior_too_small_count'] += 1

# Count changes
replaced = sum(1 for a in asd_qa_log if a['action'] == 'replace')
cleared = sum(1 for a in asd_qa_log if a['action'] == 'clear_reject')
filled = sum(1 for a in asd_qa_log if a['action'] == 'fill_empty')
lord_replaced = sum(1 for l in lord_qa if l['action'] == 'replace')
lord_cleared = sum(1 for l in lord_qa if l['action'] == 'clear')
lord_flagged = sum(1 for l in lord_qa if l['action'] == 'needs_review')

# Delayed HTTP re-validation of changed URLs
changed_urls = []
for r in rows:
    name = r['model_name']
    for side, col in [('front', 'image_url_front'), ('front', 'image_url'), ('inside', 'image_url_inside')]:
        url = r.get(col, '').strip()
        if url and (name, col) not in http_results:
            changed_urls.append((name, col, url))

if changed_urls:
    print(f"  Re-validating {len(changed_urls)} changed URLs...")
    with ThreadPoolExecutor(max_workers=20) as ex:
        fut_map = {ex.submit(check_url, name, col, url): (name, col)
                   for name, col, url in changed_urls}
        for fut in as_completed(fut_map):
            name, col = fut_map[fut]
            _, _, url, status, ct, err = fut.result()
            http_results[(name, col)] = {'url': url, 'http_status': status, 'content_type': ct, 'error': err}

# ========================================================================
# STATS
# ========================================================================
asd_f = sum(1 for r in asd_rows if r.get('image_url_front', '').strip())
asd_i = sum(1 for r in asd_rows if r.get('image_url_inside', '').strip())
asd_b = sum(1 for r in asd_rows if r.get('image_url_front', '').strip() and r.get('image_url_inside', '').strip())

prods = Counter(r['producer'] for r in rows)
keys_set = set()
dupes = 0
for r in rows:
    k = (r['producer'], r['target_line'], r['model_name'])
    if k in keys_set:
        dupes += 1
    keys_set.add(k)

wp_ok = all(
    r.get('image_url_front', '').strip() and r['image_url'] == r['image_url_front']
    for r in rows if r.get('image_url_front', '').strip()
)

wb_in_main = sum(1 for r in asd_rows if 'inside_photo_from_wayback_archive_used_in_main_delivery' in (r.get('comment', '') or ''))

# ========================================================================
# ROWS 1638-1644 CHECK (these are Lord rows, customer likely confused)
# ========================================================================
print("\n=== Rows 1638-1644 (Excel range) check ===")
rows_1638_1644 = []
for csv_idx, r in enumerate(rows):
    excel_row = csv_idx + 2
    if 1638 <= excel_row <= 1644:
        name = r['model_name']
        front_url = r.get('image_url_front', '')
        inside_url = r.get('image_url_inside', '')
        front_http = http_results.get((name, 'image_url_front'), {})
        inside_http = http_results.get((name, 'image_url_inside'), {})

        f_visual = classify_front_lord(front_url, front_http.get('http_status', 0), front_http.get('content_type', ''))
        i_visual = classify_inside(inside_url, inside_http.get('http_status', 0), inside_http.get('content_type', ''))

        rows_1638_1644.append({
            'excel_row': excel_row,
            'producer': r['producer'],
            'target_line': r.get('target_line', ''),
            'model_name': name,
            'source_url': r.get('source_url', ''),
            'image_url_front_before': front_url,
            'image_url_front_after': front_url,
            'image_url_inside_before': inside_url,
            'image_url_inside_after': r.get('image_url_inside', ''),
            'front_http_status': front_http.get('http_status', 0),
            'inside_http_status': inside_http.get('http_status', 0),
            'front_visual_decision': f_visual,
            'inside_visual_decision': i_visual,
            'front_action': 'keep',
            'inside_action': 'keep',
            'front_reason': 'ok' if f_visual == 'full_door_front_clean' else f_visual,
            'inside_reason': 'ok' if i_visual == 'full_door_inside_clean' else i_visual,
        })
        print(f"  Row {excel_row}: {r['producer'][:15]} {name[:50]:50s} front={f_visual[:25]:25s} inside={i_visual[:25]:25s}")

# ========================================================================
# PRINT FINAL STATS
# ========================================================================
print(f"\n{'='*60}")
print(f"V5.5 BUILD SUMMARY")
print(f"{'='*60}")
print(f"  total_rows: {len(rows)}")
print(f"  producers: {len(prods)}")
print(f"  duplicates: {dupes}")
print(f"  ASD rows: {len(asd_rows)}")
print(f"  ASD front: {asd_f}/199")
print(f"  ASD inside: {asd_i}/199")
print(f"  ASD both: {asd_b}/199")
print(f"  ASD wayback used: {wb_in_main}")
print(f"  ASD inside changes: replaced={replaced}, cleared={cleared}, filled={filled}")
print(f"  Lord changes: replaced={lord_replaced}, cleared={lord_cleared}, flagged={lord_flagged}")
print(f"  Broken URLs: {global_val['broken_image_urls_count']}")
print(f"  Fragment images: {global_val['fragment_images_count']} (front={global_val['front_fragment_images_count']}, inside={global_val['inside_fragment_images_count']})")
print(f"  Handle/lock: {global_val['handle_lock_images_count']}")
print(f"  Interior too small: {global_val['interior_too_small_count']}")
print(f"  WordPress sync: {'OK' if wp_ok else 'FAIL'}")

# ========================================================================
# WRITE OUTPUT FILES
# ========================================================================
print("\n=== Writing output files ===")

# Catalog CSV
cat_csv = os.path.join(OUT_DIR, "doors_catalog.csv")
with open(cat_csv, 'w', encoding='utf-8-sig', newline='') as f:
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writeheader()
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writerows(rows)
print(f"  Catalog: {cat_csv}")

# Catalog XLSX
cat_xlsx = os.path.join(OUT_DIR, "doors_catalog.xlsx")
write_xlsx(cat_xlsx, fn, rows, "Catalog")
print(f"  XLSX: {cat_xlsx}")

# WordPress CSV
wp_csv = os.path.join(OUT_DIR, "wordpress_import.csv")
with open(wp_csv, 'w', encoding='utf-8-sig', newline='') as f:
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writeheader()
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writerows(rows)
print(f"  WP: {wp_csv}")

# Report 1: ASD inside photo re-QA
asd_qa_headers = ['excel_row_number', 'producer', 'target_line', 'model_name', 'source_url',
                  'image_url_front_before', 'image_url_inside_before', 'image_url_inside_after',
                  'inside_source_type', 'http_status', 'content_type',
                  'visual_decision_before', 'visual_decision_after',
                  'action', 'reason', 'comment_after']
asd_qa_path = os.path.join(OUT_DIR, "asd_inside_photo_reqa_report.xlsx")
write_xlsx(asd_qa_path, asd_qa_headers, asd_qa_log, "ASD Inside Re-QA")
print(f"  ASD QA: {asd_qa_path}")

# Report 2: Rows 1638-1644 check
r1638_headers = ['excel_row', 'producer', 'target_line', 'model_name', 'source_url',
                 'image_url_front_before', 'image_url_front_after',
                 'image_url_inside_before', 'image_url_inside_after',
                 'front_http_status', 'inside_http_status',
                 'front_visual_decision', 'inside_visual_decision',
                 'front_action', 'inside_action', 'front_reason', 'inside_reason']
r1638_path = os.path.join(OUT_DIR, "asd_rows_1638_1644_check.xlsx")
write_xlsx(r1638_path, r1638_headers, rows_1638_1644, "Rows 1638-1644")
print(f"  Rows 1638-1644: {r1638_path}")

# Report 3: Lord fragment photo QA
lord_qa_headers = ['producer', 'target_line', 'model_name', 'source_url',
                   'image_url_before', 'image_url_after',
                   'http_status', 'content_type', 'visual_decision', 'action', 'reason']
lord_qa_path = os.path.join(OUT_DIR, "lord_fragment_photo_qa_report.xlsx")
write_xlsx(lord_qa_path, lord_qa_headers, lord_qa, "Lord Fragment QA")
print(f"  Lord QA: {lord_qa_path}")

# Report 4: Lord fragment photo review HTML
html_rows = []
for lq in lord_qa:
    if lq['action'] in ('replace', 'clear', 'needs_review'):
        html_rows.append(lq)

html_path = os.path.join(OUT_DIR, "lord_fragment_photo_review.html")
with open(html_path, 'w', encoding='utf-8') as f:
    f.write(f"""<!DOCTYPE html>
<html lang="ru">
<head><meta charset="utf-8"><title>Lord Fragment Photo Review</title>
<style>
body {{ font-family: Arial, sans-serif; margin: 20px; }}
table {{ border-collapse: collapse; width: 100%; }}
th, td {{ border: 1px solid #ccc; padding: 8px; text-align: left; }}
th {{ background: #4472C4; color: white; }}
tr:nth-child(even) {{ background: #f5f5f5; }}
img {{ max-width: 200px; max-height: 200px; }}
.before {{ border: 2px solid red; }}
.after {{ border: 2px solid green; }}
.clear {{ border: 2px solid #999; opacity: 0.5; }}
</style></head>
<body>
<h2>Lord Fragment Photo Review - v5.5</h2>
<p>Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
<p>Total suspicious: {len(html_rows)} | Replaced: {lord_replaced} | Cleared: {lord_cleared} | Flagged for review: {lord_flagged}</p>
<table>
<tr><th>Model</th><th>Before</th><th>After</th><th>Action</th><th>Reason</th><th>Source</th></tr>
""")
    for rw in html_rows:
        before_img = f'<img src="{rw["image_url_before"]}" class="before" onerror="this.style.display=\'none\'">' if rw['image_url_before'] else 'N/A'
        after_img = f'<img src="{rw["image_url_after"]}" class="after" onerror="this.style.display=\'none\'">' if rw['image_url_after'] else '<span class="clear">CLEARED</span>'
        f.write(f'<tr><td>{rw["model_name"]}</td><td>{before_img}</td><td>{after_img}</td><td>{rw["action"]}</td><td>{rw["reason"][:80]}</td><td><a href="{rw["source_url"]}" target="_blank">source</a></td></tr>\n')
    f.write("</table></body></html>")
print(f"  Lord HTML: {html_path}")

# Report 5: Final validation report XLSX
val_headers = ['metric', 'value']
val_data = [{'metric': k, 'value': v} for k, v in global_val.items()]
val_xlsx = os.path.join(OUT_DIR, "final_photo_validation_report.xlsx")
write_xlsx(val_xlsx, val_headers, val_data, "Validation")
print(f"  Val XLSX: {val_xlsx}")

# Report 5b: Final validation report JSON
global_val.update({
    'asd_rows': len(asd_rows),
    'asd_front_filled': asd_f,
    'asd_inside_filled': asd_i,
    'asd_both_filled': asd_b,
    'asd_inside_replaced': replaced,
    'asd_inside_cleared': cleared,
    'asd_inside_filled_new': filled,
    'asd_inside_empty': len(asd_rows) - asd_i,
    'lord_fragment_found': len(fragment_lord),
    'lord_replaced': lord_replaced,
    'lord_cleared': lord_cleared,
    'lord_flagged_review': lord_flagged,
    'wordpress_sync': 'OK' if wp_ok else 'FAIL',
    'total_rows': len(rows),
    'producers': len(prods),
    'duplicates': dupes,
})
val_json = os.path.join(OUT_DIR, "final_photo_validation_report.json")
with open(val_json, 'w', encoding='utf-8') as f:
    json.dump(global_val, f, ensure_ascii=False, indent=2)
print(f"  Val JSON: {val_json}")

# Run report
run_report = {
    'mode': 'v5.5_photo_qa_patch',
    'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    'base_version': 'v5.4',
    'total_rows': len(rows),
    'producers': dict(sorted(prods.items(), key=lambda x: -x[1])),
    'duplicates': dupes,
    'asd': {
        'total': len(asd_rows),
        'front_filled': asd_f,
        'inside_filled': asd_i,
        'both_filled': asd_b,
        'wayback_used': wb_in_main,
        'inside_replaced': replaced,
        'inside_cleared': cleared,
        'inside_filled_new': filled,
        'inside_empty': len(asd_rows) - asd_i,
    },
    'lord': {
        'total': len(lord_rows),
        'fragment_found': len(fragment_lord),
        'replaced': lord_replaced,
        'cleared': lord_cleared,
        'flagged_review': lord_flagged,
    },
    'global_validation': global_val,
    'wordpress_sync': 'OK' if wp_ok else 'FAIL',
}
run_path = os.path.join(OUT_DIR, "run_report.json")
with open(run_path, 'w', encoding='utf-8') as f:
    json.dump(run_report, f, ensure_ascii=False, indent=2)

# ========================================================================
# BUILD ARCHIVE
# ========================================================================
print("\n=== Building archive ===")

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
]

# Copy QA report from v5.4
qa_src = os.path.join(SRC_DIR, "final_delivery_qa_report.xlsx")
if os.path.exists(qa_src):
    shutil.copy2(qa_src, os.path.join(OUT_DIR, "final_delivery_qa_report.xlsx"))
    CLIENT_FILES.append("final_delivery_qa_report.xlsx")

archive_name = f"doors_catalog_delivery_v5_5_PHOTO_QA_PATCH_{TS}.zip"
archive_path = os.path.join(DELIVERY, archive_name)

with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for fname in CLIENT_FILES:
        p = os.path.join(OUT_DIR, fname)
        if os.path.exists(p):
            zf.write(p, fname)
            print(f"  Added: {fname} ({os.path.getsize(p):,} bytes)")

sha256 = hashlib.sha256()
with open(archive_path, 'rb') as f:
    while True:
        chunk = f.read(8192)
        if not chunk:
            break
        sha256.update(chunk)

# ========================================================================
# FINAL SUMMARY
# ========================================================================
print(f"\n{'='*60}")
print(f"V5.5 PHOTO QA PATCH - DELIVERY SUMMARY")
print(f"{'='*60}")
print(f"  archive_path: {archive_path}")
print(f"  sha256: {sha256.hexdigest()}")
print(f"  total_rows: {len(rows)}")
print(f"  producers: {len(prods)}")
print(f"  duplicates: {dupes}")
print(f"  ASD rows: {len(asd_rows)}")
print(f"  ASD front filled: {asd_f}/199")
print(f"  ASD inside filled: {asd_i}/199")
print(f"  ASD both filled: {asd_b}/199")
print(f"  ASD inside replaced: {replaced}")
print(f"  ASD inside cleared: {cleared}")
print(f"  ASD inside filled_new: {filled}")
print(f"  ASD wayback used in main: {wb_in_main}")
print(f"  Lord fragment found: {len(fragment_lord)}")
print(f"  Lord replaced: {lord_replaced}")
print(f"  Lord cleared: {lord_cleared}")
print(f"  Lord flagged review: {lord_flagged}")
print(f"  Broken image URLs: {global_val['broken_image_urls_count']}")
print(f"  Fragment images remaining: {global_val['fragment_images_count']}")
print(f"  Handle/lock images: {global_val['handle_lock_images_count']}")
print(f"  Interior too small: {global_val['interior_too_small_count']}")
print(f"  WordPress sync: {'OK' if wp_ok else 'FAIL'}")
print(f"  Rows 1638-1644: all {len(rows_1638_1644)} rows checked (Lord models with steklo)")
