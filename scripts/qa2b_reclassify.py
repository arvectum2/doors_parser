"""Fix classification: separate pattern_matched from manufacturer_site"""
import csv, json, os, re

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CAT = os.path.join(BASE, "output", "delivery_v5_2", "doors_catalog.csv")
VALID = os.path.join(BASE, "output", "delivery_v5_2", "asd_url_validation_results.json")
WB_JSON = os.path.join(BASE, "output", "delivery_v5_2", "wayback_inside_all_results.json")
OUT = os.path.join(BASE, "output", "delivery_v5_2", "asd_final_photo_source_validation.json")

ASD_KEY = "\u0410\u0421\u0414"

with open(CAT, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
asd = [r for r in rows if r['producer'] == ASD_KEY]

with open(VALID, encoding='utf-8') as f:
    val_raw = json.load(f)
validation = {}
for k, v in val_raw.items():
    parts = k.split('||')
    if len(parts) == 2:
        validation[(parts[0], parts[1])] = v

# Load wayback URLs
with open(WB_JSON, encoding='utf-8') as f:
    wb_raw = json.load(f)
wb_urls = set()
for v in wb_raw.values():
    u = v.get('inside_url','').strip().replace('https://','http://')
    if u: wb_urls.add(u)

# Patch URLs
V2_PATCH = os.path.join(BASE, "output", "asd_photo_patch_v2_clean", "asd_photo_patch_v2_clean.csv")
patch_urls = set()
with open(V2_PATCH, encoding='utf-8-sig') as f:
    for r in csv.DictReader(f, delimiter=';'):
        u = r.get('selected_image_url','').strip().replace('https://','http://')
        if u: patch_urls.add(u)

def is_pattern_matched(front_url, inside_url, wb_urls):
    """Check if inside_url is a pattern-generated variant of front_url"""
    if not front_url or not inside_url:
        return False
    if inside_url in wb_urls:
        return False
    
    # Normalize
    f = front_url.replace('https://','http://')
    i = inside_url.replace('https://','http://')
    
    # Same domain?
    if f.split('/')[2] != i.split('/')[2]:
        return False
    
    # Different file name but same base path?
    f_parts = f.rsplit('.', 1)
    i_parts = i.rsplit('.', 1)
    if len(f_parts) != 2 or len(i_parts) != 2:
        return False
    
    f_base = f_parts[0]
    i_base = i_parts[0]
    f_ext = f_parts[1]
    i_ext = i_parts[1]
    
    if f_ext != i_ext:
        return False
    
    # Pattern: front is model.ext, inside is model-N.ext or model-N.ext
    # Strip trailing numbers from front base
    f_clean = re.sub(r'-?\d+$', '', f_base)
    i_clean = re.sub(r'-?\d+$', '', i_base)
    
    if f_clean == i_clean or (i_base.startswith(f_clean) and i_base != f_base):
        return True
    
    # Also check: front is model-N.ext, inside is model-M.ext (different numbers)
    f_num = re.search(r'-(\d+)$', f_base)
    i_num = re.search(r'-(\d+)$', i_base)
    if f_num and i_num:
        f_prefix = f_base[:f_base.rfind('-')]
        i_prefix = i_base[:i_base.rfind('-')]
        if f_prefix == i_prefix and f_num.group(1) != i_num.group(1):
            return True
    
    return False

def classify_source(url, side, front_url, wb_set, patch_set):
    if not url:
        return 'empty'
    url_norm = url.replace('https://', 'http://')
    
    if 'web.archive.org' in url_norm:
        return 'wayback'
    if url_norm in wb_set:
        return 'wayback'
    if url_norm in patch_set:
        return 'v2_patch'
    
    # For inside URLs: check if pattern-matched from front
    if side == 'inside' and front_url:
        if is_pattern_matched(front_url, url_norm, wb_set):
            return 'pattern_matched'
    
    if 'as-doors.ru' in url_norm:
        return 'manufacturer_site'
    
    return 'pattern_matched'  # fallback

def classify_visual(url, status, ct, side='front'):
    if not url or status != 200:
        return 'unknown'
    if not ct.startswith('image/'):
        return 'unknown'
    url_lower = url.lower()
    
    if any(kw in url_lower for kw in ['handle', 'ruchka']):
        return 'handle'
    if any(kw in url_lower for kw in ['lock', 'zamok']):
        return 'lock'
    if any(kw in url_lower for kw in ['fragment', 'detail', 'element', 'close-up', 'closeup']):
        return 'fragment'
    if any(kw in url_lower for kw in ['interior', 'room', 'hall', 'corridor', 'vnutri', 'pomesh']):
        return 'interior_room'
    if any(kw in url_lower for kw in ['technical', 'tech', 'scheme', 'drawing', 'chertezh']):
        return 'technical'
    if any(kw in url_lower for kw in ['badge', 'logo', 'brand']):
        return 'badge'
    if any(kw in url_lower for kw in ['inside', 'vnutri']):
        return 'full_door_inside'
    if side == 'front':
        return 'full_door_front'
    return 'full_door_inside'

# Rebuild QA data
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
    
    f_source = classify_source(f_url, 'front', None, wb_urls, patch_urls)
    i_source = classify_source(i_url, 'inside', f_url, wb_urls, patch_urls)
    
    f_visual = classify_visual(f_url, f_status, f_ct, 'front')
    i_visual = classify_visual(i_url, i_status, i_ct, 'inside')
    
    f_opens = 'yes' if (f_status == 200 and f_ct.startswith('image/')) else 'no'
    i_opens = 'yes' if (i_status == 200 and i_ct.startswith('image/')) else 'no'
    
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
        reasons.append('front_not_image')
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
    elif i_source == 'wayback':
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
    })

# Stats
from collections import Counter
src_front = Counter(q['front_source_type'] for q in qa_rows)
src_inside = Counter(q['inside_source_type'] for q in qa_rows)
action_cnt = Counter()
for q in qa_rows:
    for a in q['action'].split('; '):
        action_cnt[a] += 1

print(f"Front sources: {dict(src_front)}")
print(f"Inside sources: {dict(src_inside)}")

# Separate pattern-matched
pm = [q for q in qa_rows if q['inside_source_type'] == 'pattern_matched']
wb = [q for q in qa_rows if q['inside_source_type'] == 'wayback']
print(f"\nPattern-matched inside: {len(pm)}")
print(f"Wayback inside: {len(wb)}")
print(f"Manufacturer site inside: {src_inside.get('manufacturer_site', 0)}")

# Save re-classified JSON
with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(qa_rows, f, ensure_ascii=False, indent=2)
print(f"\nSaved: {OUT}")
