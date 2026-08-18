"""V5.3-1: Analyze current state - find invalid URLs, technical images, wayback usage"""
import csv, json, os
from collections import Counter

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
OUT_DIR = os.path.join(BASE, "output", "delivery_v5_2")
CAT = os.path.join(OUT_DIR, "doors_catalog.csv")
QA_JSON = os.path.join(OUT_DIR, "asd_final_photo_source_validation.json")

ASD_KEY = "\u0410\u0421\u0414"

with open(CAT, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
fn = list(rows[0].keys())
asd = [r for r in rows if r['producer'] == ASD_KEY]

with open(QA_JSON, encoding='utf-8') as f:
    qa_rows = json.load(f)
qa_by_name = {r['model_name']: r for r in qa_rows}

print("=" * 60)
print("V5.3 STATE ANALYSIS")
print("=" * 60)

# 1. Invalid image URLs (non-200 or non-image that are still in catalog)
print("\n1. INVALID IMAGE URLs (non-200 / non-image still in catalog):")
invalid = []
for q in qa_rows:
    name = q['model_name']
    r = next((x for x in asd if x['model_name'] == name), None)
    if not r: continue
    
    for side, col, src_key in [('front', 'image_url_front', 'front_source_type'),
                                ('inside', 'image_url_inside', 'inside_source_type')]:
        url = q.get(f'image_url_{side}', '')
        if not url: continue
        # Check if still in catalog
        cat_url = r.get(col, '')
        if not cat_url: continue
        
        opens = q.get(f'{side}_opens_as_image', '') == 'yes'
        visual = q.get(f'{side}_visual_decision', 'unknown')
        status = q.get(f'{side}_http_status', 0)
        ct = q.get(f'{side}_content_type', '')
        src_type = q.get(src_key, '')
        
        if not opens:
            invalid.append({
                'name': name, 'side': side, 'url': url,
                'status': status, 'ct': ct, 'source': src_type,
                'visual': visual
            })
            print(f"  [{side}] {name[:50]:50s} status={status} ct={ct[:20]:20s} source={src_type}")

print(f"\n  Total invalid: {len(invalid)}")

# 2. Technical/fragment images
print("\n2. TECHNICAL/FRAGMENT IMAGES:")
tech_images = []
for q in qa_rows:
    name = q['model_name']
    r = next((x for x in asd if x['model_name'] == name), None)
    if not r: continue
    
    for side in ('front', 'inside'):
        url = q.get(f'image_url_{side}', '')
        if not url: continue
        cat_url = r.get(f'image_url_{side}', '')
        if not cat_url: continue
        
        visual = q.get(f'{side}_visual_decision', '')
        if visual in ('handle', 'lock', 'fragment', 'technical', 'badge', 'interior_room'):
            tech_images.append({
                'name': name, 'side': side, 'url': url,
                'visual': visual, 'source': q.get(f'{side}_source_type', '')
            })
            print(f"  [{side}] {name[:50]:50s} visual={visual}")

print(f"\n  Total technical/fragment: {len(tech_images)}")

# 3. Wayback URLs in main catalog
print("\n3. WAYBACK URLs IN MAIN CATALOG:")
wayback_main = []
for r in asd:
    name = r['model_name']
    q = qa_by_name.get(name, {})
    
    for side, col in [('front', 'image_url_front'), ('inside', 'image_url_inside')]:
        url = r.get(col, '').strip()
        if not url: continue
        src = q.get(f'{side}_source_type', '')
        if src == 'wayback':
            wayback_main.append({
                'name': name, 'side': side, 'url': url,
                'comment': r.get('comment', '')
            })
            print(f"  [{side}] {name[:50]:50s} comment={r.get('comment','')[:60]}")

print(f"\n  Total wayback in main: {len(wayback_main)}")

# 4. Summary of current state
f_filled = sum(1 for r in asd if r.get('image_url_front','').strip())
i_filled = sum(1 for r in asd if r.get('image_url_inside','').strip())
b_filled = sum(1 for r in asd if r.get('image_url_front','').strip() and r.get('image_url_inside','').strip())

print(f"\n4. CURRENT STATS:")
print(f"  ASD front filled: {f_filled}/199")
print(f"  ASD inside filled: {i_filled}/199")
print(f"  ASD both filled: {b_filled}/199")

# Count inside by source
inside_sources = Counter()
for r in asd:
    name = r['model_name']
    q = qa_by_name.get(name, {})
    if r.get('image_url_inside','').strip():
        inside_sources[q.get('inside_source_type','unknown')] += 1
print(f"  Inside by source: {dict(inside_sources)}")

# 5. Check all other producers preserved
prods = Counter(r['producer'] for r in rows)
print(f"\n5. ALL PRODUCERS:")
for p, c in sorted(prods.items(), key=lambda x: -x[1]):
    print(f"  {p}: {c}")

# 6. Check non-ASD for issues
print(f"\n6. NON-ASD CHECKS:")
for p in [k for k in prods.keys() if k != ASD_KEY]:
    prod_rows = [r for r in rows if r['producer'] == p]
    f = sum(1 for r in prod_rows if r.get('image_url_front','').strip())
    i = sum(1 for r in prod_rows if r.get('image_url_inside','').strip())
    print(f"  {p}: {len(prod_rows)} rows, front={f}, inside={i}")
