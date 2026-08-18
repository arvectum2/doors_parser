"""V5.6-1: Analyze current ASD front/inside state"""
import csv, json, os, re
from collections import Counter

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
V55_DIR = os.path.join(BASE, "output", "delivery_v5_5")
OUT_DIR = os.path.join(BASE, "output", "delivery_v5_6")
os.makedirs(OUT_DIR, exist_ok=True)

ASD = chr(0x0410)+chr(0x0421)+chr(0x0414)

with open(os.path.join(V55_DIR, "doors_catalog.csv"), encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
fn = list(rows[0].keys())

asd_rows = [r for r in rows if r['producer'] == ASD]
print(f"Total rows: {len(rows)}")
print(f"ASD rows: {len(asd_rows)}")

# Load v5.2 QA data for gallery info
qa_json = os.path.join(BASE, "output", "delivery_v5_2", "asd_final_photo_source_validation.json")
with open(qa_json, encoding='utf-8') as f:
    qa_data = json.load(f)
qa_by_name = {r['model_name']: r for r in qa_data}

# Analyze each ASD row
issues = {
    'same_url_front_inside': [],
    'front_possible_inside': [],
    'inside_possible_front': [],
    'inside_empty': [],
    'front_empty': [],
    'both_empty': [],
    'likely_ok': [],
}

for r in asd_rows:
    name = r['model_name']
    excel_row = rows.index(r) + 2
    f_url = r.get('image_url_front', '').strip()
    i_url = r.get('image_url_inside', '').strip()
    q = qa_by_name.get(name, {})
    
    # Check 1: Same URL in front and inside
    if f_url and i_url and f_url == i_url:
        issues['same_url_front_inside'].append({
            'excel_row': excel_row,
            'name': name,
            'front': f_url[:80],
            'inside': i_url[:80],
        })
    
    # Check 2: URL-side indicators
    f_lower = f_url.lower() if f_url else ''
    i_lower = i_url.lower() if i_url else ''
    
    # Inside indicators in front URL
    inside_indicators_in_front = []
    for kw in ['vnutri', 'iznutri', 'obratnaya', 'inside-view', 'reverse', 'panel-inside']:
        if kw in f_lower:
            inside_indicators_in_front.append(kw)
    if inside_indicators_in_front:
        issues['front_possible_inside'].append({
            'excel_row': excel_row,
            'name': name,
            'url': f_url[:80],
            'indicators': ','.join(inside_indicators_in_front),
            'source': q.get('front_source_type', ''),
        })
    
    # Front indicators in inside URL
    front_indicators_in_inside = []
    for kw in ['vnesh', 'naruzh', 'front', 'exterior', 'outside', 'metal']:
        if kw in i_lower:
            front_indicators_in_inside.append(kw)
    if front_indicators_in_inside:
        issues['inside_possible_front'].append({
            'excel_row': excel_row,
            'name': name,
            'url': i_url[:80],
            'indicators': ','.join(front_indicators_in_inside),
            'source': q.get('inside_source_type', ''),
        })
    
    if not i_url:
        issues['inside_empty'].append({'excel_row': excel_row, 'name': name})
    if not f_url:
        issues['front_empty'].append({'excel_row': excel_row, 'name': name})
    if not f_url and not i_url:
        issues['both_empty'].append({'excel_row': excel_row, 'name': name})
    if f_url and i_url and f_url != i_url and not inside_indicators_in_front and not front_indicators_in_inside:
        issues['likely_ok'].append({'excel_row': excel_row, 'name': name})

print(f"\n=== ASD ISSUES ANALYSIS ===")
print(f"Same URL front==inside: {len(issues['same_url_front_inside'])}")
for s in issues['same_url_front_inside']:
    print(f"  Row {s['excel_row']}: {s['name'][:50]}")

print(f"\nFront might be inside: {len(issues['front_possible_inside'])}")
for s in issues['front_possible_inside']:
    print(f"  Row {s['excel_row']}: {s['name'][:50]} url={s['url'][:60]} indicators={s['indicators']}")

print(f"\nInside might be front: {len(issues['inside_possible_front'])}")
for s in issues['inside_possible_front']:
    print(f"  Row {s['excel_row']}: {s['name'][:50]} url={s['url'][:60]} indicators={s['indicators']}")

print(f"\nInside empty: {len(issues['inside_empty'])}")
print(f"Front empty: {len(issues['front_empty'])}")
print(f"Both empty: {len(issues['both_empty'])}")
print(f"Likely OK: {len(issues['likely_ok'])}")

# Check source types for URLs
print(f"\n=== SOURCE TYPES ===")
st = Counter(q.get('front_source_type', '') for q in qa_data)
print(f"Front source types: {dict(st)}")
st2 = Counter(q.get('inside_source_type', '') for q in qa_data)
print(f"Inside source types: {dict(st2)}")

# Print all current front/inside URLs for review
print(f"\n=== ALL ASD CURRENT URLS ===")
for r in asd_rows:
    name = r['model_name']
    excel_row = rows.index(r) + 2
    f = r.get('image_url_front', '')[:50]
    i = r.get('image_url_inside', '')[:50]
    q = qa_by_name.get(name, {})
    fs = q.get('front_source_type', '')
    is_ = q.get('inside_source_type', '')
    if f or i:
        print(f"  Row {excel_row:4d} | {name[:40]:40s} | F={fs[:15]:15s} {f[:40]:40s} | I={is_[:15]:15s} {i[:40]:40s}")
