"""V5.5-1: Analyze ASD rows 1638-1644 and Lord fragment photos"""
import csv, json, os, re, openpyxl
from collections import Counter

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CAT = os.path.join(BASE, "output", "delivery_v5_4", "doors_catalog.csv")
OUT_DIR = os.path.join(BASE, "output", "delivery_v5_5")
os.makedirs(OUT_DIR, exist_ok=True)

ASD_KEY = "\u0410\u0421\u0414"
LORD_KEY = "\u041b\u043e\u0440\u0434"

with open(CAT, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
fn = list(rows[0].keys())

print(f"Total rows: {len(rows)}")

# === ASD rows 1638-1644 (Excel rows = CSV index + 2) ===
print("\n=== ASD ROWS 1638-1644 (Excel range) ===")
for csv_idx, r in enumerate(rows):
    excel_row = csv_idx + 2  # header=row1, data starts row2
    if 1638 <= excel_row <= 1644:
        is_asd = r['producer'] == ASD_KEY
        print(f"  Excel row {excel_row}: producer={r['producer'][:20]} model={r['model_name'][:50]}")
        if is_asd:
            print(f"    front={r.get('image_url_front','')[:80]}")
            print(f"    inside={r.get('image_url_inside','')[:80]}")
            print(f"    comment={r.get('comment','')[:80]}")

# Also find ALL ASD rows and their excel positions
print("\n=== ALL ASD ROWS WITH EXCEL POSITIONS ===")
asd_positions = []
for csv_idx, r in enumerate(rows):
    if r['producer'] == ASD_KEY:
        excel_row = csv_idx + 2
        has_f = bool(r.get('image_url_front', '').strip())
        has_i = bool(r.get('image_url_inside', '').strip())
        asd_positions.append((excel_row, r['model_name'], has_f, has_i, r.get('image_url_inside', '')[:60]))

print(f"Total ASD rows: {len(asd_positions)}")
print(f"ASD excel range: {asd_positions[0][0]} - {asd_positions[-1][0]}")
missing_inside = [(er, n, i) for er, n, f, i, _ in asd_positions if not i]
print(f"ASD rows with empty inside: {len(missing_inside)}")
for er, n, i in missing_inside[:20]:
    print(f"  Excel {er}: {n[:50]}")

# Also show ASD rows near 1638-1644
print("\n=== ASD rows near 1638-1644 ===")
for er, name, has_f, has_i, i_url in asd_positions:
    if 1635 <= er <= 1650:
        print(f"  Excel {er}: {name[:50]} front={'Y' if has_f else 'N'} inside={'Y' if has_i else 'N'} url={i_url}")

# === Лорд: find fragment patterns ===
print("\n=== LORD PHOTOS - fragment patterns ===")
lord_rows = [r for r in rows if r['producer'] == LORD_KEY]
print(f"Lord rows: {len(lord_rows)}")

# Fragment keywords
fragment_kws = ['glass', 'steklo', 'vitrazh', 'fragment', 'detail', 'handle', 'lock', 'ruchka', 'zamok',
                'close-up', 'closeup', 'furnitura', 'ruchka-krupno']

suspicious_lord = []
for r in lord_rows:
    name = r['model_name']
    for side, col in [('front', 'image_url_front'), ('inside', 'image_url_inside')]:
        url = r.get(col, '').strip()
        if not url: continue
        url_lower = url.lower()
        reason = []
        for kw in fragment_kws:
            if kw in url_lower:
                reason.append(kw)
        if reason:
            suspicious_lord.append({
                'name': name,
                'side': side,
                'url': url[:100],
                'reason': ','.join(reason),
                'target_line': r.get('target_line', ''),
            })
            print(f"  [{side}] {name[:50]:50s} {url[:80]:80s} kw={','.join(reason)}")

print(f"\nSuspicious Lord images: {len(suspicious_lord)}")

# Check Lord images more broadly: look for resize/cache patterns that might indicate fragments
print("\n=== LORD ALL IMAGES ===")
for r in lord_rows:
    name = r['model_name']
    for side, col in [('front', 'image_url_front'), ('inside', 'image_url_inside')]:
        url = r.get(col, '').strip()
        if url:
            # Check for cache_image patterns (resized/dimension-limited images)
            is_cache = 'cache_image' in url.lower()
            print(f"  [{side}] {name[:45]:45s} cache={'Y' if is_cache else 'N'} {url[:80]}")

# Count lord files with/without images
lord_front = sum(1 for r in lord_rows if r.get('image_url_front', '').strip())
lord_inside = sum(1 for r in lord_rows if r.get('image_url_inside', '').strip())
print(f"\nLord photos: front={lord_front}/{len(lord_rows)}, inside={lord_inside}/{len(lord_rows)}")
