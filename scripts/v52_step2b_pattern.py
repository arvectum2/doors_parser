import csv, json, os, re

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CAT = os.path.join(BASE, "output", "delivery_v5_2", "doors_catalog.csv")
OUT = os.path.join(BASE, "output", "delivery_v5_2", "doors_catalog.csv")

ASD_KEY = "\u0410\u0421\u0414"

with open(CAT, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
fn = list(rows[0].keys())
print(f"Rows: {len(rows)}")

# Load known inside URLs from wayback (to avoid duplicates)
wb_all = {}
with open(os.path.join(BASE, "output", "delivery_v5_2", "wayback_inside_all_results.json"), encoding='utf-8') as f:
    wb_data = json.load(f)
known_inside = set()
for v in wb_data.values():
    url = v.get('inside_url', '').strip()
    if url:
        known_inside.add(url.replace('https://', 'http://'))

print(f"Known inside URLs: {len(known_inside)}")

# Pattern matching for ASD rows without inside images
restored = 0
missing = 0
patterns_tried = {}

for r in rows:
    if r['producer'] != ASD_KEY:
        continue
    if r.get('image_url_inside', '').strip():
        continue  # already has inside
    
    front_url = r.get('image_url_front', '').strip()
    if not front_url:
        missing += 1
        continue

    # Try various patterns
    # Pattern 1: front -> inside by replacing filename suffix
    # e.g., model-1.jpg -> model-2.jpg, model-3.jpg, etc.
    base = front_url.rsplit('.', 1)
    if len(base) != 2:
        missing += 1
        continue
    base_url = base[0]
    ext = base[1]

    found = False
    for n in range(1, 20):
        # Remove trailing number from base and add new number
        clean_base = base_url.rstrip('0123456789').rstrip('-')
        # Try: clean_base-n.ext and base_url-n.ext
        for candidate in [f"{clean_base}-{n}.{ext}", f"{base_url}-{n}.{ext}"]:
            if candidate == front_url:
                continue
            if candidate not in known_inside:
                r['image_url_inside'] = candidate
                restored += 1
                found = True
                break
        if found:
            break
    
    if not found:
        missing += 1
        patterns_tried[front_url] = 'no_match'

# Stats
asd_rows = [r for r in rows if r['producer'] == ASD_KEY]
f_ok = sum(1 for r in asd_rows if r.get('image_url_front','').strip())
i_ok = sum(1 for r in asd_rows if r.get('image_url_inside','').strip())
b_ok = sum(1 for r in asd_rows if r.get('image_url_front','').strip() and r.get('image_url_inside','').strip())

print(f"\nASD: {len(asd_rows)} rows")
print(f"Front: {f_ok}/199")
print(f"Inside: {i_ok}/199 (pattern restored: {restored}, still missing: {missing})")
print(f"Both: {b_ok}/199")

# Save
with open(OUT, 'w', encoding='utf-8-sig', newline='') as f:
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writeheader()
    csv.DictWriter(f, fieldnames=fn, delimiter=';').writerows(rows)
print(f"Saved: {OUT}")
