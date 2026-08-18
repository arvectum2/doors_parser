"""Final verification after fix pack"""
import csv, os

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
with open(os.path.join(BASE, "output", "delivery_v5_5", "doors_catalog.csv"), encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

LORD = chr(0x041B)+chr(0x043E)+chr(0x0440)+chr(0x0434)
ASD  = chr(0x0410)+chr(0x0421)+chr(0x0414)

# 4 Lord models still empty
lord_empty = [r for r in rows if r['producer']==LORD and not r.get('image_url_front','').strip()]
print("LORD EMPTY FRONT (4):")
for l in lord_empty:
    c = l.get('comment','')[:80] if l.get('comment') else ''
    print(f"  {l['model_name']} comment={c}")

# Rows 1638-1644
print("\nROWS 1638-1644:")
for idx, r in enumerate(rows):
    er = idx + 2
    if 1638 <= er <= 1644:
        f = r.get('image_url_front','')[:60]
        print(f"  Row {er}: {r['producer']} {r['model_name']} front={f}")

# Stats
asd = [r for r in rows if r['producer']==ASD]
print(f"\nASD front filled: {sum(1 for r in asd if r.get('image_url_front','').strip())}/199")
print(f"ASD inside filled: {sum(1 for r in asd if r.get('image_url_inside','').strip())}/199")

keys=set(); dupes=0
for r in rows:
    k=(r['producer'],r['target_line'],r['model_name'])
    if k in keys: dupes+=1
    keys.add(k)
print(f"Duplicates: {dupes}")

wp = all(r.get('image_url_front','').strip() and r['image_url']==r['image_url_front']
         for r in rows if r.get('image_url_front','').strip())
print(f"WP sync: {wp}")
print(f"Total rows: {len(rows)}")
print(f"Producers: {len({r['producer'] for r in rows})}")

# Lord with steklo still in comment
lord_steklo_remaining = [r for r in rows if r['producer']==LORD and 'steklo' in (r.get('comment','') or '').lower() and 'lord_steklo_variant_replaced' not in (r.get('comment','') or '')]
print(f"\nLord with steklo comment (not replaced): {len(lord_steklo_remaining)}")

# Fragment/handle/lock remaining
frag = [r for r in rows if 'fragment' in (r.get('comment','') or '')]
bl = [r for r in rows if 'handle' in (r.get('comment','') or '') or 'lock' in (r.get('comment','') or '')]
print(f"Fragment comments: {len(frag)}")
print(f"Handle/lock comments: {len(bl)}")

# No-image total
no_img = [r for r in rows if not r.get('image_url_front','').strip() and not r.get('image_url_inside','').strip()]
print(f"No-image rows total: {len(no_img)}")
by_prod = {}
for r in no_img:
    p = r['producer']
    by_prod[p] = by_prod.get(p, 0) + 1
print(f"No-image by producer: {by_prod}")

# Lord front count
lord = [r for r in rows if r['producer']==LORD]
lord_f = sum(1 for r in lord if r.get('image_url_front','').strip())
print(f"Lord front: {lord_f}/{len(lord)}")
