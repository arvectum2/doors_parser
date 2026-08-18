"""Diagnose remaining issues in v5.5"""
import csv, os, json

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
OUT_DIR = os.path.join(BASE, "output", "delivery_v5_5")

with open(os.path.join(OUT_DIR, "doors_catalog.csv"), encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

# Comments summary
from collections import Counter
comments = Counter()
for r in rows:
    c = r.get('comment', '')
    if c:
        for part in c.split('; '):
            comments[part] += 1

print("=== COMMENT TAGS COUNT ===")
for k, v in sorted(comments.items(), key=lambda x: -x[1]):
    if v > 0:
        print(f"  {v:4d} {k[:80]}")

# Techno models
print("\n=== TECHNO MODELS ===")
techno = [r for r in rows if r['model_name'].lower().startswith('techno')]
print(f"Total: {len(techno)}")
for t in techno:
    print(f"  {t['model_name'][:40]} front={'Y' if t.get('image_url_front','') else 'N'} comment={t.get('comment','')[:60]}")

# Lord with empty front
print("\n=== LORD WITH EMPTY FRONT ===")
lord_empty = [r for r in rows if r['producer'] == '\u041b\u043e\u0440\u0434' and not r.get('image_url_front', '').strip()]
print(f"Total: {len(lord_empty)}")
for l in lord_empty[:10]:
    print(f"  {l['model_name'][:50]} target={l.get('target_line','')[:20]} comment={l.get('comment','')[:60]}")

# Remaining fragment/handle comments
frag = [r for r in rows if 'fragment' in (r.get('comment', '') or '') or 'steklo' in (r.get('comment', '') or '').lower()]
lock = [r for r in rows if 'handle' in (r.get('comment', '') or '') or 'lock' in (r.get('comment', '') or '') or 'technical_reject' in (r.get('comment', '') or '')]
print(f"\nRows with fragment/steklo comment: {len(frag)}")
for f in frag[:5]:
    print(f"  {f['producer'][:10]} {f['model_name'][:40]}")

print(f"Rows with handle/lock/tech comment: {len(lock)}")
for l in lock[:5]:
    print(f"  {l['producer'][:10]} {l['model_name'][:40]}")

# Check for models that lost both front and inside
no_img = [r for r in rows if not r.get('image_url_front', '').strip() and not r.get('image_url_inside', '').strip()]
print(f"\nRows with NO images at all: {len(no_img)}")
for n in no_img:
    print(f"  {n['producer'][:15]} {n['model_name'][:50]} target={n.get('target_line','')[:20]}")

# Producer counts
from collections import Counter
prods = Counter(r['producer'] for r in rows)
print(f"\nProducers: {dict(sorted(prods.items(), key=lambda x: -x[1]))}")

# Total row count
print(f"\nTotal rows: {len(rows)}")

# ASD stats
ASD_KEY = chr(0x0410)+chr(0x0421)+chr(0x0414)
asd = [r for r in rows if r['producer'] == ASD_KEY]
f = sum(1 for r in asd if r.get('image_url_front','').strip())
i = sum(1 for r in asd if r.get('image_url_inside','').strip())
b = sum(1 for r in asd if r.get('image_url_front','').strip() and r.get('image_url_inside','').strip())
print(f"ASD: {len(asd)} front={f}/199 inside={i}/199 both={b}/199")

# V5.4 front count for comparison
with open(os.path.join(BASE, "output", "delivery_v5_4", "doors_catalog.csv"), encoding='utf-8-sig') as fh:
    rows54 = list(csv.DictReader(fh, delimiter=';'))
asd54 = [r for r in rows54 if r['producer'] == ASD_KEY]
f54 = sum(1 for r in asd54 if r.get('image_url_front','').strip())
i54 = sum(1 for r in asd54 if r.get('image_url_inside','').strip())
print(f"V5.4 ASD: front={f54}/199 inside={i54}/199")
