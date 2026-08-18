"""V5.4-1: ASD row count reconciliation - check 199 vs 200"""
import csv, json, os, re
from collections import Counter, defaultdict

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CAT = os.path.join(BASE, "output", "delivery_v5_3", "doors_catalog.csv")
OUT_DIR = os.path.join(BASE, "output", "delivery_v5_4")
os.makedirs(OUT_DIR, exist_ok=True)

ASD_KEY = "\u0410\u0421\u0414"

# Load current catalog
with open(CAT, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
fn = list(rows[0].keys())

asd = [r for r in rows if r['producer'] == ASD_KEY]
print(f"Current ASD rows: {len(asd)} (expected 199 or 200)")

# Check producers/duplicates
prods = Counter(r['producer'] for r in rows)
keys = Counter((r['producer'], r['target_line'], r['model_name']) for r in rows)
dupes = {k:v for k,v in keys.items() if v > 1}
print(f"Producers: {len(prods)}")
print(f"Duplicates: {len(dupes)}")

# ========================================================================
# Search for 200th ASD row across historical data
# ========================================================================
# Check v5.2 catalog
cat_v52 = os.path.join(BASE, "output", "delivery_v5_2", "doors_catalog.csv")
with open(cat_v52, encoding='utf-8-sig') as f:
    v52 = list(csv.DictReader(f, delimiter=';'))
v52_asd = [r for r in v52 if r['producer'] == ASD_KEY]
print(f"\nV5.2 ASD: {len(v52_asd)}")

# Check v5.1 catalog
cat_v51 = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")
with open(cat_v51, encoding='utf-8-sig') as f:
    v51 = list(csv.DictReader(f, delimiter=';'))
v51_asd = [r for r in v51 if r['producer'] == ASD_KEY]
print(f"V5.1 ASD: {len(v51_asd)}")

# Check v4 catalog (from v5 build zip)
pass

# Check patch CSV for model names
patch_csv = os.path.join(BASE, "output", "asd_photo_patch_v2_clean", "asd_photo_patch_v2_clean.csv")
with open(patch_csv, encoding='utf-8-sig') as f:
    patch = list(csv.DictReader(f, delimiter=';'))
patch_models = set(r['model_name'] for r in patch)
print(f"\nPatch models: {len(patch_models)}")

# Check wayback JSON for model names
wb_json = os.path.join(BASE, "output", "delivery_v5_2", "wayback_inside_all_results.json")
try:
    with open(wb_json, encoding='utf-8') as f:
        wb = json.load(f)
    wb_models = {k for k in wb.keys()}
    print(f"Wayback models: {len(wb_models)}")
except:
    wb_models = set()

# Cross-reference: which models are in patch or wayback but NOT in current catalog?
cat_models = set(r['model_name'] for r in asd)
only_in_patch = patch_models - cat_models
only_in_wb = wb_models - cat_models
print(f"\nModels in patch but NOT in catalog: {len(only_in_patch)}")
for m in sorted(only_in_patch):
    print(f"  PATCH: {m}")

print(f"\nModels in wayback but NOT in catalog: {len(only_in_wb)}")
for m in sorted(only_in_wb):
    print(f"  WB: {m}")

# Check models that might be lost between versions
v51_models = set(r['model_name'] for r in v51_asd)
lost_since_v51 = v51_models - cat_models
print(f"\nModels lost since v5.1: {len(lost_since_v51)}")
for m in sorted(lost_since_v51):
    print(f"  LOST: {m}")

v52_models = set(r['model_name'] for r in v52_asd)
lost_since_v52 = v52_models - cat_models
print(f"\nModels lost since v5.2: {len(lost_since_v52)}")
for m in sorted(lost_since_v52):
    v52_row = next((r for r in v52_asd if r['model_name'] == m), None)
    v53_comm = next((r.get('comment','') for r in asd if r['model_name']==m), 'N/A')
    v51_count = sum(1 for r in v51_asd if r['model_name'] == m)
    print(f"  {m[:60]:60s} v53_comment={v53_comm[:40]} found_in_v51={v51_count}")

# Check if any v5.3 ASD rows have empty producer or bad data
bad = [r for r in rows if r['producer'] == '' or r['model_name'] == '']
print(f"\nRows with empty producer/model: {len(bad)}")

# Model normalized overlap analysis
def norm_name(n):
    n = n.lower()
    n = n.replace('"', '').replace('«', '').replace('»', '').replace('"', '')
    n = re.sub(r'\s+', ' ', n).strip()
    return n

cat_norm = {norm_name(r['model_name']): r['model_name'] for r in asd}
patch_norm = {norm_name(r['model_name']): r['model_name'] for r in patch}
wb_norm = {norm_name(k): k for k in wb_models}

# Find non-overlapping normalized names
patch_extra = set(patch_norm.keys()) - set(cat_norm.keys())
wb_extra = set(wb_norm.keys()) - set(cat_norm.keys())
print(f"\nNormalized: patch extra = {len(patch_extra)}, wb extra = {len(wb_extra)}")
for n in sorted(patch_extra):
    print(f"  PATCH NORM: {patch_norm[n]}")
for n in sorted(wb_extra):
    print(f"  WB NORM: {wb_norm[n]}")

# Output reconciliation report
print(f"\n{'='*60}")
print(f"ASD ROW COUNT RECONCILIATION")
print(f"{'='*60}")
print(f"  Current ASD rows: {len(asd)}")
print(f"  V5.1 ASD rows: {len(v51_asd)}")
print(f"  V5.2 ASD rows: {len(v52_asd)}")
print(f"  Patch models: {len(patch_models)}")
print(f"  Wayback models: {len(wb_models)}")
print(f"  Models overlapping all sources: {len(cat_models & patch_models & wb_models)}")
print(f"  Models only in patch: {len(only_in_patch)}")
print(f"  Models only in wayback: {len(only_in_wb)}")
print(f"  Models lost since v5.1: {len(lost_since_v51)}")
print(f"  Models lost since v5.2: {len(lost_since_v52)}")
