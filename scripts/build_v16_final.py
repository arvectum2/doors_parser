"""1. Sanitize ASD Photo Patch v2 + 2. Build v16 Archive"""
import csv, os, json, hashlib, shutil, re
from collections import Counter, defaultdict
from datetime import datetime
from zipfile import ZipFile

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
V14_DIR = os.path.join(BASE, "output", "final_hybrid_client_ready_v14")
V14_CSV = os.path.join(V14_DIR, "doors_catalog.csv")

PATCH_DIR = os.path.join(BASE, "output", "asd_photo_patch_v2")
PATCH_CSV = os.path.join(PATCH_DIR, "asd_photo_patch_v2.csv")
PATCH_DEBUG = os.path.join(PATCH_DIR, "asd_photo_candidates_debug.csv")

CLEAN_DIR = os.path.join(BASE, "output", "asd_photo_patch_v2_clean")
os.makedirs(CLEAN_DIR, exist_ok=True)

BAD_PATTERNS = ["interier","interior","banner","badge","hit","logo","icon","handle",
                "ruchka","ручка","zamok","lock","замок","fragment","texture","sample","color"]

# ─── Load data ───

with open(PATCH_CSV, encoding="utf-8-sig") as f:
    patch_rows = list(csv.DictReader(f, delimiter=";"))

with open(PATCH_DEBUG, encoding="utf-8-sig") as f:
    debugs = list(csv.DictReader(f, delimiter=";"))

with open(V14_CSV, encoding="utf-8-sig") as f:
    v14_rows = list(csv.DictReader(f, delimiter=";"))

# Build debug lookup by key
debug_by_key = defaultdict(list)
for d in debugs:
    key = (d["model_name"], d["source_url"])
    debug_by_key[key].append(d)

# Build selected lookup by key
selected_by_key = {}
for d in debugs:
    if d.get("is_selected") == "YES":
        selected_by_key[(d["model_name"], d["source_url"])] = d

# ─── Collect changes for report ───
changes = {
    "fragment_to_manual": 0,
    "not_found_url_cleared": 0,
    "replace_downgraded": 0,
    "manual_handled": 0,
}

# ─── Process each row ───
clean_rows = []
for r in patch_rows:
    row = dict(r)  # make a copy
    key = (row["model_name"], row["source_url"])
    sel_debug = selected_by_key.get(key)
    decision = row["decision"]
    new_decision = decision
    
    if decision == "replace":
        # Strict check
        url = row.get("selected_image_url","") or ""
        score_str = row.get("full_door_score","0") or "0"
        score = int(score_str) if score_str.isdigit() else 0
        
        issues = []
        if not url:
            issues.append("empty_url")
        if score < 100:
            issues.append(f"score={score}<100")
        if any(b in url.lower() for b in BAD_PATTERNS):
            issues.append("bad_pattern")
        if sel_debug:
            vd = sel_debug.get("visual_decision","")
            if vd not in ("full_door", "full_door_short"):
                issues.append(f"visual_decision={vd}")
            ar_str = sel_debug.get("aspect_ratio","0")
            try:
                ar = float(ar_str)
            except:
                ar = 0
            if ar < 1.5:
                issues.append(f"ar={ar}<1.5")
        
        if issues:
            changes["replace_downgraded"] += 1
            new_decision = "manual_check"
            row["decision"] = "manual_check"
            row["reason"] = f"replace_rejected_{'_'.join(issues)}"
            row["patch_comment"] = "manual_check_replace_vetoed"
            row["selected_image_url"] = ""
            row["full_door_score"] = "0"
    
    if new_decision == "not_found":
        if row.get("selected_image_url",""):
            changes["not_found_url_cleared"] += 1
            row["selected_image_url"] = ""
    
    if new_decision == "manual_check":
        # Clear selected_image_url unless old was already a known good full-door
        old_url = row.get("old_image_url","") or ""
        row["selected_image_url"] = ""
        row["status_before"] = row.get("status_before","")
        row["comment_before"] = (row.get("comment_before","") or "") + "; asd_photo_patch_manual_check_required"
        if row.get("comment_before","").startswith("; "):
            row["comment_before"] = row["comment_before"][2:]
        changes["manual_handled"] += 1
    
    if new_decision == "keep":
        # Keep as-is - image_url already correct
        pass
    
    clean_rows.append(row)

# ─── Write cleaned patch ───
fieldnames = [
    "producer", "target_line", "model_name", "source_url",
    "old_image_url", "selected_image_url", "decision", "reason",
    "full_door_score", "candidates_checked",
    "rejected_fragment_count", "rejected_interior_count",
    "rejected_banner_count", "selected_candidate_source",
    "status_before", "comment_before", "patch_comment"
]

clean_csv = os.path.join(CLEAN_DIR, "asd_photo_patch_v2_clean.csv")
with open(clean_csv, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";")
    w.writeheader()
    w.writerows(clean_rows)

# XLSX
try:
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ASD Clean"
    for ci, h in enumerate(fieldnames, 1):
        ws.cell(row=1, column=ci, value=h)
    for ri, r in enumerate(clean_rows, 2):
        for ci, h in enumerate(fieldnames, 1):
            ws.cell(row=ri, column=ci, value=r.get(h, ""))
    clean_xlsx = os.path.join(CLEAN_DIR, "asd_photo_patch_v2_clean.xlsx")
    wb.save(clean_xlsx)
except ImportError:
    pass

# ─── Generate HTML review ───
decisions_final = Counter(r["decision"] for r in clean_rows)

def html_review(rows, debugs):
    sections = defaultdict(list)
    for r in rows:
        sections[r["decision"]].append(r)
    
    debug_lookup = defaultdict(list)
    for d in debugs:
        debug_lookup[(d["model_name"], d["source_url"])].append(d)
    
    total = len(rows)
    html = f"""<!DOCTYPE html>
<html lang="ru">
<head><meta charset="UTF-8">
<title>ASD Photo Patch v2 Clean Review</title>
<style>
body {{ font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }}
h1 {{ color: #333; }}
.summary {{ background: #fff; padding: 20px; border-radius: 8px; margin-bottom: 20px; }}
.grid {{ display: flex; flex-wrap: wrap; gap: 15px; }}
.card {{ background: #fff; border: 1px solid #ddd; border-radius: 8px; padding: 12px; width: 360px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
.card img {{ max-width: 100%; height: auto; margin: 5px 0; border: 1px solid #eee; border-radius: 4px; }}
.card .selected-img {{ border: 3px solid #4CAF50; }}
.label {{ font-weight: bold; color: #666; font-size: 11px; text-transform: uppercase; margin-top: 8px; }}
.value {{ margin-bottom: 4px; font-size: 12px; }}
.decision-replace {{ border-left: 4px solid #4CAF50; }}
.decision-keep {{ border-left: 4px solid #2196F3; }}
.decision-not_found {{ border-left: 4px solid #f44336; }}
.decision-manual_check {{ border-left: 4px solid #FF9800; }}
.badge {{ display: inline-block; padding: 2px 8px; border-radius: 3px; font-size: 11px; color: #fff; }}
.badge-replace {{ background: #4CAF50; }}
.badge-keep {{ background: #2196F3; }}
.badge-not_found {{ background: #f44336; }}
.badge-manual_check {{ background: #FF9800; }}
</style></head><body>
<h1>ASD Photo Patch v2 Clean Review</h1>
<div class="summary">
<p>Total ASD rows: <strong>{total}</strong></p>
<p>Replace: <strong>{decisions_final.get('replace',0)}</strong> | Keep: <strong>{decisions_final.get('keep',0)}</strong> | Not found: <strong>{decisions_final.get('not_found',0)}</strong> | Manual check: <strong>{decisions_final.get('manual_check',0)}</strong></p>
</div>
"""
    for section_name, label, badge_cls in [
        ("replace", "Replace", "replace"), ("keep", "Keep", "keep"),
        ("manual_check", "Manual Check", "manual_check"), ("not_found", "Not Found", "not_found"),
    ]:
        items = sections.get(section_name, [])
        if not items:
            continue
        html += f'<h2><span class="badge badge-{badge_cls}">{label}</span> {len(items)} rows</h2><div class="grid">\n'
        for r in items:
            new_img = r.get("selected_image_url", "")
            old_img = r.get("old_image_url", "")
            old_thumb = f'<img src="{old_img}" onerror="this.style.display=\'none\'">' if old_img else '<div style="color:#999;font-style:italic;">(empty)</div>'
            new_thumb = f'<img src="{new_img}" onerror="this.style.display=\'none\'" class="selected-img">' if new_img else '<div style="color:#999;font-style:italic;">(empty)</div>'
            html += f'''
<div class="card decision-{section_name}">
<div class="label">Model ({r.get("full_door_score","")})</div>
<div class="value"><b>{r["model_name"][:50]}</b></div>
<div class="label">Source</div>
<div class="value" style="font-size:11px;"><a href="{r["source_url"]}" target="_blank">{r["source_url"][:40]}...</a></div>
<div class="label">OLD</div>
{old_thumb}
<div class="label">SELECTED {r.get("decision","").upper()}</div>
{new_thumb}
</div>'''
        html += '</div>\n'
    html += "</body></html>"
    return html

html_path = os.path.join(CLEAN_DIR, "asd_photo_patch_v2_clean_review.html")
with open(html_path, "w", encoding="utf-8") as f:
    f.write(html_review(clean_rows, debugs))

# ─── Report ───
report = {
    "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "total_asd_rows": len(clean_rows),
    "replace": decisions_final.get("replace",0),
    "keep": decisions_final.get("keep",0),
    "not_found": decisions_final.get("not_found",0),
    "manual_check": decisions_final.get("manual_check",0),
    "changes": {
        "fragment_selected_moved_to_manual": changes["fragment_to_manual"],
        "not_found_urls_cleared": changes["not_found_url_cleared"],
        "replace_downgraded": changes["replace_downgraded"],
        "manual_handled": changes["manual_handled"]
    }
}
report_path = os.path.join(CLEAN_DIR, "asd_photo_patch_v2_clean_report.json")
with open(report_path, "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)

print("=== Sanitization Report ===")
print(f"Total: {len(clean_rows)}")
print(f"Replace: {decisions_final.get('replace',0)}")
print(f"Keep: {decisions_final.get('keep',0)}")
print(f"Not found: {decisions_final.get('not_found',0)}")
print(f"Manual check: {decisions_final.get('manual_check',0)}")
print(f"Changes: {json.dumps(changes, ensure_ascii=False)}")

# ─── Build v16 ───
V16_DIR = os.path.join(BASE, "output", "final_hybrid_client_ready_v16") + "_" + datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
os.makedirs(V16_DIR, exist_ok=True)

# Build lookup by model_name+source_url
clean_lookup = {}
for r in clean_rows:
    key = (r["model_name"], r["source_url"])
    clean_lookup[key] = r

v16_rows = []
applied_replace = 0
applied_keep = 0
applied_skip = 0

for r in v14_rows:
    row = dict(r)
    producer = row.get("producer","")
    
    if producer == "АСД":
        key = (row["model_name"], row["source_url"])
        patch_row = clean_lookup.get(key)
        if patch_row:
            dec = patch_row["decision"]
            if dec == "replace":
                new_url = patch_row.get("selected_image_url","") or ""
                if new_url:
                    row["image_url"] = new_url
                    applied_replace += 1
                else:
                    applied_skip += 1
            elif dec == "keep":
                applied_keep += 1
                # Keep existing image_url
            else:
                applied_skip += 1
    # else: non-ASD producers stay unchanged
    
    v16_rows.append(row)

# Verify constraints
errors = []
if len(v16_rows) != 3084:
    errors.append(f"Total rows: {len(v16_rows)} (expected 3084)")

producers = [r["producer"] for r in v16_rows]
producer_counts = Counter(producers)
if len(producer_counts) != 10:
    errors.append(f"Producer count: {len(producer_counts)} (expected 10)")

# Check duplicates
seen = set()
dups = []
for r in v16_rows:
    key = (r.get("producer",""), r.get("target_line",""), r.get("model_name",""))
    if key in seen:
        dups.append(key)
    seen.add(key)
if dups:
    errors.append(f"Duplicates: {len(dups)}")

# Compare non-ASD producers
v14_producers = [r["producer"] for r in v14_rows]
v14_prod_counts = Counter(v14_producers)
for p, c in v14_prod_counts.items():
    if p != "АСД":
        v16_c = producer_counts.get(p, 0)
        if v16_c != c:
            errors.append(f"Producer {p}: v14={c} v16={v16_c}")

# Exact match check for non-ASD
for i, (v14_r, v16_r) in enumerate(zip(v14_rows, v16_rows)):
    if v14_r["producer"] != "АСД":
        for col in ["door_type","producer","target_line","model_name","image_url","source_url"]:
            if v14_r.get(col,"") != v16_r.get(col,""):
                errors.append(f"Row {i+2} non-ASD changed: {col}: '{v14_r.get(col,'')[:30]}' -> '{v16_r.get(col,'')[:30]}'")
                break

# Check ASD replaced URLs for bad patterns (only changed ones, not pre-existing v14 URLs)
asd_bad = []
for r in clean_rows:
    if r["decision"] == "replace":
        url = r.get("selected_image_url","") or ""
        if any(b in url.lower() for b in BAD_PATTERNS):
            asd_bad.append((r["model_name"], url))
            errors.append(f"ASD replace bad pattern: {r['model_name'][:25]} {url[:55]}")

# Check ASD fragment visual_decision in applied rows
applied_keys = set()
for r in clean_rows:
    if r["decision"] in ("replace", "keep"):
        applied_keys.add((r["model_name"], r["source_url"]))

asd_fragment = []
for d in debugs:
    key = (d["model_name"], d["source_url"])
    if d.get("is_selected") == "YES" and key in applied_keys:
        if d.get("visual_decision","") not in ("full_door", "full_door_short"):
            asd_fragment.append(d)
            errors.append(f"ASD applied fragment vd: {d['model_name'][:25]} vd={d['visual_decision']}")

if not errors:
    print("\n=== v16 Verified: No errors ===")
else:
    print(f"\n=== v16 ERRORS ({len(errors)}) ===")
    for e in errors:
        print(f"  ERROR: {e}")

# Write v16 catalog
v16_csv = os.path.join(V16_DIR, "doors_catalog.csv")
fieldnames_v16 = ["door_type","producer","target_line","model_name","description","material","coating","image_url","source_url","status","comment"]
with open(v16_csv, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=fieldnames_v16, delimiter=";")
    w.writeheader()
    w.writerows(v16_rows)

# XLSX
try:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Catalog v16"
    for ci, h in enumerate(fieldnames_v16, 1):
        ws.cell(row=1, column=ci, value=h)
    for ri, r in enumerate(v16_rows, 2):
        for ci, h in enumerate(fieldnames_v16, 1):
            ws.cell(row=ri, column=ci, value=r.get(h, ""))
    v16_xlsx = os.path.join(V16_DIR, "doors_catalog.xlsx")
    wb.save(v16_xlsx)
except ImportError:
    pass

# Create ZIP
zip_name = f"doors_parser_final_HYBRID_CLIENT_READY_v16_ASD_FULL_DOOR_ONLY_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.zip"
zip_path = os.path.join(BASE, "output", zip_name)
with ZipFile(zip_path, "w") as zf:
    for fn in os.listdir(V16_DIR):
        fp = os.path.join(V16_DIR, fn)
        zf.write(fp, arcname=fn)

# SHA256
sha = hashlib.sha256()
with open(zip_path, "rb") as f:
    for chunk in iter(lambda: f.read(65536), b""):
        sha.update(chunk)

print(f"\n=== v16 Summary ===")
print(f"Total rows: {len(v16_rows)}")
print(f"Producers: {len(producer_counts)}")
print(f"ASD applied replace: {applied_replace}")
print(f"ASD applied keep: {applied_keep}")
print(f"ASD skipped (manual/not_found): {applied_skip}")
print(f"Errors: {len(errors)}")
if not errors:
    print("ALL CHECKS PASSED")
    print(f"ZIP: {zip_path}")
    print(f"SHA256: {sha.hexdigest()}")
else:
    for e in errors:
        print(f"  FAIL: {e}")
