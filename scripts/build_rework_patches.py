"""Stage 2: Build patches based on audit findings"""
import csv, json, os
from collections import defaultdict

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CATALOG_CSV = os.path.join(BASE, "output", "final_hybrid_client_ready_v16_3_asd_full_door_only", "doors_catalog.csv")
AUDIT_DIR = os.path.join(BASE, "output", "final_rework_audit")
OUT_DIR = os.path.join(BASE, "output", "final_rework_patches")
os.makedirs(OUT_DIR, exist_ok=True)

# Load catalog
with open(CATALOG_CSV, encoding="utf-8-sig") as f:
    catalog = list(csv.DictReader(f, delimiter=";"))

# Load audit
with open(os.path.join(AUDIT_DIR, "current_vs_expected_summary.json"), encoding="utf-8") as f:
    audit = json.load(f)

with open(os.path.join(AUDIT_DIR, "audit_missing_lines_models.csv"), encoding="utf-8-sig") as f:
    sample = f.read(2000)
    audit_delim = ";" if sample.count(";") > sample.count(",") else ","
with open(os.path.join(AUDIT_DIR, "audit_missing_lines_models.csv"), encoding="utf-8-sig") as f:
    audit_rows = list(csv.DictReader(f, delimiter=audit_delim))

# Build catalog index
cat_index = defaultdict(lambda: defaultdict(list))
for r in catalog:
    cat_index[r["producer"]][r["target_line"]].append(r)

# === PATCH 1: Missing models ===
patch_models = []
for row in audit_rows:
    producer = row["producer"]
    issue = row["issue_type"]
    ref_line = row["expected_line"]
    
    if issue == "missing_line":
        # Need to add line + models from site
        patch_models.append({
            "producer": producer,
            "target_line": ref_line,
            "model_name": "",
            "covering": row["expected_covering"],
            "source_url": "",
            "proposed_status": "needs_review",
            "proposed_comment": f"missing_line_from_reference_xlsx",
            "reason": f"Line '{ref_line}' exists in reference xlsx but not in catalog"
        })
    elif issue == "missing_models":
        # Line exists but has only 1 model
        cat_line = row.get("matched_catalog_line") or ref_line
        existing = cat_index.get(producer, {}).get(cat_line, [])
        existing_count = len(existing)
        existing_urls = set(r.get("source_url","") for r in existing)
        
        patch_models.append({
            "producer": producer,
            "target_line": cat_line,
            "model_name": f"[need_more_models] (currently {existing_count})",
            "covering": row["expected_covering"],
            "source_url": existing[0].get("source_url","") if existing else "",
            "proposed_status": "needs_review",
            "proposed_comment": f"missing_models_currently_{existing_count}_check_site_for_more",
            "reason": f"Reference expects multiple models, catalog has {existing_count}"
        })

# Write patch_missing_models.csv
with open(os.path.join(OUT_DIR, "patch_missing_models.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=["producer","target_line","model_name","covering",
        "source_url","proposed_status","proposed_comment","reason"])
    w.writeheader()
    w.writerows(patch_models)
print(f"patch_missing_models.csv: {len(patch_models)} rows")

# === PATCH 2: Photo front/inside ===
patch_photos = []
for r in catalog:
    producer = r["producer"]
    img = r.get("image_url","") or ""
    img_front = r.get("image_url_front","") or ""
    img_inside = r.get("image_url_inside","") or ""
    door_type = r.get("door_type","") or ""
    is_entry = "входн" in door_type.lower()
    
    issue = "ok"
    decision = "keep_old"
    reason = ""
    
    if not img and not img_front:
        issue = "missing_front"
        decision = "not_found"
        reason = "no_image_at_all"
    elif is_entry and not img_inside:
        issue = "missing_inside"
        decision = "manual_check"
        reason = "entry_door_without_inside_photo"
    elif "cache_image" in img.lower():
        issue = "needs_manual_review"
        decision = "manual_check"
        reason = "old_cache_image_may_be_fragment"
    
    if issue != "ok":
        patch_photos.append({
            "producer": producer,
            "target_line": r.get("target_line",""),
            "model_name": r.get("model_name",""),
            "source_url": r.get("source_url",""),
            "old_image_url": img,
            "new_image_url_front": "",
            "new_image_url_inside": "",
            "decision": decision,
            "reason": reason,
            "front_candidates_checked": 0,
            "inside_candidates_checked": 0,
            "rejected_technical_count": 0,
            "rejected_fragment_count": 0,
            "rejected_interior_count": 0,
            "notes": ""
        })

# Write patch_photo_front_inside.csv
with open(os.path.join(OUT_DIR, "patch_photo_front_inside.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=["producer","target_line","model_name","source_url",
        "old_image_url","new_image_url_front","new_image_url_inside","decision","reason",
        "front_candidates_checked","inside_candidates_checked","rejected_technical_count",
        "rejected_fragment_count","rejected_interior_count","notes"])
    w.writeheader()
    w.writerows(patch_photos)
print(f"patch_photo_front_inside.csv: {len(patch_photos)} rows")

# === PATCH 3: Type/covering fixes ===
# Check for mismatches between reference and catalog
patch_type = []
for row in audit_rows:
    if row["issue_type"] == "ok":
        continue
    producer = row["producer"]
    ref_line = row["expected_line"]
    ref_covering = row["expected_covering"]
    cat_line = row.get("matched_catalog_line") or ref_line
    
    # Check covering match
    existing = cat_index.get(producer, {}).get(cat_line, [])
    if existing:
        cat_covering = existing[0].get("coating","")
        if ref_covering and cat_covering and ref_covering.lower() != cat_covering.lower():
            patch_type.append({
                "producer": producer,
                "target_line": cat_line,
                "model_name": existing[0].get("model_name",""),
                "old_type": existing[0].get("door_type",""),
                "new_type": "",
                "old_covering": cat_covering,
                "new_covering": ref_covering,
                "decision": "manual_check",
                "reason": f"covering_mismatch_ref={ref_covering}_catalog={cat_covering}"
            })

with open(os.path.join(OUT_DIR, "patch_type_covering_fixes.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=["producer","target_line","model_name","old_type",
        "new_type","old_covering","new_covering","decision","reason"])
    w.writeheader()
    w.writerows(patch_type)
print(f"patch_type_covering_fixes.csv: {len(patch_type)} rows")

# === patch_report.json ===
report = {
    "generated_at": "2026-06-22",
    "version": "stage2_patches",
    "base_catalog": "v16.3",
    "reference": "Линейки производителей 3.xlsx",
    "summary": {
        "total_missing_lines": sum(1 for p in patch_models if p["reason"].startswith("Line")),
        "total_single_model_lines": sum(1 for p in patch_models if "need_more_models" in p["model_name"]),
        "total_photo_issues": len(patch_photos),
        "total_type_covering_fixes": len(patch_type),
        "manual_check_count": sum(1 for p in patch_photos if p["decision"] == "manual_check"),
        "not_found_count": sum(1 for p in patch_photos if p["decision"] == "not_found")
    },
    "special_rules": {
        "asd": "fully_covered_199_models",
        "thermobreak": "keep_old_do_not_modify"
    },
    "by_producer": {}
}

# Group by producer
for p in patch_models:
    prod = p["producer"]
    if prod not in report["by_producer"]:
        report["by_producer"][prod] = {"missing_lines": 0, "single_model_lines": 0}
    if p["reason"].startswith("Line"):
        report["by_producer"][prod]["missing_lines"] += 1
    else:
        report["by_producer"][prod]["single_model_lines"] += 1

with open(os.path.join(OUT_DIR, "patch_report.json"), "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)
print("patch_report.json: saved")

# === patch_preview.html ===
html = """<!DOCTYPE html>
<html lang="ru">
<head><meta charset="utf-8"><title>Patch Preview</title>
<style>
body { font-family: sans-serif; margin: 20px; }
table { border-collapse: collapse; width: 100%; margin: 10px 0; }
th, td { border: 1px solid #ddd; padding: 6px 10px; text-align: left; font-size: 13px; }
th { background: #f5f5f5; }
.missing { background: #fff3cd; }
.photo-issue { background: #f8d7da; }
.ok { background: #d4edda; }
h2 { margin-top: 30px; }
</style></head>
<body>
<h1>Patch Preview — Stage 2</h1>
<p>Based on audit of v16.3 against Линейки производителей 3.xlsx</p>

<h2>Missing Lines (need to add from site)</h2>
<table>
<tr><th>Producer</th><th>Line</th><th>Covering</th><th>Action</th></tr>
"""
for p in patch_models:
    if p["reason"].startswith("Line"):
        html += f'<tr class="missing"><td>{p["producer"]}</td><td>{p["target_line"]}</td><td>{p["covering"]}</td><td>{p["reason"]}</td></tr>\n'
html += "</table>\n"

html += """
<h2>Single-Model Lines (check site for more)</h2>
<table>
<tr><th>Producer</th><th>Line</th><th>Current Models</th><th>Covering</th></tr>
"""
for p in patch_models:
    if "need_more_models" in p["model_name"]:
        html += f'<tr class="photo-issue"><td>{p["producer"]}</td><td>{p["target_line"]}</td><td>{p["model_name"]}</td><td>{p["covering"]}</td></tr>\n'
html += "</table>\n"

html += """
<h2>Photo Issues</h2>
<table>
<tr><th>Producer</th><th>Model</th><th>Issue</th><th>Decision</th><th>Reason</th></tr>
"""
for p in patch_photos[:50]:  # Limit to first 50
    cls = "missing" if p["decision"] == "not_found" else "photo-issue"
    html += f'<tr class="{cls}"><td>{p["producer"]}</td><td>{p["model_name"][:40]}</td><td>{p["decision"]}</td><td>{p["reason"]}</td><td></td></tr>\n'
if len(patch_photos) > 50:
    html += f'<tr><td colspan="5">... and {len(patch_photos)-50} more</td></tr>\n'
html += "</table>\n"

html += """
<h2>Type/Covering Fixes</h2>
<table>
<tr><th>Producer</th><th>Line</th><th>Old Covering</th><th>New Covering</th></tr>
"""
for p in patch_type:
    html += f'<tr><td>{p["producer"]}</td><td>{p["target_line"]}</td><td>{p["old_covering"]}</td><td>{p["new_covering"]}</td></tr>\n'
html += "</table>\n"

html += """
<h2>Special Rules</h2>
<ul>
<li><strong>АСД</strong>: Considered fully covered (199 models). Reference says "Все модели".</li>
<li><strong>Термодур</strong>: Old thermobreak data preserved as-is. No aggressive reparsing.</li>
</ul>
</body></html>
"""

with open(os.path.join(OUT_DIR, "patch_preview.html"), "w", encoding="utf-8") as f:
    f.write(html)
print("patch_preview.html: saved")

# XLSX versions
try:
    import openpyxl
    
    # patch_missing_models.xlsx
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Missing Models"
    fields = ["producer","target_line","model_name","covering","source_url",
              "proposed_status","proposed_comment","reason"]
    for ci, h in enumerate(fields, 1):
        ws.cell(row=1, column=ci, value=h)
    for ri, row in enumerate(patch_models, 2):
        for ci, h in enumerate(fields, 1):
            ws.cell(row=ri, column=ci, value=row.get(h, ""))
    wb.save(os.path.join(OUT_DIR, "patch_missing_models.xlsx"))
    print("patch_missing_models.xlsx: saved")
    
    # patch_photo_front_inside.xlsx
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Photo Front Inside"
    pfields = ["producer","target_line","model_name","source_url","old_image_url",
               "new_image_url_front","new_image_url_inside","decision","reason",
               "front_candidates_checked","inside_candidates_checked",
               "rejected_technical_count","rejected_fragment_count",
               "rejected_interior_count","notes"]
    for ci, h in enumerate(pfields, 1):
        ws.cell(row=1, column=ci, value=h)
    for ri, row in enumerate(patch_photos, 2):
        for ci, h in enumerate(pfields, 1):
            ws.cell(row=ri, column=ci, value=row.get(h, ""))
    wb.save(os.path.join(OUT_DIR, "patch_photo_front_inside.xlsx"))
    print("patch_photo_front_inside.xlsx: saved")
except ImportError:
    print("openpyxl not available, skipping xlsx")

# Summary
print(f"\n{'='*60}")
print(f"STAGE 2 PATCH SUMMARY")
print(f"{'='*60}")
print(f"Missing lines to add: {report['summary']['total_missing_lines']}")
print(f"Single-model lines to check: {report['summary']['total_single_model_lines']}")
print(f"Photo issues: {report['summary']['total_photo_issues']}")
print(f"Type/covering fixes: {report['summary']['total_type_covering_fixes']}")
print(f"Manual check needed: {report['summary']['manual_check_count']}")
print(f"Not found: {report['summary']['not_found_count']}")
print(f"\nBy producer:")
for p, s in sorted(report["by_producer"].items()):
    parts = []
    if s["missing_lines"]: parts.append(f"{s['missing_lines']} missing_lines")
    if s["single_model_lines"]: parts.append(f"{s['single_model_lines']} single_model")
    print(f"  {p}: {', '.join(parts) if parts else 'OK'}")
