"""Build reconciliation report"""
import csv, json, os
from collections import Counter, defaultdict

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
FINAL_DIR = os.path.join(BASE, "output", "final_rework_ready")
AUDIT_DIR = os.path.join(BASE, "output", "final_rework_audit")
PATCH_DIR = os.path.join(BASE, "output", "final_rework_patches")
REFERENCE_JSON = os.path.join(BASE, "output", "reference_lines_decoded.json")

# Load all data
with open(os.path.join(FINAL_DIR, "doors_catalog.csv"), encoding="utf-8-sig") as f:
    final = list(csv.DictReader(f, delimiter=";"))

with open(os.path.join(BASE, "output", "final_hybrid_client_ready_v16_3_asd_full_door_only", "doors_catalog.csv"), encoding="utf-8-sig") as f:
    v163 = list(csv.DictReader(f, delimiter=";"))

with open(os.path.join(AUDIT_DIR, "current_vs_expected_summary.json"), encoding="utf-8") as f:
    audit = json.load(f)

with open(REFERENCE_JSON, encoding="utf-8") as f:
    reference = json.load(f)

def read_csv_auto(path):
    with open(path, encoding="utf-8-sig") as f:
        sample = f.read(2000)
    delim = ";" if sample.count(";") > sample.count(",") else ","
    with open(path, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f, delimiter=delim))

model_patches = read_csv_auto(os.path.join(PATCH_DIR, "patch_missing_models.csv"))
photo_patches = read_csv_auto(os.path.join(PATCH_DIR, "patch_photo_front_inside.csv"))

# Build indices
v163_keys = set((r["producer"], r["target_line"], r["model_name"]) for r in v163)
final_keys = set((r["producer"], r["target_line"], r["model_name"]) for r in final)
new_keys = final_keys - v163_keys

final_index = defaultdict(lambda: defaultdict(list))
for r in final:
    final_index[r["producer"]][r["target_line"]].append(r)

v163_index = defaultdict(lambda: defaultdict(list))
for r in v163:
    v163_index[r["producer"]][r["target_line"]].append(r)

# Build reconciliation rows
recon_rows = []

for producer, ref_lines in reference.items():
    for ref_entry in ref_lines:
        ref_line = ref_entry["line"]
        ref_covering = ref_entry["covering"]
        
        # Find in audit
        audit_entry = None
        for a in audit["by_producer"].get(producer, {}).get("lines_detail", []):
            if a["ref_line"] == ref_line:
                audit_entry = a
                break
        
        audit_status = audit_entry["issue"] if audit_entry else "unknown"
        
        # Count rows in v16.3 and final
        v163_count = len(v163_index.get(producer, {}).get(ref_line.strip(), []))
        final_count = len(final_index.get(producer, {}).get(ref_line.strip(), []))
        
        # Find matching patch entries
        patch_count = sum(1 for p in model_patches 
                         if p["producer"] == producer and p["target_line"].strip() == ref_line.strip())
        
        # Determine final_action and explanation
        final_action = ""
        explanation = ""
        
        if producer == "АСД":
            # АСД: reference says "Все модели" x11, catalog has 199 models
            final_action = "skipped_keep_old_thermobreak" if "терм" in ref_line.lower() else "mapped_to_existing_line"
            explanation = "Reference says 'Все модели' — catalog has 199 АСД rows. Considered fully covered."
            if v163_count > 0:
                final_action = "already_existed"
                explanation = f"АСД catalog has {v163_count} rows matching this placeholder. Skipped adding stubs."
        elif audit_status == "ok":
            final_action = "already_existed"
            explanation = f"Line found in catalog with {v163_count} rows in v16.3, {final_count} in final."
        elif audit_status == "missing_line":
            # Check if added as [MISSING] stub
            stub = None
            for r in final:
                if r["producer"] == producer and r["target_line"].strip() == ref_line.strip() and r["model_name"].startswith("[MISSING]"):
                    stub = r
                    break
            if stub:
                final_action = "added_placeholder"
                explanation = f"Added [MISSING] stub row. Needs site parsing to fill models."
            else:
                # Check if line exists under slightly different name
                found = False
                for fl in final_index.get(producer, {}).keys():
                    if ref_line.strip().lower() in fl.strip().lower() or fl.strip().lower() in ref_line.strip().lower():
                        final_action = "mapped_to_existing_line"
                        explanation = f"Found matching line '{fl}' in catalog with {len(final_index[producer][fl])} rows."
                        found = True
                        break
                if not found:
                    final_action = "left_missing_needs_review"
                    explanation = "Line not found in catalog and no stub added."
        elif audit_status == "missing_models":
            # Line exists but with few models
            final_action = "already_existed"
            explanation = f"Line exists with {v163_count} rows in v16.3. Patch marked as needs_review for checking more models on site."
        else:
            final_action = "left_missing_needs_review"
            explanation = f"Audit status: {audit_status}"
        
        recon_rows.append({
            "producer": producer,
            "expected_line": ref_line,
            "expected_covering": ref_covering,
            "audit_status": audit_status,
            "current_catalog_rows_before": v163_count,
            "final_catalog_rows_after": final_count,
            "patch_entries_count": patch_count,
            "final_status": "found" if final_count > 0 else "missing",
            "final_action": final_action,
            "explanation": explanation
        })

# Write reconciliation_report.csv
with open(os.path.join(FINAL_DIR, "reconciliation_report.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=["producer","expected_line","expected_covering",
        "audit_status","current_catalog_rows_before","final_catalog_rows_after",
        "patch_entries_count","final_status","final_action","explanation"])
    w.writeheader()
    w.writerows(recon_rows)
print(f"reconciliation_report.csv: {len(recon_rows)} rows")

# Write reconciliation_report.xlsx
try:
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reconciliation"
    fields = ["producer","expected_line","expected_covering","audit_status",
              "current_catalog_rows_before","final_catalog_rows_after",
              "patch_entries_count","final_status","final_action","explanation"]
    for ci, h in enumerate(fields, 1):
        ws.cell(row=1, column=ci, value=h)
    for ri, row in enumerate(recon_rows, 2):
        for ci, h in enumerate(fields, 1):
            ws.cell(row=ri, column=ci, value=row.get(h, ""))
    wb.save(os.path.join(FINAL_DIR, "reconciliation_report.xlsx"))
    print("reconciliation_report.xlsx: saved")
except ImportError:
    print("openpyxl not available")

# === Detailed checks ===

# 1. Missing lines that became needs_review
print("\n=== 5 MISSING LINES ADDED AS NEEDS_REVIEW ===")
missing_stubs = [r for r in final if r.get("model_name","").startswith("[MISSING]")]
for r in missing_stubs:
    print(f"  {r['producer']} | {r['target_line']} | {r['model_name']} | {r.get('comment','')[:50]}")

# 2. Patch entries breakdown
print("\n=== 57 PATCH ENTRIES BREAKDOWN ===")
patch_missing_line = [p for p in model_patches if p["reason"].startswith("Line")]
patch_missing_model = [p for p in model_patches if "need_more_models" in p.get("model_name","")]
print(f"  missing_line type: {len(patch_missing_line)}")
print(f"  missing_models type: {len(patch_missing_model)}")

# For each type, check final status
added_as_stub = 0
skipped_asd = 0
already_existed = 0
still_needs_review = 0

for p in model_patches:
    if p["producer"] == "АСД":
        skipped_asd += 1
    elif p["reason"].startswith("Line"):
        # Check if stub was added
        stub_exists = any(r["producer"] == p["producer"] and r["target_line"].strip() == p["target_line"].strip() 
                         and r["model_name"].startswith("[MISSING]") for r in final)
        if stub_exists:
            added_as_stub += 1
        else:
            # Check if line exists
            line_rows = [r for r in final if r["producer"] == p["producer"] and r["target_line"].strip() == p["target_line"].strip()]
            if line_rows:
                already_existed += 1
            else:
                still_needs_review += 1
    else:
        # missing_models type — these don't add rows, just mark existing for review
        line_rows = [r for r in final if r["producer"] == p["producer"] and r["target_line"].strip() == p["target_line"].strip()]
        if line_rows:
            already_existed += 1
        else:
            still_needs_review += 1

print(f"  Added as [MISSING] stub: {added_as_stub}")
print(f"  Skipped (АСД): {skipped_asd}")
print(f"  Already existed (line present): {already_existed}")
print(f"  Still missing/needs_review: {still_needs_review}")

# 3. АСД photo stats
print("\n=== АСД DETAILED CHECK ===")
asd_final = [r for r in final if r["producer"] == "АСД"]
print(f"Total: {len(asd_final)}")
print(f"With image_url_front: {sum(1 for r in asd_final if r.get('image_url_front',''))}")
print(f"With image_url_inside: {sum(1 for r in asd_final if r.get('image_url_inside',''))}")
print(f"Both: {sum(1 for r in asd_final if r.get('image_url_front','') and r.get('image_url_inside',''))}")
print(f"Front only: {sum(1 for r in asd_final if r.get('image_url_front','') and not r.get('image_url_inside',''))}")
print(f"Inside only: {sum(1 for r in asd_final if not r.get('image_url_front','') and r.get('image_url_inside',''))}")
print(f"No photo: {sum(1 for r in asd_final if not r.get('image_url_front','') and not r.get('image_url_inside',''))}")

# Bad patterns
bad = ["cache_image", "fragment", "handle", "ruchka", "lock", "zamok", "interior", "badge", "banner"]
bad_count = 0
for r in asd_final:
    for bp in bad:
        if bp in (r.get("image_url_front","") or "").lower() or bp in (r.get("image_url_inside","") or "").lower():
            bad_count += 1
            break
print(f"Bad patterns in front/inside: {bad_count}")

# 4. Термодур check
print("\n=== ТЕРМОДУР DETAILED CHECK ===")
termo = [r for r in final if r["producer"] == "Термодур"]
print(f"Total: {len(termo)}")
termo_v163 = [r for r in v163 if r["producer"] == "Термодур"]
print(f"v16.3 total: {len(termo_v163)}")
# Check if any rows changed
changed = 0
for r in termo:
    v163_match = [v for v in termo_v163 if v["model_name"] == r["model_name"] and v["target_line"] == r["target_line"]]
    if v163_match:
        if v163_match[0].get("image_url","") != r.get("image_url",""):
            changed += 1
print(f"Rows with changed image_url: {changed}")
print(f"Lines: {len(set(r['target_line'] for r in termo))}")

# 5. Энтро Aria
print("\n=== ЭНТРО ARIA CHECK ===")
aria_v163 = [r for r in v163 if r["producer"] == "Энтро" and "aria" in r.get("target_line","").lower()]
aria_final = [r for r in final if r["producer"] == "Энтро" and "aria" in r.get("target_line","").lower()]
print(f"Aria v16.3: {len(aria_v163)} rows")
print(f"Aria final: {len(aria_final)} rows")
for r in aria_final:
    print(f"  {r['model_name'][:50]} | img={bool(r.get('image_url',''))}")

# 6. Лорд coverage
print("\n=== ЛОРД COVERAGE ===")
lord_expected = [e["line"] for e in reference.get("Лорд", [])]
lord_final_lines = set(r["target_line"] for r in final if r["producer"] == "Лорд")
missing_lord = []
for line in lord_expected:
    stripped = line.strip()
    found = any(stripped.lower() in fl.strip().lower() for fl in lord_final_lines)
    if not found:
        missing_lord.append(line)
        print(f"  MISSING: {line}")
if not missing_lord:
    print("  All 41 expected lines covered")
else:
    print(f"  {len(missing_lord)} lines not matched")

# 7. Хаусдорс coverage
print("\n=== ХАУСДОРС COVERAGE ===")
hd_expected = [e["line"] for e in reference.get("Хаусдорс", [])]
hd_final_lines = set(r["target_line"] for r in final if r["producer"] == "Хаусдорс")
hd_missing = []
for line in hd_expected:
    found = any(line.strip().lower() in fl.strip().lower() for fl in hd_final_lines)
    if not found:
        hd_missing.append(line)
        print(f"  MISSING: {line}")
if not hd_missing:
    print(f"  All {len(hd_expected)} expected lines covered or stubbed")
else:
    print(f"  {len(hd_missing)} lines not found")

# 8. front_inside_photo_stats_by_producer
print("\n=== FRONT/INSIDE PHOTO STATS BY PRODUCER ===")
photo_stats = {}
for producer in set(r["producer"] for r in final):
    prod_rows = [r for r in final if r["producer"] == producer]
    entry_rows = [r for r in prod_rows if "входн" in (r.get("door_type","") or "").lower()]
    photo_stats[producer] = {
        "total": len(prod_rows),
        "entry_doors": len(entry_rows),
        "with_front": sum(1 for r in prod_rows if r.get("image_url_front","")),
        "with_inside": sum(1 for r in prod_rows if r.get("image_url_inside","")),
        "both_sides": sum(1 for r in prod_rows if r.get("image_url_front","") and r.get("image_url_inside","")),
        "front_only": sum(1 for r in prod_rows if r.get("image_url_front","") and not r.get("image_url_inside","")),
        "inside_only": sum(1 for r in prod_rows if not r.get("image_url_front","") and r.get("image_url_inside","")),
        "no_photo": sum(1 for r in prod_rows if not r.get("image_url_front","") and not r.get("image_url_inside",""))
    }
    ps = photo_stats[producer]
    print(f"  {producer}: total={ps['total']} front={ps['with_front']} inside={ps['with_inside']} both={ps['both_sides']} none={ps['no_photo']}")

# 9. Update run_report.json
with open(os.path.join(FINAL_DIR, "run_report.json"), encoding="utf-8") as f:
    report = json.load(f)

report["expected_lines_total"] = 194
report["expected_lines_found"] = sum(1 for r in recon_rows if r["final_status"] == "found")
report["expected_lines_missing_audit"] = audit["missing_line"]
report["expected_lines_missing_final"] = sum(1 for r in recon_rows if r["final_status"] == "missing")
report["patch_missing_entries_total"] = len(model_patches)
report["patch_missing_entries_added"] = added_as_stub
report["patch_missing_entries_skipped"] = skipped_asd
report["patch_missing_entries_mapped_existing"] = already_existed
report["front_inside_photo_stats_by_producer"] = photo_stats
report["asd_detail"] = {
    "total": len(asd_final),
    "with_front": sum(1 for r in asd_final if r.get("image_url_front","")),
    "with_inside": sum(1 for r in asd_final if r.get("image_url_inside","")),
    "both_sides": sum(1 for r in asd_final if r.get("image_url_front","") and r.get("image_url_inside","")),
    "front_only": sum(1 for r in asd_final if r.get("image_url_front","") and not r.get("image_url_inside","")),
    "no_photo": sum(1 for r in asd_final if not r.get("image_url_front","") and not r.get("image_url_inside","")),
    "bad_patterns_in_front_inside": bad_count,
    "note": "АСД reference says 'Все модели' x11 — catalog has 199 rows, considered fully covered"
}
report["thermobreak_detail"] = {
    "total": len(termo),
    "changed_image_url": changed,
    "note": "Old thermobreak data preserved as-is per client instruction"
}
report["lord_detail"] = {
    "expected": len(lord_expected),
    "covered": len(lord_expected) - len(missing_lord),
    "missing": missing_lord
}
report["haushdorss_detail"] = {
    "expected": len(hd_expected),
    "covered": len(hd_expected) - len(hd_missing),
    "missing": hd_missing,
    "stubs_added": sum(1 for r in missing_stubs if r["producer"] == "Хаусдорс")
}

with open(os.path.join(FINAL_DIR, "run_report.json"), "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)
print("\nrun_report.json: updated")
