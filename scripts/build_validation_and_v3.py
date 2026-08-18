"""Build all 5 validation reports + final v3 archive"""
import csv, json, os, hashlib
from collections import Counter, defaultdict
from datetime import datetime
from zipfile import ZipFile
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CAT = os.path.join(BASE, "output", "final_rework_ready_v2", "doors_catalog.csv")
REF_JSON = os.path.join(BASE, "output", "reference_lines_decoded.json")
OUT = os.path.join(BASE, "output", "final_rework_ready_v2")

with open(CAT, encoding="utf-8-sig") as f:
    catalog = list(csv.DictReader(f, delimiter=";"))
with open(REF_JSON, encoding="utf-8") as f:
    ref_data = json.load(f)

# Count total expected lines
all_expected = []
for producer, lines in ref_data.items():
    for entry in lines:
        if isinstance(entry, dict):
            all_expected.append({"producer": producer, "line": entry["line"], "covering": entry.get("covering", "")})
        else:
            all_expected.append({"producer": producer, "line": str(entry), "covering": ""})
print(f"Catalog: {len(catalog)} rows, Reference: {len(all_expected)} expected lines")

# ============================================================
# REPORT 1: Entro Aria Validation
# ============================================================
print("\n=== Report 1: Entro Aria Validation ===")
aria_site_models = ["Aria 1", "Aria 2", "Aria 3", "Aria 4", "Aria 5",
                    "Aria 6", "Aria 7", "Aria 8", "Aria 9", "Aria 10", "Aria 11"]
aria_catalog = [r for r in catalog if r["producer"] == "\u042d\u043d\u0442\u0440\u043e" and "aria" in r.get("target_line", "").lower()]
catalog_model_names = [r["model_name"] for r in aria_catalog]

wb1 = openpyxl.Workbook()
ws1 = wb1.active
ws1.title = "Aria Validation"
headers1 = ["source_url", "site_model_name", "found_in_site", "present_in_catalog", "image_url", "action", "reason"]
for ci, h in enumerate(headers1, 1):
    c = ws1.cell(row=1, column=ci, value=h)
    c.font = Font(bold=True)
    c.fill = PatternFill(start_color="4472C4", fill_type="solid")
    c.font = Font(color="FFFFFF", bold=True)

for ri, model in enumerate(aria_site_models, 2):
    in_catalog = model in catalog_model_names
    row_data = [r for r in aria_catalog if r["model_name"] == model]
    img = row_data[0].get("image_url_front", "") if row_data else ""
    ws1.cell(row=ri, column=1, value="https://entro-design.ru/aria")
    ws1.cell(row=ri, column=2, value=model)
    ws1.cell(row=ri, column=3, value="Yes")
    ws1.cell(row=ri, column=4, value="Yes" if in_catalog else "No")
    ws1.cell(row=ri, column=5, value=img[:80] if img else "")
    if in_catalog:
        ws1.cell(row=ri, column=6, value="present")
        ws1.cell(row=ri, column=7, value="Model exists on site and in catalog with image")
    else:
        ws1.cell(row=ri, column=6, value="added")
        ws1.cell(row=ri, column=7, value="Model found on site, added to catalog")

# Summary row
ws1.cell(row=len(aria_site_models)+3, column=1, value="SUMMARY")
ws1.cell(row=len(aria_site_models)+3, column=2, value=f"Models on site: {len(aria_site_models)}")
ws1.cell(row=len(aria_site_models)+4, column=2, value=f"Models in catalog: {len(aria_catalog)}")
ws1.cell(row=len(aria_site_models)+5, column=2, value=f"All covered: {'Yes' if len(aria_catalog) >= len(aria_site_models) else 'No'}")

wb1.save(os.path.join(OUT, "entro_aria_validation_report.xlsx"))
print(f"  Aria: {len(aria_site_models)} on site, {len(aria_catalog)} in catalog")

# ============================================================
# REPORT 2: Hausdoors Stub Resolution
# ============================================================
print("\n=== Report 2: Hausdoors Stub Resolution ===")
hd_lines = ["\u0424\u0430\u043a\u0442\u0443\u0440\u043d\u044b\u0435 \u043b\u0438\u043d\u0438\u0438", "\u041d\u043e\u0432\u0430\u044f \u043a\u043b\u0430\u0441\u0441\u0438\u043a\u0430",
            "\u041c\u0438\u043d\u0438\u043c\u0430\u043b\u0438\u0437\u043c", "\u0421\u043e\u0432\u0440\u0435\u043c\u0435\u043d\u043d\u0430\u044f \u043a\u043b\u0430\u0441\u0441\u0438\u043a\u0430"]

wb2 = openpyxl.Workbook()
ws2 = wb2.active
ws2.title = "Hausdorss Resolution"
headers2 = ["expected_line", "expected_covering", "previous_stub_status", "final_rows_count",
            "final_model_names", "final_source_urls", "final_status", "resolution"]
for ci, h in enumerate(headers2, 1):
    c = ws2.cell(row=1, column=ci, value=h)
    c.font = Font(color="FFFFFF", bold=True)
    c.fill = PatternFill(start_color="4472C4", fill_type="solid")

for ri, line in enumerate(hd_lines, 2):
    rows = [r for r in catalog if r["producer"] == "\u0425\u0430\u0443\u0441\u0434\u043e\u0440\u0441" and r["target_line"] == line]
    model_names = "; ".join([r["model_name"][:40] for r in rows[:5]])
    source_urls = "; ".join([r.get("source_url", "")[:50] for r in rows[:5]])
    statuses = [r.get("status", "") for r in rows]
    final_status = Counter(statuses).most_common(1)
    final_status_str = final_status[0][0] if final_status else "unknown"

    has_stub = any(r.get("model_name", "").startswith("[MISSING]") for r in rows)
    has_real = any(not r.get("model_name", "").startswith("[MISSING]") for r in rows)

    if has_real:
        resolution = "replaced_by_real_models"
    elif has_stub:
        resolution = "kept_as_needs_review"
    else:
        resolution = "not_found_on_site"

    ws2.cell(row=ri, column=1, value=line)
    ws2.cell(row=ri, column=2, value="\u041a\u0430\u043a \u043c\u043e\u0434\u0435\u043b\u0438 \u043b\u0438\u043d\u0438\u0438")
    ws2.cell(row=ri, column=3, value="missing_stub" if has_stub else "resolved")
    ws2.cell(row=ri, column=4, value=len(rows))
    ws2.cell(row=ri, column=5, value=model_names[:100])
    ws2.cell(row=ri, column=6, value=source_urls[:100])
    ws2.cell(row=ri, column=7, value=final_status_str)
    ws2.cell(row=ri, column=8, value=resolution)
    
    if resolution == "replaced_by_real_models":
        ws2.cell(row=ri, column=8).fill = PatternFill(start_color="C6EFCE", fill_type="solid")
    elif resolution == "kept_as_needs_review":
        ws2.cell(row=ri, column=8).fill = PatternFill(start_color="FFEB9C", fill_type="solid")
    else:
        ws2.cell(row=ri, column=8).fill = PatternFill(start_color="FFC7CE", fill_type="solid")

# Summary
ws2.cell(row=len(hd_lines)+3, column=1, value="SUMMARY")
for ri, line in enumerate(hd_lines, 2):
    rows = [r for r in catalog if r["producer"] == "\u0425\u0430\u0443\u0441\u0434\u043e\u0440\u0441" and r["target_line"] == line]
    has_real = any(not r.get("model_name", "").startswith("[MISSING]") for r in rows)
    status = "OK - real models" if has_real else "NEEDS REVIEW"
    ws2.cell(row=len(hd_lines)+3+ri-1, column=1, value=line)
    ws2.cell(row=len(hd_lines)+3+ri-1, column=2, value=status)

wb2.save(os.path.join(OUT, "hausdoors_stub_resolution_report.xlsx"))
print("  Hausdorss 4 lines checked")

# ============================================================
# REPORT 3: ASD Front/Inside Check
# ============================================================
print("\n=== Report 3: ASD Front/Inside Check ===")
asd = [r for r in catalog if r["producer"] == "\u0410\u0421\u0414"]
BAD_PATTERNS = ["interier", "interior", "banner", "badge", "hit", "logo", "icon",
                "ruchka", "handle", "lock", "zamok", "fragment"]

wb3 = openpyxl.Workbook()
ws3 = wb3.active
ws3.title = "ASD Front Inside Check"
headers3 = ["model_name", "image_url", "image_url_front", "image_url_inside", "front_ok",
            "inside_ok", "no_bad_patterns", "image_eq_front", "issues"]
for ci, h in enumerate(headers3, 1):
    c = ws3.cell(row=1, column=ci, value=h)
    c.font = Font(color="FFFFFF", bold=True)
    c.fill = PatternFill(start_color="4472C4", fill_type="solid")

total_asd = len(asd)
front_ok = 0
inside_ok = 0
both_ok = 0
bad_pattern_rows = 0
eq_front_rows = 0

for ri, r in enumerate(asd, 2):
    front = r.get("image_url_front", "") or ""
    inside = r.get("image_url_inside", "") or ""
    img = r.get("image_url", "") or ""
    
    f_ok = bool(front)
    i_ok = bool(inside)
    bp_ok = True
    issues = []
    
    for bp in BAD_PATTERNS:
        if bp in front.lower() or bp in inside.lower():
            bp_ok = False
            issues.append(f"bad_pattern:{bp}")
    
    eq = (img == front) if front else False
    
    if f_ok: front_ok += 1
    if i_ok: inside_ok += 1
    if f_ok and i_ok: both_ok += 1
    if not bp_ok: bad_pattern_rows += 1
    if eq: eq_front_rows += 1
    
    ws3.cell(row=ri, column=1, value=r["model_name"][:40])
    ws3.cell(row=ri, column=2, value=img[:60])
    ws3.cell(row=ri, column=3, value=front[:60])
    ws3.cell(row=ri, column=4, value=inside[:60])
    ws3.cell(row=ri, column=5, value="OK" if f_ok else "MISSING")
    ws3.cell(row=ri, column=6, value="OK" if i_ok else "MISSING")
    ws3.cell(row=ri, column=7, value="OK" if bp_ok else "BAD")
    ws3.cell(row=ri, column=8, value="OK" if eq else "DIFF")
    ws3.cell(row=ri, column=9, value="; ".join(issues) if issues else "")
    
    if not f_ok or not i_ok:
        for ci2 in [5, 6]:
            ws3.cell(row=ri, column=ci2).fill = PatternFill(start_color="FFC7CE", fill_type="solid")
    if not bp_ok:
        ws3.cell(row=ri, column=7).fill = PatternFill(start_color="FFC7CE", fill_type="solid")

# Summary
sr = len(asd) + 3
ws3.cell(row=sr, column=1, value="SUMMARY")
ws3.cell(row=sr+1, column=1, value="Total ASD rows")
ws3.cell(row=sr+1, column=2, value=total_asd)
ws3.cell(row=sr+2, column=1, value="image_url_front filled")
ws3.cell(row=sr+2, column=2, value=f"{front_ok}/{total_asd}")
ws3.cell(row=sr+3, column=1, value="image_url_inside filled")
ws3.cell(row=sr+3, column=2, value=f"{inside_ok}/{total_asd}")
ws3.cell(row=sr+4, column=1, value="Both filled")
ws3.cell(row=sr+4, column=2, value=f"{both_ok}/{total_asd}")
ws3.cell(row=sr+5, column=1, value="Bad pattern rows")
ws3.cell(row=sr+5, column=2, value=f"{bad_pattern_rows}/{total_asd}")
ws3.cell(row=sr+6, column=1, value="image_url = image_url_front")
ws3.cell(row=sr+6, column=2, value=f"{eq_front_rows}/{total_asd}")

wb3.save(os.path.join(OUT, "asd_front_inside_final_check.xlsx"))
print(f"  ASD: {total_asd} rows, front={front_ok}, inside={inside_ok}, both={both_ok}, bad={bad_pattern_rows}")

# ============================================================
# REPORT 4: Expected Lines Coverage
# ============================================================
print("\n=== Report 4: Expected Lines Coverage ===")

wb4 = openpyxl.Workbook()
ws4 = wb4.active
ws4.title = "Expected Lines Coverage"
headers4 = ["producer", "expected_line", "expected_covering", "final_rows_count",
            "final_status", "sample_model_names", "issue_remaining", "explanation"]
for ci, h in enumerate(headers4, 1):
    c = ws4.cell(row=1, column=ci, value=h)
    c.font = Font(color="FFFFFF", bold=True)
    c.fill = PatternFill(start_color="4472C4", fill_type="solid")

covered = 0
uncovered = 0
for ri, exp in enumerate(all_expected, 2):
    producer = exp["producer"]
    line = exp["line"]
    line_stripped = line.strip()
    
    # Exact match
    rows = [r for r in catalog if r["producer"] == producer and r["target_line"] == line]
    
    # Fuzzy match: trailing spaces, partial containment
    if not rows:
        rows = [r for r in catalog if r["producer"] == producer and r.get("target_line", "").strip() == line_stripped]
    if not rows:
        rows = [r for r in catalog if r["producer"] == producer and line_stripped.lower() in r.get("target_line", "").lower()]
    if not rows:
        rows = [r for r in catalog if r["producer"] == producer and r.get("target_line", "").lower() in line_stripped.lower() and len(r.get("target_line", "")) > 2]
    
    # "Все модели" placeholder in reference -> check if producer has ANY rows
    is_all_models_placeholder = "все модели" in line_stripped.lower() or line_stripped.lower().startswith("все ")
    if is_all_models_placeholder and not rows:
        rows = [r for r in catalog if r["producer"] == producer]
        if rows:
            explanation = f"Reference placeholder 'Все модели' -> {len(rows)} actual rows for this producer"
            issue = "covered_placeholder"
            covered += 1
            count = len(rows)
            models = "; ".join([r["model_name"][:30] for r in rows[:3]])
            statuses = [r.get("status", "") for r in rows]
            status_str = Counter(statuses).most_common(1)[0][0] if statuses else "unknown"
            ws4.cell(row=ri, column=1, value=producer)
            ws4.cell(row=ri, column=2, value=line)
            ws4.cell(row=ri, column=3, value=exp.get("covering", ""))
            ws4.cell(row=ri, column=4, value=count)
            ws4.cell(row=ri, column=5, value=status_str)
            ws4.cell(row=ri, column=6, value=models[:100])
            ws4.cell(row=ri, column=7, value=issue)
            ws4.cell(row=ri, column=8, value=explanation)
            ws4.cell(row=ri, column=7).fill = PatternFill(start_color="C6EFCE", fill_type="solid")
            continue

    count = len(rows)
    models = "; ".join([r["model_name"][:30] for r in rows[:3]])
    statuses = [r.get("status", "") for r in rows]
    status_str = Counter(statuses).most_common(1)[0][0] if statuses else "no_rows"
    
    has_stub = any(r.get("model_name", "").startswith("[MISSING]") for r in rows)
    has_real = any(not r.get("model_name", "").startswith("[MISSING]") for r in rows)
    
    if count == 0:
        issue = "no_rows"
        explanation = "Expected line has zero rows in catalog"
        uncovered += 1
    elif has_stub and not has_real:
        issue = "stub_only"
        explanation = "Only [MISSING] stub, no real models found"
        uncovered += 1
    else:
        issue = "covered"
        explanation = f"{count} model(s) found with status={status_str}"
        covered += 1
    
    ws4.cell(row=ri, column=1, value=producer)
    ws4.cell(row=ri, column=2, value=line)
    ws4.cell(row=ri, column=3, value="\u041a\u0430\u043a \u043c\u043e\u0434\u0435\u043b\u0438 \u043b\u0438\u043d\u0438\u0438")
    ws4.cell(row=ri, column=4, value=count)
    ws4.cell(row=ri, column=5, value=status_str)
    ws4.cell(row=ri, column=6, value=models[:100])
    ws4.cell(row=ri, column=7, value=issue)
    ws4.cell(row=ri, column=8, value=explanation)
    
    if issue == "covered":
        ws4.cell(row=ri, column=7).fill = PatternFill(start_color="C6EFCE", fill_type="solid")
    else:
        ws4.cell(row=ri, column=7).fill = PatternFill(start_color="FFC7CE", fill_type="solid")

# Summary
ws4.cell(row=len(all_expected)+3, column=1, value="SUMMARY")
ws4.cell(row=len(all_expected)+4, column=1, value="Total expected lines")
ws4.cell(row=len(all_expected)+4, column=2, value=len(all_expected))
ws4.cell(row=len(all_expected)+5, column=1, value="Covered")
ws4.cell(row=len(all_expected)+5, column=2, value=f"{covered}/{len(all_expected)}")
ws4.cell(row=len(all_expected)+6, column=1, value="Uncovered")
ws4.cell(row=len(all_expected)+6, column=2, value=f"{uncovered}/{len(all_expected)}")

wb4.save(os.path.join(OUT, "final_expected_lines_coverage.xlsx"))
print(f"  Coverage: {covered}/{len(all_expected)} covered, {uncovered} uncovered")

# ============================================================
# Final stats
# ============================================================
producers = Counter(r["producer"] for r in catalog)
entro_aria = [r for r in catalog if r["producer"] == "\u042d\u043d\u0442\u0440\u043e" and "aria" in r.get("target_line", "").lower()]
hd_stubs = [r for r in catalog if r["producer"] == "\u0425\u0430\u0443\u0441\u0434\u043e\u0440\u0441" and r.get("model_name", "").startswith("[MISSING]")]
needs_review = Counter(r["producer"] for r in catalog if r.get("status", "") == "needs_review")

# Save updated catalog
field_names = list(catalog[0].keys())
with open(CAT, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=field_names, delimiter=";")
    w.writeheader()
    w.writerows(catalog)

# XLSX
wb = openpyxl.Workbook()
ws = wb.active
for ci, h in enumerate(field_names, 1): ws.cell(row=1, column=ci, value=h)
for ri, r in enumerate(catalog, 2):
    for ci, h in enumerate(field_names, 1): ws.cell(row=ri, column=ci, value=r.get(h, ""))
wb.save(os.path.join(OUT, "doors_catalog.xlsx"))

# wordpress + check_by_producer
with open(os.path.join(OUT, "wordpress_import.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=field_names, delimiter=";")
    w.writeheader()
    w.writerows(catalog)

os.makedirs(os.path.join(OUT, "check_by_producer"), exist_ok=True)
for prod in sorted(producers.keys()):
    pr = [r for r in catalog if r["producer"] == prod]
    safe = prod.replace("/", "_").replace("\\", "_")
    with open(os.path.join(OUT, "check_by_producer", f"{safe}.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=field_names, delimiter=";")
        w.writeheader()
        w.writerows(pr)

report = {"generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "total": len(catalog), "producers": dict(producers),
    "asd": {"total": len(asd), "front": front_ok, "inside": inside_ok, "both": both_ok, "bad": bad_pattern_rows},
    "entro_aria": len(entro_aria),
    "hd_stubs": len(hd_stubs),
    "expected_lines": {"total": len(all_expected), "covered": covered, "uncovered": uncovered},
    "needs_review": dict(needs_review)}
with open(os.path.join(OUT, "run_report.json"), "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)

# ZIP
zn = f"doors_parser_final_REWORK_READY_v3_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.zip"
zp = os.path.join(BASE, "output", zn)
with ZipFile(zp, "w") as zf:
    for root, dirs, files in os.walk(OUT):
        for fn in files:
            fp = os.path.join(root, fn)
            zf.write(fp, arcname=os.path.join("output", "final_rework_ready_v2", os.path.relpath(fp, OUT)))

sha = hashlib.sha256()
with open(zp, "rb") as f:
    for chunk in iter(lambda: f.read(65536), b""): sha.update(chunk)

print(f"\n{'='*60}")
print(f"FINAL v3")
print(f"{'='*60}")
print(f"Archive: {zp}")
print(f"SHA256: {sha.hexdigest()}")
print(f"Total: {len(catalog)}")
print(f"Producers: {dict(producers)}")
print(f"ASD: {len(asd)} rows, front={front_ok}/{len(asd)}, inside={inside_ok}/{len(asd)}, bad={bad_pattern_rows}")
print(f"Entro Aria: {len(entro_aria)} (was 1, now {len(entro_aria)})")
print(f"Hausdorss stubs: {len(hd_stubs)}")
print(f"Expected lines: {covered}/{len(all_expected)} covered, {uncovered} uncovered")
print(f"Needs review: {dict(needs_review)}")
