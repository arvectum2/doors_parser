"""Stage 1: Audit catalog against reference xlsx"""
import csv, json, os, re
from collections import Counter, defaultdict

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CATALOG_CSV = os.path.join(BASE, "output", "final_hybrid_client_ready_v16_3_asd_full_door_only", "doors_catalog.csv")
REFERENCE_JSON = os.path.join(BASE, "output", "reference_lines_decoded.json")
OUT_DIR = os.path.join(BASE, "output", "final_rework_audit")
os.makedirs(OUT_DIR, exist_ok=True)

# Load reference
with open(REFERENCE_JSON, encoding="utf-8") as f:
    reference = json.load(f)

# Load catalog
with open(CATALOG_CSV, encoding="utf-8-sig") as f:
    catalog = list(csv.DictReader(f, delimiter=";"))

print(f"Reference producers: {len(reference)}")
print(f"Catalog rows: {len(catalog)}")

# Build catalog index: producer -> target_line -> list of rows
cat_index = defaultdict(lambda: defaultdict(list))
for r in catalog:
    cat_index[r["producer"]][r["target_line"]].append(r)

# Normalize line name for matching
def norm_line(s):
    return re.sub(r'\s+', ' ', s.strip().lower())

# Build reference -> catalog line mapping
# Try exact match first, then fuzzy
def find_catalog_line(producer, ref_line, cat_lines):
    """Find matching catalog target_line for a reference line"""
    rn = norm_line(ref_line)
    # Exact match
    for cl in cat_lines:
        if norm_line(cl) == rn:
            return cl
    # Partial match (ref in catalog or catalog in ref)
    for cl in cat_lines:
        cn = norm_line(cl)
        if rn in cn or cn in rn:
            return cl
    # Word overlap
    ref_words = set(rn.split())
    best_score = 0
    best_match = None
    for cl in cat_lines:
        cn = norm_line(cl)
        cat_words = set(cn.split())
        overlap = len(ref_words & cat_words)
        if overlap > best_score:
            best_score = overlap
            best_match = cl
    if best_score > 0:
        return best_match
    return None

# Determine if a line is entry door based on door_type in catalog
def is_entry_line(producer, line_rows):
    types = set(r.get("door_type","") for r in line_rows)
    return any("входн" in t.lower() for t in types)

# Audit each producer
audit_rows = []
photo_rows = []
summary = {
    "total_expected_lines": 0,
    "total_in_catalog": 0,
    "missing_line": 0,
    "missing_models": 0,
    "wrong_door_type": 0,
    "missing_front_photo": 0,
    "missing_inside_photo": 0,
    "needs_manual_review": 0,
    "by_producer": {}
}

for producer, ref_lines in reference.items():
    cat_lines = list(cat_index.get(producer, {}).keys())
    producer_summary = {
        "expected": len(ref_lines),
        "in_catalog": 0,
        "missing_line": 0,
        "missing_models": 0,
        "wrong_door_type": 0,
        "missing_front_photo": 0,
        "missing_inside_photo": 0,
        "lines_detail": []
    }
    
    for ref_entry in ref_lines:
        ref_line = ref_entry["line"]
        ref_covering = ref_entry["covering"]
        summary["total_expected_lines"] += 1
        
        matched_cat_line = find_catalog_line(producer, ref_line, cat_lines)
        
        if matched_cat_line:
            rows = cat_index[producer][matched_cat_line]
            model_count = len(rows)
            has_image = sum(1 for r in rows if r.get("image_url",""))
            entry_door = is_entry_line(producer, rows)
            
            issue = "ok"
            action = "none"
            notes = ""
            
            if model_count == 0:
                issue = "missing_models"
                action = "add_models_from_site"
                summary["missing_models"] += 1
                producer_summary["missing_models"] += 1
            elif model_count == 1 and producer not in ("АСД", "Термодур"):
                # Check if site has more models
                issue = "missing_models"
                action = "check_site_for_more_models"
                summary["missing_models"] += 1
                producer_summary["missing_models"] += 1
                notes = "single_model_check_if_more_on_site"
            
            # Photo checks
            missing_front = 0
            missing_inside = 0
            for r in rows:
                img = r.get("image_url","") or ""
                if not img:
                    missing_front += 1
            
            if entry_door:
                missing_inside = sum(1 for r in rows if not r.get("image_url_inside",""))
                summary["missing_inside_photo"] += missing_inside
                producer_summary["missing_inside_photo"] += missing_inside
            
            summary["missing_front_photo"] += missing_front
            producer_summary["missing_front_photo"] += missing_front
            
            producer_summary["in_catalog"] += 1
            summary["total_in_catalog"] += 1
            
            producer_summary["lines_detail"].append({
                "ref_line": ref_line,
                "ref_covering": ref_covering,
                "matched_catalog_line": matched_cat_line,
                "model_count": model_count,
                "has_image_count": has_image,
                "is_entry_door": entry_door,
                "issue": issue,
                "action": action,
                "notes": notes
            })
            
            audit_rows.append({
                "producer": producer,
                "expected_line": ref_line,
                "expected_covering": ref_covering,
                "exists_in_current_catalog": "yes",
                "current_rows_count": model_count,
                "found_on_site": "yes" if model_count > 0 else "no",
                "found_site_models_count": model_count,
                "issue_type": issue,
                "suggested_action": action,
                "source_reference": "Линейки производителей 3.xlsx",
                "notes": notes
            })
        else:
            # Missing line entirely
            issue = "missing_line"
            action = "add_line_and_models_from_site"
            summary["missing_line"] += 1
            producer_summary["missing_line"] += 1
            
            producer_summary["lines_detail"].append({
                "ref_line": ref_line,
                "ref_covering": ref_covering,
                "matched_catalog_line": None,
                "model_count": 0,
                "has_image_count": 0,
                "is_entry_door": False,
                "issue": issue,
                "action": action,
                "notes": "line_not_found_in_catalog"
            })
            
            audit_rows.append({
                "producer": producer,
                "expected_line": ref_line,
                "expected_covering": ref_covering,
                "exists_in_current_catalog": "no",
                "current_rows_count": 0,
                "found_on_site": "unknown",
                "found_site_models_count": 0,
                "issue_type": issue,
                "suggested_action": action,
                "source_reference": "Линейки производителей 3.xlsx",
                "notes": "line_not_found_in_catalog"
            })
    
    summary["by_producer"][producer] = producer_summary

# Photo audit for all catalog rows
print("\n=== PHOTO AUDIT ===")
for r in catalog:
    producer = r["producer"]
    img = r.get("image_url","") or ""
    img_front = r.get("image_url_front","") or ""
    img_inside = r.get("image_url_inside","") or ""
    
    # Detect issues
    photo_issue = "ok"
    notes = ""
    
    if not img and not img_front:
        photo_issue = "missing_front"
        notes = "no_image_at_all"
    elif not img_front and img:
        # Has image_url but no front designation
        photo_issue = "needs_manual_review"
        notes = "has_image_url_but_no_front_inside_designation"
    elif img_front and not img_inside:
        photo_issue = "only_front_available"
        notes = "inside_missing"
    
    # Check for fragment/technical patterns in URL
    bad_patterns = ["cache_image", "fragment", "handle", "ruchka", "lock", "zamok", 
                    "interior", "badge", "banner", "logo", "404"]
    for bp in bad_patterns:
        if bp in img.lower():
            photo_issue = "wrong_photo_kind"
            notes = f"bad_pattern_{bp}_in_image_url"
            break
    
    photo_rows.append({
        "producer": producer,
        "target_line": r.get("target_line",""),
        "model_name": r.get("model_name",""),
        "source_url": r.get("source_url",""),
        "current_image_url": img,
        "current_image_url_front": img_front,
        "current_image_url_inside": img_inside,
        "detected_front_candidate": "",
        "detected_inside_candidate": "",
        "front_decision": "",
        "inside_decision": "",
        "rejected_technical_count": 0,
        "rejected_fragment_count": 0,
        "rejected_interior_count": 0,
        "photo_issue_type": photo_issue,
        "notes": notes
    })

# Write outputs
print("\n=== WRITING OUTPUTS ===")

# audit_missing_lines_models.csv
with open(os.path.join(OUT_DIR, "audit_missing_lines_models.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=["producer","expected_line","expected_covering",
        "exists_in_current_catalog","current_rows_count","found_on_site",
        "found_site_models_count","issue_type","suggested_action","source_reference","notes"])
    w.writeheader()
    w.writerows(audit_rows)
print(f"audit_missing_lines_models.csv: {len(audit_rows)} rows")

# photo_audit.csv
with open(os.path.join(OUT_DIR, "photo_audit.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=["producer","target_line","model_name","source_url",
        "current_image_url","current_image_url_front","current_image_url_inside",
        "detected_front_candidate","detected_inside_candidate","front_decision","inside_decision",
        "rejected_technical_count","rejected_fragment_count","rejected_interior_count",
        "photo_issue_type","notes"])
    w.writeheader()
    w.writerows(photo_rows)
print(f"photo_audit.csv: {len(photo_rows)} rows")

# expected_lines_from_xlsx.csv
with open(os.path.join(OUT_DIR, "expected_lines_from_xlsx.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=["producer","line","covering"])
    w.writeheader()
    for producer, ref_lines in reference.items():
        for entry in ref_lines:
            w.writerow({"producer": producer, "line": entry["line"], "covering": entry["covering"]})
print(f"expected_lines_from_xlsx.csv: {sum(len(v) for v in reference.values())} rows")

# current_vs_expected_summary.json
with open(os.path.join(OUT_DIR, "current_vs_expected_summary.json"), "w", encoding="utf-8") as f:
    json.dump(summary, f, ensure_ascii=False, indent=2)
print("current_vs_expected_summary.json: saved")

# manual_questions_for_client.md
questions = []
for producer, ps in summary["by_producer"].items():
    if ps["missing_line"] > 0:
        missing = [d["ref_line"] for d in ps["lines_detail"] if d["issue"] == "missing_line"]
        questions.append(f"### {producer}: {ps['missing_line']} отсутствующих линеек\n"
                        f"Линейки: {', '.join(missing)}")
    if ps["missing_models"] > 0:
        single = [d["ref_line"] for d in ps["lines_detail"] 
                 if d["issue"] == "missing_models" and d["model_count"] == 1]
        if single:
            questions.append(f"### {producer}: линейки с 1 моделью (проверить, есть ли ещё на сайте)\n"
                           f"Линейки: {', '.join(single)}")

with open(os.path.join(OUT_DIR, "manual_questions_for_client.md"), "w", encoding="utf-8") as f:
    f.write("# Вопросы к заказчику\n\n")
    f.write(f"Сформированы на основе аудита каталога v16.3 по эталонному файлу.\n\n")
    f.write(f"## Сводка\n\n")
    f.write(f"- Всего обязательных линеек в xlsx: {summary['total_expected_lines']}\n")
    f.write(f"- Есть в текущем каталоге: {summary['total_in_catalog']}\n")
    f.write(f"- Отсутствующие линейки (missing_line): {summary['missing_line']}\n")
    f.write(f"- Линейки с малым числом моделей (missing_models): {summary['missing_models']}\n")
    f.write(f"- Входных дверей без front фото: {summary['missing_front_photo']}\n")
    f.write(f"- Входных дверей без inside фото: {summary['missing_inside_photo']}\n\n")
    for q in questions:
        f.write(q + "\n\n")
print("manual_questions_for_client.md: saved")

# XLSX versions
try:
    import openpyxl
    
    # audit_missing_lines_models.xlsx
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Audit Missing Lines"
    fields = ["producer","expected_line","expected_covering","exists_in_current_catalog",
              "current_rows_count","found_on_site","found_site_models_count","issue_type",
              "suggested_action","source_reference","notes"]
    for ci, h in enumerate(fields, 1):
        ws.cell(row=1, column=ci, value=h)
    for ri, row in enumerate(audit_rows, 2):
        for ci, h in enumerate(fields, 1):
            ws.cell(row=ri, column=ci, value=row.get(h, ""))
    wb.save(os.path.join(OUT_DIR, "audit_missing_lines_models.xlsx"))
    print("audit_missing_lines_models.xlsx: saved")
    
    # photo_audit.xlsx
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Photo Audit"
    pfields = ["producer","target_line","model_name","source_url","current_image_url",
               "current_image_url_front","current_image_url_inside","detected_front_candidate",
               "detected_inside_candidate","front_decision","inside_decision",
               "rejected_technical_count","rejected_fragment_count","rejected_interior_count",
               "photo_issue_type","notes"]
    for ci, h in enumerate(pfields, 1):
        ws.cell(row=1, column=ci, value=h)
    for ri, row in enumerate(photo_rows, 2):
        for ci, h in enumerate(pfields, 1):
            ws.cell(row=ri, column=ci, value=row.get(h, ""))
    wb.save(os.path.join(OUT_DIR, "photo_audit.xlsx"))
    print("photo_audit.xlsx: saved")
except ImportError:
    print("openpyxl not available, skipping xlsx")

# Print summary
print(f"\n{'='*60}")
print(f"AUDIT SUMMARY")
print(f"{'='*60}")
print(f"Expected lines in xlsx: {summary['total_expected_lines']}")
print(f"In current catalog: {summary['total_in_catalog']}")
print(f"Missing line: {summary['missing_line']}")
print(f"Missing models: {summary['missing_models']}")
print(f"Wrong door type: {summary['wrong_door_type']}")
print(f"Missing front photo: {summary['missing_front_photo']}")
print(f"Missing inside photo: {summary['missing_inside_photo']}")
print(f"\nBy producer:")
for p, s in sorted(summary["by_producer"].items()):
    issues = []
    if s["missing_line"]: issues.append(f"{s['missing_line']} missing_lines")
    if s["missing_models"]: issues.append(f"{s['missing_models']} missing_models")
    if s["missing_front_photo"]: issues.append(f"{s['missing_front_photo']} no_front")
    if s["missing_inside_photo"]: issues.append(f"{s['missing_inside_photo']} no_inside")
    status = "; ".join(issues) if issues else "OK"
    print(f"  {p}: {s['expected']} expected, {s['in_catalog']} in catalog — {status}")
