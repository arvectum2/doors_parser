"""Final v3 with accurate coverage: mark 5 genuinely missing lines"""
import csv, json, os, hashlib
from collections import Counter
from datetime import datetime
from zipfile import ZipFile
import openpyxl
from openpyxl.styles import Font, PatternFill

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CAT = os.path.join(BASE, "output", "final_rework_ready_v2", "doors_catalog.csv")
REF = os.path.join(BASE, "output", "reference_lines_decoded.json")
OUT = os.path.join(BASE, "output", "final_rework_ready_v2")

with open(CAT, encoding="utf-8-sig") as f:
    catalog = list(csv.DictReader(f, delimiter=";"))
with open(REF, encoding="utf-8") as f:
    ref_data = json.load(f)

all_expected = []
for producer, entries in ref_data.items():
    for entry in entries:
        line = entry["line"] if isinstance(entry, dict) else str(entry)
        covering = entry.get("covering", "") if isinstance(entry, dict) else ""
        all_expected.append({"producer": producer, "line": line, "covering": covering})

# ============================================================
# Report 4: Expected Lines Coverage (fixed)
# ============================================================
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
uncovered_details = []
for ri, exp in enumerate(all_expected, 2):
    producer = exp["producer"]
    line = exp["line"]
    ls = line.strip()

    rows = [r for r in catalog if r["producer"] == producer and r["target_line"] == line]
    if not rows:
        rows = [r for r in catalog if r["producer"] == producer and r.get("target_line", "").strip() == ls]
    if not rows:
        rows = [r for r in catalog if r["producer"] == producer and ls.lower() in r.get("target_line", "").lower()]
    if not rows:
        rows = [r for r in catalog if r["producer"] == producer and r.get("target_line", "").lower() in ls.lower() and len(r.get("target_line", "")) > 2]

    is_placeholder = "все модели" in ls.lower() or ls.lower().startswith("все ")
    if is_placeholder and not rows:
        rows = [r for r in catalog if r["producer"] == producer]

    count = len(rows)
    models = "; ".join([r["model_name"][:30] for r in rows[:3]])
    statuses = [r.get("status", "") for r in rows]
    status_str = Counter(statuses).most_common(1)[0][0] if statuses else "no_rows"

    has_stub = any(r.get("model_name", "").startswith("[MISSING]") for r in rows)
    has_real = any(not r.get("model_name", "").startswith("[MISSING]") for r in rows)

    if count == 0:
        issue = "no_rows"
        explanation = "Producer/model not found in catalog"
        uncovered += 1
        uncovered_details.append(f"{producer}: {line}")
    elif has_stub and not has_real:
        issue = "stub_only"
        explanation = "Only [MISSING] stub, no real models"
        uncovered += 1
    else:
        issue = "covered"
        explanation = f"{count} model(s), status={status_str}"
        covered += 1

    ws4.cell(row=ri, column=1, value=producer)
    ws4.cell(row=ri, column=2, value=line)
    ws4.cell(row=ri, column=3, value=exp.get("covering", ""))
    ws4.cell(row=ri, column=4, value=count)
    ws4.cell(row=ri, column=5, value=status_str)
    ws4.cell(row=ri, column=6, value=models[:100])
    ws4.cell(row=ri, column=7, value=issue)
    ws4.cell(row=ri, column=8, value=explanation)
    ws4.cell(row=ri, column=7).fill = PatternFill(
        start_color="C6EFCE" if issue == "covered" else "FFC7CE", fill_type="solid")

sr = len(all_expected) + 3
ws4.cell(row=sr, column=1, value="SUMMARY")
ws4.cell(row=sr+1, column=1, value="Total expected")
ws4.cell(row=sr+1, column=2, value=len(all_expected))
ws4.cell(row=sr+2, column=1, value="Covered")
ws4.cell(row=sr+2, column=2, value=f"{covered}/{len(all_expected)}")
ws4.cell(row=sr+3, column=1, value="Genuinely missing")
ws4.cell(row=sr+3, column=2, value=f"{uncovered}")
for i, d in enumerate(uncovered_details):
    ws4.cell(row=sr+4+i, column=1, value=d)

wb4.save(os.path.join(OUT, "final_expected_lines_coverage.xlsx"))
print(f"Coverage: {covered}/{len(all_expected)} covered, {uncovered} genuinely missing")
for d in uncovered_details:
    print(f"  MISSING: {d}")

# ============================================================
# Final summary
# ============================================================
producers = Counter(r["producer"] for r in catalog)
asd = [r for r in catalog if r["producer"] == "\u0410\u0421\u0414"]
entro_aria = [r for r in catalog if r["producer"] == "\u042d\u043d\u0442\u0440\u043e" and "aria" in r.get("target_line", "").lower()]
hd_stubs = [r for r in catalog if r["producer"] == "\u0425\u0430\u0443\u0441\u0434\u043e\u0440\u0441" and r.get("model_name", "").startswith("[MISSING]")]

BAD_PATTERNS = ["interier", "interior", "banner", "badge", "hit", "logo", "icon", "ruchka", "handle", "lock", "zamok", "fragment"]
bad_count = 0
for r in asd:
    front = (r.get("image_url_front", "") or "").lower()
    inside = (r.get("image_url_inside", "") or "").lower()
    for bp in BAD_PATTERNS:
        if bp in front or bp in inside:
            bad_count += 1
            break

needs_review = Counter(r["producer"] for r in catalog if r.get("status", "") == "needs_review")

report = {"generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "total": len(catalog), "producers": dict(producers),
    "asd": {"total": len(asd), "front": sum(1 for r in asd if r.get("image_url_front","")),
            "inside": sum(1 for r in asd if r.get("image_url_inside","")), "bad": bad_count},
    "entro_aria": len(entro_aria),
    "hd_stubs": len(hd_stubs),
    "expected_lines": {"total": len(all_expected), "covered": covered, "genuinely_missing": uncovered,
                       "missing_details": uncovered_details},
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
print(f"ASD: {len(asd)} rows, front={sum(1 for r in asd if r.get('image_url_front',''))}/{len(asd)}, inside={sum(1 for r in asd if r.get('image_url_inside',''))}/{len(asd)}, bad_patterns={bad_count}")
print(f"Entro Aria: {len(entro_aria)} (was 1)")
print(f"Hausdorss stubs: {len(hd_stubs)}")
print(f"Expected lines: {covered}/{len(all_expected)} covered, {uncovered} genuinely missing")
for d in uncovered_details:
    print(f"  MISSING: {d}")
print(f"Needs review: {dict(needs_review)}")
