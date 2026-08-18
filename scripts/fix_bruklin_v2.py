"""Final fix: Bruklin correct URL + recount bad"""
import csv, os

CAT = r"C:\Users\GN\Desktop\doors_parser_ready_fix\output\final_rework_ready_v2\doors_catalog.csv"

with open(CAT, encoding="utf-8-sig") as f:
    catalog = list(csv.DictReader(f, delimiter=";"))

FIELD_NAMES = list(catalog[0].keys())
BAD_PATTERNS = ["cache_image", "fragment", "handle", "ruchka", "lock", "zamok",
                "interior", "badge", "banner", "logo", "404"]

# Fix Bruklin: find the one with cache_image
for r in catalog:
    front = r.get("image_url_front", "") or ""
    if "cache_image" in front.lower() and "\u0411\u0440\u0443\u043a\u043b\u0438\u043d" in r.get("model_name", ""):
        print(f"Found Bruklin with cache_image: {front[:60]}")
        # Use the gallery image we found earlier
        new_front = "https://as-doors.ru/img/2025/bruklin/bruklin-2-25-new14.jpg"
        new_inside = "https://as-doors.ru/img/2025/bruklin/bruklin-2-25-new1.jpg"
        r["image_url"] = new_front
        r["image_url_front"] = new_front
        r["image_url_inside"] = new_inside
        r["source_url"] = "https://as-doors.ru/doors/zakaznyie/bruklin.html"
        r["status"] = "found"
        print(f"  FIXED -> front={new_front[:50]}, inside={new_inside[:50]}")
        break

# Recount bad patterns (per-row, not per-pattern)
asd = [r for r in catalog if r["producer"] == "\u0410\u0421\u0414"]
bad_rows = 0
for r in asd:
    front = (r.get("image_url_front", "") or "").lower()
    inside = (r.get("image_url_inside", "") or "").lower()
    for bp in BAD_PATTERNS:
        if bp in front or bp in inside:
            bad_rows += 1
            print(f"  Still bad: {r['model_name'][:30]} bp={bp}")
            break

print(f"\nBad rows: {bad_rows}")

# Save
with open(CAT, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=FIELD_NAMES, delimiter=";")
    w.writeheader()
    w.writerows(catalog)

print("Saved. Re-check:")
with open(CAT, encoding="utf-8-sig") as f:
    catalog2 = list(csv.DictReader(f, delimiter=";"))
asd2 = [r for r in catalog2 if r["producer"] == "\u0410\u0421\u0414"]
bad2 = 0
for r in asd2:
    front = (r.get("image_url_front", "") or "").lower()
    inside = (r.get("image_url_inside", "") or "").lower()
    for bp in BAD_PATTERNS:
        if bp in front or bp in inside:
            bad2 += 1
            break
print(f"After save: ASD={len(asd2)}, front={sum(1 for r in asd2 if r.get('image_url_front',''))}, inside={sum(1 for r in asd2 if r.get('image_url_inside',''))}, bad={bad2}")
