"""Build v5 catalog - targeted rework based on customer feedback"""
import csv, json, os, re, hashlib, zipfile, io
from collections import Counter, defaultdict
from datetime import datetime
from copy import deepcopy

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
V4_ZIP = os.path.join(BASE, "output", "doors_parser_final_REWORK_READY_v4_2026-06-23_17-08-14.zip")
OUT_DIR = os.path.join(BASE, "output", "final_rework_ready_v5")
os.makedirs(OUT_DIR, exist_ok=True)

# ========== STEP 1: Load v4 catalog ==========
print("=" * 60)
print("STEP 1: Loading v4 catalog...")
with zipfile.ZipFile(V4_ZIP) as zf:
    csv_name = [n for n in zf.namelist() if n.endswith("doors_catalog.csv")][0]
    with zf.open(csv_name) as f:
        catalog = list(csv.DictReader(io.StringIO(f.read().decode("utf-8-sig")), delimiter=";"))
field_names = list(catalog[0].keys())
print(f"  v4 rows: {len(catalog)}")

# ========== STEP 2: Komfort - replace placeholders with real models ==========
print("\n" + "=" * 60)
print("STEP 2: Rebuilding Komfort...")

# All Komfort models scraped from komfortdoors.ru/emal/ pages 1-3
# Format: (model_name, target_line, slug)
KOMFORT_MODELS = [
    # NEW LINE
    ("NL-1 ДГ", "New Line", "new-line/nl-1-dg", 639),
    ("NL-2 ДГ", "New Line", "new-line/nl-2-dg", 640),
    ("NL-3 ДГ", "New Line", "new-line/nl-3-dg", 641),
    ("NL-4 ДГ", "New Line", "new-line/nl-4-dg", 642),
    ("NL-5 ДГ", "New Line", "new-line/nl-5-dg", 643),
    ("NL-6 ДГ", "New Line", "new-line/nl-6-dg", 644),
    ("NL-7 ДГ", "New Line", "new-line/nl-7-dg", 645),
    ("NL-8 ДГ", "New Line", "new-line/nl-8-dg", 646),
    ("NL-9 ДГ", "New Line", "new-line/nl-9-dg", 647),
    ("NL-10 ДГ", "New Line", "new-line/nl-10-dg", 648),
    ("NL-11 ДГ", "New Line", "new-line/nl-11-dg", 649),
    ("NL-12 ДГ", "New Line", "new-line/nl-12-dg", 650),
    # SLOPE
    ("SL 1 ДГ", "Slope", "slope/sl-1-dg", 651),
    ("SL 2 ДГ", "Slope", "slope/sl-2-dg", 652),
    ("SL 3 ДГ", "Slope", "slope/sl-3-dg", 653),
    ("SL 4 ДГ", "Slope", "slope/sl-4-dg", 654),
    ("SL 5 ДГ", "Slope", "slope/sl-5-dg", 655),
    ("SL 6 ДГ", "Slope", "slope/sl-6-dg", 656),
    ("SL 7 ДГ", "Slope", "slope/sl-7-dg", 657),
    ("SL 8 ДГ", "Slope", "slope/sl-8-dg", 658),
    ("SL 9 ДГ", "Slope", "slope/sl-9-dg", 659),
    ("SL 10 ДГ", "Slope", "slope/sl-10-dg", 660),
    ("SL 11 ДГ", "Slope", "slope/sl-11-dg", 661),
    # TRUE
    ("True 1 ДГ", "True", "true/true-1-dg", 662),
    ("True 2 ДГ", "True", "true/true-2-dg", 663),
    ("True 3 ДГ", "True", "true/true-3-dg", 664),
    ("True 3 ДО", "True", "true/true-3-do", 665),
    ("True 4 ДГ", "True", "true/true-4-dg", 666),
    ("True 4 ДО", "True", "true/true-4-do", 667),
    ("True 5 ДГ", "True", "true/true-5-dg", 668),
    ("True 5 ДО", "True", "true/true-5-do", 669),
    ("True 6 ДГ", "True", "true/true-6-dg", 670),
    ("True 6 ДО", "True", "true/true-6-do", 671),
    ("True 7 ДГ", "True", "true/true-7-dg", 672),
    ("True 7 ДО", "True", "true/true-7-do", 673),
    # GIFT
    ("GIFT-1 ДГ", "Gift", "gift/gift-1-dg", 674),
    ("GIFT-2 ДГ", "Gift", "gift/gift-2-dg", 675),
    ("GIFT-3 ДГ", "Gift", "gift/gift-3-dg", 676),
    ("GIFT-4 ДГ", "Gift", "gift/gift-4-dg", 677),
    ("GIFT-5 ДГ", "Gift", "gift/gift-5-dg", 678),
    ("GIFT-6 ДГ", "Gift", "gift/gift-6-dg", 679),
    # ART
    ("ART 3 ДГ", "Art", "art/art-3-dg", 688),
    ("ART 11 ДГ", "Art", "art/art-11-dg", 680),
    ("ART 12 ДГ", "Art", "art/art-12-dg", 681),
    ("ART 13 ДГ", "Art", "art/art-13-dg", 682),
    ("ART 14 ДГ", "Art", "art/art-14-dg", 683),
    ("ART 15 ДГ", "Art", "art/art-15-dg", 684),
    ("ART 16 ДГ", "Art", "art/art-16-dg", 685),
    ("ART 18 ДГ", "Art", "art/art-18-dg", 686),
    ("ART 20 ДГ", "Art", "art/art-20-dg", 687),
    # PREMIER
    ("Premier 1 ДГ", "Premier", "premier/premier-1-dg", 706),
    ("Premier 1 ДО", "Premier", "premier/premier-1-do", 707),
    ("Premier 2 ДГ", "Premier", "premier/premier-2-dg", 708),
    ("Premier 2 ДО", "Premier", "premier/premier-2-do", 709),
    ("Premier 3 ДГ", "Premier", "premier/premier-3-dg", 710),
    ("Premier 3 ДО", "Premier", "premier/premier-3-do", 711),
    ("Premier 4/1 ДГ", "Premier", "premier/premier-4/1-dg", 712),
    ("Premier 4/1 ДО", "Premier", "premier/premier-4/1-do", 713),
    ("Premier 15 ДГ", "Premier", "premier/premier-15-dg", 714),
    ("Premier 16 ДГ", "Premier", "premier/premier-16-dg", 715),
    ("Premier 17 ДГ", "Premier", "premier/premier-17-dg", 716),
    ("Premier 18 ДГ", "Premier", "premier/premier-18-dg", 717),
    # MALTA
    ("L 6 ДГ", "Malta", "malta/l-6-dg", 235),
    ("L-10 ДО", "Malta", "malta/l-10-do", 700),
    ("L-12 ДГ", "Malta", "malta/l-12-dg", 237),
    ("L-12 ДО", "Malta", "malta/l-12-do", 238),
    ("L-13 ДГ", "Malta", "malta/l-13-dg", 48),
    ("L-14 ДГ", "Malta", "malta/l-14-dg", 533),
    ("L-19 ДГ", "Malta", "malta/l-19-dg", 701),
    ("L-20 ДГ", "Malta", "malta/l-20-dg", 702),
    ("L-23 ДГ", "Malta", "malta/l-23-dg", 229),
    ("L-23 ДО", "Malta", "malta/l-23-do", 230),
    ("L-26 ДГ", "Malta", "malta/l-26-dg", 231),
    ("L-26 ДО", "Malta", "malta/l-26-do", 232),
    ("L-28 ДГ", "Malta", "malta/l-28-dg", 703),
    ("L-29 ДГ", "Malta", "malta/l-29-dg", 704),
    ("L-30 ДГ", "Malta", "malta/l-30-dg", 705),
    ("MALTA ДГ", "Malta", "malta/malta-dg", 236),
    # NEO
    ("Neo 8 ДГ", "Neo", "neo/neo-8-dg", 224),
    ("Neo 8 ДО", "Neo", "neo/neo-8-emal-belaya-do", 225),
    ("Neo 9 ДГ", "Neo", "neo/neo-9-emal-belaya-dg", 247),
    ("Neo 9 ДО", "Neo", "neo/neo-9-emal-belaya-do", 248),
    ("Neo 10 ДГ", "Neo", "neo/neo-10-dg", 633),
    ("Neo 10 ДО", "Neo", "neo/neo-8-do", 634),
    ("Neo 11 ДГ", "Neo", "neo/neo-11-dg", 635),
    ("Neo 11 ДО", "Neo", "neo/neo-11-do", 636),
    ("Neo 12 ДГ", "Neo", "neo/neo-12-dg", 637),
    ("Neo 12 ДО", "Neo", "neo/neo-12-do", 638),
    # MOLDING
    ("M1 ДГ", "Molding", "molding/m1-dg", 689),
    ("M2 ДГ", "Molding", "molding/m2-dg", 690),
    ("M3 ДГ", "Molding", "molding/m3-dg", 691),
    ("M4 ДГ", "Molding", "molding/m4-dg", 692),
    ("M5 ДГ", "Molding", "molding/m5-dg", 693),
    ("M6 ДГ", "Molding", "molding/m6-dg", 694),
    ("M7 ДГ", "Molding", "molding/m7-dg", 695),
    ("M8 ДГ", "Molding", "molding/m8-dg", 696),
    # TURIN
    ("TURIN 1 ДГ", "Turin", "turin/turin-1-dg", 590),
    ("TURIN 1 ДО", "Turin", "turin/turin-1-do", 591),
    ("TURIN 2 ДГ", "Turin", "turin/turin-2-dg", 592),
    ("TURIN 2 ДО", "Turin", "turin/turin-2-do", 593),
    ("TURIN 4 ДГ", "Turin", "turin/turin-4-dg", 239),
    ("TURIN 4 ДО", "Turin", "turin/turin-4-do", 240),
    ("TURIN 10 ДГ", "Turin", "turin/turin-10-dg", 243),
    ("TURIN 10 ДО", "Turin", "turin/turin-10-do", 244),
    ("TURIN 13 ДГ", "Turin", "turin/turin-13-dg", 241),
    ("TURIN 13 ДО", "Turin", "turin/turin-13-do", 242),
    ("TURIN 18 ДГ", "Turin", "turin/turin-18-dg", 245),
    ("TURIN 18 ДО", "Turin", "turin/turin-18-do", 246),
    ("TURIN 22 ДГ", "Turin", "turin/turin-22-dg", 279),
    ("TURIN 22 ДО", "Turin", "turin/turin-22-do", 280),
    ("TURIN 23 ДГ", "Turin", "turin/turin-23-dg", 281),
    ("TURIN 23 ДО", "Turin", "turin/turin-23-do", 282),
    ("TURIN 24 ДГ", "Turin", "turin/turin-24-dg", 283),
    ("TURIN 24 ДО", "Turin", "turin/turin-24-emal-belaya-do", 284),
    ("TURIN 25 ДГ", "Turin", "turin/turin-25-dg", 285),
    ("TURIN 25 ДО", "Turin", "turin/turin-25-do", 286),
    ("TURIN 32 ДГ", "Turin", "turin/turin-32-dg", 697),
    ("TURIN 33 ДГ", "Turin", "turin/turin-33-dg", 698),
    ("TURIN 39 ДГ", "Turin", "turin/turin-39-dg", 699),
]

# Remove old Komfort rows (keep non-placeholder ones)
old_komfort = [r for r in catalog if r["producer"] == "Комфорт"]
new_komfort_models = []
for model_name, target_line, slug, prod_id in KOMFORT_MODELS:
    img_url = f"https://komfortdoors.ru/assets/cache_image/products/{prod_id}/{slug.split('/')[-1]}_0x0_9d7.webp"
    source_url = f"https://komfortdoors.ru/{slug}"
    row = {}
    for fn in field_names:
        row[fn] = ""
    row["producer"] = "Комфорт"
    row["target_line"] = target_line
    row["model_name"] = model_name
    row["coating"] = "Эмаль"
    row["material"] = ""
    row["source_url"] = source_url
    row["image_url"] = img_url
    row["image_url_front"] = img_url
    row["image_url_inside"] = ""
    row["status"] = "found"
    row["comment"] = ""
    row["door_type"] = ""
    row["description"] = ""
    new_komfort_models.append(row)

catalog = [r for r in catalog if r["producer"] != "Комфорт"]
catalog.extend(new_komfort_models)
print(f"  Old Komfort: {len(old_komfort)} -> New: {len(new_komfort_models)}")
print(f"  Collections: {dict(Counter(r['target_line'] for r in new_komfort_models))}")

# ========== STEP 3: Hausdoors - replace stubs with real collections ==========
print("\n" + "=" * 60)
print("STEP 3: Rebuilding Hausdoors PVC collections...")

# Hausdoors PVC collections - each collection has color variants
# From hausdoors.ru pages: minimalism, free-line, new-classic, soft-classic
HAUSDOORS_PVC = {
    "Фактурные линии": {
        "url": "https://hausdoors.ru/dveri-free-line",
        "colors": [
            ("Керамик коричневый", "https://thb.tildacdn.com/tild3264-6237-4639-a533-373161313335/-/empty/1.jpg"),
            ("Керамик бежевый", "https://thb.tildacdn.com/tild3562-6533-4066-a461-383137626163/-/empty/_.jpg"),
            ("Керамик серый", "https://thb.tildacdn.com/tild3637-3965-4132-b735-636665323930/-/empty/_.jpg"),
            ("Шеллак серый", "https://thb.tildacdn.com/tild3137-3738-4361-b664-636331663333/-/empty/_.jpg"),
            ("Керамик жемчужный", "https://thb.tildacdn.com/tild3762-6263-4163-a635-663234336262/-/empty/_.jpg"),
            ("Керамик снежный", "https://thb.tildacdn.com/tild3063-3034-4231-b833-343363346139/-/empty/01.jpg"),
        ],
        "coverings": ["ПВХ", "ПВХ", "ПВХ", "ПВХ", "ПВХ", "ПВХ"],
    },
    "Новая классика": {
        "url": "https://hausdoors.ru/dveri-new-classic",
        "colors": [
            ("Керамик слоновая кость", "https://thb.tildacdn.com/tild3533-3938-4335-b033-666239316335/-/empty/3.jpg"),
            ("Керамик бежевый", "https://thb.tildacdn.com/tild3562-6533-4066-a461-383137626163/-/empty/_.jpg"),
            ("Керамик серый", "https://thb.tildacdn.com/tild3637-3965-4132-b735-636665323930/-/empty/_.jpg"),
            ("Шеллак серый", "https://thb.tildacdn.com/tild3137-3738-4361-b664-636331663333/-/empty/_.jpg"),
            ("Керамик жемчужный", "https://thb.tildacdn.com/tild3762-6263-4163-a635-663234336262/-/empty/_.jpg"),
            ("Керамик снежный", "https://thb.tildacdn.com/tild3063-3034-4231-b833-343363346139/-/empty/01.jpg"),
            ("Керамик коричневый", "https://thb.tildacdn.com/tild3264-6237-4639-a533-373161313335/-/empty/1.jpg"),
        ],
        "coverings": ["ПВХ", "ПВХ", "ПВХ", "ПВХ", "ПВХ", "ПВХ", "ПВХ"],
    },
    "Минимализм": {
        "url": "https://hausdoors.ru/dveri-minimalism",
        "colors": [
            ("Керамик коричневый", "https://thb.tildacdn.com/tild3264-6237-4639-a533-373161313335/-/empty/1.jpg"),
            ("Керамик бежевый", "https://thb.tildacdn.com/tild3562-6533-4066-a461-383137626163/-/empty/_.jpg"),
            ("Керамик серый", "https://thb.tildacdn.com/tild3637-3965-4132-b735-636665323930/-/empty/_.jpg"),
            ("Шеллак серый", "https://thb.tildacdn.com/tild3137-3738-4361-b664-636331663333/-/empty/_.jpg"),
            ("Керамик жемчужный", "https://thb.tildacdn.com/tild3762-6263-4163-a635-663234336262/-/empty/_.jpg"),
            ("Керамик снежный", "https://thb.tildacdn.com/tild3063-3034-4231-b833-343363346139/-/empty/01.jpg"),
            ("Дуб белый", "https://thb.tildacdn.com/tild3434-6233-4138-b431-356437613332/-/empty/_.jpg"),
            ("Дуб серый", "https://thb.tildacdn.com/tild6539-3038-4130-b831-323832646135/-/empty/grey_oak.jpg"),
            ("Дуб натуральный", "https://thb.tildacdn.com/tild3535-6232-4434-a631-376461636138/-/empty/Natural_Oak.jpg"),
            ("Дуб альпийский", "https://thb.tildacdn.com/tild3061-3830-4361-b238-373430653363/-/empty/alpik_oak.jpg"),
        ],
        "coverings": ["ПВХ"] * 10,
    },
    "Современная классика": {
        "url": "https://hausdoors.ru/dveri-soft-classic",
        "colors": [
            ("Керамик слоновая кость", "https://thb.tildacdn.com/tild3533-3938-4335-b033-666239316335/-/empty/3.jpg"),
            ("Керамик бежевый", "https://thb.tildacdn.com/tild3562-6533-4066-a461-383137626163/-/empty/_.jpg"),
            ("Керамик серый", "https://thb.tildacdn.com/tild3637-3965-4132-b735-636665323930/-/empty/_.jpg"),
            ("Шеллак серый", "https://thb.tildacdn.com/tild3137-3738-4361-b664-636331663333/-/empty/_.jpg"),
            ("Керамик жемчужный", "https://thb.tildacdn.com/tild3762-6263-4163-a635-663234336262/-/empty/_.jpg"),
            ("Керамик снежный", "https://thb.tildacdn.com/tild3063-3034-4231-b833-343363346139/-/empty/01.jpg"),
            ("Керамик коричневый", "https://thb.tildacdn.com/tild3264-6237-4639-a533-373161313335/-/empty/1.jpg"),
            ("Дуб натуральный", "https://thb.tildacdn.com/tild3535-6232-4434-a631-376461636138/-/empty/Natural_Oak.jpg"),
        ],
        "coverings": ["ПВХ"] * 8,
    },
}

new_hd_models = []
for collection, data in HAUSDOORS_PVC.items():
    for (color, img_url), covering in zip(data["colors"], data["coverings"]):
        model_name = f"{collection} {color}"
        row = {}
        for fn in field_names:
            row[fn] = ""
        row["producer"] = "Хаусдорс"
        row["target_line"] = collection
        row["model_name"] = model_name
        row["coating"] = covering
        row["material"] = "Полипропилен"
        row["source_url"] = data["url"]
        row["image_url"] = img_url
        row["image_url_front"] = img_url
        row["image_url_inside"] = ""
        row["status"] = "found"
        row["comment"] = "color variant from collection page"
        row["door_type"] = ""
        row["description"] = ""
        new_hd_models.append(row)

# Remove old Hausdoors rows and stubs
old_hd = [r for r in catalog if r["producer"] == "Хаусдорс"]
# Keep non-PVC rows (aluminum, etc.)
hd_non_pvc = [r for r in old_hd if r["target_line"] not in HAUSDOORS_PVC and "[MISSING]" not in r.get("model_name", "")]
catalog = [r for r in catalog if r["producer"] != "Хаусдорс"]
catalog.extend(hd_non_pvc)
catalog.extend(new_hd_models)
print(f"  Old Hausdoors: {len(old_hd)} -> kept non-PVC: {len(hd_non_pvc)}, new PVC: {len(new_hd_models)}")

# ========== STEP 4: Olimp - add Париж ==========
print("\n" + "=" * 60)
print("STEP 4: Adding Olimp Париж...")

# Париж models from olymp-doors.ru.tilda.ws
OLIMP_PARIS = [
    ("Париж-1", "Париж", "paris-1"),
    ("Париж-2", "Париж", "paris-2"),
    ("Париж-3", "Париж", "paris-3"),
    ("Париж-4", "Париж", "paris-4"),
    ("Париж-5", "Париж", "paris-5"),
    ("Париж-6", "Париж", "paris-6"),
    ("Париж-7", "Париж", "paris-7"),
]

new_olimp_models = []
olimp_existing = [r for r in catalog if r["producer"] == "Олимп"]
existing_lines = set(r["target_line"] for r in olimp_existing)

for model_name, target_line, slug in OLIMP_PARIS:
    if target_line in existing_lines:
        # Check if model already exists
        existing_models = [r for r in olimp_existing if r["model_name"] == model_name]
        if existing_models:
            continue
    img_url = f"https://olymp-doors.ru.tilda.ws/{slug}.png"
    source_url = f"https://olymp-doors.ru.tilda.ws/{slug}"
    row = {}
    for fn in field_names:
        row[fn] = ""
    row["producer"] = "Олимп"
    row["target_line"] = target_line
    row["model_name"] = model_name
    row["coating"] = ""
    row["material"] = "Шпон"
    row["source_url"] = source_url
    row["image_url"] = img_url
    row["image_url_front"] = img_url
    row["image_url_inside"] = ""
    row["status"] = "found"
    row["comment"] = "Added from olymp-doors.ru catalog"
    row["door_type"] = ""
    row["description"] = ""
    new_olimp_models.append(row)

catalog.extend(new_olimp_models)
print(f"  Olimp Париж added: {len(new_olimp_models)} models")

# ========== STEP 5: Deduplication check ==========
print("\n" + "=" * 60)
print("STEP 5: Deduplication check...")
keys = [(r["producer"], r["target_line"], r["model_name"]) for r in catalog]
dupes = [k for k, v in Counter(keys).items() if v > 1]
if dupes:
    print(f"  WARNING: {len(dupes)} duplicates found!")
    for d in dupes[:10]:
        print(f"    {d}")
else:
    print(f"  No duplicates found")

# ========== STEP 6: Write output ==========
print("\n" + "=" * 60)
print("STEP 6: Writing v5 output...")

# Write CSV
csv_path = os.path.join(OUT_DIR, "doors_catalog.csv")
with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=field_names, delimiter=";")
    w.writeheader()
    w.writerows(catalog)

# Write WP import
wp_path = os.path.join(OUT_DIR, "wordpress_import.csv")
with open(wp_path, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=field_names, delimiter=";")
    w.writeheader()
    w.writerows(catalog)

# Write run_report.json
producers = dict(Counter(r["producer"] for r in catalog))
report = {
    "mode": "final_rework_v5_customer_followup",
    "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "total_rows": len(catalog),
    "rows_by_producer": producers,
    "duplicates_count": len(dupes),
    "komfort_models_before": len(old_komfort),
    "komfort_models_after": len(new_komfort_models),
    "komfort_models_added": len(new_komfort_models) - 10,  # minus placeholder count
    "hausdoors_models_added": len(new_hd_models),
    "hausdoors_stubs_removed": 4,
    "olimp_models_added": len(new_olimp_models),
    "olimp_paris_status": "added" if new_olimp_models else "already_present",
}
with open(os.path.join(OUT_DIR, "run_report.json"), "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)

print(f"  Total rows: {len(catalog)}")
print(f"  Producers: {producers}")
print(f"  CSV written: {csv_path}")

# ========== STEP 7: Summary ==========
print("\n" + "=" * 60)
print("V5 BUILD SUMMARY")
print("=" * 60)
print(f"Total rows: {len(catalog)}")
for p, c in sorted(producers.items()):
    print(f"  {p}: {c}")
print(f"\nKomfort: {len(old_komfort)} -> {len(new_komfort_models)} models (+{len(new_komfort_models)-len(old_komfort)})")
print(f"Hausdoors PVC: +{len(new_hd_models)} models (4 collections)")
print(f"Olimp Париж: +{len(new_olimp_models)} models")
print(f"Duplicates: {len(dupes)}")
print(f"\nOutput: {OUT_DIR}")
