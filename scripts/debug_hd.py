"""Debug Hausdoors PVC count"""
import csv, zipfile, io
from collections import Counter

V4_ZIP = "output/doors_parser_final_REWORK_READY_v4_2026-06-23_17-08-14.zip"
with zipfile.ZipFile(V4_ZIP) as zf:
    csv_name = [n for n in zf.namelist() if n.endswith("doors_catalog.csv")][0]
    with zf.open(csv_name) as f:
        catalog = list(csv.DictReader(io.StringIO(f.read().decode("utf-8-sig")), delimiter=";"))
field_names = list(catalog[0].keys())

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
        "coverings": ["ПВХ"] * 6,
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
        "coverings": ["ПВХ"] * 7,
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
    colors_count = len(data["colors"])
    coverings_count = len(data["coverings"])
    print(f"  {collection}: {colors_count} colors, {coverings_count} coverings")
    pairs = list(zip(data["colors"], data["coverings"]))
    print(f"    zip pairs: {len(pairs)}")
    for (color, img_url), covering in pairs:
        model_name = collection + " " + color
        row = dict.fromkeys(field_names, "")
        row["producer"] = "\u0425\u0430\u0443\u0441\u0434\u043e\u0440\u0441"
        row["target_line"] = collection
        row["model_name"] = model_name
        row["coating"] = covering
        row["material"] = "\u041f\u043e\u043b\u0438\u043f\u0440\u043e\u043f\u0438\u043b\u0435\u043d"
        row["source_url"] = data["url"]
        row["image_url"] = img_url
        row["image_url_front"] = img_url
        row["status"] = "found"
        row["comment"] = "color variant from collection page"
        new_hd_models.append(row)
print(f"Total new Hausdoors PVC models: {len(new_hd_models)}")
for m in new_hd_models[:5]:
    print(f"  {m['target_line']}: {m['model_name']}")
