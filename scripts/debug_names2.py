import csv, json, unicodedata

ASD = '\u0410\u0421\u0414'
with open('output/final_rework_ready_v5/doors_catalog.csv', encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
asd_rows = [r for r in rows if r['producer'] == ASD]

cat_names = set()
for r in asd_rows:
    cat_names.add(unicodedata.normalize('NFC', r['model_name'].strip()))
print(f"Catalog unique names: {len(cat_names)}")

with open('output/delivery_v5_2/wayback_inside_all_results.json') as f:
    raw = f.read()
wb = json.loads(raw)
wb_names = set()
for k in wb:
    wb_names.add(unicodedata.normalize('NFC', k.strip()))
print(f"Wayback unique names: {len(wb_names)}")

# Check overlap
overlap = cat_names & wb_names
print(f"Exact overlap: {len(overlap)}")

# Check case-insensitive
cat_lower = {n.lower() for n in cat_names}
wb_lower = {n.lower() for n in wb_names}
overlap_lower = cat_lower & wb_lower
print(f"Case-insensitive overlap: {len(overlap_lower)}")

# Sample comparison
cat_sample = list(cat_names)[0]
wb_sample = list(wb_names)[0]
print(f"\nCat sample [{repr(cat_sample[:60])}]")
print(f"WB  sample [{repr(wb_sample[:60])}]")
print(f"Cat hex: {cat_sample[:20].encode('utf-8').hex(':')}")
print(f"WB  hex: {wb_sample[:20].encode('utf-8').hex(':')}")

# Check if wb keys have extra quotes
print(f"\nWB first key repr: {repr(list(wb.keys())[0][:60])}")
