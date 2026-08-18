import csv, zipfile, io, os, glob, sys
sys.stdout.reconfigure(encoding='utf-8')

zips = glob.glob("output/doors_parser_final_REWORK_READY_v4_*.zip")
ZIP = zips[0]
with zipfile.ZipFile(ZIP) as zf:
    csv_name = [n for n in zf.namelist() if n.endswith("doors_catalog.csv")][0]
    with zf.open(csv_name) as f:
        catalog = list(csv.DictReader(io.StringIO(f.read().decode("utf-8-sig")), delimiter=";"))
olimp = [r for r in catalog if r["producer"] == "Олимп"]
print(f"Current Olimp rows: {len(olimp)}")
collections = {}
for r in olimp:
    coll = r["target_line"]
    collections[coll] = collections.get(coll, 0) + 1
print("Existing Olimp collections:")
for coll, cnt in sorted(collections.items()):
    print(f"  {coll}: {cnt} models")
print("\nSample source URLs:")
urls = set()
for r in olimp:
    url = r.get("source_url", "")
    if url:
        urls.add(url)
for u in sorted(urls):
    print(f"  {u}")
print(f"\nTotal unique URLs: {len(urls)}")
