import csv, zipfile, io, os

ZIP = os.path.join("output", "doors_parser_final_REWORK_READY_v4_2026-06-23_17-08-14.zip")
with zipfile.ZipFile(ZIP) as zf:
    csv_name = [n for n in zf.namelist() if n.endswith("doors_catalog.csv")][0]
    with zf.open(csv_name) as f:
        catalog = list(csv.DictReader(io.StringIO(f.read().decode("utf-8-sig")), delimiter=";"))

hd = [r for r in catalog if r["producer"] == "\u0425\u0430\u0443\u0441\u0434\u043e\u0440\u0441"]
print(f"Current Hausdoors rows: {len(hd)}")
for r in hd:
    tl = r["target_line"][:30]
    mn = r["model_name"][:40]
    st = r.get("status", "")
    su = r.get("source_url", "")[:80]
    print(f"  {tl} | {mn} | status={st} | {su}")
print()
# Also print all field names from first row
if hd:
    print("Fields:", list(hd[0].keys()))
    print()
    # Show full details for each
    for i, r in enumerate(hd):
        print(f"--- Row {i+1} ---")
        for k, v in r.items():
            print(f"  {k}: {v[:100] if isinstance(v, str) else v}")
        print()
