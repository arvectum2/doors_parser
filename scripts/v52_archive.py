"""Create v5.2 delivery archive"""
import os, zipfile, hashlib
from datetime import datetime

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
IN = os.path.join(BASE, "output", "delivery_v5_2")
OUT = os.path.join(BASE, "output", "delivery")

os.makedirs(OUT, exist_ok=True)

# Files for client delivery
CLIENT_FILES = [
    "doors_catalog.csv",
    "doors_catalog.xlsx",
    "wordpress_import.csv",
    "run_report.json",
    "final_delivery_qa_report.xlsx",
    "asd_front_inside_restore_report.xlsx",
    "asd_missing_front_inside_final.xlsx",
]

# Internal/auxiliary files (not client-facing but useful for auditing)
AUX_FILES = [
    "wayback_inside_all_results.json",
]

print("Files to include:")
all_ok = True
for f in CLIENT_FILES:
    p = os.path.join(IN, f)
    exists = os.path.exists(p)
    status = "OK" if exists else "MISSING"
    print(f"  {f}: {status} ({os.path.getsize(p) if exists else 0:,} bytes)")
    if not exists: all_ok = False

if not all_ok:
    print("\nERROR: Some client files missing!")
    exit(1)

# Create archive
ts = datetime.now().strftime("%Y-%m-%d_%H-%M")
archive_name = f"doors_catalog_delivery_v5_2_{ts}.zip"
archive_path = os.path.join(OUT, archive_name)

print(f"\nCreating archive: {archive_name}")
with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for f in CLIENT_FILES:
        p = os.path.join(IN, f)
        zf.write(p, f)
        print(f"  Added: {f} ({os.path.getsize(p):,} bytes)")
    for f in AUX_FILES:
        p = os.path.join(IN, f)
        if os.path.exists(p):
            zf.write(p, f"aux/{f}")
            print(f"  Added: aux/{f} ({os.path.getsize(p):,} bytes)")

# SHA256
sha256 = hashlib.sha256()
with open(archive_path, 'rb') as f:
    while True:
        chunk = f.read(8192)
        if not chunk: break
        sha256.update(chunk)

size = os.path.getsize(archive_path)
print(f"\nArchive: {archive_name}")
print(f"Size: {size:,} bytes")
print(f"SHA256: {sha256.hexdigest()}")
print(f"Path: {archive_path}")
