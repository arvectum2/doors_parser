"""Create clean delivery archive for v5"""
import os, zipfile, hashlib, shutil
from datetime import datetime

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
OUT_DIR = os.path.join(BASE, "output", "final_rework_ready_v5")
DELIVERY = os.path.join(BASE, "output", "delivery")

# Files to include in delivery
CLIENT_FILES = [
    "doors_catalog.csv",
    "run_report.json",
    "komfort_rebuild_report.json",
    "hausdoors_pvh_rebuild_report.json",
    "olimp_missing_collections_report.json",
    "image_url_validation_report.json",
]

# Check all files exist
print("Checking files:")
all_ok = True
for f in CLIENT_FILES:
    path = os.path.join(OUT_DIR, f)
    exists = os.path.exists(path)
    status = "OK" if exists else "MISSING"
    print(f"  {f}: {status}")
    if not exists:
        all_ok = False

if not all_ok:
    print("\nSome files missing. Available files:")
    for f in os.listdir(OUT_DIR):
        print(f"  {f}")

# Create timestamp
ts = datetime.now().strftime("%Y-%m-%d_%H-%M")
archive_name = f"doors_catalog_delivery_v5_{ts}.zip"
archive_path = os.path.join(DELIVERY, archive_name)

os.makedirs(DELIVERY, exist_ok=True)

print(f"\nCreating archive: {archive_name}")
with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for f in CLIENT_FILES:
        path = os.path.join(OUT_DIR, f)
        if os.path.exists(path):
            zf.write(path, f)
            size = os.path.getsize(path)
            print(f"  Added: {f} ({size:,} bytes)")
        else:
            print(f"  Skipped (missing): {f}")

# Calculate SHA256
sha256 = hashlib.sha256()
with open(archive_path, 'rb') as f:
    while True:
        chunk = f.read(8192)
        if not chunk:
            break
        sha256.update(chunk)

archive_size = os.path.getsize(archive_path)
print(f"\nArchive: {archive_name}")
print(f"Size: {archive_size:,} bytes")
print(f"SHA256: {sha256.hexdigest()}")
print(f"Path: {archive_path}")
