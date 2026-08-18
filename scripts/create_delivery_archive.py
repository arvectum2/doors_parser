"""Create clean delivery archive for client"""
import os, zipfile, hashlib, shutil
from datetime import datetime

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
OUT_DIR = os.path.join(BASE, "output", "final_rework_ready_v4")
DELIVERY = os.path.join(BASE, "output", "delivery")

# Files to include in delivery
CLIENT_FILES = [
    "doors_catalog.csv",
    "doors_catalog.xlsx",
    "wordpress_import.csv",
    "reference_lines_clean.csv",
    "reference_lines_clean.xlsx",
    "reference_lines_clean.json",
    "final_expected_lines_coverage.xlsx",
    "final_stub_mapping_check.xlsx",
    "final_stub_mapping_check.csv",
    "run_report.json",
    "entro_aria_validation_report.xlsx",
    "hausdoors_stub_resolution_report.xlsx",
    "asd_bad_pattern_review.xlsx",
    "asd_bad_pattern_review.html",
]

# Check all files exist
print("Checking files:")
for f in CLIENT_FILES:
    path = os.path.join(OUT_DIR, f)
    exists = os.path.exists(path)
    size = os.path.getsize(path) if exists else 0
    print(f"  {'OK' if exists else 'MISSING'} {f}: {size:,} bytes")

# Create delivery zip
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
zip_name = f"doors_catalog_delivery_{timestamp}.zip"
zip_path = os.path.join(DELIVERY, zip_name)
os.makedirs(DELIVERY, exist_ok=True)

with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
    for f in CLIENT_FILES:
        src = os.path.join(OUT_DIR, f)
        if os.path.exists(src):
            zf.write(src, f"doors_catalog_v4/{f}")
            print(f"  + {f}")

# Calculate SHA256
with open(zip_path, "rb") as f:
    sha = hashlib.sha256(f.read()).hexdigest()
print(f"\nSHA256: {sha}")

# Also copy the old zip for reference
old_zip = os.path.join(BASE, "output", "doors_parser_final_REWORK_READY_v4_2026-06-23_17-08-14.zip")
if os.path.exists(old_zip):
    old_sha = hashlib.sha256(open(old_zip, "rb").read()).hexdigest()
    print(f"Old zip SHA256: {old_sha}")

# Summary
print(f"\nDelivery archive: {zip_path}")
print(f"Files included: {len(CLIENT_FILES)}")
print(f"SHA256: {sha}")
