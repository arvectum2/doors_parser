import os
import zipfile
import hashlib
from datetime import datetime

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
OUT_DIR = os.path.join(BASE, "output", "delivery")
FINAL_DIR = os.path.join(BASE, "output", "final_rework_ready_v5")

timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M')
archive_name = f"doors_catalog_delivery_v5_QA_{timestamp}.zip"
archive_path = os.path.join(OUT_DIR, archive_name)

files_to_include = [
    # Catalog files
    (os.path.join(FINAL_DIR, "doors_catalog.csv"), "doors_catalog.csv"),
    (os.path.join(FINAL_DIR, "doors_catalog.xlsx"), "doors_catalog.xlsx"),
    (os.path.join(FINAL_DIR, "wordpress_import.csv"), "wordpress_import.csv"),
    (os.path.join(FINAL_DIR, "run_report.json"), "run_report.json"),
    # QA reports
    (os.path.join(OUT_DIR, "final_delivery_qa_report.json"), "final_delivery_qa_report.json"),
    (os.path.join(OUT_DIR, "final_delivery_qa_report.xlsx"), "final_delivery_qa_report.xlsx"),
    # Specific checks
    (os.path.join(OUT_DIR, "no_image_rows_v5.csv"), "no_image_rows_v5.csv"),
    # Sub-reports
    (os.path.join(FINAL_DIR, "komfort_rebuild_report.json"), "komfort_rebuild_report.json"),
    (os.path.join(FINAL_DIR, "hausdoors_pvh_rebuild_report.json"), "hausdoors_pvh_rebuild_report.json"),
    (os.path.join(FINAL_DIR, "olimp_missing_collections_report.json"), "olimp_missing_collections_report.json"),
    (os.path.join(FINAL_DIR, "image_url_validation_report.json"), "image_url_validation_report.json"),
    # Image mappings
    (os.path.join(FINAL_DIR, "komfort_images.json"), "komfort_images.json"),
    (os.path.join(FINAL_DIR, "komfort_slug_image_map.json"), "komfort_slug_image_map.json"),
]

print(f"Creating archive: {archive_name}")
print(f"Files to include: {len(files_to_include)}")

with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for src, dst in files_to_include:
        if os.path.exists(src):
            size = os.path.getsize(src)
            zf.write(src, dst)
            print(f"  Added: {dst} ({size:,} bytes)")
        else:
            print(f"  MISSING: {src}")

# Calculate SHA256
with open(archive_path, 'rb') as f:
    sha256 = hashlib.sha256(f.read()).hexdigest()

archive_size = os.path.getsize(archive_path)

# Save SHA256
with open(os.path.join(OUT_DIR, 'delivery_sha256.txt'), 'w') as f:
    f.write(f"Archive: {archive_name}\n")
    f.write(f"SHA256: {sha256}\n")
    f.write(f"Size: {archive_size:,} bytes\n")
    f.write(f"Created: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

print(f"\nArchive: {archive_path}")
print(f"Size: {archive_size:,} bytes")
print(f"SHA256: {sha256}")
