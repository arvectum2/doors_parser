"""
Stage A Runner: Execute all Stage A steps sequentially.
Usage: python scripts/run_stage_a.py [--skip-download] [--skip-lord]

Stage A builds the review package WITHOUT modifying the catalog.
Place VLM output at output/manual_photo_review_v5_7/reports/asd_vlm_output.json
and output/manual_photo_review_v5_7/reports/lord_vlm_output.json
then re-run to regenerate HTML with classifications.
"""
import os, sys, subprocess, json
from datetime import datetime

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
SCRIPTS = os.path.join(BASE, "scripts")
STAGE_A = os.path.join(SCRIPTS, "stage_a")
REPORTS_DIR = os.path.join(BASE, "output", "manual_photo_review_v5_7", "reports")
VENV_PYTHON = os.path.join(BASE, ".venv", "Scripts", "python.exe")

skip_download = '--skip-download' in sys.argv
skip_lord = '--skip-lord' in sys.argv

def run_step(name, script, args=None):
    print(f"\n{'='*60}")
    print(f"STEP: {name}")
    print(f"{'='*60}")
    cmd = [VENV_PYTHON, script]
    if args: cmd.extend(args)
    result = subprocess.run(cmd, cwd=BASE, capture_output=False)
    if result.returncode != 0:
        print(f"ERROR: Step '{name}' failed with code {result.returncode}")
        sys.exit(1)
    print(f"[OK] Step '{name}' complete")

# Ensure reports dir exists
os.makedirs(REPORTS_DIR, exist_ok=True)

# Check for VLM outputs
asd_vlm = os.path.join(REPORTS_DIR, "asd_vlm_output.json")
lord_vlm = os.path.join(REPORTS_DIR, "lord_vlm_output.json")

if os.path.exists(asd_vlm):
    with open(asd_vlm) as f:
        vlm_count = len(json.load(f))
    print(f"  Found ASD VLM output: {vlm_count} classifications")
else:
    print(f"  ASD VLM output not found at {asd_vlm}")
    print(f"  (Review HTML will show candidates without VLM data)")

if os.path.exists(lord_vlm):
    with open(lord_vlm) as f:
        vlm_count = len(json.load(f))
    print(f"  Found Lord VLM output: {vlm_count} classifications")
else:
    print(f"  Lord VLM output not found at {lord_vlm}")

# Step 1: Collect candidates and download
if not skip_download:
    run_step("Collect ASD candidates", os.path.join(STAGE_A, "a01_collect_candidates.py"))
else:
    print("\nSkipping download (--skip-download)")

# Step 2: Build review HTML, contact sheet, risky rows
run_step("Build review HTML + contact sheet + risky rows",
         os.path.join(STAGE_A, "a02_build_review_html.py"))

# Step 3: Lord audit
if not skip_lord:
    run_step("Lord visual audit", os.path.join(STAGE_A, "a03_lord_audit.py"))
else:
    print("\nSkipping Lord audit (--skip-lord)")

# ========================================================================
# Final Stage A summary
# ========================================================================
print(f"\n{'='*60}")
print(f"STAGE A COMPLETE")
print(f"{'='*60}")
print(f"  Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
print()

# Count items
reports_dir = REPORTS_DIR
for fname in sorted(os.listdir(reports_dir)):
    fpath = os.path.join(reports_dir, fname)
    size = os.path.getsize(fpath)
    print(f"  {fname} ({size:,} bytes)")

print()
print("NEXT STEPS:")
print("  1. Open asd_manual_review.html in browser")
print("  2. Manually review each model's candidates")
print("  3. Select Front/Inside using the buttons")
print("  4. Approve pair or mark Needs Review")
print("  5. Export decisions to asd_manual_decisions.json")
print("  6. For VLM integration:")
print(f"     - Run VLM on candidates from {asd_vlm}")
print(f"     - Save results to {asd_vlm}")
print(f"     - Re-run this script to regenerate HTML")
print("  7. For Lord:")
print(f"     - Open {os.path.join(reports_dir, 'lord_flagged_review.html')}")
print(f"     - Review flagged rows")
print("  8. When decisions JSON is ready, proceed to Stage B")
print()
print("REMINDER: Stage A does NOT modify the catalog.")
print("          No delivery ZIP was built.")
print("          All outputs are in the review package.")
