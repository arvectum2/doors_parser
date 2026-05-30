"""Batch re-run script for 6 problem producers.
Run when sites are accessible. Saves each result separately, then merges with good data.

Usage:
    python scripts/rerun_problem_sites.py
"""
import subprocess
import sys
import time
from pathlib import Path

PROJECT = Path(__file__).resolve().parent.parent
VENV_PYTHON = PROJECT / ".venv" / "Scripts" / "python.exe"
PARSER = PROJECT / "src" / "doors_parser.py"
MERGE = PROJECT / "merge_results.py"
TEMP_DIR = PROJECT / "output" / "problem_sites"
GOOD_DATA = PROJECT / "output" / "good_data_backup.csv"

# Map manufacturer names to --only argument (use exact names from target_lines.csv)
PROBLEM_PRODUCERS = [
    "Скамбио Порте",
    "Олимп",
    "АСД",
    "Комфорт",
    "Двери регионов",
    "Термодур",
]

def run_parser(manufacturer: str) -> bool:
    """Run parser for one manufacturer. Returns True if success."""
    out_dir = TEMP_DIR / manufacturer.replace(" ", "_").lower()
    out_dir.mkdir(parents=True, exist_ok=True)
    
    cmd = [
        str(VENV_PYTHON),
        str(PARSER),
        "--only", manufacturer,
        "--out", str(out_dir),
        "--max-pages-per-site", "200",
        "--debug",
    ]
    
    print(f"\n{'='*60}")
    print(f"  Running: {manufacturer}")
    print(f"  Output: {out_dir}")
    print(f"{'='*60}")
    
    result = subprocess.run(cmd, cwd=str(PROJECT), capture_output=False, text=True, timeout=300)
    success = result.returncode == 0
    
    # Check if output has data
    csv_file = out_dir / "doors_catalog.csv"
    if csv_file.exists():
        import csv
        with csv_file.open("r", encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f, delimiter=";"))
        print(f"  Result: {len(rows)} rows")
        if len(rows) > 0:
            return True
    
    print(f"  WARNING: {manufacturer} produced {0 if not csv_file.exists() else len(rows)} rows")
    return False


def main():
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    
    print("=" * 60)
    print("  Batch Re-run: 6 Problem Producers")
    print(f"  Python: {VENV_PYTHON}")
    print(f"  Parser: {PARSER}")
    print(f"  Good data: {GOOD_DATA}")
    print(f"  Temp dir: {TEMP_DIR}")
    print("=" * 60)
    
    if not VENV_PYTHON.exists():
        print(f"ERROR: Python not found at {VENV_PYTHON}")
        print("Make sure the virtual environment exists.")
        sys.exit(1)
    
    results = {}
    for manufacturer in PROBLEM_PRODUCERS:
        print(f"\nWaiting 2 seconds before processing {manufacturer}...")
        time.sleep(2)
        success = run_parser(manufacturer)
        results[manufacturer] = success
    
    # Summary
    print("\n" + "=" * 60)
    print("  Results Summary")
    print("=" * 60)
    for man, ok in results.items():
        status = "OK" if ok else "FAILED / EMPTY"
        print(f"  {man:25s}: {status}")
    
    # Merge
    print(f"\n{'='*60}")
    print(f"  Merging results with good data...")
    print(f"{'='*60}")
    
    # Collect all new CSV files from temp directories
    new_csvs = list(TEMP_DIR.glob("*/doors_catalog.csv"))
    
    merge_cmd = [
        str(VENV_PYTHON),
        str(MERGE),
        "--good", str(GOOD_DATA),
        "--out", str(PROJECT / "output"),
    ] + [str(c) for c in new_csvs]
    
    print(f"  Merging {len(new_csvs)} new files with {GOOD_DATA.name}")
    subprocess.run(merge_cmd, cwd=str(PROJECT))
    
    print(f"\n  Done. Final output: {PROJECT / 'output' / 'doors_catalog.csv'}")


if __name__ == "__main__":
    main()
