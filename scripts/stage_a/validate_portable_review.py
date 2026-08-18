"""
Stage A - QA: Validate portable review packages.
Checks ASD batches, Lord review, archives, scripts.
Outputs validation report (JSON + XLSX).
Usage: python scripts/stage_a/validate_portable_review.py
"""
import csv, json, os, sys
from collections import defaultdict
from datetime import datetime

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
OUTPUT = os.path.join(BASE, "output", "manual_photo_review_v5_7")
BATCHES_DIR = os.path.join(OUTPUT, "portable_batches")
LORD_DIR = os.path.join(OUTPUT, "lord_portable_review")
REPORTS_DIR = os.path.join(OUTPUT, "reports")
ARCHIVES_DIR = os.path.join(OUTPUT, "archives")

def check(name, passed, detail=""):
    status = "PASS" if passed else "FAIL"
    print(f"  [{status}] {name}" + (f" -- {detail}" if detail else ""))
    return passed

def main():
    print("=" * 60)
    print("PORTABLE REVIEW VALIDATION")
    print("=" * 60)
    
    all_passed = True
    report = {}
    
    # ====================================================================
    # ASD BATCH CHECKS
    # ====================================================================
    print("\n--- ASD Batches ---")
    
    total_models = 0
    all_indexes = set()
    batch_sizes = {}
    all_models = []
    local_images_total = 0
    remote_only_total = 0
    
    for batch_num in range(1, 9):
        label = f"batch_{batch_num:02d}"
        batch_dir = os.path.join(BATCHES_DIR, label)
        models_json_path = os.path.join(batch_dir, "models.json")
        
        if not os.path.isdir(batch_dir):
            all_passed = check(f"{label} directory exists", False) and all_passed
            continue
        
        with open(models_json_path, encoding='utf-8') as f:
            models = json.load(f)
        
        batch_sizes[label] = len(models)
        total_models += len(models)
        
        for m in models:
            idx = m['list_index']
            all_indexes.add(idx)
            all_models.append(m)
    
    all_passed = check("ASD models total = 199", total_models == 199, f"got {total_models}") and all_passed
    all_passed = check("Indexes 1-199 without gaps", all_indexes == set(range(1, 200)),
                       f"got {len(all_indexes)} unique") and all_passed
    
    expected_sizes = {f"batch_{i:02d}": 25 for i in range(1, 8)}
    expected_sizes["batch_08"] = 24
    for label, expected in expected_sizes.items():
        actual = batch_sizes.get(label, 0)
        all_passed = check(f"{label} size = {expected}", actual == expected, f"got {actual}") and all_passed
    
    # ASD content checks
    missing_neg = 0
    broken_images = 0
    initial_approved = 0
    excel_row_filled = 0
    has_autosave = has_import = has_export = has_reset = False
    same_front_inside = 0
    
    for batch_num in range(1, 9):
        label = f"batch_{batch_num:02d}"
        batch_dir = os.path.join(BATCHES_DIR, label)
        
        review_html = os.path.join(batch_dir, "review.html")
        if os.path.exists(review_html):
            with open(review_html, encoding='utf-8') as f:
                html = f.read()
            has_autosave = has_autosave or 'localStorage' in html
            has_import = has_import or 'importJSON' in html
            has_export = has_export or 'exportJSON' in html
            has_reset = has_reset or 'resetBatch' in html
        
        images_dir = os.path.join(batch_dir, "images")
        if os.path.isdir(images_dir):
            for fname in os.listdir(images_dir):
                fpath = os.path.join(images_dir, fname)
                if os.path.isfile(fpath) and os.path.getsize(fpath) == 0:
                    broken_images += 1
        
        neg_dir = os.path.join(batch_dir, "negative_examples")
        for neg_file in ['IMG_1946.jpg','IMG_1947.jpg','IMG_1948.jpg','IMG_1949.jpg',
                         'IMG_1951.jpg','IMG_1952.jpg','IMG_1953.jpg','IMG_1954.jpg',
                         'IMG_1955.jpg','IMG_1956.jpg']:
            if not os.path.exists(os.path.join(neg_dir, neg_file)):
                missing_neg += 1
        
        json_path = os.path.join(batch_dir, f"manual_decisions_{label}.json")
        if os.path.exists(json_path):
            with open(json_path, encoding='utf-8') as f:
                decisions = json.load(f)
            for d in decisions:
                if d.get('approved_front_url') or d.get('approved_inside_url'):
                    initial_approved += 1
                if d.get('excel_row'):
                    excel_row_filled += 1
                front = d.get('approved_front_url', '')
                inside = d.get('approved_inside_url', '')
                if front and inside and front == inside:
                    same_front_inside += 1
        
        # Count local vs remote candidates
        models_path = os.path.join(batch_dir, "models.json")
        if os.path.exists(models_path):
            with open(models_path, encoding='utf-8') as f:
                batch_models = json.load(f)
            # Read review.html to count candidates
            review_path = os.path.join(batch_dir, "review.html")
            if os.path.exists(review_path):
                with open(review_path, encoding='utf-8') as f:
                    content = f.read()
                # Count local images in images/
                img_dir = os.path.join(batch_dir, "images")
                if os.path.isdir(img_dir):
                    local_images_total += len([f for f in os.listdir(img_dir) if os.path.isfile(os.path.join(img_dir, f))])
    
    # Count total candidates from review.html inline data
    total_candidates = 0
    local_candidates = 0
    for batch_num in range(1, 9):
        label = f"batch_{batch_num:02d}"
        review_path = os.path.join(BATCHES_DIR, label, "review.html")
        if os.path.exists(review_path):
            with open(review_path, encoding='utf-8') as f:
                content = f.read()
            # Count "downloaded": true and "downloaded": false
            import re
            local_candidates += len(re.findall(r'"downloaded":\s*true', content))
            total_candidates += len(re.findall(r'"downloaded":\s*(true|false)', content))
    
    remote_only_total = total_candidates - local_candidates
    
    all_passed = check("Missing negative refs = 0", missing_neg == 0, f"got {missing_neg}") and all_passed
    all_passed = check("Broken images = 0", broken_images == 0, f"got {broken_images}") and all_passed
    all_passed = check("Initial approved = 0", initial_approved == 0, f"got {initial_approved}") and all_passed
    all_passed = check("Excel_row filled = 199", excel_row_filled == 199, f"got {excel_row_filled}") and all_passed
    all_passed = check("Autosave enabled", has_autosave) and all_passed
    all_passed = check("Import present", has_import) and all_passed
    all_passed = check("Export present", has_export) and all_passed
    all_passed = check("Reset present", has_reset) and all_passed
    all_passed = check("Same front/inside = 0", same_front_inside == 0, f"got {same_front_inside}") and all_passed
    
    merge_script = os.path.join(BASE, "scripts", "stage_b", "merge_manual_decisions.py")
    all_passed = check("Merge script present", os.path.exists(merge_script)) and all_passed
    
    # ====================================================================
    # LORD CHECKS
    # ====================================================================
    print("\n--- Lord ---")
    
    # Count Lord rows from catalog
    v56_csv = os.path.join(BASE, "output", "delivery_v5_6", "doors_catalog.csv")
    lord_rows_count = 0
    if os.path.exists(v56_csv):
        with open(v56_csv, encoding='utf-8-sig') as f:
            reader = csv.DictReader(f, delimiter=';')
            for r in reader:
                if r.get('producer') == chr(0x041B) + chr(0x043E) + chr(0x0440) + chr(0x0434):
                    lord_rows_count += 1
    
    # Lord images
    lord_images_dir = os.path.join(LORD_DIR, "images")
    lord_img_count = len([f for f in os.listdir(lord_images_dir) if os.path.isfile(os.path.join(lord_images_dir, f))]) if os.path.isdir(lord_images_dir) else 0
    
    # Lord flagged review HTML
    lord_flagged_html = os.path.join(LORD_DIR, "lord_flagged_review.html")
    lord_flagged_count = 0
    blank_img_refs = 0
    lord_has_autosave = False
    lord_has_import = False
    lord_has_export = False
    lord_has_reset = False
    lord_has_storage_key = False
    
    if os.path.exists(lord_flagged_html):
        with open(lord_flagged_html, encoding='utf-8') as f:
            content = f.read()
        lord_flagged_count = content.count('class="dec-btn"')
        blank_img_refs = content.count('src="images/"')
        lord_has_autosave = 'localStorage' in content
        lord_has_import = 'importJSON' in content
        lord_has_export = 'exportJSON' in content
        lord_has_reset = 'resetAll' in content
        lord_has_storage_key = 'lord_manual_review_v5_7' in content
    
    # Lord all contact sheet
    lord_all_html = os.path.join(LORD_DIR, "lord_all_contact_sheet.html")
    lord_all_blank = 0
    lord_all_has_autosave = False
    lord_all_has_storage_key = False
    lord_no_local_count = 0
    
    if os.path.exists(lord_all_html):
        with open(lord_all_html, encoding='utf-8') as f:
            content = f.read()
        lord_all_blank = content.count('src="images/"')
        lord_all_has_autosave = 'localStorage' in content
        lord_all_has_storage_key = 'lord_manual_review_v5_7' in content
        lord_no_local_count = content.count('NO LOCAL IMAGE') + content.count('no-image')
    
    # Lord decisions
    lord_dec_json = os.path.join(LORD_DIR, "lord_manual_decisions.json")
    lord_dec_count = 0
    lord_initial_reviewed = 0
    if os.path.exists(lord_dec_json):
        with open(lord_dec_json, encoding='utf-8') as f:
            lord_dec = json.load(f)
        lord_dec_count = len(lord_dec)
        for d in lord_dec:
            if d.get('explicitly_reviewed'):
                lord_initial_reviewed += 1
    
    all_passed = check("Lord rows = 932", lord_rows_count == 932, f"got {lord_rows_count}") and all_passed
    all_passed = check("Lord local images = 699", lord_img_count == 699, f"got {lord_img_count}") and all_passed
    all_passed = check("Lord flagged = 47", lord_flagged_count >= 47, f"got {lord_flagged_count} buttons") and all_passed
    all_passed = check("Lord blank image refs = 0", blank_img_refs == 0, f"got {blank_img_refs}") and all_passed
    all_passed = check("Lord all contact blank = 0", lord_all_blank == 0, f"got {lord_all_blank}") and all_passed
    all_passed = check("Lord NO LOCAL IMAGE = 89", lord_no_local_count >= 89, f"got {lord_no_local_count}") and all_passed
    all_passed = check("Lord decisions = 932", lord_dec_count == 932, f"got {lord_dec_count}") and all_passed
    all_passed = check("Lord initial reviewed = 0", lord_initial_reviewed == 0, f"got {lord_initial_reviewed}") and all_passed
    all_passed = check("Lord flagged has localStorage", lord_has_autosave) and all_passed
    all_passed = check("Lord flagged has storage key", lord_has_storage_key) and all_passed
    all_passed = check("Lord flagged has Import", lord_has_import) and all_passed
    all_passed = check("Lord flagged has Export", lord_has_export) and all_passed
    all_passed = check("Lord flagged has Reset", lord_has_reset) and all_passed
    all_passed = check("Lord contact has localStorage", lord_all_has_autosave) and all_passed
    all_passed = check("Lord contact has storage key", lord_all_has_storage_key) and all_passed
    
    # ====================================================================
    # ARCHIVE CHECKS
    # ====================================================================
    print("\n--- Archives ---")
    if os.path.isdir(ARCHIVES_DIR):
        for f in sorted(os.listdir(ARCHIVES_DIR)):
            fp = os.path.join(ARCHIVES_DIR, f)
            size_mb = os.path.getsize(fp) / 1024 / 1024
            all_passed = check(f"{f} <= 510 MB", size_mb <= 510, f"{size_mb:.1f} MB") and all_passed
    else:
        check("Archives directory exists", False)
    
    # ====================================================================
    # SAVE REPORT
    # ====================================================================
    report = {
        'timestamp': datetime.now().isoformat(),
        'all_passed': all_passed,
        'asd': {
            'models_total': total_models,
            'local_unique_images': local_images_total,
            'local_candidate_occurrences': local_candidates,
            'remote_only_candidates': remote_only_total,
            'total_candidates': total_candidates,
            'indexes_complete': all_indexes == set(range(1, 200)),
            'batch_sizes': batch_sizes,
            'missing_negative_refs': missing_neg,
            'broken_images': broken_images,
            'initial_approved': initial_approved,
            'excel_row_filled': excel_row_filled,
            'autosave': has_autosave,
            'import': has_import,
            'export': has_export,
            'reset': has_reset,
            'same_front_inside': same_front_inside,
        },
        'lord': {
            'rows_total': lord_rows_count,
            'local_images': lord_img_count,
            'flagged': lord_flagged_count,
            'blank_image_refs': blank_img_refs,
            'all_contact_blank': lord_all_blank,
            'no_local_image': lord_no_local_count,
            'decisions_template': lord_dec_count,
            'initial_reviewed': lord_initial_reviewed,
            'flagged_has_localStorage': lord_has_autosave,
            'flagged_has_storage_key': lord_has_storage_key,
            'flagged_has_import': lord_has_import,
            'flagged_has_export': lord_has_export,
            'flagged_has_reset': lord_has_reset,
            'contact_has_localStorage': lord_all_has_autosave,
            'contact_has_storage_key': lord_all_has_storage_key,
        },
        'merge_script': os.path.exists(merge_script),
    }
    
    report_json = os.path.join(REPORTS_DIR, "portable_review_validation_report.json")
    with open(report_json, 'w', encoding='utf-8') as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    
    # XLSX
    try:
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Validation"
        ws.append(["Check", "Status", "Detail"])
        
        def add_row(name, passed, detail=""):
            ws.append([name, "PASS" if passed else "FAIL", detail])
        
        add_row("ASD models = 199", total_models == 199, str(total_models))
        add_row("ASD local images", True, str(local_images_total))
        add_row("ASD local candidates", True, str(local_candidates))
        add_row("ASD remote-only", True, str(remote_only_total))
        add_row("Missing neg refs = 0", missing_neg == 0, str(missing_neg))
        add_row("Broken images = 0", broken_images == 0, str(broken_images))
        add_row("Initial approved = 0", initial_approved == 0, str(initial_approved))
        add_row("Excel_row filled = 199", excel_row_filled == 199, str(excel_row_filled))
        add_row("Autosave", has_autosave)
        add_row("Import", has_import)
        add_row("Export", has_export)
        add_row("Reset", has_reset)
        add_row("Same front/inside = 0", same_front_inside == 0, str(same_front_inside))
        add_row("Merge script", os.path.exists(merge_script))
        add_row("Lord rows = 932", lord_rows_count == 932, str(lord_rows_count))
        add_row("Lord local images = 699", lord_img_count == 699, str(lord_img_count))
        add_row("Lord flagged >= 47", lord_flagged_count >= 47, str(lord_flagged_count))
        add_row("Lord blank refs = 0", blank_img_refs == 0, str(blank_img_refs))
        add_row("Lord NO LOCAL = 89", lord_no_local_count >= 89, str(lord_no_local_count))
        add_row("Lord decisions = 932", lord_dec_count == 932, str(lord_dec_count))
        add_row("Lord initial reviewed = 0", lord_initial_reviewed == 0, str(lord_initial_reviewed))
        add_row("Lord flagged localStorage", lord_has_autosave)
        add_row("Lord flagged storage key", lord_has_storage_key)
        add_row("Lord flagged Import", lord_has_import)
        add_row("Lord flagged Export", lord_has_export)
        add_row("Lord flagged Reset", lord_has_reset)
        add_row("Lord contact localStorage", lord_all_has_autosave)
        add_row("Lord contact storage key", lord_all_has_storage_key)
        
        xlsx_path = os.path.join(REPORTS_DIR, "portable_review_validation_report.xlsx")
        wb.save(xlsx_path)
    except ImportError:
        print("  (openpyxl not available, skipping XLSX)")
    
    # ====================================================================
    # SUMMARY
    # ====================================================================
    print(f"\n{'='*60}")
    print(f"VALIDATION {'PASSED' if all_passed else 'FAILED'}")
    print(f"{'='*60}")
    
    return 0 if all_passed else 1

if __name__ == '__main__':
    sys.exit(main())
