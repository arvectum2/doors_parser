"""
Stage B - Merge manual decisions from 8 ASD batches + Lord.
Strict validation: only explicitly approved decisions accepted.
Output: merged JSON + validation report.
Usage: python scripts/stage_b/merge_manual_decisions.py
"""
import json, csv, os, sys
from collections import defaultdict
from datetime import datetime

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"

def parse_args():
    args = {
        'batches_dir': os.path.join(BASE, "output", "manual_photo_review_v5_7", "portable_batches"),
        'lord_dir': os.path.join(BASE, "output", "manual_photo_review_v5_7", "lord_portable_review"),
        'output': os.path.join(BASE, "output", "manual_photo_review_v5_7", "reports"),
    }
    i = 1
    while i < len(sys.argv):
        if sys.argv[i] == '--batches-dir' and i + 1 < len(sys.argv):
            args['batches_dir'] = sys.argv[i + 1]; i += 2
        elif sys.argv[i] == '--lord-dir' and i + 1 < len(sys.argv):
            args['lord_dir'] = sys.argv[i + 1]; i += 2
        elif sys.argv[i] == '--output' and i + 1 < len(sys.argv):
            args['output'] = sys.argv[i + 1]; i += 2
        else:
            i += 1
    return args

def load_decision_file(path):
    if not os.path.exists(path):
        return None
    ext = os.path.splitext(path)[1].lower()
    if ext == '.json':
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    elif ext == '.csv':
        with open(path, encoding='utf-8-sig') as f:
            return list(csv.DictReader(f, delimiter=';'))
    return None

def validate_asd_decisions(decisions, candidates_db):
    """Strict ASD validation."""
    issues = []
    model_names = set()
    excel_rows = set()
    
    # Build candidate URL sets per model
    candidate_urls = {}
    for m in candidates_db:
        name = m['model_name']
        urls = set()
        for c in m.get('candidates', []):
            urls.add(c['url'])
        candidate_urls[name] = urls
    
    for i, d in enumerate(decisions):
        name = d.get('model_name', '')
        front = d.get('approved_front_url', '').strip()
        inside = d.get('approved_inside_url', '').strip()
        front_explicit = d.get('front_explicitly_approved', False) or str(d.get('front_approved', 0)) in ('1', 'True', 'true')
        inside_explicit = d.get('inside_explicitly_approved', False) or str(d.get('inside_approved', 0)) in ('1', 'True', 'true')
        pair_approved = str(d.get('pair_approved', 0)) in ('1', 'True', 'true')
        needs_review = str(d.get('needs_review', 0)) in ('1', 'True', 'true')
        excel_row = d.get('excel_row', '')
        
        if not name:
            issues.append(f"Row {i}: missing model_name")
            continue
        
        if name in model_names:
            issues.append(f"Duplicate model_name: {name}")
        model_names.add(name)
        
        if excel_row:
            if excel_row in excel_rows:
                issues.append(f"Duplicate excel_row: {excel_row} for {name}")
            excel_rows.add(excel_row)
        
        # Check: at least one explicit action required
        has_action = front_explicit or inside_explicit or needs_review
        if not has_action:
            issues.append(f"{name}: not reviewed (no explicit action)")
        
        # Check: same front/inside
        if front and inside and front == inside:
            issues.append(f"{name}: same front and inside URL")
        
        # Check: pair_approved requires both explicit
        if pair_approved and (not front_explicit or not inside_explicit):
            issues.append(f"{name}: pair_approved but missing explicit front/inside")
        
        # Check: cannot be both approved and needs_review
        if pair_approved and needs_review:
            issues.append(f"{name}: both pair_approved and needs_review")
        
        # Check: approved URL must be in candidates (unless manual replacement)
        if front_explicit and front:
            manual = d.get('manual_replacement_approved', False) or d.get('front_manual_replacement', False)
            if not manual and name in candidate_urls and front not in candidate_urls[name]:
                issues.append(f"{name}: front URL not in candidates: {front[:60]}")
        if inside_explicit and inside:
            manual = d.get('manual_replacement_approved', False) or d.get('inside_manual_replacement', False)
            if not manual and name in candidate_urls and inside not in candidate_urls[name]:
                issues.append(f"{name}: inside URL not in candidates: {inside[:60]}")
    
    if len(model_names) != 199:
        issues.append(f"Expected 199 ASD models, got {len(model_names)}")
    
    return issues

def validate_lord_decisions(decisions):
    """Strict Lord validation."""
    issues = []
    allowed_decisions = {
        'keep_full_door', 'reject_fragment', 'reject_detail', 'reject_wrong_model',
        'clear_no_replacement', 'replace_with_approved_url', 'needs_review', ''
    }
    
    for i, d in enumerate(decisions):
        name = d.get('model_name', '')
        reviewed = d.get('explicitly_reviewed', False) or str(d.get('explicitly_reviewed', 0)) in ('1', 'True', 'true')
        decision = d.get('decision', '').strip()
        
        if not reviewed:
            issues.append(f"Lord row {i} ({name}): not reviewed")
            continue
        
        if decision not in allowed_decisions:
            issues.append(f"Lord row {i} ({name}): invalid decision '{decision}'")
        
        if decision == 'replace_with_approved_url':
            if not d.get('approved_replacement_url', '').strip():
                issues.append(f"Lord row {i} ({name}): replace_with_approved_url but no replacement URL")
        
        if decision == 'keep_full_door':
            if not d.get('current_image_url', '').strip():
                issues.append(f"Lord row {i} ({name}): keep_full_door but no current_image_url")
    
    if len(decisions) != 932:
        issues.append(f"Expected 932 Lord rows, got {len(decisions)}")
    
    return issues

def main():
    args = parse_args()
    batches_dir = args['batches_dir']
    lord_dir = args['lord_dir']
    output_dir = args['output']
    os.makedirs(output_dir, exist_ok=True)
    
    print("=" * 60)
    print("STAGE B - MERGE MANUAL DECISIONS (strict)")
    print("=" * 60)
    
    # Load candidates DB for URL validation
    reports_dir = os.path.join(BASE, "output", "manual_photo_review_v5_7", "reports")
    candidates_db = []
    cdb_path = os.path.join(reports_dir, "asd_candidate_database.json")
    if os.path.exists(cdb_path):
        with open(cdb_path, encoding='utf-8') as f:
            candidates_db = json.load(f)
    
    # ====================================================================
    # 1. Merge ASD decisions
    # ====================================================================
    print("\n--- ASD batches ---")
    all_asd = []
    
    for batch_num in range(1, 9):
        label = f"batch_{batch_num:02d}"
        batch_dir = os.path.join(batches_dir, label)
        
        json_path = os.path.join(batch_dir, f"manual_decisions_{label}.json")
        csv_path = os.path.join(batch_dir, f"manual_decisions_{label}.csv")
        
        decisions = load_decision_file(json_path) or load_decision_file(csv_path)
        if decisions is None:
            print(f"  {label}: NO DECISION FILE FOUND")
            continue
        
        print(f"  {label}: {len(decisions)} decisions loaded")
        all_asd.extend(decisions)
    
    print(f"\n  Total ASD decisions: {len(all_asd)}")
    
    if not all_asd:
        print("  ERROR: No ASD decisions found. Aborting.")
        sys.exit(1)
    
    # Validate
    print("\n  Validating ASD...")
    asd_issues = validate_asd_decisions(all_asd, candidates_db)
    
    # Count statuses
    asd_approved = sum(1 for d in all_asd if str(d.get('pair_approved', 0)) in ('1', 'True', 'true'))
    asd_needs_review = sum(1 for d in all_asd if str(d.get('needs_review', 0)) in ('1', 'True', 'true'))
    asd_front_only = sum(1 for d in all_asd
                         if (d.get('front_explicitly_approved', False) or str(d.get('front_approved', 0)) in ('1', 'True', 'true'))
                         and not (d.get('inside_explicitly_approved', False) or str(d.get('inside_approved', 0)) in ('1', 'True', 'true')))
    asd_inside_only = sum(1 for d in all_asd
                          if not (d.get('front_explicitly_approved', False) or str(d.get('front_approved', 0)) in ('1', 'True', 'true'))
                          and (d.get('inside_explicitly_approved', False) or str(d.get('inside_approved', 0)) in ('1', 'True', 'true')))
    asd_not_reviewed = sum(1 for d in all_asd
                           if not (d.get('front_explicitly_approved', False) or str(d.get('front_approved', 0)) in ('1', 'True', 'true'))
                           and not (d.get('inside_explicitly_approved', False) or str(d.get('inside_approved', 0)) in ('1', 'True', 'true'))
                           and not str(d.get('needs_review', 0)) in ('1', 'True', 'true'))
    
    if asd_issues:
        print(f"  ASD issues ({len(asd_issues)}):")
        for iss in asd_issues[:20]:
            print(f"    - {iss}")
        if len(asd_issues) > 20:
            print(f"    ... and {len(asd_issues) - 20} more")
    
    # Save ASD merged
    asd_output = os.path.join(output_dir, "asd_manual_decisions_merged.json")
    with open(asd_output, 'w', encoding='utf-8') as f:
        json.dump(all_asd, f, ensure_ascii=False, indent=2)
    print(f"\n  Saved: {asd_output}")
    
    # ====================================================================
    # 2. Merge Lord decisions
    # ====================================================================
    print("\n--- Lord ---")
    lord_json = os.path.join(lord_dir, "lord_manual_decisions.json")
    lord_csv = os.path.join(lord_dir, "lord_manual_decisions.csv")
    
    lord_decisions = load_decision_file(lord_json) or load_decision_file(lord_csv)
    
    lord_issues = []
    lord_kept = 0
    lord_rejected = 0
    lord_replaced = 0
    lord_needs_review = 0
    lord_not_reviewed = 0
    
    if lord_decisions:
        print(f"  Lord decisions loaded: {len(lord_decisions)}")
        
        lord_issues = validate_lord_decisions(lord_decisions)
        
        for d in lord_decisions:
            reviewed = d.get('explicitly_reviewed', False) or str(d.get('explicitly_reviewed', 0)) in ('1', 'True', 'true')
            decision = d.get('decision', '').strip()
            if not reviewed:
                lord_not_reviewed += 1
            elif decision == 'keep_full_door':
                lord_kept += 1
            elif decision.startswith('reject'):
                lord_rejected += 1
            elif decision == 'replace_with_approved_url':
                lord_replaced += 1
            elif decision == 'needs_review':
                lord_needs_review += 1
        
        if lord_issues:
            print(f"  Lord issues ({len(lord_issues)}):")
            for iss in lord_issues[:20]:
                print(f"    - {iss}")
        
        # Save Lord merged
        lord_output = os.path.join(output_dir, "lord_manual_decisions_merged.json")
        with open(lord_output, 'w', encoding='utf-8') as f:
            json.dump(lord_decisions, f, ensure_ascii=False, indent=2)
        print(f"  Saved: {lord_output}")
    else:
        print(f"  No Lord decision file found")
    
    # ====================================================================
    # 3. Save validation report
    # ====================================================================
    all_issues = asd_issues + lord_issues
    validation = {
        'timestamp': datetime.now().isoformat(),
        'passed': len(all_issues) == 0,
        'asd': {
            'total': len(all_asd),
            'approved_both': asd_approved,
            'front_only': asd_front_only,
            'inside_only': asd_inside_only,
            'needs_review': asd_needs_review,
            'not_reviewed': asd_not_reviewed,
            'issues': len(asd_issues),
        },
        'lord': {
            'total': len(lord_decisions) if lord_decisions else 0,
            'kept': lord_kept,
            'rejected': lord_rejected,
            'replaced': lord_replaced,
            'needs_review': lord_needs_review,
            'not_reviewed': lord_not_reviewed,
            'issues': len(lord_issues),
        },
        'issues': all_issues[:50],
    }
    
    validation_json = os.path.join(output_dir, "manual_decisions_validation.json")
    with open(validation_json, 'w', encoding='utf-8') as f:
        json.dump(validation, f, ensure_ascii=False, indent=2)
    
    # Try XLSX
    try:
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Validation"
        ws.append(["Category", "Metric", "Value"])
        ws.append(["ASD", "Total", len(all_asd)])
        ws.append(["ASD", "Approved both", asd_approved])
        ws.append(["ASD", "Front only", asd_front_only])
        ws.append(["ASD", "Inside only", asd_inside_only])
        ws.append(["ASD", "Needs review", asd_needs_review])
        ws.append(["ASD", "Not reviewed", asd_not_reviewed])
        ws.append(["ASD", "Issues", len(asd_issues)])
        ws.append(["Lord", "Total", len(lord_decisions) if lord_decisions else 0])
        ws.append(["Lord", "Kept", lord_kept])
        ws.append(["Lord", "Rejected", lord_rejected])
        ws.append(["Lord", "Replaced", lord_replaced])
        ws.append(["Lord", "Needs review", lord_needs_review])
        ws.append(["Lord", "Not reviewed", lord_not_reviewed])
        ws.append(["Lord", "Issues", len(lord_issues)])
        ws.append([])
        ws.append(["Issues"])
        for iss in all_issues:
            ws.append([iss])
        xlsx_path = os.path.join(output_dir, "manual_decisions_validation.xlsx")
        wb.save(xlsx_path)
    except ImportError:
        pass
    
    # ====================================================================
    # 4. Summary
    # ====================================================================
    print(f"\n{'='*60}")
    print("MERGE COMPLETE")
    print(f"{'='*60}")
    print(f"\n  ASD: {len(all_asd)} models")
    print(f"    Approved both:    {asd_approved}")
    print(f"    Front only:       {asd_front_only}")
    print(f"    Inside only:      {asd_inside_only}")
    print(f"    Needs review:     {asd_needs_review}")
    print(f"    Not reviewed:     {asd_not_reviewed}")
    print(f"    Issues:           {len(asd_issues)}")
    print(f"\n  Lord: {len(lord_decisions) if lord_decisions else 0} rows")
    print(f"    Kept:             {lord_kept}")
    print(f"    Rejected:         {lord_rejected}")
    print(f"    Replaced:         {lord_replaced}")
    print(f"    Needs review:     {lord_needs_review}")
    print(f"    Not reviewed:     {lord_not_reviewed}")
    print(f"    Issues:           {len(lord_issues)}")
    print(f"\n  Output files:")
    print(f"    {asd_output}")
    if lord_decisions:
        print(f"    {lord_output}")
    print(f"    {validation_json}")
    
    if all_issues:
        print(f"\n  BLOCKED: {len(all_issues)} validation issue(s). Fix before Stage B.")
        sys.exit(1)
    else:
        print(f"\n  Ready for Stage B apply.")

if __name__ == '__main__':
    main()
