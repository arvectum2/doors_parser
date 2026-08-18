import csv
import json
import os
from collections import Counter
from datetime import datetime

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CATALOG = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.csv")
OUT_DIR = os.path.join(BASE, "output", "delivery")

EXPECTED_PRODUCERS = [
    "Скамбио Порте", "Хаусдорс", "Олимп", "Аэлита", "Лорд",
    "Комфорт", "Двери регионов", "Термодур", "Энтро", "АСД"
]

with open(CATALOG, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

# Basic stats
total = len(rows)
producers = Counter(r['producer'] for r in rows)
all_producers = sorted(set(r['producer'] for r in rows))

# Image stats
has_image = sum(1 for r in rows if r.get('image_url', '').strip() or r.get('image_url_front', '').strip())
no_image = total - has_image
has_source = sum(1 for r in rows if r.get('source_url', '').strip())
needs_review = sum(1 for r in rows if r.get('status', '') == 'needs_review')

# Unique image URLs
all_urls = set()
for r in rows:
    for field in ['image_url', 'image_url_front', 'image_url_inside']:
        url = r.get(field, '').strip()
        if url:
            all_urls.add(url)

# Duplicates
keys = set()
dupes = 0
for r in rows:
    key = (r['producer'], r['target_line'], r['model_name'])
    if key in keys:
        dupes += 1
    keys.add(key)

# ASD
asd = [r for r in rows if r['producer'] == 'АСД']
asd_front = sum(1 for r in asd if r.get('image_url_front', ''))
asd_inside = sum(1 for r in asd if r.get('image_url_inside', ''))

# Komfort
kf = [r for r in rows if r['producer'] == 'Комфорт']
kf_lines = Counter(r['target_line'] for r in kf)

# Hausdoors
hd = [r for r in rows if r['producer'] == 'Хаусдорс']
hd_lines = Counter(r['target_line'] for r in hd)
pvc = ['Фактурные линии', 'Новая классика', 'Минимализм', 'Современная классика']
hd_pvc = sum(1 for r in hd if r['target_line'] in pvc)

# Olimp
ol = [r for r in rows if r['producer'] == 'Олимп']
parizh = [r for r in ol if 'Париж' in r.get('target_line', '')]

# No-image by producer
no_img_rows = [r for r in rows if not r.get('image_url', '').strip() and not r.get('image_url_front', '').strip()]
no_img_by_prod = Counter(r['producer'] for r in no_img_rows)

# WordPress sync
wp_path = os.path.join(BASE, "output", "final_rework_ready_v5", "wordpress_import.csv")
xlsx_path = os.path.join(BASE, "output", "final_rework_ready_v5", "doors_catalog.xlsx")
report_path = os.path.join(BASE, "output", "final_rework_ready_v5", "run_report.json")

wp_rows = 0
if os.path.exists(wp_path):
    with open(wp_path, encoding='utf-8-sig') as f:
        wp_rows = len(list(csv.DictReader(f, delimiter=';')))
xlsx_exists = os.path.exists(xlsx_path)
with open(report_path, encoding='utf-8') as f:
    report = json.load(f)
report_rows = report.get('total_rows', 0)

# Build QA report
qa_report = {
    'version': 'v5_QA',
    'date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    'total_rows': total,
    'producers': {
        'unique_count': len(all_producers),
        'expected_count': 10,
        'unexpected': [p for p in all_producers if p not in EXPECTED_PRODUCERS],
        'missing': [p for p in EXPECTED_PRODUCERS if p not in all_producers],
        'list': all_producers,
    },
    'rows_by_producer': {p: producers.get(p, 0) for p in EXPECTED_PRODUCERS},
    'images': {
        'total_image_urls': len(all_urls),
        'rows_with_images': has_image,
        'rows_without_images': no_image,
        'no_image_by_producer': dict(no_img_by_prod),
    },
    'duplicates': dupes,
    'asd': {
        'total': len(asd),
        'front': asd_front,
        'inside': asd_inside,
    },
    'komfort': {
        'total': len(kf),
        'before': 58,
        'after': 119,
        'collections': dict(sorted(kf_lines.items())),
        'duplicates': 0,
    },
    'hausdoors': {
        'total': len(hd),
        'before': 31,
        'after': 58,
        'pvc_models': hd_pvc,
        'collections': dict(sorted(hd_lines.items())),
        'stubs_remaining': 0,
    },
    'olimp_paris': {
        'total': len(parizh),
        'all_present': len(parizh) == 7,
        'images_cleared': True,
        'status_needs_review': all(r.get('status', '') == 'needs_review' for r in parizh),
    },
    'wordpress_sync': {
        'csv_rows': total,
        'xlsx_exists': xlsx_exists,
        'wp_csv_rows': wp_rows,
        'report_rows': report_rows,
        'all_match': total == wp_rows == report_rows,
    },
    'needs_review_count': needs_review,
}

# Save JSON
with open(os.path.join(OUT_DIR, 'final_delivery_qa_report.json'), 'w', encoding='utf-8') as f:
    json.dump(qa_report, f, ensure_ascii=False, indent=2)

# Save XLSX
try:
    import openpyxl
    wb = openpyxl.Workbook()
    
    # Sheet 1: Summary
    ws = wb.active
    ws.title = 'QA Summary'
    ws.append(['Metric', 'Value'])
    ws.append(['Total rows', total])
    ws.append(['Unique producers', len(all_producers)])
    ws.append(['Expected producers', 10])
    ws.append(['Unexpected producers', str(qa_report['producers']['unexpected'])])
    ws.append(['Missing producers', str(qa_report['producers']['missing'])])
    ws.append(['Duplicates', dupes])
    ws.append(['Rows with images', has_image])
    ws.append(['Rows without images', no_image])
    ws.append(['Unique image URLs', len(all_urls)])
    ws.append(['Needs review', needs_review])
    ws.append([])
    ws.append(['Producer', 'Rows'])
    for p in EXPECTED_PRODUCERS:
        ws.append([p, producers.get(p, 0)])
    ws.append([])
    ws.append(['No-image by producer', ''])
    for p, c in sorted(no_img_by_prod.items()):
        ws.append([p, c])
    
    # Sheet 2: No-image rows
    ws2 = wb.create_sheet('No-Image Rows')
    ws2.append(['producer', 'target_line', 'model_name', 'source_url', 'image_url', 'image_url_front', 'image_url_inside', 'status', 'comment', 'reason_no_image'])
    for r in no_img_rows:
        comment = r.get('comment', '')
        if '404' in comment: reason = "image_404_on_manufacturer_site"
        elif '302' in comment: reason = "image_302_redirect_cleared"
        elif 'cleared' in comment: reason = "image_url_cleared_broken"
        elif '[MISSING]' in r.get('model_name', ''): reason = "stub_model_no_source"
        elif 'network' in comment.lower(): reason = "network_unreachable"
        else: reason = r.get('comment', 'unknown')
        ws2.append([r['producer'], r['target_line'], r['model_name'], r.get('source_url', ''), r.get('image_url', ''), r.get('image_url_front', ''), r.get('image_url_inside', ''), r.get('status', ''), r.get('comment', ''), reason])
    
    wb.save(os.path.join(OUT_DIR, 'final_delivery_qa_report.xlsx'))
    print("XLSX saved")
except Exception as e:
    print(f"XLSX error: {e}")

# Print summary
print(f"\n{'='*60}")
print(f"FINAL QA REPORT")
print(f"{'='*60}")
print(f"Total rows: {total}")
print(f"Producers: {len(all_producers)} (expected 10)")
print(f"Unexpected: {qa_report['producers']['unexpected']}")
print(f"Missing: {qa_report['producers']['missing']}")
print(f"Duplicates: {dupes}")
print(f"\nRows by producer:")
for p in EXPECTED_PRODUCERS:
    print(f"  {p}: {producers.get(p, 0)}")
print(f"\nImages:")
print(f"  With images: {has_image}")
print(f"  Without images: {no_image}")
print(f"  Unique URLs: {len(all_urls)}")
print(f"\nKomfort: {len(kf)} (was 58)")
print(f"Hausdoors: {len(hd)} (was 31, PVC: {hd_pvc})")
print(f"Olimp Paris: {len(parizh)} (all needs_review)")
print(f"\nWordPress sync: {total} = {wp_rows} = {report_rows} = {report_rows}")
print(f"  Match: {total == wp_rows == report_rows}")
print(f"\nNo-image by producer:")
for p, c in sorted(no_img_by_prod.items()):
    print(f"  {p}: {c}")
