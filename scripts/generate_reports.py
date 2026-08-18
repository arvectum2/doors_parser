import csv
import json
from collections import Counter

with open('output/final_rework_ready_v5/doors_catalog.csv', encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

# ========== Report 1: Komfort Rebuild ==========
kf = [r for r in rows if r['producer'] == 'Комфорт']
kf_lines = Counter(r['target_line'] for r in kf)
komfort_report = {
    'title': 'Komfort Rebuild Report (v5)',
    'total_models': len(kf),
    'collections': dict(sorted(kf_lines.items())),
    'coating': Counter(r['coating'] for r in kf),
    'image_url_ok': sum(1 for r in kf if r.get('image_url', '')),
    'source_url_ok': sum(1 for r in kf if r.get('source_url', '')),
    'details': []
}
for r in kf:
    komfort_report['details'].append({
        'model_name': r['model_name'],
        'target_line': r['target_line'],
        'coating': r['coating'],
        'image_url': r.get('image_url', ''),
        'source_url': r.get('source_url', ''),
    })
with open('output/final_rework_ready_v5/komfort_rebuild_report.json', 'w', encoding='utf-8') as f:
    json.dump(komfort_report, f, ensure_ascii=False, indent=2)
print(f"Komfort report: {len(kf)} models across {len(kf_lines)} collections")

# ========== Report 2: Hausdoors PVC Rebuild ==========
hd = [r for r in rows if 'Хаусдорс' in r['producer'] or 'аусд' in r['producer']]
hd_lines = Counter(r['target_line'] for r in hd)
pvc_collections = ['Фактурные линии', 'Новая классика', 'Минимализм', 'Современная классика']
pvc_models = [r for r in hd if r['target_line'] in pvc_collections]
non_pvc = [r for r in hd if r['target_line'] not in pvc_collections]

hausdoors_report = {
    'title': 'Hausdoors PVC Rebuild Report (v5)',
    'total_models': len(hd),
    'total_collections': len(hd_lines),
    'collections': dict(sorted(hd_lines.items())),
    'pvc_collections': {
        'total_models': len(pvc_models),
        'collections': {c: sum(1 for r in pvc_models if r['target_line'] == c) for c in pvc_collections},
    },
    'non_pvc_models': len(non_pvc),
    'pvc_details': []
}
for r in pvc_models:
    hausdoors_report['pvc_details'].append({
        'model_name': r['model_name'],
        'target_line': r['target_line'],
        'coating': r['coating'],
        'material': r['material'],
        'image_url': r.get('image_url_front', ''),
        'source_url': r.get('source_url', ''),
    })
with open('output/final_rework_ready_v5/hausdoors_pvh_rebuild_report.json', 'w', encoding='utf-8') as f:
    json.dump(hausdoors_report, f, ensure_ascii=False, indent=2)
print(f"Hausdoors report: {len(hd)} total, {len(pvc_models)} PVC, {len(non_pvc)} non-PVC")

# ========== Report 3: Olimp Missing Collections ==========
ol = [r for r in rows if r['producer'] == 'Олимп']
ol_lines = Counter(r['target_line'] for r in ol)
parizh = [r for r in ol if 'Париж' in r['target_line']]
olimp_report = {
    'title': 'Olimp Missing Collections Report (v5)',
    'total_models': len(ol),
    'collections': dict(sorted(ol_lines.items())),
    'parizh_added': {
        'count': len(parizh),
        'models': [{'model_name': r['model_name'], 'target_line': r['target_line']} for r in parizh],
    },
    'all_collections_present': sorted(ol_lines.keys()),
}
with open('output/final_rework_ready_v5/olimp_missing_collections_report.json', 'w', encoding='utf-8') as f:
    json.dump(olimp_report, f, ensure_ascii=False, indent=2)
print(f"Olimp report: {len(ol)} models, {len(parizh)} Paris models")

# ========== Report 4: Image URL Validation ==========
total_urls = set()
for r in rows:
    for field in ['image_url', 'image_url_front', 'image_url_inside']:
        url = r.get(field, '').strip()
        if url:
            total_urls.add(url)

rows_with_images = sum(1 for r in rows if r.get('image_url', '') or r.get('image_url_front', ''))
rows_without_images = sum(1 for r in rows if not r.get('image_url', '') and not r.get('image_url_front', ''))
parizh_no_img = sum(1 for r in parizh if not r.get('image_url', '') and not r.get('image_url_front', ''))

image_report = {
    'title': 'Image URL Validation Report (v5)',
    'total_rows': len(rows),
    'unique_urls': len(total_urls),
    'rows_with_images': rows_with_images,
    'rows_without_images': rows_without_images,
    'rows_without_images_breakdown': {},
    'note': 'Full HTTP validation done separately. Komfort URLs fixed via scraping. Olimp Paris images cleared (404 on manufacturer site).',
}

# Breakdown of rows without images
no_img_rows = [r for r in rows if not r.get('image_url', '') and not r.get('image_url_front', '')]
no_img_producers = Counter(r['producer'] for r in no_img_rows)
image_report['rows_without_images_breakdown'] = dict(no_img_producers)

with open('output/final_rework_ready_v5/image_url_validation_report.json', 'w', encoding='utf-8') as f:
    json.dump(image_report, f, ensure_ascii=False, indent=2)
print(f"Image report: {rows_with_images} with images, {rows_without_images} without")
print(f"  Breakdown: {dict(no_img_producers)}")
