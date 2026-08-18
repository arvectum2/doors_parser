import csv
import json
import os
from collections import Counter
from datetime import datetime

with open('output/final_rework_ready_v5/doors_catalog.csv', encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

# Basic stats
total = len(rows)
producers = Counter(r['producer'] for r in rows)
has_image = sum(1 for r in rows if r.get('image_url', ''))
has_front = sum(1 for r in rows if r.get('image_url_front', ''))
has_inside = sum(1 for r in rows if r.get('image_url_inside', ''))
has_source = sum(1 for r in rows if r.get('source_url', ''))
stubs = sum(1 for r in rows if '[MISSING]' in r.get('model_name', '') or '[MISSING]' in r.get('target_line', ''))

# Duplicate check
keys = set()
dupes = 0
for r in rows:
    key = (r['producer'], r['target_line'], r['model_name'])
    if key in keys:
        dupes += 1
    keys.add(key)

# ASD stats
asd = [r for r in rows if r['producer'] == 'АСД']
asd_front = sum(1 for r in asd if r.get('image_url_front', ''))
asd_inside = sum(1 for r in asd if r.get('image_url_inside', ''))

report = {
    'version': 'v5',
    'date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    'total_rows': total,
    'producers': dict(sorted(producers.items())),
    'images': {
        'has_image_url': has_image,
        'has_image_url_front': has_front,
        'has_image_url_inside': has_inside,
    },
    'has_source_url': has_source,
    'stubs_missing': stubs,
    'duplicates': dupes,
    'asd': {
        'total': len(asd),
        'front': asd_front,
        'inside': asd_inside,
    },
    'changes_vs_v4': {
        'komfort': '58 -> 119 (scraped from komfortdoors.ru)',
        'hausdoors_pvc': '4 stubs -> 31 real models (scraped from hausdoors.ru)',
        'olimp_paris': '0 -> 7 models (added from olymp-doors.ru)',
    },
}

with open('output/final_rework_ready_v5/run_report.json', 'w', encoding='utf-8') as f:
    json.dump(report, f, ensure_ascii=False, indent=2)

print(json.dumps(report, ensure_ascii=False, indent=2))
