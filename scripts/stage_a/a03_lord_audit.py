"""
Stage A - Step 3: Full visual audit of all 932 Lord rows.
Downloads images, prepares VLM input, builds flagged report.
"""
import csv, json, os, hashlib, re, html as html_mod
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import defaultdict
from datetime import datetime
import requests
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import openpyxl
import urllib3
urllib3.disable_warnings()

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
V56_DIR = os.path.join(BASE, "output", "delivery_v5_6")
CACHE_DIR = os.path.join(BASE, "output", "manual_photo_review_v5_7", "cache", "lord")
REPORTS_DIR = os.path.join(BASE, "output", "manual_photo_review_v5_7", "reports")
os.makedirs(CACHE_DIR, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)

session = requests.Session()
session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
session.verify = False
session.timeout = 20

# ========================================================================
# 1. Load catalog
# ========================================================================
print("Loading v5.6 catalog...")
with open(os.path.join(V56_DIR, "doors_catalog.csv"), encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))

LORD = chr(0x041B) + chr(0x043E) + chr(0x0440) + chr(0x0434)
lord_rows = [r for r in rows if r['producer'] == LORD]
print(f"Lord rows: {len(lord_rows)}")

# ========================================================================
# 2. Download Lord images
# ========================================================================
def safe_model_dir(name):
    safe = re.sub(r'[<>:"/\\|?*]', '_', name)
    safe = re.sub(r'\s+', '_', safe)
    safe = safe.strip('._') or 'unnamed'
    return safe[:120]

def download_image(url, dest_path):
    meta = {
        'url': url, 'downloaded': False, 'sha256': None,
        'width': None, 'height': None, 'content_type': None,
        'http_status': 0, 'error': None, 'file_size': 0,
    }
    try:
        resp = session.get(url, timeout=30, stream=True)
        meta['http_status'] = resp.status_code
        if resp.status_code != 200:
            meta['error'] = f'HTTP {resp.status_code}'
            return meta
        ct = resp.headers.get('Content-Type', '')
        meta['content_type'] = ct
        if not ct.startswith('image/'):
            meta['error'] = f'Not image: {ct}'
            return meta
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        sha = hashlib.sha256()
        size = 0
        with open(dest_path, 'wb') as f:
            for chunk in resp.iter_content(8192):
                if chunk:
                    f.write(chunk)
                    sha.update(chunk)
                    size += len(chunk)
        meta['sha256'] = sha.hexdigest()
        meta['file_size'] = size
        meta['downloaded'] = True
        try:
            from PIL import Image
            with Image.open(dest_path) as img:
                meta['width'], meta['height'] = img.size
        except Exception:
            pass
    except Exception as e:
        meta['error'] = str(e)
    return meta

print("\nDownloading Lord images...")
lord_images = []
download_tasks = []

for r in lord_rows:
    name = r['model_name']
    safe = safe_model_dir(name)
    model_dir = os.path.join(CACHE_DIR, safe)
    
    # Collect all image URLs for this Lord model
    seen = set()
    for col in ['image_url_front', 'image_url_inside']:
        url = r.get(col, '').strip()
        if url and url not in seen:
            seen.add(url)
            ext = (url.rsplit('.', 1)[-1] if '.' in url else 'jpg').split('?')[0][:10]
            filename = f"{col}_{hash(url)%10000:04x}.{ext}"
            dest = os.path.join(CACHE_DIR, safe, filename)
            download_tasks.append((url, dest, name, col))

print(f"Total Lord images to download: {len(download_tasks)}")

with ThreadPoolExecutor(max_workers=10) as ex:
    fut_map = {ex.submit(download_image, url, dest): (url, dest, name, col)
               for url, dest, name, col in download_tasks}
    completed = 0
    for fut in as_completed(fut_map):
        url, dest, name, col = fut_map[fut]
        try:
            meta = fut.result()
            meta['model_name'] = name
            meta['column'] = col
            meta['destination'] = dest
            lord_images.append(meta)
        except Exception as e:
            lord_images.append({
                'url': url, 'downloaded': False, 'error': str(e),
                'model_name': name, 'column': col, 'destination': dest
            })
        completed += 1
        if completed % 100 == 0:
            print(f"  Downloaded {completed}/{len(download_tasks)}")

dl_ok = sum(1 for d in lord_images if d.get('downloaded'))
dl_fail = sum(1 for d in lord_images if not d.get('downloaded'))
print(f"Downloaded: {dl_ok}, Failed: {dl_fail}")

# ========================================================================
# 3. Build VLM input for Lord
# ========================================================================
lord_vlm_input = []
for d in lord_images:
    if not d.get('downloaded'):
        continue
    lord_vlm_input.append({
        'model_name': d['model_name'],
        'column': d['column'],
        'image_url': d['url'],
        'local_path': d['destination'],
        'sha256': d.get('sha256', ''),
        'width': d.get('width'),
        'height': d.get('height'),
        'file_size': d.get('file_size', 0),
        'content_type': d.get('content_type', ''),
        'producer': 'Lord',
    })

lord_vlm_path = os.path.join(REPORTS_DIR, "lord_vlm_input.json")
with open(lord_vlm_path, 'w', encoding='utf-8') as f:
    json.dump(lord_vlm_input, f, ensure_ascii=False, indent=2)
print(f"Lord VLM input: {lord_vlm_path} ({len(lord_vlm_input)} candidates)")

# ========================================================================
# 4. Build Lord flagged report (using VLM output if available, or pattern-based)
# ========================================================================
vlm_path = os.path.join(REPORTS_DIR, "lord_vlm_output.json")
vlm_map = {}
if os.path.exists(vlm_path):
    with open(vlm_path, encoding='utf-8') as f:
        for v in json.load(f):
            vlm_map[v.get('url', v.get('image_url', ''))] = v
    print(f"Loaded Lord VLM output: {len(vlm_map)}")

# Build audit data
lord_audit = []
flagged = []

for r in lord_rows:
    name = r['model_name']
    f_url = r.get('image_url_front', '').strip()
    i_url = r.get('image_url_inside', '').strip()
    comment = r.get('comment', '')
    
    # Pattern-based flags (fallback when no VLM)
    pattern_flags = []
    for url, side in [(f_url, 'front'), (i_url, 'inside')]:
        if not url:
            continue
        url_lower = url.lower()
        # Fragment indicators
        if any(kw in url_lower for kw in ['fragment', 'detail', 'skol', 'closeup']):
            pattern_flags.append(f'{side}_url_fragment_keyword')
        # Handle/lock indicators
        if any(kw in url_lower for kw in ['handle', 'lock', 'ruchk', 'zamok', 'close-up']):
            pattern_flags.append(f'{side}_url_handle_keyword')
        # Video thumbnail
        if any(kw in url_lower for kw in ['video', 'youtube', 'vimeo', 'thumb']):
            pattern_flags.append(f'{side}_url_video_keyword')
        # Glass/detail indicators
        if any(kw in url_lower for kw in ['steklo', 'glass', 'vitrin', 'window']):
            pattern_flags.append(f'{side}_url_glass_keyword')
    
    # VLM-based flags
    vlm_flags = []
    for url, side in [(f_url, 'front'), (i_url, 'inside')]:
        if not url:
            continue
        v = vlm_map.get(url, {})
        if v:
            if v.get('has_door') == False:
                vlm_flags.append(f'{side}_vlm_no_door')
            if v.get('full_door_visible') == False:
                vlm_flags.append(f'{side}_vlm_not_full_door')
            if v.get('door_is_primary_object') == False:
                vlm_flags.append(f'{side}_vlm_not_primary')
            scene = v.get('scene_type', '')
            if scene in ('fragment', 'technical_detail'):
                vlm_flags.append(f'{side}_vlm_{scene}')
            if scene in ('handle_lock', 'hinge'):
                vlm_flags.append(f'{side}_vlm_handle_detail')
            if scene == 'banner_collage':
                vlm_flags.append(f'{side}_vlm_collage')
            if scene == 'video_thumbnail':
                vlm_flags.append(f'{side}_vlm_video')
            if scene == 'wide_interior':
                vlm_flags.append(f'{side}_vlm_wide_interior')
            if v.get('confidence', 1) < 0.9:
                vlm_flags.append(f'{side}_vlm_low_conf_{v.get("confidence",0):.2f}')
    
    all_flags = pattern_flags + vlm_flags
    
    audit_entry = {
        'model_name': name,
        'target_line': r.get('target_line', ''),
        'front_url': f_url,
        'inside_url': i_url,
        'has_front': bool(f_url),
        'has_inside': bool(i_url),
        'pattern_flags': '; '.join(pattern_flags) if pattern_flags else '',
        'vlm_flags': '; '.join(vlm_flags) if vlm_flags else '',
        'all_flags': '; '.join(all_flags) if all_flags else '',
        'flagged': len(all_flags) > 0,
        'comment': comment,
        'status': r.get('status', ''),
    }
    lord_audit.append(audit_entry)
    if audit_entry['flagged']:
        flagged.append(audit_entry)

# Write audit
audit_path = os.path.join(REPORTS_DIR, "lord_full_visual_audit.xlsx")
hdr_font = Font(bold=True, color="FFFFFF", size=10)
hdr_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Lord Audit"

headers = ['model_name', 'target_line', 'front_url', 'inside_url',
           'has_front', 'has_inside', 'pattern_flags', 'vlm_flags',
           'all_flags', 'flagged', 'comment', 'status']

for ci, h in enumerate(headers, 1):
    c = ws.cell(row=1, column=ci, value=h)
    c.font = hdr_font; c.fill = hdr_fill; c.alignment = hdr_align; c.border = thin_border

for ri, entry in enumerate(lord_audit, 2):
    for ci, h in enumerate(headers, 1):
        val = entry.get(h, '')
        if isinstance(val, bool): val = 1 if val else 0
        cell = ws.cell(row=ri, column=ci, value=val)
        cell.border = thin_border
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        if entry['flagged']:
            cell.fill = red_fill

wb.save(audit_path)
print(f"Lord audit XLSX: {audit_path} ({len(lord_audit)} rows, {len(flagged)} flagged)")

# ========================================================================
# 5. Build Lord flagged HTML review
# ========================================================================
html_parts = []
html_parts.append(f'''<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Lord Flagged Review v5.7</title>
<style>
body {{ font-family: Arial, sans-serif; margin: 20px; }}
h1 {{ font-size: 18px; }}
.stats {{ font-size: 12px; color: #666; margin-bottom: 12px; }}
table {{ border-collapse: collapse; width: 100%; font-size: 11px; }}
th, td {{ border: 1px solid #ccc; padding: 4px 6px; text-align: left; }}
th {{ background: #4472C4; color: white; position: sticky; top: 0; }}
tr.flagged {{ background: #fff5f5; }}
tr:nth-child(even) {{ background: #f9f9f9; }}
tr.flagged:nth-child(even) {{ background: #ffebeb; }}
img {{ max-width: 80px; max-height: 60px; object-fit: cover; }}
.flag {{ display: inline-block; font-size: 9px; padding: 1px 4px; border-radius: 2px; margin: 1px; }}
.flag-fragment {{ background: #f8d7da; color: #721c24; }}
.flag-handle {{ background: #fff3cd; color: #856404; }}
.flag-video {{ background: #d1ecf1; color: #0c5460; }}
.flag-vlm {{ background: #e2e3e5; color: #383d41; }}
.flag-glass {{ background: #cce5ff; color: #004085; }}
.filter-bar {{ margin-bottom: 8px; }}
.filter-bar input, .filter-bar select {{ font-size: 11px; padding: 3px 6px; }}
</style>
</head>
<body>
<h1>Lord Flagged Review — Visual Audit v5.7</h1>
<div class="stats">
    Total Lord: {len(lord_rows)} | Flagged: {len(flagged)} | Clean: {len(lord_rows) - len(flagged)}
    | VLM available: {len(vlm_map)} images
</div>
<div class="filter-bar">
    <input type="text" id="search" placeholder="Search model..." oninput="filterTable()">
    <select id="flagFilter" onchange="filterTable()">
        <option value="all">All flagged</option>
        <option value="fragment">Fragment</option>
        <option value="handle">Handle/lock</option>
        <option value="video">Video</option>
        <option value="glass">Glass detail</option>
        <option value="vlm">VLM rejected</option>
    </select>
</div>
<table id="lordTable">
<tr>
    <th>#</th><th>Model</th><th>Line</th><th>Front</th><th>Inside</th>
    <th>Pattern Flags</th><th>VLM Flags</th><th>Comment</th>
</tr>
''')

for idx, entry in enumerate(lord_audit):
    if not entry['flagged']:
        continue
    
    front_img = f'<img src="{html_mod.escape(entry["front_url"])}" onerror="this.style.display=\'none\'">' if entry.get('front_url') else '-'
    inside_img = f'<img src="{html_mod.escape(entry["inside_url"])}" onerror="this.style.display=\'none\'">' if entry.get('inside_url') else '-'
    
    # Categorize flags
    flag_html = ''
    for flag in entry.get('all_flags', '').split('; '):
        flag = flag.strip()
        if not flag: continue
        if any(kw in flag for kw in ['fragment', 'detail']):
            flag_html += f'<span class="flag flag-fragment">{html_mod.escape(flag)}</span> '
        elif any(kw in flag for kw in ['handle', 'lock', 'hinge', 'ruchk', 'zamok']):
            flag_html += f'<span class="flag flag-handle">{html_mod.escape(flag)}</span> '
        elif any(kw in flag for kw in ['video', 'youtube', 'vimeo']):
            flag_html += f'<span class="flag flag-video">{html_mod.escape(flag)}</span> '
        elif any(kw in flag for kw in ['glass', 'steklo', 'vitrin']):
            flag_html += f'<span class="flag flag-glass">{html_mod.escape(flag)}</span> '
        elif 'vlm_' in flag:
            flag_html += f'<span class="flag flag-vlm">{html_mod.escape(flag)}</span> '
        else:
            flag_html += f'<span class="flag">{html_mod.escape(flag)}</span> '
    
    pattern_flags = '; '.join(f for f in entry.get('pattern_flags','').split('; ') if f) if entry.get('pattern_flags') else ''
    vlm_flags = '; '.join(f for f in entry.get('vlm_flags','').split('; ') if f) if entry.get('vlm_flags') else ''
    
    html_parts.append(f'''<tr class="flagged" data-model="{html_mod.escape(entry["model_name"])}"
         data-flags="{html_mod.escape(entry.get('all_flags',''))}">
    <td>{idx+1}</td>
    <td title="{html_mod.escape(entry['model_name'])}">{html_mod.escape(entry['model_name'][:50])}</td>
    <td>{html_mod.escape(entry.get('target_line','')[:30])}</td>
    <td>{front_img}</td>
    <td>{inside_img}</td>
    <td><small>{html_mod.escape(pattern_flags[:80])}</small></td>
    <td><small>{html_mod.escape(vlm_flags[:80])}</small></td>
    <td><small>{html_mod.escape(entry.get('comment','')[:60])}</small></td>
</tr>''')

html_parts.append('''</table>
<script>
function filterTable() {
    const q = document.getElementById('search').value.toLowerCase();
    const flagFilter = document.getElementById('flagFilter').value;
    document.querySelectorAll('#lordTable tr.flagged').forEach(row => {
        const name = (row.cells[1]?.textContent || '').toLowerCase();
        const flags = row.dataset.flags || '';
        let show = true;
        if (q && !name.includes(q)) show = false;
        if (flagFilter !== 'all' && !flags.includes(flagFilter)) show = false;
        row.style.display = show ? '' : 'none';
    });
}
</script>
<p style="margin-top:16px;font-size:11px;color:#666;">
Generated: ''' + datetime.now().strftime('%Y-%m-%d %H:%M') + ''' — Lord Visual Audit v5.7<br>
Place VLM output at: ''' + vlm_path.replace('\\', '/') + ''' and re-run for classification results.
</p>
</body>
</html>''')

html_path = os.path.join(REPORTS_DIR, "lord_flagged_review.html")
with open(html_path, 'w', encoding='utf-8') as f:
    f.write(''.join(html_parts))
print(f"Lord flagged HTML: {html_path}")

# ========================================================================
# SUMMARY
# ========================================================================
print(f"\n{'='*60}")
print(f"STAGE A - STEP 3: LORD AUDIT COMPLETE")
print(f"{'='*60}")
print(f"  Lord total:           {len(lord_rows)}")
print(f"  Images downloaded:    {dl_ok}")
print(f"  Failed downloads:     {dl_fail}")
print(f"  Pattern-flagged:      {len(flagged)}")
print(f"  VLM input:            {lord_vlm_path}")
print(f"  Audit XLSX:           {audit_path}")
print(f"  Flagged HTML:         {html_path}")

# Pattern-flag breakdown
flag_counts = defaultdict(int)
for e in lord_audit:
    for flag in e.get('pattern_flags', '').split('; '):
        if flag: flag_counts[flag] += 1
if flag_counts:
    print(f"\n  Pattern flag breakdown:")
    for flag, count in sorted(flag_counts.items(), key=lambda x: -x[1]):
        print(f"    {flag}: {count}")
