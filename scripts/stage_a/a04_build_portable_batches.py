"""
Stage A - Step 4: Build portable manual review batches (v5.7 rebuild).
Creates 8 self-contained ASD batches + Lord portable review package.
All paths relative, no server required, works from file://.
Fixes: state architecture, negative paths, autosave, thumbnails, Lord images, QA.
"""
import csv, json, os, re, shutil, hashlib, html as html_mod, io, zipfile
from collections import defaultdict
from datetime import datetime
from PIL import Image

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
V56_DIR = os.path.join(BASE, "output", "delivery_v5_6")
CACHE_ASD = os.path.join(BASE, "output", "manual_photo_review_v5_7", "cache", "asd")
CACHE_LORD = os.path.join(BASE, "output", "manual_photo_review_v5_7", "cache", "lord")
REPORTS_DIR = os.path.join(BASE, "output", "manual_photo_review_v5_7", "reports")
PORTABLE_DIR = os.path.join(BASE, "output", "manual_photo_review_v5_7", "portable_batches")
LORD_DIR = os.path.join(BASE, "output", "manual_photo_review_v5_7", "lord_portable_review")
NEGATIVE_DIR = os.path.join(PORTABLE_DIR, "negative_examples")
OUTPUT_BASE = os.path.join(BASE, "output", "manual_photo_review_v5_7")

NEGATIVE_EXAMPLES = {
    'IMG_1946.jpg': {'label': 'video_thumbnail_fragment', 'desc': 'Video thumbnail / fragment panel'},
    'IMG_1947.jpg': {'label': 'full_door_side_unknown', 'desc': 'Full door candidate, side unknown'},
    'IMG_1948.jpg': {'label': 'banner_collage', 'desc': 'Banner/collage'},
    'IMG_1949.jpg': {'label': 'wide_interior', 'desc': 'Wide interior scene'},
    'IMG_1951.jpg': {'label': 'full_door_side_unknown', 'desc': 'Full door candidate, side unknown'},
    'IMG_1952.jpg': {'label': 'full_door_mirror_unknown', 'desc': 'Full door with mirror, side unknown'},
    'IMG_1953.jpg': {'label': 'wide_interior', 'desc': 'Wide interior scene'},
    'IMG_1954.jpg': {'label': 'wide_interior', 'desc': 'Wide interior scene'},
    'IMG_1955.jpg': {'label': 'wide_interior', 'desc': 'Wide interior scene'},
    'IMG_1956.jpg': {'label': 'handle_lock_closeup', 'desc': 'Handle/lock close-up'},
}
NEGATIVE_ORDER = ['IMG_1946.jpg','IMG_1948.jpg','IMG_1949.jpg','IMG_1953.jpg','IMG_1954.jpg',
                  'IMG_1955.jpg','IMG_1956.jpg','IMG_1947.jpg','IMG_1951.jpg','IMG_1952.jpg']

THUMB_MAX_SIDE = 1200
THUMB_QUALITY = 82

# ========================================================================
# Helpers
# ========================================================================
def safe_filename(url, maxlen=60):
    name = url.rsplit('/', 1)[-1] if '/' in url else url
    name = name.split('?')[0][:maxlen]
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    name = re.sub(r'\s+', '_', name)
    return name or 'image.jpg'

def short_id(url):
    return hashlib.md5(url.encode()).hexdigest()[:8]

def make_thumbnail(src_path, dst_path, max_side=THUMB_MAX_SIDE, quality=THUMB_QUALITY):
    try:
        img = Image.open(src_path)
        w, h = img.size
        if max(w, h) <= max_side:
            shutil.copy2(src_path, dst_path)
            return True
        if w >= h:
            new_w, new_h = max_side, int(h * max_side / w)
        else:
            new_w, new_h = int(w * max_side / h), max_side
        img = img.resize((new_w, new_h), Image.LANCZOS)
        ext = os.path.splitext(dst_path)[1].lower()
        if ext in ('.jpg', '.jpeg'):
            img.convert('RGB').save(dst_path, 'JPEG', quality=quality, optimize=True)
        elif ext == '.webp':
            img.save(dst_path, 'WEBP', quality=quality, method=4)
        elif ext == '.png':
            img.save(dst_path, 'PNG', optimize=True)
        else:
            img.convert('RGB').save(dst_path, 'JPEG', quality=quality, optimize=True)
        return True
    except Exception:
        try:
            shutil.copy2(src_path, dst_path)
        except Exception:
            return False
        return False

def file_sha256(path):
    h = hashlib.sha256()
    try:
        with open(path, 'rb') as f:
            for chunk in iter(lambda: f.read(65536), b''):
                h.update(chunk)
        return h.hexdigest()[:16]
    except Exception:
        return ''

def img_dimensions(path):
    try:
        with Image.open(path) as img:
            return img.size
    except Exception:
        return (0, 0)

def esc(s):
    return String(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;') if False else \
        str(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')

# ========================================================================
# Load data
# ========================================================================
print("Loading data...")
with open(os.path.join(V56_DIR, "doors_catalog.csv"), encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
asd_rows = [r for r in rows if r['producer'] == 'АСД']
asd_map = {r['model_name']: r for r in asd_rows}

with open(os.path.join(REPORTS_DIR, "asd_candidate_database.json"), encoding='utf-8') as f:
    candidates_db = json.load(f)

with open(os.path.join(REPORTS_DIR, "asd_download_results.json"), encoding='utf-8') as f:
    dl_results = json.load(f)

dl_by_url = {}
for d in dl_results:
    dl_by_url[d['url']] = d

with open(os.path.join(BASE, "output", "delivery_v5_2", "asd_final_photo_source_validation.json"), encoding='utf-8') as f:
    qa_v52 = {r['model_name']: r for r in json.load(f)}

# ========================================================================
# Build model list
# ========================================================================
model_list = sorted(candidates_db, key=lambda m: (
    0 if m.get('current_front') or m.get('current_inside') else 1,
    m.get('model_name', '')
))
for i, m in enumerate(model_list):
    m['list_index'] = i + 1

BATCH_SIZE = 25
batches = []
for i in range(0, len(model_list), BATCH_SIZE):
    batch_num = len(batches) + 1
    batch_models = model_list[i:i+BATCH_SIZE]
    batches.append({
        'num': batch_num,
        'label': f'batch_{batch_num:02d}',
        'dir': os.path.join(PORTABLE_DIR, f'batch_{batch_num:02d}'),
        'models': batch_models,
        'start_idx': batch_models[0]['list_index'],
        'end_idx': batch_models[-1]['list_index'],
    })

print(f"ASD rows: {len(asd_rows)}, Batches: {len(batches)}")
for b in batches:
    print(f"  {b['label']}: models {b['start_idx']}-{b['end_idx']} ({len(b['models'])} models)")

# ========================================================================
# Copy negative examples
# ========================================================================
print("\nCopying negative examples...")
os.makedirs(NEGATIVE_DIR, exist_ok=True)
for fname in NEGATIVE_EXAMPLES:
    for src_dir in [os.path.join(BASE, '..', 'Downloads'), 'C:\\Users\\GN\\Downloads']:
        src = os.path.join(src_dir, fname)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(NEGATIVE_DIR, fname))
            break

# ========================================================================
# Build negative section HTML (FIX #2: relative path)
# ========================================================================
def build_negative_section():
    parts = ['<div class="negative-panel">',
             '<h3>Не выбирать — примеры неподходящих изображений</h3>',
             '<p class="negative-desc">Следующие типы изображений НЕ ДОЛЖНЫ быть выбраны как front или inside:</p>',
             '<div class="negative-grid">']
    for fname in NEGATIVE_ORDER:
        info = NEGATIVE_EXAMPLES.get(fname, {})
        label = info.get('label', 'unknown')
        desc = info.get('desc', '')
        parts.append(f'<div class="negative-card"><img src="negative_examples/{fname}" loading="lazy">'
                     f'<div class="neg-label">{html_mod.escape(label)}</div>'
                     f'<div class="neg-desc">{html_mod.escape(desc)}</div></div>')
    parts.append('</div></div>')
    return '\n'.join(parts)

# ========================================================================
# FIX #11: Generate thumbnails for batch images
# ========================================================================
def generate_batch_thumbnails(batch_dir):
    images_dir = os.path.join(batch_dir, 'images')
    thumbs_dir = os.path.join(batch_dir, 'thumbs')
    os.makedirs(thumbs_dir, exist_ok=True)
    thumb_map = {}
    for fname in os.listdir(images_dir):
        fpath = os.path.join(images_dir, fname)
        if not os.path.isfile(fpath):
            continue
        thumb_path = os.path.join(thumbs_dir, fname)
        if not os.path.exists(thumb_path):
            make_thumbnail(fpath, thumb_path)
        thumb_map[fname] = 'thumbs/' + fname
    return thumb_map

# ========================================================================
# FIX #1: Build review.html with correct state architecture
# ========================================================================
def build_review_html(batch, negative_section, thumb_map):
    models_js = []
    for m in batch['models']:
        name = m['model_name']
        row = asd_map.get(name, {})
        qa = qa_v52.get(name, {})
        
        candidates_js = []
        for cand in m.get('candidates', []):
            url = cand['url']
            dl = dl_by_url.get(url, {})
            local_file = ''
            thumb_file = ''
            dest = dl.get('destination', '')
            if dest and os.path.exists(dest):
                fname_full = safe_filename(url) + '_' + short_id(url) + os.path.splitext(dest)[1]
                local_file = 'images/' + fname_full
                thumb_file = thumb_map.get(fname_full, 'images/' + fname_full)
            
            candidates_js.append({
                'url': url,
                'local': local_file,
                'thumb': thumb_file,
                'source_type': cand.get('source_type', ''),
                'gallery_index': cand.get('gallery_index', 0),
                'width': dl.get('width'),
                'height': dl.get('height'),
                'file_size': dl.get('file_size', 0),
                'content_type': dl.get('content_type', ''),
                'downloaded': dl.get('downloaded', False),
                'original_filename': cand.get('original_filename', ''),
                'gallery_info': cand.get('gallery_info', {}),
                'parser_info': cand.get('parser_info', {}),
            })
        
        excel_row = ''
        for r_idx, r in enumerate(rows):
            if r['model_name'] == name:
                excel_row = r_idx + 2
                break
        
        models_js.append({
            'idx': m['list_index'],
            'excel_row': excel_row,
            'producer': 'ASD',
            'target_line': row.get('target_line', ''),
            'model_name': name,
            'source_url': row.get('source_url', ''),
            'current_front': m.get('current_front', ''),
            'current_inside': m.get('current_inside', ''),
            'comment': row.get('comment', ''),
            'status': row.get('status', ''),
            'qa_front': qa.get('front_visual_decision', ''),
            'qa_inside': qa.get('inside_visual_decision', ''),
            'candidates': candidates_js,
        })
    
    models_json = json.dumps(models_js, ensure_ascii=False)
    batch_num = batch['num']
    
    html = f'''<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ASD Batch {batch_num} — Manual Photo Review</title>
<style>
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
       background: #f5f6fa; color: #2c3e50; font-size: 13px; }}
.header {{ background: #2c3e50; color: #ecf0f1; padding: 12px 16px; position: sticky; top: 0; z-index: 1000; }}
.header h1 {{ font-size: 16px; }}
.header .sub {{ font-size: 11px; opacity: 0.8; }}
.controls {{ background: #fff; padding: 8px 16px; border-bottom: 1px solid #ddd;
            position: sticky; top: 52px; z-index: 999; display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }}
.controls input, .controls select {{ font-size: 12px; padding: 3px 8px; border: 1px solid #ccc; border-radius: 3px; }}
.controls button {{ font-size: 11px; padding: 4px 10px; border: 1px solid #3498db; background: #3498db;
                   color: #fff; border-radius: 3px; cursor: pointer; }}
.controls button:hover {{ background: #2980b9; }}
.controls .btn-export {{ background: #27ae60; border-color: #27ae60; }}
.controls .btn-export:hover {{ background: #219a52; }}
.controls .btn-import {{ background: #8e44ad; border-color: #8e44ad; }}
.controls .btn-import:hover {{ background: #7d3c98; }}
.controls .btn-reset {{ background: #e74c3c; border-color: #e74c3c; }}
.controls .btn-reset:hover {{ background: #c0392b; }}
.model-card {{ background: #fff; border: 1px solid #ddd; border-radius: 5px; margin: 10px 16px;
              padding: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.06); }}
.model-header {{ display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 4px;
                margin-bottom: 6px; }}
.model-title {{ font-size: 14px; font-weight: 600; flex: 1; min-width: 200px; }}
.model-meta {{ font-size: 11px; color: #7f8c8d; }}
.model-status {{ font-size: 10px; padding: 2px 8px; border-radius: 3px; white-space: nowrap; }}
.status-ok {{ background: #d5f5e3; color: #1e8449; }}
.status-review {{ background: #fdebd0; color: #a04000; }}
.status-empty {{ background: #fadbd8; color: #922b21; }}
.status-partial {{ background: #d6eaf8; color: #1a5276; }}
.current-pair {{ display: flex; gap: 10px; margin: 6px 0 8px; flex-wrap: wrap; }}
.current-pair .pair-item {{ text-align: center; }}
.current-pair img {{ max-width: 110px; max-height: 90px; object-fit: contain; border: 1px solid #ddd;
                    border-radius: 3px; background: #fafafa; }}
.pair-label {{ font-size: 10px; color: #7f8c8d; margin-top: 2px; }}
.pair-label.not-approved {{ color: #e74c3c; font-weight: 600; }}
.approve-current {{ font-size: 9px; padding: 2px 6px; border: 1px solid #27ae60; background: #d5f5e3;
                    color: #1e8449; border-radius: 3px; cursor: pointer; margin-top: 3px; }}
.approve-current:hover {{ background: #27ae60; color: #fff; }}
.candidates-section {{ margin-top: 6px; }}
.candidates-section h4 {{ font-size: 12px; margin-bottom: 4px; color: #34495e; }}
.candidates-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(165px, 1fr)); gap: 6px; }}
.candidate {{ border: 1px solid #e0e0e0; border-radius: 4px; padding: 4px; background: #fafafa;
             position: relative; transition: all 0.15s; }}
.candidate:hover {{ border-color: #3498db; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }}
.candidate.selected-as-front {{ border-color: #2980b9; background: #ebf5fb; }}
.candidate.selected-as-inside {{ border-color: #27ae60; background: #eafaf1; }}
.candidate.rejected {{ opacity: 0.5; border-color: #e74c3c; background: #fdedec; }}
.candidate-img {{ width: 100%; height: 100px; object-fit: contain; border-radius: 2px; background: #f0f0f0; }}
.candidate.remote-missing {{ border: 2px dashed #e67e22; background: #fef9e7; }}
.remote-placeholder {{ background: #fef9e7; border: 1px solid #f9e79f; border-radius: 4px; padding: 8px;
                       text-align: center; font-size: 10px; color: #7d6608; }}
.remote-placeholder .remote-title {{ font-weight: 600; color: #e67e22; margin-bottom: 4px; }}
.remote-placeholder a {{ color: #3498db; font-size: 9px; }}
.candidate-info {{ font-size: 9px; color: #7f8c8d; margin-top: 2px; }}
.candidate-info .src-type {{ white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
.candidate-actions {{ display: flex; gap: 2px; margin-top: 3px; }}
.candidate-actions button {{ font-size: 9px; padding: 2px 5px; border: none; border-radius: 3px;
                             cursor: pointer; flex: 1; white-space: nowrap; }}
.btn-front {{ background: #d4e6f1; color: #1a5276; }}
.btn-front.active, .btn-front:hover {{ background: #2980b9; color: #fff; }}
.btn-inside {{ background: #d5f5e3; color: #1e8449; }}
.btn-inside.active, .btn-inside:hover {{ background: #27ae60; color: #fff; }}
.btn-reject {{ background: #fadbd8; color: #922b21; }}
.btn-reject.active, .btn-reject:hover {{ background: #e74c3c; color: #fff; }}
.proposal {{ margin-top: 6px; padding: 6px; background: #fef9e7; border: 1px solid #f9e79f;
             border-radius: 4px; display: flex; gap: 12px; align-items: center; flex-wrap: wrap; font-size: 12px; }}
.proposal img {{ max-width: 80px; max-height: 60px; object-fit: contain; border: 1px solid #ddd;
                border-radius: 2px; background: #fafafa; }}
.proposal-actions {{ display: flex; gap: 6px; flex-wrap: wrap; }}
.proposal-actions button {{ font-size: 10px; padding: 3px 8px; border: 1px solid #ccc; border-radius: 3px;
                            cursor: pointer; background: #fff; }}
.proposal-actions button:hover {{ background: #eee; }}
.proposal-actions .btn-approve {{ border-color: #27ae60; color: #27ae60; }}
.proposal-actions .btn-approve.active {{ background: #27ae60; color: #fff; }}
.proposal-actions .btn-review {{ border-color: #e67e22; color: #e67e22; }}
.proposal-actions .btn-review.active {{ background: #e67e22; color: #fff; }}
.proposal-actions .btn-clear {{ border-color: #e74c3c; color: #e74c3c; }}
.comment-box {{ margin-top: 4px; }}
.comment-box textarea {{ width: 100%; min-height: 32px; font-size: 11px; padding: 3px 6px;
                         border: 1px solid #ddd; border-radius: 3px; resize: vertical; font-family: inherit; }}
.negative-panel {{ background: #fff; border: 2px solid #e74c3c; border-radius: 5px; margin: 10px 16px;
                   padding: 12px; }}
.negative-panel h3 {{ color: #e74c3c; font-size: 14px; margin-bottom: 6px; }}
.negative-desc {{ font-size: 11px; color: #7f8c8d; margin-bottom: 8px; }}
.negative-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(140px, 1fr)); gap: 4px; }}
.negative-card {{ border: 1px solid #f5c6cb; border-radius: 3px; padding: 3px; background: #fff5f5; }}
.negative-card img {{ width: 100%; height: 80px; object-fit: contain; background: #f0f0f0; }}
.neg-label {{ font-size: 9px; font-weight: 600; color: #c0392b; margin-top: 1px; }}
.neg-desc {{ font-size: 8px; color: #7f8c8d; }}
.footer {{ text-align: center; padding: 12px; font-size: 10px; color: #95a5a6; }}
</style>
</head>
<body>
<div class="header">
    <h1>ASD Batch {batch_num} / 8 — Manual Photo Review v5.7</h1>
    <div class="sub">Models {batch['start_idx']}–{batch['end_idx']} ({len(batch['models'])} models) | '
        + 'Open review.html from folder — no server needed | '
        + 'Export decisions → Stage B</div>
</div>
<div class="controls">
    <input type="text" id="search" placeholder="Search model..." oninput="filterModels()">
    <select id="filterStatus" onchange="filterModels()">
        <option value="all">All</option>
        <option value="approved">Approved</option>
        <option value="needs_review">Needs review</option>
        <option value="unreviewed">Unreviewed</option>
        <option value="front_ok">Front selected</option>
        <option value="inside_ok">Inside selected</option>
        <option value="both_ok">Both selected</option>
    </select>
    <span id="statsDisplay" style="font-size:11px;color:#7f8c8d;"></span>
    <button onclick="exportCSV()" class="btn-export">Export CSV</button>
    <button onclick="exportJSON()" class="btn-export">Export JSON</button>
    <button onclick="importJSON()" class="btn-import">Import JSON</button>
    <button onclick="resetBatch()" class="btn-reset">Reset batch</button>
    <button onclick="showNegative()">Не выбирать (примеры)</button>
</div>
<div id="negativePanel" style="display:none;">{negative_section}</div>
<div id="modelsContainer"></div>
<div class="footer">
    Generated {datetime.now().strftime('%Y-%m-%d %H:%M')} | Stage A — review only, no changes applied
</div>
<input type="file" id="importFileInput" accept=".json" style="display:none" onchange="handleImport(event)">

<script>
var BATCH_ID = {batch_num};
var STORAGE_KEY = 'asd_manual_review_v5_7_batch_{batch_num:02d}';
var MODELS = {models_json};

// ===== FIX #1: State architecture — no auto-approval =====
var state = {{}};
var dirty = false;

function initState() {{
    MODELS.forEach(function(m, i) {{
        state[i] = {{
            current_front_url: m.current_front || '',
            current_inside_url: m.current_inside || '',
            approved_front_url: '',
            approved_inside_url: '',
            front_explicitly_approved: false,
            inside_explicitly_approved: false,
            pair_explicitly_approved: false,
            needs_review: false,
            rejected_urls: [],
            comment: ''
        }};
    }});
}}

function saveState() {{
    try {{
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
        dirty = false;
    }} catch(e) {{}}
}}

function loadState() {{
    try {{
        var saved = localStorage.getItem(STORAGE_KEY);
        if (saved) {{
            var parsed = JSON.parse(saved);
            Object.keys(parsed).forEach(function(k) {{ state[k] = parsed[k]; }});
            return true;
        }}
    }} catch(e) {{}}
    return false;
}}

// ===== RENDER =====
function render() {{
    var html = '';
    MODELS.forEach(function(m, i) {{
        var s = state[i];
        var statusClass = 'status-empty';
        var statusLabel = 'Not reviewed';
        if (s.pair_explicitly_approved) {{ statusClass = 'status-ok'; statusLabel = 'Pair approved'; }}
        else if (s.needs_review) {{ statusClass = 'status-review'; statusLabel = 'Needs review'; }}
        else if (s.front_explicitly_approved && s.inside_explicitly_approved) {{ statusClass = 'status-partial'; statusLabel = 'Partially approved'; }}
        else if (s.front_explicitly_approved) {{ statusClass = 'status-partial'; statusLabel = 'Front only'; }}
        else if (s.inside_explicitly_approved) {{ statusClass = 'status-partial'; statusLabel = 'Inside only'; }}
        
        html += '<div class="model-card" data-idx="' + i + '" data-status="' + statusLabel.toLowerCase().replace(/ /g,'_') + '">';
        html += '<div class="model-header">';
        html += '<div><div class="model-title">#' + m.idx + ' ' + esc(m.model_name) + '</div>';
        html += '<div class="model-meta">Row ' + m.excel_row + ' | ' + esc(m.target_line) + ' | <a href="' + esc(m.source_url) + '" target="_blank">source</a></div></div>';
        html += '<div><span class="model-status ' + statusClass + '">' + statusLabel + '</span></div></div>';
        
        // Current pair — NOT APPROVED by default
        html += '<div class="current-pair">';
        if (m.current_front) {{
            html += '<div class="pair-item"><img src="' + esc(m.current_front) + '" onerror="this.style.display=\'none\'">';
            html += '<div class="pair-label not-approved">CURRENT FRONT — NOT APPROVED</div>';
            html += '<button class="approve-current" onclick="approveCurrentAsFront(' + i + ')">Approve current as FRONT</button></div>';
        }}
        if (m.current_inside) {{
            html += '<div class="pair-item"><img src="' + esc(m.current_inside) + '" onerror="this.style.display=\'none\'">';
            html += '<div class="pair-label not-approved">CURRENT INSIDE — NOT APPROVED</div>';
            html += '<button class="approve-current" onclick="approveCurrentAsInside(' + i + ')">Approve current as INSIDE</button></div>';
        }}
        html += '</div>';
        
        if (m.qa_front || m.qa_inside) {{
            html += '<div style="font-size:10px;color:#7f8c8d;margin-bottom:4px;">v5.2 QA: front=' + esc(m.qa_front||'none') + ', inside=' + esc(m.qa_inside||'none') + '</div>';
        }}
        
        // Candidates
        html += '<div class="candidates-section"><h4>Candidates (' + m.candidates.length + '):</h4>';
        html += '<div class="candidates-grid">';
        m.candidates.forEach(function(c, ci) {{
            var isRejected = s.rejected_urls.indexOf(c.url) >= 0;
            var isFront = s.approved_front_url === c.url && s.front_explicitly_approved;
            var isInside = s.approved_inside_url === c.url && s.inside_explicitly_approved;
            var cls = 'candidate';
            if (isFront) cls += ' selected-as-front';
            if (isInside) cls += ' selected-as-inside';
            if (isRejected) cls += ' rejected';
            if (!c.local) cls += ' remote-missing';
            
            html += '<div class="' + cls + '" data-url="' + esc(c.url) + '">';
            if (c.local) {{
                html += '<img src="' + esc(c.thumb || c.local) + '" class="candidate-img" loading="lazy">';
            }} else {{
                html += '<div class="remote-placeholder">';
                html += '<div class="remote-title">REMOTE IMAGE NOT DOWNLOADED</div>';
                html += '<div>' + esc(c.source_type) + '</div>';
                html += '<a href="' + esc(c.url) + '" target="_blank">Open remote URL</a>';
                html += '</div>';
            }}
            html += '<div class="candidate-info">';
            html += '<div class="src-type" title="' + esc(c.url) + '">' + esc(c.source_type).substring(0,35) + '</div>';
            var dims = (c.width && c.height) ? c.width + 'x' + c.height : '';
            html += '<div>Idx #' + c.gallery_index + (dims ? ' | ' + dims : '') + '</div>';
            html += '</div><div class="candidate-actions">';
            html += '<button class="btn-front' + (isFront ? ' active' : '') + '" onclick="selectFront(' + i + ',' + ci + ')">Front</button>';
            html += '<button class="btn-inside' + (isInside ? ' active' : '') + '" onclick="selectInside(' + i + ',' + ci + ')">Inside</button>';
            html += '<button class="btn-reject' + (isRejected ? ' active' : '') + '" onclick="toggleReject(' + i + ',' + ci + ')">Reject</button>';
            html += '</div></div>';
        }});
        html += '</div></div>';
        
        // Proposal — shows explicitly approved URLs only
        html += '<div class="proposal">';
        html += '<div><strong>Approved:</strong></div>';
        html += '<div class="pair-item"><div class="pair-label">Front' + (s.front_explicitly_approved ? '' : ' (none)') + '</div>';
        if (s.front_explicitly_approved) {{
            html += '<img src="' + esc(s.approved_front_url) + '" onerror="this.style.display=\\'none\\'"><br>';
            html += '<button class="btn-clear" onclick="clearFront(' + i + ')">Clear Front</button>';
        }} else {{
            html += '<span style="color:#999">—</span>';
        }}
        html += '</div>';
        html += '<div class="pair-item"><div class="pair-label">Inside' + (s.inside_explicitly_approved ? '' : ' (none)') + '</div>';
        if (s.inside_explicitly_approved) {{
            html += '<img src="' + esc(s.approved_inside_url) + '" onerror="this.style.display=\\'none\\'"><br>';
            html += '<button class="btn-clear" onclick="clearInside(' + i + ')">Clear Inside</button>';
        }} else {{
            html += '<span style="color:#999">—</span>';
        }}
        html += '</div>';
        html += '<div class="proposal-actions">';
        html += '<button class="btn-approve' + (s.pair_explicitly_approved ? ' active' : '') + '" onclick="toggleApprove(' + i + ')">Approve pair</button>';
        html += '<button class="btn-review' + (s.needs_review ? ' active' : '') + '" onclick="toggleReview(' + i + ')">Needs review</button>';
        html += '</div></div>';
        
        html += '<div class="comment-box"><textarea placeholder="Reviewer comment..." onchange="updateComment(' + i + ',this.value)" id="comment' + i + '">' + esc(s.comment) + '</textarea></div>';
        html += '</div>';
    }});
    document.getElementById('modelsContainer').innerHTML = html;
    updateStats();
}}

function esc(s) {{ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }}

// FIX #1: Explicit approval of current URLs
function approveCurrentAsFront(i) {{
    var s = state[i];
    var m = MODELS[i];
    if (!m.current_front) return;
    if (s.approved_inside_url === m.current_front) {{
        alert('This URL is already approved as INSIDE. Clear inside first.');
        return;
    }}
    s.approved_front_url = m.current_front;
    s.front_explicitly_approved = true;
    dirty = true;
    saveState();
    render();
}}

function approveCurrentAsInside(i) {{
    var s = state[i];
    var m = MODELS[i];
    if (!m.current_inside) return;
    if (s.approved_front_url === m.current_inside) {{
        alert('This URL is already approved as FRONT. Clear front first.');
        return;
    }}
    s.approved_inside_url = m.current_inside;
    s.inside_explicitly_approved = true;
    dirty = true;
    saveState();
    render();
}}

function selectFront(modelIdx, candIdx) {{
    var m = MODELS[modelIdx];
    var c = m.candidates[candIdx];
    var s = state[modelIdx];
    if (!c.local) {{ alert('Cannot approve: this is a remote-only candidate. Download the image first.'); return; }}
    if (s.inside_explicitly_approved && s.approved_inside_url === c.url) {{
        alert('This image is already approved as INSIDE. Clear inside first.'); return;
    }}
    s.approved_front_url = c.url;
    s.front_explicitly_approved = true;
    dirty = true;
    saveState();
    render();
}}

function selectInside(modelIdx, candIdx) {{
    var m = MODELS[modelIdx];
    var c = m.candidates[candIdx];
    var s = state[modelIdx];
    if (!c.local) {{ alert('Cannot approve: this is a remote-only candidate. Download the image first.'); return; }}
    if (s.front_explicitly_approved && s.approved_front_url === c.url) {{
        alert('This image is already approved as FRONT. Clear front first.'); return;
    }}
    s.approved_inside_url = c.url;
    s.inside_explicitly_approved = true;
    dirty = true;
    saveState();
    render();
}}

function toggleReject(modelIdx, candIdx) {{
    var m = MODELS[modelIdx];
    var c = m.candidates[candIdx];
    var s = state[modelIdx];
    var idx = s.rejected_urls.indexOf(c.url);
    if (idx >= 0) {{
        s.rejected_urls.splice(idx, 1);
    }} else {{
        s.rejected_urls.push(c.url);
        if (s.front_explicitly_approved && s.approved_front_url === c.url) {{
            s.approved_front_url = ''; s.front_explicitly_approved = false;
        }}
        if (s.inside_explicitly_approved && s.approved_inside_url === c.url) {{
            s.approved_inside_url = ''; s.inside_explicitly_approved = false;
        }}
    }}
    dirty = true;
    saveState();
    render();
}}

function clearFront(modelIdx) {{
    state[modelIdx].approved_front_url = '';
    state[modelIdx].front_explicitly_approved = false;
    if (!state[modelIdx].inside_explicitly_approved) state[modelIdx].pair_explicitly_approved = false;
    dirty = true;
    saveState();
    render();
}}

function clearInside(modelIdx) {{
    state[modelIdx].approved_inside_url = '';
    state[modelIdx].inside_explicitly_approved = false;
    if (!state[modelIdx].front_explicitly_approved) state[modelIdx].pair_explicitly_approved = false;
    dirty = true;
    saveState();
    render();
}}

function toggleApprove(modelIdx) {{
    var s = state[modelIdx];
    if (!s.front_explicitly_approved || !s.inside_explicitly_approved) {{
        alert('Cannot approve: both front and inside must be explicitly approved first.'); return;
    }}
    if (s.approved_front_url === s.approved_inside_url) {{
        alert('Cannot approve: front and inside URLs are identical.'); return;
    }}
    s.pair_explicitly_approved = !s.pair_explicitly_approved;
    if (s.pair_explicitly_approved) s.needs_review = false;
    dirty = true;
    saveState();
    render();
}}

function toggleReview(modelIdx) {{
    var s = state[modelIdx];
    s.needs_review = !s.needs_review;
    if (s.needs_review) s.pair_explicitly_approved = false;
    dirty = true;
    saveState();
    render();
}}

function updateComment(modelIdx, val) {{
    state[modelIdx].comment = val;
    dirty = true;
    saveState();
}}

function updateStats() {{
    var total = MODELS.length;
    var approved = 0, review = 0, unreviewed = 0, partial = 0;
    for (var i = 0; i < MODELS.length; i++) {{
        var s = state[i];
        if (s.pair_explicitly_approved) approved++;
        else if (s.needs_review) review++;
        else if (s.front_explicitly_approved || s.inside_explicitly_approved) partial++;
        else unreviewed++;
    }}
    document.getElementById('statsDisplay').textContent =
        'Total: ' + total + ' | Approved: ' + approved + ' | Partial: ' + partial +
        ' | Review: ' + review + ' | Unreviewed: ' + unreviewed;
}}

function filterModels() {{
    var q = document.getElementById('search').value.toLowerCase();
    var statusFilter = document.getElementById('filterStatus').value;
    var cards = document.querySelectorAll('.model-card');
    cards.forEach(function(card) {{
        var idx = parseInt(card.dataset.idx);
        var s = state[idx];
        var name = MODELS[idx].model_name.toLowerCase();
        var show = name.indexOf(q) >= 0;
        if (show && statusFilter !== 'all') {{
            if (statusFilter === 'approved') show = s.pair_explicitly_approved;
            else if (statusFilter === 'needs_review') show = s.needs_review;
            else if (statusFilter === 'unreviewed') show = !s.pair_explicitly_approved && !s.needs_review && !s.front_explicitly_approved && !s.inside_explicitly_approved;
            else if (statusFilter === 'front_ok') show = s.front_explicitly_approved;
            else if (statusFilter === 'inside_ok') show = s.inside_explicitly_approved;
            else if (statusFilter === 'both_ok') show = s.front_explicitly_approved && s.inside_explicitly_approved;
        }}
        card.style.display = show ? '' : 'none';
    }});
}}

function showNegative() {{
    var panel = document.getElementById('negativePanel');
    panel.style.display = panel.style.display === 'none' ? 'block' : 'none';
}}

// ===== FIX #4: Autosave + Import/Export/Reset =====
function getDecisions() {{
    var result = [];
    MODELS.forEach(function(m, i) {{
        var s = state[i];
        result.push({{
            batch_id: BATCH_ID,
            list_index: m.idx,
            excel_row: m.excel_row,
            producer: m.producer,
            target_line: m.target_line,
            model_name: m.model_name,
            source_url: m.source_url,
            approved_front_url: s.front_explicitly_approved ? s.approved_front_url : '',
            approved_inside_url: s.inside_explicitly_approved ? s.approved_inside_url : '',
            front_approved: s.front_explicitly_approved ? 1 : 0,
            inside_approved: s.inside_explicitly_approved ? 1 : 0,
            pair_approved: s.pair_explicitly_approved ? 1 : 0,
            needs_review: s.needs_review ? 1 : 0,
            rejected_urls: s.rejected_urls.join('; '),
            reviewer_comment: s.comment || ''
        }});
    }});
    return result;
}}

function exportJSON() {{
    var data = getDecisions();
    var blob = new Blob([JSON.stringify(data, null, 2)], {{type: 'application/json'}});
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = 'manual_decisions_batch_' + String(BATCH_ID).padStart(2,'0') + '.json';
    a.click(); URL.revokeObjectURL(url);
    dirty = false;
}}

function exportCSV() {{
    var data = getDecisions();
    var headers = ['batch_id','list_index','excel_row','producer','target_line','model_name','source_url',
                   'approved_front_url','approved_inside_url','front_approved','inside_approved',
                   'pair_approved','needs_review','rejected_urls','reviewer_comment'];
    var csv = headers.join(';') + '\\n';
    data.forEach(function(r) {{
        var row = headers.map(function(h) {{
            var v = (r[h] || '').toString();
            return v.replace(/;/g, ',').replace(/\\n/g, ' ');
        }});
        csv += row.join(';') + '\\n';
    }});
    var blob = new Blob(['\\ufeff' + csv], {{type: 'text/csv;charset=utf-8;'}});
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = 'manual_decisions_batch_' + String(BATCH_ID).padStart(2,'0') + '.csv';
    a.click(); URL.revokeObjectURL(url);
    dirty = false;
}}

function importJSON() {{
    document.getElementById('importFileInput').click();
}}

function handleImport(event) {{
    var file = event.target.files[0];
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function(e) {{
        try {{
            var data = JSON.parse(e.target.result);
            if (!Array.isArray(data)) {{ alert('Invalid format: expected array'); return; }}
            data.forEach(function(d) {{
                for (var i = 0; i < MODELS.length; i++) {{
                    if (MODELS[i].model_name === d.model_name) {{
                        var s = state[i];
                        s.approved_front_url = d.approved_front_url || '';
                        s.approved_inside_url = d.approved_inside_url || '';
                        s.front_explicitly_approved = !!d.front_approved;
                        s.inside_explicitly_approved = !!d.inside_approved;
                        s.pair_explicitly_approved = !!d.pair_approved;
                        s.needs_review = !!d.needs_review;
                        s.rejected_urls = d.rejected_urls ? d.rejected_urls.split('; ').filter(Boolean) : [];
                        s.comment = d.reviewer_comment || '';
                        break;
                    }}
                }}
            }});
            saveState();
            render();
            alert('Imported ' + data.length + ' decisions.');
        }} catch(err) {{
            alert('Import error: ' + err.message);
        }}
    }};
    reader.readAsText(file);
    event.target.value = '';
}}

function resetBatch() {{
    if (!confirm('Reset all decisions in this batch? This cannot be undone.')) return;
    initState();
    saveState();
    render();
}}

// Warn before close if dirty
window.addEventListener('beforeunload', function(e) {{
    if (dirty) {{ e.preventDefault(); e.returnValue = ''; }}
}});

// Init
initState();
loadState();
render();
</script>
</body>
</html>'''
    return html

# ========================================================================
# Contact sheet builder
# ========================================================================
def build_contact_sheet_html(batch):
    parts = ['<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8">',
             '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
             f'<title>ASD Batch {batch["num"]} Contact Sheet</title>',
             '<style>body{font-family:Arial,sans-serif;margin:10px;font-size:11px}',
             'h1{font-size:16px}.pair-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:6px}',
             '.pair-card{border:1px solid #ddd;border-radius:4px;padding:6px;background:#fafafa}',
             '.pair-card h4{font-size:11px;margin:0 0 3px}.pair-imgs{display:flex;gap:4px;flex-wrap:wrap}',
             '.pair-imgs img{width:110px;height:85px;object-fit:contain;border:1px solid #ddd;border-radius:2px;background:#f0f0f0}',
             '.img-col{display:flex;flex-direction:column;align-items:center}',
             '.img-label{font-size:9px;color:#7f8c8d}.tag{display:inline-block;font-size:8px;padding:1px 4px;border-radius:2px;margin-top:2px}',
             '.tag-ok{background:#d5f5e3;color:#1e8449}.tag-review{background:#fdebd0;color:#a04000}.tag-empty{background:#fadbd8;color:#922b21}',
             '</style></head><body>',
             f'<h1>ASD Batch {batch["num"]} Contact Sheet</h1>',
             f'<p>Models {batch["start_idx"]}–{batch["end_idx"]}</p>',
             '<div class="pair-grid">']
    
    for m in batch['models']:
        name = m['model_name']
        front_url = m.get('current_front', '')
        inside_url = m.get('current_inside', '')
        has_front = bool(front_url)
        has_inside = bool(inside_url)
        status = 'tag-empty'
        label = 'Empty'
        if has_front and has_inside:
            status = 'tag-ok' if front_url != inside_url else 'tag-review'
            label = 'OK' if front_url != inside_url else 'Same URL'
        elif has_front:
            status = 'tag-review'; label = 'No inside'
        elif has_inside:
            status = 'tag-review'; label = 'No front'
        
        front_html = f'<img src="{html_mod.escape(front_url)}" onerror="this.style.display=\'none\'">' if front_url else '<span style="color:#999">—</span>'
        inside_html = f'<img src="{html_mod.escape(inside_url)}" onerror="this.style.display=\'none\'">' if inside_url else '<span style="color:#999">—</span>'
        
        parts.append(f'<div class="pair-card"><h4 title="{html_mod.escape(name)}">#{m["list_index"]} {html_mod.escape(name)[:55]}</h4>'
                     f'<div class="pair-imgs"><div class="img-col">{front_html}<span class="img-label">Front</span></div>'
                     f'<div class="img-col">{inside_html}<span class="img-label">Inside</span></div></div>'
                     f'<span class="tag {status}">{label}</span></div>')
    
    parts.append('</div></body></html>')
    return '\n'.join(parts)

# ========================================================================
# FIX #5: Decision template with excel_row filled
# ========================================================================
def build_decisions_template(batch):
    decisions = []
    for m in batch['models']:
        name = m['model_name']
        row = asd_map.get(name, {})
        excel_row = ''
        for r_idx, r in enumerate(rows):
            if r['model_name'] == name:
                excel_row = r_idx + 2
                break
        decisions.append({
            'batch_id': batch['num'],
            'list_index': m['list_index'],
            'excel_row': excel_row,
            'producer': 'ASD',
            'target_line': row.get('target_line', ''),
            'model_name': name,
            'source_url': row.get('source_url', ''),
            'approved_front_url': '',
            'approved_inside_url': '',
            'front_approved': 0,
            'inside_approved': 0,
            'pair_approved': 0,
            'needs_review': 0,
            'rejected_urls': '',
            'reviewer_comment': '',
        })
    return decisions

# ========================================================================
# Build all batches
# ========================================================================
print("\n=== Building portable batches ===")
negative_section = build_negative_section()

for batch in batches:
    batch_dir = batch['dir']
    images_dir = os.path.join(batch_dir, 'images')
    os.makedirs(images_dir, exist_ok=True)
    
    copied_count = 0
    for m in batch['models']:
        for cand in m.get('candidates', []):
            url = cand['url']
            dl = dl_by_url.get(url, {})
            dest = dl.get('destination', '')
            if dest and os.path.exists(dest):
                fname = safe_filename(url) + '_' + short_id(url) + os.path.splitext(dest)[1]
                batch_img = os.path.join(images_dir, fname)
                if not os.path.exists(batch_img):
                    try:
                        shutil.copy2(dest, batch_img)
                        copied_count += 1
                    except Exception:
                        pass
    
    neg_dir_batch = os.path.join(batch_dir, 'negative_examples')
    if not os.path.exists(neg_dir_batch):
        try:
            shutil.copytree(NEGATIVE_DIR, neg_dir_batch)
        except Exception:
            pass
    
    # Generate thumbnails
    thumb_map = generate_batch_thumbnails(batch_dir)
    
    # Generate review.html
    review_html = build_review_html(batch, negative_section, thumb_map)
    with open(os.path.join(batch_dir, 'review.html'), 'w', encoding='utf-8') as f:
        f.write(review_html)
    
    # Generate contact_sheet.html
    contact_html = build_contact_sheet_html(batch)
    with open(os.path.join(batch_dir, 'contact_sheet.html'), 'w', encoding='utf-8') as f:
        f.write(contact_html)
    
    # Generate decisions template
    decisions = build_decisions_template(batch)
    csv_path = os.path.join(batch_dir, f'manual_decisions_batch_{batch["num"]:02d}.csv')
    with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=list(decisions[0].keys()), delimiter=';')
        writer.writeheader()
        writer.writerows(decisions)
    json_path = os.path.join(batch_dir, f'manual_decisions_batch_{batch["num"]:02d}.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(decisions, f, ensure_ascii=False, indent=2)
    
    # Generate models.json
    models_data = []
    for m in batch['models']:
        row = asd_map.get(m['model_name'], {})
        qa = qa_v52.get(m['model_name'], {})
        models_data.append({
            'list_index': m['list_index'],
            'producer': 'ASD',
            'target_line': row.get('target_line', ''),
            'model_name': m['model_name'],
            'source_url': row.get('source_url', ''),
            'current_front': m.get('current_front', ''),
            'current_inside': m.get('current_inside', ''),
            'comment': row.get('comment', ''),
            'qa_front_decision': qa.get('front_visual_decision', ''),
            'qa_inside_decision': qa.get('inside_visual_decision', ''),
            'candidate_count': len(m.get('candidates', [])),
        })
    with open(os.path.join(batch_dir, 'models.json'), 'w', encoding='utf-8') as f:
        json.dump(models_data, f, ensure_ascii=False, indent=2)
    
    img_count = len([f for f in os.listdir(images_dir) if os.path.isfile(os.path.join(images_dir, f))])
    thumb_count = len([f for f in os.listdir(os.path.join(batch_dir, 'thumbs')) if os.path.isfile(os.path.join(batch_dir, 'thumbs', f))]) if os.path.isdir(os.path.join(batch_dir, 'thumbs')) else 0
    print(f"  {batch['label']}: {len(batch['models'])} models, {img_count} images, {thumb_count} thumbs")

# ========================================================================
# FIX #6 & #7: Build Lord portable review with correct images
# ========================================================================
print("\n=== Building Lord portable review ===")
LORD = chr(0x041B) + chr(0x043E) + chr(0x0440) + chr(0x0434)
lord_rows = [r for r in rows if r['producer'] == LORD]

os.makedirs(LORD_DIR, exist_ok=True)
lord_images_dir = os.path.join(LORD_DIR, 'images')
os.makedirs(lord_images_dir, exist_ok=True)

# Build Lord image mapping: model_name -> local file path
lord_img_map = {}  # model_name -> {col: flat_filename}
lord_cache_files = {}

for root, dirs, files in os.walk(CACHE_LORD):
    for f in files:
        src = os.path.join(root, f)
        rel = os.path.relpath(src, CACHE_LORD)
        flat_name = rel.replace('\\', '_').replace('/', '_')
        dst = os.path.join(lord_images_dir, flat_name)
        if not os.path.exists(dst):
            try:
                shutil.copy2(src, dst)
            except Exception:
                pass
        
        parts = rel.split(os.sep)
        if len(parts) >= 1:
            model_subdir = parts[0]
            col_prefix = 'front' if 'front' in f.lower() else 'inside'
            lord_img_map.setdefault(model_subdir, {})[col_prefix] = flat_name
            lord_cache_files.setdefault(model_subdir, []).append(flat_name)

print(f"  Lord image dirs mapped: {len(lord_img_map)}")

# Build a comprehensive lookup: by model_name, by safe_name, by partial match
lord_row_to_local = {}  # row_index -> local_filename

def find_lord_local_image(row):
    name = row['model_name']
    f_url = row.get('image_url_front', '').strip()
    
    # Try exact model_name match in map
    if name in lord_img_map:
        info = lord_img_map[name]
        if 'front' in info:
            return info['front']
    
    # Try safe_name match
    safe = re.sub(r'[<>:"/\\|?*]', '_', name)[:120]
    if safe in lord_img_map:
        info = lord_img_map[safe]
        if 'front' in info:
            return info['front']
    
    # Try partial match
    for key in lord_img_map:
        if key.startswith(safe[:30]) or safe.startswith(key[:30]):
            info = lord_img_map[key]
            if 'front' in info:
                return info['front']
    
    # Try matching by URL hash
    if f_url:
        url_hash = short_id(f_url)
        for flat_name in os.listdir(lord_images_dir):
            if url_hash in flat_name:
                return flat_name
    
    return ''

# Generate Lord flagged review HTML (FIX #6)
flagged = []
for r_idx, r in enumerate(lord_rows):
    name = r['model_name']
    f_url = r.get('image_url_front', '').strip()
    i_url = r.get('image_url_inside', '').strip()
    comment = r.get('comment', '')
    
    flags = []
    for url, side in [(f_url, 'front'), (i_url, 'inside')]:
        if not url: continue
        url_lower = url.lower()
        if any(kw in url_lower for kw in ['steklo', 'glass', 'vitrin', 'window']):
            flags.append(f'{side}_glass_keyword')
        if any(kw in url_lower for kw in ['fragment', 'detail']):
            flags.append(f'{side}_fragment_keyword')
        if any(kw in url_lower for kw in ['handle', 'lock', 'ruchk', 'zamok']):
            flags.append(f'{side}_handle_keyword')
    
    if flags:
        local_img = find_lord_local_image(r)
        flagged.append({
            'row_index': r_idx,
            'model_name': name,
            'target_line': r.get('target_line', ''),
            'front_url': f_url,
            'inside_url': i_url,
            'flags': '; '.join(flags),
            'comment': comment,
            'local_image': local_img,
        })

print(f"  Lord flagged rows: {len(flagged)}")

# Build flagged review HTML
lord_html_parts = ['<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8">',
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
    '<title>Lord Flagged Review</title>',
    '<style>body{font-family:Arial,sans-serif;margin:10px;font-size:12px}',
    'h1{font-size:16px}table{border-collapse:collapse;width:100%}',
    'th,td{border:1px solid #ccc;padding:4px 6px;text-align:left}',
    'th{background:#2c3e50;color:#fff;position:sticky;top:0}',
    'tr.flagged{background:#fff5f5}',
    'img{max-width:100px;max-height:85px;object-fit:contain;border:1px solid #ddd;border-radius:2px;background:#f0f0f0}',
    '.flag{display:inline-block;font-size:9px;padding:1px 4px;border-radius:2px;margin:1px}',
    '.flag-glass{background:#cce5ff;color:#004085}',
    '.flag-fragment{background:#f8d7da;color:#721c24}',
    '.flag-handle{background:#fff3cd;color:#856404}',
    '.missing{background:#f8d7da;padding:4px;font-size:10px;color:#721c24;border:1px dashed #e74c3c}',
    '</style></head><body>',
    f'<h1>Lord Flagged Review — portable batch</h1>',
    f'<p>{len(flagged)} flagged rows out of {len(lord_rows)} total</p>',
    '<table><tr><th>#</th><th>Model</th><th>Line</th><th>Front</th><th>Flags</th><th>Action</th></tr>']

for idx, f in enumerate(flagged):
    flag_class = 'glass' if 'glass' in f.get('flags', '') else ('fragment' if 'fragment' in f.get('flags', '') else 'handle')
    if f['local_image']:
        img_html = f'<img src="images/{html_mod.escape(f["local_image"])}">'
    else:
        img_html = f'<div class="missing">LOCAL IMAGE MISSING<br><a href="{html_mod.escape(f["front_url"])}" target="_blank" style="font-size:9px">Open original</a></div>'
    
    lord_html_parts.append(f'<tr class="flagged"><td>{idx+1}</td>'
        f'<td title="{html_mod.escape(f["model_name"])}">{html_mod.escape(f["model_name"][:45])}</td>'
        f'<td>{html_mod.escape(f.get("target_line","")[:25])}</td>'
        f'<td>{img_html}</td>'
        f'<td><span class="flag flag-{flag_class}">{html_mod.escape(f.get("flags","")[:50])}</span></td>'
        f'<td><button onclick="this.style.background=\'#27ae60\';this.style.color=\'#fff\';this.textContent=\'Keep\'">Keep</button> '
        f'<button onclick="this.style.background=\'#e74c3c\';this.style.color=\'#fff\';this.textContent=\'Reject\'">Reject</button> '
        f'<button onclick="this.style.background=\'#e67e22\';this.style.color=\'#fff\';this.textContent=\'Review\'">Review</button></td></tr>')

lord_html_parts.append('</table></body></html>')
with open(os.path.join(LORD_DIR, 'lord_flagged_review.html'), 'w', encoding='utf-8') as f:
    f.write(''.join(lord_html_parts))

# Generate Lord all contact sheet (FIX #7: proper image mapping)
lord_all_parts = ['<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8">',
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
    '<title>Lord All Images Contact Sheet</title>',
    '<style>body{font-family:Arial,sans-serif;margin:10px;font-size:10px}',
    'h1{font-size:14px}.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:4px}',
    '.card{border:1px solid #ddd;border-radius:3px;padding:3px;background:#fafafa}',
    '.card img{width:100%;height:90px;object-fit:contain;border:1px solid #eee;border-radius:2px;background:#f0f0f0}',
    '.card .name{font-size:9px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}',
    '.page{page-break-after:always}.actions{display:flex;gap:2px;margin-top:2px}',
    '.actions button{font-size:8px;padding:1px 4px;border:1px solid #ccc;border-radius:2px;cursor:pointer}',
    '.no-image{background:#f8d7da;height:90px;display:flex;align-items:center;justify-content:center;font-size:9px;color:#721c24;border:1px dashed #e74c3c}',
    '</style></head><body>',
    f'<h1>Lord All Images Contact Sheet — {len(lord_rows)} models</h1>']

page_size = 50
for page_start in range(0, len(lord_rows), page_size):
    page_rows = lord_rows[page_start:page_start + page_size]
    page_num = page_start // page_size + 1
    
    lord_all_parts.append(f'<div class="page"><h2>Page {page_num} ({page_start+1}–{page_start+len(page_rows)})</h2><div class="grid">')
    
    for r in page_rows:
        name = r['model_name']
        local_img = find_lord_local_image(r)
        
        if local_img:
            img_html = f'<img src="images/{html_mod.escape(local_img)}" loading="lazy">'
        else:
            img_html = '<div class="no-image">NO LOCAL IMAGE</div>'
        
        lord_all_parts.append(f'<div class="card">{img_html}'
            f'<div class="name" title="{html_mod.escape(name)}">{html_mod.escape(name[:35])}</div>'
            f'<div class="actions"><button onclick="this.style.background=\'#27ae60\';this.textContent=\'Keep\'">Keep</button>'
            f'<button onclick="this.style.background=\'#e74c3c\';this.textContent=\'Reject\'">Reject</button></div></div>')
    
    lord_all_parts.append('</div></div>')

lord_all_parts.append('</body></html>')
with open(os.path.join(LORD_DIR, 'lord_all_contact_sheet.html'), 'w', encoding='utf-8') as f:
    f.write(''.join(lord_all_parts))

# Lord decisions template (FIX #8)
lord_decisions = []
for r_idx, r in enumerate(lord_rows):
    local_img = find_lord_local_image(r)
    lord_decisions.append({
        'excel_row': r_idx + 2,
        'producer': LORD,
        'target_line': r.get('target_line', ''),
        'model_name': r['model_name'],
        'source_url': r.get('source_url', ''),
        'current_image_url': r.get('image_url_front', ''),
        'local_image': local_img,
        'decision': '',
        'approved_replacement_url': '',
        'reviewer_comment': '',
        'explicitly_reviewed': 0,
    })

lord_csv_path = os.path.join(LORD_DIR, 'lord_decisions_template.csv')
with open(lord_csv_path, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=list(lord_decisions[0].keys()), delimiter=';')
    writer.writeheader()
    writer.writerows(lord_decisions)
lord_json_path = os.path.join(LORD_DIR, 'lord_decisions_template.json')
with open(lord_json_path, 'w', encoding='utf-8') as f:
    json.dump(lord_decisions, f, ensure_ascii=False, indent=2)

print(f"  Lord decisions template: {len(lord_decisions)} rows")

# ========================================================================
# FIX #10: README
# ========================================================================
readme = f'''# Manual Photo Review v5.7 — Portable Batches

Всего сформировано {len(batches)} ASD batches + 1 Lord review.

## Как пользоваться

1. **Распаковать архив** batch_NN.zip — все ссылки относительные, работает без сервера.

2. **Открыть review.html** в папке batch_NN/ в браузере (Chrome/Edge/Firefox).
   Файл полностью автономен — не нужен Python, localhost или интернет.

3. **Для каждой модели АСД:**
   - Просмотреть все кандидаты (сетка изображений).
   - Нажать «Front» под подходящим изображением.
   - Нажать «Inside» под подходящим изображением.
   - Если изображение не подходит — нажать «Reject».
   - Для текущих фото: нажать «Approve current as FRONT/INSIDE».
   - «Clear Front» / «Clear Inside» — сбросить выбор.
   - «Approve pair» — утвердить пару.
   - «Needs review» — если нужна дополнительная проверка.
   - Комментарий в текстовом поле.

4. **Автосохранение** — решения сохраняются в localStorage автоматически.
   При повторном открытии решения восстанавливаются.

5. **Импорт/Экспорт:**
   - «Export JSON» / «Export CSV» — скачать решения.
   - «Import JSON» — загрузить ранее сохранённые решения.
   - «Reset batch» — сбросить все решения в batch.

6. **Запрещено** выбирать:
   - Видео-миниатюры, баннеры/коллажи, две двери на одном фото
   - Широкие интерьерные сцены, фрагменты/стекло крупным планом
   - Ручки/замки/петли, технические фотографии

7. **Объединить решения:**
   python scripts/stage_b/merge_manual_decisions.py

## Структура архивов

```
asd_batch_01.zip ... asd_batch_08.zip    # Каждый автономен
lord_manual_review.zip                     # Lord flagged + all contact sheet
review_reports_and_scripts.zip             # Отчёты + merge скрипт + QA
```

## Статистика

- ASD моделей: 199
- Batches: {len(batches)}
- Локальных изображений ASD: ~700
- Lord строк: {len(lord_rows)}
- Lord flagged: {len(flagged)}
'''
with open(os.path.join(OUTPUT_BASE, 'README.md'), 'w', encoding='utf-8') as f:
    f.write(readme)

# ========================================================================
# FIX #10 & #13: Build per-batch ZIPs + Lord ZIP + reports ZIP
# ========================================================================
print("\n=== Building ZIP archives ===")

zip_output = os.path.join(OUTPUT_BASE, 'archives')
os.makedirs(zip_output, exist_ok=True)

# Per-batch ZIPs
for batch in batches:
    zip_name = os.path.join(zip_output, f'asd_batch_{batch["num"]:02d}.zip')
    batch_dir = batch['dir']
    with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(batch_dir):
            for f in files:
                fp = os.path.join(root, f)
                arcname = os.path.relpath(fp, batch_dir)
                zf.write(fp, arcname)
    size_mb = os.path.getsize(zip_name) / 1024 / 1024
    print(f"  {os.path.basename(zip_name)}: {size_mb:.1f} MB")

# Lord ZIP — split into 2 parts to stay under 510 MB
lord_images = sorted([f for f in os.listdir(os.path.join(LORD_DIR, 'images'))
                      if os.path.isfile(os.path.join(LORD_DIR, 'images', f))])
lord_cum = 0
lord_split = 0
for li, lf in enumerate(lord_images):
    lord_cum += os.path.getsize(os.path.join(LORD_DIR, 'images', lf))
    if lord_cum > 320 * 1024 * 1024:
        lord_split = li
        break

for part_num, (start, end) in enumerate([(0, lord_split), (lord_split, len(lord_images))], 1):
    lord_zip = os.path.join(zip_output, f'lord_manual_review_part{part_num}.zip')
    part_images = set(lord_images[start:end])
    with zipfile.ZipFile(lord_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
        # Non-image files (HTML, CSV, JSON)
        for root, dirs, files in os.walk(LORD_DIR):
            for f in files:
                fp = os.path.join(root, f)
                rel = os.path.relpath(fp, LORD_DIR)
                if not rel.startswith('images' + os.sep):
                    zf.write(fp, 'lord_portable_review/' + rel.replace('\\', '/'))
        # Only images for this part
        for f in part_images:
            fp = os.path.join(LORD_DIR, 'images', f)
            zf.write(fp, f'lord_portable_review/images/{f}')
    size_mb = os.path.getsize(lord_zip) / 1024 / 1024
    print(f"  {os.path.basename(lord_zip)}: {size_mb:.1f} MB ({len(part_images)} images)")

# Reports + scripts ZIP
reports_zip = os.path.join(zip_output, 'review_reports_and_scripts.zip')
with zipfile.ZipFile(reports_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
    # Reports
    if os.path.isdir(REPORTS_DIR):
        for f in os.listdir(REPORTS_DIR):
            fp = os.path.join(REPORTS_DIR, f)
            if os.path.isfile(fp):
                zf.write(fp, f'reports/{f}')
    # merge script
    merge_script = os.path.join(BASE, 'scripts', 'stage_b', 'merge_manual_decisions.py')
    if os.path.exists(merge_script):
        zf.write(merge_script, 'scripts/stage_b/merge_manual_decisions.py')
    # QA script (will be created next)
    qa_script = os.path.join(BASE, 'scripts', 'stage_a', 'validate_portable_review.py')
    if os.path.exists(qa_script):
        zf.write(qa_script, 'scripts/stage_a/validate_portable_review.py')
    # README
    readme_path = os.path.join(OUTPUT_BASE, 'README.md')
    if os.path.exists(readme_path):
        zf.write(readme_path, 'README.md')
size_mb = os.path.getsize(reports_zip) / 1024 / 1024
print(f"  {os.path.basename(reports_zip)}: {size_mb:.1f} MB")

# ========================================================================
# Build asset manifest for ASD
# ========================================================================
print("\n=== Building asset manifest ===")
manifest_rows = []
for batch in batches:
    for m in batch['models']:
        name = m['model_name']
        for cand in m.get('candidates', []):
            url = cand['url']
            dl = dl_by_url.get(url, {})
            dest = dl.get('destination', '')
            local_path = dest if (dest and os.path.exists(dest)) else ''
            downloaded = bool(local_path)
            sha = file_sha256(local_path) if local_path else ''
            w, h = (dl.get('width', 0), dl.get('height', 0)) if dl else (0, 0)
            manifest_rows.append({
                'batch_id': batch['num'],
                'model_index': m['list_index'],
                'model_name': name,
                'candidate_url': url,
                'local_path': local_path,
                'downloaded': downloaded,
                'download_status': 'ok' if downloaded else 'not_downloaded',
                'width': w,
                'height': h,
                'sha256': sha,
                'visible_in_review': 'yes' if downloaded else 'remote_only',
            })

manifest_path = os.path.join(OUTPUT_BASE, 'reports', 'portable_asset_manifest.csv')
with open(manifest_path, 'w', encoding='utf-8-sig', newline='') as f:
    if manifest_rows:
        writer = csv.DictWriter(f, fieldnames=list(manifest_rows[0].keys()), delimiter=';')
        writer.writeheader()
        writer.writerows(manifest_rows)
print(f"  Asset manifest: {len(manifest_rows)} candidates")

# ========================================================================
# FINAL SUMMARY
# ========================================================================
total_local = sum(len([f for f in os.listdir(os.path.join(b['dir'],'images')) if os.path.isfile(os.path.join(b['dir'],'images',f))]) for b in batches)
total_remote = sum(1 for r in manifest_rows if r['downloaded'] == False)

print(f"\n{'='*60}")
print(f"PORTABLE BATCHES REBUILD COMPLETE")
print(f"{'='*60}")
print(f"  ASD batches:        {len(batches)}")
print(f"  Total ASD models:   {sum(len(b['models']) for b in batches)}")
for b in batches:
    print(f"    {b['label']}: models {b['start_idx']}-{b['end_idx']} ({len(b['models'])} models)")
print(f"  ASD local images:   {total_local}")
print(f"  ASD remote-only:    {total_remote}")
print(f"  Lord rows:          {len(lord_rows)}")
print(f"  Lord flagged:       {len(flagged)}")
print(f"  Lord local images:  {len(os.listdir(lord_images_dir))}")
print(f"")
print(f"  Archives:")
for f in sorted(os.listdir(zip_output)):
    fp = os.path.join(zip_output, f)
    print(f"    {f}: {os.path.getsize(fp)/1024/1024:.1f} MB")
print(f"")
print(f"  Ready for manual review — no catalog changes.")
print(f"  Next: unzip each batch_XX, open review.html, approve, export.")
print(f"  Then: run Stage B: scripts/stage_b/merge_manual_decisions.py")
