"""
Stage A - Step 2: Build comprehensive HTML review page + contact sheet + risky rows.
Loads candidate DB + download results + (optional) VLM output.
Generates:
  - asd_manual_review.html     (full review with manual approve UI)
  - asd_pair_contact_sheet.html (compact 199-pair overview)
  - asd_risky_rows.xlsx         (all flagged rows)
  - asd_vlm_input.json          (VLM batch input file)
"""
import csv, json, os, hashlib, html as html_mod
from collections import defaultdict
from datetime import datetime
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import openpyxl

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CACHE_DIR = os.path.join(BASE, "output", "manual_photo_review_v5_7", "cache", "asd")
REPORTS_DIR = os.path.join(BASE, "output", "manual_photo_review_v5_7", "reports")
V56_DIR = os.path.join(BASE, "output", "delivery_v5_6")

# Load candidate DB
db_path = os.path.join(REPORTS_DIR, "asd_candidate_database.json")
with open(db_path, encoding='utf-8') as f:
    all_models = json.load(f)

# Load download results
dl_path = os.path.join(REPORTS_DIR, "asd_download_results.json")
dl_map = {}
if os.path.exists(dl_path):
    with open(dl_path, encoding='utf-8') as f:
        for d in json.load(f):
            dl_map[d['url']] = d

# Load optional VLM output
vlm_path = os.path.join(REPORTS_DIR, "asd_vlm_output.json")
vlm_map = {}
if os.path.exists(vlm_path):
    with open(vlm_path, encoding='utf-8') as f:
        for v in json.load(f):
            vlm_map[v.get('url', v.get('image_url', ''))] = v

# ========================================================================
# Build VLM input file (for external VLM tool)
# ========================================================================
vlm_input = []
for model in all_models:
    for cand in model['candidates']:
        if cand['url'] not in dl_map:
            continue
        dl = dl_map[cand['url']]
        if not dl.get('downloaded'):
            continue
        vlm_input.append({
            'model_name': model['model_name'],
            'source_url': model['source_url'],
            'excel_row': model['excel_row'],
            'image_url': cand['url'],
            'local_path': dl.get('destination', ''),
            'source_type': cand['source_type'],
            'gallery_index': cand['gallery_index'],
            'original_filename': cand.get('original_filename', ''),
            'sha256': dl.get('sha256', ''),
            'width': dl.get('width'),
            'height': dl.get('height'),
            'file_size': dl.get('file_size', 0),
            'content_type': dl.get('content_type', ''),
        })

vlm_input_path = os.path.join(REPORTS_DIR, "asd_vlm_input.json")
with open(vlm_input_path, 'w', encoding='utf-8') as f:
    json.dump(vlm_input, f, ensure_ascii=False, indent=2)
print(f"VLM input: {vlm_input_path} ({len(vlm_input)} candidates)")

# ========================================================================
# HELPER: Generate HTML for an image thumbnail with VLM overlay
# ========================================================================
def img_thumbnail_html(cand_url, dl, vlm, is_current_front=False, is_current_inside=False, checked=False):
    """Generate HTML for a candidate image in the review grid."""
    local_path = dl.get('destination', '') if dl else ''
    # Compute relative path from HTML location to cache
    if local_path:
        # HTML is at reports_dir/asd_manual_review.html, cache is at cache_dir/...
        rel = os.path.relpath(local_path, REPORTS_DIR)
        thumb_src = rel.replace('\\', '/')
    else:
        thumb_src = cand_url
    
    status_badges = []
    if is_current_front:
        status_badges.append('<span class="badge badge-current-front">Front</span>')
    if is_current_inside:
        status_badges.append('<span class="badge badge-current-inside">Inside</span>')
    
    vlm_class = ''
    vlm_confidence = ''
    vlm_side = ''
    vlm_scene = ''
    vlm_reject = ''
    
    if vlm:
        vlm_confidence = vlm.get('confidence', '')
        vlm_side = vlm.get('side', '')
        vlm_scene = vlm.get('scene_type', '')
        rejected = vlm.get('rejection_reason', '')
        if vlm.get('has_door') == False or vlm.get('full_door_visible') == False or vlm.get('door_is_primary_object') == False:
            rejected = rejected or 'rejected_by_vlm'
        if vlm.get('confidence', 1) < 0.9:
            rejected = (rejected + '; low_confidence' if rejected else 'low_confidence')
        vlm_reject = rejected
        
        if rejected:
            vlm_class = 'vlm-rejected'
        elif vlm_confidence and float(vlm_confidence) >= 0.9:
            vlm_class = 'vlm-accepted'
        else:
            vlm_class = 'vlm-unknown'
    
    # Build data attributes for JS
    data_url = html_mod.escape(cand_url, quote=True)
    short_url = cand_url[:80] + '...' if len(cand_url) > 80 else cand_url
    
    return f'''<div class="candidate {vlm_class}" data-url="{data_url}">
    <div class="candidate-img-wrapper">
        <img src="{html_mod.escape(thumb_src, quote=True)}" 
             onerror="this.onerror=null;this.src='data:image/svg+xml,<svg xmlns=%22http://www.w3.org/2000/svg%22 width=%22120%22 height=%22120%22><rect fill=%22%23ddd%22 width=%22120%22 height=%22120%22/><text x=%225%22 y=%2260%22 font-size=%2210%22>LOAD ERROR</text></svg>'"
             loading="lazy" class="candidate-img" 
             onclick="toggleSelect(this, '{data_url}')">
        <div class="candidate-badges">{" ".join(status_badges)}</div>
    </div>
    <div class="candidate-info">
        <div class="candidate-source" title="{html_mod.escape(short_url)}">{html_mod.escape(cand.get('source_type','unknown')[:40])}</div>
        <div class="candidate-gallery-idx">#{cand.get('gallery_index','?')}</div>
    </div>
    <div class="candidate-vlm">
        <span class="vlm-side">{html_mod.escape(str(vlm_side))}</span>
        <span class="vlm-scene">{html_mod.escape(str(vlm_scene)[:15])}</span>
        <span class="vlm-conf">{html_mod.escape(str(vlm_confidence))}</span>
        {f'<div class="vlm-reject-reason">{html_mod.escape(vlm_reject)}</div>' if vlm_reject else ''}
    </div>
    <div class="candidate-actions">
        <button class="btn-front" onclick="selectFront('{data_url}', this)">← Front</button>
        <button class="btn-inside" onclick="selectInside('{data_url}', this)">Inside →</button>
    </div>
</div>'''

# ========================================================================
# BUILD HTML REVIEW PAGE
# ========================================================================
print("Building HTML review page...")

# Counters
stats = defaultdict(int)

html_parts = []
html_parts.append('''<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ASD Manual Photo Review v5.7 - Stage A</title>
<style>
:root { --bg: #f8f9fa; --card-bg: #fff; --border: #dee2e6; --text: #212529; --text-muted: #6c757d; }
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
       background: var(--bg); color: var(--text); padding: 0; margin: 0; }
.header { background: #1a1a2e; color: #fff; padding: 16px 24px; position: sticky; top: 0; z-index: 1000;
          display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px; }
.header h1 { font-size: 18px; font-weight: 600; }
.header .stats { font-size: 12px; color: #aab; }
.header .actions { display: flex; gap: 8px; flex-wrap: wrap; }
.header button, .header .btn { background: #16213e; color: #fff; border: 1px solid #0f3460;
       padding: 6px 14px; border-radius: 4px; cursor: pointer; font-size: 12px; }
.header button:hover { background: #0f3460; }
.model-card { background: var(--card-bg); border: 1px solid var(--border); border-radius: 6px;
              margin: 12px 16px; padding: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
.model-header { display: flex; justify-content: space-between; align-items: flex-start;
                margin-bottom: 8px; flex-wrap: wrap; gap: 4px; }
.model-title { font-size: 14px; font-weight: 600; flex: 1; }
.model-row { font-size: 11px; color: var(--text-muted); }
.model-status { font-size: 11px; padding: 2px 8px; border-radius: 3px; }
.status-ok { background: #d4edda; color: #155724; }
.status-review { background: #fff3cd; color: #856404; }
.status-empty { background: #f8d7da; color: #721c24; }
.current-pair { display: flex; gap: 12px; margin: 6px 0 8px; flex-wrap: wrap; }
.current-pair img { max-width: 100px; max-height: 80px; border: 2px solid var(--border);
                    border-radius: 3px; object-fit: cover; }
.current-pair .pair-label { font-size: 10px; color: var(--text-muted); text-align: center; }
.current-pair .pair-item { display: flex; flex-direction: column; align-items: center; gap: 2px; }
.gallery-index-label { font-size: 10px; color: var(--text-muted); margin: 4px 0; }
.candidates-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(150px, 1fr));
                   gap: 8px; margin-top: 6px; }
.candidate { border: 1px solid var(--border); border-radius: 4px; padding: 4px;
             background: #fafafa; position: relative; transition: all 0.15s; }
.candidate:hover { border-color: #0f3460; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
.candidate-img-wrapper { position: relative; }
.candidate-img { width: 100%; height: 100px; object-fit: cover; border-radius: 2px;
                 cursor: pointer; border: 2px solid transparent; }
.candidate-img.selected { border-color: #007bff; }
.candidate-badges { position: absolute; top: 2px; left: 2px; display: flex; gap: 2px; flex-wrap: wrap; }
.badge { font-size: 8px; padding: 1px 4px; border-radius: 2px; font-weight: 600; }
.badge-current-front { background: #007bff; color: white; }
.badge-current-inside { background: #28a745; color: white; }
.candidate-info { display: flex; justify-content: space-between; font-size: 9px;
                  color: var(--text-muted); margin-top: 2px; }
.candidate-source { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 100px; }
.candidate-gallery-idx { font-weight: 600; }
.candidate-vlm { font-size: 9px; margin-top: 1px; display: flex; gap: 4px; flex-wrap: wrap; }
.vlm-side { color: #6f42c1; }
.vlm-scene { color: #fd7e14; }
.vlm-conf { color: #28a745; }
.vlm-reject-reason { color: #dc3545; font-size: 8px; width: 100%; }
.vlm-rejected { background: #fff5f5; border-color: #f5c6cb; }
.vlm-accepted { background: #f0fff4; border-color: #c3e6cb; }
.candidate-actions { display: flex; gap: 4px; margin-top: 3px; }
.btn-front, .btn-inside { font-size: 9px; padding: 2px 6px; border: none; border-radius: 3px;
                           cursor: pointer; flex: 1; }
.btn-front { background: #cce5ff; color: #004085; }
.btn-front:hover, .btn-front.active { background: #007bff; color: white; }
.btn-inside { background: #d4edda; color: #155724; }
.btn-inside:hover, .btn-inside.active { background: #28a745; color: white; }
.model-proposal { margin-top: 6px; padding: 6px; background: #e8f4f8; border-radius: 4px;
                  display: flex; gap: 12px; align-items: center; flex-wrap: wrap; font-size: 12px; }
.model-proposal img { max-width: 80px; max-height: 60px; border-radius: 2px; }
.proposal-label { font-size: 10px; font-weight: 600; width: 100%; }
.model-comment { margin-top: 4px; }
.model-comment textarea { width: 100%; min-height: 40px; font-size: 11px; padding: 4px;
                          border: 1px solid var(--border); border-radius: 3px; resize: vertical; }
.approval-toggle { margin-top: 4px; display: flex; gap: 12px; align-items: center; font-size: 11px; }
.approval-toggle label { display: flex; align-items: center; gap: 4px; cursor: pointer; }
.export-json { white-space: pre-wrap; font-family: monospace; font-size: 10px;
               max-height: 200px; overflow: auto; background: #f8f9fa; padding: 8px;
               border: 1px solid var(--border); margin-top: 8px; display: none; }
.section-title { font-size: 16px; font-weight: 600; margin: 16px 16px 8px;
                 padding: 8px; border-bottom: 2px solid #1a1a2e; }
.footer { text-align: center; padding: 16px; font-size: 11px; color: var(--text-muted); }
.filter-bar { display: flex; gap: 8px; padding: 8px 16px; flex-wrap: wrap;
              background: #fff; border-bottom: 1px solid var(--border); position: sticky; top: 58px; z-index: 999; }
.filter-bar input, .filter-bar select { font-size: 12px; padding: 4px 8px; border: 1px solid var(--border);
                                        border-radius: 3px; }
.filter-bar input { flex: 1; min-width: 150px; }
@media print { .header, .filter-bar, .candidate-actions, .btn-front, .btn-inside { display: none; } }
</style>
</head>
<body>
<div class="header">
<div>
<h1>ASD Manual Photo Review v5.7 — Stage A</h1>
<div class="stats" id="globalStats">Loading...</div>
</div>
<div class="actions">
<button onclick="toggleFilter()">🔍 Filter</button>
<button onclick="exportDecisions()">📤 Export JSON</button>
<button onclick="exportCSV()">📤 Export CSV</button>
<button onclick="showOnly('all')">All</button>
<button onclick="showOnly('unreviewed')">Unreviewed</button>
<button onclick="showOnly('approved')">Approved</button>
<button onclick="showOnly('needs_review')">⚠ Needs Review</button>
</div>
</div>
<div class="filter-bar" id="filterBar">
<input type="text" id="searchInput" placeholder="Search model name..." oninput="filterModels()">
<select id="statusFilter" onchange="filterModels()">
<option value="all">All statuses</option>
<option value="approved">Approved</option>
<option value="needs_review">Needs Review</option>
<option value="unreviewed">Unreviewed</option>
<option value="front_empty">Front empty</option>
<option value="inside_empty">Inside empty</option>
<option value="both_empty">Both empty</option>
</select>
<select id="riskFilter" onchange="filterModels()">
<option value="all">All risk levels</option>
<option value="vlm_rejected">VLM Rejected</option>
<option value="low_confidence">Low confidence</option>
<option value="side_unknown">Side unknown</option>
<option value="wide_interior">Wide interior</option>
<option value="fragment">Fragment</option>
<option value="handle">Handle/lock</option>
<option value="collage">Collage</option>
<option value="video">Video</option>
</select>
</div>
<div id="content">''')

# Sort models: risky first, then by model name
def risk_score(m):
    current_front = m.get('current_front', '')
    current_inside = m.get('current_inside', '')
    score = 0
    if not current_front: score += 10
    if not current_inside: score += 5
    if current_front == current_inside and current_front: score += 20
    return -score

all_models_sorted = sorted(all_models, key=lambda m: (risk_score(m), m['model_name']))

for model_idx, model in enumerate(all_models_sorted):
    name = model['model_name']
    src_url = model.get('source_url', '')
    excel_row = model.get('excel_row', '?')
    current_front = model.get('current_front', '')
    current_inside = model.get('current_inside', '')
    comment = model.get('comment', '')
    
    # Determine status
    has_front = bool(current_front)
    has_inside = bool(current_inside)
    same_url = current_front and current_inside and current_front == current_inside
    
    if not has_front and not has_inside:
        status_class = 'status-empty'
        status_label = 'Both empty'
    elif same_url:
        status_class = 'status-review'
        status_label = 'Same URL'
    elif not has_front:
        status_class = 'status-review'
        status_label = 'Front empty'
    elif not has_inside:
        status_class = 'status-review'
        status_label = 'Inside empty'
    elif has_front and has_inside:
        status_class = 'status-ok'
        status_label = 'Both filled'
    else:
        status_class = 'status-review'
        status_label = 'Needs review'
    
    stats['total'] += 1
    if not has_front: stats['front_empty'] += 1
    if not has_inside: stats['inside_empty'] += 1
    if same_url: stats['same_url'] += 1
    
    html_parts.append(f'''<div class="model-card" data-model="{html_mod.escape(name, quote=True)}"
     data-risk="{risk_score(model)}" data-has-front="{'1' if has_front else '0'}"
     data-has-inside="{'1' if has_inside else '0'}"
     data-approved="0" id="model-{model_idx}">
<div class="model-header">
    <div>
        <div class="model-title">{html_mod.escape(name)}</div>
        <div class="model-row">Row {excel_row} | <a href="{html_mod.escape(src_url)}" target="_blank">Source</a></div>
    </div>
    <div><span class="model-status {status_class}">{status_label}</span></div>
</div>
<div class="model-comment"><small>Comment: {html_mod.escape(comment)}</small></div>
<div class="current-pair">
    <div class="pair-item">
        <span class="pair-label">Current Front</span>
        {f'<img src="{html_mod.escape(current_front)}" onerror="this.style.display=\'none\'">' if current_front else '<span style="font-size:10px;color:#999">(empty)</span>'}
    </div>
    <div class="pair-item">
        <span class="pair-label">Current Inside</span>
        {f'<img src="{html_mod.escape(current_inside)}" onerror="this.style.display=\'none\'">' if current_inside else '<span style="font-size:10px;color:#999">(empty)</span>'}
    </div>
</div>
<div class="gallery-index-label">Candidates ({len(model.get("candidates", []))}) »</div>
<div class="candidates-grid">''')
    
    for cand in model.get('candidates', []):
        url = cand['url']
        dl = dl_map.get(url, {})
        vlm = vlm_map.get(url, {})
        
        # Track VLM stats
        if vlm:
            scene = vlm.get('scene_type', '')
            side = vlm.get('side', '')
            conf = vlm.get('confidence', 0)
            rejected = vlm.get('rejection_reason', '')
            if vlm.get('has_door') == False or vlm.get('full_door_visible') == False or vlm.get('door_is_primary_object') == False:
                rejected = '1'
            if conf < 0.9:
                stats['low_confidence'] += 1
            if scene == 'video_thumbnail': stats['video'] += 1
            if scene == 'banner_collage': stats['collage'] += 1
            if scene == 'wide_interior': stats['wide_interior'] += 1
            if scene == 'fragment': stats['fragment'] += 1
            if scene in ('handle_lock', 'technical_detail', 'hinge'): stats['handle'] += 1
            if side == 'unknown': stats['side_unknown'] += 1
            if rejected:
                stats['vlm_rejected'] += 1
        
        is_front = (url == current_front)
        is_inside = (url == current_inside)
        
        html_parts.append(img_thumbnail_html(url, dl, vlm, is_front, is_inside))
    
    # Recommended proposal area
    proposed_front = current_front if has_front else ''
    proposed_inside = current_inside if has_inside else ''
    
    html_parts.append(f'''</div>
<div class="model-proposal">
    <div class="proposal-label">Proposed Pair:</div>
    <div class="pair-item">
        <span class="pair-label">Front</span>
        <img id="prop-front-{model_idx}" src="{html_mod.escape(proposed_front)}" 
             onerror="this.style.display='none'" data-url="{html_mod.escape(proposed_front, quote=True)}">
        <button class="btn-front active" onclick="clearProposal({model_idx}, 'front')">Clear</button>
    </div>
    <div class="pair-item">
        <span class="pair-label">Inside</span>
        <img id="prop-inside-{model_idx}" src="{html_mod.escape(proposed_inside)}"
             onerror="this.style.display='none'" data-url="{html_mod.escape(proposed_inside, quote=True)}">
        <button class="btn-inside active" onclick="clearProposal({model_idx}, 'inside')">Clear</button>
    </div>
    <div class="approval-toggle">
        <label><input type="checkbox" class="approve-check" data-model-idx="{model_idx}"
                      onchange="toggleApprove({model_idx})"> ✅ Approve pair</label>
        <label><input type="checkbox" class="review-check" data-model-idx="{model_idx}"
                      onchange="toggleReview({model_idx})"> ⚠ Manual review required</label>
    </div>
</div>
<div class="model-comment">
    <textarea placeholder="Reviewer comment..." onchange="updateComment({model_idx}, this.value)"
              id="comment-{model_idx}"></textarea>
</div>
</div>''')

html_parts.append('''
</div>
<div class="footer">
Generated: ''' + datetime.now().strftime('%Y-%m-%d %H:%M') + ''' — Stage A: Manual Photo Review v5.7<br>
<em>No changes applied to catalog. Export decisions to JSON/CSV, then proceed to Stage B.</em>
</div>

<!-- Export area -->
<pre class="export-json" id="exportArea"></pre>

<script>
// Global decisions storage
let decisions = {};
let modelData = {};

function updateStats() {
    const total = document.querySelectorAll('.model-card').length;
    const approved = document.querySelectorAll('.approve-check:checked').length;
    const needsReview = document.querySelectorAll('.review-check:checked').length;
    const frontEmpty = document.querySelectorAll('.model-card[data-has-front="0"]').length;
    const insideEmpty = document.querySelectorAll('.model-card[data-has-inside="0"]').length;
    document.getElementById('globalStats').textContent =
        `Total: ${total} | Approved: ${approved} | Review: ${needsReview} | Front empty: ${frontEmpty} | Inside empty: ${insideEmpty}`;
}
updateStats();

function collectModelData(idx) {
    const card = document.getElementById('model-' + idx);
    if (!card) return null;
    const frontImg = document.getElementById('prop-front-' + idx);
    const insideImg = document.getElementById('prop-inside-' + idx);
    const approveCheck = card.querySelector('.approve-check');
    const reviewCheck = card.querySelector('.review-check');
    const commentEl = document.getElementById('comment-' + idx);
    
    return {
        model_name: card.querySelector('.model-title').textContent.trim(),
        excel_row: card.querySelector('.model-row').textContent.split('|')[0].replace('Row ', '').trim(),
        source_url: card.querySelector('.model-row a')?.href || '',
        proposed_front: frontImg?.dataset?.url || '',
        proposed_inside: insideImg?.dataset?.url || '',
        approved: approveCheck?.checked || false,
        manual_review_required: reviewCheck?.checked || false,
        comment: commentEl?.value || '',
        current_front: card.dataset.hasFront === '1' ? (frontImg?.src || '') : '',
        current_inside: card.dataset.hasInside === '1' ? (insideImg?.src || '') : '',
    };
}

function updateComment(idx, val) {
    if (!decisions['model-' + idx]) decisions['model-' + idx] = {};
    decisions['model-' + idx].comment = val;
}

function toggleApprove(idx) {
    const card = document.getElementById('model-' + idx);
    const cb = card.querySelector('.approve-check');
    card.dataset.approved = cb.checked ? '1' : '0';
    if (cb.checked) {
        card.querySelector('.review-check').checked = false;
    }
    updateStats();
}

function toggleReview(idx) {
    const card = document.getElementById('model-' + idx);
    const cb = card.querySelector('.review-check');
    if (cb.checked) {
        card.querySelector('.approve-check').checked = false;
        card.dataset.approved = '0';
    }
    updateStats();
}

function selectFront(url, btn) {
    const card = btn.closest('.model-card');
    const idx = card.id.replace('model-', '');
    const img = document.getElementById('prop-front-' + idx);
    img.src = url;
    img.dataset.url = url;
    img.style.display = 'block';
    card.querySelectorAll('.btn-front').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    updateStats();
}

function selectInside(url, btn) {
    const card = btn.closest('.model-card');
    const idx = card.id.replace('model-', '');
    const img = document.getElementById('prop-inside-' + idx);
    img.src = url;
    img.dataset.url = url;
    img.style.display = 'block';
    card.querySelectorAll('.btn-inside').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    updateStats();
}

function clearProposal(idx, side) {
    const img = document.getElementById('prop-' + side + '-' + idx);
    img.src = '';
    img.dataset.url = '';
    img.style.display = 'none';
}

function toggleSelect(imgEl, url) {
    imgEl.classList.toggle('selected');
}

function showOnly(filter) {
    document.querySelectorAll('.model-card').forEach(c => {
        const approved = c.querySelector('.approve-check')?.checked;
        const review = c.querySelector('.review-check')?.checked;
        const front = c.dataset.hasFront === '1';
        const inside = c.dataset.hasInside === '1';
        let show = true;
        if (filter === 'approved') show = approved;
        else if (filter === 'needs_review') show = review;
        else if (filter === 'unreviewed') show = !approved && !review;
        else if (filter === 'front_empty') show = !front;
        else if (filter === 'inside_empty') show = !inside;
        else if (filter === 'both_empty') show = !front && !inside;
        c.style.display = show ? '' : 'none';
    });
}

function filterModels() {
    const q = document.getElementById('searchInput').value.toLowerCase();
    const status = document.getElementById('statusFilter').value;
    const risk = document.getElementById('riskFilter').value;
    document.querySelectorAll('.model-card').forEach(c => {
        const name = c.querySelector('.model-title').textContent.toLowerCase();
        let show = name.includes(q);
        if (show && status !== 'all') {
            const approved = c.querySelector('.approve-check')?.checked;
            const review = c.querySelector('.review-check')?.checked;
            const front = c.dataset.hasFront === '1';
            const inside = c.dataset.hasInside === '1';
            if (status === 'approved') show = approved;
            else if (status === 'needs_review') show = review;
            else if (status === 'unreviewed') show = !approved && !review;
            else if (status === 'front_empty') show = !front;
            else if (status === 'inside_empty') show = !inside;
            else if (status === 'both_empty') show = !front && !inside;
        }
        c.style.display = show ? '' : 'none';
    });
}

function exportDecisions() {
    const result = [];
    document.querySelectorAll('.model-card').forEach(card => {
        const idx = card.id.replace('model-', '');
        const d = collectModelData(parseInt(idx));
        if (d) result.push(d);
    });
    const json = JSON.stringify(result, null, 2);
    const area = document.getElementById('exportArea');
    area.textContent = json;
    area.style.display = 'block';
    // Download
    const blob = new Blob([json], {type: 'application/json'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'asd_manual_decisions.json';
    a.click();
    URL.revokeObjectURL(url);
}

function exportCSV() {
    const headers = ['excel_row','model_name','source_url','proposed_front','proposed_inside','approved','manual_review_required','comment'];
    let csv = headers.join(';') + '\\n';
    document.querySelectorAll('.model-card').forEach(card => {
        const idx = parseInt(card.id.replace('model-', ''));
        const d = collectModelData(idx);
        if (!d) return;
        const row = [d.excel_row, d.model_name, d.source_url, d.proposed_front, d.proposed_inside,
                     d.approved ? '1' : '0', d.manual_review_required ? '1' : '0',
                     (d.comment || '').replace(/;/g, ',')];
        csv += row.join(';') + '\\n';
    });
    const blob = new Blob([csv], {type: 'text/csv;charset=utf-8;'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'asd_manual_decisions.csv';
    a.click();
    URL.revokeObjectURL(url);
}

function toggleFilter() {
    const bar = document.getElementById('filterBar');
    bar.style.display = bar.style.display === 'none' ? 'flex' : 'none';
}

// Init
document.addEventListener('DOMContentLoaded', function() {
    updateStats();
});
</script>
</body>
</html>''')

html_path = os.path.join(REPORTS_DIR, "asd_manual_review.html")
with open(html_path, 'w', encoding='utf-8') as f:
    f.write(''.join(html_parts))
print(f"HTML review: {html_path}")

# ========================================================================
# BUILD CONTACT SHEET (compact pair overview)
# ========================================================================
print("Building contact sheet...")

contact_parts = []
contact_parts.append('''<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ASD Pair Contact Sheet v5.7</title>
<style>
body { font-family: Arial, sans-serif; margin: 10px; font-size: 11px; }
h1 { font-size: 16px; margin-bottom: 8px; }
.pair-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 6px; }
.pair-card { border: 1px solid #ccc; border-radius: 4px; padding: 6px; background: #fafafa; }
.pair-card h3 { font-size: 11px; margin: 0 0 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.img-row { display: flex; gap: 4px; }
.img-row img { width: 130px; height: 90px; object-fit: cover; border: 1px solid #ddd; border-radius: 2px; }
.img-label { font-size: 9px; color: #666; text-align: center; }
.img-col { display: flex; flex-direction: column; align-items: center; }
.status-tag { display: inline-block; font-size: 9px; padding: 1px 5px; border-radius: 2px; margin-top: 2px; }
.status-ok { background: #d4edda; color: #155724; }
.status-review { background: #fff3cd; color: #856404; }
.status-empty { background: #f8d7da; color: #721c24; }
@media print { .pair-grid { break-inside: avoid; } }
</style>
</head>
<body>
<h1>ASD Pair Contact Sheet — 199 pairs</h1>
<p>Generated: ''' + datetime.now().strftime('%Y-%m-%d %H:%M') + '''</p>
<div class="pair-grid">''')

for model in all_models_sorted:
    name = model['model_name']
    front_url = model.get('current_front', '')
    inside_url = model.get('current_inside', '')
    has_front = bool(front_url)
    has_inside = bool(inside_url)
    same = front_url and inside_url and front_url == inside_url
    
    if not has_front and not has_inside:
        status = 'status-empty'
        label = 'Both empty'
    elif same:
        status = 'status-review'
        label = 'Same URL'
    elif not has_front:
        status = 'status-review'
        label = 'No front'
    elif not has_inside:
        status = 'status-review'
        label = 'No inside'
    else:
        status = 'status-ok'
        label = 'OK'
    
    front_html = f'<img src="{html_mod.escape(front_url)}" onerror="this.style.display=\'none\'">' if front_url else '<span style="color:#999">—</span>'
    inside_html = f'<img src="{html_mod.escape(inside_url)}" onerror="this.style.display=\'none\'">' if inside_url else '<span style="color:#999">—</span>'
    
    contact_parts.append(f'''<div class="pair-card">
    <h3 title="{html_mod.escape(name)}">{html_mod.escape(name)[:60]}</h3>
    <div class="img-row">
        <div class="img-col">{front_html}<span class="img-label">Front</span></div>
        <div class="img-col">{inside_html}<span class="img-label">Inside</span></div>
    </div>
    <span class="status-tag {status}">{label}</span>
</div>''')

contact_parts.append('''</div></body></html>''')

contact_path = os.path.join(REPORTS_DIR, "asd_pair_contact_sheet.html")
with open(contact_path, 'w', encoding='utf-8') as f:
    f.write(''.join(contact_parts))
print(f"Contact sheet: {contact_path}")

# ========================================================================
# BUILD RISKY ROWS XLSX
# ========================================================================
print("Building risky rows report...")

hdr_font = Font(bold=True, color="FFFFFF", size=10)
hdr_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
yellow_fill = PatternFill(start_color="FFEECC", end_color="FFEECC", fill_type="solid")

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Risky Rows"

headers = ['excel_row', 'model_name', 'source_url',
           'has_front', 'has_inside', 'same_url',
           'front_url', 'inside_url',
           'candidate_count', 'downloaded_count',
           'risk_flags', 'comment',
           'vlm_rejected_count', 'vlm_low_conf_count',
           'vlm_side_unknown_count', 'vlm_wide_interior_count',
           'vlm_fragment_count', 'vlm_handle_count',
           'vlm_collage_count', 'vlm_video_count']

for ci, h in enumerate(headers, 1):
    c = ws.cell(row=1, column=ci, value=h)
    c.font = hdr_font; c.fill = hdr_fill; c.alignment = hdr_align; c.border = thin_border

row_num = 2
for model in all_models:
    name = model['model_name']
    front_url = model.get('current_front', '')
    inside_url = model.get('current_inside', '')
    has_front = bool(front_url)
    has_inside = bool(inside_url)
    same = front_url and inside_url and front_url == inside_url
    
    risk_flags = []
    if not has_front: risk_flags.append('front_empty')
    if not has_inside: risk_flags.append('inside_empty')
    if same: risk_flags.append('same_url')
    
    # Count VLM issues
    vlm_rejected = 0
    vlm_low_conf = 0
    vlm_side_unknown = 0
    vlm_wide_interior = 0
    vlm_fragment = 0
    vlm_handle = 0
    vlm_collage = 0
    vlm_video = 0
    
    for cand in model.get('candidates', []):
        url = cand['url']
        vlm = vlm_map.get(url, {})
        if not vlm:
            continue
        scene = vlm.get('scene_type', '')
        conf = vlm.get('confidence', 1)
        side = vlm.get('side', '')
        rejected = vlm.get('rejection_reason', '')
        if vlm.get('has_door') == False or vlm.get('full_door_visible') == False or vlm.get('door_is_primary_object') == False:
            vlm_rejected += 1
        if conf < 0.9:
            vlm_low_conf += 1
        if side == 'unknown':
            vlm_side_unknown += 1
        if scene == 'wide_interior':
            vlm_wide_interior += 1
            if 'wide_interior' not in risk_flags: risk_flags.append('wide_interior')
        if scene == 'fragment':
            vlm_fragment += 1
            if 'fragment' not in risk_flags: risk_flags.append('fragment')
        if scene in ('handle_lock', 'technical_detail', 'hinge'):
            vlm_handle += 1
            if 'handle' not in risk_flags: risk_flags.append('handle')
        if scene == 'banner_collage':
            vlm_collage += 1
            if 'collage' not in risk_flags: risk_flags.append('collage')
        if scene == 'video_thumbnail':
            vlm_video += 1
            if 'video' not in risk_flags: risk_flags.append('video')
    
    # Mark as risky if any flags
    if not risk_flags and not vlm_low_conf and not vlm_side_unknown:
        continue
    
    downloaded = sum(1 for c in model.get('candidates', []) if dl_map.get(c['url'], {}).get('downloaded'))
    
    vals = [
        model.get('excel_row', ''), name, model.get('source_url', ''),
        1 if has_front else 0, 1 if has_inside else 0, 1 if same else 0,
        front_url, inside_url,
        len(model.get('candidates', [])), downloaded,
        '; '.join(risk_flags) if risk_flags else '',
        model.get('comment', ''),
        vlm_rejected, vlm_low_conf, vlm_side_unknown,
        vlm_wide_interior, vlm_fragment, vlm_handle,
        vlm_collage, vlm_video,
    ]
    
    for ci, val in enumerate(vals, 1):
        cell = ws.cell(row=row_num, column=ci, value=val)
        cell.border = thin_border
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        if risk_flags:
            cell.fill = red_fill if 'empty' in ' '.join(risk_flags) else yellow_fill
    row_num += 1

risky_path = os.path.join(REPORTS_DIR, "asd_risky_rows.xlsx")
wb.save(risky_path)
print(f"Risky rows: {risky_path} ({row_num - 2} rows)")

# ========================================================================
# SUMMARY
# ========================================================================
print(f"\n{'='*60}")
print(f"STAGE A - STEP 2 COMPLETE")
print(f"{'='*60}")
print(f"  ASD models:           {stats['total']}")
print(f"  Front empty:          {stats['front_empty']}")
print(f"  Inside empty:         {stats['inside_empty']}")
print(f"  Same URL:             {stats['same_url']}")
print(f"  HTML review:          {html_path}")
print(f"  Contact sheet:        {contact_path}")
print(f"  Risky rows:           {risky_path}")
print(f"  VLM input file:       {vlm_input_path}")
print(f"  (Place VLM output at: {vlm_path})")
