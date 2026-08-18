"""
Stage A - Step 4b: Rebuild Lord review HTML files only.
Full state management, localStorage, import/export, cross-page sync.
ASD batches NOT touched.
"""
import csv, json, os, re, html as html_mod
from datetime import datetime

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
V56_DIR = os.path.join(BASE, "output", "delivery_v5_6")
CACHE_LORD = os.path.join(BASE, "output", "manual_photo_review_v5_7", "cache", "lord")
OUTPUT = os.path.join(BASE, "output", "manual_photo_review_v5_7")
LORD_DIR = os.path.join(OUTPUT, "lord_portable_review")
LORD_IMAGES = os.path.join(LORD_DIR, "images")
LORD = chr(0x041B) + chr(0x043E) + chr(0x0440) + chr(0x0434)

STORAGE_KEY = "lord_manual_review_v5_7"

# Load catalog
with open(os.path.join(V56_DIR, "doors_catalog.csv"), encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
lord_rows = [r for r in rows if r['producer'] == LORD]

# Build image mapping
def find_local_image(row):
    name = row['model_name']
    safe = re.sub(r'[<>:"/\\|?*]', '_', name)[:120]
    images_dir = LORD_IMAGES
    if not os.path.isdir(images_dir):
        return ''
    # Exact match
    for ff in os.listdir(images_dir):
        if ff.startswith(safe + '_') or ff.startswith(name + '_'):
            return ff
    # Partial match
    for ff in os.listdir(images_dir):
        if ff.startswith(safe[:30]):
            return ff
    return ''

# Build full row data
lord_data = []
for r_idx, r in enumerate(lord_rows):
    local_img = find_local_image(r)
    lord_data.append({
        'excel_row': r_idx + 2,
        'producer': LORD,
        'target_line': r.get('target_line', ''),
        'model_name': r['model_name'],
        'source_url': r.get('source_url', ''),
        'current_image_url': r.get('image_url_front', ''),
        'local_image': local_img,
    })

print(f"Lord rows: {len(lord_data)}")
print(f"With local image: {sum(1 for d in lord_data if d['local_image'])}")
print(f"Without local image: {sum(1 for d in lord_data if not d['local_image'])}")

# Build flagged subset
flagged_indices = []
for d in lord_data:
    f_url = d['current_image_url']
    if not f_url: continue
    flags = []
    url_lower = f_url.lower()
    if any(kw in url_lower for kw in ['steklo', 'glass', 'vitrin', 'window']):
        flags.append('front_glass_keyword')
    if any(kw in url_lower for kw in ['fragment', 'detail']):
        flags.append('front_fragment_keyword')
    if any(kw in url_lower for kw in ['handle', 'lock', 'ruchk', 'zamok']):
        flags.append('front_handle_keyword')
    if flags:
        d['flags'] = '; '.join(flags)
        flagged_indices.append(len(flagged_indices))
        flagged_data.append(d) if False else None

# Rebuild flagged data properly
flagged_data = []
for d in lord_data:
    f_url = d['current_image_url']
    if not f_url: continue
    flags = []
    url_lower = f_url.lower()
    if any(kw in url_lower for kw in ['steklo', 'glass', 'vitrin', 'window']):
        flags.append('front_glass_keyword')
    if any(kw in url_lower for kw in ['fragment', 'detail']):
        flags.append('front_fragment_keyword')
    if any(kw in url_lower for kw in ['handle', 'lock', 'ruchk', 'zamok']):
        flags.append('front_handle_keyword')
    if flags:
        d['flags'] = '; '.join(flags)
        flagged_data.append(d)

print(f"Flagged: {len(flagged_data)}")

# ========================================================================
# JavaScript state management (shared between both HTML files)
# ========================================================================
JS_STATE = r"""
var STORAGE_KEY = 'lord_manual_review_v5_7';
var ROWS = %ROWS%;
var state = {};
var dirty = false;

function rowKey(r) { return String(r.excel_row); }

function initState() {
    ROWS.forEach(function(r) {
        var k = rowKey(r);
        state[k] = {
            excel_row: r.excel_row,
            producer: r.producer,
            target_line: r.target_line,
            model_name: r.model_name,
            source_url: r.source_url,
            current_image_url: r.current_image_url,
            local_image: r.local_image,
            decision: '',
            approved_replacement_url: '',
            reviewer_comment: '',
            explicitly_reviewed: false
        };
    });
}

function saveState() {
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
        dirty = false;
    } catch(e) {}
}

function loadState() {
    try {
        var saved = localStorage.getItem(STORAGE_KEY);
        if (saved) {
            var parsed = JSON.parse(saved);
            Object.keys(parsed).forEach(function(k) {
                if (state[k]) {
                    state[k].decision = parsed[k].decision || '';
                    state[k].approved_replacement_url = parsed[k].approved_replacement_url || '';
                    state[k].reviewer_comment = parsed[k].reviewer_comment || '';
                    state[k].explicitly_reviewed = parsed[k].explicitly_reviewed || false;
                }
            });
            return true;
        }
    } catch(e) {}
    return false;
}

function setDecision(k, decision) {
    if (!state[k]) return;
    var r = state[k];
    // Keep requires local image
    if (decision === 'keep_full_door' && !r.local_image) {
        if (!confirm('This row has no local image. The remote image must be opened first. Continue with keep?')) return;
    }
    r.decision = decision;
    r.explicitly_reviewed = true;
    dirty = true;
    saveState();
    updateRowVisual(k);
    updateStats();
}

function setReplacement(k, url) {
    if (!state[k]) return;
    state[k].approved_replacement_url = url;
    if (url) {
        state[k].decision = 'replace_with_approved_url';
        state[k].explicitly_reviewed = true;
    }
    dirty = true;
    saveState();
    updateRowVisual(k);
    updateStats();
}

function setComment(k, val) {
    if (!state[k]) return;
    state[k].reviewer_comment = val;
    dirty = true;
    saveState();
}

function updateRowVisual(k) {
    var r = state[k];
    var row = document.getElementById('row-' + k);
    if (!row) return;
    var btns = row.querySelectorAll('.dec-btn');
    btns.forEach(function(b) {
        b.classList.remove('active');
        if (b.dataset.decision === r.decision) b.classList.add('active');
    });
    var statusEl = row.querySelector('.row-status');
    if (statusEl) {
        if (r.explicitly_reviewed) {
            statusEl.textContent = r.decision || 'reviewed';
            statusEl.className = 'row-status reviewed';
        } else {
            statusEl.textContent = 'not reviewed';
            statusEl.className = 'row-status not-reviewed';
        }
    }
}

function updateStats() {
    var total = ROWS.length;
    var reviewed = 0, notReviewed = 0, keep = 0, reject = 0, needsReview = 0;
    ROWS.forEach(function(r) {
        var k = rowKey(r);
        var s = state[k];
        if (s.explicitly_reviewed) reviewed++; else notReviewed++;
        if (s.decision === 'keep_full_door') keep++;
        else if (s.decision && s.decision.indexOf('reject') >= 0) reject++;
        else if (s.decision === 'needs_review') needsReview++;
    });
    var el = document.getElementById('statsDisplay');
    if (el) el.textContent = 'Total: ' + total + ' | Reviewed: ' + reviewed + ' | Not reviewed: ' + notReviewed +
        ' | Keep: ' + keep + ' | Reject: ' + reject + ' | Needs review: ' + needsReview;
}

function getDecisions() {
    var result = [];
    ROWS.forEach(function(r) {
        var k = rowKey(r);
        var s = state[k];
        result.push({
            excel_row: s.excel_row,
            producer: s.producer,
            target_line: s.target_line,
            model_name: s.model_name,
            source_url: s.source_url,
            current_image_url: s.current_image_url,
            local_image: s.local_image,
            decision: s.decision,
            approved_replacement_url: s.approved_replacement_url,
            reviewer_comment: s.reviewer_comment,
            explicitly_reviewed: s.explicitly_reviewed ? 1 : 0
        });
    });
    return result;
}

function exportJSON() {
    var data = getDecisions();
    var blob = new Blob([JSON.stringify(data, null, 2)], {type: 'application/json'});
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = 'lord_manual_decisions.json';
    a.click(); URL.revokeObjectURL(url);
    dirty = false;
}

function exportCSV() {
    var data = getDecisions();
    var headers = ['excel_row','producer','target_line','model_name','source_url',
                   'current_image_url','local_image','decision','approved_replacement_url',
                   'reviewer_comment','explicitly_reviewed'];
    var csv = headers.join(';') + '\n';
    data.forEach(function(r) {
        var row = headers.map(function(h) {
            var v = (r[h] || '').toString();
            return v.replace(/;/g, ',').replace(/\n/g, ' ');
        });
        csv += row.join(';') + '\n';
    });
    var blob = new Blob(['\ufeff' + csv], {type: 'text/csv;charset=utf-8;'});
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = 'lord_manual_decisions.csv';
    a.click(); URL.revokeObjectURL(url);
    dirty = false;
}

function importJSON() {
    document.getElementById('importFileInput').click();
}

function handleImport(event) {
    var file = event.target.files[0];
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function(e) {
        try {
            var data = JSON.parse(e.target.result);
            if (!Array.isArray(data)) { alert('Invalid format: expected array'); return; }
            var count = 0;
            data.forEach(function(d) {
                var k = String(d.excel_row);
                if (state[k]) {
                    state[k].decision = d.decision || '';
                    state[k].approved_replacement_url = d.approved_replacement_url || '';
                    state[k].reviewer_comment = d.reviewer_comment || '';
                    state[k].explicitly_reviewed = !!d.explicitly_reviewed;
                    count++;
                }
            });
            saveState();
            ROWS.forEach(function(r) { updateRowVisual(rowKey(r)); });
            updateStats();
            alert('Imported ' + count + ' decisions.');
        } catch(err) {
            alert('Import error: ' + err.message);
        }
    };
    reader.readAsText(file);
    event.target.value = '';
}

function resetAll() {
    if (!confirm('Reset ALL decisions? This cannot be undone.')) return;
    initState();
    saveState();
    ROWS.forEach(function(r) { updateRowVisual(rowKey(r)); });
    updateStats();
}

window.addEventListener('beforeunload', function(e) {
    if (dirty) { e.preventDefault(); e.returnValue = ''; }
});
"""

# ========================================================================
# Build lord_flagged_review.html
# ========================================================================
print("\n=== Building lord_flagged_review.html ===")

rows_json = json.dumps(lord_data, ensure_ascii=False)
js_code = JS_STATE.replace('%ROWS%', rows_json)

# Build flagged rows HTML
flagged_html_rows = []
for idx, d in enumerate(flagged_data):
    k = str(d['excel_row'])
    has_local = bool(d['local_image'])
    img_html = f'<img src="images/{html_mod.escape(d["local_image"])}">' if has_local else \
               f'<div class="no-image"><a href="{html_mod.escape(d["current_image_url"])}" target="_blank">Open remote</a></div>'
    
    flags_html = ''.join(f'<span class="flag flag-glass">{html_mod.escape(f)}</span>' for f in d.get('flags', '').split('; '))
    
    # Comment box
    comment_html = f'<input type="text" class="row-comment" placeholder="Comment..." onchange="setComment(\'{k}\',this.value)">'
    
    # Replacement URL
    replacement_html = f'<input type="text" class="row-replacement" placeholder="Replacement URL" onchange="setReplacement(\'{k}\',this.value)">'
    
    flagged_html_rows.append(f'''<tr id="row-{k}">
<td>{idx+1}</td>
<td title="{html_mod.escape(d['model_name'])}">{html_mod.escape(d['model_name'][:40])}</td>
<td>{html_mod.escape(d.get('target_line','')[:20])}</td>
<td>{img_html}</td>
<td>{flags_html}</td>
<td>
<button class="dec-btn" data-decision="keep_full_door" onclick="setDecision('{k}','keep_full_door')">Keep full door</button>
<button class="dec-btn" data-decision="reject_fragment" onclick="setDecision('{k}','reject_fragment')">Reject fragment</button>
<button class="dec-btn" data-decision="reject_detail" onclick="setDecision('{k}','reject_detail')">Reject detail</button>
<button class="dec-btn" data-decision="reject_wrong_model" onclick="setDecision('{k}','reject_wrong_model')">Wrong model</button>
<button class="dec-btn" data-decision="clear_no_replacement" onclick="setDecision('{k}','clear_no_replacement')">Clear</button>
<button class="dec-btn" data-decision="needs_review" onclick="setDecision('{k}','needs_review')">Needs review</button>
</td>
<td>{replacement_html}</td>
<td>{comment_html}</td>
<td><span class="row-status not-reviewed">not reviewed</span></td>
</tr>''')

flagged_html = f'''<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Lord Flagged Review</title>
<style>
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
       background: #f5f6fa; color: #2c3e50; font-size: 12px; margin: 0; }}
.header {{ background: #2c3e50; color: #ecf0f1; padding: 12px 16px; position: sticky; top: 0; z-index: 1000; }}
.header h1 {{ font-size: 16px; margin: 0; }}
.controls {{ background: #fff; padding: 8px 16px; border-bottom: 1px solid #ddd;
            position: sticky; top: 52px; z-index: 999; display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }}
.controls button {{ font-size: 11px; padding: 4px 10px; border: 1px solid #3498db; background: #3498db;
                   color: #fff; border-radius: 3px; cursor: pointer; }}
.controls button:hover {{ background: #2980b9; }}
.btn-export {{ background: #27ae60 !important; border-color: #27ae60 !important; }}
.btn-export:hover {{ background: #219a52 !important; }}
.btn-import {{ background: #8e44ad !important; border-color: #8e44ad !important; }}
.btn-import:hover {{ background: #7d3c98 !important; }}
.btn-reset {{ background: #e74c3c !important; border-color: #e74c3c !important; }}
.btn-reset:hover {{ background: #c0392b !important; }}
table {{ border-collapse: collapse; width: 100%; margin-top: 10px; }}
th, td {{ border: 1px solid #ccc; padding: 4px 6px; text-align: left; vertical-align: top; }}
th {{ background: #34495e; color: #fff; position: sticky; top: 96px; z-index: 998; font-size: 11px; }}
tr.flagged {{ background: #fff5f5; }}
img {{ max-width: 100px; max-height: 85px; object-fit: contain; border: 1px solid #ddd; border-radius: 2px; background: #f0f0f0; }}
.no-image {{ background: #f8d7da; padding: 8px; font-size: 10px; color: #721c24; border: 1px dashed #e74c3c; text-align: center; }}
.no-image a {{ color: #3498db; }}
.flag {{ display: inline-block; font-size: 9px; padding: 1px 4px; border-radius: 2px; margin: 1px; }}
.flag-glass {{ background: #cce5ff; color: #004085; }}
.dec-btn {{ font-size: 9px; padding: 2px 5px; border: 1px solid #ccc; border-radius: 3px; cursor: pointer; background: #fff; margin: 1px; }}
.dec-btn:hover {{ background: #eee; }}
.dec-btn.active {{ background: #27ae60; color: #fff; border-color: #27ae60; }}
.row-comment {{ width: 120px; font-size: 10px; padding: 2px 4px; border: 1px solid #ddd; border-radius: 2px; }}
.row-replacement {{ width: 150px; font-size: 10px; padding: 2px 4px; border: 1px solid #ddd; border-radius: 2px; }}
.row-status {{ font-size: 9px; padding: 1px 4px; border-radius: 2px; white-space: nowrap; }}
.row-status.reviewed {{ background: #d5f5e3; color: #1e8449; }}
.row-status.not-reviewed {{ background: #fadbd8; color: #922b21; }}
#statsDisplay {{ font-size: 11px; color: #7f8c8d; }}
</style>
</head>
<body>
<div class="header">
    <h1>Lord Flagged Review — {len(flagged_data)} flagged rows out of {len(lord_data)} total</h1>
</div>
<div class="controls">
    <span id="statsDisplay"></span>
    <button onclick="exportJSON()" class="btn-export">Export JSON</button>
    <button onclick="exportCSV()" class="btn-export">Export CSV</button>
    <button onclick="importJSON()" class="btn-import">Import JSON</button>
    <button onclick="resetAll()" class="btn-reset">Reset all</button>
</div>
<table>
<tr><th>#</th><th>Model</th><th>Line</th><th>Front</th><th>Flags</th><th>Decision</th><th>Replacement URL</th><th>Comment</th><th>Status</th></tr>
{''.join(flagged_html_rows)}
</table>
<input type="file" id="importFileInput" accept=".json" style="display:none" onchange="handleImport(event)">
<script>
{js_code}
initState();
loadState();
ROWS.forEach(function(r) {{ updateRowVisual(rowKey(r)); }});
updateStats();
</script>
</body>
</html>'''

with open(os.path.join(LORD_DIR, 'lord_flagged_review.html'), 'w', encoding='utf-8') as f:
    f.write(flagged_html)
print(f"  Written: lord_flagged_review.html ({len(flagged_html)} bytes)")

# ========================================================================
# Build lord_all_contact_sheet.html
# ========================================================================
print("\n=== Building lord_all_contact_sheet.html ===")

# Build card HTML for each row
cards_html_parts = []
page_size = 50
for page_start in range(0, len(lord_data), page_size):
    page_rows = lord_data[page_start:page_start + page_size]
    page_num = page_start // page_size + 1
    
    cards_html_parts.append(f'<div class="page"><h2>Page {page_num} ({page_start+1}\u2013{page_start+len(page_rows)})</h2><div class="grid">')
    
    for d in page_rows:
        k = str(d['excel_row'])
        has_local = bool(d['local_image'])
        
        if has_local:
            img_html = f'<img src="images/{html_mod.escape(d["local_image"])}" loading="lazy">'
        else:
            img_html = f'<div class="no-image"><a href="{html_mod.escape(d["current_image_url"])}" target="_blank" onclick="event.stopPropagation()">Remote</a></div>'
        
        cards_html_parts.append(f'''<div class="card" id="row-{k}">
{img_html}
<div class="name" title="{html_mod.escape(d['model_name'])}">{html_mod.escape(d['model_name'][:35])}</div>
<div class="card-actions">
<button class="dec-btn" data-decision="keep_full_door" onclick="setDecision('{k}','keep_full_door')">Keep</button>
<button class="dec-btn" data-decision="reject_fragment" onclick="setDecision('{k}','reject_fragment')">Reject frag</button>
<button class="dec-btn" data-decision="reject_detail" onclick="setDecision('{k}','reject_detail')">Reject detail</button>
<button class="dec-btn" data-decision="reject_wrong_model" onclick="setDecision('{k}','reject_wrong_model')">Wrong</button>
<button class="dec-btn" data-decision="needs_review" onclick="setDecision('{k}','needs_review')">Review</button>
</div>
<div class="card-comment"><input type="text" placeholder="Comment..." onchange="setComment('{k}',this.value)"></div>
<span class="row-status not-reviewed">not reviewed</span>
</div>''')
    
    cards_html_parts.append('</div></div>')

contact_html = f'''<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Lord All Images Contact Sheet</title>
<style>
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
       background: #f5f6fa; color: #2c3e50; font-size: 10px; margin: 0; }}
.header {{ background: #2c3e50; color: #ecf0f1; padding: 12px 16px; position: sticky; top: 0; z-index: 1000; }}
.header h1 {{ font-size: 14px; margin: 0; }}
.controls {{ background: #fff; padding: 8px 16px; border-bottom: 1px solid #ddd;
            position: sticky; top: 46px; z-index: 999; display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }}
.controls button {{ font-size: 11px; padding: 4px 10px; border: 1px solid #3498db; background: #3498db;
                   color: #fff; border-radius: 3px; cursor: pointer; }}
.controls button:hover {{ background: #2980b9; }}
.btn-export {{ background: #27ae60 !important; border-color: #27ae60 !important; }}
.btn-import {{ background: #8e44ad !important; border-color: #8e44ad !important; }}
.btn-reset {{ background: #e74c3c !important; border-color: #e74c3c !important; }}
.page {{ page-break-after: always; margin-bottom: 20px; }}
.grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(145px, 1fr)); gap: 4px; }}
.card {{ border: 1px solid #ddd; border-radius: 3px; padding: 3px; background: #fafafa; }}
.card img {{ width: 100%; height: 90px; object-fit: contain; border: 1px solid #eee; border-radius: 2px; background: #f0f0f0; }}
.card .name {{ font-size: 9px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; margin-top: 2px; }}
.no-image {{ background: #f8d7da; height: 90px; display: flex; align-items: center; justify-content: center;
             font-size: 9px; color: #721c24; border: 1px dashed #e74c3c; }}
.no-image a {{ color: #3498db; }}
.card-actions {{ display: flex; gap: 1px; flex-wrap: wrap; margin-top: 2px; }}
.card-comment input {{ width: 100%; font-size: 8px; padding: 1px 3px; border: 1px solid #ddd; border-radius: 2px; }}
.dec-btn {{ font-size: 8px; padding: 1px 3px; border: 1px solid #ccc; border-radius: 2px; cursor: pointer; background: #fff; }}
.dec-btn:hover {{ background: #eee; }}
.dec-btn.active {{ background: #27ae60; color: #fff; border-color: #27ae60; }}
.row-status {{ font-size: 8px; padding: 1px 3px; border-radius: 2px; display: block; margin-top: 1px; }}
.row-status.reviewed {{ background: #d5f5e3; color: #1e8449; }}
.row-status.not-reviewed {{ background: #fadbd8; color: #922b21; }}
#statsDisplay {{ font-size: 11px; color: #7f8c8d; }}
</style>
</head>
<body>
<div class="header">
    <h1>Lord All Images Contact Sheet \u2014 {len(lord_data)} models</h1>
</div>
<div class="controls">
    <span id="statsDisplay"></span>
    <button onclick="exportJSON()" class="btn-export">Export JSON</button>
    <button onclick="exportCSV()" class="btn-export">Export CSV</button>
    <button onclick="importJSON()" class="btn-import">Import JSON</button>
    <button onclick="resetAll()" class="btn-reset">Reset all</button>
</div>
{''.join(cards_html_parts)}
<input type="file" id="importFileInput" accept=".json" style="display:none" onchange="handleImport(event)">
<script>
{js_code}
initState();
loadState();
ROWS.forEach(function(r) {{ updateRowVisual(rowKey(r)); }});
updateStats();
</script>
</body>
</html>'''

with open(os.path.join(LORD_DIR, 'lord_all_contact_sheet.html'), 'w', encoding='utf-8') as f:
    f.write(contact_html)
print(f"  Written: lord_all_contact_sheet.html ({len(contact_html)} bytes)")

# ========================================================================
# Rebuild decision templates
# ========================================================================
print("\n=== Rebuilding decision templates ===")

# JSON
json_path = os.path.join(LORD_DIR, 'lord_manual_decisions.json')
with open(json_path, 'w', encoding='utf-8') as f:
    json.dump(lord_data, f, ensure_ascii=False, indent=2)
print(f"  Written: lord_manual_decisions.json ({len(lord_data)} rows)")

# CSV
csv_path = os.path.join(LORD_DIR, 'lord_manual_decisions.csv')
fields = ['excel_row','producer','target_line','model_name','source_url',
          'current_image_url','local_image','decision','approved_replacement_url',
          'reviewer_comment','explicitly_reviewed']
with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=fields, delimiter=';')
    writer.writeheader()
    for d in lord_data:
        writer.writerow({k: d.get(k, '' if k != 'explicitly_reviewed' else 0) for k in fields})
print(f"  Written: lord_manual_decisions.csv ({len(lord_data)} rows)")

# Remove old templates
for old in ['lord_decisions_template.json', 'lord_decisions_template.csv']:
    p = os.path.join(LORD_DIR, old)
    if os.path.exists(p):
        os.remove(p)
        print(f"  Removed: {old}")

print("\nLord review rebuild complete.")
print(f"  Flagged rows: {len(flagged_data)}")
print(f"  Total rows: {len(lord_data)}")
print(f"  With local image: {sum(1 for d in lord_data if d['local_image'])}")
print(f"  Without local image: {sum(1 for d in lord_data if not d['local_image'])}")
