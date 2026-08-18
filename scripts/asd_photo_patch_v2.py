"""ASD Photo Patch v2 - strict full-door image selection with dimension analysis"""
import csv, json, os, re, requests, io, time, sys
from bs4 import BeautifulSoup
from urllib.parse import urljoin, unquote
from collections import defaultdict, Counter
from concurrent.futures import ThreadPoolExecutor, TimeoutError
from datetime import datetime
from PIL import Image

BASE_DIR = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
OUT_DIR = os.path.join(BASE_DIR, "output", "asd_photo_patch_v2")
os.makedirs(OUT_DIR, exist_ok=True)

SOURCE_CSV = os.path.join(BASE_DIR, "output", "asd_photo_patch", "asd_rows_v14_source.csv")
CACHE_FILE = os.path.join(OUT_DIR, "image_dim_cache.json")

MAX_WORKERS = 4
REQUEST_TIMEOUT = 10
ROW_TIMEOUT = 120  # max seconds per row
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
ASD_ROOT = "https://as-doors.ru"

# Global session + dimension cache
_session = None
_image_dim_cache = {}  # url -> (w, h)

def _get_session():
    global _session
    if _session is None:
        _session = requests.Session()
        _session.headers.update({"User-Agent": USER_AGENT})
        adapter = requests.adapters.HTTPAdapter(
            max_retries=2,
            pool_connections=10,
            pool_maxsize=20
        )
        _session.mount("https://", adapter)
    return _session

def _load_cache():
    global _image_dim_cache
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, encoding="utf-8") as f:
                raw = json.load(f)
                _image_dim_cache = {k: (v[0], v[1]) for k, v in raw.items()}
        except Exception:
            _image_dim_cache = {}

def _save_cache():
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump({k: list(v) for k, v in _image_dim_cache.items()}, f)
    except Exception:
        pass

def _get_cache():
    return _image_dim_cache

_load_cache()

# Extract filename prefix from a URL (e.g., "protivo-pozharnie-3.jpg" → "protivo-pozharnie")
def filename_prefix(url):
    fn = url.split("/")[-1] if "/" in url else url
    fn = re.sub(r'\.(jpg|jpeg|png|webp|gif)$', '', fn, flags=re.I)
    # Remove trailing numbers (size, year, index)
    fn = re.sub(r'[-_]\d+x\d+.*$', '', fn)  # Remove "1058x0_3a2" type suffixes
    fn = re.sub(r'[-_]\d{4}[-_].*$', '', fn)  # Remove "2026-12" type suffixes  
    fn = re.sub(r'[-_]\d+$', '', fn)  # Remove trailing "-3", "_12" etc.
    fn = re.sub(r'[-_]+$', '', fn)  # Clean trailing dashes
    return fn.lower()

def load_asd_rows():
    with open(SOURCE_CSV, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f, delimiter=";"))

def fetch_page(url):
    try:
        r = requests.get(url, timeout=REQUEST_TIMEOUT, headers={"User-Agent": USER_AGENT})
        return r if r.status_code == 200 else None
    except Exception:
        return None

def make_abs(src):
    if src.startswith("http"):
        return src.split("?")[0]
    if src.startswith("/"):
        return ASD_ROOT + src
    if re.match(r'^(img|images|img202)', src):
        return ASD_ROOT + "/" + src
    return urljoin(ASD_ROOT + "/", src)

def get_image_info(url):
    cache = _get_cache()
    if url in cache:
        return cache[url][0], cache[url][1], True
    try:
        sess = _get_session()
        r = sess.get(url, timeout=8, stream=True)
        if r.status_code != 200:
            return None, None, False
        content = io.BytesIO()
        for chunk in r.iter_content(8192):
            content.write(chunk)
            if content.tell() > 65536:
                break
        r.close()
        content.seek(0)
        img = Image.open(content)
        w, h = img.size[0], img.size[1]
        cache[url] = (w, h)
        _save_cache()
        return w, h, True
    except Exception:
        cache[url] = (None, None)
        return None, None, False

def is_product_path(src_lower):
    """Check if a URL path is a product image path (not icon/star/etc)."""
    if any(x in src_lower for x in ["/star/", "/ico/", "capcha", ".svg", "yandex", 
                                     "mc.yandex", "arrow", "logo", "novinka", 
                                     "phone", "whatsapp", "review"]):
        return False
    if any(x in src_lower for x in ["/img/", "/images/", "/img202", "/assets/cache_image/", "img/", "images/"]):
        return True
    return False

def extract_gallery_prefix(soup):
    """
    Extract the most common filename prefix from gallery/slider images.
    Returns the prefix string or None.
    """
    prefixes = []
    for selector in [".gallery a", ".door-slider img", ".door-slider a",
                     ".gallery img", "a[data-gallery]", "a[rel*='gallery']"]:
        for el in soup.select(selector):
            href = el.get("href","") or el.get("src","") or el.get("data-src","") or ""
            if not href or href.startswith("#"):
                continue
            abs_url = make_abs(href)
            if is_product_path(abs_url.lower()):
                p = filename_prefix(abs_url)
                if p and len(p) > 1:
                    prefixes.append(p)
    if not prefixes:
        return None
    return Counter(prefixes).most_common(1)[0][0]

def extract_candidates(soup, gallery_prefix):
    """Extract all candidate image URLs with metadata."""
    candidates = []  # (url, source_tag, is_gallery, prefix)
    seen = set()
    
    def add(url, tag, is_gallery=False):
        if not url:
            return
        abs_url = make_abs(url)
        if abs_url in seen:
            return
        seen.add(abs_url)
        p = filename_prefix(abs_url)
        candidates.append((abs_url, tag, is_gallery, p))
    
    # 1. Gallery <a> links (highest priority - direct full-size product images)
    for selector in [".gallery a", "a[data-gallery]", "a[rel*='gallery']"]:
        for a in soup.select(selector):
            href = a.get("href","")
            if href and not href.startswith("#") and not href.startswith("javascript"):
                add(href, "gallery_link", is_gallery=True)
    
    # 2. Gallery/door-slider img tags
    for selector in [".door-slider img", ".gallery img", ".door-slider source"]:
        for el in soup.select(selector):
            src = el.get("src","") or el.get("data-src","") or el.get("data-lazy","") or ""
            if src:
                add(src, "slider_img", is_gallery=True)
    
    # 3. All product-related img tags
    for img in soup.find_all("img"):
        src = img.get("src","") or img.get("data-src","") or ""
        if not src or len(src) < 10:
            continue
        src_l = src.lower()
        if not is_product_path(src_l):
            continue
        # Skip if already added via gallery
        abs_url = make_abs(src)
        if abs_url in seen:
            continue
        add(src, "img_tag", is_gallery=False)
    
    # 4. og:image
    og = soup.select_one('meta[property="og:image"]')
    if og:
        add(og.get("content",""), "og_image", is_gallery=False)
    
    return candidates

def score_candidate(url, w, h, source_tag, is_gallery, prefix, gallery_prefix, url_slug):
    """Score a candidate image for 'full-door-ness'."""
    score = 0
    url_lower = url.lower()
    
    # ─── Source type scoring ───
    if source_tag in ("gallery_link",):
        # Gallery <a> links point to popup images (often square/wide fragments)
        score -= 30  
    elif source_tag in ("og_image",):
        score += 10
    elif source_tag.startswith("img_tag") or source_tag.startswith("slider"):
        # Direct <img> tags are the main product photos (full door)
        if gallery_prefix and prefix and prefix == gallery_prefix:
            score += 40  # Main product photo matching model
    
    # ─── Prefix matching (model identity) ───
    if gallery_prefix and prefix:
        if prefix == gallery_prefix:
            score += 100  # Perfect match to gallery prefix
        elif gallery_prefix in prefix or prefix in gallery_prefix:
            score += 60   # Partial match
        elif len(prefix) > 3 and len(gallery_prefix) > 3:
            # Check character overlap 
            common = sum(1 for i in range(min(len(prefix), len(gallery_prefix))) if prefix[i] == gallery_prefix[i])
            if common >= min(len(prefix), len(gallery_prefix)) * 0.6:
                score += 40
        # Cross-sell penalty - strong for known cross-sell prefixes
        if prefix in ("dz1", "y23", "y12", "k061", "r061", "dz888", "06"):
            score -= 200  # Ledo smart lock cross-sells
        if prefix in ("trento", "forest", "ledo"):
            score -= 150  # Other product cross-sells
        if prefix and gallery_prefix and prefix != gallery_prefix and len(prefix) > 2:
            if len(gallery_prefix) > 2 and prefix not in gallery_prefix and gallery_prefix not in prefix:
                score -= 80
    
    # ─── URL slug matching ───
    if url_slug:
        slug_clean = re.sub(r'[^a-z0-9]', '', url_slug.lower())
        prefix_clean = re.sub(r'[^a-z0-9]', '', prefix or "")
        if slug_clean and prefix_clean:
            if slug_clean[:5] == prefix_clean[:5]:
                score += 80
            elif slug_clean[:5] in prefix_clean or prefix_clean[:5] in slug_clean:
                score += 50
    
    # ─── Aspect ratio (most important for full-door) ───
    if w and h:
        ratio = h / w if w > 0 else 0
        if ratio > 1.8:
            score += 100
        elif ratio > 1.5:
            score += 80
        elif ratio > 1.2:
            score += 30
        elif ratio > 0.8:
            score -= 60
        else:
            score -= 200
        
        if w >= 400 and h >= 800:
            score += 30
        elif w < 100 and h < 100:
            score -= 80
    
    # ─── File path patterns ───
    if "/assets/cache_image/" in url_lower:
        score -= 10
    else:
        score += 5
    
    # ─── Strong URL reject ───
    reject_pats = ["interier", "interior", "banner", "badge", "hit", "logo", 
                   "icon", "handle", "lock", "ruchk", "ручк", "замок", "петл",
                   "fragment", "detail", "texture", "sample", "profile"]
    if any(p in url_lower for p in reject_pats):
        score -= 300
    
    # ─── Fragment indicator in ASD numbering ───
    fn = url_lower.split("/")[-1] if "/" in url_lower else url_lower
    # ASD uses -14, -16 as fragment images
    if re.search(r'[-_](14|16|18)(\.[a-z]+)?$', fn):
        score -= 100
    
    return score

def get_url_slug(source_url):
    """Extract model identifier from page URL."""
    path = source_url.replace(".html", "").rstrip("/")
    slug = path.split("/")[-1] if path else ""
    return unquote(slug).lower()

def classify_reject_reason(url, w, h, prefix, gallery_prefix):
    """Classify why an image was rejected."""
    url_lower = url.lower()
    
    if any(x in url_lower for x in ["banner", "interier", "interior"]):
        return "banner"
    if any(x in url_lower for x in ["handle", "lock", "ruchk", "ручк", "замок"]):
        return "handle_lock"
    if prefix in ("dz1", "y23", "y12", "k061", "r061", "dz888", "06", "trento", "forest", "ledo"):
        return "cross_sell"
    if prefix and gallery_prefix and prefix != gallery_prefix and len(prefix) > 2 and len(gallery_prefix) > 2:
        return "different_model"
    
    if w and h:
        ratio = h / w if w > 0 else 0
        if ratio < 0.8:
            return "fragment_wide"
        if w * h < 50000:
            return "icon_small"
        if ratio < 1.2:
            return "fragment_squarish"
    
    if not w or not h:
        return "no_dimensions"
    
    return "full_door"

def visual_decision_str(w, h):
    if not w or not h:
        return "unknown"
    ratio = h / w if w > 0 else 0
    if ratio > 1.5 and w >= 300:
        return "full_door"
    if ratio > 1.2:
        return "full_door_short"
    if ratio > 0.8:
        return "squarish"
    return "fragment"

def process_asd_row(row):
    source_url = row.get("source_url", "")
    model_name = row.get("model_name", "")
    target_line = row.get("target_line", "")
    old_image_url = row.get("image_url", "")
    url_slug = get_url_slug(source_url)
    
    result = {
        "producer": "АСД", "target_line": target_line, "model_name": model_name,
        "source_url": source_url, "old_image_url": old_image_url,
        "selected_image_url": "", "decision": "not_found", "reason": "",
        "full_door_score": 0, "candidates_checked": 0,
        "rejected_fragment_count": 0, "rejected_interior_count": 0,
        "rejected_banner_count": 0, "selected_candidate_source": "",
        "status_before": row.get("status", ""), "comment_before": row.get("comment", ""),
        "patch_comment": ""
    }
    candidate_debugs = []
    
    if not source_url:
        result["reason"] = "no_source_url"
        result["patch_comment"] = "no_source_url"
        return result, candidate_debugs
    
    resp = fetch_page(source_url)
    if not resp:
        result["reason"] = "page_not_accessible"
        result["patch_comment"] = "page_not_accessible"
        result["selected_image_url"] = old_image_url
        return result, candidate_debugs
    
    soup = BeautifulSoup(resp.text, "lxml")
    gallery_prefix = extract_gallery_prefix(soup)
    candidates = extract_candidates(soup, gallery_prefix)
    
    result["candidates_checked"] = len(candidates)
    
    if not candidates:
        result["reason"] = "no_candidates_found"
        result["patch_comment"] = "no_full_door_candidate_found"
        result["selected_image_url"] = old_image_url
        return result, candidate_debugs
    
    # Score each candidate
    scored = []
    for url, tag, is_gallery, prefix in candidates:
        w, h, ok = get_image_info(url)
        s = score_candidate(url, w, h, tag, is_gallery, prefix, gallery_prefix, url_slug)
        vd = visual_decision_str(w, h)
        rej = classify_reject_reason(url, w, h, prefix, gallery_prefix)
        
        candidate_debugs.append({
            "model_name": model_name, "source_url": source_url,
            "candidate_url": url, "candidate_source": tag,
            "width": str(w or ""), "height": str(h or ""),
            "aspect_ratio": f"{(h/w):.2f}" if w and h else "",
            "score": str(s), "is_selected": "",
            "reject_reason": rej, "visual_decision": vd,
            "gallery_prefix": gallery_prefix or "",
            "prefix": prefix or ""
        })
        
        if w and h:
            scored.append((s, url, tag, w, h, is_gallery, prefix))
    
    # Count rejections
    rej_types = Counter(d["reject_reason"] for d in candidate_debugs)
    result["rejected_fragment_count"] = rej_types.get("fragment_wide",0) + rej_types.get("fragment_squarish",0)
    result["rejected_banner_count"] = rej_types.get("banner",0)
    
    if not scored:
        result["reason"] = "no_valid_dimensions"
        result["patch_comment"] = "no_full_door_candidate_found"
        result["selected_image_url"] = old_image_url
        return result, candidate_debugs
    
    # Sort by score descending
    scored.sort(key=lambda x: -x[0])
    
    best_score, best_url, best_source, best_w, best_h, best_gallery, best_prefix = scored[0]
    result["full_door_score"] = best_score
    best_ratio = best_h / best_w if best_w > 0 else 0
    
    # Check if best candidate actually matches this model
    model_matches = False
    if gallery_prefix and best_prefix:
        if best_prefix == gallery_prefix:
            model_matches = True
        elif len(best_prefix) > 3 and len(gallery_prefix) > 3:
            common_prefix = os.path.commonprefix([best_prefix, gallery_prefix.lower()])
            if len(common_prefix) >= 3:
                model_matches = True
    if not model_matches and url_slug:
        slug_base = re.sub(r'[^a-z0-9]', '', url_slug)[:10]
        prefix_base = re.sub(r'[^a-z0-9]', '', best_prefix or "")[:10]
        if slug_base and prefix_base and (slug_base[:5] == prefix_base[:5] or slug_base in best_prefix or best_prefix in slug_base):
            model_matches = True
    if not model_matches and best_gallery:
        model_matches = True  # Gallery images are model-specific
    
    if not model_matches and best_score > 0:
        highest_matching = -999
        for s, u, tag, w, h, is_gal, pre in scored:
            if is_gal or (gallery_prefix and pre == gallery_prefix):
                highest_matching = max(highest_matching, s)
            if url_slug:
                slug_base = re.sub(r'[^a-z0-9]', '', url_slug)[:10]
                pre_base = re.sub(r'[^a-z0-9]', '', pre or "")[:10]
                if slug_base and pre_base and slug_base[:4] in pre_base:
                    highest_matching = max(highest_matching, s)
        
        if highest_matching < 50 and highest_matching > -999:
            # No model-specific images found
            best_score = -100
            result["reason"] = "no_model_specific_images_on_page"
        elif best_score < highest_matching - 50:
            # There's a better matching image
            best_score = highest_matching
            # Find the best matching candidate
            for s, u, tag, w, h, is_gal, pre in scored:
                if s == highest_matching:
                    best_url, best_source, best_w, best_h = u, tag, w, h
                    best_ratio = best_h / best_w if best_w > 0 else 0
                    break
    
    # Make decision
    if best_score >= 100 and best_ratio >= 1.5:
        if old_image_url and old_image_url == best_url:
            result["decision"] = "keep"
            result["reason"] = "old_image_already_full_door"
            result["patch_comment"] = "asd_photo_already_full_door"
            result["selected_image_url"] = old_image_url
        else:
            result["decision"] = "replace"
            result["reason"] = "full_door_product_image_selected"
            result["patch_comment"] = "asd_full_door_photo_selected"
            result["selected_image_url"] = best_url
            result["selected_candidate_source"] = best_source
    elif best_score >= 60 and best_ratio >= 1.2:
        result["decision"] = "manual_check"
        result["reason"] = "ambiguous_candidates_manual_check"
        result["patch_comment"] = "asd_photo_manual_check_needed"
        result["selected_image_url"] = best_url
        result["selected_candidate_source"] = best_source
    else:
        result["reason"] = "no_full_door_candidate_found"
        result["patch_comment"] = "no_full_door_candidate_found"
        result["selected_image_url"] = old_image_url if old_image_url else ""
    
    # Mark selected
    for d in candidate_debugs:
        if d["candidate_url"] == result["selected_image_url"]:
            d["is_selected"] = "YES"
    
    return result, candidate_debugs

def generate_html_v2(results, all_debugs):
    sections = defaultdict(list)
    for r in results:
        sections[r["decision"]].append(r)
    
    total = len(results)
    replace_c = len(sections.get("replace", []))
    keep_c = len(sections.get("keep", []))
    not_found_c = len(sections.get("not_found", []))
    manual_c = len(sections.get("manual_check", []))
    
    debug_lookup = defaultdict(list)
    for d in all_debugs:
        debug_lookup[(d["model_name"], d["source_url"])].append(d)
    
    html = f"""<!DOCTYPE html>
<html lang="ru">
<head><meta charset="UTF-8">
<title>ASD Photo Patch v2 Review</title>
<style>
body {{ font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }}
h1 {{ color: #333; }}
.summary {{ background: #fff; padding: 20px; border-radius: 8px; margin-bottom: 20px; }}
.grid {{ display: flex; flex-wrap: wrap; gap: 15px; }}
.card {{ background: #fff; border: 1px solid #ddd; border-radius: 8px; padding: 12px; width: 360px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
.card img {{ max-width: 100%; height: auto; margin: 5px 0; border: 1px solid #eee; border-radius: 4px; }}
.card .selected-img {{ border: 3px solid #4CAF50; }}
.label {{ font-weight: bold; color: #666; font-size: 11px; text-transform: uppercase; margin-top: 8px; }}
.value {{ margin-bottom: 4px; font-size: 12px; }}
.decision-replace {{ border-left: 4px solid #4CAF50; }}
.decision-keep {{ border-left: 4px solid #2196F3; }}
.decision-not_found {{ border-left: 4px solid #f44336; }}
.decision-manual_check {{ border-left: 4px solid #FF9800; }}
.badge {{ display: inline-block; padding: 2px 8px; border-radius: 3px; font-size: 11px; color: #fff; }}
.badge-replace {{ background: #4CAF50; }}
.badge-keep {{ background: #2196F3; }}
.badge-not_found {{ background: #f44336; }}
.badge-manual_check {{ background: #FF9800; }}
.rejected-list {{ font-size: 11px; color: #999; margin-top: 5px; }}
.rejected-item {{ padding: 1px 0; border-bottom: 1px solid #f0f0f0; }}
</style></head><body>
<h1>ASD Photo Patch v2 - Full Door Review</h1>
<div class="summary">
<p>Total ASD rows: <strong>{total}</strong></p>
<p>Replace: <strong>{replace_c}</strong> | Keep: <strong>{keep_c}</strong> | Not found: <strong>{not_found_c}</strong> | Manual check: <strong>{manual_c}</strong></p>
</div>
"""
    
    for section_name, label, badge_cls in [
        ("replace", "Replace", "replace"), ("keep", "Keep", "keep"),
        ("not_found", "Not Found", "not_found"), ("manual_check", "Manual Check", "manual_check"),
    ]:
        items = sections.get(section_name, [])
        if not items:
            continue
        html += f'<div class="section-title"><span class="badge badge-{badge_cls}">{label}</span> {len(items)} rows</div><div class="grid">\n'
        for r in items:
            new_img = r.get("selected_image_url", "")
            old_img = r.get("old_image_url", "")
            old_thumb = f'<img src="{old_img}" onerror="this.style.display=\'none\'">' if old_img else '<div style="color:#999;font-style:italic;">(empty)</div>'
            new_thumb = f'<img src="{new_img}" onerror="this.style.display=\'none\'" class="selected-img">' if new_img else '<div style="color:#999;font-style:italic;">(empty)</div>'
            
            key = (r["model_name"], r["source_url"])
            debugs_raw = debug_lookup.get(key, [])
            rejected = [d for d in debugs_raw if d.get("is_selected") != "YES" and d.get("reject_reason") not in ("", "unknown", "full_door")]
            rejected_html = ""
            if rejected:
                rejected_html = '<div class="label">Rejected</div><div class="rejected-list">'
                for d in rejected[:5]:
                    fn = d["candidate_url"].split("/")[-1][:30]
                    rejected_html += f'<div class="rejected-item">{d["reject_reason"]}: {fn} (ar={d["aspect_ratio"]})</div>'
                rejected_html += '</div>'
            
            html += f'''
<div class="card decision-{section_name}">
<div class="label">Model (score={r.get("full_door_score","")})</div>
<div class="value"><b>{r["model_name"][:50]}</b></div>
<div class="label">Source</div>
<div class="value" style="font-size:11px;"><a href="{r["source_url"]}" target="_blank">{r["source_url"][:40]}...</a></div>
<div class="label">SELECTED {r.get("decision","").upper()}</div>
{new_thumb}
<div class="label">OLD</div>
{old_thumb}
{rejected_html}
</div>'''
        html += '</div>\n'
    
    html += "</body></html>"
    return html

def process_with_timeout(row):
    """Run process_asd_row with a timeout. Returns (result_dict, debugs_list)."""
    with ThreadPoolExecutor(max_workers=1) as pool:
        fut = pool.submit(process_asd_row, row)
        try:
            result, debugs = fut.result(timeout=ROW_TIMEOUT)
            return result, debugs
        except TimeoutError:
            return None, "timeout"
        except Exception as e:
            return None, str(e)

def main():
    global _session
    print("=== ASD Photo Patch v2 ===")
    rows = load_asd_rows()
    print(f"ASD rows: {len(rows)}")
    print(f"Cache: {len(_image_dim_cache)} entries")
    
    all_results = []
    all_debugs = []
    processed = 0
    
    # Load previous results for resume
    tmp_csv = os.path.join(OUT_DIR, "asd_photo_patch_v2_tmp.csv")
    if os.path.exists(tmp_csv):
        with open(tmp_csv, encoding="utf-8-sig") as f:
            done = list(csv.DictReader(f, delimiter=";"))
        done_keys = set((r["model_name"], r["source_url"]) for r in done)
        rows = [r for r in rows if (r["model_name"], r["source_url"]) not in done_keys]
        all_results = done
        processed = len(done)
        print(f"Resumed: {processed} already done, {len(rows)} remaining")
        
        # Load previous debugs
        tmp_debug = os.path.join(OUT_DIR, "asd_photo_candidates_debug_tmp.csv")
        if os.path.exists(tmp_debug):
            with open(tmp_debug, encoding="utf-8-sig") as f:
                all_debugs = list(csv.DictReader(f, delimiter=";"))
    
    total = len(rows)
    for idx, row in enumerate(rows):
        processed += 1
        print(f"  [{processed:3d}/{processed+total-idx-1:3d}] {row.get('model_name','')[:30]:30} ...", end="")
        
        result_obj, debugs_or_err = process_with_timeout(row)
        if debugs_or_err and isinstance(debugs_or_err, str):
            print(f" ERROR: {debugs_or_err}")
            result_obj = {
                "producer": "АСД", "target_line": row.get("target_line",""),
                "model_name": row.get("model_name",""), "source_url": row.get("source_url",""),
                "old_image_url": row.get("image_url",""), "selected_image_url": row.get("image_url",""),
                "decision": "not_found", "reason": f"error: {debugs_or_err}", "full_door_score": 0,
                "candidates_checked": 0, "rejected_fragment_count": 0,
                "rejected_interior_count": 0, "rejected_banner_count": 0,
                "selected_candidate_source": "", "status_before": row.get("status",""),
                "comment_before": row.get("comment",""), "patch_comment": "processing_error"
            }
        else:
            all_debugs.extend(debugs_or_err)
            s = f"score={result_obj['full_door_score']:4d}" if result_obj.get('full_door_score') else ""
            sel = result_obj.get('selected_image_url','')[:40] if result_obj.get('selected_image_url') else "none"
            print(f" {result_obj['decision']:10} {s:10} {sel}")
        
        all_results.append(result_obj)
        
        # Save every 10 rows
        if len(all_results) % 10 == 0 or idx == len(rows) - 1:
            fieldnames = [
                "producer", "target_line", "model_name", "source_url",
                "old_image_url", "selected_image_url", "decision", "reason",
                "full_door_score", "candidates_checked",
                "rejected_fragment_count", "rejected_interior_count",
                "rejected_banner_count", "selected_candidate_source",
                "status_before", "comment_before", "patch_comment"
            ]
            with open(tmp_csv, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";")
                w.writeheader()
                w.writerows(all_results)
            
            debug_fields = [
                "model_name", "source_url", "candidate_url", "candidate_source",
                "width", "height", "aspect_ratio", "score", "is_selected",
                "reject_reason", "visual_decision", "gallery_prefix", "prefix"
            ]
            tmp_debug = os.path.join(OUT_DIR, "asd_photo_candidates_debug_tmp.csv")
            with open(tmp_debug, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.DictWriter(f, fieldnames=debug_fields, delimiter=";")
                w.writeheader()
                w.writerows(all_debugs)
    
    # Close session
    if _session:
        _session.close()
        _session = None
    
    assert len(all_results) == 199, f"Expected 199, got {len(all_results)}"
    
    # Final outputs
    fieldnames = [
        "producer", "target_line", "model_name", "source_url",
        "old_image_url", "selected_image_url", "decision", "reason",
        "full_door_score", "candidates_checked",
        "rejected_fragment_count", "rejected_interior_count",
        "rejected_banner_count", "selected_candidate_source",
        "status_before", "comment_before", "patch_comment"
    ]
    csv_path = os.path.join(OUT_DIR, "asd_photo_patch_v2.csv")
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";")
        w.writeheader()
        w.writerows(all_results)
    print(f"\nCSV: {csv_path}")
    
    try:
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "ASD Photo Patch v2"
        for ci, h in enumerate(fieldnames, 1):
            ws.cell(row=1, column=ci, value=h)
        for ri, r in enumerate(all_results, 2):
            for ci, h in enumerate(fieldnames, 1):
                ws.cell(row=ri, column=ci, value=r.get(h, ""))
        xlsx_path = os.path.join(OUT_DIR, "asd_photo_patch_v2.xlsx")
        wb.save(xlsx_path)
        print(f"XLSX: {xlsx_path}")
    except ImportError:
        pass
    
    debug_fields = [
        "model_name", "source_url", "candidate_url", "candidate_source",
        "width", "height", "aspect_ratio", "score", "is_selected",
        "reject_reason", "visual_decision", "gallery_prefix", "prefix"
    ]
    debug_path = os.path.join(OUT_DIR, "asd_photo_candidates_debug.csv")
    with open(debug_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=debug_fields, delimiter=";")
        w.writeheader()
        w.writerows(all_debugs)
    print(f"Debug: {debug_path} ({len(all_debugs)} candidates)")
    
    html = generate_html_v2(all_results, all_debugs)
    html_path = os.path.join(OUT_DIR, "asd_photo_patch_v2_review.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"HTML: {html_path}")
    
    decisions = Counter(r["decision"] for r in all_results)
    total_candidates = sum(int(r.get("candidates_checked",0) or 0) for r in all_results)
    total_fragment = sum(int(r.get("rejected_fragment_count",0) or 0) for r in all_results)
    total_banner = sum(int(r.get("rejected_banner_count",0) or 0) for r in all_results)
    
    print(f"\n=== Statistics ===")
    print(f"Total: {len(all_results)}")
    print(f"Replace: {decisions.get('replace',0)}")
    print(f"Keep: {decisions.get('keep',0)}")
    print(f"Not found: {decisions.get('not_found',0)}")
    print(f"Manual check: {decisions.get('manual_check',0)}")
    print(f"Candidates: {total_candidates}")
    print(f"Fragments rejected: {total_fragment}")
    print(f"Banners rejected: {total_banner}")
    
    report = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "total_asd_rows": len(all_results),
        "replace": decisions.get("replace",0),
        "keep": decisions.get("keep",0),
        "not_found": decisions.get("not_found",0),
        "manual_check": decisions.get("manual_check",0),
        "candidates_checked": total_candidates,
        "fragments_rejected": total_fragment,
        "banners_rejected": total_banner
    }
    report_path = os.path.join(OUT_DIR, "asd_photo_patch_v2_report.json")
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    print(f"Report: {report_path}")
    
    return all_results

if __name__ == "__main__":
    main()
