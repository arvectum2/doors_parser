"""
Stage A - Step 1: Collect all candidates for all 199 ASD rows.
Scrapes source pages, gathers historical candidates, downloads images.
No VLM classification — pure data collection.
"""
import csv, json, os, re, hashlib, time, shutil
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import defaultdict
from datetime import datetime
import requests
from bs4 import BeautifulSoup
import urllib3
urllib3.disable_warnings()

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
V56_DIR = os.path.join(BASE, "output", "delivery_v5_6")
CACHE_DIR = os.path.join(BASE, "output", "manual_photo_review_v5_7", "cache", "asd")
REPORTS_DIR = os.path.join(BASE, "output", "manual_photo_review_v5_7", "reports")
os.makedirs(CACHE_DIR, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)

session = requests.Session()
session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
session.verify = False
session.timeout = 15

# ========================================================================
# 1. Load catalog and historical QA/candidate data
# ========================================================================
print("Loading v5.6 catalog...")
with open(os.path.join(V56_DIR, "doors_catalog.csv"), encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
asd_rows = [r for r in rows if r['producer'] == 'АСД']
print(f"ASD rows: {len(asd_rows)}")

# Load v5.2 QA
print("Loading v5.2 QA...")
with open(os.path.join(BASE, "output", "delivery_v5_2", "asd_final_photo_source_validation.json"), encoding='utf-8') as f:
    qa_v52 = {r['model_name']: r for r in json.load(f)}

# Load image_candidates from v2 (latest candidate scrape)
print("Loading image candidates...")
candidates_by_model = defaultdict(list)
for csv_path in [
    os.path.join(BASE, "output", "image_candidates_v2", "asd_candidates.csv"),
    os.path.join(BASE, "output", "image_candidates", "asd_candidates.csv"),
]:
    if os.path.exists(csv_path):
        with open(csv_path, encoding='utf-8-sig') as f:
            for r in csv.DictReader(f, delimiter=';'):
                candidates_by_model[r['model_name']].append(r)

# Load delivery_v5_4 catalog (second pass with many fixes)
v54_rows = {}
v54_path = os.path.join(BASE, "output", "delivery_v5_4", "doors_catalog.csv")
if os.path.exists(v54_path):
    with open(v54_path, encoding='utf-8-sig') as f:
        for r in csv.DictReader(f, delimiter=';'):
            v54_rows[r['model_name']] = r

# Load delivery_v5_5 catalog (QA patch baseline)
v55_rows = {}
v55_path = os.path.join(BASE, "output", "delivery_v5_5", "doors_catalog.csv")
if os.path.exists(v55_path):
    with open(v55_path, encoding='utf-8-sig') as f:
        for r in csv.DictReader(f, delimiter=';'):
            v55_rows[r['model_name']] = r

# ========================================================================
# 2. Scrape ASD source pages for gallery structure
# ========================================================================
def scrape_gallery(source_url, model_name):
    """Extract gallery images from ASD source page."""
    result = {
        'model_name': model_name,
        'source_url': source_url,
        'http_status': 0,
        'slides': [],
        'error': None
    }
    try:
        resp = session.get(source_url, timeout=20)
        result['http_status'] = resp.status_code
        if resp.status_code != 200:
            result['error'] = f'HTTP {resp.status_code}'
            return result
        soup = BeautifulSoup(resp.text, 'html.parser')
        
        # ASD-specific gallery detection
        gallery_selectors = [
            'div.product-gallery', 'div.gallery', 'div.product__slider',
            'div.product-slider', 'div.catalog-card__slider', 'div.img-wrapper',
            'div.product-photo', 'div.product__photo', 'div.item-photo',
            'div.splide__list', 'div.owl-carousel', 'div.swiper-wrapper',
            'ul.product-gallery', 'ul.gallery',
        ]
        
        seen_urls = set()
        slide_idx = 0
        
        # Try known gallery containers first
        for sel in gallery_selectors:
            container = soup.select_one(sel)
            if container:
                imgs = container.find_all('img')
                for img in imgs:
                    src = (img.get('data-src') or img.get('src') or '').strip()
                    if not src or not src.startswith('http'):
                        continue
                    if src in seen_urls:
                        continue
                    seen_urls.add(src)
                    slide_idx += 1
                    result['slides'].append({
                        'gallery_index': slide_idx,
                        'src': src,
                        'href': (img.get('data-href') or img.get('href') or ''),
                        'alt': (img.get('alt') or '').strip(),
                        'title': (img.get('title') or '').strip(),
                        'width': img.get('width', ''),
                        'height': img.get('height', ''),
                        'selector': sel,
                        'is_video': bool(img.get('data-video') or 'youtube' in src.lower() or 'vimeo' in src.lower()),
                        'container_type': 'gallery',
                    })
                if result['slides']:
                    break  # Found gallery slides
        
        # Also try global img search for images with gallery-like patterns
        if not result['slides']:
            all_imgs = soup.find_all('img')
            for img in all_imgs:
                src = (img.get('data-src') or img.get('src') or '').strip()
                if not src or not src.startswith('http'):
                    continue
                # Filter: only images in main content area
                parent_classes = ' '.join(img.parent.get('class', []) if img.parent else [])
                parent_id = img.parent.get('id', '') if img.parent else ''
                content_signals = ['content', 'product', 'main', 'card', 'item', 'gallery',
                                   'slider', 'photo', 'detail', 'body']
                if not any(s in parent_classes.lower() or s in parent_id.lower() for s in content_signals):
                    # Still include if the image seems product-related
                    alt_lower = (img.get('alt') or '').lower()
                    src_lower = src.lower()
                    if not any(s in alt_lower or s in src_lower for s in ['door', 'двер', 'photo', 'фото',
                                                                          'product', 'товар', 'image', 'изображ']):
                        continue
                if src in seen_urls:
                    continue
                seen_urls.add(src)
                slide_idx += 1
                result['slides'].append({
                    'gallery_index': slide_idx,
                    'src': src,
                    'href': '',
                    'alt': (img.get('alt') or '').strip(),
                    'title': (img.get('title') or '').strip(),
                    'width': img.get('width', ''),
                    'height': img.get('height', ''),
                    'selector': 'global_img',
                    'is_video': 'youtube' in src.lower() or 'vimeo' in src.lower(),
                    'container_type': 'page',
                })
        
        # Detect video from iframes
        iframes = soup.find_all('iframe')
        for iframe in iframes:
            src = (iframe.get('data-src') or iframe.get('src') or '').strip()
            if src and ('youtube' in src or 'vimeo' in src):
                slide_idx += 1
                result['slides'].append({
                    'gallery_index': slide_idx,
                    'src': src,
                    'href': '',
                    'alt': '',
                    'title': '',
                    'width': iframe.get('width', ''),
                    'height': iframe.get('height', ''),
                    'selector': 'iframe',
                    'is_video': True,
                    'container_type': 'video',
                })
        
    except Exception as e:
        result['error'] = str(e)
    return result

# ========================================================================
# 3. Collect and download all candidates
# ========================================================================
def download_image(url, dest_path):
    """Download image to local path. Returns metadata dict."""
    meta = {
        'url': url,
        'downloaded': False,
        'sha256': None,
        'width': None,
        'height': None,
        'content_type': None,
        'http_status': 0,
        'error': None,
        'file_size': 0,
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
        # Download
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
        # Get dimensions
        try:
            from PIL import Image
            with Image.open(dest_path) as img:
                meta['width'], meta['height'] = img.size
        except Exception:
            pass  # Dimensions unknown
    except Exception as e:
        meta['error'] = str(e)
    return meta

def safe_model_dir(name):
    """Create safe directory name from model name."""
    safe = re.sub(r'[<>:"/\\|?*]', '_', name)
    safe = re.sub(r'\s+', '_', safe)
    safe = safe.strip('._') or 'unnamed'
    return safe[:120]

# ========================================================================
# MAIN: Collect all candidates for all ASD rows
# ========================================================================
print("\n=== Scraping gallery pages and collecting candidates ===")
collection = {}

# Use thread pool for gallery scraping
with ThreadPoolExecutor(max_workers=5) as ex:
    fut_map = {}
    for r in asd_rows:
        name = r['model_name']
        src_url = r.get('source_url', '').strip()
        if src_url and src_url.startswith('http'):
            fut = ex.submit(scrape_gallery, src_url, name)
            fut_map[fut] = (name, src_url)
    
    for fut in as_completed(fut_map):
        name, src_url = fut_map[fut]
        try:
            result = fut.result()
            collection[name] = result
        except Exception as e:
            collection[name] = {'model_name': name, 'source_url': src_url, 'error': str(e), 'slides': []}

# Fill missing (for rows where scraping failed)
for r in asd_rows:
    name = r['model_name']
    if name not in collection:
        collection[name] = {
            'model_name': name,
            'source_url': r.get('source_url', ''),
            'http_status': 0,
            'slides': [],
            'error': 'not_scraped'
        }

print(f"Scraped {len(collection)} models")

# ========================================================================
# Build candidate list for each model
# ========================================================================
print("\n=== Building candidate database ===")
all_candidates = []

for r in asd_rows:
    name = r['model_name']
    safe = safe_model_dir(name)
    model_dir = os.path.join(CACHE_DIR, safe)
    
    # Collect all unique candidate URLs
    seen_urls = set()
    model_candidates = []
    
    # 1. Current front + inside from v5.6
    for col, source_type in [('image_url_front', 'v56_current_front'), ('image_url_inside', 'v56_current_inside')]:
        url = r.get(col, '').strip()
        if url and url not in seen_urls:
            seen_urls.add(url)
            model_candidates.append({
                'url': url, 'source_type': source_type, 'gallery_index': 0,
                'original_filename': url.rsplit('/', 1)[-1] if '/' in url else url,
            })
    
    # 2. Historical v5.2 QA candidates
    qa = qa_v52.get(name, {})
    for col, source_type in [('image_url_front', 'v52_qa_front'), ('image_url_inside', 'v52_qa_inside')]:
        url = qa.get(col, '').strip()
        if url and url not in seen_urls:
            seen_urls.add(url)
            model_candidates.append({
                'url': url, 'source_type': source_type, 'gallery_index': 0,
                'original_filename': url.rsplit('/', 1)[-1] if '/' in url else url,
            })
    
    # 3. v5.4 / v5.5 historical
    for label, hist in [('v54', v54_rows), ('v55', v55_rows)]:
        h = hist.get(name, {})
        for col in ['image_url_front', 'image_url_inside']:
            url = h.get(col, '').strip()
            if url and url not in seen_urls:
                seen_urls.add(url)
                model_candidates.append({
                    'url': url, 'source_type': f'{label}_{col}', 'gallery_index': 0,
                    'original_filename': url.rsplit('/', 1)[-1] if '/' in url else url,
                })
    
    # 4. Image candidates from parser
    for cand in candidates_by_model.get(name, []):
        url = cand.get('candidate_image_url', '').strip()
        if url and url not in seen_urls:
            seen_urls.add(url)
            model_candidates.append({
                'url': url,
                'source_type': f"parser_gallery_idx_{cand.get('candidate_index','?')}",
                'gallery_index': int(cand.get('candidate_index', 0)) if cand.get('candidate_index','').isdigit() else 0,
                'original_filename': url.rsplit('/', 1)[-1] if '/' in url else url,
                'parser_info': {
                    'candidate_text': cand.get('candidate_text', ''),
                    'candidate_alt': cand.get('candidate_alt', ''),
                    'candidate_title': cand.get('candidate_title', ''),
                    'candidate_width': cand.get('candidate_width', ''),
                    'candidate_height': cand.get('candidate_height', ''),
                    'parser_decision': cand.get('parser_decision', ''),
                    'reason': cand.get('reason', ''),
                }
            })
    
    # 5. Gallery slides from current scrape
    gal = collection.get(name, {})
    for slide in gal.get('slides', []):
        url = slide.get('src', '')
        if url and url not in seen_urls:
            seen_urls.add(url)
            model_candidates.append({
                'url': url,
                'source_type': f"gallery_slide_{slide.get('gallery_index','?')}",
                'gallery_index': slide.get('gallery_index', 0),
                'original_filename': url.rsplit('/', 1)[-1] if '/' in url else url,
                'gallery_info': {
                    'alt': slide.get('alt', ''),
                    'title': slide.get('title', ''),
                    'width': slide.get('width', ''),
                    'height': slide.get('height', ''),
                    'is_video': slide.get('is_video', False),
                    'container_type': slide.get('container_type', ''),
                }
            })
    
    all_candidates.append({
        'model_name': name,
        'safe_dir': safe,
        'model_dir': model_dir,
        'source_url': r.get('source_url', ''),
        'excel_row': rows.index(r) + 2,
        'current_front': r.get('image_url_front', ''),
        'current_inside': r.get('image_url_inside', ''),
        'comment': r.get('comment', ''),
        'candidates': model_candidates,
        'gallery_scrape': {
            'http_status': gal.get('http_status', 0),
            'slide_count': len(gal.get('slides', [])),
            'error': gal.get('error'),
        }
    })

# Save candidate DB
db_path = os.path.join(REPORTS_DIR, "asd_candidate_database.json")
with open(db_path, 'w', encoding='utf-8') as f:
    json.dump(all_candidates, f, ensure_ascii=False, indent=2)
print(f"Candidate DB: {db_path}")

# ========================================================================
# Download all images to local cache
# ========================================================================
print("\n=== Downloading images to local cache ===")

total_candidates = sum(len(m['candidates']) for m in all_candidates)
print(f"Total candidates to download: {total_candidates}")

dl_results = []
download_tasks = []

for model in all_candidates:
    safe = model['safe_dir']
    for cand in model['candidates']:
        url = cand['url']
        ext = (url.rsplit('.', 1)[-1] if '.' in url else 'jpg').split('?')[0][:10]
        filename = f"idx{cand['gallery_index']:03d}_{hash(url)%10000:04x}.{ext}"
        dest = os.path.join(CACHE_DIR, safe, filename)
        download_tasks.append((url, dest, cand['source_type'], model['model_name']))

# Download in parallel
with ThreadPoolExecutor(max_workers=10) as ex:
    fut_map = {ex.submit(download_image, url, dest): (url, dest, stype, mname)
               for url, dest, stype, mname in download_tasks}
    completed = 0
    for fut in as_completed(fut_map):
        url, dest, stype, mname = fut_map[fut]
        try:
            meta = fut.result()
            meta['destination'] = dest
            meta['source_type'] = stype
            meta['model_name'] = mname
            dl_results.append(meta)
        except Exception as e:
            dl_results.append({
                'url': url, 'downloaded': False, 'error': str(e),
                'destination': dest, 'source_type': stype, 'model_name': mname
            })
        completed += 1
        if completed % 50 == 0:
            print(f"  Downloaded {completed}/{total_candidates}")

# Save download results
dl_path = os.path.join(REPORTS_DIR, "asd_download_results.json")
with open(dl_path, 'w', encoding='utf-8') as f:
    json.dump(dl_results, f, ensure_ascii=False, indent=2)

downloaded_ok = sum(1 for d in dl_results if d.get('downloaded'))
downloaded_fail = sum(1 for d in dl_results if not d.get('downloaded'))
print(f"Downloaded: {downloaded_ok}, Failed: {downloaded_fail}")

# ========================================================================
# Summary
# ========================================================================
print(f"\n{'='*60}")
print(f"STAGE A - STEP 1 COMPLETE")
print(f"{'='*60}")
print(f"  Models processed: {len(all_candidates)}")
print(f"  Total candidates: {total_candidates}")
print(f"  Downloaded OK:    {downloaded_ok}")
print(f"  Failed:           {downloaded_fail}")
print(f"  Candidate DB:     {db_path}")
print(f"  Download results: {dl_path}")
print(f"  Cache dir:        {CACHE_DIR}")
