"""Fix and improve ASD photo patch script"""
import csv, json, os, re, requests, time
from bs4 import BeautifulSoup
from urllib.parse import urljoin, unquote
from collections import defaultdict, Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

BASE_DIR = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
OUT_DIR = os.path.join(BASE_DIR, "output", "asd_photo_patch")
os.makedirs(OUT_DIR, exist_ok=True)

SOURCE_CSV = os.path.join(OUT_DIR, "asd_rows_v14_source.csv")
MAX_WORKERS = 10
REQUEST_TIMEOUT = 10
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"

REJECT_PATTERNS = re.compile(
    r"(interier|interior|интерьер|roomset|room\b|banner|badge|"
    r"hit\b|sale\b|logo|icon|sprite|placeholder|no_image|"
    r"handle|lock\b|zamok|ручка|замок|петля|texture|sample|color\b|"
    r"ral\b|ncs\b|filter|menu|star|capcha|arrow|svg|"
    r"thumb|small|mini|preview|loader|pixel)", re.I
)

# ASD root URL for image resolution
ASD_ROOT = "https://as-doors.ru"

def load_asd_rows():
    with open(SOURCE_CSV, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f, delimiter=";"))

def fetch_page(url):
    try:
        r = requests.get(url, timeout=REQUEST_TIMEOUT, headers={"User-Agent": USER_AGENT})
        return r if r.status_code == 200 else None
    except requests.Timeout:
        return None
    except Exception:
        return None

def make_abs_url(src):
    """Convert relative image path to absolute URL using ASD root."""
    if src.startswith("http"):
        return src
    if src.startswith("/"):
        return ASD_ROOT + src
    # Relative path - should start from root
    if src.startswith("img/") or src.startswith("images/") or src.startswith("img202"):
        return ASD_ROOT + "/" + src
    return urljoin(ASD_ROOT + "/", src)

def verify_image_url(url):
    """Check if URL is accessible and returns image content type."""
    try:
        # Use GET with stream=True to avoid downloading full content
        r = requests.get(url, timeout=5, headers={"User-Agent": USER_AGENT}, stream=True)
        if r.status_code == 200:
            ct = r.headers.get("content-type", "")
            if ct.startswith("image/"):
                r.close()
                return True, ct
        r.close()
        return False, str(r.status_code)
    except Exception as e:
        return False, str(e)

def extract_model_name_from_url(url):
    """Extract the model identifier from the source URL."""
    path = url.replace(".html", "").rstrip("/")
    parts = path.split("/")
    # Last meaningful part is the model
    last = parts[-1] if parts else ""
    # Decode URL encoding
    last = unquote(last)
    return last.lower()

def extract_product_images_from_page(soup, page_url):
    """
    Extract product images from ASD product page.
    Returns list of (full_url, source_tag, priority).
    """
    candidates = []
    seen = set()
    
    def add(url, tag, priority):
        if not url:
            return
        full = make_abs_url(url)
        full = full.split("?")[0]  # Remove query params
        if full in seen:
            return
        seen.add(full)
        candidates.append((full, tag, priority))
    
    # 1. All <img> tags with meaningful src
    for img in soup.find_all("img"):
        src = img.get("src", "") or img.get("data-src", "") or ""
        alt = img.get("alt", "") or ""
        
        if not src or len(src) < 10:
            continue
        
        src_lower = src.lower()
        
        # Skip icons, stars, captcha, SVG, trackers
        if any(x in src_lower for x in ["/star/", "/ico/", "capcha", ".svg", "yandex", "mc.yandex"]):
            continue
        
        # Skip banner, interier, and other non-product images
        if any(x in src_lower for x in ["/banner", "banner/", "interier", "badge", "hit", "sale"]):
            continue
        
        # Classify
        if "/assets/cache_image/" in src_lower:
            # Cached image - medium priority
            if "banner" not in src_lower:
                add(src, "cache_img", priority=30)
        elif any(x in src_lower for x in ["/img/", "/img", "img/", "img202"]):
            # Direct product image - high priority
            add(src, "direct_img", priority=10)
        elif src_lower.startswith("images/") or "/images/" in src_lower:
            # Gallery image path
            add(src, "gallery_img", priority=15)
        else:
            add(src, "other_img", priority=50)
    
    # 2. Gallery <a> links (these are direct links to full-size images)
    for a_tag in soup.select(".gallery a, a[data-gallery], a[rel*='gallery'], a[class*='gallery'], a[class*='fancy']"):
        href = a_tag.get("href", "")
        if href and not href.startswith("#") and not href.startswith("javascript"):
            href_lower = href.lower()
            if any(x in href_lower for x in ["/banner", "banner/", "interier", "badge", "hit", "sale"]):
                continue
            add(href, "gallery_link", priority=5)
    
    # 3. og:image
    og = soup.select_one('meta[property="og:image"]')
    if og:
        add(og.get("content", ""), "og_image", priority=8)
    
    # 4. JSON-LD structured data
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string)
            if isinstance(data, dict):
                img = data.get("image", "")
                if img:
                    add(img if isinstance(img, str) else (img[0] if isinstance(img, list) else ""), "jsonld", priority=6)
        except:
            pass
    
    return candidates

def determine_model_folder(soup, page_url, model_name):
    """
    Determine the image folder name for the current model.
    Strategy: find the most frequent image folder in img/YYYY/<folder>/ paths.
    """
    folder_counts = Counter()
    for img in soup.find_all("img"):
        src = img.get("src", "") or ""
        # Match img/YYYY/folder/... pattern
        m = re.search(r"/(?:img|images)/(\d{4})/([^/]+)/", src)
        if m:
            folder = m.group(2).lower()
            folder_counts[folder] += 1
    
    if not folder_counts:
        return None
    
    # Return most frequent folder
    return folder_counts.most_common(1)[0][0]

def filter_relevant_candidates(candidates, model_folder_name, model_name, page_url=""):
    """
    Filter candidates to keep only the most relevant ones.
    Priority: images matching the model folder, then the model name.
    """
    if not candidates:
        return []
    
    # Normalize model name for matching
    model_lower = model_name.lower()
    # Extract the core door name (remove "Дверь ", "Стальная дверь ", etc.)
    core_model = re.sub(r'^(дверь|стальная дверь|дверь\s+)', '', model_lower).strip()
    
    scored = []
    for url, tag, priority in candidates:
        url_lower = url.lower()
        score = 0
        
        # Bonus for matching model folder in URL
        if model_folder_name and model_folder_name in url_lower:
            score += 100
        
        # Bonus for matching model name in URL
        if core_model and core_model[:10] in url_lower:
            score += 80
        
        # Direct paths preferred over cache
        if "/assets/cache_image/" not in url_lower:
            score += 20
        
        # Bonus for gallery link source
        if tag == "gallery_link":
            score += 30
        elif tag in ("direct_img", "og_image", "jsonld"):
            score += 15
        
        # Penalty for non-door paths
        if any(p in url_lower for p in ["/ico/", "/star/", "capcha", ".svg"]):
            score = -100
        
        # Strong reject for banner/interier
        if any(p in url_lower for p in ["/banner", "banner/", "interier", "badge"]):
            score = -1000
        
        # Bonus for img/ path from root
        if url_lower.startswith("https://as-doors.ru/img/"):
            score += 10
        
        scored.append((score, url, tag, priority))
    
    # Sort by score descending, then priority ascending
    scored.sort(key=lambda x: (-x[0], x[3]))
    
    # Return only positive-scored candidates, or top 3 if none positive
    positive = [s for s in scored if s[0] > 0]
    if positive:
        return [(s[1], s[2], s[3]) for s in positive]
    return [(s[1], s[2], s[3]) for s in scored[:3]]

def process_asd_row(row):
    """Process a single ASD row and return patch decision."""
    source_url = row.get("source_url", "")
    model_name = row.get("model_name", "")
    target_line = row.get("target_line", "")
    old_image_url = row.get("image_url", "")
    
    result = {
        "producer": "АСД",
        "target_line": target_line,
        "model_name": model_name,
        "source_url": source_url,
        "old_image_url": old_image_url,
        "new_image_url": "",
        "decision": "not_found",
        "reason": "",
        "candidates_checked": 0,
        "selected_candidate_source": "",
        "status_before": row.get("status", ""),
        "comment_before": row.get("comment", ""),
        "patch_comment": ""
    }
    
    if not source_url:
        result["reason"] = "no_source_url"
        result["patch_comment"] = "no_source_url"
        return result
    
    # Fetch page
    resp = fetch_page(source_url)
    if not resp:
        result["reason"] = "page_not_accessible"
        result["patch_comment"] = "page_not_accessible"
        result["new_image_url"] = old_image_url
        return result
    
    soup = BeautifulSoup(resp.text, "lxml")
    
    # Determine model folder
    model_folder = determine_model_folder(soup, source_url, model_name)
    
    # Extract all image candidates
    all_candidates = extract_product_images_from_page(soup, source_url)
    
    # Filter relevant candidates
    candidates = filter_relevant_candidates(all_candidates, model_folder, model_name, source_url)
    
    result["candidates_checked"] = len(all_candidates)
    
    if not candidates:
        result["reason"] = "no_product_images_found"
        result["patch_comment"] = "no_full_door_photo_found"
        result["new_image_url"] = old_image_url
        return result
    
    # Pick best candidate (sorted by relevance score)
    best_candidate = candidates[0]
    best_url = best_candidate[0]
    best_source = best_candidate[1]
    result["selected_candidate_source"] = best_source
    
    # Decision logic
    if best_url:
        if old_image_url and old_image_url == best_url:
            result["decision"] = "keep"
            result["patch_comment"] = "asd_photo_already_optimal"
            result["new_image_url"] = old_image_url
        else:
            result["decision"] = "replace"
            result["patch_comment"] = "asd_full_door_photo_from_product_page"
            result["new_image_url"] = best_url
            result["reason"] = f"selected_{best_source}"
    else:
        result["new_image_url"] = old_image_url if old_image_url else ""
        result["reason"] = "no_good_candidates"
        result["patch_comment"] = "no_full_door_photo_found"
    
    return result

def generate_html_review(results):
    """Generate HTML review page."""
    sections = defaultdict(list)
    for r in results:
        sections[r["decision"]].append(r)
    
    total = len(results)
    replace_c = len(sections.get("replace", []))
    keep_c = len(sections.get("keep", []))
    not_found_c = len(sections.get("not_found", []))
    manual_c = len(sections.get("manual_check", []))
    
    html = f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<title>ASD Photo Patch Review</title>
<style>
body {{ font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }}
h1 {{ color: #333; }}
.summary {{ background: #fff; padding: 20px; border-radius: 8px; margin-bottom: 20px; }}
.grid {{ display: flex; flex-wrap: wrap; gap: 15px; }}
.card {{ background: #fff; border: 1px solid #ddd; border-radius: 8px; padding: 12px; width: 340px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
.card img {{ max-width: 100%; height: auto; margin: 5px 0; border: 1px solid #eee; }}
.label {{ font-weight: bold; color: #666; font-size: 11px; text-transform: uppercase; margin-top: 8px; }}
.value {{ margin-bottom: 4px; font-size: 13px; }}
.decision-replace {{ border-left: 4px solid #4CAF50; }}
.decision-keep {{ border-left: 4px solid #2196F3; }}
.decision-not_found {{ border-left: 4px solid #f44336; }}
.section-title {{ font-size: 18px; margin: 20px 0 10px; padding: 8px; border-radius: 4px; }}
.section-replace {{ background: #e8f5e9; }}
.section-keep {{ background: #e3f2fd; }}
.section-not_found {{ background: #ffebee; }}
.badge {{ display: inline-block; padding: 2px 8px; border-radius: 3px; font-size: 11px; color: #fff; }}
.badge-replace {{ background: #4CAF50; }}
.badge-keep {{ background: #2196F3; }}
.badge-not_found {{ background: #f44336; }}
</style>
</head>
<body>
<h1>ASD Photo Patch Review</h1>
<div class="summary">
<p>Total ASD rows: <strong>{total}</strong></p>
<p>Replace: <strong>{replace_c}</strong> | Keep: <strong>{keep_c}</strong> | Not found: <strong>{not_found_c}</strong> | Manual check: <strong>{manual_c}</strong></p>
</div>
"""
    
    for section_name, label, badge_cls in [
        ("replace", "Replace", "replace"),
        ("keep", "Keep", "keep"),
        ("not_found", "Not Found", "not_found"),
    ]:
        items = sections.get(section_name, [])
        if not items:
            continue
        html += f'<div class="section-title section-{section_name}"><span class="badge badge-{badge_cls}">{label}</span> {len(items)} rows</div>\n'
        html += '<div class="grid">\n'
        for r in items:
            new_img = r.get("new_image_url", "")
            old_img = r.get("old_image_url", "")
            old_thumb = f'<img src="{old_img}" onerror="this.style.display=\'none\'">' if old_img else '<div style="color:#999;font-style:italic;">(empty)</div>'
            new_thumb = f'<img src="{new_img}" onerror="this.style.display=\'none\'">' if new_img else '<div style="color:#999;font-style:italic;">(empty)</div>'
            html += f'''
<div class="card decision-{section_name}">
<div class="label">Model</div>
<div class="value">{r["model_name"][:60]}</div>
<div class="label">Source</div>
<div class="value" style="font-size:11px;word-break:break-all;"><a href="{r["source_url"]}" target="_blank">{r["source_url"][:45]}...</a></div>
<div class="label">NEW Image</div>
{new_thumb}
<div class="label">OLD Image</div>
{old_thumb}
<div class="label">Decision</div>
<div class="value"><span class="badge badge-{badge_cls}">{r["decision"]}</span> {r.get("patch_comment","")}</div>
</div>'''
        html += '</div>\n'
    
    html += "</body>\n</html>"
    return html

def main():
    print("=== ASD Photo Patch ===")
    
    rows = load_asd_rows()
    print(f"ASD rows loaded: {len(rows)}")
    
    results = []
    processed = 0
    print("Processing all 199 ASD rows...")
    
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(process_asd_row, row): row for row in rows}
        for future in as_completed(futures):
            row = futures[future]
            processed += 1
            try:
                result = future.result()
                results.append(result)
            except Exception as e:
                print(f"  ERROR [{row.get('model_name','')[:40]}]: {e}")
                results.append({
                    "producer": "АСД",
                    "target_line": row.get("target_line", ""),
                    "model_name": row.get("model_name", ""),
                    "source_url": row.get("source_url", ""),
                    "old_image_url": row.get("image_url", ""),
                    "new_image_url": row.get("image_url", ""),
                    "decision": "not_found",
                    "reason": f"exception: {e}",
                    "candidates_checked": 0,
                    "selected_candidate_source": "",
                    "status_before": row.get("status", ""),
                    "comment_before": row.get("comment", ""),
                    "patch_comment": "processing_error"
                })
            if processed % 30 == 0:
                print(f"  Processed: {processed}/{len(rows)}")
    
    assert len(results) == 199, f"Expected 199, got {len(results)}"
    print(f"Done processing: {len(results)} results")
    
    # Write outputs
    fieldnames = [
        "producer", "target_line", "model_name", "source_url",
        "old_image_url", "new_image_url", "decision", "reason",
        "candidates_checked", "selected_candidate_source",
        "status_before", "comment_before", "patch_comment"
    ]
    
    csv_path = os.path.join(OUT_DIR, "asd_photo_patch.csv")
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";")
        w.writeheader()
        w.writerows(results)
    print(f"CSV: {csv_path}")
    
    xlsx_path = None
    try:
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "ASD Photo Patch"
        for ci, h in enumerate(fieldnames, 1):
            ws.cell(row=1, column=ci, value=h)
        for ri, r in enumerate(results, 2):
            for ci, h in enumerate(fieldnames, 1):
                ws.cell(row=ri, column=ci, value=r.get(h, ""))
        xlsx_path = os.path.join(OUT_DIR, "asd_photo_patch.xlsx")
        wb.save(xlsx_path)
        print(f"XLSX: {xlsx_path}")
    except ImportError:
        print("openpyxl not available")
    
    # Update from previous run_html var
    html = generate_html_review(results)
    html_path = os.path.join(OUT_DIR, "asd_photo_patch_review.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"HTML: {html_path}")
    
    # Statistics
    decisions = Counter(r["decision"] for r in results)
    total_candidates = sum(r["candidates_checked"] for r in results)
    new_img_count = sum(1 for r in results if r.get("new_image_url"))
    
    print(f"\n=== Statistics ===")
    print(f"Total: {len(results)}")
    print(f"Replace: {decisions.get('replace', 0)}")
    print(f"Keep: {decisions.get('keep', 0)}")
    print(f"Not found: {decisions.get('not_found', 0)}")
    print(f"Manual check: {decisions.get('manual_check', 0)}")
    print(f"Candidates checked: {total_candidates}")
    print(f"Rows with new image: {new_img_count}")
    
    report = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "total_asd_rows": len(results),
        "replace": decisions.get("replace", 0),
        "keep": decisions.get("keep", 0),
        "not_found": decisions.get("not_found", 0),
        "manual_check": decisions.get("manual_check", 0),
        "candidates_checked": total_candidates,
        "rows_with_new_image": new_img_count
    }
    report_path = os.path.join(OUT_DIR, "asd_photo_patch_report.json")
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    print(f"Report: {report_path}")
    
    return results

if __name__ == "__main__":
    main()
