"""Fix bad images in doors_catalog.csv — replace with real product photos.

Strategy per row:
1. If image_url is good → keep
2. If bad → check same (producer, target_line, model_name) group for good image
3. If none in group → fetch source_url, extract og:image / gallery / product images
4. If still nothing → clear image, mark needs_review

Creates image_audit.csv with full change log.
"""
from __future__ import annotations

import csv, json, os, re, sys, time
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse, urljoin

try:
    import requests
except ImportError:
    requests = None

from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")

PROJECT = Path(__file__).resolve().parent.parent
OUT = PROJECT / "output"
CATALOG_CSV = OUT / "doors_catalog.csv"
TARGETS_CSV = PROJECT / "data" / "target_lines.csv"

FIELDS = [
    "door_type", "producer", "target_line", "model_name", "description",
    "material", "coating", "image_url", "source_url", "status", "comment",
]

# ── Strict bad-image detection ──────────────────────────────────────

# Check PATH only, not domain — to avoid false-positives like "termo-door.ru"
BAD_PATH_KEYWORDS = [
    "logo", "icon", "badge", "label", "sticker",
    "cart", "search", "arrow", "marker",
    "catalog.svg", "search_loop", "colors_images", "open.png",
    "no_image", "no-image", "noimage", "placeholder", "blank",
    "sprite", "loader", "pixel", "favicon",
    # Thumbnail indicators (not part of domain)
    "cache/thumb", "thumbnail", "/thumb", "miniatur",
    # Sale/promo badges
    "_sale", "sale_badge", "sale.", "/sale/",
    "hit_prodazh", "hit.png", "/hit/", "new_badge",
    "new.", "_new.", "/new/",
]

# Small-size patterns in URL path (dimensions in path)
SMALL_SIZE_RE = re.compile(r"[/_](\d{1,3})x(\d{1,3})[./]")


def _path_only(url: str) -> str:
    """Return the path portion of a URL for keyword matching."""
    try:
        parsed = urlparse(url)
        return parsed.path + "?" + (parsed.query or "")
    except Exception:
        return url


def is_bad_image_url(url: str) -> bool:
    """Strict check — only flag if the PATH contains bad keywords."""
    if not url:
        return False
    path = _path_only(url).lower()

    # Check path keywords
    for kw in BAD_PATH_KEYWORDS:
        if kw in path:
            return True

    # Check small dimensions (path-based like 96x96, 150x150)
    for m in SMALL_SIZE_RE.finditer(path):
        try:
            w, h = int(m.group(1)), int(m.group(2))
            if w < 400 or h < 400:
                return True
        except ValueError:
            pass

    # Check for Tilda /resize/ with small sizes
    if "/resize/" in path:
        m2 = re.search(r"/resize/(\d+)x", path)
        if m2:
            try:
                sz = int(m2.group(1))
                if sz < 400:
                    return True
            except ValueError:
                pass

    # Check for WordPress cache/thumb which are always resized versions
    # cache/thumb with 0xN dimension (auto-width) is caught by SMALL_SIZE_RE above

    return False


def is_tilda_empty_or_resize(url: str) -> bool:
    """Check for Tilda CDN /empty/ (interior renders) or tiny /resize/."""
    u = url.lower()
    if "tildacdn.com" not in u:
        return False
    # /empty/ URLs are interior renders, not door photos
    if "/empty/" in u:
        return True
    # /resize/Nx with N < 400
    m = re.search(r"/resize/(\d+)x", u)
    if m:
        try:
            return int(m.group(1)) < 400
        except ValueError:
            pass
    return False


def is_tilda_empty_or_resize(url: str) -> bool:
    """Check for Tilda CDN /empty/ which are interior renders, not door photos."""
    u = url.lower()
    if "tildacdn.com" in u and "/empty/" in u:
        return True
    return False


def is_qr_or_icon(url: str) -> bool:
    """Strict check for known trash files."""
    u = url.lower()
    for bad in ["qr.png", "qr.svg", "favicon", "/icons/", "/images/arrow"]:
        if bad in u:
            return True
    return False


def has_good_size(url: str) -> bool:
    """Check if URL has explicit dimensions ≥ 400px."""
    m = re.search(r"[_/](\d{3,4})x(\d{2,4})[./]", url)
    if m:
        try:
            w, h = int(m.group(1)), int(m.group(2))
            return w >= 400 or h >= 400
        except ValueError:
            pass
    # Tilda /resize/Nx where N≥400
    m2 = re.search(r"/resize/(\d+)x", url)
    if m2:
        try:
            return int(m2.group(1)) >= 400
        except ValueError:
            pass
    return False


def _check_file_size(url: str, session = None) -> bool:
    """Check if image file is plausibly large enough via HEAD request.
    Returns True if file seems OK, False if suspicious."""
    if not session or not url:
        return True  # can't check, assume OK
    try:
        r = session.head(url, timeout=5, allow_redirects=True)
        if r.status_code >= 400:
            return False  # broken link
        size = int(r.headers.get("Content-Length", 0))
        if 0 < size < 10240:  # less than 10KB -> suspicious
            return False
        return True
    except Exception:
        return True  # can't check, assume OK


def image_quality(url: str, session = None) -> str:
    """Classify image: 'good', 'bad', or 'empty'."""
    if not url:
        return "empty"
    if is_qr_or_icon(url):
        return "bad"
    if is_tilda_empty_or_resize(url):
        return "bad"
    if is_bad_image_url(url):
        return "bad"
    # For WordPress cache/thumb with 0xN dimensions where N ≥ 400, check original
    # For all others with good size indicators, consider good
    if has_good_size(url):
        return "good"
    return "good"  # default: if not bad, assume good


# ── Image extraction from page ──────────────────────────────────────

def extract_page_images(url: str, session) -> list[str]:
    """Extract all plausible product images from a page."""
    images = []
    try:
        r = session.get(url, timeout=(5, 10))
        if r.status_code >= 400:
            return images
        html = r.text
    except Exception:
        return images

    soup = BeautifulSoup(html, "lxml")

    # 1. Open Graph image
    og = soup.find("meta", property="og:image")
    if og and og.get("content"):
        img = urljoin(url, og["content"].strip())
        if img and not is_bad_image_url(img) and not is_tilda_empty_or_resize(img):
            images.append(("og:image", img))

    # 2. JSON-LD product images
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string)
            if isinstance(data, dict):
                for img_data in _walk_json_ld(data, "image"):
                    if isinstance(img_data, str):
                        full = urljoin(url, img_data)
                        if full and not is_bad_image_url(full):
                            images.append(("ld+json", full))
                    elif isinstance(img_data, dict):
                        for key in ("url", "contentUrl", "image"):
                            if img_data.get(key):
                                full = urljoin(url, img_data[key])
                                if full and not is_bad_image_url(full):
                                    images.append(("ld+json", full))
        except (json.JSONDecodeError, AttributeError, TypeError):
            pass

    # 3. Product gallery images (common selectors)
    for sel in [
        "div.product-gallery img", "div.gallery img", "div.product__gallery img",
        "div.swiper img", "div.owl-carousel img", "div.product-images img",
        "ul.thumbnails img", "div.product-slider img",
        "div.fotorama img", "div.slick-slide img",
        "div.simple-slider-list img", "div.image-list img",
        "div.product-photo img", "div.product__photo img",
        "div.card-gallery img", "div.card__gallery img",
        "[data-gallery] img", ".product-media img",
        # Shop-Script specific
        ".product-gallery__item img", ".product__slide img",
    ]:
        for img in soup.select(sel):
            src = img.get("src") or img.get("data-src") or img.get("data-lazy") or ""
            if src:
                full = urljoin(url, src.strip())
                if full and not is_bad_image_url(full) and not is_tilda_empty_or_resize(full):
                    images.append(("gallery", full))

    # 5. Any large images (fallback — filter by plausible path patterns)
    for img in soup.find_all("img"):
        src = img.get("src") or img.get("data-src") or ""
        if not src:
            continue
        full = urljoin(url, src.strip())
        low = full.lower()
        # Skip trash
        if is_bad_image_url(full) or is_tilda_empty_or_resize(full) or is_qr_or_icon(full):
            continue
        # Only accept images with plausible product paths
        if any(kw in low for kw in [
            "product", "upload", "uploads", "gallery", "catalog", "photo",
            "img/", "images/", "door", "doors",
        ]):
            images.append(("fallback", full))

    return images


def _walk_json_ld(data, target_key, max_depth=5):
    """Walk JSON-LD structure to find all values for target_key."""
    if max_depth <= 0:
        return
    if isinstance(data, dict):
        for k, v in data.items():
            if k == target_key:
                if isinstance(v, list):
                    yield from v
                else:
                    yield v
            else:
                yield from _walk_json_ld(v, target_key, max_depth - 1)
    elif isinstance(data, list):
        for item in data:
            yield from _walk_json_ld(item, target_key, max_depth - 1)


def select_best_image(images: list[tuple[str, str]]) -> str | None:
    """From extracted images (source, url), pick the best one."""
    scored = []
    for src, url in images:
        score = 0
        # Prefer og:image
        if src == "og:image":
            score += 200
        elif src == "ld+json":
            score += 150
        elif src == "gallery":
            score += 100
        elif src == "product":
            score += 80
        elif src == "fallback":
            score += 30

        # Prefer static.tildacdn.com (original, not resized)
        if "static.tildacdn.com" in url:
            score += 100
        # Prefer .jpg
        if any(url.lower().endswith(ext) for ext in [".jpg", ".jpeg"]):
            score += 30
        # Prefer larger dimensions
        m = re.search(r"[_/](\d{3,4})x(\d{2,4})[./]", url)
        if m:
            try:
                w, h = int(m.group(1)), int(m.group(2))
                score += min(w, h) // 10
            except ValueError:
                pass
        # Prefer deeper path (more specific)
        score += url.count("/")
        scored.append((score, url))

    scored.sort(key=lambda x: -x[0])
    return scored[0][1] if scored else None


# ── Main processing ─────────────────────────────────────────────────

def main():
    session = requests.Session() if requests else None
    session.headers["User-Agent"] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"

    # 1. Load catalog
    with open(CATALOG_CSV, encoding="utf-8-sig") as f:
        cat = list(csv.DictReader(f, delimiter=";"))
    print(f"Loaded catalog: {len(cat)} rows")

    # 2. Classify current images
    audit = []
    total = len(cat)
    bad_rows = []

    for r in cat:
        url = r.get("image_url", "") or ""
        q = image_quality(url)
        if q == "bad":
            bad_rows.append(r)

    print(f"Rows with bad images: {len(bad_rows)}")

    # 3. Build groups by (producer, target_line, model_name)
    groups = defaultdict(list)
    for r in cat:
        key = (r.get("producer", ""), r.get("target_line", ""), r.get("model_name", ""))
        groups[key].append(r)

    # 4. Fix bad images
    fixed = 0
    cleared = 0
    from_dup = 0
    from_page = 0
    skipped = 0

    for r in bad_rows:
        key = (r.get("producer", ""), r.get("target_line", ""), r.get("model_name", ""))
        old_url = r.get("image_url", "") or ""
        old_quality = "bad"
        source_url = r.get("source_url", "") or ""

        # Step A: check same-model group for good image
        group = groups.get(key, [])
        best_in_group = None
        for gr in group:
            if gr is r:
                continue
            u = gr.get("image_url", "") or ""
            if image_quality(u) == "good":
                best_in_group = u
                break

        if best_in_group:
            # Found a good image in the same group
            r["image_url"] = best_in_group
            if image_quality(best_in_group) == "good":
                r["status"] = "found"
            fixed += 1
            from_dup += 1
            action = "replaced_from_duplicate"
            audit.append({
                "producer": r.get("producer", ""),
                "target_line": r.get("target_line", ""),
                "model_name": r.get("model_name", ""),
                "old_image_url": old_url,
                "new_image_url": best_in_group,
                "source_url": source_url,
                "image_status_before": "bad",
                "image_status_after": "good",
                "action": "replaced_from_duplicate",
                "comment": "found good image in same model group",
            })
            continue

        # Step B: fetch source_url to extract product images
        if session and source_url and not is_bad_image_url(source_url):
            try:
                page_images = extract_page_images(source_url, session)
                if page_images:
                    best = select_best_image(page_images)
                    if best and image_quality(best) == "good":
                        r["image_url"] = best
                        r["status"] = "found"
                        fixed += 1
                        from_page += 1
                        audit.append({
                            "producer": r.get("producer", ""),
                            "target_line": r.get("target_line", ""),
                            "model_name": r.get("model_name", ""),
                            "old_image_url": old_url,
                            "new_image_url": best,
                            "source_url": source_url,
                            "image_status_before": "bad",
                            "image_status_after": "good",
                            "action": "replaced_from_source_page",
                            "comment": f"extracted from page ({page_images[0][0]})",
                        })
                        continue
            except Exception as e:
                pass

        # Step C: no alternative found — clear the image
        r["image_url"] = ""
        existing_comment = r.get("comment", "")
        if "image_not_found_after_cleanup" not in existing_comment:
            parts = [c for c in existing_comment.split("; ") if c]
            parts.append("image_not_found_after_cleanup")
            r["comment"] = "; ".join(parts)
        r["status"] = "needs_review"
        cleared += 1
        audit.append({
            "producer": r.get("producer", ""),
            "target_line": r.get("target_line", ""),
            "model_name": r.get("model_name", ""),
            "old_image_url": old_url,
            "new_image_url": "",
            "source_url": source_url,
            "image_status_before": "bad",
            "image_status_after": "empty",
            "action": "cleared_bad_image",
            "comment": "no good image found in group or on source page",
        })

    print(f"Fixed from duplicate group: {from_dup}")
    print(f"Fixed from source page: {from_page}")
    print(f"Cleared (no alternative): {cleared}")
    print(f"Total bad processed: {len(bad_rows)}")

    # 5. Handle remaining rows without images
    for r in cat:
        url = r.get("image_url", "") or ""
        q = image_quality(url)
        if q == "bad":
            # Shouldn't happen after fix, but just in case
            r["image_url"] = ""
            r["status"] = "needs_review"
            if "image_not_found_after_cleanup" not in r.get("comment", ""):
                r["comment"] = "; ".join(filter(None, [r.get("comment", ""), "image_not_found_after_cleanup"]))
            audit.append({
                "producer": r.get("producer", ""),
                "target_line": r.get("target_line", ""),
                "model_name": r.get("model_name", ""),
                "old_image_url": url,
                "new_image_url": "",
                "source_url": r.get("source_url", "") or "",
                "image_status_before": "bad",
                "image_status_after": "empty",
                "action": "cleared_bad_image",
                "comment": "still bad after processing",
            })

    # 6. Dedup by (producer, target_line, model_name) — keep best image
    # Rebuild groups from updated catalog
    groups_after = defaultdict(list)
    for r in cat:
        key = (r.get("producer", ""), r.get("target_line", ""), r.get("model_name", ""))
        groups_after[key].append(r)

    deduped = []
    dup_report = []
    seen = set()
    for r in cat:
        key = (r.get("producer", ""), r.get("target_line", ""), r.get("model_name", ""))
        if key in seen:
            # Duplicate row — record in report
            dup_report.append({
                "producer": r.get("producer", ""),
                "target_line": r.get("target_line", ""),
                "model_name": r.get("model_name", ""),
                "duplicates_count_before": len(groups_after.get(key, [])),
                "kept_image_url": "",
                "removed_image_urls": r.get("image_url", "") or "",
                "reason": "photo_duplicate_removed",
            })
            continue
        seen.add(key)
        deduped.append(r)

    # Final dedup: within each group, choose best image row
    final_rows = []
    final_groups = defaultdict(list)
    for r in deduped:
        key = (r.get("producer", ""), r.get("target_line", ""), r.get("model_name", ""))
        final_groups[key].append(r)

    for key, grp in final_groups.items():
        if len(grp) == 1:
            final_rows.append(grp[0])
        else:
            # Score by image quality
            scored = []
            for r in grp:
                url = r.get("image_url", "") or ""
                q = image_quality(url)
                score = 0
                if q == "good":
                    score += 1000
                    # Prefer larger images
                    m = re.search(r"[_/](\d{3,4})x(\d{2,4})[./]", url)
                    if m:
                        try:
                            score += (int(m.group(1)) + int(m.group(2)))
                        except ValueError:
                            pass
                elif q == "empty":
                    score -= 500
                else:
                    score -= 1000
                scored.append((score, r))
            scored.sort(key=lambda x: -x[0])
            best = scored[0][1]
            final_rows.append(best)
            # If there were others, record
            if len(scored) > 1:
                removed_urls = []
                for s, r2 in scored[1:]:
                    removed_urls.append(r2.get("image_url", "") or "")
                dup_report.append({
                    "producer": best.get("producer", ""),
                    "target_line": best.get("target_line", ""),
                    "model_name": best.get("model_name", ""),
                    "duplicates_count_before": len(grp),
                    "kept_image_url": best.get("image_url", "") or "",
                    "removed_image_urls": "; ".join(filter(None, removed_urls)),
                    "reason": "photo_duplicate_removed",
                })

    print(f"\nAfter dedup: {len(cat)} → {len(final_rows)}")

    # 7. Write outputs
    OUT.mkdir(parents=True, exist_ok=True)

    # doors_catalog.csv
    with open(OUT / "doors_catalog.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS, delimiter=";")
        w.writeheader()
        for r in final_rows:
            w.writerow({k: r.get(k, "") for k in FIELDS})

    # Final counts
    final_total = len(final_rows)
    final_with_img = sum(1 for r in final_rows if image_quality(r.get("image_url", "")) == "good")
    final_no_img = sum(1 for r in final_rows if not r.get("image_url", ""))
    final_bad_img = sum(1 for r in final_rows if image_quality(r.get("image_url", "")) == "bad")
    final_empty = sum(1 for r in final_rows if not r.get("image_url", ""))

    # wordpress_import.csv
    wp_fields = ["post_title", "post_content", "producer", "target_line",
                 "material", "coating", "image_url", "source_url",
                 "door_type", "status"]
    with open(OUT / "wordpress_import.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=wp_fields, delimiter=";")
        w.writeheader()
        for r in final_rows:
            w.writerow({
                "post_title": r.get("model_name", ""),
                "post_content": r.get("description", ""),
                "producer": r.get("producer", ""),
                "target_line": r.get("target_line", ""),
                "material": r.get("material", "") or "",
                "coating": r.get("coating", "") or "",
                "image_url": r.get("image_url", "") or "",
                "source_url": r.get("source_url", "") or "",
                "door_type": r.get("door_type", ""),
                "status": r.get("status", ""),
            })

    # image_audit.csv
    audit_fields = [
        "producer", "target_line", "model_name",
        "old_image_url", "new_image_url", "source_url",
        "image_status_before", "image_status_after", "action", "comment",
    ]
    with open(OUT / "image_audit.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=audit_fields, delimiter=";")
        w.writeheader()
        for a in audit:
            w.writerow({k: a.get(k, "") for k in audit_fields})

    # duplicates_report.csv
    dup_fields = [
        "producer", "target_line", "model_name",
        "duplicates_count_before", "kept_image_url",
        "removed_image_urls", "reason",
    ]
    with open(OUT / "duplicates_report.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=dup_fields, delimiter=";")
        w.writeheader()
        for d in dup_report:
            w.writerow({k: d.get(k, "") for k in dup_fields})

    # Producer counts
    prod_counts = defaultdict(int)
    for r in final_rows:
        p = r.get("producer", "")
        if p:
            prod_counts[p] += 1

    # Found lines
    found_lines = set()
    for r in final_rows:
        found_lines.add((r.get("producer", "").lower().strip(), r.get("target_line", "").lower().strip()))

    targets = []
    if TARGETS_CSV.exists():
        with open(TARGETS_CSV, encoding="utf-8-sig") as f:
            targets = list(csv.DictReader(f, delimiter=";"))

    total_targets = len(targets) if targets else 0
    total_found = len(found_lines)

    # not_found.csv
    not_found = []
    for t in targets:
        key = (t.get("manufacturer", "").lower().strip(), t.get("target_line", "").lower().strip())
        if t.get("manufacturer", "").lower().strip() == "асд":
            continue
        if key not in found_lines:
            not_found.append({
                "door_type": t.get("door_type", ""),
                "manufacturer": t.get("manufacturer", ""),
                "target_line": t.get("target_line", ""),
                "source_site": t.get("source_site", ""),
                "status": "not_found",
                "comment": "still not found after fixes",
            })
    nf_fields = ["door_type", "manufacturer", "target_line", "source_site", "status", "comment"]
    with open(OUT / "not_found.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=nf_fields, delimiter=";")
        w.writeheader()
        w.writerows(not_found)

    # qa_report.csv
    qa_fields = [
        "producer", "target_line", "rows_count", "unique_models_count",
        "image_missing_count", "bad_images_removed_count",
        "duplicates_removed_count", "status", "comment",
    ]
    with open(OUT / "qa_report.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=qa_fields, delimiter=";")
        w.writeheader()
        qa_grouped = defaultdict(list)
        for r in final_rows:
            qa_grouped[(r.get("producer", "").lower().strip(), r.get("target_line", "").lower().strip())].append(r)
        for t in targets:
            man = t.get("manufacturer", "")
            line = t.get("target_line", "")
            key = (man.lower().strip(), line.lower().strip())
            grp = qa_grouped.get(key, [])
            rows_count = len(grp)
            unique_models = len(set(r.get("model_name", "") for r in grp))
            img_missing = sum(1 for r in grp if not r.get("image_url", ""))
            bad_imgs = sum(1 for r in grp if image_quality(r.get("image_url", "")) == "bad")
            status = "found"
            comments = []
            if not grp:
                status = "not_found"
                comments.append("no rows")
            elif all("collection_not_found" in r.get("comment", "") for r in grp):
                status = "needs_manual_review"
                comments.append("collection not found on site")
            elif img_missing > rows_count * 0.5 and rows_count > 0:
                status = "needs_manual_review"
                comments.append(f"{img_missing}/{rows_count} no image")

            if man.lower().strip() == "асд":
                comments = [c for c in comments if "collection not found" not in c.lower()]
                comments.append("collection_to_fill_manually")
                status = "needs_manual_review"

            if not comments:
                comments.append("ok")

            w.writerow({
                "producer": man,
                "target_line": line,
                "rows_count": rows_count,
                "unique_models_count": unique_models,
                "image_missing_count": img_missing,
                "bad_images_removed_count": bad_imgs,
                "duplicates_removed_count": 0,
                "status": status,
                "comment": "; ".join(comments),
            })

    # run_report.json
    report = {
        "fixed_at": datetime.now().isoformat(),
        "total_targets": total_targets,
        "total_lines_found": total_found,
        "total_items": final_total,
        "lines_not_found": len(not_found),
        "producers": dict(prod_counts),
        "image_cleanup": {
            "rows_before": len(cat),
            "rows_after": final_total,
            "image_urls_checked": total,
            "bad_images_found": len(bad_rows),
            "bad_images_replaced_from_duplicate": from_dup,
            "bad_images_replaced_from_source_page": from_page,
            "bad_images_cleared": cleared,
            "rows_with_image_after": final_with_img,
            "rows_without_image_after": final_no_img,
            "rows_with_bad_image_after": final_bad_img,
            "duplicates_removed": len(dup_report),
            "image_audit_path": str(OUT / "image_audit.csv"),
        },
    }
    with open(OUT / "run_report.json", "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)

    # XLSX
    _make_xlsx(OUT / "doors_catalog.xlsx", final_rows, targets)

    # Summary
    print(f"\n{'='*60}")
    print(f"  IMAGE CLEANUP DONE")
    print(f"  Rows: {total} → {final_total}")
    print(f"  Bad images found: {len(bad_rows)}")
    print(f"    Replaced from duplicate: {from_dup}")
    print(f"    Replaced from source page: {from_page}")
    print(f"    Cleared: {cleared}")
    print(f"  After cleanup:")
    print(f"    With good image: {final_with_img}")
    print(f"    Without image: {final_no_img}")
    print(f"    Bad remaining: {final_bad_img}")
    print(f"  Lines: {total_found}/{total_targets}")
    print(f"  Audit: {len(audit)} entries")
    print(f"{'='*60}")


def _make_xlsx(path, rows, targets):
    from collections import defaultdict
    wb = Workbook()
    ws = wb.active
    ws.title = "catalog"
    headers = ["Тип двери", "Производитель", "Линейка", "Название", "Описание",
               "Материал", "Покрытие", "Ссылка на фото", "Ссылка на карточку",
               "Статус", "Комментарий"]
    ws.append(headers)
    for r in rows:
        ws.append([r.get("door_type", ""), r.get("producer", ""),
                   r.get("target_line", ""), r.get("model_name", ""),
                   r.get("description", ""), r.get("material", ""),
                   r.get("coating", ""), r.get("image_url", ""),
                   r.get("source_url", ""), r.get("status", ""),
                   r.get("comment", "")])
    widths = [16, 22, 28, 40, 70, 18, 16, 60, 60, 14, 40]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"
    wb.save(path)


if __name__ == "__main__":
    main()
