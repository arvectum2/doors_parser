"""Merge good data backup with newly parsed problem-producer results.
Only reads doors_catalog.csv from subdirectories. Filters garbage model names and clears bad image URLs.
Keeps rows even with empty image_url — text data is still valuable.

Usage:
    python merge_results.py [--good output/good_data_backup.csv] [--new-dir output/problem_sites] [--out output]
"""
import csv
import json
import re
import sys
import argparse
from collections import Counter
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

FIELDS = [
    "door_type", "producer", "target_line", "model_name", "description",
    "material", "coating", "image_url", "source_url", "status", "comment",
]

BAD_IMAGE_KEYWORDS = [
    "logo", "icon", "pixel", "placeholder", "sprite", "loader",
    "blank", "no-photo", "nophoto", "favicon", "banner",
    "payment", "social", "rutub", "capcha", "captcha",
    "open.png", "close.png", "play.png",
    "arrow_up", "arrow", "catalog.svg", "search_loop", "colors_images",
    "star", "novinka", "printest", "tele2", "icon_ft_soc",
    ".svg", "yandex", "mc.yandex", "watch/", "no_image_glass",
    "cart_mob.webp", "/empty/",
]

BAD_MODEL_KEYWORDS = [
    "каталог", "поиск", "меню", "корзина", "избранное", "сравнение",
    "коробка", "наличник", "добор", "погонаж", "телескоп", "притвор",
    "капитель", "карниз", "розетка", "плинтус", "лакобель", "дуб скай",
    "цвет", "покрытие",
]


def normalize(text):
    text = (text or "").lower().replace("ё", "е")
    text = re.sub(r"[^a-zа-я0-9]+", " ", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


def normalize_key(text):
    return (text or "").lower().strip()


def has_price(model_name):
    mn = (model_name or "").strip()
    if "₽" in mn or "руб" in mn.lower():
        return True
    mn_no_space = mn.replace(" ", "")
    if re.search(r"^\d{1,6}[.,]?\d{0,2}₽", mn_no_space):
        return True
    if re.search(r"^\d{1,6}[.,]?\d{0,2}руб", mn_no_space, re.I):
        return True
    if re.match(r"^от\s+\d[\d\s]*?[₽p]", mn, re.I):
        return True
    if re.match(r"^\d[\d\s]*\d\s*₽", mn):
        return True
    return False


def is_generic_page(title, url):
    ntitle = normalize(title)
    path = urlparse(url).path.strip("/")
    if not path:
        return True
    GENERIC_TITLES = ["межкомнатные двери", "входные двери", "каталог", "главная", "главная страница"]
    for g in GENERIC_TITLES:
        if normalize(g) == ntitle:
            return True
    return False


def _is_bad_image(url="", alt="", klass="", low=""):
    hay = (url + " " + alt + " " + klass + " " + low).lower()
    return any(k in hay for k in BAD_IMAGE_KEYWORDS)


ADDITIONAL_COMPONENT_KEYWORDS = [
    "коробка", "наличник", "добор", "погонаж", "телескоп", "притвор",
    "капитель", "карниз", "розетка", "плинтус",
]

COMPONENT_START_KEYWORDS = [
    "подпятник", "капитель", "коробка", "наличник", "добор",
    "погонаж", "телескоп", "притвор", "карниз", "розетка", "плинтус",
]

PRICE_CLEAN_PATTERNS = [
    re.compile(r"\(\*цена указана[^)]*\)", re.I),
    re.compile(r"\(цена указана[^)]*\)", re.I),
    re.compile(r"цена указана\s*:\s*[^)]+", re.I),
    re.compile(r"Цена указана[^.]*\.?\s*", re.I),
    re.compile(r"от\s+\d[\d\s]*[₽р]"),
    re.compile(r"\d[\d\s]*[₽р]"),
]


def clean_description(desc):
    if not desc:
        return ""
    result = desc
    for pat in PRICE_CLEAN_PATTERNS:
        result = pat.sub("", result)
    result = re.sub(r"\s+", " ", result).strip()
    result = re.sub(r"[;,\s]*$", "", result)
    return result


def has_bad_model_name(model_name):
    """Remove only if model name is SHORT (≤3 words) AND contains a bad keyword.
    This prevents removing valid collection descriptions like
    'Коллекция X Комплектация: полотно, коробка, наличник' which happen to contain
    garbage words as part of a longer description."""
    mn = (model_name or "").lower().strip()
    words = [w for w in re.split(r'[\s,;:./|–—\-]+', mn) if w]
    if len(words) > 3:
        return False
    for kw in BAD_MODEL_KEYWORDS:
        if kw == mn:
            return True
        if mn.startswith(kw + " ") or mn.startswith(kw + "—") or mn.startswith(kw + "-"):
            return True
        if mn.endswith(" " + kw) or mn.endswith("—" + kw) or mn.endswith("-" + kw):
            return True
        if " " + kw + " " in mn:
            return True
    return False


def is_component_row(model_name):
    """Check if row is a component (not a door) by model name starting with a component keyword.
    Keeps rows where model_name contains 'дверь' (those are actual door models)."""
    mn = (model_name or "").lower().strip()
    if "дверь" in mn:
        return False
    mn_lower = mn.lower()
    for kw in ADDITIONAL_COMPONENT_KEYWORDS:
        if mn_lower == kw:
            return True
        if mn_lower.startswith(kw + " ") or mn_lower.startswith(kw + "—") or mn_lower.startswith(kw + "-"):
            return True
    return False


def is_component_start(model_name):
    mn = (model_name or "").lower().strip()
    if "дверь" in mn:
        return False
    for kw in COMPONENT_START_KEYWORDS:
        if mn == kw:
            return True
        if mn.startswith(kw + " ") or mn.startswith(kw + "-") or mn.startswith(kw + "—"):
            return True
    return False


def read_csv(path):
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        rows = []
        for row in reader:
            d = {}
            for k, v in row.items():
                k = k.strip()
                d[k] = (v or "").strip()
            if "producer" not in d and "manufacturer" in d:
                d["producer"] = d.pop("manufacturer")
            if "target_line" not in d and "line" in d:
                d["target_line"] = d.pop("line")
            for f in FIELDS:
                d.setdefault(f, "")
            rows.append(d)
    return rows


def write_csv(path, rows):
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS, delimiter=";")
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in FIELDS})


def main():
    ap = argparse.ArgumentParser(description="Merge good data with new parser results")
    ap.add_argument("--good", type=Path, default=Path("output/good_data_backup.csv"))
    ap.add_argument("--new", type=Path, nargs="*")
    ap.add_argument("--new-dir", type=Path, default=None)
    ap.add_argument("--out", type=Path, default=Path("output"))
    ap.add_argument("--targets", type=Path, default=Path("data/target_lines.csv"))
    args = ap.parse_args()

    out_dir = args.out
    out_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 60)
    print("  Merge Results")
    print("=" * 60)

    good = read_csv(args.good)
    print(f"\nGood data: {len(good)} rows")

    new_files = list(args.new) if args.new else []
    if args.new_dir:
        # Read the merged problem_sites CSV (latest single-producer run)
        merged_new = args.new_dir / "doors_catalog.csv"
        if merged_new.exists():
            nrows = read_csv(merged_new)
            print(f"  {merged_new.parent.name}/{merged_new.name}: {nrows} rows (merged new)")
            new_files.append(merged_new)
        # Also read per-producer subdirectories (backward compat)
        new_files.extend(sorted(args.new_dir.glob("*/doors_catalog.csv")))

    all_new = []
    for nf in new_files:
        nrows = read_csv(nf)
        print(f"  {nf.parent.name}/{nf.name}: {len(nrows)} rows")
        all_new.extend(nrows)
    print(f"New data total: {len(all_new)} rows")

    combined = good + all_new
    print(f"Combined: {len(combined)} rows")

    seen = set()
    deduped = []
    for r in combined:
        key = (
            normalize(r.get("producer", "")),
            normalize(r.get("target_line", "")),
            normalize(r.get("model_name", "")),
            normalize(r.get("image_url", "")),
            normalize(r.get("source_url", "")),
        )
        if key in seen:
            continue
        seen.add(key)
        deduped.append(r)
    print(f"After dedup: {len(deduped)} rows (removed {len(combined) - len(deduped)})")

    # Filter price rows
    before_price = len(deduped)
    deduped = [r for r in deduped if not has_price(r.get("model_name", ""))]
    removed_price = before_price - len(deduped)
    print(f"Price rows removed: {removed_price}")

    # Filter component-start rows
    before_comp_start = len(deduped)
    deduped = [r for r in deduped if not is_component_start(r.get("model_name", ""))]
    removed_comp_start = before_comp_start - len(deduped)
    print(f"Component-start rows removed: {removed_comp_start}")

    # Filter Коллекция+Комплектация rows
    before_koll = len(deduped)
    deduped = [r for r in deduped if not (r.get("model_name","").lower().startswith("коллекция") and "комплектация" in r.get("model_name","").lower())]
    removed_koll = before_koll - len(deduped)
    print(f"Коллекция+Комплектация rows removed: {removed_koll}")

    # Clean descriptions (remove price phrases)
    for r in deduped:
        r["description"] = clean_description(r.get("description", ""))

    # Filter bad images + empty rows after bad image removal
    cleaned = []
    removed_bad_img = 0
    for r in deduped:
        url = r.get("image_url", "")
        alt = r.get("model_name", "")
        klass = r.get("comment", "")
        if url and _is_bad_image(url, alt, klass, normalize(url)):
            r["image_url"] = ""
            r["comment"] = ("; ".join(filter(None, [r.get("comment", ""), "bad_image_url_removed"]))).strip()
            removed_bad_img += 1
        cleaned.append(r)
    print(f"Bad image URLs cleared: {removed_bad_img}")

    # Filter garbage model names (remove rows entirely)
    before_model = len(cleaned)
    cleaned = [r for r in cleaned if not has_bad_model_name(r.get("model_name", ""))]
    removed_model = before_model - len(cleaned)
    print(f"Garbage model names removed: {removed_model}")

    # Filter component rows (model name starts with Коробка/Наличник/Добор etc., NOT containing Дверь)
    before_comp = len(cleaned)
    cleaned = [r for r in cleaned if not is_component_row(r.get("model_name", ""))]
    removed_comp = before_comp - len(cleaned)
    print(f"Component rows removed: {removed_comp}")

    # Count (but KEEP) rows with empty image_url — text data is still valuable
    empty_img_count = sum(1 for r in cleaned if not r.get("image_url", ""))
    print(f"Empty-image rows kept: {empty_img_count} (text data preserved)")

    # Ensure all rows have description
    for r in cleaned:
        if not r.get("description"):
            r["description"] = f"Модель из линейки {r.get('target_line', '')}, производитель {r.get('producer', '')}"

    write_csv(out_dir / "doors_catalog.csv", cleaned)
    print(f"\nWritten: {out_dir / 'doors_catalog.csv'} ({len(cleaned)} rows)")

    # Write WP import
    wp_fields = ["post_title", "post_content", "producer", "target_line",
                 "material", "coating", "image_url", "source_url",
                 "door_type", "status"]
    with (out_dir / "wordpress_import.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=wp_fields, delimiter=";")
        w.writeheader()
        for r in cleaned:
            w.writerow({
                "post_title": r.get("model_name", ""),
                "post_content": r.get("description", ""),
                "producer": r.get("producer", ""),
                "target_line": r.get("target_line", ""),
                "material": r.get("material", ""),
                "coating": r.get("coating", ""),
                "image_url": r.get("image_url", ""),
                "source_url": r.get("source_url", ""),
                "door_type": r.get("door_type", ""),
                "status": r.get("status", ""),
            })

    # Write not_found.csv
    targets = list(csv.DictReader(open(args.targets, encoding="utf-8-sig"), delimiter=";"))
    found = set()
    for r in cleaned:
        found.add((normalize_key(r.get("producer", "")), normalize_key(r.get("target_line", ""))))
    
    nf_count = 0
    with (out_dir / "not_found.csv").open("w", encoding="utf-8-sig", newline="") as f:
        nf_fields = ["door_type", "manufacturer", "target_line", "source_site", "status", "comment"]
        w = csv.DictWriter(f, fieldnames=nf_fields, delimiter=";")
        w.writeheader()
        for t in targets:
            key = (normalize_key(t.get("manufacturer", "")), normalize_key(t.get("target_line", "")))
            if key not in found:
                nf_count += 1
                w.writerow({
                    "door_type": t.get("door_type", ""),
                    "manufacturer": t.get("manufacturer", ""),
                    "target_line": t.get("target_line", ""),
                    "source_site": t.get("source_site", ""),
                    "status": "not_found",
                    "comment": "page_not_found — нет подходящей страницы на сайте",
                })
    print(f"Not found entries: {nf_count}")

    # QA report
    qa_fields = [
        "producer", "target_line", "rows_count", "unique_models_count",
        "unique_source_urls_count", "bad_images_count",
        "generic_pages_count", "duplicate_rows_count",
        "needs_review_count", "status", "comment",
    ]
    with (out_dir / "qa_report.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=qa_fields, delimiter=";")
        w.writeheader()
        
        grouped = {}
        for r in cleaned:
            grouped.setdefault((normalize_key(r.get("producer", "")), normalize_key(r.get("target_line", ""))), []).append(r)
        
        for t in targets:
            man = t.get("manufacturer", "")
            line = t.get("target_line", "")
            rows = grouped.get((normalize_key(man), normalize_key(line)), [])
            
            rows_count = len(rows)
            unique_models = len(set(normalize(r.get("model_name", "")) for r in rows))
            unique_urls = len(set(r.get("source_url", "") for r in rows))
            bad_imgs = sum(1 for r in rows if any(k in (r.get("image_url", "") or "").lower() for k in BAD_IMAGE_KEYWORDS))
            generic_pgs = sum(1 for r in rows if is_generic_page(r.get("model_name", ""), r.get("source_url", "")))
            needs_review = sum(1 for r in rows if r.get("status") == "needs_review")
            
            status = "found" if rows else "not_found"
            comments = []
            if not rows:
                comments.append("no rows")
            elif generic_pgs > rows_count * 0.3:
                status = "needs_manual_review"
                comments.append(f"{generic_pgs}/{rows_count} generic pages")
            elif bad_imgs > rows_count * 0.5:
                status = "needs_manual_review"
                comments.append(f"{bad_imgs}/{rows_count} bad images")
            elif needs_review > rows_count * 0.5:
                status = "needs_manual_review"
                comments.append(f"{needs_review}/{rows_count} needs_review")
            if not comments:
                comments.append("ok")
            
            w.writerow({
                "producer": man,
                "target_line": line,
                "rows_count": rows_count,
                "unique_models_count": unique_models,
                "unique_source_urls_count": unique_urls,
                "bad_images_count": bad_imgs,
                "generic_pages_count": generic_pgs,
                "duplicate_rows_count": max(0, rows_count - unique_models),
                "needs_review_count": needs_review,
                "status": status,
                "comment": "; ".join(comments),
            })

    total_targets = len(targets)
    total_found = len({(r.get("producer", ""), r.get("target_line", "")) for r in cleaned})
    not_found_count = total_targets - total_found
    print(f"\n{'=' * 60}")
    print(f"  DONE:")
    print(f"    Total rows: {len(cleaned)}")
    print(f"    Lines: {total_found}/{total_targets} ({total_found/total_targets*100:.1f}%)")
    print(f"    Not found: {not_found_count}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
