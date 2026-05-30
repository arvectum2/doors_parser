"""Postprocessing v2: deduplicate by model_name, upgrade bad URLs, pick best image.

Usage:
    python scripts/cleanup_table_images.py --input output/doors_catalog.csv --out output
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

try:
    import requests
except ImportError:
    requests = None

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")

FIELDS = [
    "door_type", "producer", "target_line", "model_name", "description",
    "material", "coating", "image_url", "source_url", "status", "comment",
]

SERVICE_MODEL_NAMES = {
    "каталог", "поиск", "меню", "корзина", "избранное", "сравнение",
    "фото", "галерея", "смотреть", "подробнее", "следующая", "предыдущая",
}

COMPONENT_START_KEYWORDS = [
    "коробка", "наличник", "добор", "погонаж", "телескоп", "притвор",
    "капитель", "карниз", "розетка", "плинтус", "подпятник",
]

CLEAN_DESC_PATTERNS = [
    re.compile(r"\(\*цена указана[^)]*\)", re.I),
    re.compile(r"\(цена указана[^)]*\)", re.I),
    re.compile(r"Цена указана[^.]*\.?\s*", re.I),
]

# Tilda CDN URL patterns
TILDA_RESIZE_RE = re.compile(r"^(https://)(thb|static)\.tildacdn\.com/([^/]+)/-/resize/\d+x/(.+)$")
TILDA_STATIC_RE = re.compile(r"^(https://)thb\.tildacdn\.com/([^/]+)/-/(.+)$")

# Regex to extract image dimensions from URL patterns like _640x480.jpg or /resize/800x/
DIM_RE = re.compile(r"[_/](\d{1,4})x(\d{2,4})[./]")


def _get_dimensions(url: str) -> tuple[int, int] | None:
    """Try to extract width,height from URL. Returns (w,h) or None."""
    m = DIM_RE.search(url)
    if m:
        w, h = int(m.group(1)), int(m.group(2))
        return w, h
    return None


def _is_large_enough(url: str) -> bool:
    """Check if URL contains explicit dimensions and they are ≥ 400px."""
    dims = _get_dimensions(url)
    if dims:
        w, h = dims
        return w >= 400 or h >= 400
    return False


# Bad image keywords — images will be cleared only if upgrade also fails
BAD_IMAGE_KEYWORDS = [
    "thumbnail", "small", "mini", "preview",
    "icon", "logo", "sprite", "placeholder", "no_image",
    "no-image", "noimage", "blank", "loader", "arrow",
    "cart_mob.webp", "cart", "search", "open.png",
    "colors_images", "pixel", ".svg",
]

SMALL_SIZE_PATTERNS = [
    re.compile(r"/20x[./]"), re.compile(r"/50x[./]"), re.compile(r"/60x[./]"),
    re.compile(r"/70x[./]"), re.compile(r"x96[./]"), re.compile(r"/96x"),
    re.compile(r"/96/"), re.compile(r"x100[./]"), re.compile(r"/100x"),
    re.compile(r"/100/"), re.compile(r"x120[./]"), re.compile(r"/120x"),
    re.compile(r"/120/"), re.compile(r"x150[./]"), re.compile(r"/150x"),
    re.compile(r"/150/"), re.compile(r"x200[./]"), re.compile(r"/200x"),
    re.compile(r"/200/"), re.compile(r"x300[./]"), re.compile(r"/300x"),
    re.compile(r"/300/"),
]


def norm_model(mn: str) -> str:
    if not mn:
        return ""
    mn = mn.lower().replace("ё", "е")
    mn = re.sub(r'["\'«»]', "", mn)
    mn = re.sub(r"\s+", " ", mn).strip()
    return mn


def normalize_key(text: str) -> str:
    return (text or "").lower().strip()


def has_bad_image(url: str) -> bool:
    if not url:
        return False
    u = url.lower()

    # If URL has explicit dimensions ≥ 400px, it's not a bad image
    if _is_large_enough(u):
        return False

    for kw in BAD_IMAGE_KEYWORDS:
        if kw in u:
            return True
    for pat in SMALL_SIZE_PATTERNS:
        if pat.search(u):
            return True
    return False


def is_service_model(mn: str) -> bool:
    return norm_model(mn) in SERVICE_MODEL_NAMES


def is_component_start(mn: str) -> bool:
    if not mn:
        return False
    m = mn.lower().strip()
    if "дверь" in m:
        return False
    for kw in COMPONENT_START_KEYWORDS:
        if m == kw:
            return True
        if m.startswith(kw + " ") or m.startswith(kw + "-") or m.startswith(kw + "—"):
            return True
    return False


def has_price(mn: str) -> bool:
    m = (mn or "").strip()
    if "₽" in m or "руб" in m.lower():
        return True
    if re.match(r"^от\s+\d[\d\s]*[₽р]", m, re.I):
        return True
    if re.match(r"^\d[\d\s]*\d\s*₽", m):
        return True
    return False


def clean_description(desc: str) -> str:
    if not desc:
        return ""
    result = desc
    for pat in CLEAN_DESC_PATTERNS:
        result = pat.sub("", result)
    result = re.sub(r"\s+", " ", result).strip()
    return result


def upgrade_image_url(url: str) -> str:
    """Try to upgrade a bad/thumbnail URL to a full-size version."""
    if not url:
        return url

    # Tilda CDN: thb.tildacdn.com/{id}/-/resize/{size}x/{filename}
    m = TILDA_RESIZE_RE.match(url)
    if m:
        scheme = m.group(1)
        img_id = m.group(3)
        filename = m.group(4)
        upgraded = f"{scheme}static.tildacdn.com/{img_id}/{filename}"
        return upgraded

    # Tilda CDN: thb.tildacdn.com/{id}/-/empty/{filename} — already original
    # No upgrade needed for /empty/ URLs

    return url


def score_image(url: str) -> int:
    """Score an image URL: higher is better."""
    if not url:
        return -1000
    u = url.lower()
    score = 0

    # Strong positive signals
    for kw in ["full", "original", "product/", "products/", "uploads/",
               "upload/", "gallery/", "catalog/", "doors/", "door/",
               "img/", "image/", "images/"]:
        if kw in u:
            score += 100

    # static.tildacdn.com is original size (upgraded from resize)
    if "static.tildacdn.com" in u:
        score += 200

    # File extension
    if u.endswith(".jpg") or u.endswith(".jpeg"):
        score += 50
    elif u.endswith(".png"):
        score += 30
    elif u.endswith(".webp"):
        score += 20

    # Path depth — deeper paths often have more specific images
    depth = u.count("/")
    if depth >= 4:
        score += 40

    # Size hints
    if "/big/" in u or "/large/" in u or "/full/" in u:
        score += 200
    if "/800x" in u or "/1024x" in u or "/1200x" in u:
        score += 150

    # Thumbnail hints — penalty
    if "/resize/" in u:
        score -= 100

    # HTTPS bonus
    if u.startswith("https://"):
        score += 20

    return score


def select_best_row(rows: list[dict]) -> dict:
    """From a group of rows (same producer+target_line+model_name), pick the best one."""
    scored = []
    for r in rows:
        url = r.get("image_url", "") or ""
        img_score = score_image(url)
        desc_bonus = 50 if r.get("description", "").strip() else 0
        empty_penalty = -200 if not url else 0
        total = img_score + desc_bonus + empty_penalty
        scored.append((total, r))
    scored.sort(key=lambda x: -x[0])
    return scored[0][1]


def read_csv(path: Path) -> list[dict]:
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


def write_csv(path: Path, rows: list[dict]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS, delimiter=";")
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in FIELDS})


def read_targets(path: Path):
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f, delimiter=";"))


def make_xlsx(path: Path, rows: list[dict], targets: list[dict],
              total_found: int, not_found: list[dict],
              producers: dict[str, int]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "parsed_catalog"
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
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    nf = wb.create_sheet("not_found")
    nf.append(["Тип двери", "Производитель", "Линейка", "Сайт", "Статус", "Комментарий"])
    found_set = {(normalize_key(r.get("producer", "")), normalize_key(r.get("target_line", ""))) for r in rows}
    for t in targets:
        key = (normalize_key(t.get("manufacturer", "")), normalize_key(t.get("target_line", "")))
        if key not in found_set:
            nf.append([t.get("door_type", ""), t.get("manufacturer", ""),
                       t.get("target_line", ""), t.get("source_site", ""),
                       "not_found", "проверить вручную / дописать селектор"])
    for cell in nf[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="9C6500")
    for col, w in enumerate([16, 24, 32, 45, 18, 45], 1):
        nf.column_dimensions[get_column_letter(col)].width = w
    nf.freeze_panes = "A2"
    nf.auto_filter.ref = nf.dimensions

    rp = wb.create_sheet("run_report")
    rp.append(["Показатель", "Значение"])
    rp.append(["Дата", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    rp.append(["Целевых линеек", len(targets)])
    rp.append(["Найдено строк", len(rows)])
    rp.append(["Найдено линеек", total_found])
    rp.append(["Не найдено линеек", len(not_found)])
    if rows:
        rp.append(["С моделями (status=found)", sum(1 for r in rows if r.get("status") == "found")])
        rp.append(["С моделями (status=needs_review)", sum(1 for r in rows if r.get("status") == "needs_review")])
    rp.append(["Комментарий", "Парсятся только линейки из target_lines.csv. Материал/покрытие пустые."])
    rp.append([])
    rp.append(["Производитель", "Моделей", "", ""])
    for man, cnt in sorted(producers.items()):
        rp.append([man, cnt, "", ""])
    for cell in rp[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="548235")
    rp.column_dimensions["A"].width = 32
    rp.column_dimensions["B"].width = 60
    wb.save(path)


def cleanup(args: argparse.Namespace) -> None:
    out_dir = args.out
    out_dir.mkdir(parents=True, exist_ok=True)

    targets_path = args.targets if args.targets else out_dir.parent / "data" / "target_lines.csv"
    targets_path = Path(targets_path)
    targets = read_targets(targets_path) if targets_path.exists() else []

    rows = read_csv(Path(args.input))
    before = len(rows)
    print(f"Rows loaded: {before}")

    # ── Step 1: Remove service model_names ──
    before_svc = len(rows)
    rows = [r for r in rows if not is_service_model(r.get("model_name", ""))]
    removed_svc = before_svc - len(rows)
    print(f"1. Service model_name removed: {removed_svc}")

    # ── Step 2: Remove component-start model_names ──
    before_comp = len(rows)
    rows = [r for r in rows if not is_component_start(r.get("model_name", ""))]
    removed_comp = before_comp - len(rows)
    print(f"2. Component-start rows removed: {removed_comp}")

    # ── Step 2b: Remove price model_names ──
    before_price = len(rows)
    rows = [r for r in rows if not has_price(r.get("model_name", ""))]
    removed_price = before_price - len(rows)
    print(f"2b. Price model_name removed: {removed_price}")

    # ── Step 2c: Clean descriptions ──
    for r in rows:
        r["description"] = clean_description(r.get("description", ""))
    for r in rows:
        for k in r:
            if r[k] and "₽" in r[k]:
                r[k] = r[k].replace("₽", "")

    # ── Step 3: Upgrade bad image URLs ──
    before_upgrade = len(rows)
    upgraded_count = 0
    for r in rows:
        url = r.get("image_url", "") or ""
        if has_bad_image(url):
            upgraded = upgrade_image_url(url)
            if upgraded != url:
                r["image_url"] = upgraded
                upgraded_count += 1
    print(f"3. Image URLs upgraded: {upgraded_count}")

    # ── Step 4: Group by (producer, target_line, model_name) ──
    # Different "вариант N" are different model_names → preserved separately
    # Same model_name across different source_urls → dedup to one row
    groups = defaultdict(list)
    for r in rows:
        key = (r.get("producer", ""), r.get("target_line", ""), r.get("model_name", ""))
        groups[key].append(r)

    dup_report_rows = []
    cleaned = []
    duplicate_groups_found = 0
    duplicate_rows_removed = 0
    groups_with_good_image = 0
    groups_without_good_image = 0

    for key, grp in sorted(groups.items()):
        producer, tline, mdl = key
        group_size = len(grp)

        if group_size == 1:
            r = grp[0]
            url = r.get("image_url", "") or ""
            if not url:
                r["status"] = "needs_review"
                r["comment"] = ("; ".join(filter(None, [r.get("comment", ""), "image_missing"]))).strip()
            elif has_bad_image(url):
                r["image_url"] = ""
                r["status"] = "needs_review"
                r["comment"] = ("; ".join(filter(None, [r.get("comment", ""), "bad_image_removed"]))).strip()
            else:
                r["status"] = "found"
            cleaned.append(r)
        else:
            duplicate_groups_found += 1

            # Check if ANY row in the group has a good image
            has_good = False
            for r in grp:
                url = r.get("image_url", "") or ""
                if url and not has_bad_image(url):
                    has_good = True
                    break

            if has_good:
                groups_with_good_image += 1
                best = select_best_row(grp)
                best_comment = best.get("comment", "")
                if "selected_best_image_from_duplicates" not in best_comment:
                    best_comment = ("; ".join(filter(None, [best_comment,
                        f"selected_best_image_from_duplicates:{group_size}"]))).strip()
                    best["comment"] = best_comment

                best_url = best.get("image_url", "") or ""
                if not best_url:
                    best["status"] = "needs_review"
                elif has_bad_image(best_url):
                    best["image_url"] = ""
                    best["status"] = "needs_review"
                else:
                    best["status"] = "found"

                removed_urls = [r.get("image_url", "") for r in grp if r is not best]
                dup_report_rows.append({
                    "producer": producer,
                    "target_line": tline,
                    "normalized_model_name": norm_model(mdl),
                    "model_name": mdl,
                    "source_url": best.get("source_url", ""),
                    "duplicates_count_before": group_size,
                    "kept_model_name": best.get("model_name", ""),
                    "kept_image_url": best_url or "",
                    "removed_image_urls": "; ".join(str(u) for u in removed_urls if u),
                    "reason": "photo_duplicate",
                })
                duplicate_rows_removed += group_size - 1
                cleaned.append(best)
            else:
                # No good image — keep only 1 row, mark needs_review
                groups_without_good_image += 1
                best = select_best_row(grp)
                best["image_url"] = ""
                best["status"] = "needs_review"
                best_comment = best.get("comment", "")
                if "no_good_image_found" not in best_comment:
                    best_comment = ("; ".join(filter(None, [best_comment,
                        f"no_good_image_found; selected_best_image_from_duplicates:{group_size}"]))).strip()
                    best["comment"] = best_comment

                removed_urls = [r.get("image_url", "") for r in grp if r is not best]
                dup_report_rows.append({
                    "producer": producer,
                    "target_line": tline,
                    "normalized_model_name": norm_model(mdl),
                    "model_name": mdl,
                    "source_url": best.get("source_url", ""),
                    "duplicates_count_before": group_size,
                    "kept_model_name": best.get("model_name", ""),
                    "kept_image_url": "",
                    "removed_image_urls": "; ".join(str(u) for u in removed_urls if u),
                    "reason": "no_good_image",
                })
                duplicate_rows_removed += group_size - 1
                cleaned.append(best)

    print(f"4. Groups by (producer, target_line, model_name): {len(groups)}")
    print(f"   Groups with >1 row: {duplicate_groups_found}")
    print(f"   With good image: {groups_with_good_image}")
    print(f"   Without good image: {groups_without_good_image}")
    print(f"   Rows removed: {duplicate_rows_removed}")

    # ── Step 5: Write output files ──
    found_set = set()
    producer_counts: dict[str, int] = {}
    for r in cleaned:
        man = r.get("producer", "")
        line = r.get("target_line", "")
        found_set.add((normalize_key(man), normalize_key(line)))
        if man:
            producer_counts[man] = producer_counts.get(man, 0) + 1

    total_found = len(found_set)
    not_found_entries = []
    for t in targets:
        key = (normalize_key(t.get("manufacturer", "")), normalize_key(t.get("target_line", "")))
        if key not in found_set:
            not_found_entries.append({
                "door_type": t.get("door_type", ""),
                "manufacturer": t.get("manufacturer", ""),
                "target_line": t.get("target_line", ""),
                "source_site": t.get("source_site", ""),
                "status": "not_found",
                "comment": "page_not_found — нет подходящей страницы на сайте",
            })

    for r in cleaned:
        if not r.get("description"):
            r["description"] = f"Модель из линейки {r.get('target_line', '')}, производитель {r.get('producer', '')}"

    # Write doors_catalog.csv
    write_csv(out_dir / "doors_catalog.csv", cleaned)

    # Write wordpress_import.csv
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
    nf_fields = ["door_type", "manufacturer", "target_line", "source_site", "status", "comment"]
    with (out_dir / "not_found.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=nf_fields, delimiter=";")
        w.writeheader()
        w.writerows(not_found_entries)
    print(f"5. Not found: {len(not_found_entries)}")

    # Write duplicates_report.csv
    dup_fields = ["producer", "target_line", "normalized_model_name", "model_name",
                  "source_url", "duplicates_count_before", "kept_model_name",
                  "kept_image_url", "removed_image_urls", "reason"]
    with (out_dir / "duplicates_report.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=dup_fields, delimiter=";")
        w.writeheader()
        w.writerows(dup_report_rows)
    print(f"6. Duplicates report: {len(dup_report_rows)} entries")

    # Count images
    bad_images_removed = sum(1 for r in cleaned if "bad_image_removed" in r.get("comment", ""))
    image_missing_count = sum(1 for r in cleaned if "image_missing" in r.get("comment", ""))
    no_good_image_count = sum(1 for r in cleaned if "no_good_image_found" in r.get("comment", ""))

    # Write qa_report.csv
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
            grp = grouped.get((normalize_key(man), normalize_key(line)), [])
            rows_count = len(grp)
            unique_models = len(set(norm_model(r.get("model_name", "")) for r in grp))
            unique_urls = len(set(r.get("source_url", "") for r in grp))
            bad_imgs = sum(1 for r in grp if "bad_image" in r.get("comment", ""))
            generic_pgs = 0
            needs_review = sum(1 for r in grp if r.get("status") == "needs_review")
            status = "found" if grp else "not_found"
            comments = []
            if not grp:
                comments.append("no rows")
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
    print(f"7. QA report written")

    # Write run_report.json
    after = len(cleaned)
    report = {
        "start_time": "cleanup_table_images.py v2",
        "end_time": datetime.now().isoformat(),
        "total_targets": len(targets),
        "total_lines_found": total_found,
        "total_items": after,
        "lines_not_found": len(not_found_entries),
        "producers": producer_counts,
        "errors": [],
        "postprocess": {
            "rows_before_cleanup": before,
            "rows_after_cleanup": after,
            "service_rows_removed": removed_svc,
            "component_rows_removed": removed_comp,
            "price_rows_removed": removed_price,
            "image_urls_upgraded": upgraded_count,
            "duplicate_groups_found": duplicate_groups_found,
            "duplicate_rows_removed": duplicate_rows_removed,
            "groups_with_good_image": groups_with_good_image,
            "groups_without_good_image": groups_without_good_image,
            "bad_images_removed": bad_images_removed,
            "image_missing_count": image_missing_count,
            "no_good_image_count": no_good_image_count,
            "duplicates_report_path": str(out_dir / "duplicates_report.csv"),
            "cleanup_finished_at": datetime.now().isoformat(),
        },
    }
    (out_dir / "run_report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"8. Report written")

    # Write XLSX
    make_xlsx(out_dir / "doors_catalog.xlsx", cleaned, targets, total_found, not_found_entries, producer_counts)

    # Summary
    print(f"\n{'=' * 60}")
    print(f"  DONE:")
    print(f"    Rows: {before} → {after} (removed {before-after})")
    print(f"    Lines: {total_found}/{len(targets)} ({total_found/len(targets)*100:.1f}%)")
    print(f"    Not found: {len(not_found_entries)}")
    print(f"    Producers: {len(producer_counts)}")
    for man, cnt in sorted(producer_counts.items()):
        print(f"      {man}: {cnt}")
    print(f"{'=' * 60}")

    # Verification
    print(f"\n--- Verification ---")
    svc_rem = sum(1 for r in cleaned if is_service_model(r.get("model_name", "")))
    print(f"  Service model_name: {svc_rem} (should be 0)")

    comp_rem = sum(1 for r in cleaned if is_component_start(r.get("model_name", "")))
    print(f"  Component-start: {comp_rem} (should be 0)")

    # Check dupes by (producer, target_line, model_name)
    seen = set()
    dupes = 0
    for r in cleaned:
        key = (r["producer"], r["target_line"], r["model_name"])
        if key in seen:
            dupes += 1
        else:
            seen.add(key)
    print(f"  Dupes by (producer, target_line, model_name): {dupes} (should be 0)")

    bad_remaining = sum(1 for r in cleaned if has_bad_image(r.get("image_url", "")))
    print(f"  Bad image_url remaining: {bad_remaining} (should be 0)")


def main():
    ap = argparse.ArgumentParser(description="Cleanup v2: dedup by model_name, upgrade bad URLs")
    ap.add_argument("--input", type=Path, default=Path("output/doors_catalog.csv"),
                    help="Input CSV path")
    ap.add_argument("--out", type=Path, default=Path("output"),
                    help="Output directory")
    ap.add_argument("--targets", type=Path, default=None,
                    help="Target lines CSV (for QA report)")
    args = ap.parse_args()
    cleanup(args)


if __name__ == "__main__":
    main()
